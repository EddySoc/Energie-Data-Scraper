"""
VREG V-test Open Data Importer
===============================
Downloads the official monthly V-test tariff data published by the Vlaamse
Nutsregulator (VNR/VREG) and exports a clean, normalised Excel comparison.

Source: https://www.vlaamsenutsregulator.be/cijfers/v-test-data-en-energieprijscurves
Data:   All residential fixed, variable and dynamic products (excl. VAT).
        The current workbook is discovered from the regulator's publication page.

All prices in the source are in c€/kWh or €/MWh; this script converts
everything to €/kWh so you can compare on a level playing field.
"""

from __future__ import annotations

import argparse
import io
import html
import re
import sys
import urllib.request
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urljoin, urlparse

import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter

# ── Source configuration ──────────────────────────────────────────────────────

VREG_DATA_PAGE = (
    "https://www.vlaamsenutsregulator.be/cijfers/v-test-data-en-energieprijscurves"
)

# Standard residential consumption profile (kWh/year) used for the estimate.
ANNUAL_KWH_ELECTRICITY = 3500.0
ANNUAL_KWH_GAS = 17000.0

# Month ordering for sorting
MONTHS_NL = {
    "jan": 1, "feb": 2, "mrt": 3, "apr": 4, "mei": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "okt": 10, "nov": 11, "dec": 12,
}

# Price components that form the real energy cost (excl. VAT, excl. grid).
# The VREG data splits the tariff into raw energy + obligatory green/WKK levies.
ELEC_RATE_COMPONENTS = {
    "Enkelvoudige meter dagtarief (c€/kWh)": "enkel_rate",
    "Kosten groene stroom (c€/kWh)":         "groene_stroom",
    "Kosten WKK (c€/kWh)":                   "wkk",
    "Tweevoudige meter dagtarief (c€/kWh)":  "dag_rate",
    "Tweevoudige meter nachttarief (c€/kWh)":"nacht_rate",
    "Uitsluitend nachtmeter (c€/kWh)":       "excl_nacht_rate",
    "Vaste vergoeding (€)":                  "vaste_vergoeding",
}

GAS_RATE_COMPONENTS = {
    "Afname (c€/kWh)":          "gas_rate",
    "Vaste vergoeding (€)":     "gas_vaste_vergoeding",
    # Some older records use different label
    "Dagtarief (c€/kWh)":       "gas_rate",
    "Enkelvoudige meter dagtarief (c€/kWh)": "gas_rate",
}

APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
DEFAULT_DATA_DIR = APP_DIR / "scraper_data"
OUTPUT_FILE = DEFAULT_DATA_DIR / "be_vreg_vergelijking.xlsx"


# ── Download ──────────────────────────────────────────────────────────────────

def discover_latest_vreg_data_url(page_url: str = VREG_DATA_PAGE) -> str:
    """Find the newest monthly V-test workbook linked from the regulator's page."""
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(page_url, headers=headers, timeout=45)
    response.raise_for_status()
    candidates: list[tuple[int, str]] = []
    for match in re.finditer(
        r"""href\s*=\s*["']([^"']+\.xlsx(?:\?[^"']*)?)["']""",
        response.text,
        flags=re.IGNORECASE,
    ):
        href = html.unescape(match.group(1))
        url = urljoin(response.url, href)
        decoded_url = unquote(url).casefold()
        if "v-test" not in decoded_url or "energieprijscurve" in decoded_url:
            continue
        filename = unquote(urlparse(url).path).rsplit("/", 1)[-1]
        period_match = re.search(r"(?P<year>20\d{2})(?P<month>0[1-9]|1[0-2])", filename)
        if period_match is None:
            continue
        period = int(period_match.group("year") + period_match.group("month"))
        candidates.append((period, url))

    if not candidates:
        raise RuntimeError(
            f"Geen maandelijkse V-test Excel-link gevonden op de VNR-pagina: {page_url}"
        )
    latest_period = max(period for period, _ in candidates)
    latest_urls = sorted({url for period, url in candidates if period == latest_period})
    if len(latest_urls) != 1:
        raise RuntimeError(
            f"Meerdere V-testbestanden gevonden voor {latest_period}: {latest_urls}"
        )
    print(f"  Latest V-test workbook: {latest_period} ({latest_urls[0]})")
    return latest_urls[0]


def download_vreg_data(url: str | None = None) -> pd.ExcelFile:
    headers = {"User-Agent": "Mozilla/5.0"}
    selected_url = url or discover_latest_vreg_data_url()
    response = requests.get(selected_url, headers=headers, timeout=45)
    response.raise_for_status()
    if not response.content.startswith(b"PK"):
        raise ValueError(f"VNR download is not an Excel workbook: {selected_url}")
    print(f"  Downloaded {len(response.content) // 1024} KB")
    return pd.ExcelFile(io.BytesIO(response.content))


# ── Parse ─────────────────────────────────────────────────────────────────────

def parse_fixed_products(xl: pd.ExcelFile) -> pd.DataFrame:
    """Combine fixed-price sheets across every year in the published workbook."""
    frames = []
    for sheet in xl.sheet_names:
        if "vast" in sheet.lower() and "excl" in sheet.lower():
            df = xl.parse(sheet)
            frames.append(df)

    if not frames:
        raise ValueError("No fixed-price sheets found in the workbook.")

    combined = pd.concat(frames, ignore_index=True)
    combined.columns = [str(c).strip() for c in combined.columns]
    return combined


def get_latest_month_residential(df: pd.DataFrame) -> pd.DataFrame:
    """Keep only the latest year/month for residential (Woning) customers."""
    woning = df[df["Segment"] == "Woning"].copy()

    if woning.empty:
        raise ValueError("No 'Woning' segment found in the data.")

    woning["maand_nr"] = woning["Maand"].str.lower().map(MONTHS_NL).fillna(0).astype(int)
    woning["sort_key"] = woning["Jaar"].astype(int) * 100 + woning["maand_nr"]
    latest_key = woning["sort_key"].max()
    latest = woning[woning["sort_key"] == latest_key].copy()

    latest_label = f"{latest['Maand'].iloc[0]} {int(latest['Jaar'].iloc[0])}"
    print(f"  Latest data period: {latest_label}  ({len(latest)} rows)")
    return latest, latest_label


def pivot_to_wide(df: pd.DataFrame) -> pd.DataFrame:
    """Pivot from long (one row per price component) to wide (one row per product)."""
    group_cols = ["Handelsnaam", "Productnaam", "Energietype", "Contracttype",
                  "Vast/variabel/dynamisch"]
    group_cols = [c for c in group_cols if c in df.columns]

    pivot = df.pivot_table(
        index=group_cols,
        columns="Prijsonderdeel",
        values="Prijs",
        aggfunc="first",
    ).reset_index()
    pivot.columns.name = None
    pivot.columns = [str(c).strip() for c in pivot.columns]
    return pivot


# ── Normalise ─────────────────────────────────────────────────────────────────

def safe_float(val) -> float | None:
    if pd.isna(val):
        return None
    try:
        return float(val)
    except Exception:
        return None


def build_electricity_comparison(pivot: pd.DataFrame) -> pd.DataFrame:
    """Extract electricity products and calculate total rates in €/kWh."""
    elec = pivot[pivot["Energietype"] == "Elektriciteit"].copy()
    rows = []

    for _, row in elec.iterrows():
        green_raw = safe_float(row.get("Kosten groene stroom (c€/kWh)"))
        wkk_raw = safe_float(row.get("Kosten WKK (c€/kWh)"))
        groene = green_raw or 0.0
        wkk = wkk_raw or 0.0
        extra  = groene + wkk

        enkel_raw  = safe_float(row.get("Enkelvoudige meter dagtarief (c€/kWh)"))
        dag_raw    = safe_float(row.get("Tweevoudige meter dagtarief (c€/kWh)"))
        nacht_raw  = safe_float(row.get("Tweevoudige meter nachttarief (c€/kWh)"))
        nachtm_raw = safe_float(row.get("Uitsluitend nachtmeter (c€/kWh)"))
        vaste      = safe_float(row.get("Vaste vergoeding (€)"))

        # Convert c€/kWh → €/kWh
        enkel  = (enkel_raw  + extra) / 100.0 if enkel_raw  is not None else None
        dag    = (dag_raw    + extra) / 100.0 if dag_raw    is not None else None
        nacht  = (nacht_raw  + extra) / 100.0 if nacht_raw  is not None else None
        nachtm = (nachtm_raw + extra) / 100.0 if nachtm_raw is not None else None

        rows.append({
            "Leverancier":      row.get("Handelsnaam"),
            "Product stroom":   row.get("Productnaam"),
            "Contracttype":     row.get("Contracttype"),
            "Enkel €/kWh":      enkel,
            "Dag €/kWh":        dag,
            "Nacht €/kWh":      nacht,
            "Excl.nacht €/kWh": nachtm,
            "Stroom vaste €/jr": vaste,
            "Groenestroom €/kWh": green_raw / 100.0 if green_raw is not None else None,
            "WKK €/kWh": wkk_raw / 100.0 if wkk_raw is not None else None,
        })

    return pd.DataFrame(rows)


def build_gas_comparison(pivot: pd.DataFrame) -> pd.DataFrame:
    """Extract gas products and calculate total rates in €/kWh."""
    gas = pivot[pivot["Energietype"] == "Gas"].copy()
    rows = []

    # Find the gas rate column (label differs between old and new data)
    gas_rate_candidates = [
        "Afname (c€/kWh)",
        "Enkelvoudige meter dagtarief (c€/kWh)",
        "Dagtarief (c€/kWh)",
    ]

    for _, row in gas.iterrows():
        raw = None
        for col in gas_rate_candidates:
            raw = safe_float(row.get(col))
            if raw is not None:
                break

        gas_eur_kwh = raw / 100.0 if raw is not None else None
        vaste       = safe_float(row.get("Vaste vergoeding (€)"))

        rows.append({
            "Leverancier":    row.get("Handelsnaam"),
            "Product gas":    row.get("Productnaam"),
            "Gas €/kWh":      gas_eur_kwh,
            "Gas vaste €/jr": vaste,
        })

    return pd.DataFrame(rows)


def build_dual_fuel_comparison(
    elec_df: pd.DataFrame, gas_df: pd.DataFrame
) -> pd.DataFrame:
    """
    Create a dual-fuel comparison: for each electricity product find the
    cheapest available gas product from the same supplier.
    One row = one electricity product × cheapest gas product.
    """
    # Find cheapest gas per supplier
    gas_df_clean = gas_df.dropna(subset=["Gas €/kWh"]).copy()
    cheapest_gas = (
        gas_df_clean
        .sort_values("Gas €/kWh")
        .drop_duplicates(subset=["Leverancier"], keep="first")
        [["Leverancier", "Product gas", "Gas €/kWh", "Gas vaste €/jr"]]
    )

    merged = elec_df.merge(cheapest_gas, on="Leverancier", how="left")

    # Estimate: 3500 kWh single rate + 17000 kWh gas + standing charges
    def estimate(row):
        elec_rate = row.get("Enkel €/kWh") or (
            0.6 * row.get("Dag €/kWh", 0 or 0) + 0.4 * row.get("Nacht €/kWh", 0 or 0)
            if row.get("Dag €/kWh") and row.get("Nacht €/kWh") else None
        )
        gas_rate  = row.get("Gas €/kWh")
        if elec_rate is None or gas_rate is None:
            return None
        e_vaste = row.get("Stroom vaste €/jr") or 0.0
        g_vaste = row.get("Gas vaste €/jr") or 0.0
        return (
            ANNUAL_KWH_ELECTRICITY * elec_rate
            + ANNUAL_KWH_GAS       * gas_rate
            + e_vaste + g_vaste
        )

    merged["Schatting €/jr*"] = merged.apply(estimate, axis=1)
    merged["Volledig?"] = merged.apply(
        lambda r: "Ja" if (r.get("Enkel €/kWh") or r.get("Dag €/kWh")) and r.get("Gas €/kWh") else "Nee",
        axis=1
    )

    return merged.sort_values(
        ["Volledig?", "Schatting €/jr*"],
        ascending=[False, True],
        na_position="last"
    ).reset_index(drop=True)


# ── Export ────────────────────────────────────────────────────────────────────

def style_comparison_sheet(ws, nrows: int, ncols: int, data_period: str) -> None:
    hdr_fill   = PatternFill(start_color="1F5C2E", end_color="1F5C2E", fill_type="solid")
    hdr_font   = Font(bold=True, color="FFFFFF", size=11)
    green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    orange_fill= PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    link_font  = Font(color="0563C1", underline="single")
    border     = Border(
        left  =Side(style="thin"), right =Side(style="thin"),
        top   =Side(style="thin"), bottom=Side(style="thin"),
    )

    # Header row
    for ci in range(1, ncols + 1):
        cell = ws.cell(row=1, column=ci)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    # Data rows
    headers = [ws.cell(row=1, column=ci).value for ci in range(1, ncols + 1)]
    volledig_col  = next((i+1 for i, h in enumerate(headers) if h == "Volledig?"),  None)
    schat_col     = next((i+1 for i, h in enumerate(headers) if h == "Schatting €/jr*"), None)
    eur_kwh_cols  = [i+1 for i, h in enumerate(headers) if h and "€/kWh" in str(h)]
    vaste_cols    = [i+1 for i, h in enumerate(headers) if h and "vaste" in str(h).lower()]

    for ri in range(2, nrows + 2):
        for ci in range(1, ncols + 1):
            cell = ws.cell(row=ri, column=ci)
            cell.border = border
            cell.alignment = Alignment(horizontal="left", vertical="center")

        if volledig_col:
            vc = ws.cell(row=ri, column=volledig_col)
            vc.fill = green_fill if vc.value == "Ja" else orange_fill
            vc.alignment = Alignment(horizontal="center")

        if schat_col:
            sc = ws.cell(row=ri, column=schat_col)
            if sc.value is not None:
                sc.number_format = "€#,##0.00"

        for ci in eur_kwh_cols:
            ws.cell(row=ri, column=ci).number_format = "0.00000"
        for ci in vaste_cols:
            ws.cell(row=ri, column=ci).number_format = "€#,##0.00"

    # Column widths
    col_widths = {
        "Leverancier": 24, "Product stroom": 32, "Product gas": 30,
        "Contracttype": 14, "Enkel €/kWh": 13, "Dag €/kWh": 12,
        "Nacht €/kWh": 12, "Excl.nacht €/kWh": 14, "Gas €/kWh": 12,
        "Stroom vaste €/jr": 16, "Gas vaste €/jr": 16,
        "Schatting €/jr*": 16, "Volledig?": 11,
    }
    for ci, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = col_widths.get(str(h), 16)

    ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{nrows + 1}"
    ws.freeze_panes = "D2"

    # Disclaimer
    disc_row = nrows + 3
    disc = ws.cell(row=disc_row, column=1)
    disc.value = (
        f"Bron: Vlaamse Nutsregulator (VNR) — V-test® open data | Periode: {data_period} | "
        f"Schatting op basis van {int(ANNUAL_KWH_ELECTRICITY)} kWh stroom (enkeltarief) + "
        f"{int(ANNUAL_KWH_GAS)} kWh gas per jaar | "
        "Alle prijzen excl. btw, excl. distributienettarieven en heffingen. "
        "Groenestroomcertificaten en WKK-kosten zijn WEL inbegrepen in het energietarief."
    )
    disc.font = Font(italic=True, size=9, color="666666")
    try:
        ws.merge_cells(start_row=disc_row, start_column=1,
                       end_row=disc_row, end_column=ncols)
    except Exception:
        pass


def export_to_excel(
    dual_fuel_df: pd.DataFrame,
    elec_df: pd.DataFrame,
    gas_df: pd.DataFrame,
    data_period: str,
    output_path: str = OUTPUT_FILE,
) -> str:
    try:
        writer = pd.ExcelWriter(output_path, engine="openpyxl")
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = output_path.replace(".xlsx", f"_{ts}.xlsx")
        writer = pd.ExcelWriter(output_path, engine="openpyxl")

    with writer:
        dual_fuel_df.to_excel(writer, sheet_name="Vergelijking", index=False)
        elec_df.to_excel(writer, sheet_name="Stroom_producten", index=False)
        gas_df.to_excel(writer, sheet_name="Gas_producten", index=False)

        workbook = writer.book

        ws = workbook["Vergelijking"]
        style_comparison_sheet(ws, len(dual_fuel_df), len(dual_fuel_df.columns), data_period)

        for sheet_name, df in [("Stroom_producten", elec_df), ("Gas_producten", gas_df)]:
            ws2 = workbook[sheet_name]
            hdr_fill = PatternFill(start_color="2E5396", end_color="2E5396", fill_type="solid")
            hdr_font = Font(bold=True, color="FFFFFF", size=10)
            border   = Border(left=Side(style="thin"), right=Side(style="thin"),
                               top=Side(style="thin"), bottom=Side(style="thin"))
            for ci in range(1, len(df.columns) + 1):
                c = ws2.cell(row=1, column=ci)
                c.fill = hdr_fill; c.font = hdr_font
                c.alignment = Alignment(horizontal="center", wrap_text=True)
                c.border = border
            for row in ws2.iter_rows(min_row=2, max_row=len(df)+1):
                for cell in row:
                    cell.border = border
            ws2.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df)+1}"
            ws2.freeze_panes = "C2"
            for ci in range(1, len(df.columns)+1):
                ws2.column_dimensions[get_column_letter(ci)].width = 20

    return output_path


# ── Main ──────────────────────────────────────────────────────────────────────

def main(output_file: Path = OUTPUT_FILE) -> None:
    print("VREG V-test Open Data Importer")
    print()
    print("1. Downloading official VREG tariff data...")
    xl = download_vreg_data()

    print("2. Parsing fixed-price product sheets...")
    fixed_df = parse_fixed_products(xl)

    print("3. Filtering latest month, residential (Woning)...")
    latest_df, data_period = get_latest_month_residential(fixed_df)

    print("4. Pivoting to wide format...")
    pivot_df = pivot_to_wide(latest_df)

    print("5. Building electricity comparison table...")
    elec_df = build_electricity_comparison(pivot_df)
    print(f"   {len(elec_df)} electricity products from {elec_df['Leverancier'].nunique()} suppliers")

    print("6. Building gas comparison table...")
    gas_df = build_gas_comparison(pivot_df)
    print(f"   {len(gas_df)} gas products from {gas_df['Leverancier'].nunique()} suppliers")

    print("7. Building dual-fuel comparison (stroom + goedkoopste gas per leverancier)...")
    dual_df = build_dual_fuel_comparison(elec_df, gas_df)
    complete = int((dual_df["Volledig?"] == "Ja").sum())
    print(f"   {complete} of {len(dual_df)} products are fully comparable (both gas + electricity)")

    print("8. Exporting to Excel...")
    output_file.parent.mkdir(parents=True, exist_ok=True)
    out_path = export_to_excel(dual_df, elec_df, gas_df, data_period, str(output_file))
    print(f"\nSaved: {out_path}")
    print()
    print("Open the tab 'Vergelijking' voor de gesorteerde vergelijking.")
    print("Kolom 'Schatting €/jr*' is sorteerbaar van goedkoopst naar duurste.")
    print()
    print("NOOT: Prijzen excl. btw, excl. distributienet, excl. heffingen.")
    print("      Groenestroomkosten + WKK zijn WEL meegerekend in het tarief.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Download and normalize the latest VREG V-test data.")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_DATA_DIR,
        help="Directory for the generated VREG comparison workbook.",
    )
    args = parser.parse_args()
    main(args.output_dir / OUTPUT_FILE.name)
