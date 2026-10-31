import contextlib
import math
import re
import shutil
import subprocess
import sys
import threading
import traceback
import unicodedata
import webbrowser
import tkinter as tk
from zipfile import BadZipFile
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.properties import PageSetupProperties

from energy_scraper import (
    main as run_energy_scraper,
    parse_energy_rate_to_eur_per_kwh,
    parse_standing_charge_to_eur_per_year,
)
from fluvius_nettarieven_importer import (
    OUTPUT_CSV as FLUVIUS_OUTPUT_CSV,
    update_fluvius_rates,
)
from vreg_importer import OUTPUT_FILE as VREG_OUTPUT_FILE
from vreg_importer import main as run_vreg_importer

BASE_DIR = Path(__file__).resolve().parent
APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else BASE_DIR
)


def _prepare_portable_data_dirs() -> None:
    scraper_data_dir = APP_DIR / "scraper_data"
    user_data_dir = APP_DIR / "user_data"
    scraper_data_dir.mkdir(parents=True, exist_ok=True)
    user_data_dir.mkdir(parents=True, exist_ok=True)

    if getattr(sys, "frozen", False):
        bundle_root = getattr(sys, "_MEIPASS", None)
        bundled_data_dir = Path(bundle_root) / "scraper_data" if bundle_root else None
        if bundled_data_dir is not None and bundled_data_dir.is_dir():
            for source in bundled_data_dir.rglob("*"):
                if source.is_file():
                    destination = scraper_data_dir / source.relative_to(bundled_data_dir)
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    if not destination.exists():
                        shutil.copy2(source, destination)

CSV_TABS = [
    ("Leveranciers", "be_supplier_qualification.csv"),
    ("Gevonden URLs", "be_supplier_discovered_urls.csv"),
    ("Curated URLs", "be_supplier_curated_urls.csv"),
    ("Beste URL", "be_supplier_best_urls.csv"),
    ("Vergelijkers", "be_comparators_regulators_urls.csv"),
    ("Ruwe prijzen", "be_energy_prices_raw.csv"),
    ("Prijzen", "be_energy_price_comparison.csv"),
    ("Extractie-details", "be_energy_price_extraction_details.csv"),
]

VREG_COLUMNS = {
    "Enkel €/kWh": "electricity_single_rate",
    "Dag €/kWh": "electricity_day_rate",
    "Nacht €/kWh": "electricity_night_rate",
    "Gas €/kWh": "gas_rate",
}

DASHBOARD_SNAPSHOT_VERSION = 1
SCAN_LOG_NAME = "scan_log.txt"


class _LogWriter:
    """File-like object that forwards text written by in-process scans to the scan log."""

    def __init__(self, target) -> None:
        self.target = target

    def write(self, text: str) -> int:
        if self.target is not None and text:
            self.target.write(text)
            self.target.flush()
        return len(text)

    def flush(self) -> None:
        pass


def _finite_number(value: object) -> float | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    try:
        number = float(str(value))
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _supplier_key(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"\s+", " ", text).strip().casefold()
    if text.startswith("mega"):
        return "mega"
    if text.startswith("engie"):
        return "engie"
    if text.startswith("eneco"):
        return "eneco"
    if text.startswith("totalenergies"):
        return "totalenergies"
    if text.startswith("power online"):
        return "mega"
    if text.startswith("octa+"):
        return "octa+"
    return text


def _vreg_price(value: object, unit: str) -> str | None:
    number = _finite_number(value)
    if number is None or number < 0:
        return None
    return f"{number:.6f}".rstrip("0").rstrip(".") + f" {unit}"


PRICE_DATA_COLUMNS = (
    "electricity_single_rate",
    "electricity_day_rate",
    "electricity_night_rate",
    "gas_rate",
    "standing_charge",
    "electricity_standing_charge",
    "gas_standing_charge",
)


def load_supplier_discounts(path: Path) -> dict[str, tuple[float, float]]:
    """Read per-supplier yearly discounts (electricity, gas) in EUR from a CSV file."""
    if not path.is_file():
        return {}
    try:
        df = pd.read_csv(path, encoding="utf-8-sig", sep=None, engine="python")
    except (OSError, pd.errors.ParserError, pd.errors.EmptyDataError):
        return {}
    if "supplier" not in df.columns:
        return {}
    discounts: dict[str, tuple[float, float]] = {}
    for row in df.to_dict("records"):
        if row.get("supplier") is None or pd.isna(row.get("supplier")):
            continue
        values = []
        for column in ("electricity_discount_eur_year", "gas_discount_eur_year"):
            value = _finite_number(str(row.get(column, "")).replace(",", "."))
            values.append(value if value is not None and value > 0 else 0.0)
        key = _supplier_key(row["supplier"])
        previous = discounts.get(key, (0.0, 0.0))
        discounts[key] = (previous[0] + values[0], previous[1] + values[1])
    return discounts


def drop_suppliers_without_data(prices: pd.DataFrame) -> pd.DataFrame:
    """Remove suppliers for which no price or standing charge is known."""
    columns = [column for column in PRICE_DATA_COLUMNS if column in prices.columns]
    if prices.empty or not columns:
        return prices
    has_data = prices[columns].apply(
        lambda col: col.notna() & col.astype(str).str.strip().ne("")
    ).any(axis=1)
    return prices[has_data].reset_index(drop=True)


def combine_price_sources(
    scraped: pd.DataFrame,
    vreg: pd.DataFrame,
) -> pd.DataFrame:
    """Use the best official VREG product per supplier, with scrape-only fallbacks."""
    if vreg.empty:
        return scraped.copy()

    vreg = vreg.copy()
    required = {"Leverancier", "Product stroom", "Product gas", *VREG_COLUMNS}
    missing = required.difference(vreg.columns)
    if missing:
        raise ValueError(f"VREG-vergelijking mist kolommen: {', '.join(sorted(missing))}")

    vreg = vreg[vreg["Leverancier"].notna()].copy()
    vreg["Leverancier"] = vreg["Leverancier"].astype(str).str.strip()
    vreg = vreg[~vreg["Leverancier"].str.startswith("Bron:", na=False)]
    # Teruglevering-rijen bevatten injectievergoedingen, geen afnametarieven.
    is_feed_in = vreg["Product stroom"].astype(str).str.contains(
        "terugleververgoeding|teruglevering",
        case=False,
        na=False,
    )
    if "Contracttype" in vreg.columns:
        is_feed_in |= vreg["Contracttype"].astype(str).str.contains(
            "teruglever|injectie",
            case=False,
            na=False,
        )
    vreg = vreg[~is_feed_in].copy()
    vreg["_supplier_key"] = vreg["Leverancier"].map(_supplier_key)

    def annual_estimate(row: pd.Series) -> float:
        single = _finite_number(row.get("Enkel €/kWh"))
        day = _finite_number(row.get("Dag €/kWh"))
        night = _finite_number(row.get("Nacht €/kWh"))
        gas = _finite_number(row.get("Gas €/kWh"))
        if single is not None:
            electricity = single
        elif day is not None and night is not None:
            electricity = 0.6 * day + 0.4 * night
        elif day is not None:
            electricity = day
        elif night is not None:
            electricity = night
        else:
            electricity = math.inf
        if gas is None:
            return math.inf
        fixed = sum(
            value
            for value in (
                _finite_number(row.get("Stroom vaste €/jr")),
                _finite_number(row.get("Gas vaste €/jr")),
            )
            if value is not None
        )
        return 3500 * electricity + 17000 * gas + fixed

    vreg["_annual_estimate"] = vreg.apply(annual_estimate, axis=1)
    vreg["_available_rates"] = vreg[list(VREG_COLUMNS)].notna().sum(axis=1)
    vreg = vreg.sort_values(
        ["_supplier_key", "_annual_estimate", "_available_rates"],
        ascending=[True, True, False],
    )
    best_vreg = vreg.drop_duplicates("_supplier_key", keep="first").set_index("_supplier_key")

    scraped_rows: dict[str, dict] = {}
    if "supplier" in scraped.columns:
        for row in scraped.to_dict("records"):
            supplier_value = row.get("supplier")
            if supplier_value is None or pd.isna(supplier_value):
                continue
            key = _supplier_key(supplier_value)
            existing = scraped_rows.get(key)
            if existing is None or sum(
                value is not None and not pd.isna(value)
                for column in (*VREG_COLUMNS.values(), "standing_charge")
                for value in (row.get(column),)
            ) > sum(
                value is not None and not pd.isna(value)
                for column in (*VREG_COLUMNS.values(), "standing_charge")
                for value in (existing.get(column),)
            ):
                scraped_rows[key] = row

    suppliers: dict[str, str] = {}
    for key, row in best_vreg.iterrows():
        suppliers[str(key)] = str(row["Leverancier"]).strip()
    for row in scraped.to_dict("records"):
        supplier_value = row.get("supplier")
        if supplier_value is None or pd.isna(supplier_value):
            continue
        supplier = str(supplier_value).strip()
        if supplier:
            suppliers.setdefault(_supplier_key(supplier), supplier)

    merged_rows = []
    for key, supplier in suppliers.items():
        row = dict(scraped_rows.get(key, {}))
        row["supplier"] = supplier
        if row.get("standing_charge") is not None and not pd.isna(row.get("standing_charge")):
            row["standing_charge_source"] = "Leverancierspagina (scrape)"
        else:
            row["standing_charge_source"] = "Niet beschikbaar"
        vreg_row = best_vreg.loc[key] if key in best_vreg.index else None

        for vreg_column, target_column in VREG_COLUMNS.items():
            vreg_value = _vreg_price(
                vreg_row.get(vreg_column) if vreg_row is not None else None,
                "€/kWh",
            )
            scraped_value = row.get(target_column)
            if vreg_value is not None:
                row[target_column] = vreg_value
                row[f"{target_column}_source"] = "VREG"
            elif scraped_value is not None and not pd.isna(scraped_value):
                row[f"{target_column}_source"] = "Leverancierspagina (scrape)"
            else:
                row[target_column] = None
                row[f"{target_column}_source"] = "Niet beschikbaar"

        for source_column, target_column in (
            ("Stroom vaste €/jr", "electricity_standing_charge"),
            ("Gas vaste €/jr", "gas_standing_charge"),
        ):
            vreg_value = _vreg_price(
                vreg_row.get(source_column) if vreg_row is not None else None,
                "€/jaar",
            )
            if vreg_value is not None:
                row[target_column] = vreg_value
                row[f"{target_column}_source"] = "VREG"
            else:
                row[target_column] = None
                row[f"{target_column}_source"] = "Niet beschikbaar"

        if vreg_row is not None:
            row["VREG product stroom"] = vreg_row.get("Product stroom")
            row["VREG product gas"] = vreg_row.get("Product gas")
            row["green_power_rate"] = _vreg_price(
                vreg_row.get("Groenestroom €/kWh"),
                "€/kWh",
            )
            row["wkk_rate"] = _vreg_price(vreg_row.get("WKK €/kWh"), "€/kWh")
            row["vreg_source_url"] = (
                "https://www.vlaamsenutsregulator.be/cijfers/"
                "v-test-data-en-energieprijscurves"
            )
        else:
            row["VREG product stroom"] = None
            row["VREG product gas"] = None
            row["green_power_rate"] = None
            row["wkk_rate"] = None
            row["vreg_source_url"] = None
        sources = {
            row.get(f"{column}_source")
            for column in (
                "electricity_single_rate",
                "electricity_day_rate",
                "electricity_night_rate",
                "gas_rate",
                "electricity_standing_charge",
                "gas_standing_charge",
                "standing_charge",
            )
        }
        row["price_sources"] = ", ".join(
            source for source in ("VREG", "Leverancierspagina (scrape)") if source in sources
        ) or "Niet beschikbaar"
        row["vreg_data_period"] = (
            vreg_row.get("VREG data periode") if vreg_row is not None else None
        )
        merged_rows.append(row)
    return pd.DataFrame(merged_rows)


def fmt(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float):
        return f"{value:.5f}".rstrip("0").rstrip(".")
    return str(value)


class DataTable(ttk.Frame):
    """Sortable, filterable table with clickable URLs (double click)."""

    def __init__(self, master):
        super().__init__(master)
        self.df = pd.DataFrame()
        self.view_df = pd.DataFrame()
        self.sort_col = None
        self.sort_desc = False

        top = ttk.Frame(self)
        top.pack(fill="x", pady=(4, 2))
        ttk.Label(top, text="Filter:").pack(side="left", padx=(4, 4))
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(top, textvariable=self.filter_var, width=40).pack(side="left")
        self.info = ttk.Label(top, text="")
        self.info.pack(side="right", padx=6)

        wrap = ttk.Frame(self)
        wrap.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(wrap, show="headings")
        ys = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.pack(side="right", fill="y")
        xs.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.tag_configure("odd", background="#F3F6FB")

    def set_data(self, df: pd.DataFrame):
        df = df.loc[:, ~df.columns.duplicated()]
        df.columns = [str(c) for c in df.columns]
        self.df = df.reset_index(drop=True)
        self.sort_col = None
        self.tree["columns"] = list(self.df.columns)
        for col in self.df.columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by(c))
            self.tree.column(col, width=self._width(col), anchor="w")
        self.refresh()

    def _width(self, col) -> int:
        if self.df.empty:
            return 120
        sample = self.df[col].astype(str).head(200)
        if isinstance(sample, pd.DataFrame):
            sample = sample.iloc[:, 0]
        longest = max([len(str(col))] + [len(fmt(s)) for s in sample])
        return min(max(90, longest * 7), 420)

    def sort_by(self, col):
        self.sort_desc = (not self.sort_desc) if self.sort_col == col else False
        self.sort_col = col
        self.refresh()

    def refresh(self):
        df = self.df
        text = self.filter_var.get().strip().lower()
        if text and not df.empty:
            mask = df.astype(str).apply(lambda s: s.str.lower().str.contains(text, regex=False)).any(axis=1)
            df = df[mask]
        if self.sort_col in df.columns:
            numeric = pd.to_numeric(df[self.sort_col], errors="coerce")
            if numeric.notna().sum() >= df[self.sort_col].notna().sum() / 2 and numeric.notna().any():
                df = df.assign(_k=numeric).sort_values("_k", ascending=not self.sort_desc, na_position="last").drop(columns="_k")
            else:
                df = df.sort_values(self.sort_col, ascending=not self.sort_desc, key=lambda s: s.astype(str).str.lower())
        self.view_df = df
        self.tree.delete(*self.tree.get_children())
        for i, (_, row) in enumerate(df.iterrows()):
            self.tree.insert("", "end", values=[fmt(v) for v in row], tags=("odd",) if i % 2 else ())
        self.info.config(text=f"{len(df)} / {len(self.df)} rijen — dubbelklik op een rij met URL opent de link")

    def on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item:
            return
        values = self.tree.item(item, "values")
        cols = list(self.tree["columns"])
        idx = int(col.replace("#", "")) - 1
        candidates = [values[idx]] if 0 <= idx < len(values) else []
        candidates += list(values)
        for v in candidates:
            if isinstance(v, str) and v.startswith(("http://", "https://")):
                webbrowser.open(v)
                return
        _ = cols


class ComparisonTab(ttk.Frame):
    """Same cost model as energy_scraper.estimate_annual_cost_eur, with adjustable profile."""

    COLS = [
        ("supplier", "Leverancier"),
        ("vreg_electricity_product", "VREG-product stroom"),
        ("vreg_gas_product", "VREG-product gas"),
        ("single", "Enkel €/kWh"),
        ("day", "Dag €/kWh"),
        ("night", "Nacht €/kWh"),
        ("gas", "Gas €/kWh"),
        ("standing", "Vaste kost €/jaar"),
        ("elec_cost", "Elektriciteit €/jaar"),
        ("gas_cost", "Gas €/jaar"),
        ("total", "Totaal €/jaar"),
        ("complete", "Compleet?"),
        ("price_sources", "Prijsbron"),
        ("url", "Bron URL"),
    ]

    def __init__(self, master):
        super().__init__(master)
        self.overview = pd.DataFrame()

        box = ttk.LabelFrame(self, text="Verbruiksprofiel")
        box.pack(fill="x", padx=4, pady=4)
        self.vars = {
            "elec": tk.StringVar(value="3500"),
            "gas": tk.StringVar(value="17000"),
            "day": tk.StringVar(value="60"),
        }
        labels = [
            ("Elektriciteit (kWh/jaar):", "elec"),
            ("Gas (kWh/jaar):", "gas"),
            ("Aandeel dagtarief (%):", "day"),
        ]
        for i, (label, key) in enumerate(labels):
            ttk.Label(box, text=label).grid(row=0, column=i * 2, padx=(8, 2), pady=6)
            e = ttk.Entry(box, textvariable=self.vars[key], width=9)
            e.grid(row=0, column=i * 2 + 1, padx=(0, 8))
            self.vars[key].trace_add("write", lambda *_: self.recalc())
        ttk.Button(box, text="Standaard (3500 / 17000 / 60%)", command=self.reset).grid(row=0, column=6, padx=8)

        self.table = DataTable(self)
        self.table.pack(fill="both", expand=True)

        self.note = ttk.Label(
            self,
            text="Schatting = kWh × tarief + vaste kost. Enkel tarief heeft voorrang, anders dag/nacht-mix. "
                 "Enkel gepubliceerde lijstprijzen, excl. distributie & taksen.",
            foreground="#666666",
        )
        self.note.pack(fill="x", padx=6, pady=2)

    def reset(self):
        self.vars["elec"].set("3500")
        self.vars["gas"].set("17000")
        self.vars["day"].set("60")

    def set_data(self, overview: pd.DataFrame):
        self.overview = overview
        periods = (
            overview["vreg_data_period"].dropna().astype(str).unique().tolist()
            if "vreg_data_period" in overview.columns
            else []
        )
        if periods:
            self.note.config(
                text=(
                    f"VREG-prijzen ({', '.join(periods)}) vullen waar mogelijk aan; ontbrekende velden "
                    "komen van leverancierspagina's. VREG geldt voor Vlaanderen. "
                    "Schatting = kWh × tarief + vaste kosten; "
                    "excl. distributie & taksen. Controleer altijd de getoonde prijsbron."
                )
            )
        else:
            self.note.config(
                text=(
                    "Schatting = kWh × tarief + vaste kost. Enkel tarief heeft voorrang, anders dag/nacht-mix. "
                    "Enkel gepubliceerde lijstprijzen, excl. distributie & taksen."
                )
            )
        self.recalc()

    def _profile(self):
        try:
            elec = float(self.vars["elec"].get().replace(",", "."))
            gas = float(self.vars["gas"].get().replace(",", "."))
            day = min(max(float(self.vars["day"].get().replace(",", ".")), 0.0), 100.0) / 100.0
            return elec, gas, day
        except ValueError:
            return None

    def recalc(self):
        profile = self._profile()
        headers = [h for _, h in self.COLS]
        if profile is None or self.overview.empty:
            self.table.set_data(pd.DataFrame(columns=headers))
            return
        elec_kwh, gas_kwh, day_share = profile

        rows = []
        for _, r in self.overview.iterrows():
            single = parse_energy_rate_to_eur_per_kwh(r.get("electricity_single_rate"))
            day = parse_energy_rate_to_eur_per_kwh(r.get("electricity_day_rate"))
            night = parse_energy_rate_to_eur_per_kwh(r.get("electricity_night_rate"))
            gas = parse_energy_rate_to_eur_per_kwh(r.get("gas_rate"))
            electricity_standing = parse_standing_charge_to_eur_per_year(
                r.get("electricity_standing_charge")
            )
            gas_standing = parse_standing_charge_to_eur_per_year(r.get("gas_standing_charge"))
            legacy_standing = parse_standing_charge_to_eur_per_year(r.get("standing_charge"))
            if electricity_standing is None and gas_standing is None:
                standing = legacy_standing or 0.0
            else:
                standing = (electricity_standing or 0.0) + (gas_standing or 0.0)

            if single is not None:
                elec_rate = single
            elif day is not None and night is not None:
                elec_rate = day_share * day + (1 - day_share) * night
            else:
                elec_rate = day if day is not None else night

            elec_cost = elec_kwh * elec_rate if elec_rate is not None else None
            gas_cost = gas_kwh * gas if gas is not None else None
            total = None
            if elec_cost is not None and gas_cost is not None:
                total = elec_cost + gas_cost + (standing or 0.0)

            rows.append({
                "Leverancier": r.get("supplier"),
                "VREG-product stroom": r.get("VREG product stroom"),
                "VREG-product gas": r.get("VREG product gas"),
                "Enkel €/kWh": single,
                "Dag €/kWh": day,
                "Nacht €/kWh": night,
                "Gas €/kWh": gas,
                "Vaste kost €/jaar": standing,
                "Elektriciteit €/jaar": round(elec_cost, 2) if elec_cost is not None else None,
                "Gas €/jaar": round(gas_cost, 2) if gas_cost is not None else None,
                "Totaal €/jaar": round(total, 2) if total is not None else None,
                "Compleet?": "Ja" if total is not None else "Nee",
                "Prijsbron": r.get("price_sources", "Leverancierspagina (scrape)"),
                "Bron URL": (
                    r.get("offer_url")
                    or r.get("matched_url")
                    or r.get("source_url")
                    or r.get("vreg_source_url")
                ),
            })

        out = pd.DataFrame(rows, columns=headers)
        out = out.sort_values("Totaal €/jaar", na_position="last").reset_index(drop=True)
        self.table.set_data(out)
        self.table.sort_col = "Totaal €/jaar"


class DashboardTab(ttk.Frame):
    """Interactive annual-cost dashboard using the prices available in the CSV data."""

    MODES = ("Enkelvoudig", "Dag/Nacht")
    NO_NETWORK_ZONE = "Niet meegerekend"

    def __init__(self, master):
        super().__init__(master, padding=8)
        self.overview = pd.DataFrame()
        self.network_rates = pd.DataFrame()
        self.discounts: dict[str, tuple[float, float]] = {}
        self.network_source = None
        self.network_load_error = ""
        self.electric_row = None
        self.gas_row = None

        title = ttk.Label(
            self,
            text="ENERGIE DASHBOARD",
            anchor="center",
            font=("TkDefaultFont", 14, "bold"),
            padding=6,
        )
        title.pack(fill="x", pady=(0, 6))

        top = ttk.Frame(self)
        top.pack(fill="x")
        selections = ttk.LabelFrame(top, text="Selecties")
        selections.pack(side="left", fill="both", expand=True, padx=(0, 4))
        profile = ttk.LabelFrame(top, text="Verbruik en extra opties")
        profile.pack(side="left", fill="both", expand=True, padx=(4, 0))

        self.electric_supplier = tk.StringVar()
        self.gas_supplier = tk.StringVar()
        self.network_zone = tk.StringVar(value=self.NO_NETWORK_ZONE)
        self.tariff_mode = tk.StringVar(value="Enkelvoudig")
        self.profile_vars = {
            "day": tk.StringVar(value="2200"),
            "night": tk.StringVar(value="1300"),
            "gas": tk.StringVar(value="17000"),
            "peak": tk.StringVar(value="4.5"),
            "pv": tk.StringVar(value="3500"),
            "self_use": tk.StringVar(value="35"),
            "battery_boost": tk.StringVar(value="20"),
            "ev": tk.StringVar(value="1800"),
            "heat_pump": tk.StringVar(value="2500"),
            "vat": tk.StringVar(value="6"),
            "discount": tk.StringVar(value="0"),
        }
        self.options = {
            "solar": tk.BooleanVar(value=False),
            "battery": tk.BooleanVar(value=False),
            "ev_enabled": tk.BooleanVar(value=False),
            "heat_pump_enabled": tk.BooleanVar(value=False),
        }
        self._add_combo(selections, 0, "Leverancier stroom", self.electric_supplier)
        ttk.Label(selections, text="Tariefmodus stroom").grid(row=1, column=0, sticky="w", padx=6, pady=3)
        mode = ttk.Combobox(
            selections,
            textvariable=self.tariff_mode,
            values=self.MODES,
            state="readonly",
            width=22,
        )
        mode.grid(row=1, column=1, sticky="ew", padx=6, pady=3)
        mode.bind("<<ComboboxSelected>>", lambda *_: self.recalc())
        self._add_combo(selections, 2, "Leverancier gas", self.gas_supplier)
        ttk.Label(selections, text="Netbeheerder (Fluvius-zone)").grid(
            row=3, column=0, sticky="w", padx=6, pady=3
        )
        self.network_combo = ttk.Combobox(
            selections,
            textvariable=self.network_zone,
            state="readonly",
            width=25,
        )
        self.network_combo.grid(row=3, column=1, sticky="ew", padx=6, pady=3)
        self.network_combo.bind("<<ComboboxSelected>>", lambda *_: self.recalc())
        selections.columnconfigure(1, weight=1)

        for row, (label, key) in enumerate(
            (
                ("Stroom dag (kWh/jaar)", "day"),
                ("Stroom nacht (kWh/jaar)", "night"),
                ("Gas (kWh/jaar)", "gas"),
                ("Gemiddelde maandpiek (kW)", "peak"),
                ("BTW (%)", "vat"),
                ("Korting (€/jaar)", "discount"),
            )
        ):
            self._add_entry(profile, row, 0, label, key)
        for row, (label, key) in enumerate(
            (
                ("Zonnepanelen", "solar"),
                ("PV productie (kWh/jaar)", "pv"),
                ("Zelfverbruik (%)", "self_use"),
                ("Batterij", "battery"),
                ("Battery boost zelfverbruik (%)", "battery_boost"),
                ("EV laden thuis", "ev_enabled"),
                ("Extra stroom EV (kWh/jaar)", "ev"),
                ("Warmtepomp", "heat_pump_enabled"),
                ("Extra stroom warmtepomp (kWh/jaar)", "heat_pump"),
            )
        ):
            self._add_option(profile, row, label, key)
        profile.columnconfigure(1, weight=1)

        rate_box = ttk.LabelFrame(self, text="Geselecteerde tarieven")
        rate_box.pack(fill="x", pady=(8, 4))
        self.rate_vars = {
            "electricity_single": tk.StringVar(value="—"),
            "electricity_day": tk.StringVar(value="—"),
            "electricity_night": tk.StringVar(value="—"),
            "electricity_fixed": tk.StringVar(value="—"),
            "gas": tk.StringVar(value="—"),
            "gas_fixed": tk.StringVar(value="—"),
        }
        rate_labels = (
            ("Stroom enkel (€/kWh)", "electricity_single"),
            ("Stroom dag (€/kWh)", "electricity_day"),
            ("Stroom nacht (€/kWh)", "electricity_night"),
            ("Vaste stroomkost (€/jaar)", "electricity_fixed"),
            ("Gas (€/kWh)", "gas"),
            ("Vaste gaskost (€/jaar)", "gas_fixed"),
        )
        for index, (label, key) in enumerate(rate_labels):
            column = (index // 3) * 2
            row = index % 3
            ttk.Label(rate_box, text=label).grid(row=row, column=column, sticky="w", padx=6, pady=2)
            ttk.Label(rate_box, textvariable=self.rate_vars[key], width=18).grid(
                row=row, column=column + 1, sticky="w", padx=6, pady=2
            )

        self.network_source_var = tk.StringVar(value="Nettarieven: niet meegerekend.")
        source_line = ttk.Frame(self)
        source_line.pack(fill="x", pady=(2, 2))
        ttk.Label(
            source_line,
            textvariable=self.network_source_var,
            foreground="#555555",
            wraplength=1100,
        ).pack(side="left", fill="x", expand=True)
        ttk.Button(source_line, text="Open Fluvius-bron", command=self.open_network_source).pack(
            side="right", padx=(6, 0)
        )

        result_box = ttk.LabelFrame(self, text="Berekening (per jaar)")
        result_box.pack(fill="both", expand=True, pady=(4, 4))
        self.result_vars = {
            "solar_used": tk.StringVar(value="—"),
            "net_day": tk.StringVar(value="—"),
            "net_night": tk.StringVar(value="—"),
            "electricity_cost": tk.StringVar(value="—"),
            "green_power": tk.StringVar(value="—"),
            "wkk": tk.StringVar(value="—"),
            "electricity_fixed": tk.StringVar(value="—"),
            "gas_cost": tk.StringVar(value="—"),
            "gas_fixed": tk.StringVar(value="—"),
            "discount": tk.StringVar(value="—"),
            "subtotal": tk.StringVar(value="—"),
            "vat": tk.StringVar(value="—"),
            "electricity_distribution": tk.StringVar(value="—"),
            "electricity_odv": tk.StringVar(value="—"),
            "electricity_surcharges": tk.StringVar(value="—"),
            "network_capacity": tk.StringVar(value="—"),
            "electricity_data": tk.StringVar(value="—"),
            "gas_distribution": tk.StringVar(value="—"),
            "gas_odv": tk.StringVar(value="—"),
            "gas_other_levies": tk.StringVar(value="—"),
            "gas_fixed_term": tk.StringVar(value="—"),
            "gas_data": tk.StringVar(value="—"),
        }
        self.result_labels = (
            ("Solar zelfverbruik (kWh)", "solar_used"),
            ("Netto stroom dag (kWh)", "net_day"),
            ("Netto stroom nacht (kWh)", "net_night"),
            ("Leverancierstarief excl. uitgesplitste posten", "electricity_cost"),
            ("Groenestroomkosten (deel tarief)", "green_power"),
            ("WKK-kosten (deel tarief)", "wkk"),
            ("Vaste leverancierskosten", "electricity_fixed"),
            ("Gas energiekost", "gas_cost"),
            ("Vaste gaskosten", "gas_fixed"),
            ("Korting", "discount"),
            ("Subtotaal excl. btw", "subtotal"),
            ("Btw bedrag", "vat"),
            ("Distributie elektriciteit", "electricity_distribution"),
            ("ODV-heffing elektriciteit", "electricity_odv"),
            ("Toeslagen elektriciteit", "electricity_surcharges"),
            ("Capaciteit stroomnet (incl. minimum)", "network_capacity"),
            ("Databeheer elektriciteit", "electricity_data"),
            ("Distributie gas", "gas_distribution"),
            ("ODV-heffing gas", "gas_odv"),
            ("Overige heffingen gas", "gas_other_levies"),
            ("Vaste netvergoeding gas", "gas_fixed_term"),
            ("Databeheer gas", "gas_data"),
        )
        result_rows_per_column = (len(self.result_labels) + 2) // 3
        for index, (label, key) in enumerate(self.result_labels):
            column = (index // result_rows_per_column) * 2
            row = index % result_rows_per_column
            ttk.Label(result_box, text=label).grid(row=row, column=column, sticky="w", padx=6, pady=2)
            ttk.Label(result_box, textvariable=self.result_vars[key], width=18).grid(
                row=row, column=column + 1, sticky="w", padx=6, pady=2
            )

        summary = ttk.Frame(self)
        summary.pack(fill="x", pady=(2, 0))
        self.total_var = tk.StringVar(value="Totaal: —")
        ttk.Label(summary, textvariable=self.total_var, font=("TkDefaultFont", 12, "bold")).pack(side="left")
        self.monthly_var = tk.StringVar(value="Per maand: —")
        ttk.Label(summary, textvariable=self.monthly_var, font=("TkDefaultFont", 12, "bold")).pack(side="right")

        self.status_var = tk.StringVar()
        ttk.Label(
            self,
            textvariable=self.status_var,
            foreground="#9C5700",
            wraplength=1100,
        ).pack(fill="x", pady=(3, 0))

        for variable in self.profile_vars.values():
            variable.trace_add("write", lambda *_: self.recalc())
        for variable in self.options.values():
            variable.trace_add("write", lambda *_: self.recalc())
        self.electric_supplier.trace_add("write", lambda *_: self._apply_supplier_discount())
        self.gas_supplier.trace_add("write", lambda *_: self._apply_supplier_discount())
        self.network_zone.trace_add("write", lambda *_: self.recalc())
        self.recalc()

    def set_discounts(self, discounts: dict[str, tuple[float, float]]) -> None:
        self.discounts = discounts
        self._apply_supplier_discount()

    def _apply_supplier_discount(self) -> None:
        """Prefill the discount field from the discount file for the chosen suppliers."""
        electric = self.discounts.get(_supplier_key(self.electric_supplier.get()), (0.0, 0.0))[0]
        gas = self.discounts.get(_supplier_key(self.gas_supplier.get()), (0.0, 0.0))[1]
        total = electric + gas
        text = f"{total:.2f}".rstrip("0").rstrip(".").replace(".", ",") if total else "0"
        if self.profile_vars["discount"].get() != text:
            self.profile_vars["discount"].set(text)
        else:
            self.recalc()

    def _add_combo(self, parent, row: int, label: str, variable: tk.StringVar) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=3)
        combo = ttk.Combobox(parent, textvariable=variable, state="readonly", width=25)
        combo.grid(row=row, column=1, sticky="ew", padx=6, pady=3)
        combo.bind("<<ComboboxSelected>>", lambda *_: self.recalc())
        if label == "Leverancier stroom":
            self.electric_combo = combo
        else:
            self.gas_combo = combo

    def _add_entry(self, parent, row: int, column: int, label: str, key: str) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=column, sticky="w", padx=6, pady=2)
        entry = ttk.Entry(parent, textvariable=self.profile_vars[key], width=12)
        entry.grid(row=row, column=column + 1, sticky="w", padx=6, pady=2)
        entry.bind("<KP_Decimal>", self._insert_decimal_comma)

    @staticmethod
    def _insert_decimal_comma(event: tk.Event) -> str:
        entry = event.widget
        if not isinstance(entry, ttk.Entry):
            return "break"
        if entry.selection_present():
            entry.delete("sel.first", "sel.last")
        entry.insert("insert", ",")
        return "break"

    def _add_option(self, parent, row: int, label: str, key: str) -> None:
        if key in self.options:
            ttk.Label(parent, text=label).grid(row=row, column=2, sticky="w", padx=6, pady=2)
            ttk.Checkbutton(parent, variable=self.options[key]).grid(row=row, column=3, sticky="w", padx=6)
        else:
            self._add_entry(parent, row, 2, label, key)

    def set_data(self, overview: pd.DataFrame) -> None:
        self.overview = overview.loc[:, ~overview.columns.duplicated()].copy()
        suppliers = []
        if "supplier" in self.overview.columns:
            suppliers = sorted(
                self.overview["supplier"].dropna().astype(str).unique().tolist(),
                key=str.casefold,
            )
        self.electric_combo["values"] = suppliers
        self.gas_combo["values"] = suppliers
        current_electric = self.electric_supplier.get()
        current_gas = self.gas_supplier.get()
        if current_electric not in suppliers:
            preferred = self._first_supplier_with_rate(
                ("electricity_single_rate", "electricity_day_rate", "electricity_night_rate")
            )
            self.electric_supplier.set(preferred or (suppliers[0] if suppliers else ""))
        if current_gas not in suppliers:
            self.gas_supplier.set(self._first_supplier_with_rate(("gas_rate",)) or (suppliers[0] if suppliers else ""))
        self.recalc()

    def set_network_data(
        self,
        rates: pd.DataFrame,
        source: Path | None = None,
        load_error: str = "",
    ) -> None:
        self.network_rates = rates.loc[:, ~rates.columns.duplicated()].copy()
        self.network_source = source
        self.network_load_error = load_error
        zones = []
        if "zone" in self.network_rates.columns:
            zones = sorted(
                self.network_rates["zone"].dropna().astype(str).unique().tolist(),
                key=str.casefold,
            )
        self.network_combo["values"] = [self.NO_NETWORK_ZONE, *zones]
        if self.network_zone.get() not in zones and self.network_zone.get() != self.NO_NETWORK_ZONE:
            self.network_zone.set("Fluvius Imewo" if "Fluvius Imewo" in zones else self.NO_NETWORK_ZONE)
        elif self.network_zone.get() == self.NO_NETWORK_ZONE and "Fluvius Imewo" in zones:
            self.network_zone.set("Fluvius Imewo")
        self.recalc()

    def _first_supplier_with_rate(self, columns: tuple[str, ...]) -> str | None:
        for _, row in self.overview.iterrows():
            if any(parse_energy_rate_to_eur_per_kwh(row.get(column)) is not None for column in columns):
                return str(row.get("supplier"))
        return None

    def _row_for(self, supplier: str):
        if not supplier or "supplier" not in self.overview.columns:
            return None
        rows = self.overview[self.overview["supplier"].astype(str) == supplier]
        return rows.iloc[0] if not rows.empty else None

    def _network_row(self):
        if self.network_zone.get() == self.NO_NETWORK_ZONE or "zone" not in self.network_rates.columns:
            return None
        rows = self.network_rates[self.network_rates["zone"].astype(str) == self.network_zone.get()]
        return rows.iloc[0] if not rows.empty else None

    def open_network_source(self) -> None:
        row = self._network_row()
        if row is None:
            return
        for field in ("elec_source", "gas_source"):
            url = row.get(field)
            if isinstance(url, str) and url.startswith(("https://", "http://")):
                webbrowser.open(url)
                return

    @staticmethod
    def _rate_value(row, key: str) -> float | None:
        if row is None:
            return None
        value = row.get(key)
        if value is None or pd.isna(value):
            return None
        try:
            number = float(value)
        except (TypeError, ValueError):
            return None
        return number if math.isfinite(number) else None

    @staticmethod
    def _gas_tier(consumption_kwh: float) -> str:
        if consumption_kwh <= 5000:
            return "t1"
        if consumption_kwh <= 150000:
            return "t2"
        if consumption_kwh <= 1000000:
            return "t3"
        return "t4"

    @staticmethod
    def _euro(value: float | None) -> str:
        if value is None:
            return "Niet beschikbaar"
        return f"€ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    @staticmethod
    def _rate(value: float | None) -> str:
        if value is None:
            return "Niet beschikbaar"
        return f"{value:.5f}".replace(".", ",")

    def _number(self, key: str, label: str, *, maximum: float | None = None) -> float:
        try:
            value = float(self.profile_vars[key].get().replace(",", "."))
        except ValueError as exc:
            raise ValueError(f"{label}: voer een geldig getal in.") from exc
        if not math.isfinite(value) or value < 0 or (maximum is not None and value > maximum):
            if maximum is None:
                raise ValueError(f"{label} moet een positief getal zijn.")
            raise ValueError(f"{label} moet tussen 0 en {maximum} liggen.")
        return value

    def recalc(self) -> None:
        self.electric_row = self._row_for(self.electric_supplier.get())
        self.gas_row = self._row_for(self.gas_supplier.get())
        network_row = self._network_row()

        electric_rates = {
            "single": parse_energy_rate_to_eur_per_kwh(
                self.electric_row.get("electricity_single_rate") if self.electric_row is not None else None
            ),
            "day": parse_energy_rate_to_eur_per_kwh(
                self.electric_row.get("electricity_day_rate") if self.electric_row is not None else None
            ),
            "night": parse_energy_rate_to_eur_per_kwh(
                self.electric_row.get("electricity_night_rate") if self.electric_row is not None else None
            ),
        }
        gas_rate = parse_energy_rate_to_eur_per_kwh(
            self.gas_row.get("gas_rate") if self.gas_row is not None else None
        )
        electric_fixed = parse_standing_charge_to_eur_per_year(
            self.electric_row.get("electricity_standing_charge")
            if self.electric_row is not None
            else None
        )
        gas_fixed = parse_standing_charge_to_eur_per_year(
            self.gas_row.get("gas_standing_charge") if self.gas_row is not None else None
        )
        same_supplier = self.electric_supplier.get() == self.gas_supplier.get()
        has_vreg_fixed_charges = electric_fixed is not None or gas_fixed is not None
        if not has_vreg_fixed_charges:
            electric_fixed = parse_standing_charge_to_eur_per_year(
                self.electric_row.get("standing_charge") if self.electric_row is not None else None
            )
            gas_fixed = parse_standing_charge_to_eur_per_year(
                self.gas_row.get("standing_charge") if self.gas_row is not None else None
            )
        elif same_supplier:
            if electric_fixed is None:
                electric_fixed = 0.0
            if gas_fixed is None:
                gas_fixed = 0.0

        for key in ("electricity_single", "electricity_day", "electricity_night"):
            self.rate_vars[key].set(self._rate(electric_rates[key.removeprefix("electricity_")]))
        self.rate_vars["gas"].set(self._rate(gas_rate))
        self.rate_vars["electricity_fixed"].set(self._euro(electric_fixed))
        self.rate_vars["gas_fixed"].set(self._euro(gas_fixed))

        try:
            day_kwh = self._number("day", "Stroom dag")
            night_kwh = self._number("night", "Stroom nacht")
            gas_kwh = self._number("gas", "Gas")
            peak_kw = self._number("peak", "Gemiddelde maandpiek")
            vat_rate = self._number("vat", "Btw", maximum=100.0)
            discount = self._number("discount", "Korting")
            pv_kwh = self._number("pv", "PV productie")
            self_use = self._number("self_use", "Zelfverbruik", maximum=100.0)
            battery_boost = self._number("battery_boost", "Battery boost", maximum=100.0)
            ev_kwh = self._number("ev", "Extra stroom EV")
            heat_pump_kwh = self._number("heat_pump", "Extra stroom warmtepomp")
        except ValueError as exc:
            self.status_var.set(str(exc) if str(exc) else "Vul geldige niet-negatieve getallen in.")
            self._clear_results()
            return

        extra_day = (
            (ev_kwh if self.options["ev_enabled"].get() else 0.0)
            + (heat_pump_kwh if self.options["heat_pump_enabled"].get() else 0.0)
        )
        gross_day = day_kwh + extra_day
        solar_used = 0.0
        if self.options["solar"].get():
            boost = battery_boost if self.options["battery"].get() else 0.0
            solar_used = min(gross_day, pv_kwh * (self_use / 100.0) * (1.0 + boost / 100.0))
        net_day = max(0.0, gross_day - solar_used)
        net_night = night_kwh
        net_total = net_day + net_night

        mode = self.tariff_mode.get()
        if mode == "Enkelvoudig" and electric_rates["single"] is not None:
            electricity_cost = net_total * electric_rates["single"]
        elif mode == "Dag/Nacht" and electric_rates["day"] is not None and electric_rates["night"] is not None:
            electricity_cost = net_day * electric_rates["day"] + net_night * electric_rates["night"]
        elif electric_rates["single"] is not None:
            electricity_cost = net_total * electric_rates["single"]
        elif electric_rates["day"] is not None and electric_rates["night"] is not None:
            electricity_cost = net_day * electric_rates["day"] + net_night * electric_rates["night"]
        else:
            fallback_rate = electric_rates["day"] or electric_rates["night"]
            electricity_cost = net_total * fallback_rate if fallback_rate is not None else None
        if net_total == 0:
            electricity_cost = 0.0
        gas_cost = gas_kwh * gas_rate if gas_rate is not None else (0.0 if gas_kwh == 0 else None)

        green_power_rate = parse_energy_rate_to_eur_per_kwh(
            self.electric_row.get("green_power_rate") if self.electric_row is not None else None
        )
        wkk_rate = parse_energy_rate_to_eur_per_kwh(
            self.electric_row.get("wkk_rate") if self.electric_row is not None else None
        )
        green_power_cost = net_total * green_power_rate if green_power_rate is not None else None
        wkk_cost = net_total * wkk_rate if wkk_rate is not None else None
        electricity_cost_ex_green_wkk = (
            electricity_cost - (green_power_cost or 0.0) - (wkk_cost or 0.0)
            if electricity_cost is not None
            else None
        )

        network_costs: dict[str, float | None] = {
            "electricity_distribution": 0.0,
            "electricity_odv": 0.0,
            "electricity_surcharges": 0.0,
            "capacity": 0.0,
            "electricity_data": 0.0,
            "gas_distribution": 0.0,
            "gas_odv": 0.0,
            "gas_other_levies": 0.0,
            "gas_fixed_term": 0.0,
            "gas_data": 0.0,
        }
        tariff_year = None
        if network_row is not None:
            tier = self._gas_tier(gas_kwh)
            distribution_day = self._rate_value(network_row, "elec_dist_day_eur_kwh")
            distribution_night = self._rate_value(network_row, "elec_dist_night_eur_kwh")
            odv_day = self._rate_value(network_row, "elec_heffingen_odv_eur_kwh")
            odv_night = self._rate_value(network_row, "elec_heffingen_odv_night_eur_kwh")
            surcharges = self._rate_value(network_row, "elec_heffingen_toeslagen_eur_kwh")
            capacity_rate = self._rate_value(network_row, "elec_capacity_eur_kw_year")
            minimum_network = self._rate_value(network_row, "elec_min_network_eur_year")
            electricity_data = self._rate_value(network_row, "elec_databeheer_eur_year")
            gas_distribution = self._rate_value(network_row, f"gas_dist_eur_kwh_{tier}")
            gas_odv = self._rate_value(network_row, f"gas_odv_eur_kwh_{tier}")
            gas_other_levies = self._rate_value(network_row, f"gas_other_levies_eur_kwh_{tier}")
            gas_fixed_term = self._rate_value(network_row, f"gas_fixed_term_eur_year_{tier}")
            gas_data = self._rate_value(network_row, "gas_databeheer_eur_year")
            tariff_year = str(network_row.get("imported_at", "") or "")

            if (
                distribution_day is not None
                and distribution_night is not None
                and odv_day is not None
                and odv_night is not None
                and surcharges is not None
                and capacity_rate is not None
                and minimum_network is not None
                and electricity_data is not None
            ):
                network_costs["electricity_distribution"] = (
                    gross_day * distribution_day + night_kwh * distribution_night
                )
                network_costs["electricity_odv"] = (
                    gross_day * odv_day + night_kwh * odv_night
                )
                network_costs["electricity_surcharges"] = (gross_day + night_kwh) * surcharges
                variable_electricity_network = sum(
                    network_costs[key] or 0.0
                    for key in (
                        "electricity_distribution",
                        "electricity_odv",
                        "electricity_surcharges",
                    )
                )
                network_costs["capacity"] = max(
                    peak_kw * capacity_rate,
                    minimum_network - variable_electricity_network,
                    0.0,
                )
                network_costs["electricity_data"] = electricity_data
            else:
                network_costs["electricity_distribution"] = None
                network_costs["electricity_odv"] = None
                network_costs["electricity_surcharges"] = None
                network_costs["capacity"] = None
                network_costs["electricity_data"] = None

            if (
                gas_distribution is not None
                and gas_odv is not None
                and gas_other_levies is not None
                and gas_fixed_term is not None
                and gas_data is not None
            ):
                network_costs["gas_distribution"] = gas_kwh * gas_distribution
                network_costs["gas_odv"] = gas_kwh * gas_odv
                network_costs["gas_other_levies"] = gas_kwh * gas_other_levies
                network_costs["gas_fixed_term"] = gas_fixed_term
                network_costs["gas_data"] = gas_data
            else:
                network_costs["gas_distribution"] = None
                network_costs["gas_odv"] = None
                network_costs["gas_other_levies"] = None
                network_costs["gas_fixed_term"] = None
                network_costs["gas_data"] = None

        if same_supplier and not has_vreg_fixed_charges:
            fixed_cost = electric_fixed if electric_fixed is not None else gas_fixed
            fixed_electric = fixed_cost or 0.0
            fixed_gas = 0.0
        else:
            fixed_electric = electric_fixed or 0.0
            fixed_gas = gas_fixed or 0.0

        subtotal = None
        vat_amount = None
        total = None
        missing = []
        if electricity_cost is None:
            missing.append("stroomtarief")
        if gas_cost is None:
            missing.append("gastarief")
        if network_row is not None and any(value is None for value in network_costs.values()):
            missing.append("Fluvius-tarieven")
        if (
            electricity_cost_ex_green_wkk is not None
            and gas_cost is not None
            and not missing
        ):
            subtotal = max(
                0.0,
                electricity_cost_ex_green_wkk
                + (green_power_cost or 0.0)
                + (wkk_cost or 0.0)
                + fixed_electric + gas_cost + fixed_gas
                + sum(value or 0.0 for value in network_costs.values())
                - discount,
            )
            vat_amount = subtotal * (vat_rate / 100.0)
            total = subtotal + vat_amount

        values = {
            "solar_used": f"{solar_used:,.0f} kWh".replace(",", "."),
            "net_day": f"{net_day:,.0f} kWh".replace(",", "."),
            "net_night": f"{net_night:,.0f} kWh".replace(",", "."),
            "electricity_cost": self._euro(electricity_cost_ex_green_wkk),
            "green_power": self._euro(green_power_cost),
            "wkk": self._euro(wkk_cost),
            "electricity_fixed": self._euro(fixed_electric),
            "gas_cost": self._euro(gas_cost),
            "gas_fixed": self._euro(fixed_gas),
            "discount": self._euro(discount),
            "subtotal": self._euro(subtotal),
            "vat": self._euro(vat_amount),
            "electricity_distribution": self._euro(network_costs["electricity_distribution"]),
            "electricity_odv": self._euro(network_costs["electricity_odv"]),
            "electricity_surcharges": self._euro(network_costs["electricity_surcharges"]),
            "network_capacity": self._euro(network_costs["capacity"]),
            "electricity_data": self._euro(network_costs["electricity_data"]),
            "gas_distribution": self._euro(network_costs["gas_distribution"]),
            "gas_odv": self._euro(network_costs["gas_odv"]),
            "gas_other_levies": self._euro(network_costs["gas_other_levies"]),
            "gas_fixed_term": self._euro(network_costs["gas_fixed_term"]),
            "gas_data": self._euro(network_costs["gas_data"]),
        }
        for key, value in values.items():
            self.result_vars[key].set(value)
        self.total_var.set(f"Totaal incl. btw: {self._euro(total)}")
        monthly = total / 12 if total is not None else None
        self.monthly_var.set(f"Per maand: {self._euro(monthly)}")
        self._publish_breakdown(
            electricity_cost_ex_green_wkk, green_power_cost, wkk_cost, fixed_electric,
            gas_cost, fixed_gas, discount, network_costs, network_row is not None, vat_rate,
        )

        if network_row is not None:
            tariff_stamp = f" (import: {tariff_year})" if tariff_year and tariff_year != "nan" else ""
            self.network_source_var.set(
                f"Fluvius-nettarieven en gepubliceerde heffingen voor {self.network_zone.get()}{tariff_stamp}; "
                "excl. btw; laagspanning met piekmeting. Stroomnetkosten rekenen met bruto-afname."
            )
        else:
            self.network_source_var.set(
                "Geen Fluvius-zone gekozen of geldige tariefdata ontbreekt; nettarieven en heffingen zijn niet meegerekend."
            )
            if self.network_load_error:
                self.network_source_var.set(f"{self.network_source_var.get()} {self.network_load_error}")

        caveats = [
            "Raming op basis van gevonden leveranciersprijzen en, indien gekozen, Fluvius-nettarieven."
            " Federale accijnzen/overige overheidsheffingen zijn niet beschikbaar in de brondata."
            " De btw-basis van leveranciersprijzen hangt af van de bron en moet worden gecontroleerd."
        ]
        selected_sources = []
        for row in (self.electric_row, self.gas_row):
            if row is not None:
                source = row.get("price_sources")
                if isinstance(source, str) and source and source not in selected_sources:
                    selected_sources.append(source)
        if selected_sources:
            caveats.append(f"Prijsbron: {'; '.join(selected_sources)}.")
        selected_periods = {
            str(row.get("vreg_data_period"))
            for row in (self.electric_row, self.gas_row)
            if row is not None and pd.notna(row.get("vreg_data_period"))
        }
        if selected_periods:
            caveats.append(
                f"VREG-dataset: {', '.join(sorted(selected_periods))} (alleen Vlaanderen)."
            )
        selected_products = []
        if self.electric_row is not None and pd.notna(self.electric_row.get("VREG product stroom")):
            selected_products.append(f"stroom: {self.electric_row.get('VREG product stroom')}")
        if self.gas_row is not None and pd.notna(self.gas_row.get("VREG product gas")):
            selected_products.append(f"gas: {self.gas_row.get('VREG product gas')}")
        if selected_products:
            caveats.append(
                "VREG-referentie voor deze leveranciers (niet noodzakelijk je eigen contract): "
                f"{'; '.join(selected_products)}. De koppeling kiest per leverancier het goedkoopste "
                "geschatte product op basis van 3.500 kWh stroom en 17.000 kWh gas per jaar."
            )
        if self.electric_row is not None and (
            green_power_rate is not None or wkk_rate is not None
        ):
            caveats.append(
                "Beschikbare groenestroom- en WKK-bedragen zijn uitgesplitst uit het VREG-stroomtarief "
                "en niet dubbel geteld."
            )
        elif self.electric_row is not None:
            caveats.append(
                "Groenestroom/WKK zijn niet afzonderlijk beschikbaar in deze prijsbron; "
                "de getoonde stroomprijs blijft inclusief eventuele componenten."
            )
        if missing:
            caveats.append(f"Onvolledig: ontbrekend {' en '.join(missing)} voor de gekozen leveranciers.")
        if (
            (
                self.electric_row is not None
                and parse_standing_charge_to_eur_per_year(self.electric_row.get("standing_charge")) is not None
            )
            or (
                self.gas_row is not None
                and parse_standing_charge_to_eur_per_year(self.gas_row.get("standing_charge")) is not None
                and not same_supplier
            )
        ):
            caveats.append("De gevonden vaste vergoeding is niet per stroom/gas uitgesplitst.")
        self.status_var.set(" ".join(caveats))

    def _publish_breakdown(
        self, electricity, green, wkk, fixed_electric, gas, fixed_gas, discount,
        network_costs, has_network, vat_rate,
    ) -> None:
        """Group all costs into supplier / network / taxes columns for the summary tab."""
        callback = getattr(self, "on_breakdown", None)
        if callback is None:
            return

        def total(*values):
            if any(value is None for value in values):
                return None
            return sum(values)

        n = network_costs
        rows = {
            "Stroom": (
                total(electricity, green, wkk, fixed_electric),
                total(n["electricity_distribution"], n["capacity"], n["electricity_data"])
                if has_network else None,
                total(n["electricity_odv"], n["electricity_surcharges"]) if has_network else None,
            ),
            "Gas": (
                total(gas, fixed_gas),
                total(n["gas_distribution"], n["gas_fixed_term"], n["gas_data"])
                if has_network else None,
                total(n["gas_odv"], n["gas_other_levies"]) if has_network else None,
            ),
            "Korting": (-discount if discount else 0.0, 0.0, 0.0),
        }
        subtotal = []
        for column in range(3):
            known = [rows[name][column] for name in rows if rows[name][column] is not None]
            has_values = any(rows[name][column] is not None for name in ("Stroom", "Gas"))
            subtotal.append(sum(known) if has_values else None)
        excl_vat = sum(value for value in subtotal if value is not None)
        vat = excl_vat * vat_rate / 100.0
        callback((rows, subtotal, vat, vat_rate))

    def _clear_results(self) -> None:
        callback = getattr(self, "on_breakdown", None)
        if callback is not None:
            callback(None)
        for variable in self.result_vars.values():
            variable.set("—")
        self.total_var.set("Totaal incl. btw: —")
        self.monthly_var.set("Per maand: —")

    def snapshot_settings(self) -> dict[str, str]:
        settings = {
            "version": str(DASHBOARD_SNAPSHOT_VERSION),
            "electric_supplier": self.electric_supplier.get(),
            "gas_supplier": self.gas_supplier.get(),
            "network_zone": self.network_zone.get(),
            "tariff_mode": self.tariff_mode.get(),
        }
        settings.update(
            {f"profile_{key}": variable.get() for key, variable in self.profile_vars.items()}
        )
        settings.update(
            {f"option_{key}": "1" if variable.get() else "0" for key, variable in self.options.items()}
        )
        return settings

    def restore_snapshot_settings(self, settings: dict[str, str]) -> None:
        if settings.get("version") != str(DASHBOARD_SNAPSHOT_VERSION):
            raise ValueError("Het dashboardbestand heeft een onbekende of niet-ondersteunde versie.")
        expected = {
            "electric_supplier",
            "gas_supplier",
            "network_zone",
            "tariff_mode",
        }
        expected.update(f"profile_{key}" for key in self.profile_vars)
        expected.update(f"option_{key}" for key in self.options)
        missing = expected.difference(settings)
        if missing:
            raise ValueError(f"Dashboardbestand mist instellingen: {', '.join(sorted(missing))}.")
        if settings["tariff_mode"] not in self.MODES:
            raise ValueError("Het opgeslagen dashboard bevat een onbekende tariefmodus.")
        for key in self.options:
            if settings[f"option_{key}"] not in {"0", "1"}:
                raise ValueError(f"De opgeslagen optie '{key}' is ongeldig.")

        self.electric_supplier.set(settings["electric_supplier"])
        self.gas_supplier.set(settings["gas_supplier"])
        self.network_zone.set(settings["network_zone"])
        self.tariff_mode.set(settings["tariff_mode"])
        for key, variable in self.profile_vars.items():
            variable.set(settings[f"profile_{key}"])
        for key, variable in self.options.items():
            variable.set(settings[f"option_{key}"] == "1")
        self.recalc()


class DiscountTab(ttk.Frame):
    """View and edit yearly supplier discounts stored in leverancier_kortingen.csv."""

    COLUMNS = (
        ("supplier", "Leverancier", 300),
        ("electricity", "Korting stroom (€/jaar)", 160),
        ("gas", "Korting gas (€/jaar)", 160),
        ("note", "Opmerking", 400),
    )

    def __init__(self, master, on_saved):
        super().__init__(master)
        self.on_saved = on_saved
        self.path: Path | None = None
        self.rows: dict[str, dict[str, str]] = {}

        ttk.Label(
            self,
            text=(
                "Kortingen zitten niet in de VREG-data. Selecteer een leverancier, pas de bedragen aan "
                "en klik 'Toepassen'; 'Opslaan' schrijft ze naar leverancier_kortingen.csv. "
                "Ze worden automatisch ingevuld in het dashboard."
            ),
            wraplength=1100,
        ).pack(fill="x", padx=8, pady=(8, 4))

        self.tree = ttk.Treeview(
            self, columns=[c[0] for c in self.COLUMNS], show="headings", height=14, selectmode="browse"
        )
        for key, title, width in self.COLUMNS:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=8, pady=4)
        self.tree.bind("<<TreeviewSelect>>", lambda *_: self._load_selected())

        form = ttk.Frame(self)
        form.pack(fill="x", padx=8, pady=4)
        self.supplier_var = tk.StringVar()
        self.electricity_var = tk.StringVar()
        self.gas_var = tk.StringVar()
        self.note_var = tk.StringVar()
        ttk.Label(form, textvariable=self.supplier_var, font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, 4)
        )
        for column, (label, variable, width) in enumerate(
            (
                ("Stroom €/jaar", self.electricity_var, 10),
                ("Gas €/jaar", self.gas_var, 10),
                ("Opmerking", self.note_var, 40),
            )
        ):
            ttk.Label(form, text=label).grid(row=1, column=column * 2, sticky="w", padx=(0, 4))
            ttk.Entry(form, textvariable=variable, width=width).grid(
                row=1, column=column * 2 + 1, sticky="w", padx=(0, 12)
            )
        ttk.Button(form, text="Toepassen", command=self._apply).grid(row=1, column=6, padx=4)
        ttk.Button(form, text="Opslaan", command=self.save).grid(row=1, column=7, padx=4)
        self.status_var = tk.StringVar()
        ttk.Label(self, textvariable=self.status_var, foreground="#9C5700").pack(fill="x", padx=8, pady=(0, 6))

    def load(self, path: Path, suppliers: list[str]) -> None:
        self.path = path
        self.rows = {}
        if path.is_file():
            try:
                df = pd.read_csv(path, encoding="utf-8-sig", sep=None, engine="python", dtype=str)
            except (OSError, pd.errors.ParserError, pd.errors.EmptyDataError):
                df = pd.DataFrame()
            if "supplier" in df.columns:
                for record in df.fillna("").to_dict("records"):
                    name = str(record["supplier"]).strip()
                    if name:
                        self.rows[_supplier_key(name)] = {
                            "supplier": name,
                            "electricity": str(record.get("electricity_discount_eur_year", "")).strip(),
                            "gas": str(record.get("gas_discount_eur_year", "")).strip(),
                            "note": str(record.get("note", "")).strip(),
                        }
        for name in suppliers:
            self.rows.setdefault(
                _supplier_key(name), {"supplier": name, "electricity": "", "gas": "", "note": ""}
            )
        self._render()

    def _render(self, select: str | None = None) -> None:
        self.tree.delete(*self.tree.get_children())
        for key, row in sorted(self.rows.items(), key=lambda item: item[1]["supplier"].casefold()):
            self.tree.insert(
                "", "end", iid=key,
                values=(row["supplier"], row["electricity"], row["gas"], row["note"]),
            )
        if select and select in self.rows:
            self.tree.selection_set(select)

    def _load_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        row = self.rows[selection[0]]
        self.supplier_var.set(row["supplier"])
        self.electricity_var.set(row["electricity"])
        self.gas_var.set(row["gas"])
        self.note_var.set(row["note"])

    @staticmethod
    def _amount(text: str, label: str) -> str:
        text = text.strip()
        if not text:
            return ""
        value = _finite_number(text.replace(",", "."))
        if value is None or value < 0:
            raise ValueError(f"{label}: voer een positief bedrag in.")
        return f"{value:.2f}".rstrip("0").rstrip(".")

    def _apply(self) -> bool:
        selection = self.tree.selection()
        if not selection:
            self.status_var.set("Selecteer eerst een leverancier.")
            return False
        try:
            electricity = self._amount(self.electricity_var.get(), "Stroom")
            gas = self._amount(self.gas_var.get(), "Gas")
        except ValueError as exc:
            self.status_var.set(str(exc))
            return False
        key = selection[0]
        self.rows[key].update(
            electricity=electricity, gas=gas, note=self.note_var.get().strip()
        )
        self._render(select=key)
        self.status_var.set("Toegepast; klik 'Opslaan' om te bewaren.")
        return True

    def save(self) -> None:
        if self.tree.selection() and not self._apply():
            return
        if self.path is None:
            return
        records = [
            {
                "supplier": row["supplier"],
                "electricity_discount_eur_year": row["electricity"],
                "gas_discount_eur_year": row["gas"],
                "note": row["note"],
            }
            for row in self.rows.values()
            if row["electricity"] or row["gas"] or row["note"]
        ]
        try:
            pd.DataFrame(
                records,
                columns=["supplier", "electricity_discount_eur_year", "gas_discount_eur_year", "note"],
            ).to_csv(self.path, index=False, encoding="utf-8-sig")
        except OSError as exc:
            self.status_var.set(f"Opslaan mislukt: {exc}")
            return
        self.status_var.set(f"Opgeslagen in {self.path}")
        self.on_saved(load_supplier_discounts(self.path))


class SummaryTab(ttk.Frame):
    """Cost overview grouped in supplier, network and taxes/levies columns."""

    HEADERS = ("", "Leverancier", "Netwerk (Fluvius)", "Taksen, accijnzen & heffingen", "Totaal")

    def __init__(self, master):
        super().__init__(master)
        ttk.Label(
            self,
            text=(
                "Leverancier: energie, groene stroom/WKK, vaste kosten en korting. "
                "Netwerk: distributie, capaciteit, vaste netvergoeding en databeheer. "
                "Taksen: ODV-heffing, toeslagen, overige heffingen en btw. "
                "Alle bedragen per jaar, gebaseerd op het Dashboard."
            ),
            wraplength=1100,
        ).pack(fill="x", padx=8, pady=(8, 4))
        self.grid_frame = ttk.Frame(self)
        self.grid_frame.pack(anchor="nw", padx=8, pady=8)
        self.render(None)

    @staticmethod
    def _euro(value: float | None) -> str:
        if value is None:
            return "—"
        return f"€ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def render(self, data) -> None:
        for child in self.grid_frame.winfo_children():
            child.destroy()
        for column, header in enumerate(self.HEADERS):
            ttk.Label(self.grid_frame, text=header, font=("TkDefaultFont", 10, "bold"), width=28 if column else 14,
                      anchor="e" if column else "w").grid(row=0, column=column, padx=6, pady=4, sticky="e")
        if data is None:
            return
        rows, subtotal, vat, vat_rate = data

        def add_row(row_index: int, label: str, values, bold: bool = False) -> None:
            font = ("TkDefaultFont", 10, "bold") if bold else ("TkDefaultFont", 10)
            ttk.Label(self.grid_frame, text=label, font=font).grid(row=row_index, column=0, sticky="w", padx=6, pady=3)
            known = [value for value in values if value is not None]
            cells = [*values, sum(known) if known else None]
            for column, value in enumerate(cells, start=1):
                ttk.Label(self.grid_frame, text=self._euro(value), font=font, anchor="e", width=28).grid(
                    row=row_index, column=column, sticky="e", padx=6, pady=3
                )

        row_index = 1
        for name, values in rows.items():
            add_row(row_index, name, values)
            row_index += 1
        ttk.Separator(self.grid_frame, orient="horizontal").grid(
            row=row_index, column=0, columnspan=5, sticky="ew", pady=3
        )
        add_row(row_index + 1, "Totaal excl. btw", subtotal, bold=True)
        add_row(row_index + 2, f"Btw ({vat_rate:g}%)", (None, None, vat))
        grand = [*subtotal]
        grand[2] = (grand[2] or 0.0) + vat if (grand[2] is not None or vat) else None
        add_row(row_index + 3, "Totaal incl. btw", grand, bold=True)
        monthly = [None if value is None else value / 12 for value in grand]
        add_row(row_index + 4, "Per maand", monthly)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Belgische energieprijzen — overzicht")
        self.geometry("1300x900+0+0")
        self.folder = APP_DIR
        try:
            _prepare_portable_data_dirs()
        except OSError as exc:
            messagebox.showerror(
                "Portable gegevensmappen",
                f"Gegevensmappen naast de app konden niet worden klaargemaakt:\n{exc}",
                parent=self,
            )
            raise

        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=6, pady=6)
        self.reload_btn = ttk.Button(bar, text="Herladen", command=self.load_all)
        self.reload_btn.pack(side="left", padx=4)
        self.update_rates_btn = ttk.Button(
            bar,
            text="Tarieven bijwerken",
            command=self.update_all_rates,
        )
        self.update_rates_btn.pack(side="left", padx=4)
        self.full_scan_btn = ttk.Button(
            bar,
            text="Volledige bronzoektocht",
            command=lambda: self.update_all_rates(full_discovery=True),
        )
        self.full_scan_btn.pack(side="left", padx=4)
        ttk.Button(bar, text="Dashboard opslaan (Excel)", command=self.export_dashboard).pack(
            side="left", padx=4
        )
        ttk.Button(bar, text="Dashboard laden", command=self.import_dashboard).pack(side="left", padx=4)
        ttk.Button(bar, text="Vergelijking export (CSV)", command=self.export_comparison).pack(
            side="left", padx=4
        )

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        status_bar = ttk.Frame(self, relief="sunken", padding=(6, 3))
        status_bar.pack(fill="x", side="bottom", padx=6, pady=(0, 4))
        self.status = ttk.Label(status_bar, text="", anchor="w")
        self.status.pack(side="left", fill="x", expand=True)
        ttk.Button(status_bar, text="Details", command=self.show_status_details).pack(
            side="right",
            padx=(6, 0),
        )

        self.dashboard = DashboardTab(self.nb)
        self.nb.add(self.dashboard, text="Dashboard")
        self.summary_tab = SummaryTab(self.nb)
        self.nb.add(self.summary_tab, text="Kostenoverzicht")
        self.dashboard.on_breakdown = self.summary_tab.render
        self.dashboard.recalc()
        self.comparison = ComparisonTab(self.nb)
        self.nb.add(self.comparison, text="Vergelijking")
        self.discount_tab = DiscountTab(self.nb, lambda d: self.dashboard.set_discounts(d))
        self.nb.add(self.discount_tab, text="Kortingen")
        self.tables: dict[str, DataTable] = {}
        for title, _ in CSV_TABS:
            t = DataTable(self.nb)
            self.tables[title] = t
            self.nb.add(t, text=title)

        self.load_all()
        self._refresh_vreg_on_start()

    def show_status_details(self):
        details = self.status.cget("text")
        if details:
            messagebox.showinfo("Statusdetails", details, parent=self)

    def load_all(self):
        loaded = 0
        scraped_prices = pd.DataFrame()
        for title, name in CSV_TABS:
            path = self.folder / "scraper_data" / name
            if not path.is_file():
                path = self.folder / name
            try:
                df = pd.read_csv(path, encoding="utf-8-sig")
                loaded += 1
            except (FileNotFoundError, pd.errors.EmptyDataError):
                df = pd.DataFrame()
            if name == "be_energy_price_comparison.csv":
                scraped_prices = df
                df = drop_suppliers_without_data(df)
            self.tables[title].set_data(df)

        vreg_paths = (
            self.folder / "scraper_data" / "be_vreg_vergelijking.xlsx",
            self.folder / "be_vreg_vergelijking.xlsx",
        )
        vreg_path = next((path for path in vreg_paths if path.is_file()), None)
        vreg_status = ""
        combined_prices = scraped_prices
        if vreg_path is not None:
            try:
                vreg_df = pd.read_excel(vreg_path, sheet_name="Vergelijking")
                period = ""
                if "Leverancier" in vreg_df.columns:
                    periods = (
                        vreg_df["Leverancier"]
                        .astype(str)
                        .str.extract(r"Periode:\s*([^|]+)", expand=False)
                        .dropna()
                        .str.strip()
                    )
                    if not periods.empty:
                        period = periods.iloc[0]
                vreg_df["VREG data periode"] = period or "Onbekend"
                combined_prices = combine_price_sources(scraped_prices, vreg_df)
                vreg_status = f"VREG {period or 'periode onbekend'} gekoppeld."
            except (OSError, ValueError, ImportError, pd.errors.ParserError) as exc:
                vreg_status = f"VREG-data niet gebruikt: {exc}"
        else:
            vreg_status = "Geen VREG-vergelijking gevonden; alleen leveranciersscrape gebruikt."
        combined_prices = drop_suppliers_without_data(combined_prices)
        self.dashboard.set_data(combined_prices)
        self.comparison.set_data(combined_prices)

        network_paths = (
            self.folder / "scraper_data" / "fluvius_nettarieven_latest.csv",
            self.folder / "fluvius_nettarieven_latest.csv",
        )
        network_path = next((path for path in network_paths if path.is_file()), None)
        network_df = pd.DataFrame()
        network_error = ""
        if network_path is not None:
            try:
                network_df = pd.read_csv(network_path, encoding="utf-8-sig")
                required_columns = {
                    "zone",
                    "elec_heffingen_odv_night_eur_kwh",
                    "elec_min_network_eur_year",
                    "gas_dist_eur_kwh_t2",
                    "gas_fixed_term_eur_year_t2",
                    "gas_odv_eur_kwh_t2",
                    "gas_other_levies_eur_kwh_t2",
                }
                if not required_columns.issubset(network_df.columns):
                    network_error = (
                        f"Bestand {network_path.name} heeft een verouderd formaat; "
                        "werk de nettarieven bij."
                    )
                    network_df = pd.DataFrame()
            except (FileNotFoundError, pd.errors.EmptyDataError, pd.errors.ParserError) as exc:
                network_error = f"Nettarieven konden niet worden geladen: {exc}"
                network_df = pd.DataFrame()
        else:
            network_error = "Geen Fluvius-tariefbestand gevonden; kies 'Nettarieven bijwerken'."
        self.dashboard.set_network_data(network_df, network_path, network_error)
        discount_paths = (
            self.folder / "scraper_data" / "leverancier_kortingen.csv",
            self.folder / "leverancier_kortingen.csv",
        )
        discount_path = next((path for path in discount_paths if path.is_file()), discount_paths[-1])
        self.dashboard.set_discounts(load_supplier_discounts(discount_path))
        supplier_names = (
            combined_prices["supplier"].dropna().astype(str).tolist()
            if "supplier" in combined_prices.columns
            else []
        )
        self.discount_tab.load(discount_path, supplier_names)
        self.status.config(
            text=f"{loaded}/{len(CSV_TABS)} CSV's geladen uit {self.folder} — {vreg_status}"
        )

    def update_all_rates(self, *, full_discovery: bool = False):
        scripts = (
            BASE_DIR / "fluvius_nettarieven_importer.py",
            BASE_DIR / "vreg_importer.py",
            BASE_DIR / "energy_scraper.py",
        )
        missing = [str(script) for script in scripts if not script.is_file()]
        if missing and not getattr(sys, "frozen", False):
            messagebox.showerror(
                "Tarieven bijwerken",
                "Deze benodigde scripts zijn niet gevonden:\n" + "\n".join(missing),
                parent=self,
            )
            return

        self.update_rates_btn.state(["disabled"])
        self.full_scan_btn.state(["disabled"])
        self.reload_btn.state(["disabled"])
        self.status.config(text="Tarieven bijwerken wordt voorbereid…")

        data_dir = self.folder / "scraper_data"
        steps = (
            "Fluvius-nettarieven",
            "VREG-data",
            (
                "Volledige leveranciersscan"
                if full_discovery
                else "Leveranciersprijzen (URL-cache)"
            ),
        )

        def run_frozen_step(label: str) -> None:
            writer = _LogWriter(log_file)
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                if label == "Fluvius-nettarieven":
                    update_fluvius_rates(data_dir / FLUVIUS_OUTPUT_CSV)
                elif label == "VREG-data":
                    run_vreg_importer(data_dir / VREG_OUTPUT_FILE.name)
                else:
                    run_energy_scraper(reuse_discovery=not full_discovery)

        log_file = None

        def log(text: str) -> None:
            if log_file is not None:
                log_file.write(text.rstrip("\n") + "\n")
                log_file.flush()

        def worker():
            nonlocal log_file
            failures = []
            try:
                data_dir.mkdir(parents=True, exist_ok=True)
                try:
                    log_file = open(data_dir / SCAN_LOG_NAME, "w", encoding="utf-8")
                except OSError:
                    log_file = None
                log(f"Scan gestart: {datetime.now():%Y-%m-%d %H:%M:%S}")
                log(f"Modus: {'volledige bronzoektocht' if full_discovery else 'URL-cache'}")
                for label in steps:
                    log(f"\n=== {label} ({datetime.now():%H:%M:%S}) ===")
                    self.after(
                        0,
                        lambda current_step=label: self.status.config(
                            text=f"{current_step} bijwerken…"
                        ),
                    )
                    try:
                        if getattr(sys, "frozen", False):
                            run_frozen_step(label)
                            return_code = 0
                            output = ""
                        elif label.startswith(("Volledige leveranciersscan", "Leveranciersprijzen")):
                            command = [sys.executable, "-u", str(scripts[2])]
                            if not full_discovery:
                                command.append("--reuse-discovery")
                            recent_output = []
                            with subprocess.Popen(
                                command,
                                cwd=self.folder,
                                stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT,
                                text=True,
                                encoding="utf-8",
                                errors="replace",
                                bufsize=1,
                            ) as process:
                                if process.stdout is not None:
                                    for output_line in process.stdout:
                                        log(output_line)
                                        progress = output_line.strip()
                                        if progress:
                                            recent_output.append(progress)
                                            del recent_output[:-20]
                                            self.after(
                                                0,
                                                lambda detail=progress: self.status.config(
                                                    text=f"{label}: {detail[:180]}"
                                                ),
                                            )
                                return_code = process.wait()
                                output = "\n".join(recent_output)
                        else:
                            script = scripts[0] if label == "Fluvius-nettarieven" else scripts[1]
                            result = subprocess.run(
                                [
                                    sys.executable,
                                    str(script),
                                    "--output-dir",
                                    str(data_dir.resolve()),
                                ],
                                cwd=self.folder,
                                capture_output=True,
                                text=True,
                            )
                            return_code = result.returncode
                            log(result.stdout)
                            log(result.stderr)
                            output = result.stderr.strip() or result.stdout.strip()
                    except Exception as exc:
                        failures.append(f"{label}: {exc}")
                        log(f"[ERROR] {label}: {exc}")
                        log(traceback.format_exc())
                        continue
                    log(f"--- {label}: exitcode {return_code}")
                    if return_code != 0:
                        output = output or "Geen foutdetails ontvangen."
                        failures.append(f"{label} mislukt:\n{output[-1200:]}")
            except OSError as exc:
                failures.append(f"Datamap scraper_data kon niet worden klaargemaakt:\n{exc}")
            log(f"\nScan klaar: {datetime.now():%Y-%m-%d %H:%M:%S}; {len(failures)} mislukte stap(pen)")
            if log_file is not None:
                log_file.close()
            self.after(0, lambda: self._rates_update_done(failures))

        threading.Thread(target=worker, daemon=True).start()

    def _refresh_vreg_on_start(self) -> None:
        importer = BASE_DIR / "vreg_importer.py"
        if not importer.is_file() and not getattr(sys, "frozen", False):
            messagebox.showerror("VREG-data", f"Importer niet gevonden:\n{importer}", parent=self)
            return
        self.update_rates_btn.state(["disabled"])
        self.full_scan_btn.state(["disabled"])
        self.reload_btn.state(["disabled"])
        self.status.config(text="Nieuwste VREG-data controleren bij het opstarten…")

        def worker():
            try:
                output_file = self.folder / "scraper_data" / VREG_OUTPUT_FILE.name
                if getattr(sys, "frozen", False):
                    run_vreg_importer(output_file)
                    error = ""
                else:
                    result = subprocess.run(
                        [
                            sys.executable,
                            str(importer),
                            "--output-dir",
                            str((self.folder / "scraper_data").resolve()),
                        ],
                        cwd=self.folder,
                        capture_output=True,
                        text=True,
                    )
                    error = ""
                    if result.returncode != 0:
                        output = result.stderr.strip() or result.stdout.strip() or "Geen foutdetails ontvangen."
                        error = f"VREG-data bijwerken mislukt:\n{output[-1200:]}"
            except Exception as exc:
                error = f"VREG-data bijwerken mislukt: {exc}"
            self.after(0, lambda: self._startup_vreg_done(error))

        threading.Thread(target=worker, daemon=True).start()

    def _startup_vreg_done(self, error: str) -> None:
        self.update_rates_btn.state(["!disabled"])
        self.full_scan_btn.state(["!disabled"])
        self.reload_btn.state(["!disabled"])
        self.load_all()
        if error:
            self.status.config(text="Automatische VREG-update mislukt; je kunt later 'Tarieven bijwerken' kiezen.")
            messagebox.showerror("VREG-data", error, parent=self)

    def _rates_update_done(self, failures: list[str]) -> None:
        self.update_rates_btn.state(["!disabled"])
        self.full_scan_btn.state(["!disabled"])
        self.reload_btn.state(["!disabled"])
        self.load_all()
        if failures:
            details = "\n\n".join(failures)
            self.status.config(text="Tarieven bijwerken afgerond met fouten; zie details.")
            messagebox.showerror(
                "Tarieven bijwerken",
                "Niet alle gegevens konden worden bijgewerkt:\n\n"
                + details
                + f"\n\nVolledig logboek: {self.folder / 'scraper_data' / SCAN_LOG_NAME}",
                parent=self,
            )
        else:
            self.status.config(
                text=f"Alles bijgewerkt. Logboek: {self.folder / 'scraper_data' / SCAN_LOG_NAME}"
            )

    def _dashboard_data_dir(self) -> Path:
        return self.folder / "user_data"

    def export_dashboard(self):
        data_dir = self._dashboard_data_dir()
        try:
            data_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror("Dashboard opslaan", f"Datamap kon niet worden aangemaakt:\n{exc}", parent=self)
            return
        path = filedialog.asksaveasfilename(
            title="Dashboard opslaan",
            defaultextension=".xlsx",
            filetypes=[("Excel-dashboard", "*.xlsx")],
            initialdir=data_dir,
            initialfile="dashboard_snapshot.xlsx",
        )
        if not path:
            return

        settings = self.dashboard.snapshot_settings()
        workbook = Workbook()
        sheet = workbook.active
        if sheet is None:
            raise RuntimeError("Er kan geen Excelblad voor het dashboard worden aangemaakt.")
        sheet.title = "Dashboard"
        sheet.merge_cells("A1:B1")
        sheet["A1"] = "Energie-dashboard — kostenraming"
        sheet["A1"].font = Font(bold=True, size=16, color="FFFFFF")
        sheet["A1"].fill = PatternFill("solid", fgColor="1F5C2E")
        sheet["A1"].alignment = Alignment(horizontal="center")
        sheet.append([f"Opgeslagen op: {datetime.now().strftime('%d-%m-%Y %H:%M')}"])
        section_rows = []

        def add_section(title: str) -> None:
            sheet.append([title])
            row = sheet.max_row
            sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            section_rows.append(row)

        def add_item(label: str, value: object) -> None:
            if value is None or (isinstance(value, str) and not value.strip()):
                value = "Niet beschikbaar"
            sheet.append([label, str(value)])

        add_section("Geselecteerde producten")
        add_item("Leverancier stroom", self.dashboard.electric_supplier.get())
        add_item("Product stroom", self._dashboard_product("VREG product stroom", self.dashboard.electric_row))
        add_item("Leverancier gas", self.dashboard.gas_supplier.get())
        add_item("Product gas", self._dashboard_product("VREG product gas", self.dashboard.gas_row))
        add_item("Tariefmodus stroom", self.dashboard.tariff_mode.get())
        add_item("Fluvius-zone", self.dashboard.network_zone.get())

        add_section("Verbruik en instellingen")
        profile_labels = {
            "day": "Stroom dag (kWh/jaar)",
            "night": "Stroom nacht (kWh/jaar)",
            "gas": "Gas (kWh/jaar)",
            "peak": "Gemiddelde maandpiek (kW)",
            "vat": "BTW (%)",
            "discount": "Korting (€/jaar)",
            "pv": "PV productie (kWh/jaar)",
            "self_use": "Zelfverbruik (%)",
            "battery_boost": "Battery boost zelfverbruik (%)",
            "ev": "Extra stroom EV (kWh/jaar)",
            "heat_pump": "Extra stroom warmtepomp (kWh/jaar)",
        }
        for key, label in profile_labels.items():
            add_item(label, self.dashboard.profile_vars[key].get())
        for key, label in (
            ("solar", "Zonnepanelen ingeschakeld"),
            ("battery", "Batterij ingeschakeld"),
            ("ev_enabled", "EV laden thuis ingeschakeld"),
            ("heat_pump_enabled", "Warmtepomp ingeschakeld"),
        ):
            add_item(label, "Ja" if self.dashboard.options[key].get() else "Nee")

        add_section("Tarieven")
        rate_labels = {
            "electricity_single": "Stroom enkel (€/kWh)",
            "electricity_day": "Stroom dag (€/kWh)",
            "electricity_night": "Stroom nacht (€/kWh)",
            "electricity_fixed": "Vaste stroomkost (€/jaar)",
            "gas": "Gas (€/kWh)",
            "gas_fixed": "Vaste gaskost (€/jaar)",
        }
        for key, label in rate_labels.items():
            add_item(label, self.dashboard.rate_vars[key].get())

        add_section("Jaarlijkse berekening")
        for label, key in self.dashboard.result_labels:
            add_item(label, self.dashboard.result_vars[key].get())
        add_item("Totaal", self.dashboard.total_var.get())
        add_item("Per maand", self.dashboard.monthly_var.get())

        add_section("Bronnen en toelichting")
        add_item("Nettarievenbron", self.dashboard.network_source_var.get())
        add_item("Prijs- en berekeningsnotities", self.dashboard.status_var.get())

        settings_sheet = workbook.create_sheet("Instellingen")
        settings_sheet.append(["key", "value"])
        for key, value in settings.items():
            settings_sheet.append([key, value])
        settings_sheet.sheet_state = "hidden"

        for cell in sheet[2]:
            cell.font = Font(italic=True, color="666666")
        for row in section_rows:
            for cell in sheet[row]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="2E5396")
        for row in sheet.iter_rows(min_row=1):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
        sheet.column_dimensions["A"].width = 43
        sheet.column_dimensions["B"].width = 72
        sheet.freeze_panes = "A3"
        sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True, autoPageBreaks=False)
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.page_margins.left = 0.3
        sheet.page_margins.right = 0.3
        sheet.page_margins.top = 0.5
        sheet.page_margins.bottom = 0.5
        sheet.print_title_rows = "1:2"
        sheet.print_area = f"A1:B{sheet.max_row}"
        workbook.active = 0

        try:
            workbook.save(path)
        except OSError as exc:
            messagebox.showerror("Dashboard opslaan", f"Opslaan mislukt:\n{exc}", parent=self)
            return
        self.status.config(text=f"Dashboard opgeslagen (Excel, afdruk-klaar): {path}")

    @staticmethod
    def _dashboard_product(column: str, row: pd.Series | None) -> object:
        if row is None:
            return "Niet beschikbaar"
        value = row.get(column)
        return value if value is not None and pd.notna(value) else "Niet beschikbaar"

    def import_dashboard(self):
        data_dir = self._dashboard_data_dir()
        if not data_dir.is_dir():
            data_dir = self.folder
        path = filedialog.askopenfilename(
            title="Dashboard laden",
            filetypes=[("Excel-dashboard", "*.xlsx")],
            initialdir=data_dir,
        )
        if not path:
            return

        try:
            workbook = load_workbook(path, read_only=True, data_only=True)
            if "Instellingen" not in workbook.sheetnames:
                raise ValueError("Dit bestand bevat geen opgeslagen dashboardinstellingen.")
            settings_sheet = workbook["Instellingen"]
            settings = {
                str(key): str(value)
                for key, value in settings_sheet.iter_rows(min_row=2, max_col=2, values_only=True)
                if key is not None and value is not None
            }
            workbook.close()
            self.dashboard.restore_snapshot_settings(settings)
        except (OSError, ValueError, KeyError, BadZipFile, InvalidFileException) as exc:
            messagebox.showerror("Dashboard laden", f"Dashboard kon niet worden geladen:\n{exc}", parent=self)
            return
        self.status.config(
            text=(
                f"Dashboardinstellingen geladen uit {path}. "
                "De bedragen zijn opnieuw berekend met de momenteel geladen tarieven."
            )
        )

    def export_comparison(self):
        df = self.comparison.table.view_df
        if df.empty:
            messagebox.showinfo("Export", "Geen gegevens om te exporteren.")
            return
        data_dir = self.folder / "scraper_data"
        data_dir.mkdir(parents=True, exist_ok=True)
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialdir=data_dir,
            initialfile="vergelijking.csv",
        )
        if path:
            df.to_csv(path, index=False, encoding="utf-8-sig")
            self.status.config(text=f"Opgeslagen: {path}")

if __name__ == "__main__":
    App().mainloop()
