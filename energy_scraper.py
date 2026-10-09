import argparse
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections.abc import Callable
from urllib.parse import urljoin, urlparse
from time import sleep
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from pypdf import PdfReader
from requests.exceptions import SSLError
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
    PLAYWRIGHT_AVAILABLE = True
except Exception:
    PLAYWRIGHT_AVAILABLE = False

DATA_DIR = Path.cwd() / "scraper_data"

# Broad Belgium candidate list. The script auto-filters to suppliers that
# appear to offer both electricity and gas (dual fuel).
CANDIDATE_SUPPLIERS = [
    {"supplier": "ENGIE", "start_url": "https://www.engie.be/nl/", "force_include": True},
    {"supplier": "Luminus", "start_url": "https://www.luminus.be/nl/", "force_include": True},
    {"supplier": "Eneco België", "start_url": "https://www.eneco.be/nl/", "force_include": True},
    {"supplier": "TotalEnergies Power & Gas Belgium", "start_url": "https://totalenergies.be/nl/particulieren", "force_include": True},
    {"supplier": "EnergyVision", "start_url": "https://www.energyvision.be/nl/", "force_include": True},
    {"supplier": "Mega", "start_url": "https://www.mega.be/nl/", "force_include": True},
    {"supplier": "Octa+", "start_url": "https://www.octaplus.be/nl/particulieren", "force_include": True},
    {"supplier": "Bolt", "start_url": "https://www.boltenergy.be/", "force_include": True},
    {"supplier": "Elegant", "start_url": "https://www.elegant.be/nl/", "force_include": True},
    {"supplier": "DATS 24", "start_url": "https://www.dats24.be/nl/particulier/energie.html", "force_include": True},
    {"supplier": "Trevion", "start_url": "https://www.trevion.be/", "force_include": True},
    {"supplier": "Watz", "start_url": "https://www.watz.be/", "force_include": True},
    {"supplier": "EBEM", "start_url": "https://www.ebem.be/", "force_include": True},
    {"supplier": "Power Online", "start_url": "https://www.poweronline.be/", "force_include": True},
    {"supplier": "Ecopower", "start_url": "https://www.ecopower.be/", "force_include": False},
    {"supplier": "Cociter", "start_url": "https://www.cociter.be/", "force_include": False},
    {"supplier": "Aspiravi Energy", "start_url": "https://www.aspiravi-energy.be/nl/", "force_include": False},
]

# Official Belgian comparison tools/regulators (not supplier pages, but useful).
COMPARATORS_AND_REGULATORS = {
    "VREG V-test (Vlaanderen)": "https://vtest.vreg.be/",
    "CWaPE comparator (Wallonie)": "https://www.cwape.be/consommateurs/comparer",
    "BRUGEL BRUSIM (Brussel)": "https://www.brugel.brussels/",
    "CREG": "https://www.creg.be/",
}

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

INSECURE_SSL_DOMAINS = {
    "www.octaplus.be",
    "www.watz.be",
}

# Known stable tariff pages per supplier — tried before heuristic discovery.
# These are publicly required regulatory documents (tariefkaarten).
# If a supplier moves a page, heuristic discovery still kicks in as fallback.
DIRECT_TARIFF_SEEDS: dict[str, list[str]] = {
    "ENGIE": [
        "https://www.engie.be/nl/woning/elektriciteit-gas/contracten/tariefkaarten/",
        "https://www.engie.be/nl/woning/elektriciteit-gas/contracten/tarieven/",
    ],
    "Luminus": [
        "https://www.luminus.be/nl/particulieren/producten/elektriciteit-gas/tarieven",
        "https://www.luminus.be/nl/tarieven",
    ],
    "Eneco België": [
        "https://www.eneco.be/nl/tarieven/",
        "https://www.eneco.be/nl/energie/tariefkaarten/",
    ],
    "TotalEnergies Power & Gas Belgium": [
        "https://totalenergies.be/nl/particulieren/elektriciteit-en-gas/onze-aanbiedingen/tarieven",
        "https://totalenergies.be/nl/particulieren/tarieven",
    ],
    "EnergyVision": [
        "https://www.energyvision.be/nl/tarieven/",
        "https://www.energyvision.be/nl/tariefkaart/",
    ],
    "Mega": [
        "https://www.mega.be/nl/tariefkaarten",
        "https://www.mega.be/nl/tarieven",
    ],
    "Octa+": [
        "https://www.octaplus.be/nl/particulieren/tarieven",
        "https://www.octaplus.be/nl/tariefkaarten",
    ],
    "Bolt": [
        "https://www.boltenergy.be/nl/tarieven",
        "https://www.boltenergy.be/nl/prijzen",
    ],
    "Elegant": [
        "https://www.elegant.be/nl/tarieven/",
        "https://www.elegant.be/nl/tariefkaarten/",
    ],
    "DATS 24": [
        "https://www.dats24.be/nl/particulier/energie/tarieven.html",
        "https://www.dats24.be/nl/particulier/energie/elektriciteit-en-gas.html",
    ],
    "Trevion": [
        "https://www.trevion.be/tarieven",
        "https://www.trevion.be/tariefkaarten",
    ],
    "Watz": [
        "https://www.watz.be/tarieven",
    ],
    "EBEM": [
        "https://www.ebem.be/nl/tarieven",
        "https://www.ebem.be/tarieven",
    ],
}

KEYWORD_PATTERN = re.compile(
    r"(tarief|tarif|tariff|prijs|price|prijzen|prijslijst|documents|brochure|fiche|productfiche|contract)",
    flags=re.IGNORECASE,
)

ELECTRICITY_PATTERN = re.compile(
    r"(elektriciteit|electricite|electricity|stroom|power|kwh)",
    flags=re.IGNORECASE,
)

GAS_PATTERN = re.compile(
    r"(aardgas|gas|gaz\s+naturel|natural\s+gas)",
    flags=re.IGNORECASE,
)

ENERGY_PRICE_REGEX = re.compile(
    r"(?<![\d.,])(?P<price>\d{1,3}(?:[.,]\d{1,8})?)(?![\d.,])"
    r"\s*(?P<currency>c€|c euro|ct|cent|€|eur|euro)\s*(?:/|per)?\s*(?P<unit>kwh|mwh)",
    flags=re.IGNORECASE,
)

# Also match formats like "0,12345 €/kWh" where number > typical cent values
ENERGY_PRICE_REGEX_ALT = re.compile(
    r"(?P<price>\d{1,3}(?:[.,]\d{1,6})?)\s*(?P<unit>kwh|mwh)",
    flags=re.IGNORECASE,
)

ANNUAL_FEE_REGEX = re.compile(
    r"(?<![\d.,])(?P<price>\d{1,4}(?:[.,]\d{1,8})?)(?![\d.,])"
    r"\s*(?P<currency>€|eur|euro)\s*(?:/|per)?\s*(?P<unit>jaar|an|maand|month|year)",
    flags=re.IGNORECASE,
)


def parse_energy_rate_to_eur_per_kwh(raw_value: str | None) -> float | None:
    """Convert values like '97,64 €/MWh' or '1,06c€/kWh' to numeric EUR/kWh."""
    if not raw_value or not isinstance(raw_value, str):
        return None

    match = ENERGY_PRICE_REGEX.search(raw_value)
    if not match:
        return None

    try:
        number = float(match.group("price").replace(",", "."))
    except Exception:
        return None

    currency = (match.group("currency") or "").lower()
    unit = (match.group("unit") or "").lower()

    # Convert cents to EUR.
    if currency in {"c€", "ct", "cent"}:
        number = number / 100.0

    # Convert EUR/MWh to EUR/kWh.
    if unit == "mwh":
        number = number / 1000.0

    return number


def parse_standing_charge_to_eur_per_year(raw_value: str | None) -> float | None:
    """Convert values like '10 €/maand' or '120 €/jaar' to EUR/year."""
    if not raw_value or not isinstance(raw_value, str):
        return None

    match = ANNUAL_FEE_REGEX.search(raw_value)
    if not match:
        return None

    try:
        number = float(match.group("price").replace(",", "."))
    except Exception:
        return None

    unit = (match.group("unit") or "").lower()
    if unit == "maand":
        return number * 12.0
    return number


def estimate_annual_cost_eur(row: dict) -> float | None:
    """Create one comparable annual estimate with a fixed household profile."""
    single_rate = parse_energy_rate_to_eur_per_kwh(row.get("electricity_single_rate"))
    day_rate = parse_energy_rate_to_eur_per_kwh(row.get("electricity_day_rate"))
    night_rate = parse_energy_rate_to_eur_per_kwh(row.get("electricity_night_rate"))
    gas_rate = parse_energy_rate_to_eur_per_kwh(row.get("gas_rate"))
    standing_year = parse_standing_charge_to_eur_per_year(row.get("standing_charge")) or 0.0

    # Simple fixed profile for comparability.
    annual_kwh_electricity = 3500.0
    annual_kwh_gas = 17000.0

    elec_rate = None
    if single_rate is not None:
        elec_rate = single_rate
    elif day_rate is not None and night_rate is not None:
        elec_rate = 0.6 * day_rate + 0.4 * night_rate
    elif day_rate is not None:
        elec_rate = day_rate
    elif night_rate is not None:
        elec_rate = night_rate

    if elec_rate is None or gas_rate is None:
        return None

    return (annual_kwh_electricity * elec_rate) + (annual_kwh_gas * gas_rate) + standing_year


def build_comparison_ready_df(overview_df: pd.DataFrame) -> pd.DataFrame:
    """Add normalized numeric columns and a sortable estimate for easier supplier choice."""
    if overview_df.empty:
        return overview_df.copy()

    rows = []
    for _, row in overview_df.iterrows():
        row_dict = row.to_dict()

        single_num = parse_energy_rate_to_eur_per_kwh(row_dict.get("electricity_single_rate"))
        day_num = parse_energy_rate_to_eur_per_kwh(row_dict.get("electricity_day_rate"))
        night_num = parse_energy_rate_to_eur_per_kwh(row_dict.get("electricity_night_rate"))
        gas_num = parse_energy_rate_to_eur_per_kwh(row_dict.get("gas_rate"))
        standing_year = parse_standing_charge_to_eur_per_year(row_dict.get("standing_charge"))
        est_annual = estimate_annual_cost_eur(row_dict)

        row_dict["electricity_single_eur_per_kwh"] = single_num
        row_dict["electricity_day_eur_per_kwh"] = day_num
        row_dict["electricity_night_eur_per_kwh"] = night_num
        row_dict["gas_eur_per_kwh"] = gas_num
        row_dict["standing_charge_eur_per_year"] = standing_year
        row_dict["estimated_annual_cost_eur"] = est_annual

        has_elec = any(value is not None for value in [single_num, day_num, night_num])
        has_gas = gas_num is not None
        row_dict["is_comparable"] = bool(has_elec and has_gas)

        rows.append(row_dict)

    out = pd.DataFrame(rows)
    out = out.sort_values(
        ["is_comparable", "estimated_annual_cost_eur", "fields_found", "supplier"],
        ascending=[False, True, False, True],
        na_position="last",
    ).reset_index(drop=True)
    return out


def qualifies_as_dual_fuel(base_url: str) -> tuple[bool, str]:
    """Heuristic check: supplier page mentions both electricity and gas."""
    try:
        response = request_with_retry(base_url)
    except Exception as exc:
        return False, f"fetch_error: {exc}"

    soup = BeautifulSoup(response.text, "html.parser")
    visible_text = soup.get_text(" ", strip=True)
    sample_text = visible_text[:120000]

    has_electricity = bool(ELECTRICITY_PATTERN.search(sample_text))
    has_gas = bool(GAS_PATTERN.search(sample_text))

    if has_electricity and has_gas:
        return True, "matched_text_terms"
    if has_electricity and not has_gas:
        return False, "electricity_only_signal"
    if has_gas and not has_electricity:
        return False, "gas_only_signal"
    return False, "no_clear_energy_terms"


def discover_tariff_links(base_url: str, max_links: int = 40) -> list[str]:
    """Discover likely tariff pages by scanning anchor text and href on a start page."""
    try:
        response = request_with_retry(base_url)
    except Exception as exc:
        print(f"[DISCOVERY ERROR] {base_url} -> {exc}")
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    links: list[str] = []
    seen = set()

    for anchor in soup.select("a[href]"):
        href = (anchor.get("href") or "").strip()
        text = anchor.get_text(" ", strip=True)
        haystack = f"{text} {href}"

        if not href:
            continue
        if href.startswith("#"):
            continue
        if href.lower().startswith("javascript:"):
            continue
        if not KEYWORD_PATTERN.search(haystack):
            continue

        absolute = urljoin(base_url, href)
        if absolute in seen:
            continue

        seen.add(absolute)
        links.append(absolute)

        if len(links) >= max_links:
            break

    return links


def discover_tariff_links_playwright(base_url: str, max_links: int = 40) -> list[str]:
    """Discover likely tariff links by rendering JS with Edge (fallback: Firefox) via Playwright."""
    if not PLAYWRIGHT_AVAILABLE:
        return []

    links: list[str] = []
    seen = set()

    try:
        with sync_playwright() as p:
            try:
                browser = p.chromium.launch(channel="msedge", headless=True)
            except Exception:
                browser = p.firefox.launch(headless=True)
            context = browser.new_context(user_agent=HEADERS["User-Agent"])
            page = context.new_page()

            try:
                page.goto(base_url, wait_until="domcontentloaded", timeout=35000)
                page.wait_for_timeout(2200)
            except PlaywrightTimeoutError:
                context.close()
                browser.close()
                return []

            try:
                page.evaluate(
                    """
                    async () => {
                        for (let i = 0; i < 3; i++) {
                            window.scrollBy(0, Math.max(window.innerHeight, 700));
                            await new Promise(r => setTimeout(r, 450));
                        }
                    }
                    """
                )
            except Exception:
                pass

            try:
                anchor_rows = page.eval_on_selector_all(
                    "a[href]",
                    """
                    (nodes) => nodes.map(a => ({
                        href: a.href || a.getAttribute('href') || '',
                        text: (a.textContent || '').trim(),
                        title: a.title || ''
                    }))
                    """,
                )
            except Exception:
                anchor_rows = []

            for row in anchor_rows:
                href = (row.get("href") or "").strip()
                text = row.get("text") or ""
                title = row.get("title") or ""
                haystack = f"{text} {title} {href}"

                if not href:
                    continue
                if href.startswith("#"):
                    continue
                if href.lower().startswith("javascript:"):
                    continue
                if not KEYWORD_PATTERN.search(haystack) and ".pdf" not in href.lower():
                    continue

                absolute = urljoin(base_url, href)
                if absolute in seen:
                    continue

                seen.add(absolute)
                links.append(absolute)
                if len(links) >= max_links:
                    break

            if len(links) < max_links:
                try:
                    resource_urls = page.evaluate(
                        """
                        () => {
                            return performance
                                .getEntriesByType('resource')
                                .map(r => (r && r.name) ? String(r.name) : '')
                                .filter(Boolean);
                        }
                        """
                    )
                except Exception:
                    resource_urls = []

                for resource_url in resource_urls:
                    candidate = str(resource_url)
                    lowered = candidate.lower()
                    if not KEYWORD_PATTERN.search(lowered) and ".pdf" not in lowered:
                        continue
                    if candidate in seen:
                        continue

                    seen.add(candidate)
                    links.append(candidate)
                    if len(links) >= max_links:
                        break

            context.close()
            browser.close()

    except Exception as exc:
        print(f"[PLAYWRIGHT DISCOVERY ERROR] {base_url} -> {exc}")
        return []

    return links


def request_with_retry(url: str, attempts: int = 3, timeout: int = 25) -> requests.Response:
    """Retry transient network failures so supplier pages are less likely to be skipped."""
    last_error = None
    hostname = urlparse(url).netloc.lower()
    verify_ssl = hostname not in INSECURE_SSL_DOMAINS

    for attempt in range(1, attempts + 1):
        try:
            response = requests.get(url, headers=HEADERS, timeout=timeout, verify=verify_ssl)
            response.raise_for_status()
            return response
        except SSLError as exc:
            last_error = exc
            verify_ssl = False
        except Exception as exc:
            last_error = exc
            if attempt < attempts:
                sleep(1.2 * attempt)

    raise RuntimeError(str(last_error))


def fetch_source_text(
    url: str,
    *,
    attempts: int = 3,
    timeout: int = 25,
) -> tuple[str, str, str]:
    """Fetch text from HTML or PDF pages for downstream heuristic extraction."""
    response = request_with_retry(url, attempts=attempts, timeout=timeout)
    content_type = response.headers.get("Content-Type", "").lower()
    source_type = "pdf" if url.lower().endswith(".pdf") or "pdf" in content_type else "html"

    if source_type == "pdf":
        reader = PdfReader(BytesIO(response.content))
        pages = [page.extract_text() or "" for page in reader.pages]
        text = "\n".join(pages)
        return text, source_type, content_type

    soup = BeautifulSoup(response.text, "html.parser")

    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    text = soup.get_text("\n", strip=True)
    return text, source_type, content_type


def normalize_text(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = re.sub(r"\r", "\n", text)
    text = re.sub(r"\n{2,}", "\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()


def discover_child_document_links(base_url: str, max_links: int = 12) -> list[str]:
    """Find likely tariff-card child links, preferring PDFs and explicit pricing docs."""
    try:
        response = request_with_retry(base_url)
    except Exception:
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    candidates: list[tuple[int, str]] = []
    seen = set()

    for anchor in soup.select("a[href]"):
        href = (anchor.get("href") or "").strip()
        text = anchor.get_text(" ", strip=True)
        haystack = f"{text} {href}".lower()

        if not href or href.startswith("#") or href.startswith("javascript:"):
            continue

        if not KEYWORD_PATTERN.search(haystack) and ".pdf" not in haystack:
            continue

        absolute = urljoin(base_url, href)
        if absolute in seen:
            continue

        score = 0
        if absolute.lower().endswith(".pdf"):
            score += 120
        if "tariefkaart" in haystack or "tariefkaarten" in haystack:
            score += 100
        if "prijslijst" in haystack or "prijslijsten" in haystack:
            score += 95
        if "prijzen-voorwaarden" in haystack:
            score += 90
        if "pdf" in haystack:
            score += 70
        if "simulatie" in haystack:
            score -= 30
        if "blog" in haystack or "inspiratie" in haystack:
            score -= 80

        seen.add(absolute)
        candidates.append((score, absolute))

    candidates.sort(key=lambda item: (-item[0], item[1]))
    return [url for _, url in candidates[:max_links]]


def extract_value_from_context(text: str, keywords: list[str], regex: re.Pattern[str], window: int = 180) -> tuple[str | None, str | None]:
    """Find the first price-like match near keyword occurrences."""
    lowered = text.lower()

    for keyword in keywords:
        start = 0
        keyword_lower = keyword.lower()

        while True:
            index = lowered.find(keyword_lower, start)
            if index == -1:
                break

            snippet_start = max(0, index - 60)
            snippet_end = min(len(text), index + window)
            snippet = text[snippet_start:snippet_end]
            match = regex.search(snippet)
            if match:
                return match.group(0).strip(), snippet.strip().replace("\n", " | ")

            start = index + len(keyword_lower)

    return None, None


def extract_fields_from_text(text: str, supplier: str, source_url: str, source_type: str) -> tuple[dict, list[dict]]:
    """Extract target fields from one text blob."""
    field_values = {
        "electricity_single_rate": None,
        "electricity_day_rate": None,
        "electricity_night_rate": None,
        "gas_rate": None,
        "standing_charge": None,
    }
    detail_rows = []

    target_fields = [
        ("electricity_single_rate", ["enkel", "enkelvoudig", "simple", "mono", "single", "eénuurs", "enkeltarief", "enkeltarieven", "elektriciteit"], ENERGY_PRICE_REGEX),
        ("electricity_day_rate", ["dag", "day", "peak", "dagtarief", "piek"], ENERGY_PRICE_REGEX),
        ("electricity_night_rate", ["nacht", "night", "off-peak", "dal", "nachttarief", "nachtuur"], ENERGY_PRICE_REGEX),
        ("gas_rate", ["aardgas", "gas", "gaz naturel", "gasprijs", "gaz", "natural gas", "aardgastarief"], ENERGY_PRICE_REGEX),
        ("standing_charge", ["vaste vergoeding", "abonnement", "vergoeding fixe", "annual fee", "forfait", "vast tarief", "vaste kost", "aansluitkost", "vaste bijdrage"], ANNUAL_FEE_REGEX),
    ]

    for field_name, keywords, regex in target_fields:
        value, snippet = extract_value_from_context(text, keywords, regex)
        field_values[field_name] = value
        if value:
            detail_rows.append(
                {
                    "supplier": supplier,
                    "field": field_name,
                    "value": value,
                    "source_url": source_url,
                    "source_type": source_type,
                    "snippet": snippet,
                }
            )

    return field_values, detail_rows


def extract_best_effort_prices(
    best_urls_df: pd.DataFrame,
    progress_callback: Callable[[str], None] | None = None,
    *,
    fast_mode: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Extract price-like values from the best supplier pages into wide and detailed outputs."""
    indexed_rows = list(best_urls_df.iterrows())

    def extract_one(index: int, row: pd.Series) -> tuple[int, dict, list[dict]]:
        supplier = row["supplier"]
        source_url = row["discovered_url"]
        if progress_callback:
            progress_callback(f"Tarieven uitlezen ({index + 1}/{len(indexed_rows)}): {supplier}")
        extracted = {
            "supplier": supplier,
            "source_url": source_url,
            "matched_url": None,
            "source_type": None,
            "electricity_single_rate": None,
            "electricity_day_rate": None,
            "electricity_night_rate": None,
            "gas_rate": None,
            "standing_charge": None,
            "fields_found": 0,
            "extraction_note": None,
        }
        detail_rows: list[dict] = []

        try:
            if fast_mode:
                candidate_urls = [source_url] if isinstance(source_url, str) and source_url else []
            else:
                # Full mode checks known seed pages and follows tariff documents.
                seed_urls = DIRECT_TARIFF_SEEDS.get(supplier, [])
                candidate_urls = []
                seen_candidates: set[str] = set()
                for candidate_url in seed_urls + [source_url]:
                    if candidate_url and candidate_url not in seen_candidates:
                        seen_candidates.add(candidate_url)
                        candidate_urls.append(candidate_url)

                child_links: list[str] = []
                for start_url in candidate_urls[:3]:
                    for child in discover_child_document_links(start_url):
                        if child not in seen_candidates:
                            seen_candidates.add(child)
                            child_links.append(child)
                candidate_urls.extend(child_links)

            best_candidate = None
            best_candidate_details = []

            for candidate_url in candidate_urls:
                try:
                    if fast_mode:
                        source_text, source_type, content_type = fetch_source_text(
                            candidate_url,
                            attempts=1,
                            timeout=12,
                        )
                    else:
                        source_text, source_type, content_type = fetch_source_text(candidate_url)
                except Exception:
                    continue

                normalized_text = normalize_text(source_text)
                if not normalized_text:
                    continue

                candidate_values, candidate_details = extract_fields_from_text(
                    normalized_text,
                    supplier=supplier,
                    source_url=candidate_url,
                    source_type=source_type,
                )
                fields_found = sum(1 for value in candidate_values.values() if value)
                candidate_result = {
                    **candidate_values,
                    "matched_url": candidate_url,
                    "source_type": source_type,
                    "content_type": content_type,
                    "fields_found": fields_found,
                }

                if best_candidate is None or candidate_result["fields_found"] > best_candidate["fields_found"]:
                    best_candidate = candidate_result
                    best_candidate_details = candidate_details

                if best_candidate is not None and best_candidate["fields_found"] >= 3:
                    break

            if best_candidate is None:
                extracted["extraction_note"] = "no_text_extracted"
            else:
                extracted.update(best_candidate)
                detail_rows.extend(best_candidate_details)

                if extracted["fields_found"] == 0:
                    extracted["extraction_note"] = "no_target_price_found"
                else:
                    extracted["extraction_note"] = f"{extracted['fields_found']}_fields_found"

        except Exception as exc:
            extracted["extraction_note"] = f"error: {exc}"

        return index, extracted, detail_rows

    comparison_rows: list[dict | None] = [None] * len(indexed_rows)
    detail_rows: list[dict] = []
    if fast_mode and len(indexed_rows) > 1:
        with ThreadPoolExecutor(max_workers=min(8, len(indexed_rows))) as executor:
            futures = [
                executor.submit(extract_one, index, row)
                for index, (_, row) in enumerate(indexed_rows)
            ]
            for future in as_completed(futures):
                index, extracted, supplier_details = future.result()
                comparison_rows[index] = extracted
                detail_rows.extend(supplier_details)
    else:
        for index, (_, row) in enumerate(indexed_rows):
            _, extracted, supplier_details = extract_one(index, row)
            comparison_rows[index] = extracted
            detail_rows.extend(supplier_details)

    comparison_df = pd.DataFrame([row for row in comparison_rows if row is not None])
    detail_df = pd.DataFrame(detail_rows)
    return comparison_df, detail_df


def scrape_tables_from_url(url: str) -> pd.DataFrame:
    """Read HTML tables from a page and flatten into a row-based table."""
    try:
        tables = pd.read_html(url)
    except Exception:
        return pd.DataFrame()

    rows = []
    for index, table in enumerate(tables, start=1):
        table = table.copy()
        table.columns = [str(col).strip() for col in table.columns]
        table["source_url"] = url
        table["table_index"] = index
        rows.append(table)

    if not rows:
        return pd.DataFrame()

    return pd.concat(rows, ignore_index=True)


def collect_supplier_qualification(
    progress_callback: Callable[[str], None] | None = None,
) -> pd.DataFrame:
    rows = []

    for candidate in CANDIDATE_SUPPLIERS:
        supplier = candidate["supplier"]
        base_url = candidate["start_url"]
        force_include = candidate["force_include"]
        if progress_callback:
            progress_callback(f"Leverancier controleren: {supplier}")

        is_dual, reason = qualifies_as_dual_fuel(base_url)
        qualification_mode = "heuristic"

        if not is_dual and force_include:
            is_dual = True
            reason = f"manual_override_after_{reason}"
            qualification_mode = "manual_override"

        rows.append(
            {
                "supplier": supplier,
                "start_url": base_url,
                "is_dual_fuel": is_dual,
                "qualification_reason": reason,
                "qualification_mode": qualification_mode,
                "force_include": force_include,
            }
        )

    return pd.DataFrame(rows)


def collect_all_discovered_links(
    qualified_suppliers_df: pd.DataFrame,
    progress_callback: Callable[[str], None] | None = None,
) -> pd.DataFrame:
    discovered_rows = []

    eligible = qualified_suppliers_df.loc[
        qualified_suppliers_df["is_dual_fuel"] == True, ["supplier", "start_url"]
    ]

    for _, supplier_row in eligible.iterrows():
        supplier = supplier_row["supplier"]
        base_url = supplier_row["start_url"]
        if progress_callback:
            progress_callback(f"Tariefpagina's zoeken: {supplier}")
        static_links = discover_tariff_links(base_url)
        dynamic_links = []

        # JS-heavy supplier sites often hide tariff URLs in rendered content.
        if len(static_links) < 3:
            dynamic_links = discover_tariff_links_playwright(base_url)

        links = []
        seen = set()
        for candidate in static_links + dynamic_links:
            if candidate in seen:
                continue
            seen.add(candidate)
            links.append(candidate)

        if not links:
            discovered_rows.append(
                {
                    "supplier": supplier,
                    "start_url": base_url,
                    "discovered_url": None,
                }
            )
            continue

        for link in links:
            discovered_rows.append(
                {
                    "supplier": supplier,
                    "start_url": base_url,
                    "discovered_url": link,
                }
            )

    return pd.DataFrame(discovered_rows)


def scrape_discovered_pages(discovered_df: pd.DataFrame) -> pd.DataFrame:
    output_frames = []

    for _, row in discovered_df.dropna(subset=["discovered_url"]).iterrows():
        supplier = row["supplier"]
        url = row["discovered_url"]

        page_df = scrape_tables_from_url(url)
        if page_df.empty:
            continue

        page_df["supplier"] = supplier
        output_frames.append(page_df)

    if not output_frames:
        return pd.DataFrame()

    return pd.concat(output_frames, ignore_index=True)


def build_curated_urls(discovered_df: pd.DataFrame) -> pd.DataFrame:
    """Keep links that are likely useful for pricing extraction."""
    if discovered_df.empty:
        return discovered_df.copy()

    include_pattern = re.compile(
        r"(?:tarief|tarif|tariff|prijslijst|tariefkaart|prijzen|price|contract|simulatie|estimate|calculator|energy-price-calculator|elektriciteit-gas|electricite-gaz|pdf)",
        flags=re.IGNORECASE,
    )
    exclude_pattern = re.compile(
        r"(?:blog|voordelen|encyclopedie|inspiratie|faq|veelgestelde-vragen|hulp|contact)",
        flags=re.IGNORECASE,
    )

    curated = discovered_df.dropna(subset=["discovered_url"]).copy()
    curated["is_relevant"] = curated["discovered_url"].str.contains(include_pattern)
    curated["is_noise"] = curated["discovered_url"].str.contains(exclude_pattern)

    curated = curated[(curated["is_relevant"]) & (~curated["is_noise"])].copy()
    curated = curated.drop(columns=["is_relevant", "is_noise"])
    curated = curated.drop_duplicates(subset=["supplier", "discovered_url"]).reset_index(drop=True)
    return curated


def select_best_url_per_supplier(curated_urls_df: pd.DataFrame, qualification_df: pd.DataFrame) -> pd.DataFrame:
    """Pick the most likely tariff page per supplier based on URL scoring."""
    if curated_urls_df.empty:
        fallback_rows = qualification_df.loc[
            qualification_df["is_dual_fuel"] == True, ["supplier", "start_url"]
        ].copy()
        fallback_rows["discovered_url"] = fallback_rows["start_url"]
        fallback_rows["url_score"] = 0
        fallback_rows["selection_reason"] = "start_url_fallback"
        return fallback_rows[["supplier", "start_url", "discovered_url", "url_score", "selection_reason"]]

    def score_url(url: str) -> int:
        value = url.lower()
        score = 0

        priority_terms = {
            "tariefkaart": 100,
            "tariefkaarten": 95,
            "prijslijst": 90,
            "prijslijsten": 90,
            "prijs-berekenen": 88,
            "price-calc": 88,
            "prijzen-voorwaarden": 85,
            "overzicht-tariefkaarten": 85,
            "tarief": 70,
            "prijzen": 65,
            "simulatie": 50,
            "estimate": 50,
            "calculator": 45,
            ".pdf": 80,
        }

        penalty_terms = {
            "blog": -80,
            "voordelen": -80,
            "encyclopedie": -70,
            "inspiratie": -70,
            "hulp": -40,
            "contact": -40,
            "injectie": -90,
        }

        for term, points in priority_terms.items():
            if term in value:
                score += points

        for term, points in penalty_terms.items():
            if term in value:
                score += points

        return score

    best_urls = curated_urls_df.copy()
    best_urls["url_score"] = best_urls["discovered_url"].apply(score_url)
    best_urls = best_urls.sort_values(
        ["supplier", "url_score", "discovered_url"],
        ascending=[True, False, True],
    )
    best_urls = best_urls.drop_duplicates(subset=["supplier"], keep="first").reset_index(drop=True)
    best_urls["selection_reason"] = "ranked_curated_url"

    missing_suppliers = qualification_df.loc[
        (qualification_df["is_dual_fuel"] == True)
        & (~qualification_df["supplier"].isin(best_urls["supplier"])),
        ["supplier", "start_url"],
    ].copy()
    if not missing_suppliers.empty:
        missing_suppliers["discovered_url"] = missing_suppliers["start_url"]
        missing_suppliers["url_score"] = 0
        missing_suppliers["selection_reason"] = "start_url_fallback"
        best_urls = pd.concat([best_urls, missing_suppliers], ignore_index=True)

    return best_urls[["supplier", "start_url", "discovered_url", "url_score", "selection_reason"]].sort_values(["supplier"]).reset_index(drop=True)


def export_excel_workbook(
    qualification_df: pd.DataFrame,
    discovered_df: pd.DataFrame,
    curated_urls_df: pd.DataFrame,
    best_urls_df: pd.DataFrame,
    comparators_df: pd.DataFrame,
    prices_df: pd.DataFrame,
    comparison_df: pd.DataFrame,
    extraction_details_df: pd.DataFrame,
    output_path: str = "be_energy_overview.xlsx",
) -> str:
    """Export clean Excel workbook with separate tabs and a combined overview tab."""
    workbook_path = output_path

    try:
        writer = pd.ExcelWriter(workbook_path, engine="openpyxl")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        workbook_path = f"be_energy_overview_{timestamp}.xlsx"
        writer = pd.ExcelWriter(workbook_path, engine="openpyxl")

    with writer:
        q = qualification_df.sort_values(["is_dual_fuel", "supplier"], ascending=[False, True]).reset_index(drop=True)
        d = discovered_df.sort_values(["supplier", "discovered_url"], na_position="last").reset_index(drop=True)
        c = curated_urls_df.sort_values(["supplier", "discovered_url"], na_position="last").reset_index(drop=True)
        b = best_urls_df.sort_values(["supplier"]).reset_index(drop=True)
        p = prices_df.reset_index(drop=True)
        x = comparison_df.sort_values(["supplier"]).reset_index(drop=True)
        xd = extraction_details_df.sort_values(["supplier", "field"]).reset_index(drop=True)

        overview_df = x.merge(
            b[["supplier", "discovered_url", "selection_reason"]],
            on="supplier",
            how="left",
        )
        overview_df = overview_df.rename(
            columns={
                "discovered_url": "offer_url",
            }
        )
        overview_df = overview_df[
            [
                "supplier",
                "electricity_single_rate",
                "electricity_day_rate",
                "electricity_night_rate",
                "gas_rate",
                "standing_charge",
                "fields_found",
                "extraction_note",
                "offer_url",
                "selection_reason",
            ]
        ]
        comparison_ready_df = build_comparison_ready_df(overview_df)

        q.to_excel(writer, sheet_name="Suppliers_Qualification", index=False)
        d.to_excel(writer, sheet_name="Discovered_URLs", index=False)
        c.to_excel(writer, sheet_name="Curated_URLs", index=False)
        b.to_excel(writer, sheet_name="Best_URL_Per_Supplier", index=False)
        comparators_df.to_excel(writer, sheet_name="Comparators_Regulators", index=False)
        p.to_excel(writer, sheet_name="Raw_Prices", index=False)
        x.to_excel(writer, sheet_name="Best_Effort_Prices", index=False)
        xd.to_excel(writer, sheet_name="Extraction_Details", index=False)
        overview_df.to_excel(writer, sheet_name="Comparison_Clickable", index=False)
        comparison_ready_df.to_excel(writer, sheet_name="Comparison_Ready", index=False)

        sheet = "Combined_Overview"

        # Put qualification, best URLs and extracted prices next to each other.
        q.to_excel(writer, sheet_name=sheet, index=False, startrow=0, startcol=0)
        startcol_urls = len(q.columns) + 2
        b.to_excel(writer, sheet_name=sheet, index=False, startrow=0, startcol=startcol_urls)
        startcol_prices = startcol_urls + len(b.columns) + 2
        x.to_excel(writer, sheet_name=sheet, index=False, startrow=0, startcol=startcol_prices)

        # Put comparator links under the overview tables.
        max_rows = max(len(q), len(b), len(x)) + 3
        comparators_df.to_excel(writer, sheet_name=sheet, index=False, startrow=max_rows, startcol=0)

        workbook = writer.book

        # ── Style helper ──────────────────────────────────────────────────────
        header_fill_blue  = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_fill_green = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
        header_font_white = Font(bold=True, color="FFFFFF", size=11)
        hyperlink_font    = Font(color="0563C1", underline="single")
        green_fill        = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        orange_fill       = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"),  bottom=Side(style="thin"),
        )

        def style_sheet_headers(ws, ncols: int, fill):
            for ci in range(1, ncols + 1):
                cell = ws.cell(row=1, column=ci)
                cell.fill = fill
                cell.font = header_font_white
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border

        def style_sheet_data(ws, nrows: int, ncols: int):
            for row in ws.iter_rows(min_row=2, max_row=nrows + 1, min_col=1, max_col=ncols):
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=False)

        # ── Vergelijking sheet (primary, clean, sorted) ────────────────────
        verg_cols = {
            "Leverancier":         "supplier",
            "Enkel €/kWh":         "electricity_single_eur_per_kwh",
            "Dag €/kWh":           "electricity_day_eur_per_kwh",
            "Nacht €/kWh":         "electricity_night_eur_per_kwh",
            "Gas €/kWh":           "gas_eur_per_kwh",
            "Vaste kost €/jaar":   "standing_charge_eur_per_year",
            "Schatting €/jaar*":   "estimated_annual_cost_eur",
            "Compleet?":           "is_comparable",
            "Bron URL":            "offer_url",
        }
        verg_df = comparison_ready_df.rename(columns={v: k for k, v in verg_cols.items()})
        verg_available = [k for k in verg_cols if k in verg_df.columns]
        verg_df = verg_df[verg_available]
        verg_df["Compleet?"] = verg_df["Compleet?"].map({True: "Ja", False: "Nee"}).fillna("Nee")

        verg_df.to_excel(writer, sheet_name="Vergelijking", index=False)
        verg_sheet = workbook["Vergelijking"]
        style_sheet_headers(verg_sheet, len(verg_available), header_fill_green)
        style_sheet_data(verg_sheet, len(verg_df), len(verg_available))

        supplier_col_verg = list(verg_available).index("Leverancier") + 1
        bron_col_verg     = list(verg_available).index("Bron URL") + 1 if "Bron URL" in verg_available else None
        compleet_col      = list(verg_available).index("Compleet?") + 1 if "Compleet?" in verg_available else None
        schat_col         = list(verg_available).index("Schatting €/jaar*") + 1 if "Schatting €/jaar*" in verg_available else None

        for ri in range(2, len(verg_df) + 2):
            if bron_col_verg:
                url_cell = verg_sheet.cell(row=ri, column=bron_col_verg)
                sup_cell = verg_sheet.cell(row=ri, column=supplier_col_verg)
                if url_cell.value:
                    sup_cell.hyperlink = str(url_cell.value)
                    sup_cell.font = hyperlink_font
                    url_cell.hyperlink = str(url_cell.value)
                    url_cell.font = hyperlink_font
            if compleet_col:
                cc = verg_sheet.cell(row=ri, column=compleet_col)
                if cc.value == "Ja":
                    cc.fill = green_fill
                else:
                    cc.fill = orange_fill
            # Highlight estimated-cost cells that have a value (sortable ranking)
            if schat_col:
                sc = verg_sheet.cell(row=ri, column=schat_col)
                if sc.value is not None:
                    sc.number_format = "€#,##0.00"

        # Number format for €/kWh columns
        kwh_cols_idx = [i+1 for i, k in enumerate(verg_available) if "€/kWh" in k]
        for ri in range(2, len(verg_df) + 2):
            for ci in kwh_cols_idx:
                verg_sheet.cell(row=ri, column=ci).number_format = "0.00000"

        col_widths = {
            "Leverancier": 26, "Enkel €/kWh": 14, "Dag €/kWh": 13,
            "Nacht €/kWh": 13, "Gas €/kWh": 13, "Vaste kost €/jaar": 17,
            "Schatting €/jaar*": 18, "Compleet?": 11, "Bron URL": 50,
        }
        for ci, col_name in enumerate(verg_available, start=1):
            verg_sheet.column_dimensions[get_column_letter(ci)].width = col_widths.get(col_name, 18)

        verg_sheet.auto_filter.ref = f"A1:{get_column_letter(len(verg_available))}{len(verg_df)+1}"
        verg_sheet.freeze_panes = "B2"

        # Disclaimer row after data
        disclaimer_row = len(verg_df) + 3
        verg_sheet.cell(row=disclaimer_row, column=1).value = (
            "* Schatting op basis van: 3500 kWh elektriciteit(enkel) + 17000 kWh gas per jaar. "
            "Alleen publiceerde lijstprijzen — geen kortingen. Netto energieprijs, excl. distributie & taksen."
        )
        verg_sheet.cell(row=disclaimer_row, column=1).font = Font(italic=True, size=9, color="666666")
        verg_sheet.merge_cells(
            start_row=disclaimer_row, start_column=1,
            end_row=disclaimer_row, end_column=len(verg_available)
        )

        # ── Comparison_Clickable sheet styling ────────────────────────────
        clickable_sheet = workbook["Comparison_Clickable"]
        offer_col = list(overview_df.columns).index("offer_url") + 1
        for row_index in range(2, len(overview_df) + 2):
            supplier_cell = clickable_sheet.cell(row=row_index, column=1)
            offer_cell    = clickable_sheet.cell(row=row_index, column=offer_col)
            if offer_cell.value:
                supplier_cell.hyperlink = offer_cell.value
                supplier_cell.font = hyperlink_font
                offer_cell.hyperlink = offer_cell.value
                offer_cell.font = hyperlink_font

        style_sheet_headers(clickable_sheet, len(overview_df.columns), header_fill_blue)
        style_sheet_data(clickable_sheet, len(overview_df), len(overview_df.columns))
        for ci, col_name in enumerate(overview_df.columns, start=1):
            clickable_sheet.column_dimensions[get_column_letter(ci)].width = (
                16 if col_name in ["electricity_single_rate","electricity_day_rate",
                                   "electricity_night_rate","gas_rate","standing_charge"] else 20
            )
        clickable_sheet.auto_filter.ref = f"A1:{get_column_letter(len(overview_df.columns))}{len(overview_df)+1}"
        clickable_sheet.freeze_panes = "A2"

        combined_sheet = workbook["Combined_Overview"]
        combined_sheet.freeze_panes = "A2"

    return workbook_path


def _load_discovery_cache(data_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    cached_files = {
        "qualification": (data_dir / "be_supplier_qualification.csv", {"supplier", "start_url", "is_dual_fuel"}),
        "discovered": (data_dir / "be_supplier_discovered_urls.csv", {"supplier", "start_url", "discovered_url"}),
        "curated": (data_dir / "be_supplier_curated_urls.csv", {"supplier", "start_url", "discovered_url"}),
        "best": (data_dir / "be_supplier_best_urls.csv", {"supplier", "start_url", "discovered_url"}),
    }
    missing = [str(path.name) for path, _ in cached_files.values() if not path.is_file()]
    if missing:
        raise FileNotFoundError(
            "Geen volledige cache van leveranciers-URL’s gevonden; ontbreekt: "
            + ", ".join(missing)
        )

    frames = {}
    for name, (path, required_columns) in cached_files.items():
        frame = pd.read_csv(path, encoding="utf-8-sig")
        missing_columns = required_columns.difference(frame.columns)
        if missing_columns:
            raise ValueError(
                f"URL-cache {path.name} mist kolommen: {', '.join(sorted(missing_columns))}"
            )
        frames[name] = frame
    return frames["qualification"], frames["discovered"], frames["curated"], frames["best"]


def main(*, reuse_discovery: bool = False) -> None:
    print("Leveranciersprijzen bijwerken...", flush=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    qualification_df = pd.DataFrame()
    discovered_df = pd.DataFrame()
    curated_urls_df = pd.DataFrame()
    best_urls_df = pd.DataFrame()

    # 1-4) Reuse saved URL discovery for a faster price refresh, or rediscover everything.
    if reuse_discovery:
        try:
            (
                qualification_df,
                discovered_df,
                curated_urls_df,
                best_urls_df,
            ) = _load_discovery_cache(DATA_DIR)
            print(
                "Bestaande leveranciers-URL’s hergebruiken; "
                "gebruik --refresh-discovery voor een volledige bronzoektocht.",
                flush=True,
            )
        except (FileNotFoundError, ValueError, pd.errors.ParserError) as exc:
            print(f"URL-cache niet bruikbaar ({exc}); voer volledige bronzoektocht uit.", flush=True)
            reuse_discovery = False

    if not reuse_discovery:
        print("1/8 Leveranciers controleren...", flush=True)
        qualification_df = collect_supplier_qualification(
            progress_callback=lambda message: print(message, flush=True)
        )
        print("2/8 Tariefpagina’s zoeken...", flush=True)
        discovered_df = collect_all_discovered_links(
            qualification_df,
            progress_callback=lambda message: print(message, flush=True),
        )
        print("3/8 URL’s filteren en rangschikken...", flush=True)
        curated_urls_df = build_curated_urls(discovered_df)
        best_urls_df = select_best_url_per_supplier(curated_urls_df, qualification_df)

    qualification_df.to_csv(DATA_DIR / "be_supplier_qualification.csv", index=False, encoding="utf-8-sig")
    print("Saved: be_supplier_qualification.csv")

    discovered_df.to_csv(DATA_DIR / "be_supplier_discovered_urls.csv", index=False, encoding="utf-8-sig")
    print("Saved: be_supplier_discovered_urls.csv")

    curated_urls_df.to_csv(DATA_DIR / "be_supplier_curated_urls.csv", index=False, encoding="utf-8-sig")
    print("Saved: be_supplier_curated_urls.csv")

    best_urls_df.to_csv(DATA_DIR / "be_supplier_best_urls.csv", index=False, encoding="utf-8-sig")
    print("Saved: be_supplier_best_urls.csv")

    # 5) Save comparator/regulator references as static helper list.
    comparators_df = pd.DataFrame(
        [
            {"name": name, "url": url}
            for name, url in COMPARATORS_AND_REGULATORS.items()
        ]
    )
    comparators_df.to_csv(DATA_DIR / "be_comparators_regulators_urls.csv", index=False, encoding="utf-8-sig")
    print("Saved: be_comparators_regulators_urls.csv")

    # 6) Table scraping is only refreshed during a full discovery run.
    raw_prices_path = DATA_DIR / "be_energy_prices_raw.csv"
    if reuse_discovery:
        if raw_prices_path.is_file():
            prices_df = pd.read_csv(raw_prices_path, encoding="utf-8-sig")
            print(
                "HTML-tabellen niet opnieuw opgehaald; bestaande ruwe tabeldata behouden.",
                flush=True,
            )
        else:
            prices_df = pd.DataFrame()
            print("Geen bestaande ruwe tabeldata; HTML-tabellen overgeslagen in snelle modus.", flush=True)
    else:
        print("6/8 HTML-tabellen van tariefpagina’s ophalen...", flush=True)
        prices_df = scrape_discovered_pages(curated_urls_df)
        if prices_df.empty:
            print("No HTML tables found on discovered pages.")
        else:
            prices_df.to_csv(raw_prices_path, index=False, encoding="utf-8-sig")
            print("Saved: be_energy_prices_raw.csv")

    # 7) Extract best-effort price values from the strongest URL per supplier.
    print("7/8 Leverancierstarieven uitlezen...", flush=True)
    comparison_df, extraction_details_df = extract_best_effort_prices(
        best_urls_df,
        progress_callback=lambda message: print(message, flush=True),
        fast_mode=reuse_discovery,
    )
    comparison_df.to_csv(DATA_DIR / "be_energy_price_comparison.csv", index=False, encoding="utf-8-sig")
    extraction_details_df.to_csv(DATA_DIR / "be_energy_price_extraction_details.csv", index=False, encoding="utf-8-sig")
    print("Saved: be_energy_price_comparison.csv")
    print("Saved: be_energy_price_extraction_details.csv")

    # 8) Export everything to one Excel workbook.
    print("8/8 Excel-overzicht bijwerken...", flush=True)
    workbook_path = export_excel_workbook(
        qualification_df=qualification_df,
        discovered_df=discovered_df,
        curated_urls_df=curated_urls_df,
        best_urls_df=best_urls_df,
        comparators_df=comparators_df,
        prices_df=prices_df,
        comparison_df=comparison_df,
        extraction_details_df=extraction_details_df,
        output_path=str(DATA_DIR / "be_energy_overview.xlsx"),
    )
    print(f"Saved: {workbook_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scrape Belgian supplier energy prices.")
    parser.add_argument(
        "--reuse-discovery",
        action="store_true",
        help="Reuse saved supplier URL lists and refresh prices without rediscovering all pages.",
    )
    parser.add_argument(
        "--refresh-discovery",
        action="store_true",
        help="Force a full supplier and tariff URL discovery before refreshing prices.",
    )
    args = parser.parse_args()
    main(reuse_discovery=args.reuse_discovery and not args.refresh_discovery)