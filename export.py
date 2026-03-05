"""VectorWorks Import TXT/XLSX generator."""

import io

import geopandas as gpd
from pyproj import CRS, Transformer

LS320_WKT = (
    'PROJCS["ETRS89 / Gauss-Kruger CM 9E (LS320)",'
    'GEOGCS["ETRS89",'
    'DATUM["European_Terrestrial_Reference_System_1989",'
    'SPHEROID["GRS 1980",6378137,298.257222101]],'
    'PRIMEM["Greenwich",0],'
    'UNIT["degree",0.0174532925199433]],'
    'PROJECTION["Transverse_Mercator"],'
    'PARAMETER["latitude_of_origin",0],'
    'PARAMETER["central_meridian",9],'
    'PARAMETER["scale_factor",1],'
    'PARAMETER["false_easting",3500000],'
    'PARAMETER["false_northing",0],'
    'UNIT["metre",1],'
    'AXIS["Easting",EAST],'
    'AXIS["Northing",NORTH]]'
)
LS320_CRS = CRS.from_wkt(LS320_WKT)

CRS_OPTIONS = {
    "EPSG:25832": "ETRS89 / UTM zone 32N",
    "EPSG:31467": "DHDN / 3-degree Gauss-Kruger zone 3 (3xxxxxx)",
    "EPSG:4647": "ETRS89 / UTM zone 32N (zE-N, 32xxxxxx)",
    "LS320": "Hamburg LS320 — ETRS89 / GK CM 9E (FE=3500000)",
}


def resolve_crs(key: str):
    """Return a pyproj-compatible CRS for a given key."""
    if key == "LS320":
        return LS320_CRS
    return key


def estimate_ansatzhoehe(hoehe, kronendurchmesser, method="ratio", ratio=0.25):
    """Estimate Ansatzhöhe (crown base height) from tree height and crown diameter.

    Methods:
    - "ratio": Ansatzhöhe = Höhe × ratio  (default 0.25, typical range 0.15-0.40)
    - "kd":    Ansatzhöhe = Höhe - Kronendurchmesser (crude, often overestimates)
    """
    try:
        h = float(hoehe)
    except (TypeError, ValueError):
        return ""

    if method == "kd":
        try:
            kd = float(kronendurchmesser)
            result = max(1.0, h - kd)
        except (TypeError, ValueError):
            result = max(1.0, h * ratio)
    else:
        result = max(1.0, h * ratio)

    return round(result)


def _fmt(f):
    """Format a field value for TXT output."""
    if f is None:
        return ""
    s = str(f)
    return "" if s in ("nan", "None") else s


# VW parameter names (German) — exact match to VW Import Tree Survey dropdown.
# Every parameter is included so VW can auto-map by column header name.
# Empty columns are fine — VW maps them to <Kein>.
VW_COLUMNS = [
    # Coordinates
    "x-Koordinate",
    "y-Koordinate",
    # Identity
    "Baum-ID",
    "Botanischer Name",
    "Deutscher Name",
    "Familie",
    "Herkunft",
    "Invasive Art",
    # Dimensions
    "Höhe",
    "Kronendurchmesser",
    "Krone Nord (N)",
    "Krone Nordost (NO)",
    "Krone Ost (O)",
    "Krone Südost (SO)",
    "Krone Süd (S)",
    "Krone Südwest (SW)",
    "Krone West (W)",
    "Krone Nordwest (NW)",
    "Lichte Höhe",
    "Astansatzhöhe",
    "Ausrichtung Hauptast",
    "Stammumfang",
    # Status
    "Maßnahme",
    "Veteranenbaum",
    "Alter Baum",
    "Pflanzjahr",
    "Letzte Überprüfung am",
    "Standort",
    "Bemerkungen",
    "Form",
    "Struktur",
    "Vitalität",
    "Stammfuß-ø",
    # Custom fields
    "Feld 1",
    "Feld 2",
    "Feld 3",
    "Feld 4",
    "Feld 5",
    "Feld 6",
    "Zusatz 07",
    "Zusatz 08",
    "Zusatz 09",
    "Zusatz 10",
]


# VW Maßnahme values — map our action terms to exact VW dropdown values.
# Values already matching VW pass through; legacy terms get mapped.
_ACTION_TO_VW = {
    # Legacy terms (from old exports or English usage)
    "Retain": "Sichern",
    "Remove": "Entfernen",
    "Transplant": "Umpflanzen - Am Standort",
    "Roden": "Entfernen",
    "Umpflanzen": "Umpflanzen - Am Standort",
}

# VW Vitalitätsstufe — map numeric grades to exact VW dropdown strings
_VITALITAET_TO_VW = {
    "0": "0 Krone harmonisch geschlossen, fast kein Totholz in der Krone",
    "1": "1 Kronenmantel an wenigen Stellen zerklüftet, wenig Totholz im Dünnast- und Starkastbereich",
    "2": "2 Vermehrt Totholz, Bildung einer Sekundärkrone",
    "3": "3 Absterben von Ästen, sehr viel Totholz in der Krone",
}


def _map_action(val):
    """Map action value to exact VW Maßnahme dropdown value."""
    s = _fmt(val)
    return _ACTION_TO_VW.get(s, s)


def _map_vitalitaet(val):
    """Map vitalitaet value to exact VW Vitalitätsstufe dropdown value.

    Handles single values (0, 1, 2, 3) and ranges (0-1, 1-2).
    For ranges, uses the worse (higher) value.
    """
    s = _fmt(val).strip().replace("–", "-")
    if not s:
        return ""
    # If it's a range like "1-2", take the worse (higher) value
    if "-" in s:
        parts = s.split("-")
        try:
            worst = str(max(int(p.strip()) for p in parts))
            return _VITALITAET_TO_VW.get(worst, s)
        except ValueError:
            pass
    return _VITALITAET_TO_VW.get(s, s)


def _build_vw_row(row, x, y, ansatzhoehe="", standort=""):
    """Build a row list matching VW_COLUMNS from a GeoDataFrame row."""
    deutsch = row.get("art_deutsch", "") or row.get("gattung_deutsch", "")
    latein = row.get("art_latein", "") or row.get("gattung_latein", "")

    return [
        # Coordinates
        f"{x:.3f}",
        f"{y:.3f}",
        # Identity
        row.get("baum_id", ""),
        latein,
        deutsch,
        row.get("familie", ""),
        row.get("herkunft", ""),
        row.get("invasive_art", ""),
        # Dimensions
        row.get("baumhoehe", ""),
        row.get("kronendurchmesser", ""),
        row.get("krone_n", ""),
        row.get("krone_no", ""),
        row.get("krone_o", ""),
        row.get("krone_so", ""),
        row.get("krone_s", ""),
        row.get("krone_sw", ""),
        row.get("krone_w", ""),
        row.get("krone_nw", ""),
        row.get("lichte_hoehe", ""),
        ansatzhoehe,
        row.get("ausrichtung_hauptast", ""),
        row.get("stammumfang", ""),
        # Status — mapped to exact VW dropdown values
        _map_action(row.get("action", "")),
        row.get("veteranenbaum", ""),
        row.get("alter_baum", ""),
        row.get("pflanzjahr", ""),
        row.get("letzte_ueberpruefung", ""),
        standort,
        row.get("bemerkungen", ""),
        row.get("kronenform", ""),
        row.get("struktur", ""),
        _map_vitalitaet(row.get("vitalitaet", "")),
        row.get("stammfuss_durchmesser", ""),
        # Custom fields — PDF extras that VW has no native field for
        row.get("erhaltung", ""),           # Feld 1: Erhaltungswürdigkeit
        row.get("verkehrssicherheit", ""),   # Feld 2: Verkehrssicherheit
        row.get("schutzstatus", ""),         # Feld 3
        row.get("anzahl_staemme", ""),       # Feld 4
        row.get("ersatzpflanzungen", ""),    # Feld 5
        row.get("feld6", ""),
        row.get("zusatz07", ""),
        row.get("zusatz08", ""),
        row.get("zusatz09", ""),
        row.get("zusatz10", ""),
    ]


def _export_vw_txt(all_rows: list[list]) -> str:
    """Given a list of row-lists (each matching VW_COLUMNS), drop empty columns and build TXT."""
    # Find which columns have at least one non-empty value
    n_cols = len(VW_COLUMNS)
    has_data = [False] * n_cols
    for row in all_rows:
        for i in range(n_cols):
            if _fmt(row[i]):
                has_data[i] = True

    # Always keep x/y coordinates
    has_data[0] = True  # x-Koordinate
    has_data[1] = True  # y-Koordinate

    keep = [i for i in range(n_cols) if has_data[i]]

    header = "\t".join(VW_COLUMNS[i] for i in keep)
    lines = [header]
    for row in all_rows:
        lines.append("\t".join(_fmt(row[i]) for i in keep))

    return "\n".join(lines)


def _export_vw_xlsx(all_rows: list[list]) -> bytes:
    """Given a list of row-lists (each matching VW_COLUMNS), drop empty columns and build XLSX."""
    import openpyxl

    n_cols = len(VW_COLUMNS)
    has_data = [False] * n_cols
    for row in all_rows:
        for i in range(n_cols):
            if _fmt(row[i]):
                has_data[i] = True

    has_data[0] = True  # x-Koordinate
    has_data[1] = True  # y-Koordinate

    keep = [i for i in range(n_cols) if has_data[i]]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Baumkataster"

    # Header row
    for col_idx, i in enumerate(keep, 1):
        ws.cell(row=1, column=col_idx, value=VW_COLUMNS[i])

    # Data rows — try to write numbers as numbers for VW
    for row_idx, row in enumerate(all_rows, 2):
        for col_idx, i in enumerate(keep, 1):
            val = _fmt(row[i])
            if val:
                try:
                    val = float(val)
                except ValueError:
                    pass
            else:
                val = ""
            ws.cell(row=row_idx, column=col_idx, value=val)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _build_wfs_rows(trees_gdf, output_crs_key, ansatz_method="none", ansatz_ratio=0.25):
    """Build VW row data from WFS/REST GeoDataFrame."""
    transformer = Transformer.from_crs("EPSG:4326", resolve_crs(output_crs_key), always_xy=True)
    all_rows = []
    for _, row in trees_gdf.iterrows():
        x, y = transformer.transform(row.geometry.x, row.geometry.y)
        hoehe = row.get("baumhoehe", "")
        kd = row.get("kronendurchmesser", "")
        if ansatz_method == "none":
            ansatzhoehe = row.get("ansatzhoehe", "")
        else:
            ansatzhoehe = estimate_ansatzhoehe(hoehe, kd, method=ansatz_method, ratio=ansatz_ratio)
        strasse = row.get("strasse", "")
        hausnr = row.get("hausnummer", "")
        standort = f"{strasse} {hausnr}".strip() if strasse else ""
        all_rows.append(_build_vw_row(row, x, y, ansatzhoehe=ansatzhoehe, standort=standort))
    return all_rows


def _build_pdf_rows(trees_gdf, output_crs_key, ansatz_method="none", ansatz_ratio=0.25):
    """Build VW row data from PDF-merged GeoDataFrame."""
    transformer = Transformer.from_crs("EPSG:4326", resolve_crs(output_crs_key), always_xy=True)
    all_rows = []
    for _, row in trees_gdf.iterrows():
        x, y = transformer.transform(row.geometry.x, row.geometry.y)
        ansatzhoehe = row.get("ansatzhoehe", "")
        if not _fmt(ansatzhoehe) and ansatz_method != "none":
            hoehe = row.get("baumhoehe", "")
            kd = row.get("kronendurchmesser", "")
            ansatzhoehe = estimate_ansatzhoehe(hoehe, kd, method=ansatz_method, ratio=ansatz_ratio)
        standort = row.get("standort", "")
        all_rows.append(_build_vw_row(row, x, y, ansatzhoehe=ansatzhoehe, standort=standort))
    return all_rows


# --- TXT exports ---

def trees_to_vw_txt(trees_gdf, output_crs_key, ansatz_method="none", ansatz_ratio=0.25):
    """WFS/REST pipeline → VW import TXT."""
    return _export_vw_txt(_build_wfs_rows(trees_gdf, output_crs_key, ansatz_method, ansatz_ratio))


def pdf_trees_to_vw_txt(trees_gdf, output_crs_key, ansatz_method="none", ansatz_ratio=0.25):
    """PDF pipeline → VW import TXT."""
    return _export_vw_txt(_build_pdf_rows(trees_gdf, output_crs_key, ansatz_method, ansatz_ratio))


# --- XLSX exports ---

def trees_to_vw_xlsx(trees_gdf, output_crs_key, ansatz_method="none", ansatz_ratio=0.25):
    """WFS/REST pipeline → VW import XLSX."""
    return _export_vw_xlsx(_build_wfs_rows(trees_gdf, output_crs_key, ansatz_method, ansatz_ratio))


def pdf_trees_to_vw_xlsx(trees_gdf, output_crs_key, ansatz_method="none", ansatz_ratio=0.25):
    """PDF pipeline → VW import XLSX."""
    return _export_vw_xlsx(_build_pdf_rows(trees_gdf, output_crs_key, ansatz_method, ansatz_ratio))
