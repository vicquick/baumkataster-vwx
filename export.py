"""VectorWorks Import TXT generator."""

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


def trees_to_vw_txt(trees_gdf: gpd.GeoDataFrame, output_crs_key: str,
                    ansatz_method: str = "none", ansatz_ratio: float = 0.25,
                    include_extra_cols: bool = False) -> str:
    """Convert normalized trees GeoDataFrame (EPSG:4326) to VectorWorks import TXT.

    VW expects 12 tab-separated columns (by position):
    Baum-ID, Koordinaten, Deutscher Name, Botanischer Name,
    Stammumfang, Kronendurchmesser, Höhe, Ansatzhöhe,
    Form, Vitalitätsstufe, Erhaltungsstufe, Bemerkungen

    If include_extra_cols is True, adds Pflanzjahr and Standort (14 columns).
    """
    transformer = Transformer.from_crs("EPSG:4326", resolve_crs(output_crs_key), always_xy=True)

    columns = [
        "Baum-ID", "Koordinaten", "Deutscher Name", "Botanischer Name",
        "Stammumfang", "Kronendurchmesser", "Höhe", "Ansatzhöhe",
        "Form", "Vitalitätsstufe", "Erhaltungsstufe", "Bemerkungen",
    ]
    if include_extra_cols:
        # Insert before Bemerkungen
        columns = columns[:11] + ["Pflanzjahr", "Standort"] + columns[11:]

    header = "\t".join(columns)
    lines = [header]

    for _, row in trees_gdf.iterrows():
        x, y = transformer.transform(row.geometry.x, row.geometry.y)
        coord_str = f"[{x:.3f}, {y:.3f}]"

        baum_id = row.get("baum_id", "")
        deutsch = row.get("art_deutsch", "") or row.get("gattung_deutsch", "")
        latein = row.get("art_latein", "") or row.get("gattung_latein", "")
        stammumfang = row.get("stammumfang", "")
        kronendurchmesser = row.get("kronendurchmesser", "")
        hoehe = row.get("baumhoehe", "")

        # Ansatzhöhe
        if ansatz_method == "none":
            ansatzhoehe = ""
        else:
            ansatzhoehe = estimate_ansatzhoehe(hoehe, kronendurchmesser,
                                               method=ansatz_method, ratio=ansatz_ratio)

        form = ""
        vitalitaet = ""
        erhaltung = ""

        # Bemerkungen: include Pflanzjahr + Standort here so VW always gets them
        bemerkung_parts = []
        pflanzjahr = row.get("pflanzjahr", "")
        strasse = row.get("strasse", "")
        hausnr = row.get("hausnummer", "")
        standort = f"{strasse} {hausnr}".strip() if strasse else ""
        if pflanzjahr and str(pflanzjahr) not in ("", "0", "nan", "None"):
            bemerkung_parts.append(f"Pflanzjahr {pflanzjahr}")
        if standort:
            bemerkung_parts.append(standort)
        bemerkungen = ". ".join(str(p) for p in bemerkung_parts)

        fields = [
            baum_id, coord_str, deutsch, latein,
            stammumfang, kronendurchmesser, hoehe, ansatzhoehe,
            form, vitalitaet, erhaltung,
        ]

        if include_extra_cols:
            fields.extend([pflanzjahr, standort])

        fields.append(bemerkungen)

        def fmt(f):
            if f is None:
                return ""
            s = str(f)
            if s in ("nan", "None"):
                return ""
            return s

        lines.append("\t".join(fmt(f) for f in fields))

    return "\n".join(lines)


def pdf_trees_to_vw_txt(trees_gdf: gpd.GeoDataFrame, output_crs_key: str) -> str:
    """Convert PDF-parsed trees (with real Ansatzhöhe, Kronenform, etc.) to VW import TXT.

    Uses the standard 12-column VW format with actual survey data.
    """
    transformer = Transformer.from_crs("EPSG:4326", resolve_crs(output_crs_key), always_xy=True)

    header = "\t".join([
        "Baum-ID", "Koordinaten", "Deutscher Name", "Botanischer Name",
        "Stammumfang", "Kronendurchmesser", "Höhe", "Ansatzhöhe",
        "Form", "Vitalitätsstufe", "Erhaltungsstufe", "Bemerkungen",
    ])
    lines = [header]

    def fmt(f):
        if f is None:
            return ""
        s = str(f)
        return "" if s in ("nan", "None") else s

    for _, row in trees_gdf.iterrows():
        x, y = transformer.transform(row.geometry.x, row.geometry.y)
        coord_str = f"[{x:.3f}, {y:.3f}]"

        fields = [
            row.get("baum_id", ""),
            coord_str,
            row.get("art_deutsch", ""),
            row.get("art_latein", ""),
            row.get("stammumfang", ""),
            row.get("kronendurchmesser", ""),
            row.get("baumhoehe", ""),
            row.get("ansatzhoehe", ""),
            row.get("kronenform", ""),
            row.get("vitalitaet", ""),
            row.get("erhaltung", ""),
            row.get("bemerkungen", ""),
        ]
        lines.append("\t".join(fmt(f) for f in fields))

    return "\n".join(lines)
