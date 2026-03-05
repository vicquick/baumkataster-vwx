"""Parse tree assessment PDFs (Baumgutachten) to extract per-tree data.

Supports two formats:
1. Narrative: per-tree blocks with "Baum Nr. XX – Latin – German" headers
2. Tabular: spreadsheet-style tables with one row per tree
"""

import re

import pdfplumber


def parse_tree_pdf(pdf_file) -> list[dict]:
    """Auto-detect format and parse. Returns list of dicts with normalized fields."""
    with pdfplumber.open(pdf_file) as pdf:
        # Detect format: check if first data pages have tables
        has_tables = False
        full_text = ""
        all_tables = []

        for page in pdf.pages:
            text = page.extract_text() or ""
            full_text += text + "\n"
            tables = page.extract_tables()
            if tables:
                for t in tables:
                    if len(t) > 1 and len(t[0]) >= 5:
                        has_tables = True
                        all_tables.append(t)

        # Decide format
        has_narrative = bool(re.search(r'\d+\.\d+\s+Baum\s+Nr\.', full_text))

        if has_narrative:
            return _parse_narrative(full_text)
        elif has_tables:
            return _parse_tabular(all_tables)
        else:
            return []


# ==========================================================================
# FORMAT 1: Narrative (e.g., Zemke Baumgutachten)
# ==========================================================================

def _parse_narrative(text: str) -> list[dict]:
    """Parse narrative-style Baumgutachten with per-tree blocks."""
    pattern = r'(?=\d+\.\d+\s+Baum\s+Nr\.\s*\S+)'
    blocks = re.split(pattern, text)

    trees = []
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        tree = _parse_narrative_block(block)
        if tree and tree.get("baum_id"):
            trees.append(tree)
    return trees


def _parse_narrative_block(block: str) -> dict | None:
    """Parse a single narrative tree block."""
    header_match = re.match(
        r'\d+\.\d+\s+Baum\s+Nr\.\s*(\S+)\s*[\u2013\-–]\s*(.+?)[\u2013\-–]\s*(.+)',
        block
    )
    if not header_match:
        return None

    baum_id = header_match.group(1).strip()
    art_latein = header_match.group(2).strip()
    art_deutsch = header_match.group(3).split("\n")[0].strip()

    tree = {
        "baum_id": baum_id,
        "art_latein": art_latein,
        "art_deutsch": art_deutsch,
    }

    # Stammumfang: XXX cm
    m = re.search(r'Stammumfang:\s*([\d,.]+)\s*cm', block)
    if m:
        tree["stammumfang"] = _parse_german_float(m.group(1))

    # Kronendurchmesser: X,X m
    m = re.search(r'Kronendurchmesser:\s*([\d,.]+)\s*m', block)
    if m:
        tree["kronendurchmesser"] = _parse_german_float(m.group(1))

    # Höhe: X,X m
    m = re.search(r'Höhe:\s*([\d,.]+)\s*m', block)
    if m:
        tree["baumhoehe"] = _parse_german_float(m.group(1))

    # Kronenansatz: in X,X m Höhe
    m = re.search(r'Kronenansatz:\s*in\s*([\d,.]+)\s*m', block)
    if m:
        tree["ansatzhoehe"] = _parse_german_float(m.group(1))

    # Kronenform
    m = re.search(r'Kronenform:\s*(.+?)(?:\n|Vitalität)', block, re.DOTALL)
    if m:
        tree["kronenform"] = _clean_multiline(m.group(1))

    # Vitalität: X or X-Y
    m = re.search(r'Vitalität:\s*([\d\-–]+)', block)
    if m:
        tree["vitalitaet"] = m.group(1).replace("\u2013", "-").strip()

    # Verkehrssicherheit
    m = re.search(r'Verkehrssicherheit:\s*(.+?)(?:\n)', block)
    if m:
        tree["verkehrssicherheit"] = m.group(1).strip()

    # Bemerkungen
    m = re.search(r'Bemerkungen:\s*(.+?)(?:Erhaltungswürdigkeit:|$)', block, re.DOTALL)
    if m:
        tree["bemerkungen"] = _clean_multiline(m.group(1))

    # Erhaltungswürdigkeit
    m = re.search(r'Erhaltungswürdigkeit:\s*(.+?)(?:\n|Ansicht|$)', block)
    if m:
        tree["erhaltung"] = m.group(1).strip()

    return tree


# ==========================================================================
# FORMAT 2: Tabular (e.g., Schnelsen Datenabgleich)
# ==========================================================================

# Known header patterns mapped to normalized field names
_TABULAR_FIELD_PATTERNS = {
    r"B(?:aum|Baauumm)\s*N(?:r|Nrr)": "baum_id",
    r"Deutscher\s*Name": "art_deutsch",
    r"Botanischer\s*Name": "art_latein",
    r"Stamm-?\s*umfang": "stammumfang",
    r"Kronen\s*durch\s*messer.*geschätzt": "kronendurchmesser_est",
    r"Kronen\s*durch\s*messer.*Vermes": "kronendurchmesser",
    r"Baum\s*höhe": "baumhoehe",
    r"Vitalität": "vitalitaet",
    r"Anmerkungen|Mängel|Defekt": "bemerkungen",
    r"Erhaltens": "erhaltung",
    r"Schutzstatus": "schutzstatus",
    r"Anzahl\s*Stämm": "anzahl_staemme",
    r"Ersatzpflan": "ersatzpflanzungen",
}


def _parse_tabular(all_tables: list) -> list[dict]:
    """Parse tabular PDF tables into tree records."""
    trees = []

    for table in all_tables:
        if len(table) < 2:
            continue

        # Map column indices to normalized field names
        header_row = table[0]
        col_map = _map_columns(header_row)

        if "baum_id" not in col_map.values():
            continue

        for row in table[1:]:
            tree = _parse_tabular_row(row, col_map)
            if tree and tree.get("baum_id"):
                trees.append(tree)

    return trees


def _map_columns(header_row: list) -> dict[int, str]:
    """Map column indices to normalized field names based on header text."""
    col_map = {}
    for i, cell in enumerate(header_row):
        if not cell:
            continue
        # Normalize cell text for matching
        cell_clean = re.sub(r'\s+', ' ', cell.strip())
        for pattern, field_name in _TABULAR_FIELD_PATTERNS.items():
            if re.search(pattern, cell_clean, re.IGNORECASE):
                # Don't overwrite if already mapped (first match wins)
                if field_name not in col_map.values():
                    col_map[i] = field_name
                break
    return col_map


def _parse_tabular_row(row: list, col_map: dict[int, str]) -> dict | None:
    """Parse a single table row into a tree dict."""
    tree = {}
    for i, field_name in col_map.items():
        if i >= len(row):
            continue
        val = (row[i] or "").strip()
        # Clean up multiline cell values
        val = re.sub(r'\s+', ' ', val)

        if not val:
            continue

        if field_name in ("stammumfang", "baumhoehe"):
            tree[field_name] = _extract_number(val)
        elif field_name in ("kronendurchmesser", "kronendurchmesser_est"):
            parsed = _extract_number(val)
            # Prefer surveyed (Vermesser) over estimated
            if field_name == "kronendurchmesser" or "kronendurchmesser" not in tree:
                tree["kronendurchmesser"] = parsed
        elif field_name == "baum_id":
            tree["baum_id"] = val
        elif field_name == "art_latein":
            tree["art_latein"] = val
        elif field_name == "art_deutsch":
            tree["art_deutsch"] = val
        else:
            tree[field_name] = val

    # Skip empty/summary rows
    if not tree.get("baum_id") or tree.get("baum_id", "").lower() in ("summe", ""):
        return None

    return tree


# ==========================================================================
# Helpers
# ==========================================================================

def _parse_german_float(s: str) -> str:
    """Convert German decimal (comma) to value string."""
    if not s:
        return ""
    s = s.strip().replace(",", ".")
    try:
        return str(float(s))
    except ValueError:
        return s


def _extract_number(s: str) -> str:
    """Extract the first number from a string like '15 cm' or '8m'."""
    if not s:
        return ""
    m = re.search(r'([\d,.]+)', s)
    if m:
        return _parse_german_float(m.group(1))
    return ""


def _clean_multiline(text: str) -> str:
    """Clean multi-line text: remove dashes, normalize whitespace."""
    text = re.sub(r'\s*[\u2500─]+\s*', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()
