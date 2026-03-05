"""Parse tree assessment PDFs (Baumgutachten) to extract per-tree data."""

import re

import pdfplumber


def parse_tree_pdf(pdf_file) -> list[dict]:
    """Parse a tree assessment PDF and extract structured data for each tree.

    Expects the common German Baumgutachten format:
        X.Y Baum Nr. ZZ – Lat. Name – Dt. Name
        Stammumfang: XXX cm    Kronendurchmesser: X,X m
        Höhe: X,X m            Kronenansatz: in X,X m Höhe
        Kronenform: ...
        Vitalität: X           Verkehrssicherheit: ...
        Bemerkungen: ...
        Erhaltungswürdigkeit: ...

    Returns list of dicts with normalized field names.
    """
    with pdfplumber.open(pdf_file) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text() or ""
            full_text += text + "\n"

    return _parse_tree_blocks(full_text)


def _parse_tree_blocks(text: str) -> list[dict]:
    """Split full text into per-tree blocks and parse each one."""
    # Pattern: "4.X Baum Nr. XX" or just "Baum Nr. XX"
    # Split on tree headers
    pattern = r'(?=\d+\.\d+\s+Baum\s+Nr\.\s*\S+)'
    blocks = re.split(pattern, text)

    trees = []
    for block in blocks:
        block = block.strip()
        if not block:
            continue

        tree = _parse_single_tree(block)
        if tree and tree.get("baum_id"):
            trees.append(tree)

    return trees


def _parse_german_float(s: str) -> str:
    """Convert German decimal (comma) to value string, or return empty."""
    if not s:
        return ""
    s = s.strip().replace(",", ".")
    try:
        return str(float(s))
    except ValueError:
        return s


def _parse_single_tree(block: str) -> dict | None:
    """Parse a single tree block into a dict."""
    # Header: "4.X Baum Nr. XX – Latin – German"
    header_match = re.match(
        r'\d+\.\d+\s+Baum\s+Nr\.\s*(\S+)\s*[\u2013\-–]\s*(.+?)[\u2013\-–]\s*(.+)',
        block
    )
    if not header_match:
        return None

    baum_id = header_match.group(1).strip()
    art_latein = header_match.group(2).strip()
    art_deutsch = header_match.group(3).strip()
    # Clean trailing newline junk from art_deutsch
    art_deutsch = art_deutsch.split("\n")[0].strip()

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

    # Kronenform: ...
    m = re.search(r'Kronenform:\s*(.+?)(?:\n|Vitalität)', block, re.DOTALL)
    if m:
        tree["kronenform"] = _clean_multiline(m.group(1))

    # Vitalität: X or X-Y
    m = re.search(r'Vitalität:\s*([\d\-–]+)', block)
    if m:
        tree["vitalitaet"] = m.group(1).replace("–", "-").strip()

    # Verkehrssicherheit
    m = re.search(r'Verkehrssicherheit:\s*(.+?)(?:\n)', block)
    if m:
        tree["verkehrssicherheit"] = m.group(1).strip()

    # Bemerkungen: everything between "Bemerkungen:" and "Erhaltungswürdigkeit:"
    m = re.search(r'Bemerkungen:\s*(.+?)(?:Erhaltungswürdigkeit:|$)', block, re.DOTALL)
    if m:
        tree["bemerkungen"] = _clean_multiline(m.group(1))

    # Erhaltungswürdigkeit
    m = re.search(r'Erhaltungswürdigkeit:\s*(.+?)(?:\n|Ansicht|$)', block)
    if m:
        tree["erhaltung"] = m.group(1).strip()

    return tree


def _clean_multiline(text: str) -> str:
    """Clean multi-line text: remove dashes, normalize whitespace."""
    # Remove isolated dashes (─) used as line continuation
    text = re.sub(r'\s*[─\u2500]+\s*', ' ', text)
    # Collapse whitespace
    text = re.sub(r'\s+', ' ', text)
    return text.strip()
