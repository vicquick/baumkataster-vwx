# Baumkataster → Vectorworks Import Generator

Streamlit tool that fetches municipal tree-cadastre (Baumkataster) data from
public geodata services, normalizes it, and exports a **Vectorworks tree-survey
import** (TXT / XLSX) plus a VW Python fix-up script that batch-fills the
*Existing Tree* PIO fields.

```
streamlit run app.py        # -> http://localhost:8501
```

## Data sources (`presets.py`)

Each preset declares a `service_type` and a `field_map` (city field → normalized
field). Normalized fields: `baum_id, art_deutsch, gattung_deutsch, art_latein,
gattung_latein, stammumfang, kronendurchmesser, baumhoehe, pflanzjahr, strasse,
hausnummer, bezirk`.

| `service_type` | Source | Notes |
|----------------|--------|-------|
| `wfs`        | OGC WFS GetFeature | Hamburg, Berlin, Köln, Leipzig, Rostock |
| `arcgis_rest`| ArcGIS REST `/query` | Frankfurt |
| `3dtiles`    | Cesium 3D Tiles (i3dm/cmpt) | Hamburg Sommerbäume (height from ALS) |
| `overpass`   | **OpenStreetMap via Overpass** | any town without an official cadastre |

### OpenStreetMap / Overpass source

`fetcher._fetch_overpass()` pulls `natural=tree` nodes for the bbox and maps OSM
tags to normalized fields:

| OSM tag | normalized |
|---------|-----------|
| `species:de` / `species` | `art_deutsch` / `art_latein` |
| `genus:de` / `genus`     | `gattung_deutsch` / `gattung_latein` |
| `height` | `baumhoehe` |
| `circumference` | `stammumfang` |
| `diameter_crown` | `kronendurchmesser` |
| `start_date` | `pflanzjahr` |
| `ref` | `baum_id` |

Presets: `OpenStreetMap (Overpass) — beliebige Stadt` and
`Oldenburg in Holstein (OSM)`.

> **Coverage reality:** OSM tree tagging is sparse. Expect position + presence
> for most nodes, but `genus` on ~10 % and `species` near-zero in most German
> towns. Good for QA / cross-check, weak for attribute enrichment. For accurate
> height use a LiDAR canopy-height model (nDOM = DSM − DTM); for crown size use
> the canopy polygon area. See the Oldenburg in Holstein pipeline.

## Export (`export.py`)

- `trees_to_vw_txt` / `trees_to_vw_xlsx` — VW import table (German headers VW
  auto-maps; reprojects 4326 → chosen output CRS; `EPSG:4647` supported).
- `generate_fixup_script` — VW Python script; matches imported trees by
  coordinate proximity and sets all *Existing Tree* fields (works around VW
  import bugs).
- Missing dimensions can be estimated: height & trunk from crown diameter
  (`estimate_baumhoehe` / `estimate_stammumfang`), crown-base height from
  species-specific ratios (`SPECIES_ANSATZ_RATIOS`).

`species_ratios` passed to the export functions must be `{botanical_name: float}`
(the UI editor flattens `build_species_ratio_table`, which returns
`{name: (ratio, desc)}`).

## Modes (`app.py`)

WFS/REST/OSM fetch · 3D Tiles · PDF parser merge · attribute-table join ·
manual table entry. Fetched WFS records are de-duplicated by `baum_id`
(or coordinate) before export.
