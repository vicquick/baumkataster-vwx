import io
import tempfile
import zipfile
from pathlib import Path

import folium
import geopandas as gpd
import pandas as pd
import streamlit as st
from folium.plugins import Draw
from shapely.geometry import shape
from streamlit_folium import st_folium

from export import (
    CRS_OPTIONS, VW_COLUMNS, resolve_crs,
    trees_to_vw_txt, pdf_trees_to_vw_txt,
    trees_to_vw_xlsx, pdf_trees_to_vw_xlsx,
    generate_fixup_script,
)
from fetcher import (
    discover_fields,
    discover_typenames,
    fetch_custom_wfs,
    fetch_trees,
)
from pdf_parser import parse_tree_pdf
from presets import PRESETS

st.set_page_config(page_title="Baumkataster Tool", layout="wide")
st.title("Baumkataster — VectorWorks Import Generator")

# --- Sidebar: CRS + Export ---
st.sidebar.header("CRS Settings")
crs_labels = [f"{code} — {name}" for code, name in CRS_OPTIONS.items()]

input_crs_label = st.sidebar.selectbox(
    "Input SHP CRS (override auto-detected)",
    ["Auto-detect from .prj"] + crs_labels,
    index=0,
)
input_crs_override = None
if input_crs_label != "Auto-detect from .prj":
    input_crs_override = input_crs_label.split(" — ")[0]

output_crs_label = st.sidebar.selectbox(
    "Output CRS (VW Import coordinates)",
    crs_labels,
    index=1,
)
output_crs = output_crs_label.split(" — ")[0]

st.sidebar.header("VW Export Values")
vw_lang = st.sidebar.radio(
    "Enum values (Maßnahme, Vitalität)",
    ["en", "de"],
    format_func=lambda x: "English (Retain/Remove/Good/Poor)" if x == "en" else "Deutsch (Sichern/Entfernen/0-3)",
    index=0,
    help="Column headers are always German for VW auto-mapping. "
         "VW may match enum values in English internally (Retain, Remove, Excellent, Good, Average, Poor).",
)
st.sidebar.caption(
    "**Note:** VW 2026 has known bugs with Action/Maßnahme import "
    "([VB-214920](https://forum.vectorworks.net), [VB-216372](https://forum.vectorworks.net)). "
    "Values may import as 'Custom/Eigen' despite correct format. "
    "Workaround: set Action manually in VW after import."
)

# ============================================================================
# TOP-LEVEL TABS
# ============================================================================
mode_wfs, mode_pdf, mode_join, mode_table = st.tabs(["City WFS / REST", "PDF Baumgutachten", "Label → Point Join", "Table → VW Import"])

# ============================================================================
# MODE 1: City WFS / REST
# ============================================================================
with mode_wfs:
    st.sidebar.header("WFS Settings")

    city_options = list(PRESETS.keys()) + ["Custom WFS"]
    selected_city = st.sidebar.selectbox("Select city", city_options)

    is_custom = selected_city == "Custom WFS"
    preset = PRESETS.get(selected_city)

    custom_config = {}
    if is_custom:
        custom_url = st.sidebar.text_input("WFS URL")
        if custom_url:
            with st.sidebar.expander("Discover & Configure", expanded=True):
                if st.button("Discover TypeNames"):
                    try:
                        names = discover_typenames(custom_url)
                        st.session_state["custom_typenames"] = names
                    except Exception as e:
                        st.error(f"GetCapabilities failed: {e}")

                typenames = st.session_state.get("custom_typenames", [])
                if typenames:
                    chosen_type = st.selectbox("TypeName", typenames)
                else:
                    chosen_type = st.text_input("TypeName (manual)")

                typename_param = st.selectbox("typename param", ["typeNames", "typeName"])
                output_format = st.selectbox("outputFormat", [
                    "application/json",
                    "application/geo+json",
                    "application/json; subtype=geojson",
                ])
                native_crs = st.selectbox("Native CRS", ["EPSG:25832", "EPSG:25833", "EPSG:4326"])

                if chosen_type and st.button("Discover Fields"):
                    try:
                        fields = discover_fields(custom_url, chosen_type, typename_param, output_format)
                        st.session_state["custom_fields"] = fields
                    except Exception as e:
                        st.error(f"Field discovery failed: {e}")

                available_fields = st.session_state.get("custom_fields", [])
                if available_fields:
                    st.markdown("**Map fields** (leave blank to skip):")
                    normalized = [
                        "baum_id", "art_deutsch", "gattung_deutsch", "art_latein",
                        "gattung_latein", "stammumfang", "kronendurchmesser",
                        "baumhoehe", "pflanzjahr", "strasse", "hausnummer", "bezirk",
                    ]
                    field_map = {}
                    for norm in normalized:
                        choice = st.selectbox(
                            norm, ["(none)"] + available_fields,
                            key=f"custom_field_{norm}",
                        )
                        if choice != "(none)":
                            field_map[choice] = norm

                    custom_config = {
                        "url": custom_url,
                        "type_name": chosen_type,
                        "typename_param": typename_param,
                        "output_format": output_format,
                        "native_crs": native_crs,
                        "field_map": field_map,
                    }

    max_features = st.sidebar.number_input("Max features to fetch", 100, 50000, 5000)

    st.sidebar.header("Ansatzhöhe (Kronenansatz)")
    ansatz_method = st.sidebar.selectbox(
        "Estimation method",
        ["none", "ratio", "kd"],
        format_func=lambda x: {
            "none": "Leave empty",
            "ratio": "Höhe × Ratio (recommended)",
            "kd": "Höhe − Kronendurchmesser",
        }[x],
        index=1,
    )
    ansatz_ratio = 0.25
    if ansatz_method == "ratio":
        ansatz_ratio = st.sidebar.slider("Ratio (Ansatzhöhe / Höhe)", 0.10, 0.50, 0.25, 0.05)


    # --- Define Area ---
    st.header("1. Define Area")
    tab_draw, tab_shp = st.tabs(["Draw on Map", "Upload Shapefile"])

    boundary_gdf = None

    with tab_draw:
        if preset:
            center = preset["center"]
        elif is_custom:
            center = [51.5, 10.0]
        else:
            center = [53.55, 10.0]

        m = folium.Map(location=center, zoom_start=13)
        Draw(
            draw_options={
                "polyline": False, "circlemarker": False, "marker": False,
                "circle": False, "polygon": True, "rectangle": True,
            },
            edit_options={"edit": False},
        ).add_to(m)

        output = st_folium(m, width=700, height=500, key="draw_map")

        if output and output.get("all_drawings"):
            drawings = output["all_drawings"]
            if drawings:
                last_drawing = drawings[-1]
                drawn_geom = shape(last_drawing["geometry"])
                boundary_gdf = gpd.GeoDataFrame(geometry=[drawn_geom], crs="EPSG:4326")
                st.success(f"Area defined: {drawn_geom.geom_type} with {len(drawings)} drawing(s)")

    with tab_shp:
        uploaded_files = st.file_uploader(
            "Upload shapefile components (.shp, .shx, .dbf, .prj) or a .zip",
            type=["shp", "shx", "dbf", "prj", "cpg", "zip"],
            accept_multiple_files=True,
            key="wfs_shp_upload",
        )

        if uploaded_files:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmppath = Path(tmpdir)
                for uf in uploaded_files:
                    if uf.name.endswith(".zip"):
                        with zipfile.ZipFile(io.BytesIO(uf.read())) as zf:
                            zf.extractall(tmppath)
                    else:
                        (tmppath / uf.name).write_bytes(uf.read())

                shp_files = list(tmppath.glob("**/*.shp"))
                if shp_files:
                    shp_gdf = gpd.read_file(shp_files[0])
                    detected_crs = shp_gdf.crs
                    if input_crs_override:
                        shp_gdf = shp_gdf.set_crs(resolve_crs(input_crs_override), allow_override=True)
                        st.info(f"CRS overridden: {detected_crs} → {input_crs_override}")
                    boundary_gdf = shp_gdf.to_crs("EPSG:4326")
                    st.success(
                        f"Loaded: {shp_files[0].name} — {len(boundary_gdf)} feature(s), "
                        f"CRS: {detected_crs}"
                    )

                    centroid = boundary_gdf.union_all().centroid
                    m_shp = folium.Map(location=[centroid.y, centroid.x], zoom_start=16)
                    folium.GeoJson(
                        boundary_gdf.__geo_interface__,
                        style_function=lambda x: {
                            "fillColor": "blue", "color": "blue",
                            "weight": 2, "fillOpacity": 0.1,
                        },
                    ).add_to(m_shp)
                    st_folium(m_shp, width=700, height=400, key="shp_map")
                else:
                    st.error("No .shp file found in uploaded files.")

    # --- Fetch & Filter Trees ---
    if boundary_gdf is not None:
        st.header("2. Fetch Trees")
        city_label = selected_city if not is_custom else "Custom WFS"
        st.caption(f"Source: **{city_label}**")

        if st.button("Fetch trees within boundary"):
            with st.spinner("Fetching trees..."):
                bounds = boundary_gdf.total_bounds
                bbox_4326 = tuple(bounds)

                try:
                    if is_custom and custom_config:
                        trees_gdf = fetch_custom_wfs(
                            url=custom_config["url"],
                            type_name=custom_config["type_name"],
                            field_map=custom_config["field_map"],
                            bbox_4326=bbox_4326,
                            max_features=max_features,
                            typename_param=custom_config["typename_param"],
                            output_format=custom_config["output_format"],
                            native_crs=custom_config["native_crs"],
                        )
                    elif preset:
                        trees_gdf = fetch_trees(preset, bbox_4326, max_features)
                    else:
                        st.error("No valid data source configured.")
                        st.stop()

                    if trees_gdf.empty:
                        st.warning("No trees found in this area.")
                    else:
                        total_bbox = len(trees_gdf)
                        st.info(f"Fetched {total_bbox} trees in bounding box")

                        boundary_union = boundary_gdf.union_all()
                        mask = trees_gdf.within(boundary_union)
                        filtered = trees_gdf[mask].copy()

                        st.success(f"**{len(filtered)} trees** within the boundary polygon")

                        if not filtered.empty:
                            st.session_state["filtered_trees"] = filtered
                            st.session_state["boundary_gdf"] = boundary_gdf

                except Exception as e:
                    st.error(f"Fetch failed: {e}")

        # --- Display & Export ---
        if "filtered_trees" in st.session_state:
            filtered = st.session_state["filtered_trees"]
            boundary_for_map = st.session_state.get("boundary_gdf", boundary_gdf)

            st.header("3. Results")

            display_cols = [
                "baum_id", "art_deutsch", "gattung_deutsch", "art_latein",
                "stammumfang", "kronendurchmesser", "baumhoehe",
                "pflanzjahr", "strasse", "hausnummer", "bezirk",
            ]
            existing_cols = [c for c in display_cols if c in filtered.columns]
            st.dataframe(filtered[existing_cols], use_container_width=True)

            centroid = boundary_for_map.union_all().centroid
            m2 = folium.Map(location=[centroid.y, centroid.x], zoom_start=17)
            folium.GeoJson(
                boundary_for_map.__geo_interface__,
                style_function=lambda x: {
                    "fillColor": "blue", "color": "blue",
                    "weight": 2, "fillOpacity": 0.1,
                },
            ).add_to(m2)

            for _, tree in filtered.iterrows():
                popup_text = (
                    f"<b>{tree.get('baum_id', '')}</b><br>"
                    f"{tree.get('art_deutsch', '')} ({tree.get('art_latein', '')})<br>"
                    f"StU: {tree.get('stammumfang', '')} cm, "
                    f"KD: {tree.get('kronendurchmesser', '')} m<br>"
                    f"Höhe: {tree.get('baumhoehe', '')}<br>"
                    f"Pflanzjahr: {tree.get('pflanzjahr', '')}"
                )
                folium.CircleMarker(
                    location=[tree.geometry.y, tree.geometry.x],
                    radius=5, color="green", fill=True,
                    fill_color="green", fill_opacity=0.7,
                    popup=folium.Popup(popup_text, max_width=250),
                ).add_to(m2)

            st_folium(m2, width=700, height=500, key="trees_map")

            st.header("4. Export VectorWorks Import TXT")
            st.caption(f"Output coordinates in: **{output_crs}** — {CRS_OPTIONS[output_crs]}")

            vw_txt = trees_to_vw_txt(filtered, output_crs,
                                      ansatz_method=ansatz_method,
                                      ansatz_ratio=ansatz_ratio,
                                      lang=vw_lang)
            st.text_area("Preview (first 10 lines)", "\n".join(vw_txt.split("\n")[:11]), height=300)

            dl_txt, dl_xlsx = st.columns(2)
            with dl_txt:
                st.download_button(
                    label="Download TXT",
                    data=vw_txt.encode("utf-8"),
                    file_name="Baumkataster_VW_Import.txt",
                    mime="text/plain",
                )
            with dl_xlsx:
                vw_xlsx = trees_to_vw_xlsx(filtered, output_crs,
                                           ansatz_method=ansatz_method,
                                           ansatz_ratio=ansatz_ratio,
                                           lang=vw_lang)
                st.download_button(
                    label="Download XLSX",
                    data=vw_xlsx,
                    file_name="Baumkataster_VW_Import.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


# ============================================================================
# MODE 2: PDF Baumgutachten
# ============================================================================
with mode_pdf:
    st.header("PDF Baumgutachten → VW Import")
    st.caption(
        "Upload a tree assessment PDF and a point shapefile. "
        "The PDF is parsed for tree data (Ansatzhöhe, Vitalität, Kronenform, etc.) "
        "and matched to SHP points by Baum-ID. Result: a complete VW import TXT."
    )

    col_pdf, col_shp = st.columns(2)

    with col_pdf:
        st.subheader("1. Upload PDF")
        pdf_file = st.file_uploader("Tree assessment PDF", type=["pdf"], key="pdf_upload")

        if pdf_file:
            with st.spinner("Parsing PDF..."):
                try:
                    pdf_trees = parse_tree_pdf(pdf_file)
                    st.session_state["pdf_trees"] = pdf_trees
                    st.success(f"Parsed **{len(pdf_trees)} trees** from PDF")

                    # Preview
                    pdf_df = pd.DataFrame(pdf_trees)
                    preview_cols = [c for c in [
                        "baum_id", "art_deutsch", "art_latein", "stammumfang",
                        "kronendurchmesser", "baumhoehe", "ansatzhoehe",
                        "kronenform", "vitalitaet", "erhaltung",
                    ] if c in pdf_df.columns]
                    st.dataframe(pdf_df[preview_cols], use_container_width=True, height=300)
                except Exception as e:
                    st.error(f"PDF parsing failed: {e}")

    with col_shp:
        st.subheader("2. Upload Point SHP")
        shp_files_pdf = st.file_uploader(
            "Point shapefile (.shp/.shx/.dbf/.prj or .zip)",
            type=["shp", "shx", "dbf", "prj", "cpg", "zip"],
            accept_multiple_files=True,
            key="pdf_shp_upload",
        )

        points_gdf = None
        if shp_files_pdf:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmppath = Path(tmpdir)
                for uf in shp_files_pdf:
                    if uf.name.endswith(".zip"):
                        with zipfile.ZipFile(io.BytesIO(uf.read())) as zf:
                            zf.extractall(tmppath)
                    else:
                        (tmppath / uf.name).write_bytes(uf.read())

                found_shp = list(tmppath.glob("**/*.shp"))
                if found_shp:
                    points_gdf = gpd.read_file(found_shp[0])
                    if input_crs_override:
                        points_gdf = points_gdf.set_crs(
                            resolve_crs(input_crs_override), allow_override=True
                        )
                    points_gdf = points_gdf.to_crs("EPSG:4326")
                    st.session_state["pdf_points_gdf"] = points_gdf
                    st.success(f"Loaded **{len(points_gdf)} points**, fields: {list(points_gdf.columns)}")
                else:
                    st.error("No .shp file found.")

    # --- Match & Export ---
    pdf_trees = st.session_state.get("pdf_trees")
    points_gdf = st.session_state.get("pdf_points_gdf")

    if pdf_trees and points_gdf is not None and not points_gdf.empty:
        st.header("3. Match & Export")

        # Select ID field from shapefile
        non_geom_cols = [c for c in points_gdf.columns if c != "geometry"]
        id_field = st.selectbox("SHP field containing Baum-ID", non_geom_cols)

        if id_field:
            # Show SHP IDs vs PDF IDs for comparison
            shp_ids = set(str(v) for v in points_gdf[id_field].dropna().unique())
            pdf_ids = set(t["baum_id"] for t in pdf_trees)

            matched = shp_ids & pdf_ids
            shp_only = shp_ids - pdf_ids
            pdf_only = pdf_ids - shp_ids

            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("Matched", len(matched))
            col_m2.metric("SHP only (no PDF)", len(shp_only))
            col_m3.metric("PDF only (no SHP)", len(pdf_only))

            if shp_only:
                with st.expander(f"SHP IDs not found in PDF ({len(shp_only)})"):
                    st.write(sorted(shp_only))
            if pdf_only:
                with st.expander(f"PDF IDs not found in SHP ({len(pdf_only)})"):
                    st.write(sorted(pdf_only, key=lambda x: (not x[0].isdigit(), x)))

            if matched:
                # Build merged GeoDataFrame: SHP geometry + PDF attributes
                pdf_dict = {t["baum_id"]: t for t in pdf_trees}

                # Columns to carry over from input SHP (e.g. action from Join tab)
                shp_carry_cols = [c for c in points_gdf.columns
                                  if c not in ("geometry", id_field)]

                rows = []
                for _, pt in points_gdf.iterrows():
                    sid = str(pt[id_field])
                    if sid in pdf_dict:
                        tree_data = dict(pdf_dict[sid])
                        tree_data["geometry"] = pt.geometry
                        # Carry over SHP fields not already in PDF data
                        for col in shp_carry_cols:
                            if col not in tree_data:
                                tree_data[col] = pt.get(col, "")
                        rows.append(tree_data)

                merged_gdf = gpd.GeoDataFrame(rows, crs="EPSG:4326")
                st.success(f"**{len(merged_gdf)} trees** matched and ready for export")

                # Preview table
                preview_cols = [c for c in [
                    "baum_id", "art_deutsch", "art_latein", "stammumfang",
                    "kronendurchmesser", "baumhoehe", "ansatzhoehe",
                    "kronenform", "vitalitaet", "erhaltung", "action", "bemerkungen",
                ] if c in merged_gdf.columns]
                st.dataframe(merged_gdf[preview_cols], use_container_width=True)

                # Map
                centroid = merged_gdf.union_all().centroid
                m3 = folium.Map(location=[centroid.y, centroid.x], zoom_start=17)
                for _, tree in merged_gdf.iterrows():
                    action_str = tree.get("action", "")
                    popup = (
                        f"<b>{tree.get('baum_id', '')}</b><br>"
                        f"{tree.get('art_deutsch', '')} ({tree.get('art_latein', '')})<br>"
                        f"H: {tree.get('baumhoehe', '')} m, "
                        f"Ansatz: {tree.get('ansatzhoehe', '')} m<br>"
                        f"Vit: {tree.get('vitalitaet', '')}, "
                        f"Erh: {tree.get('erhaltung', '')}"
                    )
                    if action_str and str(action_str) not in ("", "nan", "None"):
                        popup += f"<br>Maßnahme: {action_str}"
                    if action_str and str(action_str).startswith(("Entfernen", "Entnehmen")):
                        color = "red"
                    else:
                        color = "green"
                    folium.CircleMarker(
                        location=[tree.geometry.y, tree.geometry.x],
                        radius=5, color=color, fill=True,
                        fill_color=color, fill_opacity=0.7,
                        popup=folium.Popup(popup, max_width=300),
                    ).add_to(m3)
                st_folium(m3, width=700, height=500, key="pdf_trees_map")

                # Export
                st.header("4. Export VectorWorks Import")
                st.caption(f"Output CRS: **{output_crs}** — {CRS_OPTIONS[output_crs]}")

                vw_txt = pdf_trees_to_vw_txt(merged_gdf, output_crs,
                                             ansatz_method=ansatz_method,
                                             ansatz_ratio=ansatz_ratio,
                                             lang=vw_lang)
                st.text_area(
                    "Preview (first 10 lines)",
                    "\n".join(vw_txt.split("\n")[:11]),
                    height=300,
                    key="pdf_preview",
                )

                dl_txt, dl_xlsx, dl_script = st.columns(3)
                with dl_txt:
                    st.download_button(
                        label="Download TXT",
                        data=vw_txt.encode("utf-8"),
                        file_name="Baumkataster_VW_Import.txt",
                        mime="text/plain",
                        key="pdf_download",
                    )
                with dl_xlsx:
                    vw_xlsx = pdf_trees_to_vw_xlsx(merged_gdf, output_crs,
                                                   ansatz_method=ansatz_method,
                                                   ansatz_ratio=ansatz_ratio,
                                                   lang=vw_lang)
                    st.download_button(
                        label="Download XLSX",
                        data=vw_xlsx,
                        file_name="Baumkataster_VW_Import.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="pdf_xlsx_download",
                    )
                with dl_script:
                    fixup_script = generate_fixup_script(
                        merged_gdf, output_crs,
                        lang=vw_lang, ansatz_method=ansatz_method,
                        ansatz_ratio=ansatz_ratio)
                    if fixup_script:
                        st.download_button(
                            label="VW Fix-up Script (.py)",
                            data=fixup_script.encode("utf-8"),
                            file_name="fixup_trees.py",
                            mime="text/x-python",
                            key="pdf_script_download",
                            help="Run in VW Script Editor after import — sets ALL fields by coordinate matching",
                        )


# ============================================================================
# MODE 3: Label → Point Nearest Join
# ============================================================================
with mode_join:
    st.header("Label → Point Nearest Join")
    st.caption(
        "Upload **tree points** (SHP) and **labels** (DXF or SHP). "
        "Each point gets the ID of its nearest label text. "
        "Download the result as a SHP ready for the PDF pipeline."
    )

    col_pts, col_lbl = st.columns(2)

    def _load_shp(uploaders, key_prefix):
        """Load SHP from uploaded files."""
        if not uploaders:
            return None
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            for uf in uploaders:
                if uf.name.endswith(".zip"):
                    with zipfile.ZipFile(io.BytesIO(uf.read())) as zf:
                        zf.extractall(tmppath)
                else:
                    (tmppath / uf.name).write_bytes(uf.read())
            found = list(tmppath.glob("**/*.shp"))
            if found:
                gdf = gpd.read_file(found[0])
                if input_crs_override:
                    gdf = gdf.set_crs(resolve_crs(input_crs_override), allow_override=True)
                return gdf
        return None

    def _load_dxf_texts(dxf_file) -> gpd.GeoDataFrame:
        """Extract TEXT and MTEXT entities from a DXF as a point GeoDataFrame."""
        import ezdxf
        from shapely.geometry import Point as ShapelyPoint

        with tempfile.NamedTemporaryFile(suffix=".dxf", delete=False) as tmp:
            tmp.write(dxf_file.read())
            tmp_path = tmp.name

        doc = ezdxf.readfile(tmp_path)
        Path(tmp_path).unlink(missing_ok=True)

        records = []
        msp = doc.modelspace()
        for entity in msp:
            if entity.dxftype() == "TEXT":
                ins = entity.dxf.insert
                records.append({
                    "text": entity.dxf.text.strip(),
                    "layer": entity.dxf.layer,
                    "geometry": ShapelyPoint(ins.x, ins.y),
                })
            elif entity.dxftype() == "MTEXT":
                ins = entity.dxf.insert
                records.append({
                    "text": entity.text.strip(),
                    "layer": entity.dxf.layer,
                    "geometry": ShapelyPoint(ins.x, ins.y),
                })

        if not records:
            return gpd.GeoDataFrame()
        return gpd.GeoDataFrame(records)

    with col_pts:
        st.subheader("Tree Points")
        pts_files = st.file_uploader(
            "Baumstämme point SHP",
            type=["shp", "shx", "dbf", "prj", "cpg", "zip"],
            accept_multiple_files=True,
            key="join_pts_upload",
        )
        pts_gdf = _load_shp(pts_files, "pts")
        if pts_gdf is not None:
            st.success(f"**{len(pts_gdf)} points**, CRS: {pts_gdf.crs}")
            st.dataframe(pts_gdf.drop(columns="geometry").head(5), use_container_width=True)

    with col_lbl:
        st.subheader("Labels (DXF or SHP)")
        lbl_files = st.file_uploader(
            "Label text DXF or SHP",
            type=["dxf", "shp", "shx", "dbf", "prj", "cpg", "zip"],
            accept_multiple_files=True,
            key="join_lbl_upload",
        )

        lbl_gdf = None
        if lbl_files:
            # Check if any file is DXF
            dxf_files = [f for f in lbl_files if f.name.lower().endswith(".dxf")]
            if dxf_files:
                lbl_gdf = _load_dxf_texts(dxf_files[0])
                if lbl_gdf is not None and not lbl_gdf.empty:
                    # DXF has no CRS — assume same as points
                    if pts_gdf is not None and pts_gdf.crs:
                        lbl_gdf = lbl_gdf.set_crs(pts_gdf.crs)
                    st.success(f"**{len(lbl_gdf)} text entities** from DXF")
                    # Show unique layers for filtering
                    layers = lbl_gdf["layer"].unique().tolist()
                    if len(layers) > 1:
                        selected_layers = st.multiselect(
                            "Filter by DXF layer (optional)", layers, default=layers,
                        )
                        lbl_gdf = lbl_gdf[lbl_gdf["layer"].isin(selected_layers)]
                    st.dataframe(lbl_gdf.drop(columns="geometry").head(10), use_container_width=True)
                else:
                    st.warning("No TEXT/MTEXT entities found in DXF.")
            else:
                lbl_gdf = _load_shp(lbl_files, "lbl")
                if lbl_gdf is not None:
                    st.success(f"**{len(lbl_gdf)} labels**, CRS: {lbl_gdf.crs}")
                    st.dataframe(lbl_gdf.drop(columns="geometry").head(5), use_container_width=True)

    if pts_gdf is not None and lbl_gdf is not None and not lbl_gdf.empty:
        # Pick which label field contains the Baum-ID text
        lbl_cols = [c for c in lbl_gdf.columns if c != "geometry"]
        default_idx = lbl_cols.index("text") if "text" in lbl_cols else 0
        lbl_id_field = st.selectbox("Label field containing Baum-ID text", lbl_cols, index=default_idx, key="join_lbl_field")

        max_dist = st.number_input(
            "Max match distance (m)", value=10.0, min_value=0.1, max_value=100.0, step=1.0,
            help="Labels farther than this from any point are flagged as unmatched.",
        )

        # Layer → VW Action mapping (DXF only)
        layer_action_map = {}
        if "layer" in lbl_gdf.columns:
            # Exact VW Maßnahme dropdown values
            vw_actions = [
                "",
                "Sichern",
                "Sichern und Baumpflegemaßnahmen",
                "Umpflanzen - Am Standort",
                "Umpflanzen - Neuer Standort",
                "Entfernen",
                "Entnehmen - innerhalb 4 Wochen",
                "Entnehmen - im nächsten Pflegezyklus",
            ]
            default_map = {
                "LL-VEG-Baum-Nummern": "Sichern",
                "LL-VEG-Baum-Nummern_Roden": "Entfernen",
            }
            with st.expander("Layer → VW Action mapping (Maßnahme)", expanded=True):
                st.caption("Map DXF layers to VectorWorks **Existing Tree** Action field.")
                unique_layers = lbl_gdf["layer"].unique().tolist()
                for layer_name in unique_layers:
                    default = default_map.get(layer_name, "")
                    idx = vw_actions.index(default) if default in vw_actions else 0
                    chosen = st.selectbox(
                        f"`{layer_name}`", vw_actions, index=idx,
                        key=f"action_map_{layer_name}",
                    )
                    if chosen:
                        layer_action_map[layer_name] = chosen

        if st.button("Run Nearest Join", key="join_btn"):
            with st.spinner("Joining..."):
                # Project to a metric CRS for distance calculation
                # Use UTM zone based on centroid
                centroid = pts_gdf.to_crs("EPSG:4326").union_all().centroid
                utm_zone = int((centroid.x + 180) / 6) + 1
                hemisphere = "north" if centroid.y >= 0 else "south"
                utm_epsg = 32600 + utm_zone if hemisphere == "north" else 32700 + utm_zone
                metric_crs = f"EPSG:{utm_epsg}"

                pts_m = pts_gdf.to_crs(metric_crs).copy()
                lbl_m = lbl_gdf.to_crs(metric_crs).copy()

                has_layer = "layer" in lbl_m.columns

                # Two-round nearest assignment:
                # Round 1: Greedy 1-to-1 (each label used at most once)
                # Round 2: Leftover points get nearest label (allows reuse for mehrstämmige)

                pts_geom = pts_m.geometry.values
                lbl_geom = lbl_m.geometry.values

                # Build candidate pairs (pt_idx, lbl_idx, dist)
                candidates = []
                for i, pg in enumerate(pts_geom):
                    for j, lg in enumerate(lbl_geom):
                        d = pg.distance(lg)
                        if d <= max_dist:
                            candidates.append((i, j, d))

                # Sort by distance (closest first)
                candidates.sort(key=lambda c: c[2])

                # Round 1: greedy 1-to-1
                used_pts = set()
                used_lbls = set()
                assignments = {}  # pt_idx -> (lbl_idx, dist)

                for pt_i, lbl_j, dist in candidates:
                    if pt_i not in used_pts and lbl_j not in used_lbls:
                        assignments[pt_i] = (lbl_j, dist)
                        used_pts.add(pt_i)
                        used_lbls.add(lbl_j)

                # Round 2: unmatched points get nearest label (reuse allowed)
                unmatched_pts = set(range(len(pts_geom))) - used_pts
                for pt_i, lbl_j, dist in candidates:
                    if pt_i in unmatched_pts:
                        assignments[pt_i] = (lbl_j, dist)
                        unmatched_pts.discard(pt_i)

                # Build result
                result = pts_gdf.copy()
                baum_ids = [None] * len(result)
                match_dists = [None] * len(result)
                actions = [""] * len(result)

                for pt_i, (lbl_j, dist) in assignments.items():
                    baum_ids[pt_i] = lbl_m.iloc[lbl_j][lbl_id_field]
                    match_dists[pt_i] = dist
                    if has_layer and layer_action_map:
                        layer_val = lbl_m.iloc[lbl_j].get("layer", "")
                        actions[pt_i] = layer_action_map.get(layer_val, "")

                result["baum_id"] = baum_ids
                result["match_dist_m"] = match_dists
                if has_layer and layer_action_map:
                    result["action"] = actions

                matched = result["baum_id"].notna().sum()
                unmatched = result["baum_id"].isna().sum()
                dupes = result["baum_id"].dropna().duplicated().sum()

                n_metric = 5 if "action" in result.columns else 3
                metric_cols = st.columns(n_metric)
                metric_cols[0].metric("Matched", int(matched))
                metric_cols[1].metric("Unmatched", int(unmatched))
                metric_cols[2].metric("Mehrstämmig", int(dupes),
                                      help="Points sharing a label (round 2)")
                if "action" in result.columns:
                    retain_n = result["action"].astype(str).str.startswith("Sichern").sum()
                    remove_n = result["action"].isin(["Entfernen", "Entnehmen - innerhalb 4 Wochen", "Entnehmen - im nächsten Pflegezyklus"]).sum()
                    metric_cols[3].metric("Sichern", int(retain_n))
                    metric_cols[4].metric("Entfernen", int(remove_n))

                st.session_state["join_result"] = result

        if "join_result" in st.session_state:
            result = st.session_state["join_result"]

            st.dataframe(
                result.drop(columns="geometry"),
                use_container_width=True,
            )

            # Map
            result_4326 = result.to_crs("EPSG:4326")
            centroid = result_4326.union_all().centroid
            m_join = folium.Map(location=[centroid.y, centroid.x], zoom_start=17)
            for _, row in result_4326.iterrows():
                bid = row.get("baum_id", "")
                action = row.get("action", "")
                has_action = pd.notna(action) and action
                if has_action and action.startswith(("Entfernen", "Entnehmen")):
                    color = "red"
                elif pd.notna(bid) and bid:
                    color = "green"
                else:
                    color = "gray"
                label = str(bid) if pd.notna(bid) else "?"
                dist = row.get("match_dist_m", 0)
                popup_text = f"<b>{label}</b><br>dist: {dist:.1f} m"
                if has_action:
                    popup_text += f"<br>Action: {action}"
                folium.CircleMarker(
                    location=[row.geometry.y, row.geometry.x],
                    radius=5, color=color, fill=True,
                    fill_color=color, fill_opacity=0.7,
                    popup=folium.Popup(popup_text, max_width=200),
                ).add_to(m_join)
            st_folium(m_join, width=700, height=500, key="join_map")

            # Download as SHP (zip)
            zip_bytes = None
            with tempfile.TemporaryDirectory() as tmpdir:
                out_path = Path(tmpdir) / "trees_with_ids.shp"
                export_gdf = result.to_crs(resolve_crs(output_crs))
                export_gdf.to_file(str(out_path), driver="ESRI Shapefile")
                # Zip all shapefile components
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for f in Path(tmpdir).glob("trees_with_ids.*"):
                        zf.write(str(f), f.name)
                zip_bytes = zip_buffer.getvalue()

            if zip_bytes:
                st.download_button(
                    label="Download SHP (Baum-IDs + Action)",
                    data=zip_bytes,
                    file_name="trees_with_ids.zip",
                    mime="application/zip",
                    key="join_download",
                )

            # Also offer VW TXT/XLSX export (with coordinates + action)
            result_4326_export = result.to_crs("EPSG:4326")
            dl_txt, dl_xlsx, dl_script = st.columns(3)
            with dl_txt:
                vw_txt = trees_to_vw_txt(result_4326_export, output_crs,
                                          ansatz_method=ansatz_method,
                                          ansatz_ratio=ansatz_ratio,
                                          lang=vw_lang)
                st.download_button(
                    label="Download VW TXT",
                    data=vw_txt.encode("utf-8"),
                    file_name="Baumkataster_VW_Import.txt",
                    mime="text/plain",
                    key="join_txt_download",
                )
            with dl_xlsx:
                vw_xlsx = trees_to_vw_xlsx(result_4326_export, output_crs,
                                            ansatz_method=ansatz_method,
                                            ansatz_ratio=ansatz_ratio,
                                            lang=vw_lang)
                st.download_button(
                    label="Download VW XLSX",
                    data=vw_xlsx,
                    file_name="Baumkataster_VW_Import.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="join_xlsx_download",
                )
            with dl_script:
                fixup_script = generate_fixup_script(
                    result_4326_export, output_crs,
                    lang=vw_lang, ansatz_method=ansatz_method,
                    ansatz_ratio=ansatz_ratio)
                if fixup_script:
                    st.download_button(
                        label="VW Fix-up Script (.py)",
                        data=fixup_script.encode("utf-8"),
                        file_name="fixup_trees.py",
                        mime="text/x-python",
                        key="join_script_download",
                        help="Run in VW Script Editor after import — sets ALL fields by coordinate matching",
                    )


# ============================================================================
# MODE 4: Table → VW Import (polygon / non-point SHP with attribute coords)
# ============================================================================
with mode_table:
    st.header("Table → VW Import")
    st.caption(
        "Upload a shapefile exported from VectorWorks (polygon circles, etc.) where "
        "the **attribute table** contains the coordinates and tree dimensions — "
        "geometry type doesn't matter, only the table data is used."
    )

    tbl_files = st.file_uploader(
        "Upload shapefile (.shp/.shx/.dbf/.prj or .zip)",
        type=["shp", "shx", "dbf", "prj", "cpg", "zip"],
        accept_multiple_files=True,
        key="tbl_shp_upload",
    )

    tbl_gdf = None
    if tbl_files:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            for uf in tbl_files:
                if uf.name.endswith(".zip"):
                    with zipfile.ZipFile(io.BytesIO(uf.read())) as zf:
                        zf.extractall(tmppath)
                else:
                    (tmppath / uf.name).write_bytes(uf.read())

            found_shp = list(tmppath.glob("**/*.shp"))
            if found_shp:
                tbl_gdf = gpd.read_file(found_shp[0])
                st.session_state["tbl_gdf"] = tbl_gdf
                st.success(
                    f"Loaded **{len(tbl_gdf)} features** — "
                    f"Geometry: {tbl_gdf.geom_type.iloc[0] if len(tbl_gdf) else '?'}, "
                    f"Fields: {[c for c in tbl_gdf.columns if c != 'geometry']}"
                )
            else:
                st.error("No .shp file found.")

    tbl_gdf = st.session_state.get("tbl_gdf")

    if tbl_gdf is not None and not tbl_gdf.empty:
        st.subheader("Attribute Table Preview")
        st.dataframe(
            tbl_gdf.drop(columns="geometry", errors="ignore").head(20),
            use_container_width=True,
        )

        non_geom = [c for c in tbl_gdf.columns if c != "geometry"]

        st.subheader("Field Mapping")
        col_map1, col_map2 = st.columns(2)

        with col_map1:
            x_field = st.selectbox("X coordinate field", non_geom, key="tbl_x_field")
            y_field = st.selectbox("Y coordinate field", non_geom,
                                   index=min(1, len(non_geom) - 1), key="tbl_y_field")
            coord_units = st.selectbox(
                "Coordinate units in table",
                ["m (Meters)", "mm (Millimeters)"],
                index=1,
                key="tbl_coord_units",
                help="VectorWorks often exports coordinates in mm when document is set to mm.",
            )
            # CRS uses sidebar Input/Output settings (same as other tabs)
            tbl_input_crs = input_crs_override if input_crs_override else "EPSG:31467"
            st.info(f"Input CRS: **{tbl_input_crs}** / Output CRS: **{output_crs}** — change in sidebar")

        with col_map2:
            # Optional field mappings for tree attributes
            none_opt = ["(none)"] + non_geom
            d_field = st.selectbox("Kronendurchmesser (D) field", none_opt, key="tbl_d_field")
            id_field = st.selectbox("Baum-ID field", none_opt, key="tbl_id_field")
            hoehe_field = st.selectbox("Baumhöhe field", none_opt, key="tbl_h_field")
            stu_field = st.selectbox("Stammumfang field", none_opt, key="tbl_stu_field")

        if st.button("Build VW Import", key="tbl_build_btn"):
            with st.spinner("Building..."):
                from shapely.geometry import Point as ShapelyPoint

                divisor = 1000.0 if "mm" in coord_units else 1.0

                rows_out = []
                for _, row in tbl_gdf.iterrows():
                    try:
                        rx = float(row[x_field]) / divisor
                        ry = float(row[y_field]) / divisor
                    except (TypeError, ValueError):
                        continue

                    tree_data = {
                        "geometry": ShapelyPoint(rx, ry),
                    }
                    if d_field != "(none)":
                        tree_data["kronendurchmesser"] = row.get(d_field, "")
                    if id_field != "(none)":
                        tree_data["baum_id"] = row.get(id_field, "")
                    if hoehe_field != "(none)":
                        tree_data["baumhoehe"] = row.get(hoehe_field, "")
                    if stu_field != "(none)":
                        tree_data["stammumfang"] = row.get(stu_field, "")

                    rows_out.append(tree_data)

                if not rows_out:
                    st.error("No valid rows — check your X/Y field mapping.")
                else:
                    result_gdf = gpd.GeoDataFrame(rows_out, crs=resolve_crs(tbl_input_crs))
                    result_4326 = result_gdf.to_crs("EPSG:4326")
                    st.session_state["tbl_result"] = result_4326
                    st.success(f"**{len(result_4326)} trees** ready for export")

        if "tbl_result" in st.session_state:
            result_4326 = st.session_state["tbl_result"]

            # Preview table
            preview = result_4326.drop(columns="geometry", errors="ignore")
            st.dataframe(preview, use_container_width=True)

            # Map
            centroid = result_4326.union_all().centroid
            m_tbl = folium.Map(location=[centroid.y, centroid.x], zoom_start=17)
            for _, tree in result_4326.iterrows():
                bid = tree.get("baum_id", "")
                kd = tree.get("kronendurchmesser", "")
                popup = f"<b>{bid}</b><br>KD: {kd}"
                folium.CircleMarker(
                    location=[tree.geometry.y, tree.geometry.x],
                    radius=5, color="green", fill=True,
                    fill_color="green", fill_opacity=0.7,
                    popup=folium.Popup(popup, max_width=200),
                ).add_to(m_tbl)
            st_folium(m_tbl, width=700, height=500, key="tbl_map")

            # Export
            st.header("Export VectorWorks Import")
            st.caption(f"Output CRS: **{output_crs}** — {CRS_OPTIONS[output_crs]}")

            dl_txt, dl_xlsx = st.columns(2)
            with dl_txt:
                vw_txt = trees_to_vw_txt(result_4326, output_crs,
                                          ansatz_method=ansatz_method,
                                          ansatz_ratio=ansatz_ratio,
                                          lang=vw_lang)
                st.download_button(
                    label="Download TXT",
                    data=vw_txt.encode("utf-8"),
                    file_name="Baumkataster_VW_Import.txt",
                    mime="text/plain",
                    key="tbl_txt_download",
                )
            with dl_xlsx:
                vw_xlsx = trees_to_vw_xlsx(result_4326, output_crs,
                                            ansatz_method=ansatz_method,
                                            ansatz_ratio=ansatz_ratio,
                                            lang=vw_lang)
                st.download_button(
                    label="Download XLSX",
                    data=vw_xlsx,
                    file_name="Baumkataster_VW_Import.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="tbl_xlsx_download",
                )
