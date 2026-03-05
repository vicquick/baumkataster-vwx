import io
import tempfile
import zipfile
from pathlib import Path

import folium
import geopandas as gpd
import streamlit as st
from folium.plugins import Draw
from shapely.geometry import shape
from streamlit_folium import st_folium

from export import CRS_OPTIONS, resolve_crs, trees_to_vw_txt
from fetcher import (
    discover_fields,
    discover_typenames,
    fetch_custom_wfs,
    fetch_trees,
)
from presets import PRESETS

st.set_page_config(page_title="Baumkataster Tool", layout="wide")
st.title("Baumkataster — VectorWorks Import Generator")

# --- Sidebar ---
st.sidebar.header("City / Data Source")

city_options = list(PRESETS.keys()) + ["Custom WFS"]
selected_city = st.sidebar.selectbox("Select city", city_options)

is_custom = selected_city == "Custom WFS"
preset = PRESETS.get(selected_city)

# Custom WFS configuration
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
                        norm,
                        ["(none)"] + available_fields,
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
    index=1,  # default to EPSG:31467
)
output_crs = output_crs_label.split(" — ")[0]

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

st.sidebar.header("Export Options")
include_extra_cols = st.sidebar.checkbox(
    "Extra columns (Pflanzjahr, Standort)", value=False,
    help="Adds 2 extra columns. Disable for standard VW 12-column import.",
)

# --- Main Area: Define Area ---
st.header("1. Define Area")

tab_draw, tab_shp = st.tabs(["Draw on Map", "Upload Shapefile"])

boundary_gdf = None

# --- Tab: Draw on Map ---
with tab_draw:
    # Determine map center from preset
    if preset:
        center = preset["center"]
    elif is_custom:
        center = [51.5, 10.0]  # Germany center
    else:
        center = [53.55, 10.0]

    m = folium.Map(location=center, zoom_start=13)
    Draw(
        draw_options={
            "polyline": False,
            "circlemarker": False,
            "marker": False,
            "circle": False,
            "polygon": True,
            "rectangle": True,
        },
        edit_options={"edit": False},
    ).add_to(m)

    output = st_folium(m, width=700, height=500, key="draw_map")

    if output and output.get("all_drawings"):
        drawings = output["all_drawings"]
        if drawings:
            # Take the last drawn shape
            last_drawing = drawings[-1]
            drawn_geom = shape(last_drawing["geometry"])
            boundary_gdf = gpd.GeoDataFrame(geometry=[drawn_geom], crs="EPSG:4326")
            st.success(f"Area defined: {drawn_geom.geom_type} with {len(drawings)} drawing(s)")

# --- Tab: Upload Shapefile ---
with tab_shp:
    uploaded_files = st.file_uploader(
        "Upload shapefile components (.shp, .shx, .dbf, .prj) or a .zip",
        type=["shp", "shx", "dbf", "prj", "cpg", "zip"],
        accept_multiple_files=True,
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
                # Convert to 4326 for unified handling
                boundary_gdf = shp_gdf.to_crs("EPSG:4326")
                st.success(
                    f"Loaded: {shp_files[0].name} — {len(boundary_gdf)} feature(s), "
                    f"CRS: {detected_crs}"
                )

                # Show boundary on map
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
            # Get bbox from boundary in EPSG:4326
            bounds = boundary_gdf.total_bounds  # minx, miny, maxx, maxy
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

                    # Filter to within polygon
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

        # Table with normalized field names
        display_cols = [
            "baum_id", "art_deutsch", "gattung_deutsch", "art_latein",
            "stammumfang", "kronendurchmesser", "baumhoehe",
            "pflanzjahr", "strasse", "hausnummer", "bezirk",
        ]
        existing_cols = [c for c in display_cols if c in filtered.columns]
        st.dataframe(filtered[existing_cols], use_container_width=True)

        # Map with trees
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
                radius=5,
                color="green",
                fill=True,
                fill_color="green",
                fill_opacity=0.7,
                popup=folium.Popup(popup_text, max_width=250),
            ).add_to(m2)

        st_folium(m2, width=700, height=500, key="trees_map")

        # --- Export ---
        st.header("4. Export VectorWorks Import TXT")
        st.caption(f"Output coordinates in: **{output_crs}** — {CRS_OPTIONS[output_crs]}")

        vw_txt = trees_to_vw_txt(filtered, output_crs,
                                  ansatz_method=ansatz_method,
                                  ansatz_ratio=ansatz_ratio,
                                  include_extra_cols=include_extra_cols)
        st.text_area("Preview (first 10 lines)", "\n".join(vw_txt.split("\n")[:11]), height=300)

        st.download_button(
            label="Download Baumkataster_VW_Import.txt",
            data=vw_txt.encode("utf-8"),
            file_name="Baumkataster_VW_Import.txt",
            mime="text/plain",
        )
