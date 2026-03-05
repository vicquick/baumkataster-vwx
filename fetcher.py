"""Unified fetch logic for WFS, ArcGIS REST, and 3D Tiles tree services."""

import math
import struct

import geopandas as gpd
import requests
from pyproj import Transformer
from shapely.geometry import MultiPoint, Point, shape


def fetch_trees(preset: dict, bbox_4326: tuple, max_features: int = 10000) -> gpd.GeoDataFrame:
    """Fetch trees from any preset. Returns GeoDataFrame in EPSG:4326 with normalized fields."""
    if preset["service_type"] == "arcgis_rest":
        gdf = _fetch_arcgis(preset, bbox_4326, max_features)
    elif preset["service_type"] == "3dtiles":
        gdf = _fetch_3dtiles(preset, bbox_4326, max_features)
    else:
        gdf = _fetch_wfs(preset, bbox_4326, max_features)

    if gdf.empty:
        return gdf

    # Normalize field names using field_map
    gdf = _apply_field_map(gdf, preset.get("field_map", {}))
    return gdf


def fetch_custom_wfs(url: str, type_name: str, field_map: dict,
                     bbox_4326: tuple, max_features: int = 10000,
                     wfs_version: str = "2.0.0",
                     output_format: str = "application/json",
                     typename_param: str = "typeNames",
                     native_crs: str = "EPSG:25832") -> gpd.GeoDataFrame:
    """Fetch from a custom WFS with user-provided field mapping."""
    preset = {
        "service_type": "wfs",
        "wfs_url": url,
        "type_name": type_name,
        "wfs_version": wfs_version,
        "output_format": output_format,
        "typename_param": typename_param,
        "native_crs": native_crs,
        "field_map": field_map,
    }
    return fetch_trees(preset, bbox_4326, max_features)


def discover_typenames(wfs_url: str) -> list[str]:
    """Do GetCapabilities to discover available typeNames."""
    params = {
        "service": "WFS",
        "request": "GetCapabilities",
    }
    resp = requests.get(wfs_url, params=params, timeout=30)
    resp.raise_for_status()

    # Parse XML to extract FeatureType names
    import xml.etree.ElementTree as ET
    root = ET.fromstring(resp.text)

    # Handle WFS namespaces
    names = []
    for elem in root.iter():
        if elem.tag.endswith("}Name") or elem.tag == "Name":
            parent_tag = ""
            # Walk up to check if parent is FeatureType
            # Since ElementTree doesn't support parent, search differently
            pass

    # Simpler approach: find all FeatureType elements
    for ft in root.iter():
        if ft.tag.endswith("}FeatureType") or ft.tag == "FeatureType":
            for child in ft:
                if child.tag.endswith("}Name") or child.tag == "Name":
                    if child.text:
                        names.append(child.text.strip())
                    break
    return names


def discover_fields(wfs_url: str, type_name: str, typename_param: str = "typeNames",
                    output_format: str = "application/json") -> list[str]:
    """Fetch a single feature to discover available field names."""
    params = {
        "service": "WFS",
        "version": "2.0.0",
        "request": "GetFeature",
        typename_param: type_name,
        "outputFormat": output_format,
        "count": "1",
    }
    resp = requests.get(wfs_url, params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    features = data.get("features", [])
    if features:
        return list(features[0].get("properties", {}).keys())
    return []


def _fetch_wfs(preset: dict, bbox_4326: tuple, max_features: int) -> gpd.GeoDataFrame:
    """Fetch trees via WFS. bbox_4326 is (minx, miny, maxx, maxy) in EPSG:4326."""
    native_crs = preset["native_crs"]

    # Convert bbox from 4326 to native CRS
    if native_crs != "EPSG:4326":
        transformer = Transformer.from_crs("EPSG:4326", native_crs, always_xy=True)
        minx, miny = transformer.transform(bbox_4326[0], bbox_4326[1])
        maxx, maxy = transformer.transform(bbox_4326[2], bbox_4326[3])
    else:
        minx, miny, maxx, maxy = bbox_4326

    params = {
        "service": "WFS",
        "version": preset.get("wfs_version", "2.0.0"),
        "request": "GetFeature",
        preset.get("typename_param", "typeName"): preset["type_name"],
        "outputFormat": preset.get("output_format", "application/json"),
        "srsName": native_crs,
        "count": str(max_features),
        "bbox": f"{minx},{miny},{maxx},{maxy},{native_crs}",
    }

    resp = requests.get(preset["wfs_url"], params=params, timeout=120)
    resp.raise_for_status()
    data = resp.json()

    features = data.get("features", [])
    if not features:
        return gpd.GeoDataFrame()

    records = []
    for feat in features:
        props = dict(feat.get("properties", {}))
        geom = shape(feat["geometry"])
        if isinstance(geom, MultiPoint):
            geom = Point(geom.geoms[0].x, geom.geoms[0].y)
        props["geometry"] = geom
        records.append(props)

    gdf = gpd.GeoDataFrame(records, crs=native_crs)

    # Reproject to EPSG:4326
    if native_crs != "EPSG:4326":
        gdf = gdf.to_crs("EPSG:4326")

    return gdf


def _fetch_arcgis(preset: dict, bbox_4326: tuple, max_features: int) -> gpd.GeoDataFrame:
    """Fetch trees via ArcGIS REST API."""
    rest_url = preset["rest_url"]
    minx, miny, maxx, maxy = bbox_4326

    # ArcGIS REST query endpoint
    query_url = f"{rest_url}/query"
    params = {
        "where": "1=1",
        "geometry": f"{minx},{miny},{maxx},{maxy}",
        "geometryType": "esriGeometryEnvelope",
        "inSR": "4326",
        "outSR": "4326",
        "spatialRel": "esriSpatialRelIntersects",
        "outFields": "*",
        "returnGeometry": "true",
        "f": "geojson",
        "resultRecordCount": str(max_features),
    }

    resp = requests.get(query_url, params=params, timeout=120)
    resp.raise_for_status()
    data = resp.json()

    features = data.get("features", [])
    if not features:
        return gpd.GeoDataFrame()

    records = []
    for feat in features:
        props = dict(feat.get("properties", {}))
        geom = shape(feat["geometry"])
        if isinstance(geom, MultiPoint):
            geom = Point(geom.geoms[0].x, geom.geoms[0].y)
        props["geometry"] = geom
        records.append(props)

    return gpd.GeoDataFrame(records, crs="EPSG:4326")


def _apply_field_map(gdf: gpd.GeoDataFrame, field_map: dict) -> gpd.GeoDataFrame:
    """Rename columns from city-specific names to normalized names."""
    # Build reverse map: city_field -> normalized_field
    rename = {}
    for city_field, norm_field in field_map.items():
        if city_field in gdf.columns:
            rename[city_field] = norm_field

    if rename:
        gdf = gdf.rename(columns=rename)
    return gdf


# --- 3D Tiles (i3dm) support ---

def _ecef_to_wgs84(x, y, z):
    """Convert ECEF coordinates to WGS84 (lat, lon) in degrees."""
    a = 6378137.0
    f = 1 / 298.257223563
    b = a * (1 - f)
    e2 = 1 - (b / a) ** 2
    lon = math.atan2(y, x)
    p = math.sqrt(x ** 2 + y ** 2)
    lat = math.atan2(z, p * (1 - e2))
    for _ in range(10):
        N = a / math.sqrt(1 - e2 * math.sin(lat) ** 2)
        lat = math.atan2(z + e2 * N * math.sin(lat), p)
    return math.degrees(lat), math.degrees(lon)


def _region_intersects_bbox(region, bbox_4326):
    """Check if a 3D Tiles region [west, south, east, north, ...] in radians intersects bbox in degrees."""
    west_deg = math.degrees(region[0])
    south_deg = math.degrees(region[1])
    east_deg = math.degrees(region[2])
    north_deg = math.degrees(region[3])
    minx, miny, maxx, maxy = bbox_4326
    return not (east_deg < minx or west_deg > maxx or north_deg < miny or south_deg > maxy)


def _parse_i3dm(data: bytes) -> list[dict]:
    """Parse an i3dm file and extract tree records with lat/lon from ECEF positions."""
    if len(data) < 32 or data[:4] != b"i3dm":
        return []

    _version = struct.unpack("<I", data[4:8])[0]
    ft_json_len = struct.unpack("<I", data[12:16])[0]
    ft_bin_len = struct.unpack("<I", data[16:20])[0]
    bt_json_len = struct.unpack("<I", data[20:24])[0]

    header_size = 32

    # Feature table (positions)
    import json
    ft_json = json.loads(data[header_size:header_size + ft_json_len])
    ft_bin_start = header_size + ft_json_len
    ft_bin = data[ft_bin_start:ft_bin_start + ft_bin_len]

    n = ft_json.get("INSTANCES_LENGTH", 0)
    if n == 0:
        return []

    # Extract positions (ECEF)
    positions = []
    if "POSITION_QUANTIZED" in ft_json:
        offset = ft_json["QUANTIZED_VOLUME_OFFSET"]
        scale = ft_json["QUANTIZED_VOLUME_SCALE"]
        pos_off = ft_json["POSITION_QUANTIZED"]["byteOffset"]
        for i in range(n):
            qx, qy, qz = struct.unpack_from("<HHH", ft_bin, pos_off + i * 6)
            x = offset[0] + qx / 65535.0 * scale[0]
            y = offset[1] + qy / 65535.0 * scale[1]
            z = offset[2] + qz / 65535.0 * scale[2]
            positions.append((x, y, z))
    elif "POSITION" in ft_json:
        pos_off = ft_json["POSITION"]["byteOffset"]
        for i in range(n):
            x, y, z = struct.unpack_from("<fff", ft_bin, pos_off + i * 12)
            positions.append((x, y, z))
    else:
        return []

    # Batch table (attributes)
    bt_json_start = ft_bin_start + ft_bin_len
    attributes_list = [{}] * n
    if bt_json_len > 0:
        bt_json = json.loads(data[bt_json_start:bt_json_start + bt_json_len])
        # Attributes may be in an "attributes" array or as direct arrays
        if "attributes" in bt_json and isinstance(bt_json["attributes"], list):
            attributes_list = bt_json["attributes"]
        else:
            # Direct arrays: each key maps to a list of values
            attributes_list = []
            for i in range(n):
                row = {}
                for key, val in bt_json.items():
                    if isinstance(val, list) and len(val) == n:
                        row[key] = val[i]
                attributes_list.append(row)

    records = []
    for i, (x, y, z) in enumerate(positions):
        lat, lon = _ecef_to_wgs84(x, y, z)
        attrs = attributes_list[i] if i < len(attributes_list) else {}
        attrs["geometry"] = Point(lon, lat)
        records.append(attrs)

    return records


def _parse_cmpt(data: bytes) -> list[dict]:
    """Parse a cmpt (composite) tile — extract inner i3dm tiles."""
    if len(data) < 16 or data[:4] != b"cmpt":
        return []

    _version = struct.unpack("<I", data[4:8])[0]
    _byte_length = struct.unpack("<I", data[8:12])[0]
    tiles_length = struct.unpack("<I", data[12:16])[0]

    records = []
    offset = 16
    for _ in range(tiles_length):
        if offset + 12 > len(data):
            break
        inner_magic = data[offset:offset + 4]
        inner_length = struct.unpack("<I", data[offset + 8:offset + 12])[0]
        inner_data = data[offset:offset + inner_length]

        if inner_magic == b"i3dm":
            records.extend(_parse_i3dm(inner_data))
        elif inner_magic == b"cmpt":
            records.extend(_parse_cmpt(inner_data))

        offset += inner_length

    return records


def _fetch_3dtiles(preset: dict, bbox_4326: tuple, max_features: int) -> gpd.GeoDataFrame:
    """Fetch trees from a 3D Tiles tileset, traversing the tile tree and parsing i3dm/cmpt tiles."""
    import json

    tileset_url = preset["tileset_url"]
    base_url = tileset_url.rsplit("/", 1)[0]

    resp = requests.get(tileset_url, timeout=30)
    resp.raise_for_status()
    root_tileset = resp.json()

    all_records = []
    tiles_to_visit = [(root_tileset["root"], base_url)]

    while tiles_to_visit and len(all_records) < max_features:
        node, node_base = tiles_to_visit.pop()

        # Check bounding volume intersection
        bv = node.get("boundingVolume", {})
        if "region" in bv:
            if not _region_intersects_bbox(bv["region"], bbox_4326):
                continue

        # If node has content, fetch it
        content = node.get("content", {})
        if "uri" in content:
            uri = content["uri"]
            # Resolve relative paths
            if uri.startswith(".."):
                parts = uri.split("/")
                base_parts = node_base.split("/")
                while parts and parts[0] == "..":
                    parts.pop(0)
                    if base_parts:
                        base_parts.pop()
                tile_url = "/".join(base_parts + parts)
            elif uri.startswith("http"):
                tile_url = uri
            else:
                tile_url = f"{node_base}/{uri}"

            try:
                tile_resp = requests.get(tile_url, timeout=30)
                tile_resp.raise_for_status()
                tile_data = tile_resp.content

                if uri.endswith(".i3dm"):
                    all_records.extend(_parse_i3dm(tile_data))
                elif uri.endswith(".cmpt"):
                    all_records.extend(_parse_cmpt(tile_data))
                elif uri.endswith(".json"):
                    # Subtileset — add its root to the visit queue
                    sub_tileset = tile_resp.json()
                    sub_base = tile_url.rsplit("/", 1)[0]
                    tiles_to_visit.append((sub_tileset["root"], sub_base))
            except Exception:
                pass  # Skip failed tiles

        # Add children to visit
        for child in node.get("children", []):
            tiles_to_visit.append((child, node_base))

    if not all_records:
        return gpd.GeoDataFrame()

    gdf = gpd.GeoDataFrame(all_records, crs="EPSG:4326")

    # Filter points to those actually inside the bbox (tiles are coarse)
    minx, miny, maxx, maxy = bbox_4326
    gdf = gdf.cx[minx:maxx, miny:maxy].copy()

    # Truncate to max_features
    if len(gdf) > max_features:
        gdf = gdf.iloc[:max_features]
    return gdf
