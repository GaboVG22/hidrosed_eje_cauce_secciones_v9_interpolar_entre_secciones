# -*- coding: utf-8 -*-
"""
HidroSed · Módulo Eje Cauce y Secciones v8 · Interpolación por curvas y canal por tramo
Aplicación Streamlit independiente para validar eje de cauce, tramo útil, perfil longitudinal,
secciones transversales y exportaciones KMZ/Excel/CSV/JSON.

Diseñada para integrarse posteriormente en HidroSed.
"""

from __future__ import annotations

import io
import json
import math
import os
import tempfile
import zipfile
import requests
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple, Any
import xml.etree.ElementTree as ET
from html import escape

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from shapely.geometry import LineString, Point, MultiLineString
from shapely.ops import transform as shp_transform
from pyproj import CRS, Transformer
import matplotlib.pyplot as plt

# Dependencias opcionales controladas
try:
    import rasterio
    from rasterio.transform import rowcol
except Exception:  # pragma: no cover
    rasterio = None

try:
    import ezdxf
except Exception:  # pragma: no cover
    ezdxf = None

try:
    import plotly.graph_objects as go
except Exception:  # pragma: no cover
    go = None


# --------------------------------------------------------------------------------------
# Configuración Streamlit
# --------------------------------------------------------------------------------------
st.set_page_config(
    page_title="HidroSed · Eje Cauce y Secciones",
    page_icon="🌊",
    layout="wide",
)


# --------------------------------------------------------------------------------------
# Modelos de datos
# --------------------------------------------------------------------------------------
@dataclass
class ProfilePoint:
    km: float
    x: float
    y: float
    z_dem: Optional[float] = None
    z_respaldo: Optional[float] = None
    diferencia: Optional[float] = None
    pendiente_local: Optional[float] = None


@dataclass
class CrossSection:
    km: float
    center_x: float
    center_y: float
    left_x: float
    left_y: float
    right_x: float
    right_y: float
    stations: List[float]
    elevations: List[float]
    source: str
    status: str = "aceptada"
    notes: str = ""


@dataclass
class AxisValidation:
    axis_length_m: float
    useful_length_m: float
    km_pc_hidrologico: float
    km_pc_soporte: float
    direction: str
    inverted: bool
    left_right_defined: bool
    warnings: List[str]


# --------------------------------------------------------------------------------------
# Utilidades de archivos KML/KMZ/DXF/Excel
# --------------------------------------------------------------------------------------
def uploaded_bytes(uploaded_file) -> bytes:
    uploaded_file.seek(0)
    return uploaded_file.read()


def read_kml_text_from_upload(uploaded_file) -> str:
    """Lee archivo KML o KMZ y retorna texto KML."""
    data = uploaded_bytes(uploaded_file)
    name = uploaded_file.name.lower()
    if name.endswith(".kmz"):
        with zipfile.ZipFile(io.BytesIO(data), "r") as z:
            kml_names = [n for n in z.namelist() if n.lower().endswith(".kml")]
            if not kml_names:
                raise ValueError("El KMZ no contiene archivo .kml interno.")
            # preferir doc.kml si existe
            kml_name = "doc.kml" if "doc.kml" in kml_names else kml_names[0]
            return z.read(kml_name).decode("utf-8", errors="ignore")
    return data.decode("utf-8", errors="ignore")


def parse_kml_coordinates(coord_text: str) -> List[Tuple[float, float, Optional[float]]]:
    coords: List[Tuple[float, float, Optional[float]]] = []
    if not coord_text:
        return coords
    for token in coord_text.replace("\n", " ").replace("\t", " ").split():
        parts = token.split(",")
        if len(parts) >= 2:
            try:
                lon = float(parts[0])
                lat = float(parts[1])
                z = float(parts[2]) if len(parts) >= 3 and parts[2] != "" else None
                coords.append((lon, lat, z))
            except ValueError:
                continue
    return coords


def extract_kml_geometries(uploaded_file) -> Dict[str, Any]:
    """Extrae puntos y líneas de KML/KMZ. Coordenadas en lon,lat,z."""
    kml_text = read_kml_text_from_upload(uploaded_file)
    root = ET.fromstring(kml_text.encode("utf-8"))
    ns = {"kml": "http://www.opengis.net/kml/2.2"}

    def findall_any(parent, tag):
        found = parent.findall(f".//kml:{tag}", ns)
        if found:
            return found
        return parent.findall(f".//{tag}")

    points = []
    lines = []
    placemarks = findall_any(root, "Placemark")
    for pm in placemarks:
        name_el = pm.find("kml:name", ns)
        if name_el is None:
            name_el = pm.find("name")
        name = name_el.text.strip() if name_el is not None and name_el.text else "sin_nombre"
        for pt in findall_any(pm, "Point"):
            coord_el = pt.find("kml:coordinates", ns)
            if coord_el is None:
                coord_el = pt.find("coordinates")
            coords = parse_kml_coordinates(coord_el.text if coord_el is not None else "")
            if coords:
                points.append({"name": name, "coords": coords[0]})
        for line in findall_any(pm, "LineString"):
            coord_el = line.find("kml:coordinates", ns)
            if coord_el is None:
                coord_el = line.find("coordinates")
            coords = parse_kml_coordinates(coord_el.text if coord_el is not None else "")
            if len(coords) >= 2:
                lines.append({"name": name, "coords": coords})
        # gx:Track no implementado formalmente, pero se intenta leer gx:coord si existe
        for track in pm.findall(".//{http://www.google.com/kml/ext/2.2}Track"):
            coords = []
            for coord in track.findall("{http://www.google.com/kml/ext/2.2}coord"):
                vals = coord.text.split() if coord.text else []
                if len(vals) >= 2:
                    lon, lat = float(vals[0]), float(vals[1])
                    z = float(vals[2]) if len(vals) >= 3 else None
                    coords.append((lon, lat, z))
            if len(coords) >= 2:
                lines.append({"name": name, "coords": coords})
    return {"points": points, "lines": lines, "kml_text": kml_text}


def pick_first_point(uploaded_file, label="punto") -> Tuple[float, float, Optional[float], str]:
    geoms = extract_kml_geometries(uploaded_file)
    if not geoms["points"]:
        raise ValueError(f"El archivo de {label} no contiene puntos válidos.")
    p = geoms["points"][0]
    lon, lat, z = p["coords"]
    return lon, lat, z, p["name"]


def pick_longest_line(uploaded_file, label="línea") -> Tuple[List[Tuple[float, float, Optional[float]]], str]:
    geoms = extract_kml_geometries(uploaded_file)
    if not geoms["lines"]:
        raise ValueError(f"El archivo de {label} no contiene una línea válida.")
    # largo aproximado en grados solo para elegir; después se reproyecta
    def approx_len(coords):
        total = 0.0
        for a, b in zip(coords[:-1], coords[1:]):
            total += math.hypot(b[0] - a[0], b[1] - a[1])
        return total
    line = max(geoms["lines"], key=lambda d: approx_len(d["coords"]))
    return line["coords"], line["name"]


def determine_local_utm_crs(lon: float, lat: float) -> CRS:
    zone = int(math.floor((lon + 180.0) / 6.0) + 1)
    epsg = 32600 + zone if lat >= 0 else 32700 + zone
    return CRS.from_epsg(epsg)


def transformers_for_crs(crs_proj: CRS):
    to_proj = Transformer.from_crs("EPSG:4326", crs_proj, always_xy=True)
    to_wgs = Transformer.from_crs(crs_proj, "EPSG:4326", always_xy=True)
    return to_proj, to_wgs


def coords_to_linestring_projected(coords_lonlatz, to_proj) -> Tuple[LineString, List[Optional[float]]]:
    pts = []
    zs = []
    for lon, lat, z in coords_lonlatz:
        x, y = to_proj.transform(lon, lat)
        pts.append((x, y))
        zs.append(z)
    return LineString(pts), zs


def point_projected(lon: float, lat: float, to_proj) -> Point:
    x, y = to_proj.transform(lon, lat)
    return Point(x, y)


def load_support_curves(uploaded_file, to_proj) -> List[Dict[str, Any]]:
    """Carga curvas de nivel de apoyo desde KMZ/KML o DXF.
    Retorna lista {line: LineString proyectada, elev: float|None, name: str, source:str}
    """
    if uploaded_file is None:
        return []
    name = uploaded_file.name.lower()
    curves: List[Dict[str, Any]] = []
    if name.endswith((".kml", ".kmz")):
        geoms = extract_kml_geometries(uploaded_file)
        for ln in geoms["lines"]:
            line, zs = coords_to_linestring_projected(ln["coords"], to_proj)
            elev = None
            zs_non = [z for z in zs if z is not None]
            if zs_non:
                elev = float(np.nanmedian(zs_non))
            else:
                # Intentar extraer número del nombre
                import re
                m = re.search(r"(-?\d+(?:\.\d+)?)", ln.get("name", ""))
                if m:
                    elev = float(m.group(1))
            curves.append({"line": line, "elev": elev, "name": ln.get("name", "curva"), "source": "kml/kmz"})
    elif name.endswith(".dxf"):
        if ezdxf is None:
            raise RuntimeError("ezdxf no está instalado. Agrega ezdxf al requirements.txt.")
        data = uploaded_bytes(uploaded_file)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        try:
            doc = ezdxf.readfile(tmp_path)
            msp = doc.modelspace()
            for e in msp:
                etype = e.dxftype()
                elev = None
                pts = []
                layer = getattr(e.dxf, "layer", "dxf")
                if etype == "LWPOLYLINE":
                    try:
                        elev = float(getattr(e.dxf, "elevation", 0.0))
                    except Exception:
                        elev = None
                    for p in e.get_points():
                        # p: x, y, start_width, end_width, bulge
                        pts.append((float(p[0]), float(p[1])))
                elif etype in ("POLYLINE", "3DPOLYLINE"):
                    try:
                        for v in e.vertices:
                            loc = v.dxf.location
                            pts.append((float(loc.x), float(loc.y)))
                            if elev is None:
                                elev = float(loc.z)
                    except Exception:
                        continue
                elif etype == "LINE":
                    s = e.dxf.start
                    t = e.dxf.end
                    pts = [(float(s.x), float(s.y)), (float(t.x), float(t.y))]
                    elev = float(np.nanmean([s.z, t.z])) if hasattr(s, "z") else None
                if len(pts) >= 2:
                    # Asumimos DXF en coordenadas UTM WGS84 local. No reproyectamos.
                    curves.append({"line": LineString(pts), "elev": elev, "name": layer, "source": "dxf"})
        finally:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass
    else:
        raise ValueError("Formato de curvas de apoyo no soportado. Use KMZ/KML/DXF.")
    return curves


def load_profile_file(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    name = uploaded_file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(uploaded_file)
    elif name.endswith((".csv", ".txt")):
        data = uploaded_bytes(uploaded_file)
        text = data.decode("utf-8", errors="ignore")
        # autodetectar separador básico
        sep = ";" if text.count(";") > text.count(",") else ","
        df = pd.read_csv(io.StringIO(text), sep=sep)
    elif name.endswith((".kml", ".kmz")):
        coords, _ = pick_longest_line(uploaded_file, "perfil longitudinal")
        rows = []
        cum = 0.0
        prev = None
        # Distancia aproximada en grados no sirve; se ajusta después si se usa como XY/latlon.
        for lon, lat, z in coords:
            if prev is not None:
                cum += math.hypot(lon - prev[0], lat - prev[1])
            rows.append({"lon": lon, "lat": lat, "cota": z, "distancia": cum})
            prev = (lon, lat)
        df = pd.DataFrame(rows)
    elif name.endswith(".dxf"):
        if ezdxf is None:
            raise RuntimeError("ezdxf no está instalado. Agrega ezdxf al requirements.txt.")
        data = uploaded_bytes(uploaded_file)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        rows = []
        try:
            doc = ezdxf.readfile(tmp_path)
            msp = doc.modelspace()
            longest = []
            for e in msp:
                pts = []
                if e.dxftype() == "LWPOLYLINE":
                    elev = float(getattr(e.dxf, "elevation", 0.0))
                    for p in e.get_points():
                        pts.append((float(p[0]), float(p[1]), elev))
                elif e.dxftype() in ("POLYLINE", "3DPOLYLINE"):
                    for v in e.vertices:
                        loc = v.dxf.location
                        pts.append((float(loc.x), float(loc.y), float(loc.z)))
                if len(pts) > len(longest):
                    longest = pts
            cum = 0.0
            prev = None
            for x, y, z in longest:
                if prev is not None:
                    cum += math.hypot(x - prev[0], y - prev[1])
                rows.append({"x": x, "y": y, "cota": z, "distancia": cum})
                prev = (x, y)
        finally:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass
        df = pd.DataFrame(rows)
    else:
        raise ValueError("Formato de perfil no soportado.")
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def normalize_profile_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    # Identificación flexible
    cols = list(df.columns)
    cota_col = next((c for c in cols if c in ["cota", "z", "elev", "elevacion", "elevación", "altitud"]), None)
    dist_col = next((c for c in cols if c in ["km", "dist", "distancia", "distancia_m", "chainage", "station"]), None)
    if cota_col is None:
        raise ValueError("El perfil debe contener una columna de cota: cota, z, elev, elevacion o altitud.")
    out = pd.DataFrame()
    if dist_col is not None:
        dist = pd.to_numeric(df[dist_col], errors="coerce")
        # Si parece estar en km, convertir a m para interpolación interna
        if dist.max(skipna=True) is not None and dist.max(skipna=True) < 1000 and "km" in dist_col:
            dist_m = dist * 1000.0
        else:
            dist_m = dist
        out["dist_m"] = dist_m
    elif "x" in cols and "y" in cols:
        x = pd.to_numeric(df["x"], errors="coerce")
        y = pd.to_numeric(df["y"], errors="coerce")
        dist_m = [0.0]
        for i in range(1, len(df)):
            dist_m.append(dist_m[-1] + float(math.hypot(x.iloc[i] - x.iloc[i-1], y.iloc[i] - y.iloc[i-1])))
        out["dist_m"] = dist_m
    else:
        # fallback índice
        out["dist_m"] = np.arange(len(df), dtype=float)
    out["cota"] = pd.to_numeric(df[cota_col], errors="coerce")
    out = out.dropna().sort_values("dist_m")
    out = out.drop_duplicates("dist_m")
    return out




# --------------------------------------------------------------------------------------
# Descarga DEM desde OpenTopography
# --------------------------------------------------------------------------------------
class BytesUpload:
    """Objeto mínimo compatible con uploaded_bytes para usar DEM descargado por API."""
    def __init__(self, name: str, data: bytes):
        self.name = name
        self._bio = io.BytesIO(data)

    def seek(self, pos: int):
        return self._bio.seek(pos)

    def read(self, *args, **kwargs):
        return self._bio.read(*args, **kwargs)


def bbox_from_lonlat_inputs(axis_coords, pc_h_lon, pc_h_lat, pc_s_lon, pc_s_lat, margin_km: float):
    """Calcula bbox WGS84 para OpenTopography a partir del eje y PCs, con margen en km."""
    lons = [float(c[0]) for c in axis_coords] + [float(pc_h_lon), float(pc_s_lon)]
    lats = [float(c[1]) for c in axis_coords] + [float(pc_h_lat), float(pc_s_lat)]
    lat_mid = float(np.nanmean(lats)) if lats else 0.0
    deg_lat = float(margin_km) / 111.32
    cos_lat = max(0.15, abs(math.cos(math.radians(lat_mid))))
    deg_lon = float(margin_km) / (111.32 * cos_lat)
    south = max(-90.0, min(lats) - deg_lat)
    north = min(90.0, max(lats) + deg_lat)
    west = max(-180.0, min(lons) - deg_lon)
    east = min(180.0, max(lons) + deg_lon)
    return west, south, east, north


def estimate_bbox_area_km2(west: float, south: float, east: float, north: float) -> float:
    lat_mid = (south + north) / 2.0
    width = max(0.0, east - west) * 111.32 * max(0.15, abs(math.cos(math.radians(lat_mid))))
    height = max(0.0, north - south) * 111.32
    return width * height


def download_opentopography_dem(api_key: str, demtype: str, west: float, south: float, east: float, north: float) -> bytes:
    """Descarga DEM global de OpenTopography como GeoTIFF.

    No imprime ni guarda la API Key. Lanza excepción con mensaje controlado si falla.
    """
    if not api_key or not str(api_key).strip():
        raise ValueError("Debe ingresar una API Key de OpenTopography.")
    if not (south < north and west < east):
        raise ValueError("Bounding box inválido: verifique west/east/south/north.")
    url = "https://portal.opentopography.org/API/globaldem"
    params = {
        "demtype": demtype,
        "south": f"{south:.8f}",
        "north": f"{north:.8f}",
        "west": f"{west:.8f}",
        "east": f"{east:.8f}",
        "outputFormat": "GTiff",
        "API_Key": api_key.strip(),
    }
    try:
        r = requests.get(url, params=params, timeout=180)
    except requests.RequestException as exc:
        raise RuntimeError(f"No se pudo conectar con OpenTopography: {exc}") from exc
    if r.status_code != 200:
        txt = r.text[:600] if hasattr(r, 'text') else ''
        raise RuntimeError(f"OpenTopography respondió código {r.status_code}. Revise API Key, DEM, bbox o límites del servicio. Respuesta: {txt}")
    data = r.content
    if len(data) < 1024:
        txt = data.decode('utf-8', errors='ignore')[:600]
        raise RuntimeError(f"La respuesta de OpenTopography no parece ser un GeoTIFF válido: {txt}")
    head = data[:20].lower()
    if head.startswith(b'{') or head.startswith(b'<') or b'error' in head:
        txt = data.decode('utf-8', errors='ignore')[:600]
        raise RuntimeError(f"OpenTopography no entregó un GeoTIFF válido: {txt}")
    return data


# --------------------------------------------------------------------------------------
# Geometría del eje y tramo útil
# --------------------------------------------------------------------------------------
def orient_axis_by_dem_or_user(axis: LineString, dem_sampler, user_invert: bool = False) -> Tuple[LineString, bool, str, Optional[float], Optional[float]]:
    coords = list(axis.coords)
    z0 = z1 = None
    inverted = False
    direction_note = "Sentido conservado según digitalización original."
    if dem_sampler is not None:
        z0 = dem_sampler(coords[0][0], coords[0][1])
        z1 = dem_sampler(coords[-1][0], coords[-1][1])
        if z0 is not None and z1 is not None and np.isfinite(z0) and np.isfinite(z1):
            # Aguas arriba mayor cota. Línea se orienta aguas arriba -> aguas abajo.
            if z0 < z1:
                coords = list(reversed(coords))
                inverted = True
                direction_note = "Eje invertido automáticamente: extremo final tenía mayor cota DEM."
            else:
                direction_note = "Eje conservado: extremo inicial tiene mayor cota DEM."
        else:
            direction_note = "No se pudo determinar sentido por DEM; se conserva digitalización."
    if user_invert:
        coords = list(reversed(coords))
        inverted = not inverted
        direction_note += " Inversión manual aplicada por usuario."
    return LineString(coords), inverted, direction_note, z0, z1


def extract_subline(line: LineString, d0: float, d1: float) -> LineString:
    """Extrae sublínea entre distancias d0 y d1 sobre una LineString."""
    if d1 < d0:
        d0, d1 = d1, d0
    d0 = max(0.0, min(float(d0), line.length))
    d1 = max(0.0, min(float(d1), line.length))
    if d1 <= d0:
        raise ValueError("El tramo útil tiene longitud cero. Revise PCs y eje.")
    coords = list(line.coords)
    new_pts = [line.interpolate(d0).coords[0]]
    dist_acc = 0.0
    for a, b in zip(coords[:-1], coords[1:]):
        seg_len = math.hypot(b[0] - a[0], b[1] - a[1])
        next_acc = dist_acc + seg_len
        if next_acc > d0 and dist_acc < d1:
            if dist_acc >= d0 and next_acc <= d1:
                new_pts.append(b)
        dist_acc = next_acc
    new_pts.append(line.interpolate(d1).coords[0])
    # eliminar duplicados consecutivos
    cleaned = []
    for p in new_pts:
        if not cleaned or math.hypot(p[0] - cleaned[-1][0], p[1] - cleaned[-1][1]) > 1e-6:
            cleaned.append(p)
    if len(cleaned) < 2:
        raise ValueError("No se pudo extraer tramo útil válido.")
    return LineString(cleaned)


def tangent_at_distance(line: LineString, d: float, eps: float = 1.0) -> Tuple[float, float]:
    d0 = max(0.0, d - eps)
    d1 = min(line.length, d + eps)
    if d1 <= d0:
        d0 = max(0.0, d - 5.0)
        d1 = min(line.length, d + 5.0)
    p0 = line.interpolate(d0)
    p1 = line.interpolate(d1)
    dx = p1.x - p0.x
    dy = p1.y - p0.y
    n = math.hypot(dx, dy)
    if n == 0:
        return (1.0, 0.0)
    return (dx / n, dy / n)


def cross_section_line(line: LineString, d: float, half_width: float) -> LineString:
    p = line.interpolate(d)
    tx, ty = tangent_at_distance(line, d)
    # Normal izquierda mirando aguas abajo
    nx, ny = -ty, tx
    left = (p.x + nx * half_width, p.y + ny * half_width)
    right = (p.x - nx * half_width, p.y - ny * half_width)
    # HEC-RAS: puntos left-to-right looking downstream
    return LineString([left, right])


def make_station_points(section: LineString, station_step: float) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
    length = section.length
    n = max(2, int(math.ceil(length / station_step)) + 1)
    ds = np.linspace(0.0, length, n)
    xs = np.array([section.interpolate(float(d)).x for d in ds])
    ys = np.array([section.interpolate(float(d)).y for d in ds])
    return ds, xs, ys


# --------------------------------------------------------------------------------------
# DEM sampler
# --------------------------------------------------------------------------------------
class DemSampler:
    def __init__(self, uploaded_file=None):
        self.dataset = None
        self.path = None
        self.crs = None
        self.transformer_from_proj = None
        self.nodata = None
        if uploaded_file is not None:
            if rasterio is None:
                raise RuntimeError("rasterio no está instalado.")
            data = uploaded_bytes(uploaded_file)
            suffix = ".tif" if uploaded_file.name.lower().endswith(".tif") else ".tiff"
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            tmp.write(data)
            tmp.close()
            self.path = tmp.name
            self.dataset = rasterio.open(self.path)
            self.crs = self.dataset.crs
            self.nodata = self.dataset.nodata

    def close(self):
        try:
            if self.dataset:
                self.dataset.close()
        finally:
            if self.path and os.path.exists(self.path):
                try:
                    os.unlink(self.path)
                except Exception:
                    pass

    def sample_xy(self, x: float, y: float, source_crs: CRS) -> Optional[float]:
        if self.dataset is None:
            return None
        if self.crs is None:
            return None
        if CRS.from_user_input(source_crs) != CRS.from_user_input(self.crs):
            transformer = Transformer.from_crs(source_crs, self.crs, always_xy=True)
            xx, yy = transformer.transform(x, y)
        else:
            xx, yy = x, y
        try:
            val = next(self.dataset.sample([(xx, yy)]))[0]
            if self.nodata is not None and float(val) == float(self.nodata):
                return None
            if not np.isfinite(val):
                return None
            return float(val)
        except Exception:
            return None

    def sample_many_xy(self, xs: Sequence[float], ys: Sequence[float], source_crs: CRS) -> np.ndarray:
        if self.dataset is None:
            return np.full(len(xs), np.nan)
        if CRS.from_user_input(source_crs) != CRS.from_user_input(self.crs):
            transformer = Transformer.from_crs(source_crs, self.crs, always_xy=True)
            pts = [transformer.transform(float(x), float(y)) for x, y in zip(xs, ys)]
        else:
            pts = list(zip(xs, ys))
        vals = []
        try:
            for v in self.dataset.sample(pts):
                val = float(v[0])
                if self.nodata is not None and val == float(self.nodata):
                    vals.append(np.nan)
                elif np.isfinite(val):
                    vals.append(val)
                else:
                    vals.append(np.nan)
        except Exception:
            vals = [np.nan] * len(xs)
        return np.array(vals, dtype=float)


# --------------------------------------------------------------------------------------
# Perfil longitudinal y secciones
# --------------------------------------------------------------------------------------
def generate_longitudinal_profile(line: LineString, dem: Optional[DemSampler], crs_proj: CRS, step_m: float = 25.0, respaldo: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    n = max(2, int(math.ceil(line.length / step_m)) + 1)
    dists = np.linspace(0.0, line.length, n)
    xs = np.array([line.interpolate(float(d)).x for d in dists])
    ys = np.array([line.interpolate(float(d)).y for d in dists])
    if dem is not None and dem.dataset is not None:
        z_dem = dem.sample_many_xy(xs, ys, crs_proj)
    else:
        z_dem = np.full(n, np.nan)
    z_res = np.full(n, np.nan)
    if respaldo is not None and not respaldo.empty:
        r = normalize_profile_df(respaldo)
        if not r.empty and r["dist_m"].max() > r["dist_m"].min():
            # Escala perfil respaldo al largo del tramo útil si difiere razonablemente.
            rr_dist = r["dist_m"].to_numpy(dtype=float)
            rr_z = r["cota"].to_numpy(dtype=float)
            scale = line.length / rr_dist.max() if rr_dist.max() > 0 else 1.0
            rr_scaled = rr_dist * scale
            z_res = np.interp(dists, rr_scaled, rr_z, left=np.nan, right=np.nan)
    diff = z_dem - z_res
    slope = np.full(n, np.nan)
    if np.isfinite(z_dem).sum() >= 2:
        for i in range(1, n):
            if np.isfinite(z_dem[i]) and np.isfinite(z_dem[i-1]) and dists[i] > dists[i-1]:
                slope[i] = (z_dem[i-1] - z_dem[i]) / (dists[i] - dists[i-1])
    return pd.DataFrame({
        "km": dists / 1000.0,
        "dist_m": dists,
        "x": xs,
        "y": ys,
        "cota_dem": z_dem,
        "cota_respaldo": z_res,
        "diferencia_dem_menos_respaldo": diff,
        "pendiente_local_m_m": slope,
    })


def evaluate_profile(profile_df: pd.DataFrame) -> List[str]:
    warnings = []
    if profile_df.empty:
        return ["No se generó perfil longitudinal."]
    z = profile_df["cota_dem"].to_numpy(dtype=float)
    d = profile_df["dist_m"].to_numpy(dtype=float)
    finite = np.isfinite(z)
    if finite.sum() < max(3, len(z) // 4):
        warnings.append("El perfil DEM tiene muchas cotas sin dato. Revise cobertura del DEM.")
    if finite.sum() >= 2:
        # Eje orientado aguas arriba -> aguas abajo: debería tender a descender.
        dz_total = z[finite][0] - z[finite][-1]
        if dz_total < 0:
            warnings.append("El perfil DEM aumenta aguas abajo. Revise sentido del eje o calidad del DEM.")
        # Saltos locales grandes relativos
        dz = np.diff(z)
        dd = np.diff(d)
        with np.errstate(invalid="ignore", divide="ignore"):
            grad = np.abs(dz / dd)
        if np.nanmax(grad) > 0.50:
            warnings.append("Se detectan saltos de pendiente local muy altos en el perfil DEM (>50%). Revise eje/DEM/topografía.")
    if "cota_respaldo" in profile_df.columns and np.isfinite(profile_df["cota_respaldo"]).sum() > 2:
        diff = profile_df["diferencia_dem_menos_respaldo"].to_numpy(dtype=float)
        mad = np.nanmean(np.abs(diff))
        mx = np.nanmax(np.abs(diff))
        if mad > 5:
            warnings.append(f"Diferencia media DEM vs perfil respaldo alta: {mad:.2f} m.")
        if mx > 15:
            warnings.append(f"Diferencia máxima DEM vs perfil respaldo alta: {mx:.2f} m.")
    return warnings


def generate_natural_sections(line: LineString, dem: DemSampler, crs_proj: CRS, spacing_m: float, half_width_m: float, station_step_m: float) -> List[CrossSection]:
    if dem is None or dem.dataset is None:
        return []
    dists = np.arange(0.0, line.length + 0.001, spacing_m)
    if dists[-1] < line.length:
        dists = np.append(dists, line.length)
    sections: List[CrossSection] = []
    for d in dists:
        sec_line = cross_section_line(line, float(d), half_width_m)
        st, xs, ys = make_station_points(sec_line, station_step_m)
        zs = dem.sample_many_xy(xs, ys, crs_proj)
        status = "aceptada" if np.isfinite(zs).sum() >= max(3, len(zs) // 2) else "revisar"
        notes = "Generada desde DEM."
        if status == "revisar":
            notes += " Muchas cotas sin dato."
        center = line.interpolate(float(d))
        sections.append(CrossSection(
            km=float(d) / 1000.0,
            center_x=center.x,
            center_y=center.y,
            left_x=sec_line.coords[0][0],
            left_y=sec_line.coords[0][1],
            right_x=sec_line.coords[-1][0],
            right_y=sec_line.coords[-1][1],
            stations=st.tolist(),
            elevations=[float(v) if np.isfinite(v) else np.nan for v in zs],
            source="DEM",
            status=status,
            notes=notes,
        ))
    return sections



def _intersection_points_for_profile(geom) -> List[Point]:
    """Convierte intersecciones Shapely sección-curva en puntos útiles.

    Replica la lógica robusta de la aplicación v13: si una curva coincide parcialmente
    con la sección, toma el punto medio del tramo de superposición.
    """
    if geom is None or geom.is_empty:
        return []
    gt = geom.geom_type
    if gt == "Point":
        return [geom]
    if gt == "MultiPoint":
        return list(geom.geoms)
    if gt == "LineString":
        return [geom.interpolate(geom.length / 2.0)] if geom.length > 0 else []
    if gt == "MultiLineString":
        return [part.interpolate(part.length / 2.0) for part in geom.geoms if part.length > 0]
    if gt == "GeometryCollection":
        pts: List[Point] = []
        for part in geom.geoms:
            pts.extend(_intersection_points_for_profile(part))
        return pts
    return []


def _interp_axis_elevation_from_section(stations: Sequence[float], elevations: Sequence[float], section_width: float) -> Optional[float]:
    """Interpola la cota en el eje de la sección usando puntos estación-cota."""
    if not stations or not elevations or len(stations) < 2:
        return None
    df = pd.DataFrame({"s": pd.to_numeric(pd.Series(stations), errors="coerce"), "z": pd.to_numeric(pd.Series(elevations), errors="coerce")})
    df = df.dropna().sort_values("s")
    if len(df) < 2:
        return None
    axis_s = float(section_width) / 2.0
    if float(df["s"].min()) <= axis_s <= float(df["s"].max()):
        return float(np.interp(axis_s, df["s"].to_numpy(dtype=float), df["z"].to_numpy(dtype=float)))
    return None


def generate_sections_from_contours_v13(
    line: LineString,
    curves: List[Dict[str, Any]],
    spacing_m: float,
    half_width_m: float,
    min_points_each_bank: int = 2,
    min_total_points: int = 4,
) -> List[CrossSection]:
    """Genera secciones por intersección sección-curva, replicando el modelo v13.

    En vez de muestrear únicamente un DEM, corta cada sección transversal con las curvas
    de nivel de apoyo y construye el perfil station-elevation de izquierda a derecha,
    mirando aguas abajo. Este método fue el que dio mejores resultados en la app
    `app_secciones_kmz_v13_fix_km_final_utm19s_3d`.
    """
    if not curves:
        return []
    usable_curves = [c for c in curves if c.get("elev") is not None and c.get("line") is not None]
    if not usable_curves:
        return []
    dists = np.arange(0.0, line.length + 0.001, spacing_m)
    if len(dists) == 0 or dists[-1] < line.length:
        dists = np.append(dists, line.length)

    sections: List[CrossSection] = []
    for d in dists:
        sec_line = cross_section_line(line, float(d), half_width_m)
        width = float(sec_line.length)
        axis_station = width / 2.0
        raw_pts: List[Tuple[float, float, Point, str]] = []
        for c in usable_curves:
            try:
                inter = sec_line.intersection(c["line"])
            except Exception:
                continue
            for pt in _intersection_points_for_profile(inter):
                sta = float(sec_line.project(pt))
                if -1e-6 <= sta <= width + 1e-6:
                    raw_pts.append((sta, float(c["elev"]), pt, str(c.get("name", "curva"))))
        raw_pts.sort(key=lambda t: (t[0], t[1]))
        clean: List[Tuple[float, float, Point, str]] = []
        for sta, elev, pt, name in raw_pts:
            if clean and abs(sta - clean[-1][0]) <= 0.05 and abs(elev - clean[-1][1]) <= 0.01:
                continue
            clean.append((sta, elev, pt, name))

        stations = [float(t[0]) for t in clean]
        elevs = [float(t[1]) for t in clean]
        left_count = sum(1 for s0 in stations if s0 < axis_station - 1e-6)
        right_count = sum(1 for s0 in stations if s0 > axis_station + 1e-6)
        axis_z = _interp_axis_elevation_from_section(stations, elevs, width)
        reasons: List[str] = []
        if len(stations) < min_total_points:
            reasons.append(f"pocos puntos curva-sección ({len(stations)}<{min_total_points})")
        if left_count < min_points_each_bank:
            reasons.append(f"ribera izquierda insuficiente ({left_count}<{min_points_each_bank})")
        if right_count < min_points_each_bank:
            reasons.append(f"ribera derecha insuficiente ({right_count}<{min_points_each_bank})")
        if axis_z is None:
            reasons.append("sin interpolación de cota en eje")
        status = "aceptada" if not reasons else "revisar"
        notes = "Modelo v13: intersección sección-curvas de nivel. "
        notes += f"Puntos={len(stations)}, izquierda={left_count}, derecha={right_count}."
        if reasons:
            notes += " Revisar: " + "; ".join(reasons)
        center = line.interpolate(float(d))
        sections.append(CrossSection(
            km=float(d) / 1000.0,
            center_x=center.x,
            center_y=center.y,
            left_x=sec_line.coords[0][0],
            left_y=sec_line.coords[0][1],
            right_x=sec_line.coords[-1][0],
            right_y=sec_line.coords[-1][1],
            stations=stations,
            elevations=elevs,
            source="Curvas nivel v13",
            status=status,
            notes=notes,
        ))
    return sections


def longitudinal_profile_from_sections_v13(sections: List[CrossSection], fallback_profile: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """Construye perfil longitudinal de modelación usando cota interpolada en el eje de cada sección."""
    rows: List[Dict[str, Any]] = []
    for sec in sections:
        z_axis = _interp_axis_elevation_from_section(sec.stations, sec.elevations, abs(sec.right_x - sec.left_x) if False else max(sec.stations) - min(sec.stations) if sec.stations else 0)
        # Si las estaciones no parten en 0 por edición manual, usa el rango real para centro local.
        if z_axis is None and len(sec.elevations) > 0:
            zz = np.array(sec.elevations, dtype=float)
            z_axis = float(np.nanmin(zz)) if np.isfinite(zz).any() else np.nan
        rows.append({
            "km": float(sec.km),
            "dist_m": float(sec.km) * 1000.0,
            "x": float(sec.center_x),
            "y": float(sec.center_y),
            "cota_dem": z_axis,
            "cota_respaldo": np.nan,
            "diferencia_dem_menos_respaldo": np.nan,
            "pendiente_local_m_m": np.nan,
            "fuente_perfil": "secciones_curvas_v13",
        })
    out = pd.DataFrame(rows).sort_values("dist_m") if rows else pd.DataFrame()
    if not out.empty:
        z = out["cota_dem"].to_numpy(dtype=float)
        d = out["dist_m"].to_numpy(dtype=float)
        slope = np.full(len(out), np.nan)
        for i in range(1, len(out)):
            if np.isfinite(z[i]) and np.isfinite(z[i-1]) and d[i] > d[i-1]:
                slope[i] = (z[i-1] - z[i]) / (d[i] - d[i-1])
        out["pendiente_local_m_m"] = slope
        return out.reset_index(drop=True)
    return fallback_profile if fallback_profile is not None else pd.DataFrame()

def generate_prismatic_sections(line: LineString, profile_df: pd.DataFrame, spacing_m: float, kind: str, bottom_width: float, height: float, side_slope_l: float, side_slope_r: float, manning_n: float) -> List[CrossSection]:
    dists = np.arange(0.0, line.length + 0.001, spacing_m)
    if dists[-1] < line.length:
        dists = np.append(dists, line.length)
    # interp cota fondo desde perfil DEM si existe, si no 0
    if profile_df is not None and not profile_df.empty and np.isfinite(profile_df.get("cota_dem", pd.Series(dtype=float))).sum() > 1:
        ref_d = profile_df["dist_m"].to_numpy(dtype=float)
        ref_z = profile_df["cota_dem"].to_numpy(dtype=float)
        mask = np.isfinite(ref_z)
        z_beds = np.interp(dists, ref_d[mask], ref_z[mask], left=ref_z[mask][0], right=ref_z[mask][-1])
    else:
        z_beds = np.zeros(len(dists))
    sections: List[CrossSection] = []
    for d, zbed in zip(dists, z_beds):
        if kind.lower().startswith("rect"):
            stations = [0.0, bottom_width, bottom_width]
            elevations = [zbed + height, zbed + height, zbed]
            # Mejor como 4 puntos left wall, bottom L, bottom R, right wall
            stations = [0.0, 0.0, bottom_width, bottom_width]
            elevations = [zbed + height, zbed, zbed, zbed + height]
            top_width = bottom_width
        else:
            wl = side_slope_l * height
            wr = side_slope_r * height
            top_width = wl + bottom_width + wr
            stations = [0.0, wl, wl + bottom_width, top_width]
            elevations = [zbed + height, zbed, zbed, zbed + height]
        sec_line = cross_section_line(line, float(d), max(top_width / 2, bottom_width / 2, 1.0))
        center = line.interpolate(float(d))
        sections.append(CrossSection(
            km=float(d) / 1000.0,
            center_x=center.x,
            center_y=center.y,
            left_x=sec_line.coords[0][0],
            left_y=sec_line.coords[0][1],
            right_x=sec_line.coords[-1][0],
            right_y=sec_line.coords[-1][1],
            stations=[float(s) for s in stations],
            elevations=[float(e) for e in elevations],
            source=f"{kind} n={manning_n}",
            status="aceptada",
            notes="Sección prismática generada por parámetros.",
        ))
    return sections


def load_prismatic_sections_excel(uploaded_file, line: LineString, profile_df: pd.DataFrame) -> List[CrossSection]:
    if uploaded_file is None:
        return []
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip().lower() for c in df.columns]
    required = ["km", "tipo_seccion", "ancho_fondo_m", "altura_m"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"El Excel de secciones no contiene columnas: {', '.join(missing)}")
    sections = []
    ref_d = ref_z = None
    if profile_df is not None and not profile_df.empty and np.isfinite(profile_df.get("cota_dem", pd.Series(dtype=float))).sum() > 1:
        ref_d = profile_df["dist_m"].to_numpy(dtype=float)
        ref_z = profile_df["cota_dem"].to_numpy(dtype=float)
        mask = np.isfinite(ref_z)
        ref_d, ref_z = ref_d[mask], ref_z[mask]
    for _, row in df.iterrows():
        km = float(row["km"])
        d = km * 1000.0
        if d < 0 or d > line.length:
            continue
        kind = str(row["tipo_seccion"]).lower()
        b = float(row["ancho_fondo_m"])
        h = float(row["altura_m"])
        zl = float(row.get("talud_izq_hv", 0.0) or 0.0)
        zr = float(row.get("talud_der_hv", 0.0) or 0.0)
        zbed = row.get("cota_fondo", np.nan)
        try:
            zbed = float(zbed)
        except Exception:
            zbed = np.nan
        if not np.isfinite(zbed):
            if ref_d is not None and len(ref_d) > 1:
                zbed = float(np.interp(d, ref_d, ref_z))
            else:
                zbed = 0.0
        n = float(row.get("manning_n", 0.035) or 0.035)
        if kind.startswith("rect"):
            top_width = b
            stations = [0.0, 0.0, b, b]
            elevations = [zbed + h, zbed, zbed, zbed + h]
        else:
            wl = zl * h
            wr = zr * h
            top_width = wl + b + wr
            stations = [0.0, wl, wl + b, top_width]
            elevations = [zbed + h, zbed, zbed, zbed + h]
        sec_line = cross_section_line(line, d, max(top_width / 2, 1.0))
        center = line.interpolate(d)
        sections.append(CrossSection(
            km=km,
            center_x=center.x,
            center_y=center.y,
            left_x=sec_line.coords[0][0],
            left_y=sec_line.coords[0][1],
            right_x=sec_line.coords[-1][0],
            right_y=sec_line.coords[-1][1],
            stations=[float(s) for s in stations],
            elevations=[float(e) for e in elevations],
            source=f"excel_{kind} n={n}",
            status="aceptada",
            notes=str(row.get("observacion", "")),
        ))
    return sections


def evaluate_curve_support(line: LineString, sections: List[CrossSection], curves: List[Dict[str, Any]], corridor_m: float = 1000.0) -> Dict[str, Any]:
    if not curves:
        return {"estado": "sin_curvas", "detalle": "No se cargaron curvas de apoyo.", "secciones_cubiertas_pct": 0.0}
    covered = 0
    checked = 0
    for sec in sections:
        sec_line = LineString([(sec.left_x, sec.left_y), (sec.right_x, sec.right_y)])
        center = Point(sec.center_x, sec.center_y)
        left_hits = 0
        right_hits = 0
        tx, ty = tangent_at_distance(line, sec.km * 1000.0)
        nx, ny = -ty, tx
        for c in curves:
            # Filtro rápido por bbox antes de calcular intersección
            sb = sec_line.bounds
            cb = c["line"].bounds
            if sb[2] < cb[0] or sb[0] > cb[2] or sb[3] < cb[1] or sb[1] > cb[3]:
                continue
            inter = sec_line.intersection(c["line"])
            if inter.is_empty:
                continue
            pts = []
            if inter.geom_type == "Point":
                pts = [inter]
            elif inter.geom_type == "MultiPoint":
                pts = list(inter.geoms)
            elif inter.geom_type in ("LineString", "MultiLineString"):
                pts = [inter.interpolate(0.5, normalized=True)]
            for p in pts:
                side = (p.x - center.x) * nx + (p.y - center.y) * ny
                if side >= 0:
                    left_hits += 1
                else:
                    right_hits += 1
        checked += 1
        if left_hits >= 1 and right_hits >= 1:
            covered += 1
    pct = 100.0 * covered / checked if checked else 0.0
    if pct >= 80:
        estado = "suficiente"
    elif pct >= 40:
        estado = "parcial"
    else:
        estado = "insuficiente"
    return {"estado": estado, "detalle": f"{covered}/{checked} secciones con curvas a ambos lados del eje.", "secciones_cubiertas_pct": pct}


# --------------------------------------------------------------------------------------
# Exportaciones
# --------------------------------------------------------------------------------------
def _kml_coord(lon: float, lat: float, z: Optional[float] = None) -> str:
    if z is None or not np.isfinite(z):
        return f"{lon:.8f},{lat:.8f}"
    return f"{lon:.8f},{lat:.8f},{float(z):.3f}"


def kml_linestring(name: str, coords_lonlat: List[Tuple[float, float, Optional[float]]], style_url: str = "") -> str:
    coord_text = " ".join([_kml_coord(lon, lat, z) for lon, lat, z in coords_lonlat])
    alt_mode = "<altitudeMode>absolute</altitudeMode>" if any(z is not None for _, _, z in coords_lonlat) else ""
    return f"""
    <Placemark><name>{escape(name)}</name>{style_url}<LineString><tessellate>1</tessellate>{alt_mode}<coordinates>{coord_text}</coordinates></LineString></Placemark>"""


def kml_point(name: str, lon: float, lat: float, z: Optional[float] = None, style_url: str = "") -> str:
    coord = _kml_coord(lon, lat, z)
    return f"""
    <Placemark><name>{escape(name)}</name>{style_url}<Point><coordinates>{coord}</coordinates></Point></Placemark>"""


def line_to_lonlatz(line: LineString, to_wgs, z_values: Optional[Sequence[float]] = None) -> List[Tuple[float, float, Optional[float]]]:
    coords = list(line.coords)
    out = []
    for i, (x, y) in enumerate(coords):
        lon, lat = to_wgs.transform(x, y)
        z = None
        if z_values is not None and i < len(z_values):
            z = float(z_values[i]) if np.isfinite(z_values[i]) else None
        out.append((lon, lat, z))
    return out


def make_kmz(axis_full: LineString, axis_useful: LineString, pc_h: Point, pc_s: Point, sections: List[CrossSection], to_wgs, out_name="hidrosed_eje_secciones.kmz") -> bytes:
    styles = """
    <Style id="axis_full"><LineStyle><color>ff777777</color><width>2</width></LineStyle></Style>
    <Style id="axis_useful"><LineStyle><color>ffff0000</color><width>4</width></LineStyle></Style>
    <Style id="pc_h"><IconStyle><color>ff0000ff</color><scale>1.2</scale></IconStyle></Style>
    <Style id="pc_s"><IconStyle><color>ff00ff00</color><scale>1.2</scale></IconStyle></Style>
    <Style id="section_ok"><LineStyle><color>ff00ffff</color><width>2</width></LineStyle></Style>
    <Style id="section_review"><LineStyle><color>ff00a5ff</color><width>2</width></LineStyle></Style>
    """
    placemarks = []
    placemarks.append(kml_linestring("Eje completo", line_to_lonlatz(axis_full, to_wgs), "<styleUrl>#axis_full</styleUrl>"))
    placemarks.append(kml_linestring("Eje útil PC hidrológico - PC cuenca soporte", line_to_lonlatz(axis_useful, to_wgs), "<styleUrl>#axis_useful</styleUrl>"))
    lon, lat = to_wgs.transform(pc_h.x, pc_h.y)
    placemarks.append(kml_point("PC hidrológico proyectado", lon, lat, None, "<styleUrl>#pc_h</styleUrl>"))
    lon, lat = to_wgs.transform(pc_s.x, pc_s.y)
    placemarks.append(kml_point("PC cuenca soporte proyectado", lon, lat, None, "<styleUrl>#pc_s</styleUrl>"))
    for i, sec in enumerate(sections, start=1):
        line = LineString([(sec.left_x, sec.left_y), (sec.right_x, sec.right_y)])
        style = "#section_ok" if sec.status == "aceptada" else "#section_review"
        placemarks.append(kml_linestring(f"XS {i:03d} km {sec.km:.3f} {sec.status}", line_to_lonlatz(line, to_wgs), f"<styleUrl>{style}</styleUrl>"))
    kml = f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>HidroSed Eje Cauce y Secciones</name>{styles}{''.join(placemarks)}</Document></kml>"""
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("doc.kml", kml)
    mem.seek(0)
    return mem.read()


def sections_to_hecras_excel(sections: List[CrossSection]) -> bytes:
    mem = io.BytesIO()
    rows = []
    for sec in sections:
        for st, elev in zip(sec.stations, sec.elevations):
            rows.append({
                "River Station_km": sec.km,
                "Station_m": st,
                "Elevation_m": elev,
                "Source": sec.source,
                "Status": sec.status,
                "Notes": sec.notes,
            })
    df_points = pd.DataFrame(rows)
    df_summary = pd.DataFrame([{
        "km": s.km,
        "center_x": s.center_x,
        "center_y": s.center_y,
        "left_x": s.left_x,
        "left_y": s.left_y,
        "right_x": s.right_x,
        "right_y": s.right_y,
        "source": s.source,
        "status": s.status,
        "notes": s.notes,
    } for s in sections])
    with pd.ExcelWriter(mem, engine="openpyxl") as writer:
        df_summary.to_excel(writer, index=False, sheet_name="secciones_resumen")
        df_points.to_excel(writer, index=False, sheet_name="hecras_station_elevation")
    mem.seek(0)
    return mem.read()


def make_prismatic_template() -> bytes:
    df = pd.DataFrame({
        "km": [0.0, 0.1, 0.2],
        "tipo_seccion": ["trapecial", "trapecial", "rectangular"],
        "ancho_fondo_m": [4.0, 4.5, 3.0],
        "altura_m": [2.0, 2.0, 1.5],
        "talud_izq_HV": [1.5, 1.5, 0.0],
        "talud_der_HV": [1.5, 1.5, 0.0],
        "manning_n": [0.035, 0.035, 0.030],
        "cota_fondo": [np.nan, np.nan, np.nan],
        "observacion": ["ejemplo", "ejemplo", "ejemplo"],
    })
    mem = io.BytesIO()
    with pd.ExcelWriter(mem, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="secciones")
    mem.seek(0)
    return mem.read()


# --------------------------------------------------------------------------------------
# Visualización
# --------------------------------------------------------------------------------------
def plot_axis_sections(axis_full, axis_useful, pc_h, pc_s, sections, curves=None):
    fig, ax = plt.subplots(figsize=(9, 7))
    xf, yf = axis_full.xy
    ax.plot(xf, yf, linewidth=1.5, label="Eje completo")
    xu, yu = axis_useful.xy
    ax.plot(xu, yu, linewidth=3, label="Eje útil")
    ax.scatter([pc_h.x], [pc_h.y], s=60, marker="o", label="PC hidrológico proyectado")
    ax.scatter([pc_s.x], [pc_s.y], s=60, marker="s", label="PC soporte proyectado")
    if curves:
        for c in curves[:200]:
            try:
                x, y = c["line"].xy
                ax.plot(x, y, linewidth=0.5, alpha=0.35)
            except Exception:
                pass
    for sec in sections[:300]:
        ax.plot([sec.left_x, sec.right_x], [sec.left_y, sec.right_y], linewidth=0.8, alpha=0.75)
    ax.set_aspect("equal", adjustable="box")
    ax.grid(True, alpha=0.3)
    ax.legend(loc="best")
    ax.set_title("Eje útil y secciones transversales")
    ax.set_xlabel("X UTM [m]")
    ax.set_ylabel("Y UTM [m]")
    return fig


def plot_profile(profile_df):
    fig, ax = plt.subplots(figsize=(10, 4))
    if "cota_dem" in profile_df:
        ax.plot(profile_df["km"], profile_df["cota_dem"], label="DEM")
    if "cota_respaldo" in profile_df and np.isfinite(profile_df["cota_respaldo"]).sum() > 1:
        ax.plot(profile_df["km"], profile_df["cota_respaldo"], label="Perfil respaldo")
    ax.set_xlabel("km")
    ax.set_ylabel("Cota [m]")
    ax.set_title("Perfil longitudinal del tramo útil")
    ax.grid(True, alpha=0.3)
    ax.legend()
    return fig


def plot_sections_3d(axis_useful, sections):
    if go is None:
        return None
    fig = go.Figure()
    x, y = axis_useful.xy
    fig.add_trace(go.Scatter3d(x=list(x), y=list(y), z=[0] * len(x), mode="lines", name="Eje útil XY"))
    for i, sec in enumerate(sections[:80]):
        xs = np.linspace(sec.left_x, sec.right_x, len(sec.elevations))
        ys = np.linspace(sec.left_y, sec.right_y, len(sec.elevations))
        z = np.array(sec.elevations, dtype=float)
        fig.add_trace(go.Scatter3d(x=xs, y=ys, z=z, mode="lines", name=f"XS {sec.km:.3f} km", showlegend=(i < 5)))
    fig.update_layout(scene=dict(xaxis_title="X", yaxis_title="Y", zaxis_title="Cota"), height=650, title="Vista 3D referencial eje + secciones")
    return fig


# --------------------------------------------------------------------------------------
# Visualización avanzada, edición y estado de la aplicación
# --------------------------------------------------------------------------------------
def section_xyz(sec: CrossSection) -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    """Retorna station, x, y, z de una sección, de izquierda a derecha mirando aguas abajo."""
    st = np.array(sec.stations, dtype=float)
    z = np.array(sec.elevations, dtype=float)
    if len(st) == 0:
        return st, np.array([]), np.array([]), z
    total = max(float(st[-1] - st[0]), 1e-9)
    frac = (st - st[0]) / total
    xs = sec.left_x + frac * (sec.right_x - sec.left_x)
    ys = sec.left_y + frac * (sec.right_y - sec.left_y)
    return st, xs, ys, z


def section_to_edit_df(sec: CrossSection) -> pd.DataFrame:
    return pd.DataFrame({
        "station_m": [float(v) for v in sec.stations],
        "cota_m": [float(v) if np.isfinite(v) else np.nan for v in sec.elevations],
    })


def apply_section_edit(sec: CrossSection, edit_df: pd.DataFrame, status: str, notes: str) -> CrossSection:
    df = edit_df.copy()
    df["station_m"] = pd.to_numeric(df["station_m"], errors="coerce")
    df["cota_m"] = pd.to_numeric(df["cota_m"], errors="coerce")
    df = df.dropna(subset=["station_m", "cota_m"]).sort_values("station_m", kind="mergesort")
    if len(df) < 2:
        raise ValueError("La sección debe mantener al menos dos puntos station-cota.")
    sec.stations = [float(v) for v in df["station_m"].to_list()]
    sec.elevations = [float(v) for v in df["cota_m"].to_list()]
    sec.status = status
    sec.notes = notes
    return sec


def smooth_section(sec: CrossSection, window: int = 3) -> CrossSection:
    z = pd.Series(sec.elevations, dtype="float64")
    if len(z) >= 3:
        zz = z.rolling(window=max(3, int(window)), center=True, min_periods=1).mean().to_numpy(dtype=float)
        # Mantener extremos para no cambiar artificialmente márgenes
        zz[0] = float(sec.elevations[0])
        zz[-1] = float(sec.elevations[-1])
        sec.elevations = [float(v) for v in zz]
        sec.notes = (sec.notes or "") + " | suavizada"
    return sec


def interpolate_section_from_neighbors(sections: List[CrossSection], idx: int) -> Optional[CrossSection]:
    if idx <= 0 or idx >= len(sections) - 1:
        return None
    prev_s = sections[idx - 1]
    next_s = sections[idx + 1]
    sec = sections[idx]
    st = np.array(sec.stations, dtype=float)
    if len(st) < 2:
        return None
    zp = np.interp(st, np.array(prev_s.stations, dtype=float), np.array(prev_s.elevations, dtype=float))
    zn = np.interp(st, np.array(next_s.stations, dtype=float), np.array(next_s.elevations, dtype=float))
    sec.elevations = [float(v) for v in (zp + zn) / 2.0]
    sec.status = "corregida"
    sec.notes = (sec.notes or "") + " | interpolada entre secciones vecinas"
    return sec


def plot_single_section(sec: CrossSection):
    fig, ax = plt.subplots(figsize=(10, 4.8))
    stn = np.array(sec.stations, dtype=float)
    elev = np.array(sec.elevations, dtype=float)
    ok = np.isfinite(stn) & np.isfinite(elev)
    if ok.sum() >= 2:
        stn = stn[ok]
        elev = elev[ok]
        ax.plot(stn, elev, marker="o", linewidth=1.8)
        base = float(np.nanmin(elev)) - max(1.0, 0.04 * max(1.0, float(np.nanmax(elev) - np.nanmin(elev))))
        ax.fill_between(stn, elev, base, alpha=0.12)
        axis_station = (float(np.nanmin(stn)) + float(np.nanmax(stn))) / 2.0
        ax.axvline(axis_station, linestyle="--", linewidth=1.0, alpha=0.75)
        ax.text(axis_station, float(np.nanmax(elev)), " eje", va="top", ha="left")
        ax.text(float(np.nanmin(stn)), base, "Ribera izquierda", va="bottom", ha="left", fontsize=9)
        ax.text(float(np.nanmax(stn)), base, "Ribera derecha", va="bottom", ha="right", fontsize=9)
    ax.set_xlabel("Station [m] · izquierda → derecha mirando aguas abajo")
    ax.set_ylabel("Cota [m]")
    ax.set_title(f"Sección independiente · km {sec.km:.3f} · {sec.status} · {sec.source}")
    ax.grid(True, alpha=0.3)
    return fig


def plot_profile_interactive(profile_df: pd.DataFrame, sections: List[CrossSection]):
    if go is None:
        return None
    fig = go.Figure()
    if profile_df is not None and not profile_df.empty:
        if "cota_dem" in profile_df.columns:
            fig.add_trace(go.Scatter(
                x=profile_df["km"], y=profile_df["cota_dem"], mode="lines", name="Perfil DEM",
                hovertemplate="km=%{x:.3f}<br>cota DEM=%{y:.2f} m<extra></extra>"
            ))
        if "cota_respaldo" in profile_df.columns and np.isfinite(profile_df["cota_respaldo"]).sum() > 1:
            fig.add_trace(go.Scatter(
                x=profile_df["km"], y=profile_df["cota_respaldo"], mode="lines", name="Perfil respaldo",
                hovertemplate="km=%{x:.3f}<br>cota respaldo=%{y:.2f} m<extra></extra>"
            ))
    if sections:
        xs = [s.km for s in sections]
        ys = []
        colors = []
        labels = []
        for s in sections:
            z = np.array(s.elevations, dtype=float)
            ys.append(float(np.nanmin(z)) if np.isfinite(z).any() else np.nan)
            colors.append(s.status)
            labels.append(f"km {s.km:.3f}<br>estado: {s.status}<br>{escape(s.notes or '')}")
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="markers", name="Secciones",
            marker=dict(size=7), text=labels, hovertemplate="%{text}<br>cota fondo/ref=%{y:.2f} m<extra></extra>"
        ))
    fig.update_layout(
        title="Perfil longitudinal interactivo con ubicación de secciones",
        xaxis_title="km tramo útil",
        yaxis_title="Cota [m]",
        height=450,
        hovermode="closest",
    )
    return fig


def _section_offsets_z(sec: CrossSection) -> Tuple[np.ndarray, np.ndarray]:
    stn = np.array(sec.stations, dtype=float)
    z = np.array(sec.elevations, dtype=float)
    ok = np.isfinite(stn) & np.isfinite(z)
    stn = stn[ok]
    z = z[ok]
    if len(stn) == 0:
        return np.array([]), np.array([])
    # centro geométrico en station. Para secciones del DEM empieza en 0 y termina en ancho;
    # para secciones editadas o de curvas, se toma el centro del rango efectivo.
    center = (float(np.nanmin(stn)) + float(np.nanmax(stn))) / 2.0
    return stn - center, z


def plot_sections_3d_interactive(axis_useful: LineString, profile_df: pd.DataFrame, sections: List[CrossSection]):
    """Vista 3D replicando el modelo claro de la app v13.

    Usa coordenadas hidráulicas, no coordenadas UTM planas:
    X = progresiva sobre eje, Y = offset transversal desde eje, Z = cota.
    Esta representación evita que las secciones se vean aplastadas por la escala UTM.
    """
    if go is None:
        return None
    fig = go.Figure()

    # Perfil longitudinal central en y=0.
    if profile_df is not None and not profile_df.empty and {"dist_m", "cota_dem"}.issubset(profile_df.columns):
        p = profile_df.copy()
        p["dist_m"] = pd.to_numeric(p["dist_m"], errors="coerce")
        p["cota_dem"] = pd.to_numeric(p["cota_dem"], errors="coerce")
        p = p.dropna(subset=["dist_m", "cota_dem"]).sort_values("dist_m")
        if len(p) >= 2:
            fig.add_trace(go.Scatter3d(
                x=p["dist_m"], y=[0.0] * len(p), z=p["cota_dem"],
                mode="lines", name="Perfil longitudinal eje", line=dict(width=6),
                hovertemplate="Progresiva=%{x:.2f} m<br>Offset=0 m<br>Cota=%{z:.2f} m<extra></extra>",
            ))

    # Secciones transversales como líneas verticales en la progresiva.
    for i, sec in enumerate(sections[:500]):
        off, z = _section_offsets_z(sec)
        if len(off) < 2:
            continue
        order = np.argsort(off)
        off = off[order]
        z = z[order]
        x = np.full(len(off), float(sec.km) * 1000.0)
        name = f"XS km {sec.km:.3f}"
        fig.add_trace(go.Scatter3d(
            x=x, y=off, z=z, mode="lines+markers", name=name,
            showlegend=(i < 8),
            line=dict(width=4), marker=dict(size=3),
            text=[f"{name}<br>{sec.status}<br>{escape(sec.source)}"] * len(off),
            hovertemplate="%{text}<br>Progresiva=%{x:.2f} m<br>Offset=%{y:.2f} m<br>Cota=%{z:.2f} m<extra></extra>",
        ))

    # Línea de fondo/ref por sección para control longitudinal.
    rows = []
    for sec in sections:
        off, z = _section_offsets_z(sec)
        if len(z) >= 2 and np.isfinite(z).any():
            rows.append({"dist_m": float(sec.km) * 1000.0, "zmin": float(np.nanmin(z)), "status": sec.status})
    if len(rows) >= 2:
        df = pd.DataFrame(rows).sort_values("dist_m")
        fig.add_trace(go.Scatter3d(
            x=df["dist_m"], y=[0.0] * len(df), z=df["zmin"],
            mode="lines+markers", name="Fondo/ref. secciones", line=dict(width=3, dash="dash"),
            hovertemplate="Progresiva=%{x:.2f} m<br>Cota fondo/ref=%{z:.2f} m<extra></extra>",
        ))

    fig.update_layout(
        title="Modelo 3D hidráulico: X=progresiva, Y=offset transversal, Z=cota",
        scene=dict(
            xaxis_title="Progresiva sobre eje [m]",
            yaxis_title="Offset transversal desde eje [m]",
            zaxis_title="Cota [m]",
            aspectmode="manual",
            aspectratio=dict(x=2.4, y=1.0, z=0.65),
        ),
        height=760,
        margin=dict(l=0, r=0, t=45, b=0),
    )
    return fig


def compute_app_result(inputs: Dict[str, Any]) -> Dict[str, Any]:
    axis_file = inputs["axis_file"]
    pc_h_file = inputs["pc_h_file"]
    pc_s_file = inputs["pc_s_file"]
    dem_source = inputs["dem_source"]
    dem_file = inputs.get("dem_file")
    opentopo_api_key = inputs.get("opentopo_api_key", "")
    opentopo_demtype = inputs.get("opentopo_demtype", "COP30")
    opentopo_margin_km = inputs.get("opentopo_margin_km", 3.0)
    profile_file = inputs.get("profile_file")
    curves_file = inputs.get("curves_file")
    sections_excel = inputs.get("sections_excel")
    manual_invert = inputs.get("manual_invert", False)
    profile_step = inputs.get("profile_step", 25.0)
    section_mode = inputs.get("section_mode", "Desde DEM natural")
    section_spacing = inputs.get("section_spacing", 100.0)
    half_width = inputs.get("half_width", 100.0)
    station_step = inputs.get("station_step", 5.0)
    prism_kind = inputs.get("prism_kind", "trapecial")
    bottom_width = inputs.get("bottom_width", 5.0)
    height = inputs.get("height", 2.0)
    talud_l = inputs.get("talud_l", 1.5)
    talud_r = inputs.get("talud_r", 1.5)
    manning_n = inputs.get("manning_n", 0.035)

    dem = None
    try:
        axis_coords, axis_name = pick_longest_line(axis_file, "eje del cauce")
        pc_h_lon, pc_h_lat, pc_h_z, pc_h_name = pick_first_point(pc_h_file, "PC hidrológico")
        pc_s_lon, pc_s_lat, pc_s_z, pc_s_name = pick_first_point(pc_s_file, "PC cuenca soporte")

        crs_proj = determine_local_utm_crs(pc_h_lon, pc_h_lat)
        to_proj, to_wgs = transformers_for_crs(crs_proj)
        axis_proj, axis_zs = coords_to_linestring_projected(axis_coords, to_proj)
        pc_h = point_projected(pc_h_lon, pc_h_lat, to_proj)
        pc_s = point_projected(pc_s_lon, pc_s_lat, to_proj)

        dem_download_info = None
        if dem_source == "Descargar desde OpenTopography":
            west, south, east, north = bbox_from_lonlat_inputs(axis_coords, pc_h_lon, pc_h_lat, pc_s_lon, pc_s_lat, opentopo_margin_km)
            area_bbox = estimate_bbox_area_km2(west, south, east, north)
            dem_download_info = {
                "demtype": opentopo_demtype,
                "west": west, "south": south, "east": east, "north": north,
                "area_bbox_km2_aprox": area_bbox,
                "margen_km": opentopo_margin_km,
            }
            dem_bytes = download_opentopography_dem(opentopo_api_key, opentopo_demtype, west, south, east, north)
            dem_file_runtime = BytesUpload(f"DEM_{opentopo_demtype}_opentopography.tif", dem_bytes)
            dem_size_mb = len(dem_bytes) / 1024 / 1024
        else:
            dem_file_runtime = dem_file
            dem_size_mb = None

        dem = DemSampler(dem_file_runtime) if dem_file_runtime is not None else None
        dem_sampler_func = None
        if dem is not None and dem.dataset is not None:
            dem_sampler_func = lambda x, y: dem.sample_xy(x, y, crs_proj)

        axis_oriented, inverted, direction_note, z_start_raw, z_end_raw = orient_axis_by_dem_or_user(axis_proj, dem_sampler_func, manual_invert)
        d_h = axis_oriented.project(pc_h)
        d_s = axis_oriented.project(pc_s)
        pc_h_on = axis_oriented.interpolate(d_h)
        pc_s_on = axis_oriented.interpolate(d_s)
        dist_pc_h = pc_h.distance(pc_h_on)
        dist_pc_s = pc_s.distance(pc_s_on)
        axis_useful = extract_subline(axis_oriented, d_h, d_s)

        respaldo_df = load_profile_file(profile_file) if profile_file is not None else pd.DataFrame()
        profile_df = generate_longitudinal_profile(axis_useful, dem, crs_proj, step_m=profile_step, respaldo=respaldo_df)
        warnings = []
        warnings.extend(evaluate_profile(profile_df))
        if dist_pc_h > 250:
            warnings.append(f"PC hidrológico está a {dist_pc_h:.1f} m del eje. Revisar proyección.")
        if dist_pc_s > 250:
            warnings.append(f"PC cuenca soporte está a {dist_pc_s:.1f} m del eje. Revisar proyección.")
        if axis_useful.length < 10:
            warnings.append("El tramo útil es demasiado corto.")
        if dem is None and section_mode == "Desde DEM natural":
            raise RuntimeError("Seleccionaste Desde DEM natural, pero no hay DEM disponible.")

        curves = load_support_curves(curves_file, to_proj) if curves_file is not None else []

        # Modo robusto: las curvas de apoyo son opcionales.
        # Si existen, se usa el método v13 por intersección sección-curva.
        # Si no existen, la app cae automáticamente al DEM natural.
        # Si tampoco hay DEM disponible, usa sección prismática por parámetros
        # para que el módulo siga entregando un eje/secciones revisables.
        effective_section_mode = section_mode
        sections = []

        if section_mode == "Automático recomendado":
            if curves:
                effective_section_mode = "Desde curvas de nivel v13"
            elif dem is not None:
                effective_section_mode = "Desde DEM natural"
            elif sections_excel is not None:
                effective_section_mode = "Desde Excel prismático"
            else:
                effective_section_mode = "Prismática por parámetros"
                warnings.append("No se cargaron curvas ni DEM. Se generaron secciones prismáticas por parámetros.")

        if effective_section_mode == "Desde curvas de nivel v13":
            if curves:
                sections = generate_sections_from_contours_v13(axis_useful, curves, section_spacing, half_width)
                if not sections:
                    warnings.append("Las curvas de apoyo fueron cargadas, pero no generaron secciones válidas. Se usará DEM natural si está disponible.")
                    effective_section_mode = "Desde DEM natural" if dem is not None else "Prismática por parámetros"
                else:
                    # Para la vista longitudinal 3D se prioriza la cota interpolada desde las secciones v13.
                    profile_from_sections = longitudinal_profile_from_sections_v13(sections, fallback_profile=profile_df)
                    if profile_from_sections is not None and not profile_from_sections.empty:
                        profile_df = profile_from_sections
            else:
                warnings.append("Modo v13 seleccionado sin curvas de nivel de apoyo. Como las curvas son opcionales, se usará DEM natural si está disponible.")
                effective_section_mode = "Desde DEM natural" if dem is not None else "Prismática por parámetros"

        if effective_section_mode == "Desde DEM natural":
            if dem is None:
                warnings.append("No hay DEM disponible para secciones naturales. Se usará sección prismática por parámetros.")
                effective_section_mode = "Prismática por parámetros"
            else:
                sections = generate_natural_sections(axis_useful, dem, crs_proj, section_spacing, half_width, station_step)
                if not sections:
                    warnings.append("No se pudieron generar secciones naturales desde DEM. Se usará sección prismática por parámetros.")
                    effective_section_mode = "Prismática por parámetros"

        if effective_section_mode == "Prismática por parámetros":
            sections = generate_prismatic_sections(axis_useful, profile_df, section_spacing, prism_kind, bottom_width, height, talud_l, talud_r, manning_n)

        if effective_section_mode == "Desde Excel prismático":
            sections = load_prismatic_sections_excel(sections_excel, axis_useful, profile_df) if sections_excel is not None else []
            if sections_excel is None:
                warnings.append("Modo Excel seleccionado, pero no se cargó Excel de secciones. Se usará sección prismática por parámetros.")
                sections = generate_prismatic_sections(axis_useful, profile_df, section_spacing, prism_kind, bottom_width, height, talud_l, talud_r, manning_n)
                effective_section_mode = "Prismática por parámetros"

        warnings.append(f"Modo efectivo de generación de secciones: {effective_section_mode}.")

        curve_eval = evaluate_curve_support(axis_useful, sections, curves) if sections else {"estado": "sin_secciones", "detalle": "No hay secciones para evaluar.", "secciones_cubiertas_pct": 0.0}

        validation = AxisValidation(
            axis_length_m=float(axis_oriented.length),
            useful_length_m=float(axis_useful.length),
            km_pc_hidrologico=float(d_h / 1000.0),
            km_pc_soporte=float(d_s / 1000.0),
            direction=direction_note,
            inverted=bool(inverted),
            left_right_defined=True,
            warnings=warnings,
        )
        return {
            "axis_oriented": axis_oriented,
            "axis_useful": axis_useful,
            "pc_h_on": pc_h_on,
            "pc_s_on": pc_s_on,
            "sections": sections,
            "profile_df": profile_df,
            "curves": curves,
            "curve_eval": curve_eval,
            "validation": validation,
            "to_wgs": to_wgs,
            "crs_proj": crs_proj,
            "warnings": warnings,
            "axis_name": axis_name,
            "pc_h_name": pc_h_name,
            "pc_s_name": pc_s_name,
            "pc_h_original_lonlat": [pc_h_lon, pc_h_lat],
            "pc_s_original_lonlat": [pc_s_lon, pc_s_lat],
            "dist_pc_h_m": dist_pc_h,
            "dist_pc_s_m": dist_pc_s,
            "section_mode": section_mode,
            "section_spacing": section_spacing,
            "half_width": half_width,
            "station_step": station_step,
            "prism_kind": prism_kind,
            "bottom_width": bottom_width,
            "height": height,
            "talud_l": talud_l,
            "talud_r": talud_r,
            "manning_n": manning_n,
            "dem_source": dem_source,
            "dem_download_info": dem_download_info,
            "dem_size_mb": dem_size_mb,
        }
    finally:
        if dem is not None:
            dem.close()



def rescale_section_width(sec: CrossSection, new_half_width: float) -> CrossSection:
    """Ajusta ancho geométrico de una sección manteniendo su forma relativa station-cota."""
    new_half_width = float(new_half_width)
    if new_half_width <= 0:
        raise ValueError("El semi-ancho debe ser mayor que cero.")
    dx = float(sec.right_x - sec.left_x)
    dy = float(sec.right_y - sec.left_y)
    L = math.hypot(dx, dy)
    if L <= 1e-9:
        raise ValueError("La sección no tiene geometría transversal válida.")
    ux, uy = dx / L, dy / L
    cx, cy = float(sec.center_x), float(sec.center_y)
    sec.left_x = cx - ux * new_half_width
    sec.left_y = cy - uy * new_half_width
    sec.right_x = cx + ux * new_half_width
    sec.right_y = cy + uy * new_half_width
    st = np.array(sec.stations, dtype=float)
    z = np.array(sec.elevations, dtype=float)
    if len(st) >= 2 and np.isfinite(st).any():
        old_min, old_max = float(np.nanmin(st)), float(np.nanmax(st))
        old_half = max((old_max - old_min) / 2.0, 1e-9)
        old_center = (old_min + old_max) / 2.0
        rel = np.clip((st - old_center) / old_half, -1.0, 1.0)
        sec.stations = [float((r + 1.0) * new_half_width) for r in rel]
    else:
        n = max(len(z), 2)
        sec.stations = [float(v) for v in np.linspace(0.0, 2.0 * new_half_width, n)]
    sec.status = "corregida"
    sec.notes = (sec.notes or "") + f" | semi-ancho ajustado a {new_half_width:.1f} m"
    return sec


def apply_batch_section_improvement(sections: List[CrossSection], km_from: float, km_to: float,
                                    new_half_width: Optional[float] = None,
                                    smooth: bool = False,
                                    status: Optional[str] = None,
                                    note: str = "") -> Tuple[List[CrossSection], int]:
    """Aplica mejora por tramo entre km_from y km_to."""
    a, b = sorted([float(km_from), float(km_to)])
    count = 0
    for i, sec in enumerate(sections):
        if a <= float(sec.km) <= b:
            if new_half_width is not None and float(new_half_width) > 0:
                sec = rescale_section_width(sec, float(new_half_width))
            if smooth:
                sec = smooth_section(sec)
            if status:
                sec.status = status
            if note:
                sec.notes = (sec.notes or "") + " | " + note
            sections[i] = sec
            count += 1
    return sections, count



def _clone_section(sec: CrossSection) -> CrossSection:
    return CrossSection(
        km=float(sec.km),
        center_x=float(sec.center_x), center_y=float(sec.center_y),
        left_x=float(sec.left_x), left_y=float(sec.left_y),
        right_x=float(sec.right_x), right_y=float(sec.right_y),
        stations=[float(v) for v in sec.stations],
        elevations=[float(v) for v in sec.elevations],
        source=str(sec.source), status=str(sec.status), notes=str(sec.notes or ""),
    )


def _section_profile_normalized(sec: CrossSection) -> Tuple[np.ndarray, np.ndarray, float]:
    """Devuelve coordenada normalizada 0-1, cotas y semi-ancho de una sección."""
    st = np.array(sec.stations, dtype=float)
    z = np.array(sec.elevations, dtype=float)
    ok = np.isfinite(st) & np.isfinite(z)
    st, z = st[ok], z[ok]
    if len(st) < 2:
        raise ValueError("Sección sin puntos suficientes para interpolar.")
    order = np.argsort(st)
    st, z = st[order], z[order]
    st_min, st_max = float(st[0]), float(st[-1])
    width = max(st_max - st_min, 1e-9)
    t = (st - st_min) / width
    # eliminar estaciones normalizadas duplicadas
    tu, idx = np.unique(np.round(t, 8), return_index=True)
    zu = z[idx]
    half = width / 2.0
    return tu.astype(float), zu.astype(float), float(half)


def _interpolate_section_at_km_from_brackets(axis_useful: LineString,
                                             ref_sections: List[CrossSection],
                                             target_km: float,
                                             station_step_m: float,
                                             fixed_half_width_m: Optional[float] = None) -> Optional[CrossSection]:
    """Crea una sección interpolada en target_km a partir de las secciones vecinas aguas arriba/abajo."""
    refs = sorted([s for s in ref_sections if s.stations and s.elevations], key=lambda s: float(s.km))
    if len(refs) < 2:
        return None
    km = float(target_km)
    prev_sec = None
    next_sec = None
    for s in refs:
        if float(s.km) <= km:
            prev_sec = s
        if float(s.km) >= km and next_sec is None:
            next_sec = s
    if prev_sec is None or next_sec is None:
        return None

    # Si coincide exactamente con una sección existente, usarla como base pero reubicar geometría al eje.
    if abs(float(next_sec.km) - float(prev_sec.km)) < 1e-9:
        base = _clone_section(prev_sec)
        d = min(max(km * 1000.0, 0.0), float(axis_useful.length))
        t0, z0, hw0 = _section_profile_normalized(base)
        hw = float(fixed_half_width_m) if fixed_half_width_m and fixed_half_width_m > 0 else hw0
        step = max(float(station_step_m or 5.0), 0.5)
        n = max(2, int(math.ceil((2.0 * hw) / step)) + 1)
        stations = np.linspace(0.0, 2.0 * hw, n)
        tn = stations / max(2.0 * hw, 1e-9)
        elev = np.interp(tn, t0, z0)
        sec_line = cross_section_line(axis_useful, d, hw)
        center = axis_useful.interpolate(d)
        base.km = km
        base.center_x, base.center_y = center.x, center.y
        base.left_x, base.left_y = sec_line.coords[0][0], sec_line.coords[0][1]
        base.right_x, base.right_y = sec_line.coords[-1][0], sec_line.coords[-1][1]
        base.stations = [float(v) for v in stations]
        base.elevations = [float(v) for v in elev]
        base.source = "Sección interpolada entre secciones"
        base.status = "corregida"
        base.notes = (base.notes or "") + " | reconstruida en malla regular por interpolación entre secciones"
        return base

    k0, k1 = float(prev_sec.km), float(next_sec.km)
    w = (km - k0) / max(k1 - k0, 1e-9)
    w = float(np.clip(w, 0.0, 1.0))
    t0, z0, hw0 = _section_profile_normalized(prev_sec)
    t1, z1, hw1 = _section_profile_normalized(next_sec)
    hw = float(fixed_half_width_m) if fixed_half_width_m and fixed_half_width_m > 0 else (1.0 - w) * hw0 + w * hw1
    hw = max(float(hw), 1.0)
    step = max(float(station_step_m or 5.0), 0.5)
    n = max(2, int(math.ceil((2.0 * hw) / step)) + 1)
    stations = np.linspace(0.0, 2.0 * hw, n)
    tn = stations / max(2.0 * hw, 1e-9)
    z_prev = np.interp(tn, t0, z0)
    z_next = np.interp(tn, t1, z1)
    elev = (1.0 - w) * z_prev + w * z_next

    d = min(max(km * 1000.0, 0.0), float(axis_useful.length))
    sec_line = cross_section_line(axis_useful, d, hw)
    center = axis_useful.interpolate(d)
    return CrossSection(
        km=km,
        center_x=center.x, center_y=center.y,
        left_x=sec_line.coords[0][0], left_y=sec_line.coords[0][1],
        right_x=sec_line.coords[-1][0], right_y=sec_line.coords[-1][1],
        stations=[float(v) for v in stations],
        elevations=[float(v) for v in elev],
        source="Sección interpolada entre secciones",
        status="corregida",
        notes=f"Interpolada entre km {k0:.3f} y km {k1:.3f}, separación definida por usuario.",
    )


def interpolate_sections_by_spacing(sections: List[CrossSection], axis_useful: LineString,
                                    km_from: float, km_to: float,
                                    spacing_m: float, station_step_m: float,
                                    fixed_half_width_m: Optional[float] = None,
                                    replace_existing: bool = True) -> Tuple[List[CrossSection], int, int]:
    """Genera secciones interpoladas entre secciones existentes con una separación fija.

    Si replace_existing=True, reemplaza todas las secciones del tramo por una nueva serie regular.
    Si False, agrega solo las intermedias y conserva las existentes.
    """
    if not sections:
        return sections, 0, 0
    a, b = sorted([float(km_from), float(km_to)])
    sep_km = max(float(spacing_m), 1.0) / 1000.0
    ref_sections = sorted([_clone_section(s) for s in sections], key=lambda s: float(s.km))

    targets = []
    k = a
    while k <= b + 1e-9:
        targets.append(round(k, 6))
        k += sep_km
    if not targets or abs(targets[-1] - b) > max(sep_km * 0.25, 1e-6):
        targets.append(round(b, 6))
    targets = sorted(set(targets))

    new_secs: List[CrossSection] = []
    skipped = 0
    for km in targets:
        sec = _interpolate_section_at_km_from_brackets(axis_useful, ref_sections, km, station_step_m, fixed_half_width_m)
        if sec is None:
            skipped += 1
        else:
            new_secs.append(sec)

    if replace_existing:
        out = [s for s in sections if not (a <= float(s.km) <= b)] + new_secs
    else:
        out = list(sections)
        existing_kms = np.array([float(s.km) for s in out], dtype=float) if out else np.array([])
        tol = max(sep_km * 0.20, 0.0005)
        added = []
        for ns in new_secs:
            if existing_kms.size and float(np.min(np.abs(existing_kms - float(ns.km)))) <= tol:
                continue
            added.append(ns)
        new_secs = added
        out.extend(new_secs)
    out = sorted(out, key=lambda s: float(s.km))
    return out, len(new_secs), skipped




def _interpolate_bed_from_profile_or_section(profile_df: pd.DataFrame, sec: CrossSection) -> float:
    """Obtiene cota de fondo por perfil longitudinal o por mínimo de la sección actual."""
    d = float(sec.km) * 1000.0
    if profile_df is not None and not profile_df.empty and {"dist_m", "cota_dem"}.issubset(profile_df.columns):
        ref = profile_df[["dist_m", "cota_dem"]].copy()
        ref["dist_m"] = pd.to_numeric(ref["dist_m"], errors="coerce")
        ref["cota_dem"] = pd.to_numeric(ref["cota_dem"], errors="coerce")
        ref = ref.dropna().sort_values("dist_m")
        if len(ref) >= 2:
            return float(np.interp(d, ref["dist_m"].to_numpy(dtype=float), ref["cota_dem"].to_numpy(dtype=float)))
    z = np.array(sec.elevations, dtype=float)
    if len(z) and np.isfinite(z).any():
        return float(np.nanmin(z))
    return 0.0


def _contour_hits_on_section(sec_line: LineString, curves: List[Dict[str, Any]]) -> List[Tuple[float, float]]:
    """Retorna pares station-cota donde la sección cruza curvas de nivel."""
    if not curves:
        return []
    width = float(sec_line.length)
    hits: List[Tuple[float, float]] = []
    for c in curves:
        if c.get("elev") is None or c.get("line") is None:
            continue
        try:
            sb = sec_line.bounds
            cb = c["line"].bounds
            if sb[2] < cb[0] or sb[0] > cb[2] or sb[3] < cb[1] or sb[1] > cb[3]:
                continue
            inter = sec_line.intersection(c["line"])
        except Exception:
            continue
        for pt in _intersection_points_for_profile(inter):
            sta = float(sec_line.project(pt))
            if -1e-6 <= sta <= width + 1e-6:
                hits.append((sta, float(c["elev"])))
    hits.sort(key=lambda t: (t[0], t[1]))
    clean: List[Tuple[float, float]] = []
    for sta, elev in hits:
        if clean and abs(sta - clean[-1][0]) <= 0.10:
            # Si hay dos cotas casi en la misma estación, conservar promedio simple.
            old_sta, old_elev = clean[-1]
            clean[-1] = ((old_sta + sta) / 2.0, (old_elev + elev) / 2.0)
        else:
            clean.append((sta, elev))
    return clean


def rebuild_section_by_contour_interpolation(axis_useful: LineString, sec: CrossSection,
                                             curves: List[Dict[str, Any]],
                                             half_width_m: Optional[float],
                                             station_step_m: float) -> Tuple[CrossSection, bool, str]:
    """Reconstruye una sección interpolando progresivamente entre curvas de nivel.

    A diferencia del modo v13 básico, que solo deja los puntos de cruce sección-curva,
    esta mejora densifica el perfil entre curvas sucesivas. Así la sección queda ajustada
    progresivamente entre dos cotas conocidas, evitando perfiles quebrados o mal graficados.
    """
    if not curves:
        return sec, False, "no hay curvas de apoyo cargadas"
    d = float(sec.km) * 1000.0
    current_half = None
    try:
        current_half = (float(max(sec.stations)) - float(min(sec.stations))) / 2.0 if sec.stations else None
    except Exception:
        current_half = None
    hw = float(half_width_m) if half_width_m and half_width_m > 0 else float(current_half or 100.0)
    step = max(float(station_step_m or 5.0), 0.5)
    sec_line = cross_section_line(axis_useful, d, hw)
    width = float(sec_line.length)
    hits = _contour_hits_on_section(sec_line, curves)
    if len(hits) < 2:
        sec.status = "revisar"
        sec.notes = (sec.notes or "") + f" | interpolación por curvas no aplicada: solo {len(hits)} cruce(s)"
        return sec, False, f"solo {len(hits)} cruce(s) con curvas"

    hit_s = np.array([h[0] for h in hits], dtype=float)
    hit_z = np.array([h[1] for h in hits], dtype=float)
    order = np.argsort(hit_s)
    hit_s, hit_z = hit_s[order], hit_z[order]

    # Construcción progresiva: puntos regulares + puntos exactos de curvas.
    regular = np.arange(0.0, width + 0.001, step)
    if regular.size == 0 or regular[-1] < width:
        regular = np.append(regular, width)
    stations = np.unique(np.round(np.concatenate([regular, hit_s]), 3))
    # Interpola entre curvas. Fuera del rango de curvas, mantiene la cota de la curva extrema
    # y lo deja advertido como extrapolación conservadora.
    elevations = np.interp(stations, hit_s, hit_z, left=hit_z[0], right=hit_z[-1])

    center = axis_useful.interpolate(d)
    sec.left_x = sec_line.coords[0][0]
    sec.left_y = sec_line.coords[0][1]
    sec.right_x = sec_line.coords[-1][0]
    sec.right_y = sec_line.coords[-1][1]
    sec.center_x = center.x
    sec.center_y = center.y
    sec.stations = [float(v) for v in stations]
    sec.elevations = [float(v) for v in elevations]
    sec.source = "Curvas interpoladas entre cotas"
    sec.status = "corregida"
    outside = ""
    if hit_s[0] > 0.5 or hit_s[-1] < width - 0.5:
        outside = " Bordes fuera del rango de curvas extrapolados con cota extrema; revisar en imagen/perfil."
        sec.status = "revisar"
    sec.notes = (sec.notes or "") + f" | sección reconstruida interpolando entre {len(hits)} cruces de curvas.{outside}"
    return sec, True, f"{len(hits)} cruces interpolados"


def apply_contour_interpolation_by_km(sections: List[CrossSection], axis_useful: LineString,
                                      curves: List[Dict[str, Any]], km_from: float, km_to: float,
                                      half_width_m: Optional[float], station_step_m: float,
                                      note: str = "") -> Tuple[List[CrossSection], int, int]:
    """Aplica interpolación progresiva entre curvas a secciones dentro de un tramo."""
    a, b = sorted([float(km_from), float(km_to)])
    ok_count = 0
    fail_count = 0
    for i, sec in enumerate(sections):
        if a <= float(sec.km) <= b:
            new_sec, ok, detail = rebuild_section_by_contour_interpolation(axis_useful, sec, curves, half_width_m, station_step_m)
            if note:
                new_sec.notes = (new_sec.notes or "") + " | " + note
            new_sec.notes = (new_sec.notes or "") + f" | {detail}"
            sections[i] = new_sec
            ok_count += 1 if ok else 0
            fail_count += 0 if ok else 1
    return sections, ok_count, fail_count


def apply_prismatic_channel_by_km(sections: List[CrossSection], axis_useful: LineString,
                                  profile_df: pd.DataFrame, km_from: float, km_to: float,
                                  kind: str, bottom_width: float, height: float,
                                  side_slope_l: float, side_slope_r: float,
                                  manning_n: float, note: str = "") -> Tuple[List[CrossSection], int]:
    """Reemplaza las secciones de un tramo por canal rectangular/trapecial progresivo.

    La cota de fondo se obtiene desde el perfil longitudinal disponible; si no existe,
    usa el mínimo de la sección actual. Mantiene la progresiva de cada sección existente.
    """
    a, b = sorted([float(km_from), float(km_to)])
    count = 0
    kind_l = str(kind).lower()
    for i, sec in enumerate(sections):
        if not (a <= float(sec.km) <= b):
            continue
        d = float(sec.km) * 1000.0
        zbed = _interpolate_bed_from_profile_or_section(profile_df, sec)
        if kind_l.startswith("rect"):
            top_width = float(bottom_width)
            stations = [0.0, 0.0, top_width, top_width]
            elevations = [zbed + height, zbed, zbed, zbed + height]
            source = f"canal rectangular tramo n={manning_n:.3f}"
        else:
            wl = float(side_slope_l) * float(height)
            wr = float(side_slope_r) * float(height)
            top_width = wl + float(bottom_width) + wr
            stations = [0.0, wl, wl + float(bottom_width), top_width]
            elevations = [zbed + height, zbed, zbed, zbed + height]
            source = f"canal trapecial tramo n={manning_n:.3f}"
        sec_line = cross_section_line(axis_useful, d, max(top_width / 2.0, 1.0))
        center = axis_useful.interpolate(d)
        sec.center_x = center.x
        sec.center_y = center.y
        sec.left_x = sec_line.coords[0][0]
        sec.left_y = sec_line.coords[0][1]
        sec.right_x = sec_line.coords[-1][0]
        sec.right_y = sec_line.coords[-1][1]
        sec.stations = [float(v) for v in stations]
        sec.elevations = [float(v) for v in elevations]
        sec.source = source
        sec.status = "corregida"
        sec.notes = (sec.notes or "") + f" | reemplazada por {source} entre km {a:.3f}-{b:.3f}. {note}".strip()
        sections[i] = sec
        count += 1
    return sections, count

def _latlon_line_from_xy(line: LineString, to_wgs) -> List[List[float]]:
    out = []
    try:
        for x, y in list(line.coords):
            lon, lat = to_wgs.transform(float(x), float(y))
            out.append([float(lat), float(lon)])
    except Exception:
        return []
    return out


def _latlon_section(sec: CrossSection, to_wgs) -> List[List[float]]:
    out = []
    try:
        for x, y in [(sec.left_x, sec.left_y), (sec.center_x, sec.center_y), (sec.right_x, sec.right_y)]:
            lon, lat = to_wgs.transform(float(x), float(y))
            out.append([float(lat), float(lon)])
    except Exception:
        return []
    return out


def render_satellite_reference(axis_useful: LineString, sections: List[CrossSection], pc_h: Point, pc_s: Point, to_wgs, height: int = 720):
    """Mapa satelital en Leaflet con eje útil y secciones."""
    axis_ll = _latlon_line_from_xy(axis_useful, to_wgs)
    sec_ll = []
    for sec in sections[:800]:
        pts = _latlon_section(sec, to_wgs)
        if len(pts) == 3:
            sec_ll.append({"km": float(sec.km), "status": sec.status, "pts": pts})
    points = []
    for name, pt in [("PC hidrológico", pc_h), ("PC cuenca soporte", pc_s)]:
        try:
            lon, lat = to_wgs.transform(float(pt.x), float(pt.y))
            points.append({"name": name, "lat": float(lat), "lon": float(lon)})
        except Exception:
            pass
    payload = json.dumps({"axis": axis_ll, "sections": sec_ll, "points": points}, ensure_ascii=False)
    html_template = """
<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8' />
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<link rel='stylesheet' href='https://unpkg.com/leaflet@1.9.4/dist/leaflet.css' />
<script src='https://unpkg.com/leaflet@1.9.4/dist/leaflet.js'></script>
<style>
  html, body, #map { height: HEIGHTPXpx; margin: 0; padding: 0; }
  .legend { background: rgba(255,255,255,0.9); padding: 8px 10px; border-radius: 6px; font: 12px Arial; }
</style>
</head>
<body>
<div id='map'></div>
<script>
const data = PAYLOAD_JSON;
const map = L.map('map', {preferCanvas: true});
const img = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
  maxZoom: 19,
  attribution: 'Tiles © Esri — imagen satelital de referencia. Abrir KMZ en Google Earth para revisión final.'
}).addTo(map);
const osm = L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', {maxZoom: 19, attribution: '© OpenStreetMap'});
L.control.layers({'Satélite referencia': img, 'OpenStreetMap': osm}, null, {collapsed: false}).addTo(map);
let bounds = [];
if (data.axis && data.axis.length > 1) {
  L.polyline(data.axis, {color: '#ffd400', weight: 5, opacity: 0.95}).addTo(map).bindPopup('Eje útil del cauce');
  bounds = bounds.concat(data.axis);
}
(data.sections || []).forEach(s => {
  const color = s.status === 'rechazada' ? '#ff3333' : (s.status === 'corregida' ? '#00ff66' : '#00e5ff');
  L.polyline([s.pts[0], s.pts[2]], {color: color, weight: 2, opacity: 0.85}).addTo(map)
    .bindPopup('Sección km ' + s.km.toFixed(3) + '<br>Estado: ' + s.status);
  L.circleMarker(s.pts[1], {radius: 2, color: '#ffffff', fillColor:'#000000', fillOpacity:0.8, weight:1}).addTo(map);
  bounds = bounds.concat([s.pts[0], s.pts[2]]);
});
(data.points || []).forEach(p => {
  L.marker([p.lat, p.lon]).addTo(map).bindPopup(p.name);
  bounds.push([p.lat, p.lon]);
});
const legend = L.control({position: 'bottomleft'});
legend.onAdd = function() {
  const div = L.DomUtil.create('div', 'legend');
  div.innerHTML = '<b>Referencia satelital</b><br><span style="color:#ffd400">━━</span> Eje útil<br><span style="color:#00e5ff">━━</span> Sección aceptada/revisar<br><span style="color:#00ff66">━━</span> Sección corregida<br><span style="color:#ff3333">━━</span> Sección rechazada';
  return div;
};
legend.addTo(map);
if (bounds.length > 0) { map.fitBounds(bounds, {padding: [25,25]}); } else { map.setView([-30.0, -71.0], 10); }
</script>
</body>
</html>
"""
    html = html_template.replace("HEIGHTPX", str(int(height))).replace("PAYLOAD_JSON", payload)
    components.html(html, height=height + 20, scrolling=False)

def render_results(result: Dict[str, Any]):
    axis_oriented = result["axis_oriented"]
    axis_useful = result["axis_useful"]
    pc_h_on = result["pc_h_on"]
    pc_s_on = result["pc_s_on"]
    sections = result["sections"]
    profile_df = result["profile_df"]
    curves = result["curves"]
    curve_eval = result["curve_eval"]
    validation = result["validation"]
    to_wgs = result["to_wgs"]
    crs_proj = result["crs_proj"]
    warnings = result["warnings"]

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Etapa 1 · Eje",
        "Etapa 2 · Secciones",
        "Ventana sección por km",
        "3D interactivo / Corrección",
        "Imagen satelital / ancho",
        "Descargas",
    ])

    with tab1:
        st.subheader("Validación del eje del cauce")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Longitud eje completo", f"{axis_oriented.length/1000:.3f} km")
        col2.metric("Longitud tramo útil", f"{axis_useful.length/1000:.3f} km")
        col3.metric("PC hidrológico sobre eje", f"km {validation.km_pc_hidrologico:.3f}")
        col4.metric("PC soporte sobre eje", f"km {validation.km_pc_soporte:.3f}")
        st.write(validation.direction)
        st.write("**Ribera izquierda/derecha:** definida mirando en sentido aguas arriba → aguas abajo del eje orientado.")
        if result.get("dem_download_info"):
            st.caption(f"DEM descargado: {result['dem_download_info']['demtype']} · {result.get('dem_size_mb', 0):.2f} MB · bbox aprox. {result['dem_download_info']['area_bbox_km2_aprox']:,.1f} km²")
        if warnings:
            st.warning("\n".join([f"- {w}" for w in warnings]))
        else:
            st.success("Eje útil validado sin advertencias críticas.")
        st.pyplot(plot_axis_sections(axis_oriented, axis_useful, pc_h_on, pc_s_on, sections, curves))
        st.subheader("Perfil longitudinal")
        st.pyplot(plot_profile(profile_df))
        figp = plot_profile_interactive(profile_df, sections)
        if figp is not None:
            st.plotly_chart(figp, use_container_width=True)
        st.dataframe(profile_df.head(300), use_container_width=True)

    with tab2:
        st.subheader("Resumen de secciones transversales")
        if not sections:
            st.error("No se generaron secciones. Revisa modo de generación e insumos.")
        else:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("N° secciones", len(sections))
            c2.metric("Aceptadas", sum(1 for s in sections if s.status == "aceptada"))
            c3.metric("Corregidas", sum(1 for s in sections if s.status == "corregida"))
            c4.metric("Revisar/Rechazadas", sum(1 for s in sections if s.status not in ["aceptada", "corregida"]))
            st.write(f"**Evaluación curvas de apoyo:** {curve_eval['estado']} · {curve_eval['detalle']}")
            df_secs = pd.DataFrame([{
                "índice": i,
                "km": s.km,
                "source": s.source,
                "status": s.status,
                "n_puntos": len(s.stations),
                "cota_min": float(np.nanmin(s.elevations)) if np.isfinite(np.array(s.elevations, dtype=float)).any() else np.nan,
                "cota_max": float(np.nanmax(s.elevations)) if np.isfinite(np.array(s.elevations, dtype=float)).any() else np.nan,
                "notes": s.notes,
            } for i, s in enumerate(sections)])
            st.dataframe(df_secs, use_container_width=True, height=450)

    with tab3:
        st.subheader("Ventana independiente de sección por km")
        if not sections:
            st.info("No hay secciones para visualizar.")
        else:
            labels = [f"km {s.km:.3f} · {s.status} · #{i}" for i, s in enumerate(sections)]
            current_idx = int(st.selectbox("Seleccionar sección por km", options=list(range(len(sections))), format_func=lambda i: labels[i], key="section_select_idx"))
            sec = sections[current_idx]
            col_a, col_b = st.columns([1.2, 1.0])
            with col_a:
                st.pyplot(plot_single_section(sec))
                st.caption("La estación va de ribera izquierda a ribera derecha, mirando aguas abajo.")
            with col_b:
                st.markdown(f"**Sección km {sec.km:.3f}**")
                st.write(f"Fuente: `{sec.source}`")
                new_status = st.selectbox("Estado", ["aceptada", "corregida", "revisar", "rechazada"], index=["aceptada", "corregida", "revisar", "rechazada"].index(sec.status) if sec.status in ["aceptada", "corregida", "revisar", "rechazada"] else 2, key=f"status_{current_idx}")
                new_notes = st.text_area("Observación", value=sec.notes or "", key=f"notes_{current_idx}", height=90)
                edit_df = st.data_editor(
                    section_to_edit_df(sec),
                    num_rows="dynamic",
                    use_container_width=True,
                    key=f"editor_section_{current_idx}",
                    column_config={
                        "station_m": st.column_config.NumberColumn("Station [m]", step=0.1),
                        "cota_m": st.column_config.NumberColumn("Cota [m]", step=0.01),
                    },
                )
                c1, c2, c3, c4 = st.columns(4)
                if c1.button("Aplicar edición", key=f"apply_{current_idx}"):
                    try:
                        sections[current_idx] = apply_section_edit(sec, edit_df, new_status, new_notes)
                        result["sections"] = sections
                        st.session_state["eje_result"] = result
                        st.success("Sección actualizada.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))
                if c2.button("Suavizar", key=f"smooth_{current_idx}"):
                    sections[current_idx] = smooth_section(sec)
                    result["sections"] = sections
                    st.session_state["eje_result"] = result
                    st.rerun()
                if c3.button("Interpolar", key=f"interp_{current_idx}", disabled=(current_idx == 0 or current_idx == len(sections)-1)):
                    out = interpolate_section_from_neighbors(sections, current_idx)
                    if out is not None:
                        sections[current_idx] = out
                        result["sections"] = sections
                        st.session_state["eje_result"] = result
                        st.rerun()
                if c4.button("Rechazar", key=f"reject_{current_idx}"):
                    sec.status = "rechazada"
                    sec.notes = (sec.notes or "") + " | rechazada por revisión"
                    sections[current_idx] = sec
                    result["sections"] = sections
                    st.session_state["eje_result"] = result
                    st.rerun()

            st.divider()
            st.subheader("Mejorar secciones por tramo km")
            st.caption("Permite corregir en bloque las secciones entre un km inicial y un km final. Útil cuando el ancho del cauce cambia en un tramo visible en imagen satelital.")
            min_km = float(min(s.km for s in sections))
            max_km = float(max(s.km for s in sections))
            b1, b2, b3 = st.columns(3)
            km_from = b1.number_input("km inicial", min_value=min_km, max_value=max_km, value=min_km, step=0.001, format="%.3f", key="batch_km_from_tab3")
            km_to = b2.number_input("km final", min_value=min_km, max_value=max_km, value=max_km, step=0.001, format="%.3f", key="batch_km_to_tab3")
            default_hw = float(max((max(sec.stations) - min(sec.stations)) / 2.0 if sec.stations else 100.0, 1.0))
            batch_half_width = b3.number_input("Nuevo semi-ancho [m]", min_value=1.0, max_value=5000.0, value=default_hw, step=5.0, key="batch_half_tab3")
            bb1, bb2, bb3 = st.columns(3)
            batch_smooth = bb1.checkbox("Suavizar cotas del tramo", value=False, key="batch_smooth_tab3")
            batch_status = bb2.selectbox("Estado a asignar", ["corregida", "aceptada", "revisar", "rechazada"], index=0, key="batch_status_tab3")
            batch_note = bb3.text_input("Nota de corrección", value="corrección por tramo", key="batch_note_tab3")
            if st.button("Aplicar mejora al tramo seleccionado", key="batch_apply_tab3"):
                sections, nedit = apply_batch_section_improvement(sections, km_from, km_to, batch_half_width, batch_smooth, batch_status, batch_note)
                result["sections"] = sections
                st.session_state["eje_result"] = result
                st.success(f"Se actualizaron {nedit} secciones entre km {min(km_from, km_to):.3f} y km {max(km_from, km_to):.3f}.")
                st.rerun()

            with st.expander("Interpolar entre secciones con separación fija", expanded=True):
                st.markdown(
                    """
Esta herramienta genera una nueva serie de secciones **entre dos km**, usando como referencia las
secciones existentes aguas arriba y aguas abajo. Es diferente a interpolar entre curvas de nivel:
aquí la forma de la sección cambia progresivamente entre secciones vecinas y se crea con una
**separación longitudinal definida**.
"""
                )
                is1, is2, is3, is4 = st.columns(4)
                km_from_is = is1.number_input("km inicial", min_value=min_km, max_value=max_km, value=km_from, step=0.001, format="%.3f", key="is_km_from")
                km_to_is = is2.number_input("km final", min_value=min_km, max_value=max_km, value=km_to, step=0.001, format="%.3f", key="is_km_to")
                sep_is = is3.number_input("Separación entre secciones [m]", min_value=1.0, max_value=2000.0, value=float(result.get("section_spacing", 100.0)), step=5.0, key="is_sep")
                stp_is = is4.number_input("Paso station transversal [m]", min_value=0.5, max_value=100.0, value=float(result.get("station_step", 5.0)), step=0.5, key="is_station_step")
                is5, is6, is7 = st.columns(3)
                use_fixed_hw = is5.checkbox("Usar semi-ancho fijo", value=False, key="is_fixed_hw_on")
                fixed_hw = is6.number_input("Semi-ancho fijo [m]", min_value=1.0, max_value=5000.0, value=batch_half_width, step=5.0, key="is_fixed_hw", disabled=not use_fixed_hw)
                replace_is = is7.checkbox("Reemplazar secciones existentes del tramo", value=True, key="is_replace")
                if st.button("Interpolar secciones del tramo", key="is_apply"):
                    sections, nnew, nskip = interpolate_sections_by_spacing(
                        sections, axis_useful, km_from_is, km_to_is, sep_is, stp_is,
                        fixed_half_width_m=(fixed_hw if use_fixed_hw else None),
                        replace_existing=replace_is,
                    )
                    result["sections"] = sections
                    st.session_state["eje_result"] = result
                    st.success(f"Se generaron {nnew} secciones interpoladas con separación {sep_is:.1f} m. Omitidas: {nskip}.")
                    st.rerun()

            with st.expander("Mejora avanzada: interpolar sección entre curvas de nivel", expanded=False):
                st.markdown(
                    """
Esta herramienta reconstruye las secciones del tramo usando los cruces con curvas de nivel y
rellena el perfil **interpolando progresivamente entre dos curvas sucesivas**. Es útil cuando
una sección queda quebrada, con pocos puntos o mal definida por la topografía de apoyo.
"""
                )
                if not curves:
                    st.warning("No hay curvas de nivel de apoyo cargadas. Para usar esta mejora debes cargar KMZ/KML/DXF de curvas en el panel lateral.")
                ci1, ci2, ci3, ci4 = st.columns(4)
                km_from_ci = ci1.number_input("km inicial", min_value=min_km, max_value=max_km, value=km_from, step=0.001, format="%.3f", key="ci_km_from")
                km_to_ci = ci2.number_input("km final", min_value=min_km, max_value=max_km, value=km_to, step=0.001, format="%.3f", key="ci_km_to")
                ci_half = ci3.number_input("Semi-ancho para reinterpolar [m]", min_value=1.0, max_value=5000.0, value=batch_half_width, step=5.0, key="ci_half")
                ci_step = ci4.number_input("Paso interpolado [m]", min_value=0.5, max_value=100.0, value=float(result.get("station_step", 5.0)), step=0.5, key="ci_step")
                ci_note = st.text_input("Nota", value="interpolación progresiva entre curvas", key="ci_note")
                if st.button("Interpolar tramo entre curvas de nivel", key="ci_apply", disabled=not bool(curves)):
                    sections, nok, nfail = apply_contour_interpolation_by_km(sections, axis_useful, curves, km_from_ci, km_to_ci, ci_half, ci_step, ci_note)
                    result["sections"] = sections
                    st.session_state["eje_result"] = result
                    st.success(f"Interpolación aplicada: {nok} secciones corregidas; {nfail} sin cruces suficientes con curvas.")
                    st.rerun()

            with st.expander("Insertar canal trapecial o rectangular entre km X y km Y", expanded=False):
                st.markdown(
                    """
Esta herramienta reemplaza las secciones existentes dentro del tramo por un canal prismático
rectangular o trapecial. La cota de fondo se toma del perfil longitudinal del DEM o del mínimo
de cada sección, por lo que el canal sigue progresivamente la pendiente del cauce.
"""
                )
                pc1, pc2, pc3, pc4 = st.columns(4)
                km_from_ch = pc1.number_input("km inicial", min_value=min_km, max_value=max_km, value=km_from, step=0.001, format="%.3f", key="ch_km_from")
                km_to_ch = pc2.number_input("km final", min_value=min_km, max_value=max_km, value=km_to, step=0.001, format="%.3f", key="ch_km_to")
                ch_kind = pc3.selectbox("Tipo canal", ["trapecial", "rectangular"], index=0, key="ch_kind")
                ch_n = pc4.number_input("Manning n", min_value=0.010, max_value=0.200, value=float(result.get("manning_n", 0.035)), step=0.001, format="%.3f", key="ch_n")
                pg1, pg2, pg3, pg4 = st.columns(4)
                ch_b = pg1.number_input("Ancho fondo [m]", min_value=0.1, max_value=1000.0, value=float(result.get("bottom_width", 5.0)), step=0.5, key="ch_b")
                ch_h = pg2.number_input("Altura [m]", min_value=0.1, max_value=100.0, value=float(result.get("height", 2.0)), step=0.1, key="ch_h")
                ch_zl = pg3.number_input("Talud izq H:V", min_value=0.0, max_value=50.0, value=float(result.get("talud_l", 1.5)), step=0.1, key="ch_zl")
                ch_zr = pg4.number_input("Talud der H:V", min_value=0.0, max_value=50.0, value=float(result.get("talud_r", 1.5)), step=0.1, key="ch_zr")
                ch_note = st.text_input("Nota canal", value="canal prismático por tramo", key="ch_note")
                if st.button("Aplicar canal al tramo", key="ch_apply"):
                    sections, ncanal = apply_prismatic_channel_by_km(sections, axis_useful, profile_df, km_from_ch, km_to_ch, ch_kind, ch_b, ch_h, ch_zl, ch_zr, ch_n, ch_note)
                    result["sections"] = sections
                    st.session_state["eje_result"] = result
                    st.success(f"Se reemplazaron {ncanal} secciones por canal {ch_kind} entre km {min(km_from_ch, km_to_ch):.3f} y km {max(km_from_ch, km_to_ch):.3f}.")
                    st.rerun()

    with tab4:
        st.subheader("Perfil longitudinal del cauce con secciones en 3D interactivo")
        fig3d = plot_sections_3d_interactive(axis_useful, profile_df, sections)
        if fig3d is not None and sections:
            st.plotly_chart(fig3d, use_container_width=True)
        else:
            st.info("Vista 3D no disponible. Verifica que Plotly esté instalado y que existan secciones.")
        st.markdown(
            """
**Uso recomendado de revisión:**
1. Recorre el perfil longitudinal 3D y detecta secciones muy planas, invertidas o con saltos de cota.
2. Abre la pestaña **Ventana sección por km**.
3. Edita cotas station-elevación, suaviza o interpola entre secciones vecinas.
4. Descarga nuevamente Excel/KMZ; las descargas usan las secciones ya modificadas.
"""
        )

    with tab5:
        st.subheader("Eje útil y secciones sobre imagen satelital")
        st.caption("Vista de referencia para corregir ancho y posición relativa de secciones. Para revisión final en Google Earth, descarga el KMZ actualizado.")
        if sections:
            render_satellite_reference(axis_useful, sections, pc_h_on, pc_s_on, to_wgs, height=720)
            st.divider()
            st.subheader("Corrección rápida de ancho por tramo")
            st.info("Para interpolar secciones entre curvas o insertar canal trapecial/rectangular por tramo, usa la pestaña ‘Ventana sección por km’." )
            min_km = float(min(s.km for s in sections))
            max_km = float(max(s.km for s in sections))
            c1, c2, c3, c4 = st.columns(4)
            km_from2 = c1.number_input("km inicial", min_value=min_km, max_value=max_km, value=min_km, step=0.001, format="%.3f", key="batch_km_from_sat")
            km_to2 = c2.number_input("km final", min_value=min_km, max_value=max_km, value=max_km, step=0.001, format="%.3f", key="batch_km_to_sat")
            default_hw2 = float(max((max(sections[0].stations) - min(sections[0].stations)) / 2.0 if sections[0].stations else 100.0, 1.0))
            new_half2 = c3.number_input("Nuevo semi-ancho [m]", min_value=1.0, max_value=5000.0, value=default_hw2, step=5.0, key="batch_half_sat")
            st_state2 = c4.selectbox("Estado", ["corregida", "aceptada", "revisar", "rechazada"], index=0, key="batch_status_sat")
            smooth2 = st.checkbox("Suavizar cotas luego de ajustar ancho", value=False, key="batch_smooth_sat")
            if st.button("Aplicar ancho al tramo en vista satelital", key="batch_apply_sat"):
                sections, nedit = apply_batch_section_improvement(sections, km_from2, km_to2, new_half2, smooth2, st_state2, "ajuste desde vista satelital")
                result["sections"] = sections
                st.session_state["eje_result"] = result
                st.success(f"Se actualizaron {nedit} secciones. Vuelve a esta pestaña para verlas proyectadas sobre la imagen.")
                st.rerun()
        else:
            st.info("No hay secciones para proyectar sobre imagen satelital.")

    with tab6:
        st.subheader("Descargas actualizadas")
        kmz_bytes = make_kmz(axis_oriented, axis_useful, pc_h_on, pc_s_on, sections, to_wgs)
        st.download_button("Descargar KMZ eje + secciones", kmz_bytes, file_name="hidrosed_eje_cauce_secciones_v7_excel_tramo_satelite.kmz", mime="application/vnd.google-earth.kmz")
        st.download_button("Descargar perfil longitudinal CSV", profile_df.to_csv(index=False).encode("utf-8"), file_name="perfil_longitudinal_tramo_util.csv", mime="text/csv")
        st.download_button("Descargar Excel tipo HEC-RAS actualizado", sections_to_hecras_excel(sections), file_name="secciones_tipo_hecras_editadas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        summary = {
            "modulo": "HidroSed Eje Cauce y Secciones v7 excel tramo satélite",
            "crs_proyectado": crs_proj.to_string(),
            "axis_validation": asdict(validation),
            "pc_hidrologico_original_lonlat": result.get("pc_h_original_lonlat"),
            "pc_soporte_original_lonlat": result.get("pc_s_original_lonlat"),
            "dist_pc_hidrologico_a_eje_m": result.get("dist_pc_h_m"),
            "dist_pc_soporte_a_eje_m": result.get("dist_pc_s_m"),
            "numero_secciones": len(sections),
            "secciones_por_estado": pd.Series([s.status for s in sections]).value_counts().to_dict() if sections else {},
            "evaluacion_curvas_apoyo": curve_eval,
            "modo_secciones": result.get("section_mode"),
            "fuente_dem": result.get("dem_source"),
            "descarga_opentopography": result.get("dem_download_info"),
        }
        st.download_button("Descargar JSON técnico actualizado", json.dumps(summary, indent=2, ensure_ascii=False).encode("utf-8"), file_name="resumen_tecnico_eje_secciones_v8.json", mime="application/json")


# --------------------------------------------------------------------------------------
# App UI
# --------------------------------------------------------------------------------------
st.title("🌊 HidroSed · Módulo Eje Cauce y Secciones v8")
st.caption("Eje útil obligatorio · DEM OpenTopography/manual · curvas opcionales · interpolación por curvas · canal prismático por tramo · perfil 3D hidráulico")

with st.expander("Criterio técnico del módulo", expanded=False):
    st.markdown(
        """
Este módulo construye y revisa el insumo hidráulico del cauce:

- eje completo y eje útil entre PC hidrológico y PC cuenca soporte;
- sentido aguas arriba → aguas abajo;
- ribera izquierda y derecha mirando aguas abajo;
- DEM descargado desde OpenTopography o GeoTIFF cargado por el usuario;
- perfil longitudinal DEM y perfil topográfico de respaldo opcional;
- secciones transversales naturales o prismáticas;
- **ventana independiente de revisión por km**;
- **edición de secciones** antes de exportar;
- **perfil longitudinal 3D interactivo con secciones**;
- **mejora de secciones por tramo km inicial–km final**;
- **interpolación progresiva entre curvas de nivel por tramo**;
- **canal trapecial o rectangular entre km X y km Y**;
- **proyección de eje útil y secciones sobre imagen satelital de referencia**.
"""
    )

st.sidebar.header("1. Entradas obligatorias")
axis_file = st.sidebar.file_uploader("Eje del cauce obligatorio · KMZ/KML con línea", type=["kmz", "kml"])
pc_h_file = st.sidebar.file_uploader("PC hidrológico · KMZ/KML con punto", type=["kmz", "kml"])
pc_s_file = st.sidebar.file_uploader("PC cuenca soporte · KMZ/KML con punto", type=["kmz", "kml"])

st.sidebar.header("2. DEM")
st.sidebar.markdown("El DEM puede cargarse manualmente como GeoTIFF o descargarse automáticamente desde **OpenTopography** usando API Key.")
dem_source = st.sidebar.radio("Fuente DEM", ["Descargar desde OpenTopography", "Cargar DEM GeoTIFF manual"], index=0)
dem_file = None
opentopo_api_key = ""
opentopo_demtype = "COP30"
opentopo_margin_km = 3.0
if dem_source == "Cargar DEM GeoTIFF manual":
    dem_file = st.sidebar.file_uploader("DEM GeoTIFF · cargar aquí", type=["tif", "tiff"])
else:
    opentopo_demtype = st.sidebar.selectbox("DEM OpenTopography", ["COP30", "NASADEM", "SRTMGL1", "SRTMGL3", "AW3D30"], index=0)
    opentopo_api_key = st.sidebar.text_input("API Key OpenTopography", type="password", help="La API Key no se guarda ni se imprime en archivos de salida.")
    opentopo_margin_km = st.sidebar.number_input("Margen de descarga DEM alrededor del eje y PCs [km]", min_value=0.5, max_value=100.0, value=3.0, step=0.5)

st.sidebar.header("3. Entradas opcionales de respaldo")
profile_file = st.sidebar.file_uploader("Perfil topográfico longitudinal de respaldo opcional", type=["csv", "txt", "xlsx", "xls", "kmz", "kml", "dxf"])
curves_file = st.sidebar.file_uploader("Curvas de nivel de apoyo opcionales · KMZ/KML/DXF", type=["kmz", "kml", "dxf"])
st.sidebar.header("3B. Excel prismático")
st.sidebar.caption("Úsalo cuando selecciones el modo **Desde Excel prismático**. La planilla debe tener km, tipo_seccion, ancho_fondo_m y altura_m.")
sections_excel = st.sidebar.file_uploader(
    "Cargar planilla Excel prismática aquí",
    type=["xlsx", "xls"],
    help="Obligatorio solo para el modo Desde Excel prismático. También puedes descargar la plantilla más abajo."
)

st.sidebar.header("4. Parámetros eje")
manual_invert = st.sidebar.checkbox("Invertir sentido del eje manualmente", value=False)
profile_step = st.sidebar.number_input("Paso perfil longitudinal DEM [m]", min_value=5.0, max_value=500.0, value=25.0, step=5.0)

st.sidebar.header("5. Parámetros secciones")
section_mode = st.sidebar.selectbox(
    "Modo de generación de secciones",
    ["Automático recomendado", "Desde curvas de nivel v13", "Desde DEM natural", "Prismática por parámetros", "Desde Excel prismático"],
    index=0,
    help=(
        "Automático recomendado: usa curvas v13 si se cargan; si no, usa DEM natural; "
        "si no hay DEM, genera secciones prismáticas por parámetros. "
        "Las curvas de nivel de apoyo son opcionales."
    )
)
section_spacing = st.sidebar.number_input("Separación entre secciones s [m]", min_value=5.0, max_value=1000.0, value=100.0, step=5.0)
half_width = st.sidebar.number_input("Semi-ancho de sección natural [m]", min_value=5.0, max_value=2000.0, value=100.0, step=5.0)
station_step = st.sidebar.number_input("Paso de muestreo transversal [m]", min_value=1.0, max_value=100.0, value=5.0, step=1.0)

if section_mode == "Desde curvas de nivel v13":
    if curves_file is None:
        st.sidebar.info("Modo v13 seleccionado sin curvas de apoyo: la app usará DEM natural automáticamente si existe DEM disponible.")
elif section_mode == "Automático recomendado":
    if curves_file is None:
        st.sidebar.info("Curvas de apoyo no cargadas: se usarán secciones naturales desde DEM, o secciones prismáticas si no hay DEM.")
elif section_mode == "Desde DEM natural":
    if dem_source == "Cargar DEM GeoTIFF manual" and dem_file is None:
        st.sidebar.warning("Para usar secciones naturales debes cargar un DEM GeoTIFF.")
    if dem_source == "Descargar desde OpenTopography" and not opentopo_api_key:
        st.sidebar.warning("Para usar secciones naturales con OpenTopography debes ingresar API Key.")

with st.sidebar.expander("Parámetros cauce prismático"):
    prism_kind = st.selectbox("Tipo", ["trapecial", "rectangular"])
    bottom_width = st.number_input("Ancho de fondo [m]", min_value=0.1, max_value=500.0, value=5.0, step=0.5)
    height = st.number_input("Altura [m]", min_value=0.1, max_value=50.0, value=2.0, step=0.1)
    talud_l = st.number_input("Talud izquierdo H:V", min_value=0.0, max_value=20.0, value=1.5, step=0.1)
    talud_r = st.number_input("Talud derecho H:V", min_value=0.0, max_value=20.0, value=1.5, step=0.1)
    manning_n = st.number_input("Manning n", min_value=0.010, max_value=0.200, value=0.035, step=0.001, format="%.3f")

st.sidebar.download_button(
    "Descargar plantilla Excel secciones prismáticas",
    data=make_prismatic_template(),
    file_name="plantilla_secciones_prismaticas_hidrosed.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

if st.sidebar.button("Limpiar resultados guardados"):
    st.session_state.pop("eje_result", None)
    st.rerun()

run = st.sidebar.button("Procesar eje y secciones", type="primary")

if run:
    if axis_file is None or pc_h_file is None or pc_s_file is None:
        st.error("Faltan entradas obligatorias: eje del cauce, PC hidrológico y PC cuenca soporte.")
        st.stop()
    try:
        inputs = dict(
            axis_file=axis_file,
            pc_h_file=pc_h_file,
            pc_s_file=pc_s_file,
            dem_source=dem_source,
            dem_file=dem_file,
            opentopo_api_key=opentopo_api_key,
            opentopo_demtype=opentopo_demtype,
            opentopo_margin_km=opentopo_margin_km,
            profile_file=profile_file,
            curves_file=curves_file,
            sections_excel=sections_excel,
            manual_invert=manual_invert,
            profile_step=profile_step,
            section_mode=section_mode,
            section_spacing=section_spacing,
            half_width=half_width,
            station_step=station_step,
            prism_kind=prism_kind,
            bottom_width=bottom_width,
            height=height,
            talud_l=talud_l,
            talud_r=talud_r,
            manning_n=manning_n,
        )
        with st.spinner("Procesando eje, DEM, perfil y secciones..."):
            result = compute_app_result(inputs)
        st.session_state["eje_result"] = result
        st.success("Procesamiento finalizado. Puedes revisar y editar las secciones sin reprocesar el DEM.")
    except Exception as e:
        st.error("No se pudo procesar la información.")
        st.exception(e)
        st.stop()

if "eje_result" not in st.session_state:
    st.info("Carga el eje del cauce, PC hidrológico y PC cuenca soporte. Luego elige DEM OpenTopography o GeoTIFF manual y presiona **Procesar eje y secciones**.")
    st.stop()

render_results(st.session_state["eje_result"])
