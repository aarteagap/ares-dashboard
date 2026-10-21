#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Agrega Kg reales de las 3 temporadas cerradas (T23-24, T24-25, T25-26) a
sobrecostos_data.json, usando los programas de despacho históricos (BD_2324,
BD_2425, BD_2526) en vez del Registro de Gastos (que ya no tenemos en este
contenedor). Mismo criterio que T26-27: Kg = "Peso Neto"/"Net weight" de las
filas con Status/Estado = Confirmado, agregado por (temporada, semana, vía,
exportador). No se cruza por Booking contra el gasto (el archivo ya agregado
no guarda bookings suficientes -se truncó a los 3000 más grandes-); se usa
el programa de despacho como fuente independiente de Kg, igual que Horizon
para T26-27.
"""
import json, math, unicodedata
from collections import defaultdict
from datetime import datetime
import openpyxl

SCRATCH = "/tmp/claude-0/-home-user-axis-2-0-dashboard/2fb4a3aa-b69c-5e44-bf67-dc7a2b1ec063/scratchpad"
DATA_JSON = f"{SCRATCH}/sobrecostos_data.json"
BASE = "/root/.claude/uploads/2fb4a3aa-b69c-5e44-bf67-dc7a2b1ec063"

FILES = {
    "T23-24": (f"{BASE}/c33b9745-BD_2324.xlsx", "BD Despachos"),
    "T24-25": (f"{BASE}/e02c51fa-BD_2425.xlsx", "BD PLANNING"),
    "T25-26": (f"{BASE}/51d0a793-BD_2526.xlsx", "BD PLANNING"),
}
SEASON_ANCHOR = {
    "T23-24": datetime(2023, 5, 1), "T24-25": datetime(2024, 5, 1),
    "T25-26": datetime(2025, 5, 1),
}


def strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


def norm(s):
    if s is None:
        return ""
    return strip_accents(str(s)).strip().upper()


VIA_MAP = {
    "MARITIMO": "Marítimo", "SEA": "Marítimo",
    "AEREO": "Aéreo", "AIR": "Aéreo",
    "TERRESTRE": "Terrestre", "LAND": "Terrestre",
}


def map_via(v):
    return VIA_MAP.get(norm(v))


def map_shipper(v):
    n = norm(v)
    if "HORTIFRUT" in n or n == "HFPE":
        return "Hortifrut Perú S.A.C."
    if "JORDIE" in n:
        return "Inversiones Jordie S.A."
    if "TAL S.A" in n or "TALSA" in n:
        return "Tal S.A. (Talsa)"
    if "HFE BERRIES" in n or "HFE_BERRIES" in n:
        return "HFE Berries Perú S.A.C."
    return "Sin especificar"


def find_col(header, *aliases):
    normed = [norm(h) for h in header]
    for alias in aliases:
        a = norm(alias)
        for i, h in enumerate(normed):
            if h == a:
                return i
    for alias in aliases:
        a = norm(alias)
        for i, h in enumerate(normed):
            if a in h:
                return i
    return None


def numeric(v):
    return isinstance(v, (int, float))


def to_campaign_week(d, season):
    return math.floor((d - SEASON_ANCHOR[season]).days / 7) + 1


def find_header_row(ws, max_probe=3):
    """Algunas hojas (BD_2324) traen una fila de título fusionada antes del
    encabezado real; se detecta probando si la fila trae 'Shipper'/'Booking'."""
    rows_iter = ws.iter_rows(values_only=True)
    for _ in range(max_probe):
        row = next(rows_iter)
        normed = [norm(h) for h in row]
        if any("SHIPPER" in h for h in normed) and any("BOOKING" in h for h in normed):
            return row, rows_iter
    raise ValueError("No se encontró fila de encabezado")


def load_season_kg(season, path, sheet_name):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name]
    header, rows_iter = find_header_row(ws)
    col = lambda *a: find_col(header, *a)
    i_status = col("Estado", "Status")
    i_mode = col("Vía", "Mode")
    i_shipper = col("Shipper")
    i_dispatch = col("F. de Despacho", "Dispatch Date")
    i_weight = col("Peso Neto", "Net weight")

    # 1ra pasada: junta todas las instrucciones Confirmado (con o sin peso) y
    # el promedio de Kg/instrucción por Vía, a partir de las que SÍ traen peso.
    instructions = []
    kg_sum_by_via = defaultdict(float)
    n_con_peso_by_via = defaultdict(int)
    n_rows_confirmado = 0
    for r in rows_iter:
        status = norm(r[i_status])
        if status != "CONFIRMADO":
            continue
        n_rows_confirmado += 1
        via = map_via(r[i_mode])
        if via not in ("Marítimo", "Aéreo"):
            continue
        d = r[i_dispatch] if isinstance(r[i_dispatch], datetime) else None
        if d is None:
            continue
        kg = r[i_weight] if numeric(r[i_weight]) else None
        if kg and kg > 0:
            if kg > 100000:
                # mismo error de unidad detectado en Horizon (peso en gramos)
                kg = kg / 1000
            kg_sum_by_via[via] += kg
            n_con_peso_by_via[via] += 1
        else:
            kg = None
        shipper = map_shipper(r[i_shipper])
        week = d.isocalendar()[1]
        cweek = to_campaign_week(d, season)
        instructions.append({"via": via, "shipper": shipper, "week": week,
                              "cweek": cweek, "kg": kg})

    avg_kg_by_via = {
        via: kg_sum_by_via[via] / n_con_peso_by_via[via]
        for via in kg_sum_by_via if n_con_peso_by_via[via] > 0
    }

    # 2da pasada: agrega Kg real donde existe, o el promedio de esa Vía/
    # temporada donde falta (pedido por el usuario para que no se note tan
    # irregular la cobertura parcial de peso registrado).
    kg_agg = defaultdict(float)
    n_reales, n_imputados = 0, 0
    for ins in instructions:
        kg = ins["kg"]
        if kg is None:
            kg = avg_kg_by_via.get(ins["via"])
            if kg is None:
                continue  # sin ningún dato de peso para esa vía en la temporada
            n_imputados += 1
        else:
            n_reales += 1
        key = (season, ins["week"], ins["cweek"], ins["via"], ins["shipper"])
        kg_agg[key] += kg
    n_bookings_con_kg = n_reales + n_imputados

    print(f"{season}: {n_rows_confirmado} filas Confirmado, "
          f"{n_bookings_con_kg} con Kg ({n_reales} reales + {n_imputados} "
          f"imputados con promedio Kg/instrucción de su Vía), "
          f"{len(kg_agg)} buckets semana/vía/exportador, "
          f"promedio Kg/instrucción: {({v: round(k,1) for v,k in avg_kg_by_via.items()})}")
    return kg_agg, n_reales, n_imputados


def main():
    with open(DATA_JSON, encoding="utf-8") as f:
        data = json.load(f)

    # quita cualquier Kg histórico previo de estas 3 temporadas (por si se
    # vuelve a correr este script) sin tocar T26-27.
    data["weekly_kg"] = [r for r in data.get("weekly_kg", [])
                          if r["s"] not in FILES]

    all_rows = []
    totales_imputacion = {}
    for season, (path, sheet) in FILES.items():
        kg_agg, n_reales, n_imputados = load_season_kg(season, path, sheet)
        totales_imputacion[season] = (n_reales, n_imputados)
        all_rows.extend(
            {"s": s, "w": w, "cw": cw, "v": v, "ex": ex, "kg": round(kg, 1), "p": False}
            for (s, w, cw, v, ex), kg in kg_agg.items()
        )

    data["weekly_kg"].extend(all_rows)

    total_kg_by_season = defaultdict(float)
    for r in all_rows:
        total_kg_by_season[r["s"]] += r["kg"]

    # actualiza la nota de metodología del modo Kg
    partes = " · ".join(f"{s}: {total_kg_by_season[s]:,.0f} kg" for s in FILES)
    imput_txt = " · ".join(
        f"{s}: {totales_imputacion[s][0]} reales + {totales_imputacion[s][1]} imputados"
        for s in FILES
    )
    data["meta"]["kg_note"] = (
        "El modo USD/Kg usa el peso neto real de los despachos cuando está "
        "registrado, y donde falta se completa con el promedio de Kg por "
        "PO/instrucción Confirmada de esa Vía y temporada (para que la "
        "cobertura parcial de peso registrado no se note irregular semana a "
        "semana). Fuente: para T23-24/T24-25/T25-26, los programas de "
        "despacho históricos (BD_2324/BD_2425/BD_2526, filas Status/"
        f"Estado='Confirmado'; {partes} en total — {imput_txt}); para T26-27, "
        "Horizon (despachos confirmados). El tramo "
        "proyectado de T26-27 (línea punteada) usa Kg estimado = promedio "
        "histórico de Kg por PO/instrucción confirmada de cada Vía × Nº de "
        "instrucciones proyectadas de esa semana, ya que Horizon no trae peso "
        "para instrucciones 'Proyectado' todavía. Se detectaron y corrigieron "
        "filas con el peso cargado en gramos en vez de kg (valores > 100,000) "
        "tanto en los programas históricos como en Horizon."
    )

    with open(DATA_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    print("OK ->", DATA_JSON)
    print("total kg by season:", dict(total_kg_by_season))


if __name__ == "__main__":
    main()
