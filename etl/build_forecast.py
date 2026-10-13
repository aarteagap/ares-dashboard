#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Agrega una proyección (forecast) de sobrecostos para el resto de T26-27 al
sobrecostos_data.json ya generado por etl_sobrecostos.py.

No depende de los Excel de Registro de Gastos (T23-24/T24-25/T25-26) ni de
GAOPEX_T2627: toma el histórico ya agregado desde el propio DATA.weekly /
DATA.bookings (generados en una corrida anterior del ETL principal) y lo
combina con las instrucciones PROYECTADAS del archivo Horizon más reciente.

Metodología (pedida por el usuario):
  1. Tasa histórica = USD histórico de sobrecostos (solo conceptos activos
     por defecto, T23-24+T24-25+T25-26) / Nº de bookings históricos,
     calculada por (Vía, Exportador, Clasificación OPEX).
  2. Nº de instrucciones futuras = filas de la hoja HORIZON_FORECAST con
     Status = "Proyectado", agrupadas por (Vía, Exportador, Semana).
  3. USD proyectado(semana, vía, exportador, opex) = tasa × nº instrucciones.

Las filas proyectadas se agregan a DATA.weekly con "p": true (vs "p": false
en las filas reales), misma forma que las filas reales para que el frontend
las pueda filtrar igual (Vía/Exportador/OPEX/Semana), pero SIN granularidad
de Concepto individual (se usa el grupo OPEX como pseudo-concepto) y
EXCLUIDAS por defecto de los KPIs de "Sobrecosto total" (ver app.js).
"""
import json, math, unicodedata
from collections import defaultdict
from datetime import datetime, date
import openpyxl

SCRATCH = "/tmp/claude-0/-home-user-axis-2-0-dashboard/2fb4a3aa-b69c-5e44-bf67-dc7a2b1ec063/scratchpad"
DATA_JSON = f"{SCRATCH}/sobrecostos_data.json"
HORIZON_FILE = "/root/.claude/uploads/2fb4a3aa-b69c-5e44-bf67-dc7a2b1ec063/53f19d39-HORIZON_FORECAST_2627_V1.xlsm"

SEASON = "T26-27"
SEASON_ANCHOR = datetime(2026, 5, 1)
CLOSED_SEASONS = ["T23-24", "T24-25", "T25-26"]

DEFAULT_EXCLUDE_PATTERNS = [
    "GESTION DOCUMENTAL", "SENASA", "SERVICIO INTEGRAL",
    "COMISION DE AG", "CERTIFICADO DE ORIGEN", "FEE DE SEGURIDAD",
    "FLETE INTERNACIONAL", "FLETE NACIONAL", "MANEJO DE PLATAFORMA",
    "COURIER", "VISTO BUENO", "MOVILIDAD", "MULTA", "NOTA DE CREDITO",
]


def strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


def norm(s):
    if s is None:
        return ""
    return strip_accents(str(s)).strip().upper()


def is_default_active(concepto):
    c = norm(concepto)
    return not any(p in c for p in DEFAULT_EXCLUDE_PATTERNS)


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


def to_campaign_week(d):
    delta_days = (d - SEASON_ANCHOR).days
    return math.floor(delta_days / 7) + 1


# ---------------------------------------------------------------------------
# 1) Tasa histórica USD/booking por (Vía, Exportador, OPEX) — solo conceptos
#    "sobrecosto" activos por defecto, T23-24 a T25-26.
# ---------------------------------------------------------------------------
def build_historical_rates(data):
    usd_by_key = defaultdict(float)
    for r in data["weekly"]:
        if r.get("p"):
            continue  # no reutilizar una proyección previa como base
        if r["s"] not in CLOSED_SEASONS:
            continue
        if not is_default_active(r["c"]):
            continue
        usd_by_key[(r["v"], r["ex"], r["o"])] += r["usd"]

    bookings_by_via_ex = defaultdict(set)
    for b in data["bookings"]:
        if b["s"] not in CLOSED_SEASONS:
            continue
        bookings_by_via_ex[(b["v"], b["ex"])].add((b["s"], b["bk"]))

    rate = {}
    for (v, ex, o), usd in usd_by_key.items():
        n_bookings = len(bookings_by_via_ex.get((v, ex), set()))
        if n_bookings > 0:
            rate[(v, ex, o)] = usd / n_bookings
    return rate, {k: len(v) for k, v in bookings_by_via_ex.items()}


# ---------------------------------------------------------------------------
# 2) Instrucciones PROYECTADAS de Horizon (hoja HORIZON_FORECAST)
# ---------------------------------------------------------------------------
def load_projected_instructions():
    wb = openpyxl.load_workbook(HORIZON_FILE, read_only=True, data_only=True)
    ws = wb["HORIZON_FORECAST"]
    rows_iter = ws.iter_rows(values_only=True)
    header = next(rows_iter)
    col = lambda *a: find_col(header, *a)
    i_status = col("Status")
    i_mode = col("Mode")
    i_shipper = col("Shipper")
    i_dispatch = col("Dispatch Date")
    i_etd = col("ETD")
    i_packplan = col("Pack Plan")
    i_instruction = col("Instruction")

    instructions = []
    n_sin_fecha = 0
    for r in rows_iter:
        if r[i_status] != "Proyectado":
            continue
        via = map_via(r[i_mode])
        if via not in ("Marítimo", "Aéreo"):
            continue  # Terrestre no genera sobrecostos / no se proyecta
        shipper = map_shipper(r[i_shipper])

        d = r[i_dispatch] if isinstance(r[i_dispatch], datetime) else None
        if d is None and isinstance(r[i_etd], datetime):
            d = r[i_etd]
        week = None
        if d is not None:
            week = d.isocalendar()[1]
            cweek = to_campaign_week(d)
        elif isinstance(r[i_packplan], (int, float)):
            # Fallback: "Pack Plan" es el Nº de semana ISO de empaque (ETD
            # real cae ~1 semana después); se usa el lunes de esa semana
            # ISO-2026 como fecha aproximada solo para la semana de campaña.
            week = int(r[i_packplan])
            approx = date.fromisocalendar(2026, week, 1)
            cweek = to_campaign_week(datetime(approx.year, approx.month, approx.day))
        else:
            n_sin_fecha += 1
            continue

        instructions.append({
            "via": via, "shipper": shipper, "week": week, "cweek": cweek,
            "instruction": r[i_instruction],
        })
    return instructions, n_sin_fecha


def main():
    with open(DATA_JSON, encoding="utf-8") as f:
        data = json.load(f)

    rate, n_bookings_by_via_ex = build_historical_rates(data)
    instructions, n_sin_fecha = load_projected_instructions()
    print(f"Instrucciones proyectadas en Horizon: {len(instructions)} "
          f"({n_sin_fecha} sin fecha/semana utilizable, descartadas)")

    n_instr_by_bucket = defaultdict(int)
    for ins in instructions:
        n_instr_by_bucket[(ins["via"], ins["shipper"], ins["week"], ins["cweek"])] += 1

    projected_rows = []
    for (via, shipper, week, cweek), n_instr in n_instr_by_bucket.items():
        for (v, ex, o), r in rate.items():
            if v != via or ex != shipper:
                continue
            usd = r * n_instr
            if usd <= 0:
                continue
            projected_rows.append({
                "s": SEASON, "w": week, "cw": cweek, "v": via, "ex": shipper,
                "o": o, "c": o, "usd": round(usd, 2), "n": n_instr,
                "active": True, "p": True,
            })

    total_proj_usd = sum(r["usd"] for r in projected_rows)
    print(f"Filas proyectadas generadas: {len(projected_rows)}, "
          f"USD proyectado total: {total_proj_usd:,.2f}")

    # marca explícitamente las filas existentes como no-proyectadas
    for r in data["weekly"]:
        r.setdefault("p", False)
    data["weekly"].extend(projected_rows)

    n_instr_total = len(instructions)
    semanas = sorted({ins["week"] for ins in instructions})
    data["meta"]["forecast"] = {
        "n_instrucciones_proyectadas": n_instr_total,
        "usd_proyectado_total": round(total_proj_usd, 2),
        "semanas_cubiertas": [semanas[0], semanas[-1]] if semanas else None,
        "n_bookings_historicos_por_via_exportador": {
            f"{v} · {ex}": n for (v, ex), n in n_bookings_by_via_ex.items()
        },
        "metodologia": (
            "Para cada combinación Vía×Exportador×Clasificación OPEX se calculó "
            "una tasa histórica = USD de sobrecostos (solo conceptos activos por "
            "defecto, T23-24+T24-25+T25-26) / Nº de bookings históricos de esa "
            "Vía×Exportador. Las instrucciones con Status='Proyectado' en la hoja "
            "HORIZON_FORECAST de Horizon (temporada T26-27) se agruparon por "
            "Vía×Exportador×Semana (semana = ISO de 'Dispatch Date', o si falta, "
            "de 'ETD', o si falta, el Nº de 'Pack Plan' como semana ISO "
            "aproximada). USD proyectado = tasa × Nº de instrucciones proyectadas "
            "de esa semana. Es un estimado: asume que el patrón de sobrecostos por "
            "despacho del resto de la temporada se parece al histórico de las 3 "
            "temporadas cerradas; no reemplaza al dato real una vez ejecutado."
        ),
    }

    with open(DATA_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    print("OK ->", DATA_JSON)


if __name__ == "__main__":
    main()
