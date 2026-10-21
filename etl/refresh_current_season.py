#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Refresca por completo la temporada actual (T26-27) en sobrecostos_data.json:
  - Reemplaza las líneas "actuales" (reales) cruzando GAOPEX_T2627_En_curso-Opex
    (clasificación OPEX) contra la hoja HORIZON_FORECAST del Horizon más
    reciente (Status='Confirmado' con Dispatch Date -> Vía + Fecha + Kg).
  - Reemplaza la proyección (líneas p:true) con las instrucciones
    Status='Proyectado' de ese mismo Horizon, usando la tasa histórica ya
    calculada sobre T23-24/T24-25/T25-26 (que no se tocan).
  - Agrega "weekly_kg": Kg reales despachados por (temporada, semana, vía,
    exportador), tomados de Horizon (Net weight), para el modo USD/Kg del
    gráfico. Solo hay dato real para T26-27 confirmado (Horizon no trae Kg en
    las filas Proyectado, y no hay Kg histórico en este dataset).

No requiere los Excel de Registro de Gastos originales: parte del
sobrecostos_data.json ya generado (histórico embebido) + los 2 archivos de
esta sesión (GAOPEX y Horizon).
"""
import json, math, unicodedata
from collections import defaultdict
from datetime import datetime, date
import openpyxl

SCRATCH = "/tmp/claude-0/-home-user-axis-2-0-dashboard/2fb4a3aa-b69c-5e44-bf67-dc7a2b1ec063/scratchpad"
DATA_JSON = f"{SCRATCH}/sobrecostos_data.json"
GAOPEX_FILE = "/root/.claude/uploads/2fb4a3aa-b69c-5e44-bf67-dc7a2b1ec063/cf512946-GAOPEX_T2627_En_curso-Opex.xlsm"
HORIZON_FILE = "/root/.claude/uploads/2fb4a3aa-b69c-5e44-bf67-dc7a2b1ec063/c2cad423-HORIZON_FORECAST_2627_V1.xlsm"

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


def to_iso_week(d):
    return d.isocalendar()[1]


def to_campaign_week(d):
    return math.floor((d - SEASON_ANCHOR).days / 7) + 1


def numeric(v):
    return isinstance(v, (int, float))


# ---------------------------------------------------------------------------
# 1) Horizon: despachos confirmados (Vía+Fecha+Kg reales) e instrucciones
#    proyectadas, desde la hoja HORIZON_FORECAST.
# ---------------------------------------------------------------------------
def load_horizon():
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
    i_booking = col("Booking|AWB|CRT", "Booking")
    i_po = col("PO")
    i_weight = col("Net weight")

    confirmed_by_bk = {}  # booking_norm -> {via, shipper, week, cweek, kg}
    confirmed_by_po = {}  # po_norm -> {via, shipper, week, cweek, kg}
    projected = []   # [{via, shipper, week, cweek}]
    n_confirmed_sin_fecha = 0
    n_proyectado_sin_fecha = 0

    for r in rows_iter:
        status = r[i_status]
        via = map_via(r[i_mode])
        if via not in ("Marítimo", "Aéreo"):
            continue
        shipper = map_shipper(r[i_shipper])

        d = r[i_dispatch] if isinstance(r[i_dispatch], datetime) else None
        if d is None and isinstance(r[i_etd], datetime):
            d = r[i_etd]
        week = cweek = None
        if d is not None:
            week = to_iso_week(d)
            cweek = to_campaign_week(d)
        elif isinstance(r[i_packplan], (int, float)):
            week = int(r[i_packplan])
            approx = date.fromisocalendar(2026, week, 1)
            cweek = to_campaign_week(datetime(approx.year, approx.month, approx.day))

        if status == "Confirmado":
            bk = r[i_booking]
            po = r[i_po] if i_po is not None else None
            if week is None or (bk in (None, "N/A") and po in (None, "N/A")):
                n_confirmed_sin_fecha += 1
                continue
            kg = r[i_weight] if numeric(r[i_weight]) else None
            if kg is not None and kg > 100000:
                # Error de unidad en la fuente: algunas filas registran el
                # peso en gramos en vez de kg (ej. 12,150,000 en vez de
                # 12,150). Se detecta por outlier extremo y se normaliza.
                kg = kg / 1000
            info = {"via": via, "shipper": shipper, "week": week, "cweek": cweek, "kg": kg}
            if bk not in (None, "N/A"):
                confirmed_by_bk[norm(bk)] = info
            if po not in (None, "N/A"):
                confirmed_by_po[norm(po)] = info
        elif status == "Proyectado":
            if week is None:
                n_proyectado_sin_fecha += 1
                continue
            projected.append({"via": via, "shipper": shipper, "week": week, "cweek": cweek})

    return confirmed_by_bk, confirmed_by_po, projected, n_confirmed_sin_fecha, n_proyectado_sin_fecha


# ---------------------------------------------------------------------------
# 2) GAOPEX: líneas OPEX clasificadas de la temporada actual.
# ---------------------------------------------------------------------------
def load_gaopex():
    wb = openpyxl.load_workbook(GAOPEX_FILE, read_only=True, data_only=True)
    ws = wb["OPEX_EXPENSES_2627"]
    rows_iter = ws.iter_rows(values_only=True)
    header = next(rows_iter)
    col = lambda *a: find_col(header, *a)
    i_shipper = col("Shipper")
    i_supplier = col("Supplier")
    i_concepto = col("OPEX Expense Description")
    i_grupo = col("OPEX Expense Group")
    i_usd = col("Total Expense USD")
    i_booking = col("Booking")
    i_po = col("PO")

    rows = []
    n_rows_total = 0
    total_usd_all_rows = 0.0
    for r in rows_iter:
        n_rows_total += 1
        usd_all = r[i_usd]
        try:
            total_usd_all_rows += float(usd_all) if usd_all not in (None, "") else 0.0
        except Exception:
            pass
        concepto = r[i_concepto]
        if concepto is None:
            continue
        concepto = str(concepto).strip()
        usd = r[i_usd]
        try:
            usd = float(usd) if usd not in (None, "") else 0.0
        except Exception:
            usd = 0.0
        grupo = str(r[i_grupo]).strip() if r[i_grupo] else "SIN CLASIFICAR"
        supplier = str(r[i_supplier]).strip() if r[i_supplier] else "Sin especificar"
        shipper = map_shipper(r[i_shipper]) if i_shipper is not None else "Sin especificar"
        booking_raw = r[i_booking]
        po_raw = r[i_po] if i_po is not None else None
        rows.append({
            "concepto": concepto, "opex": grupo, "usd": usd, "shipper": shipper,
            "supplier": supplier, "booking": booking_raw, "po": po_raw,
        })
    return rows, n_rows_total, total_usd_all_rows


# ---------------------------------------------------------------------------
# 3) Tasa histórica USD/booking por (Vía, Exportador, OPEX) — temporadas
#    cerradas (no se tocan en esta corrida).
# ---------------------------------------------------------------------------
def build_historical_rates(data):
    usd_by_key = defaultdict(float)
    for r in data["weekly"]:
        if r.get("p") or r["s"] not in CLOSED_SEASONS or not is_default_active(r["c"]):
            continue
        usd_by_key[(r["v"], r["ex"], r["o"])] += r["usd"]
    bookings_by_via_ex = defaultdict(set)
    for b in data["bookings"]:
        if b["s"] in CLOSED_SEASONS:
            bookings_by_via_ex[(b["v"], b["ex"])].add((b["s"], b["bk"]))
    rate = {}
    for (v, ex, o), usd in usd_by_key.items():
        n = len(bookings_by_via_ex.get((v, ex), set()))
        if n > 0:
            rate[(v, ex, o)] = usd / n
    return rate, {k: len(v) for k, v in bookings_by_via_ex.items()}


def main():
    with open(DATA_JSON, encoding="utf-8") as f:
        data = json.load(f)

    # -- limpia toda la data vieja de T26-27 (actual + proyectada) --
    data["weekly"] = [r for r in data["weekly"] if r["s"] != SEASON]
    data["bookings"] = [b for b in data["bookings"] if b["s"] != SEASON]
    data["weekly_kg"] = [r for r in data.get("weekly_kg", []) if r["s"] != SEASON]

    confirmed_by_bk, confirmed_by_po, projected_instr, n_conf_sf, n_proy_sf = load_horizon()
    confirmed_map = confirmed_by_bk  # alias retro-compat
    # Unión deduplicada de despachos confirmados (por identidad de objeto:
    # una misma fila de Horizon puede estar indexada por Booking y por PO a
    # la vez) -> usada para estadísticas de Kg y conteos globales.
    confirmed_all = list({id(v): v for v in list(confirmed_by_bk.values()) + list(confirmed_by_po.values())}.values())
    gaopex_rows, n_rows_total, total_usd_all_rows = load_gaopex()
    print(f"Horizon: {len(confirmed_by_bk)} bookings + {len(confirmed_by_po)} PO "
          f"confirmados con Vía+Fecha ({n_conf_sf} sin fecha utilizable), "
          f"{len(projected_instr)} instrucciones proyectadas ({n_proy_sf} sin fecha utilizable)")
    print(f"GAOPEX: {n_rows_total} líneas totales, "
          f"{sum(1 for r in gaopex_rows)} clasificadas")

    # -- cruce GAOPEX x Horizon confirmado: primero por Booking, si no
    #    matchea se intenta por PO (algunas líneas traen uno pero no el otro;
    #    usar ambos como identificador sube la cobertura del cruce) --
    matched = []
    n_matched = 0
    n_matched_by_po = 0
    for r in gaopex_rows:
        bk = r["booking"]
        m = confirmed_by_bk.get(norm(bk)) if bk is not None else None
        if not m:
            po = r.get("po")
            m = confirmed_by_po.get(norm(po)) if po is not None else None
            if m:
                n_matched_by_po += 1
        if not m:
            continue
        n_matched += 1
        matched.append({
            "s": SEASON, "w": m["week"], "cw": m["cweek"], "v": m["via"],
            "ex": r["shipper"], "o": r["opex"], "c": r["concepto"],
            "usd": round(r["usd"], 2), "booking": str(bk).strip() if bk else f"PO {r.get('po')}",
            "proveedor": r["supplier"],
        })

    # -- agregación semanal (actual) --
    weekly_agg = defaultdict(lambda: {"usd": 0.0, "n": 0})
    for r in matched:
        key = (r["s"], r["w"], r["cw"], r["v"], r["ex"], r["o"], r["c"])
        weekly_agg[key]["usd"] += r["usd"]
        weekly_agg[key]["n"] += 1
    actual_weekly_rows = [
        {"s": s, "w": w, "cw": cw, "v": v, "ex": ex, "o": o, "c": c,
         "usd": round(a["usd"], 2), "n": a["n"], "active": is_default_active(c), "p": False}
        for (s, w, cw, v, ex, o, c), a in weekly_agg.items()
    ]

    # -- agregación por booking (actual) --
    bk_agg = defaultdict(lambda: {"usd": 0.0, "n": 0, "opex": defaultdict(float), "prov": set()})
    for r in matched:
        key = (r["s"], r["v"], r["ex"], r["booking"])
        b = bk_agg[key]
        b["usd"] += r["usd"]; b["n"] += 1
        b["opex"][r["o"]] += r["usd"]
        if r["proveedor"]:
            b["prov"].add(r["proveedor"])
    actual_booking_rows = [
        {"s": s, "v": v, "ex": ex, "bk": bk, "usd": round(a["usd"], 2), "n": a["n"],
         "opex": {k: round(val, 2) for k, val in a["opex"].items()}, "np": len(a["prov"])}
        for (s, v, ex, bk), a in bk_agg.items()
    ]

    # -- proyección (tasa histórica x instrucciones proyectadas) --
    rate, n_bookings_hist = build_historical_rates(data)
    n_instr_bucket = defaultdict(int)
    for ins in projected_instr:
        n_instr_bucket[(ins["via"], ins["shipper"], ins["week"], ins["cweek"])] += 1
    projected_rows = []
    for (via, shipper, week, cweek), n_instr in n_instr_bucket.items():
        for (v, ex, o), rt in rate.items():
            if v != via or ex != shipper:
                continue
            usd = rt * n_instr
            if usd <= 0:
                continue
            projected_rows.append({
                "s": SEASON, "w": week, "cw": cweek, "v": via, "ex": shipper,
                "o": o, "c": o, "usd": round(usd, 2), "n": n_instr,
                "active": True, "p": True,
            })
    total_proj_usd = sum(r["usd"] for r in projected_rows)

    # -- Kg despachados (de Horizon confirmado): real donde existe, o
    #    promedio de Kg/PO de esa Vía donde falta (pedido por el usuario para
    #    que la cobertura parcial de peso registrado no se note irregular),
    #    por semana/vía/exportador --
    kg_sum_by_via = defaultdict(float)
    n_bk_con_kg_by_via = defaultdict(int)
    n_bookings_con_kg = 0
    for m in confirmed_all:
        if m["kg"]:
            n_bookings_con_kg += 1
            kg_sum_by_via[m["via"]] += m["kg"]
            n_bk_con_kg_by_via[m["via"]] += 1
    avg_kg_by_via = {
        via: kg_sum_by_via[via] / n_bk_con_kg_by_via[via]
        for via in kg_sum_by_via if n_bk_con_kg_by_via[via] > 0
    }
    n_bookings_imputados = 0
    kg_agg = defaultdict(float)
    kg_agg_real = defaultdict(float)
    for m in confirmed_all:
        kg = m["kg"]
        key = (SEASON, m["week"], m["cweek"], m["via"], m["shipper"])
        if not kg:
            kg = avg_kg_by_via.get(m["via"])
            if not kg:
                continue
            n_bookings_imputados += 1
            kg_agg[key] += kg
        else:
            kg_agg[key] += kg
            kg_agg_real[key] += kg
    weekly_kg_rows = [
        {"s": s, "w": w, "cw": cw, "v": v, "ex": ex, "kg": round(kg, 1),
         "kg_real": round(kg_agg_real.get((s, w, cw, v, ex), 0.0), 1), "p": False}
        for (s, w, cw, v, ex), kg in kg_agg.items()
    ]

    # -- Kg PROYECTADO: promedio histórico de Kg por PO/instrucción confirmada,
    #    por Vía (Marítimo/Aéreo), × Nº de instrucciones proyectadas de esa
    #    semana. Pedido por el usuario como respaldo mientras no haya Kg real
    #    para las instrucciones 'Proyectado' de Horizon. 100% estimado -> kg_real=0.
    projected_kg_rows = []
    for (via, shipper, week, cweek), n_instr in n_instr_bucket.items():
        avg_kg = avg_kg_by_via.get(via)
        if not avg_kg:
            continue
        projected_kg_rows.append({
            "s": SEASON, "w": week, "cw": cweek, "v": via, "ex": shipper,
            "kg": round(avg_kg * n_instr, 1), "kg_real": 0.0, "p": True,
        })
    total_proj_kg = sum(r["kg"] for r in projected_kg_rows)

    data["weekly"].extend(actual_weekly_rows)
    data["weekly"].extend(projected_rows)
    data["bookings"].extend(actual_booking_rows)
    data["weekly_kg"] = data.get("weekly_kg", []) + weekly_kg_rows + projected_kg_rows

    # -- resumen "current_season" (panel Temporada actual) --
    by_concepto = defaultdict(lambda: {"usd": 0.0, "n": 0})
    by_grupo = defaultdict(lambda: {"usd": 0.0, "n": 0})
    by_supplier = defaultdict(lambda: {"usd": 0.0, "n": 0})
    by_shipper = defaultdict(lambda: {"usd": 0.0, "n": 0})
    total_usd = 0.0
    for r in gaopex_rows:
        by_concepto[r["concepto"]]["usd"] += r["usd"]; by_concepto[r["concepto"]]["n"] += 1
        by_grupo[r["opex"]]["usd"] += r["usd"]; by_grupo[r["opex"]]["n"] += 1
        by_supplier[r["supplier"]]["usd"] += r["usd"]; by_supplier[r["supplier"]]["n"] += 1
        by_shipper[r["shipper"]]["usd"] += r["usd"]; by_shipper[r["shipper"]]["n"] += 1
        total_usd += r["usd"]
    top_suppliers = sorted(by_supplier.items(), key=lambda kv: -kv[1]["usd"])[:20]
    data["current_season"] = {
        "status": "en_revision",
        "total_usd": round(total_usd, 2),
        "n_lineas": len(gaopex_rows),
        "n_rows_total": n_rows_total,
        "n_rows_sin_clasificar": n_rows_total - len(gaopex_rows),
        "total_usd_all_rows": round(total_usd_all_rows, 2),
        "n_matched_con_via_fecha": n_matched,
        "n_bookings_en_horizon": len(confirmed_all),
        "n_bookings_solo_en_bd2526": 0,
        "n_bookings_con_kg": n_bookings_con_kg,
        "matched_usd": round(sum(r["usd"] for r in matched), 2),
        "by_concepto": [{"concepto": k, "usd": round(v["usd"], 2), "n": v["n"]}
                         for k, v in sorted(by_concepto.items(), key=lambda kv: -kv[1]["usd"])],
        "by_opex": [{"opex": k, "usd": round(v["usd"], 2), "n": v["n"]}
                    for k, v in sorted(by_grupo.items(), key=lambda kv: -kv[1]["usd"])],
        "by_shipper": [{"shipper": k, "usd": round(v["usd"], 2), "n": v["n"]}
                       for k, v in sorted(by_shipper.items(), key=lambda kv: -kv[1]["usd"])],
        "top_proveedores": [{"proveedor": k, "usd": round(v["usd"], 2), "n": v["n"]}
                            for k, v in top_suppliers],
    }

    # -- totales / meta --
    total_usd_by_season = defaultdict(float)
    n_by_season = defaultdict(int)
    n_bookings_by_season = defaultdict(set)
    for r in data["weekly"]:
        if r.get("p"):
            continue
        total_usd_by_season[r["s"]] += r["usd"]
        n_by_season[r["s"]] += r["n"]
    for b in data["bookings"]:
        n_bookings_by_season[b["s"]].add(b["bk"])
    data["meta"]["n_lineas_total"] = sum(n_by_season.values())
    data["meta"]["n_bookings_total"] = sum(len(v) for v in n_bookings_by_season.values())
    data["meta"]["total_usd_total"] = round(sum(total_usd_by_season.values()), 2)
    data["meta"]["por_temporada"]["T26-27"] = {
        "usd": round(total_usd_by_season.get(SEASON, 0), 2),
        "n_lineas": n_by_season.get(SEASON, 0),
        "n_bookings": len(n_bookings_by_season.get(SEASON, set())),
    }
    data["meta"]["opex_groups"] = sorted({c["opex"] for c in data["concepts"]} |
                                          {r["o"] for r in data["weekly"] if r["s"] == SEASON})

    semanas = sorted({r["w"] for r in projected_rows})
    data["meta"]["forecast"] = {
        "n_instrucciones_proyectadas": len(projected_instr),
        "usd_proyectado_total": round(total_proj_usd, 2),
        "semanas_cubiertas": [semanas[0], semanas[-1]] if semanas else None,
        "n_bookings_historicos_por_via_exportador": {
            f"{v} · {ex}": n for (v, ex), n in n_bookings_hist.items()
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
    data["meta"]["current_season_alert"] = (
        f"GAOPEX_T2627 (temporada 2026-2027) está en proceso de revisión y "
        f"aprobación por el equipo de liquidación. De {n_rows_total} líneas del "
        f"archivo, {len(gaopex_rows)} ya tienen Concepto/Clasificación OPEX asignados. "
        f"El archivo GAOPEX no trae Vía ni Fecha de Despacho por sí solo: se cruza por "
        f"Booking (y, si no matchea, por PO como respaldo — {n_matched_by_po} líneas se "
        f"recuperaron así) contra Horizon (hoja HORIZON_FORECAST), que solo registra "
        f"Vía/Fecha/Kg "
        f"para embarques ya CONFIRMADOS ({len(confirmed_all)} bookings con fecha real al "
        f"día de hoy; el resto de Horizon es forecast a futuro). Del total clasificado, "
        f"{n_matched} líneas ({fmt_usd(sum(r['usd'] for r in matched))}) ya cruzaron con "
        f"Vía+Fecha y se incorporan a la línea de tiempo semanal, el comparativo anual y "
        f"la tabla de bookings junto con T23-24/T24-25/T25-26; el resto sigue solo en el "
        f"panel 'Temporada actual' hasta que tenga booking confirmado en Horizon o "
        f"clasificación OPEX. Actualiza estos archivos y vuelve a correr el ETL para "
        f"ampliar la cobertura del cruce."
    )
    avg_kg_txt = " · ".join(f"{v}: {avg_kg_by_via[v]:,.0f} kg/PO" for v in avg_kg_by_via)
    data["meta"]["kg_note"] = (
        "El modo USD/Kg usa el peso neto (Net weight) real de los despachos "
        "CONFIRMADOS de T26-27 en Horizon donde está registrado "
        f"({n_bookings_con_kg} de {len(confirmed_all)} bookings confirmados; "
        "se detectaron ~20 filas con el peso cargado en gramos en vez de kg "
        "-valores >100,000- y se normalizaron ÷1000), y donde falta se completa "
        f"con el promedio de Kg/PO de esa Vía ({n_bookings_imputados} bookings "
        "imputados). Para el tramo PROYECTADO de T26-27 (línea punteada), se "
        f"estima con ese mismo promedio ({avg_kg_txt}) × Nº de instrucciones "
        "proyectadas de esa semana — mismo criterio que la proyección en USD."
    )

    with open(DATA_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    print("OK ->", DATA_JSON)
    print("n_matched:", n_matched, f"(de los cuales {n_matched_by_po} solo matchearon por PO)",
          "total matched usd:", round(sum(r['usd'] for r in matched), 2))
    print("projected usd:", round(total_proj_usd, 2), "n projected rows:", len(projected_rows))
    print("kg buckets:", len(weekly_kg_rows), "bookings con kg:", n_bookings_con_kg)
    print("avg kg/PO by via:", {k: round(v,1) for k,v in avg_kg_by_via.items()})
    print("projected kg rows:", len(projected_kg_rows), "total projected kg:", round(total_proj_kg,1))


def fmt_usd(v):
    return f"${v:,.0f}"


if __name__ == "__main__":
    main()
