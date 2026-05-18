"""
pipeline.py
───────────
Procesa el Excel "Follow Up clientes_2026.xlsm" y devuelve un dict
JSON-ready con todas las estructuras que consume el dashboard.

Está pensado para correr dentro de Pyodide (browser). La función pública es:

    compute_data(xlsm_input) -> dict

donde xlsm_input puede ser:
    - bytes (contenido del Excel — caso Pyodide)
    - path str/Path (caso local para testing)

Refactor de generar_dashboard.py — sólo la parte de COMPUTE.
El HTML del 13-may ya tiene todas las modificaciones estructurales; este
módulo entrega data, el JS la inyecta en variables y re-renderiza.
"""

from __future__ import annotations

import hashlib
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd

# ════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ════════════════════════════════════════════════════════════════════════════

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

THRESH = {"HP": 75, "MP": 120, "LP": 210}
PRIO_ORD = {"HP": 0, "MP": 1, "LP": 2, "?": 3}

SUBTEMA_MAP = {
    "update": "Follow Up / Catch Up", "follow up": "Follow Up / Catch Up", "catch up": "Follow Up / Catch Up",
    "update follow up": "Follow Up / Catch Up", "update seguimiento": "Follow Up / Catch Up",
    "mantencion seguimiento": "Follow Up / Catch Up", "seguimiento mantencion": "Follow Up / Catch Up",
    "seguimiento": "Follow Up / Catch Up", "mantencion": "Follow Up / Catch Up",
    "almuerzo": "Follow Up / Catch Up", "almuerzo socios": "Follow Up / Catch Up",
    "almuerzo con socio": "Follow Up / Catch Up", "desayuno encuesta ameris 2026": "Follow Up / Catch Up",
    "primera reunion": "Follow Up / Catch Up", "reunion inicial": "Follow Up / Catch Up",
    "introduccion": "Follow Up / Catch Up", "rebates": "Follow Up / Catch Up",
    "planilla rebates": "Follow Up / Catch Up", "planilla": "Follow Up / Catch Up",
    "contrato": "Follow Up / Catch Up", "pipeline": "Follow Up / Catch Up",
    "pipeline 2023": "Follow Up / Catch Up", "pipeline 2024": "Follow Up / Catch Up",
    "pipeline 2025": "Follow Up / Catch Up", "gobierno corporativo": "Follow Up / Catch Up",
    "after office": "Follow Up / Catch Up", "cafe": "Follow Up / Catch Up",
    "pichanga futbol": "Follow Up / Catch Up", "evento": "Follow Up / Catch Up",
    "webinar": "Follow Up / Catch Up", "operacional": "Follow Up / Catch Up",
    "mail": "Follow Up / Catch Up", "llamado": "Follow Up / Catch Up",
    "reu con pm": "Follow Up / Catch Up", "meeting": "Follow Up / Catch Up",
    "seguimiento y mantencion": "Follow Up / Catch Up",
    "adi 6": "ADI 6", "adi vi": "ADI 6", "adi 6 reunion cn pm": "ADI 6", "adi 6 / newcred": "ADI 6",
    "adi 6 y fcp": "ADI 6", "ddq adi vi y fogape": "ADI 6",
    "reu con cliente. va a entrar con 50 mm al adi vi": "ADI 6",
    "pm adi 6 y fcp": "ADI 6", "roadshow con pm adi 6": "ADI 6", "reu adi 6 con pm": "ADI 6",
    "adi 6 y deuda directa": "ADI 6", "deuda privada y adi 6": "ADI 6",
    "deuda directa y adi vi": "ADI 6", "adi ix": "ADI IX",
    "fogape": "Fogape", "deuda con fogape": "Fogape", "deuda fogape": "Fogape",
    "deuda directa": "Deuda Directa", "deuda directa usd": "Deuda Directa USD",
    "deuda directa dólar": "Deuda Directa USD", "deuda usd": "Deuda Directa USD",
    "deuda dólar": "Deuda Directa USD", "deuda privada": "Deuda Privada",
    "aa tech financiamiento corto plazo": "AA Tech / FCP",
    "financiamiento corto plazo": "AA Tech / FCP", "aa tech fcp": "AA Tech / FCP",
    "aa tech": "AA Tech / FCP", "bts ii": "BTS II",
    "newcred": "NewCred", "new cred": "NewCred",
    "salfacorp preferente": "Salfacorp Preferente", "salfa corp preferente": "Salfacorp Preferente",
    "salfa": "Salfacorp Preferente", "nmg iv": "NMG IV", "nmgiv": "NMG IV", "nmg iii": "NMG III",
    "crescent iii": "Crescent CESL III", "crescent cesl iii": "Crescent CESL III", "cesl iii": "Crescent CESL III",
    "inmobiliario españa": "Inmobiliario España",
    "renta residencial serire preferente": "Renta Residencial",
    "serie preferente renta residencial": "Renta Residencial", "renta residencial": "Renta Residencial",
    "nordic": "Nordic", "nordic roadshow": "Nordic", "harbourvest": "HarbourVest",
    "warburg pincus": "Warburg Pincus", "alternativos 3": "Alternativos III",
    "alternativos iii": "Alternativos III", "alternativos 1": "Alternativos I",
    "alternativos 2": "Alternativos II", "alternativos i": "Alternativos I",
    "alternativos ii": "Alternativos II", "best ideas": "Best Ideas", "aip": "AIP",
    "notas estructuradas": "Notas Estructuradas", "renta directa": "Renta Directa",
    "renta comercial": "Renta Comercial", "electromovillidad": "Electromovilidad",
    "electromovilidad": "Electromovilidad", "dva": "DVA", "trg": "TRG",
    "garantias inmobiliarias": "Garantías Inmobiliarias", "southwind": "Southwind",
    "view deuda privada": "Deuda Privada", "zañartu": "Zañartu", "megacentro": "Megacentro",
    "parauco": "Parauco", "parauco outlets": "Parauco", "iberia": "Iberia",
    "deuda iberia reu pm": "Iberia",
}

SUB_COLORS_FIXED = {
    "Follow Up / Catch Up": "#8FA3BE", "Deuda Directa": "#0E7A4E", "Deuda Directa USD": "#0A5C3A",
    "Deuda Privada": "#1A7A5A", "ADI 6": "#2E6BAF", "ADI IX": "#3A7BC8", "AA Tech / FCP": "#1B4B9B",
    "NewCred": "#B33A2E", "Fogape": "#D4890A", "Renta Directa": "#4A9B7A",
    "Notas Estructuradas": "#6B4C9B", "Electromovilidad": "#895200", "BTS II": "#0D2D6B",
    "NMG IV": "#7FB3E8", "Warburg Pincus": "#7FB3E8", "Nordic": "#A8C5E0",
    "HarbourVest": "#B8D5E8", "Sin tema": "#E0E4EA",
}

BASE_COLORS = [
    "#1B4B9B", "#2E6BAF", "#0E7A4E", "#5A93CC", "#D4890A", "#B33A2E",
    "#7FB3E8", "#8FA3BE", "#6B4C9B", "#C0C9D8", "#0D2D6B", "#3A7BC8",
]

AP_ALIASES = {
    "itau": "itau banca privada", "larrain vial": "larraín vial banca privada",
    "augustar": "augustar peru", "credicorp": "credicorp banca privada", "fynsa": "fynsa banca privada",
}


# ════════════════════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════════════════════

def _clean_sub(s) -> str | None:
    if pd.isna(s) or str(s).strip() == "" or len(str(s).strip()) > 60:
        return None
    sl = str(s).strip().lower()
    r = SUBTEMA_MAP.get(sl)
    if r:
        return r
    s2 = str(s).strip()
    return s2.title() if len(s2) <= 35 else None


def _sub_color(name: str) -> str:
    if name in SUB_COLORS_FIXED:
        return SUB_COLORS_FIXED[name]
    return BASE_COLORS[int(hashlib.md5(name.encode()).hexdigest(), 16) % len(BASE_COLORS)]


def _load_foto(xls, sheet: str) -> pd.DataFrame:
    """
    Lee una hoja 'Clientes foto YYYY' y agrega `_tipo` ('Dist' / 'Inst')
    según las filas separadoras "Distribuidores" / "Institucionales" que
    el Excel usa como section headers en la columna Cliente.
    """
    df = pd.read_excel(xls, sheet_name=sheet)

    # Recorrer en orden y trackear la sección actual basada en las filas
    # separadoras del Excel. Cualquier fila de cliente real hereda la
    # última sección vista.
    current_section: str | None = None
    tipos: list[str | None] = []
    for cliente_val in df["Cliente"]:
        cl = str(cliente_val).strip().lower() if pd.notna(cliente_val) else ""
        if "distribu" in cl:
            current_section = "Dist"
        elif cl.startswith("institucional"):  # "Institucionales" / "Institucional"
            current_section = "Inst"
        tipos.append(current_section)
    df["_tipo"] = tipos

    # Ahora sí filtramos los separadores y totales para quedarnos sólo con clientes
    df = df[
        df["Cliente"].notna()
        & ~df["Cliente"].astype(str).str.contains("Total|Distribu|Inst", case=False, na=False)
    ].fillna(0)
    df["_cn"] = df["Cliente"].astype(str).str.strip().str.lower()
    return df


def _monthly(df: pd.DataFrame) -> list[int]:
    return [int(df[m].sum()) if m in df.columns else 0 for m in MESES]


def _top10(df: pd.DataFrame, cols: list[str]) -> list:
    df = df.copy()
    df["_t"] = df[[c for c in cols if c in df.columns]].sum(axis=1).astype(float)
    top = df.nlargest(10, "_t")[["Cliente", "_t"]]
    return [[str(r["Cliente"]), int(r["_t"])] for _, r in top.iterrows() if r["_t"] > 0]


def _foto_counts(df_f: pd.DataFrame, yr: int) -> dict:
    c = {}
    for _, row in df_f.iterrows():
        for i, m in enumerate(MESES, 1):
            if m in df_f.columns and row[m] > 0:
                c[(yr, i, row["_cn"])] = c.get((yr, i, row["_cn"]), 0) + int(row[m])
    return c


# ════════════════════════════════════════════════════════════════════════════
# COMPUTE
# ════════════════════════════════════════════════════════════════════════════

def compute_data(xlsm_input, today: date | None = None) -> dict:
    """
    Procesa el Excel y devuelve un dict con todas las variables que el
    dashboard usa.

    Args:
        xlsm_input: bytes, path str, o Path
        today: fecha de corte (default: date.today())
    """
    if isinstance(xlsm_input, (bytes, bytearray, memoryview)):
        xls_handle = BytesIO(bytes(xlsm_input))
    elif isinstance(xlsm_input, (str, Path)):
        xls_handle = str(xlsm_input)
    else:
        xls_handle = xlsm_input  # asumir file-like

    today_ts = pd.Timestamp(today or date.today())

    # ── LOAD ────────────────────────────────────────────────────────────
    xls = pd.ExcelFile(xls_handle, engine="openpyxl")

    # Auto-detectar años disponibles desde los nombres de las hojas
    # "Clientes foto YYYY". Soporta cualquier año (2023, 2024, ... 2027, ...).
    import re as _re
    foto_sheets = {}
    for s in xls.sheet_names:
        m = _re.match(r"Clientes foto (\d{4})", s)
        if m:
            foto_sheets[int(m.group(1))] = s
    years_avail = sorted(foto_sheets.keys())
    if not years_avail:
        raise ValueError("No encontré ninguna hoja 'Clientes foto YYYY' en el Excel.")

    year_dfs = {yr: _load_foto(xls, foto_sheets[yr]) for yr in years_avail}

    # Año "actual" = año de hoy si tiene hoja, si no el más reciente disponible
    cur_year = today_ts.year if today_ts.year in year_dfs else max(year_dfs.keys())
    df_cur = year_dfs[cur_year]
    # Aliases por legibilidad: y0=actual, y1=1 año atrás, etc.
    def _df_offset(off):
        return year_dfs.get(cur_year - off, pd.DataFrame())
    df_p1 = _df_offset(1)
    df_p2 = _df_offset(2)
    df_p3 = _df_offset(3)

    prio_dict = dict(zip(df_cur["_cn"], df_cur.get("Prioridad", pd.Series(dtype=object))))
    name_dict = dict(zip(df_cur["_cn"], df_cur["Cliente"]))
    tipo_dict = {cn: (t or "Dist") for cn, t in zip(df_cur["_cn"], df_cur["_tipo"])}

    # Apuntes — los nombres de las hojas tienen espacios variables, los probamos
    # con un patrón laxo (`Apuntes\s+YYYY\s*`).
    apuntes_sheets = {}
    for s in xls.sheet_names:
        m = _re.match(r"Apuntes\s+(\d{4})\s*$", s)
        if m:
            apuntes_sheets[int(m.group(1))] = s
    ap_dfs = []
    for yr, sheet_name in apuntes_sheets.items():
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df["_year"] = yr
            ap_dfs.append(df)
        except Exception:
            pass

    if ap_dfs:
        df_ap = pd.concat(ap_dfs, ignore_index=True)
    else:
        df_ap = pd.DataFrame(columns=["Fecha", "Empresa/Cliente", "Subtema"])

    df_ap["Fecha"] = pd.to_datetime(df_ap["Fecha"], errors="coerce")
    df_ap = df_ap[df_ap["Fecha"].notna()].copy()
    df_ap["month_num"] = df_ap["Fecha"].dt.month
    df_ap["year_num"] = df_ap["Fecha"].dt.year
    df_ap["cn"] = (
        df_ap["Empresa/Cliente"].astype(str).str.strip().str.lower().replace(AP_ALIASES)
    )
    last_contact = df_ap.groupby("cn")["Fecha"].max() if len(df_ap) else pd.Series(dtype="datetime64[ns]")

    # ── MONTHLY SUMS (por año disponible) ───────────────────────────────
    monthly_by_year = {yr: _monthly(year_dfs[yr]) for yr in years_avail}
    d_cur = monthly_by_year[cur_year]
    ytd = sum(d_cur[: today_ts.month])

    # ── ÚLTIMO TRIMESTRE COMPLETO ───────────────────────────────────────
    # Hoy en QN → último completo es Q(N-1) del mismo año (Q4 año anterior si hoy en Q1)
    cur_month = today_ts.month
    if cur_month <= 3:
        last_q_year, last_q = cur_year - 1, 4
    elif cur_month <= 6:
        last_q_year, last_q = cur_year, 1
    elif cur_month <= 9:
        last_q_year, last_q = cur_year, 2
    else:
        last_q_year, last_q = cur_year, 3
    last_q_months = list(range((last_q - 1) * 3 + 1, last_q * 3 + 1))
    last_q_month_names = [MESES[m - 1] for m in last_q_months]
    last_q_label = f"Q{last_q} {last_q_year}"

    # Q1 del año actual (siempre los 3 primeros meses) — para compat con KPIs viejos
    q1_cur = sum(d_cur[:3])

    # Suma de reuniones del trimestre completo (en el df del año del Q)
    df_q = year_dfs.get(last_q_year, df_cur)
    reuniones_q = sum(int(df_q[m].sum()) if m in df_q.columns else 0 for m in last_q_month_names)
    df_q_prev = year_dfs.get(last_q_year - 1)
    reuniones_q_prev = (
        sum(int(df_q_prev[m].sum()) if m in df_q_prev.columns else 0 for m in last_q_month_names)
        if df_q_prev is not None and not df_q_prev.empty else 0
    )
    pct_change_q = (
        round((reuniones_q - reuniones_q_prev) / reuniones_q_prev * 100)
        if reuniones_q_prev else 0
    )

    # ── TOPS (por año, anual) ───────────────────────────────────────────
    clients_by_year = {yr: _top10(year_dfs[yr], MESES) for yr in years_avail}

    all_t = {}
    # Suma todos los años disponibles para "clientsData.all"
    for dfx in year_dfs.values():
        for _, row in dfx.iterrows():
            all_t[row["_cn"]] = all_t.get(row["_cn"], 0) + sum(
                float(row.get(m, 0)) for m in MESES if m in dfx.columns
            )
    clients_all = sorted(
        [[name_dict.get(cn, cn.title()), round(v)] for cn, v in all_t.items() if v > 0],
        key=lambda x: -x[1],
    )[:10]

    # ── TOP CLIENTES DEL ÚLTIMO TRIMESTRE COMPLETO ──────────────────────
    q_cols = [m for m in last_q_month_names if m in df_q.columns]
    df_q_local = df_q.copy()
    df_q_local["_qsum"] = df_q_local[q_cols].sum(axis=1).astype(float) if q_cols else 0.0
    top_q_df = df_q_local.nlargest(10, "_qsum")[["Cliente", "_qsum"]]
    clients_q = [
        [str(r["Cliente"]), int(r["_qsum"])]
        for _, r in top_q_df.iterrows() if r["_qsum"] > 0
    ]

    # ── PRIO DEL ÚLTIMO TRIMESTRE (top 8 por prio HP/MP/LP) ─────────────
    prio_data = {}
    if "Prioridad" in df_q_local.columns:
        for prio in ["HP", "MP", "LP"]:
            s = df_q_local[df_q_local["Prioridad"] == prio].nlargest(8, "_qsum")
            prio_data[prio] = [
                [str(r["Cliente"]), int(r["_qsum"])]
                for _, r in s.iterrows() if r["_qsum"] > 0
            ]
    else:
        prio_data = {"HP": [], "MP": [], "LP": []}

    # ── RECONTACT (año anterior → año actual YTD) ───────────────────────
    if not df_p1.empty:
        df_p1_local = df_p1.copy()
        df_p1_local["_cn2"] = df_p1_local["Cliente"].astype(str).str.strip().str.lower()
        df_p1_local["_totp1"] = df_p1_local[[m for m in MESES if m in df_p1_local.columns]].sum(axis=1).astype(float)
        had_prev = set(df_p1_local[df_p1_local["_totp1"] > 0]["_cn2"])
    else:
        had_prev = set()

    df_cur = df_cur.copy()
    df_cur["_ytd"] = df_cur[[m for m in MESES[: today_ts.month] if m in df_cur.columns]].sum(axis=1).astype(float)
    had_ytd = set(df_cur[df_cur["_ytd"] > 0]["_cn"])

    recontact_data = {}
    for prio_low in ["hp", "mp", "lp"]:
        if "Prioridad" not in df_cur.columns:
            recontact_data[prio_low] = {"yes": [], "no": []}
            continue
        sub = df_cur[df_cur["Prioridad"] == prio_low.upper()]
        recontact_data[prio_low] = {
            "yes": [str(x) for x in sub[sub["_cn"].isin(had_prev) & sub["_cn"].isin(had_ytd)]["Cliente"]],
            "no": [str(x) for x in sub[sub["_cn"].isin(had_prev) & ~sub["_cn"].isin(had_ytd)]["Cliente"]],
        }

    # ── SEMÁFORO ────────────────────────────────────────────────────────
    semaforo = {}
    for prio, meta in [("HP", 2), ("MP", 1), ("LP", 1)]:
        if "Prioridad" not in df_cur.columns:
            semaforo[prio] = {"en_meta": 0, "parcial": 0, "sin_cont": 0, "total": 0}
            continue
        sub = df_cur[df_cur["Prioridad"] == prio]["_ytd"]
        semaforo[prio] = {
            "en_meta": int((sub >= meta).sum()),
            "parcial": int((sub == 1).sum()) if meta == 2 else 0,
            "sin_cont": int((sub == 0).sum()),
            "total": len(sub),
        }

    # ── RISK (año actual) + RISK ALL ────────────────────────────────────
    risk_cur = {"HP": [], "MP": [], "LP": []}
    risk_all = {"HP": [], "MP": [], "LP": []}
    for _, row in df_cur.iterrows():
        cn = row["_cn"]
        prio = prio_dict.get(cn)
        if prio not in THRESH:
            continue
        name = str(row["Cliente"])
        last = last_contact.get(cn) if len(last_contact) else None
        dias = int((today_ts - last).days) if last is not None and not pd.isna(last) else None
        last_str = str(last.date()) if last is not None and not pd.isna(last) else "Sin historial"
        last_yr = last.year if last is not None and not pd.isna(last) else 0
        reun_cur = int(sum(float(row.get(m, 0)) for m in MESES if m in df_cur.columns))
        tipo = row.get("_tipo") or "Dist"
        # "activo" si tuvo reuniones en el año actual. "pendiente" si último contacto en
        # el año anterior. "inactivo" si nada en el año anterior tampoco.
        status = "activo" if reun_cur > 0 else ("pendiente" if last_yr >= cur_year - 1 else "inactivo")
        risk_cur[prio].append({
            "name": name, "prio": prio, "dias": dias, "last": last_str,
            "tipo": tipo, "reun26": reun_cur, "status": status, "last_yr": last_yr,
        })
        risk_all[prio].append({"name": name, "dias": dias, "last": last_str, "tipo": tipo})

    for prio in risk_cur:
        risk_cur[prio].sort(
            key=lambda x: ({"inactivo": 0, "pendiente": 1, "activo": 2}[x["status"]], -(x["dias"] or 9999))
        )
    for prio in risk_all:
        risk_all[prio].sort(key=lambda x: -(x["dias"] or 9999))

    kpi = {}
    for prio in ["HP", "MP", "LP"]:
        kpi[prio] = {k: sum(1 for r in risk_cur[prio] if r["status"] == k) for k in ["activo", "pendiente", "inactivo"]}
        kpi[prio]["total"] = len(risk_cur[prio])

    # ── CRITICAL LIST ───────────────────────────────────────────────────
    critical = []
    for prio, items in risk_all.items():
        for r in items:
            if (r["dias"] or 9999) > THRESH[prio]:
                critical.append({
                    "name": r["name"], "prio": prio,
                    "dias": r["dias"] or 0, "last": r["last"],
                })
    critical.sort(key=lambda x: -x["dias"])
    critical = critical[:25]

    # ── MONTHLY DETAIL ──────────────────────────────────────────────────
    monthly_detail = {}
    if len(df_ap):
        for (yr, mo, cn), grp in df_ap.groupby(["year_num", "month_num", "cn"]):
            key = f"{yr}_{mo}"
            if key not in monthly_detail:
                monthly_detail[key] = []
            subs = list(dict.fromkeys(
                [str(s).strip() for s in grp["Subtema"].dropna() if str(s).strip()]
            ))[:3]
            nombre = name_dict.get(cn, cn.title())
            monthly_detail[key].append({
                "nombre": nombre,
                "prio": prio_dict.get(cn, "?"),
                "tipo": tipo_dict.get(cn, "Dist"),
                "reun": len(grp),
                "subtemas": subs,
            })
    for key in monthly_detail:
        monthly_detail[key].sort(key=lambda x: (PRIO_ORD.get(x["prio"], 3), -x["reun"]))

    # ── SUBTEMAS ────────────────────────────────────────────────────────
    foto_all = {}
    for yr, dfx in year_dfs.items():
        if not dfx.empty:
            foto_all.update(_foto_counts(dfx, yr))

    df_ap_local = df_ap.copy()
    df_ap_local["sub_clean"] = df_ap_local["Subtema"].apply(_clean_sub)
    df_ap_local["_sort"] = df_ap_local["sub_clean"].apply(
        lambda x: 0 if (x and x != "Follow Up / Catch Up") else (1 if x == "Follow Up / Catch Up" else 2)
    )
    df_ap_s = df_ap_local.sort_values(["year_num", "month_num", "cn", "_sort"])

    final_rows = []
    for (yr, mo, cn), fc in foto_all.items():
        ap_subs = df_ap_s[
            (df_ap_s["year_num"] == yr) & (df_ap_s["month_num"] == mo)
            & (df_ap_s["cn"] == cn) & (df_ap_s["sub_clean"].notna())
        ]["sub_clean"].tolist()
        for i in range(fc):
            final_rows.append({
                "year_num": yr, "month_num": mo, "cn": cn,
                "sub_clean": ap_subs[i] if i < len(ap_subs) else "Sin tema",
            })
    df_sub = pd.DataFrame(final_rows) if final_rows else pd.DataFrame(columns=["year_num", "month_num", "cn", "sub_clean"])

    sm, sa, scm, sca = {}, {}, {}, {}
    if len(df_sub):
        for (yr, mo), g in df_sub.groupby(["year_num", "month_num"]):
            sm[f"{yr}_{mo}"] = g["sub_clean"].value_counts().to_dict()
        for yr in years_avail:
            sa[str(yr)] = df_sub[df_sub["year_num"] == yr]["sub_clean"].value_counts().to_dict()

        def _build_cli(dfg):
            r = {}
            for (sub, cn), g in dfg.groupby(["sub_clean", "cn"]):
                r.setdefault(sub, []).append({
                    "name": name_dict.get(cn, cn.title()),
                    "prio": prio_dict.get(cn, "?"),
                    "reun": len(g),
                })
            for sub in r:
                r[sub].sort(key=lambda x: (PRIO_ORD.get(x["prio"], 3), -x["reun"]))
            return r

        for (yr, mo), g in df_sub.groupby(["year_num", "month_num"]):
            scm[f"{yr}_{mo}"] = _build_cli(g)
        for yr in years_avail:
            sca[str(yr)] = _build_cli(df_sub[df_sub["year_num"] == yr])
        all_subs = sorted(set(df_sub["sub_clean"].dropna().unique()))
        sub_colors = {s: _sub_color(s) for s in all_subs}
    else:
        sub_colors = {}

    # ── RECENT 5 LABELS / VALS (para el peakChart) ──────────────────────
    # Toma los 5 meses anteriores al mes en curso (incluyendo si cruza años).
    MES_ABBR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    cur_m = today_ts.month
    recent_labels: list[str] = []
    recent_vals: list[int] = []
    for offset in range(5, 0, -1):
        m = cur_m - offset
        y = cur_year
        while m <= 0:
            m += 12
            y -= 1
        recent_labels.append(f"{MES_ABBR[m-1]} {y}")
        recent_vals.append(int(monthly_by_year.get(y, [0]*12)[m-1]) if y in monthly_by_year else 0)

    # ── REUNIONES YTD POR PRIORIDAD (para reunYtdChart) ─────────────────
    reun_ytd_by_prio = {
        prio: sum(int(r["reun26"]) for r in risk_cur[prio])
        for prio in ["HP", "MP", "LP"]
    }

    # ── RESULT ──────────────────────────────────────────────────────────
    # data y clientsData keyed por año (string) — dinámico según años disponibles.
    data_out = {str(yr): monthly_by_year[yr] for yr in years_avail}
    clients_out = {"all": clients_all}
    for yr in years_avail:
        clients_out[str(yr)] = _top10(year_dfs[yr], MESES)

    # Reuniones del año anterior en el MISMO mes en curso (para "vs YTD año anterior")
    df_prev_year = year_dfs.get(cur_year - 1)
    ytd_prev = (
        sum(int(df_prev_year[m].sum()) if m in df_prev_year.columns else 0
            for m in MESES[: today_ts.month])
        if df_prev_year is not None and not df_prev_year.empty else 0
    )
    pct_change_ytd = round((ytd - ytd_prev) / ytd_prev * 100) if ytd_prev else 0

    return {
        "data": data_out,
        "clientsData": clients_out,
        "prioData": prio_data,
        "recontactData": recontact_data,
        "riskAll": risk_all,
        "criticalList": critical,
        "monthlyDetailData": monthly_detail,
        "subtemaMoData": sm,
        "subteamAnnual": sa,
        "subtemasColors": sub_colors,
        "subtemasClientsM": scm,
        "subtemasClientsA": sca,
        "riskData2026": risk_cur,  # legacy key
        "riskCurrent": risk_cur,
        "topClientesQ": clients_q,
        "recentLabels": recent_labels,
        "recentVals": recent_vals,
        "years": years_avail,
        "currentYear": cur_year,
        "currentQ": {
            "year": last_q_year,
            "q": last_q,
            "label": last_q_label,           # "Q1 2026"
            "label_year_prev": f"Q{last_q} {last_q_year - 1}",
            "months": last_q_months,         # [1,2,3]
            "monthNames": last_q_month_names,
            "reuniones": reuniones_q,
            "reuniones_prev": reuniones_q_prev,
            "pct_change": pct_change_q,
        },
        "kpi": {
            "ytd": ytd,
            "ytd_prev": ytd_prev,
            "pct_change_ytd": pct_change_ytd,
            "q1": q1_cur,
            "current_q_total": reuniones_q,
            "current_q_prev": reuniones_q_prev,
            "abr": d_cur[3] if len(d_cur) > 3 else 0,
            "may": d_cur[4] if len(d_cur) > 4 else 0,
            "riesgo_totals": {
                "activo": sum(kpi[p]["activo"] for p in ["HP", "MP", "LP"]),
                "pendiente": sum(kpi[p]["pendiente"] for p in ["HP", "MP", "LP"]),
                "inactivo": sum(kpi[p]["inactivo"] for p in ["HP", "MP", "LP"]),
                "total": sum(kpi[p]["total"] for p in ["HP", "MP", "LP"]),
            },
            "riesgo_by_prio": kpi,
            "semaforo": semaforo,
            "reun_ytd_by_prio": reun_ytd_by_prio,
        },
        "meta": {
            "n_clientes": int(len(df_cur)),
            "fecha_corte": str(today_ts.date()),
            "years_disponibles": years_avail,
        },
    }
