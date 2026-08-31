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
import re
from datetime import date, timedelta
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


def _norm_entity(s) -> str:
    """
    Normaliza el nombre de una entidad (cliente / advisor / distribuidor) para
    cruzar entre fuentes (gastos, foto de reuniones, saldos por corte).

    Quita sufijos legales y de tipo de cuenta ('S.A.', 'SpA', 'Banca Privada',
    'Wealth Management'…) y deja sólo alfanuméricos en minúscula.

    NO colapsa 'GR Capital' con 'Grey Capital' — quedan 'grcapital' vs
    'greycapital' (son entidades distintas, confirmado con el área).
    """
    s = str(s).strip().lower()
    s = re.sub(
        r"\b(s\.?a\.?|spa|ltda|limitada|s\.?l\.?|banca privada|wealth management|wmg)\b",
        "", s,
    )
    return re.sub(r"[^a-z0-9]", "", s)


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


def _corte_y_trimestre(
    d_cur: list[int], cur_month: int, es_anio_actual: bool
) -> tuple[int, int, list[int]]:
    """Decide hasta qué mes corta el dashboard y qué trimestre muestra.

    Dos reglas distintas, a propósito:

    - El TRIMESTRE lo manda el calendario. Si abro el dashboard en diciembre
      quiero ver Q4, aunque el Excel venga con un mes de rezago.
    - El MES DE CORTE lo manda la data: es el mes en curso, pero acotado al
      último mes con reuniones cargadas. Así el YoY nunca compara un tramo
      parcial de este año contra meses completos del anterior, ni el foto
      muestra columnas en blanco.

    Devuelve (ref_month, trimestre, meses_del_trimestre_con_data). La lista
    de meses queda vacía si el trimestre ya empezó pero no tiene data — el
    caller lo reporta como "sin datos cargados" en vez de mostrar un 0.

    >>> _corte_y_trimestre([5]*7 + [0]*5, 8, True)   # ago, data hasta jul
    (7, 3, [7])
    >>> _corte_y_trimestre([5]*7 + [0]*5, 12, True)  # dic, data hasta jul
    (7, 4, [])
    >>> _corte_y_trimestre([5]*11 + [0], 12, True)   # dic, data hasta nov
    (11, 4, [10, 11])
    >>> _corte_y_trimestre([5]*12, 12, True)         # dic, año completo
    (12, 4, [10, 11, 12])
    >>> _corte_y_trimestre([5]*7 + [0]*5, 3, False)  # foto de un año anterior
    (7, 4, [])
    """
    # Techo: el mes en curso si el foto es del año de hoy; si el foto es de un
    # año anterior, el año ya terminó y el techo es diciembre.
    cap_month = cur_month if es_anio_actual else 12
    last_month_data = max((i + 1 for i, v in enumerate(d_cur) if v > 0), default=0)
    ref_month = min(cap_month, last_month_data) if last_month_data else cap_month
    q = (cap_month - 1) // 3 + 1
    meses = [m for m in range((q - 1) * 3 + 1, q * 3 + 1) if m <= ref_month]
    return ref_month, q, meses


def _top10(df: pd.DataFrame, cols: list[str]) -> list:
    df = df.copy()
    df["_t"] = df[[c for c in cols if c in df.columns]].sum(axis=1).astype(float)
    top = df.nlargest(10, "_t")[["Cliente", "_t"]]
    return [[str(r["Cliente"]), int(r["_t"])] for _, r in top.iterrows() if r["_t"] > 0]


class ExcelValidationError(Exception):
    """Error que se muestra al usuario en la UI con un mensaje claro."""
    pass


def _validate_excel(xls: pd.ExcelFile, today_ts: pd.Timestamp) -> list[str]:
    """
    Valida que el Excel tenga la estructura mínima esperada.
    - Lanza ExcelValidationError si falta algo crítico.
    - Devuelve una lista de warnings (no críticos).
    """
    warnings: list[str] = []

    # 1) Al menos una hoja "Clientes foto YYYY"
    foto_years = sorted(
        int(re.match(r"Clientes foto (\d{4})", s).group(1))
        for s in xls.sheet_names if re.match(r"Clientes foto (\d{4})", s)
    )
    if not foto_years:
        raise ExcelValidationError(
            "No encontré ninguna hoja con el nombre 'Clientes foto YYYY' "
            "(ej. 'Clientes foto 2026'). El Excel debe tener al menos una."
        )

    # 2) La hoja del año actual (o más reciente disponible) debe existir
    cur_year = today_ts.year
    if cur_year not in foto_years:
        warnings.append(
            f"No encontré 'Clientes foto {cur_year}'. Usando '{max(foto_years)}' "
            f"como año más reciente — los datos pueden estar desactualizados."
        )

    # 3) Validar columnas críticas en la hoja del año más reciente
    main_year = cur_year if cur_year in foto_years else max(foto_years)
    main_sheet = f"Clientes foto {main_year}"
    try:
        df_main = pd.read_excel(xls, sheet_name=main_sheet, nrows=5)
    except Exception as e:
        raise ExcelValidationError(
            f"No pude leer la hoja '{main_sheet}': {e}"
        )
    if "Cliente" not in df_main.columns:
        raise ExcelValidationError(
            f"La hoja '{main_sheet}' no tiene columna 'Cliente'. "
            f"Columnas encontradas: {list(df_main.columns)[:8]}..."
        )
    if "Prioridad" not in df_main.columns:
        warnings.append(
            f"La hoja '{main_sheet}' no tiene columna 'Prioridad'. "
            f"Las segmentaciones HP/MP/LP no van a funcionar."
        )
    # Verificar las filas separadoras "Distribuidores" / "Institucionales"
    df_main_full = pd.read_excel(xls, sheet_name=main_sheet)
    clientes_col = df_main_full["Cliente"].astype(str).str.strip().str.lower()
    has_dist = clientes_col.str.contains("distribu", na=False).any()
    has_inst = clientes_col.str.startswith("institucional", na=False).any()
    if not has_dist:
        warnings.append(
            f"En '{main_sheet}' no encontré la fila separadora 'Distribuidores'. "
            "Todos los clientes van a quedar etiquetados como Dist por default."
        )
    if not has_inst:
        warnings.append(
            f"En '{main_sheet}' no encontré la fila separadora 'Institucionales'. "
            "La clasificación Dist/Inst puede estar mal."
        )

    # 4) Al menos una hoja Apuntes YYYY
    apuntes_count = sum(1 for s in xls.sheet_names if re.match(r"Apuntes\s+\d{4}", s))
    if apuntes_count == 0:
        raise ExcelValidationError(
            "No encontré ninguna hoja 'Apuntes YYYY' (ej. 'Apuntes  2026'). "
            "El Excel debe tener al menos una para las reuniones."
        )

    # 5) Validar columnas críticas en Apuntes del año principal
    ap_sheet = next(
        (s for s in xls.sheet_names if re.match(rf"Apuntes\s+{main_year}\s*$", s)),
        None,
    )
    if ap_sheet:
        try:
            df_ap = pd.read_excel(xls, sheet_name=ap_sheet, nrows=5)
            crit_ap_cols = ["Fecha", "Empresa/Cliente", "Subtema"]
            missing_ap = [c for c in crit_ap_cols if c not in df_ap.columns]
            if missing_ap:
                raise ExcelValidationError(
                    f"La hoja '{ap_sheet}' no tiene columnas: {', '.join(missing_ap)}. "
                    f"Esperaba: {', '.join(crit_ap_cols)}."
                )
        except ExcelValidationError:
            raise
        except Exception as e:
            warnings.append(f"Error leyendo '{ap_sheet}': {e}")

    return warnings


def _foto_counts(df_f: pd.DataFrame, yr: int) -> dict:
    c = {}
    for _, row in df_f.iterrows():
        for i, m in enumerate(MESES, 1):
            if m in df_f.columns and row[m] > 0:
                c[(yr, i, row["_cn"])] = c.get((yr, i, row["_cn"]), 0) + int(row[m])
    return c


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

    # ── LOAD + VALIDACIÓN ───────────────────────────────────────────────
    xls = pd.ExcelFile(xls_handle, engine="openpyxl")

    # Validar estructura del Excel antes de procesar (lanza ExcelValidationError
    # si falta algo crítico; warnings los retornamos en el dict de salida)
    validation_warnings = _validate_excel(xls, today_ts)

    # Auto-detectar años disponibles desde los nombres de las hojas
    # "Clientes foto YYYY". Soporta cualquier año (2023, 2024, ... 2027, ...).
    _re = re
    foto_sheets = {}
    for s in xls.sheet_names:
        m = _re.match(r"Clientes foto (\d{4})", s)
        if m:
            foto_sheets[int(m.group(1))] = s
    years_avail = sorted(foto_sheets.keys())

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

    # ── MES DE CORTE + TRIMESTRE EN CURSO ───────────────────────────────
    ref_month, last_q, last_q_months = _corte_y_trimestre(
        d_cur, today_ts.month, today_ts.year == cur_year
    )
    ytd = sum(d_cur[:ref_month])
    last_q_year = cur_year
    last_q_month_names = [MESES[m - 1] for m in last_q_months]
    last_q_label = f"Q{last_q} {last_q_year}"
    last_q_parcial = len(last_q_months) < 3
    # El trimestre ya empezó pero el Excel todavía no trae ninguno de sus meses.
    last_q_sin_datos = not last_q_months

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
    df_cur["_ytd"] = df_cur[[m for m in MESES[:ref_month] if m in df_cur.columns]].sum(axis=1).astype(float)
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

    # ── ACTIVACIÓN DE CARTERA (4 categorías derivadas del histórico) ────
    def _cn_with_reuniones(df_in, months=None):
        if df_in is None or df_in.empty:
            return set()
        cols = [m for m in (months or MESES) if m in df_in.columns]
        if not cols:
            return set()
        totals = df_in[cols].sum(axis=1)
        return set(df_in[totals > 0]["_cn"])

    cur_q_cn = _cn_with_reuniones(df_cur, last_q_month_names) if last_q_year == cur_year else set()
    cur_any_cn = _cn_with_reuniones(df_cur)
    prev_any_cn = _cn_with_reuniones(df_p1)
    older_any_cn = _cn_with_reuniones(df_p2) | _cn_with_reuniones(df_p3)

    # nunca: HP sin ningún historial de contacto
    nunca_items = []
    if "Prioridad" in df_cur.columns:
        for _, row in df_cur[df_cur["Prioridad"] == "HP"].iterrows():
            cn = row["_cn"]
            if cn not in cur_any_cn and cn not in prev_any_cn and cn not in older_any_cn:
                nunca_items.append([str(row["Cliente"]), "HP"])

    # nuevos: primera reunión EVER en el Q actual (no historial antes)
    nuevos_items = []
    for cn in sorted(cur_q_cn):
        if cn not in prev_any_cn and cn not in older_any_cn:
            nuevos_items.append([name_dict.get(cn, cn.title()), prio_dict.get(cn, "?")])

    # reactiv: reunión en el Q actual, sin reunión en año anterior, pero con historial previo
    reactiv_items = []
    for cn in sorted(cur_q_cn):
        if cn not in prev_any_cn and cn in older_any_cn:
            reactiv_items.append([name_dict.get(cn, cn.title()), prio_dict.get(cn, "?")])

    # inactivos: sin reunión en año anterior NI en año actual (todo el año hasta hoy)
    inactivos_items = []
    inactivos_by_prio = {"HP": 0, "MP": 0, "LP": 0}
    for _, row in df_cur.iterrows():
        cn = row["_cn"]
        if cn not in cur_any_cn and cn not in prev_any_cn:
            prio = prio_dict.get(cn, "?")
            inactivos_items.append([str(row["Cliente"]), prio])
            if prio in inactivos_by_prio:
                inactivos_by_prio[prio] += 1

    activation_data = {
        "nunca":     {"label": f"HP sin historial de contacto ({len(nunca_items)})",
                      "color": "#B33A2E", "items": nunca_items,
                      "kpi_label": "Nunca contactados (HP)",
                      "kpi_sub": "HP sin historial · ▼ ver lista"},
        "nuevos":    {"label": f"Nuevos activados {last_q_label} ({len(nuevos_items)})",
                      "color": "#0E7A4E", "items": nuevos_items,
                      "kpi_label": f"Nuevos activados {last_q_label.split()[0]}",
                      "kpi_sub": f"Primera reunión en {cur_year} · ▼ ver lista"},
        "reactiv":   {"label": f"Reactivados {last_q_label} ({len(reactiv_items)})",
                      "color": "#0E7A4E", "items": reactiv_items,
                      "kpi_label": f"Reactivados (dormant {cur_year - 1})",
                      "kpi_sub": f"Tenían historial, sin reunión en {cur_year - 1} · ▼ ver lista"},
        "inactivos": {"label": f"Sin reunión en {cur_year - 1} ni {cur_year} ({len(inactivos_items)})",
                      "color": "#B33A2E", "items": inactivos_items,
                      "kpi_label": f"Sin reunión {cur_year - 1} ni {cur_year}",
                      "kpi_sub": f"HP:{inactivos_by_prio['HP']} · MP:{inactivos_by_prio['MP']} · LP:{inactivos_by_prio['LP']} · ▼ ver lista"},
    }
    activation_counts = {
        "nunca": len(nunca_items),
        "nuevos": len(nuevos_items),
        "reactiv": len(reactiv_items),
        "inactivos": len(inactivos_items),
        "inactivos_by_prio": inactivos_by_prio,
    }

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
    # Antes sumaba los 12 meses del foto y descartaba las contrapartes sin
    # prioridad válida, así que no cuadraba con kpi.ytd. Ahora corta en el mes
    # de corte y manda las no clasificadas al bucket "?", de modo que
    # HP+MP+LP+? == ytd exactamente.
    reun_ytd_by_prio = {"HP": 0, "MP": 0, "LP": 0, "?": 0}
    _cols_ytd = [m for m in MESES[:ref_month] if m in df_cur.columns]
    for _, _row in df_cur.iterrows():
        _pr = prio_dict.get(_row["_cn"])
        _pr = _pr if _pr in THRESH else "?"
        reun_ytd_by_prio[_pr] += int(sum(float(_row.get(m, 0) or 0) for m in _cols_ytd))

    # ── NUEVOS CONTACTOS (vs año anterior) ──────────────────────────────
    # Contrapartes que están en el foto del año actual pero NO estaban en el
    # del año anterior, Y tuvieron al menos una reunión este año. Para cada una
    # buscamos el subtema/fondo de su primera reunión.
    #
    # Antes esto filtraba _tipo == "Dist" y se llamaba "nuevos distribuidores",
    # pero dejaba fuera aseguradoras, AFPs y demás institucionales, que también
    # son contactos nuevos. Ahora entran todas y cada fila lleva su tipo.
    nuevos_distribuidores = []
    if not df_p1.empty:
        candidatos = sorted(set(df_cur["_cn"]) - set(df_p1["_cn"]))
        for cn in candidatos:
            row_arr = df_cur[df_cur["_cn"] == cn]
            if row_arr.empty:
                continue
            row = row_arr.iloc[0]
            # ¿Tuvo reunión en algún mes del año actual?
            primera_mes = None
            for i, m in enumerate(MESES, 1):
                if m in df_cur.columns and float(row.get(m, 0) or 0) > 0:
                    primera_mes = i
                    break
            if primera_mes is None:
                continue
            # Subtema / fondo de la primera reunión en apuntes (si existe)
            fondo = "—"
            ap_cn = df_ap[(df_ap["cn"] == cn) & (df_ap["year_num"] == cur_year)].sort_values("Fecha") \
                    if len(df_ap) else pd.DataFrame()
            if not ap_cn.empty:
                # Primer subtema "limpio" disponible
                for _, ap_row in ap_cn.iterrows():
                    s = _clean_sub(ap_row.get("Subtema"))
                    if s and s != "Follow Up / Catch Up":
                        fondo = s
                        break
                if fondo == "—":
                    # Si todos son Follow Up, usar Follow Up
                    s = _clean_sub(ap_cn.iloc[0].get("Subtema"))
                    if s:
                        fondo = s
            q_num = (primera_mes - 1) // 3 + 1
            nuevos_distribuidores.append({
                "name":   str(row["Cliente"]).strip(),
                "tipo":   row.get("_tipo") or "Dist",
                "fondo":  fondo,
                "q":      f"Q{q_num}",
                "q_num":  q_num,
                "mes":    primera_mes,
            })

    # ── COMPARATIVA Q YoY POR PRIORIDAD ─────────────────────────────────
    # Reuniones del Q completo: año actual del Q vs año anterior, separado HP/MP/LP
    q_comparativa = {}
    df_q_prev_yr = year_dfs.get(last_q_year - 1)
    for prio in ["HP", "MP", "LP"]:
        cur_total = 0
        if "Prioridad" in df_q.columns:
            sub_cur = df_q[df_q["Prioridad"] == prio]
            cur_total = int(sum(
                sub_cur[m].sum() for m in last_q_month_names if m in sub_cur.columns
            ))
        prev_total = 0
        if df_q_prev_yr is not None and not df_q_prev_yr.empty and "Prioridad" in df_q_prev_yr.columns:
            sub_prev = df_q_prev_yr[df_q_prev_yr["Prioridad"] == prio]
            prev_total = int(sum(
                sub_prev[m].sum() for m in last_q_month_names if m in sub_prev.columns
            ))
        pct = round((cur_total - prev_total) / prev_total * 100) if prev_total else 0
        q_comparativa[prio] = {"cur": cur_total, "prev": prev_total, "pct": pct}

    # ── CROSS-SELL RADAR ────────────────────────────────────────────────
    # Clientes con >= 3 reuniones en el año actual pero <= 1 producto distinto
    # (excluyendo Follow Up / Catch Up y Sin tema). Posibles oportunidades de
    # ofrecer otros productos del portafolio.
    cross_sell_radar = []
    if len(df_ap):
        df_ap_cur_yr = df_ap[df_ap["year_num"] == cur_year]
        for cn, grp in df_ap_cur_yr.groupby("cn"):
            reun = len(grp)
            if reun < 3:
                continue
            subs = set()
            for s in grp["Subtema"].dropna():
                cleaned = _clean_sub(s)
                if cleaned and cleaned not in ("Follow Up / Catch Up", "Sin tema"):
                    subs.add(cleaned)
            if len(subs) <= 1:
                cross_sell_radar.append({
                    "name":       name_dict.get(cn, cn.title()),
                    "prio":       prio_dict.get(cn, "?"),
                    "reun":       reun,
                    "products":   sorted(subs),
                    "n_products": len(subs),
                })
        cross_sell_radar.sort(
            key=lambda x: (PRIO_ORD.get(x["prio"], 3), -x["reun"], x["name"])
        )
        cross_sell_radar = cross_sell_radar[:25]

    # ── ÁREAS / PRODUCTOS (subtemas del año actual) ─────────────────────
    year_str = str(cur_year)
    sa_year = sa.get(year_str, {})
    sca_year = sca.get(year_str, {})
    areas_data = []
    for sub_name, count in sorted(sa_year.items(), key=lambda x: -x[1]):
        clientes_list = sca_year.get(sub_name, [])
        areas_data.append({
            "name": sub_name,
            "reun": int(count),
            "cp": len(clientes_list),
            "color": sub_colors.get(sub_name, "#7D93B5"),
            "clientes": [c["name"] for c in clientes_list],
        })

    # ── CRUCE GASTO vs REUNIONES + REUNIONES POR ENTIDAD ────────────────
    # Reuniones del año actual por entidad (normalizada), desde el foto. Incluye
    # reun=0 (un cliente puede estar en el foto sin reuniones este año).
    reun_all: dict = {}  # norm -> {nombre, reun, tipo}
    for _, row in df_cur.iterrows():
        nombre = str(row["Cliente"]).strip()
        n = _norm_entity(nombre)
        if not n:
            continue
        reun = int(sum(float(row.get(m, 0) or 0) for m in MESES if m in df_cur.columns))
        cur = reun_all.get(n)
        if cur:
            cur["reun"] += reun
        else:
            reun_all[n] = {"nombre": nombre, "reun": reun, "tipo": (row.get("_tipo") or "Dist")}

    # Para el cruce saldo vs reuniones: sólo entidades con reuniones > 0.
    reun_por_entidad = {n: v for n, v in reun_all.items() if v["reun"] > 0}

    # Gasto por entidad (hoja 'Gastos', columnas Cliente / Monto).
    gasto_por_entidad: dict = {}  # norm -> {nombre, monto}
    if "Gastos" in xls.sheet_names:
        try:
            dfg = pd.read_excel(xls, sheet_name="Gastos")
            if "Cliente" in dfg.columns and "Monto" in dfg.columns:
                for _, row in dfg.iterrows():
                    nombre = str(row["Cliente"]).strip()
                    if not nombre or nombre.lower() in ("nan", "total"):
                        continue
                    n = _norm_entity(nombre)
                    if not n:
                        continue
                    try:
                        monto = float(row["Monto"])
                    except (ValueError, TypeError):
                        continue
                    if pd.isna(monto):
                        continue
                    g = gasto_por_entidad.get(n)
                    if g:
                        g["monto"] += monto
                    else:
                        gasto_por_entidad[n] = {"nombre": nombre, "monto": monto}
        except Exception as e:
            print(f"[gastos] error parseando hoja Gastos: {e}")

    # Cruce gasto vs reuniones: entidades presentes en gastos Y en el foto.
    # La reunión puede ser 0 (gasto sin reuniones es una señal en sí misma).
    cross_gasto_reun = []
    for n, g in gasto_por_entidad.items():
        info = reun_all.get(n)
        if info is None:
            continue  # sólo distribuidores que también están en el foto
        reun = info["reun"]
        cross_gasto_reun.append({
            "nombre": info["nombre"],
            "gasto": g["monto"],
            "reun": reun,
            "gastoPorReun": (g["monto"] / reun) if reun else None,
        })
    cross_gasto_reun.sort(key=lambda x: -x["gasto"])

    # ── CLIENTES FOTO YTD ───────────────────────────────────────────────
    # Reproduce la hoja "Clientes foto {cur_year}" tal cual la ve el Excel:
    # una fila por contraparte con sus reuniones mes a mes, cortado en el mes
    # en curso. Es la vista cruda que el área usa para revisar cobertura.
    ytd_month = ref_month
    df_prev_foto = year_dfs.get(cur_year - 1)
    prev_ytd_by_cn: dict = {}
    if df_prev_foto is not None and not df_prev_foto.empty:
        for _, row in df_prev_foto.iterrows():
            v = int(sum(
                float(row.get(m, 0) or 0)
                for m in MESES[:ytd_month] if m in df_prev_foto.columns
            ))
            prev_ytd_by_cn[row["_cn"]] = prev_ytd_by_cn.get(row["_cn"], 0) + v

    foto_rows = []
    foto_excluidas = []
    for _, row in df_cur.iterrows():
        cn = row["_cn"]
        meses_vals = [
            int(float(row.get(m, 0) or 0)) if m in df_cur.columns else 0
            for m in MESES
        ]
        prio = prio_dict.get(cn)
        prio = prio if prio in THRESH else "?"
        last = last_contact.get(cn) if len(last_contact) else None
        has_last = last is not None and not pd.isna(last)
        rec = {
            "name": str(row["Cliente"]).strip(),
            "prio": prio,
            "tipo": tipo_dict.get(cn) or "Dist",
            "m": meses_vals[:ytd_month],          # Ene → mes en curso
            "ytd": sum(meses_vals[:ytd_month]),
            "anio": sum(meses_vals),              # año completo cargado en el foto
            "prevYtd": prev_ytd_by_cn.get(cn, 0),
            "last": str(last.date()) if has_last else None,
            "dias": int((today_ts - last).days) if has_last else None,
        }
        # Fuera del listado las contrapartes sin historial de contacto. Se
        # conservan las que sí registran reuniones en el foto aunque no
        # aparezcan en Apuntes (pasa cuando el nombre no calza entre hojas):
        # sacarlas descuadraría los totales contra el resto del dashboard.
        if not has_last and rec["anio"] == 0 and rec["prevYtd"] == 0:
            foto_excluidas.append(rec["name"])
            continue
        foto_rows.append(rec)
    # Orden por defecto: más reuniones YTD primero, luego prioridad y nombre.
    foto_rows.sort(key=lambda r: (-r["ytd"], PRIO_ORD.get(r["prio"], 3), r["name"].lower()))

    # El YTD del año anterior por fila sólo cubre contrapartes que siguen en el
    # foto actual. `prevYtdTotal` es el total real del año anterior (incluye las
    # que salieron del foto) — la diferencia se muestra como nota en la UI.
    foto_ytd = {
        "year": cur_year,
        "prevYear": cur_year - 1,
        "month": ytd_month,
        "monthNames": MESES[:ytd_month],
        "rows": foto_rows,
        "prevYtdTotal": int(sum(prev_ytd_by_cn.values())),
        "prevYtdMatched": int(sum(r["prevYtd"] for r in foto_rows)),
        "excluidas": sorted(foto_excluidas, key=str.lower),
    }

    # ── CONSOLIDADO POR TRIMESTRE Y SEGMENTO ────────────────────────────
    # Reuniones y contrapartes distintas (clientes únicos contactados) por
    # prioridad en cada trimestre del año. Los trimestres que aún no empiezan
    # quedan en None y la UI los muestra como "—". Antes esta tabla estaba
    # escrita a mano con los datos de Q1 2026.
    def _q_meses(qi):
        """Meses del trimestre `qi` ya transcurridos y presentes en el foto."""
        return [MESES[m - 1] for m in range((qi - 1) * 3 + 1, qi * 3 + 1)
                if m <= ref_month and MESES[m - 1] in df_cur.columns]

    def _celda(sub, qi):
        cols = _q_meses(qi)
        if not cols:
            return None
        suma = sub[cols].sum(axis=1)
        return {"reun": int(suma.sum()), "cp": int((suma > 0).sum())}

    df_cons = df_cur.copy()
    df_cons["_prio"] = [
        (prio_dict.get(cn) if prio_dict.get(cn) in THRESH else "?") for cn in df_cons["_cn"]
    ]
    cols_ytd = [m for m in MESES[:ref_month] if m in df_cons.columns]

    def _fila(sub, etiqueta):
        ytd = sub[cols_ytd].sum(axis=1) if cols_ytd else sub.iloc[:, :0].sum(axis=1)
        total_reun = int(ytd.sum())
        return {
            "prio": etiqueta,
            "clientes": int(len(sub)),
            "q": [_celda(sub, qi) for qi in (1, 2, 3, 4)],
            "totalReun": total_reun,
            "totalCp": int((ytd > 0).sum()),
            # Proyección anual: ritmo YTD llevado a 12 meses (antes era Q1 × 4).
            "proy": int(round(total_reun / ref_month * 12)) if ref_month else 0,
        }

    consolidado_q = {
        "year": cur_year,
        "currentQ": last_q,
        "refMonth": ref_month,
        "rows": [_fila(df_cons[df_cons["_prio"] == p], p) for p in ("HP", "MP", "LP")],
        "total": _fila(df_cons, "Total"),
        # Contrapartes sin prioridad en el foto: explican que el Total no sea
        # exactamente HP+MP+LP.
        "sinPrio": int((df_cons["_prio"] == "?").sum()),
    }

    # ── RITMO SEMANAL DEL TRIMESTRE EN CURSO ────────────────────────────
    # Promedio de reuniones por semana en el trimestre, cortado en la última
    # semana completa (lunes a domingo). Sale de Apuntes y no del foto: el foto
    # es mensual y no da granularidad semanal.
    ritmo_semanal = None
    if last_q_months:
        q_ini = date(cur_year, last_q_months[0], 1)
        hoy_d = today_ts.date()
        # Domingo de la última semana cerrada (la semana en curso no cuenta).
        fin_sem = hoy_d - timedelta(days=hoy_d.weekday() + 1)
        # Si el Excel viene con rezago, cortar en la semana de la última reunión
        # registrada: si no, las semanas sin cargar diluirían el promedio.
        ap_cur = df_ap[df_ap["year_num"] == cur_year] if len(df_ap) else df_ap
        if len(ap_cur):
            ult = ap_cur["Fecha"].max().date()
            fin_sem = min(fin_sem, ult + timedelta(days=6 - ult.weekday()))
        if fin_sem >= q_ini and len(ap_cur):
            fechas = ap_cur["Fecha"].dt.date
            n_reun = int(((fechas >= q_ini) & (fechas <= fin_sem)).sum())
            semanas = ((fin_sem - q_ini).days + 1) / 7
            ritmo_semanal = {
                "valor": round(n_reun / semanas, 1) if semanas else 0.0,
                "reuniones": n_reun,
                "semanas": round(semanas, 1),
                "hasta": str(fin_sem),
            }

    # ── CLIENTES ACTIVOS DEL TRIMESTRE ──────────────────────────────────
    # Contrapartes del foto con al menos una reunión en los meses del Q ya
    # transcurridos. Usa el foto, igual que el resto de KPIs del trimestre.
    q_cols_foto = [m for m in last_q_month_names if m in df_cur.columns]
    clientes_activos_q = {
        "activos": int((df_cur[q_cols_foto].sum(axis=1) > 0).sum()) if q_cols_foto else 0,
        "total": int(len(df_cur)),
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
            for m in MESES[:ref_month])
        if df_prev_year is not None and not df_prev_year.empty else 0
    )
    pct_change_ytd = round((ytd - ytd_prev) / ytd_prev * 100) if ytd_prev else 0

    return {
        "data": data_out,
        "clientsData": clients_out,
        "prioData": prio_data,
        "recontactData": recontact_data,
        "monthlyDetailData": monthly_detail,
        "subtemaMoData": sm,
        "subteamAnnual": sa,
        "subtemasColors": sub_colors,
        "subtemasClientsM": scm,
        "subtemasClientsA": sca,
        "topClientesQ": clients_q,
        "activationData": activation_data,
        "activationCounts": activation_counts,
        "areasData": areas_data,
        "nuevosContactos": nuevos_distribuidores,
        "qComparativa": q_comparativa,
        "crossSellRadar": cross_sell_radar,
        "crossGastoReun": cross_gasto_reun,
        "fotoYTD": foto_ytd,
        "consolidadoQ": consolidado_q,
        "recentLabels": recent_labels,
        "recentVals": recent_vals,
        "years": years_avail,
        "currentYear": cur_year,
        "currentQ": {
            "year": last_q_year,
            "q": last_q,
            "label": last_q_label,           # "Q3 2026"
            "label_year_prev": f"Q{last_q} {last_q_year - 1}",
            "months": last_q_months,         # meses transcurridos del Q
            "monthNames": last_q_month_names,
            "reuniones": reuniones_q,
            "reuniones_prev": reuniones_q_prev,
            "pct_change": pct_change_q,
            "parcial": last_q_parcial,       # el Q todavía está en curso
            "sinDatos": last_q_sin_datos,    # el Q no tiene ningún mes cargado
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
            "riesgo_by_prio": kpi,
            "ritmo_semanal": ritmo_semanal,
            "clientes_activos_q": clientes_activos_q,
            "reun_ytd_by_prio": reun_ytd_by_prio,
        },
        "meta": {
            "n_clientes": int(len(df_cur)),
            "fecha_corte": str(today_ts.date()),
            "mes_corte": ref_month,          # último mes con data (1-12)
            "years_disponibles": years_avail,
            "warnings": validation_warnings,
        },
    }
