import os
import time
import unicodedata
from datetime import date, datetime, time as dt_time, timedelta
from html import escape
from typing import Optional

import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
from plotly.offline import get_plotlyjs
import pyodbc
import streamlit as st
import streamlit.components.v1 as components

# =========================================================
# CONFIG STREAMLIT
# =========================================================
st.set_page_config(
    page_title="One Page • Diretoria — Futurista",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =========================================================
# CONEXÕES — DRIVER ANTIGO {SQL Server}
# =========================================================
CONN_STR_PROTHEUS = (
    "DRIVER={SQL Server};"
    "SERVER=200.201.241.3;"
    "DATABASE=PROTHEUSLOBO;"
    "UID=leitura;"
    "PWD=54321;"
)

CONN_STR_ACESSOTA = (
    "DRIVER={SQL Server};"
    "SERVER=200.201.241.3;"
    "DATABASE=ACESSOTA;"
    "UID=leitura;"
    "PWD=54321;"
)


def connect_sql_safe(conn_str: str, timeout: int = 10, retries: int = 3, delay: int = 2, msg: str = "SQL Server inacessível"):
    last_err = None
    for attempt in range(retries):
        try:
            return pyodbc.connect(conn_str, timeout=timeout, autocommit=True)
        except pyodbc.Error as e:
            last_err = e
            st.warning(f"⚠️ {msg} (tentativa {attempt + 1}/{retries})")
            time.sleep(delay)

    st.error(f"❌ {msg}. Último erro: {last_err}")
    return None


def read_sql_safe(
    query: str,
    conn_str: str,
    params=None,
    retries: int = 3,
    delay: int = 2,
    msg: str = "Falha ao consultar SQL",
):
    last_err = None

    for i in range(retries):
        conn = connect_sql_safe(conn_str, timeout=10, retries=1, delay=1, msg=msg)
        if conn is None:
            return pd.DataFrame()

        try:
            return pd.read_sql_query(query, conn, params=params)
        except (pd.errors.DatabaseError, pyodbc.Error) as e:
            last_err = e
            st.warning(f"⚠️ {msg} (tentativa {i + 1}/{retries})")
            time.sleep(delay)
        finally:
            try:
                conn.close()
            except Exception:
                pass

    st.error(f"❌ {msg}. Último erro: {last_err}")
    return pd.DataFrame()


# =========================================================
# CAMINHOS ABS
# =========================================================
PASTA_PESSOAS = r"Z:\IBERO\OEE\2026 -\Nova pasta\2.0 - PESSOAS"
ARQ_MAO_OBRA = "BASE DE MÃO DE OBRA.xlsx"
ARQ_FERIAS = "FERIAS GERAL.xlsx"

TIPO_AUSENCIA = {
    1: "DOENÇA",
    2: "SOCIAL",
    3: "ESTRUTURA",
}

# =========================================================
# METAS
# =========================================================
META_GGF = 1_000_000.00
META_GGM = 300_000.00
META_GGM_FAT = 0.006
META_FAT = 50_000_000.00

META_ABS = None
META_HE_RS = None
META_HE_HRS = None
META_GGF_ITEM = None
META_FAC_TOTAL = None
META_ENERGIA = None
META_AGUA = None
META_PESSOAS_HE = None

# =========================================================
# REGRAS / FIXOS
# =========================================================
CONTAS_GGF_SD1 = [
    "313030030", "313030031", "313030033", "313030040",
    "313030043", "313030046", "313030047", "313030076",
    "313020103", "313030057", "313030070", "313030058",
    "313030072", "313030059",
]

CONTAS_GGM_SD1 = ["313020129", "313030021"]
CONTAS_FAC_ENERGIA = ["313020020", "313030020"]
CONTAS_FAC_AGUA = ["313020019"]
CONTAS_FAC_ALL = list(dict.fromkeys(CONTAS_FAC_ENERGIA + CONTAS_FAC_AGUA))

CC_PERMITIDOS = [
    "1310", "1400", "1401", "3001", "3403", "3404", "3502",
    "4102", "4103", "4104", "4108", "4109", "4112",
    "4201", "4202", "4203", "4204", "4205",
    "4301", "4402",
    "4501", "4502", "4503",
    "4701", "4702", "4703", "4704", "4708", "4709",
    "4711", "4712", "4716", "4717", "4720", "4721",
    "4993", "4994", "4995",
]

CONSUMIVEL_PARA_CONTA = {
    "CON.000.000.001": "313030043",
    "CON.000.000.002": "313030043",
    "CON.000.000.003": "313030059",
    "CON.000.000.004": "313030076",
    "CON.000.000.005": "313030031",
    "CON.000.000.006": "313030076",
    "CON.000.000.007": "313030076",
    "CON.000.000.008": "313030043",
    "CON.000.000.009": "313030043",
}
CODS_SD3_CONSUMIVEIS = list(CONSUMIVEL_PARA_CONTA.keys())

CF_SUBLIST = ["101", "102", "103", "107", "109", "110", "118", "122", "124", "401", "403", "933"]
VERBAS_HE = ["030", "080", "082", "085", "087", "090", "138", "421", "423"]

# =========================================================
# CORES
# =========================================================
COLORS = {
    "fat": "#12D7A6",
    "fac": "#FF7C7C",
    "he": "#A775FF",
    "abs": "#F0B53A",
    "ggf": "#228CFF",
    "ggm": "#21D39D",
    "facd": "#F2A988",
    "pessoas": "#B58CFF",
}

# =========================================================
# CSS BASE STREAMLIT
# =========================================================
st.markdown(
    """
    <style>
      .block-container{
        max-width: 1800px;
        padding-top: .6rem;
        padding-bottom: .8rem;
      }
      [data-testid="stAppViewContainer"]{
        background:
          radial-gradient(1000px 500px at 8% -8%, rgba(34,140,255,.08), transparent 58%),
          radial-gradient(1000px 520px at 100% 0%, rgba(167,117,255,.07), transparent 55%),
          linear-gradient(180deg, #f7f8fb 0%, #eef2f7 100%);
      }
      [data-testid="stHeader"]{
        background: transparent;
      }
      .hero{
        border-radius: 18px;
        padding: 14px 18px;
        background: rgba(255,255,255,.72);
        border: 1px solid rgba(255,255,255,.9);
        box-shadow: 0 12px 32px rgba(15,23,42,.07);
        backdrop-filter: blur(10px);
        margin-bottom: 14px;
      }
      .hero-title{
        font-size: 26px;
        font-weight: 900;
        color: #111827;
        margin: 0;
      }
      .hero-sub{
        color: #6b7280;
        font-size: 12px;
        margin-top: 4px;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# UTIL
# =========================================================
def yyyymmdd(d: date) -> str:
    return d.strftime("%Y%m%d")


def fmt_brl(v) -> str:
    if v is None or pd.isna(v):
        return "—"
    return "R$ " + f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_brl0(v) -> str:
    if v is None or pd.isna(v):
        return "—"
    return "R$ " + f"{float(v):,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_num(v) -> str:
    if v is None or pd.isna(v):
        return "—"
    return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_num0(v) -> str:
    if v is None or pd.isna(v):
        return "—"
    return f"{float(v):,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(v) -> str:
    if v is None or pd.isna(v):
        return "—"
    return f"{float(v) * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_value(v, kind: str) -> str:
    if kind == "brl":
        return fmt_brl(v)
    if kind == "brl0":
        return fmt_brl0(v)
    if kind == "pct":
        return fmt_pct(v)
    if kind == "num0":
        return fmt_num0(v)
    return fmt_num(v)


def safe_float(x) -> float:
    try:
        if x is None or pd.isna(x):
            return 0.0
        return float(x)
    except Exception:
        return 0.0


def month_range_last_7() -> list[str]:
    end = pd.Timestamp.today().replace(day=1)
    return [(end - pd.DateOffset(months=i)).to_period("M").strftime("%Y-%m") for i in range(6, -1, -1)]


def month_first_day(m: str) -> date:
    return pd.to_datetime(m + "-01").date()


def month_last_day(m: str) -> date:
    return (pd.to_datetime(m + "-01") + pd.offsets.MonthEnd(1)).date()


def normalizar_nome_coluna(col):
    col = str(col).strip()
    col = unicodedata.normalize("NFKD", col).encode("ascii", "ignore").decode("utf-8")
    col = col.upper()
    col = " ".join(col.split())
    return col


def encontrar_coluna(df: pd.DataFrame, candidatos: list[str]) -> str:
    mapa = {normalizar_nome_coluna(c): c for c in df.columns}
    for cand in candidatos:
        chave = normalizar_nome_coluna(cand)
        if chave in mapa:
            return mapa[chave]
    raise RuntimeError(f"Coluna não encontrada. Colunas disponíveis: {list(df.columns)}")


def ajustar_data(dt):
    if pd.isna(dt):
        return None
    if dt.time() <= dt_time(5, 0):
        return dt.date() - timedelta(days=1)
    return dt.date()


def padronizar_matricula(df: pd.DataFrame, col: str) -> pd.DataFrame:
    df[col] = (
        df[col]
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
        .str.zfill(6)
    )
    return df


def dia_util_padrao(d: date) -> bool:
    return d.weekday() < 5


def periodo_abs_fechado(mes_ref: str):
    dt_mes = pd.to_datetime(f"{mes_ref}-01")
    inicio = (dt_mes - pd.DateOffset(months=1)).replace(day=16).date()
    fim = dt_mes.replace(day=15).date()
    return inicio, fim


def periodo_abs_mes_atual_aberto(mes_ref: str):
    hoje = date.today()
    mes_atual = hoje.strftime("%Y-%m")

    if mes_ref != mes_atual:
        return periodo_abs_fechado(mes_ref)

    if hoje.day <= 15:
        base = pd.Timestamp(hoje) - pd.DateOffset(months=1)
        inicio = date(base.year, base.month, 16)
    else:
        inicio = date(hoje.year, hoje.month, 16)

    return inicio, hoje


def month_short_label(m: str) -> str:
    meses = {
        "01": "jan", "02": "fev", "03": "mar", "04": "abr", "05": "mai", "06": "jun",
        "07": "jul", "08": "ago", "09": "set", "10": "out", "11": "nov", "12": "dez",
    }
    ano, mes = m.split("-")
    return f"{meses.get(mes, mes)}/{ano[2:]}"


def hex_to_rgba(hex_color: str, alpha: float) -> str:
    hex_color = hex_color.replace("#", "").strip()
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


def pct_meta_text(value_now, meta, kind):
    if meta is None or safe_float(meta) <= 0 or value_now is None or pd.isna(value_now):
        return "Sem meta"
    pct = safe_float(value_now) / safe_float(meta)
    return f"{pct * 100:,.2f}% da meta".replace(",", "X").replace(".", ",").replace("X", ".")


# =========================================================
# ÍCONES
# =========================================================
def svg_icon_receipt():
    return """
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <path d="M6 3h10l2 2v16l-2-1.2L14 21l-2-1.2L10 21l-2-1.2L6 21V3z"></path>
      <path d="M9 8h6"></path><path d="M9 12h6"></path><path d="M9 16h4"></path>
      <path d="M16 4v4h4"></path>
    </svg>
    """


def svg_icon_building():
    return """
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <path d="M3 21h18"></path>
      <path d="M5 21V7l7-4 7 4v14"></path>
      <path d="M9 21v-4h6v4"></path>
      <path d="M8 10h.01"></path><path d="M12 10h.01"></path><path d="M16 10h.01"></path>
      <path d="M8 13h.01"></path><path d="M12 13h.01"></path><path d="M16 13h.01"></path>
    </svg>
    """


def svg_icon_clock():
    return """
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="12" cy="12" r="8"></circle>
      <path d="M12 7v5l3 2"></path>
      <path d="M4 4l2 2"></path><path d="M20 4l-2 2"></path>
      <path d="M4 20l2-2"></path><path d="M20 20l-2-2"></path>
    </svg>
    """


def svg_icon_abs():
    return """
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <path d="M4 16h16"></path>
      <path d="M4 16V9h5a3 3 0 0 1 3 3v4"></path>
      <path d="M12 12h8v4"></path>
      <circle cx="8" cy="7" r="2"></circle>
      <path d="M18 4l4 4"></path><path d="M22 4l-4 4"></path>
    </svg>
    """


def svg_icon_box():
    return """
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <path d="M12 3l8 4-8 4-8-4 8-4z"></path>
      <path d="M4 7v10l8 4 8-4V7"></path>
      <path d="M12 11v10"></path>
    </svg>
    """


def svg_icon_tools():
    return """
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <path d="M14 7a4 4 0 0 0 5 5l-8 8a2 2 0 1 1-3-3l8-8a4 4 0 0 0-2-2z"></path>
      <path d="M6 6l3 3"></path>
      <path d="M3 9l6-6"></path>
    </svg>
    """


def svg_icon_energy():
    return """
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <path d="M13 2L6 13h5l-1 9 8-12h-5l0-8z"></path>
      <circle cx="12" cy="12" r="10"></circle>
    </svg>
    """


def svg_icon_water():
    return """
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <path d="M4 8h9"></path>
      <path d="M10 8V5a2 2 0 0 1 2-2h2"></path>
      <path d="M13 8v4a3 3 0 0 0 3 3h1"></path>
      <path d="M17 15c0 0-2 2.2-2 4a2 2 0 0 0 4 0c0-1.8-2-4-2-4z"></path>
      <path d="M18 8h2a2 2 0 0 1 2 2v1"></path>
    </svg>
    """


def svg_icon_people():
    return """
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="9" cy="8" r="3"></circle>
      <path d="M4 20a5 5 0 0 1 10 0"></path>
      <circle cx="17" cy="9" r="2.5"></circle>
      <path d="M15 20a4 4 0 0 1 5 0"></path>
    </svg>
    """


# =========================================================
# QUERIES FINANCEIRAS
# =========================================================
@st.cache_data(ttl=240)
def q_sd1_sum_by_month(data_ini_yyyymmdd: str, data_fim_yyyymmdd: str, contas: list[str], ccs: list[str], exigir_pedido: bool):
    if not contas or not ccs:
        return pd.DataFrame(columns=["MES_REF", "VALOR"])

    contas_in = ",".join(["?"] * len(contas))
    cc_in = ",".join(["?"] * len(ccs))

    pedido_clause = ""
    if exigir_pedido:
        pedido_clause = " AND LTRIM(RTRIM(ISNULL(SD1010.D1_PEDIDO,''))) <> ''"

    sql = f"""
        SELECT
            LEFT(SD1010.D1_DTDIGIT, 6) AS MES_YYYYMM,
            SUM(CAST(SD1010.D1_TOTAL AS float)) AS VALOR
        FROM SD1010 WITH (NOLOCK)
        WHERE SD1010.D_E_L_E_T_ = ''
          AND SD1010.D1_DTDIGIT BETWEEN ? AND ?
          AND LTRIM(RTRIM(SD1010.D1_CONTA)) IN ({contas_in})
          AND LTRIM(RTRIM(SD1010.D1_CC)) IN ({cc_in})
          {pedido_clause}
        GROUP BY LEFT(SD1010.D1_DTDIGIT, 6)
        ORDER BY LEFT(SD1010.D1_DTDIGIT, 6)
    """
    params = [data_ini_yyyymmdd, data_fim_yyyymmdd] + list(contas) + list(ccs)

    df = read_sql_safe(sql, CONN_STR_PROTHEUS, params=tuple(params), msg="Falha ao consultar SD1010")

    if df.empty:
        return pd.DataFrame(columns=["MES_REF", "VALOR"])

    df["MES_REF"] = df["MES_YYYYMM"].astype(str).str.slice(0, 4) + "-" + df["MES_YYYYMM"].astype(str).str.slice(4, 6)
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce").fillna(0.0)
    return df[["MES_REF", "VALOR"]]


@st.cache_data(ttl=240)
def q_sd3_ggf_sum_by_month(data_ini_yyyymmdd: str, data_fim_yyyymmdd: str, cods: list[str], ccs: list[str], filial: str, tms: list[str], custo1_unitario: bool):
    if not cods or not ccs:
        return pd.DataFrame(columns=["MES_REF", "VALOR"])

    mes_aberto_yyyymm = date.today().strftime("%Y%m")
    cod_in = ",".join(["?"] * len(cods))
    cc_in = ",".join(["?"] * len(ccs))

    tm_clause = ""
    tm_params = []
    if tms:
        tm_in = ",".join(["?"] * len(tms))
        tm_clause = f" AND LTRIM(RTRIM(SD3010.D3_TM)) IN ({tm_in})"
        tm_params = list(tms)

    filial_clause = ""
    filial_params = []
    if filial:
        filial_clause = " AND LTRIM(RTRIM(SD3010.D3_FILIAL)) = ?"
        filial_params = [filial]

    base_where = f"""
        SD3010.D_E_L_E_T_ = ''
        AND SD3010.D3_EMISSAO BETWEEN ? AND ?
        AND LTRIM(RTRIM(SD3010.D3_COD)) IN ({cod_in})
        {tm_clause}
        {filial_clause}
    """

    valor_expr = "SUM(CAST(SD3010.D3_CUSTO1 AS float))"
    if custo1_unitario:
        valor_expr = "SUM(CAST(SD3010.D3_QUANT AS float) * CAST(SD3010.D3_CUSTO1 AS float))"

    sql = f"""
        SELECT
            LEFT(SD3010.D3_EMISSAO, 6) AS MES_YYYYMM,
            {valor_expr} AS VALOR
        FROM SD3010 WITH (NOLOCK)
        WHERE ({base_where})
          AND (
            (LEFT(SD3010.D3_EMISSAO, 6) <> ? AND LTRIM(RTRIM(ISNULL(SD3010.D3_CC,''))) <> '' AND LTRIM(RTRIM(SD3010.D3_CC)) IN ({cc_in}))
            OR
            (LEFT(SD3010.D3_EMISSAO, 6) = ? AND (
                (LTRIM(RTRIM(ISNULL(SD3010.D3_CC,''))) <> '' AND LTRIM(RTRIM(SD3010.D3_CC)) IN ({cc_in}))
                OR
                (LTRIM(RTRIM(ISNULL(SD3010.D3_CC,''))) = '' AND LTRIM(RTRIM(ISNULL(SD3010.D3_LOCAL,''))) = 'W1')
            ))
          )
        GROUP BY LEFT(SD3010.D3_EMISSAO, 6)
        ORDER BY LEFT(SD3010.D3_EMISSAO, 6)
    """

    params = (
        [data_ini_yyyymmdd, data_fim_yyyymmdd]
        + list(cods)
        + tm_params
        + filial_params
        + [mes_aberto_yyyymm]
        + list(ccs)
        + [mes_aberto_yyyymm]
        + list(ccs)
    )

    df = read_sql_safe(sql, CONN_STR_PROTHEUS, params=tuple(params), msg="Falha ao consultar SD3010")

    if df.empty:
        return pd.DataFrame(columns=["MES_REF", "VALOR"])

    df["MES_REF"] = df["MES_YYYYMM"].astype(str).str.slice(0, 4) + "-" + df["MES_YYYYMM"].astype(str).str.slice(4, 6)
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce").fillna(0.0)
    return df[["MES_REF", "VALOR"]]


@st.cache_data(ttl=240)
def q_faturamento_sum_by_month(data_ini_yyyymmdd: str, data_fim_yyyymmdd: str, filial: str):
    filial = (filial or "01").strip() or "01"
    cf_in = ",".join(["?"] * len(CF_SUBLIST))

    sql = f"""
        SELECT
            LEFT(SD2.D2_EMISSAO, 6) AS MES_YYYYMM,
            SUM(CAST(SD2.D2_TOTAL AS float) + CAST(SD2.D2_VALIPI AS float)) AS VALMERC
        FROM SD2010 SD2 WITH (NOLOCK)
        WHERE SD2.D_E_L_E_T_ = ''
          AND SD2.D2_FILIAL = ?
          AND SD2.D2_EMISSAO BETWEEN ? AND ?
          AND SUBSTRING(SD2.D2_CF, 2, 3) IN ({cf_in})
        GROUP BY LEFT(SD2.D2_EMISSAO, 6)
        ORDER BY LEFT(SD2.D2_EMISSAO, 6)
    """
    params = [filial, data_ini_yyyymmdd, data_fim_yyyymmdd] + list(CF_SUBLIST)

    df = read_sql_safe(sql, CONN_STR_PROTHEUS, params=tuple(params), msg="Falha ao consultar SD2010")

    if df.empty:
        return pd.DataFrame(columns=["MES_REF", "VALMERC"])

    df["MES_REF"] = df["MES_YYYYMM"].astype(str).str.slice(0, 4) + "-" + df["MES_YYYYMM"].astype(str).str.slice(4, 6)
    df["VALMERC"] = pd.to_numeric(df["VALMERC"], errors="coerce").fillna(0.0)
    return df[["MES_REF", "VALMERC"]]


@st.cache_data(ttl=240)
def q_he_horas_sum_by_month(data_ini_yyyymmdd: str, data_fim_yyyymmdd: str, verbas: list[str]):
    verbas = [(v or "").strip() for v in (verbas or []) if str(v).strip() != ""]
    verbas_3 = [str(v).zfill(3) for v in verbas]
    if not verbas_3:
        return pd.DataFrame(columns=["MES_REF", "HE_HORAS"])

    verba_norm_expr = "RIGHT('000' + LTRIM(RTRIM(CONVERT(varchar(20), SP9.P9_CODFOL))), 3)"
    verba_in = ",".join(["?"] * len(verbas_3))

    sql = f"""
        SELECT
            LEFT(CONVERT(varchar(8), SPC.PC_DATA), 6) AS MES_YYYYMM,
            SUM(CAST(SPC.PC_QUANTC AS float)) AS HE_HORAS
        FROM SPC010 SPC WITH (NOLOCK)
        INNER JOIN SP9010 SP9 WITH (NOLOCK)
            ON SP9.D_E_L_E_T_ = ''
           AND CAST(SPC.PC_PD AS bigint) = CAST(SP9.P9_CODIGO AS bigint)
        WHERE SPC.D_E_L_E_T_ = ''
          AND CONVERT(varchar(8), SPC.PC_DATA) BETWEEN ? AND ?
          AND {verba_norm_expr} IN ({verba_in})
        GROUP BY LEFT(CONVERT(varchar(8), SPC.PC_DATA), 6)
        ORDER BY LEFT(CONVERT(varchar(8), SPC.PC_DATA), 6)
    """

    params = [data_ini_yyyymmdd, data_fim_yyyymmdd] + verbas_3
    df = read_sql_safe(sql, CONN_STR_PROTHEUS, params=tuple(params), msg="Falha ao consultar hora extra (horas)")

    if df.empty:
        return pd.DataFrame(columns=["MES_REF", "HE_HORAS"])

    df["MES_REF"] = df["MES_YYYYMM"].astype(str).str.slice(0, 4) + "-" + df["MES_YYYYMM"].astype(str).str.slice(4, 6)
    df["HE_HORAS"] = pd.to_numeric(df["HE_HORAS"], errors="coerce").fillna(0.0)
    return df[["MES_REF", "HE_HORAS"]]


@st.cache_data(ttl=14400)
def descobrir_coluna_valor_he():
    sql = "SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('dbo.SPC010')"
    df = read_sql_safe(sql, CONN_STR_PROTHEUS, msg="Falha ao descobrir coluna de valor da HE")
    if df.empty or "name" not in df.columns:
        return None

    cols = df["name"].astype(str).str.upper().tolist()
    for cand in ["PC_VALOR", "PC_VALCAL", "PC_VAL", "PC_VALCALC", "PC_VLR"]:
        if cand.upper() in cols:
            return cand
    return None


@st.cache_data(ttl=240)
def q_he_rs_sum_by_month(data_ini_yyyymmdd: str, data_fim_yyyymmdd: str, verbas: list[str]):
    col_valor = descobrir_coluna_valor_he()
    if not col_valor:
        return pd.DataFrame(columns=["MES_REF", "HE_RS"])

    verbas = [(v or "").strip() for v in (verbas or []) if str(v).strip() != ""]
    verbas_3 = [str(v).zfill(3) for v in verbas]
    if not verbas_3:
        return pd.DataFrame(columns=["MES_REF", "HE_RS"])

    verba_norm_expr = "RIGHT('000' + LTRIM(RTRIM(CONVERT(varchar(20), SP9.P9_CODFOL))), 3)"
    verba_in = ",".join(["?"] * len(verbas_3))

    sql = f"""
        SELECT
            LEFT(CONVERT(varchar(8), SPC.PC_DATA), 6) AS MES_YYYYMM,
            SUM(CAST(SPC.[{col_valor}] AS float)) AS HE_RS
        FROM SPC010 SPC WITH (NOLOCK)
        INNER JOIN SP9010 SP9 WITH (NOLOCK)
            ON SP9.D_E_L_E_T_ = ''
           AND CAST(SPC.PC_PD AS bigint) = CAST(SP9.P9_CODIGO AS bigint)
        WHERE SPC.D_E_L_E_T_ = ''
          AND CONVERT(varchar(8), SPC.PC_DATA) BETWEEN ? AND ?
          AND {verba_norm_expr} IN ({verba_in})
        GROUP BY LEFT(CONVERT(varchar(8), SPC.PC_DATA), 6)
        ORDER BY LEFT(CONVERT(varchar(8), SPC.PC_DATA), 6)
    """

    params = [data_ini_yyyymmdd, data_fim_yyyymmdd] + verbas_3
    df = read_sql_safe(sql, CONN_STR_PROTHEUS, params=tuple(params), msg="Falha ao consultar hora extra (R$)")

    if df.empty:
        return pd.DataFrame(columns=["MES_REF", "HE_RS"])

    df["MES_REF"] = df["MES_YYYYMM"].astype(str).str.slice(0, 4) + "-" + df["MES_YYYYMM"].astype(str).str.slice(4, 6)
    df["HE_RS"] = pd.to_numeric(df["HE_RS"], errors="coerce").fillna(0.0)
    return df[["MES_REF", "HE_RS"]]


@st.cache_data(ttl=240)
def q_itens_prod_sum_by_month(data_ini_yyyymmdd: str, data_fim_yyyymmdd: str):
    return pd.DataFrame(columns=["MES_REF", "ITENS_PROD"])


# =========================================================
# ABSENTEÍSMO
# =========================================================
@st.cache_data(ttl=1800)
def carregar_base_pessoas():
    caminho = os.path.join(PASTA_PESSOAS, ARQ_MAO_OBRA)
    df = pd.read_excel(caminho)
    df.columns = [str(c).strip() for c in df.columns]

    col_matricula = encontrar_coluna(df, ["Matricula", "Matrícula", "MATRICULA", "MATRÍCULA"])
    df = df.rename(columns={col_matricula: "MATRICULA"})
    df = padronizar_matricula(df, "MATRICULA")
    df = df[df["MATRICULA"].astype(str).str.strip() != ""].copy()
    df = df.drop_duplicates(subset=["MATRICULA"]).reset_index(drop=True)
    return df


@st.cache_data(ttl=1800)
def carregar_base_ferias():
    caminho = os.path.join(PASTA_PESSOAS, ARQ_FERIAS)
    df = pd.read_excel(caminho)
    df.columns = [str(c).strip() for c in df.columns]

    col_matricula = encontrar_coluna(df, ["Matricula", "Matrícula", "MATRICULA", "MATRÍCULA"])
    col_ini = encontrar_coluna(df, ["Inic. Ferias", "Inic Ferias", "Início Férias", "Inicio Ferias", "INIC. FERIAS"])
    col_fim = encontrar_coluna(df, ["Fim   Ferias", "Fim Ferias", "Fim Férias", "FIM FERIAS"])

    df = df.rename(columns={
        col_matricula: "MATRICULA",
        col_ini: "Inic. Ferias",
        col_fim: "Fim   Ferias",
    })

    df = padronizar_matricula(df, "MATRICULA")
    df["Inic. Ferias"] = pd.to_datetime(df["Inic. Ferias"], dayfirst=True, errors="coerce")
    df["Fim   Ferias"] = pd.to_datetime(df["Fim   Ferias"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["Inic. Ferias", "Fim   Ferias"]).copy()
    return df


@st.cache_data(ttl=900)
def q_abs_marcacoes(data_ini_str: str, data_fim_str: str):
    data_ini = pd.to_datetime(data_ini_str).date()
    data_fim = pd.to_datetime(data_fim_str).date()

    dt_inicio = datetime.combine(data_ini, dt_time(0, 0)) - timedelta(hours=5)
    dt_fim = datetime.combine(data_fim + timedelta(days=1), dt_time(5, 0))

    sql = """
    SELECT
        P.PES_MATRICULA AS MATRICULA,
        M.MAR_DATAHORA
    FROM ACESSOTA.TELESSVR.GAC_MARCACAO M WITH (NOLOCK)
    JOIN ACESSOTA.TELESSVR.GAC_PESSOA P WITH (NOLOCK)
        ON M.MAR_PESSOA = P.PES_ID
    WHERE
        P.PES_MATRICULA IS NOT NULL
        AND M.MAR_DATAHORA >= ?
        AND M.MAR_DATAHORA < ?
    """

    df = read_sql_safe(sql, CONN_STR_ACESSOTA, params=(dt_inicio, dt_fim), msg="Falha ao consultar marcações")

    if df.empty:
        return pd.DataFrame(columns=["DATA", "MATRICULA", "QTD_MARCACOES"])

    df = padronizar_matricula(df, "MATRICULA")
    df["MAR_DATAHORA"] = pd.to_datetime(df["MAR_DATAHORA"], errors="coerce")
    df = df.dropna(subset=["MAR_DATAHORA"]).copy()
    df["DATA"] = df["MAR_DATAHORA"].apply(ajustar_data)
    df = df[(df["DATA"] >= data_ini) & (df["DATA"] <= data_fim)].copy()

    if df.empty:
        return pd.DataFrame(columns=["DATA", "MATRICULA", "QTD_MARCACOES"])

    return df.groupby(["DATA", "MATRICULA"]).size().reset_index(name="QTD_MARCACOES")


@st.cache_data(ttl=900)
def q_abs_ponto(data_ini_str: str, data_fim_str: str):
    data_ini = pd.to_datetime(data_ini_str).date()
    data_fim = pd.to_datetime(data_fim_str).date()

    dt_inicio = datetime.combine(data_ini, dt_time(0, 0)) - timedelta(hours=5)
    dt_fim = datetime.combine(data_fim + timedelta(days=1), dt_time(5, 0))

    sql = """
    SELECT
        P.PES_MATRICULA AS MATRICULA,
        PT.PON_DATAHORA,
        PT.PON_TIPO
    FROM ACESSOTA.TELESSVR.GAC_PONTO PT WITH (NOLOCK)
    JOIN ACESSOTA.TELESSVR.GAC_PESSOA P WITH (NOLOCK)
        ON PT.PON_EMP_ID = P.PES_ID
    WHERE
        P.PES_MATRICULA IS NOT NULL
        AND PT.PON_DATAHORA >= ?
        AND PT.PON_DATAHORA < ?
    """

    df = read_sql_safe(sql, CONN_STR_ACESSOTA, params=(dt_inicio, dt_fim), msg="Falha ao consultar ponto/ausência")

    if df.empty:
        return pd.DataFrame(columns=["DATA", "MATRICULA", "QTD_AUSENCIA"])

    df = padronizar_matricula(df, "MATRICULA")
    df["PON_DATAHORA"] = pd.to_datetime(df["PON_DATAHORA"], errors="coerce")
    df = df.dropna(subset=["PON_DATAHORA"]).copy()
    df["DATA"] = df["PON_DATAHORA"].apply(ajustar_data)
    df = df[(df["DATA"] >= data_ini) & (df["DATA"] <= data_fim)].copy()

    df["PON_TIPO"] = pd.to_numeric(df["PON_TIPO"], errors="coerce")
    df["TIPO_AUSENCIA"] = df["PON_TIPO"].map(TIPO_AUSENCIA)
    df = df[df["TIPO_AUSENCIA"].notna()].copy()

    if df.empty:
        return pd.DataFrame(columns=["DATA", "MATRICULA", "QTD_AUSENCIA"])

    return df.groupby(["DATA", "MATRICULA"]).size().reset_index(name="QTD_AUSENCIA")


@st.cache_data(ttl=900)
def montar_base_mo_abs(data_ini_str: str, data_fim_str: str):
    data_ini = pd.to_datetime(data_ini_str).date()
    data_fim = pd.to_datetime(data_fim_str).date()

    pessoas = carregar_base_pessoas()
    ferias = carregar_base_ferias()
    marc = q_abs_marcacoes(data_ini_str, data_fim_str)
    aus = q_abs_ponto(data_ini_str, data_fim_str)

    if pessoas.empty:
        return pd.DataFrame(columns=["DATA", "MATRICULA", "STATUS_FINAL", "DIA_UTIL"])

    dias = [d.date() for d in pd.date_range(data_ini, data_fim, freq="D") if dia_util_padrao(d.date())]
    if not dias:
        return pd.DataFrame(columns=["DATA", "MATRICULA", "STATUS_FINAL", "DIA_UTIL"])

    base = pd.MultiIndex.from_product(
        [dias, pessoas["MATRICULA"].astype(str).unique()],
        names=["DATA", "MATRICULA"]
    ).to_frame(index=False)

    if not marc.empty:
        marc2 = marc.copy()
        marc2["MATRICULA"] = marc2["MATRICULA"].astype(str)
        base = base.merge(marc2, on=["DATA", "MATRICULA"], how="left")
    else:
        base["QTD_MARCACOES"] = 0

    if not aus.empty:
        aus2 = aus.copy()
        aus2["MATRICULA"] = aus2["MATRICULA"].astype(str)
        base = base.merge(aus2, on=["DATA", "MATRICULA"], how="left")
    else:
        base["QTD_AUSENCIA"] = 0

    base["QTD_MARCACOES"] = pd.to_numeric(base.get("QTD_MARCACOES", 0), errors="coerce").fillna(0)
    base["QTD_AUSENCIA"] = pd.to_numeric(base.get("QTD_AUSENCIA", 0), errors="coerce").fillna(0)
    base["FERIAS"] = 0

    if not ferias.empty:
        ferias_exp = []
        for _, row in ferias.iterrows():
            mat = str(row["MATRICULA"]).strip()
            ini = row["Inic. Ferias"]
            fim = row["Fim   Ferias"]

            if pd.isna(ini) or pd.isna(fim):
                continue

            ini_d = max(ini.date(), data_ini)
            fim_d = min(fim.date(), data_fim)
            if ini_d > fim_d:
                continue

            for d in pd.date_range(ini_d, fim_d, freq="D"):
                dd = d.date()
                if dia_util_padrao(dd):
                    ferias_exp.append((dd, mat, 1))

        if ferias_exp:
            df_fer = pd.DataFrame(ferias_exp, columns=["DATA", "MATRICULA", "FERIAS"])
            df_fer = df_fer.drop_duplicates(subset=["DATA", "MATRICULA"])
            base = base.drop(columns=["FERIAS"]).merge(df_fer, on=["DATA", "MATRICULA"], how="left")
            base["FERIAS"] = pd.to_numeric(base["FERIAS"], errors="coerce").fillna(0)

    def status_final(row):
        if row.get("FERIAS", 0) > 0:
            return "FÉRIAS"
        if row.get("QTD_MARCACOES", 0) > 0:
            return "PRESENTE"
        if row.get("QTD_AUSENCIA", 0) > 0:
            return "AUSENTE"
        return "AUSÊNCIA NÃO JUSTIFICADA"

    base["STATUS_FINAL"] = base.apply(status_final, axis=1)
    base["DIA_UTIL"] = 1
    return base[["DATA", "MATRICULA", "STATUS_FINAL", "DIA_UTIL"]]


@st.cache_data(ttl=900)
def q_absenteismo_mes(meses_ref: tuple[str, ...]):
    if not meses_ref:
        return pd.DataFrame(columns=["MES_REF", "ABS_PCT", "ABS_AUSENTES", "ABS_BASE_OK"])

    datas_ini = []
    datas_fim = []
    for mes in meses_ref:
        ini, fim = periodo_abs_mes_atual_aberto(mes)
        datas_ini.append(ini)
        datas_fim.append(fim)

    data_ini_global = min(datas_ini)
    data_fim_global = max(datas_fim)

    base_abs = montar_base_mo_abs(
        data_ini_global.strftime("%Y-%m-%d"),
        data_fim_global.strftime("%Y-%m-%d"),
    )

    if base_abs.empty:
        return pd.DataFrame(columns=["MES_REF", "ABS_PCT", "ABS_AUSENTES", "ABS_BASE_OK"])

    base_abs["DATA"] = pd.to_datetime(base_abs["DATA"]).dt.date

    out = []
    for mes in meses_ref:
        ini, fim = periodo_abs_mes_atual_aberto(mes)
        f = base_abs[
            (base_abs["DATA"] >= ini) &
            (base_abs["DATA"] <= fim) &
            (base_abs["DIA_UTIL"] == 1)
        ].copy()

        ausentes = int((f["STATUS_FINAL"] == "AUSENTE").sum())
        base_ok = int(f["STATUS_FINAL"].isin(["PRESENTE", "FÉRIAS"]).sum())
        abs_pct = (ausentes / base_ok) if base_ok > 0 else 0.0

        out.append({
            "MES_REF": mes,
            "ABS_PCT": abs_pct,
            "ABS_AUSENTES": ausentes,
            "ABS_BASE_OK": base_ok,
        })

    return pd.DataFrame(out)


# =========================================================
# RENDER HTML DASHBOARD
# =========================================================
def build_fig_html(series_df: pd.DataFrame, ycol: str, color: str, kind: str, target_value=None, has_data: bool = True):
    d = series_df.copy()
    if "MES_REF" not in d.columns:
        d = pd.DataFrame({"MES_REF": [], ycol: []})

    d = d[["MES_REF", ycol]].copy() if ycol in d.columns else pd.DataFrame({"MES_REF": d.get("MES_REF", []), ycol: []})
    if d.empty:
        d = pd.DataFrame({"MES_REF": [], ycol: []})

    if len(d) == 0:
        d = pd.DataFrame({"MES_REF": [""], ycol: [0]})

    d["MES_LABEL"] = d["MES_REF"].astype(str).apply(lambda x: month_short_label(x) if x and x != "" and len(x) == 7 else x)
    d[ycol] = pd.to_numeric(d[ycol], errors="coerce")

    if not has_data or d[ycol].notna().sum() == 0:
        d[ycol] = 0.0

    custom_vals = [fmt_value(v, kind) for v in d[ycol]]

    fig = go.Figure()
    fill_color = hex_to_rgba(color, 0.14)
    line_color = color if has_data else "#cfd6e3"

    fig.add_trace(
        go.Scatter(
            x=d["MES_LABEL"],
            y=d[ycol].fillna(0.0),
            mode="lines+markers",
            line=dict(color=line_color, width=2.2, shape="spline"),
            marker=dict(size=5, color=line_color),
            fill="tozeroy" if has_data else None,
            fillcolor=fill_color,
            customdata=custom_vals,
            hovertemplate="<b>%{x}</b><br>Valor: %{customdata}<extra></extra>",
            showlegend=False,
        )
    )

    if target_value is not None and safe_float(target_value) > 0:
        fig.add_trace(
            go.Scatter(
                x=d["MES_LABEL"],
                y=[safe_float(target_value)] * len(d),
                mode="lines",
                line=dict(color="rgba(157,169,189,.95)", width=1.1, dash="dot"),
                hovertemplate=f"<b>Meta</b><br>{fmt_value(target_value, kind)}<extra></extra>",
                showlegend=False,
            )
        )

    fig.update_layout(
        height=78,
        margin=dict(l=0, r=0, t=4, b=0),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
        hovermode="x unified",
        xaxis=dict(
            showgrid=False,
            showline=False,
            zeroline=False,
            ticks="",
            tickfont=dict(size=9, color="rgba(107,114,128,.95)"),
        ),
        yaxis=dict(
            visible=False,
            showgrid=False,
            showline=False,
            zeroline=False,
            rangemode="tozero",
        ),
        hoverlabel=dict(bgcolor="white", font_size=11, font_family="Inter, Arial, sans-serif"),
    )

    return pio.to_html(
        fig,
        full_html=False,
        include_plotlyjs=False,
        config={"displayModeBar": False, "responsive": True},
    )


def card_html(
    title: str,
    value_now,
    value_prev,
    series_df: pd.DataFrame,
    ycol: str,
    color: str,
    icon_svg: str,
    area_class: str,
    kind: str = "brl",
    meta=None,
    reverse_good: bool = False,
    note: Optional[str] = None,
    has_data: bool = True,
):
    del reverse_good

    value_txt = fmt_value(value_now, kind) if has_data else "Sem base"
    prev_txt = fmt_value(value_prev, kind) if has_data else "Sem base"
    meta_txt = fmt_value(meta, kind) if meta is not None else ""
    meta_pct_txt = pct_meta_text(value_now, meta, kind)

    badge_bg = hex_to_rgba(color, 0.18)
    badge_text_color = color
    meta_line = f"Meta: {meta_txt}" if meta is not None else ""
    note_html = f"<div class='card-note'>{escape(note)}</div>" if note else ""
    fig_html = build_fig_html(series_df, ycol, color, kind, target_value=meta, has_data=has_data)

    return f"""
    <div class="kpi-card {area_class}">
      <div class="card-topline" style="background:{color};"></div>
      <div class="card-head">
        <div class="card-title">{escape(title)}</div>
        <div class="card-icon" style="color:{color};">{icon_svg}</div>
      </div>
      <div class="card-value">{escape(value_txt)}</div>
      <div class="card-meta-row">
        <span class="pill" style="background:{badge_bg}; color:{badge_text_color}; border-color:{hex_to_rgba(color, .28)};">{escape(meta_pct_txt)}</span>
        <span class="meta-label">{escape(meta_line)}</span>
      </div>
      <div class="pill pill-gray">Mês anterior: {escape(prev_txt)}</div>
      {note_html}
      <div class="plot-wrap">{fig_html}</div>
    </div>
    """


def render_dashboard_html(base: pd.DataFrame):
    m2 = base.iloc[-1]
    m1 = base.iloc[-2] if len(base) >= 2 else base.iloc[-1]

    has_ggf_item = bool(base["GGF_ITEM"].notna().sum())

    cards = []
    cards.append(card_html("FATURAMENTO", m2["VALMERC"], m1["VALMERC"], base, "VALMERC", COLORS["fat"], svg_icon_receipt(), "area-fat", "brl", META_FAT))
    cards.append(card_html("FACILITIES", m2["FAC_TOTAL"], m1["FAC_TOTAL"], base, "FAC_TOTAL", COLORS["fac"], svg_icon_building(), "area-fac", "brl", META_FAC_TOTAL))
    cards.append(card_html("HORA EXTRA (R$)", m2["HE_RS"], m1["HE_RS"], base, "HE_RS", COLORS["he"], svg_icon_clock(), "area-he-rs", "brl", META_HE_RS))
    cards.append(card_html("ABSENTEÍSMO", m2["ABS_PCT"], m1["ABS_PCT"], base, "ABS_PCT", COLORS["abs"], svg_icon_abs(), "area-abs", "pct", META_ABS))
    cards.append(card_html("GGF", m2["GGF"], m1["GGF"], base, "GGF", COLORS["ggf"], svg_icon_box(), "area-ggf", "brl", META_GGF))
    cards.append(card_html("GGM", m2["GGM"], m1["GGM"], base, "GGM", COLORS["ggm"], svg_icon_tools(), "area-ggm", "brl", META_GGM))
    cards.append(card_html("ENERGIA", m2["FAC_ENERGIA"], m1["FAC_ENERGIA"], base, "FAC_ENERGIA", COLORS["facd"], svg_icon_energy(), "area-energia", "brl", META_ENERGIA))
    cards.append(card_html("HORA EXTRA (HRS)", m2["HE_HORAS"], m1["HE_HORAS"], base, "HE_HORAS", COLORS["he"], svg_icon_clock(), "area-he-hrs", "num", META_HE_HRS))
    cards.append(card_html("GGF/ITENS PROD", m2["GGF_ITEM"], m1["GGF_ITEM"], base, "GGF_ITEM", COLORS["ggf"], svg_icon_box(), "area-ggf-item", "brl", META_GGF_ITEM, note=None if has_ggf_item else "Sem base de itens produzidos configurada neste script.", has_data=has_ggf_item))
    cards.append(card_html("GGM/FAT", m2["GGM_PCT"], m1["GGM_PCT"], base, "GGM_PCT", COLORS["ggm"], svg_icon_tools(), "area-ggm-fat", "pct", META_GGM_FAT))
    cards.append(card_html("ÁGUA", m2["FAC_AGUA"], m1["FAC_AGUA"], base, "FAC_AGUA", COLORS["facd"], svg_icon_water(), "area-agua", "brl", META_AGUA))
    cards.append(card_html("PESSOAS (HE)", m2["PESSOAS_HE"], m1["PESSOAS_HE"], base, "PESSOAS_HE", COLORS["pessoas"], svg_icon_people(), "area-pessoas", "num", META_PESSOAS_HE))

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8" />
      <style>
        * {{ box-sizing: border-box; }}
        body {{ margin: 0; font-family: Inter, Segoe UI, Arial, sans-serif; background: transparent; }}
        .dashboard {{
          display: grid;
          grid-template-columns: repeat(5, minmax(0, 1fr));
          grid-template-areas:
            "fat fat fac he abs"
            "ggf ggm energia hehrs ."
            "ggfitem ggmfat agua pessoas .";
          gap: 12px;
          padding: 2px 2px 4px 2px;
        }}
        .area-fat {{ grid-area: fat; }}
        .area-fac {{ grid-area: fac; }}
        .area-he-rs {{ grid-area: he; }}
        .area-abs {{ grid-area: abs; }}
        .area-ggf {{ grid-area: ggf; }}
        .area-ggm {{ grid-area: ggm; }}
        .area-energia {{ grid-area: energia; }}
        .area-he-hrs {{ grid-area: hehrs; }}
        .area-ggf-item {{ grid-area: ggfitem; }}
        .area-ggm-fat {{ grid-area: ggmfat; }}
        .area-agua {{ grid-area: agua; }}
        .area-pessoas {{ grid-area: pessoas; }}
        .kpi-card {{ position: relative; height: 198px; padding: 10px 12px 8px 12px; border-radius: 16px; background: rgba(255,255,255,.84); border: 1px solid rgba(17,24,39,.10); box-shadow: 0 10px 26px rgba(15,23,42,.08); overflow: hidden; }}
        .card-topline {{ position: absolute; top: 0; left: 0; right: 0; height: 3px; }}
        .card-head {{ display: flex; align-items: flex-start; justify-content: space-between; gap: 8px; margin-top: 1px; }}
        .card-title {{ font-size: 12px; font-weight: 900; letter-spacing: .35px; color: #1f2937; text-transform: uppercase; }}
        .card-icon {{ width: 32px; height: 32px; flex: 0 0 32px; opacity: .95; }}
        .card-icon svg {{ width: 100%; height: 100%; }}
        .card-value {{ font-size: 29px; line-height: 1.05; font-weight: 950; color: #111827; margin: 2px 0 6px 0; letter-spacing: -.3px; }}
        .card-meta-row {{ display: flex; align-items: center; justify-content: space-between; gap: 8px; margin-bottom: 6px; }}
        .pill {{ display: inline-flex; align-items: center; border-radius: 999px; padding: 4px 8px; font-size: 10px; font-weight: 900; border: 1px solid transparent; white-space: nowrap; }}
        .pill-gray {{ background: rgba(17,24,39,.06); color: #6b7280; border-color: rgba(17,24,39,.04); margin-bottom: 6px; }}
        .meta-label {{ font-size: 10px; font-weight: 800; color: #7c8698; text-align: right; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 44%; }}
        .card-note {{ font-size: 10px; color: #8a93a6; margin-top: -2px; margin-bottom: 4px; }}
        .plot-wrap {{ width: 100%; height: 82px; }}
        .plot-wrap .plotly-graph-div {{ width: 100% !important; height: 82px !important; }}
      </style>
      <script>{get_plotlyjs()}</script>
    </head>
    <body>
      <div class="dashboard">{''.join(cards)}</div>
    </body>
    </html>
    """


# =========================================================
# PARÂMETROS
# =========================================================
months = month_range_last_7()
mes_ini = months[0]
mes_fim = months[-1]

data_ini = month_first_day(mes_ini)
data_fim = month_last_day(mes_fim)
data_ini_str = yyyymmdd(data_ini)
data_fim_str = yyyymmdd(data_fim)

cc_sel = CC_PERMITIDOS
filial_fat = "01"
filial_sd3 = ""
tms_sel = ["504"]
custo1_eh_unit = False
he_verbas = VERBAS_HE

# =========================================================
# HEADER
# =========================================================
st.markdown(
    f"""
    <div class="hero">
      <div class="hero-title">📊 One Page • Diretoria — Futurista Clean</div>
      <div class="hero-sub">
        Período automático: <b>{data_ini.strftime('%d/%m/%Y')}</b> até <b>{data_fim.strftime('%d/%m/%Y')}</b>
        • Últimos <b>6 meses + atual</b>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# COLETA
# =========================================================
with st.spinner("Carregando indicadores..."):
    ggf_sd1 = q_sd1_sum_by_month(data_ini_str, data_fim_str, CONTAS_GGF_SD1, cc_sel, exigir_pedido=False)
    ggf_sd3 = q_sd3_ggf_sum_by_month(data_ini_str, data_fim_str, CODS_SD3_CONSUMIVEIS, cc_sel, filial_sd3, tms_sel, custo1_eh_unit)

    ggf = (
        pd.concat(
            [ggf_sd1.rename(columns={"VALOR": "VAL"}), ggf_sd3.rename(columns={"VALOR": "VAL"})],
            ignore_index=True,
        )
        if (not ggf_sd1.empty or not ggf_sd3.empty)
        else pd.DataFrame(columns=["MES_REF", "VAL"])
    )
    ggf = ggf.groupby("MES_REF", as_index=False)["VAL"].sum().rename(columns={"VAL": "GGF"})

    ggm = q_sd1_sum_by_month(data_ini_str, data_fim_str, CONTAS_GGM_SD1, cc_sel, exigir_pedido=True).rename(columns={"VALOR": "GGM"})
    fat = q_faturamento_sum_by_month(data_ini_str, data_fim_str, filial=filial_fat)

    fac_total = q_sd1_sum_by_month(data_ini_str, data_fim_str, CONTAS_FAC_ALL, cc_sel, exigir_pedido=False).rename(columns={"VALOR": "FAC_TOTAL"})
    fac_ener = q_sd1_sum_by_month(data_ini_str, data_fim_str, CONTAS_FAC_ENERGIA, cc_sel, exigir_pedido=False).rename(columns={"VALOR": "FAC_ENERGIA"})
    fac_agua = q_sd1_sum_by_month(data_ini_str, data_fim_str, CONTAS_FAC_AGUA, cc_sel, exigir_pedido=False).rename(columns={"VALOR": "FAC_AGUA"})

    he_horas = q_he_horas_sum_by_month(data_ini_str, data_fim_str, verbas=he_verbas)
    he_rs = q_he_rs_sum_by_month(data_ini_str, data_fim_str, verbas=he_verbas)

    he = he_horas.merge(he_rs, on="MES_REF", how="outer")
    he["HE_HORAS"] = pd.to_numeric(he.get("HE_HORAS", 0), errors="coerce").fillna(0.0)
    he["HE_RS"] = pd.to_numeric(he.get("HE_RS", 0), errors="coerce").fillna(0.0)
    he["PESSOAS_HE"] = he["HE_HORAS"].apply(lambda h: safe_float(h) / 220.0 if safe_float(h) > 0 else 0.0)

    itens_prod = q_itens_prod_sum_by_month(data_ini_str, data_fim_str)

    try:
        abs_df = q_absenteismo_mes(tuple(months))
    except Exception as e:
        abs_df = pd.DataFrame(columns=["MES_REF", "ABS_PCT"])
        st.warning(f"ABS não carregado: {e}")

# =========================================================
# BASE UNIFICADA
# =========================================================
base = pd.DataFrame({"MES_REF": months})
base = base.merge(ggf, on="MES_REF", how="left")
base = base.merge(ggm, on="MES_REF", how="left")
base = base.merge(fat, on="MES_REF", how="left")
base = base.merge(fac_total, on="MES_REF", how="left")
base = base.merge(fac_ener, on="MES_REF", how="left")
base = base.merge(fac_agua, on="MES_REF", how="left")
base = base.merge(he[["MES_REF", "HE_HORAS", "HE_RS", "PESSOAS_HE"]], on="MES_REF", how="left")
base = base.merge(itens_prod, on="MES_REF", how="left")

if not abs_df.empty:
    base = base.merge(abs_df[["MES_REF", "ABS_PCT"]], on="MES_REF", how="left")
else:
    base["ABS_PCT"] = 0.0

for c in ["GGF", "GGM", "VALMERC", "FAC_TOTAL", "FAC_ENERGIA", "FAC_AGUA", "HE_HORAS", "HE_RS", "PESSOAS_HE", "ABS_PCT"]:
    if c not in base.columns:
        base[c] = 0.0
    base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0.0)

if "ITENS_PROD" not in base.columns:
    base["ITENS_PROD"] = pd.NA
else:
    base["ITENS_PROD"] = pd.to_numeric(base["ITENS_PROD"], errors="coerce")

base["GGM_PCT"] = base.apply(lambda r: (r["GGM"] / r["VALMERC"]) if safe_float(r["VALMERC"]) > 0 else 0.0, axis=1)
base["GGF_ITEM"] = base.apply(
    lambda r: (r["GGF"] / r["ITENS_PROD"]) if pd.notna(r["ITENS_PROD"]) and safe_float(r["ITENS_PROD"]) > 0 else pd.NA,
    axis=1,
)

# =========================================================
# DASHBOARD HTML
# =========================================================
components.html(render_dashboard_html(base), height=650, scrolling=False)


