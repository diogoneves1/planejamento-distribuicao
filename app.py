import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from io import BytesIO
from datetime import date
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Liberacao de Pedidos v2",
    layout="wide",
    initial_sidebar_state="expanded"
)

VERSION = "2.0"

# ──────────────────────────────────────────────────────────────────────────────
# ESTILOS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"], .stApp {
    font-family: 'Inter', system-ui, -apple-system, sans-serif !important;
    font-size: 13px;
    color: #1C2B4A;
    -webkit-font-smoothing: antialiased;
}
.stApp { background-color: #F4F6FA; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background-color: #FFFFFF !important;
    border-right: 1px solid #E8ECF2 !important;
}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] h4 {
    color: #1C2B4A !important;
    font-size: 10.5px !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 0.09em;
    margin: 0 !important;
}
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] .stMarkdown p {
    color: #FFFFF !important;
    font-size: 11.5px !important;
    line-height: 1.55;
}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] label,
section[data-testid="stSidebar"] .stWidgetLabel p {
    color: #1C2B4A !important;
    font-size: 11.5px !important;
    font-weight: 500 !important;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label p,
section[data-testid="stSidebar"] div[role="radiogroup"] label span,
section[data-testid="stSidebar"] div[role="radiogroup"] label {
    color: #1C2B4A !important;
    font-size: 12px !important;
    font-weight: 400 !important;
}
section[data-testid="stSidebar"] input {
    background: #FFFFFF !important;
    color: #1C2B4A !important;
    border: 1px solid #C8D0E4 !important;
    border-radius: 6px !important;
    font-size: 12px !important;
}
section[data-testid="stSidebar"] input:focus {
    border-color: #0054A6 !important;
    box-shadow: 0 0 0 3px rgba(0,84,166,0.12) !important;
}
section[data-testid="stSidebar"] input::placeholder { color: #9CAABE !important; }
section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background: #FFFFFF !important;
    border: 1px solid #C8D0E4 !important;
    border-radius: 6px !important;
    color: #1C2B4A !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] > div:focus-within {
    border-color: #0054A6 !important;
    box-shadow: 0 0 0 3px rgba(0,84,166,0.12) !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] input {
    color: #1C2B4A !important;
    background: transparent !important;
    border: none !important;
}
section[data-testid="stSidebar"] span[data-baseweb="tag"] {
    background-color: #EBF2FF !important;
    border: 1px solid #BDD3F5 !important;
    border-radius: 4px !important;
}
section[data-testid="stSidebar"] span[data-baseweb="tag"] span {
    color: #003D80 !important;
    font-size: 11px !important;
    font-weight: 500 !important;
}
section[data-testid="stSidebar"] [data-baseweb="tag"] svg { fill: #5A8AC8 !important; }
section[data-testid="stSidebar"] hr {
    border-color: #E8ECF2 !important;
    margin: 14px 0 !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] {
    background: #F7F9FC !important;
    border: 1.5px dashed #C8D0E4 !important;
    border-radius: 8px !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] p,
section[data-testid="stSidebar"] [data-testid="stFileUploader"] span,
section[data-testid="stSidebar"] [data-testid="stFileUploader"] small {
    color: #5A6A8A !important;
    font-size: 11.5px !important;
}
section[data-testid="stSidebar"] .stAlert {
    background: #0054A6 !important;
    border: 1px solid #003D80 !important;
    border-radius: 6px !important;
    font-size: 11.5px !important;
    font-weight: 600 !important;
}
section[data-testid="stSidebar"] .stAlert p,
section[data-testid="stSidebar"] .stAlert span,
section[data-testid="stSidebar"] .stAlert div {
    color: #FFFFFF !important;
}
section[data-testid="stSidebar"] .stAlert svg {
    fill: #FFFFFF !important;
}

/* ── Botões ── */
.stButton > button {
    background-color: #0054A6 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    font-size: 12.5px !important;
    padding: 9px 22px !important;
    box-shadow: 0 1px 4px rgba(0,84,166,0.25) !important;
    transition: background-color 0.15s !important;
}
.stButton > button:hover { background-color: #003D80 !important; }
.stDownloadButton > button {
    background-color: #FFFFFF !important;
    color: #0054A6 !important;
    border: 1.5px solid #0054A6 !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    font-size: 12.5px !important;
    padding: 9px 22px !important;
}
.stDownloadButton > button:hover { background-color: #EBF2FF !important; }

/* ── Métricas ── */
[data-testid="stMetric"] {
    background: #FFFFFF !important;
    border-radius: 8px !important;
    padding: 16px 18px !important;
    border: 1px solid #E8ECF2 !important;
    border-top: 3px solid #0054A6 !important;
    box-shadow: 0 1px 4px rgba(28,43,74,0.06) !important;
}
[data-testid="stMetric"] label,
[data-testid="stMetric"] [data-testid="stMetricLabel"] p {
    color: #5A6A8A !important;
    font-size: 10px !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.08em !important;
}
[data-testid="stMetric"] [data-testid="stMetricValue"] {
    color: #1C2B4A !important;
    font-size: 1.65rem !important;
    font-weight: 700 !important;
    line-height: 1.15 !important;
}

/* ── Abas ── */
.stTabs [data-baseweb="tab-list"] {
    background: #FFFFFF !important;
    border-bottom: 2px solid #E8ECF2 !important;
    padding: 0 !important;
    gap: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border-radius: 0 !important;
    font-size: 12.5px !important;
    font-weight: 500 !important;
    color: #5A6A8A !important;
    padding: 11px 18px !important;
    border-bottom: 2px solid transparent !important;
    margin-bottom: -2px !important;
}
.stTabs [aria-selected="true"] {
    color: #0054A6 !important;
    font-weight: 600 !important;
    border-bottom: 2px solid #0054A6 !important;
}
.stTabs [data-baseweb="tab"]:hover { color: #0054A6 !important; background: #F4F8FF !important; }

/* ── Tabelas / Expander / Hr ── */
[data-testid="stDataFrame"] {
    border-radius: 8px !important;
    border: 1px solid #E8ECF2 !important;
    overflow: hidden !important;
    box-shadow: 0 1px 4px rgba(28,43,74,0.05) !important;
}
[data-testid="stExpander"] {
    background: #FFFFFF !important;
    border: 1px solid #E8ECF2 !important;
    border-radius: 8px !important;
}
hr { border-color: #E8ECF2 !important; margin: 16px 0 !important; }
.stAlert { border-radius: 6px !important; font-size: 12.5px !important; }
.stCaption, [data-testid="stCaptionContainer"] p {
    color: #5A6A8A !important;
    font-size: 11px !important;
}

/* ── Cards de gráfico ── */
.g-card {
    background: #FFFFFF;
    border: 1px solid #E8ECF2;
    border-radius: 8px;
    padding: 16px 18px 10px 18px;
    margin-bottom: 12px;
    box-shadow: 0 1px 4px rgba(28,43,74,0.05);
}
.g-label {
    font-size: 10px;
    font-weight: 700;
    color: #5A6A8A;
    text-transform: uppercase;
    letter-spacing: 0.09em;
    margin-bottom: 10px;
    padding-bottom: 8px;
    border-bottom: 1px solid #F0F3F8;
}

/* ── Topbar ── */
.topbar {
    background: #0054A6;
    padding: 0 28px;
    height: 52px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin: -1rem -1rem 0 -1rem;
    border-bottom: 1px solid #003D80;
}
.topbar-left { display: flex; align-items: center; gap: 16px; }
.topbar-logo { font-size: 14px; font-weight: 700; color: #FFFFFF; }
.topbar-logo span { font-weight: 300; opacity: 0.7; }
.topbar-divider { width: 1px; height: 20px; background: rgba(255,255,255,0.25); }
.topbar-page { font-size: 13px; font-weight: 500; color: rgba(255,255,255,0.88); }
.topbar-right { font-size: 11px; color: rgba(255,255,255,0.5); }
.topbar-badge {
    background: rgba(255,255,255,0.15);
    color: #FFFFFF;
    font-size: 10px;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 10px;
    letter-spacing: 0.05em;
}

/* ── Breadcrumb ── */
.breadcrumb {
    font-size: 11px; color: #9CAABE;
    margin: 14px 0 6px 0;
    display: flex; align-items: center; gap: 6px;
}
.breadcrumb-sep { color: #C8D0E4; }
.breadcrumb-current { color: #1C2B4A; font-weight: 500; }

/* ── Sidebar brand ── */
.sb-brand {
    display: flex; align-items: center; gap: 10px;
    padding: 16px 0 14px 0;
    border-bottom: 1px solid #E8ECF2;
    margin-bottom: 16px;
}
.sb-brand-icon {
    width: 32px; height: 32px;
    background: #0054A6; border-radius: 6px;
    display: flex; align-items: center; justify-content: center;
    font-size: 13px; color: #FFFFFF; font-weight: 700; flex-shrink: 0;
}
.sb-brand-name { font-size: 12.5px; font-weight: 600; color: #1C2B4A; line-height: 1.2; }
.sb-brand-sub  { font-size: 10px; color: #5A6A8A; margin-top: 1px; }
.sb-section {
    font-size: 10px; font-weight: 600; color: #9CAABE;
    text-transform: uppercase; letter-spacing: 0.11em;
    margin: 16px 0 7px 0;
}

/* ── Painel KPI ── */
.kpi-panel {
    background: #FFFFFF; border: 1px solid #E8ECF2;
    border-radius: 8px; padding: 16px 18px 12px 18px;
    margin-bottom: 16px;
    box-shadow: 0 1px 4px rgba(28,43,74,0.05);
}
.kpi-panel-title {
    font-size: 10px; font-weight: 700; color: #5A6A8A;
    text-transform: uppercase; letter-spacing: 0.09em;
    margin-bottom: 12px; padding-bottom: 8px;
    border-bottom: 1px solid #F0F3F8;
}

/* ── Badge prioridade ── */
.badge-p0 { background:#FFF3CD;color:#856404;border-radius:4px;padding:2px 7px;font-size:10px;font-weight:700; }
.badge-p1 { background:#D1E7DD;color:#0A3622;border-radius:4px;padding:2px 7px;font-size:10px;font-weight:700; }
.badge-p2 { background:#CFE2FF;color:#084298;border-radius:4px;padding:2px 7px;font-size:10px;font-weight:700; }
.badge-p3 { background:#F0F3F8;color:#5A6A8A;border-radius:4px;padding:2px 7px;font-size:10px;font-weight:700; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ──────────────────────────────────────────────────────────────────────────────
COLUNAS_OBRIGATORIAS = [
    "PEDIDO", "CLIENTE", "ESTADO", "REGIÃO",
    "FATURAR EM", "ITEM", "QNTD PROGRAMADA", "ESTOQUE INICIAL"
]

C_BLUE  = "#0054A6"
C_LGRAY = "#E8ECF2"
C_RED   = "#C0392B"
C_AMBER = "#B85C00"
C_GREEN = "#217A3C"
C_TEXT  = "#1C2B4A"
C_MUTED = "#5A6A8A"

def br(v: float, decimais: int = 2) -> str:
    """Formata número no padrão brasileiro: 1.234,56"""
    fmt = f"{v:,.{decimais}f}"
    return fmt.replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_peso(kg: float) -> str:
    """Exibe em kg se < 1000, caso contrário converte para ton."""
    if kg < 1000:
        return br(kg, 0) + " kg"
    else:
        ton = kg / 1000
        return f"{ton:,.2f} ton".replace(",", ".")


# ──────────────────────────────────────────────────────────────────────────────
# FUNÇÕES DE NEGÓCIO
# ──────────────────────────────────────────────────────────────────────────────

def preparar_base(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip().str.upper()
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()
    df["FATURAR EM"]      = pd.to_datetime(df["FATURAR EM"], errors="coerce")
    df["QNTD PROGRAMADA"] = pd.to_numeric(df["QNTD PROGRAMADA"], errors="coerce").fillna(0)
    df["ESTOQUE INICIAL"] = pd.to_numeric(df["ESTOQUE INICIAL"], errors="coerce").fillna(0)
    if "PESO" in df.columns:
        df["PESO"] = pd.to_numeric(df["PESO"], errors="coerce").fillna(0)
    return df


def validar_colunas(df: pd.DataFrame) -> list:
    return [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]


def montar_estoque(df: pd.DataFrame) -> dict:
    estoque = {}
    for _, row in df.iterrows():
        if row["ITEM"] not in estoque:
            estoque[row["ITEM"]] = row["ESTOQUE INICIAL"]
    return estoque


def aplicar_filtros(df, regioes, estados, clientes):
    if regioes:  df = df[df["REGIÃO"].isin(regioes)]
    if estados:  df = df[df["ESTADO"].isin(estados)]
    if clientes: df = df[df["CLIENTE"].isin(clientes)]
    return df


def ordenar_prioridade(df, p_cliente, p_regiao, p_estado):
    """
    P0 = Clientes prioritários   (máxima prioridade)
    P1 = Regiões prioritárias
    P2 = Estados prioritários
    P3 = Demais
    Dentro de cada prioridade: ordem por data de faturamento → número do pedido
    """
    df = df.copy()

    pedidos_p0 = set(df[df["CLIENTE"].isin(p_cliente)]["PEDIDO"].unique()) if p_cliente else set()

    pedidos_p1 = set(
        df[df["REGIÃO"].isin(p_regiao) & ~df["PEDIDO"].isin(pedidos_p0)]["PEDIDO"].unique()
    ) if p_regiao else set()

    pedidos_p2 = set(
        df[
            df["ESTADO"].isin(p_estado) &
            ~df["PEDIDO"].isin(pedidos_p0) &
            ~df["PEDIDO"].isin(pedidos_p1)
        ]["PEDIDO"].unique()
    ) if p_estado else set()

    df["ORDEM_PRIORIDADE"] = 3
    df.loc[df["PEDIDO"].isin(pedidos_p0), "ORDEM_PRIORIDADE"] = 0
    df.loc[df["PEDIDO"].isin(pedidos_p1), "ORDEM_PRIORIDADE"] = 1
    df.loc[df["PEDIDO"].isin(pedidos_p2), "ORDEM_PRIORIDADE"] = 2

    return df.sort_values(["ORDEM_PRIORIDADE", "FATURAR EM", "PEDIDO"], ascending=True)


def analisar_pedidos(df_ordem, estoque, tem_peso: bool):
    liberados, bloqueados = [], []
    consumo_item, faltas_item = {}, {}

    for pedido, grupo in df_ordem.groupby("PEDIDO", sort=False):
        info   = grupo.iloc[0]
        faltas = []
        pode   = True

        for _, row in grupo.iterrows():
            item  = row["ITEM"]
            qtd   = row["QNTD PROGRAMADA"]
            saldo = estoque.get(item, 0)
            if saldo < qtd:
                pode = False
                faltas.append({"ITEM": item, "QTD": qtd, "SALDO": saldo, "FALTA": qtd - saldo})

        peso_pedido = round(grupo["PESO"].sum(), 4) if tem_peso else 0

        base = {
            "PEDIDO":      pedido,
            "CLIENTE":     info["CLIENTE"],
            "ESTADO":      info["ESTADO"],
            "REGIÃO":      info["REGIÃO"],
            "FATURAR EM":  info["FATURAR EM"].date() if pd.notnull(info["FATURAR EM"]) else None,
            "PRIORIDADE":  info["ORDEM_PRIORIDADE"],
            "PESO_TON":    peso_pedido,
        }

        if pode:
            itens = []
            for _, row in grupo.iterrows():
                item, qtd = row["ITEM"], row["QNTD PROGRAMADA"]
                estoque[item] = estoque.get(item, 0) - qtd
                consumo_item[item] = consumo_item.get(item, 0) + qtd
                itens.append(f"{item} ({int(qtd)})")
            liberados.append({**base, "ITENS": " | ".join(itens)})
        else:
            for i, f in enumerate(faltas, 1):
                base[f"ITEM_{i}"]               = f["ITEM"]
                base[f"QTD_NECESSARIA_{i}"]     = f["QTD"]
                base[f"ESTOQUE_DISPONIVEL_{i}"] = f["SALDO"]
                base[f"FALTA_{i}"]              = f["FALTA"]
                faltas_item[f["ITEM"]] = faltas_item.get(f["ITEM"], 0) + f["FALTA"]
            bloqueados.append(base)

    _ec = pd.DataFrame(columns=["ITEM", "CONSUMO_TOTAL"])
    _ef = pd.DataFrame(columns=["ITEM", "FALTA_TOTAL"])

    df_cons = pd.DataFrame([{"ITEM": k, "CONSUMO_TOTAL": v} for k, v in consumo_item.items()]
                           ).sort_values("CONSUMO_TOTAL", ascending=False) if consumo_item else _ec
    df_falt = pd.DataFrame([{"ITEM": k, "FALTA_TOTAL": v}   for k, v in faltas_item.items()]
                           ).sort_values("FALTA_TOTAL",   ascending=False) if faltas_item else _ef
    df_est  = pd.DataFrame([{"ITEM": k, "ESTOQUE_FINAL": v}  for k, v in estoque.items()]
                           ).sort_values("ESTOQUE_FINAL",  ascending=False)

    return pd.DataFrame(liberados), pd.DataFrame(bloqueados), df_cons, df_falt, df_est


def calcular_resumo(df_lib, df_bloq, df_cons, df_falt):
    total = len(df_lib) + len(df_bloq)
    pct   = round(len(df_lib) / total * 100, 1) if total else 0
    ton_lib  = round(df_lib["PESO_TON"].sum(),  2) if not df_lib.empty  and "PESO_TON" in df_lib.columns  else 0
    ton_bloq = round(df_bloq["PESO_TON"].sum(), 2) if not df_bloq.empty and "PESO_TON" in df_bloq.columns else 0
    return pd.DataFrame([
        {"METRICA": "Total pedidos analisados", "VALOR": total},
        {"METRICA": "Pedidos liberados",         "VALOR": len(df_lib)},
        {"METRICA": "Pedidos bloqueados",         "VALOR": len(df_bloq)},
        {"METRICA": "Taxa de liberacao (%)",      "VALOR": pct},
        {"METRICA": "Toneladas liberadas",        "VALOR": ton_lib},
        {"METRICA": "Toneladas retidas",          "VALOR": ton_bloq},
        {"METRICA": "Total itens consumidos",     "VALOR": int(df_cons["CONSUMO_TOTAL"].sum()) if not df_cons.empty else 0},
        {"METRICA": "Total itens em falta",       "VALOR": int(df_falt["FALTA_TOTAL"].sum())   if not df_falt.empty else 0},
    ])


def resumo_por_estado(df_lib, df_bloq):
    """Agrupa pedidos e toneladas por estado."""
    rows = []
    estados = sorted(set(
        (df_lib["ESTADO"].tolist() if not df_lib.empty else []) +
        (df_bloq["ESTADO"].tolist() if not df_bloq.empty else [])
    ))
    for estado in estados:
        lib_e  = df_lib[df_lib["ESTADO"]  == estado] if not df_lib.empty  else pd.DataFrame()
        bloq_e = df_bloq[df_bloq["ESTADO"] == estado] if not df_bloq.empty else pd.DataFrame()
        rows.append({
            "ESTADO":             estado,
            "PED_LIBERADOS":      len(lib_e),
            "PED_BLOQUEADOS":     len(bloq_e),
            "TON_LIBERADAS":      round(lib_e["PESO_TON"].sum(),  2) if not lib_e.empty  else 0,
            "TON_RETIDAS":        round(bloq_e["PESO_TON"].sum(), 2) if not bloq_e.empty else 0,
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df["TOTAL_PEDIDOS"] = df["PED_LIBERADOS"] + df["PED_BLOQUEADOS"]
        df["TOTAL_TON"]     = df["TON_LIBERADAS"] + df["TON_RETIDAS"]
        df["TX_LIBERACAO"]  = (df["PED_LIBERADOS"] / df["TOTAL_PEDIDOS"] * 100).round(1)
        df = df.sort_values("TOTAL_TON", ascending=False)
    return df


# ──────────────────────────────────────────────────────────────────────────────
# EXCEL — helpers
# ──────────────────────────────────────────────────────────────────────────────

def _estilo_aba(ws, cor):
    for cell in ws[1]:
        cell.fill      = PatternFill("solid", fgColor=cor)
        cell.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = Border(bottom=Side(style="thin", color="FFFFFF"))
    for i, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row)):
        fc = "F4F6FA" if i % 2 == 0 else "FFFFFF"
        for cell in row:
            cell.fill      = PatternFill("solid", fgColor=fc)
            cell.font      = Font(name="Calibri", size=10)
            cell.alignment = Alignment(horizontal="left", vertical="center")
    for col in ws.columns:
        letra = get_column_letter(col[0].column)
        ml = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[letra].width = min(max(ml + 3, 12), 50)
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


def _preencher(ws, r1, c1, r2, c2, fill):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).fill = fill


def _mesclar(ws, r, c1, c2, valor, font, fill=None, align=None):
    # remove merges sobrepostos antes de criar novo
    to_remove = [
        mr for mr in list(ws.merged_cells.ranges)
        if mr.min_row <= r <= mr.max_row and mr.min_col <= c2 and mr.max_col >= c1
    ]
    for mr in to_remove:
        ws.unmerge_cells(str(mr))
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cell = ws.cell(r, c1)
    cell.value     = valor
    cell.font      = font
    cell.alignment = align or Alignment(horizontal="center", vertical="center")
    if fill:
        for c in range(c1, c2 + 1):
            ws.cell(r, c).fill = fill


def _visao_gerencial(wb, df_lib, df_bloq, df_cons, df_falt, df_est, df_estado, hoje):
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.chart.series import DataPoint

    if "VISAO_GERENCIAL" in wb.sheetnames:
        del wb["VISAO_GERENCIAL"]
    ws = wb.create_sheet("VISAO_GERENCIAL", 0)

    total    = len(df_lib) + len(df_bloq)
    pct_lib  = round(len(df_lib) / total * 100, 1) if total else 0
    tot_falt = int(df_falt["FALTA_TOTAL"].sum()) if not df_falt.empty else 0
    ton_lib  = round(df_lib["PESO_TON"].sum(),  2) if not df_lib.empty  and "PESO_TON" in df_lib.columns  else 0
    ton_bloq = round(df_bloq["PESO_TON"].sum(), 2) if not df_bloq.empty and "PESO_TON" in df_bloq.columns else 0

    f_navy   = PatternFill("solid", fgColor="0054A6")
    f_titulo = PatternFill("solid", fgColor="002D6B")
    f_cinza  = PatternFill("solid", fgColor="F4F6FA")
    f_branco = PatternFill("solid", fgColor="FFFFFF")
    f_sep    = PatternFill("solid", fgColor="E8ECF2")
    f_verde  = PatternFill("solid", fgColor="D1E7DD")
    f_verm   = PatternFill("solid", fgColor="F8D7DA")

    fn_titulo = Font(name="Calibri", bold=True, color="FFFFFF", size=13)
    fn_label  = Font(name="Calibri", bold=True, color="5A6A8A", size=8)
    fn_navy   = Font(name="Calibri", bold=True, color="0054A6", size=18)
    fn_verde  = Font(name="Calibri", bold=True, color="217A3C", size=18)
    fn_verm   = Font(name="Calibri", bold=True, color="C0392B", size=18)
    fn_normal = Font(name="Calibri", size=10, color="1C2B4A")
    fn_hdr    = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
    fn_azul_v = Font(name="Calibri", bold=True, color="0054A6", size=10)
    fn_verm_v = Font(name="Calibri", bold=True, color="C0392B", size=10)
    fn_data   = Font(name="Calibri", color="FFFFFF", size=9)

    centro = Alignment(horizontal="center", vertical="center")
    esq    = Alignment(horizontal="left",   vertical="center")
    dir_   = Alignment(horizontal="right",  vertical="center")

    # larguras
    for i, w in enumerate([1,14,2,14,2,14,2,14,2,14,2,14,1], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # alturas
    for ln, h in {1:5, 2:26, 3:5, 4:12, 5:24, 6:12, 7:24, 8:5,
                  9:12, 10:16, 11:16, 12:16, 13:16, 14:16, 15:5,
                  16:12, 17:16, 18:16, 19:16, 20:16, 21:16, 22:5}.items():
        ws.row_dimensions[ln].height = h

    _preencher(ws, 1, 1, 70, 13, f_cinza)

    # cabeçalho
    _preencher(ws, 2, 1, 2, 13, f_titulo)
    _mesclar(ws, 2, 2, 9, f"  VISAO GERENCIAL — LIBERACAO DE PEDIDOS  |  v{VERSION}", fn_titulo, f_titulo,
             Alignment(horizontal="left", vertical="center"))
    c = ws.cell(2, 11)
    c.value = hoje.strftime("%d/%m/%Y")
    c.font  = fn_data
    c.alignment = Alignment(horizontal="right", vertical="center")
    c.fill  = f_titulo
    for col in [12, 13]:
        ws.cell(2, col).fill = f_titulo

    _preencher(ws, 3, 1, 3, 13, f_navy)

    # ── KPI linha 1: pedidos ──
    _preencher(ws, 4, 1, 4, 13, f_branco)
    for col, lbl in [(2,"PEDIDOS ANALISADOS"),(4,"LIBERADOS"),(6,"BLOQUEADOS"),
                     (8,"TAXA DE LIBERACAO"),(10,"ITENS EM FALTA")]:
        _mesclar(ws, 4, col, col+1, lbl, fn_label, f_branco, centro)

    _preencher(ws, 5, 1, 5, 13, f_branco)
    for col, val, fn in [
        (2,  total,           fn_navy),
        (4,  len(df_lib),     fn_verde),
        (6,  len(df_bloq),    fn_verm),
        (8,  f"{pct_lib}%",   fn_navy),
        (10, tot_falt,        fn_verm),
    ]:
        _mesclar(ws, 5, col, col+1, val, fn, f_branco, centro)

    # ── KPI linha 2: toneladas ──
    _preencher(ws, 6, 1, 6, 13, f_branco)
    for col, lbl in [(2,"TON LIBERADAS"),(6,"TON RETIDAS"),(10,"TOTAL TON")]:
        _mesclar(ws, 6, col, col+3, lbl, fn_label, f_branco, centro)

    _preencher(ws, 7, 1, 7, 13, f_branco)
    tot_ton = round(ton_lib + ton_bloq, 2)

    def _br(v):
        """Formata número no padrão brasileiro: 1.234,56"""
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    for col, val, fn in [
        (2,  _br(ton_lib),  fn_verde),
        (6,  _br(ton_bloq), fn_verm),
        (10, _br(tot_ton),  fn_navy),
    ]:
        _mesclar(ws, 7, col, col+3, val, fn, f_branco, centro)

    _preencher(ws, 8, 1, 8, 13, f_sep)

    # ── Top Faltas ──
    _preencher(ws, 9, 2, 9, 6, f_navy)
    _mesclar(ws, 9, 2, 6, "TOP ITENS EM FALTA", fn_hdr, f_navy, centro)
    top_falt = df_falt.head(5).reset_index(drop=True)
    for i in range(5):
        r = 10 + i
        _preencher(ws, r, 2, r, 6, f_branco)
        item = top_falt.iloc[i]["ITEM"]        if i < len(top_falt) else ""
        val  = top_falt.iloc[i]["FALTA_TOTAL"] if i < len(top_falt) else ""
        _mesclar(ws, r, 2, 4, item, fn_normal, f_branco, esq)
        c = ws.cell(r, 5)
        c.value = int(val) if val != "" else ""
        c.font  = fn_verm_v if val != "" else fn_normal
        c.alignment = dir_
        c.fill  = f_branco
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)

    # ── Top Consumo ──
    _preencher(ws, 9, 8, 9, 12, f_navy)
    _mesclar(ws, 9, 8, 12, "TOP ITENS POR CONSUMO", fn_hdr, f_navy, centro)
    top_cons = df_cons.head(5).reset_index(drop=True)
    for i in range(5):
        r = 10 + i
        _preencher(ws, r, 8, r, 12, f_branco)
        item = top_cons.iloc[i]["ITEM"]          if i < len(top_cons) else ""
        val  = top_cons.iloc[i]["CONSUMO_TOTAL"] if i < len(top_cons) else ""
        _mesclar(ws, r, 8, 10, item, fn_normal, f_branco, esq)
        c = ws.cell(r, 11)
        c.value = int(val) if val != "" else ""
        c.font  = fn_azul_v if val != "" else fn_normal
        c.alignment = dir_
        c.fill  = f_branco
        ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=12)

    # ── Separador + Título tabela estados (linha 15 = título, linha 16 = headers) ──
    _preencher(ws, 15, 1, 15, 13, f_navy)
    ws.row_dimensions[15].height = 13
    _mesclar(ws, 15, 2, 12, "TONELADAS E PEDIDOS POR ESTADO", fn_hdr, f_navy, centro)

    # headers na linha 16 — SEM mesclar, escrita célula a célula
    _preencher(ws, 16, 1, 16, 13, f_navy)
    ws.row_dimensions[16].height = 13
    headers = [
        (2,  "ESTADO"),
        (4,  "PED LIB"),
        (6,  "PED BLOQ"),
        (7,  "TON LIB"),
        (9,  "TON RETIDAS"),
        (11, "TX LIB%"),
    ]
    for col, lbl in headers:
        c = ws.cell(16, col)
        c.value = lbl
        c.font  = fn_hdr
        c.alignment = centro
        c.fill  = f_navy

    top_est_tab = df_estado.head(6).reset_index(drop=True) if not df_estado.empty else pd.DataFrame()
    for i in range(6):
        r = 17 + i
        ws.row_dimensions[r].height = 15
        fill_linha = f_branco if i % 2 == 0 else PatternFill("solid", fgColor="F4F6FA")
        _preencher(ws, r, 2, r, 12, fill_linha)
        if i < len(top_est_tab):
            row_d = top_est_tab.iloc[i]

            def _br_ton(v):
                return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

            dados = [
                (2,  row_d["ESTADO"],                    fn_normal, esq),
                (4,  int(row_d["PED_LIBERADOS"]),         fn_verde,  centro),
                (6,  int(row_d["PED_BLOQUEADOS"]),        fn_verm,   centro),
                (7,  _br_ton(row_d["TON_LIBERADAS"]),     fn_verde,  dir_),
                (9,  _br_ton(row_d["TON_RETIDAS"]),       fn_verm,   dir_),
                (11, f"{row_d['TX_LIBERACAO']}%",          fn_navy,   centro),
            ]
            for col, val, fn, aln in dados:
                c = ws.cell(r, col)
                c.value = val
                c.font  = fn
                c.alignment = aln
                c.fill  = fill_linha

    _preencher(ws, 23, 1, 23, 13, f_sep)

    # ── Gráficos ──
    ws.row_dimensions[24].height = 12
    _preencher(ws, 24, 2, 24, 12, f_navy)
    _mesclar(ws, 24, 2, 12, "GRAFICOS GERENCIAIS", fn_hdr, f_navy, centro)

    # dados auxiliares ocultos
    ws.cell(60, 1).value = "Liberados"
    ws.cell(61, 1).value = "Bloqueados"
    ws.cell(60, 2).value = len(df_lib)
    ws.cell(61, 2).value = len(df_bloq)

    pizza = PieChart()
    pizza.title  = "Liberados vs Bloqueados"
    pizza.style  = 10
    pizza.width  = 10
    pizza.height = 7
    pizza.add_data(Reference(ws, min_col=2, min_row=60, max_row=61))
    pizza.set_categories(Reference(ws, min_col=1, min_row=60, max_row=61))
    dp0 = DataPoint(idx=0); dp0.graphicalProperties.solidFill = "0054A6"
    dp1 = DataPoint(idx=1); dp1.graphicalProperties.solidFill = "E8ECF2"
    pizza.series[0].dPt = [dp0, dp1]
    ws.add_chart(pizza, "B25")

    # barras toneladas por estado
    if not df_estado.empty:
        est_top = df_estado.head(8)
        ws.cell(60, 4).value = "Estado"
        ws.cell(60, 5).value = "Ton Liberadas"
        ws.cell(60, 6).value = "Ton Retidas"
        for i, row_d in enumerate(est_top.itertuples()):
            ws.cell(61 + i, 4).value = row_d.ESTADO
            ws.cell(61 + i, 5).value = float(row_d.TON_LIBERADAS)
            ws.cell(61 + i, 6).value = float(row_d.TON_RETIDAS)

        n = len(est_top)
        barras = BarChart()
        barras.type     = "col"
        barras.title    = "Toneladas por Estado"
        barras.style    = 10
        barras.width    = 14
        barras.height   = 7
        barras.grouping = "clustered"
        barras.add_data(Reference(ws, min_col=5, min_row=60, max_row=60+n), titles_from_data=True)
        barras.add_data(Reference(ws, min_col=6, min_row=60, max_row=60+n), titles_from_data=True)
        barras.set_categories(Reference(ws, min_col=4, min_row=61, max_row=60+n))
        barras.series[0].graphicalProperties.solidFill = "0054A6"
        barras.series[1].graphicalProperties.solidFill = "F8D7DA"
        ws.add_chart(barras, "G25")

    for r in range(60, 80):
        ws.row_dimensions[r].hidden = True

    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False


def gerar_excel(df_lib, df_bloq, df_cons, df_falt, df_est, df_resumo, df_estado) -> bytes:
    from openpyxl import load_workbook

    output = BytesIO()
    SHEETS = {
        "RESUMO":        (df_resumo,  "0054A6"),
        "LIBERADOS":     (df_lib,     "217A3C"),
        "NAO_LIBERADOS": (df_bloq,    "C0392B"),
        "POR_ESTADO":    (df_estado,  "0054A6"),
        "CONSUMO_ITEM":  (df_cons,    "0054A6"),
        "FALTAS_ITEM":   (df_falt,    "C0392B"),
        "ESTOQUE_FINAL": (df_est,     "217A3C"),
    }
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nome, (df, cor) in SHEETS.items():
            df.to_excel(writer, sheet_name=nome, index=False)
            _estilo_aba(writer.sheets[nome], cor)

    output.seek(0)
    wb = load_workbook(output)
    _visao_gerencial(wb, df_lib, df_bloq, df_cons, df_falt, df_est, df_estado, date.today())

    if "VISAO_GERENCIAL" in wb.sheetnames:
        idx = wb.sheetnames.index("VISAO_GERENCIAL")
        wb.move_sheet("VISAO_GERENCIAL", offset=-idx)

    final = BytesIO()
    wb.save(final)
    return final.getvalue()


# ──────────────────────────────────────────────────────────────────────────────
# GRÁFICOS
# ──────────────────────────────────────────────────────────────────────────────

_L = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, system-ui, sans-serif", size=11, color=C_TEXT),
)


def fig_donut(liberados, bloqueados):
    total = liberados + bloqueados
    pct   = round(liberados / total * 100, 1) if total else 0
    fig   = go.Figure(go.Pie(
        labels=["Liberados", "Bloqueados"],
        values=[liberados, bloqueados],
        hole=0.72,
        marker=dict(colors=[C_BLUE, "#E4EAF5"], line=dict(color="#FFFFFF", width=2)),
        textinfo="none",
        hovertemplate="%{label}: %{value} (%{percent})<extra></extra>",
    ))
    fig.add_annotation(
        text=f"<b style='font-size:22px;color:{C_BLUE}'>{pct}%</b><br>"
             f"<span style='font-size:11px;color:{C_MUTED}'>liberado</span>",
        x=0.5, y=0.5, showarrow=False, align="center"
    )
    fig.update_layout(**_L, margin=dict(t=6,b=6,l=6,r=6), height=210,
                      showlegend=True,
                      legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.06,
                                  font=dict(size=11, color=C_MUTED)))
    return fig


def fig_ton_estado(df_estado):
    df_p = df_estado.head(10).sort_values("TON_LIBERADAS", ascending=True)
    fig  = go.Figure()
    fig.add_trace(go.Bar(
        name="Ton Liberadas", y=df_p["ESTADO"], x=df_p["TON_LIBERADAS"],
        orientation="h", marker_color=C_BLUE,
        text=df_p["TON_LIBERADAS"].apply(lambda v: f"{v:,.1f}"),
        textposition="outside", textfont=dict(size=10, color=C_TEXT),
    ))
    fig.add_trace(go.Bar(
        name="Ton Retidas", y=df_p["ESTADO"], x=df_p["TON_RETIDAS"],
        orientation="h", marker_color="#F8D7DA",
        text=df_p["TON_RETIDAS"].apply(lambda v: f"{v:,.1f}"),
        textposition="outside", textfont=dict(size=10, color=C_RED),
    ))
    fig.update_layout(
        **_L, barmode="group", height=max(200, len(df_p) * 38),
        margin=dict(t=6, b=6, l=6, r=60),
        xaxis=dict(showgrid=True, gridcolor="#F0F3F8", zeroline=False,
                   tickfont=dict(size=10, color=C_MUTED), title="Toneladas"),
        yaxis=dict(showgrid=False, tickfont=dict(size=11, color=C_TEXT)),
        legend=dict(orientation="h", x=1, xanchor="right", y=1.1,
                    font=dict(size=11, color=C_MUTED), bgcolor="rgba(0,0,0,0)"),
    )
    return fig


def fig_ped_estado(df_estado):
    df_p = df_estado.head(10).sort_values("PED_LIBERADOS", ascending=True)
    fig  = go.Figure()
    fig.add_trace(go.Bar(
        name="Liberados", y=df_p["ESTADO"], x=df_p["PED_LIBERADOS"],
        orientation="h", marker_color=C_BLUE,
        text=df_p["PED_LIBERADOS"], textposition="outside",
        textfont=dict(size=10, color=C_TEXT),
    ))
    fig.add_trace(go.Bar(
        name="Bloqueados", y=df_p["ESTADO"], x=df_p["PED_BLOQUEADOS"],
        orientation="h", marker_color="#C5D3E8",
        text=df_p["PED_BLOQUEADOS"], textposition="outside",
        textfont=dict(size=10, color=C_MUTED),
    ))
    fig.update_layout(
        **_L, barmode="group", height=max(200, len(df_p) * 38),
        margin=dict(t=6, b=6, l=6, r=60),
        xaxis=dict(showgrid=True, gridcolor="#F0F3F8", zeroline=False,
                   tickfont=dict(size=10, color=C_MUTED)),
        yaxis=dict(showgrid=False, tickfont=dict(size=11, color=C_TEXT)),
        legend=dict(orientation="h", x=1, xanchor="right", y=1.1,
                    font=dict(size=11, color=C_MUTED), bgcolor="rgba(0,0,0,0)"),
    )
    return fig


def fig_barras_h(df, col_label, col_valor, cor, n=10):
    df_p = df.nlargest(n, col_valor).sort_values(col_valor, ascending=True)
    fig  = go.Figure(go.Bar(
        y=df_p[col_label].astype(str), x=df_p[col_valor],
        orientation="h", marker_color=cor, marker_line=dict(width=0),
        text=df_p[col_valor].apply(lambda v: f"{int(v):,}"),
        textposition="outside", textfont=dict(size=11, color=C_TEXT),
        hovertemplate="%{y}: %{x:,}<extra></extra>",
    ))
    fig.update_layout(
        **_L, height=max(190, len(df_p) * 32),
        margin=dict(t=6, b=6, l=6, r=56),
        xaxis=dict(showgrid=True, gridcolor="#F0F3F8", zeroline=False,
                   tickfont=dict(size=10, color=C_MUTED)),
        yaxis=dict(showgrid=False, tickfont=dict(size=11, color=C_TEXT)),
    )
    return fig


def fig_estoque(df_est, n=12):
    df_p  = df_est.nlargest(n, "ESTOQUE_FINAL").sort_values("ESTOQUE_FINAL", ascending=True)
    cores = [C_RED if v == 0 else C_AMBER if v < 20 else C_BLUE for v in df_p["ESTOQUE_FINAL"]]
    fig   = go.Figure(go.Bar(
        y=df_p["ITEM"].astype(str), x=df_p["ESTOQUE_FINAL"],
        orientation="h", marker_color=cores, marker_line=dict(width=0),
        text=df_p["ESTOQUE_FINAL"].apply(lambda v: f"{int(v):,}"),
        textposition="outside", textfont=dict(size=11, color=C_TEXT),
        hovertemplate="%{y}: %{x:,}<extra></extra>",
    ))
    fig.update_layout(
        **_L, height=max(190, len(df_p) * 32),
        margin=dict(t=6, b=6, l=6, r=56),
        xaxis=dict(showgrid=True, gridcolor="#F0F3F8", zeroline=False,
                   tickfont=dict(size=10, color=C_MUTED)),
        yaxis=dict(showgrid=False, tickfont=dict(size=11, color=C_TEXT)),
    )
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────

with st.sidebar:

    st.markdown(f"""
    <div class="sb-brand">
        <div class="sb-brand-icon">LP</div>
        <div>
            <div class="sb-brand-name">Liberacao de Pedidos</div>
                <div class="sb-brand-sub">Versao {VERSION}</div>
    </div>
    """, unsafe_allow_html=True)

    arquivo = st.file_uploader("Planilha base (.xlsx)", type=["xlsx"])
    if arquivo:
        st.success("Planilha carregada.")

    st.markdown('<div class="sb-section">Periodo de Analise</div>', unsafe_allow_html=True)
    tipo_data = st.radio("Tipo", ["Data especifica", "Periodo"], label_visibility="collapsed")
    if tipo_data == "Data especifica":
        data_ref    = st.date_input("Data", value=date.today(), label_visibility="collapsed")
        data_inicio = data_fim = data_ref
    else:
        data_inicio = st.date_input("Inicio", value=date.today())
        data_fim    = st.date_input("Fim",    value=date.today())

    opcoes_regiao = opcoes_estado = opcoes_cliente = []
    df_global = None
    tem_peso  = False

    if arquivo:
        try:
            df_tmp = preparar_base(pd.read_excel(arquivo))
            if not validar_colunas(df_tmp):
                df_global      = df_tmp
                opcoes_regiao  = sorted(df_tmp["REGIÃO"].dropna().unique())
                opcoes_estado  = sorted(df_tmp["ESTADO"].dropna().unique())
                opcoes_cliente = sorted(df_tmp["CLIENTE"].dropna().unique())
                tem_peso       = "PESO" in df_tmp.columns
        except Exception:
            pass

    # ── Prioridades — P0 Cliente (novo), P1 Região, P2 Estado ──
    st.markdown('<div class="sb-section">Prioridade de Liberacao</div>', unsafe_allow_html=True)
    st.caption("P0 = clientes  ·  P1 = regioes  ·  P2 = estados  ·  P3 = demais")

    p_cliente = st.multiselect("Clientes prioritarios (P0)", opcoes_cliente, placeholder="Nenhum")
    p_regiao  = st.multiselect("Regioes prioritarias (P1)",  opcoes_regiao,  placeholder="Nenhuma")
    p_estado  = st.multiselect("Estados prioritarios (P2)",  opcoes_estado,  placeholder="Nenhum")

    st.markdown('<div class="sb-section">Filtros Operacionais</div>', unsafe_allow_html=True)
    f_regiao  = st.multiselect("Regiao",  opcoes_regiao,  placeholder="Todas")
    f_estado  = st.multiselect("Estado",  opcoes_estado,  placeholder="Todos")
    f_cliente = st.multiselect("Cliente", opcoes_cliente, placeholder="Todos")

    st.markdown("---")
    rodar = st.button("Gerar Analise", use_container_width=True)


# ──────────────────────────────────────────────────────────────────────────────
# TOPBAR
# ──────────────────────────────────────────────────────────────────────────────

st.markdown(f"""
<div class="topbar">
    <div class="topbar-left">
        <div class="topbar-logo">Gestao <span>Operacional</span></div>
        <div class="topbar-divider"></div>
        <div class="topbar-page">Planejamento de Liberacao de Pedidos</div>
        <div class="topbar-badge">v{VERSION}</div>
    </div>
    <div class="topbar-right">{date.today().strftime('%d/%m/%Y')}</div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="breadcrumb">
    Estoque <span class="breadcrumb-sep">›</span>
    Simulacao <span class="breadcrumb-sep">›</span>
    <span class="breadcrumb-current">Liberacao de Pedidos</span>
</div>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# VALIDAÇÕES
# ──────────────────────────────────────────────────────────────────────────────

if not arquivo:
    st.info("Carregue a planilha base na barra lateral para iniciar a analise.")
    st.stop()

if df_global is None:
    erros = validar_colunas(preparar_base(pd.read_excel(arquivo)))
    st.error(f"Colunas obrigatorias ausentes: {', '.join(erros)}")
    st.caption("Esperado: " + "  ·  ".join(COLUNAS_OBRIGATORIAS))
    st.stop()

df = df_global

if not tem_peso:
    st.warning("Coluna PESO nao encontrada na planilha. Toneladas serao exibidas como zero.")

with st.expander("Visualizar base carregada", expanded=False):
    st.dataframe(df.head(100), use_container_width=True, hide_index=True)
    peso_info = f"  ·  coluna PESO detectada" if tem_peso else "  ·  sem coluna PESO"
    st.caption(
        f"{len(df):,} linhas  ·  {df['PEDIDO'].nunique():,} pedidos  "
        f"·  {df['ITEM'].nunique():,} itens  ·  {df['CLIENTE'].nunique():,} clientes{peso_info}"
    )

if not rodar:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Linhas na base",  f"{len(df):,}")
    c2.metric("Pedidos unicos",  f"{df['PEDIDO'].nunique():,}")
    c3.metric("Itens unicos",    f"{df['ITEM'].nunique():,}")
    c4.metric("Clientes unicos", f"{df['CLIENTE'].nunique():,}")
    st.stop()


# ──────────────────────────────────────────────────────────────────────────────
# PROCESSAMENTO
# ──────────────────────────────────────────────────────────────────────────────

with st.spinner("Processando analise..."):
    df_f = df[
        (df["FATURAR EM"].dt.date >= data_inicio) &
        (df["FATURAR EM"].dt.date <= data_fim)
    ]
    if df_f.empty:
        st.warning("Nenhum pedido no periodo selecionado.")
        st.stop()

    df_f = aplicar_filtros(df_f, f_regiao, f_estado, f_cliente)
    if df_f.empty:
        st.warning("Nenhum pedido com os filtros aplicados.")
        st.stop()

    estoque   = montar_estoque(df)
    df_ordem  = ordenar_prioridade(df_f, p_cliente, p_regiao, p_estado)
    df_lib, df_bloq, df_cons, df_falt, df_est = analisar_pedidos(df_ordem, estoque, tem_peso)
    df_resumo = calcular_resumo(df_lib, df_bloq, df_cons, df_falt)
    df_estado = resumo_por_estado(df_lib, df_bloq)


# ──────────────────────────────────────────────────────────────────────────────
# KPIs
# ──────────────────────────────────────────────────────────────────────────────

total    = len(df_lib) + len(df_bloq)
pct      = round(len(df_lib) / total * 100, 1) if total else 0
ton_lib  = round(df_lib["PESO_TON"].sum(),  2) if not df_lib.empty  else 0
ton_bloq = round(df_bloq["PESO_TON"].sum(), 2) if not df_bloq.empty else 0

st.markdown('<div class="kpi-panel"><div class="kpi-panel-title">Resultado Geral da Simulacao</div>', unsafe_allow_html=True)

# linha 1 — pedidos
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Analisados",      f"{total:,}")
c2.metric("Liberados",       f"{len(df_lib):,}")
c3.metric("Bloqueados",      f"{len(df_bloq):,}")
c4.metric("Taxa Liberacao",  f"{pct:.1f}%")
c5.metric("SKUs em Falta",   f"{len(df_falt):,}")

st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

# linha 2 — toneladas
if tem_peso:
    t1, t2, t3 = st.columns(3)
    t1.metric("Peso Liberado",   fmt_peso(ton_lib))
    t2.metric("Peso Retido",     fmt_peso(ton_bloq))
    t3.metric("Total Analisado", fmt_peso(ton_lib + ton_bloq))

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# ABAS
# ──────────────────────────────────────────────────────────────────────────────

tab_dash, tab_estado, tab_lib, tab_bloq, tab_falt, tab_cons, tab_est = st.tabs([
    "Visao Gerencial",
    "Por Estado",
    "Liberados",
    "Nao Liberados",
    "Faltas por Item",
    "Consumo por Item",
    "Estoque Final",
])


# ── VISÃO GERENCIAL ──────────────────────────────────────────────────────────
with tab_dash:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    col_a, col_b = st.columns([1, 2])
    with col_a:
        st.markdown('<div class="g-card"><div class="g-label">Taxa de Liberacao</div>', unsafe_allow_html=True)
        st.plotly_chart(fig_donut(len(df_lib), len(df_bloq)),
                        use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    with col_b:
        st.markdown('<div class="g-card"><div class="g-label">Pedidos por Estado — Liberados vs Bloqueados</div>', unsafe_allow_html=True)
        if not df_estado.empty:
            st.plotly_chart(fig_ped_estado(df_estado),
                            use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    if tem_peso:
        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
        st.markdown('<div class="g-card"><div class="g-label">Toneladas por Estado — Liberadas vs Retidas</div>', unsafe_allow_html=True)
        if not df_estado.empty:
            st.plotly_chart(fig_ton_estado(df_estado),
                            use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)



# ── POR ESTADO ───────────────────────────────────────────────────────────────
with tab_estado:
    if df_estado.empty:
        st.info("Sem dados por estado.")
    else:
        st.caption(f"{len(df_estado):,} estados na analise")
        st.dataframe(df_estado, use_container_width=True, hide_index=True)


# ── LIBERADOS ────────────────────────────────────────────────────────────────
with tab_lib:
    if df_lib.empty:
        st.warning("Nenhum pedido foi liberado.")
    else:
        st.caption(f"{len(df_lib):,} pedidos liberados  ·  {str(round(ton_lib,2)).replace('.',',')} ton")
        # badge de prioridade
        st.dataframe(df_lib, use_container_width=True, hide_index=True)


# ── NAO LIBERADOS ────────────────────────────────────────────────────────────
with tab_bloq:
    if df_bloq.empty:
        st.success("Todos os pedidos foram liberados.")
    else:
        st.caption(f"{len(df_bloq):,} pedidos bloqueados  ·  {str(round(ton_bloq,2)).replace('.',',')} ton retidas")
        st.dataframe(df_bloq, use_container_width=True, hide_index=True)


# ── FALTAS ───────────────────────────────────────────────────────────────────
with tab_falt:
    if df_falt.empty:
        st.success("Nenhum item em falta.")
    else:
        st.caption(f"{len(df_falt):,} itens com falta registrada")
        st.dataframe(df_falt, use_container_width=True, hide_index=True)


# ── CONSUMO ──────────────────────────────────────────────────────────────────
with tab_cons:
    if df_cons.empty:
        st.info("Sem consumo registrado.")
    else:
        st.caption(f"{len(df_cons):,} itens consumidos")
        st.dataframe(df_cons, use_container_width=True, hide_index=True)


# ── ESTOQUE FINAL ─────────────────────────────────────────────────────────────
with tab_est:
    st.caption("Saldo simulado por item apos liberacao dos pedidos possiveis")
    st.dataframe(df_est, use_container_width=True, hide_index=True)


# ──────────────────────────────────────────────────────────────────────────────
# EXPORTAR
# ──────────────────────────────────────────────────────────────────────────────

st.markdown("---")
col_dl, col_txt = st.columns([1, 3])
with col_dl:
    excel_bytes = gerar_excel(df_lib, df_bloq, df_cons, df_falt, df_est, df_resumo, df_estado)
    st.download_button(
        label="Baixar Relatorio Excel",
        data=excel_bytes,
        file_name=f"liberacao_v{VERSION}_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col_txt:
    st.caption(
        f"8 abas: Visao Gerencial · Resumo · Liberados · Nao Liberados · Por Estado · Consumo · Faltas · Estoque Final"
        f"  ·  Gerado em {date.today().strftime('%d/%m/%Y')}"
    )
