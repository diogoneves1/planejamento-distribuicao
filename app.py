import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# =========================
# CONFIGURAÇÃO DA PÁGINA
# =========================
st.set_page_config(
    page_title="Planejamento de Liberação",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================
# ESTILOS VISUAIS
# =========================
st.markdown("""
<style>

/* ── FONTE BASE ── */
html, body, [class*="css"], .stApp {
    font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif !important;
    font-size: 15px;
    color: #1a2340;
}

/* ── FUNDO GERAL ── */
.stApp {
    background-color: #EFF2FA;
}

/* ══════════════════════════════════════
   SIDEBAR — fundo azul escuro
══════════════════════════════════════ */
section[data-testid="stSidebar"] {
    background: linear-gradient(170deg, #0A1648 0%, #182B78 100%) !important;
}

/* Títulos da sidebar */
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] h4 {
    color: #FFFFFF !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
}

/* Parágrafos / descrições na sidebar */
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] .stMarkdown p {
    color: #C5D0EA !important;
    font-size: 0.80rem !important;
    line-height: 1.5;
}

/* Labels dos widgets na sidebar */
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stWidgetLabel,
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {
    color: #E8EDF8 !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
}

/* Radio buttons — texto legível */
section[data-testid="stSidebar"] div[role="radiogroup"] label p,
section[data-testid="stSidebar"] div[role="radiogroup"] label span {
    color: #FFFFFF !important;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
}

/* Multiselect — texto das opções selecionadas (tags) */
section[data-testid="stSidebar"] span[data-baseweb="tag"] span {
    color: #FFFFFF !important;
    font-weight: 600 !important;
}
section[data-testid="stSidebar"] span[data-baseweb="tag"] {
    background-color: rgba(255,255,255,0.18) !important;
    border: 1px solid rgba(255,255,255,0.3) !important;
}

/* Input de texto / data na sidebar */
section[data-testid="stSidebar"] input {
    background: rgba(255,255,255,0.12) !important;
    color: #FFFFFF !important;
    border: 1.5px solid rgba(255,255,255,0.25) !important;
    border-radius: 8px !important;
    font-size: 0.9rem !important;
}
section[data-testid="stSidebar"] input::placeholder {
    color: rgba(255,255,255,0.45) !important;
}

/* Select / dropdown na sidebar */
section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background: rgba(255,255,255,0.10) !important;
    border: 1.5px solid rgba(255,255,255,0.22) !important;
    border-radius: 8px !important;
    color: #FFFFFF !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] input {
    color: #FFFFFF !important;
    background: transparent !important;
    border: none !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] [data-testid="stMarkdownContainer"] p {
    color: #FFFFFF !important;
}

/* Placeholder do select */
section[data-testid="stSidebar"] div[data-baseweb="select"] [data-testid="stWidgetLabel"] + div p {
    color: rgba(255,255,255,0.5) !important;
}

/* Ícone de fechar (x) das tags do multiselect */
section[data-testid="stSidebar"] [data-baseweb="tag"] svg {
    fill: #FFFFFF !important;
}

/* Upload button */
section[data-testid="stSidebar"] [data-testid="stFileUploader"] {
    background: rgba(255,255,255,0.08) !important;
    border-radius: 12px;
    border: 1.5px dashed rgba(255,255,255,0.25) !important;
    padding: 10px;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] p,
section[data-testid="stSidebar"] [data-testid="stFileUploader"] span,
section[data-testid="stSidebar"] [data-testid="stFileUploader"] small {
    color: #C5D0EA !important;
}

/* Divisor hr na sidebar */
section[data-testid="stSidebar"] hr {
    border-color: rgba(255,255,255,0.12) !important;
    margin: 14px 0 !important;
}

/* Success box na sidebar */
section[data-testid="stSidebar"] .stAlert {
    background: rgba(15,122,69,0.25) !important;
    border: 1px solid rgba(15,200,100,0.35) !important;
    border-radius: 10px !important;
    color: #a8f0c8 !important;
}

/* ══════════════════════════════════════
   CONTEÚDO PRINCIPAL
══════════════════════════════════════ */

/* Títulos */
h1, h2, h3 {
    color: #0D1B5E !important;
    font-weight: 800 !important;
}

/* CARDS DE MÉTRICAS */
[data-testid="stMetric"] {
    background: white;
    border-radius: 16px;
    padding: 22px 26px;
    border-top: 4px solid #182B78;
    box-shadow: 0 2px 14px rgba(24,43,120,0.09);
}
[data-testid="stMetric"] label,
[data-testid="stMetric"] [data-testid="stMetricLabel"] p {
    color: #5A6A8A !important;
    font-size: 0.75rem !important;
    font-weight: 700 !important;
    text-transform: uppercase;
    letter-spacing: 0.07em;
}
[data-testid="stMetric"] [data-testid="stMetricValue"] {
    color: #0D1B5E !important;
    font-size: 2.1rem !important;
    font-weight: 800 !important;
    line-height: 1.1;
}

/* BOTÃO PRINCIPAL */
.stButton > button {
    background: linear-gradient(135deg, #182B78, #1e3a9f) !important;
    color: white !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
    border: none !important;
    padding: 10px 24px !important;
    box-shadow: 0 4px 14px rgba(24,43,120,0.28) !important;
    transition: all 0.2s ease !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #0D1B5E, #182B78) !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(24,43,120,0.38) !important;
}

/* BOTÃO DOWNLOAD */
.stDownloadButton > button {
    background: linear-gradient(135deg, #0f7a45, #13a05a) !important;
    color: white !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    border: none !important;
    box-shadow: 0 4px 14px rgba(15,122,69,0.28) !important;
}
.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #0a5e34, #0f7a45) !important;
}

/* ABAS */
.stTabs [data-baseweb="tab-list"] {
    background: white;
    border-radius: 12px;
    padding: 5px;
    box-shadow: 0 1px 8px rgba(24,43,120,0.09);
    gap: 4px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 9px;
    font-weight: 600;
    font-size: 0.88rem;
    color: #5A6A8A;
    padding: 8px 18px;
}
.stTabs [aria-selected="true"] {
    background: #182B78 !important;
    color: white !important;
}

/* TABELAS */
[data-testid="stDataFrame"] {
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 2px 10px rgba(24,43,120,0.08);
}

/* SEPARADORES */
hr {
    border-color: #DDE3F0;
    margin: 22px 0;
}

/* SELECT área principal */
div[data-baseweb="select"] > div {
    border-radius: 8px !important;
    border-color: #C8D0E4 !important;
    color: #1a2340 !important;
}
div[data-baseweb="select"] > div:focus-within {
    border-color: #182B78 !important;
    box-shadow: 0 0 0 2px rgba(24,43,120,0.15) !important;
}

/* ALERTA */
.stAlert {
    border-radius: 12px;
}

/* BANNER HERO */
.hero-banner {
    background: linear-gradient(135deg, #0A1648 0%, #182B78 55%, #1e3a9f 100%);
    padding: 34px 44px;
    border-radius: 20px;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
    box-shadow: 0 8px 32px rgba(24,43,120,0.22);
}
.hero-banner::before {
    content: '';
    position: absolute;
    top: -50px; right: -50px;
    width: 220px; height: 220px;
    background: rgba(255,255,255,0.05);
    border-radius: 50%;
}
.hero-banner::after {
    content: '';
    position: absolute;
    bottom: -70px; right: 90px;
    width: 300px; height: 300px;
    background: rgba(255,255,255,0.03);
    border-radius: 50%;
}
.hero-title {
    font-family: 'Segoe UI', sans-serif;
    font-size: 1.85rem;
    font-weight: 800;
    color: #FFFFFF;
    margin: 0 0 8px 0;
    letter-spacing: -0.3px;
}
.hero-sub {
    color: #9DB8E8;
    font-size: 0.95rem;
    margin: 0;
    font-weight: 400;
}

/* SECTION LABEL */
.section-label {
    font-size: 0.68rem;
    font-weight: 700;
    color: #8BA0C4;
    text-transform: uppercase;
    letter-spacing: 0.13em;
    margin-bottom: 6px;
}

</style>
""", unsafe_allow_html=True)

# =========================
# CONSTANTES
# =========================
COLUNAS_OBRIGATORIAS = [
    "PEDIDO", "CLIENTE", "ESTADO", "REGIÃO",
    "FATURAR EM", "ITEM", "QNTD PROGRAMADA", "ESTOQUE INICIAL"
]

# =========================
# FUNÇÕES DE NEGÓCIO
# =========================

def preparar_base(df: pd.DataFrame) -> pd.DataFrame:
    """Padroniza colunas, tipos e valores da base."""
    df = df.copy()
    df.columns = df.columns.str.strip().str.upper()

    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()

    df["FATURAR EM"] = pd.to_datetime(df["FATURAR EM"], errors="coerce")
    df["QNTD PROGRAMADA"] = pd.to_numeric(df["QNTD PROGRAMADA"], errors="coerce").fillna(0)
    df["ESTOQUE INICIAL"] = pd.to_numeric(df["ESTOQUE INICIAL"], errors="coerce").fillna(0)

    return df


def validar_colunas(df: pd.DataFrame) -> list:
    """Retorna lista de colunas obrigatórias ausentes."""
    return [col for col in COLUNAS_OBRIGATORIAS if col not in df.columns]


def montar_estoque(df: pd.DataFrame) -> dict:
    """
    Constrói dicionário de estoque por item.
    Usa o primeiro valor de ESTOQUE INICIAL encontrado para cada item.
    """
    estoque = {}
    for _, row in df.iterrows():
        item = row["ITEM"]
        if item not in estoque:
            estoque[item] = row["ESTOQUE INICIAL"]
    return estoque


def aplicar_filtros_operacionais(
    df: pd.DataFrame,
    regioes: list,
    estados: list,
    clientes: list
) -> pd.DataFrame:
    """Aplica filtros opcionais de região, estado e cliente."""
    if regioes:
        df = df[df["REGIÃO"].isin(regioes)]
    if estados:
        df = df[df["ESTADO"].isin(estados)]
    if clientes:
        df = df[df["CLIENTE"].isin(clientes)]
    return df


def ordenar_por_prioridade(
    df: pd.DataFrame,
    prioridade_regiao: list,
    prioridade_estado: list
) -> pd.DataFrame:
    """
    Classifica pedidos em 3 níveis de prioridade:
      P1 – regiões prioritárias
      P2 – estados prioritários (excluindo os já P1)
      P3 – demais
    Em seguida ordena por prioridade → data de faturamento → número do pedido.
    """
    df = df.copy()

    pedidos_p1 = set()
    pedidos_p2 = set()

    if prioridade_regiao:
        pedidos_p1 = set(df[df["REGIÃO"].isin(prioridade_regiao)]["PEDIDO"].unique())

    if prioridade_estado:
        pedidos_p2 = set(
            df[df["ESTADO"].isin(prioridade_estado) & ~df["PEDIDO"].isin(pedidos_p1)]["PEDIDO"].unique()
        )

    df["ORDEM_PRIORIDADE"] = 3
    df.loc[df["PEDIDO"].isin(pedidos_p1), "ORDEM_PRIORIDADE"] = 1
    df.loc[df["PEDIDO"].isin(pedidos_p2), "ORDEM_PRIORIDADE"] = 2

    return df.sort_values(
        ["ORDEM_PRIORIDADE", "FATURAR EM", "PEDIDO"],
        ascending=[True, True, True]
    )


def analisar_pedidos(df_ordem: pd.DataFrame, estoque: dict) -> tuple:
    """
    Percorre os pedidos na ordem de prioridade.
    Libera um pedido somente se TODOS os itens tiverem estoque suficiente.
    Retorna 5 DataFrames: liberados, não-liberados, consumo, faltas, estoque final.
    """
    liberados = []
    nao_liberados = []
    consumo_item: dict = {}
    faltas_item: dict = {}

    for pedido, grupo in df_ordem.groupby("PEDIDO", sort=False):
        info = grupo.iloc[0]
        faltas = []
        pode_liberar = True

        for _, row in grupo.iterrows():
            item = row["ITEM"]
            qtd = row["QNTD PROGRAMADA"]
            saldo = estoque.get(item, 0)

            if saldo < qtd:
                pode_liberar = False
                faltas.append({
                    "ITEM": item,
                    "QTD_PEDIDO": qtd,
                    "ESTOQUE_DISPONIVEL": saldo,
                    "FALTA": qtd - saldo
                })

        base_linha = {
            "PEDIDO": pedido,
            "CLIENTE": info["CLIENTE"],
            "ESTADO": info["ESTADO"],
            "REGIÃO": info["REGIÃO"],
            "FATURAR EM": info["FATURAR EM"].date() if pd.notnull(info["FATURAR EM"]) else None,
            "PRIORIDADE": info["ORDEM_PRIORIDADE"],
        }

        if pode_liberar:
            # Registra e consome estoque
            itens_pedido = []
            for _, row in grupo.iterrows():
                item = row["ITEM"]
                qtd = row["QNTD PROGRAMADA"]
                estoque[item] = estoque.get(item, 0) - qtd
                consumo_item[item] = consumo_item.get(item, 0) + qtd
                itens_pedido.append(f"{item} ({int(qtd)})")

            linha = {**base_linha, "ITENS": " | ".join(itens_pedido)}
            liberados.append(linha)

        else:
            # Registra faltas
            for i, falta in enumerate(faltas, start=1):
                base_linha[f"ITEM_{i}"] = falta["ITEM"]
                base_linha[f"QTD_NECESSÁRIA_{i}"] = falta["QTD_PEDIDO"]
                base_linha[f"ESTOQUE_DISPONÍVEL_{i}"] = falta["ESTOQUE_DISPONIVEL"]
                base_linha[f"FALTA_{i}"] = falta["FALTA"]
                faltas_item[falta["ITEM"]] = faltas_item.get(falta["ITEM"], 0) + falta["FALTA"]

            nao_liberados.append(base_linha)

    df_liberados = pd.DataFrame(liberados)
    df_nao_liberados = pd.DataFrame(nao_liberados)

    df_consumo = pd.DataFrame([
        {"ITEM": k, "CONSUMO_TOTAL": v} for k, v in consumo_item.items()
    ]).sort_values("CONSUMO_TOTAL", ascending=False) if consumo_item else pd.DataFrame(columns=["ITEM", "CONSUMO_TOTAL"])

    df_faltas = pd.DataFrame([
        {"ITEM": k, "FALTA_TOTAL": v} for k, v in faltas_item.items()
    ]).sort_values("FALTA_TOTAL", ascending=False) if faltas_item else pd.DataFrame(columns=["ITEM", "FALTA_TOTAL"])

    df_estoque_final = pd.DataFrame([
        {"ITEM": k, "ESTOQUE_FINAL_SIMULADO": v} for k, v in estoque.items()
    ]).sort_values("ESTOQUE_FINAL_SIMULADO", ascending=False)

    return df_liberados, df_nao_liberados, df_consumo, df_faltas, df_estoque_final


def calcular_resumo(df_liberados, df_nao_liberados, df_consumo, df_faltas) -> pd.DataFrame:
    """Monta o DataFrame de resumo gerencial."""
    total = len(df_liberados) + len(df_nao_liberados)
    pct = (len(df_liberados) / total * 100) if total > 0 else 0

    linhas = [
        {"MÉTRICA": "Total de pedidos analisados", "VALOR": total},
        {"MÉTRICA": "Pedidos liberados", "VALOR": len(df_liberados)},
        {"MÉTRICA": "Pedidos bloqueados", "VALOR": len(df_nao_liberados)},
        {"MÉTRICA": "Taxa de liberação (%)", "VALOR": round(pct, 1)},
        {"MÉTRICA": "Total itens consumidos", "VALOR": int(df_consumo["CONSUMO_TOTAL"].sum()) if not df_consumo.empty else 0},
        {"MÉTRICA": "Total itens em falta", "VALOR": int(df_faltas["FALTA_TOTAL"].sum()) if not df_faltas.empty else 0},
    ]
    return pd.DataFrame(linhas)


# =========================
# GERAÇÃO DO EXCEL FORMATADO
# =========================

def _estilo_cabecalho(ws, row_num: int, cor_hex: str = "182B78"):
    """Aplica estilo de cabeçalho em uma linha inteira."""
    fill = PatternFill("solid", fgColor=cor_hex)
    font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        bottom=Side(style="medium", color="FFFFFF")
    )

    for cell in ws[row_num]:
        cell.fill = fill
        cell.font = font
        cell.alignment = align
        cell.border = border


def _ajustar_largura(ws):
    """Ajusta largura de colunas automaticamente."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                val = str(cell.value) if cell.value else ""
                max_len = max(max_len, len(val))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)


def _zebrar_linhas(ws, inicio: int, cor_par: str = "EEF1FA", cor_impar: str = "FFFFFF"):
    """Alterna cor de fundo das linhas de dados."""
    for i, row in enumerate(ws.iter_rows(min_row=inicio, max_row=ws.max_row)):
        fill_color = cor_par if i % 2 == 0 else cor_impar
        for cell in row:
            if cell.fill.fill_type not in ("solid",) or cell.fill.fgColor.rgb in ("182B78", "0f7a45", "9f1313"):
                cell.fill = PatternFill("solid", fgColor=fill_color)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.font = Font(name="Calibri", size=10)


def gerar_excel(
    df_liberados, df_nao_liberados,
    df_consumo, df_faltas,
    df_estoque_final, df_resumo
) -> bytes:
    """Gera relatório Excel profissional com formatação completa."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        # ----- Abas de dados -----
        sheets = {
            "RESUMO": df_resumo,
            "LIBERADOS": df_liberados,
            "NÃO_LIBERADOS": df_nao_liberados,
            "CONSUMO_ITEM": df_consumo,
            "FALTAS_ITEM": df_faltas,
            "ESTOQUE_FINAL": df_estoque_final,
        }

        cores_cabecalho = {
            "RESUMO": "182B78",
            "LIBERADOS": "0f7a45",
            "NÃO_LIBERADOS": "9f1313",
            "CONSUMO_ITEM": "182B78",
            "FALTAS_ITEM": "9f1313",
            "ESTOQUE_FINAL": "0f7a45",
        }

        for nome, df in sheets.items():
            df.to_excel(writer, sheet_name=nome, index=False, startrow=0)
            ws = writer.sheets[nome]

            # Cabeçalho colorido
            _estilo_cabecalho(ws, 1, cores_cabecalho[nome])

            # Linhas alternadas
            if ws.max_row > 1:
                _zebrar_linhas(ws, 2)

            # Ajusta larguras
            _ajustar_largura(ws)

            # Congela cabeçalho
            ws.freeze_panes = "A2"

            # Filtro automático
            if ws.max_row > 1 and ws.max_column > 0:
                ws.auto_filter.ref = ws.dimensions

    return output.getvalue()


# =========================
# SIDEBAR
# =========================

with st.sidebar:
    st.markdown("## 📦 Liberação de Pedidos")
    st.markdown("---")

    arquivo = st.file_uploader(
        "Carregar planilha base",
        type=["xlsx"],
        help="Planilha Excel com pedidos, estoque e datas de faturamento."
    )

    if arquivo:
        st.success("Planilha carregada!")

    st.markdown("---")
    st.markdown("#### ⚙️ Configurações")

    st.markdown('<p class="section-label">Filtro de Data</p>', unsafe_allow_html=True)
    tipo_data = st.radio(
        "Filtrar por",
        ["Data específica", "Período"],
        label_visibility="collapsed"
    )

    if tipo_data == "Data específica":
        data_ref = st.date_input("Data de referência", value=date.today())
        data_inicio = data_fim = data_ref
    else:
        data_inicio = st.date_input("Data inicial", value=date.today())
        data_fim = st.date_input("Data final", value=date.today())

    st.markdown("---")
    st.markdown("#### 🏆 Prioridade de Liberação")
    st.markdown('<p style="color:#9DB4D4;font-size:0.78rem;">Pedidos das regiões/estados escolhidos serão analisados primeiro.</p>', unsafe_allow_html=True)

    # Carrega opções dinâmicas se houver planilha
    opcoes_regiao = []
    opcoes_estado = []
    opcoes_cliente = []
    df_global = None

    if arquivo:
        try:
            df_temp = preparar_base(pd.read_excel(arquivo))
            erros = validar_colunas(df_temp)
            if not erros:
                df_global = df_temp
                opcoes_regiao = sorted(df_temp["REGIÃO"].dropna().unique().tolist())
                opcoes_estado = sorted(df_temp["ESTADO"].dropna().unique().tolist())
                opcoes_cliente = sorted(df_temp["CLIENTE"].dropna().unique().tolist())
        except Exception:
            pass

    prioridade_regiao = st.multiselect(
        "Regiões prioritárias (P1)",
        opcoes_regiao,
        placeholder="Todas as regiões..."
    )

    prioridade_estado = st.multiselect(
        "Estados prioritários (P2)",
        opcoes_estado,
        placeholder="Todos os estados..."
    )

    st.markdown("---")
    st.markdown("#### 🔍 Filtros Operacionais")
    st.markdown('<p style="color:#9DB4D4;font-size:0.78rem;">Restringe a análise a regiões, estados ou clientes específicos.</p>', unsafe_allow_html=True)

    filtro_regiao = st.multiselect("Região", opcoes_regiao, placeholder="Todas...")
    filtro_estado = st.multiselect("Estado", opcoes_estado, placeholder="Todos...")
    filtro_cliente = st.multiselect("Cliente", opcoes_cliente, placeholder="Todos...")

    st.markdown("---")

    rodar = st.button("▶ Gerar Análise", use_container_width=True)


# =========================
# ÁREA PRINCIPAL
# =========================

st.markdown("""
<div class="hero-banner">
    <p class="hero-title">📦 Planejamento de Liberação de Pedidos</p>
    <p class="hero-sub">Simulação de liberação por estoque disponível, prioridade e data de faturamento</p>
</div>
""", unsafe_allow_html=True)

if not arquivo:
    st.info("👈 Carregue a planilha base na barra lateral para começar.")
    st.stop()

if df_global is None:
    erros = validar_colunas(preparar_base(pd.read_excel(arquivo)))
    st.error(f"❌ Colunas obrigatórias ausentes: **{', '.join(erros)}**")
    st.markdown("A planilha deve conter: " + ", ".join(f"`{c}`" for c in COLUNAS_OBRIGATORIAS))
    st.stop()

df = df_global

# Exibe prévia da base
with st.expander("🔎 Visualizar base carregada", expanded=False):
    st.dataframe(df.head(50), use_container_width=True)
    st.caption(f"{len(df):,} linhas • {df['PEDIDO'].nunique():,} pedidos únicos • {df['ITEM'].nunique():,} itens únicos")

if not rodar:
    # Exibe estatísticas da base sem rodar análise
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total de linhas", f"{len(df):,}")
    col2.metric("Pedidos únicos", f"{df['PEDIDO'].nunique():,}")
    col3.metric("Itens únicos", f"{df['ITEM'].nunique():,}")
    col4.metric("Clientes únicos", f"{df['CLIENTE'].nunique():,}")
    st.stop()

# =========================
# EXECUÇÃO DA ANÁLISE
# =========================

with st.spinner("Processando análise..."):

    # Filtro de data
    df_filtrado = df[
        (df["FATURAR EM"].dt.date >= data_inicio) &
        (df["FATURAR EM"].dt.date <= data_fim)
    ]

    if df_filtrado.empty:
        st.warning("⚠️ Nenhum pedido encontrado no período selecionado.")
        st.stop()

    # Filtros operacionais
    df_filtrado = aplicar_filtros_operacionais(
        df_filtrado, filtro_regiao, filtro_estado, filtro_cliente
    )

    if df_filtrado.empty:
        st.warning("⚠️ Nenhum pedido encontrado com os filtros aplicados.")
        st.stop()

    # Estoque (sempre montado sobre a base completa para não perder itens)
    estoque = montar_estoque(df)

    # Prioridade e ordem
    df_ordem = ordenar_por_prioridade(df_filtrado, prioridade_regiao, prioridade_estado)

    # Análise
    df_lib, df_bloq, df_cons, df_faltas, df_est_final = analisar_pedidos(df_ordem, estoque)

    # Resumo
    df_resumo = calcular_resumo(df_lib, df_bloq, df_cons, df_faltas)

# =========================
# MÉTRICAS GERAIS
# =========================
total_analisados = len(df_lib) + len(df_bloq)
pct_lib = (len(df_lib) / total_analisados * 100) if total_analisados else 0

st.markdown("### 📊 Resultado Geral")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Pedidos Analisados", f"{total_analisados:,}")
c2.metric("✅ Liberados", f"{len(df_lib):,}")
c3.metric("🔴 Bloqueados", f"{len(df_bloq):,}")
c4.metric("Taxa de Liberação", f"{pct_lib:.1f}%")
c5.metric("Itens em Falta", f"{int(df_faltas['FALTA_TOTAL'].sum()) if not df_faltas.empty else 0:,}")

st.markdown("---")

# =========================
# ABAS DE RESULTADO
# =========================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "✅ Liberados",
    "🔴 Não Liberados",
    "📉 Faltas por Item",
    "📦 Consumo por Item",
    "🏪 Estoque Final"
])

with tab1:
    if df_lib.empty:
        st.warning("Nenhum pedido foi liberado.")
    else:
        st.caption(f"{len(df_lib):,} pedidos liberados")
        st.dataframe(df_lib, use_container_width=True, hide_index=True)

with tab2:
    if df_bloq.empty:
        st.success("🎉 Todos os pedidos foram liberados!")
    else:
        st.caption(f"{len(df_bloq):,} pedidos bloqueados por falta de estoque")
        st.dataframe(df_bloq, use_container_width=True, hide_index=True)

with tab3:
    if df_faltas.empty:
        st.success("Sem itens em falta.")
    else:
        st.caption(f"{len(df_faltas):,} itens com falta")
        st.dataframe(df_faltas, use_container_width=True, hide_index=True)

with tab4:
    if df_cons.empty:
        st.info("Sem consumo registrado.")
    else:
        st.caption(f"{len(df_cons):,} itens consumidos")
        st.dataframe(df_cons, use_container_width=True, hide_index=True)

with tab5:
    st.caption(f"Estoque simulado após liberação dos pedidos possíveis")
    st.dataframe(df_est_final, use_container_width=True, hide_index=True)

st.markdown("---")

# =========================
# EXPORTAR EXCEL
# =========================
st.markdown("### 📥 Exportar Relatório")

excel_bytes = gerar_excel(
    df_lib, df_bloq, df_cons, df_faltas, df_est_final, df_resumo
)

col_dl, col_info = st.columns([1, 3])
with col_dl:
    st.download_button(
        label="⬇️ Baixar Relatório Excel",
        data=excel_bytes,
        file_name=f"relatorio_liberacao_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
with col_info:
    st.markdown(f"""
    <div style="color:#6B7A99;font-size:0.83rem;padding-top:8px;">
        O arquivo contém <strong>6 abas</strong>: Resumo · Liberados · Não Liberados · Consumo por Item · Faltas por Item · Estoque Final<br>
        Gerado em {date.today().strftime('%d/%m/%Y')}
    </div>
    """, unsafe_allow_html=True)