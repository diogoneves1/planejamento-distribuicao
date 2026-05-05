import pandas as pd

print("Iniciando análise...")

# =========================
# LER PLANILHA
# =========================
df = pd.read_excel("base.xlsx")

# =========================
# PADRONIZAR COLUNAS
# =========================
df.columns = df.columns.str.strip().str.upper()

print("Colunas encontradas:")
print(df.columns)

# =========================
# TRATAR DATA
# =========================
df["FATURAR EM"] = pd.to_datetime(df["FATURAR EM"]).dt.date
hoje = pd.Timestamp.today().date()

# =========================
# FILTRAR (HOJE + ATRASADOS)
# =========================
df_filtrado = df[df["FATURAR EM"] <= hoje]

# =========================
# ESTOQUE POR ITEM
# =========================
estoque = {}

for _, row in df.iterrows():
    item = row["ITEM"]
    if item not in estoque:
        estoque[item] = row["ESTOQUE INICIAL"]

# =========================
# RESULTADOS
# =========================
liberados = []
nao_liberados = []

# =========================
# PROCESSAR PEDIDOS
# =========================
for pedido, grupo in df_filtrado.groupby("PEDIDO"):

    pode_liberar = True
    faltas = []

    for _, row in grupo.iterrows():
        item = row["ITEM"]
        qtd = row["QNTD PROGRAMADA"]
        saldo = estoque.get(item, 0)

        if saldo < qtd:
            pode_liberar = False
            faltas.append((item, qtd))

    if pode_liberar:
        liberados.append({"PEDIDO": pedido})

        # consumir estoque
        for _, row in grupo.iterrows():
            item = row["ITEM"]
            estoque[item] -= row["QNTD PROGRAMADA"]

    else:
        linha = {"PEDIDO": pedido}

        for i, (item, qtd) in enumerate(faltas, start=1):
            linha[f"ITEM_{i}"] = item
            linha[f"QTD_{i}"] = qtd

        nao_liberados.append(linha)

# =========================
# EXPORTAR RESULTADO
# =========================
with pd.ExcelWriter("resultado.xlsx") as writer:
    pd.DataFrame(liberados).to_excel(writer, sheet_name="LIBERADOS", index=False)
    pd.DataFrame(nao_liberados).to_excel(writer, sheet_name="NAO_LIBERADOS", index=False)

print("Análise finalizada com sucesso.")