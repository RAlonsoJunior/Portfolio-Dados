import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output, dash_table
import dash_bootstrap_components as dbc

# ======== LEITURA DOS DADOS ========
df_planocontas = pd.read_excel("BaseDados.xlsx", sheet_name="Plano Contas")
df_pag = pd.read_excel("BaseDados.xlsx", sheet_name="Pagamentos")
df_rec = pd.read_excel("BaseDados.xlsx", sheet_name="Recebimentos")

# Adiciona tipo aos dados
df_pag["Tipo"] = "Pagamento"
df_rec["Tipo"] = "Recebimento"

# Padroniza e concatena
df = pd.concat([df_pag, df_rec], ignore_index=True)
df["Valor"] = df["Valor Pago"].fillna(0) + df["Valor Recebido"].fillna(0)

# Corrige colunas para merge
df_planocontas.columns = df_planocontas.columns.str.strip()
df = df.merge(df_planocontas, how="left", left_on="Cod Plano Contas", right_on="Cod Conta")

# Datas e colunas auxiliares
df["Data"] = pd.to_datetime(df["Data Emissão"])
df["Ano"] = df["Data"].dt.year
df["Mês"] = df["Data"].dt.strftime("%b")

# Ordenação de meses
meses_ordem = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

# ======== DASH APP ========
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

app.layout = dbc.Container([
    html.H2("Dashboard Financeiro", className="text-center my-4"),

    # Filtros
    dbc.Row([
        dbc.Col(dcc.Dropdown(
            id="filtro_ano",
            options=[{"label": str(a), "value": a} for a in sorted(df["Ano"].unique())],
            value=sorted(df["Ano"].unique())[-1],
            placeholder="Ano"
        ), width=2),

        dbc.Col(dcc.Dropdown(
            id="filtro_mes",
            options=[{"label": m, "value": m} for m in df["Mês"].unique()],
            placeholder="Mês (opcional)"
        ), width=2),

        dbc.Col(dcc.Dropdown(
            id="filtro_tipo",
            options=[{"label": t, "value": t} for t in df["Tipo"].unique()],
            placeholder="Tipo"
        ), width=2),

        dbc.Col(dcc.Dropdown(
            id="filtro_fornecedor",
            options=[{"label": f, "value": f} for f in sorted(df["Fornecedor"].dropna().unique())],
            placeholder="Fornecedor"
        ), width=3),

        dbc.Col(dcc.Dropdown(
            id="filtro_conta",
            options=[{"label": c, "value": c} for c in sorted(df["Conta Nível 3"].dropna().unique())],
            placeholder="Conta Nível 3"
        ), width=3),
    ], className="mb-4"),

    # KPIs
    dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Receita"), dbc.CardBody(id="kpi_receita")]), width=2),
        dbc.Col(dbc.Card([dbc.CardHeader("Despesas"), dbc.CardBody(id="kpi_despesa")]), width=2),
        dbc.Col(dbc.Card([dbc.CardHeader("Margem"), dbc.CardBody(id="kpi_margem")]), width=2),
    ], className="mb-4"),

    # Gráficos
    dbc.Row([
        dbc.Col(dcc.Graph(id="grafico_categoria"), width=6),
        dbc.Col(dcc.Graph(id="grafico_mes"), width=6),
    ]),

    html.Hr(),

    html.H5("Tabela Financeira"),
    dash_table.DataTable(
        id="tabela_financeira",
        columns=[
            {"name": "Data", "id": "Data"},
            {"name": "Tipo", "id": "Tipo"},
            {"name": "Fornecedor", "id": "Fornecedor"},
            {"name": "Conta Nível 3", "id": "Conta Nível 3"},
            {"name": "Valor", "id": "Valor", "type": "numeric", "format": {"specifier": ".2f"}}
        ],
        page_size=10,
        style_table={"overflowX": "auto"},
    )

], fluid=True)

# ======== CALLBACKS ========
@app.callback(
    Output("kpi_receita", "children"),
    Output("kpi_despesa", "children"),
    Output("kpi_margem", "children"),
    Output("grafico_categoria", "figure"),
    Output("grafico_mes", "figure"),
    Output("tabela_financeira", "data"),
    Input("filtro_ano", "value"),
    Input("filtro_mes", "value"),
    Input("filtro_tipo", "value"),
    Input("filtro_fornecedor", "value"),
    Input("filtro_conta", "value"),
)
def atualizar(ano, mes, tipo, fornecedor, conta):
    df_filt = df[df["Ano"] == ano]
    if mes:
        df_filt = df_filt[df_filt["Mês"] == mes]
    if tipo:
        df_filt = df_filt[df_filt["Tipo"] == tipo]
    if fornecedor:
        df_filt = df_filt[df_filt["Fornecedor"] == fornecedor]
    if conta:
        df_filt = df_filt[df_filt["Conta Nível 3"] == conta]

    receita = df_filt[df_filt["Tipo"] == "Recebimento"]["Valor"].sum()
    despesa = df_filt[df_filt["Tipo"] == "Pagamento"]["Valor"].sum()
    margem = receita - despesa

    # Gráfico por categoria
    df_cat = df_filt.groupby("Conta Nível 3")["Valor"].sum().reset_index()
    fig_cat = px.bar(df_cat.sort_values("Valor"), x="Valor", y="Conta Nível 3", orientation="h", title="Valor por Categoria")

    # Gráfico mensal
    df_mes = df[df["Ano"] == ano].groupby("Mês")["Valor"].sum().reindex(meses_ordem).reset_index()
    fig_mes = px.bar(df_mes, x="Mês", y="Valor", title=f"Total Financeiro por Mês - {ano}")

    # Tabela
    tabela = df_filt[["Data", "Tipo", "Fornecedor", "Conta Nível 3", "Valor"]].copy()
    tabela["Data"] = tabela["Data"].dt.strftime("%d/%m/%Y")

    return (
        f"R$ {receita:,.2f}",
        f"R$ {despesa:,.2f}",
        f"R$ {margem:,.2f}",
        fig_cat,
        fig_mes,
        tabela.to_dict("records")
    )

# Rodar app
if __name__ == "__main__":
    app.run(debug=True)