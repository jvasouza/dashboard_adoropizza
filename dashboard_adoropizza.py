import re
from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata
from datetime import date, timedelta
import calendar
import locale
import plotly.io as pio
import numpy as np
import plotly.graph_objects as go


def estilizar_fig(fig):
    fig.update_layout(
        paper_bgcolor="#fefaf2",
        plot_bgcolor="#fefaf2",
        font=dict(color="#5f100e"),
        legend=dict(bgcolor="#fefaf2")
    )
    fig.update_xaxes(gridcolor="#eadfcb", zerolinecolor="#eadfcb")
    fig.update_yaxes(gridcolor="#eadfcb", zerolinecolor="#eadfcb")
    return fig

TONS_TERROSOS = [
    "#5F100E",
    "#A9210E",
    "#CD853F",
    "#D9C77C",
    "#DEB887",
    "#F5DEB3"
]

pio.templates["bene_tema"] = dict(
    layout=dict(
        colorway=TONS_TERROSOS,
        plot_bgcolor="#fefaf2",
        paper_bgcolor="#fefaf2",
        font=dict(color="#5f100e"),
        xaxis=dict(gridcolor="#eadfcb", zerolinecolor="#eadfcb"),
        yaxis=dict(gridcolor="#eadfcb", zerolinecolor="#eadfcb"),
        legend=dict(bgcolor="#fefaf2")
    )
)
pio.templates.default = "bene_tema"

st.set_page_config(
    page_title="Dashboard Adoro Pizza",
    page_icon="logo favicon.png",
    layout="wide"
)

col_titulo, col_logo = st.columns([6, 1])  # ajuste a proporção se quiser

with col_titulo:
    st.title("Dashboard - Adoro Pizza")

with col_logo:
    st.image("logo sidebar.png", width=120)

st.title("Dashboard - Adoro Pizza")


st.markdown("""
<style>
.stApp { background-color:#fefaf2; color:#5f100e; }
h1, h2, h3, h4, h5, h6 { color:#5f100e !important; font-weight:700; }
[data-testid="stSidebar"] {
    background-color:#5f100e !important;
    color:#fefaf2 !important;
    padding-top:0 !important;
    margin-top:0 !important;
}
section[data-testid="stSidebar"] div[role="button"] { display:none !important; }
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] h4,
[data-testid="stSidebar"] h5,
[data-testid="stSidebar"] h6,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] span { color:#fefaf2 !important; }
[data-testid="stSidebar"] .stButton>button {
    background-color:#fefaf2 !important;
    color:#5f100e !important;
    border-radius:10px !important;
    border:none !important;
    font-weight:700 !important;
    padding:0.5rem 0.75rem !important;
}
[data-testid="stSidebar"] .stButton>button:hover { background-color:#f4e9d4 !important; }
[data-testid="stSidebar"] .stButton>button * { color:#5f100e !important; }
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] .stDateInput input {
    color:#5f100e !important;
    background-color:#fefaf2 !important;
    border-radius:10px !important;
}
[data-testid="stMetricLabel"], [data-testid="stMetricValue"] { color:#5f100e !important; }
hr { border-top:2px solid #5f100e !important; }
</style>
""", unsafe_allow_html=True)

DATA = Path(__file__).parent / "data"
arq_itens = DATA / "Historico_Itens_Vendidos.xlsx"
arq_pedidos = DATA / "Todos os pedidos.xlsx"
arq_contas = DATA / "Lista-contas-receber.xlsx"
arq_custo_bebidas = DATA / "custo bebidas.xlsx"
arq_custo_pizzas = DATA / "custo_pizzas.xlsx"
arq_custos_fixos = DATA / "custos fixos.xlsx"
arq_compras = DATA / "compras.xlsx"
arq_ads_manual = DATA / "relatorio ads.xlsx"
arq_ads_manager = DATA / "relatorio-04-12-25.xlsx"

def set_locale_ptbr():
    for loc in ("pt_BR.UTF-8", "pt_BR.utf8", "pt_BR", "Portuguese_Brazil.1252"):
        try:
            locale.setlocale(locale.LC_TIME, loc)
            return loc
        except locale.Error:
            continue
    locale.setlocale(locale.LC_TIME, "C")
    return "C"

_ = set_locale_ptbr()

def filtro_periodo_global(series_dt):
    st.sidebar.header("📅 Período")
    s = pd.to_datetime(series_dt, errors="coerce").dropna()
    if s.empty:
        st.sidebar.info("Sem datas válidas para filtrar.")
        return None, None

    dmin = s.min().date()
    dmax = s.max().date()

    if "data_ini" not in st.session_state or "data_fim" not in st.session_state:
        hoje = date.today()
        fim_padrao = min(hoje, dmax)
        ini_padrao = max(fim_padrao - timedelta(days=30), dmin)
        st.session_state["data_ini"] = ini_padrao
        st.session_state["data_fim"] = fim_padrao

    anos = sorted(s.dt.year.unique())
    ano_sel = st.sidebar.selectbox("Ano", anos, index=len(anos)-1, key="ano_btns")

    nomes_pt = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    cols = st.sidebar.columns(2)
    for i, mes in enumerate(range(1, 13)):
        col = cols[i % 2]
        if col.button(nomes_pt[mes-1], key=f"btn_mes_{ano_sel}_{mes}"):
            from calendar import monthrange
            ini = date(ano_sel, mes, 1)
            fim = date(ano_sel, mes, monthrange(ano_sel, mes)[1])
            ini = max(ini, dmin)
            fim = min(fim, dmax)
            if ini <= fim:
                st.session_state["data_ini"] = ini
                st.session_state["data_fim"] = fim
                st.rerun()

    c1, c2 = st.sidebar.columns(2)
    if c1.button("Ano selecionado", key="btn_full_year"):
        ini = date(ano_sel, 1, 1)
        fim = date(ano_sel, 12, 31)
        st.session_state["data_ini"] = max(ini, dmin)
        st.session_state["data_fim"] = min(fim, dmax)
        st.rerun()
    if c2.button("Todo o período", key="btn_all_data"):
        st.session_state["data_ini"] = dmin
        st.session_state["data_fim"] = dmax
        st.rerun()

    data_ini = st.session_state.get("data_ini", dmin)
    data_fim = st.session_state.get("data_fim", dmax)
    data_ini = max(min(data_ini, dmax), dmin)
    data_fim = max(min(data_fim, dmax), dmin)
    if data_ini > data_fim:
        data_ini, data_fim = dmin, dmax

    c1, c2 = st.sidebar.columns(2)
    dini = c1.date_input("Início", value=data_ini, min_value=dmin, max_value=dmax, key="ini_input")
    dfim = c2.date_input("Fim", value=data_fim, min_value=dmin, max_value=dmax, key="fim_input")
    if dini > dfim:
        dini, dfim = dmin, dmax

    st.session_state["data_ini"], st.session_state["data_fim"] = dini, dfim

    return dini, dfim


def carregar_primeira_aba_xlsx(arquivo, caminho):
    import zipfile
    from pathlib import Path as _P
    if arquivo is not None:
        p = _P(arquivo)
    elif caminho is not None:
        p = _P(caminho)
    else:
        st.stop()
    if not p.exists():
        st.stop()
    if p.suffix.lower() != ".xlsx":
        st.stop()
    with open(p, "rb") as fh:
        if not zipfile.is_zipfile(fh):
            st.stop()
    xls = pd.ExcelFile(p, engine="openpyxl")
    primeira = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=primeira, engine="openpyxl")
    return df

def carregou(df):
    return df is not None and len(df) > 0

def br_money(x):
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def sem_acentos_upper(s):
    if pd.isna(s):
        return s
    s = str(s).strip().upper()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return " ".join(s.split())

def padroniza_pizza_nome_tamanho(nome):
    nome = str(nome).strip()
    if nome.startswith("PIZZA "):
        nome = nome.replace("PIZZA ", "", 1)
    nome = nome.replace(" Grande", " G")
    nome = nome.replace(" Média", " M")
    nome = nome.replace(" Pequena", " P")
    return nome

def nomes_legiveis(df):
    mapa = {
        "data": "Data",
        "valor_liq": "Valor",
        "forma_pagamento": "Forma de Pagamento",
        "dow": "Dia da Semana",
        "pedidos": "Pedidos",
        "receita": "Receita (R$)",
        "cliente": "Cliente",
        "gasto": "Valor (R$)",
        "cod_pedido": "Código do Pedido",
        "total_pedido": "Total do Pedido (R$)",
        "tipo_norm": "Tipo de Pedido",
        "total": "Total (R$)",
        "total_recebido": "Total Recebido (R$)",
        "categoria": "Categoria",
        "produto": "Produto",
        "qtd": "Qtd",
        "cmv": "CMV (R$)",
        "margem": "Margem (R$)",
        "margem_%": "Margem (%)"
    }
    df_formatado = df.rename(columns={c: mapa.get(c, c) for c in df.columns}).copy()
    for col in df_formatado.columns:
        if ("R$" in col or "(R$)" in col or "Valor" in col or "Receita" in col or "CMV" in col or "Margem" in col):
            if "%" not in col and pd.api.types.is_numeric_dtype(df_formatado[col]):
                df_formatado[col] = df_formatado[col].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    return df_formatado

df_periodo_base = carregar_primeira_aba_xlsx(arq_contas, None)
data_series = pd.to_datetime(df_periodo_base["Crédito"], errors="coerce") if carregou(df_periodo_base) else pd.Series([], dtype="datetime64[ns]")
if not data_series.empty and data_series.notna().any():
    data_ini, data_fim = filtro_periodo_global(data_series.dropna())
else:
    data_ini, data_fim = None, None

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Faturamento", "Pedidos", "CMV", "Metas", "Promoções"])

with tab1:
    df = carregar_primeira_aba_xlsx(arq_contas, None)
    if not carregou(df):
        st.info("Carregue a planilha de Contas a Receber para visualizar a aba Faturamento.")
    else:
        df = df.copy()
        df.columns = df.columns.str.strip()
        df = df.rename(columns={"Cód. Pedido":"cod_pedido","Valor Líq.":"valor_liq","Forma Pagamento":"forma_pagamento","Crédito":"data","Total Pedido":"total_pedido"})
        df["data"] = pd.to_datetime(df["data"], errors="coerce")
        df["valor_liq"] = pd.to_numeric(df["valor_liq"], errors="coerce")
        if data_ini is None or data_fim is None:
            data_ini = pd.to_datetime(df["data"]).min().date()
            data_fim = pd.to_datetime(df["data"]).max().date()
        def normaliza_pagto(x):
            s = str(x).strip().upper()
            if s in {"PIX", "PIX MANUAL", "A CONFIRMAR", "VALE REFEICAO", "VALE REFEIÇÃO"}:
                return "PIX"
            return s
        df["forma_pagamento"] = df["forma_pagamento"].apply(normaliza_pagto)
        mask = (df["data"] >= pd.to_datetime(data_ini)) & (df["data"] <= pd.to_datetime(data_fim))
        dff = df.loc[mask].copy()
        fat_total = float(dff["valor_liq"].sum())
        n_pedidos = int(dff["cod_pedido"].nunique())
        ticket_medio = fat_total / n_pedidos if n_pedidos else 0
        dias_periodo = max(1, (pd.to_datetime(data_fim) - pd.to_datetime(data_ini)).days + 1)
        fat_medio_dia = fat_total / dias_periodo
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Faturamento Total (R$)", br_money(fat_total))
        k2.metric("Total de Pedidos", f"{n_pedidos}")
        k3.metric("Ticket Médio (R$)", br_money(ticket_medio))
        k4.metric("Faturamento Médio/Dia (R$)", br_money(fat_medio_dia))
        st.subheader("Evolução do Faturamento Diário")
        dff["dia"] = dff["data"].dt.date
        fat_dia = dff.groupby("dia", as_index=False)["valor_liq"].sum().sort_values("dia")
        mapper = {0:"Seg",1:"Ter",2:"Qua",3:"Qui",4:"Sex",5:"Sáb",6:"Dom"}
        fat_dia["dow"] = pd.to_datetime(fat_dia["dia"]).dt.weekday.map(mapper)
        fig_fat = px.line(
            fat_dia,
            x="dia", y="valor_liq",
            markers=True,
            labels={"dia":"Data","valor_liq":"Receita (R$)"},
            color_discrete_sequence=TONS_TERROSOS
        )
        fig_fat = estilizar_fig(fig_fat)
        fig_fat.update_xaxes(type="date", tickformat="%d/%m/%Y", ticklabelmode="period", tickangle=-45)
        fig_fat.update_traces(
            hovertemplate="Data: %{x|%d/%m/%Y}<br>Dia da semana: %{customdata[0]}<br>Receita: R$ %{y:.2f}",
            customdata=fat_dia[["dow"]].to_numpy()
        )
        # após criar fat_dia e fig_fat
        periodo_dias = (pd.to_datetime(data_fim) - pd.to_datetime(data_ini)).days + 1
        if periodo_dias > 180 and len(fat_dia) > 1:
            fat_dia_sorted = fat_dia.sort_values("dia").copy()
            fat_dia_sorted["mm30"] = fat_dia_sorted["valor_liq"].rolling(30, min_periods=1).mean()
            fig_fat.add_scatter(
                x=fat_dia_sorted["dia"],
                y=fat_dia_sorted["mm30"],
                mode="lines",
                name="Média móvel (30 dias)",
    )

        st.plotly_chart(fig_fat, use_container_width=True, key="fat_linha_dia")
        st.divider()
        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("Receita por Forma de Pagamento")
            fat_pagto = dff.groupby("forma_pagamento", as_index=False)["valor_liq"].sum().sort_values("valor_liq", ascending=False)
            fig_pagto = px.pie(fat_pagto, names="forma_pagamento", values="valor_liq", hole=0.3)
            fig_pagto = estilizar_fig(fig_pagto)
            fig_pagto.update_traces(textinfo="percent+label")
            st.plotly_chart(fig_pagto, use_container_width=True, key="fat_pizza_pagto")
            st.dataframe(nomes_legiveis(fat_pagto.reset_index(drop=True)), use_container_width=True, hide_index=True)
        with col_b:
            st.subheader("Faturamento por Dia da Semana")
            mapper = {0:"Seg",1:"Ter",2:"Qua",3:"Qui",4:"Sex",5:"Sáb",6:"Dom"}
            dff["dow"] = dff["data"].dt.weekday.map(mapper)
            ordem = ["Seg","Ter","Qua","Qui","Sex","Sáb","Dom"]
            fat_dow = dff.groupby("dow", as_index=False)["valor_liq"].sum()
            fat_dow["dow"] = pd.Categorical(fat_dow["dow"], categories=ordem, ordered=True)
            fat_dow = fat_dow.sort_values("dow")
            fig_dow = px.bar(fat_dow, x="dow", y="valor_liq", labels={"dow":"Dia da Semana","valor_liq":"Receita (R$)"})
            fig_dow = estilizar_fig(fig_dow)
            st.plotly_chart(fig_dow, use_container_width=True, key="fat_barras_dow")
            st.dataframe(nomes_legiveis(fat_dow.reset_index(drop=True)), use_container_width=True, hide_index=True)

with tab2:
    dfp = carregar_primeira_aba_xlsx(arq_pedidos, None)
    if not carregou(dfp):
        st.info("Carregue a planilha de Pedidos para visualizar a aba Pedidos.")
    else:
        dfp = dfp.copy()
        dfp.columns = dfp.columns.str.strip()
        rename_map = {"Código":"codigo","Data Abertura":"data","Status":"status","Cliente":"cliente","Tipo":"tipo","Origem":"origem","Total":"total","Total Recebido":"total_recebido","Forma de Pagto":"forma_pagto"}
        dfp = dfp.rename(columns=rename_map)
        dfp["data"] = pd.to_datetime(dfp["data"], errors="coerce")
        if data_ini is None or data_fim is None:
            data_ini = pd.to_datetime(dfp["data"]).min().date()
            data_fim = pd.to_datetime(dfp["data"]).max().date()
        maskp = (dfp["data"] >= pd.to_datetime(data_ini)) & (dfp["data"] <= pd.to_datetime(data_fim))
        dpp = dfp.loc[maskp].copy()
        pedidos_total = int(dpp["codigo"].nunique())
        receita_periodo = float(dpp["total_recebido"].sum())
        ticket_medio = receita_periodo / pedidos_total if pedidos_total else 0
        k1, k2, k3 = st.columns(3)
        k1.metric("Pedidos no período", f"{pedidos_total}")
        k2.metric("Ticket Médio (R$)", br_money(ticket_medio))
        st.divider()
        st.subheader("Evolução do Nº de Pedidos por Dia")
        dpp["dia"] = dpp["data"].dt.date
        pedidos_por_dia = dpp.groupby("dia", as_index=False)["codigo"].nunique().rename(columns={"codigo": "pedidos"}).sort_values("dia")
        mapper = {0:"Seg",1:"Ter",2:"Qua",3:"Qui",4:"Sex",5:"Sáb",6:"Dom"}
        pedidos_por_dia["dow"] = pd.to_datetime(pedidos_por_dia["dia"]).dt.weekday.map(mapper)
        fig_ped_dia = px.line(
            pedidos_por_dia,
            x="dia", y="pedidos",
            markers=True,
            labels={"dia":"Data", "pedidos":"Pedidos"},
            color_discrete_sequence=TONS_TERROSOS
        )
        fig_ped_dia = estilizar_fig(fig_ped_dia)
        fig_ped_dia.update_xaxes(tickformat="%d/%m/%Y")
        fig_ped_dia.update_traces(
            hovertemplate="Data: %{x|%d/%m/%Y}<br>Dia da semana: %{customdata[0]}<br>Pedidos: %{y}",
            customdata=pedidos_por_dia[["dow"]].to_numpy()
        )
        st.plotly_chart(fig_ped_dia, use_container_width=True, key="ped_linha_dia")
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Nº de Pedidos por Tipo")
            pedidos_tipo = dpp.groupby("tipo", as_index=False)["codigo"].nunique().rename(columns={"codigo":"pedidos"})
            fig_pt = px.bar(pedidos_tipo, x="tipo", y="pedidos", labels={"tipo":"Tipo","pedidos":"Pedidos"})
            fig_pt = estilizar_fig(fig_pt)
            st.plotly_chart(fig_pt, use_container_width=True, key="ped_barras_tipo")
            st.dataframe(nomes_legiveis(pedidos_tipo.reset_index(drop=True)), use_container_width=True, hide_index=True)
        with c2:
            st.subheader("Receita por Tipo")
            receita_tipo = dpp.groupby("tipo", as_index=False)["total_recebido"].sum().rename(columns={"total_recebido":"receita"})
            fig_rt = px.pie(receita_tipo, names="tipo", values="receita", hole=0.3)
            fig_rt = estilizar_fig(fig_rt)
            fig_rt.update_traces(textinfo="percent+label")
            st.plotly_chart(fig_rt, use_container_width=True, key="ped_pizza_tipo")
            st.dataframe(nomes_legiveis(receita_tipo.reset_index(drop=True)), use_container_width=True, hide_index=True)
        st.divider()
        st.subheader("Top 10 Clientes por Nº de Pedidos")
        dpp_top = dpp[~dpp["cliente"].astype(str).str.strip().str.lower().eq("não informado")]
        top_cli = (dpp_top.groupby("cliente", as_index=False)
                    .agg(pedidos=("codigo","nunique"), gasto=("total_recebido","sum"))
                    .sort_values(["pedidos","gasto"], ascending=[False, False])
                    .head(10)
                    .reset_index(drop=True))
        st.dataframe(nomes_legiveis(top_cli), use_container_width=True, hide_index=True)

with tab3:
    itens = carregar_primeira_aba_xlsx(arq_itens, None)
    c_pizzas = carregar_primeira_aba_xlsx(arq_custo_pizzas, None)
    c_bebidas = carregar_primeira_aba_xlsx(arq_custo_bebidas, None)
    if not (carregou(itens) and carregou(c_pizzas) and carregou(c_bebidas)):
        st.info("Carregue as planilhas: Itens Vendidos, Custo Pizzas e Custo Bebidas para visualizar a aba CMV.")
    else:
        itens = itens.copy()
        itens.columns = itens.columns.str.strip()
        itens = itens.rename(columns={
            "Data/Hora Item":"data_item",
            "Qtd.":"qtd",
            "Valor. Tot. Item":"valor_tot",
            "Nome Prod":"nome_prod",
            "Cat. Prod.":"cat_prod",
            "Cod. Ped.":"codigo"
        })
        itens["nome_prod_norm"] = itens["nome_prod"].astype(str).str.strip()
        itens = itens[~itens["nome_prod_norm"].str.startswith("* Excluído *", na=False)].copy()
        itens["data_item"] = pd.to_datetime(itens["data_item"], errors="coerce")
        itens["qtd"] = pd.to_numeric(itens["qtd"], errors="coerce").fillna(0)
        itens["valor_tot"] = pd.to_numeric(itens["valor_tot"], errors="coerce").fillna(0)
        itens = itens.dropna(subset=["data_item"]).copy()

        def normalize_sizes(text):
            s = text.str.replace(r"\bGrande\b","G",regex=True)
            s = s.str.replace(r"\bM[eé]dia\b","M",regex=True)
            s = s.str.replace(r"\bPequena\b","P",regex=True)
            return s

        def normalize_key_general(s):
            t = s.astype(str)
            t = t.str.replace(r"^\s*Pizza\s+","",regex=True)
            t = normalize_sizes(t)
            t = t.str.replace(r"\s{2,}"," ",regex=True).str.strip()
            return t

        def clean_nome_prod_hist(nome_series, cat_series):
            s = nome_series.astype(str)
            s = s.str.replace(r"^PIZZA\s+", "", regex=True, case=False)
            s = normalize_sizes(s)
            mask_sucos = cat_series.astype(str).str.upper().eq("SUCOS")
            sabores = r"(LARANJA|ABACAXI|MARACUJ[ÁA])"
            s2 = s.copy()
            s2.loc[mask_sucos] = s2.loc[mask_sucos].str.replace(rf"(\bSUCO)\s+{sabores}\s+",r"\1 ",flags=re.IGNORECASE,regex=True)
            s2 = s2.str.replace(r"^carnes\s+","",regex=True, flags=re.IGNORECASE)
            s2 = s2.str.replace(r"^(?:batata frita\s+){2}", "BATATA FRITA ", flags=re.IGNORECASE, regex=True)
            mask_rodizio = s2.str.contains(r"rod[ií]zio", flags=re.IGNORECASE, regex=True)
            s2.loc[mask_rodizio] = "RODÍZIO DE PIZZA"
            s2 = s2.str.replace(r"\s{2,}"," ",regex=True).str.strip()
            return s2

        if data_ini is None or data_fim is None:
            data_ini = pd.to_datetime(itens["data_item"]).min().date()
            data_fim = pd.to_datetime(itens["data_item"]).max().date()

        mask_periodo = (itens["data_item"] >= pd.to_datetime(data_ini)) & (itens["data_item"] <= pd.to_datetime(data_fim))
        iv = itens.loc[mask_periodo].copy()
        iv["cat_norm"] = iv["cat_prod"]
        iv["nome_limpo"] = clean_nome_prod_hist(iv["nome_prod"], iv["cat_prod"])
        iv["valor_base"] = iv["valor_tot"]

        c_pizzas = c_pizzas.copy()
        c_bebidas = c_bebidas.copy()
        c_pizzas.columns = c_pizzas.columns.str.strip()
        c_bebidas.columns = c_bebidas.columns.str.strip()
        c_pizzas["_KEY"] = normalize_key_general(c_pizzas["produto"])
        c_bebidas["_KEY"] = normalize_key_general(c_bebidas["produto"])

        lookup_pizza = c_pizzas.set_index("_KEY")["custo"]
        lookup_bebida = c_bebidas.set_index("_KEY")["custo"]
        iv["custo_pizza"] = iv["nome_limpo"].map(lookup_pizza)
        iv["custo_bebida"] = iv["nome_limpo"].map(lookup_bebida)
        iv["custo_unit"] = iv["custo_pizza"].combine_first(iv["custo_bebida"])

        mask_complemento = iv["cat_norm"].astype(str).str.strip().str.upper().eq("COMPLEMENTO")
        iv["cmv_item"] = np.where(mask_complemento, 0.5 * iv["valor_base"], iv["custo_unit"] * iv["qtd"])

        if arq_compras.exists() and arq_pedidos.exists():
            dfc = carregar_primeira_aba_xlsx(arq_compras, None)
            dfp_min = carregar_primeira_aba_xlsx(arq_pedidos, None)
            if carregou(dfc) and carregou(dfp_min):
                dfc = dfc.copy()
                dfc.columns = dfc.columns.str.strip()
                alvo = ["CAIXA PIZZA G", "CAIXA PIZZA M", "CAIXA PIZZA P"]
                dfc = dfc[dfc["nome_interno"].isin(alvo)].copy()
                dfc = dfc.dropna(subset=["valor_por_unidade"]).copy()
                dfc = dfc.groupby("nome_interno", as_index=False).last()
                preco_caixa = {
                    "G": float(dfc.loc[dfc["nome_interno"].eq("CAIXA PIZZA G"), "valor_por_unidade"].iloc[0]) if (dfc["nome_interno"]=="CAIXA PIZZA G").any() else np.nan,
                    "M": float(dfc.loc[dfc["nome_interno"].eq("CAIXA PIZZA M"), "valor_por_unidade"].iloc[0]) if (dfc["nome_interno"]=="CAIXA PIZZA M").any() else np.nan,
                    "P": float(dfc.loc[dfc["nome_interno"].eq("CAIXA PIZZA P"), "valor_por_unidade"].iloc[0]) if (dfc["nome_interno"]=="CAIXA PIZZA P").any() else np.nan,
                }
                dfp_min = dfp_min.copy()
                dfp_min.columns = dfp_min.columns.str.strip()
                dfp_min = dfp_min.rename(columns={"Código": "codigo", "Tipo": "tipo"})
                dfp_min = dfp_min[["codigo", "tipo"]].dropna(subset=["codigo"]).copy()
                iv = iv.merge(dfp_min, on="codigo", how="left")
                iv["tamanho_pizza"] = iv["nome_limpo"].str.extract(r"\b([GMP])\b", expand=False)
                tipo_upper = iv["tipo"].astype(str).str.upper()
                mask_tipo = tipo_upper.isin(["DELIVERY", "BALCÃO", "BALCAO", "CAIXA"])
                cat_upper = iv["cat_norm"].astype(str).str.upper()
                mask_cat = cat_upper.isin(["PIZZAS", "CARNES", "PORÇÕES", "PORCOES"])
                mask_sz = iv["tamanho_pizza"].isin(["G", "M", "P"])
                m = mask_tipo & mask_cat & mask_sz
                if m.any():
                    iv.loc[m, "qtd_total_tmp"] = iv.loc[m].groupby(["codigo", "tamanho_pizza"])["qtd"].transform("sum")
                    iv.loc[m, "share_tmp"] = np.where(iv.loc[m, "qtd_total_tmp"] > 0, iv.loc[m, "qtd"] / iv.loc[m, "qtd_total_tmp"], 0.0)
                    grp = iv.loc[m].groupby(["codigo", "tamanho_pizza"], as_index=False).agg(qtd_total=("qtd", "sum"))
                    grp["n_caixas"] = np.floor(grp["qtd_total"] + 0.5)
                    grp["preco_caixa_unit"] = grp["tamanho_pizza"].map(preco_caixa).astype(float)
                    grp["custo_caixa_total"] = grp["n_caixas"] * grp["preco_caixa_unit"]
                    iv = iv.merge(grp[["codigo", "tamanho_pizza", "custo_caixa_total"]], on=["codigo", "tamanho_pizza"], how="left")
                    iv["custo_caixa_alocado"] = np.where(m, iv["share_tmp"] * iv["custo_caixa_total"], 0.0)
                    iv["cmv_item"] = iv["cmv_item"] + iv["custo_caixa_alocado"].fillna(0.0)
                    iv.drop(columns=["qtd_total_tmp", "share_tmp", "custo_caixa_total", "custo_caixa_alocado"], errors="ignore", inplace=True)

        cmv_total = float(iv["cmv_item"].sum(skipna=True))

        df_contas_custos = carregar_primeira_aba_xlsx(arq_contas, None)
        receita_total = 0.0
        if carregou(df_contas_custos):
            dfc2 = df_contas_custos.copy()
            dfc2.columns = dfc2.columns.str.strip()
            dfc2 = dfc2.rename(columns={"Cód. Pedido":"cod_pedido","Valor Líq.":"valor_liq","Forma Pagamento":"forma_pagamento","Crédito":"data"})
            dfc2["data"] = pd.to_datetime(dfc2["data"], errors="coerce")
            dfc2 = dfc2.dropna(subset=["data","valor_liq","cod_pedido"]).copy()
            def normaliza_pagto2(x):
                s = str(x).strip().upper()
                if s in {"PIX", "PIX MANUAL", "A CONFIRMAR", "VALE REFEICAO", "VALE REFEIÇÃO"}:
                    return "PIX"
                return s
            dfc2["forma_pagamento"] = dfc2["forma_pagamento"].apply(normaliza_pagto2)
            mask_receita = (dfc2["data"] >= pd.to_datetime(data_ini)) & (dfc2["data"] <= pd.to_datetime(data_fim))
            dfr = dfc2.loc[mask_receita].copy()
            receita_total = float(dfr["valor_liq"].sum())

        df_cfix = carregar_primeira_aba_xlsx(None, arq_custos_fixos)
        dias_periodo = (pd.to_datetime(data_fim) - pd.to_datetime(data_ini)).days + 1
        total_cfix, tabela_cfix = 0.0, pd.DataFrame()
        if dias_periodo >= 30 and carregou(df_cfix):
            dfc = df_cfix.copy()
            dfc.columns = dfc.columns.str.strip()
            dfc = dfc.rename(columns={"DATA":"data","DESCRIÇÃO":"descricao","VALOR":"valor"})
            dfc["data"] = pd.to_datetime(dfc["data"], errors="coerce")
            dfc["valor"] = pd.to_numeric(dfc["valor"], errors="coerce")
            dfc = dfc.dropna(subset=["data","valor"])
            mask_fix = (dfc["data"] >= pd.to_datetime(data_ini)) & (dfc["data"] <= pd.to_datetime(data_fim))
            aloc = dfc.loc[mask_fix, ["data","descricao","valor"]].copy()
            if not aloc.empty:
                aloc["Mês"] = aloc["data"].dt.to_period("M").astype(str)
                aloc = aloc[["Mês","descricao","valor"]].rename(columns={"descricao":"Descrição","valor":"Valor (R$)"})
                total_cfix = float(aloc["Valor (R$)"].sum())
                tabela_cfix = aloc

        margem_bruta = receita_total - cmv_total
        margem_bruta_pct = (margem_bruta / receita_total * 100) if receita_total else 0.0
        margem_liquida = receita_total - cmv_total - total_cfix
        margem_liquida_pct = (margem_liquida / receita_total * 100) if receita_total else 0.0

        kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
        kpi1.metric("Receita (R$)", br_money(receita_total))
        kpi2.metric("CMV (R$)", br_money(cmv_total))
        kpi3.metric("Margem Bruta (R$)", br_money(margem_bruta))
        kpi4.metric("Margem Bruta (%)", f"{margem_bruta_pct:.1f}%")
        kpi5.metric("Custos Fixos (R$)", br_money(total_cfix))
        kpi6.metric("Margem Líquida (R$)", br_money(margem_liquida))

        st.subheader("Custos Fixos no Período")
        if dias_periodo < 30:
            st.info("Período menor que 30 dias: custos fixos e margem líquida ignorados.")
        else:
            if not tabela_cfix.empty:
                st.dataframe(nomes_legiveis(tabela_cfix.reset_index(drop=True)), use_container_width=True, hide_index=True)
            else:
                st.info("Sem custos fixos para o período selecionado ou arquivo ausente.")

        tabela = iv.groupby(["nome_limpo"],as_index=False).agg(
            categoria=("cat_norm","first"),
            qtd=("qtd","sum"),
            receita=("valor_tot","sum"),
            cmv=("cmv_item","sum")
        )
        tabela["margem"] = tabela["receita"] - tabela["cmv"]
        tabela["margem_%"] = (tabela["margem"] / tabela["receita"] * 100).round(1)
        tabela = tabela.rename(columns={"nome_limpo":"produto"}).sort_values("cmv", ascending=False).reset_index(drop=True)
        st.dataframe(nomes_legiveis(tabela), use_container_width=True, hide_index=True)

        mask_sem_custo = iv["custo_unit"].isna() & ~mask_complemento
        diag_sem_custo = (iv.loc[mask_sem_custo, ["nome_prod","nome_limpo","cat_prod","qtd","valor_tot","valor_base"]]
                            .assign(ocorrencias=1)
                            .groupby(["nome_prod","nome_limpo","cat_prod"])
                            .agg(qtd_total=("qtd","sum"), valor_total=("valor_base","sum"), ocorrencias=("ocorrencias","sum"))
                            .reset_index()
                            .sort_values(["ocorrencias","valor_total"], ascending=[False, False]))
        if not diag_sem_custo.empty:
            st.divider()
            st.subheader("Produtos sem custo mapeado")
            st.dataframe(nomes_legiveis(diag_sem_custo.reset_index(drop=True)), use_container_width=True, hide_index=True)

            
with tab4:
    df_meta = carregar_primeira_aba_xlsx(arq_contas, None)
    if not carregou(df_meta):
        st.info("Carregue a planilha de Contas a Receber para visualizar a aba Metas.")
    else:
        df_meta = df_meta.copy()
        df_meta.columns = df_meta.columns.str.strip()
        df_meta = df_meta.rename(columns={
            "Cód. Pedido": "cod_pedido",
            "Valor Líq.": "valor_liq",
            "Forma Pagamento": "forma_pagamento",
            "Crédito": "data",
            "Total Pedido": "total_pedido"
        })
        df_meta["data"] = pd.to_datetime(df_meta["data"], errors="coerce")
        df_meta["valor_liq"] = pd.to_numeric(df_meta["valor_liq"], errors="coerce")
        df_meta = df_meta.dropna(subset=["data", "valor_liq", "cod_pedido"]).copy()

        if df_meta.empty:
            st.info("Não há dados válidos para análise de metas.")
        else:
            df_meta["semana"] = df_meta["data"] + pd.to_timedelta((6 - df_meta["data"].dt.weekday) % 7, unit="D")
            df_meta["semana"] = df_meta["semana"].dt.normalize()

            resumo_sem = (
                df_meta.groupby("semana", as_index=False)
                .agg(
                    receita=("valor_liq", "sum"),
                    pedidos=("cod_pedido", "nunique")
                )
                .sort_values("semana")
            )

            if resumo_sem.empty or len(resumo_sem) < 1:
                st.info("Dados insuficientes para análise semanal.")
            else:
                max_data = df_meta["data"].max().normalize()
                fim_semana_dados = max_data + pd.to_timedelta((6 - max_data.weekday()) % 7, unit="D")
                fim_semana_dados = fim_semana_dados.normalize()

                resumo_completo = resumo_sem[resumo_sem["semana"] < fim_semana_dados].copy()

                if resumo_completo.empty:
                    st.info("Ainda não há nenhuma semana completa fechada para análise de metas.")
                else:
                    semanas_ordenadas = resumo_completo["semana"].sort_values().unique()
                    semana_passada = semanas_ordenadas[-1]
                    linha_passada = resumo_completo[resumo_completo["semana"] == semana_passada]

                    receita_passada = float(linha_passada["receita"].iloc[0])
                    pedidos_passada = int(linha_passada["pedidos"].iloc[0])

                    diff_abs = None
                    diff_pct = None

                    if len(semanas_ordenadas) >= 2:
                        semana_retrasada = semanas_ordenadas[-2]
                        linha_retrasada = resumo_completo[resumo_completo["semana"] == semana_retrasada]
                        receita_retrasada = float(linha_retrasada["receita"].iloc[0])

                        diff_abs = receita_passada - receita_retrasada
                        if receita_retrasada != 0:
                            diff_pct = diff_abs / receita_retrasada * 100
                        else:
                            diff_pct = 0.0

                    c1, c2, c3 = st.columns(3)
                    c1.metric("Semana passada (R$)", br_money(receita_passada))
                    c2.metric("Pedidos semana passada", f"{pedidos_passada}")

                    if diff_abs is not None:
                        c3.metric(
                            "Variação vs semana retrasada",
                            br_money(diff_abs),
                            f"{diff_pct:,.1f}%"
                        )
                    else:
                        c3.metric("Variação vs semana retrasada", "-", "-")

                    if diff_abs is not None:
                        if diff_abs > 0:
                            st.caption(f"A semana passada faturou {br_money(diff_abs)} a mais (+{diff_pct:,.1f}%) que a semana retrasada.")
                        elif diff_abs < 0:
                            st.caption(f"A semana passada faturou {br_money(-diff_abs)} a menos ({diff_pct:,.1f}%) que a semana retrasada.")
                        else:
                            st.caption("Semana passada teve o mesmo faturamento da semana retrasada.")

                    st.divider()

                    st.subheader("Histórico recente de faturamento")

                    ultimas_sem = resumo_completo.sort_values("semana").tail(4).copy()
                    ultimas_sem["label"] = ultimas_sem["semana"].dt.strftime("%d/%m")

                    df_meta["mes"] = df_meta["data"].dt.to_period("M").dt.to_timestamp("M")
                    resumo_mes = (
                        df_meta.groupby("mes", as_index=False)
                        .agg(receita=("valor_liq", "sum"))
                        .sort_values("mes")
                    )
                    ultimos_mes = resumo_mes.tail(4).copy()
                    ultimos_mes["label"] = ultimos_mes["mes"].dt.strftime("%m/%Y")

                    col_sem, col_mes = st.columns(2)

                    with col_sem:
                        st.markdown("**Últimas 4 semanas completas**")
                        if len(ultimas_sem) == 0:
                            st.info("Sem semanas completas suficientes para exibir.")
                        else:
                            fig_sem = px.bar(
                                ultimas_sem,
                                y="label",
                                x="receita",
                                orientation="h",
                                labels={
                                    "label": "Semana (fim em domingo)",
                                    "receita": "Faturamento (R$)"
                                }
                            )
                            fig_sem = estilizar_fig(fig_sem)
                            st.plotly_chart(fig_sem, use_container_width=True, key="metas_barras_semanas")

                    with col_mes:
                        st.markdown("**Últimos 4 meses**")
                        if len(ultimos_mes) == 0:
                            st.info("Sem meses suficientes para exibir.")
                        else:
                            fig_mes = px.bar(
                                ultimos_mes,
                                y="label",
                                x="receita",
                                orientation="h",
                                labels={
                                    "label": "Mês",
                                    "receita": "Faturamento (R$)"
                                }
                            )
                            fig_mes = estilizar_fig(fig_mes)
                            st.plotly_chart(fig_mes, use_container_width=True, key="metas_barras_meses")


with tab5:
    itens_ads = carregar_primeira_aba_xlsx(arq_itens, None)

    if not carregou(itens_ads):
        st.info("Carregue a planilha de Itens Vendidos para visualizar a aba Promoções.")
    else:
        itens_ads = itens_ads.copy()
        itens_ads.columns = itens_ads.columns.str.strip()
        itens_ads = itens_ads.rename(columns={
            "Data/Hora Item": "data_item",
            "Qtd.": "qtd",
            "Valor. Tot. Item": "valor_tot",
            "Nome Prod": "nome_prod",
            "Cat. Prod.": "cat_prod",
            "Cod. Ped.": "cod_ped",
            "Valor Prod": "valor_prod"
        })

        itens_ads["data_item"] = pd.to_datetime(itens_ads["data_item"], errors="coerce")
        itens_ads["dia"] = itens_ads["data_item"].dt.date
        itens_ads["qtd"] = pd.to_numeric(itens_ads["qtd"], errors="coerce").fillna(0.0)
        itens_ads["valor_tot"] = pd.to_numeric(itens_ads["valor_tot"], errors="coerce").fillna(0.0)
        itens_ads["valor_prod"] = pd.to_numeric(itens_ads["valor_prod"], errors="coerce")
        itens_ads = itens_ads.dropna(subset=["dia"]).copy()

        if data_ini is not None and data_fim is not None:
            mask_itens_periodo = (itens_ads["dia"] >= data_ini) & (itens_ads["dia"] <= data_fim)
            itens_ads = itens_ads.loc[mask_itens_periodo].copy()

        def padronizar_nome_pizza(nome):
            if not isinstance(nome, str):
                return nome
            nome = nome.strip()
            if nome.startswith("PIZZA "):
                nome = nome.replace("PIZZA ", "", 1)
            nome = nome.replace(" Pequena", " P")
            nome = nome.replace(" Média", " M")
            nome = nome.replace(" Grande", " G")
            return nome.strip()

        itens_ads["nome_prod"] = itens_ads["nome_prod"].apply(padronizar_nome_pizza)
        itens_ads["nome_rodizio_chave"] = itens_ads["nome_prod"].apply(sem_acentos_upper)

        combo_config = {
            "pizzas_doce_p": [
                "BRIGADEIRO P",
                "TROPICAL P",
                "BRIGADEIRO BRANCO P",
                "PRESTIGIO P",
                "BANANA P",
                "PAÇOCA P"
            ],
            "refris_lata": [
                "COCA COLA LATA",
                "COCA COLA ZERO LATA",
                "FANTA LARANJA LATA",
                "FANTA UVA LATA"
                "SPRITE LATA"
                "SCHWEPPES CITRUS LATA"
                "GUARANÁ ANTARTICA LATA"
            ],
            "refris_2l": [
                "COCA COLA 2L",
                "GUARANÁ ANTARTICA 2L",
            ]
        }

        rodizio_nomes_todos_norm = [
            "RODIZIO DE PIZZA",
            "RODIZIO DE PIZZA ( SEXTA FEIRA )",
            "RODIZIO DIA DOS NAMORADOS",
            "RODIZIO PROMOCAO",
            "RODIZIO REFRI LIBERADO",
            "RODIZIO INFANTIL PROMO"
        ]

        rodizio_nomes_promo_norm = [
            "RODIZIO PROMOCAO",
            "RODIZIO REFRI LIBERADO",
            "RODIZIO INFANTIL PROMO"
        ]

        def marcar_itens_promocionais(df, limiar_desconto=0.8):
            df = df.copy()

            df["preco_unit_real"] = np.nan
            df["desconto_pct"] = 0.0
            df["eh_promocao"] = False
            df["preco_promocional"] = np.nan

            mask_pizza = df["Tipo de Item"] == "Produto por tamanho"
            sub = df.loc[mask_pizza].copy()

            sub["qtd"] = pd.to_numeric(sub["qtd"], errors="coerce")
            qtd_inteira = sub["qtd"].notna() & (sub["qtd"] % 1 == 0)
            multi_qtd = qtd_inteira & (sub["qtd"] > 1)

            preco_unit = pd.to_numeric(sub["Valor Un. Item"], errors="coerce")
            preco_unit.loc[multi_qtd] = sub.loc[multi_qtd, "valor_tot"] / sub.loc[multi_qtd, "qtd"]
            sub["preco_unit_real"] = preco_unit

            mask_base = sub["valor_prod"].notna() & qtd_inteira & sub["preco_unit_real"].notna()

            sub["desconto_pct"] = 0.0
            sub.loc[mask_base, "desconto_pct"] = (
                1 - sub.loc[mask_base, "preco_unit_real"] / sub.loc[mask_base, "valor_prod"]
            )

            sub["eh_promocao"] = False
            cond_promo = mask_base & (sub["preco_unit_real"] < sub["valor_prod"] * limiar_desconto)
            sub.loc[cond_promo, "eh_promocao"] = True
            sub["preco_promocional"] = np.where(sub["eh_promocao"], sub["preco_unit_real"], np.nan)

            df.loc[mask_pizza, "preco_unit_real"] = sub["preco_unit_real"]
            df.loc[mask_pizza, "desconto_pct"] = sub["desconto_pct"]
            df.loc[mask_pizza, "eh_promocao"] = sub["eh_promocao"]
            df.loc[mask_pizza, "preco_promocional"] = sub["preco_promocional"]

            return df

        def identificar_combos(df, combo_config):
            mask_combo = df["Tipo de Item"] == "Item de combo"
            df_combo = df[mask_combo].copy()
            grupos = []

            for cod_ped, g in df_combo.groupby("cod_ped"):
                nomes = g["nome_prod"]
                cats = g["cat_prod"]

                count_pizza_g = ((cats == "PIZZAS") & nomes.str.endswith(" G")).sum()
                count_pizza_p = ((cats == "PIZZAS") & nomes.str.endswith(" P")).sum()
                count_pizza_doce_p = nomes.isin(combo_config.get("pizzas_doce_p", [])).sum()
                count_refri_lata = nomes.isin(combo_config.get("refris_lata", [])).sum()
                count_refri_2l = nomes.isin(combo_config.get("refris_2l", [])).sum()

                combos = []

                n_refri2l = min(count_pizza_g, count_refri_2l)
                if n_refri2l > 0:
                    combos.append(("combo_refri_2l", n_refri2l))

                n_pizza_doce = min(count_pizza_g, count_pizza_doce_p)
                if n_pizza_doce > 0:
                    combos.append(("combo_pizza_doce", n_pizza_doce))

                n_individual = min(count_pizza_p, count_refri_lata)
                if n_individual > 0:
                    combos.append(("combo_individual", n_individual))

                for tipo, qtd_combo in combos:
                    grupos.append(
                        {
                            "cod_ped": cod_ped,
                            "tipo_combo": tipo,
                            "qtd_combo": int(qtd_combo)
                        }
                    )

            if not grupos:
                return pd.DataFrame(columns=["cod_ped", "tipo_combo", "qtd_combo"])

            return pd.DataFrame(grupos)

        def gerar_tabela_promocoes(df):
            prom = df[df["eh_promocao"]].copy()
            if prom.empty:
                return pd.DataFrame(
                    columns=[
                        "nome_prod",
                        "data_inicio",
                        "data_fim",
                        "Duracao",
                        "preço promoção",
                        "desconto_pct"
                    ]
                )

            prom["dia"] = pd.to_datetime(prom["dia"])
            linhas = []

            for nome, g in prom.groupby("nome_prod"):
                datas = sorted(g["dia"].dt.date.unique())
                if not datas:
                    continue

                run_start = datas[0]
                prev = datas[0]

                def add_linha(run_start_local, prev_local):
                    subset = g[g["dia"].dt.date.between(run_start_local, prev_local)]
                    preco_vals = subset["preco_promocional"].dropna()
                    desc_vals = subset["desconto_pct"].dropna()

                    preco_prom = float(preco_vals.iloc[0]) if not preco_vals.empty else None
                    desc_prom_pct = float(desc_vals.iloc[0]) * 100 if not desc_vals.empty else None

                    if run_start_local == prev_local:
                        duracao = run_start_local.strftime("%d-%m-%y")
                    else:
                        duracao = (
                            run_start_local.strftime("%d-%m-%y")
                            + " a "
                            + prev_local.strftime("%d-%m-%y")
                        )

                    linhas.append(
                        {
                            "nome_prod": nome,
                            "data_inicio": run_start_local,
                            "data_fim": prev_local,
                            "Duracao": duracao,
                            "preço promoção": preco_prom,
                            "desconto_pct": desc_prom_pct
                        }
                    )

                for d in datas[1:]:
                    if (d - prev).days == 1:
                        prev = d
                    else:
                        add_linha(run_start, prev)
                        run_start = d
                        prev = d

                add_linha(run_start, prev)

            return pd.DataFrame(linhas)

        def gerar_tabela_rodizio(df):
            mask_rod = df["nome_rodizio_chave"].isin(rodizio_nomes_todos_norm)
            rod = df[mask_rod].copy()
            if rod.empty:
                return pd.DataFrame(
                    columns=[
                        "dia",
                        "dia_semana",
                        "nome_prod",
                        "qtde",
                        "valor",
                        "faturamento"
                    ]
                )

            rod["valor_unit"] = pd.to_numeric(rod["Valor Un. Item"], errors="coerce")
            rod["dia"] = pd.to_datetime(rod["dia"])

            mapa_semana = {
                0: "Segunda",
                1: "Terça",
                2: "Quarta",
                3: "Quinta",
                4: "Sexta",
                5: "Sábado",
                6: "Domingo"
            }
            rod["dia_semana"] = rod["dia"].dt.dayofweek.map(mapa_semana)

            agg = (
                rod.groupby(["dia", "dia_semana", "nome_prod"], as_index=False)
                .agg(
                    qtde=("qtd", "sum"),
                    valor=("valor_unit", "max")
                )
            )
            agg["dia"] = agg["dia"].dt.date
            agg["faturamento"] = agg["qtde"] * agg["valor"]
            return agg

        def gerar_series_rodizio(df):
            mask_rod = df["nome_rodizio_chave"].isin(rodizio_nomes_todos_norm)
            rod = df[mask_rod].copy()
            if rod.empty:
                return pd.DataFrame(columns=["dia", "qtde_total", "is_promo_day"])

            rod["dia"] = pd.to_datetime(rod["dia"]).dt.date

            serie = (
                rod.groupby("dia", as_index=False)["qtd"]
                .sum()
                .rename(columns={"qtd": "qtde_total"})
            )

            dias_promo = (
                rod[rod["nome_rodizio_chave"].isin(rodizio_nomes_promo_norm)]["dia"]
                .dropna()
                .unique()
            )
            serie["is_promo_day"] = serie["dia"].isin(dias_promo)
            return serie

        def gerar_resumo_combos(df_combos):
            if df_combos.empty:
                return pd.DataFrame(columns=["tipo_combo", "qtd_total", "pedidos"])
            resumo = (
                df_combos.groupby("tipo_combo", as_index=False)
                .agg(
                    qtd_total=("qtd_combo", "sum"),
                    pedidos=("cod_ped", "nunique")
                )
            )
            return resumo

        def gerar_resumo_promos_pizza(df):
            prom = df[df["eh_promocao"]].copy()
            if prom.empty:
                return pd.DataFrame(
                    columns=[
                        "nome_prod",
                        "qtd_total",
                        "receita_promo",
                        "preço_medio_promo",
                        "desconto_medio_pct",
                        "dias_promo",
                        "qtd_media_dia",
                        "receita_media_dia",
                        "qtd_media_dia_normal"
                    ]
                )

            prom["dia"] = pd.to_datetime(prom["dia"]).dt.date

            resumo_promo = (
                prom.groupby("nome_prod", as_index=False)
                .agg(
                    qtd_total=("qtd", "sum"),
                    receita_promo=("valor_tot", "sum"),
                    preço_medio_promo=("preco_promocional", "mean"),
                    desconto_medio_pct=("desconto_pct", lambda x: x.mean() * 100),
                    dias_promo=("dia", "nunique")
                )
            )
            resumo_promo["qtd_media_dia"] = resumo_promo["qtd_total"] / resumo_promo["dias_promo"]
            resumo_promo["receita_media_dia"] = resumo_promo["receita_promo"] / resumo_promo["dias_promo"]

            base = df[(df["Tipo de Item"] == "Produto por tamanho") & (~df["eh_promocao"])].copy()
            if base.empty:
                resumo = resumo_promo.copy()
                resumo["qtd_media_dia_normal"] = np.nan
            else:
                base["dia"] = pd.to_datetime(base["dia"]).dt.date
                base_resumo = (
                    base.groupby("nome_prod", as_index=False)
                    .agg(
                        qtd_total_normal=("qtd", "sum"),
                        dias_normal=("dia", "nunique")
                    )
                )
                base_resumo["qtd_media_dia_normal"] = np.where(
                    base_resumo["dias_normal"] > 0,
                    base_resumo["qtd_total_normal"] / base_resumo["dias_normal"],
                    np.nan
                )
                resumo = resumo_promo.merge(
                    base_resumo[["nome_prod", "qtd_media_dia_normal"]],
                    on="nome_prod",
                    how="left"
                )

            resumo = resumo.sort_values("receita_media_dia", ascending=False).reset_index(drop=True)
            return resumo

        itens_ads = marcar_itens_promocionais(itens_ads)
        df_combos = identificar_combos(itens_ads, combo_config)
        tabela_promocoes = gerar_tabela_promocoes(itens_ads)
        tabela_rodizio = gerar_tabela_rodizio(itens_ads)
        serie_rodizio = gerar_series_rodizio(itens_ads)
        resumo_combos = gerar_resumo_combos(df_combos)
        resumo_promos_pizza = gerar_resumo_promos_pizza(itens_ads)

        st.subheader("Rodízio")

        if serie_rodizio.empty:
            st.info("Nenhum lançamento de rodízio encontrado no período selecionado.")
        else:
            serie_plot = serie_rodizio.copy()
            serie_plot["dia"] = pd.to_datetime(serie_plot["dia"])

            mapa_semana = {
                0: "Segunda",
                1: "Terça",
                2: "Quarta",
                3: "Quinta",
                4: "Sexta",
                5: "Sábado",
                6: "Domingo"
            }
            serie_plot["dia_semana"] = serie_plot["dia"].dt.dayofweek.map(mapa_semana)

            fig_rod = go.Figure()
            fig_rod.add_trace(
                go.Scatter(
                    x=serie_plot["dia"],
                    y=serie_plot["qtde_total"],
                    mode="lines+markers",
                    name="Rodízio - total de clientes",
                    customdata=serie_plot["dia_semana"],
                    hovertemplate="Qtde: %{y}<br>Dia: %{customdata}<extra></extra>"
                )
            )

            serie_promo = serie_plot[serie_plot["is_promo_day"]]
            if not serie_promo.empty:
                fig_rod.add_trace(
                    go.Scatter(
                        x=serie_promo["dia"],
                        y=serie_promo["qtde_total"],
                        mode="markers",
                        name="Dia de rodízio promocional",
                        marker=dict(size=10, symbol="circle-open"),
                        customdata=serie_promo["dia_semana"],
                        hovertemplate="Qtde: %{y}<br>Promo: %{customdata}<extra></extra>"
                    )
                )

            fig_rod = estilizar_fig(fig_rod)
            st.plotly_chart(fig_rod, use_container_width=True, key="rodizio_evolucao")

            serie_wd = serie_plot[serie_plot["dia_semana"].isin(["Quarta", "Quinta", "Sexta"])].copy()
            if not serie_wd.empty:
                serie_wd["tipo"] = np.where(serie_wd["is_promo_day"], "Promoção", "Normal")
                resumo_rodizio = (
                    serie_wd.groupby(["dia_semana", "tipo"], as_index=False)["qtde_total"]
                    .mean()
                    .rename(columns={"qtde_total": "media_qtde"})
                )

                fig_comp = px.bar(
                    resumo_rodizio,
                    x="dia_semana",
                    y="media_qtde",
                    color="tipo",
                    barmode="group",
                    labels={
                        "dia_semana": "Dia da semana",
                        "media_qtde": "Média de rodízios por dia",
                        "tipo": "Tipo de dia"
                    }
                )
                fig_comp = estilizar_fig(fig_comp)
                st.plotly_chart(fig_comp, use_container_width=True, key="rodizio_media_normal_vs_promo")
            else:
                st.info("Não há dados suficientes de rodízio em quarta, quinta e sexta para comparar dias normais e promocionais.")

            st.divider()
            st.subheader("Combos")

            if resumo_combos.empty:
                st.info("Nenhum combo detectado no período selecionado.")
            else:
                fig_combos = px.pie(
                    resumo_combos,
                    names="tipo_combo",
                    values="qtd_total",
                )
                fig_combos = estilizar_fig(fig_combos)
                st.plotly_chart(fig_combos, use_container_width=True, key="combos_pizza")

                st.dataframe(
                    nomes_legiveis(resumo_combos.reset_index(drop=True)),
                    use_container_width=True,
                    hide_index=True
                )

            st.divider()
            st.subheader("Promoções de pizzas")

            if resumo_promos_pizza.empty:
                st.info("Nenhuma promoção de pizza detectada no período selecionado.")
            else:
                fig_promos = px.bar(
                    resumo_promos_pizza,
                    x="nome_prod",
                    y="receita_media_dia",
                    labels={
                        "nome_prod": "Produto",
                        "receita_media_dia": "Receita média por dia em promoção (R$)"
                    }
                )
                fig_promos = estilizar_fig(fig_promos)
                st.plotly_chart(fig_promos, use_container_width=True, key="promo_barras_pizza")

                df_comp = resumo_promos_pizza.dropna(subset=["qtd_media_dia_normal"]).copy()
                if not df_comp.empty:
                    df_long = pd.melt(
                        df_comp,
                        id_vars=["nome_prod"],
                        value_vars=["qtd_media_dia_normal", "qtd_media_dia"],
                        var_name="tipo_dia",
                        value_name="qtd_media"
                    )
                    df_long["tipo_dia"] = df_long["tipo_dia"].map(
                        {
                            "qtd_media_dia_normal": "Dia normal",
                            "qtd_media_dia": "Dia com promoção"
                        }
                    )

                    fig_comp_pizza = px.bar(
                        df_long,
                        x="nome_prod",
                        y="qtd_media",
                        color="tipo_dia",
                        barmode="group",
                        labels={
                            "nome_prod": "Produto",
                            "qtd_media": "Média de pizzas por dia",
                            "tipo_dia": "Tipo de dia"
                        }
                    )
                    fig_comp_pizza = estilizar_fig(fig_comp_pizza)
                    st.plotly_chart(
                        fig_comp_pizza,
                        use_container_width=True,
                        key="promo_pizza_normal_vs_promo"
                    )
                else:
                    st.info("Não há vendas em dia normal suficientes para comparar com as promoções.")

                df_dias = itens_ads.copy()
                df_dias["dia"] = pd.to_datetime(df_dias["dia"], errors="coerce")
                df_dias = df_dias[df_dias["dia"].dt.dayofweek.isin([2, 3, 4, 5, 6])].copy()

                df_fat = (
                    df_dias.groupby("dia", as_index=False)["valor_tot"]
                    .sum()
                    .rename(columns={"valor_tot": "faturamento"})
                )

                dias_pizza_promo = (
                    df_dias[df_dias["eh_promocao"]]["dia"]
                    .dropna()
                    .unique()
                )

                mask_rod = df_dias["nome_rodizio_chave"].isin(rodizio_nomes_todos_norm)
                rod_dias = df_dias[mask_rod].copy()
                dias_rodizio_promo = (
                    rod_dias[rod_dias["nome_rodizio_chave"].isin(rodizio_nomes_promo_norm)]["dia"]
                    .dropna()
                    .unique()
                )

                dias_com_promo = set(list(dias_pizza_promo) + list(dias_rodizio_promo))

                df_fat["tipo_dia"] = np.where(df_fat["dia"].isin(dias_com_promo), "Com promoção", "Sem promoção")

                mapa_semana = {
                    0: "Segunda",
                    1: "Terça",
                    2: "Quarta",
                    3: "Quinta",
                    4: "Sexta",
                    5: "Sábado",
                    6: "Domingo"
                }
                df_fat["dow"] = df_fat["dia"].dt.dayofweek
                df_fat["dia_semana"] = df_fat["dow"].map(mapa_semana)

                df_fat = df_fat[df_fat["dow"].isin([2, 3, 4, 5, 6])].copy()

                resumo_fat = (
                    df_fat.groupby(["dia_semana", "tipo_dia"], as_index=False)["faturamento"]
                    .mean()
                    .rename(columns={"faturamento": "faturamento_medio"})
                )

                ordem = ["Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
                resumo_fat["dia_semana"] = pd.Categorical(resumo_fat["dia_semana"], categories=ordem, ordered=True)
                resumo_fat = resumo_fat.sort_values("dia_semana")

                if resumo_fat.empty:
                    st.info("Não há dados suficientes para calcular o faturamento médio por dia da semana.")
                else:
                    fig_fat = px.bar(
                        resumo_fat,
                        x="dia_semana",
                        y="faturamento_medio",
                        color="tipo_dia",
                        barmode="group",
                        labels={
                            "dia_semana": "Dia da semana",
                            "faturamento_medio": "Faturamento médio por dia (R$)",
                            "tipo_dia": "Tipo de dia"
                        }
                    )
                    fig_fat = estilizar_fig(fig_fat)
                    st.plotly_chart(
                        fig_fat,
                        use_container_width=True,
                        key="fat_medio_normal_vs_promo"
                    )

                st.markdown("**Ranking de promoções (ponderado por dia em promoção):**")
                st.dataframe(
                    nomes_legiveis(resumo_promos_pizza.reset_index(drop=True)),
                    use_container_width=True,
                    hide_index=True
                )
