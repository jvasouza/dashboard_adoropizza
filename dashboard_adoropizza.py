import re
from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata
from datetime import date
import calendar
import locale
import plotly.io as pio
import numpy as np

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

st.set_page_config(page_title="Dashboard - Adoro Pizza", layout="wide")
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
    anos = sorted(s.dt.year.unique())
    ano_sel = st.sidebar.selectbox("Ano para botões", anos, index=len(anos)-1, key="ano_btns")

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
    st.sidebar.caption(f"Filtrando: {dini.strftime('%d/%m/%Y')} → {dfim.strftime('%d/%m/%Y')}")
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
        st.divider()
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
            hoje = pd.to_datetime(date.today())
            fim_semana_atual = hoje + pd.to_timedelta((6 - hoje.weekday()) % 7, unit="D")
            fim_semana_atual = fim_semana_atual.normalize()

            resumo_passado = resumo_sem[resumo_sem["semana"] < fim_semana_atual].copy()

            if resumo_passado.empty:
                st.info("Ainda não há semanas fechadas para análise.")
            else:
                semanas_ordenadas = resumo_passado["semana"].sort_values().unique()
                semana_passada = semanas_ordenadas[-1]
                linha_passada = resumo_passado[resumo_passado["semana"] == semana_passada]

                receita_passada = float(linha_passada["receita"].iloc[0])
                pedidos_passada = int(linha_passada["pedidos"].iloc[0])

                diff_abs = None
                diff_pct = None

                if len(semanas_ordenadas) >= 2:
                    semana_retrasada = semanas_ordenadas[-2]
                    linha_retrasada = resumo_passado[resumo_passado["semana"] == semana_retrasada]
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

                st.subheader("Faturamento por Semana (histórico)")
                fig_sem = px.line(
                    resumo_passado.tail(12),
                    x="semana",
                    y="receita",
                    markers=True,
                    labels={"semana": "Fim da Semana (domingo)", "receita": "Receita (R$)"}
                )
                fig_sem = estilizar_fig(fig_sem)
                fig_sem.update_xaxes(tickformat="%d/%m/%Y")
                st.plotly_chart(fig_sem, use_container_width=True, key="metas_linha_semana")

                st.subheader("Resumo Semanal")
                df_sem_exibe = resumo_passado.tail(12).copy()
                df_sem_exibe = df_sem_exibe.rename(columns={"semana": "data"})
                st.dataframe(
                    nomes_legiveis(df_sem_exibe.reset_index(drop=True)),
                    use_container_width=True,
                    hide_index=True
                )

with tab5:
    df_ads = carregar_primeira_aba_xlsx(arq_ads_manager, None)
    itens_ads = carregar_primeira_aba_xlsx(arq_itens, None)
    df_contas_ads = carregar_primeira_aba_xlsx(arq_contas, None)

    if not carregou(df_ads):
        st.info("Carregue o relatório do Facebook Ads (relatorio-04-12-25.xlsx na pasta data) para visualizar a aba Promoções.")
    elif not carregou(itens_ads) or not carregou(df_contas_ads):
        st.info("Carregue as planilhas de Itens Vendidos e Contas a Receber para visualizar a aba Promoções.")
    else:
        df_ads = df_ads.copy()
        df_ads.columns = df_ads.columns.str.strip()
        df_ads = df_ads.rename(columns={
            "Nome do conjunto de anúncios": "adset",
            "Dia": "dia",
            "Valor usado (BRL)": "gasto",
            "Cliques no link": "cliques",
            "Visualizações da página de destino": "lp_views",
            "Alcance": "alcance",
            "Impressões": "impressoes",
            "CTR (taxa de cliques no link)": "ctr"
        })
        df_ads["dia"] = pd.to_datetime(df_ads["dia"], errors="coerce").dt.date
        df_ads["gasto"] = pd.to_numeric(df_ads["gasto"], errors="coerce").fillna(0.0)
        df_ads["cliques"] = pd.to_numeric(df_ads["cliques"], errors="coerce").fillna(0.0)
        df_ads["lp_views"] = pd.to_numeric(df_ads["lp_views"], errors="coerce").fillna(0.0)
        df_ads["alcance"] = pd.to_numeric(df_ads["alcance"], errors="coerce").fillna(0.0)
        df_ads["impressoes"] = pd.to_numeric(df_ads["impressoes"], errors="coerce").fillna(0.0)
        df_ads["ctr"] = pd.to_numeric(df_ads["ctr"], errors="coerce").fillna(0.0)
        df_ads = df_ads.dropna(subset=["dia", "adset"]).copy()

        if data_ini is not None and data_fim is not None:
            mask_ads_periodo = (df_ads["dia"] >= data_ini) & (df_ads["dia"] <= data_fim)
            df_ads = df_ads.loc[mask_ads_periodo].copy()

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

        itens_ads["produto_chave"] = itens_ads["nome_prod"].apply(sem_acentos_upper)
        df_ads["produto_chave"] = df_ads["adset"].apply(sem_acentos_upper)

        precos_normais = (
            itens_ads.dropna(subset=["valor_prod"])
            .groupby("produto_chave", as_index=False)["valor_prod"]
            .median()
            .rename(columns={"valor_prod": "preco_normal"})
        )
        itens_ads = itens_ads.merge(precos_normais, on="produto_chave", how="left")
        itens_ads["is_promo_price"] = (
            itens_ads["valor_prod"].notna()
            & itens_ads["preco_normal"].notna()
            & (itens_ads["valor_prod"] < itens_ads["preco_normal"] * 0.99)
        )

        vendas_prod = (
            itens_ads.groupby(["dia", "produto_chave"], as_index=False)
            .agg(
                qtd_vendida=("qtd", "sum"),
                receita=("valor_tot", "sum"),
                promo_price=("is_promo_price", "max")
            )
        )

        df_ads_merged = df_ads.merge(vendas_prod, on=["dia", "produto_chave"], how="left")

        df_contas_ads = df_contas_ads.copy()
        df_contas_ads.columns = df_contas_ads.columns.str.strip()
        df_contas_ads = df_contas_ads.rename(columns={
            "Cód. Pedido": "cod_pedido",
            "Valor Líq.": "valor_liq",
            "Crédito": "data"
        })
        df_contas_ads["data"] = pd.to_datetime(df_contas_ads["data"], errors="coerce")
        df_contas_ads["valor_liq"] = pd.to_numeric(df_contas_ads["valor_liq"], errors="coerce")
        df_contas_ads = df_contas_ads.dropna(subset=["data", "valor_liq"]).copy()

        if data_ini is not None and data_fim is not None:
            mask_receita_ads = (df_contas_ads["data"] >= pd.to_datetime(data_ini)) & (df_contas_ads["data"] <= pd.to_datetime(data_fim))
            df_contas_ads = df_contas_ads.loc[mask_receita_ads].copy()

        receita_total_periodo = float(df_contas_ads["valor_liq"].sum()) if not df_contas_ads.empty else 0.0

        df_sponsored = df_ads_merged[df_ads_merged["gasto"] > 0].copy()
        df_sponsored["receita"] = df_sponsored["receita"].fillna(0.0)

        gasto_total = float(df_sponsored["gasto"].sum()) if not df_sponsored.empty else 0.0
        receita_promo_sponsored = float(df_sponsored["receita"].sum()) if not df_sponsored.empty else 0.0
        pct_fat_sponsored = (receita_promo_sponsored / receita_total_periodo * 100.0) if receita_total_periodo else 0.0

        idx_ads = pd.MultiIndex.from_frame(
            df_ads.loc[df_ads["gasto"] > 0, ["dia", "produto_chave"]].drop_duplicates()
        )
        mask_sponsored_keys = vendas_prod.set_index(["dia", "produto_chave"]).index.isin(idx_ads)
        promo_organica = vendas_prod[(vendas_prod["promo_price"]) & (~mask_sponsored_keys)].copy()
        receita_promo_organica = float(promo_organica["receita"].sum()) if not promo_organica.empty else 0.0
        pct_fat_organica = (receita_promo_organica / receita_total_periodo * 100.0) if receita_total_periodo else 0.0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Gasto total em anúncios (R$)", br_money(gasto_total))
        c2.metric("Receita ligada a anúncios (R$)", br_money(receita_promo_sponsored))
        c3.metric("% fat. promo patrocinada", f"{pct_fat_sponsored:.1f}%")
        c4.metric("% fat. promo orgânica", f"{pct_fat_organica:.1f}%")

        st.divider()
        st.subheader("Resumo por produto anunciado (promoções patrocinadas)")

        if df_sponsored.empty:
            st.info("Nenhum gasto em anúncios no período selecionado.")
        else:
            resumo_prod = (
                df_sponsored.groupby("produto_chave", as_index=False)
                .agg(
                    gasto_total=("gasto", "sum"),
                    receita_promo=("receita", "sum"),
                    dias_ativos=("dia", "nunique"),
                    cliques=("cliques", "sum"),
                    lp_views=("lp_views", "sum")
                )
            )
            resumo_prod["roi"] = np.where(
                resumo_prod["gasto_total"] > 0,
                resumo_prod["receita_promo"] / resumo_prod["gasto_total"],
                np.nan
            )
            resumo_prod = resumo_prod.sort_values("receita_promo", ascending=False).reset_index(drop=True)

            fig_promos = px.bar(
                resumo_prod,
                x="produto_chave",
                y="receita_promo",
                labels={"produto_chave": "Produto", "receita_promo": "Receita ligada ao anúncio (R$)"}
            )
            fig_promos = estilizar_fig(fig_promos)
            st.plotly_chart(fig_promos, use_container_width=True, key="promo_barras_produto")

            df_exibe = resumo_prod.rename(columns={
                "produto_chave": "produto",
                "gasto_total": "gasto",
                "receita_promo": "receita",
                "dias_ativos": "dias",
                "cliques": "cliques",
                "lp_views": "lp_views",
                "roi": "roi"
            })
            st.dataframe(
                nomes_legiveis(df_exibe.reset_index(drop=True)),
                use_container_width=True,
                hide_index=True
            )

        st.divider()
        st.subheader("Promoções orgânicas (sem anúncio pago, preço abaixo do normal)")

        if promo_organica.empty:
            st.info("Nenhuma promoção orgânica detectada no período selecionado.")
        else:
            resumo_org = (
                promo_organica.groupby("produto_chave", as_index=False)
                .agg(
                    receita=("receita", "sum"),
                    dias=("dia", "nunique")
                )
                .sort_values("receita", ascending=False)
                .reset_index(drop=True)
            )
            df_org_exibe = resumo_org.rename(columns={"produto_chave": "produto"})
            st.dataframe(
                nomes_legiveis(df_org_exibe.reset_index(drop=True)),
                use_container_width=True,
                hide_index=True
            )
