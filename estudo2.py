import streamlit as st
import pandas as pd
import datetime
import pytz
import base64
from supabase import create_client
import os
from dotenv import load_dotenv
from pathlib import Path

# =============================
# Carregar variáveis de ambiente
# =============================
env_path = Path(__file__).parent / "teste.env"  # Ajuste se necessário
load_dotenv(dotenv_path=env_path)

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# =============================
# Configurações iniciais
# =============================
TZ = pytz.timezone("America/Sao_Paulo")
itens = ["Etiqueta", "Tambor + Parafuso", "Solda", "Pintura", "Borracha ABS"]
usuarios = {"joao": "1234", "maria": "abcd", "admin": "admin"}

# =============================
# Funções do Supabase
# =============================

def carregar_checklists():
    response = supabase.table("checklists").select("*").execute()
    df = pd.DataFrame(response.data)
    if not df.empty:
        df["data_hora"] = pd.to_datetime(df["data_hora"])
        if df["data_hora"].dt.tz is None:
            df["data_hora"] = df["data_hora"].dt.tz_localize('UTC').dt.tz_convert(TZ)
        else:
            df["data_hora"] = df["data_hora"].dt.tz_convert(TZ)
    return df

def salvar_checklist(serie, resultados, usuario, foto_etiqueta=None, reinspecao=False):
    existe = supabase.table("checklists").select("numero_serie").eq("numero_serie", serie).execute()
    if not reinspecao and existe.data:
        st.error("⚠️ INVÁLIDO! DUPLICIDADE – Este Nº de Série já foi inspecionado.")
        return None

    reprovado = any(info['status'] == "Não Conforme" for info in resultados.values())
    data_hora = datetime.datetime.now(TZ)

    foto_base64 = None
    if foto_etiqueta is not None:
        try:
            foto_bytes = foto_etiqueta.getvalue()
            foto_base64 = base64.b64encode(foto_bytes).decode()
        except Exception as e:
            st.error(f"Erro ao processar a foto: {e}")
            foto_base64 = None

    for item, info in resultados.items():
        supabase.table("checklists").insert({
            "numero_serie": serie,
            "item": item,
            "status": info['status'],
            "observacoes": info['obs'],
            "inspetor": usuario,
            "data_hora": data_hora.isoformat(),
            "produto_reprovado": "Sim" if reprovado else "Não",
            "reinspecao": "Sim" if reinspecao else "Não",
            "foto_etiqueta": foto_base64 if item == "Etiqueta" else None
        }).execute()

    st.success(f"Checklist salvo no Supabase para o Nº de Série {serie}")
    return True

def carregar_apontamentos():
    response = supabase.table("apontamentos").select("*").execute()
    df = pd.DataFrame(response.data)
    if not df.empty:
        df["data_hora"] = pd.to_datetime(df["data_hora"])
        if df["data_hora"].dt.tz is None:
            df["data_hora"] = df["data_hora"].dt.tz_localize('UTC').dt.tz_convert(TZ)
        else:
            df["data_hora"] = df["data_hora"].dt.tz_convert(TZ)
    return df

def salvar_apontamento(serie):
    hoje = datetime.datetime.now(TZ).date()
    # Verificar se já existe apontamento hoje para esse código
    response = supabase.table("apontamentos")\
        .select("*")\
        .eq("numero_serie", serie)\
        .gte("data_hora", datetime.datetime.combine(hoje, datetime.time.min).isoformat())\
        .lte("data_hora", datetime.datetime.combine(hoje, datetime.time.max).isoformat())\
        .execute()

    if response.data:
        return False  # Já registrado hoje

    data_hora = datetime.datetime.now(TZ).isoformat()
    res = supabase.table("apontamentos").insert({
        "numero_serie": serie,
        "data_hora": data_hora
    }).execute()

    if res.data and not getattr(res, "error", None):
        return True
    else:
        st.error(f"Erro ao inserir apontamento: {getattr(res, 'error', 'Desconhecido')}")
        return False

# =============================
# Funções do App
# =============================

def login():
    st.sidebar.title("Login")
    if 'logado' not in st.session_state:
        st.session_state['logado'] = False
        st.session_state['usuario'] = None

    if not st.session_state['logado']:
        usuario = st.sidebar.text_input("Usuário")
        senha = st.sidebar.text_input("Senha", type="password")
        if st.sidebar.button("Entrar"):
            if usuario in usuarios and usuarios[usuario] == senha:
                st.session_state['logado'] = True
                st.session_state['usuario'] = usuario
                st.sidebar.success(f"Bem vindo, {usuario}!")
            else:
                st.sidebar.error("Usuário ou senha incorretos.")
        st.stop()
    else:
        st.sidebar.write(f"Logado como: {st.session_state['usuario']}")
        if st.sidebar.button("Sair"):
            st.session_state['logado'] = False
            st.session_state['usuario'] = None
            st.experimental_rerun()

def checklist_qualidade():
    st.markdown("## ✔️ Checklist de Qualidade")

    numero_serie = st.text_input("Número de Série")
    if not numero_serie:
        st.info("Por favor, informe o número de série para iniciar o checklist.")
        return

    resultados = {}
    for item in itens:
        st.markdown(f"### {item}")
        status = st.radio(f"Status - {item}", ["Conforme", "Não Conforme", "N/A"], key=f"{numero_serie}_{item}")
        obs = st.text_area(f"Observações - {item}", key=f"obs_{numero_serie}_{item}")
        resultados[item] = {"status": status, "obs": obs}

    foto_etiqueta = st.file_uploader("Foto da Etiqueta (Opcional)", type=["png", "jpg", "jpeg"])

    if st.button("Salvar Checklist"):
        if all(r['status'] == "N/A" for r in resultados.values()):
            st.error("Checklist inválido: todos os itens estão como N/A.")
            return
        salvar_checklist(numero_serie, resultados, st.session_state['usuario'], foto_etiqueta)
        st.experimental_rerun()

def reinspecao():
    df = carregar_checklists()

    if not df.empty:
        ultimos = df.sort_values("data_hora").groupby("numero_serie").tail(1)
        reprovados = ultimos[
            (ultimos["produto_reprovado"] == "Sim") & (ultimos["reinspecao"] == "Não")
        ]["numero_serie"].unique()
    else:
        reprovados = []

    if len(reprovados) > 0:
        st.markdown("## 🔄 Reinspeção de Produtos Reprovados")
        serie_sel = st.selectbox("Selecione o Nº de Série reprovado", reprovados)

        if serie_sel:
            resultados = {}
            for item in itens:
                st.markdown(f"### {item}")
                status = st.radio(f"Status - {item} (Reinspeção)", ["Conforme", "Não Conforme", "N/A"], key=f"re_{serie_sel}_{item}")
                obs = st.text_area(f"Observações - {item}", key=f"re_obs_{serie_sel}_{item}")
                resultados[item] = {"status": status, "obs": obs}

            if st.button("Salvar Reinspeção"):
                salvar_checklist(serie_sel, resultados, st.session_state['usuario'], reinspecao=True)
    else:
        st.info("Nenhum produto reprovado para reinspeção.")

def mostrar_historico():
    df = carregar_checklists()
    if df.empty:
        st.info("Nenhum checklist registrado ainda.")
        return

    st.markdown("## 📚 Histórico de Checklists")

    col1, col2 = st.columns(2)
    with col1:
        filtro_usuario = st.selectbox("Filtrar por Inspetor", ["Todos"] + sorted(df["inspetor"].unique()))
    with col2:
        filtro_status = st.selectbox("Filtrar por Produto Reprovado", ["Todos", "Sim", "Não"])

    df_filtrado = df.copy()
    if filtro_usuario != "Todos":
        df_filtrado = df_filtrado[df_filtrado["inspetor"] == filtro_usuario]
    if filtro_status != "Todos":
        df_filtrado = df_filtrado[df_filtrado["produto_reprovado"] == filtro_status]

    df_filtrado = df_filtrado.sort_values("data_hora", ascending=False)

    st.dataframe(df_filtrado[[
        "data_hora", "numero_serie", "item", "status", "observacoes",
        "inspetor", "produto_reprovado", "reinspecao"
    ]], height=400)

    fotos = df_filtrado[df_filtrado["foto_etiqueta"].notnull()]
    if not fotos.empty:
        st.markdown("### Fotos de Etiquetas")
        for idx, row in fotos.iterrows():
            foto_b64 = row["foto_etiqueta"]
            if foto_b64:
                img_bytes = base64.b64decode(foto_b64)
                st.image(img_bytes, caption=f"Etiqueta do Nº {row['numero_serie']}", width=300)

def painel_dashboard():

    def processar_codigo_barras():
        codigo_barras = st.session_state["codigo_barras"]
        if codigo_barras:
            sucesso = salvar_apontamento(codigo_barras.strip())
            if sucesso:
                st.success(f"Código {codigo_barras} registrado com sucesso!")
            else:
                st.warning(f"Código {codigo_barras} já registrado hoje ou erro.")
            st.session_state["codigo_barras"] = ""  # limpa o campo para próxima leitura

    st.markdown("# 📊 Painel de Apontamentos")

    # Entrada do código de barras com on_change
    st.text_input("Leia o Código de Barras aqui:", key="codigo_barras", on_change=processar_codigo_barras)

    # Carregar dados do dia para o painel
    hoje = datetime.datetime.now(TZ).date()
    df_apont = carregar_apontamentos()

    # Filtrar apontamentos de hoje
    if not df_apont.empty:
        df_hoje = df_apont[
            (df_apont["data_hora"] >= datetime.datetime.combine(hoje, datetime.time.min).replace(tzinfo=TZ)) &
            (df_apont["data_hora"] <= datetime.datetime.combine(hoje, datetime.time.max).replace(tzinfo=TZ))
        ]
    else:
        df_hoje = pd.DataFrame()

    total_lidos = len(df_hoje)

    # Para % aprovação, pegamos os checklists dos produtos lidos hoje e contamos aprovados
    df_checks = carregar_checklists()
    df_checks_hoje = df_checks[df_checks["numero_serie"].isin(df_hoje["numero_serie"].unique())]

    # Pegando última inspeção por número de série
    ultimos_checks = df_checks_hoje.sort_values("data_hora").groupby("numero_serie").tail(1)

    if not ultimos_checks.empty:
        aprovados = (ultimos_checks["produto_reprovado"] == "Não").sum()
        aprovacao_perc = (aprovados / len(ultimos_checks)) * 100
    else:
        aprovacao_perc = 0.0

    # Últimos 5 códigos lidos
    ultimos_5 = df_apont.sort_values("data_hora", ascending=False).head(5)

    # Layout com duas colunas: painel info e últimos 5
    col1, col2 = st.columns([3, 2])

    with col1:
        st.markdown("### Resumo do Dia")
        st.metric("Total de Produtos Lidos Hoje", total_lidos)
        st.metric("Percentual de Aprovação", f"{aprovacao_perc:.2f}%")

    with col2:
        st.markdown("### Últimos 5 Códigos Lidos")
        if ultimos_5.empty:
            st.write("Nenhum código registrado ainda.")
        else:
            for idx, row in ultimos_5.iterrows():
                dt = row["data_hora"].strftime("%d/%m/%Y %H:%M:%S")
                st.write(f"{dt} - {row['numero_serie']}")

def app():
    st.set_page_config(page_title="Controle de Qualidade", layout="wide")

    login()

    menu = st.sidebar.selectbox("Menu", ["Dashboard", "Checklist", "Reinspeção", "Histórico"])

    if menu == "Dashboard":
        painel_dashboard()
    elif menu == "Checklist":
        checklist_qualidade()
    elif menu == "Reinspeção":
        reinspecao()
    elif menu == "Histórico":
        mostrar_historico()

if __name__ == "__main__":
    app()

