import streamlit as st
from pathlib import Path
from _ui_utils import render_versao_sidebar

APP_SENHA_PUBLICO_VERSAO = "20260516_apos_nav"

st.set_page_config(page_title="TLB · cl8us - Apoio à Gestão de Contratos", layout="wide")


# ── Proteção por senha ──────────────────────────────────────────────────────
def _verificar_acesso():
    if st.session_state.get("_acesso_liberado"):
        return
    st.markdown("### TLB · cl8us")
    st.caption("apoio à gestão de contratos")
    senha = st.text_input("Senha de acesso", type="password", key="_senha_input")
    if st.button("Entrar", type="primary"):
        import hmac
        senha_correta = st.secrets.get("SENHA_APP", "")
        if hmac.compare_digest(senha.encode(), senha_correta.encode()):
            st.session_state["_acesso_liberado"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()

# ── Fim da proteção ─────────────────────────────────────────────────────────


# Ajuste visual do menu lateral: títulos em negrito e subitens sem negrito.
st.markdown(
    """
    <style>
    section[data-testid="stSidebar"] nav a,
    section[data-testid="stSidebar"] nav a *,
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a,
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a * {
        font-weight: 400 !important;
    }
    section[data-testid="stSidebar"] nav [role="heading"],
    section[data-testid="stSidebar"] nav [role="heading"] *,
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] [role="heading"],
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] [role="heading"] *,
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] div:not(:has(a)) > span,
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] div:not(:has(a)) > p {
        font-weight: 700 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

pages_dir = Path("pages")

p0 = st.Page("pages/00_Calculadora_Reajustes.py", title="Calculadora de Reajustes", default=True)

nav = {
    "📌 Admissibilidade e Cálculo": [p0]
}

grupo_valor_global = []

if (pages_dir / "03_Valor_Global.py").exists():
    grupo_valor_global.append(
        st.Page("pages/03_Valor_Global.py", title="Valores")
    )

if (pages_dir / "04_Relatorio_Global.py").exists():
    grupo_valor_global.append(
        st.Page("pages/04_Relatorio_Global.py", title="Relatórios")
    )

if (pages_dir / "06_Central_Arquivos.py").exists():
    grupo_valor_global.append(
        st.Page("pages/06_Central_Arquivos.py", title="Central de Arquivos")
    )

if (pages_dir / "07_Checklist_Processual.py").exists():
    grupo_valor_global.append(
        st.Page("pages/07_Checklist_Processual.py", title="Checklist Processual")
    )

if grupo_valor_global:
    nav["🌐 Visão Global e Relatórios"] = grupo_valor_global

grupo_gestao = []
if (pages_dir / "05_Garantia.py").exists():
    grupo_gestao.append(
        st.Page("pages/05_Garantia.py", title="Gestão da Garantia")
    )

if (pages_dir / "12_Adequacao_Orcamentaria.py").exists():
    grupo_gestao.append(
        st.Page("pages/12_Adequacao_Orcamentaria.py", title="Adequação Orçamentária")
    )

if (pages_dir / "08_Avaliacao_Aditivos.py").exists():
    grupo_gestao.append(
        st.Page("pages/08_Avaliacao_Aditivos.py", title="Aditivos: 25%")
    )

if grupo_gestao:
    nav["🛡️ Gestão Contratual"] = grupo_gestao

grupo_instrucao = []
if (pages_dir / "09_Infos_Previas.py").exists():
    grupo_instrucao.append(
        st.Page("pages/09_Infos_Previas.py", title="Infos Prévias")
    )

if (pages_dir / "10_Saneador.py").exists():
    grupo_instrucao.append(
        st.Page("pages/10_Saneador.py", title="Saneador")
    )

if grupo_instrucao:
    nav["🧾 Instrução Processual"] = grupo_instrucao

pg = st.navigation(nav)

# A navegação é criada antes da autenticação para evitar que o Streamlit exiba a página "app" no menu lateral.
# O conteúdo das páginas só é executado depois da liberação da sessão.
_verificar_acesso()

render_versao_sidebar()
pg.run()
