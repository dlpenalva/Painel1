import streamlit as st
from pathlib import Path

st.set_page_config(page_title="TLB · cl8us - Apoio à Gestão de Contratos", layout="wide")


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

p1 = st.Page("pages/01_Calculo_Simples.py", title="Reajuste Simples", default=True)
p2 = st.Page("pages/02_Calculo_Represados.py", title="Reajustes Múltiplos")

nav = {
    "📌 Admissibilidade e Cálculo": [p1, p2]
}

grupo_valor_global = []

if (pages_dir / "03_Valor_Global.py").exists():
    grupo_valor_global.append(
        st.Page("pages/03_Valor_Global.py", title="Valor Global")
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

if (pages_dir / "08_Avaliacao_Aditivos.py").exists():
    grupo_gestao.append(
        st.Page("pages/08_Avaliacao_Aditivos.py", title="Avaliação de Aditivos")
    )

if grupo_gestao:
    nav["🛡️ Gestão Contratual"] = grupo_gestao

pg = st.navigation(nav)
pg.run()
