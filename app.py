import streamlit as st
from pathlib import Path

st.set_page_config(page_title="FARC - Telebras", layout="wide")

pages_dir = Path("pages")

p1 = st.Page("pages/01_Calculo_Simples.py", title="Reajuste Simples", icon="⚖️", default=True)
p2 = st.Page("pages/02_Calculo_Represados.py", title="Reajustes Múltiplos", icon="🔄")

nav = {
    "Admissibilidade e Cálculo": [p1, p2]
}

grupo_valor_global = []

if (pages_dir / "03_Valor_Global.py").exists():
    grupo_valor_global.append(
        st.Page("pages/03_Valor_Global.py", title="Valor Global", icon="📊")
    )

if (pages_dir / "04_Relatorio_Global.py").exists():
    grupo_valor_global.append(
        st.Page("pages/04_Relatorio_Global.py", title="Relatório Global", icon="📝")
    )

if grupo_valor_global:
    nav["Valor Global e Relatórios"] = grupo_valor_global

pg = st.navigation(nav)
pg.run()
