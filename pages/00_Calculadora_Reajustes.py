import runpy
from pathlib import Path

import streamlit as st

st.set_page_config(page_title="TLB · cl8us - Calculadora de Reajustes", layout="wide")

from _ui_utils import render_marca_topo


def executar_motor(nome_arquivo: str):
    caminho = Path(__file__).resolve().parent / nome_arquivo
    if not caminho.exists():
        st.error(f"Arquivo do motor não localizado: {nome_arquivo}")
        st.stop()

    st.session_state["_calculadora_reajustes_embedded"] = True
    try:
        runpy.run_path(str(caminho), run_name="__main__")
    finally:
        st.session_state["_calculadora_reajustes_embedded"] = False


def _obter_fluxo_query():
    try:
        fluxo = st.query_params.get("fluxo", "")
        if isinstance(fluxo, list):
            fluxo = fluxo[0] if fluxo else ""
        return str(fluxo).lower()
    except Exception:
        try:
            fluxo = st.experimental_get_query_params().get("fluxo", [""])
            return str(fluxo[0]).lower() if fluxo else ""
        except Exception:
            return ""


def selecionar_fluxo(tipo):
    st.session_state["calculadora_tipo_analise"] = tipo
    st.session_state["calculadora_mais_de_um_ciclo"] = tipo == "Múltiplos ciclos"


render_marca_topo()

st.title("Calculadora de Reajustes")
st.caption("Entrada única para análise de ciclo único ou de múltiplos ciclos.")

st.markdown(
    """
    <style>
    .calc-note {
        background: #F8FAFC;
        border: 1px solid #E5EAF0;
        border-radius: 12px;
        padding: .75rem .95rem;
        color: #334155;
        margin: .55rem 0 1rem 0;
    }
    .calc-card-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 1.15rem;
        margin: .35rem 0 1.05rem 0;
        align-items: stretch;
    }
    .calc-card-link {
        display: block;
        height: 100%;
        min-height: 182px;
        border-radius: 14px;
        padding: 1rem 1.05rem;
        text-decoration: none !important;
        box-sizing: border-box;
        transition: all .12s ease-in-out;
    }
    .calc-card-link:hover {
        transform: translateY(-1px);
        text-decoration: none !important;
    }
    .calc-card-unico {
        border: 1.6px solid #93C5FD;
        background: #EFF6FF;
        color: #1E3A8A;
    }
    .calc-card-multiplos {
        border: 1.6px solid #F6C35B;
        background: #FFF7E6;
        color: #7A4A00;
    }
    .calc-card-selected-unico {
        border: 2.2px solid #2563EB;
        box-shadow: 0 0 0 2px rgba(37, 99, 235, .10);
    }
    .calc-card-selected-multiplos {
        border: 2.2px solid #D97706;
        box-shadow: 0 0 0 2px rgba(217, 119, 6, .10);
    }
    .calc-card-title {
        font-weight: 800;
        font-size: 1.16rem;
        margin-bottom: .40rem;
    }
    .calc-card-text {
        font-size: .93rem;
        line-height: 1.42rem;
    }
    .calc-selected {
        font-weight: 800;
        font-size: .82rem;
        margin-top: .70rem;
    }
    @media (max-width: 900px) {
        .calc-card-grid {
            grid-template-columns: 1fr;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

fluxo_query = _obter_fluxo_query()
if fluxo_query in ["unico", "único"]:
    selecionar_fluxo("Ciclo único")
elif fluxo_query in ["multiplos", "múltiplos", "multiplo", "múltiplo"]:
    selecionar_fluxo("Múltiplos ciclos")

tipo = st.session_state.get("calculadora_tipo_analise")

st.subheader("A análise envolve mais de um ciclo de reajuste?")

unico_classes = "calc-card-link calc-card-unico"
mult_classes = "calc-card-link calc-card-multiplos"

if tipo == "Ciclo único":
    unico_classes += " calc-card-selected-unico"
    unico_status = "Selecionado"
else:
    unico_status = "&nbsp;"

if tipo == "Múltiplos ciclos":
    mult_classes += " calc-card-selected-multiplos"
    mult_status = "Selecionado"
else:
    mult_status = "&nbsp;"

st.markdown(
    f"""
    <div class="calc-card-grid">
        <a class="{unico_classes}" href="?fluxo=unico" target="_self">
            <div class="calc-card-title">Único ciclo</div>
            <div class="calc-card-text">
                Use quando há apenas um ciclo de reajuste a calcular.
                Pedido tardio, acordo negocial, ciclo negativo ou histórico anterior podem ser tratados dentro do próprio fluxo, se existirem.
            </div>
            <div class="calc-selected">{unico_status}</div>
        </a>
        <a class="{mult_classes}" href="?fluxo=multiplos" target="_self">
            <div class="calc-card-title">Múltiplos ciclos</div>
            <div class="calc-card-text">
                Use quando há dois ou mais ciclos, ciclos acumulados, valores represados em mais de um período,
                recomposição de histórico por ciclo ou preclusões sucessivas.
            </div>
            <div class="calc-selected">{mult_status}</div>
        </a>
    </div>
    """,
    unsafe_allow_html=True,
)

if not tipo:
    st.markdown(
        """
        <div class="calc-note">
            <b>Para iniciar:</b> selecione uma das opções acima. A área de cálculo só será carregada após essa escolha.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

st.markdown(
    """
    <div class="calc-note">
        <b>Observação:</b> nesta etapa, basta definir se a análise é de ciclo único ou de múltiplos ciclos.
        As demais variáveis — pedido tardio, acordo negocial, ciclo negativo, preclusão, histórico anterior e aditivos/supressões —
        devem ser preenchidas ou avaliadas dentro do fluxo correspondente, quando seus dados forem necessários.
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("Como interpretar esta escolha", expanded=False):
    st.write(
        "Use **Único ciclo** quando houver apenas um ciclo a calcular. "
        "Use **Múltiplos ciclos** quando houver dois ou mais ciclos, ciclos acumulados ou necessidade de recompor histórico por ciclo. "
        "A existência de pedido tardio, acordo negocial, ciclo negativo ou ciclo anterior concedido não obriga, isoladamente, o uso de Múltiplos."
    )

st.divider()
st.subheader(f"Área de cálculo — {'Único ciclo' if tipo == 'Ciclo único' else 'Múltiplos ciclos'}")

if tipo == "Ciclo único":
    executar_motor("01_Calculo_Simples.py")
else:
    executar_motor("02_Calculo_Represados.py")
