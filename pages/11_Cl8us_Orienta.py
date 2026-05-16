import pandas as pd
import streamlit as st

from _ui_utils import render_marca_topo

st.set_page_config(page_title="TLB · cl8us - Cl8us Orienta", layout="wide")

render_marca_topo()

st.title("Cl8us Orienta")
st.caption("Página de apoio. Não calcula, não altera dados e não substitui a validação técnica.")

st.info(
    "Regra de uso: escolha o modo conforme a qualidade da informação disponível. "
    "Quando houver valores mensais por competência, prefira o Modo Padrão."
)

st.subheader("1. Qual modo usar?")

df_modos = pd.DataFrame(
    [
        {
            "Modo": "Padrão",
            "Quando usar": "Quando há itens remanescentes e valores mensais.",
            "Principal cuidado": "Conferir se os valores mensais representam corretamente a execução por competência.",
        },
        {
            "Modo": "Reduzido por Itens/Estoque",
            "Quando usar": "Quando há itens remanescentes sem valores mensais.",
            "Principal cuidado": "Tratar o resultado como estimativo, salvo validação formal da fiscalização.",
        },
        {
            "Modo": "Consumo por Itens/Ciclo",
            "Quando usar": "Quando há itens consumidos por ciclo sem valores mensais.",
            "Principal cuidado": "Conferir quantidades, ciclo de consumo, fator aplicado e critério de arredondamento.",
        },
    ]
)

st.dataframe(df_modos, use_container_width=True, hide_index=True)

st.subheader("2. Modo Consumo por Itens/Ciclo")

col1, col2 = st.columns(2)

with col1:
    st.markdown(
        """
        **Por que pode divergir da planilha da fiscalização?**

        Pode haver diferença por:
        - uso de valor unitário arredondado pela fiscalização;
        - uso de fator acumulado exato pelo sistema;
        - arredondamento em etapa diferente do cálculo;
        - diferença entre quantidade consumida e quantidade faturada;
        - corte temporal diferente entre os ciclos;
        - lançamento de item em ciclo diverso.
        """
    )

with col2:
    st.markdown(
        """
        **Fator acumulado exato x valor unitário arredondado**

        O sistema pode aplicar o fator acumulado com maior precisão matemática e arredondar o resultado financeiro final.
        Uma planilha externa pode arredondar primeiro o valor unitário e depois multiplicar pela quantidade.

        Essa diferença de sequência pode gerar divergências, especialmente em contratos com muitos itens.
        """
    )

st.warning(
    "O critério de arredondamento do Modo Consumo por Itens/Ciclo não deve ser alterado apenas para coincidir com planilha externa. "
    "Qualquer mudança deve ser decisão expressa, testada e documentada."
)

st.subheader("3. Quando tratar como estimativo ou validado?")

df_validacao = pd.DataFrame(
    [
        {
            "Situação": "Sem valores mensais e sem validação das quantidades por ciclo.",
            "Classificação": "Estimativo",
            "Linguagem sugerida": "Resultado estimativo, sujeito à validação da fiscalização e da base financeira, quando aplicável.",
        },
        {
            "Situação": "Fiscalização confirmou quantidades por ciclo, mas não há base mensal.",
            "Classificação": "Estimativo validado quanto às quantidades",
            "Linguagem sugerida": "Resultado apurado com base nas quantidades validadas pela fiscalização, sem substituir validação financeira.",
        },
        {
            "Situação": "Fiscalização confirma quantidades, ciclo, valores de referência e premissas.",
            "Classificação": "Validado pela fiscalização",
            "Linguagem sugerida": "Resultado validado pela fiscalização quanto às premissas informadas, preservadas as ressalvas do método.",
        },
    ]
)

st.dataframe(df_validacao, use_container_width=True, hide_index=True)

st.subheader("4. Índice de Serviços de Telecomunicações (IST)")

st.markdown(
    """
    O cl8us utiliza o arquivo local `ist.csv`, com as colunas `MES_ANO` e `INDICE_NIVEL`.
    O dado deve ser o **número-índice do IST**, e não apenas a variação percentual mensal.

    A Calculadora exibirá um aviso discreto quando o índice selecionado for **IST (Série Local)**,
    informando a última competência constante no `ist.csv`.

    **Por que podem aparecer 13 referências?**

    Para calcular a variação de 12 meses por número-índice, compara-se o índice final com o índice inicial de 12 meses antes.
    Assim, a série pode envolver 13 referências mensais quando contada de forma inclusiva.
    Isso não significa aplicar 13 meses de reajuste; significa usar ponto inicial e ponto final para apurar 12 meses de variação.
    """
)

st.markdown(
    """
    **Links de conferência:**

    - https://www.gov.br/anatel/pt-br/regulado/competicao/tarifas-e-precos/valores-do-ist
    - https://www.gov.br/anatel/pt-br/regulado/competicao/tarifas-e-precos/calculo-do-ist
    - https://teleco.com.br/tarifafixo2.asp
    """
)


st.subheader("4.1. ICTI (Ipeadata)")

st.markdown(
    """
    O **Índice de Custo da Tecnologia da Informação (ICTI)** foi incluído como índice de reajuste consultado diretamente pela internet, sem criação de arquivo local.

    **Fonte:** Ipeadata/Ipea.  
    **Série:** `DIMAC_ICTI2`.  
    **Natureza da série:** taxa mensal de variação, em `% a.m.`.  
    **Método de cálculo:** produtório das taxas mensais do período.

    Para fins operacionais, o ICTI deve ser tratado de forma semelhante ao IPCA e ao IGP-M: o sistema acumula as taxas mensais, e não divide diretamente números-índice oficiais como ocorre no IST local.

    **Regra de intervalo adotada no cl8us:**

    - usa o mês da data-base/proposta como primeira competência acumulada;
    - soma mais 11 competências mensais posteriores;
    - totaliza 12 competências mensais no ciclo;
    - o mês anterior aparece apenas como índice-base explicativo.

    **Exemplo:**

    ```text
    Data-base/proposta: agosto/2023
    Índice-base explicativo: julho/2023
    Competências acumuladas: agosto/2023 a julho/2024
    Método: produtório das taxas mensais do ICTI/Ipeadata
    ```

    **Atenção:** o ICTI depende de consulta online ao Ipeadata. Se a fonte estiver temporariamente indisponível, o cálculo pode não ser processado naquele momento.
    """
)


st.subheader("5. Regra de cautela documental")

st.markdown(
    """
    Nos modos sem valores mensais, evite afirmar que o resultado representa, de forma definitiva,
    o valor financeiro efetivamente devido por competência. A redação deve indicar a base utilizada:
    itens remanescentes, estoque ou consumo por ciclo.

    Quando o cálculo impactar pagamento retroativo, a validação financeira continua recomendável.
    """
)

st.success(
    "Esta página é apenas orientativa. Ela não altera cálculos, fórmulas, arredondamentos nem o Valor Total Atualizado do Contrato."
)
