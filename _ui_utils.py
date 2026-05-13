import streamlit as st


def render_marca_topo():
    """Identidade visual própria do sistema, sem uso de logomarca institucional."""
    st.markdown(
        """
        <style>
        .tlb-cl8us-brand {
            display: inline-flex;
            flex-direction: column;
            gap: 1px;
            margin: 0 0 0.70rem 0;
            padding: 0;
        }
        .tlb-cl8us-brand-main {
            display: flex;
            align-items: baseline;
            gap: 0.45rem;
            line-height: 1.05;
            letter-spacing: -0.02em;
        }
        .tlb-cl8us-tlb { color: #123B63; font-size: 1.38rem; font-weight: 750; font-family: "Inter", "Segoe UI", Arial, sans-serif; }
        .tlb-cl8us-dot { color: #C0842B; font-size: 1.18rem; font-weight: 700; }
        .tlb-cl8us-name {
            color: #0F172A;
            font-size: 1.42rem;
            font-weight: 800;
            font-family: "Consolas", "SFMono-Regular", "Cascadia Mono", "Courier New", monospace;
            letter-spacing: -0.04em;
        }
        .tlb-cl8us-subtitle {
            color: #64748B;
            font-size: 0.74rem;
            font-weight: 500;
            font-family: "Inter", "Segoe UI", Arial, sans-serif;
            margin-top: 0.12rem;
            letter-spacing: 0.01em;
        }
        .tlb-cl8us-separator {
            width: 172px;
            border-bottom: 2px solid #1F4E79;
            opacity: 0.78;
            margin-top: 0.42rem;
        }
        .tlb-cl8us-aviso {
            display: block;
            margin-top: 0.42rem;
            color: #A16207;
            font-family: "Inter", "Segoe UI", Arial, sans-serif;
            font-size: 0.74rem;
            font-style: italic;
            font-weight: 400;
            line-height: 1.20;
        }
        </style>
        <div class="tlb-cl8us-brand" aria-label="TLB cl8us - apoio à gestão de contratos">
            <div class="tlb-cl8us-brand-main">
                <span class="tlb-cl8us-tlb">TLB</span>
                <span class="tlb-cl8us-dot">·</span>
                <span class="tlb-cl8us-name">cl8us</span>
            </div>
            <div class="tlb-cl8us-subtitle">apoio à gestão de contratos</div>
            <div class="tlb-cl8us-separator"></div>
            <div class="tlb-cl8us-aviso">AVISO: use apenas para docs não sigilosos e de livre acesso.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_versao_sidebar():
    """Exibe versão discreta no rodapé do menu lateral."""
    st.sidebar.markdown(
        """
        <style>
        .tlb-cl8us-sidebar-version {
            color: #94A3B8;
            font-size: 0.72rem;
            font-weight: 500;
            margin-top: 7.50rem;
            padding-top: 0;
            border-top: none;
            line-height: 1.2;
        }
        </style>
        <div class="tlb-cl8us-sidebar-version">v. 12/05/2026 - 21h</div>
        """,
        unsafe_allow_html=True,
    )


def render_indice_contrato_selectbox(key=None, index=0, options=None):
    """Renderiza o campo de índice com destaque visual consistente entre os fluxos."""
    if options is None:
        options = ["IST (Série Local)", "IPCA (433)", "IGP-M (189)"]

    with st.container(border=True):
        st.markdown(
            """
            <div style="
                background:#F5F3FF;
                border:1px solid #C4B5FD;
                border-radius:10px;
                padding:0.62rem 0.75rem;
                margin-bottom:0.55rem;
                color:#4C1D95;
            ">
                <div style="font-weight:800; font-size:0.98rem;">Índice do contrato</div>
                <div style="font-size:0.86rem; line-height:1.25rem;">
                    Revise este campo antes de prosseguir. Ele deve corresponder ao índice previsto no contrato.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        selecionado = st.selectbox(
            "Índice:",
            options,
            index=index,
            key=key,
            help="Confirme o índice contratual antes de gerar cálculos, arquivo de coleta ou relatórios.",
        )

        if selecionado == options[index]:
            st.caption(f"Conferência necessária: confirme se o índice contratual correto é **{selecionado}**.")
        else:
            st.caption(f"Índice selecionado para esta análise: **{selecionado}**.")

    return selecionado


