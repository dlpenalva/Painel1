"""Identidade visual clara e institucional compartilhada por toda a aplicacao."""

from __future__ import annotations

import streamlit as st


def render_cl8us_light_theme() -> None:
    """Aplica a camada visual global sem alterar estrutura, logica ou botoes."""

    st.markdown(
        """
        <style>
        :root {
            color-scheme: light !important;
            --cl8us-bg: #F7F5EF;
            --cl8us-bg-soft: #F4F8FB;
            --cl8us-surface: #FFFFFF;
            --cl8us-surface-alt: #F8FAFC;
            --cl8us-sidebar-light: #CFE2F1;
            --cl8us-text: #18324A;
            --cl8us-text-soft: #52697D;
            --cl8us-blue: #1F5F8B;
            --cl8us-blue-soft: #E7F1F8;
            --cl8us-blue-hover: #D9EAF6;
            --cl8us-border: #CBD8E2;
            --cl8us-border-soft: #DFE7ED;
            --cl8us-focus: #2C6F9E;
            --cl8us-info-bg: #EAF4FC;
            --cl8us-info-border: #93C5FD;
            --cl8us-warning-bg: #FFF8E1;
            --cl8us-warning-border: #F0CC6A;
            --cl8us-alert-bg: #FFF1E8;
            --cl8us-alert-border: #FDBA74;
            --cl8us-success-bg: #EAF7EF;
            --cl8us-success-border: #A7D7B8;
        }

        html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"],
        [data-testid="stMain"], [data-testid="stMainBlockContainer"] {
            color-scheme: light !important;
        }

        body, [data-testid="stAppViewContainer"] {
            color: var(--cl8us-text) !important;
        }

        /* Superficies, cards e containers */
        [data-testid="stVerticalBlockBorderWrapper"],
        [data-testid="stForm"],
        [data-testid="stMetric"],
        [data-testid="stExpander"],
        [data-testid="stFileUploaderDropzone"],
        [data-testid="stDataFrame"],
        [data-testid="stDataEditor"],
        [data-testid="stTable"] {
            background-color: var(--cl8us-surface) !important;
            border-color: var(--cl8us-border) !important;
            color: var(--cl8us-text) !important;
        }

        [data-testid="stVerticalBlockBorderWrapper"],
        [data-testid="stForm"],
        [data-testid="stMetric"],
        [data-testid="stExpander"] {
            box-shadow: 0 5px 16px rgba(31, 78, 121, .055) !important;
        }

        [data-testid="stMetricLabel"],
        [data-testid="stMetricValue"],
        [data-testid="stMetricDelta"],
        [data-testid="stFileUploaderDropzoneInstructions"],
        [data-testid="stExpander"] summary,
        [data-testid="stExpander"] summary * {
            color: var(--cl8us-text) !important;
        }

        /* Campos: selectbox, multiselect, texto, data, numero e area de texto */
        [data-baseweb="select"] > div,
        [data-testid="stTextInput"] input,
        [data-testid="stNumberInput"] input,
        [data-testid="stDateInput"] input,
        [data-testid="stTimeInput"] input,
        [data-testid="stTextArea"] textarea,
        [data-testid="stChatInput"] textarea {
            background: var(--cl8us-surface) !important;
            border-color: var(--cl8us-border) !important;
            color: var(--cl8us-text) !important;
            -webkit-text-fill-color: var(--cl8us-text) !important;
            box-shadow: none !important;
        }

        [data-baseweb="select"] > div:hover,
        [data-testid="stTextInput"] input:hover,
        [data-testid="stNumberInput"] input:hover,
        [data-testid="stDateInput"] input:hover,
        [data-testid="stTimeInput"] input:hover,
        [data-testid="stTextArea"] textarea:hover {
            border-color: #9BB5C8 !important;
        }

        [data-baseweb="select"] > div:focus-within,
        [data-testid="stTextInput"] input:focus,
        [data-testid="stNumberInput"] input:focus,
        [data-testid="stDateInput"] input:focus,
        [data-testid="stTimeInput"] input:focus,
        [data-testid="stTextArea"] textarea:focus {
            border-color: var(--cl8us-focus) !important;
            box-shadow: 0 0 0 2px rgba(44, 111, 158, .15) !important;
            outline: none !important;
        }

        [data-baseweb="select"] input,
        [data-baseweb="select"] svg,
        [data-testid="stDateInput"] svg,
        [data-testid="stNumberInput"] svg {
            color: var(--cl8us-text) !important;
            fill: currentColor !important;
        }

        [data-baseweb="tag"] {
            background: var(--cl8us-blue-soft) !important;
            border: 1px solid #BCD3E3 !important;
            color: var(--cl8us-text) !important;
        }

        input::placeholder, textarea::placeholder {
            color: #72879A !important;
            opacity: 1 !important;
        }

        input:disabled, textarea:disabled,
        [aria-disabled="true"] [data-baseweb="select"] > div {
            background: #F1F4F7 !important;
            color: #687C8D !important;
            -webkit-text-fill-color: #687C8D !important;
            opacity: 1 !important;
        }

        /* Menus e calendarios abertos: portais BaseWeb ficam fora do bloco do campo. */
        [data-baseweb="popover"],
        [data-baseweb="popover"] > div,
        [data-baseweb="menu"],
        [data-baseweb="calendar"],
        ul[role="listbox"],
        [role="listbox"] {
            background: var(--cl8us-surface) !important;
            border-color: var(--cl8us-border) !important;
            color: var(--cl8us-text) !important;
            box-shadow: 0 12px 30px rgba(31, 78, 121, .14) !important;
        }

        [role="option"],
        [data-baseweb="menu"] li,
        [data-baseweb="calendar"] * {
            color: var(--cl8us-text) !important;
        }

        [role="option"] {
            background: var(--cl8us-surface) !important;
        }

        [role="option"]:hover,
        [role="option"][aria-selected="true"] {
            background: var(--cl8us-blue-hover) !important;
            color: #123B63 !important;
        }

        [data-baseweb="calendar"] {
            border: 1px solid var(--cl8us-border) !important;
        }

        [data-baseweb="calendar"] button {
            background: transparent !important;
            color: var(--cl8us-text) !important;
            border-color: transparent !important;
            box-shadow: none !important;
        }

        [data-baseweb="calendar"] button:hover {
            background: var(--cl8us-blue-hover) !important;
        }

        [data-baseweb="calendar"] [aria-selected="true"] {
            background: var(--cl8us-blue) !important;
            color: #FFFFFF !important;
        }

        /* Tabelas HTML, dataframes e editores de dados */
        [data-testid="stTable"] table,
        [data-testid="stMarkdownContainer"] table {
            background: var(--cl8us-surface) !important;
            border-collapse: collapse !important;
            color: var(--cl8us-text) !important;
        }

        [data-testid="stTable"] thead tr th,
        [data-testid="stMarkdownContainer"] table thead tr th,
        [data-testid="stDataFrame"] [role="columnheader"],
        [data-testid="stDataEditor"] [role="columnheader"] {
            background: #E6F0F7 !important;
            border-color: #C5D6E2 !important;
            color: #173B5D !important;
            font-weight: 700 !important;
        }

        [data-testid="stTable"] tbody tr:nth-child(even),
        [data-testid="stMarkdownContainer"] table tbody tr:nth-child(even),
        [data-testid="stDataFrame"] [role="row"]:nth-child(even),
        [data-testid="stDataEditor"] [role="row"]:nth-child(even) {
            background: #F7FAFC !important;
        }

        [data-testid="stTable"] td,
        [data-testid="stTable"] th,
        [data-testid="stMarkdownContainer"] table td,
        [data-testid="stMarkdownContainer"] table th,
        [data-testid="stDataFrame"] [role="gridcell"],
        [data-testid="stDataEditor"] [role="gridcell"] {
            border-color: var(--cl8us-border-soft) !important;
            color: var(--cl8us-text) !important;
        }

        [data-testid="stDataFrame"] canvas,
        [data-testid="stDataEditor"] canvas {
            color-scheme: light !important;
        }

        /* Abas e expansores */
        [data-testid="stTabs"] [role="tablist"] {
            background: #F3F7FA !important;
            border: 1px solid var(--cl8us-border-soft) !important;
        }

        [data-testid="stTabs"] [role="tab"] {
            color: var(--cl8us-text-soft) !important;
        }

        [data-testid="stTabs"] [role="tab"][aria-selected="true"] {
            background: var(--cl8us-surface) !important;
            color: var(--cl8us-blue) !important;
        }

        [data-testid="stExpander"] details,
        [data-testid="stExpander"] summary,
        [data-testid="stExpander"] [data-testid="stExpanderDetails"] {
            background: var(--cl8us-surface) !important;
            color: var(--cl8us-text) !important;
        }

        [data-testid="stExpander"] summary:hover {
            background: var(--cl8us-bg-soft) !important;
        }

        /* Mensagens por funcao sem fundos escuros. */
        [data-testid="stAlertContainer"] {
            border-radius: 10px !important;
            box-shadow: none !important;
        }

        [data-testid="stAlertContainer"]:has([data-testid="stAlertContentInfo"]) {
            background: var(--cl8us-info-bg) !important;
            border-color: var(--cl8us-info-border) !important;
            color: #1E4E79 !important;
        }

        [data-testid="stAlertContainer"]:has([data-testid="stAlertContentWarning"]) {
            background: var(--cl8us-warning-bg) !important;
            border-color: var(--cl8us-warning-border) !important;
            color: #6B4F00 !important;
        }

        [data-testid="stAlertContainer"]:has([data-testid="stAlertContentError"]) {
            background: var(--cl8us-alert-bg) !important;
            border-color: var(--cl8us-alert-border) !important;
            color: #8A3B12 !important;
        }

        [data-testid="stAlertContainer"]:has([data-testid="stAlertContentSuccess"]) {
            background: var(--cl8us-success-bg) !important;
            border-color: var(--cl8us-success-border) !important;
            color: #245C3A !important;
        }

        [data-testid="stAlertContainer"] *,
        [data-testid="stNotification"] * {
            color: inherit !important;
        }

        /* Sidebar: azul claro, campos brancos e destaques suaves. */
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #D8E8F4 0%, var(--cl8us-sidebar-light) 100%) !important;
            color: var(--cl8us-text) !important;
        }

        [data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"],
        [data-testid="stSidebar"] [data-baseweb="select"] > div,
        [data-testid="stSidebar"] input,
        [data-testid="stSidebar"] textarea {
            background: rgba(255, 255, 255, .94) !important;
            color: var(--cl8us-text) !important;
        }

        [data-testid="stSidebar"] [data-testid="stPageLink"] a:hover,
        [data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
            background: rgba(255, 255, 255, .46) !important;
        }

        /* Blocos de codigo, json e status tambem permanecem claros. */
        [data-testid="stCode"],
        [data-testid="stJson"],
        [data-testid="stStatusWidget"],
        pre, code {
            background: #F2F6F9 !important;
            color: #173B5D !important;
            border-color: var(--cl8us-border-soft) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
