import streamlit as st
import base64
import json
import re
import pandas as pd
import requests
import io
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from dateutil.relativedelta import relativedelta

if not st.session_state.get("_calculadora_reajustes_embedded", False):
    st.set_page_config(page_title="Análises de Reajustes - Reajustes Múltiplos", layout="wide")


from _ui_utils import render_indice_contrato_selectbox, render_marca_topo
from _indice_utils import calcular_ist_numero_indice, coletar_sgs_produtorio
from _reajuste_utils import _competencias_mensais, _data_para_datetime, _formatar_data, _formatar_moeda_br, _formatar_moeda_br_md, _parse_moeda_br, _percentual_formatado

# ICTI/IPEADATA_LOCAL_FALLBACK_V1
ICTI_SERCODIGO_LOCAL = "DIMAC_ICTI2"
ICTI_API_BASES_LOCAL = [
    "https://www.ipeadata.gov.br/api/odata4",
    "http://www.ipeadata.gov.br/api/odata4",
]
MESES_PT_ABREV_ICTI = {1:"jan",2:"fev",3:"mar",4:"abr",5:"mai",6:"jun",7:"jul",8:"ago",9:"set",10:"out",11:"nov",12:"dez"}

def _icti_ipeadata_get_json_local(endpoint, timeout=20):
    ultimo_erro = None
    headers = {"User-Agent": "Mozilla/5.0 cl8us-icti", "Accept": "application/json"}
    for base_url in ICTI_API_BASES_LOCAL:
        try:
            resposta = requests.get(f"{base_url}/{endpoint}", headers=headers, timeout=timeout)
            resposta.raise_for_status()
            return resposta.json()
        except Exception as exc:
            ultimo_erro = exc
    raise RuntimeError(f"Não foi possível consultar o Ipeadata. Último erro: {ultimo_erro}")

@st.cache_data(ttl=60 * 60)
def _carregar_icti_ipeadata_local(timeout=20):
    dados = _icti_ipeadata_get_json_local(f"ValoresSerie(SERCODIGO='{ICTI_SERCODIGO_LOCAL}')", timeout=timeout)
    registros = dados.get("value", []) if isinstance(dados, dict) else []
    if not registros:
        raise RuntimeError("A API do Ipeadata retornou a série ICTI vazia.")
    linhas = []
    for item in registros:
        if not isinstance(item, dict):
            continue
        data = pd.to_datetime(item.get("VALDATA"), errors="coerce")
        valor = pd.to_numeric(item.get("VALVALOR"), errors="coerce")
        if pd.isna(data) or pd.isna(valor):
            continue
        data = pd.Timestamp(year=int(data.year), month=int(data.month), day=1).normalize()
        linhas.append({
            "data": data,
            "mes_ano": f"{MESES_PT_ABREV_ICTI[data.month]}/{str(data.year)[-2:]}",
            "taxa_mensal_percentual": float(valor),
        })
    df = pd.DataFrame(linhas)
    if df.empty:
        raise RuntimeError("Nenhuma competência válida do ICTI foi identificada no Ipeadata.")
    df = df.sort_values("data").drop_duplicates(subset=["data"], keep="last").reset_index(drop=True)
    df["fator_mensal"] = 1 + df["taxa_mensal_percentual"] / 100
    df["indice_nivel_sintetico"] = 100.0 * df["fator_mensal"].cumprod()
    return df

def calcular_icti_ipeadata(data_inicio, data_fim=None, timeout=20):
    if data_inicio is None:
        return None
    data_inicio_ts = pd.Timestamp(data_inicio)
    competencia_proposta = pd.Timestamp(data_inicio_ts.year, data_inicio_ts.month, 1).normalize()
    competencia_base = (competencia_proposta - relativedelta(months=1)).normalize()
    data_fim_ts = data_inicio_ts + relativedelta(months=11) if data_fim is None else pd.Timestamp(data_fim)
    competencia_final = pd.Timestamp(data_fim_ts.year, data_fim_ts.month, 1).normalize()
    if competencia_final < competencia_proposta:
        return None
    df = _carregar_icti_ipeadata_local(timeout=timeout)
    datas = set(df["data"])
    if competencia_base not in datas or competencia_final not in datas:
        return None
    periodo = df[(df["data"] > competencia_base) & (df["data"] <= competencia_final)].copy()
    if periodo.empty:
        return None
    fator = float(periodo["fator_mensal"].prod())
    variacao = fator - 1
    linha_base = df[df["data"] == competencia_base].iloc[0]
    linha_final = df[df["data"] == competencia_final].iloc[0]
    periodo["fator_acumulado_progressivo"] = periodo["fator_mensal"].cumprod()
    dados = periodo[["data", "taxa_mensal_percentual", "fator_mensal", "fator_acumulado_progressivo"]].copy()
    dados = dados.rename(columns={"taxa_mensal_percentual": "valor"})
    return {
        "variacao": variacao,
        "var": variacao,
        "i_ini": float(linha_base["indice_nivel_sintetico"]),
        "i_fim": float(linha_final["indice_nivel_sintetico"]),
        "d_ini": competencia_base,
        "d_fim": competencia_final,
        "p_ini": competencia_base,
        "p_fim": competencia_final,
        "competencia_proposta": competencia_proposta,
        "competencia_indice_base": competencia_base,
        "competencia_final": competencia_final,
        "d_proposta_ancora": competencia_proposta,
        "d_indice_base": competencia_base,
        "d_final_icti": competencia_final,
        "metodo": "ICTI/Ipeadata: produtório das taxas mensais; índice-base = mês anterior à proposta/âncora",
        "dados": dados,
        "sercodigo": ICTI_SERCODIGO_LOCAL,
        "serie": ICTI_SERCODIGO_LOCAL,
    }



def _data_para_date_segura(valor):
    """Converte date/datetime/Timestamp em date para validações de corte temporal."""
    try:
        if isinstance(valor, datetime):
            return valor.date()
    except Exception:
        pass
    try:
        if hasattr(valor, "date") and not isinstance(valor, str):
            return valor.date()
    except Exception:
        pass
    try:
        dt = pd.to_datetime(valor, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass
    return None


def _competencias_esperadas_indice(data_inicio, data_fim):
    """Lista competências mensais esperadas entre data_inicio e data_fim, inclusive."""
    try:
        inicio = pd.Timestamp(data_inicio).to_period("M")
        fim = pd.Timestamp(data_fim).to_period("M")
        if fim < inicio:
            return []
        return [p.strftime("%m/%Y") for p in pd.period_range(inicio, fim, freq="M")]
    except Exception:
        return []


def _competencia_de_valor(valor):
    """Extrai competência mm/aaaa de datas ou textos comuns."""
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except Exception:
        pass
    if isinstance(valor, (datetime, pd.Timestamp)):
        return pd.Timestamp(valor).to_period("M").strftime("%m/%Y")
    texto = str(valor).strip()
    if not texto:
        return None

    # dd/mm/aaaa
    m = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", texto)
    if m:
        dia, mes, ano = m.groups()
        try:
            return pd.Timestamp(int(ano), int(mes), 1).to_period("M").strftime("%m/%Y")
        except Exception:
            return None

    # mm/aaaa
    m = re.search(r"\b(\d{1,2})/(\d{4})\b", texto)
    if m:
        mes, ano = m.groups()
        try:
            return pd.Timestamp(int(ano), int(mes), 1).to_period("M").strftime("%m/%Y")
        except Exception:
            return None

    # aaaa-mm-dd ou aaaa-mm
    m = re.search(r"\b(\d{4})-(\d{1,2})(?:-\d{1,2})?\b", texto)
    if m:
        ano, mes = m.groups()
        try:
            return pd.Timestamp(int(ano), int(mes), 1).to_period("M").strftime("%m/%Y")
        except Exception:
            return None

    try:
        dt = pd.to_datetime(texto, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return pd.Timestamp(dt).to_period("M").strftime("%m/%Y")
    except Exception:
        pass
    return None


def _competencias_encontradas_no_resultado(res):
    """Extrai competências existentes no retorno do índice, especialmente para IPCA/IGP-M."""
    comps = set()
    if not res:
        return comps

    dados = res.get("dados") if isinstance(res, dict) else None

    def adicionar_de_obj(obj):
        if isinstance(obj, dict):
            for v in obj.values():
                adicionar_de_obj(v)
        elif isinstance(obj, (list, tuple, set)):
            for v in obj:
                adicionar_de_obj(v)
        else:
            comp = _competencia_de_valor(obj)
            if comp:
                comps.add(comp)

    if isinstance(dados, pd.DataFrame):
        # Dá preferência a colunas com nomes de data/competência.
        colunas_prioritarias = [
            c for c in dados.columns
            if str(c).strip().lower() in ["data", "dt", "competencia", "competência", "mes", "mês", "periodo", "período"]
        ]
        colunas = colunas_prioritarias or list(dados.columns)
        for col in colunas:
            for valor in dados[col].dropna().tolist():
                comp = _competencia_de_valor(valor)
                if comp:
                    comps.add(comp)
    elif dados is not None:
        adicionar_de_obj(dados)

    return comps


def _validar_indice_disponivel(res, data_inicio, data_fim, indice_nome):
    """Valida se o índice possui base suficiente para processar o ciclo.

    Para IPCA/IGP-M, exige todas as competências mensais do intervalo.
    Para IST, mantém a validação por retorno existente, pois o cálculo usa número-índice.
    """
    esperadas = _competencias_esperadas_indice(data_inicio, data_fim)

    if not res:
        return {
            "ok": False,
            "motivo": "sem_retorno",
            "esperadas": esperadas,
            "encontradas": [],
            "faltantes": esperadas,
        }

    if "IST" in str(indice_nome).upper():
        return {
            "ok": True,
            "motivo": "",
            "esperadas": esperadas,
            "encontradas": esperadas,
            "faltantes": [],
        }

    encontradas = sorted(_competencias_encontradas_no_resultado(res))
    faltantes = [c for c in esperadas if c not in set(encontradas)]

    return {
        "ok": len(faltantes) == 0,
        "motivo": "competencias_ausentes" if faltantes else "",
        "esperadas": esperadas,
        "encontradas": encontradas,
        "faltantes": faltantes,
    }


def _render_alerta_indice_ausente(validacao, ciclo_label="C1"):
    faltantes = validacao.get("faltantes", []) or []
    esperadas = validacao.get("esperadas", []) or []
    st.error(
        f"Processamento inviável: falta pelo menos um mês do intervalo de apuração do índice no {ciclo_label}."
    )
    st.warning(
        "Não foi possível concluir a apuração porque há competência ausente no intervalo do índice. "
        "Atualize a base de índices ou confira o período de apuração antes de prosseguir."
    )
    dados = []
    if faltantes:
        dados.append({"Item": "Competências faltantes", "Competências": ", ".join(faltantes)})
    if esperadas:
        dados.append({"Item": "Intervalo esperado", "Competências": ", ".join(esperadas)})
    if dados:
        st.dataframe(pd.DataFrame(dados), use_container_width=True, hide_index=True)

def get_data_rep(serie, d_ini, d_fim, is_ist, is_icti=False):
    try:
        if is_ist:
            resultado = calcular_ist_numero_indice(d_ini)
            if not resultado:
                r_ini = pd.Timestamp(d_ini.year, d_ini.month, 1).normalize()
                r_fim = pd.Timestamp((d_ini + relativedelta(years=1)).year, (d_ini + relativedelta(years=1)).month, 1).normalize()
                st.error(f"Dados do IST não encontrados para o período {r_ini.strftime('%m/%Y')} ou {r_fim.strftime('%m/%Y')}")
                return None
            return {
                "var": resultado["variacao"],
                "i_ini": resultado["i_ini"],
                "i_fim": resultado["i_fim"],
                "p_ini": resultado["d_ini"],
                "p_fim": resultado["d_fim"],
                "metodo": "Divisão de Número-Índice (IST)",
                "dados": resultado["dados"],
            }

        if is_icti:
            resultado = calcular_icti_ipeadata(d_ini, d_fim, timeout=15)
            if not resultado:
                st.error("Dados do ICTI/Ipeadata não encontrados para o intervalo do ciclo.")
                return None
            return {
                "var": resultado["variacao"],
                "i_ini": resultado["i_ini"],
                "i_fim": resultado["i_fim"],
                "p_ini": resultado["d_ini"],
                "p_fim": resultado["d_fim"],
                "d_indice_base": resultado.get("d_indice_base"),
                "metodo": resultado["metodo"],
                "dados": resultado["dados"],
            }

        resultado = coletar_sgs_produtorio(serie, d_ini, d_fim, timeout=10)
        if not resultado:
            return None
        return {
            "var": resultado["variacao"],
            "metodo": resultado["metodo"],
            "p_ini": d_ini,
            "p_fim": d_fim,
            "dados": resultado["dados"],
        }
    except Exception as e:
        st.error(f"Erro técnico na coleta de dados: {str(e)}")
        return None


def _render_equacao_ist(res_c):
    equacao_html = f"""
    <div style=\"background:#F4F6F8;border:1px solid #E1E6EB;border-radius:10px;padding:14px 18px;margin-top:10px;\">
        <div style=\"font-family:Consolas, Monaco, monospace;font-size:1.15rem;line-height:1.8;color:#334155;\">
            <span style=\"color:#0F766E;\">({res_c['i_fim']:.3f}</span>
            <span style=\"color:#94A3B8;\"> / </span>
            <span style=\"color:#0F766E;\">{res_c['i_ini']:.3f}</span>
            <span style=\"color:#94A3B8;\">) - 1 = </span>
            <span style=\"color:#B45309;font-weight:600;\">{res_c['var']*100:.4f}%</span>
        </div>
    </div>
    """
    st.markdown(equacao_html, unsafe_allow_html=True)


def _render_card_contexto_contrato():
    """Exibe cabeçalho executivo do bloco de contexto do contrato."""
    st.markdown(
        """
        <div style="background:#F3F6FA;border:1px solid #D9E2EC;border-left:5px solid #1F4E78;border-radius:12px;padding:14px 16px;margin:10px 0 8px 0;">
            <div style="color:#123B63;font-weight:800;font-size:1.02rem;margin-bottom:4px;">Contexto do Contrato</div>
            <div style="color:#475569;font-size:0.92rem;line-height:1.45;">
                Use este bloco quando o contrato já possuir reajustes, repactuações, aditivos ou supressões formalizados antes desta análise.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def _ciclo_para_numero(valor):
    texto = str(valor or "").strip().upper()
    if texto.startswith("C"):
        try:
            return int(texto[1:].split()[0])
        except Exception:
            return 0
    return 0

def _primeiro_ciclo_analise(contexto):
    return max(_ciclo_para_numero((contexto or {}).get('ultimo_ciclo_concedido', '')) + 1, 1)

def _data_contexto_para_datetime(valor):
    if not valor:
        return None
    try:
        dt = pd.to_datetime(valor, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            return dt.to_pydatetime()
    except Exception:
        pass
    return None

def _data_base_inicial_pelo_contexto(contexto, fallback):
    contexto = contexto or {}
    ultimo = _ciclo_para_numero(contexto.get('ultimo_ciclo_concedido', ''))
    if ultimo > 0:
        # Regra adotada: no fluxo ordinário, a data do pedido do último ciclo concedido
        # corresponde ao início dos efeitos financeiros desse reajuste.
        # Mantém-se fallback legado para arquivos antigos que ainda tenham o campo de efeitos.
        data_pedido_ultimo = _data_contexto_para_datetime(contexto.get('data_pedido_ultimo_ciclo'))
        if data_pedido_ultimo:
            return data_pedido_ultimo
        data_efeitos_legado = _data_contexto_para_datetime(contexto.get('data_efeitos_ultimo_reajuste'))
        if data_efeitos_legado:
            return data_efeitos_legado
    return fallback


def _evento_historico_principal_contexto(contexto):
    """Monta registro automático do último ciclo formalizado para exportação.

    Esse registro evita que o ciclo histórico pareça ter sumido do XLS quando
    a aba CICLOS contém apenas o objeto da análise atual.
    """
    contexto = contexto or {}
    ciclo = str(contexto.get('ultimo_ciclo_concedido', '') or '').strip()
    if not ciclo or ciclo.upper().startswith('C0') or ciclo.lower() in ['nenhum', 'c0 / nenhum']:
        return None
    data_base = str(contexto.get('data_base_ultimo_ciclo', '') or '').strip()
    data_pedido = str(contexto.get('data_pedido_ultimo_ciclo', '') or '').strip()
    valor_formalizado = contexto.get('valor_formalizado_anterior', '')
    try:
        valor_formatado = _formatar_moeda_br(float(valor_formalizado or 0.0)) if float(valor_formalizado or 0.0) > 0 else ''
    except Exception:
        valor_formatado = str(contexto.get('valor_formalizado_anterior_texto', '') or '').strip()
    percentual = contexto.get('percentual_ja_aplicado_pct', 0.0)
    try:
        percentual_txt = f"{float(percentual):.2f}%".replace('.', ',')
    except Exception:
        percentual_txt = ''
    ref = str(contexto.get('referencia_documental_historico', '') or '').strip()
    obs_base = str(contexto.get('observacao_historico', '') or '').strip()
    obs_partes = []
    if data_base:
        obs_partes.append(f"Data-base: {data_base}")
    if data_pedido:
        obs_partes.append(f"Data do pedido/efeitos financeiros: {data_pedido}")
    if percentual_txt and percentual_txt != '0,00%':
        obs_partes.append(f"Percentual aplicado: {percentual_txt}")
    if ref:
        obs_partes.append(f"Referência: {ref}")
    if obs_base:
        obs_partes.append(obs_base)
    obs_partes.append("Registro gerado automaticamente pelo Contexto do Contrato.")
    return {
        'Tipo de evento': 'Reajuste',
        'Ciclo': ciclo,
        'Data': data_pedido or data_base,
        'Valor formalizado/impacto': valor_formatado,
        'Incorporado ao valor formalizado?': 'Sim',
        'Observação': '; '.join(obs_partes),
    }


def _eventos_historicos_para_exportacao(contexto):
    """Combina o evento principal do contexto com eventos adicionais informados manualmente."""
    contexto = contexto or {}
    eventos = []
    auto = _evento_historico_principal_contexto(contexto)
    if auto:
        eventos.append(auto)
    for ev in contexto.get('eventos_historicos_anteriores') or []:
        if not isinstance(ev, dict):
            continue
        # Evita duplicação do registro automático em caso de sessões antigas.
        obs = str(ev.get('Observação', '') or '')
        if 'Registro gerado automaticamente pelo Contexto do Contrato' in obs:
            continue
        possui_conteudo = any(str(ev.get(k, '') or '').strip() for k in ['Tipo de evento', 'Data', 'Valor formalizado/impacto', 'Observação'])
        if possui_conteudo:
            eventos.append(ev)
    return eventos

def _render_contexto_contratual_anterior():
    """Coleta contexto do contrato para uso pelo Valor Global.

    Este bloco é opcional. Quando preenchido, registra memória formal anterior
    para relatórios e governança, sem alterar automaticamente o Valor Total Atualizado.
    """
    contexto_salvo = st.session_state.get('contexto_contratual_anterior', {}) or {}
    _render_card_contexto_contrato()
    with st.expander('Preencher/editar Contexto do Contrato', expanded=False):
        st.caption(
            'Preencha apenas quando o contrato já possuir eventos formalizados antes desta análise. '
            'Este contexto é memória processual e de governança; não altera automaticamente o Valor Total Atualizado, que permanece calculado por execução atualizada + saldo remanescente atualizado.'
        )
        col_ctx1, col_ctx2 = st.columns(2)
        with col_ctx1:
            valor_original_txt = st.text_input(
                'Valor original do contrato',
                value=contexto_salvo.get('valor_original_contrato_texto', ''),
                placeholder='Ex.: 20.000.000,00',
                key='ctx_valor_original_contrato',
            )
            opcoes_ciclo_concedido = ['C0 / Nenhum', 'C1', 'C2', 'C3', 'C4', 'Outro']
            ciclo_salvo = str(contexto_salvo.get('ultimo_ciclo_concedido', '') or '').strip()
            if ciclo_salvo in opcoes_ciclo_concedido:
                indice_ciclo_salvo = opcoes_ciclo_concedido.index(ciclo_salvo)
            elif ciclo_salvo == '' or ciclo_salvo.upper() == 'C0' or ciclo_salvo == 'Nenhum / C0':
                indice_ciclo_salvo = 0
            elif ciclo_salvo.upper() in opcoes_ciclo_concedido:
                indice_ciclo_salvo = opcoes_ciclo_concedido.index(ciclo_salvo.upper())
            else:
                indice_ciclo_salvo = len(opcoes_ciclo_concedido) - 1
            ultimo_ciclo = st.selectbox(
                'Último ciclo já concedido/formalizado',
                options=opcoes_ciclo_concedido,
                index=indice_ciclo_salvo,
                key='ctx_ultimo_ciclo_concedido',
            )
            data_base_ultimo = st.text_input(
                'Data-base do último ciclo concedido/formalizado',
                value=contexto_salvo.get('data_base_ultimo_ciclo', ''),
                placeholder='Ex.: 23/09/2023',
                key='ctx_data_base_ultimo_ciclo',
            )
            data_pedido_ultimo = st.text_input(
                'Data do pedido do último ciclo concedido/formalizado',
                value=contexto_salvo.get('data_pedido_ultimo_ciclo', ''),
                placeholder='Ex.: 23/09/2024',
                key='ctx_data_pedido_ultimo_ciclo',
            )
        with col_ctx2:
            modo_valor_formalizado = st.radio(
                'Como deseja informar o valor formalizado antes desta análise?',
                options=['Informar valor diretamente', 'Calcular pelo valor original + percentual já aplicado'],
                index=1 if contexto_salvo.get('modo_valor_formalizado') == 'Calcular pelo valor original + percentual já aplicado' else 0,
                key='ctx_modo_valor_formalizado',
            )
            percentual_ja_aplicado_pct = st.number_input(
                'Percentual já aplicado antes desta análise (%)',
                min_value=0.0,
                max_value=1000.0,
                value=float(contexto_salvo.get('percentual_ja_aplicado_pct', 0.0) or 0.0),
                step=0.01,
                format='%.2f',
                key='ctx_percentual_ja_aplicado_pct',
                disabled=(modo_valor_formalizado == 'Informar valor diretamente'),
            )
            valor_original_previo = _parse_moeda_br(valor_original_txt)
            valor_calculado_formalizado = round(valor_original_previo * (1 + percentual_ja_aplicado_pct / 100), 2) if valor_original_previo > 0 else 0.0

            if modo_valor_formalizado == 'Calcular pelo valor original + percentual já aplicado':
                fator_aplicado_previo = 1 + percentual_ja_aplicado_pct / 100
                valor_calculado_txt = _formatar_moeda_br(valor_calculado_formalizado)
                fator_txt = f"{fator_aplicado_previo:.6f}".replace('.', ',')
                st.info(
                    f"Memória do valor formalizado anterior: {_formatar_moeda_br(valor_original_previo)} × {fator_txt} = {valor_calculado_txt}. "
                    "O campo abaixo permanece editável para ajuste manual, se houver aditivo, supressão, repactuação ou arredondamento formal."
                )

                chave_calculo_formalizado = f"{valor_original_previo:.2f}|{percentual_ja_aplicado_pct:.6f}|auto"
                if st.session_state.get('ctx_valor_formalizado_auto_hash') != chave_calculo_formalizado:
                    st.session_state['ctx_valor_formalizado_anterior'] = valor_calculado_txt
                    st.session_state['ctx_valor_formalizado_auto_hash'] = chave_calculo_formalizado
                valor_formalizado_default = st.session_state.get('ctx_valor_formalizado_anterior', valor_calculado_txt)
            else:
                valor_formalizado_default = contexto_salvo.get('valor_formalizado_anterior_texto', '')

            valor_formalizado_txt = st.text_input(
                'Valor contratual formalizado antes desta análise',
                value=valor_formalizado_default,
                placeholder='Ex.: 22.800.000,00',
                key='ctx_valor_formalizado_anterior',
                help='Valor editável. Ajuste manualmente se houver aditivos, supressões, repactuações ou arredondamento formal.',
            )
            referencia_documental_historico = st.text_input(
                'Referência documental do histórico anterior',
                value=contexto_salvo.get('referencia_documental_historico', ''),
                placeholder='Ex.: TLB-TDA-2025/00001, Parecer, Despacho ou Ata',
                key='ctx_referencia_documental_historico',
            )
            observacao = st.text_area(
                'Observação sobre o histórico anterior',
                value=contexto_salvo.get('observacao_historico', ''),
                placeholder='Ex.: C1 e C2 já concedidos; valor inclui aditivo anterior formalizado.',
                key='ctx_observacao_historico',
                height=82,
            )

        eventos_limpos = []
        with st.expander("Eventos históricos adicionais (opcional)", expanded=False):
            st.caption(
                "Use apenas para registrar outros eventos formalizados anteriores que não estejam cobertos pelo bloco principal acima, "
                "como aditivo, supressão, repactuação, apostila anterior ou acordo negocial. "
                "O último ciclo concedido já é exportado automaticamente pelo Contexto do Contrato."
            )
            eventos_salvos = contexto_salvo.get('eventos_historicos_anteriores') or []
            # Não reapresenta na tabela registros automáticos gerados em versões anteriores.
            eventos_salvos = [
                ev for ev in eventos_salvos
                if not (isinstance(ev, dict) and 'Registro gerado automaticamente pelo Contexto do Contrato' in str(ev.get('Observação', '') or ''))
            ]
            if not eventos_salvos:
                eventos_salvos = [
                    {
                        'Tipo de evento': '',
                        'Ciclo': 'C0 / Nenhum',
                        'Data': '',
                        'Valor formalizado/impacto': '',
                        'Incorporado ao valor formalizado?': 'Sim',
                        'Observação': '',
                    }
                ]
            df_eventos_ctx = pd.DataFrame(eventos_salvos)
            if 'Valor formalizado/impacto' not in df_eventos_ctx.columns:
                if 'Valor atualizado/formalizado' in df_eventos_ctx.columns:
                    df_eventos_ctx['Valor formalizado/impacto'] = df_eventos_ctx['Valor atualizado/formalizado']
                elif 'Valor original' in df_eventos_ctx.columns:
                    df_eventos_ctx['Valor formalizado/impacto'] = df_eventos_ctx['Valor original']
                else:
                    df_eventos_ctx['Valor formalizado/impacto'] = ''
            colunas_eventos = [
                'Tipo de evento', 'Ciclo', 'Data', 'Valor formalizado/impacto',
                'Incorporado ao valor formalizado?', 'Observação'
            ]
            for col_evento in colunas_eventos:
                if col_evento not in df_eventos_ctx.columns:
                    df_eventos_ctx[col_evento] = ''
            df_eventos_ctx = df_eventos_ctx[colunas_eventos]
            eventos_editados = st.data_editor(
                df_eventos_ctx,
                hide_index=True,
                use_container_width=True,
                num_rows='dynamic',
                key='ctx_eventos_historicos_anteriores',
                column_config={
                    'Tipo de evento': st.column_config.SelectboxColumn(
                        'Tipo de evento',
                        options=['', 'Reajuste', 'Repactuação', 'Aditivo', 'Supressão', 'Apostila anterior', 'Acordo negocial', 'Outro'],
                    ),
                    'Ciclo': st.column_config.SelectboxColumn(
                        'Ciclo',
                        options=['C0 / Nenhum', 'C1', 'C2', 'C3', 'C4', 'Outro'],
                    ),
                    'Incorporado ao valor formalizado?': st.column_config.SelectboxColumn(
                        'Incorporado ao valor formalizado?',
                        options=['Sim', 'Não'],
                    ),
                },
            )
            for _, evento_row in eventos_editados.iterrows():
                evento = {col: str(evento_row.get(col, '') or '').strip() for col in colunas_eventos}
                possui_conteudo_relevante = any([
                    evento.get('Tipo de evento', ''),
                    evento.get('Data', ''),
                    evento.get('Valor formalizado/impacto', ''),
                    evento.get('Observação', ''),
                ])
                if possui_conteudo_relevante:
                    eventos_limpos.append(evento)

        valor_original = _parse_moeda_br(valor_original_txt)
        valor_formalizado = _parse_moeda_br(valor_formalizado_txt)
        contexto = {
            'valor_original_contrato': valor_original,
            'valor_original_contrato_texto': valor_original_txt,
            'valor_formalizado_anterior': valor_formalizado,
            'valor_formalizado_anterior_texto': valor_formalizado_txt,
            'modo_valor_formalizado': modo_valor_formalizado,
            'percentual_ja_aplicado_pct': float(percentual_ja_aplicado_pct),
            'valor_formalizado_calculado': float(valor_calculado_formalizado),
            'ultimo_ciclo_concedido': ultimo_ciclo.strip(),
            'data_base_ultimo_ciclo': data_base_ultimo.strip(),
            'data_pedido_ultimo_ciclo': data_pedido_ultimo.strip(),
            'referencia_documental_historico': referencia_documental_historico.strip(),
            'observacao_historico': observacao.strip(),
            'eventos_historicos_anteriores': eventos_limpos,
        }
        st.session_state['contexto_contratual_anterior'] = contexto

        if valor_original > 0 or valor_formalizado > 0 or ultimo_ciclo.strip() or observacao.strip():
            st.info(
                f"Contexto informado: valor original { _formatar_moeda_br_md(valor_original) }; "
                f"valor formalizado antes desta análise { _formatar_moeda_br_md(valor_formalizado) }."
            )
    return st.session_state.get('contexto_contratual_anterior', {})


def render_botao_download_modelo_consumo(data, file_name="Modelo_Consumo_por_Itens_Ciclo.xlsx"):
    """Renderiza botão HTML em paleta verde/terra para download do modelo Consumo por Itens/Ciclo."""
    try:
        encoded = base64.b64encode(data).decode("utf-8")
    except Exception:
        st.error("Não foi possível preparar o modelo de Consumo por Itens/Ciclo para download.")
        return

    st.markdown(
        f"""
        <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{encoded}"
           download="{file_name}"
           style="
                display:inline-flex;
                align-items:center;
                gap:0.45rem;
                background:#4E6E58;
                color:#FFFFFF;
                padding:0.56rem 0.82rem;
                border-radius:0.55rem;
                border:1px solid #3F5A48;
                font-weight:700;
                font-size:0.92rem;
                text-decoration:none;
                box-shadow:0 1px 2px rgba(15, 23, 42, 0.12);
                margin-top:0.10rem;
           ">
           🌿 Baixar modelo — Consumo por Itens/Ciclo
        </a>
        <div style="font-size:0.78rem;color:#6B7280;margin-top:0.35rem;margin-bottom:0.45rem;">
            Modelo alternativo em tons de verde/terra para apuração itemizada por consumo executado em cada ciclo.
        </div>
        """,
        unsafe_allow_html=True,
    )

def gerar_modelo_consumo_itens_ciclo_excel(dados_admissibilidade):
    """Gera modelo enxuto para o Modo Consumo por Itens/Ciclo.

    Uso previsto: após a Etapa 1 da análise, com ciclos, percentuais e fatores
    já apurados pelo cl8us. O fiscal preenche apenas dados originais dos itens
    e quantidades consumidas por ciclo. O restante é calculado automaticamente.
    """
    from xlsxwriter.utility import xl_col_to_name

    output = BytesIO()
    dados_admissibilidade = dados_admissibilidade or {}
    ciclos_origem = dados_admissibilidade.get('ciclos', []) or []
    data_geracao = datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')
    indice_etapa_1 = (
        dados_admissibilidade.get('indice')
        or dados_admissibilidade.get('indice_utilizado')
        or dados_admissibilidade.get('indice_contratual')
        or 'Não informado'
    )

    def _limpar_status(valor):
        texto = '' if valor is None else str(valor)
        for ch in ['✅', '❌', '⚠️', '⚠', '🟡', '🔴', '🟢', '🔵', '🟣', '🔻', '▲', '■', '●', '•']:
            texto = texto.replace(ch, '')
        return re.sub(r'\s+', ' ', texto).strip()

    def _valor_ciclo(ciclo, *chaves, padrao=''):
        for chave in chaves:
            valor = ciclo.get(chave) if isinstance(ciclo, dict) else None
            if valor not in [None, '']:
                return valor
        return padrao

    def _numero_seguro(valor, padrao=0.0):
        try:
            if valor is None or valor == '':
                return padrao
            return float(valor)
        except Exception:
            try:
                texto = str(valor).replace('%', '').replace('.', '').replace(',', '.').strip()
                return float(texto)
            except Exception:
                return padrao

    linhas_ciclos = []
    linhas_ciclos.append({
        'Ciclo': 'C0',
        'Data-base': dados_admissibilidade.get('data_base_original', ''),
        'Janela de admissibilidade': 'Ciclo-base inicial, sem reajuste',
        'Data do pedido': 'Não se aplica',
        'Início financeiro': 'Não se aplica',
        'Percentual aplicado': 0.0,
        'Fator do ciclo': 1.0,
        'Fator acumulado': 1.0,
        'Situação': 'Base sem reajuste',
        'Referência para preenchimento': 'Use C0 para consumo anterior ao início financeiro do C1.',
        'Observação': 'Consumo em C0 não gera retroativo de reajuste.',
    })

    for ciclo in ciclos_origem:
        if not isinstance(ciclo, dict):
            continue
        ciclo_nome = str(_valor_ciclo(ciclo, 'ciclo', 'Ciclo', 'nome', padrao='')).strip()
        if not ciclo_nome:
            continue
        if not ciclo_nome.upper().startswith('C'):
            ciclo_nome = f'C{ciclo_nome}'
        percentual = _valor_ciclo(ciclo, 'percentual_aplicado', 'Percentual aplicado', 'variacao', 'Variação', padrao=0.0)
        percentual_num = _numero_seguro(percentual, 0.0)
        if abs(percentual_num) > 1:
            percentual_num = percentual_num / 100.0
        fator = _valor_ciclo(ciclo, 'fator', 'Fator', padrao=1.0)
        fator_num = _numero_seguro(fator, 1.0)
        fator_acum = _valor_ciclo(ciclo, 'fator_acumulado', 'Fator acumulado', padrao=fator_num)
        fator_acum_num = _numero_seguro(fator_acum, fator_num)
        situacao = _limpar_status(_valor_ciclo(ciclo, 'situacao_aplicada', 'Situação aplicada', 'situacao', 'Situação', padrao=''))
        inicio_fin = _valor_ciclo(ciclo, 'financeiro_inicio', 'Início financeiro', 'Inicio financeiro', padrao='')
        ref_preenchimento = (
            f'Use {ciclo_nome} para quantitativos consumidos/executados a partir do início financeiro do ciclo ({inicio_fin}).'
            if inicio_fin else
            f'Use {ciclo_nome} para quantitativos consumidos/executados no período do ciclo, conforme validação fiscal.'
        )
        obs_partes = []
        if _valor_ciclo(ciclo, 'ciclo_ja_concedido', padrao=False):
            obs_partes.append('Ciclo já concedido/formalizado anteriormente.')
        if _valor_ciclo(ciclo, 'superacao_negocial', padrao=False):
            obs_partes.append('Tratamento negocial registrado na Etapa 1.')
        justificativa = str(_valor_ciclo(ciclo, 'justificativa_negocial', 'Justificativa negocial', padrao='') or '').strip()
        if justificativa:
            obs_partes.append(justificativa)
        linhas_ciclos.append({
            'Ciclo': ciclo_nome,
            'Data-base': _valor_ciclo(ciclo, 'data_base', 'Data-base', padrao=''),
            'Janela de admissibilidade': _valor_ciclo(ciclo, 'janela_admissibilidade', 'JanelaAdm', 'Janela de admissibilidade', padrao=''),
            'Data do pedido': _valor_ciclo(ciclo, 'data_pedido', 'Pedido', 'Data do pedido', padrao=''),
            'Início financeiro': inicio_fin,
            'Percentual aplicado': percentual_num,
            'Fator do ciclo': fator_num,
            'Fator acumulado': fator_acum_num,
            'Situação': situacao,
            'Referência para preenchimento': ref_preenchimento,
            'Observação': ' '.join(obs_partes),
        })

    qtd_linhas_itens = 240
    ciclos_modelo = [linha['Ciclo'] for linha in linhas_ciclos if linha.get('Ciclo')]
    if not ciclos_modelo:
        ciclos_modelo = ['C0', 'C1']
    ultimo_ciclo = ciclos_modelo[-1]
    ultima_linha_ciclos = len(linhas_ciclos) + 4  # cabeçalho em linha 4; dados a partir da linha 5

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Paleta verde/terra do novo modo.
        cor_verde_escuro = '#4E6E58'
        cor_verde_medio = '#7A8F63'
        cor_terra = '#8D6E63'
        cor_areia = '#F6F3EE'
        cor_areia_2 = '#EFE6D8'
        cor_input = '#FFF2CC'
        cor_auto = '#EDEDED'
        cor_alerta = '#EAF4EA'
        cor_alerta_terra = '#F4E7D3'
        cor_total = '#D9EAD3'

        fmt_title = workbook.add_format({'bold': True, 'font_size': 15, 'font_color': '#FFFFFF', 'bg_color': cor_verde_escuro, 'align': 'left', 'valign': 'vcenter'})
        fmt_subtitle = workbook.add_format({'font_color': '#3F3F3F', 'bg_color': cor_areia, 'text_wrap': True, 'valign': 'top'})
        fmt_section = workbook.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': cor_terra, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        fmt_header = workbook.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': cor_verde_escuro, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
        fmt_header_terra = workbook.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': cor_terra, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
        fmt_text = workbook.add_format({'border': 1, 'valign': 'top', 'text_wrap': True})
        fmt_input = workbook.add_format({'border': 1, 'bg_color': cor_input, 'valign': 'top', 'text_wrap': True})
        fmt_input_num = workbook.add_format({'border': 1, 'bg_color': cor_input, 'num_format': '#,##0.00'})
        fmt_input_money = workbook.add_format({'border': 1, 'bg_color': cor_input, 'num_format': 'R$ #,##0.00'})
        fmt_auto = workbook.add_format({'border': 1, 'bg_color': cor_auto, 'num_format': '#,##0.00'})
        fmt_auto_text = workbook.add_format({'border': 1, 'bg_color': cor_auto, 'text_wrap': True})
        fmt_auto_money = workbook.add_format({'border': 1, 'bg_color': cor_auto, 'num_format': 'R$ #,##0.00'})
        fmt_ciclo_a = workbook.add_format({'border': 1, 'bg_color': '#EEF6ED', 'num_format': '#,##0.00'})
        fmt_ciclo_b = workbook.add_format({'border': 1, 'bg_color': '#F4E7D3', 'num_format': '#,##0.00'})
        fmt_ciclo_money_a = workbook.add_format({'border': 1, 'bg_color': '#EEF6ED', 'num_format': 'R$ #,##0.00'})
        fmt_ciclo_money_b = workbook.add_format({'border': 1, 'bg_color': '#F4E7D3', 'num_format': 'R$ #,##0.00'})
        fmt_money = workbook.add_format({'border': 1, 'num_format': 'R$ #,##0.00'})
        fmt_percent = workbook.add_format({'border': 1, 'num_format': '0.00%'})
        fmt_factor = workbook.add_format({'border': 1, 'num_format': '0.0000'})
        fmt_total = workbook.add_format({'bold': True, 'bg_color': cor_total, 'border': 1})
        fmt_total_num = workbook.add_format({'bold': True, 'bg_color': cor_total, 'border': 1, 'num_format': '#,##0.00'})
        fmt_total_money = workbook.add_format({'bold': True, 'bg_color': cor_total, 'border': 1, 'num_format': 'R$ #,##0.00'})
        fmt_note = workbook.add_format({'bg_color': cor_alerta, 'font_color': '#274E13', 'text_wrap': True, 'valign': 'top', 'border': 1})
        fmt_note_terra = workbook.add_format({'bg_color': cor_alerta_terra, 'font_color': '#5F3B1B', 'text_wrap': True, 'valign': 'top', 'border': 1})
        fmt_kpi_label = workbook.add_format({'bold': True, 'bg_color': cor_areia_2, 'border': 1, 'text_wrap': True})
        fmt_kpi_value = workbook.add_format({'bold': True, 'bg_color': '#FFFFFF', 'border': 1, 'num_format': 'R$ #,##0.00'})
        fmt_plain = workbook.add_format({'valign': 'top', 'text_wrap': True})
        fmt_check_ok = workbook.add_format({'bg_color': '#E2F0D9', 'font_color': '#274E13'})
        fmt_check_bad = workbook.add_format({'bg_color': '#FCE4D6', 'font_color': '#9C0006'})

        # INICIO
        ws = workbook.add_worksheet('INICIO')
        writer.sheets['INICIO'] = ws
        ws.hide_gridlines(2)
        ws.set_column('A:A', 3)
        ws.set_column('B:B', 34)
        ws.set_column('C:C', 76)
        ws.set_row(1, 30)
        ws.merge_range('B2:C2', 'Modelo — Consumo por Itens/Ciclo', fmt_title)
        ws.merge_range('B3:C5', 'Use este modelo quando não houver base financeira mensal, mas a fiscalização puder informar, por item, as quantidades efetivamente consumidas/executadas em cada ciclo. O cl8us já preenche os ciclos e fatores apurados na Etapa 1. O fiscal preenche apenas os dados originais dos itens e as quantidades consumidas.', fmt_subtitle)
        orientacoes = [
            ('1. Confira os ciclos', 'A aba CICLOS_APURADOS é gerada pelo cl8us e destaca o início dos efeitos financeiros de cada ciclo.'),
            ('2. Preencha somente o essencial', 'Na aba CONSUMO_ITENS, informe item, quantidade contratada, valor unitário original e as quantidades consumidas em C0, C1, C2 etc.'),
            ('3. Não preencha cálculos', 'Valores unitários atualizados, retroativos, saldo a faturar e Valor Total Atualizado são calculados automaticamente.'),
            ('4. Valide o resumo', 'A aba RESUMO consolida execução original, execução atualizada, retroativo itemizado e saldo a faturar atualizado.'),
        ]
        row = 7
        for titulo, texto in orientacoes:
            ws.write(row, 1, titulo, fmt_section)
            ws.write(row, 2, texto, fmt_text)
            ws.set_row(row, 38)
            row += 1
        ws.write('B13', 'Data de geração', fmt_kpi_label)
        ws.write('C13', data_geracao, fmt_text)

        # PARAMETROS
        ws = workbook.add_worksheet('PARAMETROS')
        writer.sheets['PARAMETROS'] = ws
        ws.hide_gridlines(2)
        ws.set_column('A:A', 34)
        ws.set_column('B:B', 88)
        ws.set_row(0, 30)
        ws.merge_range('A1:B1', 'Parâmetros declaratórios do modo Consumo por Itens/Ciclo', fmt_title)
        ws.merge_range('A2:B3', 'Esta aba não deve repetir dados já apurados na Etapa 1. O objetivo é registrar apenas a premissa fiscal que autoriza usar consumo por item/ciclo como base da apuração.', fmt_note)
        parametros = [
            ('Método de apuração', 'Consumo por Itens/Ciclo'),
            ('Índice apurado na Etapa 1', indice_etapa_1),
            ('Premissa de equivalência fiscal', 'Confirmada pela fiscalização: o consumo por item/ciclo corresponde à execução demandada, medida/aprovada e faturável.'),
            ('Ressalvas sobre divergências financeiras', 'Sem ressalvas informadas quanto a glosas, multas, descontos, retenções, notas substituídas ou divergências entre consumo, medição, faturamento e pagamento.'),
            ('Observação do cl8us', 'Planilha destinada à apuração itemizada por consumo/ciclo, sem base financeira mensal por competência. Os efeitos financeiros constam de forma destacada na aba CICLOS_APURADOS.'),
        ]
        ws.write(4, 0, 'Campo', fmt_header)
        ws.write(4, 1, 'Valor', fmt_header)
        for idx, (campo, valor) in enumerate(parametros, start=5):
            ws.write(idx, 0, campo, fmt_kpi_label)
            fmt = fmt_input if campo in ['Premissa de equivalência fiscal', 'Ressalvas sobre divergências financeiras'] else fmt_text
            ws.write(idx, 1, valor, fmt)
            ws.set_row(idx, 42 if len(str(valor)) > 80 else 26)

        # CICLOS_APURADOS
        ws = workbook.add_worksheet('CICLOS_APURADOS')
        writer.sheets['CICLOS_APURADOS'] = ws
        ws.hide_gridlines(2)
        headers_ciclos = ['Ciclo', 'Data-base', 'Janela de admissibilidade', 'Data do pedido', 'Início financeiro', 'Percentual aplicado', 'Fator do ciclo', 'Fator acumulado', 'Situação', 'Referência para preenchimento', 'Observação']
        ws.merge_range(0, 0, 0, len(headers_ciclos)-1, 'Ciclos apurados na Etapa 1 do cl8us', fmt_title)
        ws.merge_range(1, 0, 2, len(headers_ciclos)-1, 'Efeitos financeiros: a aba CONSUMO_ITENS deve distribuir o consumo nas colunas C0, C1, C2 etc. conforme o ciclo em que a execução/consumo produziu efeito financeiro. O campo "Início financeiro" abaixo é a referência principal para esse enquadramento. C0 deve ser usado para consumo anterior ao início financeiro do C1.', fmt_note_terra)
        for col, title in enumerate(headers_ciclos):
            ws.write(3, col, title, fmt_header)
        for row_idx, linha in enumerate(linhas_ciclos, start=4):
            ws.write(row_idx, 0, linha['Ciclo'], fmt_text)
            ws.write(row_idx, 1, linha['Data-base'], fmt_text)
            ws.write(row_idx, 2, linha['Janela de admissibilidade'], fmt_text)
            ws.write(row_idx, 3, linha['Data do pedido'], fmt_text)
            ws.write(row_idx, 4, linha['Início financeiro'], fmt_text)
            ws.write_number(row_idx, 5, float(linha['Percentual aplicado'] or 0), fmt_percent)
            ws.write_number(row_idx, 6, float(linha['Fator do ciclo'] or 1), fmt_factor)
            ws.write_number(row_idx, 7, float(linha['Fator acumulado'] or 1), fmt_factor)
            ws.write(row_idx, 8, linha['Situação'], fmt_text)
            ws.write(row_idx, 9, linha['Referência para preenchimento'], fmt_text)
            ws.write(row_idx, 10, linha['Observação'], fmt_text)
        ws.set_column('A:A', 10)
        ws.set_column('B:E', 20)
        ws.set_column('F:H', 16)
        ws.set_column('I:I', 34)
        ws.set_column('J:K', 54)
        ws.freeze_panes(4, 0)

        # CONSUMO_ITENS — preenchimento em matriz simples.
        ws = workbook.add_worksheet('CONSUMO_ITENS')
        writer.sheets['CONSUMO_ITENS'] = ws
        ws.hide_gridlines(2)
        n_ciclos = len(ciclos_modelo)
        col_item = 0
        col_qtd = 1
        col_vu = 2
        col_primeiro_ciclo = 3
        col_ultimo_ciclo = col_primeiro_ciclo + n_ciclos - 1
        col_consumido_total = col_ultimo_ciclo + 1
        col_saldo = col_consumido_total + 1
        col_check = col_saldo + 1
        headers_consumo = ['Item', 'Quantidade contratada', 'Valor unitário original/base'] + [f'Consumido {c}' for c in ciclos_modelo] + ['Consumido total', 'Saldo a faturar', 'Check quantidade']
        ws.merge_range(0, 0, 0, col_check, 'Preenchimento fiscal — dados originais e consumo por ciclo', fmt_title)
        ws.merge_range(1, 0, 2, col_check, 'Preencha apenas as células amarelas: Item, Quantidade contratada, Valor unitário original/base e quantidades consumidas por ciclo. As colunas cinzas são automáticas. Não informe descrição, período, fator ou valor atualizado nesta aba.', fmt_note)
        for col, title in enumerate(headers_consumo):
            fmt = fmt_header_terra if col >= col_primeiro_ciclo and col <= col_ultimo_ciclo else fmt_header
            ws.write(3, col, title, fmt)
        for row in range(4, 4 + qtd_linhas_itens):
            excel_row = row + 1
            ws.write(row, col_item, '', fmt_input)
            ws.write(row, col_qtd, '', fmt_input_num)
            ws.write(row, col_vu, '', fmt_input_money)
            for col in range(col_primeiro_ciclo, col_ultimo_ciclo + 1):
                ws.write(row, col, '', fmt_input_num)
            first_cycle_letter = xl_col_to_name(col_primeiro_ciclo)
            last_cycle_letter = xl_col_to_name(col_ultimo_ciclo)
            qtd_letter = xl_col_to_name(col_qtd)
            total_letter = xl_col_to_name(col_consumido_total)
            saldo_letter = xl_col_to_name(col_saldo)
            ws.write_formula(row, col_consumido_total, f'=IF(A{excel_row}="","",ROUND(SUM({first_cycle_letter}{excel_row}:{last_cycle_letter}{excel_row}),2))', fmt_auto)
            ws.write_formula(row, col_saldo, f'=IF(A{excel_row}="","",ROUND({qtd_letter}{excel_row}-{total_letter}{excel_row},2))', fmt_auto)
            ws.write_formula(row, col_check, f'=IF(A{excel_row}="","",IF({saldo_letter}{excel_row}<-0.004,"DIVERGÊNCIA: consumo maior que contratado","OK"))', fmt_auto_text)
        total_row = 4 + qtd_linhas_itens
        ws.write(total_row, col_item, 'TOTAL', fmt_total)
        ws.write_formula(total_row, col_qtd, f'=ROUND(SUM({xl_col_to_name(col_qtd)}5:{xl_col_to_name(col_qtd)}{total_row}),2)', fmt_total_num)
        ws.write(total_row, col_vu, '', fmt_total)
        for col in range(col_primeiro_ciclo, col_ultimo_ciclo + 1):
            col_l = xl_col_to_name(col)
            ws.write_formula(total_row, col, f'=ROUND(SUM({col_l}5:{col_l}{total_row}),2)', fmt_total_num)
        ws.write_formula(total_row, col_consumido_total, f'=ROUND(SUM({xl_col_to_name(col_consumido_total)}5:{xl_col_to_name(col_consumido_total)}{total_row}),2)', fmt_total_num)
        ws.write_formula(total_row, col_saldo, f'=ROUND(SUM({xl_col_to_name(col_saldo)}5:{xl_col_to_name(col_saldo)}{total_row}),2)', fmt_total_num)
        ws.write(total_row, col_check, '', fmt_total)
        ws.set_column(col_item, col_item, 12)
        ws.set_column(col_qtd, col_qtd, 20)
        ws.set_column(col_vu, col_vu, 24)
        ws.set_column(col_primeiro_ciclo, col_ultimo_ciclo, 16)
        ws.set_column(col_consumido_total, col_saldo, 18)
        ws.set_column(col_check, col_check, 34)
        ws.freeze_panes(4, 0)
        ws.conditional_format(4, col_check, total_row-1, col_check, {'type': 'text', 'criteria': 'containing', 'value': 'OK', 'format': fmt_check_ok})
        ws.conditional_format(4, col_check, total_row-1, col_check, {'type': 'text', 'criteria': 'containing', 'value': 'DIVERGÊNCIA', 'format': fmt_check_bad})

        # CICLO_EM_EXECUCAO — controle declaratório da data de corte informada.
        ws = workbook.add_worksheet('CICLO_EM_EXECUCAO')
        writer.sheets['CICLO_EM_EXECUCAO'] = ws
        ws.hide_gridlines(2)
        ws.set_column('A:A', 38)
        ws.set_column('B:B', 44)
        ws.set_row(0, 30)
        ws.merge_range('A1:B1', 'Ciclo em execução e data de corte', fmt_title)
        ws.merge_range('A2:B3', 'Use esta aba para registrar até qual competência/mês a execução por itens foi informada. Este controle não substitui a distribuição das quantidades na aba CONSUMO_ITENS; ele documenta a data de corte usada para o cálculo do saldo a faturar.', fmt_note_terra)
        dados_corte = [
            ('Ciclo em execução na data de corte', ultimo_ciclo),
            ('Competência final da execução informada (mm/aaaa)', ''),
            ('Início financeiro do ciclo em execução', f'=IFERROR(VLOOKUP(B5,CICLOS_APURADOS!$A$5:$E${ultima_linha_ciclos},5,FALSE),"")'),
            ('Orientação', 'Informe a competência final até a qual o consumo foi lançado na aba CONSUMO_ITENS. O saldo a faturar será calculado como quantidade contratada menos consumo total informado.'),
        ]
        ws.write(4, 0, 'Campo', fmt_header)
        ws.write(4, 1, 'Valor', fmt_header)
        for idx, (campo, valor) in enumerate(dados_corte, start=5):
            ws.write(idx, 0, campo, fmt_kpi_label)
            if isinstance(valor, str) and valor.startswith('='):
                ws.write_formula(idx, 1, valor, fmt_text)
            else:
                ws.write(idx, 1, valor, fmt_input if idx == 6 else fmt_text)
            ws.set_row(idx, 34 if idx == 8 else 24)

        # CALCULO_AUTOMATICO — derivado da matriz simples.
        ws = workbook.add_worksheet('CALCULO_AUTOMATICO')
        writer.sheets['CALCULO_AUTOMATICO'] = ws
        ws.hide_gridlines(2)
        col = 0
        base_headers = ['Item', 'Quantidade contratada', 'Valor unitário original/base', 'Consumido total', 'Saldo a faturar']
        headers_calc = base_headers[:]
        cycle_blocks = []
        for ciclo in ciclos_modelo:
            cycle_blocks.append({
                'ciclo': ciclo,
                'qtd_col': len(headers_calc),
                'fator_col': len(headers_calc) + 1,
                'vu_atu_col': len(headers_calc) + 2,
                'orig_col': len(headers_calc) + 3,
                'atu_col': len(headers_calc) + 4,
                'retroativo_col': len(headers_calc) + 5,
            })
            headers_calc.extend([
                f'Qtd {ciclo}', f'Fator {ciclo}', f'VU atualizado {ciclo}',
                f'Valor original {ciclo}', f'Valor atualizado {ciclo}', f'Retroativo {ciclo}'
            ])
        col_saldo_fator = len(headers_calc)
        headers_calc.extend(['Fator saldo atual', 'Valor saldo original', 'Valor saldo atualizado'])
        ws.merge_range(0, 0, 0, len(headers_calc)-1, 'Cálculo automático — não preencher', fmt_title)
        ws.merge_range(1, 0, 2, len(headers_calc)-1, 'Esta aba é calculada automaticamente a partir da aba CONSUMO_ITENS e dos fatores da aba CICLOS_APURADOS. Serve de memória de cálculo para importação futura pelo cl8us.', fmt_note_terra)
        for col, title in enumerate(headers_calc):
            header_fmt = fmt_header if col < 5 else fmt_header_terra
            for bloco_idx, bloco in enumerate(cycle_blocks):
                if bloco['qtd_col'] <= col <= bloco['retroativo_col']:
                    header_fmt = fmt_header_terra if bloco_idx % 2 == 0 else fmt_header
                    break
            ws.write(3, col, title, header_fmt)
        for row in range(4, 4 + qtd_linhas_itens):
            excel_row = row + 1
            consumo_row = excel_row
            ws.write_formula(row, 0, f'=CONSUMO_ITENS!A{consumo_row}', fmt_auto_text)
            ws.write_formula(row, 1, f'=N(CONSUMO_ITENS!B{consumo_row})', fmt_auto)
            ws.write_formula(row, 2, f'=N(CONSUMO_ITENS!C{consumo_row})', fmt_auto_money)
            ws.write_formula(row, 3, f'=N(CONSUMO_ITENS!{xl_col_to_name(col_consumido_total)}{consumo_row})', fmt_auto)
            ws.write_formula(row, 4, f'=N(CONSUMO_ITENS!{xl_col_to_name(col_saldo)}{consumo_row})', fmt_auto)
            for bloco_idx, bloco in enumerate(cycle_blocks):
                ciclo = bloco['ciclo']
                consumo_col = col_primeiro_ciclo + bloco_idx
                consumo_col_l = xl_col_to_name(consumo_col)
                qtd_col_l = xl_col_to_name(bloco['qtd_col'])
                fator_col_l = xl_col_to_name(bloco['fator_col'])
                vu_atu_col_l = xl_col_to_name(bloco['vu_atu_col'])
                orig_col_l = xl_col_to_name(bloco['orig_col'])
                atu_col_l = xl_col_to_name(bloco['atu_col'])
                fmt_num_ciclo = fmt_ciclo_a if bloco_idx % 2 == 0 else fmt_ciclo_b
                fmt_money_ciclo = fmt_ciclo_money_a if bloco_idx % 2 == 0 else fmt_ciclo_money_b
                ws.write_formula(row, bloco['qtd_col'], f'=CONSUMO_ITENS!{consumo_col_l}{consumo_row}', fmt_num_ciclo)
                ws.write_formula(row, bloco['fator_col'], f'=IFERROR(VLOOKUP("{ciclo}",CICLOS_APURADOS!$A$5:$H${ultima_linha_ciclos},8,FALSE),1)', fmt_num_ciclo)
                ws.write_formula(row, bloco['vu_atu_col'], f'=IF(A{excel_row}="","",IFERROR(ROUND(N($C{excel_row})*N({fator_col_l}{excel_row}),2),0))', fmt_money_ciclo)
                ws.write_formula(row, bloco['orig_col'], f'=IF(A{excel_row}="","",IFERROR(ROUND(N({qtd_col_l}{excel_row})*N($C{excel_row}),2),0))', fmt_money_ciclo)
                ws.write_formula(row, bloco['atu_col'], f'=IF(A{excel_row}="","",IFERROR(ROUND(N({qtd_col_l}{excel_row})*N({vu_atu_col_l}{excel_row}),2),0))', fmt_money_ciclo)
                ws.write_formula(row, bloco['retroativo_col'], f'=IF(A{excel_row}="","",IFERROR(ROUND(N({atu_col_l}{excel_row})-N({orig_col_l}{excel_row}),2),0))', fmt_money_ciclo)
            saldo_fator_l = xl_col_to_name(col_saldo_fator)
            saldo_orig_l = xl_col_to_name(col_saldo_fator + 1)
            ws.write_formula(row, col_saldo_fator, f'=IFERROR(VLOOKUP("{ultimo_ciclo}",CICLOS_APURADOS!$A$5:$H${ultima_linha_ciclos},8,FALSE),1)', fmt_auto)
            ws.write_formula(row, col_saldo_fator + 1, f'=IF(A{excel_row}="","",IFERROR(ROUND(N(E{excel_row})*N($C{excel_row}),2),0))', fmt_auto_money)
            ws.write_formula(row, col_saldo_fator + 2, f'=IF(A{excel_row}="","",IFERROR(ROUND(N(E{excel_row})*ROUND(N($C{excel_row})*N({saldo_fator_l}{excel_row}),2),2),0))', fmt_auto_money)
        total_calc_row = 4 + qtd_linhas_itens
        ws.write(total_calc_row, 0, 'TOTAL', fmt_total)
        for col in range(1, len(headers_calc)):
            col_l = xl_col_to_name(col)
            fmt = fmt_total_money if 'Valor' in headers_calc[col] or 'Retroativo' in headers_calc[col] or 'VU atualizado' in headers_calc[col] else fmt_total_num
            if headers_calc[col].startswith('Fator'):
                ws.write(total_calc_row, col, '', fmt_total)
            else:
                ws.write_formula(total_calc_row, col, f'=IFERROR(ROUND(SUM({col_l}5:{col_l}{total_calc_row}),2),0)', fmt)
        ws.set_column(0, 0, 12)
        ws.set_column(1, 4, 18)
        ws.set_column(5, len(headers_calc)-1, 18)
        ws.freeze_panes(4, 0)

        # RESUMO
        ws = workbook.add_worksheet('RESUMO')
        writer.sheets['RESUMO'] = ws
        ws.hide_gridlines(2)
        ws.set_column('A:A', 42)
        ws.set_column('B:B', 26)
        ws.set_column('D:I', 22)
        ws.merge_range('A1:I1', 'Resumo — Modo Consumo por Itens/Ciclo', fmt_title)
        ws.merge_range('A2:I3', 'Resumo calculado automaticamente a partir da matriz de consumo. O Valor Total Atualizado do Contrato segue a lógica: execução atualizada por ciclo + saldo a faturar atualizado.', fmt_note)
        ws.write('A5', 'Indicador', fmt_header)
        ws.write('B5', 'Valor', fmt_header)
        orig_cols = [xl_col_to_name(b['orig_col']) for b in cycle_blocks]
        atu_cols = [xl_col_to_name(b['atu_col']) for b in cycle_blocks]
        retroativo_cols = [xl_col_to_name(b['retroativo_col']) for b in cycle_blocks]
        total_calc_excel_row = total_calc_row + 1
        exec_orig_formula = '=IFERROR(ROUND(' + '+'.join([f'CALCULO_AUTOMATICO!{c}{total_calc_excel_row}' for c in orig_cols]) + ',2),0)'
        exec_atu_formula = '=IFERROR(ROUND(' + '+'.join([f'CALCULO_AUTOMATICO!{c}{total_calc_excel_row}' for c in atu_cols]) + ',2),0)'
        retroativo_formula = '=ROUND(' + '+'.join([f'CALCULO_AUTOMATICO!{c}{total_calc_excel_row}' for c in retroativo_cols]) + ',2)'
        saldo_atu_formula = f'=IFERROR(ROUND(CALCULO_AUTOMATICO!{xl_col_to_name(col_saldo_fator + 2)}{total_calc_excel_row},2),0)'
        indicadores = [
            ('Valor original do contrato', f'=ROUND(SUMPRODUCT(CONSUMO_ITENS!B5:B{total_row},CONSUMO_ITENS!C5:C{total_row}),2)'),
            ('Execução original por itens', exec_orig_formula),
            ('Execução atualizada por itens', exec_atu_formula),
            ('Retroativo (itens consumidos/ciclo)', retroativo_formula),
            ('Saldo Remanescente Atualizado', saldo_atu_formula),
            ('Valor Total Atualizado do Contrato', '=IFERROR(ROUND(B8+B10,2),0)'),
        ]
        for idx, (label, formula) in enumerate(indicadores, start=5):
            ws.write(idx, 0, label, fmt_kpi_label)
            ws.write_formula(idx, 1, formula, fmt_kpi_value)
        ws.write('A13', 'Premissa de equivalência fiscal', fmt_kpi_label)
        ws.write_formula('B13', '=PARAMETROS!B8', fmt_text)
        ws.write('A14', 'Ressalvas financeiras', fmt_kpi_label)
        ws.write_formula('B14', '=PARAMETROS!B9', fmt_text)
        ws.write('A15', 'Efeitos financeiros', fmt_kpi_label)
        ws.write('B15', 'Consultar aba CICLOS_APURADOS. O consumo deve ser distribuído por ciclo conforme o início financeiro de cada ciclo.', fmt_text)
        ws.write('A16', 'Ciclo em execução/data de corte', fmt_kpi_label)
        ws.write_formula('B16', '=CICLO_EM_EXECUCAO!B5&" — execução informada até: "&IF(CICLO_EM_EXECUCAO!B6="","[informar mm/aaaa]",CICLO_EM_EXECUCAO!B6)', fmt_text)

        ws.write('D5', 'Resumo por ciclo', fmt_header)
        ws.write('D6', 'Ciclo', fmt_header)
        ws.write('E6', 'Qtd consumida', fmt_header)
        ws.write('F6', 'Fator acumulado', fmt_header)
        ws.write('G6', 'Execução original', fmt_header)
        ws.write('H6', 'Execução atualizada', fmt_header)
        ws.write('I6', 'Retroativo', fmt_header)
        for idx, bloco in enumerate(cycle_blocks, start=6):
            ciclo = bloco['ciclo']
            excel_row = idx + 1
            ws.write(idx, 3, ciclo, fmt_text)
            ws.write_formula(idx, 4, f'=IFERROR(ROUND(CALCULO_AUTOMATICO!{xl_col_to_name(bloco["qtd_col"])}{total_calc_excel_row},2),0)', fmt_money)
            ws.write_formula(idx, 5, f'=IFERROR(VLOOKUP(D{excel_row},CICLOS_APURADOS!$A$5:$H${ultima_linha_ciclos},8,FALSE),1)', fmt_factor)
            ws.write_formula(idx, 6, f'=IFERROR(ROUND(CALCULO_AUTOMATICO!{xl_col_to_name(bloco["orig_col"])}{total_calc_excel_row},2),0)', fmt_money)
            ws.write_formula(idx, 7, f'=IFERROR(ROUND(CALCULO_AUTOMATICO!{xl_col_to_name(bloco["atu_col"])}{total_calc_excel_row},2),0)', fmt_money)
            ws.write_formula(idx, 8, f'=IFERROR(ROUND(CALCULO_AUTOMATICO!{xl_col_to_name(bloco["retroativo_col"])}{total_calc_excel_row},2),0)', fmt_money)

        total_resumo_row = 6 + len(cycle_blocks)
        excel_total_resumo = total_resumo_row + 1
        ws.write(total_resumo_row, 3, 'TOTAL', fmt_total)
        ws.write_formula(total_resumo_row, 4, f'=IFERROR(ROUND(SUM(E7:E{excel_total_resumo-1}),2),0)', fmt_total_num)
        ws.write(total_resumo_row, 5, '', fmt_total)
        ws.write_formula(total_resumo_row, 6, f'=IFERROR(ROUND(SUM(G7:G{excel_total_resumo-1}),2),0)', fmt_total_money)
        ws.write_formula(total_resumo_row, 7, f'=IFERROR(ROUND(SUM(H7:H{excel_total_resumo-1}),2),0)', fmt_total_money)
        ws.write_formula(total_resumo_row, 8, f'=IFERROR(ROUND(SUM(I7:I{excel_total_resumo-1}),2),0)', fmt_total_money)

        ws.write('A20', 'Checklist automático', fmt_header)
        checklist = [
            ('Há itens informados?', f'=IF(COUNTA(CONSUMO_ITENS!A5:A{total_row})>0,"OK","Pendente")'),
            ('Há consumo informado?', f'=IF(SUM(CONSUMO_ITENS!{xl_col_to_name(col_consumido_total)}5:{xl_col_to_name(col_consumido_total)}{total_row})>0,"OK","Pendente")'),
            ('Há saldo negativo?', f'=IF(COUNTIF(CONSUMO_ITENS!{xl_col_to_name(col_check)}5:{xl_col_to_name(col_check)}{total_row},"*DIVERGÊNCIA*")>0,"Revisar","OK")'),
            ('Premissa fiscal preenchida?', '=IF(PARAMETROS!B8<>"","OK","Pendente")'),
        ]
        for idx, (label, formula) in enumerate(checklist, start=20):
            ws.write(idx, 0, label, fmt_kpi_label)
            ws.write_formula(idx, 1, formula, fmt_text)

        # Abas com cores distintas.
        for sheet_name in ['INICIO', 'PARAMETROS', 'CICLOS_APURADOS', 'CONSUMO_ITENS', 'CICLO_EM_EXECUCAO', 'CALCULO_AUTOMATICO', 'RESUMO']:
            writer.sheets[sheet_name].set_tab_color(cor_verde_medio if sheet_name not in ['CICLOS_APURADOS', 'CICLO_EM_EXECUCAO', 'CALCULO_AUTOMATICO'] else cor_terra)

    output.seek(0)
    return output.getvalue()

def gerar_arquivo_coleta_excel(dados_admissibilidade):
    """Gera o Arquivo de Coleta para as fases de Valor Global e Relatório.

    Regras da planilha:
    - A aba ITENS_REMANESCENTES usa linha TOTAL dinâmica por fórmula, sem tabela estruturada.
    - Ao inserir novas linhas acima da linha TOTAL, as fórmulas devem se ajustar no Excel.
    - Aditivos por item/lançamento, com acréscimo ou supressão.
    """
    output = io.BytesIO()
    ciclos = dados_admissibilidade.get('ciclos', [])
    data_geracao = datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        fmt_header = workbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': '#1F4E79',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        fmt_subheader = workbook.add_format({
            'bold': True, 'bg_color': '#D9EAD3', 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        fmt_input = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1})
        fmt_input_date = workbook.add_format({'num_format': 'dd/mm/yyyy', 'bg_color': '#FFF2CC', 'border': 1})
        fmt_date = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})
        fmt_date_no_border = workbook.add_format({'num_format': 'dd/mm/yyyy'})
        fmt_input_num = workbook.add_format({'num_format': '#,##0.00', 'bg_color': '#FFF2CC', 'border': 1})
        fmt_input_money = workbook.add_format({'num_format': 'R$ #,##0.00', 'bg_color': '#FFF2CC', 'border': 1})
        fmt_input_no_border = workbook.add_format({'bg_color': '#FFF2CC'})
        fmt_input_date_no_border = workbook.add_format({'num_format': 'dd/mm/yyyy', 'bg_color': '#FFF2CC'})
        fmt_input_num_no_border = workbook.add_format({'num_format': '#,##0.00', 'bg_color': '#FFF2CC'})
        fmt_input_money_no_border = workbook.add_format({'num_format': 'R$ #,##0.00', 'bg_color': '#FFF2CC'})
        fmt_money = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        fmt_money_no_border = workbook.add_format({'num_format': 'R$ #,##0.00'})
        fmt_money_auto = workbook.add_format({'num_format': 'R$ #,##0.00', 'bg_color': '#EDEDED', 'border': 1})
        fmt_number = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        fmt_text = workbook.add_format({'border': 1})
        fmt_auto = workbook.add_format({'bg_color': '#EDEDED', 'border': 1})
        fmt_total = workbook.add_format({'bold': True, 'bg_color': '#E2F0D9', 'border': 1})
        fmt_total_money = workbook.add_format({'bold': True, 'bg_color': '#E2F0D9', 'border': 1, 'num_format': 'R$ #,##0.00'})
        fmt_total_num = workbook.add_format({'bold': True, 'bg_color': '#E2F0D9', 'border': 1, 'num_format': '#,##0.00'})
        fmt_decrease_red = workbook.add_format({'font_color': '#C00000'})
        fmt_percent = workbook.add_format({'num_format': '0.00%', 'border': 1})
        fmt_factor = workbook.add_format({'num_format': '0.0000', 'border': 1})
        fmt_factor_auto = workbook.add_format({'num_format': '0.0000', 'bg_color': '#EDEDED', 'border': 1})
        fmt_no_border = workbook.add_format({})
        fmt_int_no_border = workbook.add_format({'num_format': '0'})
        fmt_percent_no_border = workbook.add_format({'num_format': '0.00%'})
        fmt_factor_no_border = workbook.add_format({'num_format': '0.0000'})
        fmt_text_left_no_border = workbook.add_format({'align': 'left'})
        fmt_money_left_no_border = workbook.add_format({'num_format': 'R$ #,##0.00', 'align': 'left'})
        fmt_text_wrap_no_border = workbook.add_format({'align': 'left', 'valign': 'top', 'text_wrap': True})
        fmt_header_no_border = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#1F4E79', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        fmt_text_wrap = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'text_wrap': True})
        fmt_int_left_no_border = workbook.add_format({'num_format': '0', 'align': 'left'})
        fmt_percent_left_no_border = workbook.add_format({'num_format': '0.00%', 'align': 'left'})
        fmt_factor_left_no_border = workbook.add_format({'num_format': '0.0000', 'align': 'left'})
        fmt_gap_date = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1, 'bg_color': '#F4CCCC', 'font_color': '#9C0006'})
        fmt_gap_text = workbook.add_format({'border': 1, 'bg_color': '#F4CCCC', 'font_color': '#9C0006', 'align': 'center'})
        fmt_gap_text_wrap = workbook.add_format({'border': 1, 'bg_color': '#F4CCCC', 'font_color': '#9C0006', 'text_wrap': True, 'valign': 'top'})
        fmt_gap_date_no_border = workbook.add_format({'num_format': 'dd/mm/yyyy', 'bg_color': '#F4CCCC', 'font_color': '#9C0006'})
        fmt_gap_text_no_border = workbook.add_format({'bg_color': '#F4CCCC', 'font_color': '#9C0006', 'align': 'center'})
        fmt_gap_text_wrap_no_border = workbook.add_format({'bg_color': '#F4CCCC', 'font_color': '#9C0006', 'text_wrap': True, 'valign': 'top'})
        fmt_gap_input_money = workbook.add_format({'num_format': 'R$ #,##0.00', 'bg_color': '#F4CCCC', 'font_color': '#9C0006', 'border': 1})
        fmt_gap_input_money_no_border = workbook.add_format({'num_format': 'R$ #,##0.00', 'bg_color': '#F4CCCC', 'font_color': '#9C0006'})
        fmt_total_no_border = workbook.add_format({'bold': True, 'bg_color': '#E2F0D9'})
        fmt_total_money_no_border = workbook.add_format({'bold': True, 'bg_color': '#E2F0D9', 'num_format': 'R$ #,##0.00'})
        fmt_total_num_no_border = workbook.add_format({'bold': True, 'bg_color': '#E2F0D9', 'num_format': '#,##0.00'})
        cores_ciclos_col_c = ['#EAF2F8', '#E2F0D9', '#FFF2CC', '#FCE4D6', '#E4DFEC', '#DDEBF7', '#F4CCCC', '#D9EAD3']
        fmt_ciclos_col_c = [workbook.add_format({'bg_color': cor, 'border': 1}) for cor in cores_ciclos_col_c]
        cor_aba_automatica = '#D9EAF7'

        def _periodo_mensal_seguro(valor):
            try:
                dt = pd.to_datetime(valor, dayfirst=True, errors='coerce')
                if pd.notna(dt):
                    return dt.to_period('M')
            except Exception:
                pass
            try:
                texto = str(valor or '').strip()
                m = re.search(r'(\d{1,2})[\/\-](\d{4})', texto)
                if m:
                    return pd.Period(f"{int(m.group(2))}-{int(m.group(1)):02d}", freq='M')
            except Exception:
                pass
            return None

        def _fim_intervalo_indice_periodo(intervalo):
            texto = str(intervalo or '')
            matches = re.findall(r'(\d{1,2})[\/\-](\d{4})', texto)
            if len(matches) >= 2:
                mes, ano = matches[-1]
                try:
                    return pd.Period(f"{int(ano)}-{int(mes):02d}", freq='M')
                except Exception:
                    return None
            return None

        def _competencias_sem_efeito_financeiro(ciclo):
            fim_indice = _fim_intervalo_indice_periodo(ciclo.get('intervalo_indice', ciclo.get('Janela', '')))
            inicio_fin = _periodo_mensal_seguro(ciclo.get('financeiro_inicio', ''))
            if fim_indice is None or inicio_fin is None:
                return []
            primeira_comp_habilitada = fim_indice + 1
            ultima_comp_sem_efeito = inicio_fin - 1
            if ultima_comp_sem_efeito < primeira_comp_habilitada:
                return []
            return [p.strftime('%m/%Y') for p in pd.period_range(primeira_comp_habilitada, ultima_comp_sem_efeito, freq='M')]

        def _primeira_competencia_base_execucao(ciclo):
            """Define a primeira competência da BASE_EXECUCAO_MENSAL.

            Regra contratual consolidada:
            - se o pedido/efeito financeiro ocorre no próprio mês em que o ciclo se habilita,
              a base mensal deve começar no mês do início financeiro;
            - se houver atraso no pedido, a base mensal deve começar na primeira competência
              habilitada após o intervalo do índice, para evidenciar os meses sem efeito financeiro;
            - portanto, usa-se a menor competência entre o início financeiro e a primeira
              competência habilitada. Isso preserva casos antigos, evita deslocamento de 1 mês
              e permite demonstrar competências sem retroativo quando houver lapso.
            """
            fim_indice = _fim_intervalo_indice_periodo(ciclo.get('intervalo_indice', ciclo.get('Janela', '')))
            inicio_fin = _periodo_mensal_seguro(ciclo.get('financeiro_inicio', ''))

            if fim_indice is not None and inicio_fin is not None:
                primeira_habilitada = fim_indice + 1
                return min(inicio_fin, primeira_habilitada).strftime('%d/%m/%Y')
            if inicio_fin is not None:
                return inicio_fin.strftime('%d/%m/%Y')
            if fim_indice is not None:
                return (fim_indice + 1).strftime('%d/%m/%Y')
            return ciclo.get('financeiro_inicio', '')


        def _data_inicio_teorico_ciclo_para_itens(ciclo):
            """Data da fotografia de itens/remanescentes: início teórico do ciclo.

            A aba ITENS_REMANESCENTES não deve usar a data-base do índice nem o início
            financeiro do pedido. Regra: usa o primeiro dia da janela de admissibilidade;
            se ausente, usa data_base + 12 meses.
            """
            ciclo = ciclo or {}
            for chave in ['janela_admissibilidade', 'Janela de admissibilidade', 'JanelaAdm']:
                texto = str(ciclo.get(chave, '') or '').strip()
                if not texto:
                    continue
                m = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', texto)
                if m:
                    return m.group(1)
            data_base = ciclo.get('data_base', '') or ciclo.get('Data-base', '')
            dt = pd.to_datetime(data_base, dayfirst=True, errors='coerce')
            if pd.notna(dt):
                return (dt + relativedelta(years=1)).strftime('%d/%m/%Y')
            return str(data_base or '').strip()

        # PARAMETROS_REAJUSTE
        parametros = pd.DataFrame([
            ['Origem da análise', dados_admissibilidade.get('origem', dados_admissibilidade.get('tipo', ''))],
            ['Índice utilizado', dados_admissibilidade.get('indice', '')],
            ['Data-base original', dados_admissibilidade.get('data_base_original', '')],
            ['Quantidade de ciclos', len(ciclos)],
            ['Variação acumulada final', dados_admissibilidade.get('variacao_acumulada', 0.0)],
            ['Fator acumulado total', dados_admissibilidade.get('fator_acumulado', dados_admissibilidade.get('fator', 1.0))],
            ['Valor original do contrato (contexto)', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('valor_original_contrato', '')],
            ['Valor contratual formalizado antes desta análise', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('valor_formalizado_anterior', '')],
            ['Último ciclo já concedido/formalizado', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('ultimo_ciclo_concedido', '')],
            ['Data-base do último ciclo concedido/formalizado', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('data_base_ultimo_ciclo', '')],
            ['Data do pedido do último ciclo concedido/formalizado', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('data_pedido_ultimo_ciclo', '')],
            ['Percentual já aplicado antes desta análise', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('percentual_ja_aplicado_pct', '')],
            ['Referência documental do histórico anterior', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('referencia_documental_historico', '')],
            ['Observação sobre histórico anterior', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('observacao_historico', '')],
            ['Eventos históricos anteriores', 'Ver aba EVENTOS_HISTORICOS_ANTERIORES' if dados_admissibilidade.get('contexto_contratual_anterior', {}).get('eventos_historicos_anteriores', []) else ''],
            ['Data de geração do arquivo', data_geracao],
        ], columns=['Campo', 'Valor'])
        parametros.to_excel(writer, sheet_name='PARAMETROS_REAJUSTE', index=False)
        ws = writer.sheets['PARAMETROS_REAJUSTE']
        ws.set_column('A:A', 34)
        ws.set_column('B:B', 36)
        ws.write(0, 0, 'Campo', fmt_header)
        ws.write(0, 1, 'Valor', fmt_header)
        # Linhas 5, 6 e 7: alinhamento à esquerda para leitura limpa do bloco de parâmetros.
        ws.write(4, 0, 'Quantidade de ciclos', fmt_text_left_no_border)
        ws.write_number(4, 1, len(ciclos), fmt_int_left_no_border)
        ws.write(5, 0, 'Variação acumulada final', fmt_text_left_no_border)
        ws.write_number(5, 1, float(dados_admissibilidade.get('variacao_acumulada', 0.0)), fmt_percent_left_no_border)
        ws.write(6, 0, 'Fator acumulado total', fmt_text_left_no_border)
        ws.write_number(6, 1, float(dados_admissibilidade.get('fator_acumulado', dados_admissibilidade.get('fator', 1.0))), fmt_factor_left_no_border)
        ws.set_tab_color(cor_aba_automatica)
        contexto_excel = dados_admissibilidade.get('contexto_contratual_anterior', {}) or {}
        valor_original_contexto_txt = str(contexto_excel.get('valor_original_contrato_texto') or '').strip()
        valor_formalizado_contexto_txt = str(contexto_excel.get('valor_formalizado_anterior_texto') or '').strip()
        valor_original_contexto = float(contexto_excel.get('valor_original_contrato') or 0.0)
        valor_formalizado_contexto = float(contexto_excel.get('valor_formalizado_anterior') or 0.0)
        if valor_original_contexto_txt or valor_original_contexto > 0:
            ws.write_number(7, 1, valor_original_contexto, fmt_money_no_border)
        else:
            ws.write_blank(7, 1, None, fmt_no_border)
        if valor_formalizado_contexto_txt or valor_formalizado_contexto > 0:
            ws.write_number(8, 1, valor_formalizado_contexto, fmt_money_no_border)
        else:
            ws.write_blank(8, 1, None, fmt_no_border)

        # Formatação final da aba PARAMETROS_REAJUSTE: coluna B alinhada à esquerda,
        # com preservação dos formatos numéricos relevantes.
        for row_idx in range(1, len(parametros) + 1):
            valor_param = parametros.iloc[row_idx - 1, 1]
            ws.write(row_idx, 1, valor_param, fmt_text_left_no_border)
        ws.write_number(4, 1, len(ciclos), fmt_int_left_no_border)
        ws.write_number(5, 1, float(dados_admissibilidade.get('variacao_acumulada', 0.0)), fmt_percent_left_no_border)
        ws.write_number(6, 1, float(dados_admissibilidade.get('fator_acumulado', dados_admissibilidade.get('fator', 1.0))), fmt_factor_left_no_border)
        if valor_original_contexto_txt or valor_original_contexto > 0:
            ws.write_number(7, 1, valor_original_contexto, fmt_money_left_no_border)
        else:
            ws.write_blank(7, 1, None, fmt_text_left_no_border)
        if valor_formalizado_contexto_txt or valor_formalizado_contexto > 0:
            ws.write_number(8, 1, valor_formalizado_contexto, fmt_money_left_no_border)
        else:
            ws.write_blank(8, 1, None, fmt_text_left_no_border)
        percentual_contexto = contexto_excel.get('percentual_ja_aplicado_pct', '')
        try:
            if str(percentual_contexto).strip() != '':
                ws.write_number(12, 1, float(percentual_contexto) / 100, fmt_percent_left_no_border)
        except Exception:
            ws.write(12, 1, percentual_contexto, fmt_text_left_no_border)


        # Eventos históricos anteriores em aba própria, para evitar JSON longo em B12.
        eventos_historicos_excel = _eventos_historicos_para_exportacao(contexto_excel)
        if eventos_historicos_excel:
            ws_ev = workbook.add_worksheet('EVENTOS_HISTORICOS_ANTERIORES')
            writer.sheets['EVENTOS_HISTORICOS_ANTERIORES'] = ws_ev
            ev_headers = ['Tipo de evento', 'Ciclo', 'Data', 'Valor formalizado/impacto', 'Incorporado ao valor formalizado?', 'Observação']
            for col, title in enumerate(ev_headers):
                ws_ev.write(0, col, title, fmt_header_no_border)
            for row_idx, evento in enumerate(eventos_historicos_excel, start=1):
                valor_evento = evento.get('Valor formalizado/impacto', evento.get('Valor atualizado/formalizado', evento.get('Valor original', '')))
                ws_ev.write(row_idx, 0, evento.get('Tipo de evento', ''), fmt_text_left_no_border)
                ws_ev.write(row_idx, 1, evento.get('Ciclo', ''), fmt_text_left_no_border)
                ws_ev.write(row_idx, 2, evento.get('Data', ''), fmt_text_left_no_border)
                ws_ev.write(row_idx, 3, valor_evento, fmt_text_left_no_border)
                ws_ev.write(row_idx, 4, evento.get('Incorporado ao valor formalizado?', ''), fmt_text_left_no_border)
                ws_ev.write(row_idx, 5, evento.get('Observação', ''), fmt_text_wrap_no_border)
            ws_ev.set_column('A:A', 24, fmt_text_left_no_border)
            ws_ev.set_column('B:B', 14, fmt_text_left_no_border)
            ws_ev.set_column('C:C', 18, fmt_text_left_no_border)
            ws_ev.set_column('D:E', 28, fmt_text_left_no_border)
            ws_ev.set_column('F:F', 80, fmt_text_wrap_no_border)
            ws_ev.set_tab_color(cor_aba_automatica)

        # CICLOS
        ciclos_rows = []
        for ciclo in ciclos:
            situacao = str(ciclo.get('situacao', ''))
            ciclo_ja_concedido = bool(ciclo.get('ciclo_ja_concedido', False))
            if ciclo_ja_concedido:
                tratamento = 'Já concedido/formalizado anteriormente'
            else:
                tratamento = 'Precluso' if 'PRECLUSO' in situacao.upper() else 'A apurar'
            competencias_sem_efeito = _competencias_sem_efeito_financeiro(ciclo)
            ciclos_rows.append({
                'Ciclo': ciclo.get('ciclo', ''),
                'Data-base': ciclo.get('data_base', ''),
                'Intervalo do índice': ciclo.get('intervalo_indice', ciclo.get('Janela', '')),
                'Janela de admissibilidade': ciclo.get('janela_admissibilidade', ciclo.get('JanelaAdm', '')),
                'Data do pedido': ciclo.get('data_pedido', ciclo.get('Pedido', '')),
                'Início financeiro': ciclo.get('financeiro_inicio', ''),
                'Fim financeiro': ciclo.get('financeiro_fim', ''),
                'Situação': situacao,
                'Variação': ciclo.get('variacao', 0.0),
                'Fator': ciclo.get('fator', 1.0),
                'Fator acumulado': ciclo.get('fator_acumulado', 1.0),
                'Tratamento financeiro do ciclo': tratamento,
                'Ciclo já concedido/formalizado?': 'Sim' if ciclo_ja_concedido else 'Não',
                'Objeto da análise atual?': 'Não' if ciclo_ja_concedido else 'Sim',
                'Situação automática': ciclo.get('situacao_automatica', ciclo.get('situacao', '')),
                'Acordo negocial': 'Sim' if ciclo.get('superacao_negocial', False) else 'Não',
                'Situação aplicada': ciclo.get('situacao_aplicada', ciclo.get('situacao', '')),
                'Percentual apurado pelo índice': ciclo.get('percentual_indice', ciclo.get('variacao', 0.0)),
                'Percentual aplicado': ciclo.get('percentual_aplicado', ciclo.get('variacao', 0.0)),
                'Justificativa negocial': ciclo.get('justificativa_negocial', ''),
                'Referência documental': ciclo.get('referencia_documental', ''),
                'Meses sem efeito financeiro': len(competencias_sem_efeito),
                'Competências sem efeito financeiro': ', '.join(competencias_sem_efeito),
            })
        df_ciclos = pd.DataFrame(ciclos_rows)
        df_ciclos.to_excel(writer, sheet_name='CICLOS', index=False)
        ws = writer.sheets['CICLOS']
        ws.hide_gridlines(2)
        ws.set_tab_color(cor_aba_automatica)
        ws.set_column('A:A', 12)
        ws.set_column('B:H', 24)
        ws.set_column('I:I', 12)
        ws.set_column('J:K', 14)
        ws.set_column('L:L', 34)
        ws.set_column('M:Q', 26)
        ws.set_column('R:S', 16)
        ws.set_column('T:U', 42, fmt_text_wrap_no_border)
        ws.set_column('V:V', 18)
        ws.set_column('W:W', 42, fmt_text_wrap_no_border)
        for col, title in enumerate(df_ciclos.columns):
            ws.write(0, col, title, fmt_header_no_border)
        for row in range(1, len(df_ciclos) + 1):
            data_base_excel = pd.to_datetime(df_ciclos.iloc[row-1]['Data-base'], dayfirst=True, errors='coerce')
            if pd.notna(data_base_excel):
                ws.write_datetime(row, 1, data_base_excel.to_pydatetime(), fmt_date_no_border)
            meses_sem_efeito = int(df_ciclos.iloc[row-1].get('Meses sem efeito financeiro', 0) or 0)
            inicio_financeiro_excel = pd.to_datetime(df_ciclos.iloc[row-1].get('Início financeiro', ''), dayfirst=True, errors='coerce')
            if pd.notna(inicio_financeiro_excel):
                ws.write_datetime(row, 5, inicio_financeiro_excel.to_pydatetime(), fmt_gap_date_no_border if meses_sem_efeito > 0 else fmt_date_no_border)
            elif meses_sem_efeito > 0:
                ws.write(row, 5, df_ciclos.iloc[row-1].get('Início financeiro', ''), fmt_gap_text_no_border)
            if meses_sem_efeito > 0:
                ws.write(row, df_ciclos.columns.get_loc('Meses sem efeito financeiro'), meses_sem_efeito, fmt_gap_text_no_border)
                ws.write(row, df_ciclos.columns.get_loc('Competências sem efeito financeiro'), df_ciclos.iloc[row-1].get('Competências sem efeito financeiro', ''), fmt_gap_text_wrap_no_border)
            ws.write(row, 8, df_ciclos.iloc[row-1]['Variação'], fmt_percent_no_border)
            ws.write(row, 9, df_ciclos.iloc[row-1]['Fator'], fmt_factor_no_border)
            ws.write(row, 10, df_ciclos.iloc[row-1]['Fator acumulado'], fmt_factor_no_border)
            ws.write(row, 11, df_ciclos.iloc[row-1]['Tratamento financeiro do ciclo'], fmt_no_border)
            fmt_percentual_ciclos = fmt_percent_no_border
            if 'Percentual apurado pelo índice' in df_ciclos.columns:
                col_pct_indice = df_ciclos.columns.get_loc('Percentual apurado pelo índice')
                ws.write(row, col_pct_indice, df_ciclos.iloc[row-1]['Percentual apurado pelo índice'], fmt_percentual_ciclos)
            if 'Percentual aplicado' in df_ciclos.columns:
                col_pct_aplicado = df_ciclos.columns.get_loc('Percentual aplicado')
                ws.write(row, col_pct_aplicado, df_ciclos.iloc[row-1]['Percentual aplicado'], fmt_percentual_ciclos)
            if 'Justificativa negocial' in df_ciclos.columns:
                col_justificativa = df_ciclos.columns.get_loc('Justificativa negocial')
                ws.write(row, col_justificativa, df_ciclos.iloc[row-1]['Justificativa negocial'], fmt_text_wrap_no_border)
            if 'Referência documental' in df_ciclos.columns:
                col_referencia = df_ciclos.columns.get_loc('Referência documental')
                ws.write(row, col_referencia, df_ciclos.iloc[row-1]['Referência documental'], fmt_text_wrap_no_border)

        # BASE_EXECUCAO_MENSAL
        financeiro_rows = []
        for ciclo in ciclos:
            if bool(ciclo.get('ciclo_ja_concedido', False)):
                continue
            ciclo_nome = ciclo.get('ciclo', '')
            competencias_sem_efeito = set(_competencias_sem_efeito_financeiro(ciclo))
            inicio_base_execucao = _primeira_competencia_base_execucao(ciclo)
            for competencia in _competencias_mensais(inicio_base_execucao, ciclo.get('financeiro_fim', '')):
                financeiro_rows.append({
                    'Ciclo': ciclo_nome,
                    'Competência': competencia,
                    'Valor bruto medido/aprovado por competência': '',
                    'Competência sem efeito financeiro?': 'Sim' if competencia in competencias_sem_efeito else 'Não',
                    'Observação sobre efeito financeiro': 'Competência anterior ao início dos efeitos financeiros do pedido; não compõe retroativo a pagar.' if competencia in competencias_sem_efeito else '',
                })
        if not financeiro_rows:
            financeiro_rows.append({'Ciclo': '', 'Competência': '', 'Valor bruto medido/aprovado por competência': '', 'Competência sem efeito financeiro?': '', 'Observação sobre efeito financeiro': ''})
        df_fin = pd.DataFrame(financeiro_rows)
        df_fin.to_excel(writer, sheet_name='BASE_EXECUCAO_MENSAL', index=False)
        ws = writer.sheets['BASE_EXECUCAO_MENSAL']
        ws.hide_gridlines(2)
        ws.set_column('A:A', 12, fmt_no_border)
        ws.set_column('B:B', 18, fmt_no_border)
        ws.set_column('C:C', 24, fmt_input_money_no_border)
        ws.set_column('D:D', 22, fmt_no_border)
        ws.set_column('E:E', 58, fmt_text_wrap_no_border)
        for col, title in enumerate(df_fin.columns):
            ws.write(0, col, title, fmt_header_no_border)
        for row in range(1, len(df_fin) + 1):
            sem_efeito_linha = str(df_fin.iloc[row-1].get('Competência sem efeito financeiro?', '')).strip().upper() == 'SIM'
            if sem_efeito_linha:
                ws.write(row, 0, df_fin.iloc[row-1].get('Ciclo', ''), fmt_gap_text_no_border)
                ws.write(row, 1, df_fin.iloc[row-1].get('Competência', ''), fmt_gap_text_no_border)
                ws.write(row, 2, '', fmt_gap_input_money_no_border)
                ws.write(row, 3, 'Sim', fmt_gap_text_no_border)
                ws.write(row, 4, df_fin.iloc[row-1].get('Observação sobre efeito financeiro', ''), fmt_gap_text_wrap_no_border)
            else:
                ws.write(row, 0, df_fin.iloc[row-1].get('Ciclo', ''), fmt_no_border)
                ws.write(row, 1, df_fin.iloc[row-1].get('Competência', ''), fmt_no_border)
                ws.write(row, 2, '', fmt_input_money_no_border)
                ws.write(row, 3, df_fin.iloc[row-1].get('Competência sem efeito financeiro?', ''), fmt_no_border)
                ws.write(row, 4, df_fin.iloc[row-1].get('Observação sobre efeito financeiro', ''), fmt_text_wrap_no_border)
        total_row_fin = len(df_fin) + 1
        ultima_linha_fin_excel = len(df_fin) + 1
        ws.write(total_row_fin, 0, 'TOTAL', fmt_total_no_border)
        ws.write(total_row_fin, 1, '', fmt_total_no_border)
        ws.write_formula(total_row_fin, 2, f'=ROUND(SUM(C2:C{ultima_linha_fin_excel}),2)', fmt_total_money_no_border)
        ws.insert_textbox('G2',
            'Orientação de preenchimento\n\n'
            '• Preencha a coluna C com o valor bruto demandado, medido ou aprovado por competência.\n'
            '• Competências marcadas como sem efeito financeiro ficam para memória e não compõem o retroativo a pagar.\n'
            '• Se a base vier de consumo itemizado, use o novo modelo Consumo por Itens/Ciclo e registre a premissa fiscal de equivalência.',
            {
                'width': 420,
                'height': 150,
                'fill': {'color': '#F6F3EE'},
                'line': {'color': '#8D6E63', 'width': 1.25},
                'font': {'name': 'Segoe UI', 'size': 9, 'color': '#274E13'},
                'align': {'vertical': 'top'},
                'margin': 8,
            }
        )

        # ITENS_REMANESCENTES
        ws_it = workbook.add_worksheet('ITENS_REMANESCENTES')
        writer.sheets['ITENS_REMANESCENTES'] = ws_it
        rem_cols = []
        for ciclo in ciclos:
            if bool(ciclo.get('ciclo_ja_concedido', False)):
                continue
            rem_cols.append((f"Remanescente início {ciclo.get('ciclo', '')}", _data_inicio_teorico_ciclo_para_itens(ciclo)))
        base_headers = ['Item', 'Quantidade contratada', 'Valor unitário original', 'Valor total']
        headers = base_headers + [c[0] for c in rem_cols]
        ws_it.merge_range(0, 0, 0, 3, 'Dados do item e valor original', fmt_subheader)
        for idx, (_, data_ref) in enumerate(rem_cols, start=4):
            ws_it.write(0, idx, data_ref, fmt_subheader)
        for col, title in enumerate(headers):
            ws_it.write(1, col, title, fmt_header)
        ws_it.set_column(0, 0, 12)
        ws_it.set_column(1, 1, 24)
        ws_it.set_column(2, 3, 22)
        if rem_cols:
            ws_it.set_column(4, 4 + len(rem_cols) - 1, 25)
        linhas_itens_iniciais = 180
        primeira_linha_dados = 2
        ultima_linha_dados = primeira_linha_dados + linhas_itens_iniciais - 1
        for row in range(primeira_linha_dados, ultima_linha_dados + 1):
            ws_it.write(row, 0, '', fmt_text)       # A sem preenchimento
            ws_it.write(row, 1, '', fmt_number)     # B sem preenchimento
            ws_it.write(row, 2, '', fmt_money)      # C sem preenchimento
            ws_it.write_formula(row, 3, f'=IF(OR(B{row+1}="",C{row+1}=""),"",ROUND(B{row+1}*C{row+1},2))', fmt_money_auto)
            for col in range(4, 4 + len(rem_cols)):
                ws_it.write(row, col, '', fmt_input_num)

        total_row_itens = ultima_linha_dados + 1
        # Não usar tabela estruturada do Excel aqui: em algumas instalações o XML da tabela
        # foi reparado/removido pelo Excel. A linha TOTAL permanece dinâmica porque as
        # fórmulas abaixo se ajustam quando novas linhas são inseridas acima dela.
        ws_it.write(total_row_itens, 0, 'TOTAL', fmt_total)
        ws_it.write(total_row_itens, 1, '', fmt_total)
        ws_it.write(total_row_itens, 2, '', fmt_total)
        ws_it.write_formula(total_row_itens, 3, f'=ROUND(SUM(D3:D{ultima_linha_dados + 1}),2)', fmt_total_money)
        for col in range(4, 4 + len(rem_cols)):
            col_letter = chr(ord('A') + col)
            ws_it.write_formula(total_row_itens, col, f'=SUM({col_letter}3:{col_letter}{ultima_linha_dados + 1})', fmt_total_num)
        ws_it.write(total_row_itens + 2, 0, 'Orientação', fmt_total)
        ws_it.merge_range(total_row_itens + 2, 1, total_row_itens + 2, max(3, len(headers) - 1), 'A linha TOTAL é dinâmica por fórmula. Há 180 linhas disponíveis para preenchimento. Para acrescentar mais itens, insira novas linhas acima da linha TOTAL; as fórmulas de total serão ajustadas pelo Excel. Não utilizar tabela estruturada nesta aba.', fmt_text_wrap)
        ws_it.set_row(total_row_itens + 2, 58)

        # ADITIVOS_QUANTITATIVOS
        ws_ad = workbook.add_worksheet('ADITIVOS_QUANTITATIVOS')
        writer.sheets['ADITIVOS_QUANTITATIVOS'] = ws_ad
        ad_headers = [
            'Item', 'Data do aditivo', 'Ciclo/Marco', 'Tipo de alteração',
            'Quantidade acrescida/suprimida', 'Valor unitário original',
            'Valor original da alteração', 'Aplicar reajuste acumulado? (Sim/Não)',
            'Fator acumulado aplicável', 'Valor atualizado da alteração',
            'Tratamento do aditivo'
        ]
        for col, title in enumerate(ad_headers):
            ws_ad.write(0, col, title, fmt_header)
        ws_ad.set_column('A:A', 12)
        ws_ad.set_column('B:B', 18)
        ws_ad.set_column('C:C', 16)
        ws_ad.set_column('D:D', 20)
        ws_ad.set_column('E:E', 28)
        ws_ad.set_column('F:G', 24)
        ws_ad.set_column('H:H', 32)
        ws_ad.set_column('I:I', 22)
        ws_ad.set_column('J:J', 26)
        ws_ad.set_column('K:K', 42)
        ws_ad.write(202, 0, "Orientação", fmt_total)
        ws_ad.write(202, 1, "Use 'Computar nesta análise' para aditivos/supressões que devem impactar o Valor Global atual. Use 'Informativo - já incluído no valor formalizado' quando o lançamento já estiver contemplado no campo Valor contratual formalizado antes desta análise. A coluna F (Valor unitário original) é automática: busca o item informado na coluna A na aba ITENS_REMANESCENTES e retorna o valor unitário original da coluna C.", fmt_text)
        ciclo_range = f'CICLOS!$A$2:$K${len(ciclos)+1}' if ciclos else 'CICLOS!$A$2:$K$2'
        data_range = f'CICLOS!$B$2:$B${len(ciclos)+1}' if ciclos else 'CICLOS!$B$2:$B$2'
        ciclo_nome_range = f'CICLOS!$A$2:$A${len(ciclos)+1}' if ciclos else 'CICLOS!$A$2:$A$2'
        for row in range(1, 200):
            excel_row = row + 1
            ws_ad.write(row, 0, '', fmt_text)
            ws_ad.write(row, 1, '', fmt_input_date)
            # Coluna C: identifica automaticamente o ciclo pela data do aditivo.
            # Regra: data do aditivo dentro do período financeiro do ciclo.
            # Período financeiro: Início financeiro <= Data do Aditivo <= Fim financeiro.
            if len(ciclos) > 0:
                formula_ciclo = (
                    f'=IF(B{excel_row}="","",'
                    f'IFERROR(LOOKUP(2,1/((CICLOS!$F$2:$F${len(ciclos)+1}<=B{excel_row})*'
                    f'(CICLOS!$G$2:$G${len(ciclos)+1}>=B{excel_row})),CICLOS!$A$2:$A${len(ciclos)+1}),'
                    f'IF(B{excel_row}<MIN(CICLOS!$F$2:$F${len(ciclos)+1}),"C0","Verificar ciclo")))'
                )
                ws_ad.write_formula(row, 2, formula_ciclo, fmt_auto)
            else:
                ws_ad.write(row, 2, '', fmt_auto)
            ws_ad.write(row, 3, 'Acréscimo', fmt_input)
            ws_ad.write(row, 4, '', fmt_input_num)
            ws_ad.write_formula(row, 5, f'=IF(A{excel_row}="","",IFERROR(VLOOKUP(A{excel_row},ITENS_REMANESCENTES!$A:$C,3,FALSE),""))', fmt_money_auto)
            ws_ad.write_formula(row, 6, f'=IF(OR(E{excel_row}="",F{excel_row}=""),"",ROUND(E{excel_row}*F{excel_row},2))', fmt_money_auto)
            ws_ad.write(row, 7, 'Sim', fmt_input)
            ws_ad.write_formula(row, 8, f'=IF(C{excel_row}="","",IFERROR(VLOOKUP(C{excel_row},{ciclo_range},11,FALSE),1))', fmt_factor_auto)
            ws_ad.write_formula(row, 9, f'=IF(G{excel_row}="","",ROUND(IF(OR(UPPER(D{excel_row})="DECRÉSCIMO",UPPER(D{excel_row})="DECRESCIMO",UPPER(D{excel_row})="SUPRESSÃO",UPPER(D{excel_row})="SUPRESSAO"),-ABS(G{excel_row}),ABS(G{excel_row}))*IF(OR(UPPER(H{excel_row})="NÃO",UPPER(H{excel_row})="NAO"),1,I{excel_row}),2))', fmt_money_auto)
            ws_ad.write(row, 10, 'Computar nesta análise', fmt_input)
        ws_ad.data_validation(1, 3, 199, 3, {'validate': 'list', 'source': ['Acréscimo', 'Decréscimo']})
        ws_ad.data_validation(1, 7, 199, 7, {'validate': 'list', 'source': ['Sim', 'Não']})
        ws_ad.data_validation(1, 10, 199, 10, {'validate': 'list', 'source': ['Computar nesta análise', 'Informativo - já incluído no valor formalizado']})
        # Coluna C: cor discreta por ciclo para facilitar a leitura quando houver mudança de marco.
        ws_ad.conditional_format(1, 2, 199, 2, {
            'type': 'formula',
            'criteria': '=$C2="C0"',
            'format': fmt_ciclos_col_c[0],
        })
        for idx_ciclo, ciclo_ref in enumerate(ciclos or []):
            ciclo_nome = str(ciclo_ref.get('ciclo', '') or '').strip()
            if ciclo_nome:
                ws_ad.conditional_format(1, 2, 199, 2, {
                    'type': 'formula',
                    'criteria': f'=$C2="{ciclo_nome}"',
                    'format': fmt_ciclos_col_c[(idx_ciclo + 1) % len(fmt_ciclos_col_c)],
                })
        # Se o usuário selecionar Decréscimo/Supressão, destacar toda a linha em fonte vermelha.
        ws_ad.conditional_format(1, 0, 199, 10, {
            'type': 'formula',
            # Fórmula robusta: usa início do texto para evitar falhas por acento/localidade.
            'criteria': '=OR(LEFT(UPPER($D2),4)="DECR",LEFT(UPPER($D2),6)="SUPRES")',
            'format': fmt_decrease_red,
        })
        ws_ad.write(200, 0, 'TOTAL', fmt_total)
        for col in range(1, 6):
            ws_ad.write(200, col, '', fmt_total)
        ws_ad.write_formula(200, 6, '=ROUND(SUM(G2:G200),2)', fmt_total_money)
        ws_ad.write(200, 7, '', fmt_total)
        ws_ad.write(200, 8, '', fmt_total)
        ws_ad.write_formula(200, 9, '=ROUND(SUM(J2:J200),2)', fmt_total_money)
        ws_ad.write(200, 10, '', fmt_total)

    output.seek(0)
    return output.getvalue()
if not st.session_state.get("_calculadora_reajustes_embedded", False):
    render_marca_topo()
if not st.session_state.get("_calculadora_reajustes_embedded", False):
    st.title("Múltiplos")

contexto_contratual = _render_contexto_contratual_anterior()

primeiro_ciclo_num = _primeiro_ciclo_analise(contexto_contratual)
default_dt_base_original = _data_base_inicial_pelo_contexto(contexto_contratual, datetime(2022, 10, 10))

if _ciclo_para_numero(contexto_contratual.get('ultimo_ciclo_concedido', '')) > 0 and not contexto_contratual.get('data_pedido_ultimo_ciclo'):
    st.error(
        "Processamento inviável: há ciclo anterior concedido/formalizado, mas a data do pedido do último ciclo não foi informada."
    )
    st.warning(
        "Informe a data do pedido do último ciclo no Contexto do Contrato. Esse dado será usado como âncora do próximo ciclo, pois corresponde ao início dos efeitos financeiros do último reajuste concedido/formalizado."
    )
    st.stop()

with st.sidebar:
    dt_base_original = st.date_input("Data-base/âncora inicial da análise atual:", value=default_dt_base_original, format="DD/MM/YYYY")
    qtd_ciclos = st.number_input("Ciclos a analisar:", min_value=1, max_value=5, value=2)
    idx_sel = render_indice_contrato_selectbox(key="indice_fluxo_multiplos")

# Primeira etapa: coleta dos dados de cada ciclo, sem processar automaticamente os índices.
# Isso evita que a página abra com um cenário fictício já calculado.
input_ciclos = []
containers_ciclos = []
data_atual = dt_base_original

for posicao_ciclo in range(1, int(qtd_ciclos) + 1):
    i = primeiro_ciclo_num + posicao_ciclo - 1
    st.markdown(f"### Ciclo {i}")

    d_fim = data_atual + relativedelta(months=11)
    d_aniv = data_atual + relativedelta(years=1)
    d_lim = d_aniv + relativedelta(days=90)

    col_a, col_b = st.columns(2)
    with col_a:
        st.write(f"**Data-Base do Ciclo:** {data_atual.strftime('%d/%m/%Y')}")
    with col_b:
        # A chave considera a âncora do ciclo. Assim, se um ciclo anterior for admitido
        # por negociação entre as partes e arrastar a data-base para frente, o campo do
        # pedido do ciclo seguinte também é recalculado a partir da nova âncora.
        chave_pedido = f"p{i}_{data_atual.strftime('%Y%m%d')}"
        dt_ped = st.date_input(
            f"Data do Pedido C{i}:",
            value=d_aniv,
            key=chave_pedido,
            format="DD/MM/YYYY",
        )

    ciclo_ja_concedido = st.checkbox(
        f"Ciclo C{i} já concedido/formalizado anteriormente",
        value=False,
        key=f"ciclo_ja_concedido_c{i}_{data_atual.strftime('%Y%m%d')}",
        help=(
            "Marque apenas se este ciclo já foi formalizado em instrumento anterior. "
            "O ciclo será preservado como histórico/âncora, mas não será tratado como objeto novo da análise atual."
        ),
    )
    if ciclo_ja_concedido:
        st.caption(
            f"C{i} será mantido como histórico/âncora. Não entrará como ciclo novo no relatório da análise atual nem na aba financeira da coleta."
        )

    # Lógica de Admissibilidade preservada.
    if dt_ped < d_aniv:
        if dt_ped.year == d_aniv.year and dt_ped.month == d_aniv.month:
            sit_emoji = "🟡 ADMISSÍVEL - RESSALVA"
            situacao_limpa = "ADMISSÍVEL - RESSALVA"
        else:
            sit_emoji = "⚠️ ADIANTADO"
            situacao_limpa = "ADIANTADO"
    elif dt_ped <= d_lim:
        sit_emoji = "✅ TEMPESTIVO"
        situacao_limpa = "TEMPESTIVO"
    else:
        sit_emoji = "❌ PRECLUSO"
        situacao_limpa = "PRECLUSO"

    # Regra de ancoragem do próximo ciclo preservada.
    if situacao_limpa == "TEMPESTIVO":
        data_base_proximo_ciclo = dt_ped
    else:
        data_base_proximo_ciclo = d_aniv

    inicio_efeito_financeiro = None if situacao_limpa == "PRECLUSO" else (dt_ped if dt_ped >= d_aniv else d_aniv)

    superacao_negocial = False
    percentual_negocial = 0.0
    justificativa_negocial = ""
    referencia_documental = ""
    data_inicio_efeito_negocial = None
    if situacao_limpa == "PRECLUSO":
        with st.expander(f"Acordo negocial de admissão de reajuste - C{i}", expanded=False):
            st.caption(
                "Use apenas quando houver decisão negocial fundamentada para conceder o ciclo, "
                "sem apagar o diagnóstico automático de preclusão."
            )
            superacao_negocial = st.checkbox(
                f"Ciclo admitido por negociação entre as partes - C{i}",
                value=False,
                key=f"superacao_negocial_c{i}",
            )
            if superacao_negocial:
                col_neg1, col_neg2 = st.columns(2)
                with col_neg1:
                    percentual_negocial = st.number_input(
                        f"Percentual aplicado por acordo C{i} (%)",
                        min_value=0.0,
                        max_value=100.0,
                        value=0.0,
                        step=0.01,
                        format="%.2f",
                        key=f"percentual_negocial_c{i}",
                    ) / 100
                    data_inicio_efeito_negocial = st.date_input(
                        f"Início dos efeitos financeiros por acordo C{i}",
                        value=dt_ped,
                        format="DD/MM/YYYY",
                        key=f"inicio_negocial_c{i}",
                    )
                with col_neg2:
                    referencia_documental = st.text_input(
                        f"Referência documental C{i}",
                        placeholder="Ex.: Despacho, Ata, Ofício ou Nota Técnica",
                        key=f"referencia_negocial_c{i}",
                    )
                justificativa_negocial = st.text_area(
                    f"Justificativa técnica/negocial C{i}",
                    placeholder="Registre a fundamentação da concessão por acordo negocial.",
                    key=f"justificativa_negocial_c{i}",
                    height=90,
                )
                if not justificativa_negocial.strip():
                    st.info("A justificativa deve ser preenchida para fins de memória processual antes da instrução final.")
                inicio_efeito_financeiro = data_inicio_efeito_negocial
                # Regra de ancoragem negocial:
                # se um ciclo precluso for admitido por negociação entre as partes, a âncora
                # do ciclo seguinte é arrastada para a data de início dos efeitos financeiros
                # pactuada para o ciclo admitido. Consequentemente, o próximo ciclo somente
                # estará apto após 12 meses desse novo marco.
                data_base_proximo_ciclo = data_inicio_efeito_negocial
                st.info(
                    f"Com a admissão negocial do C{i}, o próximo ciclo será ancorado em "
                    f"{data_base_proximo_ciclo.strftime('%d/%m/%Y')} e somente estará apto "
                    f"a partir de {(data_base_proximo_ciclo + relativedelta(years=1)).strftime('%d/%m/%Y')}."
                )

    input_ciclos.append({
        'numero': i,
        'data_atual': data_atual,
        'd_fim': d_fim,
        'd_aniv': d_aniv,
        'd_lim': d_lim,
        'dt_ped': dt_ped,
        'sit_emoji': sit_emoji,
        'situacao_limpa': situacao_limpa,
        'inicio_efeito_financeiro': inicio_efeito_financeiro,
        'ciclo_ja_concedido': bool(ciclo_ja_concedido),
        'superacao_negocial': bool(superacao_negocial),
        'percentual_negocial': float(percentual_negocial),
        'justificativa_negocial': justificativa_negocial.strip(),
        'referencia_documental': referencia_documental.strip(),
        'data_base_proximo_ciclo': data_base_proximo_ciclo,
    })

    containers_ciclos.append(st.container())
    data_atual = data_base_proximo_ciclo

chave_analise_multiplos = (
    primeiro_ciclo_num,
    dt_base_original.isoformat(),
    int(qtd_ciclos),
    idx_sel,
    tuple(c['dt_ped'].isoformat() for c in input_ciclos),
    tuple(str(c.get('ciclo_ja_concedido', False)) for c in input_ciclos),
    tuple(str(c.get('superacao_negocial', False)) for c in input_ciclos),
    tuple(str(c.get('percentual_negocial', 0.0)) for c in input_ciclos),
    tuple(str(c.get('justificativa_negocial', '')) for c in input_ciclos),
    str(contexto_contratual.get('valor_original_contrato', 0.0)),
    str(contexto_contratual.get('valor_formalizado_anterior', 0.0)),
    contexto_contratual.get('ultimo_ciclo_concedido', ''),
    contexto_contratual.get('data_pedido_ultimo_ciclo', ''),
    contexto_contratual.get('modo_valor_formalizado', ''),
    str(contexto_contratual.get('percentual_ja_aplicado_pct', 0.0)),
    contexto_contratual.get('observacao_historico', ''),
)

processar_multiplos = st.button(
    "Processar Análise",
    type="primary",
    use_container_width=False,
)

if processar_multiplos:
    st.session_state["processar_reajustes_multiplos_key"] = chave_analise_multiplos

if st.session_state.get("processar_reajustes_multiplos_key") != chave_analise_multiplos:
    st.info(f"Foram configurados **{int(qtd_ciclos)} ciclo(s)** para análise, iniciando em **C{primeiro_ciclo_num}**, a partir da âncora **{dt_base_original.strftime('%d/%m/%Y')}**. Confira as datas dos pedidos antes de clicar em **Processar Análise**.")
    st.stop()

data_hoje = datetime.now(ZoneInfo("America/Sao_Paulo")).date()
pedidos_futuros = []
for dados_ciclo in input_ciclos:
    data_pedido_cmp = _data_para_date_segura(dados_ciclo.get("dt_ped"))
    if data_pedido_cmp and data_pedido_cmp > data_hoje:
        pedidos_futuros.append({
            "Ciclo": f"C{dados_ciclo.get('numero')}",
            "Data do pedido": dados_ciclo.get("dt_ped").strftime("%d/%m/%Y"),
        })

if pedidos_futuros:
    st.error("Processamento inviável: há ciclo com data de pedido futura em relação à data atual.")
    st.warning("A Calculadora processa apenas pedidos já apresentados. Revise as datas dos pedidos antes de prosseguir.")
    st.dataframe(pd.DataFrame(pedidos_futuros), use_container_width=True, hide_index=True)
    st.stop()

# Segunda etapa: processamento dos ciclos, somente após o comando do usuário.
fator_acum = 1.0
historico = []
historico_coleta = []
pendencias_indice = []

for idx_ciclo, dados_ciclo in enumerate(input_ciclos):
    i = dados_ciclo['numero']
    data_atual = dados_ciclo['data_atual']
    d_fim = dados_ciclo['d_fim']
    d_aniv = dados_ciclo['d_aniv']
    d_lim = dados_ciclo['d_lim']
    dt_ped = dados_ciclo['dt_ped']
    sit_emoji = dados_ciclo['sit_emoji']
    situacao_limpa = dados_ciclo['situacao_limpa']
    inicio_efeito_financeiro = dados_ciclo['inicio_efeito_financeiro']
    ciclo_ja_concedido = bool(dados_ciclo.get('ciclo_ja_concedido', False))
    superacao_negocial = bool(dados_ciclo.get('superacao_negocial', False))
    percentual_negocial = float(dados_ciclo.get('percentual_negocial', 0.0) or 0.0)
    justificativa_negocial = dados_ciclo.get('justificativa_negocial', '')
    referencia_documental = dados_ciclo.get('referencia_documental', '')

    with containers_ciclos[idx_ciclo]:
        res_c = get_data_rep("433" if "IPCA" in idx_sel else "189", data_atual, d_fim, "IST" in idx_sel, "ICTI" in idx_sel)

        # Intervalo exibido independentemente de haver dados disponíveis para o índice.
        if res_c:
            periodo_inicio = res_c['p_ini']
            periodo_fim = res_c['p_fim']
            janela_ciclo = f"{res_c['p_ini'].strftime('%m/%Y')} a {res_c['p_fim'].strftime('%m/%Y')}"
        else:
            if "IST" in idx_sel:
                periodo_inicio = data_atual.replace(day=1)
                periodo_fim = (data_atual + relativedelta(years=1)).replace(day=1)
            else:
                periodo_inicio = data_atual
                periodo_fim = d_fim
            janela_ciclo = f"{periodo_inicio.strftime('%m/%Y')} a {periodo_fim.strftime('%m/%Y')}"

        janela_adm = f"{d_aniv.strftime('%d/%m/%Y')} a {d_lim.strftime('%d/%m/%Y')}"

        validacao_indice = _validar_indice_disponivel(res_c, data_atual, d_fim, idx_sel)

        st.markdown(f"""
        **Dados do Ciclo {i}:**
        - Intervalo do C{i}: {janela_ciclo}
        - Janela de Admissibilidade: {janela_adm}
        - Situação: {sit_emoji}
        """)

        if not validacao_indice.get("ok", False):
            _render_alerta_indice_ausente(validacao_indice, f"C{i}")
            pendencias_indice.append({
                "Ciclo": f"C{i}",
                "Intervalo do índice": janela_ciclo,
                "Competências faltantes": ", ".join(validacao_indice.get("faltantes", []) or []),
            })
            continue

        v_fmt = "Indisponível"
        v_acum_parcial = f"{(fator_acum - 1) * 100:,.2f}%".replace('.', ',')
        fator_ciclo = 1.0
        ciclo_calculado = False

        if res_c:
            fator_indice = 1 + res_c['var']
            percentual_indice = float(res_c['var'])
            ciclo_negativo = percentual_indice < 0
            percentual_aplicado = 0.0 if ciclo_negativo else percentual_indice
            situacao_aplicada = sit_emoji
            tratamento_negativo = "Ciclo negativo - percentual aplicado 0,00% no acumulado" if ciclo_negativo else ""
            if situacao_limpa == "PRECLUSO" and superacao_negocial:
                percentual_aplicado = percentual_negocial
                fator_ciclo = 1 + percentual_aplicado
                situacao_aplicada = "🟣 CICLO ADMITIDO POR NEGOCIAÇÃO ENTRE AS PARTES"
            elif situacao_limpa == "PRECLUSO":
                percentual_aplicado = 0.0
                fator_ciclo = 1.0
                if ciclo_negativo:
                    situacao_aplicada = f"{sit_emoji} | 🔻 CICLO NEGATIVO (APLICADO 0,00%)"
            elif ciclo_negativo:
                percentual_aplicado = 0.0
                fator_ciclo = 1.0
                situacao_aplicada = f"{sit_emoji} | 🔻 CICLO NEGATIVO (APLICADO 0,00%)"
            else:
                fator_ciclo = fator_indice
            fator_acum *= fator_ciclo
            v_fmt = f"{res_c['var'] * 100:,.2f}%".replace('.', ',')
            v_aplicado_fmt = f"{percentual_aplicado * 100:,.2f}%".replace('.', ',')
            v_acum_parcial = f"{(fator_acum - 1) * 100:,.2f}%".replace('.', ',')
            ciclo_calculado = True

            st.markdown(f"- Variação do Ciclo: **{v_fmt}**")
            if situacao_limpa == "PRECLUSO" and superacao_negocial:
                st.markdown(
                    f"""
                    <div style="
                        background:#F5F3FF;
                        border:1px solid #C4B5FD;
                        border-left:5px solid #7C3AED;
                        border-radius:10px;
                        padding:0.70rem 0.85rem;
                        margin:0.75rem 0 0.55rem 0;
                        color:#3B0764;
                        font-weight:750;
                    ">
                        Ciclo admitido por negociação entre as partes. Percentual aplicado: {v_aplicado_fmt}.
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            elif situacao_limpa == "PRECLUSO":
                st.caption("Variação apurada apenas para registro, sem composição no acumulado final.")
            elif ciclo_negativo:
                st.warning(
                    "A variação final apurada para este ciclo foi negativa. Para fins de composição acumulada, "
                    "o percentual aplicado neste ciclo será tratado como 0,00%. Meses negativos isolados dentro "
                    "do ciclo não são zerados; a regra somente se aplica quando o resultado final do ciclo é negativo."
                )

            with st.expander(f"🔍 Memória de Cálculo Detalhada - Ciclo {i}"):
                st.write(f"**Metodologia:** {res_c['metodo']}")
                st.write(f"**Janela de Apuração:** {res_c['p_ini'].strftime('%m/%Y')} a {res_c['p_fim'].strftime('%m/%Y')}")

                if "IST" in idx_sel:
                    st.write(
                        f"**Competência inicial:** {res_c['p_ini'].strftime('%m/%Y')} | "
                        f"**Índice inicial:** {res_c['i_ini']}"
                    )
                    st.write(
                        f"**Competência final:** {res_c['p_fim'].strftime('%m/%Y')} | "
                        f"**Índice final:** {res_c['i_fim']}"
                    )
                    _render_equacao_ist(res_c)
                elif "ICTI" in idx_sel:
                    st.write(f"**Competência da proposta/âncora:** {res_c['p_ini'].strftime('%m/%Y')}")
                    st.write(f"**Competência do índice-base utilizada:** {pd.to_datetime(res_c.get('d_indice_base')).strftime('%m/%Y')}")
                    st.write(f"**Competência final:** {res_c['p_fim'].strftime('%m/%Y')}")
                    st.dataframe(res_c['dados'], use_container_width=True)
                    st.write("Fórmula: produtório de (1 + taxa_mensal/100), com base no mês anterior à proposta/âncora.")
                else:
                    st.dataframe(res_c['dados'], use_container_width=True)
                    st.write("Fórmula: Produtório de (1 + taxa_mensal/100) - 1")

            if not ciclo_ja_concedido:
                historico.append({
                    "Ciclo": i,
                "Variação": v_fmt,
                "Percentual aplicado": v_aplicado_fmt,
                "Acumulada": v_acum_parcial,
                "Situação": situacao_aplicada,
                "Situação automática": sit_emoji,
                "Acordo negocial": bool(superacao_negocial),
                "Pedido": dt_ped.strftime('%d/%m/%Y'),
                "Janela": janela_ciclo,
                "JanelaAdm": janela_adm,
                    "Início financeiro": inicio_efeito_financeiro.strftime('%d/%m/%Y') if inicio_efeito_financeiro else "",
                    "Ciclo negativo": "Sim" if ciclo_negativo else "Não",
                })
            else:
                st.info(
                    f"C{i} foi marcado como já concedido/formalizado. O percentual foi preservado para memória e sequência lógica, "
                    "mas o ciclo não será tratado como objeto novo da análise atual."
                )
        else:
            percentual_indice = 0.0
            percentual_aplicado = 0.0
            ciclo_negativo = False
            tratamento_negativo = ""
            situacao_aplicada = sit_emoji
            st.warning(
                "Não há dados disponíveis para o índice selecionado no intervalo de apuração deste ciclo. "
                "O ciclo foi exibido para controle, mas não foi incluído no cálculo acumulado."
            )

        historico_coleta.append({
            'ciclo': f'C{i}',
            'data_base': data_atual.strftime('%d/%m/%Y'),
            'intervalo_indice': janela_ciclo,
            'janela_admissibilidade': janela_adm,
            'data_pedido': dt_ped.strftime('%d/%m/%Y'),
            'situacao': situacao_aplicada,
            'ciclo_ja_concedido': bool(ciclo_ja_concedido),
            'objeto_analise_atual': not bool(ciclo_ja_concedido),
            'situacao_automatica': sit_emoji,
            'situacao_aplicada': situacao_aplicada,
            'superacao_negocial': bool(superacao_negocial),
            'percentual_indice': float(percentual_indice),
            'percentual_aplicado': float(percentual_aplicado),
            'justificativa_negocial': justificativa_negocial.strip(),
            'referencia_documental': referencia_documental.strip(),
            'ciclo_negativo': bool(ciclo_negativo),
            'tratamento_ciclo_negativo': tratamento_negativo,
            'variacao': float(percentual_aplicado),
            'variacao_formatada': f"{percentual_aplicado * 100:,.2f}%".replace('.', ','),
            'fator': float(fator_ciclo),
            'fator_acumulado': float(fator_acum),
            'ciclo_calculado': ciclo_calculado,
            'financeiro_inicio': _formatar_data(inicio_efeito_financeiro),
            'financeiro_fim': '',
            'periodo_inicio': _formatar_data(periodo_inicio),
            'periodo_fim': _formatar_data(periodo_fim),
        })

if pendencias_indice:
    st.error("Processamento inviável: há competências ausentes em um ou mais ciclos. O cálculo acumulado e o Arquivo de Coleta foram bloqueados.")
    st.dataframe(pd.DataFrame(pendencias_indice), use_container_width=True, hide_index=True)
    st.stop()

# Finaliza os períodos financeiros com base no início financeiro do ciclo seguinte.
# Isso evita gerar 13 competências e separa o período de índice do período financeiro.
for idx, ciclo in enumerate(historico_coleta):
    inicio_txt = ciclo.get('financeiro_inicio', '')
    if not inicio_txt:
        ciclo['financeiro_fim'] = ''
        continue
    inicio_dt = pd.to_datetime(inicio_txt, dayfirst=True).to_pydatetime()
    proximo_inicio = None
    for prox in historico_coleta[idx + 1:]:
        prox_txt = prox.get('financeiro_inicio', '')
        if prox_txt:
            proximo_inicio = pd.to_datetime(prox_txt, dayfirst=True).to_pydatetime()
            break
    if proximo_inicio:
        fim_dt = proximo_inicio - relativedelta(months=1)
    else:
        fim_dt = inicio_dt + relativedelta(months=11)
    ciclo['financeiro_fim'] = fim_dt.strftime('%d/%m/%Y')

if historico:
    st.divider()
    res_final = f"{(fator_acum - 1) * 100:,.2f}%".replace('.', ',')
    st.metric("Variação Acumulada Final", res_final)

    st.subheader("Relatório de Apuração")
    corpo_relatorio = ""
    for h in historico:
        if h.get("Acordo negocial"):
            corpo_relatorio += f"""
            **C{h['Ciclo']}:** Pedido em {h['Pedido']}. Intervalo do C{h['Ciclo']}: {h['Janela']}.  
            Janela de Admissibilidade: {h['JanelaAdm']}.  
            **Resultado automático:** {h.get('Situação automática', h.get('Situação', ''))}.  
            **Tratamento aplicado:** (*) ciclo admitido por negociação entre as partes.  
            Variação apurada pelo índice: {h['Variação']}.  
            Percentual aplicado por acordo: {h.get('Percentual aplicado', h['Variação'])}.  
            Índice {idx_sel}.  
            Data de início dos efeitos financeiros por acordo: {h.get('Início financeiro', '')}.

            (*) O diagnóstico automático de preclusão foi preservado. O ciclo foi considerado aplicável por decisão negocial registrada pelo usuário.
            \n\n"""
        else:
            corpo_relatorio += f"""
            **C{h['Ciclo']}:** Pedido em {h['Pedido']}. Intervalo do C{h['Ciclo']}: {h['Janela']}.  
            Janela de Admissibilidade: {h['JanelaAdm']}.  
            Resultado: {h['Situação']}. Variação: {h['Variação']}.  
            Índice {idx_sel}.
            \n\n"""
    st.info(corpo_relatorio)

if historico_coleta:
    variacao_acumulada = fator_acum - 1
    st.session_state['dados_admissibilidade'] = {
        'origem': 'Reajustes Múltiplos',
        'tipo': 'Múltiplo',
        'indice': idx_sel,
        'data_base_original': dt_base_original.strftime('%d/%m/%Y'),
        'contexto_contratual_anterior': contexto_contratual,
        'fator': float(fator_acum),
        'fator_acumulado': float(fator_acum),
        'variacao_acumulada': float(variacao_acumulada),
        'variacao_acumulada_formatada': f"{variacao_acumulada * 100:,.2f}%".replace('.', ','),
        'ciclos': historico_coleta,
    }

    st.download_button(
        label="📥 Gerar Arquivo de Coleta",
        type="primary",
        data=gerar_arquivo_coleta_excel(st.session_state['dados_admissibilidade']),
        file_name="Coleta_Reajustes_Multiplos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )

    modelo_consumo = gerar_modelo_consumo_itens_ciclo_excel(st.session_state['dados_admissibilidade'])
    render_botao_download_modelo_consumo(modelo_consumo)
