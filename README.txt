Atualizacao: Linha do Tempo do Contrato com Data-base para reajuste.

Arquivo alterado:
- pages\03_Valor_Global.py

O evento inicial "Data-base para reajuste" passa a aparecer na timeline do site e no Relatorio Executivo em PDF da timeline.

Aplicacao:
1. Execute atualizar_homologacao_timeline_data_base.bat; ou
2. Copie pages\03_Valor_Global.py para C:\_DesktopReal\08.clausula\pages.

Depois rode:
cd /d C:\_DesktopReal\08.clausula
.venv\Scripts\activate
streamlit run app.py

Depois de validar, execute criar_ponto_restauracao_homologacao_cl8us.bat para criar backup local da homologacao.
