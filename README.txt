CL8US 2.0 - CALCULADORA DE REAJUSTES

Status: linha oficial retomada em 13/07/2026.
Aplicacao: https://reajustes.streamlit.app/
Repositorio: https://github.com/dlpenalva/Painel1
Entrada principal: app.py

DECISAO DE ARQUITETURA

O Streamlit conduz a admissibilidade, a escolha do indice e a geracao da
coleta. O XLS continua sendo o artefato operacional de calculo, conferencia e
integracao com a fiscalizacao. A retomada do 2.0 nao depende do projeto 3.0.

EXECUCAO LOCAL

1. Crie e ative um ambiente virtual.
2. Instale as dependencias: pip install -r requirements.txt
3. Execute: streamlit run app.py
4. Verifique: http://localhost:8501/_stcore/health

VALIDACAO MINIMA

python -m compileall -q app.py _ui_utils.py _reajuste_utils.py _indice_utils.py pages
python -m unittest discover -s tests -v

O plano tecnico e os criterios de seguranca da retomada estao em
docs/RETOMADA_V2.md.
