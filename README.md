# Orçamentador Dinâmico SINAPI-CE + SEINFRA-CE | Mais Engenharia

Sistema em Streamlit para criação, importação e conversão de orçamentos de obras em estrutura analítica EAP/WBS, integrado às bases oficiais **SINAPI-CE** e **SEINFRA-CE (Unificada 028)**.

## Como executar o sistema

```powershell
streamlit run app2.py
```

## Conversor via Linha de Comando (Opcional)

Além de importar e converter diretamente pela interface do sistema (`app2.py`), você também pode converter arquivos `.pdf`, `.xls`, `.xlsx` ou `.csv` em planilhas Excel `.xlsx` limpas pelo terminal:

```powershell
# Converter orçamento de terceiro (.pdf, .xls, .csv, .xlsx) para Excel (.xlsx)
python converter_orcamento.py caminho_do_orcamento.pdf -o orcamento_convertido.xlsx

# Converter uma nova tabela SEINFRA (.xls) para seinfra_ce.csv
python converter_orcamento.py Planos-de-Servicos-028.xls --seinfra -o seinfra_ce.csv
```
