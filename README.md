# ETL Cotação do Dólar

Este projeto realiza a extração da cotação do dólar dos últimos dias usando a API AwesomeAPI, transforma os dados e atualiza uma planilha Excel e atualiza gráficos.

## Etapas do ETL
- **Extrair**: via API pública
- **Transformar**: DataFrame com formatação e tipos adequados
- **Carregar**: Atualiza uma planilha XLSX com gráfico

## Requisitos
- Python 3.9+
- Bibliotecas: pandas, requests, openpyxl

## Execução
```bash
python etl_dolar.py