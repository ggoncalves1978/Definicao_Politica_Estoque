# Definicao_Politica_Estoque

Este repositório contém um script Python para cálculo de níveis de estoque, ponto de pedido, e estoque de segurança. Ele utiliza dados históricos de consumo para gerar métricas essenciais de planejamento de demanda, com base em parâmetros como lead time, curva ABC e níveis de serviço.

## Funcionalidades

O script realiza as seguintes operações:

- Importação de dados de um arquivo Excel contendo o histórico de consumo e outros parâmetros.
- Cálculo de médias móveis de consumo para diferentes períodos (24 meses, 12 meses, e 6 meses).
- Cálculo do Consumo Médio Diário (CMD).
- Cálculo do desvio padrão do consumo (com base em 24 meses de dados).
- Cálculo da demanda durante o lead time (DL).
- Cálculo do nível de serviço (NS) com base na curva ABC.
- Cálculo do estoque de segurança (ES) e ponto de pedido (PP).
- Cálculo do estoque máximo (Emax).
- Conversão dos níveis de estoque de unidades para dias de cobertura.
- Cálculo do valor de estoque máximo, considerando o Custo Médio Unitário (CMU).
- Exportação dos resultados para um arquivo Excel.

## Requisitos

Antes de executar o script, certifique-se de ter as seguintes dependências instaladas:

- **Python 3.x**
- **pandas**
- **numpy**
- **scipy**
- **plotly**
- **openpyxl** (para manipulação de arquivos Excel)

Você pode instalar as dependências executando o seguinte comando:

```bash
pip install pandas numpy scipy plotly openpyxl

