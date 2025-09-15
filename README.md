# Pipeline de Processamento de Dados de Custos e Cotação de Moedas

## Sobre o Projeto

Este projeto é um pipeline de dados em Python que automatiza a conversão de custos de pedidos de compra para o Real (BRL). Ele foi desenhado para extrair dados de arquivos CSV, buscar cotações de câmbio de forma inteligente em duas APIs e gerar um relatório consolidado e detalhado em formato Excel.

## Funcionalidades Principais

- **Extração de Dados**: Importa dados de pedidos a partir de arquivos CSV, com tratamento de erros para garantir a robustez do processo.
- **Limpeza e Padronização**: Padroniza os dados de moedas, garantindo que inconsistências como `RMB` sejam convertidas para o padrão internacional (`CNY`), evitando falhas na comunicação com as APIs.
- **Cálculo de Preço Unitário**: Adiciona uma etapa de engenharia de dados para calcular o preço unitário de cada item, dividindo o valor total pela quantidade.
- **Busca de Cotação Inteligente**: O script prioriza a **API PTAX do Banco Central do Brasil** para as moedas principais. Para o Yuan Chinês (`CNY`), ele usa uma **cotação cruzada**, buscando o valor do Dólar (`USD`) em uma API de terceiros para garantir a conversão, mesmo quando a API oficial não tem os dados.
- **Geração de Relatórios Detalhados**: Cria um único arquivo Excel com duas abas essenciais:
    - `Últimos Preços`: Relatório final com o último preço de compra de cada material, com os valores convertidos para BRL.
    - `Histórico de Conversoes RMB-CNY`: Um log completo que registra os preços originais, o valor intermediário em Dólar e o valor final em Real, garantindo a rastreabilidade do processo de conversão.

## Como Usar o Projeto

### Pré-requisitos

Certifique-se de ter o Python instalado. Em seguida, instale as bibliotecas necessárias:

```bash
pip install pandas requests openpyxl
```

## Chave de API
O script utiliza a Free Currency API para a cotação cruzada do CNY. Para funcionar, você precisa de uma chave de API gratuita.

- **Acesse https://www.freecurrencyapi.com/ e crie sua conta.**
- **Copie sua chave de API no painel de controle.**
- **Cole-a no campo API_KEY do script.**

## Arquivos de Entrada
O script espera que os arquivos de dados estejam no mesmo diretório de execução:

- **cabecalho_pedido.csv**

- **item_pedido.csv**

## Execução
Basta rodar o script diretamente no terminal:

```bash
python pipeline.py
```

## Estrutura do Código
O código foi arquitetado como um pipeline, dividido em funções com responsabilidades claras para facilitar a manutenção e o entendimento:

- **importar_dados_pedidos()**: Trata da extração dos arquivos de origem.

- **buscar_cotacao_banco_central()**: Se comunica com as APIs para obter as cotações de câmbio.

- **processar_dados()**: Filtra o último preço de cada material e realiza as conversões.

- **gerar_relatorio()**: Formata os dados finais e os salva no arquivo Excel.

- **main()**: A função principal que orquestra a chamada de todas as outras funções na ordem correta, garantindo que o fluxo de trabalho seja claro e sequencial.
