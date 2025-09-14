import pandas as pd
import requests
import openpyxl
from datetime import datetime, timedelta

API_KEY = 'fca_live_O10HzmUmIroLHQvChSlShMANNEwPBf95jRYgLjSH'

def importar_dados_pedidos(caminho_cabecalho_pedido, caminho_item_pedido):
    """
    Importa os dados de cabeçalho e itens de pedidos a partir de arquivos CSV.
    Retorna os DataFrames ou None em caso de erro.
    """
    try:
        print("Processo de importação iniciado...")
        df_cabecalho = pd.read_csv(caminho_cabecalho_pedido)
        df_item = pd.read_csv(caminho_item_pedido)
        print("Arquivos CSV importados com sucesso!")
        return df_cabecalho, df_item
    except FileNotFoundError as e:
        print(f"Erro: Arquivo não encontrado. Verifique o caminho. {e}")
        return None, None
    except Exception as e:
        print(f"Ocorreu um erro ao importar os arquivos: {e}")
        return None, None

def buscar_cotacao_banco_central(simbolo_moeda, api_key):
    """
    Busca a cotação de venda de uma moeda estrangeira em BRL (REAL) na API PTAX.
    Procura a cotação do dia atual e, se não encontrar, retrocede até o último dia útil.
    Se a moeda for CNY e a busca inicial falhar, usa uma cotação cruzada com o USD.
    """
    data_busca = datetime.now()
    dias_maximos = 7
    
    for _ in range(dias_maximos):
        data_formatada = data_busca.strftime('%m-%d-%Y')
        url = (f"https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
               f"CotacaoMoedaPeriodo(moeda='{simbolo_moeda}',dataInicial='{data_formatada}',dataFinalCotacao='{data_formatada}')"
               f"?$top=1&$orderby=dataHoraCotacao desc&$format=json")
        try:
            response = requests.get(url)
            data = response.json()
            if 'value' in data and len(data['value']) > 0:
                cotacao_venda = data['value'][0]['cotacaoVenda']
                data_cotacao = data['value'][0]['dataHoraCotacao']
                print(f"Cotação encontrada para {simbolo_moeda} na data {data_formatada}.")
                return cotacao_venda, data_cotacao
            else:
                print(f"Nenhuma cotação para {simbolo_moeda} encontrada em {data_formatada}. Tentando dia anterior...")
                data_busca -= timedelta(days=1)
        except Exception as e:
            print(f"Ocorreu um erro inesperado ao buscar a cotação para {simbolo_moeda}: {e}")
            return None, None
    
    if simbolo_moeda == 'CNY':
        print("Tentando obter a cotação cruzada para CNY via USD...")
        try:
            cotacao_usd_brl, _ = buscar_cotacao_banco_central('USD', api_key)
            
            if cotacao_usd_brl is not None:
                url_cruzada = f"https://api.freecurrencyapi.com/v1/latest?apikey={api_key}&base_currency=USD&currencies=CNY"
                response_cruzada = requests.get(url_cruzada)
                data_cruzada = response_cruzada.json()
                cotacao_usd_cny = data_cruzada['data']['CNY']
                
                cotacao_final_cny = cotacao_usd_cny * cotacao_usd_brl
                data_hoje = datetime.now().strftime('%m-%d-%Y')
                print(f"Cotação cruzada encontrada para CNY: {cotacao_final_cny}")
                return cotacao_final_cny, data_hoje
            else:
                print("Não foi possível obter a cotação de USD para realizar a cotação cruzada.")
                return None, None
        except Exception as e:
            print(f"Erro ao buscar cotação cruzada para CNY: {e}")
            return None, None
    
    print(f"Não foi possível encontrar a cotação para a moeda {simbolo_moeda} nos últimos {dias_maximos} dias.")
    return None, None

def processar_dados(df_historico, cotacoes):
    """
    Processa os dados para encontrar o último preço de compra e converter para BRL.
    """
    print("Processamento dos dados iniciado...")
    df_ultimo_preco = df_historico.sort_values(by='data_compra', ascending=False).drop_duplicates(subset='codigo_material', keep='first')

    def converter_preco(row):
        moeda = row['moeda']
        if moeda in cotacoes and cotacoes[moeda]['valor'] is not None:
            return row['preco'] * cotacoes[moeda]['valor']
        return row['preco']

    df_ultimo_preco['preco_convertido_brl'] = df_ultimo_preco.apply(converter_preco, axis=1)

    def obter_data_cotacao(row):
        moeda = row['moeda']
        if moeda in cotacoes and cotacoes[moeda]['data'] is not None:
            return cotacoes[moeda]['data']
        return None

    df_ultimo_preco['data_cotacao'] = df_ultimo_preco.apply(obter_data_cotacao, axis=1)

    return df_ultimo_preco

def gerar_relatorio(df_relatorio_principal, df_conversoes_rmb, cotacoes):
    """
    Gera o arquivo Excel com os relatórios em abas separadas,
    incluindo a conversão na aba de histórico.
    """
    print("Gerando relatório final...")
    
    # Processa e adiciona as colunas de preço convertido ao histórico de conversões
    if not df_conversoes_rmb.empty:
        cotacao_cny_em_brl = cotacoes.get('CNY', {}).get('valor')
        
        # Converte o preço original (CNY) para USD
        if 'USD' in cotacoes and cotacoes['USD']['valor'] is not None:
            cotacao_usd_brl = cotacoes['USD']['valor']
            cotacao_cny_brl = cotacao_cny_em_brl
            
            # Checa se a cotação de CNY para BRL é válida antes de tentar a conversão para USD
            if cotacao_cny_brl and cotacao_usd_brl:
                cotacao_cny_usd = cotacao_cny_brl / cotacao_usd_brl
                df_conversoes_rmb['Preço em USD'] = df_conversoes_rmb['preco'] / cotacao_cny_usd
            else:
                df_conversoes_rmb['Preço em USD'] = None
        else:
            df_conversoes_rmb['Preço em USD'] = None
        
        # Converte o preço original (CNY) para BRL
        if cotacao_cny_em_brl is not None:
            df_conversoes_rmb['Preço em BRL'] = df_conversoes_rmb['preco'] * cotacao_cny_em_brl
        else:
            df_conversoes_rmb['Preço em BRL'] = None

    df_relatorio_principal = df_relatorio_principal.rename(columns={
        'preco_convertido_brl': 'Último preço de compra convertido em BRL',
        'preco': 'Último preço de compra sem conversão',
        'moeda': 'Moeda do pedido',
        'moeda_original': 'Moeda do pedido (Original)',
        'data_compra': 'Data da última compra',
        'codigo_material': 'Código do material',
        'codigo_pedido': 'Código do pedido de referência',
        'data_cotacao': 'Data da cotação considerada (se houver)'
    })[[
        'Código do material',
        'Último preço de compra convertido em BRL',
        'Último preço de compra sem conversão',
        'Moeda do pedido',
        'Moeda do pedido (Original)',
        'Data da última compra',
        'Código do pedido de referência',
        'Data da cotação considerada (se houver)'
    ]]
    
    output_path = 'relatorio_butantan_consolidados.xlsx'
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_relatorio_principal.to_excel(writer, sheet_name='Últimos Preços', index=False)
        if not df_conversoes_rmb.empty:
            df_conversoes_rmb.to_excel(writer, sheet_name='Histórico de Conversoes RMB-CNY', index=False)

    print(f"Relatórios gerados com sucesso no arquivo: {output_path}")

def main():
    """
    Função principal que orquestra todo o pipeline de processamento de dados.
    """
    # 1. Importar dados
    caminho_cabecalho_pedido = 'cabecalho_pedido.csv'
    caminho_item_pedido = 'item_pedido.csv'
    df_cabecalho, df_item = importar_dados_pedidos(caminho_cabecalho_pedido, caminho_item_pedido)

    if df_cabecalho is None or df_item is None:
        print("Não foi possível prosseguir com o pipeline devido a um erro na importação.")
        return

    # 2. Pré-processar e limpar os dados
    df_historico = pd.merge(df_item, df_cabecalho, on='codigo_pedido', how='left')
    df_historico = df_historico.rename(columns={
        'data_pedido': 'data_compra',
        'valor_total_item_pedido': 'preco'
    })
    df_historico['moeda'] = df_historico['moeda'].str.strip().str.upper()
    df_historico['moeda_original'] = df_historico['moeda']
    df_conversoes_rmb = df_historico[df_historico['moeda'] == 'RMB'].copy()
    if not df_conversoes_rmb.empty:
        df_conversoes_rmb['Moeda Convertida'] = 'CNY'
        df_conversoes_rmb['Observacao'] = 'Conversão realizada para buscar cotação na API do BCB.'
    df_historico['moeda'] = df_historico['moeda'].replace('RMB', 'CNY')
    df_historico['data_compra'] = pd.to_datetime(df_historico['data_compra'])

    # 3. Extrair cotações das moedas via API
    print("Processo de busca dos valores de cotação iniciado...")
    cotacoes = {'BRL': {'valor': 1.0, 'data': None}}
    moedas_estrangeiras = df_historico[df_historico['moeda'] != 'BRL']['moeda'].unique()
    for moeda in moedas_estrangeiras:
        valor, data = buscar_cotacao_banco_central(moeda, API_KEY)
        if valor is not None:
            cotacoes[moeda] = {'valor': valor, 'data': data}

    # 4. Processar e converter preços
    df_relatorio_principal = processar_dados(df_historico, cotacoes)

    # 5. Gerar relatórios
    gerar_relatorio(df_relatorio_principal, df_conversoes_rmb, cotacoes)

if __name__ == "__main__":
    main()