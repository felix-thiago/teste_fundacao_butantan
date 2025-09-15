import pandas as pd 
import requests 
import openpyxl 
from datetime import datetime, timedelta

# KEY SECRET para acessar a API de moedas da Free currency api.
API_KEY = 'fca_live_O10HzmUmIroLHQvChSlShMANNEwPBf95jRYgLjSH' # minha chave, o ideal seria armazenar ela em um secret manager ou algo do tipo em nuvem

def importar_dados_pedidos(caminho_cabecalho_pedido, caminho_item_pedido):
    """
    Este é o primeiro passo do projeto: trazer os dados para dentro do script.
    É realizada a tentativa de importar os arquivos e, se algo der errado (se o arquivo não estiver lá, por exemplo),
    o código não quebra. Ele nos avisa o que aconteceu.
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
    A função mais importante do projeto! Ela vai até a API do Banco Central e nos traz o valor de cada moeda.
    Está projetada para contingências: primeiro busca a cotação oficial do Banco Central.
    Se não encontrar (porque é fim de semana, feriado, ou a moeda não é oficial), ela tem um plano B.
    """
    # A estratégia para não falhar nos fins de semana.
    data_busca = datetime.now()
    dias_maximos = 7 # Quantos dias vai retroceder na busca.

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

    # Se a moeda for CNY e o PTAX não tiver registros...
    if simbolo_moeda == 'CNY':
        print("Tentando obter a cotação cruzada para CNY via USD...")
        
        try:
            # 1. Pegar o valor do Dólar em Real, que sempre tem no PTAX.
            cotacao_usd_brl, _ = buscar_cotacao_banco_central('USD', api_key)
            
            if cotacao_usd_brl is not None:
                # 2. Coletar o valor do Yuan em Dólar para a outra API.
                url_cruzada = f"https://api.freecurrencyapi.com/v1/latest?apikey={api_key}&base_currency=USD&currencies=CNY"
                response_cruzada = requests.get(url_cruzada)
                data_cruzada = response_cruzada.json()
                cotacao_usd_cny = data_cruzada['data']['CNY']
                
                # 3. Multiplicar BRL/USD por USD/CNY para achar BRL/CNY.
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
    Aqui é onde os dados se transformam! É realizada a junção das cotações que foi coletada com os dados brutos
    e calculado o preço final.
    """
    print("Processamento dos dados iniciado...")
    
    # Primeiro, é localizada a compra mais recente de cada material para evitar dados duplicados.
    df_ultimo_preco = df_historico.sort_values(by='data_compra', ascending=False).drop_duplicates(subset='codigo_material', keep='first')

    # Esta função faz a conversão de fato. Se a moeda for estrangeira, ela usa a cotação que foi guardada.
    def converter_preco(row):
        moeda = row['moeda']
        if moeda in cotacoes and cotacoes[moeda]['valor'] is not None:
            return row['preco_unitario'] * cotacoes[moeda]['valor']
        return row['preco_unitario']

    # Aplicada a função de conversão em cada linha do dataframe.
    df_ultimo_preco['preco_convertido_brl'] = df_ultimo_preco.apply(converter_preco, axis=1)

    # Computada a data da cotação para deixar o relatório mais transparente.
    def obter_data_cotacao(row):
        moeda = row['moeda']
        if moeda in cotacoes and cotacoes[moeda]['data'] is not None:
            return cotacoes[moeda]['data']
        return None

    df_ultimo_preco['data_cotacao'] = df_ultimo_preco.apply(obter_data_cotacao, axis=1)

    return df_ultimo_preco

def gerar_relatorio(df_relatorio_principal, df_conversoes_rmb, cotacoes):
    """
    A etapa final! Compilação de tudo em um arquivo Excel limpo e organizado.
    Criado duas abas para que seja visualizado o resultado final e também
    o histórico detalhado das conversões mais complexas.
    """
    print("Gerando relatório final...")
    
    # Preparado a aba de log, calculado os valores em dólar e real para mostrar o "antes e depois".
    if not df_conversoes_rmb.empty:
        cotacao_cny_em_brl = cotacoes.get('CNY', {}).get('valor')
        
        if cotacao_cny_em_brl is not None:
            if 'USD' in cotacoes and cotacoes['USD']['valor'] is not None:
                cotacao_usd_brl = cotacoes['USD']['valor']
                cotacao_cny_usd = cotacao_cny_em_brl / cotacao_usd_brl
                df_conversoes_rmb['Preço unitário em USD'] = df_conversoes_rmb['preco_unitario'] * cotacao_cny_usd
            else:
                df_conversoes_rmb['Preço unitário em USD'] = None
            
            df_conversoes_rmb['Preço unitário em BRL'] = df_conversoes_rmb['preco_unitario'] * cotacao_cny_em_brl
        else:
            df_conversoes_rmb['Preço unitário em USD'] = None
            df_conversoes_rmb['Preço unitário em BRL'] = None

    # Renomeação das colunas para que o relatório esteja de acordo com o solicitado e fique mais fácil de ler.
    df_relatorio_principal = df_relatorio_principal.rename(columns={
        'preco_convertido_brl': 'Último preço de compra convertido em BRL',
        'preco_unitario': 'Último preço de compra unitário',
        'moeda': 'Moeda do pedido',
        'moeda_original': 'Moeda do pedido (Original)',
        'data_compra': 'Data da última compra',
        'codigo_material': 'Código do material',
        'codigo_pedido': 'Código do pedido de referência',
        'data_cotacao': 'Data da cotação considerada (se houver)'
    })[[
        'Código do material',
        'Último preço de compra convertido em BRL',
        'Último preço de compra unitário',
        'Moeda do pedido',
        'Moeda do pedido (Original)',
        'Data da última compra',
        'Código do pedido de referência',
        'Data da cotação considerada (se houver)'
    ]]
    
    output_path = 'relatorio_butantan_consolidados.xlsx'
    
    # Salvando o relatório com as duas abas.
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_relatorio_principal.to_excel(writer, sheet_name='Últimos Preços', index=False)
        if not df_conversoes_rmb.empty:
            df_conversoes_rmb.to_excel(writer, sheet_name='Histórico de Conversoes RMB-CNY', index=False)

    print(f"Relatórios gerados com sucesso no arquivo: {output_path}")

def main():
    """
    Esta é a função que orquestra todo o projeto, chamando cada etapa na ordem certa.
    """
    # Passo 1: Trazer os dados para dentro
    caminho_cabecalho_pedido = 'cabecalho_pedido.csv'
    caminho_item_pedido = 'item_pedido.csv'
    df_cabecalho, df_item = importar_dados_pedidos(caminho_cabecalho_pedido, caminho_item_pedido)

    if df_cabecalho is None or df_item is None:
        print("Não foi possível prosseguir com o pipeline devido a um erro na importação.")
        return

    # Passo 2: Preparar e Limpar os Dados
    df_historico = pd.merge(df_item, df_cabecalho, on='codigo_pedido', how='left')
    df_historico = df_historico.rename(columns={
        'data_pedido': 'data_compra',
        'valor_total_item_pedido': 'preco_total'
    })
    
    # Calculado o preço unitário para a análise ser precisa!
    df_historico['preco_unitario'] = df_historico['preco_total'] / df_historico['item_quantidade']

    df_historico['moeda'] = df_historico['moeda'].str.strip().str.upper()
    df_historico['moeda_original'] = df_historico['moeda']
    df_conversoes_rmb = df_historico[df_historico['moeda'] == 'RMB'].copy()
    if not df_conversoes_rmb.empty:
        df_conversoes_rmb['Moeda Convertida'] = 'CNY'
        df_conversoes_rmb['Observacao'] = 'Conversão realizada para buscar cotação na API do BCB.'
    df_historico['moeda'] = df_historico['moeda'].replace('RMB', 'CNY')
    df_historico['data_compra'] = pd.to_datetime(df_historico['data_compra'])

    # Passo 3: Buscar as cotações das moedas
    print("Processo de busca dos valores de cotação iniciado...")
    cotacoes = {'BRL': {'valor': 1.0, 'data': None}}
    moedas_estrangeiras = df_historico[df_historico['moeda'] != 'BRL']['moeda'].unique()
    for moeda in moedas_estrangeiras:
        valor, data = buscar_cotacao_banco_central(moeda, API_KEY)
        if valor is not None:
            cotacoes[moeda] = {'valor': valor, 'data': data}

    # Passo 4: Processar e converter os preços
    df_relatorio_principal = processar_dados(df_historico, cotacoes)

    # Passo 5: Gerar os relatórios
    gerar_relatorio(df_relatorio_principal, df_conversoes_rmb, cotacoes)

if __name__ == "__main__":
    main()