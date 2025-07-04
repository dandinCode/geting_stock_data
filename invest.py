import yfinance as yf
from tabulate import tabulate
import pandas as pd

# Caminho para o arquivo Excel local
caminho_arquivo = "./planilha.xlsx"

def resgatarDadosDaPlanilha():
    try:
        df = pd.read_excel(caminho_arquivo, engine='openpyxl')
        coluna_desejada = df.iloc[2:, 0].reset_index(drop=True)
        lista_acoes = coluna_desejada.tolist()
        lista_acoes = lista_acoes[:-2]
        return lista_acoes
    except FileNotFoundError:
        print("Erro: Arquivo não encontrado. Verifique o caminho.")
    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")

lista_acoes = resgatarDadosDaPlanilha()  # ['PETR4', 'ITUB4', 'PETR3', 'BBAS3', 'ELET3', 'SBSP3', 'WEGE3', 'BBDC4', 'B3SA3', 'ABEV3', 'ITSA4', 'EMBR3', 'BPAC11', 'EQTL3', 'SUZB3', 'JBSS3', 'RDOR3', 'PRIO3', 'RENT3', 'BBSE3', 'ENEV3', 'RADL3', 'GGBR4', 'CMIG4', 'RAIL3', 'TOTS3', 'VIVT3', 'UGPA3', 'VBBR3', 'CPLE6', 'BBDC3', 'KLBN11', 'BRFS3', 'TIMS3', 'ENGI11', 'LREN3', 'CCRO3', 'STBP3', 'ELET6', 'NTCO3', 'HAPV3', 'EGIE3', 'ISAE4', 'ASAI3', 'SANB11', 'ALOS3', 'CMIN3', 'BRAV3', 'CSAN3', 'CXSE3', 'TAEE11', 'HYPE3', 'PSSA3', 'CPFE3', 'MULT3', 'CSNA3', 'GOAU4', 'CYRE3', 'FLRY3', 'POMO4', 'RECV3', 'BRAP4', 'MRFG3', 'CRFB3', 'IRBR3', 'IGTI11', 'AZZA3', 'SLCE3', 'USIM5', 'BRKM5', 'YDUQ3', 'COGN3', 'SMTO3', 'AURE3', 'MGLU3', 'VIVA3', 'RAIZ4', 'VAMO3', 'MRVE3', 'AZUL4', 'PCAR3', 'PETZ3', 'BEEF3', 'LWSA3', 'CVCB3', 'AMOB3']

soma_desvio = 0.0
lista_dy = []
lista_desvio_padrao = []
lista_setor = []

def getCota(code):
    ticker = code + ".SA"
    acao = yf.Ticker(ticker)
    global soma_desvio, lista_dy, lista_desvio_padrao, lista_setor

    try:
        # Baixar dados históricos de preços para 2023
        dados = yf.download(ticker, start='2023-01-01', end='2024-01-02', progress=False)

        if dados.empty:
            raise ValueError("Dados históricos vazios.")

        # Calcular desvio padrão anualizado
        dados['Retorno'] = dados['Close'].pct_change()
        desvio_padrao_diario = dados['Retorno'].std()
        desvio_padrao_anual = desvio_padrao_diario * (252 ** 0.5)

        # Calcular dividendos pagos no período
        dividendos = acao.dividends.loc['2023-01-01':'2024-01-01']
        total_dividendos = dividendos.sum()

        # Preço final do período (próximo útil a 01/01/2024)
        preco_final = float(dados['Close'].iloc[-1])

        # Calcular Dividend Yield baseado na cotação do final do período
        dividend_yield_medio = (total_dividendos / preco_final) * 100 if preco_final > 0 else 0.0

        info = acao.info
        nome = info.get('longName', 'N/A')
        setor = info.get('sector', 'N/A')
        industria = info.get('industry', 'N/A')

        # Armazenar para médias
        soma_desvio += desvio_padrao_anual
        lista_dy.append(dividend_yield_medio)
        lista_desvio_padrao.append(desvio_padrao_anual)
        lista_setor.append(setor)

        return [code, nome, setor, industria, f"{dividend_yield_medio:.2f}%", f"{desvio_padrao_anual:.2f}%"]

    except Exception as e:
        print(f"Erro ao processar {code}: {e}")
        return [code, 'Erro', 'Erro', 'Erro', 'Erro', 'Erro']

# Processar todas as ações
dados = [getCota(acao) for acao in lista_acoes]

# Mostrar tabela
cabecalho = ["Código", "Nome", "Setor", "Indústria", "DY (2023)", "Desvio Padrão Anual"]
print(tabulate(dados, headers=cabecalho, tablefmt="grid"))

# Exibir listas auxiliares
print("\n\n Lista de DY\n", [float(valor) for valor in lista_dy])
print("\n Lista de Desvios Padrão\n", [float(valor) for valor in lista_desvio_padrao])
print("\n Lista de Setores\n", lista_setor)
print(f"\n Média do Desvio Padrão Anualizado da Carteira: {soma_desvio / len(lista_desvio_padrao):.2f}%")
