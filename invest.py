import yfinance as yf
from tabulate import tabulate
import pandas as pd

# Caminho para o arquivo Excel local
caminho_arquivo = "./planilha.xlsx"

def resgatarDadosDaPlanilha():
    try:
        # Ler o arquivo Excel
        df = pd.read_excel(caminho_arquivo, engine='openpyxl')

        # Acessar a primeira coluna a partir da 3ª linha
        coluna_desejada = df.iloc[2:, 0].reset_index(drop=True)

        # Converter a coluna em uma lista de strings
        lista_acoes = coluna_desejada.tolist()
        lista_acoes = lista_acoes[:-2]
        # Exibir a lista de ações
        print(lista_acoes)

        return lista_acoes
    except FileNotFoundError:
        print("Erro: Arquivo não encontrado. Verifique o caminho.")
    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")

lista_acoes = resgatarDadosDaPlanilha()     # ['ITUB4', 'PETR4', 'PETR3', 'BBAS3', 'ELET3', 'SBSP3', 'WEGE3', 'BBDC4', 'B3SA3', 'ABEV3', 'ITSA4', 'EMBR3', 'BPAC11', 'EQTL3', 'SUZB3', 'JBSS3', 'RDOR3', 'PRIO3', 'RENT3', 'BBSE3', 'ENEV3', 'RADL3', 'GGBR4', 'CMIG4', 'RAIL3', 'TOTS3', 'VIVT3', 'UGPA3', 'VBBR3', 'CPLE6', 'BBDC3', 'KLBN11', 'BRFS3', 'TIMS3', 'ENGI11', 'LREN3', 'CCRO3', 'STBP3', 'ELET6', 'NTCO3', 'HAPV3', 'EGIE3', 'ISAE4', 'ASAI3', 'SANB11', 'ALOS3', 'CMIN3', 'BRAV3', 'CSAN3', 'CXSE3', 'TAEE11', 'HYPE3', 'PSSA3', 'CPFE3', 'MULT3', 'CSNA3', 'GOAU4', 'CYRE3', 'FLRY3', 'POMO4', 'RECV3', 'BRAP4', 'MRFG3', 'CRFB3', 'IRBR3', 'IGTI11', 'AZZA3', 'SLCE3', 'USIM5', 'BRKM5', 'YDUQ3', 'COGN3', 'SMTO3', 'AURE3', 'MGLU3', 'VIVA3', 'RAIZ4', 'VAMO3', 'MRVE3', 'AZUL4', 'PCAR3', 'PETZ3', 'BEEF3', 'LWSA3', 'CVCB3', 'AMOB3']

soma_desvio = 0.0

lista_dy = []

lista_desvio_padrao = []

lista_setor = []


def getCota(code):
    acao = yf.Ticker(code+".SA")
    info = acao.info
    global soma_desvio
    global lista_dy
    global lista_desvio_padrao
    global lista_setor

    # Obtendo dados históricos dos últimos 1 ano
    dados_historicos = acao.history(period="1y")  # must be of the format 1d, 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max, etc.

    desvio_padrao = (dados_historicos['Close'].pct_change().std()) * 100

    soma_desvio += desvio_padrao
    lista_dy.append(info.get('dividendYield', 0))
    lista_desvio_padrao.append(float(desvio_padrao))
    lista_setor.append(info.get('sector', 'N/A'))

    

    # Tabela
    return [code, info.get('longName', 'N/A'), info.get('sector', 'N/A'), info.get('industry', 'N/A'), info.get('currentPrice', 'N/A'), info.get('dividendYield', 'N/A'), desvio_padrao]
    


dados = [getCota(acao) for acao in lista_acoes]
cabecalho = ["Código", "Nome", "Setor", "Indústria" "Preço", "DY", "Desvio Padrão"]
print(tabulate(dados, headers=cabecalho, tablefmt="grid"))





print("\n\n lista_dy \n")
print(lista_dy)
print("\n\n lista_desvio_padrao \n")
print(lista_desvio_padrao)
print("\n\n lista_setor \n")
print(lista_setor)

print(f"Média de desvio padrão da carteira: {soma_desvio/len(lista_acoes)}")