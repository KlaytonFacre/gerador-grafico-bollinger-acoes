from datetime import date
from openpyxl import Workbook


acao = input('Qual o código da ação a ser processada? ').upper()

with open(f'./assets/dados/{acao}.txt', 'r') as arquivo_cotacao:
    linhas = arquivo_cotacao.readlines()
    linhas = [linha.removesuffix('\n').split(';') for linha in linhas]

workbook = Workbook()
planilha_ativa = workbook.active

# PREPARANDO A PLANILHA DE 'DADOS'
planilha_ativa.title = 'Dados'

planilha_ativa.append(["DATA", "COTAÇÃO", "BANDA INFERIOR", "BANDA SUPERIOR"])

indice = 2
for linha in linhas:
    # TRATANDO E PREPARANDO A DATA, CRIANDO UM OBJETO DATE DO PYTHON
    ano_mes_dia = linha[0].split(" ")[0]
    data = date(
        year=int(ano_mes_dia.split("-")[0]),
        month=int(ano_mes_dia.split("-")[1]),
        day=int(ano_mes_dia.split("-")[2])
    )

    # TRATANDO E PREPARANDO A COTAÇÃO
    cotacao = float(linha[1])

    # ESCREVENDO DADOS NA PLANILHA ATIVA
    # Data
    planilha_ativa[f'A{indice}'] = data
    # Cotação
    planilha_ativa[f'B{indice}'] = cotacao
    # Banda inferior de Bollinger
    planilha_ativa[f'C{indice}'] = f'=AVERAGE(B{indice}:B{indice + 19}) - 2*STDEV(B{indice}:B{indice + 19})'
    # Banda superior de Bollinger
    planilha_ativa[f'D{indice}'] = f'=AVERAGE(B{indice}:B{indice + 19}) + 2*STDEV(B{indice}:B{indice + 19})'

    # PREPARA PARA A ESCRITA NA PRÓXIMA LINHA
    indice += 1


workbook.save('./saida/Planilha.xlsx')
