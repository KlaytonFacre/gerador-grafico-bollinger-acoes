from datetime import date
from openpyxl import Workbook
from openpyxl.styles import *
from openpyxl.chart import *
from openpyxl.


# acao = input('Qual o código da ação a ser processada? ').upper()
# Setando a variável temporariamente apenas para testes e agilizar a execução do programa
acao = "BIDI4"

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

planilha_grafico = workbook.create_sheet("Gráfico")
Workbook.active = planilha_grafico

# Mesclagem de células e criação do cabeçalho do gráfico
planilha_grafico.merge_cells("A1:T2")
cabecalho = planilha_grafico["A1"]
cabecalho.font = Font(b=True, sz=18, color="FFFFFF")
cabecalho.fill = PatternFill("solid", fgColor="07838f")
cabecalho.alignment = Alignment(vertical="center", horizontal="center")
cabecalho.value = "Histórico de Cotações"

# Criação do Gráfico
grafico = LineChart()
grafico.width = 33.95
grafico.height = 15.5
grafico.title = f"Cotações - {acao}"
grafico.x_axis.title = "Data da Cotação"
grafico.y_axis.title = "Valor da Cotação"

# Criando referencias e adicionando elas como séries de dados e datas
referencia_cotacoes = Reference(planilha_ativa, min_col=2, min_row=2, max_col=4, max_row=indice)
referencia_datas = Reference(planilha_ativa, min_col=1, min_row=2, max_col=1, max_row=indice)
grafico.add_data(referencia_cotacoes)
grafico.set_categories(referencia_datas)

# Editando as linhas dos gráficos, os estilos e cores
linha_cotacao = grafico.series[0]
linha_bb_inferior = grafico.series[1]
linha_bb_superior = grafico.series[2]

linha_cotacao.graphicalProperties.line.width = 0
linha_cotacao.graphicalProperties.line.solidFill = "000000"

linha_bb_inferior.graphicalProperties.line.width = 0
linha_bb_inferior.graphicalProperties.line.solidFill = "09ab11"

linha_bb_superior.graphicalProperties.line.width = 0
linha_bb_superior.graphicalProperties.line.solidFill = "ff0000"

# Adicionando de fato o gráfico à planilha
planilha_grafico.add_chart(grafico, "A3")

# Gerando o arquivo XLSX final
workbook.save(f'./saida/{acao}.xlsx')
