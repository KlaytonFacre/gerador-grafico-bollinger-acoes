from datetime import date
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import Reference
from classes.models import LeitorDeAcoes, GerenciadorPlanilha, PropriedadeSerieGrafico

try:
    acao = input('Qual o código da ação a ser processada? ').upper()

    # Busca o arquivo do Código da Ação informado pelo usuário
    leitor_acoes = LeitorDeAcoes(caminho_arquivo='./assets/dados/')
    leitor_acoes.processa_arquivo(acao)

    # Cria um Objeto gerenciador de planilhas
    gerenciador = GerenciadorPlanilha()
    planilha_dados = gerenciador.adicionar_planilha("Dados")

    # PREPARANDO A PLANILHA DE 'DADOS'
    gerenciador.adiciona_linha(["DATA", "COTAÇÃO", "BANDA INFERIOR", "BANDA SUPERIOR"])

    indice = 2
    for linha in leitor_acoes.dados:
        try:
            # TRATANDO E PREPARANDO A DATA, CRIANDO UM OBJETO DATE DO PYTHON
            ano_mes_dia = linha[0].split(" ")[0]
            data = date(
                year=int(ano_mes_dia.split("-")[0]),
                month=int(ano_mes_dia.split("-")[1]),
                day=int(ano_mes_dia.split("-")[2])
            )

            # TRATANDO E PREPARANDO A COTAÇÃO E AS FÓRMULAS DA BANDA DE BOLLINGER
            cotacao = float(linha[1])
            formula_bb_inferior = f'=AVERAGE(B{indice}:B{indice + 19}) - 2*STDEV(B{indice}:B{indice + 19})'
            formula_bb_superior = f'=AVERAGE(B{indice}:B{indice + 19}) + 2*STDEV(B{indice}:B{indice + 19})'

            # ESCREVENDO DADOS NA PLANILHA ATIVA
            # Data
            gerenciador.atualiza_celula(celula=f'A{indice}', dado=data)
            # Cotação
            gerenciador.atualiza_celula(celula=f'B{indice}', dado=cotacao)
            # Banda inferior de Bollinger
            gerenciador.atualiza_celula(celula=f'C{indice}', dado=formula_bb_inferior)
            # Banda superior de Bollinger
            gerenciador.atualiza_celula(celula=f'D{indice}', dado=formula_bb_superior)

            # PREPARA PARA A ESCRITA NA PRÓXIMA LINHA
            indice += 1
        except ValueError:
            pass

    planilha_grafico = gerenciador.adicionar_planilha(titulo_planilha="Gráfico")

    # Mesclagem de células e criação de estilos, preenchimento e definição de alinhamento
    gerenciador.mescla_celulas(celula_inicio="A1", celula_fim="T2")
    gerenciador.aplica_estilos(
        celula="A1",
        estilos=[
            ("font", Font(b=True, sz=18, color="FFFFFF")),
            ("fill", PatternFill("solid", fgColor="07838f")),
            ("alignment", Alignment(vertical="center", horizontal="center"))
        ]
    )
    gerenciador.atualiza_celula(celula="A1", dado="Histórico de Cotações")

    # Criação do Gráfico
    referencia_cotacoes = Reference(planilha_dados, min_col=2, min_row=2, max_col=4, max_row=indice)
    referencia_datas = Reference(planilha_dados, min_col=1, min_row=2, max_col=1, max_row=indice)

    # Adicionando um gráfico de linhas VERSÃO 1
    '''
    gerenciador.adiciona_grafico_linha(
        celula="A3",
        comprimento=33.87,
        altura=15.5,
        titulo=f"Cotações - {acao}",
        titulo_eixo_x="Data da Cotação",
        titulo_eixo_y="Valor da Cotação",
        referencia_eixo_x=referencia_cotacoes,
        referencia_eixo_y=referencia_datas,
        propriedades_grafico=[
            PropriedadeSerieGrafico(grossura=0, cor_preenchimento="000000"),
            PropriedadeSerieGrafico(grossura=0, cor_preenchimento="09ab11"),
            PropriedadeSerieGrafico(grossura=0, cor_preenchimento="ff0000"),
        ]
    )'''

    # Adicionando um gráfico de linhas VERSÃO 2
    gerenciador.adiciona_grafico_linha_2(
        celula="A3",
        referencia_eixo_x=referencia_cotacoes,
        referencia_eixo_y=referencia_datas,
        propriedades_grafico=[
            ("width", 33.87),
            ("height", 15.5),
            ("title", f'Cotações - {acao}'),
            ("x_axis.title", "Data da Cotação"),
            ("y_axis.title", "Valor da Cotação")
        ],
        propriedades_series=[
            (0, "width", 0),
            (0, "solidFill", "000000"),
            (1, "width", 0),
            (1, "solidFill", "09ab11"),
            (2, "width", 0),
            (2, "solidFill", "ff0000"),
        ]
    )

    # Adicionando uma imagem
    gerenciador.mescla_celulas(celula_inicio="I33", celula_fim="L36")
    gerenciador.adiciona_imagem("I33", "./assets/recursos/logo.png")

    # Gerando o arquivo XLSX final
    gerenciador.salva_arquivo(f'./saida/{acao}-REFATORADA.xlsx')
except FileNotFoundError:
    print("ERRO! Arquivo não encontrado!")
except AttributeError as e:
    print(f'Atributo inexistente! Conserte o código!')
except Exception as e:
    print(f"Aconteceu um erro inexperado no programa. Erro: {str(e)}")
