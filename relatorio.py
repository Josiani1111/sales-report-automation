
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import (
    Table,
    TableStyleInfo,
    TableColumn,
    TableFormula
)
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import AutoFilter
from openpyxl import load_workbook

# Criar a planilha
wb = Workbook()
ws = wb.active
ws.title = "Relatório de Vendas"

# Criar os cabeçalhos
ws.append(["Produto", "Quantidade", "Preço Unitário", "Total"])

# Formatar o cabeçalho
azul = PatternFill(
    start_color="17365D",
    end_color="17365D",
    fill_type="solid"
)

for celula in ws[1]:
    celula.font = Font(color="FFFFFF", bold=True)
    celula.fill = azul

# Cadastrar os produtos
# Lista de produtos

produtos = [
    ["Bolsa", 10, 50],
    ["Calçados", 5, 120],
    ["Mochila", 8, 150],
    ["Carteira", 4, 50]
]

# Cadastrar produtos e calcular totais
for produto in produtos:
    ws.append(produto)

    linha = ws.max_row

    ws[f"D{linha}"] = f"=B{linha}*C{linha}"

    print(f"Produto: {produto[0]} | Linha: {linha}")
ultima_linha_produtos = ws.max_row

print(f"Última linha dos produtos: {ultima_linha_produtos}")
# Criar linha de faturamento
ws.append([
    "Faturamento Geral",
    "",
    "",
    f"=SUBTOTAL(109,D2:D{ultima_linha_produtos})"
])
linha_faturamento = ws.max_row

# Formatar os valores monetáriosA
formato_moeda = '"R$" #,##0.00'

for linha in range(2, linha_faturamento + 1):
    ws[f"C{linha}"].number_format = formato_moeda
    ws[f"D{linha}"].number_format = formato_moeda

# Destacar o faturamento

ws[f"A{linha_faturamento}"].font = Font(bold=True)
ws[f"D{linha_faturamento}"].font = Font(bold=True)

# Criar a Tabela do Excel
tabela = Table(
    displayName="TabelaVendas",
    ref=f"A1:D{linha_faturamento}",
    totalsRowCount=1,
    totalsRowShown=True
)

# Configurar as colunas da tabela
tabela.tableColumns = [
    TableColumn(
        id=1,
        name="Produto",
        totalsRowLabel="Faturamento Geral"
    ),
    TableColumn(id=2, name="Quantidade"),
    TableColumn(id=3, name="Preço Unitário"),
    TableColumn(
        id=4,
        name="Total",
        calculatedColumnFormula=TableFormula(
            attr_text="[@Quantidade]*[@[Preço Unitário]]"
        ),
        totalsRowFunction="sum"
    )
]

# Configurar os filtros
tabela.autoFilter = AutoFilter(
    ref=f"A1:D{ultima_linha_produtos}"
)

# Aplicar estilo à tabela
tabela.tableStyleInfo = TableStyleInfo(
    name="TableStyleMedium2",
    showRowStripes=True
)

ws.add_table(tabela)

# Ajustar largura das colunas
ws.column_dimensions["A"].width = 25
ws.column_dimensions["B"].width = 15
ws.column_dimensions["C"].width = 20
ws.column_dimensions["D"].width = 20
# Validar a quantidade dos produtos

validacao = DataValidation(
    type="whole",
    operator="greaterThanOrEqual",
    formula1=0,
    allow_blank=False
)

ws.add_data_validation(validacao)

validacao.add(f"B2:B{ultima_linha_produtos}")

# Mensagem de erro para quantidade inválida
validacao.error = "Digite uma quantidade igual ou maior que zero."
validacao.errorTitle = "Quantidade inválida"
validacao.showErrorMessage = True
validacao.errorStyle = "stop"

# Validar o preço unitário dos produtos
validacao_preco = DataValidation(
    type="decimal",
    operator="greaterThanOrEqual",
    formula1=0,
    allow_blank=False
)

ws.add_data_validation(validacao_preco)
validacao_preco.add(f"C2:C{ultima_linha_produtos}")

# Mensagem de erro para preço inválido
validacao_preco.error = "O preço não pode ser negativo."
validacao_preco.errorTitle = "Preço inválido"
validacao_preco.showErrorMessage = True
validacao_preco.errorStyle = "stop"
# Salvar a planilha
wb.save("relatorio_teste.xlsx")

print("Planilha criada com sucesso!")



# Abrir novamente o arquivo salvo
arquivo = load_workbook("relatorio_teste.xlsx")

# Selecionar a planilha
planilha = arquivo["Relatório de Vendas"]

# Verificar os dados gravados
print("Dados salvos no Excel:")

for linha in planilha.iter_rows(values_only=True):
    print(linha)

arquivo.close()

    