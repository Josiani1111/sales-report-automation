from pathlib import Path
from openpyxl import load_workbook
def test_nome_planilha():
    arquivo = load_workbook("relatorio_teste.xlsx")

    assert "Relatório de Vendas" in arquivo.sheetnames

    arquivo.close()

def test_arquivo_existe():
    arquivo = Path("relatorio_teste.xlsx")

    assert arquivo.exists()
def test_cabecalhos():
    arquivo = load_workbook("relatorio_teste.xlsx")

    planilha = arquivo["Relatório de Vendas"]

    cabecalhos = [celula.value for celula in planilha[1]]

    assert cabecalhos == [
        "Produto",
        "Quantidade",
        "Preço Unitário",
        "Total"
    ]

    arquivo.close()
def test_produtos_cadastrados():
    arquivo = load_workbook("relatorio_teste.xlsx")

    planilha = arquivo["Relatório de Vendas"]

    produtos = [
        linha[0]
        for linha in planilha.iter_rows(
            min_row=2,
            max_row=planilha.max_row - 1,
            values_only=True
        )
    ]

    assert produtos == [
        "Bolsa",
        "Calçados",
        "Mochila",
        "Carteira"
    ]

    arquivo.close() 

def test_formulas():
    arquivo = load_workbook("relatorio_teste.xlsx")

    planilha = arquivo["Relatório de Vendas"]

    ultima_linha_produtos = planilha.max_row - 1

    for linha in range(2, ultima_linha_produtos + 1):
        assert planilha[f"D{linha}"].value == f"=B{linha}*C{linha}"

    assert planilha[f"D{planilha.max_row}"].value == (
        f"=SUBTOTAL(109,D2:D{ultima_linha_produtos})"
    )

    arquivo.close()
def test_tabela_filtros():
    arquivo = load_workbook("relatorio_teste.xlsx")

    planilha = arquivo["Relatório de Vendas"]

    assert "TabelaVendas" in planilha.tables

    tabela = planilha.tables["TabelaVendas"]

    assert tabela.ref == f"A1:D{planilha.max_row}"

    assert tabela.autoFilter is not None

    assert tabela.autoFilter.ref == (
        f"A1:D{planilha.max_row - 1}"
    )

    arquivo.close()    

def test_validacoes():
    arquivo = load_workbook("relatorio_teste.xlsx")

    planilha = arquivo["Relatório de Vendas"]

    validacoes = planilha.data_validations.dataValidation

    assert len(validacoes) == 2

    quantidade = validacoes[0]
    preco = validacoes[1]

    assert quantidade.type == "whole"
    assert quantidade.operator == "greaterThanOrEqual"
    assert quantidade.formula1 == "0"

    assert preco.type == "decimal"
    assert preco.operator == "greaterThanOrEqual"
    assert preco.formula1 == "0"

    ultima_linha = planilha.max_row - 1

    assert str(quantidade.sqref) == f"B2:B{ultima_linha}"
    assert str(preco.sqref) == f"C2:C{ultima_linha}"

    arquivo.close()