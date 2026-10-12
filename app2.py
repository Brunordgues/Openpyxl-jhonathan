import openpyxl

#carregando arquivo
book = openpyxl.load_workbook("Planilha de Compras.xlsx")

#selecionando uma página
frutas_page = book["Frutas"]

#imprimindo os dados de cada linha
for rows in frutas_page.iter_rows(min_row=2,max_row=5):
    for cell in rows:
        print(cell.value)

\**