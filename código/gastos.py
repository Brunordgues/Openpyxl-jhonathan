import openpyxl

gastos = openpyxl.Workbook()
gastos_page = gastos.create_sheet("Gastos")

gastos_page.append(["Descrição", "Valor", "Data"]) #cabeçalho
gastos_page.append(["Aluguel", "1200.00", "01/01/2023"])
gastos_page.append(["Compras", "300.00", "05/01/2023"])
gastos_page.append(["Combustível", "200.00", "10/01/2023"])

gastos.save("Planilha de Gastos.xlsx")
print("Planilha de gastos criada com sucesso!")
