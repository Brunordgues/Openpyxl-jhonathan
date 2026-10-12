import openpyxl

computadores = openpyxl.Workbook()
computadores_page = computadores.create_sheet("Computadores")

computadores_page.append(["Eletrônica", "Memória RAM", "Preço"])
computadores_page.append(["Computador1", "8GB RAM", "R$ 3.000,00"])
computadores_page.append(["Computador2", "16GB RAM", "R$ 5.500,00"])
computadores_page.append(["Computador3", "32GB RAM", "R$ 11.500,00"])

computadores.save("Planilha de Compras.xlsx")
print("Planilha criada com sucesso!")