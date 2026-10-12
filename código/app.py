import openpyxl

# ============================================
# CRIAR uma planilha do Excel (book)
# ============================================
book = openpyxl.Workbook()

# Visualizar as planilhas (abas) existentes
print(book.sheetnames)

# ============================================
# CRIAR uma nova página (aba) chamada "Frutas"
# ============================================
book.create_sheet("Frutas")

# ============================================
# SELECIONAR a página "Frutas"
# ============================================
frutas_page = book["Frutas"]

# Adicionar o cabeçalho
frutas_page.append(["Fruta", "Quantidade", "Preço"])

# Adicionar os dados
frutas_page.append(["Banana", "5", "R$ 3,90"])
frutas_page.append(["Mamão", "2", "R$ 5,90"])
frutas_page.append(["Kiwi", "5", "R$ 20,90"])
frutas_page.append(["Melancia", "1", "R$ 11,90"])

# ============================================
# SALVAR a planilha
# ============================================
book.save("Planilha de Compras.xlsx")

print("Planilha criada com sucesso!")