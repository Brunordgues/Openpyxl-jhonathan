import openpyxl

contato = openpyxl.Workbook()
contato_page = contato.create_sheet("Contatos")

contato_page.append(["Nome", "Telefone", "Email"]) #cabeçalho
contato_page.append(["Bruno", "123456789", "bruno@email.com"])
contato_page.append(["Maria", "987654321", "maria@email.com"])
contato_page.append(["Sabrina", "5345345", "sabrina@email.com"])
contato_page.append(["João", "53454353", "joao@email.com"])
contato_page.append(["Ana", "98453344321", "ana@email.com"])

contato.save("Planilha de Contatos.xlsx")
print("Planilha de contatos criada com sucesso!")