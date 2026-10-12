import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

# 1. Cria o workbook
gastos = openpyxl.Workbook()

# 2. Renomeia a aba padrão em vez de criar uma nova
gastos_page = gastos.active
gastos_page.title = "Gastos"

# 3. Cabeçalho
gastos_page.append(["Descrição", "Valor", "Data"])

# 4. Dados — VALORES COMO NÚMERO (sem aspas!)
gastos_page.append(["Aluguel", 1200.00, "01/01/2023"])
gastos_page.append(["Compras", 300.00, "05/01/2023"])
gastos_page.append(["Combustível", 200.00, "10/01/2023"])

# 5. Cria o gráfico
grafico = BarChart()
grafico.type = "col"           # colunas verticais
grafico.style = 10
grafico.title = "Gastos do Mês"
grafico.y_axis.title = "Valor (R$)"
grafico.x_axis.title = "Descrição"

# 6. Dados = coluna B (Valor)
dados = Reference(
    gastos_page,
    min_col=2, min_row=1,
    max_row=4
)

# 7. Categorias = coluna A (Descrição)
categorias = Reference(
    gastos_page,
    min_col=1, min_row=2,
    max_row=4
)

# 8. Monta o gráfico
grafico.add_data(dados, titles_from_data=True)
grafico.set_categories(categorias)

# 9. Mostra os valores em cima das barras
grafico.dataLabels = DataLabelList()
grafico.dataLabels.showVal = True

# 10. Tamanho do gráfico
grafico.width = 15
grafico.height = 10

# 11. Cola o gráfico na célula E2
gastos_page.add_chart(grafico, "E2")

# 12. Salva
gastos.save("Planilha de Gastos.xlsx")
print("Planilha de gastos com gráfico criada com sucesso!")