import openpyxl
from openpyxl.chart import BarChart, Reference

# 1. Cria o workbook e a aba
computadores = openpyxl.Workbook()
computadores_page = computadores.active
computadores_page.title = "Computadores"

# 2. Adiciona os dados
computadores_page.append(["Eletrônica", "Memória RAM", "Preço"])
computadores_page.append(["Computador1", "8GB RAM", 3000])
computadores_page.append(["Computador2", "16GB RAM", 5500])
computadores_page.append(["Computador3", "32GB RAM", 11500])

# 3. Cria o gráfico de barras
grafico = BarChart()
grafico.type = "col"           # "col" = colunas verticais | "bar" = barras horizontais
grafico.style = 10             # estilo visual (1 a 48)
grafico.title = "Preço dos Computadores"
grafico.y_axis.title = "Preço (R$)"
grafico.x_axis.title = "Computador"

# 4. Define os dados do gráfico
#    min_col=3 → coluna C (Preço)
#    min_row=1 → inclui o cabeçalho ("Preço") para virar o nome da série
#    max_row=4 → até a última linha de dados
dados = Reference(
    computadores_page,
    min_col=3, min_row=1,
    max_row=4
)

# 5. Define as categorias (rótulos do eixo X)
#    min_col=1 → coluna A (nomes dos computadores)
#    min_row=2 → pula o cabeçalho
categorias = Reference(
    computadores_page,
    min_col=1, min_row=2,
    max_row=4
)

# 6. Adiciona os dados e categorias ao gráfico
grafico.add_data(dados, titles_from_data=True)
grafico.set_categories(categorias)

# 7. Define o tamanho do gráfico (largura x altura em cm)
grafico.width = 15
grafico.height = 10

# 8. Adiciona o gráfico na planilha, na célula E2
computadores_page.add_chart(grafico, "E2")

# 9. Salva o arquivo
computadores.save("compras_grafico.xlsx")

print("Planilha com gráfico criada com sucesso!")