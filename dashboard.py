import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from io import BytesIO

# Dados para o DataFrame
dados = {
    'Mês': ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
    'Entrada': [10000, 12000, 15000, 13000, 11000, 14000, 16000, 15500, 14500, 12500, 17000, 18000],
    'Saída': [8000, 7000, 9000, 8500, 7500, 9500, 10000, 9800, 9200, 8300, 10500, 11000],
    'Entrada Estimada': [9500, 11500, 14000, 12500, 10500, 13500, 15500, 15000, 14000, 12000, 16500, 17500],
    'Saída Estimada': [7800, 7200, 8800, 8600, 7400, 9400, 9800, 9700, 9100, 8200, 10200, 10800],
    'Tipo de Saída': ['Dinheiro', 'Pix', 'Cartão de Crédito', 'Cartão de Débito', 'Dinheiro', 'Pix', 'Cartão de Crédito', 'Cartão de Débito', 'Dinheiro', 'Pix', 'Cartão de Crédito', 'Cartão de Débito']
}

df = pd.DataFrame(dados)
df['Lucro'] = df['Entrada'] - df['Saída']

# Salvar o DataFrame em um arquivo Excel
excel_path = 'dashboard_financeira.xlsx'
df.to_excel(excel_path, index=False)

# Criar gráficos
plt.figure(figsize=(14, 8))

# Gráfico de barras: Entrada e Saída
plt.subplot(2, 2, 1)
plt.bar(df['Mês'], df['Entrada'], color='green', label='Entrada')
plt.bar(df['Mês'], df['Saída'], color='red', label='Saída', alpha=0.7)
plt.xlabel('Mês')
plt.ylabel('Valor')
plt.title('Entrada e Saída Mensal')
plt.xticks(rotation=45)
plt.legend()

# Gráfico de linhas: Entrada e Saída Estimada vs Real
plt.subplot(2, 2, 2)
plt.plot(df['Mês'], df['Entrada'], marker='o', label='Entrada Real')
plt.plot(df['Mês'], df['Entrada Estimada'], marker='o', linestyle='--', label='Entrada Estimada')
plt.plot(df['Mês'], df['Saída'], marker='x', label='Saída Real')
plt.plot(df['Mês'], df['Saída Estimada'], marker='x', linestyle='--', label='Saída Estimada')
plt.xlabel('Mês')
plt.ylabel('Valor')
plt.title('Entrada e Saída: Real vs Estimado')
plt.xticks(rotation=45)
plt.legend()

# Gráfico de barras: Lucro Mensal
plt.subplot(2, 2, 3)
plt.bar(df['Mês'], df['Lucro'], color='blue')
plt.xlabel('Mês')
plt.ylabel('Lucro')
plt.title('Lucro Mensal')
plt.xticks(rotation=45)

# Gráfico de setores: Tipo de Saída
tipo_saida_counts = df['Tipo de Saída'].value_counts()
plt.subplot(2, 2, 4)
plt.pie(tipo_saida_counts, labels=tipo_saida_counts.index, autopct='%1.1f%%', startangle=140)
plt.title('Distribuição dos Tipos de Saída')

# Salvar os gráficos em um buffer de memória
buf = BytesIO()
plt.tight_layout()
plt.savefig(buf, format='png')
buf.seek(0)
plt.close()

# Carregar a planilha existente
workbook = load_workbook(excel_path)
sheet = workbook.active

# Adicionar o gráfico à planilha
image = Image(buf)
sheet.add_image(image, 'E5')

# Salvar a planilha atualizada
workbook.save(excel_path)

print("Dashboard criada com sucesso!")
