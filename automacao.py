import pandas as pd

# Criar um DataFrame com alguns dados
dados = {
    'Nome': ['Alice', 'Bob', 'Charlie'],
    'Idade': [25, 30, 35],
    'Cidade': ['São Paulo', 'Rio de Janeiro', 'Belo Horizonte']
}

df = pd.DataFrame(dados)

# Salvar o DataFrame em um arquivo Excel
df.to_excel('exemplo_planilha.xlsx', index=False)

print("Planilha criada com sucesso!")
