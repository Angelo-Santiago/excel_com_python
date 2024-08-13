import pandas as pd

# Carregar a planilha existente
df = pd.read_excel('exemplo_planilha.xlsx')

# Adicionar uma nova coluna
df['Email'] = ['alice@example.com', 'bob@example.com', 'charlie@example.com']

# Adicionar uma nova linha
nova_linha = pd.DataFrame({'Nome': ['Daniel'], 'Idade': [40], 'Cidade': ['Curitiba'], 'Email': ['daniel@example.com']})
df = pd.concat([df, nova_linha], ignore_index=True)

# Salvar as alterações na planilha
df.to_excel('exemplo_planilha_atualizada.xlsx', index=False)

print("Planilha atualizada com sucesso!")
