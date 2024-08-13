import pandas as pd
import os

# Criar o plano de contas
plano_de_contas = {
    "Código": [
        "1", "1.1", "1.1.1", "1.1.2", "1.1.3", "1.1.4", "1.1.5",
        "1.2", "1.2.1", "1.2.1.1", "1.2.1.2", "1.2.1.3", "1.2.2", 
        "1.2.2.1", "1.2.3", "1.2.3.1", "1.2.3.2",
        "2", "2.1", "2.1.1", "2.1.2", "2.1.3", "2.1.4",
        "2.2", "2.2.1", "2.2.2", "2.2.3",
        "3", "3.1", "3.2", "3.3",
        "4", "4.1", "4.1.1", "4.1.2", "4.2", "4.2.1", "4.2.2", "4.3", "4.3.1", "4.3.2",
        "5", "5.1", "5.1.1", "5.1.2", "5.1.3", "5.1.4", "5.2", "5.2.1", "5.2.2", "5.3", "5.3.1", "5.3.2"
    ],
    "Descrição": [
        "Ativo", "Ativo Circulante", "Caixa", "Bancos Conta Movimento", "Clientes", "Estoques", "Adiantamentos a Fornecedores",
        "Ativo Não Circulante", "Imobilizado", "Terrenos", "Edificações", "Máquinas e Equipamentos", "Investimentos",
        "Participações em Outras Empresas", "Intangível", "Marcas e Patentes", "Software",
        "Passivo", "Passivo Circulante", "Fornecedores", "Empréstimos e Financiamentos", "Salários e Encargos", "Impostos a Recolher",
        "Passivo Não Circulante", "Empréstimos e Financiamentos de Longo Prazo", "Provisões", "Obrigações com Partes Relacionadas",
        "Patrimônio Líquido", "Capital Social", "Reservas de Capital", "Lucros ou Prejuízos Acumulados",
        "Receitas", "Receita Bruta de Vendas", "Receita de Vendas de Produtos", "Receita de Vendas de Serviços", 
        "Deduções de Receitas", "Devoluções de Vendas", "Descontos Concedidos", "Outras Receitas", "Receita Financeira", "Receita de Aluguéis",
        "Despesas", "Despesas Operacionais", "Despesas com Pessoal", "Despesas Administrativas", "Despesas Comerciais", "Despesas de Marketing", 
        "Despesas Financeiras", "Juros Passivos", "Despesas com Empréstimos", "Outras Despesas", "Perdas em Processos Judiciais", "Despesas com Impostos"
    ]
}

# Converter em DataFrame
df_plano_de_contas = pd.DataFrame(plano_de_contas)

# Obter o caminho para a área de trabalho do usuário
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

# Definir o caminho completo para o arquivo Excel
file_path = os.path.join(desktop_path, "Plano_de_Contas_Exemplo.xlsx")

# Salvar como um arquivo Excel
df_plano_de_contas.to_excel(file_path, index=False)

file_path
