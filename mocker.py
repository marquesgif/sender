import pandas as pd

# Configurações
num_linhas = 100
emails_mock = ["adrianapiedade07@gmail.com", "joselubendo26@gmail.com", "jose.sad159@gmail.com"]
arquivos_mock = ["cv.pdf", "Passah.pdf"]

# Gerar dados
data = []
for i in range(1, num_linhas + 1):
    nome = f"Teste {i}"
    email = emails_mock[i % 3]         # alterna entre os dois emails
    arquivo = arquivos_mock[i % 2]     # alterna entre os dois PDFs
    data.append([nome, email, arquivo])

# Criar DataFrame
df_mock = pd.DataFrame(data, columns=["nome", "e-mail", "arquivo"])

# Salvar Excel
df_mock.to_excel("mock_envio.xlsx", index=False)
print("Mock Excel criado com sucesso!")
