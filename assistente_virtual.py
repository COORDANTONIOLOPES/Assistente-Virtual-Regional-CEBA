import openai
import pandas as pd

# Função para carregar a planilha Excel
def carregar_planilha(caminho):
    return pd.read_excel(caminho)

# Função para buscar informações na planilha
def buscar_acao(df, termo):
    resultado = df[df.apply(lambda row: row.astype(str).str.contains(termo, case=False, na=False).any(), axis=1)]
    return resultado.to_dict(orient="records")

# Função para processar perguntas e responder com IA ou dados do Excel
def chat_assistente(pergunta, df):
    openai.api_key = "SUA_API_KEY"

    if "plano de ação" in pergunta.lower():
        termo = pergunta.split("plano de ação")[-1].strip()
        resultados = buscar_acao(df, termo)
        if resultados:
            return f"Encontrei estas informações: {resultados}"
        else:
            return "Não encontrei informações correspondentes no plano de ação."

    resposta = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "system", "content": "Você é um assistente virtual útil."},
                  {"role": "user", "content": pergunta}]
    )

    return resposta['choices'][0]['message']['content']

# Execução do assistente virtual
if __name__ == "__main__":
    df = carregar_planilha("plano_de_acao.xlsx")
    print("Assistente Virtual: Olá! Como posso ajudar você hoje?")
    while True:
        pergunta = input("Você: ")
        if pergunta.lower() in ["sair", "exit", "tchau"]:
            print("Assistente Virtual: Até logo!")
            break
        resposta = chat_assistente(pergunta, df)
        print("Assistente Virtual:", resposta)
