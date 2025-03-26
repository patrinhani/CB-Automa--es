import openai

client = openai.OpenAI(api_key="SUA_API_KEY")

try:
    resposta = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": "Diga olá"}]
    )
    print(resposta.choices[0].message.content)
except Exception as e:
    print(f"Erro: {e}")
