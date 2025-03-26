import openai

# Defina sua chave de API
openai.api_key = "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"

def ask_copilot(prompt):
    try:
        # Nova interface para chamadas de completions
        response = openai.completions.create(
            model="gpt-3.5-turbo",  # Ou outro modelo, como gpt-4
            prompt=prompt,  # Mensagem que o usuário envia
            max_tokens=100,  # Número máximo de tokens na resposta
            temperature=0.7  # Aleatoriedade da resposta
        )

        # Obtenha a resposta gerada
        message = response['choices'][0]['text'].strip()
        print(f"Resposta: {message}")
        return message
    except Exception as e:
        print(f"Erro: {e}")


# Exemplo de interação
prompt = "Como posso aprender Python de forma eficaz?"
resposta = ask_copilot(prompt)
