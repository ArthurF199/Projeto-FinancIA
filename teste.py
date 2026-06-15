from ollama import chat
from pandas import read_excel

df = read_excel('data.xlsx')

resposta = (chat(

    model ="gemma4",
    messages = [
        {
            'role': 'system',
            'content': f"Você vai receber o seguinte dataframe financeiro: {df.to_string(index=False)}, você vai analisar ele e responder o que é pedido."


        },
        {
            'role': 'user',
            'content': input()
        }
    ]
))

print(resposta['message']['content'])
