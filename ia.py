# from ollama import chat
# import pandas as pd
# import os
# from dotenv import load_dotenv

# load_dotenv

# IP_DESKTOP = os.getenv("")
# OLLAMA_CLIENT = Client(host=f"http://{IP_DESKTOP}:11434")

# def Gemma4(prompt: str, num_predict: int = 300, stream: bool = False):
#     output = chat(
#         model="gemma4",
#         messages=[
#             {
#                 'role': 'system',
#                 'content': f"""
#                 Responda sem usar formatação Markdown ou qualquer outro tipo de formatação.

#                 Você é um assistente financeiro inteligente e analítico.
#                 Você tem acesso a um dataframe financeiro que contém informações sobre receitas, despesas, investimentos e outros dados financeiros relevantes.
#                 Sua tarefa é analisar esse dataframe e fornecer insights, responder perguntas e ajudar a tomar decisões financeiras informadas com base nos dados disponíveis.
#                 Seja claro, conciso e forneça respostas detalhadas quando necessário, quando não necessário, responda de forma breve e direta.

#                 Sempre que for se referir ao dataframe, o refira como planilha.
#             """
#         },
#         {
#             'role': 'user',
#             'content': prompt
#         }],

#         options={
#             'num_predict': num_predict
#         },

#         think=False,
#         stream=stream
#     )

#     if stream:
#         def gerador():
#             for token in output:
#                yield token['message']['content']   
#         return gerador()
#     else:
#         return output['message']['content'] 

import requests
import json
import pandas as pd
import os
from dotenv import load_dotenv

# Correção: adicionado os parênteses para executar a função
load_dotenv()

# Correção: adicionei o nome da variável (ex: "IP_DESKTOP") e um valor padrão de fallback
IP_DESKTOP = os.getenv("IP_DESKTOP", "127.0.0.1") 
OLLAMA_URL = f"http://{IP_DESKTOP}:11434/api/chat"

def Gemma4(prompt: str, num_predict: int = 300, stream: bool = False):
    
    # Montamos o corpo da requisição idêntico à documentação da API REST do Ollama
    payload = {
        "model": "gemma4",
        "messages": [
            {
                "role": "system",
                "content": """
                Responda sem usar formatação Markdown ou qualquer outro tipo de formatação.

                Você é um assistente financeiro inteligente e analítico.
                Você tem acesso a um dataframe financeiro que contém informações sobre receitas, despesas, investimentos e outros dados financeiros relevantes.
                Sua tarefa é analisar esse dataframe e fornecer insights, responder perguntas e ajudar a tomar decisões financeiras informadas com base nos dados disponíveis.
                Seja claro, conciso e forneça respostas detalhadas quando necessário, quando não necessário, responda de forma breve e direta.

                Sempre que for se referir ao dataframe, o refira como planilha.
                """
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "options": {
            "num_predict": num_predict
        },
        "stream": stream
    }

    if stream:
        # Lógica para o modo Stream usando o iter_lines() do requests
        def gerador():
            try:
                with requests.post(OLLAMA_URL, json=payload, stream=True, timeout=60) as response:
                    response.raise_for_status() # Verifica se deu erro HTTP 
                    
                    for linha in response.iter_lines():
                        if linha:
                            # O Ollama retorna um JSON por linha durante o stream
                            chunk = json.loads(linha.decode('utf-8'))
                            if "message" in chunk and "content" in chunk["message"]:
                                yield chunk["message"]["content"]
            except Exception as e:
                yield f"[Erro na conexão com Ollama: {e}]"
                
        return gerador()
        
    else:
        # Lógica para resposta direta (sem stream)
        try:
            response = requests.post(OLLAMA_URL, json=payload, timeout=60)
            response.raise_for_status()
            
            dados = response.json()
            return dados['message']['content']
        except Exception as e:
            return f"Erro de conexão com a IA: {e}"