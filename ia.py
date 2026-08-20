from ollama import Client
import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv

IP_DESKTOP = os.getenv("")

def Gemma4(prompt: str, num_predict: int = 300, stream: bool = False):
    output = Client.chat(
        model="gemma4",
        messages=[
            {
                'role': 'system',
                'content': f"""
                Responda sem usar formatação Markdown ou qualquer outro tipo de formatação.

                Você é um assistente financeiro inteligente e analítico.
                Você tem acesso a um dataframe financeiro que contém informações sobre receitas, despesas, investimentos e outros dados financeiros relevantes.
                Sua tarefa é analisar esse dataframe e fornecer insights, responder perguntas e ajudar a tomar decisões financeiras informadas com base nos dados disponíveis.
                Seja claro, conciso e forneça respostas detalhadas quando necessário, quando não necessário, responda de forma breve e direta.

                Sempre que for se referir ao dataframe, o refira como planilha.
            """
        },
        {
            'role': 'user',
            'content': prompt
        }],

        options={
            'num_predict': num_predict
        },

        think=False,
        stream=stream
        host=IP_DESKTOP
    )

    if stream:
        def gerador():
            for token in output:
               yield token['message']['content']   
        return gerador()
    else:
        return output['message']['content'] 
