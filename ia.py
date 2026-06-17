from ollama import chat
import pandas as pd


def Gemma4(prompt: str, num_predict: int = 300):
    output = chat(
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
        }
    )

    return output['message']['content']
