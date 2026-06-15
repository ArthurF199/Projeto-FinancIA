import pandas as pd
import ia
import json


def title(title: str):
    print('-'*10, title, '-'*10)


def options(options: list):
    for i, option in enumerate(options):
        print(f'{i+1} - {option}')
    return int(input('Digite o número da opção desejada: '))


def add_revenue():
    title('Adicionar receita')
    response = ia.Gemma4(f"""
              Vou te enviar uma descrição de uma receita, quero que você me retorne um JSON com os seguintes campos:
              - Descrição: uma breve descrição da receita(caso não seja informada, deixe em branco(null))
              - Valor: o valor da receita, seguindo o formato R$ 000.00, (caso não seja informada, deixe em branco(null))
              - Data: a data da receita, seguindo o formato DD-MM-AAAA (caso a data não for informada, deixe em branco(null))
              Siga o seguinte exemplo: {{ "Descrição": "Descrição da receita", "Valor": "R$ 000.00", "Data": "DD-MM-AAAA" }}
              Quero que me mande somente a receita, sem nenhuma explicação ou formatação adicional.
              A descrição da receita é: {input('Usuário: ')}
    """)
    print(response)
    response = json.loads(response)
    print(response)
    n = len(df)
    df.loc[n, 'Descrição'] = [response['Descrição']]
    df.loc[n, 'Valor'] = [response['Valor']]
    df.loc[n, 'Data'] = [response['Data']]
    df.to_excel('data.xlsx', index=False)
    print('Receita adicionada com sucesso!')


df = pd.read_excel('data.xlsx')

while True:
    title('FinancIA: Gestor Financeiro')
    match options(['Analisar a planilha', 'Adicionar ou remover receitas', 'Adicionar ou remover despesas', 'Viver de Renda', 'Sair']):
        case 1:
            print(ia.Gemma4('Analise o dataframe e me diga o que você acha')['message']['content'])
            break
        case 2:
            add_revenue()
        case 3:
            print(df)
        case 4:
            break