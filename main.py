import pandas as pd
import ia
import json


def title(title: str):
    print('-'*10, title, '-'*10)


def options(options: list):
    for i, option in enumerate(options):
        print(f'{i+1} - {option}')
    return int(input('Digite o número da opção desejada: '))


def add_revenue(df):
    title('Adicionar receita')
    response = ia.Gemma4(f"""
              Vou te enviar uma descrição de uma receita, quero que você me retorne um JSON com os seguintes campos:
              - Descrição: uma breve descrição da receita(caso não seja informada, deixe em branco(null))
              - Valor: o valor da receita, seguindo o formato R$ 000.00, (caso não seja informada, deixe em branco(null))
              - Data: a data da receita, seguindo o formato DD-MM-AAAA (caso a data não for informada, deixe em branco(null))
              Siga o seguinte exemplo: {{ "Descrição": "Descrição da receita", "Valor": "R$ 0000.00", "Data": "DD-MM-AAAA" }}
              Quero que me mande somente a receita, sem nenhuma explicação ou formatação adicional.
                         
              Considere como não informada apenas quando realmente o usuário não descreveu a informação, tente ao máximo extrair as informações que o usuário descrever.
              A descrição da receita é: {input('Usuário: ')}
    """, 500)

    print(repr(response))
    response = json.loads(response)
    new_df = df.copy()
    
    n = len(new_df)
    
    new_df.loc[n, 'Descrição'] = response['Descrição']
    new_df.loc[n, 'Valor'] = float(response['Valor'].split('R$ ')[1].replace(',', '.'))
    new_df.loc[n, 'Data'] = response['Data']

    return new_df


df = pd.read_excel('data.xlsx')

while True:
    title('FinancIA: Gestor Financeiro')
    match options(['Analisar a planilha', 'Adicionar ou remover receitas', 'Visualisar Planilha', 'Viver de Renda', 'Sair']):
        case 1:
            print(ia.Gemma4('Analise o dataframe e me diga o que você acha, responda de forma breve e direta, sem formatação ou explicação adicional.', 300))
            break
        case 2:
            try:
                unsaved_df = add_revenue(df)

                unsaved_df.to_excel('data.xlsx', index=False) # Unsaved DataFrame
                df = unsaved_df # Saved Dataframe
                print('Receita adicionada com sucesso!')            
            except PermissionError as e:
                print(f'Erro de permissão. Sua planilha está aberta, feche-a para que possa ser modificada. [{e}]')
                print(df)
            except Exception as e:
                print(f'Ocorreu um erro ao adicionar a receita: {e}')
        case 3:
            print(df)
        case 4:
            break
