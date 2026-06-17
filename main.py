import pandas as pd
import ia
import json


def title(title: str):
    print('-'*10, title, '-'*10)


def options(options: list):
    for i, option in enumerate(options):
        print(f'{i+1} - {option}')
    return int(input('Digite o número da opção desejada: '))


def registerData(df):
    title('Adicionar receita')
    response = ia.Gemma4(f"""
              Vou te enviar uma descrição de um registro.
              Quero que você determine se o usuário quer adicionar ou remover uma receita e me retorne um JSON com os seguintes campos:
              - Ação: Adicionar/remover o registro, deixe 0 para remover e 1 para adicionar um dado à planilha
                (caso o usuário não deixe claro se ele deseja adicionar ou remover, considere deixar em 1)
              - Descrição: uma breve descrição da receita
                (caso não seja informada, escreva null no campo)
              - Valor: o valor da receita, seguindo o formato R$ 000.00.
                (caso não seja informada, escreva null no campo)
              - Data: a data do registro, seguindo o formato DD-MM-AAAA
                (caso a data não for informada, escreva null no campo)         
              - Tipo: se é entrada ou saída ou conta.
                (caso não for informada, escreva null no campo)
              - Dia de pagamento: somente se for descrito como uma conta, siga o mesmo formato do campo Data(DD-MM-AA)
                (caso não for informada, escreva null no campo)
              Siga o seguinte exemplo: {{"Ação": 0 ou 1, "Descrição": "Descrição da receita", "Valor": "R$ 0000.00", "Data": "DD-MM-AAAA", "Tipo": "Entrada ou Saída de dinheiro", "Dia de Pagamento": "DD-MM-AAAA (somente se for necessário)"}}
              Quero que me mande somente a receita, sem nenhuma explicação ou formatação adicional.
                         
              Considere como não informada apenas quando realmente o usuário não descreveu a informação, tente ao máximo extrair as informações que o usuário descrever.
              Considere como ação remover quando o usuário diz algo como: "quero remover tal conta de tal dia com tal descrição"
              
              Considere os seguintes exemplos de extrações corretas:
              
              Entrada: Recebi meu salário de R$ 3.500,00 hoje.
              Saída: "Ação": 1, "Descrição": "Salário", "Valor": "R$ 3500.00", "Data": "16-06-2026", "Tipo": "Entrada", "Dia de Pagamento": null
                         
              Entrada: Gastei R$ 120,50 no supermercado ontem.
              Saída: "Ação": 1, "Descrição": "Supermercado", "Valor": "R$ 120.50", "Data": "15-06-2026", "Tipo": "Saída", "Dia de Pagamento": null
               
              Entrada: Todo dia 10 pago R$ 89,90 da internet.
              Saída: "Ação": 1, "Descrição": "Internet", "Valor": "R$ 89.90", "Data": null, "Tipo": "Saída", "Dia de Pagamento": "10-07-2026"
              
              Entrada: Remova o gasto de restaurante de R$ 300,00.
              Saída: "Ação": 0, "Descrição": "Restaurante", "Valor": "R$ 300.00", "Data": null, "Tipo": "Saída", "Dia de Pagamento": null
              
              Entrada: Recebi um dinheiro de um freela.
              Saída: "Ação": 1, "Descrição": "Freelance", "Valor": null, "Data": null, "Tipo": "Entrada", "Dia de Pagamento": null
              {"Descrição"}
              A descrição do registro é: {input('Usuário: ')}
    """, 1000)

    print(repr(response))
    response = json.loads(response)
    new_df = df.copy()
    n = len(new_df)
    
    match response['Ação']:
        case 0:
            print(new_df['Descrição'] == response['Descrição'])
            print(new_df['Valor'] == response['Valor'])
            print(new_df['Data do registro'] == response['Data'])
            print(new_df['Tipo'] == response['Tipo'])
            print(new_df['Data de Pagamento'] == response['Data de Pagamento'])
            # A comparação da coluna valor está errada, a IA retorna um valor em reais R$, enquanto na planilha está só o número

            filtro = (
            (new_df['Descrição'] == response['Descrição']) & 
            (new_df['Valor'] == response['Valor']) & 
            (new_df['Data do registro'] == response['Data']) & 
            (new_df['Tipo'] == response['Tipo']) &
            (new_df['Dia de Pagamento'] == response['Dia de Pagamento'])
            )

            print(new_df)
            new_df = new_df.drop(new_df[filtro].index)
            print(new_df)

        case 1:
            new_df.loc[n, 'Descrição'] = response['Descrição']
            new_df.loc[n, 'Valor'] = float(response['Valor'].split('R$ ')[1].replace(',', '.'))
            new_df.loc[n, 'Data do registro'] = response['Data']
            new_df.loc[n, 'Tipo'] = response['Tipo']
            new_df.loc[n, 'Dia de Pagamento'] = response['Dia de Pagamento']
    return new_df


while True:
    df = pd.read_excel('data.xlsx')
    title('FinancIA: Gestor Financeiro')
    match options(['Analisar a planilha', 'register', 'Visualisar Planilha', 'Viver de Renda', 'Sair']):
        case 1:
            print(ia.Gemma4('Analise o dataframe e me diga o que você acha, responda de forma direta, sem criar tópicos ou textos muito grandes.', 300))
            break
        case 2:
            try:
                unsaved_df = registerData(df)

                unsaved_df.to_excel('data.xlsx', index=False) # Unsaved DataFrame
                df = unsaved_df # Saved Dataframe
                print('Receita adicionada com sucesso!')            
            except PermissionError as e:
                print(f'Erro de permissão. Sua planilha está aberta, feche-a para que possa ser modificada. [{e}]')
                print(df)
            except Exception as e:
                print(f'Ocorreu um erro ao adicionar a receita: {e}')
        case 3:
            print(df.to_string(index=False))
        case 4:
            break
