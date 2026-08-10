import pandas as pd
import ia
import json
import os
from time import sleep
from datetime import datetime
import flet as ft


def teste():
    return 'fodase'


def saveXLSX(df):
    try:
        while True:
            if input('Deseja salvar as informações? [S/N]') in 'Ss':
                df.to_excel('data.xlsx', index=False)
                clear()
                break
            elif input('Deseja salvar as informações? [S/N]') in 'Nn':
                clear()
                break
            else:
                print('Input Incorreto!')
    except Exception as e:
        print(f'Input Incorreto {e}')


def clear():
    os.system('cls' if os.name == 'nt' else 'clear')


def title(title: str):
    print('-'*10, title, '-'*10)


def options(options: list):
    for i, option in enumerate(options):
        print(f'{i+1} - {option}')
    return int(input('Digite o número da opção desejada: '))


def registerData(df, prompt):
    print(('Carregando...'))
    response = ia.Gemma4(f"""
              Hoje: {datetime.now().strftime("%d/%m%Y")}             
            
              Vou te enviar uma descrição de um registro.
              Quero que você determine se o usuário quer adicionar ou remover uma receita e me retorne um JSON com os seguintes campos:
              - Ação: Adicionar/remover o registro, deixe 0 para remover e 1 para adicionar um dado à planilha
                (caso o usuário não deixe claro se ele deseja adicionar ou remover, considere deixar em 1)
              - Descrição: uma breve descrição da receita
                (caso não seja informada, escreva null no campo)
              - Valor: o valor da receita, seguindo o formato R$ 000.00.
                (caso não seja informada, escreva null no campo)
              - Data: a data do registro, seguindo o formato DD-MM-AAAA
                (caso a data não for informada, considere a de hoje({datetime.now().strftime("%d/%m%Y")}))         
              - Tipo: se é entrada ou saída ou conta.
                (caso não for informada, escreva null no campo)
              - Dia de pagamento: somente se for descrito como uma conta, siga o mesmo formato do campo Data(DD-MM-AA)
                (caso não for informada, escreva null no campo)
              Siga o seguinte exemplo: {{"Ação": 0 ou 1, "Descrição": "Descrição da receita", "Valor": "R$ 0000", "Data": "DD-MM-AAAA", "Tipo": "Entrada ou Saída de dinheiro", "Dia de Pagamento": "DD-MM-AAAA (somente se for necessário)"}}
              Quero que me mande somente a receita, sem nenhuma explicação ou formatação adicional.
                         
              Considere como não informada apenas quando realmente o usuário não descreveu a informação, tente ao máximo extrair as informações que o usuário descrever.
              Considere como ação remover quando o usuário diz algo como: "quero remover tal conta de tal dia com tal descrição"
              
              Considere os seguintes exemplos de extrações corretas:
              
              Entrada: Recebi meu salário de R$ 3.500,00 hoje.
              Saída: "Ação": 1, "Descrição": "Salário", "Valor": "R$ 3500", "Data": "16-06-2026", "Tipo": "Entrada", "Dia de Pagamento": null
                         
              Entrada: Gastei R$ 120,50 no supermercado ontem.
              Saída: "Ação": 1, "Descrição": "Supermercado", "Valor": "R$ 120.50", "Data": "15-06-2026", "Tipo": "Saída", "Dia de Pagamento": null
               
              Entrada: Todo dia 10 pago R$ 89,90 da internet.
              Saída: "Ação": 1, "Descrição": "Internet", "Valor": "R$ 89.90", "Data": null, "Tipo": "Saída", "Dia de Pagamento": "10-07-2026"
              
              Entrada: Remova o gasto de restaurante de R$ 300,00.
              Saída: "Ação": 0, "Descrição": "Restaurante", "Valor": "R$ 300", "Data": null, "Tipo": "Saída", "Dia de Pagamento": null
              
              Entrada: Recebi um dinheiro de um freela.
              Saída: "Ação": 1, "Descrição": "Freelance", "Valor": null, "Data": {datetime.now().strftime("%d/%m%Y")}, "Tipo": "Entrada", "Dia de Pagamento": null

              Entrada: Hoje chegou uma conta de luz no valor de 200 reais para ser paga no dia 20/05/2026.
              Saída: "Ação": 1, "Descrição": "Conta de luz", "Valor": 200, "Data": {datetime.now().strftime("%d/%m%Y")}, "Tipo": "Saída", "Dia de Pagamento": 20-05-2026

              Entrada: Paguei a conta de luz que foi registrada no dia 15/05/2026 no valor de 200 reais que era para ser paga no dia 20/05/2026.
              Saída: "Ação": 0, "Descrição": "Conta de luz", "Valor": 200, "Data": 15/05/2026, "Tipo": "Saída", "Dia de Pagamento": 20-05-2026

              Entrada: Paguei a conta de 500 reais que era para ser paga no dia 11 do 8 de 2026
              Saída: "Ação": 0, "Descrição": "Conta a pagar", "Valor": 500, "Data": {datetime.now().strftime("%d/%m%Y")}, "Tipo": "Saída", "Dia de Pagamento": 11-08-2026
              A descrição do registro é: {prompt}
    """, 1000, stream=False)

    print(repr(response))
    response = json.loads(response)
    new_df = df.fillna('').copy()
    n = len(new_df)
    
    match response['Ação']:
        case 0:
            desc_resp = (response['Descrição'] or '').lower()
            val_resp = (response['Valor'] or '').lower()
            data_resp = (response['Data'] or '').lower()
            tipo_resp = (response['Tipo'] or '').lower()
            pag_resp = (response['Dia de Pagamento'] or '').lower()


            filtro = (
            (new_df['Descrição'].astype(str).str.lower() == desc_resp) & 
            (new_df['Valor'].astype(str).str.lower() == val_resp) & 
            (new_df['Data do registro'].astype(str).str.lower() == data_resp) & 
            (new_df['Tipo'].astype(str).str.lower() == tipo_resp) &
            (new_df['Dia de Pagamento'].astype(str).str.lower() == pag_resp)
            )


            new_df = new_df.drop(new_df[filtro].index)
        case 1:
            new_df.loc[n, 'Descrição'] = response['Descrição']
            new_df.loc[n, 'Valor'] = response['Valor']
            new_df.loc[n, 'Data do registro'] = response['Data']
            new_df.loc[n, 'Tipo'] = response['Tipo']
            new_df.loc[n, 'Dia de Pagamento'] = response['Dia de Pagamento']
    return (new_df, response['Ação'])


clear()
# while True:
#     ft.app(target=app.main)
#     df = pd.read_excel('data.xlsx')
#     title('FinancIA: Gestor Financeiro')
#     print(df.fillna('').to_string(index=False))
#     sleep(.5)
#     match options(['Registrar Salário', 'Analisar a planilha', 'Registrar/Remover Dados', 'Reserva de Emergência', 'Viver de Renda', 'Sair']):
#         case 1:
#             clear()
#             df['Salário'] = df['Salário'].astype(str)
#             df.loc[0, 'Salário'] = 'R$ ' + input('Qual o seu salário: R$')
#             saveXLSX(df)
#         case 2:
#             clear()
#             title('Análise da planilha')
#             sleep(.5)
#             print('Carregando...')
#             response = ia.Gemma4(f'Analise o dataframe: {df} e me diga se eu estou fazendo um bom gerenciamento do meu dinheiro e me dê sugestões do que eu deveria fazer, responda de forma direta, sem criar tópicos ou textos muito grandes.', 1000)
#             clear()
#             title('Análise da planilha')
#             if response is not None or not '':
#                 print(response)
#                 sleep(.5)
#                 input('Pressione ENTER para continuar.')
#             else:
#                 print('Erro, tente novamente.')
#                 sleep(.5)
#                 input('Pressione ENTER para continuar.')
#                 continue
#         case 3:
#             clear()
#             title('Registro de dados')
#             sleep(.5)
#             print(df.fillna('').to_string(index=False))
#             try:
#                 temp = registerData(df, input('Descreva o registro: ')) # temp solution
#                 unsaved_df = temp[0]
#                 action = temp[1]
#                 del temp

#                 match action:
#                     case 0:
#                         print('Receita removida com sucesso!') 
#                     case 1:
#                         print('Receita adicionada com sucesso!') 
#                 saveXLSX(unsaved_df) # Unsaved DataFrame
#                 df = unsaved_df # Saved Dataframe

#             except PermissionError as e:
#                 print(f'Erro de permissão. Sua planilha está aberta, feche-a para que possa ser modificada. [{e}]')
#                 input('Pressione ENTER para continuar.')
#             except Exception as e:
#                 print(f'Ocorreu um erro ao adicionar a receita: {e}')
#                 input('Pressione ENTER para continuar.')
#         case 4:
#             clear()
#             df.loc['Reserva de Emergência'] = df.loc[0, 'Reserva de Emergência'].astype(str) # bug aqui
#             df.loc[0, 'Reserva de Emergência'] = (
#                 f" R$ {int(df.loc[0, 'Salário']
#                 .strip()
#                 .replace('R$', '')
#                 ) * 6}"
#             )

#             clear()
#             title('Reserva de Emergência')
#             sleep(.5)
#             print(f'Sua reserva de emergência é de {df.loc[0, 'Reserva de Emergência']}')
#             sleep(.5)
#             saveXLSX(df)
#         case 5:
#             df.loc['Aporte Mensal'] = df.loc[0, 'Aporte Mensal'].astype(str) # possivel bug aqui
#             df.loc[0, 'Aporte Mensal'] = (
#                 f"R$ {int(df.loc[0, 'Salário']
#                 .strip()
#                 .replace('R$', '')
#                 ) * 0.2:.0f}"
#             )

#             df.loc[0, 'Viver de Renda'] = (
#                 f"R$ {int(df.loc[0, 'Salário']
#                 .strip()
#                 .replace('R$', '')
#                 ) * 120}"
#             )

#             clear()
#             title('Viver de Renda')
#             sleep(.5)
#             print(f'Para viver de renda você precisa chegar a {df.loc[0, 'Viver de Renda']} investidos\nPara isso, você precisa de um aporte mensal de {df.loc[0, 'Aporte Mensal']}')
#             sleep(.5)
#             saveXLSX(df)
#         case 6:
#             break
