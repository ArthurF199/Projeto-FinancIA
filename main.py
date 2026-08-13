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
Você é um assistente especialista em extração de dados financeiros.
Sua tarefa é analisar a mensagem do usuário e extrair os dados para um objeto JSON estrito.
NÃO retorne nenhuma palavra adicional, explicações ou blocos de código markdown. Retorne APENAS o JSON puro.

Hoje é: {datetime.now().strftime("%d/%m/%Y")}

=== DADOS ATUAIS DA PLANILHA ===
Aqui estão os registros recentes do usuário:
{df.tail(50).fillna('').to_string(index=False)}
=================================

Regras para os campos do JSON:
- "Ação": Inteiro. 1 para adicionar, 0 para remover.
- "Descrição": String curta resumindo o gasto/receita.
- "Valor": Número (float/int) usando ponto para decimais. Ex: 150.50.
- "Data": String no formato "DD/MM/AAAA". Se não informada, use a de hoje ({datetime.now().strftime("%d/%m/%Y")}).
- "Tipo": String categórica ("Entrada", "Saída" ou "Conta").
- "Dia de Pagamento": String "DD/MM/AAAA". Usado APENAS se for uma conta para o futuro.

IMPORTANTE PARA REMOÇÕES (Ação 0): 
Se o usuário pedir para remover um registro, olhe nos "DADOS ATUAIS DA PLANILHA" acima, identifique de qual registro ele está falando e copie os dados EXATOS (Descrição, Valor, Data e Tipo) para preencher o JSON.

Exemplos de extrações perfeitas:

Entrada: Recebi meu salário de R$ 3.500,00 hoje.
Saída: {{"Ação": 1, "Descrição": "Salário", "Valor": 3500.00, "Data": "{datetime.now().strftime("%d/%m/%Y")}", "Tipo": "Entrada", "Dia de Pagamento": null}}

Entrada: Gastei R$ 120,50 no supermercado ontem.
Saída: {{"Ação": 1, "Descrição": "Supermercado", "Valor": 120.50, "Data": "09/08/2026", "Tipo": "Saída", "Dia de Pagamento": null}}

Entrada: Todo dia 10 pago R$ 89,90 da internet.
Saída: {{"Ação": 1, "Descrição": "Internet", "Valor": 89.90, "Data": "{datetime.now().strftime("%d/%m/%Y")}", "Tipo": "Saída", "Dia de Pagamento": "10/08/2026"}}

Entrada: Remova o gasto de restaurante de R$ 300,00.
Saída: {{"Ação": 0, "Descrição": "Restaurante", "Valor": 300.00, "Data": "{datetime.now().strftime("%d/%m/%Y")}", "Tipo": "Saída", "Dia de Pagamento": null}}

Entrada: Hoje chegou uma conta de luz no valor de 200 reais para ser paga no dia 20/05/2026.
Saída: {{"Ação": 1, "Descrição": "Conta de luz", "Valor": 200.00, "Data": "{datetime.now().strftime("%d/%m/%Y")}", "Tipo": "Conta", "Dia de Pagamento": "20/05/2026"}}

Entrada: Paguei a conta de luz que foi registrada no dia 15/05/2026 no valor de 200 reais que era para ser paga no dia 20/05/2026.
Saída: {{"Ação": 0, "Descrição": "Conta de luz", "Valor": 200.00, "Data": "15/05/2026", "Tipo": "Conta", "Dia de Pagamento": "20/05/2026"}}

Agora processe a seguinte entrada do usuário:
Entrada: {prompt}
    """, 1000, stream=False)

    print(repr(response))
    response = json.loads(response)
    new_df = df.fillna('').copy()
    n = len(new_df)
    
    match response['Ação']:
        case 0:
            desc_resp = (response['Descrição'] or '').lower().strip()
            val_resp = (str(response['Valor']) or '').lower().strip()
            data_resp = (response['Data'] or '').lower().strip()
            tipo_resp = (response['Tipo'] or '').lower().strip()
            pag_resp = (response['Dia de Pagamento'] or '').lower().strip()


            filtro = (
            (new_df['Descrição'].astype(str).str.lower().str.strip() == desc_resp) & 
            (new_df['Valor'].astype(str).str.lower().str.strip() == val_resp) & 
            (new_df['Data do registro'].astype(str).str.lower().str.strip() == data_resp) & 
            (new_df['Tipo'].astype(str).str.lower().str.strip() == tipo_resp) &
            (new_df['Dia de Pagamento'].astype(str).str.lower().str.strip() == pag_resp)
            )


            new_df = new_df.drop(new_df[filtro].index)
        case 1:
            new_df.loc[n, 'Descrição'] = response['Descrição'].strip()
            new_df.loc[n, 'Valor'] = float(response['Valor'])
            new_df.loc[n, 'Data do registro'] = response['Data'].strip()
            new_df.loc[n, 'Tipo'] = response['Tipo'].strip()
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
