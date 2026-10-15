import pandas as pd
import ia
import json
import os
from time import sleep
from datetime import datetime
import flet as ft


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


# Colunas de transações (uma linha por registro)
COLUNAS_REGISTRO = ['Data do registro', 'Descrição', 'Valor', 'Tipo', 'Dia de Pagamento']
# Colunas de configuração (só a linha 0 é usada)
COLUNAS_CONFIG = ['Salário', 'Reserva de Emergência', 'Viver de Renda', 'Aporte Mensal']


def numero(valor) -> float:
    # Converte qualquer célula para float; células vazias ou inválidas viram 0
    try:
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


def carregar_planilha(caminho: str) -> pd.DataFrame:
    if os.path.exists(caminho):
        df = pd.read_excel(caminho)
    else:
        df = pd.DataFrame(columns=COLUNAS_REGISTRO + COLUNAS_CONFIG)

    # Remove colunas de índice que foram salvas por engano e totais antigos (agora são calculados na hora)
    df = df.loc[:, ~df.columns.astype(str).str.startswith('Unnamed')]
    df = df.drop(columns=['Entradas', 'Despesas'], errors='ignore')

    for coluna in COLUNAS_REGISTRO + COLUNAS_CONFIG:
        if coluna not in df.columns:
            df[coluna] = ''

    df = df.astype(object).fillna('').reset_index(drop=True)

    # A linha 0 guarda as configurações, então ela precisa existir mesmo sem registros
    if df.empty:
        df.loc[0] = ''
    return df


def remover_registro(df: pd.DataFrame, indice) -> pd.DataFrame:
    # Guarda as configurações antes de apagar, pois elas vivem na linha 0
    config = df.loc[0, COLUNAS_CONFIG].copy()

    df = df.drop(indice).reset_index(drop=True)
    if df.empty:
        df.loc[0] = ''

    df.loc[0, COLUNAS_CONFIG] = config.values
    return df


def extrair_json(texto: str) -> dict:
    # O modelo às vezes envolve o JSON em ```json ... ```, então pegamos só o trecho entre chaves
    inicio, fim = texto.find('{'), texto.rfind('}')
    if inicio == -1 or fim == -1:
        raise ValueError(f'A IA não retornou um JSON válido: {texto}')
    return json.loads(texto[inicio:fim + 1])


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
- "Ação": Inteiro. 1 para adicionar, 0 para remover, 2 para dúvidas financeiras ou bate-papo
- "Resposta_Chat": Se a ação for 2, escreva aqui a sua dica financeira. Se for 0 ou 1, retorne null
- "Descrição": String curta resumindo o gasto/receita.
- "Valor": Número (float/int) usando ponto para decimais. Ex: 150.50.
- "Data": String no formato "DD/MM/AAAA". Se não informada, use a de hoje ({datetime.now().strftime("%d/%m/%Y")}).
- "Tipo": String categórica ("Entrada", "Saída" ou "Conta").
- "Dia de Pagamento": String "DD/MM/AAAA". Usado APENAS se for uma conta para o futuro.

IMPORTANTE PARA REMOÇÕES (Ação 0): 
Se o usuário pedir para remover um registro, olhe nos "DADOS ATUAIS DA PLANILHA" acima, identifique de qual registro ele está falando e copie os dados EXATOS (Descrição, Valor, Data e Tipo) para preencher o JSON.

CASO O USUÁRIO PEÇA PARA REMOVER ALGO QUE NÃO ESTEJA NA PLANILHA RETORNE A AÇÃO 2 E DIGA QUE NÃO TEM ACESSO A TAL REGISTRO.

Exemplos de extrações perfeitas:

Entrada: Recebi meu salário de R$ 3.500,00 hoje.
Saída: {{"Ação": 1, "Resposta_Chat": null, "Descrição": "Salário", "Valor": 3500, "Data": "{datetime.now().strftime("%d/%m/%Y")}", "Tipo": "Entrada", "Dia de Pagamento": null}}

Entrada: Gastei R$ 120,50 no supermercado ontem.
Saída: {{"Ação": 1, "Resposta_Chat": null, "Descrição": "Supermercado", "Valor": 120.50, "Data": "15/08/2026", "Tipo": "Saída", "Dia de Pagamento": null}}

Entrada: Todo dia 10 pago R$ 89,90 da internet.
Saída: {{"Ação": 1, "Resposta_Chat": null, "Descrição": "Internet", "Valor": 89.90, "Data": "{datetime.now().strftime("%d/%m/%Y")}", "Tipo": "Saída", "Dia de Pagamento": "10/08/2026"}}

Entrada: Remova o gasto de restaurante de R$ 300,00.
Saída: {{"Ação": 0, "Resposta_Chat": null, "Descrição": "Restaurante", "Valor": 300, "Data": "{datetime.now().strftime("%d/%m/%Y")}", "Tipo": "Saída", "Dia de Pagamento": null}}

Entrada: Como devo começar a montar minha reserva de emergência?
Saída: {{"Ação": 2, "Resposta_Chat": "Para começar, calcule seu custo de vida mensal e multiplique por 6. Guarde esse valor gradualmente em investimentos seguros e de alta liquidez, como Tesouro Selic ou um CDB que renda 100% do CDI.", "Descrição": null, "Valor": null, "Data": null, "Tipo": null, "Dia de Pagamento": null}}

Entrada: Vale a pena financiar um carro em 60 vezes?
Saída: {{"Ação": 2, "Resposta_Chat": "Geralmente não é recomendado, pois o acúmulo de juros ao longo de 60 meses fará você pagar um valor muito maior pelo veículo. O ideal é poupar para dar uma boa entrada e reduzir os juros e o tempo de financiamento.", "Descrição": null, "Valor": null, "Data": null, "Tipo": null, "Dia de Pagamento": null}}

Entrada: O que é o CDI?
Saída: {{"Ação": 2, "Resposta_Chat": "O CDI (Certificado de Depósito Interbancário) é uma taxa que os bancos usam para emprestar dinheiro entre si. Ele serve como principal referência (benchmark) para a rentabilidade dos investimentos de renda fixa no Brasil.", "Descrição": null, "Valor": null, "Data": null, "Tipo": null, "Dia de Pagamento": null}}

Agora processe a seguinte entrada do usuário:
Entrada: {prompt}
    """, 1000, stream=False)

    response = extrair_json(response)
    new_df = df.fillna('').reset_index(drop=True)
    n = len(new_df)

    match response['Ação']:
        case 2:
            return (new_df, 2, response['Resposta_Chat'] or '')
        case 0:
            desc_resp = (response['Descrição'] or '').lower().strip()
            data_resp = (response['Data'] or '').strip()
            tipo_resp = (response['Tipo'] or '').lower().strip()

            # Valor é comparado como número: "300" (IA) e 300.0 (planilha) são iguais
            filtro = (
                (new_df['Descrição'].astype(str).str.lower().str.strip() == desc_resp) &
                (new_df['Valor'].apply(numero) == numero(response['Valor'])) &
                (new_df['Tipo'].astype(str).str.lower().str.strip() == tipo_resp)
            )
            encontrados = new_df[filtro]

            # Se houver mais de um igual, a data desempata
            mesma_data = encontrados[encontrados['Data do registro'].astype(str).str.strip() == data_resp]
            if not mesma_data.empty:
                encontrados = mesma_data

            if encontrados.empty:
                return (new_df, 2, 'Não encontrei esse registro na planilha, então nada foi removido.')

            # Remove apenas um registro por pedido
            new_df = remover_registro(new_df, encontrados.index[0])
        case 1:
            # Se a planilha só tem a linha de configuração vazia, reaproveita ela
            if n == 1 and str(new_df.loc[0, 'Descrição']).strip() == '':
                n = 0

            new_df.loc[n, 'Descrição'] = (response['Descrição'] or '').strip()
            new_df.loc[n, 'Valor'] = numero(response['Valor'])
            new_df.loc[n, 'Data do registro'] = (response['Data'] or datetime.now().strftime('%d/%m/%Y')).strip()
            new_df.loc[n, 'Tipo'] = (response['Tipo'] or '').strip()
            new_df.loc[n, 'Dia de Pagamento'] = response['Dia de Pagamento'] or ''
    return (new_df.fillna(''), response['Ação'], '')


# clear()
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
