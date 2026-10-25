import pandas as pd
import ia
import json
import os
from time import sleep
from datetime import datetime, timedelta
import unicodedata
import flet as ft


def saveXLSX(df):
    try:
        while True:
            resposta = input('Deseja salvar as informações? [S/N]').strip()
            if resposta in ('S', 's'):
                df.to_excel('data.xlsx', index=False)
                clear()
                break
            elif resposta in ('N', 'n'):
                clear()
                break
            else:
                print('Input Incorreto!')
    except Exception as e:
        print(f'Input Incorreto {e}')


def clear():
    os.system('cls' if os.name == 'nt' else 'clear')


# Colunas de transações (uma linha por registro)
COLUNAS_REGISTRO = ['Data do registro', 'Descrição', 'Valor', 'Tipo', 'Categoria', 'Dia de Pagamento']
# Colunas de configuração (só a linha 0 é usada)
COLUNAS_CONFIG = ['Salário', 'Reserva de Emergência', 'Viver de Renda', 'Aporte Mensal']

TIPOS = ['Entrada', 'Saída', 'Conta']

# Fonte única das categorias: o prompt é montado a partir daqui e a resposta da IA é validada contra elas
CATEGORIAS_ENTRADA = {
    'Salário': 'salário, 13º, férias, pagamento fixo do emprego',
    'Renda Extra': 'freelas, bicos, vendas, comissões, presentes em dinheiro, reembolsos',
    'Rendimentos': 'juros, dividendos, rendimento de investimentos, aluguel recebido',
}
CATEGORIAS_DESPESA = {
    'Alimentação': 'supermercado, feira, padaria, restaurante, lanche, delivery',
    'Moradia': 'aluguel, condomínio, IPTU, manutenção e itens para a casa',
    'Contas de Casa': 'água, luz, gás, internet, telefone, celular',
    'Transporte': 'combustível, Uber/táxi, ônibus/metrô, estacionamento, manutenção do carro, IPVA',
    'Saúde': 'farmácia, consultas, exames, plano de saúde, academia',
    'Educação': 'escola, faculdade, cursos, livros, material escolar',
    'Lazer': 'cinema, viagens, passeios, festas, bares, jogos',
    'Assinaturas': 'streaming, aplicativos, softwares e outros serviços recorrentes',
    'Compras': 'roupas, calçados, eletrônicos, presentes, cuidados pessoais',
    'Dívidas': 'fatura do cartão, empréstimos, financiamentos, parcelas de acordos',
    'Investimentos': 'aportes, compra de ações, Tesouro, CDB, poupança, reserva de emergência',
}
CATEGORIA_PADRAO = 'Outros'


def numero(valor) -> float:
    # Converte qualquer célula para float; células vazias ou inválidas viram 0
    if isinstance(valor, str):
        valor = valor.replace('R$', '').strip()
        # Formato brasileiro: "1.234,56" -> "1234.56"
        if ',' in valor:
            valor = valor.replace('.', '').replace(',', '.')
    try:
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


def normalizar(texto) -> str:
    # "Saída", "saida" e " SAÍDA " ficam iguais na comparação
    texto = unicodedata.normalize('NFKD', str(texto or '').strip().lower())
    return ''.join(c for c in texto if not unicodedata.combining(c))


def escolher_opcao(valor, opcoes, padrao=''):
    # Devolve a opção com a grafia oficial, ou o padrão se a IA respondeu algo fora da lista
    for opcao in opcoes:
        if normalizar(opcao) == normalizar(valor):
            return opcao
    return padrao


def data_valida(texto, padrao: str = '') -> str:
    # Garante "DD/MM/AAAA"; datas em outro formato ficariam invisíveis nos gráficos
    try:
        return datetime.strptime(str(texto or '').strip(), '%d/%m/%Y').strftime('%d/%m/%Y')
    except ValueError:
        return padrao


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


def proximo_dia(hoje: datetime, dia: int) -> str:
    # Próxima ocorrência do dia do mês a partir de hoje (usado no exemplo de conta recorrente)
    ano, mes = hoje.year, hoje.month
    if hoje.day > dia:
        ano, mes = (ano + 1, 1) if mes == 12 else (ano, mes + 1)
    return f'{dia:02d}/{mes:02d}/{ano}'


def montar_prompt(df: pd.DataFrame, mensagem: str) -> str:
    agora = datetime.now()
    hoje = agora.strftime('%d/%m/%Y')
    ontem = (agora - timedelta(days=1)).strftime('%d/%m/%Y')
    dias_semana = ['segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado', 'domingo']

    # Só os registros de verdade e só as colunas de registro: as colunas de configuração são ruído para a IA
    registros = df[df['Descrição'].astype(str).str.strip() != ''][COLUNAS_REGISTRO].tail(50)
    planilha = registros.to_string(index=False) if not registros.empty else '(nenhum registro ainda)'

    def lista(categorias: dict) -> str:
        return '\n'.join(f'- "{nome}": {exemplos}.' for nome, exemplos in categorias.items())

    vazio = {'Descrição': None, 'Valor': None, 'Data': None, 'Tipo': None, 'Dia de Pagamento': None, 'Categoria': None}
    exemplos = [
        ('Recebi R$ 800,00 de um freela de design.',
         {'Ação': 1, 'Resposta_Chat': None, 'Descrição': 'Freela de design', 'Valor': 800, 'Data': hoje,
          'Tipo': 'Entrada', 'Dia de Pagamento': None, 'Categoria': 'Renda Extra'}),
        ('Gastei 120,50 no mercado ontem',
         {'Ação': 1, 'Resposta_Chat': None, 'Descrição': 'Supermercado', 'Valor': 120.5, 'Data': ontem,
          'Tipo': 'Saída', 'Dia de Pagamento': None, 'Categoria': 'Alimentação'}),
        ('abasteci o carro, deu 200 conto',
         {'Ação': 1, 'Resposta_Chat': None, 'Descrição': 'Combustível', 'Valor': 200, 'Data': hoje,
          'Tipo': 'Saída', 'Dia de Pagamento': None, 'Categoria': 'Transporte'}),
        ('Todo dia 10 pago R$ 89,90 da internet.',
         {'Ação': 1, 'Resposta_Chat': None, 'Descrição': 'Internet', 'Valor': 89.9, 'Data': hoje,
          'Tipo': 'Conta', 'Dia de Pagamento': proximo_dia(agora, 10), 'Categoria': 'Contas de Casa'}),
        ('Coloquei 500 reais na minha reserva de emergência.',
         {'Ação': 1, 'Resposta_Chat': None, 'Descrição': 'Reserva de emergência', 'Valor': 500, 'Data': hoje,
          'Tipo': 'Saída', 'Dia de Pagamento': None, 'Categoria': 'Investimentos'}),
        ('Apaga aquele gasto de 300 no restaurante.',
         {'Ação': 0, 'Resposta_Chat': None, 'Descrição': 'Restaurante', 'Valor': 300, 'Data': hoje,
          'Tipo': 'Saída', 'Dia de Pagamento': None, 'Categoria': 'Alimentação'}),
        ('Vale a pena financiar um carro em 60 vezes?',
         {'Ação': 2, 'Resposta_Chat': 'Geralmente não, porque os juros de 60 meses fazem o carro sair bem mais caro. '
                                      'O ideal é juntar uma boa entrada para financiar menos e por menos tempo.', **vazio}),
    ]
    texto_exemplos = '\n\n'.join(
        f'Mensagem: {entrada}\nJSON: {json.dumps(saida, ensure_ascii=False)}' for entrada, saida in exemplos
    )

    return f"""Você é um extrator de dados financeiros. Leia a mensagem do usuário e responda com UM único objeto JSON, sem nenhum texto antes ou depois e sem markdown.

Hoje é {dias_semana[agora.weekday()]}, {hoje}.

=== REGISTROS ATUAIS DA PLANILHA ===
{planilha}
====================================

PASSO 1 - DEFINA A "Ação" (número inteiro):
- 1: o usuário quer registrar um gasto, uma conta ou um dinheiro recebido.
- 0: o usuário quer remover, apagar ou excluir um registro que está na planilha acima.
- 2: qualquer outra coisa (dúvidas, conselhos, conversa) OU pedido de remoção de um registro que NÃO está na planilha.

PASSO 2 - PREENCHA OS CAMPOS:
- "Resposta_Chat": só na Ação 2. Resposta em texto simples, direta, com no máximo 3 frases. Se o pedido era remover algo que não existe, diga que não encontrou esse registro na planilha. Nas Ações 0 e 1, use null.
- "Descrição": nome curto do gasto ou recebimento, com inicial maiúscula (ex: "Supermercado", "Conta de luz").
- "Valor": número positivo com ponto decimal, sem "R$" e sem sinal de menos (ex: 120.5). Gastos também são positivos.
- "Data": data do registro em "DD/MM/AAAA". Converta datas relativas ("ontem", "sexta passada") a partir de hoje. Sem data na mensagem, use {hoje}.
- "Tipo": exatamente um destes:
  - "Entrada": dinheiro que o usuário recebeu.
  - "Saída": gasto pontual, já feito.
  - "Conta": conta recorrente ou com vencimento (aluguel, luz, internet, mensalidade, parcela, fatura).
- "Dia de Pagamento": só quando "Tipo" for "Conta": a próxima data de vencimento em "DD/MM/AAAA" (ex: "todo dia 10" = o próximo dia 10 a partir de hoje). Nos outros casos, null.
- "Categoria": veja o passo 3.
Na Ação 2, todos os campos, exceto "Ação" e "Resposta_Chat", são null.

PASSO 3 - ESCOLHA A "Categoria" (copie o nome exatamente como está, com acentos; nunca invente uma categoria nova):
Se "Tipo" for "Entrada":
{lista(CATEGORIAS_ENTRADA)}
Se "Tipo" for "Saída" ou "Conta":
{lista(CATEGORIAS_DESPESA)}
Para qualquer "Tipo":
- "{CATEGORIA_PADRAO}": somente quando nenhuma das categorias acima servir.
Como decidir:
1. Pense no que foi comprado, pago ou recebido (produto, local ou motivo), não só nas palavras usadas. Ex: "abasteci" = "Transporte"; "comprei remédio" = "Saúde"; "iFood" = "Alimentação".
2. A categoria tem que pertencer ao grupo do "Tipo" escolhido.
3. Se servirem duas categorias, escolha a do objetivo principal do gasto ou recebimento.
4. Se a planilha já tiver um registro parecido com categoria preenchida, repita essa categoria.

REMOÇÕES (Ação 0):
Ache na planilha o registro citado e copie EXATAMENTE os valores dele: "Descrição", "Valor", "Tipo", "Categoria" e "Data" (a coluna "Data do registro"). Se a categoria estiver vazia na planilha, use null. Se mais de um registro combinar, escolha o mais recente.

EXEMPLOS (as datas consideram que hoje é {hoje}):

{texto_exemplos}

Agora responda apenas com o JSON para a mensagem abaixo.
Mensagem: {mensagem}
JSON:"""


def registerData(df, prompt):
    print('Carregando...')
    # temperatura 0 deixa a extração previsível, e formato "json" obriga o Ollama a devolver um JSON válido
    resposta = ia.Gemma4(montar_prompt(df, prompt), 1000, stream=False, formato='json', temperatura=0)

    response = extrair_json(resposta)
    new_df = df.fillna('').reset_index(drop=True)
    n = len(new_df)

    # A IA às vezes manda a ação como texto ("1"); qualquer coisa inválida vira bate-papo
    try:
        acao = int(response.get('Ação', 2))
    except (TypeError, ValueError):
        acao = 2

    hoje = datetime.now().strftime('%d/%m/%Y')
    descricao = str(response.get('Descrição') or '').strip()
    valor = abs(numero(response.get('Valor')))
    tipo = escolher_opcao(response.get('Tipo'), TIPOS)

    match acao:
        case 0:
            # Valor é comparado como número: "300" (IA) e 300.0 (planilha) são iguais
            filtro = (
                (new_df['Descrição'].apply(normalizar) == normalizar(descricao)) &
                (new_df['Valor'].apply(numero) == valor) &
                (new_df['Tipo'].apply(normalizar) == normalizar(tipo))
            )
            encontrados = new_df[filtro]

            # Se houver mais de um igual, a data desempata
            data_resp = data_valida(response.get('Data'))
            mesma_data = encontrados[encontrados['Data do registro'].astype(str).str.strip() == data_resp]
            if not mesma_data.empty:
                encontrados = mesma_data

            if encontrados.empty:
                return (new_df, 2, 'Não encontrei esse registro na planilha, então nada foi removido.')

            # Remove apenas um registro por pedido (o mais recente, como o prompt pede)
            new_df = remover_registro(new_df, encontrados.index[-1])
        case 1:
            if descricao == '' or valor == 0 or tipo == '':
                return (new_df, 2, 'Não consegui entender o registro. Diga o que foi, o valor e se foi um gasto ou um recebimento.')

            # A categoria precisa ser do mesmo grupo do tipo; fora disso vira "Outros"
            categorias = CATEGORIAS_ENTRADA if tipo == 'Entrada' else CATEGORIAS_DESPESA
            categoria = escolher_opcao(response.get('Categoria'), categorias, CATEGORIA_PADRAO)

            # Se a planilha só tem a linha de configuração vazia, reaproveita ela
            if n == 1 and str(new_df.loc[0, 'Descrição']).strip() == '':
                n = 0

            new_df.loc[n, 'Descrição'] = descricao
            new_df.loc[n, 'Valor'] = valor
            new_df.loc[n, 'Data do registro'] = data_valida(response.get('Data'), hoje)
            new_df.loc[n, 'Tipo'] = tipo
            new_df.loc[n, 'Categoria'] = categoria
            # Dia de pagamento só faz sentido para contas
            new_df.loc[n, 'Dia de Pagamento'] = data_valida(response.get('Dia de Pagamento')) if tipo == 'Conta' else ''
        case _:
            return (new_df, 2, str(response.get('Resposta_Chat') or '').strip())
    return (new_df.fillna(''), acao, '')


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
