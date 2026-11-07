import pandas as pd
import ia
import json
import os
import calendar
from datetime import datetime, timedelta
import unicodedata


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
# Status das contas ("Sim"/"Não"). Fica fora de COLUNAS_REGISTRO: na tela ele vira um indicador, não uma coluna
COLUNA_PAGO = 'Pago'
# Valores do salário ao longo do tempo ("2026-01=3000.0;2026-06=3500.0"): cada valor vale a partir daquele mês.
# Sem ele, mudar o salário mudaria também os meses que já passaram
COLUNA_HISTORICO_SALARIO = 'Histórico do Salário'
# Colunas de configuração (só a linha 0 é usada)
COLUNAS_CONFIG = ['Salário', 'Reserva de Emergência', 'Viver de Renda', 'Aporte Mensal', COLUNA_HISTORICO_SALARIO]

TIPOS = ['Entrada', 'Saída', 'Conta']
TIPOS_DESPESA = ['Saída', 'Conta']
CATEGORIA_INVESTIMENTO = 'Investimentos'
CATEGORIA_SALARIO = 'Salário'
# Entradas da categoria Salário que são um extra e não o salário do mês (o salário configurado continua contando)
EXTRAS_DE_SALARIO = ('13', 'decimo', 'ferias', 'plr', 'participacao nos lucros', 'bonus', 'bonificacao', 'rescisao')
METRICAS_CONSULTA = ['Despesas', 'Entradas', 'Investido', 'Saldo']
NOMES_MESES_COMPLETOS = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                         'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

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


def moeda(valor: float) -> str:
    # Formato brasileiro: R$ 1.234,56
    return 'R$ ' + f'{valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def normalizar(texto) -> str:
    # "Saída", "saida" e " SAÍDA " ficam iguais na comparação
    texto = unicodedata.normalize('NFKD', str(texto or '').strip().lower())
    return ''.join(c for c in texto if not unicodedata.combining(c))


def esta_paga(valor) -> bool:
    # Aceita o que alguém digitaria no Excel: "Sim", "sim", "x", "pago"...
    return normalizar(valor) in ('sim', 's', 'x', 'pago', 'paga', 'true', '1')


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


def linha_vazia(colunas) -> pd.DataFrame:
    # Planilha com só uma linha vazia. Criada já como "object": com df.loc[0] = '' o pandas transformaria
    # as colunas em texto, e gravar um número (Valor, Salário) depois quebraria
    return pd.DataFrame([[''] * len(colunas)], columns=colunas, dtype=object)


def carregar_planilha(caminho: str) -> pd.DataFrame:
    if os.path.exists(caminho):
        df = pd.read_excel(caminho)
    else:
        df = pd.DataFrame(columns=COLUNAS_REGISTRO + COLUNAS_CONFIG)

    # Remove colunas de índice que foram salvas por engano e totais antigos (agora são calculados na hora)
    df = df.loc[:, ~df.columns.astype(str).str.startswith('Unnamed')]
    df = df.drop(columns=['Entradas', 'Despesas'], errors='ignore')

    for coluna in COLUNAS_REGISTRO + [COLUNA_PAGO] + COLUNAS_CONFIG:
        if coluna not in df.columns:
            df[coluna] = ''

    df = df.astype(object).fillna('').reset_index(drop=True)

    # A linha 0 guarda as configurações, então ela precisa existir mesmo sem registros
    if df.empty:
        df = linha_vazia(df.columns)
    return df


def remover_registro(df: pd.DataFrame, indice) -> pd.DataFrame:
    # Guarda as configurações antes de apagar, pois elas vivem na linha 0
    config = df.loc[0, COLUNAS_CONFIG].copy()

    df = df.drop(indice).reset_index(drop=True)
    if df.empty:
        df = linha_vazia(df.columns)

    df.loc[0, COLUNAS_CONFIG] = config.values
    return df


# ==========================================
# CÁLCULOS (usados pela tela e pelas consultas do chat)
# ==========================================
def preparar_registros(df: pd.DataFrame) -> pd.DataFrame:
    # Apenas as linhas que são transações de verdade, com colunas auxiliares já convertidas
    tabela = df[df['Descrição'].astype(str).str.strip() != ''].copy()
    tabela['_data'] = pd.to_datetime(tabela['Data do registro'], format='%d/%m/%Y', errors='coerce')
    tabela['_valor'] = tabela['Valor'].apply(numero)
    tabela['_tipo'] = tabela['Tipo'].astype(str).str.strip()
    tabela['_categoria'] = tabela['Categoria'].astype(str).str.strip().replace('', CATEGORIA_PADRAO)
    tabela['_vencimento'] = pd.to_datetime(tabela['Dia de Pagamento'], format='%d/%m/%Y', errors='coerce')
    tabela['_pago'] = tabela[COLUNA_PAGO].apply(esta_paga)
    return tabela


def eh_investimento(tabela):
    return tabela['_tipo'].isin(TIPOS_DESPESA) & (tabela['_categoria'] == CATEGORIA_INVESTIMENTO)


def eh_despesa(tabela):
    # Dinheiro investido sai da conta, mas não é gasto: ele aparece separado
    return tabela['_tipo'].isin(TIPOS_DESPESA) & (tabela['_categoria'] != CATEGORIA_INVESTIMENTO)


def eh_salario_mensal(descricao) -> bool:
    # "13º salário", "Férias", "PLR"... são extras: não substituem o salário do mês
    texto = normalizar(descricao)
    return not any(extra in texto for extra in EXTRAS_DE_SALARIO)


def historico_salario(df: pd.DataFrame) -> list:
    # Lista ordenada de (ano, mês, valor): cada valor vale daquele mês em diante
    historico = []
    for parte in str(df.loc[0, COLUNA_HISTORICO_SALARIO] or '').split(';'):
        try:
            mes_ano, valor = parte.split('=')
            ano, mes = map(int, mes_ano.split('-'))
            historico.append((ano, mes, float(valor)))
        except ValueError:
            continue
    if not historico:
        salario = numero(df.loc[0, 'Salário'])
        if salario > 0:
            # Planilha de antes do histórico: o salário passa a valer a partir do primeiro lançamento
            datas = pd.to_datetime(df['Data do registro'], format='%d/%m/%Y', errors='coerce').dropna()
            inicio = datas.min() if not datas.empty else datetime.now()
            historico.append((inicio.year, inicio.month, salario))
    return sorted(historico)


def salario_do_mes(historico: list, ano: int, mes: int) -> float:
    valor = 0.0
    for ano_inicio, mes_inicio, salario in historico:
        if (ano_inicio, mes_inicio) <= (ano, mes):
            valor = salario
    return valor


def definir_salario(df: pd.DataFrame, valor: float, hoje: datetime | None = None) -> pd.DataFrame:
    # Novo salário vale a partir deste mês; os meses anteriores mantêm o valor que tinham
    hoje = hoje or datetime.now()
    historico = [h for h in historico_salario(df) if (h[0], h[1]) != (hoje.year, hoje.month)]
    historico.append((hoje.year, hoje.month, float(valor)))

    novo_df = df.copy()
    novo_df.loc[0, COLUNA_HISTORICO_SALARIO] = ';'.join(f'{a}-{m:02d}={v}' for a, m, v in sorted(historico))
    novo_df.loc[0, 'Salário'] = valor
    # Reserva = 6 salários, Viver de Renda = 120 salários, Aporte = 20% do salário
    novo_df.loc[0, 'Reserva de Emergência'] = valor * 6
    novo_df.loc[0, 'Viver de Renda'] = valor * 120
    novo_df.loc[0, 'Aporte Mensal'] = round(valor * 0.2, 2)
    return novo_df


def totais(tabela: pd.DataFrame, historico: list | None = None, meses_salario=()) -> tuple:
    # (entradas, despesas, investido) de um recorte da tabela. O salário configurado entra em cada mês de
    # meses_salario, a não ser que o salário daquele mês já tenha sido lançado pelo chat
    entradas = float(tabela.loc[tabela['_tipo'] == 'Entrada', '_valor'].sum())
    for ano, mes in meses_salario:
        do_mes = tabela[(tabela['_data'].dt.year == ano) & (tabela['_data'].dt.month == mes)]
        salario_lancado = ((do_mes['_tipo'] == 'Entrada') & (do_mes['_categoria'] == CATEGORIA_SALARIO)
                           & do_mes['Descrição'].apply(eh_salario_mensal)).any()
        if not salario_lancado:
            entradas += salario_do_mes(historico or [], ano, mes)
    despesas = float(tabela.loc[eh_despesa(tabela), '_valor'].sum())
    investido = float(tabela.loc[eh_investimento(tabela), '_valor'].sum())
    return entradas, despesas, investido


def ultimo_dia_do_mes(data) -> pd.Timestamp:
    return pd.Timestamp(data.year, data.month, calendar.monthrange(data.year, data.month)[1])


def meses_completos(inicio: pd.Timestamp, fim: pd.Timestamp, hoje: pd.Timestamp) -> list:
    # Meses que o período cobre inteiros (no mês atual, até hoje): só neles o salário entra na conta
    meses = []
    atual = pd.Timestamp(inicio.year, inicio.month, 1)
    while atual <= min(fim, hoje):
        if inicio <= atual and fim >= min(ultimo_dia_do_mes(atual), hoje):
            meses.append((atual.year, atual.month))
        atual = ultimo_dia_do_mes(atual) + timedelta(days=1)
    return meses


def descrever_periodo(inicio: pd.Timestamp, fim: pd.Timestamp, hoje: pd.Timestamp) -> str:
    if inicio == fim:
        return 'Hoje' if inicio == hoje else f'Em {inicio:%d/%m/%Y}'
    mesmo_mes = (inicio.year, inicio.month) == (fim.year, fim.month)
    if inicio.day == 1 and mesmo_mes and fim in (ultimo_dia_do_mes(inicio), hoje):
        nome = f'{NOMES_MESES_COMPLETOS[inicio.month - 1].lower()} de {inicio.year}'
        return f'Em {nome} (até hoje)' if fim == hoje and fim != ultimo_dia_do_mes(inicio) else f'Em {nome}'
    return f'De {inicio:%d/%m/%Y} a {fim:%d/%m/%Y}'


def responder_consulta(df: pd.DataFrame, consulta: dict, agora: datetime | None = None) -> str:
    # A IA só diz o que a pessoa quer saber (métrica, período e filtros); a conta é feita aqui,
    # porque o modelo erra somas
    hoje = pd.Timestamp((agora or datetime.now()).date())
    metrica = escolher_opcao(consulta.get('Métrica'), METRICAS_CONSULTA, 'Despesas')
    texto_inicio = data_valida(consulta.get('Início'))
    if not texto_inicio:
        return 'Não entendi o período da pergunta. Tente algo como "quanto gastei em outubro?".'
    inicio = pd.Timestamp(datetime.strptime(texto_inicio, '%d/%m/%Y'))
    texto_fim = data_valida(consulta.get('Fim'))
    fim = pd.Timestamp(datetime.strptime(texto_fim, '%d/%m/%Y')) if texto_fim else min(ultimo_dia_do_mes(inicio), hoje)
    if fim < inicio:
        inicio, fim = fim, inicio

    todas = list(CATEGORIAS_ENTRADA) + list(CATEGORIAS_DESPESA) + [CATEGORIA_PADRAO]
    categoria = escolher_opcao(consulta.get('Categoria'), todas)
    busca = str(consulta.get('Busca') or '').strip()
    if categoria == CATEGORIA_INVESTIMENTO and metrica == 'Despesas':
        metrica = 'Investido'

    tabela = preparar_registros(df)
    filtrado = tabela[(tabela['_data'] >= inicio) & (tabela['_data'] <= fim)]
    if categoria:
        filtrado = filtrado[filtrado['_categoria'] == categoria]
    if busca:
        filtrado = filtrado[filtrado['Descrição'].apply(normalizar).str.contains(normalizar(busca), regex=False)]

    # O salário configurado só entra quando a pergunta não filtra outra coisa
    meses = meses_completos(inicio, fim, hoje) if not busca and categoria in ('', CATEGORIA_SALARIO) else []
    entradas, despesas, investido = totais(filtrado, historico_salario(df), meses)
    salario_somado = entradas - totais(filtrado)[0]

    quando = descrever_periodo(inicio, fim, hoje)
    filtro = f' com {busca}' if busca else (f' em {categoria}' if categoria else '')

    if metrica == 'Despesas':
        gastos = filtrado[eh_despesa(filtrado)]
        if gastos.empty:
            return f'{quando} você não teve gastos{filtro} registrados.'
        texto = f'{quando} você gastou {moeda(despesas)}{filtro} ({len(gastos)} lançamento{"s" if len(gastos) != 1 else ""}).'
        if not categoria and not busca:
            maiores = gastos.groupby('_categoria')['_valor'].sum().sort_values(ascending=False).head(3)
            texto += ' Onde mais gastou: ' + ', '.join(f'{nome} ({moeda(valor)})' for nome, valor in maiores.items()) + '.'
        return texto
    if metrica == 'Entradas':
        if entradas <= 0:
            return f'{quando} você não teve entradas{filtro} registradas.'
        extra = f', contando seu salário de {moeda(salario_somado)}' if salario_somado > 0 else ''
        return f'{quando} você recebeu {moeda(entradas)}{filtro}{extra}.'
    if metrica == 'Investido':
        if investido <= 0:
            return f'{quando} você não registrou investimentos{filtro}.'
        return f'{quando} você investiu {moeda(investido)}{filtro}.'

    saldo = entradas - despesas - investido
    return (f'{quando} {"sobraram" if saldo >= 0 else "faltaram"} {moeda(abs(saldo))}: '
            f'entradas de {moeda(entradas)}, gastos de {moeda(despesas)} e {moeda(investido)} investidos.')


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


def montar_schema() -> dict:
    # JSON Schema enviado ao Ollama: ele obriga a resposta a seguir exatamente este formato
    # (a ordem das propriedades é a ordem em que o modelo escreve, então a Ação vem primeiro e o Tipo antes da Categoria)
    categorias = list(CATEGORIAS_ENTRADA) + list(CATEGORIAS_DESPESA) + [CATEGORIA_PADRAO]
    texto_ou_nulo = {'anyOf': [{'type': 'string'}, {'type': 'null'}]}
    campos = {
        'Ação': {'type': 'integer', 'enum': [0, 1]},
        'Descrição': {'type': 'string'},
        'Valor': {'type': 'number'},
        'Data': {'type': 'string'},
        'Tipo': {'type': 'string', 'enum': TIPOS},
        'Dia de Pagamento': texto_ou_nulo,
        'Categoria': {'anyOf': [{'type': 'string', 'enum': categorias}, {'type': 'null'}]},
    }
    campos_consulta = {
        'Métrica': {'type': 'string', 'enum': METRICAS_CONSULTA},
        'Início': {'type': 'string'},
        'Fim': {'type': 'string'},
        'Categoria': {'anyOf': [{'type': 'string', 'enum': categorias}, {'type': 'null'}]},
        'Busca': texto_ou_nulo,
    }
    return {
        'type': 'object',
        'properties': {
            'Registros': {
                'type': 'array',
                'items': {'type': 'object', 'properties': campos, 'required': list(campos)},
            },
            'Consulta': {'anyOf': [
                {'type': 'object', 'properties': campos_consulta, 'required': list(campos_consulta)},
                {'type': 'null'},
            ]},
            'Resposta_Chat': texto_ou_nulo,
        },
        'required': ['Registros', 'Consulta', 'Resposta_Chat'],
    }


def montar_prompt(df: pd.DataFrame, mensagem: str) -> str:
    agora = datetime.now()
    hoje = agora.strftime('%d/%m/%Y')
    ontem = (agora - timedelta(days=1)).strftime('%d/%m/%Y')
    # Datas prontas para as consultas: o modelo erra menos copiando do que calculando
    inicio_mes = agora.replace(day=1)
    fim_mes_passado = inicio_mes - timedelta(days=1)
    inicio_mes_passado = fim_mes_passado.replace(day=1)
    inicio_semana = agora - timedelta(days=agora.weekday())
    dias_semana = ['segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado', 'domingo']

    # Só os registros de verdade e só as colunas de registro: as colunas de configuração são ruído para a IA
    registros = df[df['Descrição'].astype(str).str.strip() != ''][COLUNAS_REGISTRO].tail(50)
    planilha = registros.to_string(index=False) if not registros.empty else '(nenhum registro ainda)'

    def lista(categorias: dict) -> str:
        return '\n'.join(f'- "{nome}": {exemplos}.' for nome, exemplos in categorias.items())

    def registro(acao, descricao, valor, data, tipo, categoria, dia_pagamento=None):
        return {'Ação': acao, 'Descrição': descricao, 'Valor': valor, 'Data': data,
                'Tipo': tipo, 'Dia de Pagamento': dia_pagamento, 'Categoria': categoria}

    def consulta(metrica, inicio, fim, categoria=None, busca=None):
        return {'Métrica': metrica, 'Início': inicio, 'Fim': fim, 'Categoria': categoria, 'Busca': busca}

    # (mensagem, registros, consulta, resposta do chat)
    exemplos = [
        ('Recebi R$ 800,00 de um freela de design.',
         [registro(1, 'Freela de design', 800, hoje, 'Entrada', 'Renda Extra')], None, None),
        ('Gastei 120,50 no mercado ontem',
         [registro(1, 'Supermercado', 120.5, ontem, 'Saída', 'Alimentação')], None, None),
        ('Todo dia 10 pago R$ 89,90 da internet.',
         [registro(1, 'Internet', 89.9, hoje, 'Conta', 'Contas de Casa', proximo_dia(agora, 10))], None, None),
        ('Hoje paguei 32 de uber, almocei por 45,90 e abasteci o carro, deu 200 conto',
         [registro(1, 'Uber', 32, hoje, 'Saída', 'Transporte'),
          registro(1, 'Almoço', 45.9, hoje, 'Saída', 'Alimentação'),
          registro(1, 'Combustível', 200, hoje, 'Saída', 'Transporte')], None, None),
        ('Caiu meu salário de 3500 e coloquei 500 na reserva de emergência.',
         [registro(1, 'Salário', 3500, hoje, 'Entrada', 'Salário'),
          registro(1, 'Reserva de emergência', 500, hoje, 'Saída', 'Investimentos')], None, None),
        ('Apaga aquele gasto de 300 no restaurante.',
         [registro(0, 'Restaurante', 300, hoje, 'Saída', 'Alimentação')], None, None),
        ('Quanto gastei no mês passado?',
         [], consulta('Despesas', f'{inicio_mes_passado:%d/%m/%Y}', f'{fim_mes_passado:%d/%m/%Y}'), None),
        ('Quanto gastei de Uber este mês?',
         [], consulta('Despesas', f'{inicio_mes:%d/%m/%Y}', hoje, busca='Uber'), None),
        ('Quanto sobrou esta semana?',
         [], consulta('Saldo', f'{inicio_semana:%d/%m/%Y}', hoje), None),
        ('Vale a pena financiar um carro em 60 vezes?',
         [], None, 'Geralmente não, porque os juros de 60 meses fazem o carro sair bem mais caro. '
                   'O ideal é juntar uma boa entrada para financiar menos e por menos tempo.'),
    ]
    texto_exemplos = '\n\n'.join(
        f'Mensagem: {entrada}\nJSON: '
        f'{json.dumps({"Registros": saida, "Consulta": pergunta, "Resposta_Chat": chat}, ensure_ascii=False)}'
        for entrada, saida, pergunta, chat in exemplos
    )

    return f"""Você é um extrator de dados financeiros. Leia a mensagem do usuário e responda com UM único objeto JSON no formato {{"Registros": [...], "Consulta": ..., "Resposta_Chat": ...}}, sem nenhum texto antes ou depois e sem markdown.

Hoje é {dias_semana[agora.weekday()]}, {hoje}. Este mês começou em {inicio_mes:%d/%m/%Y}. O mês passado foi de {inicio_mes_passado:%d/%m/%Y} a {fim_mes_passado:%d/%m/%Y}. Esta semana começou na segunda-feira, {inicio_semana:%d/%m/%Y}.

=== REGISTROS ATUAIS DA PLANILHA ===
{planilha}
====================================

PASSO 1 - SEPARE OS REGISTROS:
- A mensagem pode citar um ou VÁRIOS gastos, contas ou recebimentos. Crie um item em "Registros" para CADA um deles, na ordem em que aparecem.
- Nunca junte itens diferentes num só e nunca some os valores deles (ex: "32 de uber e 45 no almoço" são dois itens).
- Não repita o mesmo item: se a mensagem fala da mesma coisa duas vezes, crie um item só.
- Se a mensagem não pede para registrar nem remover nada (perguntas, dúvidas, conselhos, conversa), "Registros" é uma lista vazia [].

PASSO 2 - DEFINA A "Ação" DE CADA ITEM (número inteiro):
- 1: registrar um gasto, uma conta ou um dinheiro recebido.
- 0: remover, apagar ou excluir um registro que está na planilha acima.
- Um pedido de remoção de um registro que NÃO está na planilha não entra em "Registros": explique em "Resposta_Chat" que não encontrou esse registro.

PASSO 3 - PREENCHA OS CAMPOS DE CADA ITEM:
- "Descrição": nome curto do gasto ou recebimento, com inicial maiúscula (ex: "Supermercado", "Conta de luz").
- "Valor": número positivo com ponto decimal, sem "R$" e sem sinal de menos (ex: 120.5). Gastos também são positivos.
- "Data": data do registro em "DD/MM/AAAA". Converta datas relativas ("ontem", "sexta passada") a partir de hoje. Sem data na mensagem, use {hoje}. Uma data citada uma vez vale para os itens que vêm junto dela.
- "Tipo": exatamente um destes:
  - "Entrada": dinheiro que o usuário recebeu.
  - "Saída": gasto pontual, já feito.
  - "Conta": conta recorrente ou com vencimento (aluguel, luz, internet, mensalidade, parcela, fatura).
- "Dia de Pagamento": só quando "Tipo" for "Conta": a próxima data de vencimento em "DD/MM/AAAA" (ex: "todo dia 10" = o próximo dia 10 a partir de hoje). Nos outros casos, null.
- "Categoria": veja o passo 4.

PASSO 4 - ESCOLHA A "Categoria" DE CADA ITEM (copie o nome exatamente como está, com acentos; nunca invente uma categoria nova):
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
Ache na planilha o registro citado e copie EXATAMENTE os valores dele: "Descrição", "Valor", "Tipo", "Categoria" e "Data" (a coluna "Data do registro"). Se a categoria estiver vazia na planilha, use null. Se mais de um registro combinar, escolha o mais recente. Para remover vários registros, crie um item de Ação 0 para cada um.

"Consulta" (perguntas sobre quanto foi gasto, recebido, investido ou quanto sobrou):
- NUNCA faça contas nem cite valores somados: preencha "Consulta" e o aplicativo calcula. Nesse caso "Registros" é [] e "Resposta_Chat" é null.
- "Métrica": "Despesas" (gastos e contas), "Entradas" (dinheiro recebido), "Investido" (aportes) ou "Saldo" (quanto sobrou ou faltou).
- "Início" e "Fim": datas em "DD/MM/AAAA", as duas incluídas no período. "este mês" = de {inicio_mes:%d/%m/%Y} a {hoje}; "mês passado" = de {inicio_mes_passado:%d/%m/%Y} a {fim_mes_passado:%d/%m/%Y}; um mês pelo nome (ex: "em março") = do dia 1 ao último dia desse mês, no ano mais recente que já começou; "este ano" = de 01/01/{agora.year} a {hoje}. Sem período na pergunta, use este mês.
- "Categoria": só quando a pergunta cita uma das categorias do passo 4 (ex: "com alimentação" = "Alimentação"); senão null.
- "Busca": quando a pergunta cita um gasto ou recebimento específico, a palavra para procurar na descrição (ex: "de Uber" = "Uber", "no iFood" = "iFood"); senão null.
- Em qualquer outra mensagem, "Consulta" é null.

"Resposta_Chat":
Resposta em texto simples, direta, com no máximo 3 frases, para dúvidas e conversa ou quando algum pedido de remoção não foi encontrado. Quando tudo foi registrado ou removido, ou quando a mensagem é uma consulta, use null.

EXEMPLOS (as datas consideram que hoje é {hoje}):

{texto_exemplos}

Agora responda apenas com o JSON para a mensagem abaixo.
Mensagem: {mensagem}
JSON:"""


def ler_registros(response) -> list:
    # Formato novo: {"Registros": [...]}. O formato antigo (um objeto só com "Ação") também é aceito,
    # caso o Ollama ignore o schema
    if not isinstance(response, dict):
        return []
    registros = response.get('Registros')
    if registros is None and 'Ação' in response:
        registros = [response]
    if not isinstance(registros, list):
        return []
    return [r for r in registros if isinstance(r, dict)]


def interpretar_mensagem(df, prompt):
    # Parte lenta: pergunta à IA o que a mensagem pede. Devolve (dicionário, None) ou (None, texto do erro).
    # Não mexe na planilha, então pode rodar enquanto a pessoa usa o resto do aplicativo
    # temperatura 0 deixa a extração previsível, e o schema obriga o Ollama a devolver o JSON no formato certo.
    # 2000 tokens cabem por volta de 15 registros numa mesma mensagem
    resposta = ia.Gemma4(montar_prompt(df, prompt), 2000, stream=False, formato=montar_schema(), temperatura=0)
    try:
        response = extrair_json(resposta)
    except ValueError:
        # Erro de conexão aparece como está; JSON cortado ou inválido vira um aviso amigável
        if str(resposta).startswith('Erro de conexão'):
            return None, str(resposta)
        return None, 'Não consegui entender a resposta da IA. Tente de novo ou envie menos itens por mensagem.'
    print(response)
    return response, None


def registerData(df, prompt):
    response, erro = interpretar_mensagem(df, prompt)
    if erro:
        return (df.fillna('').reset_index(drop=True), 2, erro)
    return aplicar_resposta(df, response)


def aplicar_resposta(df, response):
    # Parte rápida: aplica o que a IA entendeu na planilha. Deve receber a planilha mais atual, e não a que
    # existia quando a mensagem foi enviada; senão alterações feitas nesse meio-tempo seriam perdidas
    new_df = df.fillna('').reset_index(drop=True)

    hoje = datetime.now().strftime('%d/%m/%Y')
    adicionar, remover, avisos = [], [], []

    # 1ª passada: valida cada item; um item com problema não impede os outros
    for item in ler_registros(response):
        # A IA às vezes manda a ação como texto ("1")
        try:
            acao = int(item.get('Ação'))
        except (TypeError, ValueError):
            acao = None

        descricao = str(item.get('Descrição') or '').strip()
        valor = abs(numero(item.get('Valor')))
        tipo = escolher_opcao(item.get('Tipo'), TIPOS)
        nome = descricao or 'um dos itens'

        if acao not in (0, 1) or descricao == '' or valor == 0 or tipo == '':
            avisos.append(f'Não entendi {nome}. Diga o que foi, o valor e se foi um gasto ou um recebimento.')
        elif acao == 1:
            # A categoria precisa ser do mesmo grupo do tipo; fora disso vira "Outros"
            categorias = CATEGORIAS_ENTRADA if tipo == 'Entrada' else CATEGORIAS_DESPESA
            adicionar.append({
                'Descrição': descricao,
                'Valor': valor,
                'Data do registro': data_valida(item.get('Data'), hoje),
                'Tipo': tipo,
                'Categoria': escolher_opcao(item.get('Categoria'), categorias, CATEGORIA_PADRAO),
                # Dia de pagamento só faz sentido para contas
                'Dia de Pagamento': data_valida(item.get('Dia de Pagamento')) if tipo == 'Conta' else '',
                # Conta nova entra como não paga; é marcada como paga na tela de Lançamentos
                COLUNA_PAGO: 'Não' if tipo == 'Conta' else '',
            })
        else:
            remover.append((descricao, valor, tipo, data_valida(item.get('Data'))))

    # 2ª passada: remoções, sempre procuradas na planilha original
    removidos = []
    for descricao, valor, tipo, data_resp in remover:
        # Valor é comparado como número: "300" (IA) e 300.0 (planilha) são iguais
        filtro = (
            (new_df['Descrição'].apply(normalizar) == normalizar(descricao)) &
            (new_df['Valor'].apply(numero) == valor) &
            (new_df['Tipo'].apply(normalizar) == normalizar(tipo)) &
            # Uma linha já escolhida não pode ser removida de novo ("apaga os dois cafés" remove dois cafés)
            (~new_df.index.isin(removidos))
        )
        encontrados = new_df[filtro]

        # Se houver mais de um igual, a data desempata
        mesma_data = encontrados[encontrados['Data do registro'].astype(str).str.strip() == data_resp]
        if not mesma_data.empty:
            encontrados = mesma_data

        if encontrados.empty:
            avisos.append(f'Não encontrei {descricao} · {moeda(valor)} na planilha, então nada foi removido.')
        else:
            # Remove o mais recente que combina
            removidos.append(encontrados.index[-1])

    mensagens = [f"Removi: {new_df.loc[i, 'Descrição']} · {moeda(numero(new_df.loc[i, 'Valor']))} · "
                 f"{new_df.loc[i, 'Data do registro']}" for i in removidos]
    # Remove tudo de uma vez: removendo um por um, os índices das linhas mudariam no meio do caminho
    if removidos:
        new_df = remover_registro(new_df, removidos)

    # 3ª passada: adições, no fim da planilha
    for novo in adicionar:
        n = len(new_df)
        # Se a planilha só tem a linha de configuração vazia, reaproveita ela
        if n == 1 and str(new_df.loc[0, 'Descrição']).strip() == '':
            n = 0
        for coluna, valor in novo.items():
            new_df.loc[n, coluna] = valor
        # Mostra o que a IA entendeu, para o usuário conferir
        mensagens.append(f"Adicionei: {novo['Descrição']} · {moeda(novo['Valor'])} · {novo['Categoria']} · "
                         f"{novo['Data do registro']}")

    consulta = response.get('Consulta') if isinstance(response, dict) else None
    resposta_chat = str(response.get('Resposta_Chat') or '').strip()
    if isinstance(consulta, dict):
        # Com consulta, a resposta da IA é ignorada: ela poderia trazer valores que o modelo inventou
        avisos.append(responder_consulta(new_df, consulta))
    elif resposta_chat:
        avisos.append(resposta_chat)

    if not mensagens:
        # Nada mudou na planilha: é conversa ou nenhum item pôde ser aplicado
        texto = '\n'.join(avisos) or 'Não consegui entender. Diga o que foi, o valor e se foi um gasto ou um recebimento.'
        return (new_df, 2, texto)

    # 1 se adicionou algo (mensagem em verde na tela), 0 se só removeu
    acao = 1 if adicionar else 0
    return (new_df.fillna(''), acao, '\n'.join(mensagens + avisos))


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
