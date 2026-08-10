import flet as ft
import pandas as pd
from ia import *
from main import registerData


def Main(page: ft.Page):
    fullscreen = False
    def full_screen():
        nonlocal fullscreen
        if fullscreen == False:
            fullscreen = True
        else:
            fullscreen = False
        page.window.full_screen = fullscreen
    fonte = 15
    page.window.full_screen = False
    page.theme = ft.Theme(
        text_theme=ft.TextTheme(
            body_medium=ft.TextStyle(size=fonte),   # Tamanho padrão do texto comum
            body_large=ft.TextStyle(size=fonte+fonte*.3),    # Texto ligeiramente maior
            title_medium=ft.TextStyle(size=fonte+fonte*.6), # Títulos de tabelas/cards
            title_large=ft.TextStyle(size=fonte+fonte*2)   # Títulos de páginas
        )
    )


    # ==========================================
    # LÓGICA DE INSERÇÃO MANUAL DE DADOS
    # ==========================================
    # 1. Container que vai guardar o campo de texto (começa vazio)
    area_novo_dado = ft.Container() 

    # 2. Criamos os componentes de input
    campo_input = ft.TextField(label="Digite o seu salário", expand=True)
    
    def salvar_dado_manual(e):
        nonlocal df # Garante que estamos mexendo no df principal
        
        valor = campo_input.value
        if valor != "":
            df.loc[0, 'Salário'] = valor

            # Opcional: Salva no Excel para não perder
            df = df.fillna('')
            df.to_excel("data.xlsx", index=False)
            
            # Recarrega a tabela visual
            recarregar_tabela()
            
            # Limpa o campo e esconde a área de input
            campo_input.value = ""
            area_novo_dado.content = None 
            page.update()

    botao_salvar = ft.IconButton(icon=ft.Icons.CHECK, icon_color=ft.Colors.GREEN, on_click=salvar_dado_manual)
    botao_cancelar = ft.IconButton(icon=ft.Icons.CANCEL, icon_color=ft.Colors.RED, on_click=lambda e: fechar_campo())

    def fechar_campo():
        area_novo_dado.content = None
        page.update()

    # 3. Função que o botão principal chama para mostrar o campo na tela
    def mostrar_campo_insercao(e):
        # Coloca o TextField e os botões dentro da área que estava vazia
        area_novo_dado.content = ft.Row(
            controls=[
                campo_input,
                botao_salvar,
                botao_cancelar
            ]
        )
        page.update()
        campo_input.focus() # Já coloca o cursor piscando no campo

    def reserva_emergência(e):
        nonlocal df
        salario = df.loc[0, 'Salário']
        if salario != '':
            df.loc[0, 'Reserva de Emergência'] = str(int(salario) * 6)
            df.fillna('')
            df.to_excel('data.xlsx')

            recarregar_tabela()

            page.update()
        else:
            mostrar_campo_insercao


    def viver_renda(e):
        nonlocal df
        salario = df.loc[0, 'Salário']
        if salario != '':
            df.loc[0, 'Viver de Renda'] = f"{int(salario) * 120}"
            df.loc[0, 'Aporte Mensal'] = f"{int(salario) * 0.2:.0f}"
            df.fillna('')
            df.to_excel('data.xlsx')

            recarregar_tabela()

            page.update()
        else:
            mostrar_campo_insercao

    # Configurações da página
    page.title = "Dashboard Flet"
    page.padding = 20
    # O theme_mode pode ser LIGHT ou DARK
    page.theme_mode = ft.ThemeMode.DARK 

    # ==========================================
    # 1. COLUNA ESQUERDA (Botões)
    # ==========================================
    def analise(e):
        prompt = f"""Analise o dataframe: {df}.
        Esse dataframe apenas indica meus gastos, considere que eu não tenho controle sobre o dataframe, eu não consigo adicionar ou alterar as categorias, eu apenas dito os meus gastos.
        Me diga se estou fazendo um bom gerenciamento do meu dinheiro e me dê sugestões do que eu deveria fazer em relação ao meu financeiro.
        Lembre-se de responder de forma curta e resumida.
        """
        mensagem_ia = ft.Text("Sistema: Analisando dados...", color=ft.Colors.GREEN)
        lista_mensagens.controls.append(mensagem_ia)
        page.update()
        
        resposta = Gemma4(prompt)
        mensagem_ia.value = f"Sistema: {resposta}"
        page.update()
        # for token in resposta:
        #     mensagem_ia.value += token
        #     page.update()

    coluna_botoes = ft.Container(
        content=ft.Column(
            controls=[
                ft.Row(
                    ft.Button('', icon=ft.Icons.FULLSCREEN, on_click=lambda e: full_screen())),
                ft.Text("Menu", size=fonte+20, weight="bold"),
                ft.ElevatedButton(ft.Text("Salário", size=fonte+5), icon=ft.Icons.DASHBOARD, width=200, on_click=mostrar_campo_insercao),
                ft.ElevatedButton(ft.Text("Analisar Planilha", size=fonte+5), icon=ft.Icons.PIE_CHART, width=200, on_click=analise),
                ft.ElevatedButton(ft.Text("Reserva de Emergência", size=fonte+5), icon=ft.Icons.SETTINGS, width=200, on_click=reserva_emergência),
                ft.ElevatedButton(ft.Text("Viver de Renda", size=fonte+5), icon=ft.Icons.SETTINGS, width=200, on_click=viver_renda),
            ],
            spacing=15 # Espaço entre os botões
        ),
        expand=1, # Ocupa 1 parte do espaço
        padding=10,
        # bgcolor=ft.Colors.SURFACE_VARIANT,
        border_radius=10
    )

    # ==========================================
    # 2. COLUNA CENTRAL (Planilha / DataTable)
    # ==========================================

    df = pd.read_excel('data.xlsx')
    df = pd.DataFrame(df).fillna('')  # Preenche valores nulos com string vazia

    tabela_financeira = ft.DataTable(columns=[], rows=[])
    
    planilha = ft.Container(
        content=ft.Column(
            controls=[
                ft.Text("Planilha Financeira", size=fonte+5, weight="bold"),

                area_novo_dado,

                ft.Row(
                    controls=[tabela_financeira],
                    scroll=ft.ScrollMode.ADAPTIVE
                )
            ],
            scroll=ft.ScrollMode.ADAPTIVE,
            expand=True
        ),
        expand=5, # Ocupa 3 partes do espaço
        padding=10
    )
    
    def recarregar_tabela():
        tabela = tabela_financeira
        tabela.columns.clear()
        tabela.rows.clear()
        
        for coluna in df.columns:
            tabela.columns.append(ft.DataColumn(ft.Text(coluna, weight="bold")))

        for indice, linha in df.iterrows():
            linhas = []
            for valor in linha:
                linhas.append(ft.DataCell(ft.Text(str(valor))))
            tabela.rows.append(ft.DataRow(cells=linhas))


    recarregar_tabela()
    # ==========================================
    # 3. COLUNA DIREITA (Chat)
    # ==========================================
    # ListView permite rolar a tela se houver muitas mensagens
    lista_mensagens = ft.ListView(expand=True, spacing=10)
    lista_mensagens.controls.append(ft.Text("Sistema: Chat iniciado.", color=ft.Colors.GREEN))

    campo_mensagem = ft.TextField(hint_text="Digite...", expand=True)
    # Lembra da nossa regra da função com evento 'e'? Aqui está ela em ação!
    def enviar_mensagem(e):
        if campo_mensagem.value != "":
            # Adiciona o texto na lista de mensagens
            lista_mensagens.controls.append(ft.Text(f"Você: {campo_mensagem.value}", size=fonte+5))
            mensagem_usuario = campo_mensagem.value
            # Limpa o campo
            campo_mensagem.value = ""
            page.update()

            try:
                nonlocal df
                novo_df, acao = registerData(df, mensagem_usuario)
                df = novo_df.fillna('')
                df.to_excel("data.xlsx", index=False)
                recarregar_tabela()

                if acao == 0:
                    resposta = "Registro removido com sucesso."
                else:
                    resposta = "Registro adicionado com sucesso."

                lista_mensagens.controls.append(ft.Text(f"FinancIA: {resposta}", size=fonte+5, color=ft.Colors.GREEN))
            except Exception as exc:
                lista_mensagens.controls.append(ft.Text(f"FinancIA: Erro ao processar: {exc}", size=fonte+5, color=ft.Colors.RED))

            page.update()
            return campo_mensagem.value
            

    
    # O botão de enviar dispara a função acima
    botao_enviar = ft.IconButton(icon=ft.Icons.SEND, on_click=enviar_mensagem)
    # Apertar "Enter" no campo de texto também envia
    campo_mensagem.on_submit = enviar_mensagem

    coluna_chat = ft.Container(
        content=ft.Column(
            controls=[
                ft.Text("FinancIA", size=fonte+5, weight="bold"),
                lista_mensagens, # Ocupa o meio do chat
                ft.Row([campo_mensagem, botao_enviar]) # Fica na base do chat
            ]
        ),
        expand=2, # Ocupa 2 partes do espaço
        padding=10,
        # bgcolor=ft.Colors.SURFACE_VARIANT,
        border_radius=10
    )

    # ==========================================
    # JUNTANDO TUDO NA PÁGINA
    # ==========================================
    # Colocamos os três blocos dentro de uma Row principal (Linha)
    layout_principal = ft.Row(
        controls=[
            coluna_botoes,
            ft.VerticalDivider(width=1), # Uma linha vertical separando
            planilha,
            ft.VerticalDivider(width=1),
            coluna_chat
        ],
        expand=True # Faz a linha ocupar 100% da altura da tela
    )

    page.add(layout_principal)

ft.app(target=Main)