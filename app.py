import flet as ft
import pandas as pd
from ia import *
from main import registerData
import os


def Main(page: ft.Page):
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    dataxlsx = os.path.join(BASE_DIR, "data.xlsx")

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
    def salario(e):
        nonlocal df # Garante que estamos mexendo no df principal
        
        valor = campo_input.value
        if valor != "":
            df.loc[0, 'Salário'] = valor

            # Opcional: Salva no Excel para não perder
            df = df.fillna('')
            df.to_excel(dataxlsx, index=False)
            
            # Recarrega a tabela visual
            recarregar_tabela()
            
            # Limpa o campo e esconde a área de input
            campo_input.value = ""
            area_novo_dado.content = None 
            page.update()

    area_novo_dado = ft.Container() 

    # 2. Criamos os componentes de input
    campo_input = ft.TextField(label="Digite o seu salário", expand=True, color=ft.Colors.ON_SURFACE, on_submit=salario)
    


    botao_salvar = ft.IconButton(icon=ft.Icons.CHECK, icon_color=ft.Colors.GREEN, on_click=salario)
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

    def reserva_emergencia(e):
        nonlocal df
        salario = df.loc[0, 'Salário']
        if salario != '':
            df.loc[0, 'Reserva de Emergência'] = str(int(salario) * 6)
            df.fillna('')
            df.to_excel(dataxlsx, index=False)

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
            df.to_excel(dataxlsx, index=False)

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


    def hover_botao(e):
        e.control.scale = 1.1 if e.data == True else 1.0
        e.control.bgcolor = ft.Colors.BLUE_600 if e.data == True else ft.Colors.BLUE_700
        e.control.update()

    def botao(texto, icone, ao_clicar):
        return ft.Container(
            content=ft.Column(
                controls=[
                    ft.Icon(icone, color=ft.Colors.WHITE),
                    ft.Text(texto, size=fonte+5, weight="bold", color=ft.Colors.WHITE, expand=True, no_wrap=True),
                ],
                horizontal_alignment=ft.MainAxisAlignment.CENTER,

            ),
            width=270,
            height=110,
            bgcolor=ft.Colors.BLUE_700,
            border_radius=8,
            padding=10,
            ink=True,  # Mantém o efeito visual de clique do botão
            on_click=ao_clicar,
            on_hover=hover_botao,
            scale=1.0,  # Aceita o número direto nas versões recentes
            animate=ft.Animation(500, ft.AnimationCurve.EASE_OUT),
            animate_scale=ft.Animation(200, ft.AnimationCurve.EASE_OUT),
        )

    coluna_botoes = ft.Container(
    content=ft.Column(
        controls=[
            ft.Container(
                ft.Text("Menu Principal", size=20, weight="bold", color=ft.Colors.WHITE),
                alignment=ft.Alignment.CENTER,
                padding=ft.Padding.only(top=25)
            ),
            ft.Divider(color=ft.Colors.WHITE24), # Linha divisória sutil
            
            # Botão com visual mais moderno
            botao("Salário", ft.Icons.ATTACH_MONEY, mostrar_campo_insercao),
            botao("Analisar Planilha", ft.Icons.ANALYTICS, analise),
            botao("Reserva de Emergência", ft.Icons.SAVINGS, reserva_emergencia),
            botao("Viver de Renda", ft.Icons.SAVINGS, viver_renda),
        ],
        spacing=30
    ),
    width=300, # Largura fixa para o menu
    padding=20,
    bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.WHITE), # Fundo levemente translúcido
    border_radius=15,
    shadow=ft.BoxShadow(blur_radius=15, color=ft.Colors.BLACK26), # Sombra elegante
    )

    # ==========================================
    # 2. COLUNA CENTRAL (Planilha / DataTable)
    # ==========================================

    df = pd.read_excel(dataxlsx)
    df = pd.DataFrame(df).fillna('')  # Preenche valores nulos com string vazia

    tabela_financeira = ft.DataTable(columns=[], rows=[], column_spacing=105)
    
    planilha = ft.Container(
        content=ft.Column(
            controls=[
                ft.Text("Planilha Financeira", size=fonte+15, weight="bold", color=ft.Colors.WHITE),

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
        
        # 1. Defina exatamente quais colunas vão aparecer na tela e em qual ordem
        colunas_visiveis = [
            "Data do registro", 
            "Descrição", 
            "Valor", 
            "Tipo", 
            "Dia de Pagamento"
        ]
        
        # 2. Cria os cabeçalhos da tabela APENAS para as colunas visíveis
        for coluna in colunas_visiveis:
            # É bom checar se a coluna existe no df para evitar erros caso a planilha mude
            if coluna in df.columns:
                cabecalho = ft.Container(
                    content=ft.Text(coluna, weight="bold", size=fonte+5, color=ft.Colors.ON_SURFACE),
                    bgcolor=ft.Colors.BLACK_12,
                    border_radius=0
                )
                tabela.columns.append(ft.DataColumn(cabecalho))

        # 3. Preenche as linhas buscando os valores apenas dessas colunas
        for indice, linha in df.iterrows():
            linhas_celulas = []
            for coluna in colunas_visiveis:
                if coluna in df.columns:
                    valor = linha[coluna] # Pega o valor específico daquela coluna
                    if coluna == "Valor" and linha['Valor'] != "":
                        linhas_celulas.append(ft.DataCell(ft.Text("R$ "+f"{valor:.2f}", color=ft.Colors.WHITE)))
                    else:
                        linhas_celulas.append(ft.DataCell(ft.Text(str(valor), color=ft.Colors.WHITE)))

            
            tabela.rows.append(ft.DataRow(cells=linhas_celulas))


    recarregar_tabela()
    # ==========================================
    # 3. COLUNA DIREITA (Chat)
    # ==========================================
    # ListView permite rolar a tela se houver muitas mensagens
    lista_mensagens = ft.ListView(expand=True, spacing=10)
    lista_mensagens.controls.append(ft.Text("Sistema: Chat iniciado.", color=ft.Colors.GREEN, selectable=True))

    campo_mensagem = ft.TextField(hint_text="Digite...", expand=True, color=ft.Colors.ON_SURFACE)
    # Lembra da nossa regra da função com evento 'e'? Aqui está ela em ação!
    def enviar_mensagem(e):
        if campo_mensagem.value != "":
            # Adiciona o texto na lista de mensagens
            lista_mensagens.controls.append(ft.Text(f"Você: {campo_mensagem.value}", size=fonte+5, selectable=True, color=ft.Colors.ON_SURFACE))
            mensagem_usuario = campo_mensagem.value
            # Limpa o campo
            campo_mensagem.value = ""
            page.update()

            try:
                nonlocal df
                novo_df, acao = registerData(df, mensagem_usuario)
                df = novo_df.fillna('')
                df.to_excel(dataxlsx, index=False)
                recarregar_tabela()

                if acao == 0:
                    resposta = "Registro removido com sucesso."
                else:
                    resposta = "Registro adicionado com sucesso."

                lista_mensagens.controls.append(ft.Text(f"FinancIA: {resposta}", size=fonte+5, color=ft.Colors.GREEN, selectable=True))
            except Exception as exc:
                lista_mensagens.controls.append(ft.Text(f"FinancIA: Erro ao processar: {exc}", size=fonte+5, color=ft.Colors.RED, selectable=True))

            page.update()
            return campo_mensagem.value
            

    
    # O botão de enviar dispara a função acima
    botao_enviar = ft.IconButton(icon=ft.Icons.SEND, on_click=enviar_mensagem)
    # Apertar "Enter" no campo de texto também envia
    campo_mensagem.on_submit = enviar_mensagem

    coluna_chat = ft.Container(
        content=ft.Column(
            controls=[
                ft.Text("FinancIA", size=fonte+5, weight="bold", color=ft.Colors.ON_SURFACE),
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

if __name__ == "__main__":
    ft.run(Main)

# ADICIONAR GRÁFICO
