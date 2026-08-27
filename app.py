import flet as ft
import pandas as pd
from ia import *
from main import registerData
import os
import threading
from time import sleep


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
            title_large=ft.TextStyle(size=fonte+fonte*2),   # Títulos de páginas
        ),
        scrollbar_theme=ft.ScrollbarTheme(
            thickness=1
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
            try:
                # Substitui vírgula por ponto (para aceitar 1500,00) e converte para float
                valor_convertido = float(valor.replace(',', '.'))
                
                # Atualiza o DataFrame
                df.loc[0, 'Salário'] = valor_convertido

                # Opcional: Salva no Excel para não perder
                # Nota: Não precisa de df = df.fillna('') aqui se quiser manter a integridade, 
                # mas mantive como no seu original.
                # df.fillna('', inplace=True) 
                df.to_excel(dataxlsx, index=False)
                
                # Recarrega a tabela visual
                recarregar_tabela()
                
                # Limpa o campo e esconde a área de input
                campo_input.value = ""
                campo_input.error_text = None # Limpa possíveis erros anteriores
                fechar_campo() # Você já tem essa função abaixo, pode reaproveitá-la!
                
            except ValueError:
                # Se o usuário digitar letras (ex: "abc"), vai cair aqui
                campo_input.error_text = "Digite um número válido"
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
            df.loc[0, 'Reserva de Emergência'] = salario * 6
            df.fillna('')
            df.to_excel(dataxlsx, index=False)

            recarregar_tabela()
            atualizar_informacoes()

            page.update()
        else:
            mostrar_campo_insercao


    def viver_renda(e):
        nonlocal df
        salario = df.loc[0, 'Salário']
        if salario != '':
            df.loc[0, 'Viver de Renda'] = f"{salario * 120}"
            df.loc[0, 'Aporte Mensal'] = f"{salario * 0.2:.0f}"
            df.fillna('')
            df.to_excel(dataxlsx, index=False)

            recarregar_tabela()
            atualizar_informacoes()
            
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
    def relatorio(e):
        nonlocal df
        tela_central.content.controls[0].content=tela_relatorio

    tela_relatorio=ft.Container(
        ft.Text()
    )


    def hover_botao(e):
        e.control.scale = 1.1 if e.data == True else 1.0
        e.control.bgcolor = ft.Colors.BLUE_600 if e.data == True else ft.Colors.BLUE_700
        e.control.update()


    def botao(texto, icone, ao_clicar):
        return ft.Container(
            gradient=ft.LinearGradient(
                colors=[ft.Colors.BLACK38, ft.Colors.WHITE12]
            ),
            padding=3,
            border_radius=18,
            on_hover=hover_botao,
            animate=ft.Animation(500, ft.AnimationCurve.EASE_OUT),
            animate_scale=ft.Animation(200, ft.AnimationCurve.EASE_OUT),

            content=ft.Container(
                content=ft.Row(
                    controls=[
                        ft.Column(
                            ft.Icon(icone, color=ft.Colors.WHITE, size=50),
                            alignment=ft.CrossAxisAlignment.CENTER,
                        ),

                        ft.Text(texto, size=fonte+5+1, weight="bold", color=ft.Colors.WHITE, expand=True, no_wrap=True),
                    ],
                    vertical_alignment=ft.MainAxisAlignment.CENTER,
                ),
                width=375,
                height=110,
                bgcolor=ft.Colors.BLACK38,
                border_radius=16,
                padding=20,
                ink=True,  # Mantém o efeito visual de clique do botão
                on_click=ao_clicar,
                scale=1.0,  # Aceita o número direto nas versões recentes

            )
        )


    coluna_botoes = ft.Container(
    content=ft.Column(
        controls=[
            ft.Container(
                ft.Row([
                    ft.Container(
                        ft.Icon(ft.CupertinoIcons.CHART_BAR_ALT_FILL, size=50, color=ft.Colors.with_opacity(0.95, ft.Colors.WHITE)),
                        bgcolor=ft.Colors.BLACK_26,
                        border_radius=15,
                        padding=ft.Padding.only(left=7, right=5, top=5, bottom=9)
                    ),
                    ft.Column([
                        ft.Text("FinancIA", weight="bold", size=35, color=ft.Colors.WHITE),
                        ft.Text("Seu controle financeiro", size=17, color=ft.Colors.with_opacity(0.80, ft.Colors.WHITE)),
                    ], spacing=-5)
                ])
            ),

            ft.Divider(color=ft.Colors.WHITE_12), # Linha divisória sutil
            
            ft.Container(
                ft.Text("MENU PRINCIPAL",
                        size=17,
                        weight="W_300",
                        color=ft.Colors.with_opacity(0.75, ft.Colors.WHITE24)),
                alignment=ft.Alignment.CENTER_LEFT,
                padding=ft.Padding.only(left=5),
                margin=ft.Margin.only(top=10),
                height=20
            ),

            # Botão com visual mais moderno

            botao("Salário", ft.Icons.ATTACH_MONEY, mostrar_campo_insercao),
            botao("Relatório", ft.Icons.ANALYTICS, relatorio),
            botao("Reserva de Emergência", ft.Icons.SHIELD, reserva_emergencia),
            botao("Viver de Renda", ft.Icons.SAVINGS, viver_renda),
            

            ft.Container(
                ft.Container(
                    ft.Row([
                        ft.Container(
                            ft.Icon(ft.Icons.AUTO_AWESOME, size=60),
                            bgcolor=ft.Colors.with_opacity(0.20, ft.Colors.BLACK),
                            padding=5,
                            border_radius=15,
                        ),
                        ft.Column([
                            ft.Text("Organize hoje,", size=17, color=ft.Colors.WHITE, weight="bold"),
                            ft.Text("conquiste amanhã.", size=17, color=ft.Colors.WHITE)
                        ], alignment=ft.MainAxisAlignment.CENTER, spacing=0)
                    ]),

                    bgcolor=ft.Colors.BLACK26,

                    width=375,
                    height=150,
                    padding=20,
                    border_radius=15,
                ),
                bgcolor=ft.Colors.BLACK12,
                gradient=ft.LinearGradient(
                        colors=[ft.Colors.WHITE, ft.Colors.BLACK],
                        begin=ft.Alignment.CENTER_LEFT,
                        end=ft.Alignment(0.75, 0.5),
                        rotation=3.14 / 2.55
                    ),
                padding=2,
                width=375,
                height=150,
                border_radius=15,
                margin=ft.Margin.only(top=50)

            )
            
        ],
        scroll=ft.ScrollMode.ADAPTIVE
    ),
    width=375, # Largura fixa para o menu
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

    is_visao_dados = True

    area_dinamica = ft.Container(expand=True)

    def alternar_visao(e):
        nonlocal is_visao_dados
        # Não precisa do nonlocal tela_central se vamos apenas alterar uma propriedade dela

        is_visao_dados = not is_visao_dados

        if is_visao_dados:
            recarregar_tabela()
            area_dinamica.content = planilha
            tela_central.content.controls[0].controls[0].value = "Planilha Financeira"
        else:
            atualizar_informacoes()
            area_dinamica.content = tela_informacoes
            tela_central.content.controls[0].controls[0].value = "Informações"

        # Atualiza o container que segura as telas
        tela_central.update() 

    tela_central = ft.Container(
        ft.Column([
            ft.Row([
                ft.Text("", size=fonte+15, weight="bold", color=ft.Colors.WHITE),
                ft.Container(
                    content=ft.Icon(
                        ft.Icons.SYNC,
                        color=ft.Colors.with_opacity(0.5, ft.Colors.WHITE),
                    ),
                    # bgcolor=ft.Colors.BLACK_26,
                    height=30,
                    width=30,
                    border_radius=30,
                    margin=ft.Margin.only(top=5),
                    on_click=alternar_visao,
                ),
            ],
            margin=20),

            ft.Column([
                area_dinamica
            ], expand=True)
        ]),
        expand=5,
        bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.WHITE),
        border_radius=15,
        padding=10
    )


    visualizacao_atual = [
        "Data do registro",
        "Descrição",
        "Valor",
        "Tipo",
        "Dia de Pagamento"
    ]

    # Crie as variáveis para os textos que vão mudar
    texto_salario = ft.Text(f"R$ {df.loc[0, 'Salário']:.2f}",
                            size=fonte+20,
                            color=ft.Colors.WHITE,
                            text_align=ft.TextAlign.CENTER,
                            margin=ft.Margin.only(top=10, bottom=8),
                            weight="bold")
    
    texto_reserva = ft.Text(f"R$ {df.loc[0, 'Reserva de Emergência']:.2f}",
                            size=fonte+20,
                            color=ft.Colors.WHITE,
                            text_align=ft.TextAlign.CENTER,
                            margin=ft.Margin.only(top=10, bottom=8),
                            weight="bold")
    
    texto_renda = ft.Text(f"R$ {df.loc[0, 'Viver de Renda']:.2f}",
                            size=fonte+20,
                            color=ft.Colors.WHITE,
                            text_align=ft.TextAlign.CENTER,
                            margin=ft.Margin.only(top=10, bottom=8),
                            weight="bold")
    
    texto_aporte = ft.Text(f"R$ {df.loc[0, 'Aporte Mensal']:.2f}",
                            size=fonte+20,
                            color=ft.Colors.WHITE,
                            text_align=ft.TextAlign.CENTER,
                            margin=ft.Margin.only(top=10, bottom=8),
                            weight="bold")

    def atualizar_informacoes():
        # ATENÇÃO: Se o arquivo do Excel mudou no computador, 
        # você precisa ler ele de novo descomentando a linha abaixo:
        nonlocal df 
        df = pd.read_excel(dataxlsx).fillna('')
        
        # Atualiza os valores dos textos na tela
        texto_salario.value = f"R$ {df.loc[0, 'Salário']:.2f}"
        texto_reserva.value = f"R$ {df.loc[0, 'Reserva de Emergência']:.2f}"
        texto_renda.value = f"R$ {df.loc[0, 'Viver de Renda']:.2f}"
        texto_aporte.value = f"R$ {df.loc[0, 'Aporte Mensal']:.2f}"
        
        # Se a tela de informações estiver visível, atualiza ela
        # if tela_informacoes:
        #     tela_informacoes.update()

    def alternar_visao(e):
        nonlocal is_visao_dados
        # Não precisa do nonlocal tela_central se vamos apenas alterar uma propriedade dela

        is_visao_dados = not is_visao_dados

        if is_visao_dados:
            recarregar_tabela()
            tela_central.content = planilha
        else:
            atualizar_informacoes()
            tela_central.content = tela_informacoes

        # Atualiza o container que segura as telas
        tela_central.update() 


    def recarregar_tabela():
        entradas = 0
        despesas = 0
        tabela = tabela_financeira
        tabela.columns.clear()
        tabela.rows.clear()
        
        # 1. Cria os cabeçalhos da tabela APENAS para as colunas visíveis
        for coluna in visualizacao_atual:
            # É bom checar se a coluna existe no df para evitar erros caso a planilha mude
            if coluna in df.columns:
                cabecalho = ft.Container(
                    content=ft.Text(coluna, weight="bold", size=fonte+5, color=ft.Colors.ON_SURFACE),
                )
                tabela.columns.append(ft.DataColumn(ft.Container(cabecalho)))

        # 2. Preenche as linhas buscando os valores apenas dessas colunas
        for indice, linha in df.iterrows():
            linhas_celulas = []
            
            # Descobre o "Tipo" da linha ATUAL antes de preencher as colunas
            # Usamos o .get() para evitar erro caso a coluna 'Tipo' não exista ou esteja vazia
            tipo_da_linha = str(linha.get('Tipo', '')).strip()
            
            for coluna in visualizacao_atual:
                if coluna in df.columns:
                    valor = linha[coluna] # Pega o valor específico daquela coluna
                    
                    if isinstance(valor, (int, float)):
                        # Aplica a regra de cores baseada na variável tipo_da_linha
                        if tipo_da_linha == "Saída":
                            cor_texto = ft.Colors.GREEN
                            despesas+=valor
                        elif tipo_da_linha == "Entrada":
                            cor_texto = ft.Colors.BLUE
                            entradas+=valor
                        else:
                            cor_texto = ft.Colors.WHITE # Cor padrão para prevenir erros
                            
                        linhas_celulas.append(ft.DataCell(ft.Text("R$ "+f"{valor:.2f}", color=cor_texto, weight="W_600")))

                    else:
                        if coluna == "Tipo" and tipo_da_linha == "Saída":
                            linhas_celulas.append(ft.DataCell(
                                ft.Container(
                                    content=ft.Text(str(valor), color=ft.Colors.RED, weight="bold"),
                                    bgcolor=ft.Colors.with_opacity(0.1, ft.Colors.RED),
                                    border_radius=10,
                                    padding=ft.Padding.only(left=10, right=10, top=5, bottom=5)
                                )
                            ))
                        elif coluna == "Tipo" and tipo_da_linha == "Entrada":
                            linhas_celulas.append(ft.DataCell(
                                ft.Container(
                                    content=ft.Text(str(valor), color=ft.Colors.GREEN, weight="bold"),
                                    bgcolor=ft.Colors.with_opacity(0.1, ft.Colors.GREEN),
                                    border_radius=10,
                                    padding=ft.Padding.only(left=10, right=10, top=5, bottom=5)
                                )
                            ))
                        else:
                            linhas_celulas.append(ft.DataCell(ft.Text(str(valor), color=ft.Colors.WHITE, weight="W_600")))

 
            tabela.rows.append(ft.DataRow(cells=linhas_celulas))
            df.loc[0, 'Entradas'] = entradas
            df.loc[0, 'Despesas'] = despesas            
        page.update() # Atualiza a tela após inserir tudo

    planilha = ft.Container(
        content=ft.Column(
            controls=[
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
        padding=ft.Padding.only(left=10, right=10),
        bgcolor=ft.Colors.BLACK26,
        border_radius=15
    )

    recarregar_tabela()

    tela_informacoes = ft.Container(
        content=ft.Column(
            controls=[
                ft.Row([
                    ft.Container(
                        ft.Row([
                            ft.Container(
                                ft.Icon(ft.Icons.WALLET, size=50),
                                bgcolor=ft.Colors.with_opacity(0.3, ft.Colors.BLACK),
                                border_radius=20,
                                padding=15,
                                margin=ft.Margin.only(right=10)
                            ),

                            ft.Column([
                                ft.Text("Salário", color=ft.Colors.with_opacity(0.75, ft.Colors.WHITE), size=18),
                                texto_salario,
                                ft.Text("Receita Principal", color=ft.Colors.with_opacity(0.5, ft.Colors.WHITE))
                            ], spacing=-10, alignment=ft.CrossAxisAlignment.CENTER)
                        ]),
                        bgcolor=ft.Colors.BLACK26,
                        border_radius=15,
                        padding=20,
                        height=135,
                        expand=True
                    ),
                    ft.Container(
                        ft.Row([
                            ft.Container(
                                ft.Icon(ft.Icons.SHIELD, size=50),
                                bgcolor=ft.Colors.with_opacity(0.3, ft.Colors.BLACK),
                                border_radius=20,
                                padding=15,
                                margin=ft.Margin.only(right=10)
                            ),

                            ft.Column([
                                ft.Text("Reserva de Emergência", color=ft.Colors.with_opacity(0.75, ft.Colors.WHITE), size=18),
                                texto_reserva,
                                ft.Text("6 meses de segurança", color=ft.Colors.with_opacity(0.5, ft.Colors.WHITE))
                            ], spacing=-10, alignment=ft.CrossAxisAlignment.CENTER)
                        ]),
                        bgcolor=ft.Colors.BLACK26,
                        border_radius=15,
                        padding=20,
                        height=135,
                        expand=True
                    ),
                ]),
                ft.Row([
                    ft.Container(
                        ft.Row([
                            ft.Container(
                                ft.Icon(ft.Icons.SAVINGS, size=50),
                                bgcolor=ft.Colors.with_opacity(0.3, ft.Colors.BLACK),
                                border_radius=20,
                                padding=15,
                                margin=ft.Margin.only(right=10)
                            ),

                            ft.Column([
                                ft.Text("Viver de Renda", color=ft.Colors.with_opacity(0.75, ft.Colors.WHITE), size=18),
                                texto_renda,
                                ft.Text("Meta de Patrimônio", color=ft.Colors.with_opacity(0.5, ft.Colors.WHITE))
                            ], spacing=-10, alignment=ft.CrossAxisAlignment.CENTER)
                        ]),
                        bgcolor=ft.Colors.BLACK26,
                        border_radius=15,
                        padding=20,
                        height=135,
                        expand=True
                    ),
                    ft.Container(
                        ft.Row([
                            ft.Container(
                                ft.Icon(ft.Icons.SHOW_CHART, size=50),
                                bgcolor=ft.Colors.with_opacity(0.3, ft.Colors.BLACK),
                                border_radius=20,
                                padding=15,
                                margin=ft.Margin.only(right=10)
                            ),

                            ft.Column([
                                ft.Text("Aporte Mensal", color=ft.Colors.with_opacity(0.75, ft.Colors.WHITE), size=18),
                                texto_aporte,
                                ft.Text("Investimento mensal", color=ft.Colors.with_opacity(0.5, ft.Colors.WHITE))
                            ], spacing=-10, alignment=ft.CrossAxisAlignment.CENTER)
                        ]),
                        bgcolor=ft.Colors.BLACK26,
                        border_radius=15,
                        padding=20,
                        height=135,
                        expand=True
                    ),
                ]),

                ft.Container(
                    ft.Column([
                        ft.Row([
                            ft.Icon(ft.Icons.DASHBOARD, size=25),
                            ft.Text("Resumo Geral", color=ft.Colors.WHITE, weight="bold", size=25),
                        ]),
                        ft.Row([
                            ft.Column([
                                ft.Text("Receitas (mês)", color=ft.Colors.BLUE, weight="bold", size=20),
                                ft.Text(f"R$ {df.loc[0, 'Salário']+df.loc[0, 'Entradas']:.2f}",
                                    size=25,
                                    color=ft.Colors.WHITE,
                                    text_align=ft.TextAlign.CENTER,
                                    margin=ft.Margin.only(top=10, bottom=8),
                                    weight="bold"
                                )
                            ], spacing=-10),

                            ft.VerticalDivider(),

                            ft.Column([
                                ft.Text("Despesas (mês)", color=ft.Colors.RED, weight="bold", size=20),
                                ft.Text(f"R$ {df.loc[0, 'Despesas']:.2f}",
                                    size=25,
                                    color=ft.Colors.WHITE,
                                    text_align=ft.TextAlign.CENTER,
                                    margin=ft.Margin.only(top=10, bottom=8),
                                    weight="bold"
                                )
                            ], spacing=-10),

                            ft.VerticalDivider(),

                            ft.Column([
                                ft.Text("Investimentos (mês)", color=ft.Colors.GREEN, weight="bold", size=20),
                                ft.Text(f"R$ {df.loc[0, 'Aporte Mensal']:.2f}",
                                    size=25,
                                    color=ft.Colors.WHITE,
                                    text_align=ft.TextAlign.CENTER,
                                    margin=ft.Margin.only(top=10, bottom=8),
                                    weight="bold"
                                )
                            ], spacing=-10),

                            ft.VerticalDivider(width=1),

                            ft.Column([
                                ft.Text("Saldo (mês)", color=ft.Colors.BLUE, weight="bold", size=20),
                                ft.Text(f"R$ {df.loc[0, 'Salário']+df.loc[0, 'Entradas']-df.loc[0, 'Despesas']:.2f}",
                                    size=25,
                                    color=ft.Colors.WHITE if df.loc[0, 'Salário']+df.loc[0, 'Entradas']-df.loc[0, 'Despesas'] > 0 else ft.Colors.RED,
                                    text_align=ft.TextAlign.CENTER,
                                    margin=ft.Margin.only(top=10, bottom=8),
                                    weight="bold"
                                )
                            ], spacing=-10),

                        ],spacing=30,
                        margin=ft.Margin.only(left=32, right=32),
                        height=100,
                        scroll=ft.Scrollbar(
                            thumb_visibility=True,
                            thickness=4.0 if page.platform.is_mobile() and not page.web else None
                        ),
                        alignment=ft.MainAxisAlignment.CENTER)
                    ]),

                    bgcolor=ft.Colors.BLACK26,
                    border_radius=15,
                    padding=20,
                    height=150,
                )
            ]
        )
    )

    

    # ==========================================
    # 3. COLUNA DIREITA (Chat)
    # ==========================================
    # ListView permite rolar a tela se houver muitas mensagens
    lista_mensagens = ft.ListView(expand=True, spacing=10, auto_scroll=True)
    def mensagem(texto: str = "", user: bool = True):
        texto_control = ft.Text(texto,
                                size=fonte+5,
                                color=ft.Colors.WHITE,
                                selectable=True)
        
        row = ft.Row(
                    alignment=ft.MainAxisAlignment.END if user else ft.MainAxisAlignment.START,
                    controls=[
                        ft.Container(
                            width=page.width * 0.22 * 0.7, # 70% do tamanho do chat
                            bgcolor=ft.Colors.BLACK54 if user else ft.Colors.BLACK26,
                            border_radius=20,
                            padding=10,
                            content=texto_control)
                    ],
                    wrap=True
                )
        return row, texto_control
    campo_mensagem = ft.TextField(hint_text="Digite...",
                                  expand=True,
                                  color=ft.Colors.WHITE,
                                  bgcolor=ft.Colors.BLACK38,
                                  border_radius=15)
    # Lembra da nossa regra da função com evento 'e'? Aqui está ela em ação!
    def enviar_mensagem(e):
        if campo_mensagem.value != "":
            # Adiciona o texto na lista de mensagens
            mensagem_usuario = mensagem(campo_mensagem.value)[0]
            lista_mensagens.controls.append(mensagem_usuario)
            campo_mensagem.value = ""
            page.update()

            def processar_ia():
                sleep(.2) # Pequeno atraso estético

                mensagem_ia, txt_stream = mensagem("Processando...", user=False)
                lista_mensagens.controls.append(mensagem_ia)
                lista_mensagens.update()


                try:
                    nonlocal df
                    novo_df, acao, mensagem_chat = registerData(df, mensagem_usuario.controls[0].content.value)
                    df = novo_df.fillna('')

                    if acao == 0:
                        df.to_excel(dataxlsx, index=False)
                        recarregar_tabela()
                        txt_stream.value = ""
                        txt_stream.color = ft.Colors.RED
                        for char in "Registro removido com sucesso.":
                            txt_stream.value += char
                            mensagem_ia.update()
                            sleep(.01)

                    elif acao == 1:
                        df.to_excel(dataxlsx, index=False)
                        recarregar_tabela()
                        txt_stream.value = "" 
                        txt_stream.color = ft.Colors.GREEN
                        for char in "Registro adicionado com sucesso.":
                            txt_stream.value += char
                            mensagem_ia.update()
                            sleep(.01)
                    elif acao == 2:
                        txt_stream.value = ""
                        for char in mensagem_chat:
                            txt_stream.value += char
                            page.update()
                            sleep(.01)
                    page.update()
                except Exception as exc:
                    mensagem_ia.controls[0].content.value = f"Erro ao processar: {exc}"
                    mensagem_ia.controls[0].content.color = ft.Colors.RED
                    page.update()

            threading.Thread(target=processar_ia, daemon=True).start()

    
    # O botão de enviar dispara a função acima
    botao_enviar = ft.Container(
        content=ft.IconButton(icon=ft.Icons.SEND, on_click=enviar_mensagem, icon_size=35),
        bgcolor=ft.Colors.BLACK26,
        border_radius=180,
    )
    # Apertar "Enter" no campo de texto também envia
    campo_mensagem.on_submit = enviar_mensagem

    coluna_chat = ft.Container(
        content=ft.Column(
            controls=[
                ft.Row([
                    ft.Text("FinancIA", size=fonte+15, weight="bold", color=ft.Colors.ON_SURFACE),
                    ft.Icon(ft.Icons.AUTO_AWESOME, size=35)
                ], margin=20),
                lista_mensagens, # Ocupa o meio do chat
                ft.Row([campo_mensagem, botao_enviar]) # Fica na base do chat
            ],
        ),
        width=375,
        expand=2, # Ocupa 2 partes do espaço
        padding=10,
        bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.WHITE),
        border_radius=10,
    )

    # ==========================================
    # JUNTANDO TUDO NA PÁGINA
    # ==========================================
    # Colocamos os três blocos dentro de uma Row principal (Linha)
    layout_principal = ft.Row(
        controls=[
            coluna_botoes,
            # ft.VerticalDivider(width=1), # Uma linha vertical separando
            tela_central,
            # ft.VerticalDivider(width=1),
            coluna_chat
        ],
        expand=True # Faz a linha ocupar 100% da altura da tela
    )

    page.add(layout_principal)

if __name__ == "__main__":
    ft.run(Main, view=ft.AppView.WEB_BROWSER,
           host="0.0.0.0",
           port=8550)

# ADICIONAR GRÁFICO
