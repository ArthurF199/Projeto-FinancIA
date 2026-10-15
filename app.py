import flet as ft
import flet_charts as fch
import pandas as pd
from ia import *
from main import registerData, carregar_planilha, numero
import os
import threading
from time import sleep
from datetime import datetime


TIPOS_DESPESA = ["Saída", "Conta"]
NOMES_MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
CORES_PIZZA = [
    ft.Colors.RED_400,
    ft.Colors.ORANGE_400,
    ft.Colors.AMBER_400,
    ft.Colors.PURPLE_300,
    ft.Colors.CYAN_400,
    ft.Colors.BLUE_GREY_400,
]


def moeda(valor: float) -> str:
    # Formato brasileiro: R$ 1.234,56
    return "R$ " + f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def Main(page: ft.Page):
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    dataxlsx = os.path.join(BASE_DIR, "data.xlsx")

    df = carregar_planilha(dataxlsx)

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
    # FUNÇÕES DE DADOS
    # ==========================================
    def salvar_planilha(dados: pd.DataFrame):
        try:
            dados.to_excel(dataxlsx, index=False)
        except PermissionError:
            raise PermissionError("Feche o data.xlsx no Excel para que o FinancIA consiga salvar.")

    def registros():
        # Apenas as linhas que são transações de verdade, com colunas auxiliares já convertidas
        tabela = df[df["Descrição"].astype(str).str.strip() != ""].copy()
        tabela["_data"] = pd.to_datetime(tabela["Data do registro"], format="%d/%m/%Y", errors="coerce")
        tabela["_valor"] = tabela["Valor"].apply(numero)
        tabela["_tipo"] = tabela["Tipo"].astype(str).str.strip()
        return tabela

    def registros_do_mes(tabela, ano, mes):
        return tabela[(tabela["_data"].dt.year == ano) & (tabela["_data"].dt.month == mes)]

    def calcular_resumo():
        hoje = datetime.now()
        do_mes = registros_do_mes(registros(), hoje.year, hoje.month)

        salario = numero(df.loc[0, "Salário"])
        entradas = do_mes.loc[do_mes["_tipo"] == "Entrada", "_valor"].sum()
        despesas = do_mes.loc[do_mes["_tipo"].isin(TIPOS_DESPESA), "_valor"].sum()
        aporte = numero(df.loc[0, "Aporte Mensal"])
        receitas = salario + entradas

        return {
            "salario": salario,
            "reserva": numero(df.loc[0, "Reserva de Emergência"]),
            "renda": numero(df.loc[0, "Viver de Renda"]),
            "aporte": aporte,
            "receitas": receitas,
            "despesas": despesas,
            "saldo": receitas - despesas - aporte,
        }

    def calcular_metas():
        # Reserva = 6 salários, Viver de Renda = 120 salários, Aporte = 20% do salário
        salario = numero(df.loc[0, "Salário"])
        df.loc[0, "Reserva de Emergência"] = salario * 6
        df.loc[0, "Viver de Renda"] = salario * 120
        df.loc[0, "Aporte Mensal"] = round(salario * 0.2, 2)


    # ==========================================
    # LÓGICA DE INSERÇÃO MANUAL DE DADOS
    # ==========================================
    # 1. Container que vai guardar o campo de texto (começa vazio)
    def salario(e):
        valor = campo_input.value.strip()

        if valor != "":
            try:
                # Aceita "1500", "1500,00" e "1.500,00"
                if "," in valor:
                    valor = valor.replace(".", "").replace(",", ".")
                valor_convertido = float(valor.replace("R$", "").strip())

                df.loc[0, "Salário"] = valor_convertido
                calcular_metas()
                salvar_planilha(df)

                recarregar_tabela()
                atualizar_informacoes()

                # Limpa o campo e esconde a área de input
                campo_input.value = ""
                campo_input.error_text = None
                fechar_campo()

            except ValueError:
                # Se o usuário digitar letras (ex: "abc"), vai cair aqui
                campo_input.error_text = "Digite um número válido"
                page.update()
            except PermissionError as erro:
                campo_input.error_text = str(erro)
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
        # O campo fica acima da tabela, então garante que a planilha está visível
        mostrar_visao("planilha")
        area_novo_dado.content = ft.Row(
            controls=[
                campo_input,
                botao_salvar,
                botao_cancelar
            ]
        )
        page.update()
        campo_input.focus() # Já coloca o cursor piscando no campo

    # Configurações da página
    page.title = "Dashboard Flet"
    page.padding = 20
    # O theme_mode pode ser LIGHT ou DARK
    page.theme_mode = ft.ThemeMode.DARK

    # ==========================================
    # 1. COLUNA ESQUERDA (Botões)
    # ==========================================
    def relatorio(e):
        # Pede para a IA analisar a planilha e mostra a resposta no chat
        def gerar():
            mensagem_ia, txt_stream = mensagem("Gerando relatório...", user=False)
            lista_mensagens.controls.append(mensagem_ia)
            page.update()

            resumo = calcular_resumo()
            colunas = ["Data do registro", "Descrição", "Valor", "Tipo", "Dia de Pagamento"]
            resposta = Gemma4(f"""
Analise a planilha abaixo e me diga se eu estou fazendo um bom gerenciamento do meu dinheiro e me dê sugestões do que eu deveria fazer.
Responda de forma direta, sem criar tópicos ou textos muito grandes.

Salário: {moeda(resumo['salario'])}
Receitas do mês: {moeda(resumo['receitas'])}
Despesas do mês: {moeda(resumo['despesas'])}
Aporte mensal planejado: {moeda(resumo['aporte'])}
Saldo do mês: {moeda(resumo['saldo'])}

Registros:
{registros()[colunas].to_string(index=False)}
""", 1000)
            digitar(txt_stream, resposta)

        threading.Thread(target=gerar, daemon=True).start()


    def hover_botao(e):
        e.control.scale = 1.1 if e.data == True else 1.0
        e.control.bgcolor = ft.Colors.BLUE_600 if e.data == True else ft.Colors.BLUE_700
        e.control.update()


    def botao(texto, icone, ao_clicar):
        return ft.Container(
            gradient=ft.LinearGradient(
                colors=[ft.Colors.BLACK_38, ft.Colors.WHITE_12]
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
                bgcolor=ft.Colors.BLACK_38,
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
                        color=ft.Colors.with_opacity(0.75, ft.Colors.WHITE_24)),
                alignment=ft.Alignment.CENTER_LEFT,
                padding=ft.Padding.only(left=5),
                margin=ft.Margin.only(top=10),
                height=20
            ),

            # Botão com visual mais moderno

            botao("Registrar Salário", ft.Icons.ATTACH_MONEY, mostrar_campo_insercao),
            botao("Relatório", ft.Icons.ANALYTICS, relatorio),
            botao("Informações", ft.Icons.INFO, lambda e: alternar_visao(e)),


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

                    bgcolor=ft.Colors.BLACK_26,

                    width=375,
                    height=150,
                    padding=20,
                    border_radius=15,
                ),
                bgcolor=ft.Colors.BLACK_12,
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
    shadow=ft.BoxShadow(blur_radius=15, color=ft.Colors.BLACK_26), # Sombra elegante
    )

    # ==========================================
    # 2. COLUNA CENTRAL (Planilha / DataTable)
    # ==========================================

    tabela_financeira = ft.DataTable(columns=[], rows=[], column_spacing=105)

    visao_atual = "planilha"

    area_dinamica = ft.Container(expand=True)

    def mostrar_visao(nome):
        nonlocal visao_atual
        visao_atual = nome

        if nome == "planilha":
            recarregar_tabela()
            area_dinamica.content = planilha
            titulo_tela.value = "Planilha Financeira"
            coluna_botoes.content.controls[5].content.content.controls[1].value = "Informações"
        else:
            atualizar_informacoes()
            area_dinamica.content = tela_informacoes
            titulo_tela.value = "Informações"
            coluna_botoes.content.controls[5].content.content.controls[1].value = "Planilha"

        page.update()

    def alternar_visao(e):
        mostrar_visao("informacoes" if visao_atual == "planilha" else "planilha")

    titulo_tela = ft.Text("Planilha Financeira", size=fonte+15, weight="bold", color=ft.Colors.WHITE)

    tela_central = ft.Container(
        ft.Column([
            ft.Row([
                titulo_tela,
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

    # Textos que mudam: começam zerados e são preenchidos por atualizar_informacoes()
    texto_salario = ft.Text("R$ 0,00",
                            size=fonte+20,
                            color=ft.Colors.WHITE,
                            text_align=ft.TextAlign.CENTER,
                            margin=ft.Margin.only(top=10, bottom=8),
                            weight="bold")

    texto_reserva = ft.Text("R$ 0,00",
                            size=fonte+20,
                            color=ft.Colors.WHITE,
                            text_align=ft.TextAlign.CENTER,
                            margin=ft.Margin.only(top=10, bottom=8),
                            weight="bold")

    texto_renda = ft.Text("R$ 0,00",
                            size=fonte+20,
                            color=ft.Colors.WHITE,
                            text_align=ft.TextAlign.CENTER,
                            margin=ft.Margin.only(top=10, bottom=8),
                            weight="bold")

    texto_aporte = ft.Text("R$ 0,00",
                            size=fonte+20,
                            color=ft.Colors.WHITE,
                            text_align=ft.TextAlign.CENTER,
                            margin=ft.Margin.only(top=10, bottom=8),
                            weight="bold")

    texto_receitas = ft.Text("R$ 0,00",
                                        size=25,
                                        color=ft.Colors.WHITE,
                                        text_align=ft.TextAlign.CENTER,
                                        margin=ft.Margin.only(top=10, bottom=8),
                                        weight="bold",
                                        no_wrap=True)

    texto_despesas = ft.Text("R$ 0,00",
                                    size=25,
                                    color=ft.Colors.WHITE,
                                    text_align=ft.TextAlign.CENTER,
                                    margin=ft.Margin.only(top=10, bottom=8),
                                    weight="bold")

    texto_investimentos = ft.Text("R$ 0,00",
                                    size=25,
                                    color=ft.Colors.WHITE,
                                    text_align=ft.TextAlign.CENTER,
                                    margin=ft.Margin.only(top=10, bottom=8),
                                    weight="bold")

    texto_saldo = ft.Text("R$ 0,00",
                                    size=25,
                                    color=ft.Colors.WHITE,
                                    text_align=ft.TextAlign.CENTER,
                                    margin=ft.Margin.only(top=10, bottom=8),
                                    weight="bold")

    # ==========================================
    # GRÁFICOS
    # ==========================================
    RAIO_FATIA = 60
    RAIO_FATIA_DESTAQUE = 70

    def destacar_fatia(e: fch.PieChartEvent):
        # Aumenta a fatia que está sob o mouse
        for indice, fatia in enumerate(grafico_pizza.sections):
            fatia.radius = RAIO_FATIA_DESTAQUE if indice == e.section_index else RAIO_FATIA
        grafico_pizza.update()

    # Pizza (rosca): para onde foram as despesas do mês
    grafico_pizza = fch.PieChart(
        sections=[],
        sections_space=2,
        center_space_radius=40,
        expand=True,
        on_event=destacar_fatia,
    )
    legenda_pizza = ft.Column(spacing=8, alignment=ft.MainAxisAlignment.CENTER)
    conteudo_pizza = ft.Container(expand=True)

    # Barras: entradas x despesas dos últimos 6 meses
    grafico_barras = fch.BarChart(
        groups=[],
        group_spacing=20,
        interactive=True,
        expand=True,
        horizontal_grid_lines=fch.ChartGridLines(color=ft.Colors.WHITE_10, width=1, dash_pattern=[4, 4]),
        left_axis=fch.ChartAxis(label_size=60),
        bottom_axis=fch.ChartAxis(labels=[], label_size=30),
        tooltip=fch.BarChartTooltip(bgcolor=ft.Colors.BLACK_87, border_radius=8),
    )

    def atualizar_grafico_pizza():
        hoje = datetime.now()
        do_mes = registros_do_mes(registros(), hoje.year, hoje.month)
        despesas = do_mes[do_mes["_tipo"].isin(TIPOS_DESPESA)]
        gastos = despesas.groupby("Descrição")["_valor"].sum().sort_values(ascending=False)

        if gastos.sum() <= 0:
            conteudo_pizza.content = ft.Text(
                f"Nenhuma despesa registrada em {NOMES_MESES[hoje.month - 1]}/{hoje.year}.",
                color=ft.Colors.with_opacity(0.5, ft.Colors.WHITE),
            )
            return

        # Mais de 6 fatias fica ilegível: mantém as 5 maiores e junta o resto em "Outros"
        if len(gastos) > 6:
            outros = gastos.iloc[5:].sum()
            gastos = gastos.iloc[:5].copy()
            gastos["Outros"] = outros

        total = gastos.sum()
        grafico_pizza.sections = []
        legenda_pizza.controls = []
        for indice, (descricao, valor) in enumerate(gastos.items()):
            cor = CORES_PIZZA[indice % len(CORES_PIZZA)]
            grafico_pizza.sections.append(fch.PieChartSection(
                value=float(valor),
                title=f"{valor / total:.0%}",
                title_style=ft.TextStyle(size=13, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                color=cor,
                radius=RAIO_FATIA,
            ))
            legenda_pizza.controls.append(ft.Row([
                ft.Container(width=12, height=12, bgcolor=cor, border_radius=6),
                ft.Text(f"{descricao}  {moeda(valor)}", size=14, color=ft.Colors.WHITE, no_wrap=True),
            ]))

        conteudo_pizza.content = ft.Row([
            ft.Container(grafico_pizza, expand=3),
            ft.Container(legenda_pizza, expand=2),
        ])

    def atualizar_grafico_barras():
        tabela = registros()
        hoje = datetime.now()

        # Lista os últimos 6 meses (ano, mês), do mais antigo para o atual
        meses = []
        ano, mes = hoje.year, hoje.month
        for _ in range(6):
            meses.insert(0, (ano, mes))
            mes -= 1
            if mes == 0:
                ano, mes = ano - 1, 12

        grafico_barras.groups = []
        grafico_barras.bottom_axis.labels = []
        maior_valor = 0
        for indice, (ano, mes) in enumerate(meses):
            do_mes = registros_do_mes(tabela, ano, mes)
            entradas = do_mes.loc[do_mes["_tipo"] == "Entrada", "_valor"].sum()
            despesas = do_mes.loc[do_mes["_tipo"].isin(TIPOS_DESPESA), "_valor"].sum()
            maior_valor = max(maior_valor, entradas, despesas)

            grafico_barras.groups.append(fch.BarChartGroup(
                x=indice,
                spacing=4,
                rods=[
                    fch.BarChartRod(to_y=float(entradas), width=14, color=ft.Colors.GREEN_400,
                                    border_radius=4, tooltip=f"Entradas\n{moeda(entradas)}"),
                    fch.BarChartRod(to_y=float(despesas), width=14, color=ft.Colors.RED_400,
                                    border_radius=4, tooltip=f"Despesas\n{moeda(despesas)}"),
                ],
            ))
            grafico_barras.bottom_axis.labels.append(fch.ChartAxisLabel(
                value=indice,
                label=ft.Text(NOMES_MESES[mes - 1], size=13, color=ft.Colors.with_opacity(0.75, ft.Colors.WHITE)),
            ))

        # Folga de 20% no topo para a barra mais alta não encostar na borda
        grafico_barras.max_y = float(maior_valor) * 1.2 if maior_valor > 0 else 100

    def atualizar_informacoes():
        resumo = calcular_resumo()

        # Atualiza os valores dos textos na tela
        texto_salario.value = moeda(resumo["salario"])
        texto_reserva.value = moeda(resumo["reserva"])
        texto_renda.value = moeda(resumo["renda"])
        texto_aporte.value = moeda(resumo["aporte"])

        texto_receitas.value = moeda(resumo["receitas"])
        texto_despesas.value = moeda(resumo["despesas"])
        texto_investimentos.value = moeda(resumo["aporte"])
        texto_saldo.value = moeda(resumo["saldo"])
        texto_saldo.color = ft.Colors.WHITE if resumo["saldo"] >= 0 else ft.Colors.RED

        atualizar_grafico_pizza()
        atualizar_grafico_barras()


    def recarregar_tabela():
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
            # A linha 0 pode existir só para guardar o salário, sem nenhum registro
            if str(linha.get("Descrição", "")).strip() == "":
                continue

            linhas_celulas = []

            # Descobre o "Tipo" da linha ATUAL antes de preencher as colunas
            tipo_da_linha = str(linha.get('Tipo', '')).strip()

            if tipo_da_linha == "Entrada":
                cor_tipo = ft.Colors.GREEN
            elif tipo_da_linha == "Saída":
                cor_tipo = ft.Colors.RED
            elif tipo_da_linha == "Conta":
                cor_tipo = ft.Colors.ORANGE
            else:
                cor_tipo = ft.Colors.WHITE # Cor padrão para prevenir erros

            for coluna in visualizacao_atual:
                if coluna in df.columns:
                    valor = linha[coluna] # Pega o valor específico daquela coluna

                    if coluna == "Valor":
                        linhas_celulas.append(ft.DataCell(ft.Text(moeda(numero(valor)), color=cor_tipo, weight="W_600")))

                    elif coluna == "Tipo" and cor_tipo != ft.Colors.WHITE:
                        linhas_celulas.append(ft.DataCell(
                            ft.Container(
                                content=ft.Text(str(valor), color=cor_tipo, weight="bold"),
                                bgcolor=ft.Colors.with_opacity(0.1, cor_tipo),
                                border_radius=10,
                                padding=ft.Padding.only(left=10, right=10, top=5, bottom=5)
                            )
                        ))
                    else:
                        linhas_celulas.append(ft.DataCell(ft.Text(str(valor), color=ft.Colors.WHITE, weight="W_600")))

            tabela.rows.append(ft.DataRow(cells=linhas_celulas))

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
        bgcolor=ft.Colors.BLACK_26,
        border_radius=15
    )

    tela_informacoes = ft.Container(
        content=ft.Column(
            scroll=ft.ScrollMode.ALWAYS,
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
                        bgcolor=ft.Colors.BLACK_26,
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
                        bgcolor=ft.Colors.BLACK_26,
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
                        bgcolor=ft.Colors.BLACK_26,
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
                        bgcolor=ft.Colors.BLACK_26,
                        border_radius=15,
                        padding=20,
                        height=135,
                        expand=True
                    ),
                ]),

                ft.Container(
                    bgcolor=ft.Colors.BLACK_26,
                    border_radius=15,
                    padding=20,
                    content=ft.Column([
                        ft.Row([
                            ft.Icon(ft.Icons.DASHBOARD, size=30),
                            ft.Text("Resumo Geral", color=ft.Colors.WHITE, weight="bold", size=25),
                        ]),
                        # ft.Container(height=10),

                        ft.Container(
                            width=float("inf"),
                            content=ft.Row(
                                expand=True,
                                wrap=True,           # 1. Permite que os blocos pulem de linha
                                spacing=30,          # 2. Espaço horizontal entre os blocos
                                run_spacing=20,      # 3. Espaço vertical quando um bloco pular para a linha de baixo
                                alignment=ft.MainAxisAlignment.CENTER,
                                controls=[
                                    ft.Container(bgcolor=ft.Colors.WHITE_12, height=80, width=1),

                                    # Bloco de Receitas
                                    ft.Container(
                                        width=180, # 4. Largura fixa garante que o texto dentro não seja esmagado
                                        content=ft.Column([
                                            ft.Text("Receitas (mês)", color=ft.Colors.BLUE, weight="bold", size=19, no_wrap=True),
                                            texto_receitas
                                        ], spacing=-10)
                                    ),

                                    ft.Container(bgcolor=ft.Colors.WHITE_12, height=80, width=1),

                                    # Bloco de Despesas
                                    ft.Container(
                                        width=180,
                                        content=ft.Column([
                                            ft.Text("Despesas (mês)", color=ft.Colors.RED, weight="bold", size=19, no_wrap=True),
                                            texto_despesas
                                        ], spacing=-10)
                                    ),

                                    ft.Container(bgcolor=ft.Colors.WHITE_12, height=80, width=1),

                                    # Bloco de Investimentos
                                    ft.Container(
                                        width=180,
                                        content=ft.Column([
                                            ft.Text("Investimentos (mês)", color=ft.Colors.GREEN, weight="bold", size=19, no_wrap=True),
                                            texto_investimentos
                                        ], spacing=-10)
                                    ),

                                    ft.Container(bgcolor=ft.Colors.WHITE_12, height=80, width=1),

                                    # Bloco de Saldo
                                    ft.Container(
                                        width=180,
                                        content=ft.Column([
                                            ft.Text("Saldo (mês)", color=ft.Colors.BLUE, weight="bold", size=19, no_wrap=True),
                                            texto_saldo
                                        ], spacing=-10)
                                    ),

                                    ft.Container(bgcolor=ft.Colors.WHITE_12, height=80, width=1),
                                ]
                            )
                        )
                    ])
                ),

                ft.Row(
                    alignment=ft.MainAxisAlignment.CENTER,
                    controls=[
                        # Gráfico de pizza: despesas do mês por descrição
                        ft.Container(
                            bgcolor=ft.Colors.BLACK_26,
                            height=320,
                            expand=5,
                            border_radius=15,
                            padding=20,
                            content=ft.Column([
                                ft.Row([
                                    ft.Icon(ft.Icons.PIE_CHART, size=26),
                                    ft.Text("Despesas do mês", color=ft.Colors.WHITE, weight="bold", size=20),
                                ]),
                                conteudo_pizza,
                            ])
                        ),
                        # Gráfico de barras: entradas x despesas por mês
                        ft.Container(
                            bgcolor=ft.Colors.BLACK_26,
                            height=320,
                            expand=6,
                            border_radius=15,
                            padding=20,
                            content=ft.Column([
                                ft.Row([
                                    ft.Icon(ft.Icons.BAR_CHART, size=26),
                                    ft.Text("Entradas x Despesas", color=ft.Colors.WHITE, weight="bold", size=20),
                                    ft.Container(expand=True),
                                    ft.Container(width=10, height=10, bgcolor=ft.Colors.GREEN_400, border_radius=5),
                                    ft.Text("Entradas", size=13),
                                    ft.Container(width=10, height=10, bgcolor=ft.Colors.RED_400, border_radius=5),
                                    ft.Text("Despesas", size=13),
                                ]),
                                ft.Container(grafico_barras, expand=True),
                            ])
                        ),
                    ]
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
                            bgcolor=ft.Colors.BLACK_54 if user else ft.Colors.BLACK_26,
                            border_radius=20,
                            padding=10,
                            content=texto_control)
                    ],
                    wrap=True
                )
        return row, texto_control

    def digitar(txt_stream, texto: str, cor=ft.Colors.WHITE):
        # Efeito de digitação, letra por letra
        txt_stream.value = ""
        txt_stream.color = cor
        for char in texto:
            txt_stream.value += char
            page.update()
            sleep(.01)

    campo_mensagem = ft.TextField(hint_text="Digite...",
                                  expand=True,
                                  color=ft.Colors.WHITE,
                                  bgcolor=ft.Colors.BLACK_38,
                                  border_radius=15)
    # Lembra da nossa regra da função com evento 'e'? Aqui está ela em ação!
    def enviar_mensagem(e):
        texto_usuario = campo_mensagem.value.strip()
        if texto_usuario != "":
            # Adiciona o texto na lista de mensagens
            lista_mensagens.controls.append(mensagem(texto_usuario)[0])
            campo_mensagem.value = ""
            page.update()

            def processar_ia():
                nonlocal df
                sleep(.2) # Pequeno atraso estético

                mensagem_ia, txt_stream = mensagem("Processando...", user=False)
                lista_mensagens.controls.append(mensagem_ia)
                page.update()

                try:
                    novo_df, acao, mensagem_chat = registerData(df, texto_usuario)

                    if acao in (0, 1):
                        # Só troca o df depois de salvar, para a tela não mostrar algo que não está no arquivo
                        salvar_planilha(novo_df)
                        df = novo_df
                        recarregar_tabela()
                        atualizar_informacoes()

                    if acao == 0:
                        digitar(txt_stream, "Registro removido com sucesso.", ft.Colors.RED)
                    elif acao == 1:
                        digitar(txt_stream, "Registro adicionado com sucesso.", ft.Colors.GREEN)
                    else:
                        digitar(txt_stream, mensagem_chat)
                    page.update()
                except Exception as exc:
                    txt_stream.value = f"Erro ao processar: {exc}"
                    txt_stream.color = ft.Colors.RED
                    page.update()

            threading.Thread(target=processar_ia, daemon=True).start()


    # O botão de enviar dispara a função acima
    botao_enviar = ft.Container(
        content=ft.IconButton(icon=ft.Icons.SEND, on_click=enviar_mensagem, icon_size=35),
        bgcolor=ft.Colors.BLACK_26,
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

    # Prepara as duas telas antes de exibir; a planilha é a tela inicial
    recarregar_tabela()
    atualizar_informacoes()
    area_dinamica.content = planilha

    page.add(layout_principal)

if __name__ == "__main__":
    ft.run(Main, view=ft.AppView.WEB_BROWSER,
           host="0.0.0.0",
           port=8550)
