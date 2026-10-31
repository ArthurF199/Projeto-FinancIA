import flet as ft
import flet_charts as fch
import pandas as pd
from ia import *
from main import registerData, carregar_planilha, remover_registro, numero, moeda, COLUNAS_REGISTRO
import os
import re
import random
import threading
from time import sleep
from datetime import datetime, timedelta


TIPOS_DESPESA = ["Saída", "Conta"]
CATEGORIA_INVESTIMENTO = "Investimentos"
CATEGORIA_SALARIO = "Salário"
NOMES_MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
NOMES_MESES_COMPLETOS = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                         "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
# Paleta: azul-marinho profundo em camadas (fundo -> painel -> cartão -> elevado) + verde-esmeralda como destaque
AZUL_FUNDO = "#020817"
AZUL_PAINEL = "#061022"
AZUL_CARTAO = "#0A182E"
AZUL_ELEVADO = "#10223F"
AZUL_BORDA = "#172B4D"
VERDE = "#10B981"
VERDE_ESCURO = "#047857"
# Gradiente dos elementos de destaque (logo e botão de enviar), sempre com ícone branco
GRADIENTE_VERDE = [VERDE, VERDE_ESCURO]
# Cores de status: mantêm o significado (entrada/despesa/conta) em tons que combinam com o marinho
VERMELHO = "#EF4444"
AMBAR = "#F59E0B"
AZUL_CLARO = "#3B82F6"
ROXO = "#A78BFA"
TEXTO_SECUNDARIO = ft.Colors.with_opacity(0.72, ft.Colors.WHITE)
TEXTO_TERCIARIO = ft.Colors.with_opacity(0.55, ft.Colors.WHITE)

CORES_PIZZA = [
    VERDE,
    "#0D9488",
    AZUL_CLARO,
    "#6366F1",
    AMBAR,
    "#475569",
]

# Na planilha os tipos continuam "Saída"/"Conta" (a IA usa esses nomes); na tela usamos nomes mais claros
ROTULO_TIPO = {"Entrada": "Entrada", "Saída": "Gasto", "Conta": "Conta fixa"}
COR_TIPO = {"Entrada": VERDE, "Saída": VERMELHO, "Conta": ROXO}
ROTULO_COLUNA = {"Data do registro": "Data", "Dia de Pagamento": "Vencimento"}

# Períodos dos gráficos: a chave é usada no código, o texto aparece no seletor
PERIODOS_PIZZA = {"mes": "Este mês", "semana": "Esta semana", "dia": "Hoje"}
PERIODOS_BARRAS = {"mes": "Por mês", "semana": "Por semana", "dia": "Por dia"}

SUGESTOES_CHAT = [
    "Gastei 50 no mercado hoje",
    "Recebi 300 de um freela",
    "Pago 99,90 de internet todo dia 10",
    "Quanto gastei este mês?",
]

DICAS = [
    "Registre gastos escrevendo no chat, como \"gastei 30 no Uber\".",
    "Errou um lançamento? Peça no chat: \"apaga o Uber de ontem\".",
    "Contas que se repetem, como \"internet todo dia 10\", viram Conta fixa.",
    "Use \"Analisar meus gastos\" para receber sugestões da IA.",
]


def Main(page: ft.Page):
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    dataxlsx = os.path.join(BASE_DIR, "data.xlsx")

    df = carregar_planilha(dataxlsx)

    fonte = 15
    page.window.full_screen = False
    page.theme = ft.Theme(
        text_theme=ft.TextTheme(
            body_medium=ft.TextStyle(size=fonte),   # Tamanho padrão do texto comum
            body_large=ft.TextStyle(size=fonte+fonte*.3),    # Texto ligeiramente maior
            title_medium=ft.TextStyle(size=fonte+fonte*.6), # Títulos de tabelas/cards
            title_large=ft.TextStyle(size=fonte+fonte*2),   # Títulos de páginas
        ),
        # Barra de rolagem sempre visível: sem ela ninguém percebe que a tabela rola
        scrollbar_theme=ft.ScrollbarTheme(
            thickness=6,
            thumb_visibility=True,
            thumb_color=ft.Colors.with_opacity(0.25, ft.Colors.WHITE),
        ),
        color_scheme=ft.ColorScheme(
            primary=VERDE,
            on_primary=ft.Colors.WHITE,
            secondary=VERDE,
            on_secondary=ft.Colors.WHITE,
            surface=AZUL_PAINEL,
            on_surface=ft.Colors.WHITE,
            on_surface_variant=TEXTO_SECUNDARIO,
            surface_container=AZUL_CARTAO,
            surface_container_high=AZUL_ELEVADO,
            surface_container_highest=AZUL_ELEVADO,
            outline=AZUL_BORDA,
            outline_variant=AZUL_BORDA,
            error=VERMELHO,
        ),
        scaffold_bgcolor=AZUL_FUNDO,
        divider_color=AZUL_BORDA,
    )
    # Com theme_mode DARK o Flet usa o dark_theme, então os dois ficam iguais
    page.dark_theme = page.theme
    page.bgcolor = AZUL_FUNDO


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
        tabela["_categoria"] = tabela["Categoria"].astype(str).str.strip().replace("", "Outros")
        return tabela

    def eh_investimento(tabela):
        return tabela["_tipo"].isin(TIPOS_DESPESA) & (tabela["_categoria"] == CATEGORIA_INVESTIMENTO)

    def eh_despesa(tabela):
        # Dinheiro investido sai da conta, mas não é gasto: ele aparece separado
        return tabela["_tipo"].isin(TIPOS_DESPESA) & (tabela["_categoria"] != CATEGORIA_INVESTIMENTO)

    def totais(tabela, incluir_salario):
        # (entradas, despesas, investido) de um recorte da tabela
        entradas = tabela.loc[tabela["_tipo"] == "Entrada", "_valor"].sum()
        # Se o salário já foi lançado pelo chat, não soma de novo
        salario_lancado = ((tabela["_tipo"] == "Entrada") & (tabela["_categoria"] == CATEGORIA_SALARIO)).any()
        if incluir_salario and not salario_lancado:
            entradas += numero(df.loc[0, "Salário"])
        despesas = tabela.loc[eh_despesa(tabela), "_valor"].sum()
        investido = tabela.loc[eh_investimento(tabela), "_valor"].sum()
        return float(entradas), float(despesas), float(investido)

    def registros_do_mes(tabela, ano, mes):
        return tabela[(tabela["_data"].dt.year == ano) & (tabela["_data"].dt.month == mes)]

    def calcular_resumo():
        hoje = datetime.now()
        tabela = registros()
        receitas, despesas, investido = totais(registros_do_mes(tabela, hoje.year, hoje.month), True)

        return {
            "salario": numero(df.loc[0, "Salário"]),
            "reserva": numero(df.loc[0, "Reserva de Emergência"]),
            "renda": numero(df.loc[0, "Viver de Renda"]),
            "aporte": numero(df.loc[0, "Aporte Mensal"]),
            "receitas": receitas,
            "despesas": despesas,
            "investido": investido,
            "investido_total": float(tabela.loc[eh_investimento(tabela), "_valor"].sum()),
            "saldo": receitas - despesas - investido,
        }

    def calcular_metas(dados):
        # Reserva = 6 salários, Viver de Renda = 120 salários, Aporte = 20% do salário
        salario = numero(dados.loc[0, "Salário"])
        dados.loc[0, "Reserva de Emergência"] = salario * 6
        dados.loc[0, "Viver de Renda"] = salario * 120
        dados.loc[0, "Aporte Mensal"] = round(salario * 0.2, 2)

    def aplicar_alteracao(novo_df):
        # Salva primeiro: a tela só muda se o arquivo também mudou
        nonlocal df
        salvar_planilha(novo_df)
        df = novo_df
        recarregar_tabela()
        atualizar_informacoes()

    def avisar(texto, cor=VERDE, ao_desfazer=None):
        page.show_dialog(ft.SnackBar(
            content=ft.Text(texto, color=ft.Colors.WHITE, size=fonte+1),
            bgcolor=AZUL_ELEVADO,
            behavior=ft.SnackBarBehavior.FLOATING,
            width=520,
            show_close_icon=True,
            close_icon_color=TEXTO_SECUNDARIO,
            shape=ft.RoundedRectangleBorder(radius=12, side=ft.BorderSide(1, cor)),
            action=ft.SnackBarAction(label="Desfazer", text_color=VERDE, on_click=ao_desfazer) if ao_desfazer else None,
        ))

    def desfazer_para(df_anterior, df_depois):
        # Só desfaz se nada mudou desde então; senão voltaríamos alterações mais novas também
        def desfazer(e=None):
            if df is not df_depois:
                avisar("Não dá mais para desfazer: a planilha mudou depois disso.", AMBAR)
                return False
            try:
                aplicar_alteracao(df_anterior)
            except PermissionError as erro:
                avisar(str(erro), VERMELHO)
                return False
            page.update()
            return True
        return desfazer


    # ==========================================
    # SALÁRIO (configurado no card de Salário)
    # ==========================================
    area_salario = ft.Container()
    editando_salario = False

    def salvar_salario(e):
        nonlocal editando_salario
        valor = campo_salario.value.strip()
        if valor == "":
            return
        try:
            # Aceita "1500", "1500,00", "1.500,00" e "1.500"
            valor = valor.replace("R$", "").strip()
            if "," in valor or re.fullmatch(r"\d{1,3}(\.\d{3})+", valor):
                valor = valor.replace(".", "").replace(",", ".")
            valor_convertido = float(valor)
            if valor_convertido < 0:
                raise ValueError

            novo_df = df.copy()
            novo_df.loc[0, "Salário"] = valor_convertido
            calcular_metas(novo_df)
            aplicar_alteracao(novo_df)

            campo_salario.value = ""
            campo_salario.error_text = None
            editando_salario = False
            atualizar_area_salario()
            avisar(f"Salário atualizado para {moeda(valor_convertido)}. Suas metas foram recalculadas.")
        except ValueError:
            # Se o usuário digitar letras (ex: "abc"), vai cair aqui
            campo_salario.error_text = "Digite um número válido, ex.: 3500 ou 3.500,00"
        except PermissionError as erro:
            campo_salario.error_text = str(erro)
        page.update()

    campo_salario = ft.TextField(label="Seu salário mensal (R$)", hint_text="Ex.: 3.500,00", expand=True,
                                 color=ft.Colors.WHITE, on_submit=salvar_salario,
                                 bgcolor=AZUL_ELEVADO, filled=True, border_color=AZUL_BORDA,
                                 focused_border_color=VERDE, cursor_color=VERDE,
                                 label_style=ft.TextStyle(color=TEXTO_SECUNDARIO),
                                 hint_style=ft.TextStyle(color=TEXTO_TERCIARIO), border_radius=12)

    def editar_salario(e=None):
        nonlocal editando_salario
        editando_salario = True
        salario_atual = numero(df.loc[0, "Salário"])
        campo_salario.value = f"{salario_atual:.2f}".replace(".", ",") if salario_atual > 0 else ""
        campo_salario.error_text = None
        atualizar_area_salario()
        page.update()
        campo_salario.focus()

    def cancelar_salario(e=None):
        nonlocal editando_salario
        editando_salario = False
        atualizar_area_salario()
        page.update()

    def formulario_salario():
        return ft.Container(
            ft.Column([
                ft.Text("Qual é o seu salário mensal?", color=ft.Colors.WHITE, weight="bold", size=fonte+2),
                ft.Row([
                    campo_salario,
                    ft.FilledButton("Salvar", icon=ft.Icons.CHECK, on_click=salvar_salario,
                                    style=ft.ButtonStyle(bgcolor=VERDE, color=ft.Colors.WHITE)),
                    ft.TextButton("Cancelar", on_click=cancelar_salario,
                                  style=ft.ButtonStyle(color=TEXTO_SECUNDARIO)),
                ], vertical_alignment=ft.CrossAxisAlignment.START),
            ], spacing=10),
            bgcolor=AZUL_CARTAO,
            border=ft.Border.all(1, VERDE),
            border_radius=15,
            padding=20,
            expand=True,
        )

    def atualizar_area_salario():
        # O formulário abre onde a pessoa clicou: no aviso do topo (primeiro uso) ou no próprio card de Salário
        sem_salario = numero(df.loc[0, "Salário"]) <= 0
        espaco_card_salario.content = formulario_salario() if editando_salario and not sem_salario else card_salario

        if editando_salario and sem_salario:
            area_salario.content = formulario_salario()
        elif sem_salario:
            # Primeiro uso: sem salário as metas e o resumo ficam zerados, então pedimos logo de cara
            area_salario.content = ft.Container(
                ft.Row([
                    ft.Icon(ft.Icons.WAVING_HAND, color=AMBAR, size=30),
                    ft.Column([
                        ft.Text("Comece configurando seu salário", weight="bold", size=fonte+3, color=ft.Colors.WHITE),
                        ft.Text("Com ele calculamos suas metas e quanto sobra no mês.", color=TEXTO_SECUNDARIO),
                    ], spacing=2, expand=True),
                    ft.FilledButton("Configurar salário", icon=ft.Icons.EDIT, on_click=editar_salario,
                                    style=ft.ButtonStyle(bgcolor=VERDE, color=ft.Colors.WHITE)),
                ], vertical_alignment=ft.CrossAxisAlignment.CENTER),
                bgcolor=ft.Colors.with_opacity(0.08, AMBAR),
                border=ft.Border.all(1, ft.Colors.with_opacity(0.5, AMBAR)),
                border_radius=15,
                padding=20,
            )
        else:
            area_salario.content = None

    # Configurações da página
    page.title = "FinancIA"
    page.padding = 20
    # O theme_mode pode ser LIGHT ou DARK
    page.theme_mode = ft.ThemeMode.DARK

    # ==========================================
    # 1. COLUNA ESQUERDA (Navegação)
    # ==========================================
    botoes_nav = {}

    def estilo_nav(botao, selecionado, em_cima=False):
        botao.bgcolor = ft.Colors.with_opacity(0.14, VERDE) if selecionado else (AZUL_ELEVADO if em_cima else AZUL_CARTAO)
        botao.border = ft.Border.all(1.5, VERDE if selecionado else AZUL_BORDA)
        texto = botao.content.controls[1]
        texto.color = ft.Colors.WHITE if selecionado or em_cima else TEXTO_SECUNDARIO

    def hover_nav(e):
        nome = e.control.data
        estilo_nav(e.control, nome == visao_atual, e.data == True)
        e.control.update()

    def botao_nav(texto, icone, nome):
        botao = ft.Container(
            content=ft.Row([
                ft.Icon(icone, color=VERDE, size=26),
                ft.Text(texto, size=fonte+3, weight="bold", no_wrap=True),
            ], spacing=14),
            data=nome,
            height=58,
            border_radius=14,
            padding=ft.Padding.symmetric(horizontal=18),
            ink=True,
            on_click=lambda e: mostrar_visao(nome),
            on_hover=hover_nav,
            animate=ft.Animation(200, ft.AnimationCurve.EASE_OUT),
        )
        botoes_nav[nome] = botao
        return botao

    def titulo_secao(texto):
        return ft.Container(
            ft.Text(texto, size=fonte, weight="W_500", color=TEXTO_TERCIARIO),
            padding=ft.Padding.only(left=5),
            margin=ft.Margin.only(top=10),
        )

    analisando = False

    def analisar_gastos(e):
        # Pede para a IA analisar a planilha e mostra a resposta no chat
        nonlocal analisando
        if analisando:
            return
        analisando = True
        estado_analise(True)
        esconder_boas_vindas()

        def gerar():
            nonlocal analisando
            try:
                mensagem_ia, txt_stream, _ = mensagem("Analisando seus gastos...", user=False)
                lista_mensagens.controls.append(mensagem_ia)
                page.update()

                resumo = calcular_resumo()
                resposta = Gemma4(f"""
Analise a planilha abaixo e me diga se eu estou fazendo um bom gerenciamento do meu dinheiro e me dê sugestões do que eu deveria fazer.
Responda de forma direta, sem criar tópicos ou textos muito grandes.

Salário: {moeda(resumo['salario'])}
Receitas do mês: {moeda(resumo['receitas'])}
Despesas do mês: {moeda(resumo['despesas'])}
Investido no mês: {moeda(resumo['investido'])} (meta: {moeda(resumo['aporte'])})
Saldo do mês: {moeda(resumo['saldo'])}

Registros:
{registros()[COLUNAS_REGISTRO].to_string(index=False)}
""", 1000)
                digitar(txt_stream, resposta)
            finally:
                analisando = False
                estado_analise(False)
                page.update()

        threading.Thread(target=gerar, daemon=True).start()

    texto_analise = ft.Text("Analisar meus gastos", size=fonte+2, weight="bold", color=ft.Colors.WHITE, no_wrap=True)
    icone_analise = ft.Container(ft.Icon(ft.Icons.AUTO_AWESOME, color=ft.Colors.WHITE, size=24), width=24, height=24)

    def estado_analise(carregando):
        # Enquanto a IA trabalha, o botão mostra que está ocupado e o chat ganha destaque
        texto_analise.value = "Analisando..." if carregando else "Analisar meus gastos"
        icone_analise.content = (ft.ProgressRing(width=20, height=20, stroke_width=2.5, color=ft.Colors.WHITE)
                                 if carregando else ft.Icon(ft.Icons.AUTO_AWESOME, color=ft.Colors.WHITE, size=24))
        botao_analise.opacity = 0.7 if carregando else 1
        coluna_chat.border = ft.Border.all(1.5 if carregando else 1, VERDE if carregando else AZUL_BORDA)

    botao_analise = ft.Container(
        content=ft.Row([icone_analise, texto_analise], spacing=12),
        gradient=ft.LinearGradient(colors=GRADIENTE_VERDE, begin=ft.Alignment.TOP_LEFT, end=ft.Alignment.BOTTOM_RIGHT),
        height=58,
        border_radius=14,
        padding=ft.Padding.symmetric(horizontal=18),
        ink=True,
        on_click=analisar_gastos,
        tooltip="A IA lê seus lançamentos e responde no chat ao lado",
        animate_opacity=200,
    )

    texto_dica = ft.Text(random.choice(DICAS), size=fonte, color=TEXTO_SECUNDARIO)

    coluna_botoes = ft.Container(
    content=ft.Column(
        controls=[
          # Parte de cima: rola em telas pequenas e ocupa o espaço livre
          ft.Column(
            controls=[
            ft.Container(
                ft.Row([
                    ft.Container(
                        ft.Icon(ft.CupertinoIcons.CHART_BAR_ALT_FILL, size=50, color=ft.Colors.WHITE),
                        gradient=ft.LinearGradient(colors=GRADIENTE_VERDE, begin=ft.Alignment.TOP_LEFT, end=ft.Alignment.BOTTOM_RIGHT),
                        border_radius=15,
                        padding=ft.Padding.only(left=7, right=5, top=5, bottom=9)
                    ),
                    ft.Column([
                        ft.Text("FinancIA", weight="bold", size=35, color=ft.Colors.WHITE),
                        ft.Text("Seu controle financeiro", size=17, color=TEXTO_SECUNDARIO),
                    ], spacing=-5)
                ])
            ),

            ft.Divider(color=AZUL_BORDA), # Linha divisória sutil

            titulo_secao("MENU"),
            botao_nav("Visão geral", ft.Icons.DASHBOARD, "informacoes"),
            botao_nav("Lançamentos", ft.Icons.RECEIPT_LONG, "planilha"),

            titulo_secao("INTELIGÊNCIA ARTIFICIAL"),
            botao_analise,
            ],
            expand=True,
            spacing=10,
            scroll=ft.ScrollMode.ADAPTIVE
          ),

            # Parte de baixo: dica discreta (não é clicável, então não parece botão)
            ft.Container(
                ft.Row([
                    ft.Icon(ft.Icons.LIGHTBULB_OUTLINE, size=24, color=AMBAR),
                    ft.Column([
                        ft.Text("Dica", size=fonte, weight="bold", color=ft.Colors.WHITE),
                        texto_dica,
                    ], spacing=2, expand=True),
                ], vertical_alignment=ft.CrossAxisAlignment.START),
                padding=15,
                border_radius=12,
                bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.WHITE),
            )
        ],
    ),
    width=375, # Largura fixa para o menu
    padding=20,
    bgcolor=AZUL_PAINEL,
    border=ft.Border.all(1, AZUL_BORDA),
    border_radius=15,
    shadow=ft.BoxShadow(blur_radius=15, color=ft.Colors.with_opacity(0.4, ft.Colors.BLACK)), # Sombra elegante
    )

    # ==========================================
    # 2. COLUNA CENTRAL (Visão geral / Lançamentos)
    # ==========================================

    tabela_financeira = ft.DataTable(columns=[], rows=[], column_spacing=48,
                                     horizontal_lines=ft.BorderSide(1, AZUL_BORDA),
                                     heading_row_color=AZUL_ELEVADO,
                                     data_row_color={ft.ControlState.HOVERED: ft.Colors.with_opacity(0.06, VERDE)})

    visao_atual = "informacoes"

    area_dinamica = ft.Container(expand=True)
    titulo_tela = ft.Text("", size=fonte+15, weight="bold", color=ft.Colors.WHITE)
    subtitulo_tela = ft.Text("", size=fonte+1, color=TEXTO_SECUNDARIO)

    def mostrar_visao(nome, atualizar_pagina=True):
        nonlocal visao_atual
        visao_atual = nome

        if nome == "planilha":
            recarregar_tabela()
            area_dinamica.content = planilha
            titulo_tela.value = "Lançamentos"
            subtitulo_tela.value = "Tudo o que você registrou. Para adicionar, escreva no chat ao lado."
        else:
            atualizar_informacoes()
            area_dinamica.content = tela_informacoes
            titulo_tela.value = "Visão geral"
            subtitulo_tela.value = "Como está o seu dinheiro neste mês."

        for nome_botao, botao in botoes_nav.items():
            estilo_nav(botao, nome_botao == nome)

        if atualizar_pagina:
            page.update()

    tela_central = ft.Container(
        ft.Column([
            ft.Container(
                ft.Column([titulo_tela, subtitulo_tela], spacing=0),
                padding=ft.Padding.only(left=20, top=10, bottom=10),
            ),

            ft.Column([
                area_dinamica
            ], expand=True)
        ]),
        expand=5,
        bgcolor=AZUL_PAINEL,
        border=ft.Border.all(1, AZUL_BORDA),
        border_radius=15,
        padding=10
    )


    visualizacao_atual = COLUNAS_REGISTRO

    def texto_valor(tamanho=25):
        # Textos que mudam: começam zerados e são preenchidos por atualizar_informacoes()
        return ft.Text("R$ 0,00", size=tamanho, color=ft.Colors.WHITE, weight="bold", no_wrap=True)

    texto_salario = texto_valor(fonte+17)
    texto_reserva = texto_valor(fonte+17)
    texto_renda = texto_valor(fonte+17)
    texto_aporte = texto_valor(fonte+17)

    legenda_reserva = ft.Text("", color=TEXTO_TERCIARIO, size=fonte-1)
    legenda_renda = ft.Text("", color=TEXTO_TERCIARIO, size=fonte-1)
    legenda_aporte = ft.Text("", color=TEXTO_TERCIARIO, size=fonte-1)

    def barra_progresso():
        return ft.ProgressBar(value=0, color=VERDE, bgcolor=AZUL_BORDA, bar_height=6, border_radius=3)

    progresso_reserva = barra_progresso()
    progresso_renda = barra_progresso()
    progresso_aporte = barra_progresso()

    texto_receitas = texto_valor()
    texto_despesas = texto_valor()
    texto_investimentos = texto_valor()
    texto_saldo = texto_valor(fonte+23)
    rotulo_saldo = ft.Text("Sobrou no mês", color=VERDE, weight="bold", size=fonte+4, no_wrap=True)
    legenda_investimentos = ft.Text("", color=TEXTO_TERCIARIO, size=fonte-1, no_wrap=True)
    titulo_resumo = ft.Text("", color=ft.Colors.WHITE, weight="bold", size=25)

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

    # Pizza (rosca): para onde foram as despesas, por categoria
    grafico_pizza = fch.PieChart(
        sections=[],
        sections_space=2,
        center_space_radius=40,
        expand=True,
        on_event=destacar_fatia,
    )
    legenda_pizza = ft.Column(spacing=4, alignment=ft.MainAxisAlignment.CENTER)
    conteudo_pizza = ft.Container(expand=True)

    # Barras: entradas x despesas dos últimos períodos
    grafico_barras = fch.BarChart(
        groups=[],
        group_spacing=20,
        interactive=True,
        expand=True,
        horizontal_grid_lines=fch.ChartGridLines(color=AZUL_BORDA, width=1, dash_pattern=[4, 4]),
        left_axis=fch.ChartAxis(labels=[], label_size=70),
        bottom_axis=fch.ChartAxis(labels=[], label_size=40),
        tooltip=fch.BarChartTooltip(bgcolor=AZUL_FUNDO, border_radius=8),
    )

    def criar_seletor_periodo(opcoes, ao_mudar):
        return ft.Dropdown(
            value="mes",
            options=[ft.DropdownOption(key=chave, text=texto) for chave, texto in opcoes.items()],
            width=150,
            dense=True,
            text_size=14,
            color=ft.Colors.WHITE,
            border_color=AZUL_BORDA,
            focused_border_color=VERDE,
            filled=True,
            fill_color=AZUL_ELEVADO,
            border_radius=10,
            on_select=ao_mudar,
        )

    def inicio_do_dia(data):
        return pd.Timestamp(data.year, data.month, data.day)

    def intervalo_atual(modo):
        # Intervalo [início, fim) do período que contém hoje
        hoje = inicio_do_dia(datetime.now())
        if modo == "dia":
            return hoje, hoje + timedelta(days=1)
        if modo == "semana":
            inicio = hoje - timedelta(days=hoje.weekday())
            return inicio, inicio + timedelta(days=7)
        inicio = pd.Timestamp(hoje.year, hoje.month, 1)
        return inicio, pd.Timestamp(hoje.year + (hoje.month == 12), hoje.month % 12 + 1, 1)

    def periodos_barras(modo):
        # Lista de (rótulo, início, fim), do mais antigo para o atual
        hoje = inicio_do_dia(datetime.now())
        periodos = []
        if modo == "dia":
            for i in range(6, -1, -1):
                dia = hoje - timedelta(days=i)
                periodos.append((dia.strftime("%d/%m"), dia, dia + timedelta(days=1)))
        elif modo == "semana":
            segunda = hoje - timedelta(days=hoje.weekday())
            for i in range(5, -1, -1):
                inicio = segunda - timedelta(weeks=i)
                periodos.append((inicio.strftime("%d/%m"), inicio, inicio + timedelta(days=7)))
        else:
            ano, mes = hoje.year, hoje.month
            for _ in range(6):
                fim_mes = pd.Timestamp(ano + (mes == 12), mes % 12 + 1, 1)
                periodos.insert(0, (NOMES_MESES[mes - 1], pd.Timestamp(ano, mes, 1), fim_mes))
                mes -= 1
                if mes == 0:
                    ano, mes = ano - 1, 12
        return periodos

    def filtrar_periodo(tabela, inicio, fim):
        return tabela[(tabela["_data"] >= inicio) & (tabela["_data"] < fim)]

    def abreviar(valor):
        if valor >= 1_000_000:
            return f"{valor / 1_000_000:.1f}M".replace(".", ",")
        if valor >= 1_000:
            return f"{valor / 1_000:.1f}k".replace(".", ",")
        return f"{valor:.0f}"

    def atualizar_grafico_pizza():
        hoje = datetime.now()
        modo = seletor_pizza.value
        inicio, fim = intervalo_atual(modo)
        descricao_periodo = {
            "mes": f"{NOMES_MESES_COMPLETOS[hoje.month - 1].lower()}",
            "semana": "esta semana",
            "dia": "hoje",
        }[modo]

        do_periodo = filtrar_periodo(registros(), inicio, fim)
        despesas = do_periodo[eh_despesa(do_periodo)]
        gastos = despesas.groupby("_categoria")["_valor"].sum().sort_values(ascending=False)

        if gastos.sum() <= 0:
            conteudo_pizza.content = ft.Container(
                ft.Text(f"Nenhuma despesa registrada {'em ' if modo == 'mes' else ''}{descricao_periodo}.\n"
                        "Escreva no chat, por exemplo: \"gastei 50 no mercado\".",
                        color=TEXTO_TERCIARIO, text_align=ft.TextAlign.CENTER),
                alignment=ft.Alignment.CENTER,
                expand=True,
            )
            return

        # Mais de 5 fatias fica ilegível: mantém as 4 maiores e junta o resto em "Outros"
        if len(gastos) > 5:
            outros = gastos.iloc[4:].sum()
            gastos = gastos.iloc[:4].copy()
            gastos["Outros"] = gastos.get("Outros", 0) + outros

        total = gastos.sum()
        grafico_pizza.sections = []
        legenda_pizza.controls = []
        for indice, (categoria, valor) in enumerate(gastos.items()):
            cor = CORES_PIZZA[indice % len(CORES_PIZZA)]
            grafico_pizza.sections.append(fch.PieChartSection(
                value=float(valor),
                title=f"{valor / total:.0%}",
                title_style=ft.TextStyle(size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                color=cor,
                radius=RAIO_FATIA,
            ))
            # Nome e valor em linhas separadas para não cortar em telas menores
            legenda_pizza.controls.append(ft.Row([
                ft.Container(width=12, height=12, bgcolor=cor, border_radius=6),
                ft.Column([
                    ft.Text(categoria, size=14, color=ft.Colors.WHITE, no_wrap=True),
                    ft.Text(moeda(valor), size=13, color=TEXTO_SECUNDARIO, no_wrap=True),
                ], spacing=0, expand=True),
            ]))

        conteudo_pizza.content = ft.Row([
            ft.Container(grafico_pizza, expand=3),
            ft.Container(legenda_pizza, expand=2),
        ])

    def atualizar_grafico_barras():
        tabela = registros()
        modo = seletor_barras.value
        estilo_tooltip = ft.TextStyle(size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE)
        # O salário é mensal, então só entra nas barras por mês (igual ao card de Receitas)
        incluir_salario = modo == "mes"
        rotulo_entradas = "Entradas (com salário)" if incluir_salario else "Entradas"
        legenda_entradas.value = rotulo_entradas

        grafico_barras.groups = []
        rotulos_x = []
        maior_valor = 0
        for indice, (rotulo, inicio, fim) in enumerate(periodos_barras(modo)):
            entradas, despesas, _ = totais(filtrar_periodo(tabela, inicio, fim), incluir_salario)
            maior_valor = max(maior_valor, entradas, despesas)

            grafico_barras.groups.append(fch.BarChartGroup(
                x=indice,
                spacing=4,
                rods=[
                    fch.BarChartRod(to_y=entradas, width=14, color=VERDE, border_radius=4,
                                    tooltip=fch.BarChartRodTooltip(text=f"{rotulo_entradas}\n{moeda(entradas)}", text_style=estilo_tooltip)),
                    fch.BarChartRod(to_y=despesas, width=14, color=VERMELHO, border_radius=4,
                                    tooltip=fch.BarChartRodTooltip(text=f"Despesas\n{moeda(despesas)}", text_style=estilo_tooltip)),
                ],
            ))
            rotulos_x.append(fch.ChartAxisLabel(
                value=indice,
                label=ft.Container(
                    ft.Text(rotulo, size=14, color=ft.Colors.WHITE),
                    padding=ft.Padding.only(top=8),
                ),
            ))
        grafico_barras.bottom_axis.labels = rotulos_x

        # Folga de 20% no topo para a barra mais alta não encostar na borda
        topo = float(maior_valor) * 1.2 if maior_valor > 0 else 100
        grafico_barras.min_y = 0
        grafico_barras.max_y = topo

        # Eixo Y com 4 divisões e rótulos brancos (o padrão do gráfico é preto)
        passo = topo / 4
        grafico_barras.horizontal_grid_lines.interval = passo
        grafico_barras.left_axis.labels = [
            fch.ChartAxisLabel(
                value=passo * i,
                label=ft.Text(abreviar(passo * i), size=13, color=TEXTO_SECUNDARIO),
            )
            for i in range(5)
        ]

    def mudou_periodo(e):
        atualizar_grafico_pizza()
        atualizar_grafico_barras()
        page.update()

    seletor_pizza = criar_seletor_periodo(PERIODOS_PIZZA, mudou_periodo)
    seletor_barras = criar_seletor_periodo(PERIODOS_BARRAS, mudou_periodo)
    legenda_entradas = ft.Text("Entradas", size=15, color=ft.Colors.WHITE)

    def porcentagem(parte, meta):
        return min(parte / meta, 1.0) if meta > 0 else 0.0

    def atualizar_informacoes():
        resumo = calcular_resumo()
        hoje = datetime.now()

        titulo_resumo.value = f"Resumo de {NOMES_MESES_COMPLETOS[hoje.month - 1]}"

        # Metas
        texto_salario.value = moeda(resumo["salario"])
        texto_reserva.value = moeda(resumo["reserva"])
        texto_renda.value = moeda(resumo["renda"])
        texto_aporte.value = moeda(resumo["aporte"])

        progresso_reserva.value = porcentagem(resumo["investido_total"], resumo["reserva"])
        progresso_renda.value = porcentagem(resumo["investido_total"], resumo["renda"])
        progresso_aporte.value = porcentagem(resumo["investido"], resumo["aporte"])
        legenda_reserva.value = f"6x seu salário · {moeda(resumo['investido_total'])} investidos até agora"
        legenda_renda.value = f"120x seu salário · {progresso_renda.value:.0%} alcançado"
        legenda_aporte.value = f"20% do salário · {moeda(resumo['investido'])} investidos este mês"

        # Resumo do mês
        texto_receitas.value = moeda(resumo["receitas"])
        texto_despesas.value = moeda(resumo["despesas"])
        texto_investimentos.value = moeda(resumo["investido"])
        legenda_investimentos.value = f"Meta: {moeda(resumo['aporte'])}"
        texto_saldo.value = moeda(abs(resumo["saldo"]))
        positivo = resumo["saldo"] >= 0
        texto_saldo.color = VERDE if positivo else VERMELHO
        rotulo_saldo.value = "Sobrou no mês" if positivo else "Faltou no mês"
        rotulo_saldo.color = VERDE if positivo else VERMELHO

        atualizar_area_salario()
        atualizar_grafico_pizza()
        atualizar_grafico_barras()


    # ==========================================
    # LANÇAMENTOS (tabela)
    # ==========================================
    def meses_disponiveis():
        datas = registros()["_data"].dropna()
        meses = sorted({(d.year, d.month) for d in datas}, reverse=True)
        return [ft.DropdownOption(key="todos", text="Todos os meses")] + [
            ft.DropdownOption(key=f"{ano}-{mes}", text=f"{NOMES_MESES_COMPLETOS[mes - 1]}/{ano}")
            for ano, mes in meses
        ]

    def mudou_filtro(e):
        recarregar_tabela()
        page.update()

    filtro_mes = ft.Dropdown(
        value="todos",
        options=[],
        width=220,
        dense=True,
        text_size=14,
        color=ft.Colors.WHITE,
        border_color=AZUL_BORDA,
        focused_border_color=VERDE,
        filled=True,
        fill_color=AZUL_ELEVADO,
        border_radius=10,
        on_select=mudou_filtro,
    )
    contador_lancamentos = ft.Text("", color=TEXTO_SECUNDARIO, size=fonte+1)
    area_tabela = ft.Container()

    def clicou_excluir(e):
        # Primeiro clique arma, segundo confirma: evita excluir sem querer
        botao = e.control
        indice = botao.data
        if botao.icon != ft.Icons.DELETE_FOREVER:
            botao.icon = ft.Icons.DELETE_FOREVER
            botao.icon_color = VERMELHO
            botao.tooltip = "Clique de novo para confirmar"
            botao.update()
            return

        df_anterior = df
        try:
            aplicar_alteracao(remover_registro(df, indice))
        except PermissionError as erro:
            avisar(str(erro), VERMELHO)
            page.update()
            return
        linha = df_anterior.loc[indice]
        desfazer = desfazer_para(df_anterior, df)
        avisar(f"Excluído: {linha['Descrição']} · {moeda(numero(linha['Valor']))}",
               ao_desfazer=lambda e: desfazer() and avisar("Lançamento restaurado."))
        page.update()

    def recarregar_tabela():
        tabela = tabela_financeira
        tabela.columns.clear()
        tabela.rows.clear()

        # Mantém o filtro escolhido se o mês ainda existir
        filtro_mes.options = meses_disponiveis()
        if filtro_mes.value not in [opcao.key for opcao in filtro_mes.options]:
            filtro_mes.value = "todos"

        dados = registros()
        if filtro_mes.value != "todos":
            ano, mes = map(int, filtro_mes.value.split("-"))
            dados = registros_do_mes(dados, ano, mes)
        # Mais recentes primeiro; registros sem data válida vão para o fim
        dados = dados.sort_values("_data", ascending=False, na_position="last", kind="stable")

        quantidade = len(dados)
        contador_lancamentos.value = f"{quantidade} lançamento{'s' if quantidade != 1 else ''}"

        if quantidade == 0:
            area_tabela.content = ft.Container(
                ft.Column([
                    ft.Icon(ft.Icons.RECEIPT_LONG, size=50, color=TEXTO_TERCIARIO),
                    ft.Text("Nenhum lançamento por aqui ainda.", size=fonte+3, color=ft.Colors.WHITE, weight="bold"),
                    ft.Text("Escreva no chat ao lado, por exemplo: \"gastei 50 no mercado hoje\".",
                            color=TEXTO_SECUNDARIO, text_align=ft.TextAlign.CENTER),
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=8),
                alignment=ft.Alignment.CENTER,
                padding=60,
            )
            return
        area_tabela.content = ft.Row(controls=[tabela], scroll=ft.ScrollMode.ADAPTIVE)

        # 1. Cria os cabeçalhos da tabela APENAS para as colunas visíveis
        colunas = [coluna for coluna in visualizacao_atual if coluna in df.columns]
        for coluna in colunas:
            tabela.columns.append(ft.DataColumn(
                ft.Text(ROTULO_COLUNA.get(coluna, coluna), weight="bold", size=fonte+4, color=ft.Colors.WHITE)
            ))
        tabela.columns.append(ft.DataColumn(ft.Text("")))  # Coluna do botão de excluir

        # 2. Preenche as linhas buscando os valores apenas dessas colunas
        for indice, linha in dados.iterrows():
            tipo_da_linha = linha["_tipo"]
            cor_tipo = COR_TIPO.get(tipo_da_linha, ft.Colors.WHITE)
            linhas_celulas = []

            for coluna in colunas:
                valor = linha[coluna] # Pega o valor específico daquela coluna

                if coluna == "Valor":
                    sinal = "+ " if tipo_da_linha == "Entrada" else "− "
                    linhas_celulas.append(ft.DataCell(ft.Text(sinal + moeda(linha["_valor"]), color=cor_tipo, weight="W_600")))

                elif coluna == "Tipo" and tipo_da_linha in COR_TIPO:
                    linhas_celulas.append(ft.DataCell(
                        ft.Container(
                            content=ft.Text(ROTULO_TIPO[tipo_da_linha], color=cor_tipo, weight="bold"),
                            bgcolor=ft.Colors.with_opacity(0.14, cor_tipo),
                            border_radius=10,
                            padding=ft.Padding.only(left=10, right=10, top=5, bottom=5)
                        )
                    ))
                else:
                    texto = str(valor).strip() or "—"
                    linhas_celulas.append(ft.DataCell(ft.Text(texto, color=ft.Colors.WHITE, weight="W_600")))

            linhas_celulas.append(ft.DataCell(ft.IconButton(
                icon=ft.Icons.DELETE_OUTLINE,
                icon_color=TEXTO_TERCIARIO,
                icon_size=20,
                tooltip="Excluir lançamento",
                data=indice,
                on_click=clicou_excluir,
            )))
            tabela.rows.append(ft.DataRow(cells=linhas_celulas))

    planilha = ft.Container(
        content=ft.Column(
            controls=[
                ft.Container(
                    ft.Row([contador_lancamentos, filtro_mes],
                           alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                           vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=ft.Padding.only(top=15, bottom=5, left=10, right=10),
                ),
                area_tabela,
            ],
            scroll=ft.ScrollMode.ADAPTIVE,
            expand=True
        ),
        expand=5,
        padding=ft.Padding.only(left=10, right=10),
        bgcolor=AZUL_CARTAO,
        border_radius=15
    )

    # ==========================================
    # VISÃO GERAL
    # ==========================================
    def card_meta(icone, titulo, valor, legenda, progresso=None, acao=None):
        cabecalho = [ft.Text(titulo, color=TEXTO_SECUNDARIO, size=18, expand=True)]
        if acao:
            cabecalho.append(acao)
        return ft.Container(
            ft.Row([
                ft.Container(
                    ft.Icon(icone, size=40, color=VERDE),
                    bgcolor=ft.Colors.with_opacity(0.12, VERDE),
                    border_radius=18,
                    padding=12,
                ),
                ft.Column(
                    [ft.Row(cabecalho, vertical_alignment=ft.CrossAxisAlignment.CENTER), valor]
                    + ([progresso] if progresso else []) + [legenda],
                    spacing=4,
                    expand=True,
                    alignment=ft.MainAxisAlignment.CENTER,
                )
            ], spacing=16),
            bgcolor=AZUL_CARTAO,
            border=ft.Border.all(1, AZUL_BORDA),
            border_radius=15,
            padding=20,
            height=150,
            expand=True
        )

    def bloco_resumo(titulo, cor, valor, legenda):
        return ft.Container(
            width=170,
            content=ft.Column([
                ft.Text(titulo, color=cor, weight="bold", size=17, no_wrap=True),
                valor,
                legenda,
            ], spacing=2)
        )

    def divisoria():
        return ft.Container(bgcolor=AZUL_BORDA, height=70, width=1)

    botao_editar_salario = ft.IconButton(icon=ft.Icons.EDIT, icon_color=VERDE, icon_size=20,
                                         tooltip="Editar salário", on_click=editar_salario)
    card_salario = card_meta(ft.Icons.WALLET, "Salário", texto_salario,
                             ft.Text("Receita principal", color=TEXTO_TERCIARIO, size=fonte-1),
                             acao=botao_editar_salario)
    # Espaço do card de Salário: vira o formulário enquanto o salário é editado
    espaco_card_salario = ft.Container(card_salario, expand=True, height=150)

    tela_informacoes = ft.Container(
        content=ft.Column(
            scroll=ft.ScrollMode.ALWAYS,
            spacing=12,
            controls=[
                area_salario,

                # Resumo do mês: o que muda no dia a dia fica no topo
                ft.Container(
                    bgcolor=AZUL_CARTAO,
                    border=ft.Border.all(1, AZUL_BORDA),
                    border_radius=15,
                    padding=20,
                    content=ft.Column([
                        ft.Row([
                            ft.Icon(ft.Icons.CALENDAR_MONTH, size=30, color=VERDE),
                            titulo_resumo,
                        ]),
                        ft.Row(
                            wrap=True,
                            spacing=22,
                            run_spacing=20,
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                            controls=[
                                # Saldo em destaque: é a resposta para "como estou este mês?"
                                ft.Container(
                                    ft.Column([rotulo_saldo, texto_saldo,
                                               ft.Text("Receitas − despesas − investido", color=TEXTO_TERCIARIO, size=fonte-1)],
                                              spacing=2),
                                    bgcolor=ft.Colors.with_opacity(0.06, ft.Colors.WHITE),
                                    border_radius=12,
                                    padding=ft.Padding.symmetric(horizontal=20, vertical=12),
                                ),
                                bloco_resumo("Receitas", VERDE, texto_receitas,
                                             ft.Text("Salário + entradas", color=TEXTO_TERCIARIO, size=fonte-1)),
                                divisoria(),
                                bloco_resumo("Despesas", VERMELHO, texto_despesas,
                                             ft.Text("Gastos + contas fixas", color=TEXTO_TERCIARIO, size=fonte-1)),
                                divisoria(),
                                bloco_resumo("Investido", AZUL_CLARO, texto_investimentos, legenda_investimentos),
                            ]
                        )
                    ])
                ),

                ft.Row(
                    alignment=ft.MainAxisAlignment.CENTER,
                    controls=[
                        # Gráfico de pizza: despesas do período por categoria
                        ft.Container(
                            bgcolor=AZUL_CARTAO,
                            border=ft.Border.all(1, AZUL_BORDA),
                            height=320,
                            expand=5,
                            border_radius=15,
                            padding=20,
                            content=ft.Column([
                                ft.Row(
                                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                    controls=[
                                        ft.Row([ft.Icon(ft.Icons.PIE_CHART, size=26, color=VERDE),
                                                ft.Text("Despesas por categoria", color=ft.Colors.WHITE, weight="bold", size=20)]),
                                        seletor_pizza,
                                    ],
                                ),
                                conteudo_pizza,
                            ])
                        ),
                        # Gráfico de barras: entradas x despesas por período
                        ft.Container(
                            bgcolor=AZUL_CARTAO,
                            border=ft.Border.all(1, AZUL_BORDA),
                            height=320,
                            expand=6,
                            border_radius=15,
                            padding=20,
                            content=ft.Column([
                                ft.Row(
                                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                    controls=[
                                        ft.Row([
                                            ft.Icon(ft.Icons.BAR_CHART, size=26, color=VERDE),
                                            ft.Text("Entradas x Despesas", color=ft.Colors.WHITE, weight="bold", size=20),
                                        ]),
                                        seletor_barras,
                                    ],
                                ),
                                ft.Row(spacing=8, controls=[
                                    ft.Container(width=12, height=12, bgcolor=VERDE, border_radius=6),
                                    legenda_entradas,
                                    ft.Container(width=12, height=12, bgcolor=VERMELHO, border_radius=6),
                                    ft.Text("Despesas", size=15, color=ft.Colors.WHITE),
                                ]),
                                ft.Container(grafico_barras, expand=True),
                            ])
                        ),
                    ]
                ),

                ft.Container(
                    ft.Row([
                        ft.Icon(ft.Icons.FLAG, size=30, color=VERDE),
                        ft.Text("Suas metas", color=ft.Colors.WHITE, weight="bold", size=25),
                        ft.Text("calculadas a partir do seu salário", color=TEXTO_TERCIARIO, size=fonte),
                    ], vertical_alignment=ft.CrossAxisAlignment.END),
                    padding=ft.Padding.only(left=5, top=10),
                ),
                ft.Row([
                    espaco_card_salario,
                    card_meta(ft.Icons.SHOW_CHART, "Meta de investimento mensal", texto_aporte,
                              legenda_aporte, progresso_aporte),
                ]),
                ft.Row([
                    card_meta(ft.Icons.SHIELD, "Meta de reserva de emergência", texto_reserva,
                              legenda_reserva, progresso_reserva),
                    card_meta(ft.Icons.SAVINGS, "Meta para viver de renda", texto_renda,
                              legenda_renda, progresso_renda),
                ]),
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
                                size=fonte+3,
                                color=ft.Colors.WHITE,
                                selectable=True)
        conteudo = ft.Column([texto_control], spacing=6)

        balao = ft.Container(
            expand=7, # 70% da largura do chat, acompanhando o tamanho da janela
            bgcolor=ft.Colors.with_opacity(0.16, VERDE) if user else AZUL_CARTAO,
            border=ft.Border.all(1, ft.Colors.with_opacity(0.35, VERDE) if user else AZUL_BORDA),
            border_radius=20,
            padding=12,
            content=conteudo)
        espaco = ft.Container(expand=3)
        row = ft.Row(controls=[espaco, balao] if user else [balao, espaco])
        return row, texto_control, conteudo

    def digitar(txt_stream, texto: str, cor=ft.Colors.WHITE):
        # Efeito de digitação, letra por letra
        txt_stream.value = ""
        txt_stream.color = cor
        for char in texto:
            txt_stream.value += char
            page.update()
            sleep(.01)

    def chip_sugestao(texto):   
        return ft.Container(
            ft.Text(texto, size=fonte, color=ft.Colors.WHITE),
            bgcolor=AZUL_ELEVADO,
            border=ft.Border.all(1, AZUL_BORDA),
            border_radius=20,
            padding=ft.Padding.symmetric(horizontal=14, vertical=8),
            ink=True,
            on_click=lambda e: enviar_texto(texto),
            tooltip="Clique para enviar",
        )

    # Card de boas-vindas: explica o chat e some depois da primeira mensagem
    card_boas_vindas = ft.Container(
        ft.Column([
            ft.Row([
                ft.Icon(ft.Icons.AUTO_AWESOME, color=VERDE, size=26),
                ft.Text("Como posso ajudar?", size=fonte+5, weight="bold", color=ft.Colors.WHITE),
            ]),
            ft.Text("Escreva do seu jeito e eu registro na sua planilha. "
                    "Também posso remover lançamentos e tirar dúvidas sobre dinheiro.",
                    color=TEXTO_SECUNDARIO, size=fonte+1),
            ft.Text("Experimente:", color=TEXTO_TERCIARIO, size=fonte, weight="bold"),
            ft.Row([chip_sugestao(s) for s in SUGESTOES_CHAT], wrap=True, spacing=8, run_spacing=8),
        ], spacing=12),
        bgcolor=AZUL_CARTAO,
        border=ft.Border.all(1, ft.Colors.with_opacity(0.4, VERDE)),
        border_radius=16,
        padding=18,
    )
    lista_mensagens.controls.append(card_boas_vindas)

    def esconder_boas_vindas():
        if card_boas_vindas in lista_mensagens.controls:
            lista_mensagens.controls.remove(card_boas_vindas)

    campo_mensagem = ft.TextField(hint_text="Ex.: gastei 35 no almoço",
                                  expand=True,
                                  color=ft.Colors.WHITE,
                                  bgcolor=AZUL_CARTAO,
                                  filled=True,
                                  border_color=AZUL_BORDA,
                                  focused_border_color=VERDE,
                                  cursor_color=VERDE,
                                  hint_style=ft.TextStyle(color=TEXTO_TERCIARIO),
                                  border_radius=15)

    def botao_desfazer_chat(conteudo, desfazer):
        def clicou(e):
            if desfazer():
                conteudo.controls[-1] = ft.Text("Desfeito.", color=TEXTO_TERCIARIO, size=fonte)
                page.update()
        conteudo.controls.append(ft.TextButton("Desfazer", icon=ft.Icons.UNDO, on_click=clicou,
                                               style=ft.ButtonStyle(color=VERDE, padding=0)))

    def enviar_texto(texto_usuario):
        texto_usuario = texto_usuario.strip()
        if texto_usuario == "":
            return
        esconder_boas_vindas()
        # Adiciona o texto na lista de mensagens
        lista_mensagens.controls.append(mensagem(texto_usuario)[0])
        campo_mensagem.value = ""
        page.update()

        def processar_ia():
            sleep(.2) # Pequeno atraso estético

            mensagem_ia, txt_stream, conteudo = mensagem("Processando...", user=False)
            lista_mensagens.controls.append(mensagem_ia)
            page.update()

            try:
                df_anterior = df
                novo_df, acao, mensagem_chat = registerData(df, texto_usuario)

                if acao in (0, 1):
                    # Só troca o df depois de salvar, para a tela não mostrar algo que não está no arquivo
                    aplicar_alteracao(novo_df)
                    digitar(txt_stream, mensagem_chat, VERDE if acao == 1 else ft.Colors.WHITE)
                    botao_desfazer_chat(conteudo, desfazer_para(df_anterior, df))
                else:
                    digitar(txt_stream, mensagem_chat)
                page.update()
            except Exception as exc:
                txt_stream.value = f"Erro ao processar: {exc}"
                txt_stream.color = VERMELHO
                page.update()

        threading.Thread(target=processar_ia, daemon=True).start()

    def enviar_mensagem(e):
        enviar_texto(campo_mensagem.value)

    # O botão de enviar dispara a função acima
    botao_enviar = ft.Container(
        content=ft.IconButton(icon=ft.Icons.SEND, on_click=enviar_mensagem, icon_size=35, icon_color=ft.Colors.WHITE,
                              tooltip="Enviar"),
        gradient=ft.LinearGradient(colors=GRADIENTE_VERDE, begin=ft.Alignment.TOP_LEFT, end=ft.Alignment.BOTTOM_RIGHT),
        border_radius=180,
    )
    # Apertar "Enter" no campo de texto também envia
    campo_mensagem.on_submit = enviar_mensagem

    coluna_chat = ft.Container(
        content=ft.Column(
            controls=[
                ft.Container(
                    ft.Column([
                        ft.Row([
                            ft.Text("Assistente", size=fonte+15, weight="bold", color=ft.Colors.WHITE),
                            ft.Icon(ft.Icons.AUTO_AWESOME, size=30, color=VERDE)
                        ]),
                        ft.Text("Registre gastos e tire dúvidas conversando", size=fonte+1, color=TEXTO_SECUNDARIO),
                    ], spacing=0),
                    padding=ft.Padding.only(left=10, top=10, bottom=10),
                ),
                lista_mensagens, # Ocupa o meio do chat
                ft.Row([campo_mensagem, botao_enviar]) # Fica na base do chat
            ],
        ),
        width=375,
        expand=2, # Ocupa 2 partes do espaço
        padding=10,
        bgcolor=AZUL_PAINEL,
        border=ft.Border.all(1, AZUL_BORDA),
        border_radius=15,
        animate=ft.Animation(300, ft.AnimationCurve.EASE_OUT),
    )

    # ==========================================
    # JUNTANDO TUDO NA PÁGINA
    # ==========================================
    # Colocamos os três blocos dentro de uma Row principal (Linha)
    layout_principal = ft.Row(
        controls=[
            coluna_botoes,
            tela_central,
            coluna_chat
        ],
        expand=True # Faz a linha ocupar 100% da altura da tela
    )

    # Prepara as duas telas antes de exibir; a Visão geral é a tela inicial
    recarregar_tabela()
    mostrar_visao("informacoes", atualizar_pagina=False)

    page.add(layout_principal)

if __name__ == "__main__":
    ft.run(Main, view=ft.AppView.WEB_BROWSER,
           host="0.0.0.0",
           port=8550)
