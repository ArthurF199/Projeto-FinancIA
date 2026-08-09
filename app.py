# import flet as ft


# def main(page: ft.Page):
#     botoes = ft.Column(
#         controls = [
#             ft.Button(content="Registrar Salário"),
#             ft.Button(content="Analisar a planilha"),
#             ft.Button(content="Registrar/Remover Dados"),
#             ft.Button(content="Reserva de Emergência"),
#             ft.Button(content="Viver de Renda")
#         ],
#         alignment=ft.MainAxisAlignment.END,
#         horizontal_alignment=ft.CrossAxisAlignment.CENTER,

#     )

#     page.add(
#         botoes
#     )


# ft.app(target=main)


import flet as ft
import pandas as pd
from ia import *


def Main(page: ft.Page):
    # Configurações da página
    page.title = "Dashboard Flet"
    page.padding = 20
    # O theme_mode pode ser LIGHT ou DARK
    page.theme_mode = ft.ThemeMode.DARK 

    # ==========================================
    # 1. COLUNA ESQUERDA (Botões)
    # ==========================================
    coluna_botoes = ft.Container(
        content=ft.Column(
            controls=[
                ft.Text("Menu", size=20, weight="bold"),
                ft.ElevatedButton("Dashboard", icon=ft.Icons.DASHBOARD, width=200),
                ft.ElevatedButton("Relatórios", icon=ft.Icons.PIE_CHART, width=200),
                ft.ElevatedButton("Configurações", icon=ft.Icons.SETTINGS, width=200),
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
    
    planilha = ft.Container(
        content=ft.Column(
            controls=[
                ft.Text("Planilha Financeira", size=20, weight="bold"),
                ft.DataTable(
                    columns=[],
                    rows=[]
                )
            ],
            scroll=ft.ScrollMode.ADAPTIVE,
            expand=True
        ),
        expand=5, # Ocupa 3 partes do espaço
        padding=10
    )
    
    def recarregar_tabela():
        tabela = planilha.content.controls[1]
        tabela.columns.clear()
        tabela.rows.clear()
        
        for coluna in df.columns:
            tabela.columns.append(ft.DataColumn(ft.Text(coluna, weight="bold")))

        for indice, linha in df.iterrows():
            linhas = []
            for valor in linha:
                linhas.append(ft.DataCell(ft.Text(str(valor))))
            tabela.rows.append(ft.DataRow(cells=linhas))

    
    def modificar_dados(e):
        n = len(df)
        df.drop("Descrição", axis=1, inplace=True)
        
        
        recarregar_tabela()
        page.update()

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
            lista_mensagens.controls.append(ft.Text(f"Você: {campo_mensagem.value}"))
            mensagem_usuario = campo_mensagem.value
            # Limpa o campo
            campo_mensagem.value = ""
            page.update()

            try:
                from main import registerData

                nonlocal df
                novo_df, acao = registerData(df, mensagem_usuario)
                df = novo_df
                df.to_excel("data.xlsx", index=False)
                recarregar_tabela()

                if acao == 0:
                    resposta = "Registro removido com sucesso."
                else:
                    resposta = "Registro adicionado com sucesso."

                lista_mensagens.controls.append(ft.Text(f"FinancIA: {resposta}", color=ft.Colors.GREEN))
            except Exception as exc:
                lista_mensagens.controls.append(ft.Text(f"FinancIA: Erro ao processar: {exc}", color=ft.Colors.RED))

            page.update()
            return campo_mensagem.value
            

    
    # O botão de enviar dispara a função acima
    botao_enviar = ft.IconButton(icon=ft.Icons.SEND, on_click=enviar_mensagem)
    # Apertar "Enter" no campo de texto também envia
    campo_mensagem.on_submit = enviar_mensagem

    coluna_chat = ft.Container(
        content=ft.Column(
            controls=[
                ft.Text("FinancIA", size=20, weight="bold"),
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