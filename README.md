<p align="center">
  <img src="assets/logo.png" alt="Logo do FinancIA" width="160">
</p>

<h1 align="center">FinancIA</h1>

<p align="center">
  <b>Seu gestor financeiro com Inteligência Artificial local.</b><br>
  Descreva um gasto do jeito que você falaria. A IA entende, categoriza e registra na sua planilha.
</p>

<p align="center">
  <img alt="Python" src="https://img.shields.io/badge/Python-3.10+-0B1F3A?logo=python&logoColor=white">
  <img alt="Flet" src="https://img.shields.io/badge/Interface-Flet-0B1F3A">
  <img alt="Ollama" src="https://img.shields.io/badge/IA-Ollama%20%2B%20Gemma%204-1DB97A">
  <img alt="Local" src="https://img.shields.io/badge/100%25-Local-1DB97A">
</p>

---

## Sobre o projeto

O **FinancIA** é um gestor financeiro pessoal que usa uma **IA rodando na sua própria máquina** (hoje o **Gemma 4**, via [Ollama](https://ollama.com)) para tirar o trabalho chato de anotar gastos.

Em vez de preencher formulários, você só escreve no chat:

> *"Gastei 120,50 no mercado ontem"*
> *"Todo dia 10 pago R$ 89,90 da internet"*
> *"Recebi 800 de um freela de design"*

A IA interpreta a mensagem, descobre a **descrição**, o **valor**, a **data**, o **tipo** (entrada, saída ou conta) e a **categoria**, e grava tudo em uma planilha Excel (`data.xlsx`). Depois é só abrir o **Dashboard** para ver para onde o seu dinheiro está indo.

Como a IA é local, **seus dados financeiros não saem do seu computador**.

## Funcionalidades

- **Registro por linguagem natural**: escreva gastos, contas e recebimentos como se estivesse conversando. Datas relativas ("ontem", "sexta passada") são convertidas automaticamente.
- **Remoção por conversa**: *"Apaga aquele gasto de 300 no restaurante"* encontra e remove o registro certo.
- **Categorização automática**: entradas (Salário, Renda Extra, Rendimentos) e despesas (Alimentação, Moradia, Contas de Casa, Transporte, Saúde, Educação, Lazer, Assinaturas, Compras, Dívidas, Investimentos, Outros).
- **Contas recorrentes**: contas com vencimento ganham o próximo **dia de pagamento** calculado.
- **Planilha financeira**: todos os registros visíveis numa tabela e salvos em `data.xlsx`, que você pode abrir no Excel.
- **Dashboard**:
  - Salário, Reserva de Emergência (6 salários), meta de Viver de Renda (120 salários) e Aporte Mensal sugerido (20% do salário).
  - Resumo do mês: receitas, despesas, investimentos e saldo.
  - Gráfico de pizza com as despesas do mês por categoria.
  - Gráfico de barras de Entradas x Despesas por mês.
- **Relatório com IA**: a IA analisa seus números e diz, de forma direta, se você está gerenciando bem o dinheiro e o que pode melhorar.
- **Assistente financeiro**: perguntas que não são registros (*"Vale a pena financiar um carro em 60 vezes?"*) são respondidas no chat.

## Como funciona

```
 Você escreve no chat
         │
         ▼
 ┌──────────────────┐     prompt + últimos registros      ┌─────────────────────┐
 │  FinancIA (Flet) │ ──────────────────────────────────▶ │  Ollama + Gemma 4   │
 │                  │ ◀────────────────────────────────── │  (na sua máquina)   │
 └──────────────────┘        JSON estruturado             └─────────────────────┘
         │
         ▼   valida tipo, categoria, valor e datas
   data.xlsx  ──▶  Planilha  +  Dashboard
```

1. A mensagem é enviada ao modelo junto com os registros atuais e exemplos de como responder.
2. O modelo devolve um JSON com a **ação** (registrar, remover ou só conversar) e os campos do registro.
3. O FinancIA valida a resposta (tipo e categoria precisam existir, valores viram números, datas são conferidas) e só então salva na planilha.

## Tecnologias

| Camada | Ferramenta |
| --- | --- |
| Interface | [Flet](https://flet.dev) + `flet-charts` |
| Dados | `pandas` + `openpyxl` (planilha `data.xlsx`) |
| IA | [Ollama](https://ollama.com) com o modelo `gemma4:e4b` (API REST via `requests`) |
| Configuração | `python-dotenv` |

## Instalação

### Pré-requisitos

- Python 3.10 ou superior
- [Ollama](https://ollama.com/download) instalado

### Passo a passo

1. Clone o repositório:

   ```bash
   git clone https://github.com/ArthurF199/Projeto-FinancIA.git
   cd Projeto-FinancIA
   ```

2. Instale as dependências:

   ```bash
   pip install -r requirements.txt
   ```

3. Baixe o modelo de IA no Ollama:

   ```bash
   ollama pull gemma4:e4b
   ```

4. (Opcional) Se o Ollama roda em **outro computador** da rede, crie um arquivo `.env` na raiz com o IP dele:

   ```env
   IP_DESKTOP=192.168.0.10
   ```

   Sem o `.env`, o FinancIA usa `127.0.0.1` (a própria máquina).

5. Inicie o aplicativo:

   ```bash
   python app.py
   ```

   O FinancIA abre no navegador em `http://localhost:8550`.

## Como usar

| Você quer... | Faça assim |
| --- | --- |
| Registrar um gasto | *"Paguei 45 reais de Uber hoje"* |
| Registrar uma conta fixa | *"Todo dia 5 pago 1.200 de aluguel"* |
| Registrar um recebimento | *"Recebi 300 de dividendos"* |
| Remover um registro | *"Apaga o Uber de 45 reais"* |
| Tirar uma dúvida | *"Quanto devo guardar por mês?"* |
| Cadastrar o salário | Botão **Registrar Salário** no menu |
| Ver gráficos e metas | Botão **Informações** (alterna Planilha / Dashboard) |
| Receber uma análise | Botão **Relatório** |

> **Dica:** feche o `data.xlsx` no Excel enquanto usa o FinancIA, senão ele não consegue salvar.

## Trocando o modelo de IA

O modelo é definido em [`ia.py`](ia.py), no campo `"model"` do payload. Para usar outro modelo do Ollama (por exemplo `llama3.1` ou `qwen2.5`), baixe-o com `ollama pull <modelo>` e troque o nome ali.

## Estrutura do projeto

```
FinancIA/
├── app.py            # Interface (Flet): menu, planilha, dashboard e chat
├── main.py           # Regras de negócio: categorias, prompt, registro e remoção
├── ia.py             # Comunicação com o Ollama (Gemma 4)
├── data.xlsx         # Sua planilha financeira
├── assets/           # Logo e favicon
├── panfleto.html     # Panfleto de apresentação do projeto
└── requirements.txt
```

## Autor

Feito por **Arthur** ([@ArthurF199](https://github.com/ArthurF199)).

<p align="center"><i>Organize hoje, conquiste amanhã.</i></p>
