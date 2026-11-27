# cc4p_ANBD(AV1_AV2)
ENGENHARIA DE REQUISITOS / ANÁLISE DE NEGÓCIOS COM BIG DATA (ANBG)

# Sistema de Controle de Biblioteca Digital (Tema 8)

## 1. Informações do Projeto
Este projeto foi desenvolvido no âmbito da disciplina de **Análise de Dados e Big Data** (com continuidade em **Engenharia de Requisitos**). O sistema simula o ambiente de uma biblioteca digital, gerenciando o cadastro de livros e leitores, a consulta de disponibilidade, o registro de empréstimos e devoluções e a geração de indicadores de leitura.

* **Problema que resolve:** Automatiza o controle de circulação do acervo (empréstimos e devoluções) e o cálculo de indicadores de desempenho literário que seriam demorados e sujeitos a falhas se feitos manualmente.
* **Objetivo do sistema:** Fornecer uma ferramenta em Python capaz de persistir os registros do acervo e movimentações em planilhas Excel, validar regras condicionais de empréstimo e gerar visualizações gráficas (dashboards) para apoiar decisões de curadoria e gestão.
* **Relação com o Minimundo e Caso de Uso:** Atende diretamente ao minimundo de uma biblioteca digital, contemplando o caso de uso principal em que o usuário consulta a disponibilidade de uma obra antes de solicitar o empréstimo.

---

## 2. Informações da Disciplina
* **Universidade:** Universidade Veiga de Almeida (UVA)
* **Curso:** Ciência da Computação
* **Disciplina:** Análise de Dados e Big Data
* **Ano/Semestre:** 2026.2

---

## 3. Quem fez o trabalho
* **[Caiã Lopes Tenente]** — Matrícula: `[1250111001]`

---

## 4. Funcionalidades do Sistema
O sistema entrega todas as funcionalidades principais exigidas pelo projeto:

* **Cadastro de Acervo e Usuários (`cadastrar_livro` e `cadastrar_usuario`):** Registra novas obras (título, autor, categoria) e leitores com ID incremental e validação contra cadastros duplicados.
* **Consulta de Disponibilidade (`consultar_disponibilidade`):** Permite ao usuário pesquisar livros por título ou autor, informando em tempo real se a obra está **Disponível** ou **Indisponível (Emprestada)** e seu histórico de leituras.
* **Controle de Empréstimos e Devoluções (`realizar_emprestimo` e `realizar_devolucao`):** Aplica regras condicionais (`if/elif/else`) para verificar se o livro existe, se o usuário está ativo e se o exemplar está livre; atualiza o status de disponibilidade, registra o histórico com data/hora e incrementa o contador de empréstimos.
* **Geração de Indicadores e Dashboards (`gerar_relatorio_e_graficos`):** Consolida métricas gerenciais (taxa de disponibilidade, giro médio do acervo e ranking dos livros mais lidos) e gera gráficos comparativos em barras.
* **Persistência em Excel (`inicializar_base_excel` e `salvar_dados_excel`):** Salva e atualiza automaticamente as abas de `Livros`, `Usuarios`, `Historico` e `Indicadores_Mais_Lidos` em arquivo `.xlsx`.

---

## 5. Tecnologias Utilizadas
* **Linguagem:** Python 3.15
* **Manipulação de Dados:** Pandas
* **Visualização de Dados:** Matplotlib & Seaborn
* **Persistência de Dados:** Planilhas Excel (`.xlsx` com suporte do motor `openpyxl`)
* **Controle de Versão:** Git e GitHub

---

## 6. Instalação e Configuração
Siga os passos abaixo para configurar o ambiente em sua máquina local:

```bash
# 1. Clone o repositório
git clone <link-do-repositorio>
cd <nome-do-projeto>

# 2. Crie e ative um ambiente virtual Python
python -m venv venv

# No Windows:
venv\Scripts\activate

# No macOS/Linux:
# source venv/bin/activate

# 3. Instale as dependências necessárias
pip install pandas matplotlib seaborn openpyxl
