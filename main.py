# ==============================================================================
# SISTEMA DE CONTROLE DE BIBLIOTECA DIGITAL - AV1 (ANBG)
# Stack: Python + Pandas + Matplotlib + Seaborn + OpenPyXL (.xlsx)
# ==============================================================================

import os
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Configuração visual padrão dos gráficos
sns.set_theme(style="whitegrid", palette="viridis")

ARQUIVO_EXCEL = "biblioteca_digital_av1.xlsx"


# ==============================================================================
# 1. INICIALIZAÇÃO E PERSISTÊNCIA DE DADOS (.XLSX COM OPENPYXL)
# ==============================================================================
def inicializar_base_excel(caminho_arquivo=ARQUIVO_EXCEL):
    """Cria a planilha .xlsx inicial caso ainda não exista e carrega os DataFrames."""
    if not os.path.exists(caminho_arquivo):
        df_livros_init = pd.DataFrame([
            {
                "id": 1,
                "titulo": "Clean Code",
                "autor": "Robert C. Martin",
                "categoria": "Engenharia de Software",
                "disponivel": True,
                "total_emprestimos": 14,
                "emprestado_para": None,
            },
            {
                "id": 2,
                "titulo": "Big Data: A Revolution",
                "autor": "Viktor Mayer-Schönberger",
                "categoria": "Big Data",
                "disponivel": False,
                "total_emprestimos": 21,
                "emprestado_para": "Ana Souza",
            },
            {
                "id": 3,
                "titulo": "Python para Análise de Dados",
                "autor": "Wes McKinney",
                "categoria": "Ciência de Dados",
                "disponivel": True,
                "total_emprestimos": 18,
                "emprestado_para": None,
            },
            {
                "id": 4,
                "titulo": "Engenharia de Dados com Python",
                "autor": "Paul Crickard",
                "categoria": "Engenharia de Dados",
                "disponivel": True,
                "total_emprestimos": 9,
                "emprestado_para": None,
            },
            {
                "id": 5,
                "titulo": "Arquitetura Limpa",
                "autor": "Robert C. Martin",
                "categoria": "Engenharia de Software",
                "disponivel": True,
                "total_emprestimos": 11,
                "emprestado_para": None,
            },
        ])

        df_usuarios_init = pd.DataFrame([
            {"id": 101, "nome": "Ana Souza", "ativo": True},
            {"id": 102, "nome": "Carlos Lima", "ativo": True},
            {"id": 103, "nome": "Marina Reis", "ativo": False},
        ])

        df_historico_init = pd.DataFrame(columns=["data_hora", "operacao", "id_livro", "titulo_livro", "usuario"])

        salvar_dados_excel(df_livros_init, df_usuarios_init, df_historico_init, caminho_arquivo)

    df_livros = pd.read_excel(caminho_arquivo, sheet_name="Livros", engine="openpyxl")
    df_usuarios = pd.read_excel(caminho_arquivo, sheet_name="Usuarios", engine="openpyxl")
    df_historico = pd.read_excel(caminho_arquivo, sheet_name="Historico", engine="openpyxl")
    return df_livros, df_usuarios, df_historico


def salvar_dados_excel(df_livros, df_usuarios, df_historico, caminho_arquivo=ARQUIVO_EXCEL):
    """Persiste as tabelas de Livros, Usuários, Histórico e Indicadores no Excel (.xlsx)."""
    df_ranking = df_livros.sort_values(by="total_emprestimos", ascending=False).copy()
    df_ranking["posicao"] = range(1, len(df_ranking) + 1)

    with pd.ExcelWriter(caminho_arquivo, engine="openpyxl", mode="w") as writer:
        df_livros.to_excel(writer, sheet_name="Livros", index=False)
        df_usuarios.to_excel(writer, sheet_name="Usuarios", index=False)
        df_historico.to_excel(writer, sheet_name="Historico", index=False)
        df_ranking[["posicao", "id", "titulo", "autor", "categoria", "total_emprestimos", "disponivel"]].to_excel(
            writer, sheet_name="Indicadores_Mais_Lidos", index=False
        )


# ==============================================================================
# 2. FUNÇÕES DE NEGÓCIO (CADASTRO, CONSULTA, EMPRÉSTIMO E DEVOLUÇÃO)
# ==============================================================================
def cadastrar_livro(df_livros, titulo, autor, categoria):
    titulo_limpo = titulo.strip()
    autor_limpo = autor.strip()
    categoria_limpo = categoria.strip() if categoria.strip() else "Geral"

    if not titulo_limpo or not autor_limpo:
        print("[ERRO] Título e autor são obrigatórios.")
        return df_livros

    duplicado = df_livros[
        (df_livros["titulo"].str.lower() == titulo_limpo.lower())
        & (df_livros["autor"].str.lower() == autor_limpo.lower())
    ]
    if not duplicado.empty:
        print(f"[AVISO] O livro '{titulo_limpo}' de {autor_limpo} já está cadastrado.")
        return df_livros

    novo_id = int(df_livros["id"].max()) + 1 if not df_livros.empty else 1
    novo_registro = pd.DataFrame([{
        "id": novo_id,
        "titulo": titulo_limpo,
        "autor": autor_limpo,
        "categoria": categoria_limpo,
        "disponivel": True,
        "total_emprestimos": 0,
        "emprestado_para": None,
    }])

    df_livros = pd.concat([df_livros, novo_registro], ignore_index=True)
    print(f"[CADASTRO] Livro '{titulo_limpo}' cadastrado com ID {novo_id}.")
    return df_livros


def cadastrar_usuario(df_usuarios, nome):
    nome_limpo = nome.strip()
    if not nome_limpo:
        print("[ERRO] Nome do usuário não pode ser vazio.")
        return df_usuarios

    if not df_usuarios[df_usuarios["nome"].str.lower() == nome_limpo.lower()].empty:
        print(f"[AVISO] Usuário '{nome_limpo}' já existe no sistema.")
        return df_usuarios

    novo_id = int(df_usuarios["id"].max()) + 1 if not df_usuarios.empty else 101
    novo_usuario = pd.DataFrame([{"id": novo_id, "nome": nome_limpo, "ativo": True}])
    df_usuarios = pd.concat([df_usuarios, novo_usuario], ignore_index=True)
    print(f"[CADASTRO] Usuário '{nome_limpo}' cadastrado com ID {novo_id}.")
    return df_usuarios


def consultar_disponibilidade(df_livros, termo_busca):
    termo = termo_busca.strip()
    if not termo:
        print("[ERRO] Informe um título ou autor para consulta.")
        return

    resultados = df_livros[
        df_livros["titulo"].str.contains(termo, case=False, na=False)
        | df_livros["autor"].str.contains(termo, case=False, na=False)
    ]

    if resultados.empty:
        print(f"[CONSULTA] Nenhum livro encontrado para '{termo_busca}'.")
        return

    print(f"\n[CONSULTA] Resultados para '{termo_busca}':")
    for _, row in resultados.iterrows():
        status = "Disponível" if row["disponivel"] else f"Indisponível (Com: {row['emprestado_para']})"
        print(
            f"  - ID {row['id']}: '{row['titulo']}' ({row['autor']}) "
            f"[{row['categoria']}] -> {status} | Leituras: {row['total_emprestimos']}"
        )


def realizar_emprestimo(df_livros, df_usuarios, df_historico, id_livro, id_usuario):
    idx_livro = df_livros.index[df_livros["id"] == id_livro].tolist()
    usuario_match = df_usuarios[df_usuarios["id"] == id_usuario]

    if not idx_livro:
        print(f"[EMPRÉSTIMO NEGADO] Livro ID {id_livro} não encontrado.")
        return df_livros, df_historico
    if usuario_match.empty:
        print(f"[EMPRÉSTIMO NEGADO] Usuário ID {id_usuario} não encontrado.")
        return df_livros, df_historico

    i = idx_livro[0]
    usuario = usuario_match.iloc[0]

    if not bool(usuario["ativo"]):
        print(f"[EMPRÉSTIMO NEGADO] Usuário '{usuario['nome']}' está inativo.")
        return df_livros, df_historico

    if not bool(df_livros.loc[i, "disponivel"]):
        print(
            f"[EMPRÉSTIMO NEGADO] '{df_livros.loc[i, 'titulo']}' já está emprestado "
            f"para {df_livros.loc[i, 'emprestado_para']}."
        )
        return df_livros, df_historico

    # Atualiza status, vínculo e contador de empréstimos
    df_livros.loc[i, "disponivel"] = False
    df_livros.loc[i, "emprestado_para"] = usuario["nome"]
    df_livros.loc[i, "total_emprestimos"] = int(df_livros.loc[i, "total_emprestimos"]) + 1

    novo_evento = pd.DataFrame([{
        "data_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "operacao": "EMPRÉSTIMO",
        "id_livro": id_livro,
        "titulo_livro": df_livros.loc[i, "titulo"],
        "usuario": usuario["nome"],
    }])
    df_historico = pd.concat([df_historico, novo_evento], ignore_index=True)

    print(
        f"[EMPRÉSTIMO APROVADO] '{df_livros.loc[i, 'titulo']}' emprestado para {usuario['nome']} "
        f"(Total de empréstimos: {df_livros.loc[i, 'total_emprestimos']})."
    )
    return df_livros, df_historico


def realizar_devolucao(df_livros, df_historico, id_livro):
    idx_livro = df_livros.index[df_livros["id"] == id_livro].tolist()

    if not idx_livro:
        print(f"[DEVOLUÇÃO ERRO] Livro ID {id_livro} não encontrado.")
        return df_livros, df_historico

    i = idx_livro[0]
    if bool(df_livros.loc[i, "disponivel"]):
        print(f"[DEVOLUÇÃO AVISO] '{df_livros.loc[i, 'titulo']}' já consta como disponível.")
        return df_livros, df_historico

    leitor_anterior = df_livros.loc[i, "emprestado_para"]
    df_livros.loc[i, "disponivel"] = True
    df_livros.loc[i, "emprestado_para"] = None

    novo_evento = pd.DataFrame([{
        "data_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "operacao": "DEVOLUÇÃO",
        "id_livro": id_livro,
        "titulo_livro": df_livros.loc[i, "titulo"],
        "usuario": leitor_anterior,
    }])
    df_historico = pd.concat([df_historico, novo_evento], ignore_index=True)

    print(f"[DEVOLUÇÃO CONCLUÍDA] '{df_livros.loc[i, 'titulo']}' devolvido por {leitor_anterior}.")
    return df_livros, df_historico


# ==============================================================================
# 3. INDICADORES (PANDAS) E VISUALIZAÇÃO (MATPLOTLIB + SEABORN)
# ==============================================================================
def gerar_relatorio_e_graficos(df_livros, salvar_imagem=True):
    ranking = df_livros.sort_values(by="total_emprestimos", ascending=False).reset_index(drop=True)
    ranking["posicao"] = ranking.index + 1

    total_livros = len(df_livros)
    disponiveis = int(df_livros["disponivel"].sum())
    emprestados = total_livros - disponiveis
    total_leituras = int(df_livros["total_emprestimos"].sum())
    giro_medio = df_livros["total_emprestimos"].mean()

    print("\n" + "=" * 68)
    print("INDICADORES DE DESEMPENHO - BIBLIOTECA DIGITAL (AV1)")
    print("=" * 68)
    print(f"Total de obras no acervo : {total_livros}")
    print(f"Disponíveis vs Emprestados: {disponiveis} disponíveis ({disponiveis/total_livros:.1%}) | {emprestados} emprestados")
    print(f"Total de empréstimos     : {total_leituras} | Giro médio por obra: {giro_medio:.2f}")
    print("-" * 68)
    print(ranking[["posicao", "titulo", "categoria", "total_emprestimos", "disponivel"]].to_string(index=False))
    print("=" * 68)

    # Painel visual com 2 gráficos (Matplotlib + Seaborn)
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))

    # Gráfico 1: Ranking de Livros Mais Lidos
    sns.barplot(
        data=ranking,
        x="total_emprestimos",
        y="titulo",
        hue="categoria",
        dodge=False,
        ax=axes[0],
        palette="viridis",
    )
    axes[0].set_title("Indicador: Livros Mais Lidos (Total de Empréstimos)", fontsize=12, fontweight="bold")
    axes[0].set_xlabel("Quantidade de Empréstimos")
    axes[0].set_ylabel("Título do Livro")
    axes[0].legend(title="Categoria", loc="lower right")

    for container in axes[0].containers:
        axes[0].bar_label(container, padding=3, fontsize=9, fontweight="bold")

    # Gráfico 2: Empréstimos Acumulados por Categoria
    por_categoria = (
        df_livros.groupby("categoria", as_index=False)["total_emprestimos"]
        .sum()
        .sort_values(by="total_emprestimos", ascending=False)
    )
    sns.barplot(
        data=por_categoria,
        x="categoria",
        y="total_emprestimos",
        hue="categoria",
        legend=False,
        ax=axes[1],
        palette="crest",
    )
    axes[1].set_title("Volume de Leituras por Categoria", fontsize=12, fontweight="bold")
    axes[1].set_xlabel("Categoria")
    axes[1].set_ylabel("Total de Empréstimos")
    axes[1].tick_params(axis="x", rotation=20)

    for container in axes[1].containers:
        axes[1].bar_label(container, padding=3, fontsize=9, fontweight="bold")

    plt.tight_layout()
    if salvar_imagem:
        plt.savefig("indicadores_mais_lidos_av1.png", dpi=150, bbox_inches="tight")
    plt.show()


# ==============================================================================
# 4. EXECUÇÃO PRINCIPAL (FLUXO COMPLETO)
# ==============================================================================
if __name__ == "__main__":
    df_livros, df_usuarios, df_historico = inicializar_base_excel()

    # 1. Cadastros
    df_livros = cadastrar_livro(df_livros, "Estatística Prática para Cientistas de Dados", "Peter Bruce", "Estatística")
    df_usuarios = cadastrar_usuario(df_usuarios, "Lucas Mendes")

    # 2. Consultas de disponibilidade (passando df_livros!)
    consultar_disponibilidade(df_livros, "Clean Code")
    consultar_disponibilidade(df_livros, "Obra Inexistente")

    # 3. Empréstimos
    df_livros, df_historico = realizar_emprestimo(df_livros, df_usuarios, df_historico, id_livro=1, id_usuario=102)
    df_livros, df_historico = realizar_emprestimo(df_livros, df_usuarios, df_historico, id_livro=1, id_usuario=102)
    df_livros, df_historico = realizar_emprestimo(df_livros, df_usuarios, df_historico, id_livro=3, id_usuario=103)

    # 4. Devolução e novo empréstimo
    df_livros, df_historico = realizar_devolucao(df_livros, df_historico, id_livro=1)
    df_livros, df_historico = realizar_emprestimo(df_livros, df_usuarios, df_historico, id_livro=1, id_usuario=104)

    # 5. Persistência no Excel (.xlsx)
    salvar_dados_excel(df_livros, df_usuarios, df_historico)

    # 6. Indicadores e gráficos
    gerar_relatorio_e_graficos(df_livros)