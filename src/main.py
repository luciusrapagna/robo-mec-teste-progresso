import os
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches

ARQUIVO_ENTRADA = "data/raw/TS_TP.xlsx"

PASTA_PROCESSADOS = "data/processed"
PASTA_GRAFICOS = "outputs/graficos"
PASTA_RELATORIOS = "outputs/relatorios"

os.makedirs(PASTA_PROCESSADOS, exist_ok=True)
os.makedirs(PASTA_GRAFICOS, exist_ok=True)
os.makedirs(PASTA_RELATORIOS, exist_ok=True)


def carregar_dados():
    xls = pd.ExcelFile(ARQUIVO_ENTRADA)

    base = pd.read_excel(xls, "Base_completa")
    ranking_alunos = pd.read_excel(xls, "Ranking_alunos")
    ranking_periodos = pd.read_excel(xls, "Ranking_periodos")
    media_disciplina = pd.read_excel(xls, "Media_disciplina_geral")
    disciplina_periodo = pd.read_excel(xls, "Disciplina_periodo")
    alunos_risco = pd.read_excel(xls, "Alunos_em_risco")

    return base, ranking_alunos, ranking_periodos, media_disciplina, disciplina_periodo, alunos_risco


def salvar_tabelas(base, ranking_alunos, ranking_periodos, media_disciplina, disciplina_periodo, alunos_risco):
    caminho = f"{PASTA_PROCESSADOS}/tabelas_robo_mec.xlsx"

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        base.to_excel(writer, sheet_name="Base_completa", index=False)
        ranking_alunos.to_excel(writer, sheet_name="Ranking_alunos", index=False)
        ranking_periodos.to_excel(writer, sheet_name="Ranking_periodos", index=False)
        media_disciplina.to_excel(writer, sheet_name="Media_disciplina_geral", index=False)
        disciplina_periodo.to_excel(writer, sheet_name="Disciplina_periodo", index=False)
        alunos_risco.to_excel(writer, sheet_name="Alunos_em_risco", index=False)

    return caminho


def gerar_graficos(base, ranking_periodos, media_disciplina):
    caminhos = {}

    plt.figure(figsize=(10, 6))
    plt.hist(base["Percentual_acertos"], bins=15)
    plt.title("Distribuição geral do desempenho dos estudantes")
    plt.xlabel("Percentual de acertos")
    plt.ylabel("Número de estudantes")
    caminho = f"{PASTA_GRAFICOS}/histograma_desempenho.png"
    plt.savefig(caminho, dpi=300, bbox_inches="tight")
    plt.close()
    caminhos["histograma"] = caminho

    plt.figure(figsize=(12, 6))
    base.boxplot(column="Percentual_acertos", by="Período")
    plt.title("Comparação do desempenho por período")
    plt.suptitle("")
    plt.xlabel("Período")
    plt.ylabel("Percentual de acertos")
    caminho = f"{PASTA_GRAFICOS}/boxplot_periodos.png"
    plt.savefig(caminho, dpi=300, bbox_inches="tight")
    plt.close()
    caminhos["boxplot"] = caminho

    plt.figure(figsize=(10, 6))
    ranking_periodos_ordenado = ranking_periodos.sort_values("Media_percentual", ascending=True)
    plt.barh(ranking_periodos_ordenado["Período"], ranking_periodos_ordenado["Media_percentual"])
    plt.title("Ranking médio de desempenho por período")
    plt.xlabel("Média percentual de acertos")
    plt.ylabel("Período")
    caminho = f"{PASTA_GRAFICOS}/ranking_periodos.png"
    plt.savefig(caminho, dpi=300, bbox_inches="tight")
    plt.close()
    caminhos["ranking"] = caminho

    plt.figure(figsize=(10, 6))
    plt.bar(media_disciplina["Disciplina"], media_disciplina["Media_bruta"])
    plt.title("Média de acertos por disciplina")
    plt.xlabel("Disciplina")
    plt.ylabel("Média bruta de acertos")
    plt.xticks(rotation=30, ha="right")
    caminho = f"{PASTA_GRAFICOS}/media_disciplina.png"
    plt.savefig(caminho, dpi=300, bbox_inches="tight")
    plt.close()
    caminhos["disciplinas"] = caminho

    return caminhos


def texto_humanizado(ranking_periodos, alunos_risco, base):
    melhor = ranking_periodos.sort_values("Media_percentual", ascending=False).iloc[0]
    menor = ranking_periodos.sort_values("Media_percentual", ascending=True).iloc[0]

    total_alunos = len(base)
    total_risco = len(alunos_risco)
    percentual_risco = round((total_risco / total_alunos) * 100, 2)

    texto = f"""
A análise dos resultados do Teste de Progresso evidencia um panorama relevante para o acompanhamento pedagógico do curso de Medicina. Foram avaliados {total_alunos} estudantes, considerando o desempenho geral, o desempenho por período, a comparação entre turmas e a identificação de estudantes que necessitam de acompanhamento pedagógico mais próximo.

O período com melhor desempenho médio foi o {melhor['Período']}, com média de {melhor['Media_percentual']:.2f}% de acertos. Esse resultado sugere maior consolidação dos conhecimentos avaliados e pode servir como referência para a análise das práticas pedagógicas desenvolvidas ao longo da formação.

Por outro lado, o período com menor média foi o {menor['Período']}, com {menor['Media_percentual']:.2f}% de acertos. Esse achado não deve ser interpretado apenas como fragilidade discente, mas como indicador institucional para planejamento de ações de nivelamento, revisão de conteúdos essenciais e fortalecimento do acompanhamento docente.

Foram identificados {total_risco} estudantes com desempenho inferior ao ponto de atenção definido, correspondendo a {percentual_risco}% do total analisado. Esse grupo representa uma prioridade para ações de apoio pedagógico, monitorias, tutorias acadêmicas e estratégias de recuperação formativa.

De modo geral, os resultados reforçam a importância do uso sistemático de indicadores educacionais para subsidiar decisões do NDE, da coordenação do curso e dos docentes. A utilização desses dados contribui para uma cultura avaliativa baseada em evidências, favorecendo o acompanhamento longitudinal da aprendizagem e a melhoria contínua do curso, aspectos especialmente relevantes em processos de avaliação externa, como visitas do MEC.
"""
    return texto.strip()


def criar_relatorio_word(ranking_periodos, alunos_risco, base, caminhos_graficos):
    doc = Document()

    doc.add_heading("Relatório Pedagógico do Teste de Progresso", level=0)

    doc.add_paragraph(
        "Documento gerado automaticamente para subsidiar o acompanhamento pedagógico, "
        "a análise do desempenho discente e a organização de evidências institucionais "
        "para processos avaliativos, incluindo visita do MEC."
    )

    doc.add_heading("1. Síntese interpretativa dos resultados", level=1)
    doc.add_paragraph(texto_humanizado(ranking_periodos, alunos_risco, base))

    doc.add_heading("2. Distribuição geral do desempenho", level=1)
    doc.add_paragraph(
        "A Figura 1 apresenta a distribuição geral dos percentuais de acertos dos estudantes. "
        "Esse gráfico permite visualizar a concentração dos desempenhos e identificar possíveis "
        "grupos com maior necessidade de acompanhamento pedagógico."
    )
    doc.add_picture(caminhos_graficos["histograma"], width=Inches(6))
    doc.add_paragraph("Figura 1. Distribuição geral do percentual de acertos dos estudantes.")

    doc.add_heading("3. Comparação entre períodos", level=1)
    doc.add_paragraph(
        "A Figura 2 apresenta a comparação do desempenho entre os períodos. "
        "O boxplot permite observar a mediana, a dispersão dos resultados e possíveis valores extremos, "
        "favorecendo uma análise mais refinada das diferenças entre as turmas."
    )
    doc.add_picture(caminhos_graficos["boxplot"], width=Inches(6))
    doc.add_paragraph("Figura 2. Comparação do desempenho dos estudantes por período.")

    doc.add_heading("4. Ranking dos períodos", level=1)
    doc.add_paragraph(
        "A Figura 3 apresenta o ranking médio dos períodos. Esse resultado pode apoiar o planejamento "
        "de ações pedagógicas específicas, considerando os períodos com melhor desempenho e aqueles que "
        "demandam maior atenção institucional."
    )
    doc.add_picture(caminhos_graficos["ranking"], width=Inches(6))
    doc.add_paragraph("Figura 3. Ranking dos períodos segundo a média percentual de acertos.")

    doc.add_heading("5. Desempenho por disciplina", level=1)
    doc.add_paragraph(
        "A Figura 4 apresenta a média de acertos por disciplina. Essa análise permite identificar áreas "
        "com maior ou menor desempenho relativo, contribuindo para o planejamento docente e para a revisão "
        "de conteúdos no âmbito do curso."
    )
    doc.add_picture(caminhos_graficos["disciplinas"], width=Inches(6))
    doc.add_paragraph("Figura 4. Média de acertos por disciplina avaliada.")

    doc.add_heading("6. Encaminhamentos pedagógicos sugeridos", level=1)
    doc.add_paragraph(
        "Com base nos resultados, recomenda-se que a coordenação do curso, o NDE e os docentes utilizem "
        "os dados para planejar ações de acompanhamento discente, monitoria, revisão de conteúdos, oficinas "
        "de aprendizagem e devolutivas formativas por período e por disciplina."
    )

    doc.add_paragraph(
        "Para fins de visita do MEC, este relatório pode ser utilizado como evidência de gestão acadêmica "
        "baseada em dados, demonstrando acompanhamento sistemático da aprendizagem, identificação de fragilidades "
        "e proposição de melhorias contínuas."
    )

    caminho = f"{PASTA_RELATORIOS}/relatorio_pedagogico_mec.docx"
    doc.save(caminho)
    return caminho


def main():
    print("Carregando dados...")
    base, ranking_alunos, ranking_periodos, media_disciplina, disciplina_periodo, alunos_risco = carregar_dados()

    print("Salvando tabelas...")
    tabela_excel = salvar_tabelas(
        base,
        ranking_alunos,
        ranking_periodos,
        media_disciplina,
        disciplina_periodo,
        alunos_risco
    )

    print("Gerando gráficos...")
    caminhos_graficos = gerar_graficos(base, ranking_periodos, media_disciplina)

    print("Criando relatório humanizado em Word...")
    relatorio = criar_relatorio_word(
        ranking_periodos,
        alunos_risco,
        base,
        caminhos_graficos
    )

    print("\nROBÔ MEC FINALIZADO COM SUCESSO!")
    print(f"Tabelas geradas em: {tabela_excel}")
    print(f"Relatório Word gerado em: {relatorio}")
    print(f"Gráficos gerados em: {PASTA_GRAFICOS}")


if __name__ == "__main__":
    main()