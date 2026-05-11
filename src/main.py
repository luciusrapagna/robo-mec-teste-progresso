import os
import sys
import re
import json
import shutil
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from docx import Document
from docx.shared import Inches

from openpyxl import load_workbook
from openpyxl.styles import PatternFill


CAMINHO_PROJETO = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.append(CAMINHO_PROJETO)

try:
    from ia.humanizador import humanizar_relatorio
except Exception:
    def humanizar_relatorio(
        total,
        melhor_turma,
        melhor_media,
        pior_turma,
        pior_media,
        alunos_risco,
        percentual_risco,
    ):
        return (
            f"A análise dos resultados evidencia aspectos relevantes para o acompanhamento pedagógico. "
            f"Foram avaliados {total} estudantes. O melhor desempenho médio foi observado em {melhor_turma}, "
            f"com {melhor_media:.2f}%, enquanto o menor desempenho médio foi observado em {pior_turma}, "
            f"com {pior_media:.2f}%. Foram identificados {alunos_risco} estudantes em risco, "
            f"correspondendo a {percentual_risco:.2f}% do total."
        )


def carregar_config():
    caminho_config = os.path.join(CAMINHO_PROJETO, "config.json")

    if not os.path.exists(caminho_config):
        raise FileNotFoundError("Arquivo config.json não encontrado na raiz do projeto.")

    with open(caminho_config, "r", encoding="utf-8") as f:
        return json.load(f)


def criar_pasta_projeto(config):
    agora = datetime.now().strftime("%Y_%m_%d_%Hh%M")
    nome_avaliacao = config.get("nome_avaliacao", "TP").replace(" ", "_")
    pasta_base = os.path.join("projetos_gerados", f"{nome_avaliacao}_{agora}")

    pastas = {
        "base": pasta_base,
        "tabelas": os.path.join(pasta_base, "tabelas"),
        "graficos": os.path.join(pasta_base, "graficos"),
        "relatorios": os.path.join(pasta_base, "relatorios"),
        "dados": os.path.join(pasta_base, "dados_processados"),
        "original": os.path.join(pasta_base, "planilha_original"),
    }

    for pasta in pastas.values():
        os.makedirs(pasta, exist_ok=True)

    return pastas


def pedir_planilha(config):
    print(f"\nROBÔ MEC - {config.get('nome_avaliacao', 'Teste de Progresso')}")
    print("Selecione a planilha Excel com os resultados.\n")

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    caminho = askopenfilename(
        title="Selecione a planilha do Teste de Progresso",
        filetypes=[
            ("Arquivos Excel", "*.xlsx"),
            ("Arquivos Excel antigos", "*.xls"),
            ("Todos os arquivos", "*.*"),
        ],
    )

    root.destroy()

    if not caminho:
        raise Exception("Nenhuma planilha foi selecionada.")

    print(f"Planilha selecionada: {caminho}")
    return caminho


def carregar_dados(caminho_planilha):
    xls = pd.ExcelFile(caminho_planilha)
    abas = xls.sheet_names

    print("\nAbas encontradas:")
    for aba in abas:
        print(f"- {aba}")

    base = pd.read_excel(caminho_planilha, sheet_name=abas[0])
    return base


def detectar_colunas(base):
    coluna_periodo = None
    coluna_nome = None

    for col in base.columns:
        col_lower = str(col).lower()

        if "período" in col_lower or "periodo" in col_lower or "turma" in col_lower:
            coluna_periodo = col

        if "nome" in col_lower or "aluno" in col_lower or "estudante" in col_lower or "discente" in col_lower:
            coluna_nome = col

    colunas_numericas = base.select_dtypes(include=[np.number]).columns.tolist()

    if not colunas_numericas:
        raise ValueError("Não encontrei colunas numéricas com acertos ou notas.")

    coluna_total = colunas_numericas[-1]

    print("\nColunas detectadas:")
    print(f"Aluno: {coluna_nome}")
    print(f"Turma/Período: {coluna_periodo}")
    print(f"Coluna usada para desempenho geral: {coluna_total}")

    return coluna_nome, coluna_periodo, coluna_total


def preparar_base(base, coluna_nome, coluna_periodo, coluna_total, config):
    base = base.copy()

    if coluna_nome is None:
        base["Aluno"] = [f"Aluno_{i + 1}" for i in range(len(base))]
        coluna_nome = "Aluno"

    if coluna_periodo is None:
        base["Periodo"] = "Geral"
        coluna_periodo = "Periodo"

    numero_questoes = config.get("numero_questoes", 120)

    base["Total_acertos"] = pd.to_numeric(base[coluna_total], errors="coerce")
    base = base.dropna(subset=["Total_acertos"])

    maior_valor = base["Total_acertos"].max()

    if maior_valor <= 1:
        base["Percentual_acertos"] = base["Total_acertos"] * 100
    elif maior_valor <= 100:
        base["Percentual_acertos"] = base["Total_acertos"]
    else:
        base["Percentual_acertos"] = (base["Total_acertos"] / numero_questoes) * 100

    base["Media_turma"] = base.groupby(coluna_periodo)["Percentual_acertos"].transform("mean")
    base["Diferenca_media_turma"] = base["Percentual_acertos"] - base["Media_turma"]

    base["Percentual_abaixo_media_turma"] = np.where(
        base["Percentual_acertos"] < base["Media_turma"],
        ((base["Media_turma"] - base["Percentual_acertos"]) / base["Media_turma"]) * 100,
        0,
    )

    base["Percentual_acima_media_turma"] = np.where(
        base["Percentual_acertos"] > base["Media_turma"],
        ((base["Percentual_acertos"] - base["Media_turma"]) / base["Media_turma"]) * 100,
        0,
    )

    risco_config = config.get("criterios_risco", {})
    limite_pedagogico = risco_config.get("risco_pedagogico_ate_percentual_abaixo_media", 10)
    limite_critico = risco_config.get("risco_critico_a_partir_percentual_abaixo_media", 11)

    base["Nivel_risco"] = np.where(
        base["Percentual_abaixo_media_turma"] >= limite_critico,
        "Risco crítico",
        np.where(
            (base["Percentual_abaixo_media_turma"] > 0)
            & (base["Percentual_abaixo_media_turma"] <= limite_pedagogico),
            "Risco pedagógico",
            "Sem risco pela média da turma",
        ),
    )

    base["Classificacao"] = np.where(
        base["Percentual_acertos"] < 60,
        "Atenção pedagógica",
        "Desempenho adequado",
    )

    return base, coluna_nome, coluna_periodo


def gerar_pontos_extras(base, coluna_nome, coluna_periodo, config):
    criterios = config.get("criterios_pontos_extras", [])

    pontos = base.copy()

    pontos["Faixa_bonus"] = "Sem bônus"
    pontos["Ponto_extra_por_disciplina"] = 0.0
    pontos["Numero_disciplinas_bonus"] = 0
    pontos["Total_pontos_extras"] = 0.0

    for regra in criterios:
        minimo = regra.get("min_percentual_acima", 0)
        maximo = regra.get("max_percentual_acima", None)

        if maximo is None:
            mascara = pontos["Percentual_acima_media_turma"] >= minimo
        else:
            mascara = (
                (pontos["Percentual_acima_media_turma"] >= minimo)
                & (pontos["Percentual_acima_media_turma"] <= maximo)
            )

        pontos.loc[mascara, "Faixa_bonus"] = regra.get("faixa", "")
        pontos.loc[mascara, "Ponto_extra_por_disciplina"] = regra.get("ponto_por_disciplina", 0)
        pontos.loc[mascara, "Numero_disciplinas_bonus"] = regra.get("numero_disciplinas", 0)
        pontos.loc[mascara, "Total_pontos_extras"] = regra.get("total_pontos", 0)

    pontos_bonus = pontos[pontos["Faixa_bonus"] != "Sem bônus"].copy()

    colunas_saida = [
        coluna_nome,
        coluna_periodo,
        "Percentual_acertos",
        "Media_turma",
        "Percentual_acima_media_turma",
        "Faixa_bonus",
        "Ponto_extra_por_disciplina",
        "Numero_disciplinas_bonus",
        "Total_pontos_extras",
    ]

    return pontos_bonus[colunas_saida].sort_values(
        ["Total_pontos_extras", "Percentual_acima_media_turma"],
        ascending=False,
    )


def aplicar_cores_risco(caminho_excel):
    wb = load_workbook(caminho_excel)

    amarelo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    vermelho = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    verde = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")

    for aba in wb.sheetnames:
        ws = wb[aba]
        cabecalhos = [cell.value for cell in ws[1]]

        if "Nivel_risco" in cabecalhos:
            coluna_risco = cabecalhos.index("Nivel_risco") + 1

            for linha in range(2, ws.max_row + 1):
                valor = ws.cell(row=linha, column=coluna_risco).value

                if valor == "Risco pedagógico":
                    for col in range(1, ws.max_column + 1):
                        ws.cell(row=linha, column=col).fill = amarelo

                elif valor == "Risco crítico":
                    for col in range(1, ws.max_column + 1):
                        ws.cell(row=linha, column=col).fill = vermelho

        if aba == "Pontos_extras":
            for linha in range(2, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=linha, column=col).fill = verde

    wb.save(caminho_excel)


import os
import sys
import re
import json
import shutil
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from docx import Document
from docx.shared import Inches

from openpyxl import load_workbook
from openpyxl.styles import PatternFill


CAMINHO_PROJETO = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.append(CAMINHO_PROJETO)


def carregar_config():
    caminho = os.path.join(CAMINHO_PROJETO, "config.json")

    if not os.path.exists(caminho):
        raise FileNotFoundError("Arquivo config.json não encontrado na raiz do projeto.")

    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)


def criar_pasta_projeto(config):
    agora = datetime.now().strftime("%Y_%m_%d_%Hh%M")
    nome_avaliacao = config.get("nome_avaliacao", "TP").replace(" ", "_")

    pasta_base = os.path.join("projetos_gerados", f"{nome_avaliacao}_{agora}")

    pastas = {
        "base": pasta_base,
        "tabelas": os.path.join(pasta_base, "tabelas"),
        "graficos": os.path.join(pasta_base, "graficos"),
        "relatorios": os.path.join(pasta_base, "relatorios"),
        "dados": os.path.join(pasta_base, "dados_processados"),
        "original": os.path.join(pasta_base, "planilha_original")
    }

    for pasta in pastas.values():
        os.makedirs(pasta, exist_ok=True)

    return pastas


def pedir_planilha(config):
    print(f"\nROBÔ MEC - {config.get('nome_avaliacao', 'Teste de Progresso')}")
    print("Selecione a planilha Excel com os resultados.\n")

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    caminho = askopenfilename(
        title="Selecione a planilha do Teste de Progresso",
        filetypes=[
            ("Arquivos Excel", "*.xlsx"),
            ("Arquivos Excel antigos", "*.xls"),
            ("Todos os arquivos", "*.*")
        ]
    )

    root.destroy()

    if not caminho:
        raise Exception("Nenhuma planilha foi selecionada.")

    print(f"Planilha selecionada: {caminho}")
    return caminho


def carregar_dados(caminho_planilha):
    xls = pd.ExcelFile(caminho_planilha)
    abas = xls.sheet_names

    print("\nAbas encontradas:")
    for aba in abas:
        print(f"- {aba}")

    return pd.read_excel(caminho_planilha, sheet_name=abas[0])


def detectar_colunas(base):
    coluna_nome = None
    coluna_periodo = None

    for col in base.columns:
        col_lower = str(col).lower()

        if "nome" in col_lower or "aluno" in col_lower or "estudante" in col_lower or "discente" in col_lower:
            coluna_nome = col

        if "período" in col_lower or "periodo" in col_lower or "turma" in col_lower:
            coluna_periodo = col

    colunas_numericas = base.select_dtypes(include=[np.number]).columns.tolist()

    if not colunas_numericas:
        raise ValueError("Não encontrei colunas numéricas com acertos ou notas.")

    coluna_total = colunas_numericas[-1]

    print("\nColunas detectadas:")
    print(f"Aluno: {coluna_nome}")
    print(f"Turma/Período: {coluna_periodo}")
    print(f"Coluna usada para desempenho geral: {coluna_total}")

    return coluna_nome, coluna_periodo, coluna_total


def preparar_base(base, coluna_nome, coluna_periodo, coluna_total, config):
    base = base.copy()

    if coluna_nome is None:
        base["Aluno"] = [f"Aluno_{i + 1}" for i in range(len(base))]
        coluna_nome = "Aluno"

    if coluna_periodo is None:
        base["Periodo"] = "Geral"
        coluna_periodo = "Periodo"

    numero_questoes = config.get("numero_questoes", 120)

    base["Total_acertos"] = pd.to_numeric(base[coluna_total], errors="coerce")
    base = base.dropna(subset=["Total_acertos"])

    maior_valor = base["Total_acertos"].max()

    if maior_valor <= 1:
        base["Percentual_acertos"] = base["Total_acertos"] * 100
    elif maior_valor <= 100:
        base["Percentual_acertos"] = base["Total_acertos"]
    else:
        base["Percentual_acertos"] = (base["Total_acertos"] / numero_questoes) * 100

    base["Media_turma"] = base.groupby(coluna_periodo)["Percentual_acertos"].transform("mean")
    base["Diferenca_media_turma"] = base["Percentual_acertos"] - base["Media_turma"]

    base["Percentual_abaixo_media_turma"] = np.where(
        base["Percentual_acertos"] < base["Media_turma"],
        ((base["Media_turma"] - base["Percentual_acertos"]) / base["Media_turma"]) * 100,
        0
    )

    base["Percentual_acima_media_turma"] = np.where(
        base["Percentual_acertos"] > base["Media_turma"],
        ((base["Percentual_acertos"] - base["Media_turma"]) / base["Media_turma"]) * 100,
        0
    )

    risco = config.get("criterios_risco", {})
    limite_ped = risco.get("risco_pedagogico_ate_percentual_abaixo_media", 10)
    limite_crit = risco.get("risco_critico_a_partir_percentual_abaixo_media", 11)

    base["Nivel_risco"] = np.where(
        base["Percentual_abaixo_media_turma"] >= limite_crit,
        "Risco crítico",
        np.where(
            (base["Percentual_abaixo_media_turma"] > 0) &
            (base["Percentual_abaixo_media_turma"] <= limite_ped),
            "Risco pedagógico",
            "Sem risco pela média da turma"
        )
    )

    base["Classificacao"] = np.where(
        base["Percentual_acertos"] < 60,
        "Atenção pedagógica",
        "Desempenho adequado"
    )

    return base, coluna_nome, coluna_periodo


def extrair_numero_periodo(valor):
    texto = str(valor)
    numeros = re.findall(r"\d+", texto)

    if not numeros:
        return None

    return numeros[0]


def classificar_ciclo(valor_periodo, config):
    numero = extrair_numero_periodo(valor_periodo)

    ciclos = config.get("ciclos_formativos", {})

    for nome_ciclo, periodos in ciclos.items():
        if numero in [str(p) for p in periodos]:
            return nome_ciclo

    return "Não classificado"


def gerar_tabela_ciclos(base, coluna_periodo, config):
    temp = base.copy()
    temp["Ciclo_formativo"] = temp[coluna_periodo].apply(lambda x: classificar_ciclo(x, config))

    tabela = (
        temp.groupby("Ciclo_formativo")
        .agg(
            Numero_estudantes=("Percentual_acertos", "count"),
            Media_percentual=("Percentual_acertos", "mean"),
            Mediana_percentual=("Percentual_acertos", "median"),
            Desvio_padrao=("Percentual_acertos", "std"),
            Minimo=("Percentual_acertos", "min"),
            Maximo=("Percentual_acertos", "max")
        )
        .reset_index()
        .sort_values("Media_percentual", ascending=True)
    )

    return tabela


def gerar_pontos_extras(base, coluna_nome, coluna_periodo, config):
    pontos = base.copy()

    pontos["Faixa_bonus"] = "Sem bônus"
    pontos["Ponto_extra_por_disciplina"] = 0.0
    pontos["Numero_disciplinas_bonus"] = 0
    pontos["Total_pontos_extras"] = 0.0

    for regra in config.get("criterios_pontos_extras", []):
        minimo = regra.get("min_percentual_acima", 0)
        maximo = regra.get("max_percentual_acima", None)

        if maximo is None:
            mascara = pontos["Percentual_acima_media_turma"] >= minimo
        else:
            mascara = (
                (pontos["Percentual_acima_media_turma"] >= minimo) &
                (pontos["Percentual_acima_media_turma"] <= maximo)
            )

        pontos.loc[mascara, "Faixa_bonus"] = regra.get("faixa", "")
        pontos.loc[mascara, "Ponto_extra_por_disciplina"] = regra.get("ponto_por_disciplina", 0)
        pontos.loc[mascara, "Numero_disciplinas_bonus"] = regra.get("numero_disciplinas", 0)
        pontos.loc[mascara, "Total_pontos_extras"] = regra.get("total_pontos", 0)

    pontos_bonus = pontos[pontos["Faixa_bonus"] != "Sem bônus"].copy()

    colunas = [
        coluna_nome,
        coluna_periodo,
        "Percentual_acertos",
        "Media_turma",
        "Percentual_acima_media_turma",
        "Faixa_bonus",
        "Ponto_extra_por_disciplina",
        "Numero_disciplinas_bonus",
        "Total_pontos_extras"
    ]

    return pontos_bonus[colunas].sort_values(
        ["Total_pontos_extras", "Percentual_acima_media_turma"],
        ascending=False
    )


def identificar_grandes_areas(base, config):
    areas = config.get("grandes_areas", {})
    colunas_areas = {}

    for nome_area, termos in areas.items():
        for coluna in base.columns:
            coluna_lower = str(coluna).lower()

            if any(str(termo).lower() in coluna_lower for termo in termos):
                if pd.api.types.is_numeric_dtype(base[coluna]):
                    colunas_areas[nome_area] = coluna
                break

    return colunas_areas


def gerar_tabela_grandes_areas(base, coluna_periodo, colunas_areas):
    tabelas = {}

    for nome_area, coluna_area in colunas_areas.items():
        temp = base.copy()
        temp[coluna_area] = pd.to_numeric(temp[coluna_area], errors="coerce")

        max_area = temp[coluna_area].max()

        if max_area <= 1:
            temp["Percentual_area"] = temp[coluna_area] * 100
        elif max_area <= 100:
            temp["Percentual_area"] = temp[coluna_area]
        else:
            temp["Percentual_area"] = (temp[coluna_area] / max_area) * 100

        tabela = (
            temp.groupby(coluna_periodo)
            .agg(
                Numero_estudantes=(coluna_area, "count"),
                Media_acertos=(coluna_area, "mean"),
                Percentual_medio=("Percentual_area", "mean"),
                Desvio_padrao=("Percentual_area", "std")
            )
            .reset_index()
            .sort_values("Percentual_medio", ascending=False)
        )

        tabelas[nome_area] = tabela

    return tabelas


def aplicar_cores_excel(caminho_excel):
    wb = load_workbook(caminho_excel)

    amarelo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    vermelho = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    verde = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")

    for aba in wb.sheetnames:
        ws = wb[aba]
        cabecalhos = [cell.value for cell in ws[1]]

        if "Nivel_risco" in cabecalhos:
            coluna_risco = cabecalhos.index("Nivel_risco") + 1

            for linha in range(2, ws.max_row + 1):
                valor = ws.cell(row=linha, column=coluna_risco).value

                if valor == "Risco pedagógico":
                    for col in range(1, ws.max_column + 1):
                        ws.cell(row=linha, column=col).fill = amarelo

                elif valor == "Risco crítico":
                    for col in range(1, ws.max_column + 1):
                        ws.cell(row=linha, column=col).fill = vermelho

        if aba == "Pontos_extras":
            for linha in range(2, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=linha, column=col).fill = verde

    wb.save(caminho_excel)


def gerar_tabelas(base, coluna_nome, coluna_periodo, pastas, config):
    ranking_alunos = base.sort_values("Percentual_acertos", ascending=False)

    ranking_periodos = (
        base.groupby(coluna_periodo)
        .agg(
            N_estudantes=(coluna_nome, "count"),
            Media_percentual=("Percentual_acertos", "mean"),
            Mediana_percentual=("Percentual_acertos", "median"),
            Desvio_padrao=("Percentual_acertos", "std"),
            Minimo=("Percentual_acertos", "min"),
            Maximo=("Percentual_acertos", "max")
        )
        .reset_index()
        .sort_values("Media_percentual", ascending=False)
    )

    tabela_ciclos = gerar_tabela_ciclos(base, coluna_periodo, config)

    colunas_areas = identificar_grandes_areas(base, config)
    tabelas_areas = gerar_tabela_grandes_areas(base, coluna_periodo, colunas_areas)

    alunos_risco = base[
        base["Nivel_risco"].isin(["Risco pedagógico", "Risco crítico"])
    ].copy()

    pontos_extras = gerar_pontos_extras(base, coluna_nome, coluna_periodo, config)

    caminho_excel = os.path.join(pastas["tabelas"], "tabelas_resultados_mec.xlsx")

    with pd.ExcelWriter(caminho_excel, engine="openpyxl") as writer:
        base.to_excel(writer, sheet_name="Base_processada", index=False)
        ranking_alunos.to_excel(writer, sheet_name="Ranking_alunos", index=False)
        ranking_periodos.to_excel(writer, sheet_name="Ranking_periodos", index=False)
        tabela_ciclos.to_excel(writer, sheet_name="Ciclos_formativos", index=False)
        alunos_risco.to_excel(writer, sheet_name="Alunos_em_risco", index=False)
        pontos_extras.to_excel(writer, sheet_name="Pontos_extras", index=False)

        for nome_area, tabela_area in tabelas_areas.items():
            tabela_area.to_excel(writer, sheet_name=nome_area[:28], index=False)

    aplicar_cores_excel(caminho_excel)

    base.to_csv(
        os.path.join(pastas["dados"], "base_processada.csv"),
        sep=";",
        decimal=",",
        index=False,
        encoding="utf-8-sig"
    )

    return (
        ranking_alunos,
        ranking_periodos,
        tabela_ciclos,
        alunos_risco,
        pontos_extras,
        caminho_excel,
        tabelas_areas
    )


def configurar_estilo_grafico():
    plt.rcParams["font.family"] = "Arial"
    plt.rcParams["axes.spines.top"] = False
    plt.rcParams["axes.spines.right"] = False


def salvar_grafico(nome, pastas):
    png = os.path.join(pastas["graficos"], f"{nome}.png")
    svg = os.path.join(pastas["graficos"], f"{nome}.svg")

    plt.savefig(png, dpi=300, bbox_inches="tight")
    plt.savefig(svg, bbox_inches="tight")
    plt.close()

    return png


def nome_arquivo_seguro(texto):
    texto = texto.lower()
    texto = texto.replace("á", "a").replace("ã", "a").replace("â", "a").replace("à", "a")
    texto = texto.replace("é", "e").replace("ê", "e")
    texto = texto.replace("í", "i")
    texto = texto.replace("ó", "o").replace("õ", "o").replace("ô", "o")
    texto = texto.replace("ú", "u")
    texto = texto.replace("ç", "c")
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    return texto.strip("_")


def gerar_graficos(base, ranking_periodos, tabela_ciclos, coluna_periodo, pastas):
    configurar_estilo_grafico()

    graficos = {}

    tabela = ranking_periodos.sort_values("Media_percentual", ascending=True)

    plt.figure(figsize=(12, 7))
    barras = plt.bar(
        tabela[coluna_periodo].astype(str),
        tabela["Media_percentual"],
        color="#28a9d6"
    )

    plt.plot(
        tabela[coluna_periodo].astype(str),
        tabela["Media_percentual"],
        color="#222222",
        marker="o",
        linewidth=3
    )

    for barra in barras:
        valor = barra.get_height()
        plt.text(
            barra.get_x() + barra.get_width() / 2,
            valor + 1,
            f"{valor:.1f}%",
            ha="center",
            fontsize=12,
            fontweight="bold"
        )

    plt.title("Desempenho médio por turma/período", fontsize=18, fontweight="bold")
    plt.xlabel("Turma/Período")
    plt.ylabel("Média de acertos (%)")
    plt.grid(axis="y", alpha=0.2)
    plt.ylim(0, max(tabela["Media_percentual"]) + 15)
    graficos["ranking"] = salvar_grafico("ranking_periodos_visual", pastas)

    plt.figure(figsize=(11, 6))
    plt.hist(
        base["Percentual_acertos"],
        bins=15,
        color="#28a9d6",
        edgecolor="white"
    )

    media = base["Percentual_acertos"].mean()
    plt.axvline(media, color="#ec174c", linestyle="--", linewidth=3)

    plt.title("Distribuição geral do desempenho dos estudantes", fontsize=18, fontweight="bold")
    plt.xlabel("Percentual de acertos")
    plt.ylabel("Número de estudantes")
    plt.grid(axis="y", alpha=0.2)
    graficos["histograma"] = salvar_grafico("histograma_desempenho_visual", pastas)

    plt.figure(figsize=(12, 6))
    base.boxplot(column="Percentual_acertos", by=coluna_periodo, grid=False)
    plt.title("Comparação do desempenho por turma/período", fontsize=18, fontweight="bold")
    plt.suptitle("")
    plt.xlabel("Turma/Período")
    plt.ylabel("Percentual de acertos")
    plt.grid(axis="y", alpha=0.2)
    graficos["boxplot"] = salvar_grafico("boxplot_periodos_visual", pastas)

    if not tabela_ciclos.empty:
        ciclo_plot = tabela_ciclos.sort_values("Media_percentual", ascending=True)

        plt.figure(figsize=(10, 6))
        barras = plt.bar(
            ciclo_plot["Ciclo_formativo"].astype(str),
            ciclo_plot["Media_percentual"],
            color="#36d6b0"
        )

        for barra in barras:
            valor = barra.get_height()
            plt.text(
                barra.get_x() + barra.get_width() / 2,
                valor + 1,
                f"{valor:.1f}%",
                ha="center",
                fontsize=12,
                fontweight="bold"
            )

        plt.title("Desempenho médio por ciclo formativo", fontsize=18, fontweight="bold")
        plt.xlabel("Ciclo formativo")
        plt.ylabel("Média de acertos (%)")
        plt.grid(axis="y", alpha=0.2)
        plt.ylim(0, max(ciclo_plot["Media_percentual"]) + 15)
        graficos["ciclos"] = salvar_grafico("ciclos_formativos_visual", pastas)

    return graficos


def gerar_grafico_area(nome_area, tabela_area, coluna_periodo, pastas):
    configurar_estilo_grafico()

    nome_base = nome_arquivo_seguro(nome_area)
    tabela = tabela_area.sort_values("Percentual_medio", ascending=True)

    plt.figure(figsize=(12, 7))
    barras = plt.bar(
        tabela[coluna_periodo].astype(str),
        tabela["Percentual_medio"],
        color="#28a9d6"
    )

    plt.plot(
        tabela[coluna_periodo].astype(str),
        tabela["Percentual_medio"],
        color="#222222",
        marker="o",
        linewidth=3
    )

    for barra in barras:
        valor = barra.get_height()
        plt.text(
            barra.get_x() + barra.get_width() / 2,
            valor + 1,
            f"{valor:.1f}%",
            ha="center",
            fontsize=12,
            fontweight="bold"
        )

    plt.title(f"Percentual médio em {nome_area} por turma/período", fontsize=16, fontweight="bold")
    plt.xlabel("Turma/Período")
    plt.ylabel("Percentual médio (%)")
    plt.grid(axis="y", alpha=0.2)
    plt.ylim(0, max(tabela["Percentual_medio"]) + 15)

    caminho_percentual = os.path.join(pastas["graficos"], f"area_{nome_base}_percentual.png")
    plt.savefig(caminho_percentual, dpi=300, bbox_inches="tight")
    plt.close()

    plt.figure(figsize=(12, 7))
    barras = plt.bar(
        tabela[coluna_periodo].astype(str),
        tabela["Media_acertos"],
        color="#36d6b0"
    )

    for barra in barras:
        valor = barra.get_height()
        plt.text(
            barra.get_x() + barra.get_width() / 2,
            valor + 0.5,
            f"{valor:.1f}",
            ha="center",
            fontsize=12,
            fontweight="bold"
        )

    plt.title(f"Média de acertos em {nome_area} por turma/período", fontsize=16, fontweight="bold")
    plt.xlabel("Turma/Período")
    plt.ylabel("Média de acertos")
    plt.grid(axis="y", alpha=0.2)

    caminho_media = os.path.join(pastas["graficos"], f"area_{nome_base}_media_acertos.png")
    plt.savefig(caminho_media, dpi=300, bbox_inches="tight")
    plt.close()

    return caminho_percentual, caminho_media


def adicionar_titulo_tabela(doc, numero, titulo):
    p = doc.add_paragraph()
    run = p.add_run(f"Tabela {numero} – {titulo}")
    run.bold = True


def adicionar_fonte(doc, config):
    fonte = config.get("relatorio", {}).get(
        "fonte_tabelas_figuras",
        "Fonte: Elaboração própria a partir dos resultados do Teste de Progresso."
    )

    p = doc.add_paragraph()
    run = p.add_run(fonte)
    run.italic = True


def adicionar_legenda_figura(doc, texto, config):
    doc.add_paragraph(texto)
    adicionar_fonte(doc, config)


def criar_relatorio_word(
    base,
    ranking_periodos,
    tabela_ciclos,
    alunos_risco,
    pontos_extras,
    graficos,
    coluna_nome,
    coluna_periodo,
    pastas,
    tabelas_areas,
    config
):
    total = len(base)
    total_risco = len(alunos_risco)
    percentual_risco = (total_risco / total) * 100 if total > 0 else 0

    melhor = ranking_periodos.iloc[0]
    pior = ranking_periodos.iloc[-1]

    nome_instituicao = config.get("nome_instituicao", "")
    nome_curso = config.get("nome_curso", "")
    nome_avaliacao = config.get("nome_avaliacao", "Teste de Progresso")
    titulo = config.get("relatorio", {}).get("titulo", "Relatório Pedagógico do Teste de Progresso")
    finalidade = config.get("relatorio", {}).get("finalidade", "")

    estilo = config.get("estilo_mec", {})

    doc = Document()
    doc.add_heading(titulo, level=0)

    doc.add_paragraph(f"Instituição: {nome_instituicao}")
    doc.add_paragraph(f"Curso: {nome_curso}")
    doc.add_paragraph(f"Instrumento avaliativo: {nome_avaliacao}")

    doc.add_paragraph(
        f"O {nome_avaliacao} é utilizado como ferramenta de avaliação cognitiva longitudinal, "
        f"permitindo acompanhar a progressão do estudante ao longo da formação médica. "
        f"Neste relatório, os resultados são apresentados de forma didática e analítica, com foco "
        f"na evolução entre períodos, nos ciclos formativos, nas grandes áreas do conhecimento, "
        f"nos estudantes que demandam acompanhamento pedagógico e nos estudantes elegíveis à bonificação "
        f"acadêmica por desempenho acima da média da própria turma. {finalidade}"
    )

    doc.add_heading("1. Leitura pedagógica geral dos resultados", level=1)

    doc.add_paragraph(
        f"Foram avaliados {total} estudantes. O maior desempenho médio foi observado em "
        f"{melhor[coluna_periodo]}, com média de {melhor['Media_percentual']:.2f}% de acertos. "
        f"O menor desempenho médio foi observado em {pior[coluna_periodo]}, com média de "
        f"{pior['Media_percentual']:.2f}% de acertos."
    )

    doc.add_paragraph(
        estilo.get(
            "progressao",
            "Observa-se progressão longitudinal do desempenho médio ao longo dos períodos do curso."
        )
    )

    doc.add_paragraph(
        estilo.get(
            "competencias",
            "Os resultados sugerem consolidação progressiva das competências cognitivas esperadas para cada etapa formativa."
        )
    )

    doc.add_paragraph(
        "Para a gestão acadêmica, esses dados permitem identificar se há coerência entre o avanço do estudante "
        "no curso e o crescimento esperado no domínio cognitivo. Essa leitura auxilia a coordenação, o NDE "
        "e os docentes na definição de ações de reforço, revisão de conteúdos e acompanhamento longitudinal."
    )

    doc.add_heading("2. Distribuição geral do desempenho", level=1)

    doc.add_paragraph(
        "A distribuição geral do desempenho permite visualizar como os estudantes se posicionam em relação "
        "à média geral do teste. Quando há maior concentração em torno da média, observa-se maior homogeneidade "
        "do grupo. Quando há maior dispersão, evidencia-se a necessidade de acompanhamento mais individualizado."
    )

    doc.add_picture(graficos["histograma"], width=Inches(6))
    adicionar_legenda_figura(doc, "Figura 1 – Distribuição geral do percentual de acertos dos estudantes.", config)

    doc.add_heading("3. Progressão por turma/período", level=1)

    doc.add_paragraph(
        "A comparação entre turmas/períodos permite verificar a curva de crescimento esperada no Teste de Progresso. "
        "Em uma avaliação longitudinal, espera-se que os estudantes apresentem aumento gradual das taxas de acerto "
        "conforme avançam no curso e acumulam maior contato com conteúdos clínicos, práticas supervisionadas e "
        "experiências formativas."
    )

    doc.add_picture(graficos["ranking"], width=Inches(6))
    adicionar_legenda_figura(doc, "Figura 2 – Ranking dos períodos/turmas segundo a média percentual de acertos.", config)

    doc.add_picture(graficos["boxplot"], width=Inches(6))
    adicionar_legenda_figura(doc, "Figura 3 – Distribuição do desempenho por turma/período.", config)

    adicionar_titulo_tabela(doc, 1, "Síntese estatística do desempenho por turma/período")

    tabela_doc = doc.add_table(rows=1, cols=6)
    tabela_doc.style = "Table Grid"

    cab = tabela_doc.rows[0].cells
    cab[0].text = "Turma/Período"
    cab[1].text = "N"
    cab[2].text = "Média (%)"
    cab[3].text = "Mediana (%)"
    cab[4].text = "Desvio-padrão"
    cab[5].text = "Máximo (%)"

    for _, row in ranking_periodos.iterrows():
        cells = tabela_doc.add_row().cells
        cells[0].text = str(row[coluna_periodo])
        cells[1].text = str(int(row["N_estudantes"]))
        cells[2].text = f"{row['Media_percentual']:.2f}"
        cells[3].text = f"{row['Mediana_percentual']:.2f}"
        cells[4].text = f"{row['Desvio_padrao']:.2f}"
        cells[5].text = f"{row['Maximo']:.2f}"

    adicionar_fonte(doc, config)

    doc.add_paragraph(
        "A Tabela 1 demonstra a distribuição do desempenho por turma/período. A média indica o desempenho "
        "global do grupo, enquanto o desvio-padrão permite observar a variabilidade interna. Turmas com maior "
        "desvio-padrão podem demandar estratégias de acompanhamento mais individualizadas."
    )

    doc.add_heading("4. Resultados por ciclos formativos", level=1)

    doc.add_paragraph(
        "A análise por ciclos formativos permite organizar os resultados de acordo com a estrutura pedagógica "
        "do curso. O Ciclo Básico tende a refletir a consolidação dos fundamentos biomédicos e sociais; o Ciclo "
        "Clínico evidencia a ampliação da aplicação do conhecimento em contextos de cuidado; e o Internato expressa "
        "a integração das competências cognitivas com a prática clínica supervisionada."
    )

    if "ciclos" in graficos:
        doc.add_picture(graficos["ciclos"], width=Inches(6))
        adicionar_legenda_figura(doc, "Figura 4 – Desempenho médio por ciclo formativo.", config)

    adicionar_titulo_tabela(doc, 2, "Síntese do desempenho por ciclo formativo")

    tabela_ciclo_doc = doc.add_table(rows=1, cols=6)
    tabela_ciclo_doc.style = "Table Grid"

    cab = tabela_ciclo_doc.rows[0].cells
    cab[0].text = "Ciclo"
    cab[1].text = "N"
    cab[2].text = "Média (%)"
    cab[3].text = "Mediana (%)"
    cab[4].text = "Desvio-padrão"
    cab[5].text = "Máximo (%)"

    for _, row in tabela_ciclos.iterrows():
        cells = tabela_ciclo_doc.add_row().cells
        cells[0].text = str(row["Ciclo_formativo"])
        cells[1].text = str(int(row["Numero_estudantes"]))
        cells[2].text = f"{row['Media_percentual']:.2f}"
        cells[3].text = f"{row['Mediana_percentual']:.2f}"
        cells[4].text = f"{row['Desvio_padrao']:.2f}"
        cells[5].text = f"{row['Maximo']:.2f}"

    adicionar_fonte(doc, config)

    doc.add_paragraph(
        "Essa leitura por ciclos permite que o NDE acompanhe se a progressão cognitiva observada está coerente "
        "com a matriz curricular e com as competências previstas no PPC."
    )

    doc.add_heading("5. Resultados por grandes áreas do conhecimento", level=1)

    doc.add_paragraph(
        "A análise por grandes áreas do conhecimento permite identificar em quais eixos os estudantes apresentam "
        "maior consolidação do aprendizado e quais áreas representam oportunidades de reforço pedagógico. Essa "
        "leitura é fundamental para orientar planejamento docente, revisão de conteúdos e integração curricular."
    )

    numero_tabela = 3

    if not tabelas_areas:
        doc.add_paragraph("Não foram identificadas automaticamente colunas referentes às grandes áreas do conhecimento.")
    else:
        for nome_area, tabela_area in tabelas_areas.items():
            grafico_percentual, grafico_media = gerar_grafico_area(
                nome_area,
                tabela_area,
                coluna_periodo,
                pastas
            )

            melhor_area = tabela_area.sort_values("Percentual_medio", ascending=False).iloc[0]
            menor_area = tabela_area.sort_values("Percentual_medio", ascending=True).iloc[0]

            doc.add_heading(nome_area, level=2)

            doc.add_paragraph(
                f"No eixo de {nome_area}, a taxa média de acerto variou de "
                f"{menor_area['Percentual_medio']:.2f}% a {melhor_area['Percentual_medio']:.2f}% "
                f"entre as turmas/períodos avaliados. O maior desempenho foi observado em "
                f"{melhor_area[coluna_periodo]}, com média de {melhor_area['Media_acertos']:.2f} acertos. "
                f"O menor desempenho foi observado em {menor_area[coluna_periodo]}, com média de "
                f"{menor_area['Media_acertos']:.2f} acertos."
            )

            doc.add_paragraph(
                f"Do ponto de vista da gestão pedagógica, os resultados em {nome_area} permitem identificar "
                f"o alcance das competências relacionadas a esse eixo e sinalizar oportunidades de reforço "
                f"quando houver menor taxa de acerto em determinados períodos."
            )

            doc.add_picture(grafico_percentual, width=Inches(6))
            adicionar_legenda_figura(doc, f"Figura – Percentual médio em {nome_area} por turma/período.", config)

            doc.add_picture(grafico_media, width=Inches(6))
            adicionar_legenda_figura(doc, f"Figura – Média de acertos em {nome_area} por turma/período.", config)

            adicionar_titulo_tabela(
                doc,
                numero_tabela,
                f"Número de estudantes, média de acertos e percentual médio em {nome_area}"
            )

            tabela_area_doc = doc.add_table(rows=1, cols=5)
            tabela_area_doc.style = "Table Grid"

            cab = tabela_area_doc.rows[0].cells
            cab[0].text = "Turma/Período"
            cab[1].text = "Número"
            cab[2].text = "Média de acertos"
            cab[3].text = "Percentual médio (%)"
            cab[4].text = "DP (%)"

            for _, row in tabela_area.sort_values("Percentual_medio", ascending=False).iterrows():
                cells = tabela_area_doc.add_row().cells
                cells[0].text = str(row[coluna_periodo])
                cells[1].text = str(int(row["Numero_estudantes"]))
                cells[2].text = f"{row['Media_acertos']:.2f}"
                cells[3].text = f"{row['Percentual_medio']:.2f}"
                cells[4].text = f"{row['Desvio_padrao']:.2f}"

            adicionar_fonte(doc, config)
            numero_tabela += 1

    doc.add_heading("6. Estudantes em risco pedagógico", level=1)

    risco_pedagogico = alunos_risco[alunos_risco["Nivel_risco"] == "Risco pedagógico"]
    risco_critico = alunos_risco[alunos_risco["Nivel_risco"] == "Risco crítico"]

    doc.add_paragraph(
        f"Foram identificados {len(risco_pedagogico)} estudantes com desempenho até 10% abaixo da média "
        f"da própria turma e {len(risco_critico)} estudantes com desempenho a partir de 11% abaixo da média "
        f"da própria turma."
    )

    doc.add_paragraph(
        "Esses dados orientam ações de acompanhamento formativo, como devolutivas individualizadas, monitorias, "
        "revisão de conteúdos e encaminhamento para apoio psicopedagógico quando necessário."
    )

    adicionar_titulo_tabela(doc, numero_tabela, "Estudantes classificados em risco pedagógico ou crítico")

    tabela_risco = doc.add_table(rows=1, cols=5)
    tabela_risco.style = "Table Grid"

    cab = tabela_risco.rows[0].cells
    cab[0].text = "Aluno"
    cab[1].text = "Turma/Período"
    cab[2].text = "% de acertos"
    cab[3].text = "% abaixo da média da turma"
    cab[4].text = "Classificação"

    for _, row in alunos_risco.sort_values("Percentual_abaixo_media_turma", ascending=False).iterrows():
        cells = tabela_risco.add_row().cells
        cells[0].text = str(row[coluna_nome])
        cells[1].text = str(row[coluna_periodo])
        cells[2].text = f"{row['Percentual_acertos']:.2f}%"
        cells[3].text = f"{row['Percentual_abaixo_media_turma']:.2f}%"
        cells[4].text = str(row["Nivel_risco"])

    adicionar_fonte(doc, config)
    numero_tabela += 1

    doc.add_heading("7. Bonificação acadêmica por desempenho acima da média da turma", level=1)

    doc.add_paragraph(
        "A bonificação acadêmica foi calculada a partir da comparação entre o desempenho individual e a média "
        "da própria turma. Essa estratégia reconhece estudantes com desempenho acima do grupo de referência, "
        "mantendo o critério contextualizado por turma/período."
    )

    adicionar_titulo_tabela(doc, numero_tabela, "Estudantes elegíveis à bonificação acadêmica")

    tabela_bonus = doc.add_table(rows=1, cols=7)
    tabela_bonus.style = "Table Grid"

    cab = tabela_bonus.rows[0].cells
    cab[0].text = "Aluno"
    cab[1].text = "Turma/Período"
    cab[2].text = "% acertos"
    cab[3].text = "% acima da média"
    cab[4].text = "Faixa"
    cab[5].text = "Disciplinas"
    cab[6].text = "Total pontos"

    for _, row in pontos_extras.iterrows():
        cells = tabela_bonus.add_row().cells
        cells[0].text = str(row[coluna_nome])
        cells[1].text = str(row[coluna_periodo])
        cells[2].text = f"{row['Percentual_acertos']:.2f}%"
        cells[3].text = f"{row['Percentual_acima_media_turma']:.2f}%"
        cells[4].text = str(row["Faixa_bonus"])
        cells[5].text = str(int(row["Numero_disciplinas_bonus"]))
        cells[6].text = f"{row['Total_pontos_extras']:.2f}"

    adicionar_fonte(doc, config)

    doc.add_heading("8. Encaminhamentos pedagógicos sugeridos", level=1)

    doc.add_paragraph(
        estilo.get(
            "monitoramento",
            "Os indicadores produzidos pelo Teste de Progresso subsidiam ações do NDE e da coordenação do curso."
        )
    )

    doc.add_paragraph(
        estilo.get(
            "reforco",
            "As áreas com menor desempenho são analisadas como oportunidades de reforço pedagógico."
        )
    )

    doc.add_paragraph(
        estilo.get(
            "curriculo",
            "Os resultados observados reforçam o alinhamento entre a matriz curricular e as competências previstas no PPC."
        )
    )

    doc.add_paragraph(
        "Recomenda-se que os resultados sejam discutidos em reuniões do NDE e do colegiado, com registro em ata, "
        "definição de ações de melhoria, acompanhamento dos estudantes em risco e planejamento de intervenções "
        "pedagógicas por turma, ciclo formativo e área do conhecimento."
    )

    caminho = os.path.join(pastas["relatorios"], "relatorio_completo_mec.docx")
    doc.save(caminho)

    return caminho


def main():
    config = carregar_config()
    pastas = criar_pasta_projeto(config)

    caminho_planilha = pedir_planilha(config)

    shutil.copy(
        caminho_planilha,
        os.path.join(pastas["original"], os.path.basename(caminho_planilha))
    )

    base = carregar_dados(caminho_planilha)

    coluna_nome, coluna_periodo, coluna_total = detectar_colunas(base)

    base, coluna_nome, coluna_periodo = preparar_base(
        base,
        coluna_nome,
        coluna_periodo,
        coluna_total,
        config
    )

    (
        ranking_alunos,
        ranking_periodos,
        tabela_ciclos,
        alunos_risco,
        pontos_extras,
        caminho_excel,
        tabelas_areas
    ) = gerar_tabelas(
        base,
        coluna_nome,
        coluna_periodo,
        pastas,
        config
    )

    graficos = gerar_graficos(
        base,
        ranking_periodos,
        tabela_ciclos,
        coluna_periodo,
        pastas
    )

    relatorio = criar_relatorio_word(
        base,
        ranking_periodos,
        tabela_ciclos,
        alunos_risco,
        pontos_extras,
        graficos,
        coluna_nome,
        coluna_periodo,
        pastas,
        tabelas_areas,
        config
    )

    print("\n==============================================")
    print("ROBÔ MEC FINALIZADO COM SUCESSO!")
    print("==============================================")
    print(f"Pasta do projeto: {pastas['base']}")
    print(f"Tabelas: {caminho_excel}")
    print(f"Relatório: {relatorio}")
    print(f"Gráficos: {pastas['graficos']}")
    print("==============================================")


if __name__ == "__main__":
    main()