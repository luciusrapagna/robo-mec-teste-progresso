import random


def carregar_frases(caminho):
    with open(caminho, "r", encoding="utf-8") as f:
        frases = [linha.strip() for linha in f if linha.strip()]
    return frases


def frase_abertura():
    frases = carregar_frases(
        "modelos/frases/frases_abertura.txt"
    )
    return random.choice(frases)


def frase_fechamento():
    frases = carregar_frases(
        "modelos/frases/frases_fechamento.txt"
    )
    return random.choice(frases)


def conectivo():
    frases = carregar_frases(
        "modelos/frases/conectivos.txt"
    )
    return random.choice(frases)


def humanizar_relatorio(
    total,
    melhor_turma,
    melhor_media,
    pior_turma,
    pior_media,
    alunos_risco,
    percentual_risco
):

    texto = f"""
{frase_abertura()}

{conectivo()}, foram avaliados {total} estudantes no Teste de Progresso, considerando o desempenho global, as diferenças entre turmas/períodos e a identificação de estudantes que demandam acompanhamento pedagógico mais próximo.

A turma/período com melhor desempenho foi {melhor_turma}, apresentando média de {melhor_media:.2f}% de acertos. Esse resultado pode indicar maior consolidação dos conteúdos avaliados e melhor desempenho acadêmico coletivo.

{conectivo()}, a turma/período com menor desempenho médio foi {pior_turma}, com média de {pior_media:.2f}% de acertos. Esse achado deve ser interpretado como importante indicador institucional para planejamento de estratégias de reforço pedagógico, revisão de conteúdos e acompanhamento acadêmico.

Foram identificados {alunos_risco} estudantes com desempenho inferior ao critério estabelecido, representando {percentual_risco:.2f}% do total analisado.

{frase_fechamento()}
"""

    return texto.strip()