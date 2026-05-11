import os
from docx import Document
from pypdf import PdfReader


def ler_docx(caminho):
    texto = []
    doc = Document(caminho)

    for paragrafo in doc.paragraphs:
        conteudo = paragrafo.text.strip()
        if conteudo:
            texto.append(conteudo)

    return "\n".join(texto)


def ler_pdf(caminho):
    texto = []
    reader = PdfReader(caminho)

    for pagina in reader.pages:
        conteudo = pagina.extract_text()
        if conteudo:
            texto.append(conteudo)

    return "\n".join(texto)


def carregar_banco_escrita():
    pasta_base = "banco-escrita"
    pasta_word = os.path.join(pasta_base, "exemplos_word")
    pasta_pdf = os.path.join(pasta_base, "exemplos_pdf")
    pasta_saida = os.path.join(pasta_base, "textos_extraidos")

    os.makedirs(pasta_word, exist_ok=True)
    os.makedirs(pasta_pdf, exist_ok=True)
    os.makedirs(pasta_saida, exist_ok=True)

    textos = []

    for arquivo in os.listdir(pasta_word):
        if arquivo.lower().endswith(".docx"):
            caminho = os.path.join(pasta_word, arquivo)
            try:
                textos.append(ler_docx(caminho))
            except Exception as erro:
                print(f"Erro ao ler DOCX {arquivo}: {erro}")

    for arquivo in os.listdir(pasta_pdf):
        if arquivo.lower().endswith(".pdf"):
            caminho = os.path.join(pasta_pdf, arquivo)
            try:
                textos.append(ler_pdf(caminho))
            except Exception as erro:
                print(f"Erro ao ler PDF {arquivo}: {erro}")

    banco_textual = "\n\n".join(textos)

    caminho_saida = os.path.join(pasta_saida, "banco_textual_extraido.txt")

    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write(banco_textual)

    return banco_textual


def gerar_orientacao_estilo():
    texto = carregar_banco_escrita()

    if not texto.strip():
        return (
            "Não foram encontrados exemplos de escrita no banco-escrita. "
            "Utilizar linguagem institucional, pedagógica, clara e adequada para avaliação do MEC."
        )

    return (
        "Utilizar como referência de estilo os textos previamente inseridos no banco de escrita. "
        "Priorizar linguagem institucional, clareza pedagógica, tom analítico, uso de conectivos, "
        "menção ao NDE, ao PPC, ao acompanhamento longitudinal e à melhoria contínua."
    )