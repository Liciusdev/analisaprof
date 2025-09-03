import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from io import BytesIO

st.set_page_config(page_title="Analisador de Avaliação Institucional", layout="wide")
st.title("📊 Analisador de Avaliação Institucional")

# ---------------- FUNÇÃO PROCESSAR OTIMIZADA ----------------
def processar_excel(uploaded_file):
    df = pd.read_excel(uploaded_file)

    colunas_necessarias = [
        "avaliacaoinstitucional", "pergunta", "nomeprofessor",
        "nomecurso", "nomedisciplina", "nomeunidadeensino"
    ]
    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            st.error(f"A coluna '{coluna}' não foi encontrada no arquivo Excel.")
            return None

    # Normalizar respostas
    def normalizar_resposta(x):
        x = str(x).upper().strip()
        if x == "MUITO BOM":
            return "MUITO_BOM"
        elif x == "ÓTIMO":
            return "OTIMO"
        elif x == "NÃO SEI OU NÃO TENHO CONDIÇÕES DE AVALIAR":
            return "NAO_SEI_OU_NAO_TENHO_CONDICOES_DE_AVALIAR"
        elif x in ["BOM", "REGULAR", "RUIM"]:
            return x
        else:
            return "NAO_SEI_OU_NAO_TENHO_CONDICOES_DE_AVALIAR"

    df['resposta'] = df.drop(columns=colunas_necessarias).bfill(axis=1).iloc[:,0]
    df['resposta'] = df['resposta'].apply(normalizar_resposta)

    # Agrupar dados
    todos_resultados = {}
    grupos = df.groupby(['nomeprofessor','nomecurso','avaliacaoinstitucional','nomeunidadeensino','nomedisciplina','pergunta'])

    total_grupos = len(df.groupby(['nomeprofessor','nomecurso','avaliacaoinstitucional','nomeunidadeensino','nomedisciplina']))
    contador = 0
    progress_bar = st.progress(0)

    for (professor, curso, questionario, unidade, disciplina), grupo in df.groupby(['nomeprofessor','nomecurso','avaliacaoinstitucional','nomeunidadeensino','nomedisciplina']):
        resultados_disciplina = {}

        for pergunta, subgrupo in grupo.groupby('pergunta'):
            contagem = subgrupo['resposta'].value_counts().to_dict()
            resultados_disciplina[pergunta] = {
                "BOM": contagem.get("BOM",0),
                "MUITO_BOM": contagem.get("MUITO_BOM",0),
                "OTIMO": contagem.get("OTIMO",0),
                "REGULAR": contagem.get("REGULAR",0),
                "RUIM": contagem.get("RUIM",0),
                "NAO_SEI_OU_NAO_TENHO_CONDICOES_DE_AVALIAR": contagem.get("NAO_SEI_OU_NAO_TENHO_CONDICOES_DE_AVALIAR",0)
            }

        # Montar tabela
        tabela_dados = [["PERGUNTAS", "NÃO SEI(0)", "R(1)", "REG(2)", "B(3)", "MB(4)", "O(5)", "SOMA", "MÉDIA", "PORC"]]
        soma_total, respostas_total = 0,0

        for pergunta, cont in resultados_disciplina.items():
            bom = cont["BOM"]; muito_bom = cont["MUITO_BOM"]; otimo = cont["OTIMO"]
            regular = cont["REGULAR"]; ruim = cont["RUIM"]; nao_sei = cont["NAO_SEI_OU_NAO_TENHO_CONDICOES_DE_AVALIAR"]

            soma_pontos = bom*3 + muito_bom*4 + otimo*5 + regular*2 + ruim*1
            total_respostas = bom + muito_bom + otimo + regular + ruim
            media = soma_pontos/total_respostas if total_respostas>0 else 0
            porcentual = (soma_pontos/(total_respostas*5))*100 if total_respostas>0 else 0

            tabela_dados.append([pergunta,nao_sei,ruim,regular,bom,muito_bom,otimo,soma_pontos,f"{media:.2f}",f"{porcentual:.2f}%"])
            soma_total += soma_pontos
            respostas_total += total_respostas

        media_geral = (soma_total/(respostas_total*5))*100 if respostas_total>0 else 0
        chave = (professor,curso,questionario,disciplina,unidade)
        todos_resultados[chave] = (tabela_dados, media_geral)

        contador += 1
        progress_bar.progress(contador/total_grupos)

    return todos_resultados

# ---------------- FUNÇÃO GERAR PDF ----------------
def gerar_pdf(todos_resultados):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter))
    estilo_titulo = ParagraphStyle(name='Arial10', fontSize=10, leading=12, textColor=colors.black, alignment=TA_LEFT)
    estilo_media = ParagraphStyle(name='Arial12Center', fontSize=12, leading=14, textColor=colors.black, alignment=TA_CENTER)
    conteudo = []

    for (professor,curso,questionario,disciplina,unidade),(tabela,media) in todos_resultados.items():
        conteudo.append(Paragraph(f"Professor: {professor}", estilo_titulo))
        conteudo.append(Paragraph(f"Curso: {curso}", estilo_titulo))
        conteudo.append(Paragraph(f"Questionário: {questionario}", estilo_titulo))
        conteudo.append(Paragraph(f"Unidade: {unidade}", estilo_titulo))
        conteudo.append(Paragraph(f"Disciplina: {disciplina}", estilo_titulo))
        conteudo.append(Spacer(1,6))
        t = Table(tabela, colWidths=[None]*10)
        t.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),colors.grey),
            ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('FONTSIZE',(0,0),(-1,-1),6),
            ('BACKGROUND',(0,1),(-1,-1),colors.paleturquoise),
            ('GRID',(0,0),(-1,-1),0.5,colors.black)
        ]))
        conteudo.append(t)
        conteudo.append(Spacer(1,12))
        conteudo.append(Paragraph(f"Média Geral (%): {media:.2f}", estilo_media))
        conteudo.append(PageBreak())

    if conteudo:
        conteudo = conteudo[:-1]
    doc.build(conteudo)
    buffer.seek(0)
    return buffer

# ---------------- INTERFACE STREAMLIT ----------------
uploaded_file = st.file_uploader("Selecione o arquivo Excel", type=["xls","xlsx"])

if uploaded_file:
    with st.spinner("⏳ Processando a planilha..."):
        resultados = processar_excel(uploaded_file)

    if resultados:
        st.success("✅ Relatório processado!")
        for (professor,curso,questionario,disciplina,unidade),(tabela,media) in resultados.items():
            with st.expander(f"{professor} - {disciplina} ({curso}) [{questionario}] - {unidade}"):
                st.table(pd.DataFrame(tabela[1:],columns=tabela[0]))
                st.write(f"**Média Geral (%):** {media:.2f}")

        pdf_buffer = gerar_pdf(resultados)
        st.download_button("📥 Baixar PDF", pdf_buffer, file_name="relatorio_avaliacao.pdf", mime="application/pdf")
