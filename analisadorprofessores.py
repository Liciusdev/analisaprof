import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from io import BytesIO

st.set_page_config(page_title="Analisador de Avaliação Institucional", layout="wide")

st.title("📊 Analisador de Avaliação Institucional")

# Upload de arquivo
uploaded_file = st.file_uploader("Selecione o arquivo Excel", type=["xls", "xlsx"])
if uploaded_file:
    with st.spinner("⏳ Processando a planilha, aguarde..."):
        resultados = processar_excel(uploaded_file)

    if resultados:
        st.success("✅ Relatório processado com sucesso!")

def processar_excel(arquivo_excel):
    df = pd.read_excel(arquivo_excel)

    colunas_necessarias = [
        "avaliacaoinstitucional",
        "pergunta",
        "nomeprofessor",
        "nomecurso",
        "nomedisciplina",
        "nomeunidadeensino"
    ]
    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            st.error(f"A coluna '{coluna}' não foi encontrada no arquivo Excel.")
            return None

    professores = df["nomeprofessor"].unique().tolist()
    cursos = df["nomecurso"].unique().tolist()
    questionarios = df["avaliacaoinstitucional"].unique().tolist()
    unidades_ensino = df["nomeunidadeensino"].unique().tolist()

    todos_resultados = {}

    for professor in professores:
        for curso in cursos:
            for questionario in questionarios:
                for unidade_ensino in unidades_ensino:
                    df_filtrado = df[
                        (df["nomeprofessor"] == professor) &
                        (df["nomecurso"] == curso) &
                        (df["avaliacaoinstitucional"] == questionario) &
                        (df["nomeunidadeensino"] == unidade_ensino)
                    ].copy()

                    disciplinas = df_filtrado["nomedisciplina"].unique().tolist()
                    for disciplina in disciplinas:
                        df_disciplina = df_filtrado[df_filtrado["nomedisciplina"] == disciplina].copy()
                        resultados_disciplina = {}
                        for _, row in df_disciplina.iterrows():
                            pergunta = row["pergunta"]
                            colunas_restantes = row.drop(colunas_necessarias).dropna()
                            resposta = str(colunas_restantes.iloc[0]).strip().upper() if not colunas_restantes.empty else "NÃO AVALIADO"

                            if pergunta not in resultados_disciplina:
                                resultados_disciplina[pergunta] = {
                                    "BOM": 0, "MUITO BOM": 0, "OTIMO": 0, "REGULAR": 0,
                                    "RUIM": 0, "NAO_SEI_OU_NAO_TENHO_CONDICOES_DE_AVALIAR": 0
                                }
                            if resposta in ["BOM", "MUITO BOM", "ÓTIMO", "REGULAR", "RUIM", "NÃO SEI OU NÃO TENHO CONDIÇÕES DE AVALIAR"]:
                                chave = resposta.replace("Ó", "O").replace("Ã", "A").replace(" ", "_")
                                resultados_disciplina[pergunta][chave] += 1

                        tabela_dados = [["PERGUNTAS", "NÃO SEI(0)", "R(1)", "REG(2)", "B(3)", "MB(4)", "O(5)", "SOMA", "MÉDIA", "PORC"]]
                        soma_ponderada_total, total_respostas_total = 0, 0

                        for pergunta, contagens in resultados_disciplina.items():
                            bom = contagens.get("BOM", 0)
                            muito_bom = contagens.get("MUITO BOM", 0)
                            otimo = contagens.get("OTIMO", 0)
                            regular = contagens.get("REGULAR", 0)
                            ruim = contagens.get("RUIM", 0)
                            nao_sei = contagens.get("NAO_SEI_OU_NAO_TENHO_CONDICOES_DE_AVALIAR", 0)

                            soma_pontos = bom*3 + muito_bom*4 + otimo*5 + regular*2 + ruim*1
                            total_respostas = bom + muito_bom + otimo + regular + ruim
                            media = soma_pontos / total_respostas if total_respostas > 0 else 0
                            porcentual = (soma_pontos / (total_respostas * 5)) * 100 if total_respostas > 0 else 0

                            tabela_dados.append([pergunta, nao_sei, ruim, regular, bom, muito_bom, otimo,
                                                 soma_pontos, f"{media:.2f}", f"{porcentual:.2f}%"])

                            soma_ponderada_total += soma_pontos
                            tot

