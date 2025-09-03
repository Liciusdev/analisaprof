import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT

try:
    pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
except:
    print("Fonte Arial já registrada ou arquivo Arial.ttf não encontrado.")

def processar_excel():
    arquivo_excel = filedialog.askopenfilename(
        title="Selecionar arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xls;*.xlsx")]
    )

    if not arquivo_excel:
        return

    try:
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
                messagebox.showerror("Erro", f"A coluna '{coluna}' não foi encontrada no arquivo Excel.")
                return
        # Obter lista única de professores
        professores = df["nomeprofessor"].unique().tolist()
        cursos = df["nomecurso"].unique().tolist()
        questionarios = df["avaliacaoinstitucional"].unique().tolist()
        unidades_ensino = df["nomeunidadeensino"].unique().tolist()  

        # Estrutura para armazenar resultados
        todos_resultados = {}

        
        for professor in professores:
            for curso in cursos:
                for questionario in questionarios:
                    for unidade_ensino in unidades_ensino:  
                        # Filtrar o DataFrame
                        df_filtrado = df[
                            (df["nomeprofessor"] == professor) &
                            (df["nomecurso"] == curso) &
                            (df["avaliacaoinstitucional"] == questionario) &
                            (df["nomeunidadeensino"] == unidade_ensino)  

                            ].copy()

                        # Disciplinas para o professor, curso e questionário atuais
                        disciplinas = df_filtrado["nomedisciplina"].unique().tolist()
                        # Iterar sobre as disciplinas
                        for disciplina in disciplinas:
                            df_disciplina = df_filtrado[df_filtrado["nomedisciplina"] == disciplina].copy()
                            # Estrutura para armazenar resultados da disciplina
                            resultados_disciplina = {}
                            for index, row in df_disciplina.iterrows():
                                pergunta = row["pergunta"]
                                colunas_restantes = row.drop(colunas_necessarias).dropna()

                                if not colunas_restantes.empty:
                                    resposta = colunas_restantes.iloc[0]
                                else:
                                    resposta = "NÃO AVALIADO"
                                    print(
                                        f"Aviso: Nenhuma resposta encontrada para a linha {index}. Pergunta: {row['pergunta']}")

                                resposta = str(resposta).strip().upper()
                                print(f"Resposta: {resposta}")

                                if pergunta not in resultados_disciplina:
                                    resultados_disciplina[pergunta] = {
                                        "BOM": 0,
                                        "MUITO BOM": 0,
                                        "OTIMO": 0,
                                        "REGULAR": 0,
                                        "RUIM": 0,
                                        "NAO_SEI_OU_NAO_TENHO_CONDICOES_DE_AVALIAR": 0
                                    }
                                if resposta == "BOM":
                                    resultados_disciplina[pergunta]["BOM"] += 1
                                elif resposta == "MUITO BOM":
                                    resultados_disciplina[pergunta]["MUITO BOM"] += 1
                                elif resposta == "ÓTIMO":
                                    resultados_disciplina[pergunta]["OTIMO"] += 1
                                elif resposta == "REGULAR":
                                    resultados_disciplina[pergunta]["REGULAR"] += 1
                                elif resposta == "RUIM":
                                    resultados_disciplina[pergunta]["RUIM"] += 1
                                elif resposta == "NÃO SEI OU NÃO TENHO CONDIÇÕES DE AVALIAR":
                                    resultados_disciplina[pergunta]["NAO_SEI_OU_NAO_TENHO_CONDICOES_DE_AVALIAR"] += 1

                            # Criar a tabela para exibição
                            tabela_dados = [
                                ["PERGUNTAS", "NÃO SEI(0)", "R(1)", "REG(2)", "B(3)", "MB(4)", "O(5)", "SOMA", "MÉDIA",
                                 "PORC"]]

                            soma_ponderada_total = 0
                            total_respostas_total = 0
                            perguntas_lista = list(resultados_disciplina.keys())
                            for pergunta in perguntas_lista:
                                contagens = resultados_disciplina[pergunta]

                                # Imprimir as 'contagens'
                                print(f"Chaves do dicionário 'contagens': {contagens.keys()}")

                                # Debug
                                print(f"Dicionário 'contagens': {contagens}")

                                # Verificar se as chaves existem antes de usá-las
                                bom = contagens.get("BOM", 0)  
                                muito_bom = contagens.get("MUITO BOM", 0)
                                otimo = contagens.get("OTIMO", 0)
                                regular = contagens.get("REGULAR", 0)
                                ruim = contagens.get("RUIM", 0)
                                nao_sei = contagens.get("NAO_SEI_OU_NAO_TENHO_CONDICOES_DE_AVALIAR", 0)

                                soma_pontos = (
                                        bom * 3 +
                                        muito_bom * 4 +
                                        otimo * 5 +
                                        regular * 2 +
                                        ruim * 1
                                )
                                total_respostas = bom + muito_bom + otimo + regular + ruim  

                                # Cálculo da média
                                media = soma_pontos / total_respostas if total_respostas > 0 else 0
                                # Calculo do porcentual
                                porcentual = (soma_pontos / (total_respostas * 5)) * 100 if total_respostas > 0 else 0

                                tabela_dados.append([
                                    pergunta,
                                    str(nao_sei), 
                                    str(ruim),  
                                    str(regular),  
                                    str(bom),  
                                    str(muito_bom),  
                                    str(otimo),  
                                    str(soma_pontos),
                                    f"{media:.2f}",  
                                    f"{porcentual:.2f}%"  

                                ])

                                soma_ponderada_total += soma_pontos
                                total_respostas_total += total_respostas

                            # Calcular a média geral ponderada
                            media_geral_ponderada = (soma_ponderada_total / (total_respostas_total * 5)) * 100 if total_respostas_total > 0 else 0
                            chave_resultado = (
                                professor, curso, questionario, disciplina, unidade_ensino)  
                            todos_resultados[chave_resultado] = (tabela_dados, media_geral_ponderada)

        # Limpar Text
        texto_resultados.delete("1.0", tk.END)
        # Imprimir a tabela
        for (professor, curso, questionario, disciplina, unidade_ensino), (tabela_dados, media_geral_ponderada) in todos_resultados.items():  
            texto_resultados.insert(tk.END, f"Nome do Professor: {professor}\n")
            texto_resultados.insert(tk.END, f"Curso: {curso}\n")
            texto_resultados.insert(tk.END, f"Questionário: {questionario}\n")
            texto_resultados.insert(tk.END, f"Unidade de Ensino: {unidade_ensino}\n")  
            texto_resultados.insert(tk.END, f"Disciplina: {disciplina}\n\n")
            for linha in tabela_dados:
                texto_resultados.insert(tk.END, "| " + " | ".join(linha) + " |\n")
            texto_resultados.insert(tk.END, f"Média Geral (%): {media_geral_ponderada:.2f}%\n\n")
        # Gerar PDF com todos os resultados
        gerar_pdf(todos_resultados)

    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo não encontrado.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao processar o arquivo: {e}")

def gerar_pdf(todos_resultados):
    arquivo_pdf = filedialog.asksaveasfilename(
        title="Salvar arquivo PDF",
        defaultextension=".pdf",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    if not arquivo_pdf:
        return

    doc = SimpleDocTemplate(arquivo_pdf, pagesize=landscape(letter))
    estilos = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        name='Arial10',
        fontName='Arial',
        fontSize=10,
        leading=12,
        textColor=colors.black,
        alignment=TA_LEFT,
    )
    estilo_media_geral = ParagraphStyle(
        name='Arial12Center',
        fontName='Arial',
        fontSize=12,
        leading=14,
        textColor=colors.black,
        alignment=TA_CENTER,
    )

    conteudo = []
    for (professor, curso, questionario, disciplina, unidade_ensino), (tabela_dados, media_geral_ponderada) in todos_resultados.items():  

        conteudo.append(Paragraph(f"Nome do Professor: {professor}", estilo_titulo))
        conteudo.append(Paragraph(f"Curso: {curso}", estilo_titulo))
        conteudo.append(Paragraph(f"Questionário: {questionario}", estilo_titulo))
        conteudo.append(Paragraph(f"Unidade de Ensino: {unidade_ensino}", estilo_titulo))  
        conteudo.append(Paragraph(f"Disciplina: {disciplina}", estilo_titulo))
        conteudo.append(Spacer(1, 6))  

        t = Table(tabela_dados, colWidths=[None] * 10)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 3),
            ('BACKGROUND', (0, 1), (-1, -1), colors.paleturquoise),  
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),  
            ('WORDWRAP', (0, 0), (-1, -1), 1),  
        ]))
        conteudo.append(t)
        conteudo.append(Spacer(1, 12))
        conteudo.append(Paragraph(f"Média Geral (%): {media_geral_ponderada:.2f}%", estilo_media_geral)) 

        conteudo.append(PageBreak())  

    
    if len(todos_resultados) > 0:
        conteudo = conteudo[:-1]

    doc.build(conteudo)
    messagebox.showinfo("Sucesso", f"PDF salvo em {arquivo_pdf}")


# --- Interface Gráf ---
janela = tk.Tk()
janela.title("Analisador de Avaliação Institucional")
janela.geometry("800x600")

style = ttk.Style()
style.configure("TButton", padding=6, relief="flat", background="#4CAF50", foreground="white")
style.configure("TLabel", padding=6)

botao_selecionar = ttk.Button(janela, text="Clique aqui para selecionar o arquivo Excel e gerar relatório",
                             command=processar_excel)
botao_selecionar.pack(pady=20)

texto_resultados = tk.Text(janela, wrap=tk.WORD)
texto_resultados.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

barra_rolagem = ttk.Scrollbar(janela, orient=tk.VERTICAL, command=texto_resultados.yview)
barra_rolagem.pack(side=tk.RIGHT, fill=tk.Y)
texto_resultados["yscrollcommand"] = barra_rolagem.set

janela.mainloop()