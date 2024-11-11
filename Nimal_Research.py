import datetime
import mysql.connector
from openpyxl import load_workbook
import customtkinter as ctk
import os
from tkinter import filedialog
from PIL import Image
import pandas as pd
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import ttk
from datetime import datetime


# Conectando ao banco de dados
conexao = mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="nimal"
)

# Criando um cursor
cursor = conexao.cursor()


# Configuração do tema escuro
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Função para extrair informações do PDF e salvar no Excel
import fitz
import re
import customtkinter as ctk
from tkinter import messagebox, simpledialog


def extrair_informacoes_pdf(caminho_pdf):
    try:
        with fitz.open(caminho_pdf) as pdf:
            conteudo = ""
            for pagina in pdf:
                conteudo += pagina.get_text()

            linhas = conteudo.splitlines()

            # Listas de padrões para buscar diferentes variações
            padroes_venc = [r"Venc\.\s*(.*)", r"VENCIMENTO\s*(.*)"]
            padroes_nf = [r"Nº\.\s*(.*)", r"Nº\s*(.*)", r"Número Fiscal\s*(.*)"]
            padroes_dist = [r"IDENTIFICAÇÃO DO EMITENTE\s*(.*)", r"Distribuidor\s*(.*)", r"RECEBEMOS DE\s*(.*)"]
            padroes_valor = [r"V\. TOTAL DA NOTA\s*(.*)", r"VALOR TOTAL DA NOTA\s*(.*)", r"V\. Total\s*(.*)"]

            # Função auxiliar para tentar vários padrões até encontrar uma correspondência
            def buscar_texto(padroes):
                for padrao in padroes:
                    match = re.search(padrao, conteudo)
                    if match:
                        return match.group(1).strip()
                return "Não encontrado"

            # Usar a função auxiliar para buscar cada campo
            venc_extraido = buscar_texto(padroes_venc)
            nf_extraido = buscar_texto(padroes_nf)
            dist_extraido = buscar_texto(padroes_dist)
            valor_extraido = buscar_texto(padroes_valor)

            # Chamar função para editar e confirmar dados
            editar_dados(venc_extraido, nf_extraido, dist_extraido, valor_extraido)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o PDF: {e}")


def editar_dados(venc, nf, dist, valor):
    # Criar uma nova janela para edição
    editar_janela = ctk.CTkToplevel()
    editar_janela.title("Editar Informações Extraídas")
    editar_janela.geometry("700x500")

    # Forçar a nova janela a ficar sempre em cima
    editar_janela.attributes('-topmost', True)

    # Labels e campos de entrada
    ctk.CTkLabel(editar_janela, text="Vencimento:").pack(pady=5)
    venc_entry = ctk.CTkEntry(editar_janela)
    venc_entry.insert(0, venc)
    venc_entry.pack(pady=5)

    ctk.CTkLabel(editar_janela, text="NF:").pack(pady=5)
    nf_entry = ctk.CTkEntry(editar_janela)
    nf_entry.insert(0, nf)
    nf_entry.pack(pady=5)

    ctk.CTkLabel(editar_janela, text="Distribuidor:").pack(pady=5)
    dist_entry = ctk.CTkEntry(editar_janela)
    dist_entry.insert(0, dist)
    dist_entry.pack(pady=5)

    ctk.CTkLabel(editar_janela, text="Valor:").pack(pady=5)
    valor_entry = ctk.CTkEntry(editar_janela)
    valor_entry.insert(0, valor)
    valor_entry.pack(pady=5)

    # Função para confirmar e salvar os dados
    def confirmar():
        venc_editado = venc_entry.get()
        nf_editado = nf_entry.get()
        dist_editado = dist_entry.get()
        valor_editado = valor_entry.get()

        editar_janela.attributes('-topmost', False)
        # Salvar informações no Excel e no banco de dados
        salvar_informacoes_excel_base(venc_editado, nf_editado, dist_editado, valor_editado)
        editar_janela.destroy()  # Fecha a janela de edição
        # Mostrar mensagem de sucesso e fechar a janela



    # Botão de confirmação
    ctk.CTkButton(editar_janela, text="Salvar", command=confirmar, fg_color="gray", hover_color="darkred", width=300, height=40,
                                font=("Impact", 18), corner_radius=10).pack(pady=10)

    # Botão para cancelar a edição
    ctk.CTkButton(editar_janela, text="Cancelar", command=editar_janela.destroy, fg_color="gray", hover_color="darkred", width=300, height=40,
                                font=("Impact", 18), corner_radius=10).pack(pady=5)


def salvar_informacoes_excel_base(venc, nf, dist, valor):
    caminho_excel = filedialog.askopenfilename(title="Selecione o arquivo Excel de base",
                                               filetypes=[("Excel files", "*.xlsx")])
    if not caminho_excel or not os.access(caminho_excel, os.W_OK):
        messagebox.showerror("Erro", "Arquivo inválido ou sem permissão de escrita.")
        return

    try:
        wb = load_workbook(caminho_excel)
        sheet = wb.active

        # Limpar preenchimento das células a partir da segunda linha
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.fill = PatternFill(fill_type=None)  # Limpa o preenchimento

        # Criar colunas nas posições 12, 13, 14 e 15 se não existirem
        colunas_posicoes = {12: "vencimento", 13: "nf", 14: "valor nf", 15: "distribuidor"}
        preenchimento = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        for posicao, nome_coluna in colunas_posicoes.items():
            if sheet.cell(row=1, column=posicao).value != nome_coluna:
                sheet.cell(row=1, column=posicao, value=nome_coluna)
                sheet.cell(row=1, column=posicao).fill = preenchimento

        # Perguntar pelo valor do orçamento
        valor_orcamento = simpledialog.askstring("Valor do Orçamento",
                                                 "Digite o valor do orçamento que você deseja buscar:")
        if not valor_orcamento:
            return

        coluna_orcamento = next(
            (cell.column for cell in sheet[1] if str(cell.value).lower() in ["orcamento", "orçamento"]), None)
        if coluna_orcamento is None:
            messagebox.showerror("Erro", "Coluna 'Orçamento' não encontrada.")
            return

        linha_encontrada = next(
            (row[0].row for row in sheet.iter_rows(min_row=2) if row[coluna_orcamento - 1].value == valor_orcamento),
            None)
        if linha_encontrada is None:
            messagebox.showerror("Erro", "Valor de orçamento não encontrado.")
            return

        # Adiciona os dados nas colunas corretas na planilha Excel
        sheet.cell(row=linha_encontrada, column=12, value=venc)
        sheet.cell(row=linha_encontrada, column=13, value=nf)
        sheet.cell(row=linha_encontrada, column=14, value=valor)
        sheet.cell(row=linha_encontrada, column=15, value=dist)

        # Perguntar ao usuário se deseja adicionar uma nova nota fiscal ao pedido
        adicionar_nf = messagebox.askyesno("Adicionar Nota Fiscal",
                                           "Deseja adicionar uma nova nota fiscal para este pedido?")

        if adicionar_nf:
            # Incrementar o valor do orçamento para criar um novo registro
            base_orcamento = valor_orcamento.rsplit('-', 1)[0]
            numero_incremento = int(valor_orcamento.rsplit('-', 1)[1]) + 1

            # Verificar se o próximo orçamento já existe
            novo_orcamento = f"{base_orcamento}-{numero_incremento}"
            while any(row[coluna_orcamento - 1].value == novo_orcamento for row in sheet.iter_rows(min_row=2)):
                numero_incremento += 1
                novo_orcamento = f"{base_orcamento}-{numero_incremento}"

            # Mover linhas abaixo da linha encontrada para baixo
            sheet.insert_rows(linha_encontrada + 1)

            # Duplicar a linha com o novo orçamento logo abaixo da linha encontrada
            for col in range(1, sheet.max_column + 1):
                sheet.cell(row=linha_encontrada + 1, column=col,
                           value=sheet.cell(row=linha_encontrada, column=col).value)
            sheet.cell(row=linha_encontrada + 1, column=coluna_orcamento, value=novo_orcamento)

            # Adicionar os novos dados para a linha duplicada
            sheet.cell(row=linha_encontrada + 1, column=12, value=venc)
            sheet.cell(row=linha_encontrada + 1, column=13, value=nf)
            sheet.cell(row=linha_encontrada + 1, column=14, value=valor)
            sheet.cell(row=linha_encontrada + 1, column=15, value=dist)

            messagebox.showinfo("Sucesso", "Nova nota fiscal adicionada ao pedido.")

        # Salvar as alterações no Excel
        wb.save(caminho_excel)
        messagebox.showinfo("Sucesso", "Informações salvas no Excel.")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar no Excel: {e}")


# Função para abrir o seletor de PDF
def selecionar_pdf():
    caminho_pdf = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if caminho_pdf:
        extrair_informacoes_pdf(caminho_pdf)


def restaurar_menu():
    # Limpa o conteúdo anterior no menu
    for widget in menu_frame.winfo_children():
        widget.destroy()

    # Recria os botões do menu
    botao_instrucoes = ctk.CTkButton(menu_frame, text="   Instruções", command=lambda: mostrar_frame(instrucoes_frame),
                                     fg_color="#681919", hover_color="darkred", width=500, height=80,
                                     image=imagem_instrucoes, compound="left", font=("Impact", 18), corner_radius=10,
                                     anchor="w")
    botao_instrucoes.pack(padx=10, pady=20)

    botao_selecionar_pdf = ctk.CTkButton(menu_frame, text="    Selecionar PDF", command=selecionar_pdf,
                                         fg_color="#681919", hover_color="darkred", width=500, height=80,
                                         image=imagem_pdf, font=("Impact", 18), corner_radius=10, anchor="w")
    botao_selecionar_pdf.pack(padx=50, pady=20)

    botao_visao_geral = ctk.CTkButton(menu_frame, text="     Visão Geral", command=mostrar_visao_geral,
                                      fg_color="#681919", hover_color="darkred", width=500, height=80,
                                      image=imagem_visao, font=("Impact", 18), corner_radius=10, anchor="w")
    botao_visao_geral.pack(padx=50, pady=20)

    botao_importar = ctk.CTkButton(menu_frame, text="   Importar Excel", command=importar_dados_excel, fg_color="#681919",
                                   hover_color="darkred", width=500, height=80,
                                   font=("Impact", 18), corner_radius=10, anchor="w",image=imagem_excel)
    botao_importar.pack(padx=50, pady=20)


# Atualizando a função para exibir a visão geral
def mostrar_visao_geral():
    # Limpa o conteúdo anterior no frame de menu
    for widget in menu_frame.winfo_children():
        widget.destroy()

    # Frame para a visão geral
    visao_frame = ctk.CTkFrame(frame2, fg_color="#C0C0C0")
    visao_frame.place(relwidth=1, relheight=1)

    # Campo de entrada para o valor de filtro
    label_filtro = ctk.CTkLabel(visao_frame, text="Filtrar por valor:", font=("Arial", 14))
    label_filtro.pack(pady=5)
    entry_filtro = ctk.CTkEntry(visao_frame, width=200, font=("Arial", 12))
    entry_filtro.pack(pady=5)

    # Função para carregar dados com filtro por valor
    def carregar_dados(valor=None):
        # Limpa a tabela antes de carregar os novos dados
        for row in tree.get_children():
            tree.delete(row)

        # Conecta ao banco de dados
        conexao = mysql.connector.connect(
            host="localhost",
            user="root",
            password="",
            database="nimal"
        )
        cursor = conexao.cursor()

        # SQL básico com filtro opcional
        sql_query = """
            SELECT local, orcamento,pedido,data, situacao, cliente, razao, representante, itens,total, vencimento, nf, valor_nf, distribuidor 
            FROM nimalnotas
        """
        if valor:  # Se um valor for passado, filtra pela data
            sql_query += " WHERE orcamento = %s"
            cursor.execute(sql_query, (valor,))
        else:  # Se não houver filtro, executa o SELECT sem condições
            cursor.execute(sql_query)

        # Carregar os resultados e exibir na Treeview
        resultados = cursor.fetchall()
        for linha in resultados:
            tree.insert("", tk.END, values=linha)

        # Fechar o cursor e a conexão
        cursor.close()
        conexao.close()

    # Botão para aplicar o filtro
    def aplicar_filtro():
        valor_filtrado = entry_filtro.get().strip()
        if valor_filtrado:  # Verifica se há algum valor inserido para filtrar
            carregar_dados(valor_filtrado)
        else:
            carregar_dados()  # Se não houver valor, carrega todos os dados

    botao_filtrar = ctk.CTkButton(visao_frame, text="Filtrar", command=aplicar_filtro, fg_color="#681919",
                                  hover_color="darkred", width=200, height=40, font=("Impact", 14))
    botao_filtrar.pack(pady=10)

    # Criando a tabela com Treeview
    colunas = ("local", "orcamento","pedido", "data", "situacao", "cliente", "razao", "representante",
               "itens","total", "vencimento","valor_nf", "nf",  "distribuidor")
    tree = ttk.Treeview(visao_frame, columns=colunas, show="headings")
    tree.pack(fill=tk.BOTH, expand=True)

    for coluna in colunas:
        tree.heading(coluna, text=coluna)
        tree.column(coluna, width=50)

    # Carregar todos os dados inicialmente
    carregar_dados()

    # Botão exportar
    def exportar_para_excel():
        caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if caminho_arquivo:
            # Extraindo os dados da Treeview para um DataFrame
            dados = [tree.item(item)["values"] for item in tree.get_children()]
            df = pd.DataFrame(dados, columns=colunas)

            # Usando ExcelWriter para definir a linha inicial de escrita
            with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, startcol=1)  # Dados a partir da segunda linha (índice 1)

            messagebox.showinfo("Exportação", f"Dados exportados com sucesso para {caminho_arquivo}")

    # Configuração dos botões
    botao_exportar = ctk.CTkButton(visao_frame, text="Exportar para Excel", command=exportar_para_excel,
                                   fg_color="#681919", hover_color="darkred", width=200, height=40, font=("Impact", 14))
    botao_exportar.pack(pady=10)

    botao_voltar = ctk.CTkButton(visao_frame, text="Voltar",
                                 command=lambda: (restaurar_menu(), mostrar_frame(menu_frame)),
                                 fg_color="#681919", hover_color="darkred", width=200, height=50, font=("Impact", 16))
    botao_voltar.pack(pady=10)


def importar_dados_excel():
    caminho_excel = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Excel files", "*.xlsx")])
    if not caminho_excel or not os.access(caminho_excel, os.R_OK):
        messagebox.showerror("Erro", "Arquivo inválido ou sem permissão de leitura.")
        return

    try:
        wb = load_workbook(caminho_excel)
        sheet = wb.active

        # Conectar ao banco de dados
        conexao = mysql.connector.connect(
            host='localhost',
            user='root',
            password='',
            database='nimal'
        )
        cursor = conexao.cursor()

        # Percorrer todas as linhas de dados do Excel (começando pela segunda linha)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # Verificar se a linha possui colunas suficientes
            if len(row) < 14:
                continue

            # Atribuir valores com segurança
            local = row[1]
            orcamento = row[2]
            pedido = row[3]
            data = row[4]
            situacao = row[5]
            cliente = row[6]
            razao = row[7]
            representante = row[8]
            itens = row[9]
            total = row[10]
            vencimento = row[11]
            nf = row[13]
            valor_nf = row[12]
            distribuidor = row[14]

            # Verificar se o orçamento já existe no banco de dados
            cursor.execute("SELECT COUNT(*) FROM nimalnotas WHERE orcamento = %s", (orcamento,))
            resultado = cursor.fetchone()

            if resultado and resultado[0] > 0:
                # Atualizar se o orçamento já existe
                sql_update = """
                    UPDATE nimalnotas
                    SET local = %s, pedido = %s, situacao = %s, cliente = %s, razao = %s,
                        representante = %s, itens = %s,total = %s, vencimento = %s,
                        valor_nf = %s,nf = %s, distribuidor = %s, data = %s
                    WHERE orcamento = %s
                """
                cursor.execute(sql_update, (local, pedido, situacao, cliente, razao, representante, itens,total, vencimento,valor_nf,nf, distribuidor, data, orcamento))
            else:
                # Inserir novo registro
                sql_insert = """
                    INSERT INTO nimalnotas (
                        local, orcamento, pedido, situacao, cliente, razao, representante,
                        itens,total, vencimento, nf, valor_nf, distribuidor, data
                    ) VALUES (%s, %s,%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """
                cursor.execute(sql_insert, (local, orcamento, pedido, situacao, cliente, razao, representante, itens,total, vencimento, nf, valor_nf, distribuidor, data))

        conexao.commit()
        cursor.close()
        conexao.close()
        messagebox.showinfo("Sucesso", "Dados do Excel importados com sucesso.")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao importar dados do Excel: {e}")

def mostrar_frame(frame):
    frame.tkraise()

# Configuração da interface principal
janela = ctk.CTk()
janela.title("NimalResearch")
janela.state('zoomed')
janela.config(bg="#DCDCDC")
janela.resizable(False, False)
janela.iconbitmap("nimal.ico")

frame1 = ctk.CTkFrame(janela, fg_color="#C0C0C0", width=1240, height=60,corner_radius=10,background_corner_colors=["#DCDCDC", "#DCDCDC", "#DCDCDC", "#DCDCDC"])
frame1.place(x=20, y=10)

# Frame principal onde alternaremos as telas
frame2 = ctk.CTkFrame(janela, fg_color="#DCDCDC", width=1240, height=820,corner_radius=20,)
frame2.place(x=20, y=100)

# Configuração do frame1 (cabeçalho)
titulo = ctk.CTkLabel(frame1, text="NIMAL RESEARCH", font=("Impact", 40),text_color="#681919")
titulo.place(relx=0.5, rely=0.5, anchor='center')

# Carregar imagens de fundo para os botões (substitua "instrucoes.png" e "pdf.png" com os caminhos reais das imagens)
imagem_instrucoes = ctk.CTkImage(Image.open("instrucoes.png"), size=(60, 60))
imagem_pdf = ctk.CTkImage(Image.open("pdf.png"), size=(60, 60))
imagem_logo = ctk.CTkImage(Image.open("nimal.ico"), size=(90, 90))
imagem_visao = ctk.CTkImage(Image.open("visao.png"), size=(60, 60))
imagem_excel = ctk.CTkImage(Image.open("excel.png"), size=(60, 60))
imagem_passo1= ctk.CTkImage(Image.open("passo1.png"), size=(300, 300))
imagem_passo2= ctk.CTkImage(Image.open("passo2.png"), size=(300, 300))
imagem_passo3= ctk.CTkImage(Image.open("passo3.png"), size=(300, 300))
imagem_passo4= ctk.CTkImage(Image.open("passo4.png"), size=(300, 300))



logo_label = ctk.CTkLabel(frame1, text="", image=imagem_logo, font=("Impact", 30,"bold"))
logo_label.place(x=10, y=-15)

# Frame do menu principal
menu_frame = ctk.CTkFrame(frame2, fg_color="#C0C0C0",corner_radius=20)
menu_frame.place(relwidth=1, relheight=1)

# Frame das instruções
instrucoes_frame = ctk.CTkFrame(frame2, fg_color="#C0C0C0",corner_radius=20)
instrucoes_frame.place(relwidth=1, relheight=1)

# Configuração do menu principal
botao_instrucoes = ctk.CTkButton(menu_frame, text="   Instruções", command=lambda: mostrar_frame(instrucoes_frame),
                                 fg_color="#681919", hover_color="darkred", width=500, height=80,
                                 image=imagem_instrucoes,compound="left", font=("Impact", 18), corner_radius=10,anchor="w")
botao_instrucoes.pack(padx=10, pady=20)

botao_selecionar_pdf = ctk.CTkButton(menu_frame, text="    Selecionar PDF", command=selecionar_pdf,
                                     fg_color="#681919", hover_color="darkred", width=500, height=80,
                                     image=imagem_pdf, font=("Impact", 18), corner_radius=10,anchor="w")
botao_selecionar_pdf.pack(padx=50, pady=20)

botao_visao_geral = ctk.CTkButton(menu_frame, text="     Visão Geral", command=mostrar_visao_geral,
                                     fg_color="#681919", hover_color="darkred", width=500, height=80,
                                     image=imagem_visao, font=("Impact", 18), corner_radius=10,anchor="w")
botao_visao_geral.pack(padx=50, pady=20)

botao_importar = ctk.CTkButton(menu_frame, text="   Importar Excel",command=importar_dados_excel,fg_color="#681919", hover_color="darkred", width=500, height=80,
                                      font=("Impact", 18), corner_radius=10, anchor="w",image=imagem_excel)
botao_importar.pack(padx=50, pady=20)


# Configuração da tela de instruções
titulo_instrucoes = ctk.CTkLabel(instrucoes_frame, text="Instruções",text_color="#681919",font=("Impact", 30))
titulo_instrucoes.place(relx=0.5, y=15, anchor='center')

texto_instrucoes = ctk.CTkLabel(instrucoes_frame, text="Bem-vindo ao Nimal Research! Aqui você pode adicionar automaticamente as informações de notas fiscais à sua planilha no Excel. Para isso, basta seguir estes passos, mas antes, precisamos salvar o arquivo baixado do eloca novamente, para isso basta abrir o arquivo da planilha excel e ir em -> Arquivo -> Salvar Como e escolher um nome para o arquivo.\n 1. Clique em 'Selecionar PDF' e escolha o arquivo da nota fiscal (Lembre-se de que ele precisa estar no formato PDF). Quando o programa ler o arquivo, as informações extraídas aparecerão para o usuário.\n 2. Selecione o arquivo da planilha Excel para adicionar as informações extraídas. \n 3. Na aba de seleção do valor do orçamento, insira o número correspondente à ordem de serviço escolhida. Esse número é indicado na coluna 'Orçamento' e segue o formato padrão: '00-0'. ",
                                font=("Arial", 14,"bold"),text_color="#681919", wraplength=850, justify="left")
texto_instrucoes.place(relx=0.5, y=120, anchor='center')


tutorial1 = ctk.CTkLabel(instrucoes_frame, text="selecionar arquivo da nota fiscal",image=imagem_passo1,font=("Impact", 15))
tutorial1.place(x=170, y=400, anchor='center')

tutorial2 = ctk.CTkLabel(instrucoes_frame, text="confirmar os dados",image=imagem_passo2, font=("Impact", 15))
tutorial2.place(x=470, y=400, anchor='center')

tutorial3 = ctk.CTkLabel(instrucoes_frame, text="selecionar arquivo excel",image=imagem_passo3, font=("Impact", 15))
tutorial3.place(x=770, y=400, anchor='center')

tutorial4 = ctk.CTkLabel(instrucoes_frame, text="escolher o número do orçamento",image=imagem_passo4, font=("Impact", 15))
tutorial4.place(x=1070, y=400, anchor='center')

botao_voltar = ctk.CTkButton(instrucoes_frame, text="Voltar", command=lambda: mostrar_frame(menu_frame),
                             fg_color="#681919", hover_color="darkred", width=200, height=50, font=("Impact", 16))
botao_voltar.place(relx=0.5, y=700, anchor='center')

mostrar_frame(menu_frame)

janela.mainloop()


