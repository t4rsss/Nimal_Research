import mysql.connector
from openpyxl import load_workbook
import os
from PIL import Image
import pandas as pd
import tkinter as tk
from tkinter import ttk
import fitz
import re
import customtkinter as ctk
from tkinter import messagebox
from tkinter import filedialog

conexao = mysql.connector.connect(
    host="",
    user="seu_usuario",
    password="sua_senha",
    database="nimalnotas"
)

cursor = conexao.cursor()

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def mostrar_visao_geral():
    global tree, frame2, menu_frame

    frame2.configure(width=1200, height=650)
    frame2.pack_propagate(False)
    # Limpa o conteúdo anterior no frame de menu
    for widget in menu_frame.winfo_children():
        widget.destroy()

    # Frame para a visão geral
    visao_frame = ctk.CTkFrame(frame2, fg_color="#33415c")
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
            host="",  # ou o IP do servidor MySQL
            user="seu_usuario",  # substitua pelo seu usuário MySQL
            password="sua_senha",  # substitua pela senha do seu usuário
            database="nimalnotas"
        )
        cursor = conexao.cursor()

        if conexao.is_connected():
            print("Conexão bem-sucedida ao MySQL")

        def atualizar_treeview():
            # Limpar a Treeview
            for item in tree.get_children():
                tree.delete(item)

            # Recuperar os dados do banco de dados
            cursor.execute("SELECT * FROM nimal ORDER BY orcamento ASC")
            resultados = cursor.fetchall()

            # Adicionar os dados na Treeview
            for linha in resultados:
                tree.insert("", "end", values=linha)

        # SQL básico com filtro opcional
        sql_query = """
            SELECT local, orcamento,pedido,data, situacao, cliente, razao, representante, itens,total, vencimento, nf, valor_nf, distribuidor 
            FROM nimal
        """
        if valor:  # Se um valor for passado, filtra pela data
            sql_query += " WHERE orcamento LIKE %s "
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
        valor_filtrado = entry_filtro.get().strip() + '%'
        if valor_filtrado:  # Verifica se há algum valor inserido para filtrar
            carregar_dados(valor_filtrado)
        else:
            carregar_dados()  # Se não houver valor, carrega todos os dados

    botao_filtrar = ctk.CTkButton(visao_frame, text="Filtrar", command=aplicar_filtro, fg_color="#001427",
                                  hover_color="#4361ee", width=200, height=40, font=("Impact", 14))
    botao_filtrar.pack(pady=10)

    style = ttk.Style()
    style.theme_use("clam")

    style.configure("Treeview.Heading",
                    background="#33415c",  # Cor de fundo do cabeçalho
                    foreground="white",  # Cor da fonte do cabeçalho
                    font=("Arial", 10, "bold"))  # Fonte do cabeçalho

    style.map("Treeview.Heading",
              background=[("active", "#4361ee")])  # Cor de fundo ao passar o mouse



    style.configure("Treeview", background="#1C1C1C",  # Cor de fundo
                    fieldbackground="#1C1C1C",  # Cor de fundo dos campos
                    foreground="white",  # Cor do texto
                    rowheight=25)  # Altura das linhas

    # Criando a tabela com Treeview
    colunas = ("local", "orcamento", "pedido", "data", "fase", "cliente", "razao", "representante",
               "itens", "total", "vencimento", "nf", "valor_nf", "distribuidor")
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
                df.to_excel(writer, index=False, startcol=1)

            messagebox.showinfo("Exportação", f"Dados exportados com sucesso para {caminho_arquivo}")

    # Configuração dos botões
    botao_exportar = ctk.CTkButton(visao_frame, text="Gerar Relatório", command=exportar_para_excel,
                                   fg_color="#001427", hover_color="#4361ee", width=200, height=50, font=("Impact", 16))
    botao_exportar.pack(pady=5, padx=5, side=tk.LEFT)

    botao_selecionar_pdf = ctk.CTkButton(visao_frame, text="Adicionar Nota Fiscal", command=selecionar_pdf,
                                         fg_color="#001427", hover_color="#4361ee", width=200, height=50,
                                         image=imagem_pdf, font=("Impact", 16), corner_radius=10, anchor="w")
    botao_selecionar_pdf.pack(pady=5, padx=5, side=tk.LEFT)

    def duplicar_orcamento():
        try:
            global orcamento_selecionado  # Declarar como global para que seja acessível dentro da função

            # Obtendo o orcamento selecionado da Treeview
            orcamento_selecionado = tree.item(tree.selection())["values"][1]

            # Incrementar o valor do orçamento para criar um novo registro
            base_orcamento = orcamento_selecionado.rsplit('-', 1)[0]
            numero_incremento = int(orcamento_selecionado.rsplit('-', 1)[1]) + 1
            novo_orcamento = f"{base_orcamento}-{numero_incremento}"

            cursor.execute(
                "SELECT local, orcamento,pedido,data, situacao, cliente, razao, representante, itens,total, vencimento, nf, valor_nf, distribuidor FROM nimal WHERE orcamento = %s",
                (orcamento_selecionado,))
            linha_selecionada = cursor.fetchone()

            print(linha_selecionada)

            # Copiar a linha selecionada e modificar o orçamento
            dados_copia = [
                linha_selecionada[0],  # local
                novo_orcamento,  # orcamento
                linha_selecionada[2],  # pedido
                linha_selecionada[3],  # data
                linha_selecionada[4],  # situacao
                linha_selecionada[5],  # cliente
                linha_selecionada[6],  # razao
                linha_selecionada[7],  # representante
                linha_selecionada[8],  # itens
                linha_selecionada[9],  # total
                linha_selecionada[10],  # vencimento
                linha_selecionada[11],  # nf
                linha_selecionada[12],  # valor_nf
                linha_selecionada[13]  # distribuidor
            ]

            print(dados_copia)

            # Inserir a nova linha na tabela no banco de dados
            cursor.execute(
                "INSERT INTO nimal (local,orcamento, pedido, data, situacao, cliente, razao, representante, itens, total, vencimento, nf, valor_nf, distribuidor) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                dados_copia
            )
            conexao.commit()

            mostrar_visao_geral()
        except Exception as e:
            messagebox.showerror("Erro", "Por favor,selecione um pedido para duplicar.")

    # Botão para duplicar orçamento
    duplicar_button = ctk.CTkButton(visao_frame, text="Duplicar Pedido", command=duplicar_orcamento,
                                    fg_color="#001427", hover_color="#4361ee", width=200, height=50,
                                    font=("Impact", 16))
    duplicar_button.pack(pady=5, padx=5, side=tk.LEFT)

    botao_voltar = ctk.CTkButton(visao_frame, text="Voltar",
                                 command=lambda: (restaurar_menu(), mostrar_frame(menu_frame)),
                                 fg_color="#001427", hover_color="#4361ee", width=200, height=50, font=("Impact", 16))
    botao_voltar.pack(pady=5, padx=5, side=tk.RIGHT)

    def editar_dados():
        # Verifica se há uma linha selecionada
        item_selecionado = tree.selection()
        if not item_selecionado:
            messagebox.showwarning("Aviso", "Por favor, selecione uma linha para editar.")
            return

        # Recupera os valores da linha selecionada
        valores_atuais = tree.item(item_selecionado, "values")

        # Cria uma nova janela para edição
        janela_edicao = ctk.CTkToplevel()
        janela_edicao.title("Editar Dados")
        janela_edicao.geometry("500x700")

        # Labels e entradas para edição dos valores
        campos = [
            "local", "orcamento", "pedido", "data", "situacao", "cliente", "razao",
            "representante", "itens", "total", "vencimento", "nf", "valor_nf", "distribuidor"
        ]

        entradas = {}
        for i, campo in enumerate(campos):
            frame_linha = ctk.CTkFrame(janela_edicao)
            frame_linha.pack(fill=tk.X, padx=10, pady=5)

            ctk.CTkLabel(frame_linha, text=campo.capitalize(), font=("Arial", 12), width=15, anchor="w").pack(
                side=tk.LEFT, padx=5)
            entrada = ctk.CTkEntry(frame_linha, font=("Arial", 12))
            entrada.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)
            entrada.insert(0, valores_atuais[i])  # Preenche com o valor atual
            entradas[campo] = entrada

        # Função para salvar as edições
        def confirmar_edicoes():
            novos_valores = [entradas[campo].get().strip() for campo in campos]

            # Atualiza os valores na Treeview
            tree.item(item_selecionado, values=novos_valores)

            # Atualiza o banco de dados
            conexao = mysql.connector.connect(
                host="",
                user="seu_usuario",
                password="sua_senha",
                database="nimalnotas"
            )
            cursor = conexao.cursor()

            try:
                sql_update = """
                    UPDATE nimal
                    SET local = %s, orcamento = %s, pedido = %s, data = %s, situacao = %s, cliente = %s, 
                        razao = %s, representante = %s, itens = %s, total = %s, vencimento = %s, 
                        nf = %s, valor_nf = %s, distribuidor = %s
                    WHERE orcamento = %s
                """
                cursor.execute(sql_update,
                               (*novos_valores, valores_atuais[1]))  # Usa o orçamento como identificador único
                conexao.commit()
                messagebox.showinfo("Sucesso", "Dados atualizados com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao atualizar dados: {e}")
            finally:
                cursor.close()
                conexao.close()

            janela_edicao.destroy()

        # Botão para confirmar as alterações
        ctk.CTkButton(
            janela_edicao, text="Confirmar", command=confirmar_edicoes,
            fg_color="#001427", hover_color="#4361ee", width=200, height=40, font=("Impact", 14)
        ).pack(pady=20)

        # Botão para cancelar a edição
        ctk.CTkButton(
            janela_edicao, text="Cancelar", command=janela_edicao.destroy,
            fg_color="#001427", hover_color="#e63946", width=200, height=40, font=("Impact", 14)
        ).pack(pady=10)

    def remover_dados():
        global orcamento_selecionado

        orcamento_selecionado = tree.item(tree.selection())["values"][1]

        resposta = messagebox.askyesno("Confirmação",
                                       f"Tem certeza que deseja remover o pedido {orcamento_selecionado} do banco de dados?")
        if resposta:
            try:
                conexao = mysql.connector.connect(
                    host="",
                    user="seu_usuario",
                    password="sua_senha",
                    database="nimalnotas"
                )
                cursor = conexao.cursor()

                # Remover a linha selecionada
                sql_delete = "DELETE FROM nimal WHERE orcamento = %s"
                cursor.execute(sql_delete, (orcamento_selecionado,))

                conexao.commit()
                cursor.close()
                conexao.close()

                # Recarregar a visão geral
                mostrar_visao_geral()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao remover dados: {e}")

    botao_editar = ctk.CTkButton(visao_frame, text="Editar", command=editar_dados, fg_color="#1C1C1C",
                                 hover_color="#696969", width=100, height=50, font=("Impact", 16))
    botao_editar.pack(pady=5, padx=5, side=tk.LEFT)

    botao_remover = ctk.CTkButton(visao_frame, text="Remover", command=remover_dados, fg_color="#8B0000",
                                  hover_color="#A52A2A", width=100, height=50, font=("Impact", 16))
    botao_remover.pack(pady=5, padx=5, side=tk.LEFT)


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
            padroes_fase = ""

            # Função auxiliar para tentar vários padrões até encontrar uma correspondência
            def buscar_texto(padroes):
                for padrao in padroes:
                    match = re.search(padrao, conteudo)
                    if match:
                        return match.group(1).strip()
                return ""

            # Usar a função auxiliar para buscar cada campo
            venc_extraido = buscar_texto(padroes_venc)
            nf_extraido = buscar_texto(padroes_nf)
            dist_extraido = buscar_texto(padroes_dist)
            valor_extraido = buscar_texto(padroes_valor)
            fase_extraido = buscar_texto(padroes_fase)

            # Chamar função para editar e confirmar dados
            editar_dados(venc_extraido, nf_extraido, dist_extraido, valor_extraido, fase_extraido)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o PDF: {e}")


def editar_dados(venc, nf, dist, valor, fase):
    global orcamento_selecionado

    orcamento_selecionado = tree.item(tree.selection())["values"][1]


    janela_edicao = ctk.CTkToplevel()
    janela_edicao.geometry("500x500")
    janela_edicao.title("Editar Dados")

    # Campos para edição
    ctk.CTkLabel(janela_edicao, text="Vencimento:").grid(row=0, column=0, padx=10, pady=5)
    entrada_venc = ctk.CTkEntry(janela_edicao, width=300)
    entrada_venc.insert(0, venc)
    entrada_venc.grid(row=0, column=1, padx=10, pady=5)

    ctk.CTkLabel(janela_edicao, text="Nº da NF:").grid(row=1, column=0, padx=10, pady=5)
    entrada_nf = ctk.CTkEntry(janela_edicao, width=300)
    entrada_nf.insert(0, nf)
    entrada_nf.grid(row=1, column=1, padx=10, pady=5)

    ctk.CTkLabel(janela_edicao, text="Distribuidor:").grid(row=2, column=0, padx=10, pady=5)
    entrada_dist = ctk.CTkEntry(janela_edicao, width=300)
    entrada_dist.insert(0, dist)
    entrada_dist.grid(row=2, column=1, padx=10, pady=5)

    ctk.CTkLabel(janela_edicao, text="Valor:").grid(row=3, column=0, padx=10, pady=5)
    entrada_valor = ctk.CTkEntry(janela_edicao, width=300)
    entrada_valor.insert(0, valor)
    entrada_valor.grid(row=3, column=1, padx=10, pady=5)

    ctk.CTkLabel(janela_edicao, text="Fase:").grid(row=4, column=0, padx=10, pady=5)
    entrada_fase = ctk.CTkEntry(janela_edicao, width=300)
    entrada_fase.insert(0, fase)
    entrada_fase.grid(row=4, column=1, padx=10, pady=5)

    # Botão de confirmação para salvar as alterações
    def confirmar_edicao():
        # Conectar ao banco de dados
        conexao = mysql.connector.connect(
            host="",
            user="seu_usuario",
            password="sua_senha",
            database="nimalnotas"
        )
        cursor = conexao.cursor()

        # Atualizar os dados no banco de dados
        sql_update = """
            UPDATE nimal
            SET vencimento = %s, nf = %s, distribuidor = %s, valor_nf = %s , situacao = %s
            WHERE orcamento = %s
        """
        cursor.execute(sql_update, (
            entrada_venc.get(), entrada_nf.get(), entrada_dist.get(), entrada_valor.get(), entrada_fase.get(),
            orcamento_selecionado))

        conexao.commit()
        cursor.close()
        conexao.close()

        # Fechar a janela de edição
        janela_edicao.destroy()

        # Recarregar a visão geral
        mostrar_visao_geral()

    botao_confirmar = ctk.CTkButton(janela_edicao, text="Confirmar", command=confirmar_edicao, fg_color="#001427",
                                    hover_color="#4361ee", width=100, height=40, font=("Impact", 14))
    botao_confirmar.grid(pady=10)


def selecionar_pdf():
    item_selecionado = tree.selection()
    if not item_selecionado:
        messagebox.showwarning("Aviso", "Por favor, selecione um pedido para adicionar a nota fiscal.")
        return

    caminho_pdf = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    if caminho_pdf:
        extrair_informacoes_pdf(caminho_pdf)


def restaurar_menu():
    frame2.configure(width=600, height=480)
    frame2.pack_propagate(False)
    # Limpa o conteúdo anterior no menu
    for widget in menu_frame.winfo_children():
        widget.destroy()

    # Recria os botões do menu

    botao_visao_geral = ctk.CTkButton(menu_frame, text="     Visão Geral", command=mostrar_visao_geral,
                                      fg_color="#001427", hover_color="#4361ee", width=500, height=100,
                                      image=imagem_visao, font=("Impact", 18), corner_radius=10, anchor="w")
    botao_visao_geral.pack(padx=50, pady=75)

    botao_importar = ctk.CTkButton(menu_frame, text="   Importar Excel", command=importar_dados_excel,
                                   fg_color="#001427",
                                   hover_color="#4361ee", width=500, height=100,
                                   font=("Impact", 18), corner_radius=10, anchor="w", image=imagem_excel)
    botao_importar.pack(padx=50, pady=45)


def importar_dados_excel():
    caminho_excel = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Excel files", "*.xlsx")])
    if not caminho_excel or not os.access(caminho_excel, os.R_OK):
        messagebox.showerror("Erro", "Arquivo inválido ou sem permissão de leitura.")
        return

    try:
        wb = load_workbook(caminho_excel)
        sheet = wb.active

        # Verificar e adicionar colunas necessárias
        if 'vencimento' not in sheet.cell(row=1, column=sheet.max_column).value:
            sheet.cell(row=1, column=sheet.max_column + 1, value="vencimento")
        if 'nf' not in sheet.cell(row=1, column=sheet.max_column).value:
            sheet.cell(row=1, column=sheet.max_column + 1, value="nf")
        if 'valor_nf' not in sheet.cell(row=1, column=sheet.max_column).value:
            sheet.cell(row=1, column=sheet.max_column + 1, value="valor_nf")
        if 'distribuidor' not in sheet.cell(row=1, column=sheet.max_column).value:
            sheet.cell(row=1, column=sheet.max_column + 1, value="distribuidor")

        # Conectar ao banco de dados
        conexao = mysql.connector.connect(
            host="",  # ou o IP do servidor MySQL
            user="seu_usuario",  # substitua pelo seu usuário MySQL
            password="sua_senha",  # substitua pela senha do seu usuário
            database="nimalnotas"
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
            valor_nf = row[12]
            nf = row[13]
            distribuidor = row[14]

            # Verificar se o orçamento já existe no banco de dados
            cursor.execute("SELECT COUNT(*) FROM nimal WHERE orcamento = %s ", (orcamento,))
            resultado = cursor.fetchone()

            if resultado and resultado[0] > 0:
                # Atualizar se o orçamento já existe
                sql_update = """
                    UPDATE nimal
                    SET local = %s, pedido = %s, situacao = %s, cliente = %s, razao = %s,
                        representante = %s, itens = %s, total = %s, vencimento = %s,
                        valor_nf = %s, nf = %s, distribuidor = %s, data = %s
                    WHERE orcamento = %s
                """
                cursor.execute(sql_update, (
                    local, pedido, situacao, cliente, razao, representante, itens, total, vencimento, valor_nf, nf,
                    distribuidor, data, orcamento))
            else:
                # Inserir novo registro
                sql_insert = """
                    INSERT INTO nimal (
                        local, orcamento, pedido, situacao, cliente, razao, representante,
                        itens, total, vencimento, nf, valor_nf, distribuidor, data
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """
                cursor.execute(sql_insert, (
                    local, orcamento, pedido, situacao, cliente, razao, representante, itens, total, vencimento, nf,
                    valor_nf, distribuidor, data))

        conexao.commit()
        cursor.close()
        conexao.close()
        messagebox.showinfo("Sucesso", "Dados do Excel importados com sucesso.")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao importar dados do Excel: {e}")


def mostrar_frame(frame):
    if frame == menu_frame:
        frame2.configure(width=600, height=480)
        frame2.pack_propagate(False)
    else:
        frame2.configure(width=600, height=480)  # Exemplo de tamanho padrão
        frame2.pack_propagate(True)  # Reativa o ajuste automático de tamanho

    frame.tkraise()


def get_screen_size():
    screen_width = janela.winfo_screenwidth()
    screen_height = janela.winfo_screenheight()
    return screen_width, screen_height


janela = ctk.CTk()
screen_width, screen_height = get_screen_size()
janela.geometry(f"{screen_width}x{screen_height}")
janela.resizable(True, True)
janela.title("NimalResearch")
janela.state('zoomed')
janela.iconbitmap("nimal.ico")
bg = ctk.CTkImage(Image.open("bg10.png"), size=(1920, 1080))
bg_label = ctk.CTkLabel(janela, image=bg)
bg_label.place(relwidth=1, relheight=1)


imagem_instrucoes = ctk.CTkImage(Image.open("instrucoes.png"), size=(60, 60))
imagem_pdf = ctk.CTkImage(Image.open("pdf.png"), size=(30, 30))
imagem_logo = ctk.CTkImage(Image.open("nimal.ico"), size=(80, 80))
imagem_visao = ctk.CTkImage(Image.open("visao.png"), size=(60, 60))
imagem_excel = ctk.CTkImage(Image.open("excel.png"), size=(60, 60))


frame1 = ctk.CTkFrame(janela, fg_color="#33415c", width=600, height=90, corner_radius=10,background_corner_colors=["#100f3e", "#27269a", "#27269a", "#100f3e"])
frame1.place(relx=0.5, rely=0.09, anchor='center')

frame2 = ctk.CTkFrame(janela, fg_color="#33415c", width=600, height=280, corner_radius=10)
frame2.place(relx=0.5, rely=0.5, anchor='center')

titulo = ctk.CTkLabel(frame1, text="NIMAL RESEARCH", font=("Impact", 40))
titulo.place(relx=0.5, rely=0.5, anchor='center')

menu_frame = ctk.CTkFrame(frame2, fg_color="#33415c", border_color="#33415c", border_width=3, corner_radius=10,background_corner_colors=["#100f3e", "#27269a", "black", "#100f3e"])
menu_frame.place(relwidth=1, relheight=1)

logo_label = ctk.CTkLabel(frame1, text="", image=imagem_logo, font=("Impact", 30, "bold"))
logo_label.place(x=20, y=5)

botao_visao_geral = ctk.CTkButton(menu_frame, text="     Visão Geral", command=mostrar_visao_geral,fg_color="#001427", hover_color="#4361ee", width=500, height=100,image=imagem_visao, font=("Impact", 18), corner_radius=10, anchor="w")
botao_visao_geral.pack(padx=50, pady=75)

botao_importar = ctk.CTkButton(menu_frame, text="   Importar Excel", command=importar_dados_excel, fg_color="#001427",hover_color="#4361ee", width=500, height=100,font=("Impact", 18), corner_radius=10, anchor="w", image=imagem_excel)
botao_importar.pack(padx=50, pady=45)

mostrar_frame(menu_frame)

janela.mainloop()
