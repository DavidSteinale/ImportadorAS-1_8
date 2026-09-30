import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from zipfile import BadZipFile

from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import psycopg
from psycopg import sql
# import xlwt  # Exportação XLS desativada temporariamente.

load_dotenv()


class ImportadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Importador")
        self.root.geometry("600x340")
        self.root.minsize(460, 260)
        self.configuracao = None
        self.status = tk.StringVar(value="Configure os dados do PostgreSQL para começar.")

        self._criar_menu()

        conteudo = ttk.Frame(root, padding=24)
        conteudo.pack(fill="both", expand=True)
        ttk.Label(conteudo, text="Importador", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(
            conteudo,
            text="Use o menu Banco de dados para configurar ou testar a conexão.",
        ).pack(anchor="w", pady=(8, 20))
        ttk.Separator(conteudo).pack(fill="x", pady=(0, 14))
        ttk.Label(conteudo, textvariable=self.status, wraplength=520).pack(anchor="w")

    def _criar_menu(self):
        menu = tk.Menu(self.root)
        banco_menu = tk.Menu(menu, tearoff=False)
        banco_menu.add_command(label="Configurar conexão...", command=self.abrir_configuracao)
        banco_menu.add_command(label="Testar conexão", command=self.testar_conexao_salva)
        banco_menu.add_separator()
        banco_menu.add_command(label="Executar consulta SELECT...", command=self.abrir_consulta)
        menu.add_cascade(label="Banco de dados", menu=banco_menu)

        importar_menu = tk.Menu(menu, tearoff=False)
        importar_menu.add_command(
            label="Importar produtos por XLSX...", command=self.importar_produtos_xlsx
        )
        menu.add_cascade(label="Importar", menu=importar_menu)

        ajuda_menu = tk.Menu(menu, tearoff=False)
        ajuda_menu.add_command(
            label="Sobre",
            command=lambda: messagebox.showinfo(
                "Sobre", "Importador com conexão PostgreSQL.", parent=self.root
            ),
        )
        menu.add_cascade(label="Ajuda", menu=ajuda_menu)
        self.root.config(menu=menu)

    def importar_produtos_xlsx(self):
        if not self._exigir_configuracao():
            return

        caminho = filedialog.askopenfilename(
            parent=self.root,
            title="Selecionar planilha XLSX de produtos",
            filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
        )
        if not caminho:
            return

        try:
            workbook = load_workbook(caminho, read_only=True, data_only=True)
            try:
                planilha = workbook.active
                cabecalho = next(planilha.iter_rows(values_only=True), None)
            finally:
                workbook.close()
        except (OSError, BadZipFile, InvalidFileException, ValueError) as erro:
            messagebox.showerror("Erro ao ler XLSX", str(erro), parent=self.root)
            return

        if not cabecalho:
            messagebox.showerror(
                "Planilha vazia",
                "A aba ativa não contém uma linha de cabeçalho.",
                parent=self.root,
            )
            return

        colunas = [str(coluna).strip() if coluna is not None else "" for coluna in cabecalho]
        if any(not coluna for coluna in colunas):
            messagebox.showerror(
                "Cabeçalho inválido",
                "Todas as células do cabeçalho precisam ter um nome de coluna.",
                parent=self.root,
            )
            return
        if len(set(colunas)) != len(colunas):
            messagebox.showerror(
                "Cabeçalho inválido",
                "O cabeçalho contém nomes de coluna duplicados.",
                parent=self.root,
            )
            return

        nomes_colunas = ", ".join(colunas)
        confirmar = messagebox.askyesno(
            "Confirmar importação",
            f"Colunas lidas da aba ativa:\n{nomes_colunas}\n\n"
            "Inserir os registros na tabela 'produto'?",
            parent=self.root,
        )
        if not confirmar:
            return

        consulta_insert = sql.SQL("INSERT INTO {} ({}) VALUES ({})").format(
            sql.Identifier("produto"),
            sql.SQL(", ").join(sql.Identifier(coluna) for coluna in colunas),
            sql.SQL(", ").join(sql.Placeholder() for _ in colunas),
        )
        linhas_importadas = 0

        def ler_linhas():
            nonlocal linhas_importadas
            workbook = load_workbook(caminho, read_only=True, data_only=True)
            try:
                planilha = workbook.active
                for numero_linha, linha in enumerate(
                    planilha.iter_rows(min_row=2, values_only=True), start=2
                ):
                    valores = list(linha)
                    if all(valor is None for valor in valores):
                        continue
                    if len(valores) > len(colunas) and any(
                        valor is not None for valor in valores[len(colunas):]
                    ):
                        raise ValueError(
                            f"A linha {numero_linha} contém valores além das colunas do cabeçalho."
                        )
                    valores = valores[:len(colunas)]
                    valores.extend([None] * (len(colunas) - len(valores)))
                    linhas_importadas += 1
                    yield valores
            finally:
                workbook.close()

        try:
            with psycopg.connect(**self.configuracao, connect_timeout=5) as conexao:
                with conexao.cursor() as cursor:
                    cursor.executemany(consulta_insert, ler_linhas())
        except (psycopg.Error, OSError, BadZipFile, InvalidFileException, ValueError) as erro:
            self.status.set("Falha ao importar o XLSX.")
            messagebox.showerror("Erro na importação", str(erro), parent=self.root)
            return

        self.status.set(f"Importação concluída: {linhas_importadas} produto(s) inserido(s).")
        messagebox.showinfo(
            "Importação concluída",
            f"{linhas_importadas} produto(s) inserido(s) na tabela 'produto'.",
            parent=self.root,
        )

    def abrir_configuracao(self):
        janela = tk.Toplevel(self.root)
        janela.title("Configurar conexão PostgreSQL")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()

        campos = (
            ("Servidor", "host"),
            ("Porta", "port"),
            ("Banco de dados", "dbname"),
            ("Usuário", "user"),
            ("Senha", "password"),
        )
        valores_atuais = self.configuracao or {
            "host": os.getenv("PGHOST", "localhost"),
            "port": os.getenv("PGPORT", "5432"),
            "dbname": os.getenv("PGDATABASE", "autosystem"),
            "user": os.getenv("PGUSER", "postgres"),
            "password": os.getenv("PGPASSWORD", ""),
        }
        variaveis = {
            chave: tk.StringVar(value=str(valores_atuais[chave]))
            for _, chave in campos
        }

        formulario = ttk.Frame(janela, padding=16)
        formulario.grid(row=0, column=0, sticky="nsew")
        for linha, (rotulo, chave) in enumerate(campos):
            ttk.Label(formulario, text=rotulo).grid(
                row=linha, column=0, sticky="w", padx=(0, 12), pady=5
            )
            entrada = ttk.Entry(
                formulario,
                textvariable=variaveis[chave],
                width=34,
                show="*" if chave == "password" else "",
            )
            entrada.grid(row=linha, column=1, sticky="ew", pady=5)

        def ler_configuracao():
            valores = {chave: variavel.get().strip() for chave, variavel in variaveis.items()}
            try:
                valores["port"] = int(valores["port"])
                if not 1 <= valores["port"] <= 65535:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Porta inválida", "Informe uma porta entre 1 e 65535.", parent=janela)
                return None

            if (
                not valores["host"]
                or not valores["dbname"]
                or not valores["user"]
                or not valores["password"]
            ):
                messagebox.showerror(
                    "Dados incompletos",
                    "Preencha servidor, banco de dados, usuário e senha.",
                    parent=janela,
                )
                return None
            return valores

        def testar():
            valores = ler_configuracao()
            if valores is not None:
                self._testar_configuracao(valores, janela)

        def aplicar():
            valores = ler_configuracao()
            if valores is not None:
                self.configuracao = valores
                self.status.set(
                    f"Configuração aplicada para {valores['host']}:{valores['port']}/{valores['dbname']} "
                    "nesta sessão."
                )
                janela.destroy()

        botoes = ttk.Frame(formulario)
        botoes.grid(row=len(campos), column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(botoes, text="Testar conexão", command=testar).pack(side="left", padx=(0, 8))
        ttk.Button(botoes, text="Cancelar", command=janela.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(botoes, text="Aplicar", command=aplicar).pack(side="left")

    def testar_conexao_salva(self):
        if not self._exigir_configuracao():
            return
        self._testar_configuracao(self.configuracao, self.root)

    def _exigir_configuracao(self):
        if self.configuracao is not None:
            return True

        messagebox.showinfo(
            "Configuração necessária",
            "Informe e aplique os dados de acesso ao PostgreSQL para continuar.",
            parent=self.root,
        )
        self.abrir_configuracao()
        return False

    def _testar_configuracao(self, configuracao, parent):
        try:
            with psycopg.connect(**configuracao, connect_timeout=5) as conexao:
                with conexao.cursor() as cursor:
                    cursor.execute("SELECT 1")
                    cursor.fetchone()
        except psycopg.Error as erro:
            self.status.set("Falha no teste de conexão.")
            messagebox.showerror("Falha na conexão", str(erro), parent=parent)
            return False

        self.status.set(
            f"Conexão testada com sucesso: {configuracao['host']}:{configuracao['port']}/"
            f"{configuracao['dbname']}"
        )
        messagebox.showinfo(
            "Conexão OK", "Conexão com o PostgreSQL realizada com sucesso.", parent=parent
        )
        return True

    def abrir_consulta(self):
        if not self._exigir_configuracao():
            return

        janela = tk.Toplevel(self.root)
        janela.title("Consulta SELECT")
        janela.geometry("900x600")
        janela.minsize(640, 420)
        janela.transient(self.root)

        conteudo = ttk.Frame(janela, padding=12)
        conteudo.pack(fill="both", expand=True)
        ttk.Label(conteudo, text="Comando SQL (somente SELECT):").pack(anchor="w")

        editor = scrolledtext.ScrolledText(conteudo, height=7, wrap="none")
        editor.pack(fill="x", pady=(6, 8))
        editor.insert("1.0", "SELECT * FROM sua_tabela LIMIT 100;")

        barra = ttk.Frame(conteudo)
        barra.pack(fill="x", pady=(0, 8))
        status = tk.StringVar(value="Informe o SELECT e clique em Executar.")
        # Exportação XLS temporariamente desativada:
        # resultado = {"colunas": [], "linhas": []}
        # botao_salvar = ttk.Button(
        #     barra,
        #     text="Salvar resultado XLS...",
        #     command=lambda: self.salvar_resultado_xls(resultado, janela),
        #     state="disabled",
        # )
        # botao_salvar.pack(side="right")
        ttk.Button(
            barra,
            text="Executar",
            command=lambda: self.executar_select(editor, tabela, status, janela),
        ).pack(side="left")
        ttk.Label(barra, textvariable=status).pack(side="left", padx=12)

        tabela = ttk.Treeview(conteudo, show="headings")
        barra_vertical = ttk.Scrollbar(conteudo, orient="vertical", command=tabela.yview)
        barra_horizontal = ttk.Scrollbar(conteudo, orient="horizontal", command=tabela.xview)
        tabela.configure(yscrollcommand=barra_vertical.set, xscrollcommand=barra_horizontal.set)
        tabela.pack(side="top", fill="both", expand=True)
        barra_vertical.pack(side="right", fill="y")
        barra_horizontal.pack(side="bottom", fill="x")

    def executar_select(self, editor, tabela, status, parent):
        consulta = editor.get("1.0", "end").strip()
        if not consulta:
            messagebox.showwarning("Consulta vazia", "Digite um comando SELECT.", parent=parent)
            return
        if not consulta.lstrip().lower().startswith("select"):
            messagebox.showerror(
                "Comando não permitido",
                "Este menu aceita somente comandos SELECT.",
                parent=parent,
            )
            return

        tabela.delete(*tabela.get_children())

        try:
            with psycopg.connect(**self.configuracao, connect_timeout=5) as conexao:
                with conexao.cursor() as cursor:
                    cursor.execute("SET TRANSACTION READ ONLY")

                    # Exemplo de bloco de comando para consulta direta no código:
                    # cursor.execute("SELECT id, nome FROM sua_tabela LIMIT 20")
                    # resultado = cursor.fetchall()

                    cursor.execute(consulta)
                    if cursor.description is None:
                        messagebox.showerror(
                            "Consulta inválida",
                            "O comando não retornou colunas. Informe um SELECT.",
                            parent=parent,
                        )
                        return

                    nomes_colunas = [coluna.name for coluna in cursor.description]
                    linhas = cursor.fetchmany(1001)
        except psycopg.Error as erro:
            status.set("Falha ao executar a consulta.")
            messagebox.showerror("Erro na consulta", str(erro), parent=parent)
            return

        tabela.delete(*tabela.get_children())
        identificadores = [f"coluna_{indice}" for indice in range(len(nomes_colunas))]
        tabela.configure(columns=identificadores)
        for identificador, nome in zip(identificadores, nomes_colunas):
            tabela.heading(identificador, text=nome)
            tabela.column(identificador, width=150, minwidth=80, stretch=True)

        limite = 1000
        for linha in linhas[:limite]:
            tabela.insert("", "end", values=linha)

        quantidade = min(len(linhas), limite)
        complemento = " (exibindo as primeiras 1000)" if len(linhas) > limite else ""
        status.set(f"{quantidade} linha(s) retornada(s){complemento}.")

    # Exportação XLS temporariamente desativada:
    # def salvar_resultado_xls(self, resultado, parent):
    #     if not resultado["colunas"]:
    #         messagebox.showwarning(
    #             "Sem resultado", "Execute um SELECT antes de salvar.", parent=parent
    #         )
    #         return
    #
    #     caminho = filedialog.asksaveasfilename(
    #         parent=parent,
    #         title="Salvar resultado da consulta",
    #         initialfile="resultado_consulta.xls",
    #         defaultextension=".xls",
    #         filetypes=[("Planilha Excel 97-2003", "*.xls")],
    #     )
    #     if not caminho:
    #         return
    #
    #     try:
    #         pasta_trabalho = xlwt.Workbook(encoding="utf-8")
    #         planilha = pasta_trabalho.add_sheet("Resultado")
    #         estilo_cabecalho = xlwt.easyxf("font: bold on")
    #
    #         for coluna, nome in enumerate(resultado["colunas"]):
    #             planilha.write(0, coluna, nome, estilo_cabecalho)
    #
    #         for indice_linha, linha in enumerate(resultado["linhas"], start=1):
    #             for indice_coluna, valor in enumerate(linha):
    #                 if valor is None:
    #                     valor = ""
    #                 elif not isinstance(valor, (str, int, float, bool)):
    #                     valor = str(valor)
    #                 planilha.write(indice_linha, indice_coluna, valor)
    #
    #         pasta_trabalho.save(caminho)
    #     except Exception as erro:
    #         messagebox.showerror("Erro ao salvar XLS", str(erro), parent=parent)
    #         return
    #
    #     messagebox.showinfo(
    #         "Arquivo salvo", f"Resultado salvo em:\n{caminho}", parent=parent
    #     )


if __name__ == "__main__":
    janela_principal = tk.Tk()
    app = ImportadorApp(janela_principal)
    janela_principal.mainloop()