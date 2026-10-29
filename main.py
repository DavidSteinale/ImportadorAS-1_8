import csv
import hmac
import os
import re
import logging
import sys
import tkinter as tk
import unicodedata
from contextlib import closing
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk
from zipfile import BadZipFile

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import psycopg
import pyodbc
from psycopg import sql
# import xlwt  # Exportação XLS desativada temporariamente.

# Usa a pasta do executável empacotado ou, durante o desenvolvimento, a pasta deste script.
PASTA_APLICATIVO = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
ARQUIVO_LOG = PASTA_APLICATIVO / "importador.txt"
ARQUIVO_RELATORIO_VALIDACAO = PASTA_APLICATIVO / "relatorio_validacao_importacao.txt"
ARQUIVO_RELATORIO_CPFS_DUPLICADOS = PASTA_APLICATIVO / "relatorio_cpfs_duplicados.xlsx"

LOGGER = logging.getLogger("importador")
LOGGER.setLevel(logging.INFO)
if not LOGGER.handlers:
    manipulador_log = logging.FileHandler(ARQUIVO_LOG, encoding="utf-8")
    formato_log = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    manipulador_log.setFormatter(formato_log)
    LOGGER.addHandler(manipulador_log)
    LOGGER.propagate = False

# Carrega as variáveis locais do arquivo .env sem versionar credenciais.
load_dotenv()

# Início do bloco de login e centralização das janelas.
USUARIO_LOGIN = "master"


def senha_login_atual():
    # Senha = hora + minuto do relógio do computador no momento da conferência.
    agora = datetime.now()
    return str(agora.hour + agora.minute)


def centralizar_janela(janela, largura, altura):
    posicao_x = (janela.winfo_screenwidth() - largura) // 2
    posicao_y = (janela.winfo_screenheight() - altura) // 2
    janela.geometry(f"{largura}x{altura}+{posicao_x}+{posicao_y}")


def solicitar_login():
    # Exibe a tela de login e retorna o usuário autenticado, ou None se a janela for fechada.
    janela = tk.Tk()
    janela.title("Login - Importador")
    janela.resizable(False, False)
    resultado = {"usuario": None}
    senha = tk.StringVar()
    mensagem = tk.StringVar()

    def entrar(_evento=None):
        if hmac.compare_digest(
            senha.get().strip().encode("utf-8"), senha_login_atual().encode("utf-8")
        ):
            resultado["usuario"] = USUARIO_LOGIN
            LOGGER.info("Login realizado: usuario=%s.", USUARIO_LOGIN)
            janela.destroy()
            return

        LOGGER.warning("Falha de login: usuario=%s.", USUARIO_LOGIN)
        senha.set("")
        mensagem.set("Senha inválida.")

    formulario = ttk.Frame(janela, padding=20)
    formulario.pack()
    ttk.Label(formulario, text="Senha").grid(row=0, column=0, sticky="w", pady=4)
    campo_senha = ttk.Entry(formulario, textvariable=senha, show="*", width=24)
    campo_senha.grid(row=0, column=1, padx=(10, 0), pady=4)
    ttk.Label(formulario, textvariable=mensagem, foreground="#b00020").grid(
        row=1, column=0, columnspan=2, pady=(4, 0)
    )
    ttk.Button(formulario, text="Entrar", command=entrar).grid(
        row=2, column=0, columnspan=2, pady=(10, 0)
    )
    janela.bind("<Return>", entrar)
    janela.update_idletasks()
    centralizar_janela(janela, janela.winfo_reqwidth(), janela.winfo_reqheight())
    campo_senha.focus_set()
    janela.mainloop()
    return resultado["usuario"]


# Fim do bloco de login e centralização das janelas.
# Reúne a janela principal, os menus e as operações com o PostgreSQL.
class ImportadorApp:
    def __init__(self, root):
        # Prepara a janela, o estado da conexão e os elementos da tela inicial.
        self.root = root
        self.root.title("Importador")
        centralizar_janela(self.root, 600, 340)
        self.root.minsize(460, 260)
        self.configuracao = None
        self.configuracao_sqlserver = None
        self.status = tk.StringVar(value="Configure os dados do PostgreSQL para começar.")
        self.root.report_callback_exception = self._registrar_erro_interface

        self._criar_menu()

        conteudo = ttk.Frame(root, padding=24)
        conteudo.pack(fill="both", expand=True)
        ttk.Label(conteudo, text="Importador", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(
            conteudo,
            text="Use o menu Obter dados para consultar tabelas e salvar em Excel.",
        ).pack(anchor="w", pady=(8, 20))
        ttk.Separator(conteudo).pack(fill="x", pady=(0, 14))
        ttk.Label(conteudo, textvariable=self.status, wraplength=520).pack(anchor="w")
        LOGGER.info("Aplicativo iniciado.")

    @staticmethod
    def _registrar_erro_interface(tipo_erro, _valor_erro, _traceback_erro):
        LOGGER.error("Erro não tratado na interface: %s", tipo_erro.__name__)

    def _criar_menu(self):
        # Conecta cada opção de menu ao fluxo correspondente do aplicativo.
        menu = tk.Menu(self.root)
        menu_geral = tk.Menu(menu, tearoff=False)
        menu_geral.add_command(label="Configurar conexão...", command=self.abrir_configuracao)
        menu_geral.add_command(label="Testar conexão", command=self.testar_conexao_salva)
        menu_geral.add_separator()
        menu_geral.add_command(
            label="Executar comandos...", command=self.abrir_execucao_comandos
        )
        menu.add_cascade(label="Menu", menu=menu_geral)

        obter_dados_menu = tk.Menu(menu, tearoff=False)
        postgres_menu = tk.Menu(obter_dados_menu, tearoff=False)
        postgres_menu.add_command(
            label="Consultar tabela e salvar em Excel...",
            command=self.abrir_obter_dados,
        )
        obter_dados_menu.add_cascade(label="PostgreSQL", menu=postgres_menu)

        sql_server_menu = tk.Menu(obter_dados_menu, tearoff=False)
        sql_server_menu.add_command(
            label="Consultar tabela e salvar em Excel...",
            command=self.abrir_obter_dados_sqlserver,
        )
        obter_dados_menu.add_cascade(label="SQL Server", menu=sql_server_menu)
        menu.add_cascade(label="Obter dados", menu=obter_dados_menu)

        importar_menu = tk.Menu(menu, tearoff=False)
        importar_menu.add_command(
            label="Produtos", command=self.importar_produtos_csv
        )
        importar_menu.add_command(
            label="Grupos", command=self.importar_grupos_csv
        )
        importar_menu.add_command(
            label="Clientes", command=self.importar_clientes_csv
        )
        menu.add_cascade(label="Importar", menu=importar_menu)

        ajuda_menu = tk.Menu(menu, tearoff=False)
        ajuda_menu.add_command(
            label="Sobre",
            command=lambda: messagebox.showinfo(
                "Sobre",
                "Importador de dados para PostgreSQL.\nVersão Beta 1.8",
                parent=self.root,
            ),
        )
        menu.add_cascade(label="Ajuda", menu=ajuda_menu)
        self.root.config(menu=menu)

    def abrir_obter_dados(self):
        # Monta a tela para consultar uma tabela e exportar o resultado para Excel.
        janela = tk.Toplevel(self.root)
        janela.title("Obter dados do PostgreSQL")
        janela.geometry("960x640")
        janela.minsize(700, 480)
        janela.transient(self.root)

        # Reaproveita a conexão da sessão ou preenche os campos com valores do .env.
        valores_atuais = self.configuracao or {
            "host": os.getenv("PGHOST", "localhost"),
            "port": os.getenv("PGPORT", "5432"),
            "dbname": os.getenv("PGDATABASE", "autosystem"),
            "user": os.getenv("PGUSER", "postgres"),
            "password": os.getenv("PGPASSWORD", ""),
        }
        campos = (
            ("Servidor", "host"),
            ("Porta", "port"),
            ("Banco", "dbname"),
            ("Usuário", "user"),
            ("Senha", "password"),
        )
        variaveis = {
            chave: tk.StringVar(value=str(valores_atuais[chave]))
            for _, chave in campos
        }
        nome_tabela = tk.StringVar()
        status = tk.StringVar(value="Informe a conexão e a tabela para consultar.")
        resultado = {"colunas": [], "linhas": []}

        # Cria os campos de conexão e a entrada do nome da tabela.
        formulario = ttk.Frame(janela, padding=12)
        formulario.pack(fill="x")
        for indice, (rotulo, chave) in enumerate(campos):
            coluna = indice * 2
            ttk.Label(formulario, text=rotulo).grid(
                row=0, column=coluna, sticky="w", padx=(0, 6)
            )
            ttk.Entry(
                formulario,
                textvariable=variaveis[chave],
                width=16,
                show="*" if chave == "password" else "",
            ).grid(row=1, column=coluna, sticky="ew", padx=(0, 10), pady=(4, 0))
            formulario.columnconfigure(coluna, weight=1)

        ttk.Label(formulario, text="Tabela (ou esquema.tabela)").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(12, 0)
        )
        ttk.Entry(formulario, textvariable=nome_tabela, width=36).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=(0, 10), pady=(4, 0)
        )

        # Liga os botões de consulta e exportação às funções de processamento.
        barra = ttk.Frame(janela, padding=(12, 0, 12, 8))
        barra.pack(fill="x")
        botao_excel = ttk.Button(
            barra,
            text="Salvar resultado em Excel...",
            state="disabled",
            command=lambda: self.salvar_dados_excel(resultado, janela),
        )
        botao_excel.pack(side="right")
        ttk.Button(
            barra,
            text="Obter dados",
            command=lambda: self.consultar_tabela(
                variaveis,
                nome_tabela,
                resultado,
                tabela_resultados,
                status,
                botao_excel,
                janela,
            ),
        ).pack(side="left")
        ttk.Label(barra, textvariable=status).pack(side="left", padx=12)

        # Exibe o resultado em uma grade com rolagem horizontal e vertical.
        area_resultados = ttk.Frame(janela, padding=(12, 0, 12, 12))
        area_resultados.pack(fill="both", expand=True)
        tabela_resultados = ttk.Treeview(area_resultados, show="headings")
        rolagem_vertical = ttk.Scrollbar(
            area_resultados, orient="vertical", command=tabela_resultados.yview
        )
        rolagem_horizontal = ttk.Scrollbar(
            area_resultados, orient="horizontal", command=tabela_resultados.xview
        )
        tabela_resultados.configure(
            yscrollcommand=rolagem_vertical.set,
            xscrollcommand=rolagem_horizontal.set,
        )
        tabela_resultados.grid(row=0, column=0, sticky="nsew")
        rolagem_vertical.grid(row=0, column=1, sticky="ns")
        rolagem_horizontal.grid(row=1, column=0, sticky="ew")
        area_resultados.rowconfigure(0, weight=1)
        area_resultados.columnconfigure(0, weight=1)

    def abrir_obter_dados_sqlserver(self):
        # Abre a tela de consulta usando SQL Server e os drivers ODBC instalados.
        drivers = pyodbc.drivers()
        driver_preferido = os.getenv("MSSQL_DRIVER", "").strip()
        if not driver_preferido:
            driver_preferido = next(
                (
                    driver
                    for versao in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server")
                    for driver in drivers
                    if driver == versao
                ),
                drivers[0] if drivers else "ODBC Driver 18 for SQL Server",
            )
        if driver_preferido not in drivers:
            drivers.insert(0, driver_preferido)

        janela = tk.Toplevel(self.root)
        janela.title("Obter dados do SQL Server")
        janela.geometry("960x640")
        janela.minsize(700, 480)
        janela.transient(self.root)

        valores_atuais = self.configuracao_sqlserver or {
            "driver": driver_preferido,
            "server": os.getenv("MSSQL_SERVER", "localhost"),
            "port": os.getenv("MSSQL_PORT", "1433"),
            "dbname": os.getenv("MSSQL_DATABASE", ""),
            "user": os.getenv("MSSQL_USER", ""),
            "password": os.getenv("MSSQL_PASSWORD", ""),
            "trust_server_certificate": os.getenv(
                "MSSQL_TRUST_SERVER_CERTIFICATE", "false"
            ).lower() in {"1", "true", "yes"},
        }
        campos = (
            ("Driver ODBC", "driver"),
            ("Servidor", "server"),
            ("Porta", "port"),
            ("Banco", "dbname"),
            ("Usuário", "user"),
            ("Senha", "password"),
        )
        variaveis = {
            chave: tk.StringVar(value=str(valores_atuais[chave]))
            for _, chave in campos
        }
        confiar_certificado = tk.BooleanVar(
            value=valores_atuais["trust_server_certificate"]
        )
        nome_tabela = tk.StringVar()
        status = tk.StringVar(value="Informe a conexão e a tabela para consultar.")
        resultado = {"colunas": [], "linhas": []}

        formulario = ttk.Frame(janela, padding=12)
        formulario.pack(fill="x")
        for indice, (rotulo, chave) in enumerate(campos):
            coluna = indice * 2
            ttk.Label(formulario, text=rotulo).grid(
                row=0, column=coluna, sticky="w", padx=(0, 6)
            )
            if chave == "driver":
                entrada = ttk.Combobox(
                    formulario,
                    textvariable=variaveis[chave],
                    values=drivers,
                    width=24,
                    state="readonly",
                )
            else:
                entrada = ttk.Entry(
                    formulario,
                    textvariable=variaveis[chave],
                    width=16,
                    show="*" if chave == "password" else "",
                )
            entrada.grid(row=1, column=coluna, sticky="ew", padx=(0, 10), pady=(4, 0))
            formulario.columnconfigure(coluna, weight=1)

        ttk.Checkbutton(
            formulario,
            text="Confiar no certificado do servidor (desenvolvimento)",
            variable=confiar_certificado,
        ).grid(row=2, column=0, columnspan=6, sticky="w", pady=(10, 0))
        ttk.Label(formulario, text="Tabela (ou esquema.tabela)").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(10, 0)
        )
        ttk.Entry(formulario, textvariable=nome_tabela, width=36).grid(
            row=4, column=0, columnspan=2, sticky="ew", padx=(0, 10), pady=(4, 0)
        )

        barra = ttk.Frame(janela, padding=(12, 0, 12, 8))
        barra.pack(fill="x")
        botao_excel = ttk.Button(
            barra,
            text="Salvar resultado em Excel...",
            state="disabled",
            command=lambda: self.salvar_dados_excel(resultado, janela),
        )
        botao_excel.pack(side="right")
        ttk.Button(
            barra,
            text="Obter dados",
            command=lambda: self.consultar_tabela_sqlserver(
                variaveis,
                confiar_certificado,
                nome_tabela,
                resultado,
                tabela_resultados,
                status,
                botao_excel,
                janela,
            ),
        ).pack(side="left")
        ttk.Label(barra, textvariable=status).pack(side="left", padx=12)

        area_resultados = ttk.Frame(janela, padding=(12, 0, 12, 12))
        area_resultados.pack(fill="both", expand=True)
        tabela_resultados = ttk.Treeview(area_resultados, show="headings")
        rolagem_vertical = ttk.Scrollbar(
            area_resultados, orient="vertical", command=tabela_resultados.yview
        )
        rolagem_horizontal = ttk.Scrollbar(
            area_resultados, orient="horizontal", command=tabela_resultados.xview
        )
        tabela_resultados.configure(
            yscrollcommand=rolagem_vertical.set,
            xscrollcommand=rolagem_horizontal.set,
        )
        tabela_resultados.grid(row=0, column=0, sticky="nsew")
        rolagem_vertical.grid(row=0, column=1, sticky="ns")
        rolagem_horizontal.grid(row=1, column=0, sticky="ew")
        area_resultados.rowconfigure(0, weight=1)
        area_resultados.columnconfigure(0, weight=1)

    @staticmethod
    def _valor_conexao_odbc(valor):
        # Escapa valores de conexão ODBC que contenham chaves ou ponto e vírgula.
        return "{" + str(valor).replace("}", "}}") + "}"

    @staticmethod
    def _identificador_sqlserver(partes):
        # Delimita cada parte do nome para evitar interpretar o nome como SQL executável.
        return ".".join("[" + parte.replace("]", "]]" ) + "]" for parte in partes)

    def consultar_tabela_sqlserver(
        self,
        variaveis,
        confiar_certificado,
        nome_tabela,
        resultado,
        tabela_resultados,
        status,
        botao_excel,
        parent,
    ):
        # Valida os dados ODBC e o destino antes de abrir a conexão SQL Server.
        configuracao = {chave: variavel.get().strip() for chave, variavel in variaveis.items()}
        try:
            configuracao["port"] = int(configuracao["port"])
            if not 1 <= configuracao["port"] <= 65535:
                raise ValueError
        except ValueError:
            messagebox.showerror("Porta inválida", "Informe uma porta entre 1 e 65535.", parent=parent)
            return

        obrigatorios = ("driver", "server", "dbname", "user", "password")
        if any(not configuracao[chave] for chave in obrigatorios):
            messagebox.showerror(
                "Dados incompletos",
                "Preencha driver, servidor, banco, usuário e senha.",
                parent=parent,
            )
            return

        partes_tabela = [parte.strip() for parte in nome_tabela.get().strip().split(".")]
        if not partes_tabela[0] or len(partes_tabela) > 2 or any(not parte for parte in partes_tabela):
            messagebox.showerror(
                "Tabela inválida",
                "Informe o nome da tabela ou esquema.tabela.",
                parent=parent,
            )
            return

        configuracao["trust_server_certificate"] = confiar_certificado.get()
        nome_tabela_log = ".".join(partes_tabela)
        LOGGER.info(
            "Consulta SQL Server iniciada: tabela=%s banco=%s servidor=%s.",
            nome_tabela_log,
            configuracao["dbname"],
            configuracao["server"],
        )
        tabela_resultados.delete(*tabela_resultados.get_children())
        resultado["colunas"] = []
        resultado["linhas"] = []
        botao_excel.state(["disabled"])

        servidor = f"{configuracao['server']},{configuracao['port']}"
        string_conexao = ";".join(
            (
                f"DRIVER={self._valor_conexao_odbc(configuracao['driver'])}",
                f"SERVER={self._valor_conexao_odbc(servidor)}",
                f"DATABASE={self._valor_conexao_odbc(configuracao['dbname'])}",
                f"UID={self._valor_conexao_odbc(configuracao['user'])}",
                f"PWD={self._valor_conexao_odbc(configuracao['password'])}",
                "Encrypt=yes",
                "TrustServerCertificate=" + ("yes" if confiar_certificado.get() else "no"),
                "Connection Timeout=5",
            )
        )
        tabela_sql = self._identificador_sqlserver(partes_tabela)
        consulta = f"SELECT * FROM {tabela_sql}"

        try:
            with closing(pyodbc.connect(string_conexao, timeout=5)) as conexao:
                with closing(conexao.cursor()) as cursor:
                    cursor.execute(consulta)
                    colunas = [coluna[0] for coluna in cursor.description]
                    linhas = []
                    while True:
                        lote = cursor.fetchmany(2000)
                        if not lote:
                            break
                        linhas.extend(lote)
        except pyodbc.Error as erro:
            status.set("Falha ao obter dados do SQL Server.")
            LOGGER.error(
                "Consulta SQL Server falhou: tabela=%s erro=%s.",
                nome_tabela_log,
                type(erro).__name__,
            )
            messagebox.showerror("Erro na consulta SQL Server", str(erro), parent=parent)
            return

        LOGGER.info(
            "Consulta SQL Server concluída: tabela=%s registros=%s.",
            nome_tabela_log,
            len(linhas),
        )
        self.configuracao_sqlserver = configuracao
        resultado["colunas"] = colunas
        resultado["linhas"] = linhas
        identificadores = [f"coluna_{indice}" for indice in range(len(colunas))]
        tabela_resultados.configure(columns=identificadores)
        for identificador, nome in zip(identificadores, colunas):
            tabela_resultados.heading(identificador, text=nome)
            tabela_resultados.column(identificador, width=150, minwidth=80, stretch=True)
        for linha in linhas[:1000]:
            tabela_resultados.insert(
                "",
                "end",
                values=[
                    "" if valor is None else str(self._decodificar_bytes(valor))
                    for valor in linha
                ],
            )

        botao_excel.state(["!disabled"])
        status.set(
            f"{len(linhas)} registro(s) obtido(s); a tela exibe até 1.000."
        )

    def consultar_tabela(
        self,
        variaveis,
        nome_tabela,
        resultado,
        tabela_resultados,
        status,
        botao_excel,
        parent,
    ):
        # Lê e valida os dados de conexão antes de abrir qualquer conexão.
        configuracao = {chave: variavel.get().strip() for chave, variavel in variaveis.items()}
        try:
            configuracao["port"] = int(configuracao["port"])
            if not 1 <= configuracao["port"] <= 65535:
                raise ValueError
        except ValueError:
            messagebox.showerror("Porta inválida", "Informe uma porta entre 1 e 65535.", parent=parent)
            return

        if any(not configuracao[chave] for chave in ("host", "dbname", "user", "password")):
            messagebox.showerror(
                "Dados incompletos",
                "Preencha servidor, banco, usuário e senha.",
                parent=parent,
            )
            return

        # Aceita tabela ou esquema.tabela; os identificadores serão escapados pelo psycopg.
        partes_tabela = [parte.strip() for parte in nome_tabela.get().strip().split(".")]
        if not partes_tabela[0] or len(partes_tabela) > 2 or any(not parte for parte in partes_tabela):
            messagebox.showerror(
                "Tabela inválida",
                "Informe o nome da tabela ou esquema.tabela.",
                parent=parent,
            )
            return

        nome_tabela_log = ".".join(partes_tabela)
        LOGGER.info(
            "Consulta iniciada: tabela=%s banco=%s servidor=%s.",
            nome_tabela_log,
            configuracao["dbname"],
            configuracao["host"],
        )
        tabela_resultados.delete(*tabela_resultados.get_children())
        resultado["colunas"] = []
        resultado["linhas"] = []
        botao_excel.state(["disabled"])

        # Executa uma consulta somente leitura e busca registros em lotes.
        consulta = sql.SQL("SELECT * FROM {}").format(sql.Identifier(*partes_tabela))
        try:
            with psycopg.connect(**configuracao, connect_timeout=5) as conexao:
                with conexao.cursor() as cursor:
                    cursor.execute("SET TRANSACTION READ ONLY")
                    cursor.execute(consulta)
                    colunas = [coluna.name for coluna in cursor.description]
                    linhas = []
                    while True:
                        lote = cursor.fetchmany(2000)
                        if not lote:
                            break
                        linhas.extend(lote)
        except psycopg.Error as erro:
            status.set("Falha ao obter dados.")
            LOGGER.error(
                "Consulta falhou para tabela=%s (%s).",
                nome_tabela_log,
                type(erro).__name__,
            )
            messagebox.showerror("Erro na consulta", str(erro), parent=parent)
            return

        # Guarda os dados retornados, exibe até 1.000 linhas e habilita a exportação.
        LOGGER.info("Consulta concluída: tabela=%s registros=%s.", nome_tabela_log, len(linhas))
        self.configuracao = configuracao
        resultado["colunas"] = colunas
        resultado["linhas"] = linhas
        identificadores = [f"coluna_{indice}" for indice in range(len(colunas))]
        tabela_resultados.configure(columns=identificadores)
        for identificador, nome in zip(identificadores, colunas):
            tabela_resultados.heading(identificador, text=nome)
            tabela_resultados.column(identificador, width=150, minwidth=80, stretch=True)
        for linha in linhas[:1000]:
            tabela_resultados.insert(
                "",
                "end",
                values=[
                    "" if valor is None else str(self._decodificar_bytes(valor))
                    for valor in linha
                ],
            )

        botao_excel.state(["!disabled"])
        status.set(
            f"{len(linhas)} registro(s) obtido(s); a tela exibe até 1.000."
        )

    def salvar_dados_excel(self, resultado, parent):
        # Verifica se há dados e se o resultado cabe em uma planilha do Excel.
        if not resultado["colunas"]:
            LOGGER.warning("Exportação Excel solicitada sem dados de consulta.")
            messagebox.showwarning("Sem dados", "Faça uma consulta antes de salvar.", parent=parent)
            return
        if len(resultado["linhas"]) > 1_048_575:
            messagebox.showerror(
                "Limite do Excel",
                "O resultado ultrapassa o limite de 1.048.575 registros por aba.",
                parent=parent,
            )
            return

        caminho = filedialog.asksaveasfilename(
            parent=parent,
            title="Salvar resultado da consulta",
            initialfile="dados_consulta.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")],
        )
        if not caminho:
            LOGGER.info("Salvamento Excel cancelado pelo usuário.")
            return

        # Grava cabeçalho e todas as linhas em modo streaming para reduzir memória.
        try:
            workbook = Workbook(write_only=True)
            planilha = workbook.create_sheet("Dados")
            planilha.append(resultado["colunas"])
            tipos_excel = (str, int, float, bool, Decimal, date, datetime, time, timedelta)
            for linha in resultado["linhas"]:
                planilha.append(
                    [
                        self._decodificar_bytes(valor)
                        if valor is None or isinstance(valor, tipos_excel)
                        else str(self._decodificar_bytes(valor))
                        for valor in linha
                    ]
                )
            workbook.save(caminho)
        except Exception as erro:
            LOGGER.error("Falha ao salvar planilha Excel (%s).", type(erro).__name__)
            messagebox.showerror("Erro ao salvar Excel", str(erro), parent=parent)
            return

        LOGGER.info(
            "Planilha Excel salva: arquivo=%s registros=%s.",
            Path(caminho).name,
            len(resultado["linhas"]),
        )
        messagebox.showinfo(
            "Arquivo salvo",
            f"{len(resultado['linhas'])} registro(s) salvo(s) em:\n{caminho}",
            parent=parent,
        )

    @staticmethod
    def _decodificar_bytes(valor):
        # Converte conteúdo binário UTF-8 para texto legível nas grades e planilhas.
        if isinstance(valor, (bytes, bytearray, memoryview)):
            return bytes(valor).decode("utf-8", errors="replace")
        return valor

    @staticmethod
    def _normalizar_valor_unico(valor):
        # Normaliza valores e ignora espaços/capitalização em textos.
        if isinstance(valor, str):
            return "texto", valor.strip().casefold()
        if isinstance(valor, bool):
            return "booleano", valor
        if isinstance(valor, (int, float, Decimal)):
            return "numero", Decimal(str(valor)).normalize()
        return type(valor).__name__, str(valor).strip().casefold()

    # Importação XLSX desativada em favor da importação por CSV.
    """
    def _analisar_colunas_unicas_xlsx(self, caminho, colunas):
        # Exige as duas colunas no cabeçalho e valida a unicidade dos valores preenchidos.
        colunas_obrigatorias = ("codigo", "codigo_Barra")
        indices_colunas = {nome.casefold(): [] for nome in colunas_obrigatorias}
        problemas = []

        for indice, nome_coluna in enumerate(colunas):
            chave = nome_coluna.casefold()
            if chave in indices_colunas:
                indices_colunas[chave].append(indice)

        for nome_coluna in colunas_obrigatorias:
            indices = indices_colunas[nome_coluna.casefold()]
            if not indices:
                problemas.append(f'Coluna obrigatória "{nome_coluna}" ausente no cabeçalho.')
            elif len(indices) > 1:
                problemas.append(f'Coluna obrigatória "{nome_coluna}" aparece mais de uma vez.')

        primeira_ocorrencia = {nome: {} for nome in colunas_obrigatorias}
        linhas_analisadas = 0
        workbook = load_workbook(caminho, read_only=True, data_only=True)
        try:
            planilha = workbook.active
            for numero_linha, linha in enumerate(
                planilha.iter_rows(min_row=2, values_only=True), start=2
            ):
                valores = list(linha)
                if not any(valor is not None for valor in valores):
                    continue
                linhas_analisadas += 1

                for nome_coluna in colunas_obrigatorias:
                    indices = indices_colunas[nome_coluna.casefold()]
                    if len(indices) != 1:
                        continue

                    indice = indices[0]
                    valor = valores[indice] if indice < len(valores) else None
                    if valor is None or (isinstance(valor, str) and not valor.strip()):
                        continue

                    chave_valor = self._normalizar_valor_unico(valor)
                    ocorrencias = primeira_ocorrencia[nome_coluna]
                    if chave_valor in ocorrencias:
                        primeira_linha, valor_original = ocorrencias[chave_valor]
                        problemas.append(
                            f'Linha {numero_linha}, coluna "{nome_coluna}": valor '
                            f'"{valor_original}" duplicado (primeira ocorrência na linha '
                            f'{primeira_linha}).'
                        )
                    else:
                        ocorrencias[chave_valor] = (numero_linha, valor)
        finally:
            workbook.close()

        return planilha.title, linhas_analisadas, problemas

    def importar_dados_xlsx(self):
        # Solicita destino e arquivo antes de começar a preparar os registros.
        if not self._exigir_configuracao():
            return

        nome_tabela = simpledialog.askstring(
            "Tabela de destino",
            "Informe a tabela onde os dados serão inseridos (ou esquema.tabela):",
            initialvalue="produto",
            parent=self.root,
        )
        if nome_tabela is None:
            LOGGER.info("Importação XLSX cancelada antes da escolha da tabela.")
            return

        # Separa o nome da tabela para montar um identificador SQL seguro.
        partes_tabela = [parte.strip() for parte in nome_tabela.split(".")]
        if not partes_tabela[0] or len(partes_tabela) > 2 or any(not parte for parte in partes_tabela):
            messagebox.showerror(
                "Tabela inválida",
                "Informe o nome da tabela ou esquema.tabela.",
                parent=self.root,
            )
            return
        nome_tabela = ".".join(partes_tabela)
        LOGGER.info("Importação XLSX iniciada: tabela=%s.", nome_tabela)

        caminho = filedialog.askopenfilename(
            parent=self.root,
            title="Selecionar planilha XLSX para importar",
            filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
        )
        if not caminho:
            LOGGER.info("Seleção de planilha XLSX cancelada.")
            return
        LOGGER.info("Planilha selecionada: arquivo=%s.", Path(caminho).name)

        # Abre a planilha ativa e usa a primeira linha como nomes das colunas.
        try:
            workbook = load_workbook(caminho, read_only=True, data_only=True)
            try:
                planilha = workbook.active
                cabecalho = next(planilha.iter_rows(values_only=True), None)
            finally:
                workbook.close()
        except (OSError, BadZipFile, InvalidFileException, ValueError) as erro:
            LOGGER.error("Falha ao ler planilha XLSX (%s).", type(erro).__name__)
            messagebox.showerror("Erro ao ler XLSX", str(erro), parent=self.root)
            return

        if not cabecalho:
            LOGGER.warning("Importação XLSX cancelada: cabeçalho vazio.")
            messagebox.showerror(
                "Planilha vazia",
                "A aba ativa não contém uma linha de cabeçalho.",
                parent=self.root,
            )
            return

        # Valida o cabeçalho antes de pedir confirmação para inserir os dados.
        colunas = [str(coluna).strip() if coluna is not None else "" for coluna in cabecalho]
        if any(not coluna for coluna in colunas):
            LOGGER.warning("Importação XLSX cancelada: há coluna sem nome.")
            messagebox.showerror(
                "Cabeçalho inválido",
                "Todas as células do cabeçalho precisam ter um nome de coluna.",
                parent=self.root,
            )
            return
        if len(set(colunas)) != len(colunas):
            LOGGER.warning("Importação XLSX cancelada: há nomes de coluna duplicados.")
            messagebox.showerror(
                "Cabeçalho inválido",
                "O cabeçalho contém nomes de coluna duplicados.",
                parent=self.root,
            )
            return

        # Interrompe a importação se faltar uma coluna ou houver valor preenchido duplicado.
        try:
            nome_aba, linhas_analisadas, problemas = self._analisar_colunas_unicas_xlsx(
                caminho, colunas
            )
        except (OSError, BadZipFile, InvalidFileException, ValueError) as erro:
            LOGGER.error("Falha na análise prévia da planilha (%s).", type(erro).__name__)
            messagebox.showerror("Erro na análise da planilha", str(erro), parent=self.root)
            return

        if problemas:
            conteudo_relatorio = [
                "RELATORIO DE VALIDACAO DA IMPORTACAO XLSX",
                f"Data e hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"Arquivo: {Path(caminho).name}",
                f"Aba analisada: {nome_aba}",
                f"Tabela de destino: {nome_tabela}",
                f"Linhas com dados analisadas: {linhas_analisadas}",
                f"Problemas encontrados: {len(problemas)}",
                "",
                "DETALHES",
                *[f"- {problema}" for problema in problemas],
            ]
            try:
                ARQUIVO_RELATORIO_VALIDACAO.write_text(
                    "\n".join(conteudo_relatorio) + "\n",
                    encoding="utf-8",
                )
            except OSError as erro:
                LOGGER.error("Falha ao salvar relatório de validação (%s).", type(erro).__name__)
                messagebox.showerror(
                    "Erro ao salvar relatório",
                    f"Não foi possível salvar o relatório:\n{erro}",
                    parent=self.root,
                )
                return

            LOGGER.warning(
                "Importação bloqueada pela validação: tabela=%s problemas=%s relatório=%s.",
                nome_tabela,
                len(problemas),
                ARQUIVO_RELATORIO_VALIDACAO.name,
            )
            messagebox.showerror(
                "Planilha reprovada na validação",
                f"Foram encontrados {len(problemas)} problema(s) nas colunas "
                "codigo e codigo_Barra. A importação foi cancelada.\n\n"
                f"Relatório salvo em:\n{ARQUIVO_RELATORIO_VALIDACAO}",
                parent=self.root,
            )
            return

        LOGGER.info(
            "Validação de colunas e duplicidade concluída: tabela=%s linhas=%s.",
            nome_tabela,
            linhas_analisadas,
        )

        nomes_colunas = ", ".join(colunas)
        confirmar = messagebox.askyesno(
            "Confirmar importação",
            f"Colunas lidas da aba ativa:\n{nomes_colunas}\n\n"
            f"Inserir os registros na tabela '{nome_tabela}'?",
            parent=self.root,
        )
        if not confirmar:
            LOGGER.info("Importação XLSX cancelada na confirmação: tabela=%s.", nome_tabela)
            return

        # Monta o INSERT com identificadores escapados e valores parametrizados.
        consulta_insert = sql.SQL("INSERT INTO {} ({}) VALUES ({})").format(
            sql.Identifier(*partes_tabela),
            sql.SQL(", ").join(sql.Identifier(coluna) for coluna in colunas),
            sql.SQL(", ").join(sql.Placeholder() for _ in colunas),
        )
        linhas_importadas = 0

        def ler_linhas():
            # Lê as linhas da planilha sob demanda e valida a quantidade de colunas.
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

        # Executa todos os INSERTs na mesma transação para permitir rollback em erro.
        try:
            with psycopg.connect(**self.configuracao, connect_timeout=5) as conexao:
                with conexao.cursor() as cursor:
                    cursor.executemany(consulta_insert, ler_linhas())
        except (psycopg.Error, OSError, BadZipFile, InvalidFileException, ValueError) as erro:
            self.status.set("Falha ao importar o XLSX.")
            LOGGER.error(
                "Importação XLSX falhou: tabela=%s erro=%s.",
                nome_tabela,
                type(erro).__name__,
            )
            messagebox.showerror("Erro na importação", str(erro), parent=self.root)
            return

        LOGGER.info(
            "Importação XLSX concluída: tabela=%s registros=%s.",
            nome_tabela,
            linhas_importadas,
        )
        self.status.set(
            f"Importação concluída: {linhas_importadas} registro(s) inserido(s) em {nome_tabela}."
        )
        messagebox.showinfo(
            "Importação concluída",
            f"{linhas_importadas} registro(s) inserido(s) na tabela '{nome_tabela}'.",
            parent=self.root,
        )

    """

    # Início do bloco de importação e validação de arquivos CSV.
    def _analisar_colunas_unicas_csv(self, caminho, colunas, validar_codigos=True):
        if not validar_codigos:
            with open(caminho, "r", encoding="utf-8-sig", newline="") as arquivo:
                leitor = csv.reader(arquivo, delimiter=";")
                next(leitor, None)
                linhas_analisadas = sum(
                    any(valor.strip() for valor in valores)
                    for valores in leitor
                )
            return linhas_analisadas, []

        # Valida a presença e a unicidade dos valores em codigo e codigo_Barra.
        colunas_obrigatorias = ("codigo", "codigo_Barra")
        indices_colunas = {nome.casefold(): [] for nome in colunas_obrigatorias}
        problemas = []

        for indice, nome_coluna in enumerate(colunas):
            chave = nome_coluna.casefold()
            if chave in indices_colunas:
                indices_colunas[chave].append(indice)

        for nome_coluna in colunas_obrigatorias:
            indices = indices_colunas[nome_coluna.casefold()]
            if not indices:
                problemas.append(f'Coluna obrigatória "{nome_coluna}" ausente no cabeçalho.')
            elif len(indices) > 1:
                problemas.append(f'Coluna obrigatória "{nome_coluna}" aparece mais de uma vez.')

        primeira_ocorrencia = {nome: {} for nome in colunas_obrigatorias}
        linhas_analisadas = 0
        with open(caminho, "r", encoding="utf-8-sig", newline="") as arquivo:
            leitor = csv.reader(arquivo, delimiter=";")
            next(leitor, None)
            for numero_linha, valores in enumerate(leitor, start=2):
                if not any(valor.strip() for valor in valores):
                    continue
                if len(valores) > len(colunas) and any(
                    valor.strip() for valor in valores[len(colunas):]
                ):
                    raise ValueError(
                        f"A linha {numero_linha} contém valores além das colunas do cabeçalho."
                    )
                valores = valores[:len(colunas)]
                valores.extend([""] * (len(colunas) - len(valores)))
                linhas_analisadas += 1

                for nome_coluna in colunas_obrigatorias:
                    indices = indices_colunas[nome_coluna.casefold()]
                    if len(indices) != 1:
                        continue

                    valor = valores[indices[0]]
                    if not valor.strip():
                        continue

                    chave_valor = self._normalizar_valor_unico(valor)
                    ocorrencias = primeira_ocorrencia[nome_coluna]
                    if chave_valor in ocorrencias:
                        primeira_linha, valor_original = ocorrencias[chave_valor]
                        problemas.append(
                            f'Linha {numero_linha}, coluna "{nome_coluna}": valor '
                            f'"{valor_original}" duplicado (primeira ocorrência na linha '
                            f'{primeira_linha}).'
                        )
                    else:
                        ocorrencias[chave_valor] = (numero_linha, valor)

        return linhas_analisadas, problemas

    # Início do bloco de recuperação de texto UTF-8 previamente decodificado incorretamente.
    @staticmethod
    def _corrigir_mojibake(texto):
        def pontuar(valor):
            return sum(
                caractere in "ÃÂ�" or "\x80" <= caractere <= "\x9f"
                for caractere in valor
            )

        pontuacao_original = pontuar(texto)
        if not pontuacao_original:
            return texto

        for codificacao in ("cp1252", "latin-1"):
            try:
                candidato = texto.encode(codificacao).decode("utf-8")
            except (UnicodeEncodeError, UnicodeDecodeError):
                continue
            if pontuar(candidato) < pontuacao_original:
                return candidato
        return texto

    # Fim do bloco de recuperação de texto UTF-8 previamente decodificado incorretamente.
    # Início do bloco de revisão de acentos e caracteres especiais.
    # Revisa caracteres acentuados e permite normalizar uma cópia do CSV.
    @staticmethod
    def _caracteres_acentuados(texto):
        texto = unicodedata.normalize("NFC", texto)
        return sorted({
            caractere
            for caractere in texto
            if caractere.isalpha()
            and any(
                unicodedata.combining(parte)
                for parte in unicodedata.normalize("NFD", caractere)
            )
        })

    # Início do bloco de detecção e normalização de caracteres especiais.
    @staticmethod
    def _caracteres_especiais(texto):
        texto = unicodedata.normalize("NFC", texto)
        return sorted({
            caractere
            for caractere in texto
            if caractere not in "-.+,@"
            and not caractere.isalnum()
            and not caractere.isspace()
        })

    @staticmethod
    def _normalizar_texto_csv(texto):
        texto = ImportadorApp._corrigir_mojibake(texto)
        texto = texto.replace("ç", "c").replace("Ç", "C")
        resultado = []
        for caractere in unicodedata.normalize("NFD", texto):
            if unicodedata.combining(caractere):
                continue
            if (
                caractere.isalnum()
                or caractere.isspace()
                or caractere in "-.+,@"
            ):
                resultado.append(caractere.upper())
            else:
                resultado.append(" ")
        return "".join(resultado)

    # Fim do bloco de detecção e normalização de caracteres especiais.
    def _criar_csv_normalizado(self, caminho):
        origem = Path(caminho)
        destino = origem.with_name(f"{origem.stem}_sem_acentos{origem.suffix}")
        indice = 2
        while destino.exists():
            destino = origem.with_name(
                f"{origem.stem}_sem_acentos_{indice}{origem.suffix}"
            )
            indice += 1

        try:
            with open(origem, "r", encoding="utf-8-sig", newline="") as arquivo_origem:
                leitor = csv.reader(arquivo_origem, delimiter=";")
                with open(destino, "x", encoding="utf-8-sig", newline="") as arquivo_destino:
                    escritor = csv.writer(arquivo_destino, delimiter=";")
                    for numero_linha, valores in enumerate(leitor, start=1):
                        if numero_linha > 1:
                            valores = [self._normalizar_texto_csv(valor) for valor in valores]
                        escritor.writerow(valores)
        except (OSError, UnicodeError, csv.Error):
            try:
                destino.unlink(missing_ok=True)
            except OSError:
                pass
            raise

        LOGGER.info(
            "Cópia CSV normalizada criada: arquivo=%s.",
            destino.name,
        )
        return str(destino)

    # Início do bloco de varredura das linhas de dados, ignorando o cabeçalho.
    def _localizar_acentos_csv(self, caminho):
        ocorrencias = []
        with open(caminho, "r", encoding="utf-8-sig", newline="") as arquivo:
            leitor = csv.reader(arquivo, delimiter=";")
            for numero_registro, valores in enumerate(leitor, start=1):
                if numero_registro == 1:
                    continue
                valores = [self._corrigir_mojibake(valor) for valor in valores]
                numero_linha = leitor.line_num
                acentos = sorted({
                    caractere
                    for valor in valores
                    for caractere in self._caracteres_acentuados(valor)
                })
                especiais = sorted({
                    caractere
                    for valor in valores
                    for caractere in self._caracteres_especiais(valor)
                })
                if acentos or especiais:
                    trecho = " ; ".join(" ".join(valor.split()) for valor in valores)
                    ocorrencias.append((numero_linha, acentos, especiais, trecho))
        return ocorrencias

    # Fim do bloco de varredura das linhas de dados, ignorando o cabeçalho.
    def _revisar_acentos_csv(self, caminho):
        try:
            ocorrencias = self._localizar_acentos_csv(caminho)
        except (OSError, UnicodeError, csv.Error) as erro:
            LOGGER.error("Falha ao verificar acentos no CSV (%s).", type(erro).__name__)
            messagebox.showerror("Erro ao verificar CSV", str(erro), parent=self.root)
            return None

        janela = tk.Toplevel(self.root)
        janela.title("Ajuste de informações")
        janela.geometry("900x500")
        janela.minsize(680, 360)
        janela.transient(self.root)
        resultado = {"continuar": False, "caminho": caminho}

        if ocorrencias:
            texto = (
                f"Foram encontrados acentos ou caracteres especiais em "
                f"{len(ocorrencias)} linha(s). Confira antes de continuar. "
                "O cabeçalho será preservado."
            )
        else:
            texto = (
                "Nenhum caractere acentuado ou especial foi encontrado. "
                "O cabeçalho será preservado."
            )
        rotulo_status = ttk.Label(janela, text=texto, padding=(12, 12, 12, 8))
        rotulo_status.pack(anchor="w")

        area = ttk.Frame(janela, padding=(12, 0, 12, 8))
        area.pack(fill="both", expand=True)
        colunas = ("linha", "acentos", "especiais", "conteudo")
        tabela = ttk.Treeview(area, columns=colunas, show="headings")
        tabela.heading("linha", text="Linha")
        tabela.heading("acentos", text="Acentos")
        tabela.heading("especiais", text="Caracteres especiais")
        tabela.heading("conteudo", text="Conteúdo da linha")
        tabela.column("linha", width=70, minwidth=60, stretch=False, anchor="center")
        tabela.column("acentos", width=130, minwidth=100, stretch=False)
        tabela.column("especiais", width=150, minwidth=120, stretch=False)
        tabela.column("conteudo", width=460, minwidth=220, stretch=True)
        rolagem_vertical = ttk.Scrollbar(area, orient="vertical", command=tabela.yview)
        rolagem_horizontal = ttk.Scrollbar(area, orient="horizontal", command=tabela.xview)
        tabela.configure(
            yscrollcommand=rolagem_vertical.set,
            xscrollcommand=rolagem_horizontal.set,
        )
        tabela.grid(row=0, column=0, sticky="nsew")
        rolagem_vertical.grid(row=0, column=1, sticky="ns")
        rolagem_horizontal.grid(row=1, column=0, sticky="ew")
        area.rowconfigure(0, weight=1)
        area.columnconfigure(0, weight=1)
        def atualizar_lista():
            tabela.delete(*tabela.get_children())
            for numero_linha, acentos, especiais, trecho in ocorrencias:
                tabela.insert(
                    "",
                    "end",
                    values=(
                        numero_linha,
                        " ".join(acentos),
                        " ".join(especiais),
                        trecho,
                    ),
                )

        atualizar_lista()

        botoes = ttk.Frame(janela, padding=(12, 0, 12, 12))
        botoes.pack(fill="x")

        def fechar(permitir_importacao=False):
            resultado["continuar"] = permitir_importacao
            janela.destroy()

        botao_corrigir = ttk.Button(
            botoes,
            text="Ajuste automático",
        )
        botao_corrigir.pack(side="left")

        def corrigir_arquivo():
            try:
                novo_caminho = self._criar_csv_normalizado(resultado["caminho"])
                novas_ocorrencias = self._localizar_acentos_csv(novo_caminho)
            except (OSError, UnicodeError, csv.Error) as erro:
                LOGGER.error("Falha ao normalizar CSV (%s).", type(erro).__name__)
                messagebox.showerror(
                    "Erro ao corrigir CSV", str(erro), parent=janela
                )
                return

            resultado["caminho"] = novo_caminho
            ocorrencias[:] = novas_ocorrencias
            atualizar_lista()
            rotulo_status.configure(
                text=(
                    f"Cópia corrigida: {Path(novo_caminho).name}. "
                    "O cabeçalho foi preservado para manter os nomes das colunas."
                )
            )
            botao_corrigir.configure(state="disabled")

        botao_corrigir.configure(command=corrigir_arquivo)
        ttk.Button(
            botoes,
            text="Cancelar",
            command=lambda: fechar(),
        ).pack(side="right")
        ttk.Button(
            botoes,
            text="Importar dados",
            command=lambda: fechar(True),
        ).pack(side="right", padx=(0, 8))
        janela.protocol("WM_DELETE_WINDOW", lambda: fechar())
        janela.update_idletasks()
        centralizar_janela(janela, 900, 500)
        janela.grab_set()
        janela.wait_window()
        if resultado["continuar"]:
            return resultado["caminho"]
        return None

    # Fim do bloco de revisão e normalização de acentos no CSV.
    # Início do bloco que localiza CPFs duplicados no CSV de clientes.
    @staticmethod
    def _analisar_cpfs_duplicados_csv(caminho, colunas):
        indice_cpf = next(
            (
                indice
                for indice, coluna in enumerate(colunas)
                if coluna.casefold() == "cpf"
            ),
            None,
        )
        if indice_cpf is None:
            raise ValueError('O arquivo CSV de clientes precisa conter a coluna "cpf".')

        ocorrencias = {}
        duplicados = {}
        with open(caminho, "r", encoding="utf-8-sig", newline="") as arquivo:
            leitor = csv.reader(arquivo, delimiter=";")
            next(leitor, None)
            for numero_linha, valores in enumerate(leitor, start=2):
                if not any(valor.strip() for valor in valores):
                    continue
                if len(valores) > len(colunas) and any(
                    valor.strip() for valor in valores[len(colunas):]
                ):
                    raise ValueError(
                        f"A linha {numero_linha} contém valores além das colunas do cabeçalho."
                    )
                valores = valores[:len(colunas)]
                valores.extend([""] * (len(colunas) - len(valores)))
                cpf = valores[indice_cpf].strip()
                chave_cpf = re.sub(r"[.\-\s]", "", cpf).casefold()
                if not chave_cpf:
                    continue

                registro = {
                    "linha": numero_linha,
                    "cpf": cpf,
                    "valores": valores,
                }
                if chave_cpf in duplicados:
                    duplicados[chave_cpf].append(registro)
                elif chave_cpf in ocorrencias:
                    duplicados[chave_cpf] = [ocorrencias.pop(chave_cpf), registro]
                else:
                    ocorrencias[chave_cpf] = registro

        return list(duplicados.values())
    # Fim do bloco que localiza CPFs duplicados no CSV de clientes.

    # Início do bloco que gera o relatório Excel de CPFs duplicados.
    @staticmethod
    def _salvar_relatorio_cpfs_duplicados(duplicados, colunas, caminho):
        relatorio = Workbook()
        planilha = relatorio.active
        planilha.title = "CPFs duplicados"
        planilha.append(["Linha CSV", *colunas, "Ação"])
        linhas_ignoradas = set()

        for grupo in duplicados:
            for indice, item in enumerate(grupo):
                acao = "Mantido" if indice == 0 else "Ignorado como duplicado"
                planilha.append([item["linha"], *item["valores"], acao])
                if indice > 0:
                    linhas_ignoradas.add(item["linha"])

        relatorio.save(caminho)
        relatorio.close()
        return linhas_ignoradas
    # Fim do bloco que gera o relatório Excel de CPFs duplicados.

    def importar_dados_csv(
        self,
        validar_codigos=True,
        revisar_caracteres=True,
        tabela_padrao="produto",
        tabela_fixa=None,
        solicitar_configuracao=False,
        gerar_codigo_automatico=False,
        validar_cpfs_duplicados=False,
    ):
        # Solicita destino e arquivo antes de preparar os registros.
        if solicitar_configuracao:
            if not self._solicitar_configuracao_importacao():
                return
        elif not self._exigir_configuracao():
            return

        if tabela_fixa is None:
            nome_tabela = simpledialog.askstring(
                "Tabela de destino",
                "Informe a tabela onde os dados serão inseridos (ou esquema.tabela):",
                initialvalue=tabela_padrao,
                parent=self.root,
            )
            if nome_tabela is None:
                LOGGER.info("Importação CSV cancelada antes da escolha da tabela.")
                return
        else:
            nome_tabela = tabela_fixa

        partes_tabela = [parte.strip() for parte in nome_tabela.split(".")]
        if not partes_tabela[0] or len(partes_tabela) > 2 or any(not parte for parte in partes_tabela):
            messagebox.showerror(
                "Tabela inválida",
                "Informe o nome da tabela ou esquema.tabela.",
                parent=self.root,
            )
            return
        nome_tabela = ".".join(partes_tabela)
        LOGGER.info("Importação CSV iniciada: tabela=%s.", nome_tabela)

        caminho = filedialog.askopenfilename(
            parent=self.root,
            title="Selecionar arquivo CSV para importar",
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        )
        if not caminho:
            LOGGER.info("Seleção de arquivo CSV cancelada.")
            return
        LOGGER.info("Arquivo CSV selecionado: arquivo=%s.", Path(caminho).name)

        try:
            with open(caminho, "r", encoding="utf-8-sig", newline="") as arquivo:
                leitor = csv.reader(arquivo, delimiter=";")
                cabecalho = next(leitor, None)
        except (OSError, UnicodeError, csv.Error) as erro:
            LOGGER.error("Falha ao ler CSV (%s).", type(erro).__name__)
            messagebox.showerror("Erro ao ler CSV", str(erro), parent=self.root)
            return

        if not cabecalho or not any(coluna.strip() for coluna in cabecalho):
            LOGGER.warning("Importação CSV cancelada: cabeçalho vazio.")
            messagebox.showerror(
                "Arquivo vazio",
                "O arquivo CSV não contém nomes de colunas na primeira linha.",
                parent=self.root,
            )
            return

        colunas = [coluna.strip() for coluna in cabecalho]
        if any(not coluna for coluna in colunas):
            LOGGER.warning("Importação CSV cancelada: há coluna sem nome.")
            messagebox.showerror(
                "Cabeçalho inválido",
                "Todas as colunas da primeira linha precisam ter um nome.",
                parent=self.root,
            )
            return
        if len(set(colunas)) != len(colunas):
            LOGGER.warning("Importação CSV cancelada: há nomes de coluna duplicados.")
            messagebox.showerror(
                "Cabeçalho inválido",
                "A primeira linha contém nomes de coluna duplicados.",
                parent=self.root,
            )
            return

        # Início do bloco de validação de CPF duplicado e geração do relatório Excel (clientes).
        linhas_csv_ignoradas = set()
        if validar_cpfs_duplicados:
            try:
                grupos_cpfs_duplicados = self._analisar_cpfs_duplicados_csv(
                    caminho, colunas
                )
            except (OSError, UnicodeError, csv.Error, ValueError) as erro:
                LOGGER.error(
                    "Falha na validação de CPFs do CSV (%s).",
                    type(erro).__name__,
                )
                messagebox.showerror(
                    "Erro na validação de CPFs",
                    str(erro),
                    parent=self.root,
                )
                return

            if grupos_cpfs_duplicados:
                try:
                    linhas_csv_ignoradas = self._salvar_relatorio_cpfs_duplicados(
                        grupos_cpfs_duplicados,
                        colunas,
                        ARQUIVO_RELATORIO_CPFS_DUPLICADOS,
                    )
                except OSError as erro:
                    LOGGER.error(
                        "Falha ao salvar relatório Excel de CPFs duplicados (%s).",
                        type(erro).__name__,
                    )
                    messagebox.showerror(
                        "Erro ao salvar relatório",
                        f"Não foi possível salvar o relatório:\n{erro}",
                        parent=self.root,
                    )
                    return

                LOGGER.warning(
                    "CPFs duplicados encontrados: grupos=%s registros_ignorados=%s "
                    "relatorio=%s.",
                    len(grupos_cpfs_duplicados),
                    len(linhas_csv_ignoradas),
                    ARQUIVO_RELATORIO_CPFS_DUPLICADOS.name,
                )
        # Fim do bloco de validação de CPF duplicado e geração do relatório Excel (clientes).

        try:
            linhas_analisadas, problemas = self._analisar_colunas_unicas_csv(
                caminho, colunas, validar_codigos
            )
        except (OSError, UnicodeError, csv.Error, ValueError) as erro:
            LOGGER.error("Falha na análise prévia do CSV (%s).", type(erro).__name__)
            messagebox.showerror("Erro na análise do CSV", str(erro), parent=self.root)
            return

        if problemas:
            conteudo_relatorio = [
                "RELATORIO DE VALIDACAO DA IMPORTACAO CSV",
                f"Data e hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"Arquivo: {Path(caminho).name}",
                f"Tabela de destino: {nome_tabela}",
                f"Linhas com dados analisadas: {linhas_analisadas}",
                f"Problemas encontrados: {len(problemas)}",
                "",
                "DETALHES",
                *[f"- {problema}" for problema in problemas],
            ]
            try:
                ARQUIVO_RELATORIO_VALIDACAO.write_text(
                    "\n".join(conteudo_relatorio) + "\n",
                    encoding="utf-8",
                )
            except OSError as erro:
                LOGGER.error("Falha ao salvar relatório de validação (%s).", type(erro).__name__)
                messagebox.showerror(
                    "Erro ao salvar relatório",
                    f"Não foi possível salvar o relatório:\n{erro}",
                    parent=self.root,
                )
                return

            LOGGER.warning(
                "Importação bloqueada pela validação: tabela=%s problemas=%s relatório=%s.",
                nome_tabela,
                len(problemas),
                ARQUIVO_RELATORIO_VALIDACAO.name,
            )
            messagebox.showerror(
                "CSV reprovado na validação",
                f"Foram encontrados {len(problemas)} problema(s) nas colunas "
                "codigo e codigo_Barra. A importação foi cancelada.\n\n"
                f"Relatório salvo em:\n{ARQUIVO_RELATORIO_VALIDACAO}",
                parent=self.root,
            )
            return

        if revisar_caracteres:
            caminho_revisado = self._revisar_acentos_csv(caminho)
            if caminho_revisado is None:
                LOGGER.info("Importação CSV cancelada na revisão de caracteres.")
                return
            caminho = caminho_revisado

        # Início do bloco que prepara a coluna info de produto.
        incluir_info_padrao = partes_tabela[-1].casefold() == "produto"
        colunas_insert = list(colunas)
        # Início do bloco que define se o código de pessoa será gerado (CSV sem coluna codigo).
        gerar_codigo_pessoa = gerar_codigo_automatico and not any(
            coluna.casefold() == "codigo" for coluna in colunas
        )
        if gerar_codigo_pessoa:
            colunas_insert.append("codigo")
        # Fim do bloco que define se o código de pessoa será gerado.

        # Início do bloco que mapeia codigo_grupo para produto.grupo.
        indice_codigo_grupo = next(
            (
                indice
                for indice, coluna in enumerate(colunas)
                if coluna.casefold() == "codigo_grupo"
            ),
            None,
        )
        indice_grupo_origem = next(
            (
                indice
                for indice, coluna in enumerate(colunas)
                if coluna.casefold() == "grupo"
            ),
            None,
        )
        resolver_codigo_grupo = incluir_info_padrao and indice_codigo_grupo is not None
        indice_grupo_insert = None
        if resolver_codigo_grupo:
            colunas_insert = [
                coluna
                for indice, coluna in enumerate(colunas)
                if indice != indice_codigo_grupo
            ]
            if indice_grupo_origem is None:
                indice_grupo_insert = len(colunas_insert)
                colunas_insert.append("grupo")
            else:
                indice_grupo_insert = indice_grupo_origem - (
                    indice_codigo_grupo < indice_grupo_origem
                )
        # Fim do bloco que mapeia codigo_grupo para produto.grupo.

        indice_info = colunas_insert.index("info") if "info" in colunas_insert else None
        if incluir_info_padrao and indice_info is None:
            indice_info = len(colunas_insert)
            colunas_insert.append("info")
        # Fim do bloco que prepara a coluna info de produto.

        LOGGER.info(
            "Validação do CSV concluída: tabela=%s linhas=%s.",
            nome_tabela,
            linhas_analisadas,
        )
        confirmar = messagebox.askyesno(
            "Confirmar importação",
            f"Colunas do INSERT:\n{', '.join(colunas_insert)}\n\n"
            f"Inserir os registros na tabela '{nome_tabela}'?",
            parent=self.root,
        )
        if not confirmar:
            LOGGER.info("Importação CSV cancelada na confirmação: tabela=%s.", nome_tabela)
            return

        consulta_insert = sql.SQL("INSERT INTO {} ({}) VALUES ({})").format(
            sql.Identifier(*partes_tabela),
            sql.SQL(", ").join(sql.Identifier(coluna) for coluna in colunas_insert),
            sql.SQL(", ").join(sql.Placeholder() for _ in colunas_insert),
        )
        linhas_importadas = 0
        proximo_codigo = None

        def ler_linhas():
            nonlocal linhas_importadas, proximo_codigo
            # Início do bloco de leitura do CSV em UTF-8 para o INSERT.
            with open(caminho, "r", encoding="utf-8-sig", newline="") as arquivo:
                leitor = csv.reader(arquivo, delimiter=";")
                next(leitor, None)
                for numero_linha, valores in enumerate(leitor, start=2):
                    # Início/fim do bloco que pula linhas de CPF duplicado (clientes).
                    if numero_linha in linhas_csv_ignoradas:
                        continue
                    if not any(valor.strip() for valor in valores):
                        continue
                    if len(valores) > len(colunas) and any(
                        valor.strip() for valor in valores[len(colunas):]
                    ):
                        raise ValueError(
                            f"A linha {numero_linha} contém valores além das colunas do cabeçalho."
                        )
                    valores = valores[:len(colunas)]
                    valores.extend([""] * (len(colunas) - len(valores)))
                    # Início do bloco que troca codigo_grupo pelo grid da tabela grupo_produto.
                    valores_para_inserir = list(valores)
                    if resolver_codigo_grupo:
                        codigo_grupo_texto = valores_para_inserir[indice_codigo_grupo].strip()
                        if not codigo_grupo_texto:
                            raise ValueError(
                                f"Linha {numero_linha}: o campo codigo_grupo está vazio."
                            )
                        try:
                            codigo_grupo = int(codigo_grupo_texto)
                        except ValueError as erro:
                            raise ValueError(
                                f"Linha {numero_linha}: codigo_grupo '{codigo_grupo_texto}' "
                                "deve ser um número inteiro."
                            ) from erro
                        if codigo_grupo in codigos_ambiguos:
                            raise ValueError(
                                f"Linha {numero_linha}: codigo_grupo {codigo_grupo} "
                                "está associado a mais de um grid em grupo_produto."
                            )
                        if codigo_grupo not in grids_por_codigo:
                            raise ValueError(
                                f"Linha {numero_linha}: codigo_grupo {codigo_grupo} "
                                "não encontrado em grupo_produto."
                            )

                        valores_para_inserir.pop(indice_codigo_grupo)
                        if indice_grupo_origem is None:
                            valores_para_inserir.insert(
                                indice_grupo_insert,
                                grids_por_codigo[codigo_grupo],
                            )
                        else:
                            valores_para_inserir[indice_grupo_insert] = grids_por_codigo[
                                codigo_grupo
                            ]
                    # Fim do bloco que troca codigo_grupo pelo grid da tabela grupo_produto.

                    # Início do bloco que marca cada linha importada para produto.
                    if incluir_info_padrao:
                        if indice_info < len(valores_para_inserir):
                            valores_para_inserir[indice_info] = "Importado via importador AS."
                        else:
                            valores_para_inserir.append("Importado via importador AS.")
                    # Fim do bloco que marca cada linha importada para produto.
                    # Início do bloco que adiciona o código sequencial gerado para pessoa.
                    if gerar_codigo_pessoa:
                        valores_para_inserir.append(proximo_codigo)
                        proximo_codigo += 1
                    # Fim do bloco que adiciona o código sequencial gerado para pessoa.
                    linhas_importadas += 1
                    # Início do bloco de conversão para maiúsculas antes do INSERT.
                    valores_convertidos = []
                    for valor in valores_para_inserir:
                        if isinstance(valor, str):
                            valor = self._corrigir_mojibake(valor).upper() if valor else None
                        valores_convertidos.append(valor)
                    yield valores_convertidos
                    # Fim do bloco de conversão para maiúsculas antes do INSERT.
            # Fim do bloco de leitura do CSV em UTF-8 para o INSERT.

        try:
            # Início do bloco que envia os valores em UTF-8 ao INSERT.
            with psycopg.connect(
                **self.configuracao,
                connect_timeout=5,
                client_encoding="UTF8",
            ) as conexao:
                with conexao.cursor() as cursor:
                    # Início do bloco que carrega os grids de grupo_produto por codigo.
                    grids_por_codigo = {}
                    codigos_ambiguos = set()
                    if resolver_codigo_grupo:
                        esquema_grupo = partes_tabela[-2] if len(partes_tabela) == 2 else "public"
                        consulta_grupos = sql.SQL("SELECT {}, {} FROM {}").format(
                            sql.Identifier("codigo"),
                            sql.Identifier("grid"),
                            sql.Identifier(esquema_grupo, "grupo_produto"),
                        )
                        cursor.execute(consulta_grupos)
                        for codigo, grid in cursor.fetchall():
                            if codigo is None:
                                continue
                            codigo = int(codigo)
                            if codigo in grids_por_codigo and grids_por_codigo[codigo] != grid:
                                codigos_ambiguos.add(codigo)
                            else:
                                grids_por_codigo[codigo] = grid
                    # Fim do bloco que carrega os grids de grupo_produto por codigo.

                    # Início do bloco que consulta o último código de pessoa para gerar os próximos.
                    if gerar_codigo_pessoa:
                        consulta_ultimo_codigo = sql.SQL(
                            "SELECT COALESCE(MAX({}), 0) FROM {}"
                        ).format(
                            sql.Identifier("codigo"),
                            sql.Identifier(*partes_tabela),
                        )
                        cursor.execute(consulta_ultimo_codigo)
                        proximo_codigo = int(cursor.fetchone()[0]) + 1
                    # Fim do bloco que consulta o último código de pessoa para gerar os próximos.

                    # Início do INSERT dos registros do CSV (clientes: tabela pessoa).
                    cursor.executemany(consulta_insert, ler_linhas())
                    # Fim do INSERT dos registros do CSV (clientes: tabela pessoa).
            # Fim do bloco que envia os valores em UTF-8 ao INSERT.
        except (psycopg.Error, OSError, UnicodeError, csv.Error, ValueError) as erro:
            self.status.set("Falha ao importar o CSV.")
            LOGGER.error(
                "Importação CSV falhou: tabela=%s erro=%s.",
                nome_tabela,
                type(erro).__name__,
            )
            messagebox.showerror("Erro na importação", str(erro), parent=self.root)
            return

        LOGGER.info(
            "Importação CSV concluída: tabela=%s registros=%s duplicados=%s.",
            nome_tabela,
            linhas_importadas,
            len(linhas_csv_ignoradas),
        )
        # Início do bloco da tela de resumo com registros inseridos e duplicados (clientes).
        if validar_cpfs_duplicados:
            resumo_importacao = (
                f"Registros inseridos: {linhas_importadas}\n"
                f"Registros duplicados ignorados: {len(linhas_csv_ignoradas)}"
            )
        else:
            resumo_importacao = (
                f"{linhas_importadas} registro(s) inserido(s) na tabela '{nome_tabela}'."
            )
        if validar_cpfs_duplicados:
            self.status.set(
                f"Importação concluída: {linhas_importadas} registro(s) inserido(s) "
                f"em {nome_tabela}; {len(linhas_csv_ignoradas)} duplicado(s) ignorado(s)."
            )
        else:
            self.status.set(
                f"Importação concluída: {linhas_importadas} registro(s) "
                f"inserido(s) em {nome_tabela}."
            )
        messagebox.showinfo(
            "Importação concluída",
            resumo_importacao,
            parent=self.root,
        )
        # Fim do bloco da tela de resumo com registros inseridos e duplicados (clientes).

    # Início do bloco exclusivo de importação de grupos sem validações de produto.
    def importar_grupos_csv(self):
        self.importar_dados_csv(
            validar_codigos=False,
            revisar_caracteres=False,
            tabela_fixa="grupo_produto",
            solicitar_configuracao=True,
        )

    # Fim do bloco exclusivo de importação de grupos sem validações de produto.

    # Início do bloco exclusivo de importação de clientes para a tabela pessoa.
    def importar_clientes_csv(self):
        self.importar_dados_csv(
            validar_codigos=False,
            revisar_caracteres=False,
            tabela_fixa="pessoa",
            solicitar_configuracao=True,
            gerar_codigo_automatico=True,
            validar_cpfs_duplicados=True,
        )
    # Fim do bloco exclusivo de importação de clientes para a tabela pessoa.

    # Início do bloco exclusivo de importação de produtos para a tabela produto.
    def importar_produtos_csv(self):
        self.importar_dados_csv(
            tabela_fixa="produto",
            solicitar_configuracao=True,
        )

    # Fim do bloco exclusivo de importação de produtos para a tabela produto.
    # Fim do bloco de importação e validação de arquivos CSV.
    # Início do bloco que exige nova configuração e teste de conexão antes da importação.
    def _solicitar_configuracao_importacao(self):
        if self.configuracao is not None:
            return True

        self.abrir_configuracao()

        if self.configuracao is None:
            LOGGER.info("Importação cancelada antes de configurar a conexão PostgreSQL.")
            return False

        if not self._testar_configuracao(self.configuracao, self.root):
            self.configuracao = None
            LOGGER.warning("Importação cancelada após falha no teste de conexão PostgreSQL.")
            return False

        return True

    # Fim do bloco que exige nova configuração e teste de conexão antes da importação.
    def abrir_configuracao(self):
        # Abre o formulário modal usado para informar ou testar o acesso ao banco.
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
        # Preenche o formulário com os dados da sessão ou com as variáveis locais.
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
            # Valida porta e campos obrigatórios e devolve os dados prontos para uso.
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
            # Testa os dados digitados sem salvá-los na configuração da sessão.
            valores = ler_configuracao()
            if valores is not None:
                self._testar_configuracao(valores, janela)

        def aplicar():
            # Guarda os dados válidos apenas durante a sessão atual do aplicativo.
            valores = ler_configuracao()
            if valores is not None:
                self.configuracao = valores
                LOGGER.info(
                    "Configuração PostgreSQL aplicada: servidor=%s porta=%s banco=%s usuário=%s.",
                    valores["host"],
                    valores["port"],
                    valores["dbname"],
                    valores["user"],
                )
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
        # Testa a configuração atual ou pede que ela seja preenchida primeiro.
        if not self._exigir_configuracao():
            return
        self._testar_configuracao(self.configuracao, self.root)

    def _exigir_configuracao(self):
        # Impede operações no banco até os dados de acesso serem aplicados.
        if self.configuracao is not None:
            return True

        LOGGER.warning("Operação solicitada sem configuração PostgreSQL aplicada.")
        messagebox.showinfo(
            "Configuração necessária",
            "Informe e aplique os dados de acesso ao PostgreSQL para continuar.",
            parent=self.root,
        )
        self.abrir_configuracao()
        return False

    def _testar_configuracao(self, configuracao, parent):
        # Abre uma conexão curta e executa SELECT 1 para validar o acesso.
        LOGGER.info(
            "Teste de conexão iniciado: servidor=%s porta=%s banco=%s usuário=%s.",
            configuracao["host"],
            configuracao["port"],
            configuracao["dbname"],
            configuracao["user"],
        )
        try:
            with psycopg.connect(**configuracao, connect_timeout=5) as conexao:
                with conexao.cursor() as cursor:
                    cursor.execute("SELECT 1")
                    cursor.fetchone()
        except psycopg.Error as erro:
            self.status.set("Falha no teste de conexão.")
            LOGGER.error("Teste de conexão falhou (%s).", type(erro).__name__)
            messagebox.showerror("Falha na conexão", str(erro), parent=parent)
            return False

        LOGGER.info("Teste de conexão concluído com sucesso.")
        self.status.set(
            f"Conexão testada com sucesso: {configuracao['host']}:{configuracao['port']}/"
            f"{configuracao['dbname']}"
        )
        messagebox.showinfo(
            "Conexão OK", "Conexão com o PostgreSQL realizada com sucesso.", parent=parent
        )
        return True

    def abrir_execucao_comandos(self):
        # Cria a tela para executar consultas SELECT.
        if not self._exigir_configuracao():
            return

        LOGGER.info("Tela de execução de comandos SQL aberta.")
        janela = tk.Toplevel(self.root)
        janela.title("Executar comandos SQL")
        janela.geometry("900x600")
        janela.minsize(640, 420)
        janela.transient(self.root)

        conteudo = ttk.Frame(janela, padding=12)
        conteudo.pack(fill="both", expand=True)
        ttk.Label(
            conteudo,
            text="Comando permitido: SELECT.",
        ).pack(anchor="w")

        editor = scrolledtext.ScrolledText(conteudo, height=7, wrap="none")
        editor.pack(fill="x", pady=(6, 8))
        editor.insert("1.0", "SELECT * FROM sua_tabela LIMIT 100;")

        barra = ttk.Frame(conteudo)
        barra.pack(fill="x", pady=(0, 8))
        status = tk.StringVar(value="Informe um comando SQL e clique em Executar comando.")
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
            text="Executar comando",
            command=lambda: self.executar_comando_sql(editor, tabela, status, janela),
        ).pack(side="left")
        ttk.Label(barra, textvariable=status).pack(side="left", padx=12)

        tabela = ttk.Treeview(conteudo, show="headings")
        barra_vertical = ttk.Scrollbar(conteudo, orient="vertical", command=tabela.yview)
        barra_horizontal = ttk.Scrollbar(conteudo, orient="horizontal", command=tabela.xview)
        tabela.configure(yscrollcommand=barra_vertical.set, xscrollcommand=barra_horizontal.set)
        tabela.pack(side="top", fill="both", expand=True)
        barra_vertical.pack(side="right", fill="y")
        barra_horizontal.pack(side="bottom", fill="x")

    @staticmethod
    def _identificar_comando_sql(consulta):
        # Ignora comentários iniciais e identifica o primeiro comando SQL.
        restante = consulta.lstrip()
        while restante.startswith("--") or restante.startswith("/*"):
            if restante.startswith("--"):
                fim_comentario = restante.find("\n")
                if fim_comentario == -1:
                    return ""
                restante = restante[fim_comentario + 1 :].lstrip()
            else:
                fim_comentario = restante.find("*/", 2)
                if fim_comentario == -1:
                    return ""
                restante = restante[fim_comentario + 2 :].lstrip()

        correspondencia = re.match(r"[A-Za-z]+", restante)
        return correspondencia.group(0).upper() if correspondencia else ""

    def executar_comando_sql(self, editor, tabela, status, parent):
        # Valida o comando antes de enviar o SQL ao PostgreSQL.
        consulta = editor.get("1.0", "end").strip()
        if not consulta:
            LOGGER.warning("Execução SQL solicitada com o campo vazio.")
            messagebox.showwarning("Comando vazio", "Digite um comando SQL.", parent=parent)
            return
        comando = self._identificar_comando_sql(consulta)
        if comando != "SELECT":
            LOGGER.warning("Comando SQL não permitido: tipo=%s.", comando or "desconhecido")
            messagebox.showerror(
                "Comando não permitido",
                "Este menu aceita somente comandos SELECT.",
                parent=parent,
            )
            return

        tabela.delete(*tabela.get_children())
        nomes_colunas = []
        linhas = []

        # A transação somente leitura também bloqueia SELECTs com efeitos de escrita.
        LOGGER.info("Comando SQL iniciado: tipo=%s.", comando)
        try:
            with psycopg.connect(**self.configuracao, connect_timeout=5) as conexao:
                with conexao.cursor() as cursor:
                    # Exemplo de bloco de comando para consulta direta no código:
                    # cursor.execute("SELECT id, nome FROM sua_tabela LIMIT 20")
                    # resultado = cursor.fetchall()

                    cursor.execute("SET TRANSACTION READ ONLY")
                    cursor.execute(consulta, prepare=True)
                    if cursor.description is None:
                        nomes_colunas = []
                    else:
                        nomes_colunas = [coluna.name for coluna in cursor.description]
                        linhas = cursor.fetchmany(1001)
        except psycopg.Error as erro:
            status.set("Falha ao executar o comando.")
            LOGGER.error("Comando SQL falhou: tipo=%s erro=%s.", comando, type(erro).__name__)
            messagebox.showerror("Erro no comando", str(erro), parent=parent)
            return

        # Comandos sem retorno de colunas terminam após atualizar o status.
        if not nomes_colunas:
            tabela.configure(columns=())
            return

        # Mostra na grade as colunas e as primeiras 1.000 linhas retornadas.
        tabela.delete(*tabela.get_children())
        identificadores = [f"coluna_{indice}" for indice in range(len(nomes_colunas))]
        tabela.configure(columns=identificadores)
        for identificador, nome in zip(identificadores, nomes_colunas):
            tabela.heading(identificador, text=nome)
            tabela.column(identificador, width=150, minwidth=80, stretch=True)

        limite = 1000
        for linha in linhas[:limite]:
            tabela.insert(
                "",
                "end",
                values=[
                    "" if valor is None else str(self._decodificar_bytes(valor))
                    for valor in linha
                ],
            )

        quantidade = min(len(linhas), limite)
        complemento = " (exibindo as primeiras 1000)" if len(linhas) > limite else ""
        if comando == "SELECT":
            LOGGER.info("Consulta SELECT concluída: linhas=%s.", len(linhas))
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


# Cria a janela principal apenas quando este arquivo é executado diretamente.
if __name__ == "__main__":
    try:
        if solicitar_login() is None:
            LOGGER.info("Aplicativo encerrado sem login.")
            sys.exit(0)
        janela_principal = tk.Tk()
        app = ImportadorApp(janela_principal)
        janela_principal.mainloop()
    except Exception as erro:
        LOGGER.critical("Falha fatal no aplicativo (%s).", type(erro).__name__)
        raise
    else:
        LOGGER.info("Aplicativo encerrado normalmente.")