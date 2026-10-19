# Importador de Dados

Aplicativo desktop para consultar e importar dados do PostgreSQL e consultar o Microsoft SQL Server. Versão Beta 1.4.

## Menus

- **Menu > Configurar conexão...**: informa os dados PostgreSQL para a sessão atual.
- **Menu > Testar conexão**: testa os dados PostgreSQL aplicados.
- **Menu > Executar comandos...**: executa consultas `SELECT` somente leitura no PostgreSQL.
- **Obter dados > PostgreSQL**: consulta uma tabela PostgreSQL e permite exportar o resultado para XLSX.
- **Obter dados > SQL Server**: consulta uma tabela SQL Server via ODBC e permite exportar o resultado para XLSX.
- **Importar > Importar dados por CSV...**: importa um arquivo CSV separado por ponto e vírgula para uma tabela PostgreSQL.
- **Ajuda > Sobre**: mostra a descrição e versão do aplicativo.

## Preparar o ambiente

No PowerShell, na pasta do projeto:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
Copy-Item .env.example .env
python -m pip install -r requirements.txt
python main.py
```

Edite `.env` com os dados locais. O arquivo é ignorado pelo Git e não deve ser enviado ao GitHub. A conexão só é aberta quando uma operação é iniciada.

### SQL Server

Instale o Microsoft ODBC Driver 17 ou 18 para SQL Server, com arquitetura compatível com o aplicativo. Configure no `.env` `MSSQL_DRIVER`, `MSSQL_SERVER`, `MSSQL_PORT`, `MSSQL_DATABASE`, `MSSQL_USER` e `MSSQL_PASSWORD`; ou informe os campos na janela **Obter dados > SQL Server**. A opção para confiar no certificado deve ser usada somente em desenvolvimento.

## Importar arquivo CSV

Use **Importar > Importar dados por CSV...**. Informe a tabela de destino (opcionalmente `esquema.tabela`) e selecione o arquivo. O delimitador deve ser ponto e vírgula (`;`) e a primeira linha deve conter os nomes das colunas da tabela PostgreSQL.

Antes do `INSERT`, o aplicativo verifica as colunas `codigo` e `codigo_Barra`. Ambas podem ficar vazias. Quando preenchidos, os valores de cada coluna não podem se repetir. Se houver duplicidades, a importação é cancelada e as ocorrências são registradas em `relatorio_validacao_importacao.txt`, ao lado do executável ou de `main.py`. O relatório pode conter valores do arquivo; revise-o antes de compartilhar.

Todos os `INSERT`s do CSV são feitos na mesma transação: em caso de erro, a transação é revertida. Arquivos com BOM UTF-8 são aceitos; células vazias são importadas como `NULL`.

Ao importar para a tabela `produto`, a coluna `info` recebe `importado via importador AS`. Esse valor é adicionado ao `INSERT` quando a coluna não está no cabeçalho e substitui o valor do arquivo quando ela já existe.

## Consultas e exportação

As telas **Obter dados > PostgreSQL** e **Obter dados > SQL Server** recebem os dados de conexão e uma tabela (ou `esquema.tabela`), executam `SELECT *` e mostram até 1.000 linhas na grade. O botão **Salvar resultado em Excel...** exporta o resultado consultado para `.xlsx`, respeitando o limite de linhas do formato.

## Logs

O aplicativo escreve eventos e erros em `importador.txt`, ao lado de `Importador.exe` quando empacotado ou ao lado de `main.py` no modo de desenvolvimento. O log não contém senhas nem o texto literal dos comandos SQL.

## Gerar o executável

Com o ambiente virtual ativado:

```powershell
python -m pip install pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed --name Importador --specpath build --workpath build --distpath dist --collect-all psycopg --collect-all openpyxl --collect-all dotenv main.py
```

O arquivo gerado fica em `dist\Importador.exe`. Para SQL Server, o Microsoft ODBC Driver correspondente também precisa estar instalado no Windows.