# Importador

Aplicacao desktop em Python para consultar e importar dados para o PostgreSQL.

## Preparar o ambiente

No PowerShell, na pasta do projeto:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
Copy-Item .env.example .env
python -m pip install -r requirements.txt
python main.py
```

Edite o arquivo `.env` com os dados locais do PostgreSQL. Esse arquivo e ignorado pelo Git e nao deve ser enviado ao GitHub.

Para importar produtos, use **Importar > Importar produtos por XLSX...**. A primeira linha da aba ativa deve conter os nomes das colunas existentes na tabela `produto`.