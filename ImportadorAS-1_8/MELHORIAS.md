# Controle de melhorias – Importador AS

Documento de controle com as 10 últimas melhorias registradas no repositório, mais as alterações do menu Clientes.

## Menu Importar > Clientes

| # | Melhoria | Detalhe |
|---|----------|---------|
| P1 | Novo menu **Importar > Clientes** | Importa CSV para a tabela fixa `pessoa`, seguindo o conceito da importação de Grupos. |
| P2 | Código automático | Se o CSV não tem a coluna `codigo`, usa `MAX(codigo)+1` da tabela e soma 1 a cada registro; se tem, preserva os valores do arquivo. |
| P3 | Validação de CPF duplicado | Antes do INSERT, mantém a primeira ocorrência e ignora as demais; o CSV original não é alterado. |
| P4 | Relatório Excel de duplicados | Gera `relatorio_cpfs_duplicados.xlsx` com linha do CSV, dados e ação (Mantido / Ignorado como duplicado). |
| P5 | Tela de resumo | Ao final informa registros inseridos e duplicados ignorados. |
| P6 | Comentários de início/fim | Blocos do fluxo de Clientes identificados no `main.py`. |

## Últimas 10 melhorias commitadas (antes do menu Clientes)

| # | Data | Commit | Melhoria |
|---|------|--------|----------|
| 1 | 2026-10-29 | `467bffc` | Atualiza importação para Beta 1.8 |
| 2 | 2026-10-27 | `b8b161e` | Mapeamento de grupos na importação de produtos |
| 3 | 2026-10-27 | `c7114f2` | Renomeia botão de importação CSV |
| 4 | 2026-10-23 | `99b9cc0` | Correção da importação CSV |
| 5 | 2026-10-23 | `6f867b9` | Marcação de blocos do CSV e renomeação do botão de ajuste |
| 6 | 2026-10-22 | `8c70dd7` | Revisão e normalização de acentos no CSV |
| 7 | 2026-10-20 | `f798165` | Manuais e login atualizados (Beta 1.5) |
| 8 | 2026-10-19 | `836f6a6` | Tela de login e info padrão na importação de produto |
| 9 | 2026-10-16 | `6fee445` | Importação de CSV delimitado por ponto e vírgula |
| 10 | 2026-10-06 | `0c9c2f9` | Permite campos de validação de código em branco |

## Observações

- Não houve teste da importação de Clientes contra banco real; a lógica de CPF/Excel foi validada com snippet e a sintaxe com Pylance.
