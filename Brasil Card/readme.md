# Sistema de Transações - BCARD

## Descrição
Aplicação desenvolvida em VB6 para gerenciamento de transações de cartões, com funcionalidades de cadastro, consulta, exclusão e geração de relatórios em Excel.

---

## Funcionalidades
- Cadastro de transações
- Atualização de registros
- Exclusão de transações
- Busca com filtros (cartão, data e valor)
- Geração de relatório por período
- Geração de relatório do último mês

---

## Tecnologias
- VB6
- SQL Server
- ADO (ADODB)
- Excel (Automação)

---

## Banco de Dados
Os scripts estão na pasta `/SQL` e incluem:

- Criação das tabelas
- Stored Procedure
- View
- Function

---

## Relatórios
Exemplos de relatórios disponíveis na pasta `/Relatorios`:

- Relatório por período
- Relatório do último mês

---

## Como executar
1. Rodar os scripts SQL na ordem:
   - 01 - create tables.sql
   - 02 - create procedure.sql
   - 03 - create function.sql
   - 04 - create view.sql

2. Abrir o projeto VB6 (`.vbp`)

3. Configurar conexão com banco

4. Executar o sistema

---
Yuri Carvalho