USE [BCARD]
GO

/****** Object:  View [dbo].[vw_TransacoesClientes]    Script Date: 18/03/2026 22:10:29 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE VIEW [dbo].[vw_TransacoesClientes]
AS
SELECT
	C.CLI_NOME,
	T.TRA_NUM_CARTAO,
	T.TRA_VALOR,
	FORMAT(T.TRA_DATA,'dd/MM/yyyy') TRA_DATA,
	T.TRA_DESCRICAO,
	dbo.fn_CategoriaTransacao(T.TRA_VALOR) AS CATEGORIA,
	T.TRA_COD
FROM BCA_TRANSACOES T 
INNER JOIN BCA_CLIENTES C ON T.TRA_NUM_CARTAO = C.CLI_NUM_CARTAO
GO


