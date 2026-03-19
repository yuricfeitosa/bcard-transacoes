USE [BCARD]
GO

/****** Object:  UserDefinedFunction [dbo].[fn_CategoriaTransacao]    Script Date: 18/03/2026 22:08:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION	[dbo].[fn_CategoriaTransacao] (@Valor DECIMAL(10,2))
RETURNS VARCHAR(10)
AS
BEGIN
	DECLARE @Categoria VARCHAR(10)
	
	IF @Valor > 1000 
		SET @Categoria = 'Alta'
	ELSE IF @Valor BETWEEN 500 AND 1000
		SET @Categoria = 'Média'
	ELSE
		SET @Categoria = 'Baixa'

	RETURN @Categoria
END
GO

