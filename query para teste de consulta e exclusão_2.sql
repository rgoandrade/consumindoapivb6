/****** Script do comando SelectTopNRows de SSMS  ******/
SELECT TOP (1000) [IsoCode]
      ,[sNameLang]
      ,[Cod_pais]
  FROM [PaisDb].[dbo].[Linguas]

  delete FROM [PaisDb].[dbo].[Linguas]