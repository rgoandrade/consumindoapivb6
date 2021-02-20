/****** Script do comando SelectTopNRows de SSMS  ******/
SELECT TOP (1000) [Cod]
      ,[sISOCode]
      ,[sName]
      ,[sCapitalCity]
      ,[sPhoneCode]
      ,[sContinentCode]
      ,[sCurrencyISOCode]
      ,[sCountryFlag]
  FROM [PaisDb].[dbo].[InformacaoPais]



  delete FROM [PaisDb].[dbo].[InformacaoPais]

                                          