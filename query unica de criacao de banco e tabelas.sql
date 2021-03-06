USE [PaisDb]
GO
/****** Object:  Table [dbo].[InformacaoPais]    Script Date: 20/02/2021 16:53:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[InformacaoPais](
	[Cod] [int] NOT NULL,
	[sISOCode] [nchar](3) NULL,
	[sName] [nchar](50) NULL,
	[sCapitalCity] [nchar](50) NULL,
	[sPhoneCode] [int] NULL,
	[sContinentCode] [nchar](3) NULL,
	[sCurrencyISOCode] [nchar](4) NULL,
	[sCountryFlag] [nchar](150) NULL,
PRIMARY KEY CLUSTERED 
(
	[Cod] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Linguas]    Script Date: 20/02/2021 16:53:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Linguas](
	[IsoCode] [varchar](4) NULL,
	[sNameLang] [varchar](50) NULL,
	[Cod_pais] [int] NOT NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Linguas]  WITH CHECK ADD  CONSTRAINT [lang_fk] FOREIGN KEY([Cod_pais])
REFERENCES [dbo].[InformacaoPais] ([Cod])
GO
ALTER TABLE [dbo].[Linguas] CHECK CONSTRAINT [lang_fk]
GO
