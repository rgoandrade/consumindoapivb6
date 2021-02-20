USE [PaisDb]
GO

/****** Object:  Table [dbo].[Linguas]    Script Date: 20/02/2021 16:51:43 ******/
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

