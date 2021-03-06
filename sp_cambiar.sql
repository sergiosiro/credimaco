USE [creditosSQL]
GO
/****** Object:  Table [dbo].[ESTUDIOS]    Script Date: 04/08/2018 12:37:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ESTUDIOS](
	[IDEstudio] [int] NOT NULL,
	[Nombre] [nchar](50) NOT NULL,
	[IDProvincia] [int] NOT NULL,
	[Predeterminado] [bit] NOT NULL,
 CONSTRAINT [PK_ESTUDIOS] PRIMARY KEY CLUSTERED 
(
	[IDEstudio] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CREDITOSBLOQUEADOS]    Script Date: 04/08/2018 12:37:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[CREDITOSBLOQUEADOS](
	[IdCredito] [int] NOT NULL,
	[Secuencia] [tinyint] NOT NULL,
	[FechaEstado] [datetime] NOT NULL,
	[Estado] [char](2) NOT NULL,
	[IdEstudio] [int] NULL,
	[FechaEnvio] [date] NULL,
	[Observaciones] [nvarchar](200) NOT NULL,
 CONSTRAINT [PK_CREDITOSBLOQUEADOS] PRIMARY KEY CLUSTERED 
(
	[IdCredito] ASC,
	[Secuencia] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  StoredProcedure [dbo].[SeleccionarEstudio]    Script Date: 04/08/2018 12:36:58 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[SeleccionarEstudio]

	@IdEstudio				INT

AS
BEGIN
	SET NOCOUNT ON;

	SELECT * 
		FROM ESTUDIOS
		WHERE IDEstudio = @IdEstudio
	
END
GO
/****** Object:  StoredProcedure [dbo].[InfoBloqueado]    Script Date: 04/08/2018 12:36:57 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[InfoBloqueado] 

	@IdCredito			INT
	
AS
BEGIN
	SET NOCOUNT ON;

	SELECT * 
		FROM CREDITOSBLOQUEADOS 
		WHERE IdCredito = @IdCredito
		  AND Secuencia = (SELECT MAX(Secuencia) FROM CREDITOSBLOQUEADOS WHERE IdCredito = @IdCredito)

END
GO
