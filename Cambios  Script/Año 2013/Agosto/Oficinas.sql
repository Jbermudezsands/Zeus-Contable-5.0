/*
   viernes, 06 de febrero de 201511:38:40 a.m.
   Usuario: sa
   Servidor: JUAN\SQL2005
   Base de datos: SistemaContableALFA
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar esta secuencia de comandos detalladamente antes de ejecutarla fuera del contexto del diseñador de base de datos.*/
BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
CREATE TABLE dbo.Oficinas
	(
	IdReg int NOT NULL IDENTITY (1, 1),
	Descripcion nvarchar(50) NULL,
	FechaReg smalldatetime NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Oficinas ADD CONSTRAINT
	PK_Oficinas PRIMARY KEY CLUSTERED 
	(
	IdReg
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
