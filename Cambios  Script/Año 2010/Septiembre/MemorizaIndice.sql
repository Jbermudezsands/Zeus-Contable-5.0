/*
   Miércoles, 09 de Septiembre de 2009 05:23:06 p.m.
   Usuario: 
   Servidor: JUAN\SQL2000
   Base de datos: SistemaContableMarAzul
   Aplicación: MS SQLEM - Data Tools
*/

BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
CREATE TABLE dbo.IndiceMemoriza
	(
	IdMemoria int NOT NULL IDENTITY (1, 1),
	CicloNombre nvarchar(80) NULL,
	Frecuencia nvarchar(50) NULL,
	SiguienteVez datetime NULL,
	NVeces float(53) NULL,
	TipoMemoria nvarchar(50) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.IndiceMemoriza ADD CONSTRAINT
	DF_IndiceMemoriza_Frecuencia DEFAULT 1 FOR Frecuencia
GO
ALTER TABLE dbo.IndiceMemoriza ADD CONSTRAINT
	DF_IndiceMemoriza_NVeces DEFAULT 0 FOR NVeces
GO
ALTER TABLE dbo.IndiceMemoriza ADD CONSTRAINT
	PK_IndiceMemoriza PRIMARY KEY CLUSTERED 
	(
	IdMemoria
	) ON [PRIMARY]

GO
COMMIT
