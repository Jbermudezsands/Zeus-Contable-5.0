/*
   miércoles, 19 de diciembre de 201811:57:31 a.m.
   Usuario: 
   Servidor: JUANBERMUDEZ\SQL2014
   Base de datos: Contabilidad1
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar este script detalladamente antes de ejecutarlo fuera del contexto del diseñador de base de datos.*/
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
CREATE TABLE dbo.ProyectosContratistas
	(
	IdProyecto int NOT NULL IDENTITY (1, 1),
	NombreProyecto nvarchar(50) NULL,
	CodigoContratista nvarchar(50) NULL,
	FechaContrato smalldatetime NULL,
	FechaFinalizacion smalldatetime NULL,
	Descripcion_Trabajos nvarchar(250) NULL,
	MontoContratado decimal(18, 2) NULL,
	PagoAnterioresManual decimal(18, 2) NULL,
	Moneda nvarchar(20) NULL,
	Observaciones nvarchar(250) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.ProyectosContratistas ADD CONSTRAINT
	PK_ProyectosContratistas PRIMARY KEY CLUSTERED 
	(
	IdProyecto
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
ALTER TABLE dbo.ProyectosContratistas SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
select Has_Perms_By_Name(N'dbo.ProyectosContratistas', 'Object', 'ALTER') as ALT_Per, Has_Perms_By_Name(N'dbo.ProyectosContratistas', 'Object', 'VIEW DEFINITION') as View_def_Per, Has_Perms_By_Name(N'dbo.ProyectosContratistas', 'Object', 'CONTROL') as Contr_Per 