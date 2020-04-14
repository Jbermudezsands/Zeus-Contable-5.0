/*
   Sábado, 22 de Agosto de 2009 03:07:40 p.m.
   Usuario: 
   Servidor: SERVIDOR1
   Base de datos: SistemaContable
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
ALTER TABLE dbo.Cuentas ADD
	SubDivicion nvarchar(50) NULL,
	CausaIva bit NULL,
	CausaRetencion bit NULL,
	DescRetencion nvarchar(50) NULL,
	Nombre1 nvarchar(50) NULL,
	Nombre2 nvarchar(50) NULL,
	Apellido1 nvarchar(50) NULL,
	Apellido2 nvarchar(50) NULL,
	Cedula nvarchar(50) NULL,
	RUC nvarchar(50) NULL,
        Telefono nvarchar(50) NULL,
        Direccion nvarchar(250) NULL
GO
COMMIT
