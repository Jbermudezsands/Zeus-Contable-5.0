/*
   Mi�rcoles, 09 de Septiembre de 2009 06:59:04 p.m.
   Usuario: 
   Servidor: JUAN\SQL2000
   Base de datos: SistemaContableMarAzul
   Aplicaci�n: MS SQLEM - Data Tools
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
COMMIT
BEGIN TRANSACTION
ALTER TABLE dbo.DetalleMemoriza ADD CONSTRAINT
	FK_DetalleMemoriza_IndiceMemoriza FOREIGN KEY
	(
	IdMemoria
	) REFERENCES dbo.IndiceMemoriza
	(
	IdMemoria
	) ON UPDATE CASCADE
	 ON DELETE CASCADE
	
GO
COMMIT
