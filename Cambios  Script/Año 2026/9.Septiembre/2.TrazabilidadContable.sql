CREATE TABLE dbo.TrazabilidadContable
(
    IdTrazabilidad BIGINT IDENTITY(1,1) NOT NULL,

    -- Origen del documento
    ModuloOrigen NVARCHAR(30) NOT NULL,
    TipoDocumento NVARCHAR(50) NOT NULL,

    -- Identificación del documento origen
    IdDocumento NVARCHAR(100) NULL,
    NumeroDocumento NVARCHAR(100) NULL,

    -- Movimiento contable relacionado
    FechaTransaccion SMALLDATETIME NOT NULL,
    NPeriodo INT NOT NULL,
    NumeroMovimiento INT NOT NULL,

    -- Información adicional
    FechaRegistro DATETIME2(3) NOT NULL
        CONSTRAINT DF_TrazabilidadContable_FechaRegistro
        DEFAULT (SYSDATETIME()),

    UsuarioRegistro NVARCHAR(100) NULL,

    CONSTRAINT PK_TrazabilidadContable
        PRIMARY KEY CLUSTERED (IdTrazabilidad)
);