VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmContactos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro Central de Contratistas"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9330
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4200
      Picture         =   "FrmContactos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CmdJustifica 
      Caption         =   "Proyectos"
      Height          =   375
      Left            =   7200
      TabIndex        =   45
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   43
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   480
      TabIndex        =   42
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   6720
      TabIndex        =   41
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   5520
      TabIndex        =   40
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   2880
      TabIndex        =   39
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   1680
      TabIndex        =   38
      Top             =   5640
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo DBContratista 
      Height          =   315
      Left            =   2280
      TabIndex        =   17
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DBGrupos 
      Height          =   315
      Left            =   2280
      TabIndex        =   16
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaNacceso 
      Height          =   330
      Left            =   2520
      Top             =   9120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaNacceso"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaCuentas 
      Height          =   330
      Left            =   0
      Top             =   9120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaCuentas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaGrupoCuentas 
      Height          =   330
      Left            =   2520
      Top             =   8760
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaGrupoCuentas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaContratista 
      Height          =   375
      Left            =   0
      Top             =   8760
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaContratista"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CmdBuscarEmpleado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4200
      Picture         =   "FrmContactos.frx":014E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   480
      Width           =   375
   End
   Begin MSComCtl2.DTPicker TxtFechaFinaliza 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Format          =   78839809
      CurrentDate     =   37992
   End
   Begin MSComCtl2.DTPicker TxtFechaContrata 
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Format          =   78839809
      CurrentDate     =   37992
   End
   Begin VB.TextBox TxtEmail 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox TxtFax 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox TxtIdiomas 
      Height          =   615
      Left            =   6480
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox TxtCodigoPostal 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox TxtCiudad 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox TxtRazones 
      Height          =   615
      Left            =   6480
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox TxtTrabAnteriores 
      Height          =   615
      Left            =   6480
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox TxtRecomendaciones 
      Height          =   615
      Left            =   6480
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox TxtCursos 
      Height          =   645
      Left            =   6480
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Txtdireccion 
      Height          =   645
      Left            =   2280
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox TxtTelefono 
      Height          =   285
      Left            =   2280
      MaxLength       =   150
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox TxtTelEmergencia 
      Height          =   525
      Left            =   6480
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   480
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":029C
      TabIndex        =   18
      Top             =   4560
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":0328
      TabIndex        =   19
      Top             =   3840
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":0392
      TabIndex        =   20
      Top             =   3120
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":0404
      TabIndex        =   21
      Top             =   2400
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":0480
      TabIndex        =   22
      Top             =   1200
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":04F0
      TabIndex        =   23
      Top             =   480
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":057A
      TabIndex        =   24
      Top             =   840
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":0604
      TabIndex        =   25
      Top             =   1920
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":068A
      TabIndex        =   26
      Top             =   2760
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":06F8
      TabIndex        =   27
      Top             =   3480
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":0774
      TabIndex        =   28
      Top             =   4200
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   495
      Left            =   4800
      OleObjectBlob   =   "FrmContactos.frx":0800
      TabIndex        =   29
      Top             =   3840
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   375
      Left            =   4800
      OleObjectBlob   =   "FrmContactos.frx":0884
      TabIndex        =   30
      Top             =   1680
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   495
      Left            =   4800
      OleObjectBlob   =   "FrmContactos.frx":090A
      TabIndex        =   31
      Top             =   480
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "FrmContactos.frx":09AC
      TabIndex        =   32
      Top             =   120
      Width           =   3615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
      Height          =   255
      Left            =   4800
      OleObjectBlob   =   "FrmContactos.frx":0A86
      TabIndex        =   33
      Top             =   120
      Width           =   4095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   495
      Left            =   4800
      OleObjectBlob   =   "FrmContactos.frx":0B62
      TabIndex        =   34
      Top             =   1080
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   255
      Left            =   4800
      OleObjectBlob   =   "FrmContactos.frx":0C02
      TabIndex        =   35
      Top             =   2520
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
      Height          =   255
      Left            =   4800
      OleObjectBlob   =   "FrmContactos.frx":0C80
      TabIndex        =   36
      Top             =   3240
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
      Height          =   255
      Left            =   5520
      OleObjectBlob   =   "FrmContactos.frx":0CFE
      TabIndex        =   37
      Top             =   4800
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   255
      Left            =   600
      OleObjectBlob   =   "FrmContactos.frx":0D7C
      TabIndex        =   44
      Top             =   4920
      Width           =   3255
   End
End
Attribute VB_Name = "FrmContactos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
DtaContratista.Recordset.MovePrevious
If DtaContratista.Recordset.BOF Then
   DtaContratista.Recordset.MoveNext
   MsgBox "Este es el Primer Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
  Me.DBContratista.Text = DtaContratista.Recordset!CodigoCuenta
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub CmdBorrar_Click()
On Error GoTo TipoErrs
    If DtaContratista.Recordset.BOF And DtaContratista.Recordset.EOF Then
        MsgBox "No Existen Registros de Contratistas Actualmente", vbInformation
        Exit Sub
    End If
  Dim Respuesta, Rsp
  
  Set Rsp = DtaContratista.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando: " & Me.TxtNombre.Text)
   If Respuesta = 6 Then
     Criterio = "CodigoCuenta='" & Me.DBContratista & "'"
     DtaContratista.Recordset.MoveFirst
     Me.DtaContratista.Recordset.Find (Criterio)
    If Not DtaContratista.Recordset.EOF Then
     DtaContratista.Recordset.Delete
   '/////////Borra registro de cuentas/////////////
     Criterio = "CodCuentas='" & Me.DBContratista.Text & "'"
     DtaCuentas.Recordset.MoveFirst
     Me.DtaCuentas.Recordset.Find (Criterio)
    If Not DtaCuentas.Recordset.EOF Then
      DtaCuentas.Recordset.Delete
    End If
    End If
  
  End If
  DtaContratista.Refresh
Me.DBContratista.Text = ""
  Me.TxtNombre.Text = ""
 Me.TxtDireccion.Text = ""
 Me.TxtCiudad.Text = ""
 Me.TxtTelefono.Text = ""
 Me.TxtCodigoPostal.Text = ""
 Me.TxtFax.Text = ""
 Me.TxtEmail.Text = ""
 Me.TxtFechaContrata.Value = Format(Now, "dd/mm/yyyy")
 Me.TxtFechaFinaliza.Value = Format(Now, "dd/mm/yyyy")
 Me.TxtTelEmergencia.Text = ""
 Me.TxtCursos.Text = ""
 Me.TxtRazones.Text = ""
 Me.TxtTrabAnteriores.Text = ""
 Me.TxtRecomendaciones.Text = ""
 Me.TxtIdiomas.Text = ""
 Me.DBGrupos.Text = ""
 Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub CmdBuscarEmpleado_Click()
On Error GoTo TipoErrs
 QueProducto = "Contratista"
 FrmConsulta.Show 1
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo TipoErrs
Me.DtaContratista.Refresh

If Me.DBGrupos.Text = "" Then
 MsgBox "Seleccione el grupo de Cuentas", vbCritical, "Sistema Contable"
 Me.DBGrupos.SetFocus
 Exit Sub
End If

Criterio = "CodigoCuenta='" & Me.DBContratista.Text & "'"
If DtaContratista.Recordset.RecordCount <> 0 Then DtaContratista.Recordset.MoveFirst
Me.DtaContratista.Recordset.Find (Criterio)
If DtaContratista.Recordset.EOF Then
  Criterio = "CodCuentas='" & Me.DBContratista.Text & "'"
  DtaCuentas.Recordset.MoveFirst
  Me.DtaCuentas.Recordset.Find (Criterio)
  If DtaCuentas.Recordset.EOF Then
   DtaCuentas.Recordset.AddNew
   DtaCuentas.Recordset!CodCuentas = Me.DBContratista.Text
   DtaCuentas.Recordset!DescripcionCuentas = Me.TxtNombre.Text
   DtaCuentas.Recordset!TipoCuenta = "Cuentas x Pagar"
   DtaCuentas.Recordset!CodGrupo = CodGrupo
   DtaCuentas.Recordset!SaldoActual = 0#
   DtaCuentas.Recordset!TipoMoneda = "Dólares"
   DtaCuentas.Recordset.Update
  End If

  DtaContratista.Recordset.AddNew
   DtaContratista.Recordset!CodigoCuenta = Me.DBContratista.Text
   DtaContratista.Recordset!Beneficiario = Me.TxtNombre.Text
   DtaContratista.Recordset!Direccion = Me.TxtDireccion.Text
   DtaContratista.Recordset!Ciudad = Me.TxtCiudad.Text
   DtaContratista.Recordset!Telefono = Me.TxtTelefono.Text
   DtaContratista.Recordset!CP = Me.TxtCodigoPostal.Text
   DtaContratista.Recordset!Fax = Me.TxtFax.Text
   DtaContratista.Recordset!Email = Me.TxtEmail.Text
   DtaContratista.Recordset!FechaContratacion = Me.TxtFechaContrata.Value
   DtaContratista.Recordset!FechaFinalizacion = Me.TxtFechaFinaliza.Value
   DtaContratista.Recordset!TelefonoEmergencia = Me.TxtTelEmergencia.Text
   DtaContratista.Recordset!CursosRecibidos = Me.TxtCursos.Text
   DtaContratista.Recordset!RazonesContrato = Me.TxtRazones.Text
   DtaContratista.Recordset!TrabAnteriores = Me.TxtTrabAnteriores.Text
   DtaContratista.Recordset!Recomendaciones = Me.TxtRecomendaciones.Text
   DtaContratista.Recordset!IdiomaDomina = Me.TxtIdiomas.Text
 DtaContratista.Recordset.Update
Else
  'DtaContratista.Recordset.Edit
   DtaContratista.Recordset!Beneficiario = Me.TxtNombre.Text
   DtaContratista.Recordset!Direccion = Me.TxtDireccion.Text
   DtaContratista.Recordset!Ciudad = Me.TxtCiudad.Text
   DtaContratista.Recordset!Telefono = Me.TxtTelefono.Text
   DtaContratista.Recordset!CP = Me.TxtCodigoPostal.Text
   DtaContratista.Recordset!Fax = Me.TxtFax.Text
   DtaContratista.Recordset!Email = Me.TxtEmail.Text
   DtaContratista.Recordset!FechaContratacion = Me.TxtFechaContrata.Value
   DtaContratista.Recordset!FechaFinalizacion = Me.TxtFechaFinaliza.Value
   DtaContratista.Recordset!TelefonoEmergencia = Me.TxtTelEmergencia.Text
   DtaContratista.Recordset!CursosRecibidos = Me.TxtCursos.Text
   DtaContratista.Recordset!RazonesContrato = Me.TxtRazones.Text
   DtaContratista.Recordset!TrabAnteriores = Me.TxtTrabAnteriores.Text
   DtaContratista.Recordset!Recomendaciones = Me.TxtRecomendaciones.Text
   DtaContratista.Recordset!IdiomaDomina = Me.TxtIdiomas.Text
 DtaContratista.Recordset.Update
End If

Me.DBContratista.Text = ""
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdJustifica_Click()
'If Not IsNull(DtaContratista.Recordset("MontoAcordado")) Then
' MontoAcordado = DtaContratista.Recordset("MontoAcordado")
' FrmJustificacion.TxtCosto = DtaContratista.Recordset("MontoAcordado")
'End If
'FrmJustificacion.TxtFechaIni = Me.TxtFechaContrata
'FrmJustificacion.TxtFechaT = Me.TxtFechaFinaliza
'FrmJustificacion.Label1 = Me.TxtNombre.Text
'FrmJustificacion.Show 1
FrmContratistasProyectos.CodigoContratista = Me.DBContratista.Text
FrmContratistasProyectos.Show 1
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
DtaContratista.Recordset.MoveNext
If DtaContratista.Recordset.EOF Then
   DtaContratista.Recordset.MovePrevious
   MsgBox "Este es el Ultimo Registro", vbInfoContabilidadtion, "Control de contratista Contabilidad"
Else
  Me.DBContratista.Text = DtaContratista.Recordset!CodigoCuenta
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtFax.Text = FrmConsulta.Cuenta
End Sub

Private Sub DBContratista_Change()
On Error GoTo TipoErrs
If DtaContratista.Recordset.RecordCount = 0 Then
    Exit Sub
End If
Criterio = "CodigoCuenta='" & Me.DBContratista.Text & "'"
DtaContratista.Recordset.MoveFirst
Me.DtaContratista.Recordset.Find (Criterio)
If DtaContratista.Recordset.EOF Then
  Me.TxtNombre.Text = ""
 Me.TxtDireccion.Text = ""
 Me.TxtCiudad.Text = ""
 Me.TxtTelefono.Text = ""
 Me.TxtCodigoPostal.Text = ""
 Me.TxtFax.Text = ""
 Me.TxtEmail.Text = ""
 Me.TxtFechaContrata.Value = Format(Now, "dd/mm/yyyy")
 Me.TxtFechaFinaliza.Value = Format(Now, "dd/mm/yyyy")
 Me.TxtTelEmergencia.Text = ""
 Me.TxtCursos.Text = ""
 Me.TxtRazones.Text = ""
 Me.TxtTrabAnteriores.Text = ""
 Me.TxtRecomendaciones.Text = ""
 Me.TxtIdiomas.Text = ""
 Me.DBGrupos.Text = ""
Else
  Me.CmdJustifica.Enabled = True
 Me.TxtNombre.Text = DtaContratista.Recordset!Beneficiario
 Me.TxtDireccion.Text = DtaContratista.Recordset!Direccion
 Me.TxtCiudad.Text = DtaContratista.Recordset!Ciudad
 Me.TxtTelefono.Text = DtaContratista.Recordset!Telefono
 Me.TxtCodigoPostal.Text = DtaContratista.Recordset!CP
 Me.TxtFax.Text = DtaContratista.Recordset!Fax
 Me.TxtEmail.Text = DtaContratista.Recordset!Email
 Me.TxtFechaContrata.Value = DtaContratista.Recordset!FechaContratacion
 Me.TxtFechaFinaliza.Value = DtaContratista.Recordset!FechaFinalizacion
 Me.TxtTelEmergencia.Text = DtaContratista.Recordset!TelefonoEmergencia
 Me.TxtCursos.Text = DtaContratista.Recordset!CursosRecibidos
 Me.TxtRazones.Text = DtaContratista.Recordset!RazonesContrato
 Me.TxtTrabAnteriores.Text = DtaContratista.Recordset!TrabAnteriores
 Me.TxtRecomendaciones.Text = DtaContratista.Recordset!Recomendaciones
 Me.TxtIdiomas.Text = DtaContratista.Recordset!IdiomaDomina
 '/////Busco la Descripcion del Grupo/////////////////
  CodigoCuenta = Me.DBContratista.Text
  Criterio = "CodCuentas='" & CodigoCuenta & "'"
  DtaCuentas.Recordset.MoveFirst
  Me.DtaCuentas.Recordset.Find (Criterio)
  If Not DtaCuentas.Recordset.EOF Then
    CodGrupo = DtaCuentas.Recordset!CodGrupo
  End If
 
 
 '/////Busco la Descripcion del Grupo/////////////////
  Criterio = "CodGrupo='" & CodGrupo & "'"
  DtaGrupoCuentas.Recordset.MoveFirst
  Me.DtaGrupoCuentas.Recordset.Find (Criterio)
  If Not DtaGrupoCuentas.Recordset.EOF Then
    Me.DBGrupos.Text = DtaGrupoCuentas.Recordset!DescripcionGrupo
  End If
 

 
End If
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub DBGrupos_Click(Area As Integer)
On Error GoTo TipoErrs:
  If DBGrupos.Text = "" Then Exit Sub 'jonathan
  Criterio = "DescripcionGrupo='" & Me.DBGrupos.Text & "'"
  d = DtaGrupoCuentas.Recordset.RecordCount
  DtaGrupoCuentas.Recordset.MoveFirst
  Me.DtaGrupoCuentas.Recordset.Find (Criterio)
  CodGrupo = DtaGrupoCuentas.Recordset!CodGrupo
  
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Form_Activate()
On Error GoTo TipoErrs
If Not CodigoUsuario = 0 Then
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Contratistas'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False
  
 End If
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Contratistas'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdBorrar.Enabled = False
 End If
End If
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
Me.CmdJustifica.Enabled = False
'Me.CmdJustifica.BackColor = RGB(255, 255, 191)

With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Accesos"
   .Refresh
End With


With Me.DtaGrupoCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from GrupoCuentas"
   .Refresh
End With
LlenarDataCombos DtaGrupoCuentas, DBGrupos, "DescripcionGrupo", "CodGrupo"

With Me.DtaContratista
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from contactos"
   .Refresh
End With
LlenarDataCombos DtaContratista, DBContratista, "CodigoCuenta", "CodigoCuenta"

With Me.DtaCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Cuentas"
   .Refresh
End With


End Sub

