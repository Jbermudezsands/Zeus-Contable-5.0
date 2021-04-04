VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmGrupo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamentos de Cuentas"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   7995
   Begin VB.CommandButton Command3 
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
      Left            =   2640
      Picture         =   "FrmGrupo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmGrupo.frx":014E
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   3240
      OleObjectBlob   =   "FrmGrupo.frx":01C2
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin MSDataListLib.DataCombo DBCodigo 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaNacceso 
      Height          =   330
      Left            =   4200
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
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
   Begin MSAdodcLib.Adodc DtaGrupoCuentas 
      Height          =   375
      Left            =   4200
      Top             =   1800
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.TextBox TxtDescripcion 
      Height          =   315
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FrmGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
Me.DtaGrupoCuentas.Recordset.MovePrevious
If Me.DtaGrupoCuentas.Recordset.BOF Then
   Me.DtaGrupoCuentas.Recordset.MoveNext
   MsgBox "Este es el Primer Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
  Me.DBCodigo.Text = Me.DtaGrupoCuentas.Recordset!CodGrupo
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub CmdBorrar_Click()
On Error GoTo TipoErrs
  Dim Respuesta, Rsp
    If DtaGrupoCuentas.Recordset.RecordCount = 0 Then
        MsgBox "No Existen Registros de Departamentos de Cuentas Actualmente", vbInformation
        Exit Sub
    End If
  Set Rsp = Me.DtaGrupoCuentas.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando: " & Me.TxtDescripcion.Text)
   If Respuesta = 6 Then
     Criterio = "CodGrupo='" & Me.DBCodigo & "'"
     DtaGrupoCuentas.Recordset.MoveFirst
     Me.DtaGrupoCuentas.Recordset.Find (Criterio)
    If Not Me.DtaGrupoCuentas.Recordset.EOF Then
     Me.DtaGrupoCuentas.Recordset.Delete
   
    End If
      Me.DBCodigo.Text = ""
      TxtDescripcion.Text = ""
  End If
  Me.DtaGrupoCuentas.Refresh

 Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo TipoErrs
If DBCodigo.Text = "" Then
    MsgBox "El Código del Departamento de Cuentas es Requerido", vbInformation
    DBCodigo.SetFocus
    Exit Sub
End If
If TxtDescripcion.Text = "" Then
    MsgBox "La Descripción del Departamento de Cuentas es Requerido", vbInformation
    TxtDescripcion.SetFocus
    Exit Sub
End If

Criterio = "CodGrupo='" & Me.DBCodigo.Text & "'"
If DtaGrupoCuentas.Recordset.RecordCount <> 0 Then DtaGrupoCuentas.Recordset.MoveFirst
Me.DtaGrupoCuentas.Recordset.Find (Criterio)
If DtaGrupoCuentas.Recordset.EOF Then
  Me.DtaGrupoCuentas.Recordset.AddNew
  Me.DtaGrupoCuentas.Recordset!CodGrupo = Me.DBCodigo.Text
  Me.DtaGrupoCuentas.Recordset!DescripcionGrupo = Me.TxtDescripcion.Text
  Me.DtaGrupoCuentas.Recordset.Update

Else
  'Me.DtaGrupoCuentas.Recordset.Edit
  Me.DtaGrupoCuentas.Recordset!DescripcionGrupo = Me.TxtDescripcion.Text
  Me.DtaGrupoCuentas.Recordset.Update
 
End If
Me.DBCodigo.Text = ""
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdNuevo_Click()
Me.DBCodigo.Text = ""
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
Me.DtaGrupoCuentas.Recordset.MoveNext
If Me.DtaGrupoCuentas.Recordset.EOF Then
   Me.DtaGrupoCuentas.Recordset.MovePrevious
   MsgBox "Este es el Ultimo Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
  Me.DBCodigo.Text = Me.DtaGrupoCuentas.Recordset!CodGrupo
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub Command3_Click()
   QueProducto = "Departamento"
   FrmConsulta.Show 1
   
  Me.DBCodigo.Text = FrmConsulta.Codigo
End Sub

Private Sub DBCodigo_Change()
On Error GoTo TipoErrs
If DtaGrupoCuentas.Recordset.RecordCount = 0 Then
    Exit Sub
End If
Criterio = "CodGrupo='" & Me.DBCodigo.Text & "'"
DtaGrupoCuentas.Recordset.MoveFirst
Me.DtaGrupoCuentas.Recordset.Find (Criterio)
If DtaGrupoCuentas.Recordset.EOF Then
  Me.TxtDescripcion.Text = ""

Else
  Me.TxtDescripcion.Text = Me.DtaGrupoCuentas.Recordset!DescripcionGrupo
 
 
End If
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub Form_Activate()
On Error GoTo TipoErrs
MDIPrimero.Skin1.ApplySkin hWnd
If Not CodigoUsuario = 0 Then
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Grupo Cuentas'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False

 End If
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Grupo Cuentas'))"
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
With Me.DtaGrupoCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "select * from GrupoCuentas"
   .Refresh
End With
LlenarDataCombos DtaGrupoCuentas, DBCodigo, "CodGrupo", "CodGrupo"
With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Accesos"
   .Refresh
End With
Me.BackColor = RGB(236, 233, 216)
Me.CmdAnterior.BackColor = RGB(236, 233, 216)
Me.CmdBorrar.BackColor = RGB(236, 233, 216)
Me.CmdGrabar.BackColor = RGB(236, 233, 216)
Me.CmdNuevo.BackColor = RGB(236, 233, 216)
Me.CmdSalir.BackColor = RGB(236, 233, 216)
Me.CmdSiguiente.BackColor = RGB(236, 233, 216)


End Sub
