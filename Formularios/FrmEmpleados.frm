VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmEmpleados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Empleados"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   4215
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2640
      TabIndex        =   23
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   1560
      TabIndex        =   22
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   2640
      TabIndex        =   21
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   1560
      TabIndex        =   20
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   480
      TabIndex        =   19
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   4440
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEmpleados.frx":0000
      TabIndex        =   15
      Top             =   3600
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEmpleados.frx":0068
      TabIndex        =   14
      Top             =   3240
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEmpleados.frx":00EC
      TabIndex        =   13
      Top             =   240
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEmpleados.frx":0168
      TabIndex        =   12
      Top             =   720
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEmpleados.frx":01E4
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEmpleados.frx":0254
      TabIndex        =   10
      Top             =   1800
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEmpleados.frx":02C2
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo DBEmpleado 
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaNacceso 
      Height          =   375
      Left            =   5760
      Top             =   7320
      Width           =   3015
      _ExtentX        =   5318
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
      Height          =   375
      Left            =   3000
      Top             =   7320
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaCuentas "
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
   Begin MSAdodcLib.Adodc DtaEncargado 
      Height          =   375
      Left            =   0
      Top             =   6840
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "DtaEncargado"
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
   Begin VB.TextBox TxtCargo 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox TxtTelefono 
      Height          =   285
      Left            =   2040
      MaxLength       =   150
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox TxtCodigoPostal 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox TxtFax 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox TxtEmail 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Txtdireccion 
      Height          =   645
      Left            =   2040
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker TxtFechaContratacion 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Format          =   17104897
      CurrentDate     =   37992
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEmpleados.frx":033A
      TabIndex        =   16
      Top             =   2520
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEmpleados.frx":039E
      TabIndex        =   17
      Top             =   2880
      Width           =   735
   End
   Begin VB.Shape Shape1 
      DrawMode        =   6  'Mask Pen Not
      FillColor       =   &H00FFC0FF&
      Height          =   4095
      Left            =   -120
      Shape           =   5  'Rounded Square
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FrmEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
DtaEncargado.Recordset.MovePrevious
If DtaEncargado.Recordset.BOF Then
   DtaEncargado.Recordset.MoveNext
   MsgBox "Este es el Primer Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
  Me.DBEmpleado.Text = Me.DtaEncargado.Recordset!CodEncargado
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub CmdBorrar_Click()
On Error GoTo TipoErrs
  Dim Respuesta, Rsp
    If DtaEncargado.Recordset.RecordCount = 0 Then
        MsgBox "No Existen Registros de Empleados Actualmente", vbInformation
        Exit Sub
    End If
  Set Rsp = DtaEncargado.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando: " & Me.TxtNombre.Text)
   If Respuesta = 6 Then
     Criterio = "CodEncargado='" & Me.DBEmpleado.Text & "'"
     DtaEncargado.Recordset.MoveFirst
     Me.DtaEncargado.Recordset.Find (Criterio)
    If Not DtaEncargado.Recordset.EOF Then
     DtaEncargado.Recordset.Delete
    End If
   Me.DtaEncargado.Refresh
 End If
 limpiar
 Exit Sub
TipoErrs:
 ControlErrores
End Sub
Private Sub limpiar()
    Me.DBEmpleado.Text = ""
    Me.TxtNombre.Text = ""
    Me.TxtDireccion.Text = ""
    Me.TxtTelefono.Text = ""
    Me.TxtCodigoPostal.Text = ""
    Me.TxtFax.Text = ""
    Me.TxtEmail.Text = ""
    Me.txtcargo.Text = ""
    Me.TxtFechaContratacion = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo TipoErrs
If DBEmpleado.Text = "" Then
    MsgBox "El Código del Empleado es requerido", vbInformation
    DBEmpleado.SetFocus
    Exit Sub
End If
If TxtNombre.Text = "" Then
    MsgBox "El Nombre del Empleado es requerido", vbInformation
    TxtNombre.SetFocus
    Exit Sub
End If

Criterio = "CodEncargado='" & Me.DBEmpleado.Text & "'"
If DtaEncargado.Recordset.RecordCount <> 0 Then DtaEncargado.Recordset.MoveFirst
Me.DtaEncargado.Recordset.Find (Criterio)
If DtaEncargado.Recordset.EOF Then
  Me.DtaEncargado.Recordset.AddNew
  DtaEncargado.Recordset!CodEncargado = Me.DBEmpleado.Text
  DtaEncargado.Recordset!NombreEncargado = Me.TxtNombre.Text
  DtaEncargado.Recordset!Direccion = Me.TxtDireccion.Text
  DtaEncargado.Recordset!Telefono = Me.TxtTelefono.Text
  DtaEncargado.Recordset!CP = Me.TxtCodigoPostal.Text
  DtaEncargado.Recordset!Fax = Me.TxtFax.Text
  DtaEncargado.Recordset!Email = Me.TxtEmail.Text
  DtaEncargado.Recordset!Cargo = Me.txtcargo.Text
  DtaEncargado.Recordset!FechaContratacion = Me.TxtFechaContratacion
Me.DtaEncargado.Recordset.Update
Else
 'Me.DtaEncargado.Recordset.Edit
  DtaEncargado.Recordset!NombreEncargado = Me.TxtNombre.Text
  DtaEncargado.Recordset!Direccion = Me.TxtDireccion.Text
  DtaEncargado.Recordset!Telefono = Me.TxtTelefono.Text
  DtaEncargado.Recordset!CP = Me.TxtCodigoPostal.Text
  DtaEncargado.Recordset!Fax = Me.TxtFax.Text
  DtaEncargado.Recordset!Email = Me.TxtEmail.Text
  DtaEncargado.Recordset!Cargo = Me.txtcargo.Text
  DtaEncargado.Recordset!FechaContratacion = Me.TxtFechaContratacion
Me.DtaEncargado.Recordset.Update
  
 
 
End If
Me.DtaEncargado.Refresh
Me.DBEmpleado.Text = ""
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdNuevo_Click()
    limpiar
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
'On Error GoTo TipoErrs
Dim Respuesta As Integer
If DtaEncargado.Recordset.RecordCount = 0 Then
    MsgBox "No Existen Empleados Actualmente", vbInformation
    Exit Sub
End If
DtaEncargado.Recordset.MoveNext
If DtaEncargado.Recordset.EOF Then
   DtaEncargado.Recordset.MovePrevious
   MsgBox "Este es el Ultimo Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
  Me.DBEmpleado.Text = Me.DtaEncargado.Recordset!CodEncargado
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub DBEmpleado_Change()
On Error GoTo TipoErrs
If DtaEncargado.Recordset.RecordCount = 0 Then
    Exit Sub
End If
Criterio = "CodEncargado='" & Me.DBEmpleado.Text & "'"
DtaEncargado.Recordset.MoveFirst
Me.DtaEncargado.Recordset.Find (Criterio)
If DtaEncargado.Recordset.EOF Then
  Me.TxtNombre.Text = ""
  Me.TxtDireccion.Text = ""
  Me.TxtTelefono.Text = ""
  Me.TxtCodigoPostal.Text = ""
  Me.TxtFax.Text = ""
  Me.TxtEmail.Text = ""
  Me.txtcargo.Text = ""
  Me.TxtFechaContratacion = Format(Now, "dd/mm/yyyy")
Else
 Me.TxtNombre.Text = DtaEncargado.Recordset!NombreEncargado
  Me.TxtDireccion.Text = DtaEncargado.Recordset!Direccion
  Me.TxtTelefono.Text = DtaEncargado.Recordset!Telefono
  Me.TxtCodigoPostal.Text = DtaEncargado.Recordset!CP
  Me.TxtFax.Text = DtaEncargado.Recordset!Fax
  Me.TxtEmail.Text = DtaEncargado.Recordset!Email
  Me.txtcargo.Text = DtaEncargado.Recordset!Cargo
  Me.TxtFechaContratacion = Format(DtaEncargado.Recordset!FechaContratacion, "dd/mm/yyyy")
  
  
 
 
End If
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub Form_Activate()
On Error GoTo TipoErrs
Me.DtaCuentas.Refresh
Me.DtaEncargado.Refresh

If Not CodigoUsuario = 0 Then

 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Empleados'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False

 End If
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Empleados'))"
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

MDIPrimero.Skin1.ApplySkin Me.hWnd
'Me.BackColor = RGB(236, 233, 216)
'Me.CmdAnterior.BackColor = RGB(236, 233, 216)
'Me.CmdBorrar.BackColor = RGB(236, 233, 216)
'Me.CmdGrabar.BackColor = RGB(236, 233, 216)
'Me.CmdNuevo.BackColor = RGB(236, 233, 216)
'Me.CmdSalir.BackColor = RGB(236, 233, 216)
'Me.CmdSiguiente.BackColor = RGB(236, 233, 216)


With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Accesos"
   .Refresh
End With

With Me.DtaEncargado
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Encargado"
   .Refresh
End With
LlenarDataCombos DtaEncargado, DBEmpleado, "CodEncargado", "CodEncargado"

With Me.DtaCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "select * from Cuentas"
   .Refresh
End With
End Sub

Private Sub Label6_Click()

End Sub
