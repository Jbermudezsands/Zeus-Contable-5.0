VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmListaUsuario 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios Registrados"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ForeColor       =   &H00EFEFEF&
   HelpContextID   =   1
   Icon            =   "FrmListaUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3255
   Begin MSDataListLib.DataList DBLUsuario 
      Bindings        =   "FrmListaUsuario.frx":0442
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3201
      _Version        =   393216
      ListField       =   "NombreUsuario"
   End
   Begin MSAdodcLib.Adodc DtaPassword 
      Height          =   375
      Left            =   480
      Top             =   4200
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
      Caption         =   "DtaPassword"
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
   Begin MSAdodcLib.Adodc DtaFecha 
      Height          =   375
      Left            =   600
      Top             =   3360
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
      Caption         =   "DtaFecha"
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
   Begin VB.CommandButton CmdSeleccionar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seleccionar"
      DownPicture     =   "FrmListaUsuario.frx":045C
      Height          =   375
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancelar"
      DownPicture     =   "FrmListaUsuario.frx":325E
      Height          =   375
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmListaUsuario.frx":4D40
      Top             =   0
   End
End
Attribute VB_Name = "FrmListaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdSalir_Click()
Unload Me
End Sub
Private Sub CmdSeleccionar_Click()
FrmEntrada.TxtNombreUsuario.Text = DBLUsuario.Text
FrmEntrada.Show 1
End Sub

Private Sub DBLUsuario_DblClick()
FrmEntrada.TxtNombreUsuario.Text = DBLUsuario.Text
FrmEntrada.Show 1
End Sub

Private Sub DBLUsuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  FrmEntrada.TxtNombreUsuario.Text = DBLUsuario.Text
  FrmEntrada.Show 1
 End If
End Sub



Private Sub Form_Activate()
Dim var As Double, BaseEntrada As Boolean

'FechaSistema = Format(Now, "dd/mm/yyyy")
'
' Me.DtaPassword.Refresh
' If DtaPassword.Recordset.EOF Then
'    BaseEntrada = True
'    MDIPrimero.Show
'    FrmListaUsuario.CmdSalir.Value = True
'    NivelAcceso = 100
'    NombreUsuario = "Desconocido"
'
'    GravaUsuarios = True
''    MDIPrimero.SmartMenuXP1.MenuItems.Enabled(3) = False
'  Else
'    Var = DtaPassword.Recordset("CodUsuario")
' End If

End Sub

Private Sub Form_Load()
Me.Skin1.ApplySkin hWnd
Dim TextFecha As String
Dim FechaSystem As Long
Dim Unidad As String
Dim RutaServer As String, Server As String
Dim Clave As String, User As String, NombreBD As String
Dim RutaFoto As String

Me.top = 3500
Me.Left = 3500

Dim ConexionSTR1 As String
Dim TxtClaveEntrada As String
'abro el archivo para solo lectura de la cadena de conexion
Dim NextLine As String
Dim Autorizado As Boolean
Autorizado = False

'FrmListaCompañia.CmdCancelar.Value = True


'Open App.Path + "\SysInfo.dll" For Input As #1
' Do Until EOF(1)
'  Line Input #1, NextLine
'        ConexionSTR1 = Trim(NextLine)
'  Loop
'Close #1
'
'  Unidad = App.Path + "\"
'  RutaFoto = App.Path + "\fotos\"
'  RutaLogo = App.Path + "\Imagenes\logo.jpg"
'  RutaIconos = App.Path + "\Imagenes"
  
  
'  Conexion = ConexionSTR1
'
'  ConexionReporte = ConexionSTR1

'  Unidad = Mid(Unidad, 1, 3)

With Me.DtaPassword
    .ConnectionString = Conexion
    .RecordSource = "Usuarios"
    .Refresh
End With

FechaSistema = Format(Now, "dd/mm/yyyy")

 Me.DtaPassword.Refresh
 If DtaPassword.Recordset.EOF Then
    BaseEntrada = True
    Me.Hide
'    MDIPrimero.Show
    NivelAcceso = 100
    NombreUsuario = "Desconocido"
    GravaUsuarios = True
'    Unload Me
'    MDIPrimero.SmartMenuXP1.MenuItems.Enabled(3) = False
  Else
    var = DtaPassword.Recordset("CodUsuario")
 End If

'Unload Me
End Sub
