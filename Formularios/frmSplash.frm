VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7335
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DtaConsulta 
      Caption         =   "DtaConsulta"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data DtaServidor 
      Caption         =   "DtaServidor"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Height          =   450
      Left            =   1800
      TabIndex        =   0
      Top             =   7560
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc DtaPassword 
      Height          =   375
      Left            =   3240
      Top             =   7800
      Visible         =   0   'False
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
   Begin VB.Image imgLogo 
      Height          =   7425
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9975
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Versión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   885
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      Caption         =   "Producto de la compañía"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   7
      Top             =   2400
      Width           =   3000
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      Caption         =   "Autorizado a"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2400
      TabIndex        =   5
      Top             =   900
      Width           =   2430
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plataforma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5460
      TabIndex        =   4
      Top             =   2100
      Width           =   1275
   End
   Begin VB.Label lblWarning 
      Caption         =   "Advertencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   3420
      Width           =   6855
   End
   Begin VB.Label lblCompany 
      Caption         =   "Compañía"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   3030
      Width           =   2415
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   2820
      Width           =   2415
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
'    FrmListaCompañia.Show
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
Dim NumeroCia As Double
Dim RutaServer As String, var As Variant

 If Timer1.Interval < 40 Then
   Me.Timer1.Interval = Me.Timer1.Interval + 1
 Else
 

DoEvents
 
       RutaServer = App.Path + "\CntConta.dll"
        If Dir(RutaServer) <> "" Then
        
          
          With Me.DtaServidor
             .DatabaseName = RutaServer
             .RecordSource = "Servidor"
             .Refresh
          End With
          
       
          
            With Me.DtaConsulta
             .DatabaseName = RutaServer
           End With
        End If
 
 
 
   If Not Me.DtaServidor.Recordset.EOF Then
   Me.DtaServidor.Recordset.MoveLast
   NumeroCia = Me.DtaServidor.Recordset.RecordCount
    If NumeroCia = 1 Then
      Conexion = Me.DtaServidor.Recordset("Servidor")
      ConexionReporte = Me.DtaServidor.Recordset("Servidor")
      
        With Me.DtaPassword
         .ConnectionString = Conexion
         .RecordSource = "Usuarios"
         .Refresh
        End With
        
         FechaSistema = Format(Now, "dd/mm/yyyy")
    
         Me.DtaPassword.Refresh
         If DtaPassword.Recordset.EOF Then
'            BaseEntrada = True
            Me.Hide
            MDIPrimero.Show
            Unload Me
            NivelAcceso = 100
            NombreUsuario = "Desconocido"
        
            GravaUsuarios = True
        '    MDIPrimero.SmartMenuXP1.MenuItems.Enabled(3) = False
          Else
            var = DtaPassword.Recordset("CodUsuario")
            Me.Hide
            FrmListaUsuario.Show
            Unload Me
         End If
    Else
            Me.Hide
            FrmListaCompañia.Show
            Unload Me
    End If
 End If
 
 
 

 End If
End Sub
