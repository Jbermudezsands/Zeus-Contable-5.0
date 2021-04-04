VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Entrada"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   HelpContextID   =   1
   Icon            =   "FrmEntrada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   4335
      TabIndex        =   11
      Top             =   0
      Width           =   4335
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Contro de Permisos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   2520
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   120
         Top             =   120
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   6720
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   120
         Picture         =   "FrmEntrada.frx":0442
         Top             =   120
         Width           =   720
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEntrada.frx":0D37
      TabIndex        =   10
      Top             =   600
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   1080
      OleObjectBlob   =   "FrmEntrada.frx":0DB9
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEntrada.frx":0E3D
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmEntrada.frx":0EBB
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc DtaUsuario 
      Height          =   375
      Left            =   480
      Top             =   3960
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
      Caption         =   "DtaUsuario"
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
   Begin VB.TextBox TxtCodEmpleado 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "TxtCodEmpleado"
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TxtNombreUsuario 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox TxtNivel 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   1
      Text            =   "TxtNivel"
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TxtClave 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmEntrada.frx":0F3F
      Top             =   0
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H8000000F&
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   735
      Left            =   3480
      OleObjectBlob   =   "FrmEntrada.frx":25C76C
      SourceDoc       =   "C:\Icon\Space.exe"
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "FrmEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
 DtaUsuario.Refresh
       Do While Not DtaUsuario.Recordset.EOF
         If DtaUsuario.Recordset("NombreUsuario") = TxtNombreUsuario.Text Then
            If DtaUsuario.Recordset("Clave") = TxtClave.Text Then
               NivelAcceso = DtaUsuario.Recordset("Nivel")
               CodPasword = DtaUsuario.Recordset("Clave")
               NombreUsuario = TxtNombreUsuario.Text
               CodigoUsuario = DtaUsuario.Recordset("CodUsuario")
               FechaSistema = Format(Now, "dd/mm/yyyy")
                    Unload Me
                     FrmListaUsuario.CmdSalir.Value = True
                     MDIPrimero.Show
                   
                    Exit Sub
                
                   Guia = 1
              Else
                Guia = 1
            End If 'Cierre del If Pasword
         Else
          Guia = 1
         End If 'Cierre del If NombreEmpleado
       DtaUsuario.Recordset.MoveNext
       Loop
    Select Case Guia
       Case 1: MsgBox "No Tiene Permiso", vbCritical, "Sistema de Contabilidad"
               TxtClave.Text = ""
               TxtClave.SetFocus
    End Select
End Sub

Private Sub CmdCancelar_Click()
 Unload Me
End Sub

Private Sub DBNombreUsuario_Change()
 'Al ejecutar algun cambio en el combo actualizo el nombre del Empleado
   DtaUsuario.Refresh
   Do While Not DtaUsuario.Recordset.EOF
     If DtaUsuario.Recordset("NombreEmpleado") = DBNombreUsuario.Text Then
        'TxtNivel.Text = DtaUsuario.Recordset("Nivel")Acceso
        'TxtClave.SetFocus
        Exit Do
     End If
       DtaUsuario.Recordset.MoveNext
   Loop
End Sub

Private Sub DBNombreUsuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   TxtNivel.SetFocus
 End If
End Sub

Private Sub Form_Load()
Me.Skin1.ApplySkin hWnd
With Me.DtaUsuario
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Usuarios"
   .Refresh
End With

End Sub


Private Sub TxtClave_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   CmdAceptar.Value = True
 End If
End Sub
Private Sub TxtNivel_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   TxtClave.SetFocus
  End If
End Sub

