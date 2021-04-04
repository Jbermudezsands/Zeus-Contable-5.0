VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmConexiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conexiones"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6705
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "Sistema Factura"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sistema Nomina"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.TextBox TxtConexionString 
      Height          =   1515
      Left            =   1920
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1320
      Width           =   3855
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CmdConexion 
      Caption         =   "..."
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5C1A1&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   9015
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Image Image2 
         Height          =   960
         Left            =   240
         Picture         =   "FrmConexiones.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   9000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Top             =   120
         Width           =   645
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Administracion de Compañias"
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
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   4065
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmConexiones.frx":1B42
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "FrmConexiones.frx":25D36F
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "FrmConexiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAgregar_Click()

 If Me.Option2.Value = True Then
    MDIPrimero.AdoConfiguracion.Refresh
    MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion = Me.TxtConexionString.Text
    MDIPrimero.AdoConfiguracion.Recordset.Update
 Else
    MDIPrimero.AdoConfiguracion.Refresh
    MDIPrimero.AdoConfiguracion.Recordset!ConexionNomina = Me.TxtConexionString.Text
    MDIPrimero.AdoConfiguracion.Recordset.Update
 End If

MDIPrimero.AdoConfiguracion.Refresh
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdConexion_Click()
On Error GoTo TipoErrs
Dim mydlg As New MSDASC.DataLinks
Dim ADOcon As New ADODB.Connection

Me.TxtConexionString.Text = mydlg.PromptNew


Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub

Private Sub Form_Load()
Me.Skin1.ApplySkin hWnd



End Sub
