VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmEditaPresupuesto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Estructura"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   120
      Width           =   2655
      Begin VB.TextBox TxtNombre 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox TxtAnterior 
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmEditaPresupuesto.frx":0000
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmEditaPresupuesto.frx":007C
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Renombrar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc DtaGrupos 
      Height          =   375
      Left            =   240
      Top             =   3840
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "DtaGrupos"
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
End
Attribute VB_Name = "FrmEditaPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGrabar_Click()
FrmEstructuraPresupuesto.TreeView1.SelectedItem = Me.TxtNombre.Text

Me.DtaGrupos.RecordSource = "SELECT EstructuraPresupuesto.KeyGrupo, EstructuraPresupuesto.DescripcionGrupo From EstructuraPresupuesto Where (((EstructuraPresupuesto.KeyGrupo) = '" & KeyPrincipal & "'))"
Me.DtaGrupos.Refresh
If Not DtaGrupos.Recordset.EOF Then
 'Me.DtaGrupos.Recordset.Edit
  Me.DtaGrupos.Recordset("DescripcionGrupo") = Me.TxtNombre.Text
 Me.DtaGrupos.Recordset.Update
End If
Unload Me

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
With Me.DtaGrupos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With
Me.TxtAnterior.Text = DescripcionNodo

End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd

End Sub
