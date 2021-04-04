VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmServicios 
   Caption         =   "Agregar/Editar Servicios"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6300
   Icon            =   "FrmServicios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5040
      Picture         =   "FrmServicios.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   2400
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3720
      Picture         =   "FrmServicios.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mostrar Alerta  (*)"
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   6015
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmServicios.frx":0E9E
         Left            =   840
         List            =   "FrmServicios.frx":0EAE
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   100
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Antes del Vencimiento del proximo mantenimiento"
         Height          =   195
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   3480
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmServicios.frx":0ECE
      Left            =   2400
      List            =   "FrmServicios.frx":0EDE
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Repetir cada (*):"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.TextBox txttipopla 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin MSComCtl2.DTPicker DTPicker8 
      Height          =   300
      Left            =   1560
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   0
      Format          =   109707265
      CurrentDate     =   38651
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proximo Servicio:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion (*):"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "FrmServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isactualiza As Integer
Public idservicio, ismantoAF As Integer


Private Sub Combo1_Click()
Combo1.BackColor = vb3DLight
End Sub

Private Sub Combo1_LostFocus()
Combo1.BackColor = vbWhite
End Sub

Private Sub Combo2_Click()
Combo2.BackColor = vb3DLight
End Sub

Private Sub Combo2_LostFocus()
Combo2.BackColor = vbWhite
End Sub

Private Sub Command2_Click()
guardainfo
FrmMaestraServicios.cargaservicios
End Sub
Private Sub guardainfo()
If txttipopla.Text = "" Or Text1.Text = "" Or Text2.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
MsgBox ("Falta informacion requerida, por favor verifique"), vbInformation
Else
    If ismantoAF = 1 Then
         Set rsa = Nothing
         sql = "select * from dbo.MantenimientoPorActivo where idreg=" & idservicio & ""
         rsa.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
        If Text1.Text = "" Then
            rsa!repetircada = 0
        Else
            rsa!repetircada = Text1.Text
        End If
        rsa!tiporepeticion = Combo1.Text
        If Text2.Text = "" Then
            rsa!mostraralerta = 0
        Else
            rsa!mostraralerta = Text2.Text
        End If
        rsa!TipoAlerta = Combo2.Text
        rsa!Descripcion = txttipopla.Text
        rsa!proximomanto = Format(DTPicker8.Value, "DD/MM/YYYY")
        rsa.Update
        Command2.Enabled = False
        FrmMantenimientoActivos.cargamanttos (1)
    Else
        Set rsa = Nothing
        If isactualiza = 0 Then
            sql = "select * from dbo.TipoServiciosManto"
        Else
            sql = "select * from dbo.TipoServiciosManto where idreg=" & idservicio & ""
        End If
        rsa.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
        If isactualiza = 0 Then
            rsa.AddNew
        End If
        If Text1.Text = "" Then
            rsa!repetircada = 0
        Else
            rsa!repetircada = Text1.Text
        End If
        rsa!tiporepeticion = Combo1.Text
        If Text2.Text = "" Then
            rsa!mostraralerta = 0
        Else
            rsa!mostraralerta = Text2.Text
        End If
        rsa!TipoRepAlerta = Combo2.Text
        rsa!DescripcionMantto = txttipopla.Text
        rsa.Update
        Command2.Enabled = False
    End If
End If
End Sub

Private Sub Command4_Click()
ismantoAF = 0
isactualiza = 0

Unload Me
End Sub

Private Sub Form_Load()
If ismantoAF = 1 Then
    DTPicker8.Visible = True
    Label2.Visible = True
Else
    DTPicker8.Visible = False
    Label2.Visible = False
End If

If idservicio <> 0 And ismantoAF = 0 Then
    cargadatosservicio
Else
    cargadatosservicioAF
End If

End Sub
Private Sub cargadatosservicio()
Set rsa = Nothing
sql = "select * from dbo.TipoServiciosManto where idreg=" & idservicio & ""
rsa.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
txttipopla.Text = rsa!DescripcionMantto
Text1.Text = rsa!repetircada
Combo1.Text = rsa!tiporepeticion
Text2.Text = rsa!mostraralerta
Combo2.Text = rsa!TipoRepAlerta

End Sub

Private Sub cargadatosservicioAF()
If idservicio = 0 Then

Else
    Set rsa = Nothing
    sql = "select * from dbo.MantenimientoPorActivo where idreg=" & idservicio & ""
    rsa.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
    txttipopla.Text = rsa!Descripcion
    Text1.Text = rsa!repetircada
    Combo1.Text = rsa!tiporepeticion
    Text2.Text = rsa!mostraralerta
    Combo2.Text = rsa!TipoAlerta
    If IsNull(rsa!proximomanto) Then
        DTPicker8.Value = Format(Now, "DD/MM/YYYY")
    Else
        DTPicker8.Value = Format(rsa!proximomanto, "DD/MM/YYYY")
    End If
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
ismantoAF = 0
isactualiza = 0
End Sub

Private Sub Text1_Click()
Text1.BackColor = vb3DLight
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = vbWhite
End Sub
Private Sub Text2_Click()
Text2.BackColor = vb3DLight
End Sub

Private Sub Text2_LostFocus()
Text2.BackColor = vb3DLight
End Sub

Private Sub txttipopla_Click()
txttipopla.BackColor = vb3DLight
End Sub

Private Sub txttipopla_LostFocus()
txttipopla.BackColor = vbWhite
End Sub
