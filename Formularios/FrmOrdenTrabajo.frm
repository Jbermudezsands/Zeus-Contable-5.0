VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmOrdenTrabajo 
   Caption         =   "Agregar / Editar Orden de Trabajo"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6375
   Icon            =   "FrmOrdenTrabajo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Left            =   3600
      Picture         =   "FrmOrdenTrabajo.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Guardar Informacion"
      Top             =   6240
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
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
      Left            =   4920
      Picture         =   "FrmOrdenTrabajo.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Salir"
      Top             =   6240
      Width           =   1185
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   2040
      MaxLength       =   300
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   4920
      Width           =   4095
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
      Left            =   2040
      MaxLength       =   100
      TabIndex        =   16
      Top             =   4560
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado"
      Height          =   1095
      Left            =   2040
      TabIndex        =   12
      Top             =   3360
      Width           =   4095
      Begin VB.OptionButton Option3 
         Caption         =   "Cerrado"
         Height          =   375
         Left            =   2760
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "En curso"
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pendiente"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   2040
      MaxLength       =   300
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1920
      Width           =   4095
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
      Left            =   2040
      MaxLength       =   100
      TabIndex        =   8
      Top             =   1560
      Width           =   4095
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
      Left            =   2040
      MaxLength       =   100
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker DTPicker8 
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   2340
      _ExtentX        =   4128
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
      Format          =   75104257
      CurrentDate     =   38651
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "FrmOrdenTrabajo.frx":109E
      DataField       =   "codempleado"
      Height          =   360
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483640
      ListField       =   "nombrecompleto"
      BoundColumn     =   "codempleado"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc conduc 
      Height          =   330
      Left            =   3720
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   2040
      TabIndex        =   6
      Top             =   1200
      Width           =   2340
      _ExtentX        =   4128
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
      Format          =   75104257
      CurrentDate     =   38651
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota:"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia #"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   930
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor Responsable"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Requerido en :"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1050
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reportado por:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Creado el:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activo Fijo:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "FrmOrdenTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public idAF As Integer
Public idordentra As Integer
Public isactualiza As Integer


Private Sub Command2_Click()
Set rsa3 = Nothing
If isactualiza = 0 Then
    SQL = "SELECT * FROM dbo.ControlOrdenTrabajo"
End If

If isactualiza = 1 Then
    SQL = "SELECT * FROM dbo.ControlOrdenTrabajo WHERE idreg=" & idAF & " "
End If

rsa3.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
If isactualiza = 0 Then
    rsa3.AddNew
End If
rsa3!BienOrden = txttipopla.Text
rsa3!Fcreado = Format(DTPicker8.Value, "DD/MM/YYYY")
rsa3!Reportadopor = DataCombo3.Text
rsa3!frequeireOrden = Format(DTPicker1.Value, "DD/MM/YYYY")
rsa3!proveeresponsable = Text1.Text
If Text8.Text = "" Then
    rsa3!Descripcion = ""
Else
    rsa3!Descripcion = Text8.Text
End If
If Option1.Value = True Then
    rsa3!estado = "P"
Else
    If Option2.Value = True Then
        rsa3!estado = "EC"
    Else
        If Option3.Value = True Then
            rsa3!estado = "C"
        End If
    End If
End If
 If Text2.Text = "" Then
    rsa3!referencia = ""
 Else
    rsa3!referencia = Text2.Text
End If

If Text3.Text = "" Then
    rsa3!Nota = ""
Else
    rsa3!Nota = Text3.Text
End If
 rsa3!idactivo = idAF
rsa3.Update
Command2.Enabled = False
FrmMantenimientoActivos.cargarordenestrabajo
End Sub

Private Sub Command3_Click()
isactualiza = 0
Unload Me
End Sub

Private Sub Form_Load()
DTPicker8.Value = Format(Now, "DD/MM/YYYY")
DTPicker1.Value = Format(Now, "DD/MM/YYYY")
CargaADODC "Empleado", conduc, "", DataCombo3.Name, "Trim", Conexion, Me, "order by CodEmpleado"
If isactualiza = 1 Then
        cargadatosOT
End If
End Sub
Private Sub cargadatosOT()
Set rsa3 = Nothing
SQL = "SELECT * FROM dbo.ControlOrdenTrabajo where idreg=" & idordentra & ""
rsa3.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic

 txttipopla.Text = rsa3!BienOrden
DTPicker8.Value = Format(rsa3!Fcreado, "DD/MM/YYYY")
DataCombo3.Text = rsa3!Reportadopor
DTPicker1.Value = Format(rsa3!frequeireOrden, "DD/MM/YYYY")
Text1.Text = rsa3!proveeresponsable

If IsNull(rsa3!Descripcion) Then
    Text8.Text = ""
Else
    Text8.Text = rsa3!Descripcion
End If

If rsa3!estado = "P" Then
    Option1.Value = True
Else
    If rsa3!estado = "EC" Then
       Option2.Value = True
    Else
        If rsa3!estado = "C" Then
           Option3.Value = True
        End If
    End If
End If
 If IsNull(rsa3!referencia) Then
   Text2.Text = ""
 Else
     Text2.Text = rsa3!referencia
End If

If IsNull(rsa3!Nota) Then
    Text3.Text = ""
Else
    Text3.Text = rsa3!Nota
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
isactualiza = 0
End Sub

