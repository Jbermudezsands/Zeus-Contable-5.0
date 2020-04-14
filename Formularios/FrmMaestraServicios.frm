VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmMaestraServicios 
   Caption         =   "Lista Maestra de Servicios"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12615
   Icon            =   "FrmMaestraServicios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Left            =   11160
      Picture         =   "FrmMaestraServicios.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   8520
      Width           =   1185
   End
   Begin VB.CommandButton Command4 
      Enabled         =   0   'False
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
      Left            =   8160
      Picture         =   "FrmMaestraServicios.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Agregar Mantenimiento al Activo"
      Top             =   8520
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
      Left            =   9480
      Picture         =   "FrmMaestraServicios.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar programacion "
      Top             =   8520
      Width           =   1185
   End
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
      Left            =   6840
      Picture         =   "FrmMaestraServicios.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Agregar nueva programacion"
      Top             =   8520
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      Begin VB.Frame Rechazo 
         BackColor       =   &H00B3BFAC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7935
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   12165
         Begin MSDataGridLib.DataGrid DataGrid7 
            Bindings        =   "FrmMaestraServicios.frx":1BB2
            Height          =   7815
            Left            =   0
            TabIndex        =   2
            Top             =   120
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   13785
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   19466
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   19466
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               Size            =   2
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin MSAdodcLib.Adodc Adoreg 
         Height          =   330
         Left            =   0
         Top             =   7920
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
   End
End
Attribute VB_Name = "FrmMaestraServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idservicio As Integer
Public ServicioalAF As Boolean
Public activocod As Integer
Dim mostrarAlePr As Integer
Dim DescripManto As String
Dim TipoRpePr As String
Dim TipoAlPr As String
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
FrmServicios.isactualiza = 0
FrmServicios.ismantoAF = 0
FrmServicios.Show vbModal
End Sub

Private Sub Command4_Click()
If idservicio = 0 Then
    MsgBox ("Seleccione el Servicio hacer añadido"), vbInformation
Else
    Set rsa = Nothing
    SQL = "select * from dbo.MantenimientoPorActivo"
    rsa.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
    rsa.AddNew
    rsa!idactivo = activocod
    rsa!IdServici = idservicio
    rsa!repetircada = DameDatosServicio
    rsa!mostraralerta = mostrarAlePr
    rsa!proximomanto = Format(Now, "DD/MM/YYYY")
    rsa!Descripcion = DescripManto
    rsa!tiporepeticion = TipoRpePr
    rsa!TipoAlerta = TipoAlPr
    rsa.Update
    Command4.Enabled = Falsefrafrainfind
    FrmMantenimientoActivos.cargamanttos (1)
End If
End Sub
Public Function DameDatosServicio() As Integer
Set rsa2 = Nothing
SQL = "select * from dbo.TipoServiciosManto where idreg=" & idservicio & ""
rsa2.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
DameDatosServicio = rsa2!repetircada
mostrarAlePr = rsa2!mostraralerta
DescripManto = rsa2!DescripcionMantto
TipoRpePr = rsa2!tiporepeticion
TipoAlPr = rsa2!TipoRepAlerta

End Function

Private Sub DataGrid7_Click()
datosservicio
End Sub
Private Sub datosservicio()
If Adoreg.Recordset.RecordCount = 0 Then
    idservicio = 0
Else
    idservicio = Adoreg.Recordset!no
End If
End Sub

Private Sub DataGrid7_DblClick()
datosservicio
FrmServicios.isactualiza = 1
FrmServicios.idservicio = idservicio
FrmServicios.Show vbModal
End Sub

Private Sub Form_Load()
If ServicioalAF = True Then
    Command4.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = False
Else
    Command4.Enabled = False
    Command2.Enabled = True
    Command3.Enabled = True
End If
cargaservicios

End Sub
Public Sub cargaservicios()
Adoreg.ConnectionString = Conexion
Adoreg.CommandTimeout = 0
Adoreg.RecordSource = "select idreg as No, DescripcionMantto as Servicio,repetircada as Frecuencia,   TipoRepeticion  as Servicio_Cada, mostraralerta as Frec_Alerta,  TipoRepAlerta  as Mostrar_Alerta_cada from dbo.TipoServiciosManto"
Adoreg.Refresh
End Sub
