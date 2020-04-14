VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmContratistaNProyectos 
   Caption         =   "Nuevos Proyectos"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CmbMoneda 
      Height          =   315
      ItemData        =   "FrmContratistaNProyectos.frx":0000
      Left            =   1920
      List            =   "FrmContratistaNProyectos.frx":000A
      TabIndex        =   17
      Text            =   "Cordobas"
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox TxtObservaciones 
      Height          =   525
      Left            =   1920
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox TxtPagoAnteriores 
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox TxtMontoContratado 
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox TxtDescripcionProyecto 
      Height          =   525
      Left            =   1920
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox TxtNombreProyecto 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   975
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _Version        =   786432
      _ExtentX        =   2778
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Guardar "
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmContratistaNProyectos.frx":0021
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   975
      Left            =   5040
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
      _Version        =   786432
      _ExtentX        =   2778
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Salir      "
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmContratistaNProyectos.frx":1B73
   End
   Begin MSComCtl2.DTPicker TxtFechaFinaliza 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   503
      _Version        =   393216
      Format          =   78774273
      CurrentDate     =   37992
   End
   Begin MSComCtl2.DTPicker TxtFechaContrata 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   503
      _Version        =   393216
      Format          =   78774273
      CurrentDate     =   37992
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmContratistaNProyectos.frx":4B90
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmContratistaNProyectos.frx":4C14
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "FrmContratistaNProyectos.frx":4C98
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc AdoProyectos 
      Height          =   375
      Left            =   720
      Top             =   5160
      Width           =   3615
      _ExtentX        =   6376
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
      Caption         =   "AdoProyectos"
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
   Begin VB.Label Label5 
      Caption         =   "Moneda"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Pagos Anteriores"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Monto Contratado"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre Proyecto"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FrmContratistaNProyectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoContratista As String, IdProyectos As Double

Private Sub Form_Load()
    
    
    With Me.AdoProyectos
       .ConnectionString = Conexion
       .RecordSource = "SELECT  ProyectosContratistas.* From ProyectosContratistas WHERE (CodigoContratista = '" & CodigoContratista & "')"
       .Refresh
    End With
    
    Me.TxtFechaContrata.Value = Now
    Me.TxtFechaFinaliza.Value = Now
End Sub

Private Sub PushButton2_Click()

Criterio = "IdProyecto=" & IdProyectos & " "
If AdoProyectos.Recordset.RecordCount = 0 Then

      Me.AdoProyectos.Recordset.AddNew
      AdoProyectos.Recordset("NombreProyecto") = Me.TxtNombreProyecto.Text
      AdoProyectos.Recordset("CodigoContratista") = CodigoContratista
      AdoProyectos.Recordset("FechaContrato") = Me.TxtFechaContrata.Value
      AdoProyectos.Recordset("FechaFinalizacion") = Me.TxtFechaFinaliza.Value
      AdoProyectos.Recordset("Descripcion_Trabajos") = Me.TxtDescripcionProyecto.Text
      AdoProyectos.Recordset("MontoContratado") = Me.TxtMontoContratado.Text
      AdoProyectos.Recordset("PagoAnterioresManual") = Me.TxtPagoAnteriores.Text
      AdoProyectos.Recordset("Observaciones") = Me.TxtObservaciones.Text
      AdoProyectos.Recordset("Moneda") = Me.CmbMoneda.Text
    Me.AdoProyectos.Recordset.Update

Else

    AdoProyectos.Recordset.MoveFirst
    Me.AdoProyectos.Recordset.Find (Criterio)
    If AdoProyectos.Recordset.EOF Then
      AdoProyectos.Recordset("CodEncargado") = Me.TxtNombreProyecto.Text
      AdoProyectos.Recordset("CodigoContratista") = CodigoContratista
      AdoProyectos.Recordset("FechaContrato") = Me.TxtFechaContrata.Value
      AdoProyectos.Recordset("FechaFinalizacion") = Me.TxtFechaFinaliza.Value
      AdoProyectos.Recordset("Descripcion_Trabajos") = Me.TxtDescripcionProyecto.Text
      AdoProyectos.Recordset("MontoContratado") = Me.TxtMontoContratado.Text
      AdoProyectos.Recordset("PagoAnterioresManual") = Me.TxtPagoAnteriores.Text
      AdoProyectos.Recordset("Observaciones") = Me.TxtObservaciones.Text
      AdoProyectos.Recordset("Moneda") = Me.CmbMoneda.Text
      Me.AdoProyectos.Recordset.Update
    End If
End If




End Sub

Private Sub PushButton3_Click()
Unload Me
End Sub
