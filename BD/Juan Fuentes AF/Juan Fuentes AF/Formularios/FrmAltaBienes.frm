VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.1#0"; "Codejock.Controls.v12.1.1.ocx"
Begin VB.Form FrmAltaBienes 
   BackColor       =   &H80000003&
   Caption         =   "Alta de Bienes"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9600
   Icon            =   "FrmAltaBienes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Caption         =   "Consulta Rápida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6540
      TabIndex        =   25
      Top             =   2520
      Width           =   2895
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000003&
         Caption         =   "Ver Traslados"
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000003&
         Caption         =   "Ver Bajas"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Height          =   615
      Left            =   -120
      TabIndex        =   17
      Top             =   7920
      Width           =   9855
      Begin VB.TextBox txtfiltrorapido 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   21
         ToolTipText     =   "Filtrar por Codigo Activo, localizacion o Nombre del Bien"
         Top             =   150
         Width           =   4545
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   6720
         TabIndex        =   18
         Top             =   120
         Width           =   1335
         _Version        =   786433
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Guardar"
         ForeColor       =   0
         BackColor       =   -2147483633
         Appearance      =   6
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   375
         Left            =   8160
         TabIndex        =   19
         Top             =   150
         Width           =   1335
         _Version        =   786433
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Salir"
         ForeColor       =   0
         BackColor       =   -2147483633
         Appearance      =   6
         ImageAlignment  =   0
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro Rápido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Frame Rechazo 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   9285
      Begin MSAdodcLib.Adodc AdoHist 
         Height          =   330
         Left            =   60
         Top             =   5640
         Visible         =   0   'False
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   582
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
         Appearance      =   0
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
         Caption         =   "Registro 0 de 0"
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
      Begin MSDataGridLib.DataGrid DataGrid7 
         Bindings        =   "FrmAltaBienes.frx":058A
         Height          =   3495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6165
         _Version        =   393216
         HeadLines       =   1
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adoreg 
         Height          =   330
         Left            =   0
         Top             =   5880
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
      Begin MSAdodcLib.Adodc adoactivos 
         Height          =   330
         Left            =   0
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
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
         Appearance      =   0
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
         Caption         =   "Registro 0 de 0"
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
   Begin VB.TextBox txtobserva 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   1200
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2040
      Width           =   4425
   End
   Begin VB.TextBox txtfecha 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7080
      MaxLength       =   20
      TabIndex        =   5
      Top             =   720
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2280
      MaxLength       =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2985
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      Height          =   615
      Left            =   -120
      TabIndex        =   0
      Top             =   -120
      Width           =   9855
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alta de Bienes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6960
         TabIndex        =   1
         Top             =   195
         Width           =   2040
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton5 
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   720
      Width           =   615
      _Version        =   786433
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "?"
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin MSDataListLib.DataCombo cmdgrupo2 
      Bindings        =   "FrmAltaBienes.frx":05A3
      DataField       =   "Idreg"
      Height          =   360
      Left            =   5640
      TabIndex        =   8
      Top             =   1440
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "Descripcion"
      BoundColumn     =   "Idreg"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dtrespo 
      Bindings        =   "FrmAltaBienes.frx":05B6
      DataField       =   "IdReg"
      Height          =   360
      Left            =   480
      TabIndex        =   14
      Top             =   7320
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "NombreResponsable"
      BoundColumn     =   "IdReg"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dtrespo2 
      Bindings        =   "FrmAltaBienes.frx":05CD
      DataField       =   "IdReg"
      Height          =   360
      Left            =   5520
      TabIndex        =   16
      Top             =   7320
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "NombreResponsable"
      BoundColumn     =   "IdReg"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   375
      Left            =   9000
      TabIndex        =   22
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   1440
      Width           =   375
      _Version        =   786433
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "..."
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin MSAdodcLib.Adodc ofic 
      Height          =   330
      Left            =   7920
      Top             =   1440
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
   Begin XtremeSuiteControls.PushButton btnreci 
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   7320
      Width           =   375
      _Version        =   786433
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "..."
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnentre 
      Height          =   375
      Left            =   9000
      TabIndex        =   24
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   7320
      Width           =   375
      _Version        =   786433
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "..."
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin MSAdodcLib.Adodc adorespo 
      Height          =   330
      Left            =   120
      Top             =   7080
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entregado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6480
      TabIndex        =   15
      Top             =   7680
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recibido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   13
      Top             =   7680
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina Destino:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   7
      Top             =   1440
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6480
      TabIndex        =   4
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Referencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1920
   End
End
Attribute VB_Name = "FrmAltaBienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idactivo As String
Private Sub btnentre_Click()
FrmResponsablesAreas.Show vbModal
End Sub

Private Sub btnreci_Click()
FrmResponsablesAreas.Show vbModal
End Sub

Private Sub DataGrid7_Click()
tieneid
End Sub
Public Function tieneid() As Boolean
If adoactivos.Recordset.RecordCount = 0 Then
    idactivo = 0
    tieneid = False
Else
    idactivo = adoactivos.Recordset!idactivo
    tieneid = True
End If
End Function

Private Sub DataGrid7_DblClick()
If Option2.Value = True Then
'es una baja
    If tieneid = False Then
        Exit Sub
    Else
        FrmBajaBienes.isotromodulo = 1
        Set rsa2 = Nothing
        sql = "select * from BajadeBienes where IdActivoBaja='" & idactivo & "' "
        rsa2.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
        FrmBajaBienes.Text1.Text = rsa2!IdReferencia
        FrmBajaBienes.Text1.Locked = True
        FrmBajaBienes.txtfecha.Text = rsa2!FechaGraba
        FrmBajaBienes.txtfecha.Locked = True
        FrmBajaBienes.PushButton5.Enabled = False
        FrmBajaBienes.cmdgrupo2.Text = rsa2!DescriOficina
        FrmBajaBienes.cmdgrupo2.Locked = True
        PushButton3.Enabled = False
        FrmBajaBienes.txtobserva.Text = rsa2!Observaciones
        FrmBajaBienes.txtobserva.Locked = True
        FrmBajaBienes.dtrespo.Text = rsa2!NombreRecibe
        FrmBajaBienes.dtrespo.Locked = True
        FrmBajaBienes.dtrespo2.Text = rsa2!NombreEntrega
        FrmBajaBienes.dtrespo2.Locked = True
        FrmBajaBienes.dtrespo3.Text = rsa2!NombreAutoriza
        FrmBajaBienes.dtrespo3.Locked = True
        FrmBajaBienes.btnreci.Enabled = False
        FrmBajaBienes.btnentre.Enabled = False
        FrmBajaBienes.PushButton1.Enabled = False
        FrmBajaBienes.txtfiltrorapido.Locked = True
        FrmBajaBienes.PushButton2.Enabled = False
        FrmBajaBienes.PushButton3.Enabled = False
        FrmBajaBienes.Show
    End If
End If
End Sub

Private Sub Form_Activate()
cargaoficinas
generareferencia
cargaresponsables
End Sub
Private Sub cargaresponsables()
CargaADODCConta "ResponsablesAreas", adorespo, "1", dtrespo.Name, "Trim", ConexionContable, Me, "order by Idreg"
CargaADODCConta "ResponsablesAreas", adorespo, "1", dtrespo2.Name, "Trim", ConexionContable, Me, "order by Idreg"
End Sub
Private Sub generareferencia()
Set rsa = Nothing
sql = "select max (idreg) as idreg from AltadeBienes "
rsa.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
If IsNull(rsa!idreg) Then
    Text1.Text = "000000" & 1
Else
    Text1.Text = "000000" & rsa!idreg + 1
End If
End Sub

Private Sub Form_Load()
cargaoficinas
ActivosDisponiblesAlta 1 'filtra todos los activos disponibles para ser dados de alta.
                        'Se registra en el catalogo, luego deben darse de alta
End Sub
Private Sub cargaoficinas()
 With Me.ofic
    .ConnectionString = Conexion
 End With
    CargaADODCConta "Oficinas", ofic, "1", cmdgrupo2.Name, "Trim", ConexionContable, Me, "order by Idreg"
End Sub
Private Sub ActivosDisponiblesAlta(opcionFiltro As Integer)
    adoactivos.ConnectionString = ConexionContable
    adoactivos.CommandTimeout = 0
    If opcionFiltro = 1 Then 'filtra todos los activos que aun no se le han dado de alta
                             'luego de ser registrados en el catalogo
        adoactivos.RecordSource = "select idreg as IdActivo, DescripcionAF as Descripcion, Serie as Marbete,fcompragen as Fecha_Adquisicion, Marca from dbo.CatalogoActivoFijo where (datoalta is null or datoalta='False' or datoalta=0) and (dadobaja is null or dadobaja='False' or dadobaja=0) "
    End If
    If opcionFiltro = 2 Then 'filtrado rapido, busca el activo por nombre o codigo escrito
        adoactivos.RecordSource = "select idreg as IdActivo, DescripcionAF as Descripcion, Serie as Marbete,fcompragen as Fecha_Adquisicion, Marca from dbo.CatalogoActivoFijo where (datoalta is null or datoalta='False' or datoalta=0) and (dadobaja is null or dadobaja='False' or dadobaja=0) and  (DescripcionAF LIKE '" & Trim(txtfiltrorapido.Text) & "%' or cntacontable LIKE '" & Trim(txtfiltrorapido.Text) & "%' or descrigrupo LIKE '" & Trim(txtfiltrorapido.Text) & "%' ) "
    End If
    If opcionFiltro = 3 Then 'Ver altas
        adoactivos.RecordSource = "select idreg as IdActivo, DescripcionAF as Descripcion, Serie as Marbete,fcompragen as Fecha_Adquisicion, Marca from dbo.CatalogoActivoFijo where (datoalta is not null or datoalta='True' or datoalta=1) and (dadobaja is null or dadobaja='False' or dadobaja=0) and (Trasladado is null or Trasladado='False' or Trasladado=0)"
    End If
    
    If opcionFiltro = 4 Then 'Ver altas
        adoactivos.RecordSource = "select idreg as IdActivo, DescripcionAF as Descripcion, Serie as Marbete,fcompragen as Fecha_Adquisicion, Marca from dbo.CatalogoActivoFijo where (dadobaja is not null or dadobaja='True' or dadobaja=1) and ((Trasladado is null or Trasladado='False' or Trasladado=0)or (Trasladado is not null or Trasladado='True' or Trasladado=1))"
    End If
    
    If opcionFiltro = 5 Then 'Ver traslados
        adoactivos.RecordSource = "select idreg as IdActivo, DescripcionAF as Descripcion, Serie as Marbete,fcompragen as Fecha_Adquisicion, Marca from dbo.CatalogoActivoFijo where (Trasladado is not null or Trasladado='True' or Trasladado=1)"
    End If
    
    adoactivos.Refresh
End Sub


Private Sub Option1_Click()
ActivosDisponiblesAlta 3 'Ver altas
End Sub

Private Sub Option2_Click()
ActivosDisponiblesAlta 4 'Ver bajas
End Sub

Private Sub Option3_Click()
ActivosDisponiblesAlta 4 'Ver traslados
End Sub

Private Sub PushButton1_Click()
If txtfecha.Text = "" Or txtobserva.Text = "" Or dtrespo.Text = "" Or dtrespo2.Text = "" Or idactivo = "" Then
    MsgBox ("Informacion incompleta, por favor verifique"), vbInformation
Else
   guardaaltabien
   actualizaestadoactivo
   ActivosDisponiblesAlta 1
End If
End Sub
Private Sub actualizaestadoactivo()
Set rsa = Nothing
sql = "update dbo.CatalogoActivoFijo set datoalta=1, fechaalta='" & Format(Now, "YYYY/MM/DD") & "', set IdActivoAlta=" & idactivo & "  where idreg='" & Trim(idactivo) & "'"
rsa.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
End Sub

Private Sub guardaaltabien()
Set rsa = Nothing
sql = "select * from AltadeBienes "
rsa.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
rsa.AddNew
rsa!IdReferencia = Text1.Text
rsa!FechaGraba = Format(CDate(txtfecha.Text), "YYYY/MM/DD")
rsa!IdOfiDestino = cmdgrupo2.BoundText
rsa!DescriOficina = cmdgrupo2.Text
rsa!Observaciones = txtobserva.Text
rsa!IdUserRecibe = dtrespo.BoundText
rsa!NombreRecibe = dtrespo.Text
rsa!IdUserEntrega = dtrespo2.BoundText
rsa!NombreEntrega = dtrespo2.Text
rsa!IdActivoAlta = idactivo
rsa!IdOfiAlta = cmdgrupo2.BoundText
rsa.Update
PushButton1.Enabled = False
End Sub

Private Sub PushButton2_Click()
Unload Me
End Sub

Private Sub PushButton3_Click()
FrmOficinas.Show
End Sub

Private Sub PushButton5_Click()
 
    wfecha = IIf(Len(Trim(txtfecha.Text)) = 0 Or Not IsDate(txtfecha.Text), Date, txtfecha.Text)
    Set wforma = Me
    wtextc = txtfecha.Name
    whabfe = True
    On Local Error Resume Next
    Load fcalendario
    On Local Error GoTo 0
    If fcalendario.Visible = False Then fcalendario.Show vbModal
End Sub

Private Sub Text3_LostFocus()
Text3.Text = UCase(Left(Text3.Text, 1)) & LCase(Mid(Text3.Text, 2))

End Sub

Private Sub txtfiltrorapido_Change()
If Not IsNumeric(txtfiltrorapido.Text) Then 'Significa que esta escribiendo el nombre del activo
                                        'o el numero de codigo del mismo
    ActivosDisponiblesAlta 2 'busca el filtro del activo
Else
    If txtfiltrorapido.Text = "" Then
        ActivosDisponiblesAlta 1
    End If
End If
End Sub
