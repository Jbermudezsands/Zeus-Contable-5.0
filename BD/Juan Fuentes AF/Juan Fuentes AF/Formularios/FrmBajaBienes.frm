VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.1#0"; "Codejock.Controls.v12.1.1.ocx"
Begin VB.Form FrmBajaBienes 
   BackColor       =   &H80000003&
   Caption         =   "Baja de Bienes"
   ClientHeight    =   9525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9390
   Icon            =   "FrmBajaBienes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Height          =   615
      Left            =   -120
      TabIndex        =   23
      Top             =   8760
      Width           =   9735
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
         TabIndex        =   24
         ToolTipText     =   "Filtrar por Codigo Activo, localizacion o Nombre del Bien"
         Top             =   150
         Width           =   4545
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   375
         Left            =   6720
         TabIndex        =   25
         Top             =   150
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
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   375
         Left            =   8160
         TabIndex        =   26
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
      Begin VB.Label Label9 
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
         TabIndex        =   27
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
      Left            =   0
      TabIndex        =   12
      Top             =   3600
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
         Bindings        =   "FrmBajaBienes.frx":058A
         Height          =   3495
         Left            =   120
         TabIndex        =   13
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
      Left            =   2160
      MaxLength       =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2985
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
      Left            =   6960
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1080
      Width           =   1305
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
      TabIndex        =   2
      Top             =   2400
      Width           =   4425
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      Height          =   615
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baja de Bienes"
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
         Width           =   2115
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton5 
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Top             =   1080
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
      Bindings        =   "FrmBajaBienes.frx":05A3
      DataField       =   "Idreg"
      Height          =   360
      Left            =   1680
      TabIndex        =   6
      Top             =   1680
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
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   1680
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
      Left            =   4080
      Top             =   1680
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
   Begin MSDataListLib.DataCombo dtrespo 
      Bindings        =   "FrmBajaBienes.frx":05B6
      DataField       =   "IdReg"
      Height          =   360
      Left            =   120
      TabIndex        =   14
      Top             =   7560
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
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
      Bindings        =   "FrmBajaBienes.frx":05CD
      DataField       =   "IdReg"
      Height          =   360
      Left            =   5160
      TabIndex        =   15
      Top             =   7560
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
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
   Begin XtremeSuiteControls.PushButton btnreci 
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   7560
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
      Left            =   8640
      TabIndex        =   17
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   7560
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
   Begin MSDataListLib.DataCombo dtrespo3 
      Bindings        =   "FrmBajaBienes.frx":05E4
      DataField       =   "IdReg"
      Height          =   360
      Left            =   2880
      TabIndex        =   20
      Top             =   8160
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
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   375
      Left            =   6360
      TabIndex        =   21
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   8160
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
      Left            =   0
      Top             =   7440
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Autorizado"
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
      Left            =   3840
      TabIndex        =   22
      Top             =   8520
      Width           =   915
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
      Left            =   1080
      TabIndex        =   19
      Top             =   7920
      Width           =   765
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
      Left            =   6120
      TabIndex        =   18
      Top             =   7920
      Width           =   885
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
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1920
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
      Left            =   6360
      TabIndex        =   10
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina Origen:"
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
      TabIndex        =   9
      Top             =   1680
      Width           =   1290
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
      TabIndex        =   8
      Top             =   2520
      Width           =   1080
   End
End
Attribute VB_Name = "FrmBajaBienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idactivobaja As String
Public isotromodulo As Integer

Private Sub btnentre_Click()
FrmResponsablesAreas.Show vbModal

End Sub

Private Sub btnreci_Click()
FrmResponsablesAreas.Show vbModal

End Sub

Private Sub cmdgrupo2_Click(Area As Integer)
If cmdgrupo2.Text = "" Then
Else
    ActivosDisponiblesAlta 3
End If
End Sub

Private Sub DataGrid7_Click()
idactivobaja = adoactivos.Recordset!idactivo
quienrecibeyentrego
End Sub
Private Sub quienrecibeyentrego()
Set rsa = Nothing
sql = "select IdUserRecibe, IdUserEntrega from AltadeBienes where IdActivoAlta='" & idactivobaja & "'"
rsa.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
dtrespo.BoundText = rsa!IdUserRecibe
dtrespo2.BoundText = rsa!IdUserEntrega
End Sub

Private Sub Form_Activate()
If isotromodulo = 0 Then
    cargaoficinas
    generareferencia
    cargaresponsables
    ActivosDisponiblesAlta 1
End If
End Sub
Private Sub ActivosDisponiblesAlta(opcionFiltro As Integer)
    adoactivos.ConnectionString = ConexionContable
    adoactivos.CommandTimeout = 0
    If opcionFiltro = 1 Then 'filtra todos los activos que aun no se le han dado de alta
                             'luego de ser registrados en el catalogo
        adoactivos.RecordSource = "select idreg as IdActivo, DescripcionAF as Descripcion, Serie as Marbete,fcompragen as Fecha_Adquisicion, Marca from dbo.CatalogoActivoFijo  where (datoalta=1 or datoalta='True' ) and (dadobaja=0 or dadobaja='False') "
    End If
    If opcionFiltro = 2 Then 'filtrado rapido, busca el activo por nombre o codigo escrito
        adoactivos.RecordSource = "select idreg as IdActivo, DescripcionAF as Descripcion, Serie as Marbete,fcompragen as Fecha_Adquisicion, Marca from dbo.CatalogoActivoFijo  where (datoalta=1 or datoalta='True' )and (dadobaja=0 or dadobaja='False') and  (descripcionactivo LIKE '" & Trim(txtfiltrorapido.Text) & "%' or codcuenta LIKE '" & Trim(txtfiltrorapido.Text) & "%' or localizacion LIKE '" & Trim(txtfiltrorapido.Text) & "%' ) "
    End If
     If opcionFiltro = 3 Then 'filtrado rapido, busca el activo por nombre o codigo escrito
        adoactivos.RecordSource = "select idreg as IdActivo, DescripcionAF as Descripcion, Serie as Marbete,fcompragen as Fecha_Adquisicion, Marca from dbo.CatalogoActivoFijo  where IdOfiAlta=" & cmdgrupo2.BoundText & " and  (datoalta=1 or datoalta='True' )and (dadobaja=0 or dadobaja='False')  "
    End If
    adoactivos.Refresh
End Sub
Private Sub cargaresponsables()
CargaADODCConta "ResponsablesAreas", adorespo, "1", dtrespo.Name, "Trim", ConexionContable, Me, "order by Idreg"
CargaADODCConta "ResponsablesAreas", adorespo, "1", dtrespo2.Name, "Trim", ConexionContable, Me, "order by Idreg"
CargaADODCConta "ResponsablesAreas", adorespo, "1", dtrespo3.Name, "Trim", ConexionContable, Me, "order by Idreg"
End Sub

Private Sub generareferencia()
Set rsa = Nothing
sql = "select max (idreg) as idreg from BajadeBienes "
rsa.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
If IsNull(rsa!idreg) Then
    Text1.Text = "000000" & 1
Else
    Text1.Text = "000000" & rsa!idreg + 1
End If
End Sub


Private Sub Form_Load()
cargaoficinas
generareferencia
cargaresponsables
ActivosDisponiblesAlta 1
End Sub
Private Sub cargaoficinas()
 With Me.ofic
    .ConnectionString = Conexion
 End With
    CargaADODCConta "Oficinas", ofic, "1", cmdgrupo2.Name, "Trim", ConexionContable, Me, "order by Idreg"
End Sub

Private Sub Form_Unload(Cancel As Integer)
isotromodulo = 0
End Sub

Private Sub PushButton1_Click()
FrmResponsablesAreas.Show vbModal

End Sub

Private Sub PushButton2_Click()
If txtfecha.Text = "" Or txtobserva.Text = "" Or dtrespo.Text = "" Or dtrespo2.Text = "" Or idactivobaja = "" Or dtrespo3.Text = "" Then
    MsgBox ("Informacion incompleta, por favor verifique"), vbInformation
Else
   guardarbajabien
   actualizaestadoactivo
   ActivosDisponiblesAlta 1
End If
End Sub
Private Sub actualizaestadoactivo()
Set rsa = Nothing
sql = "update CatalogoActivoFijo set dadobaja=1, FechaBaja='" & Format(Now, "YYYY/MM/DD") & "'  where Idreg='" & Trim(idactivobaja) & "'"
rsa.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
End Sub
Private Sub guardarbajabien()
Set rsa = Nothing
sql = "select * from BajadeBienes "
rsa.Open sql, ConexionContable, adOpenForwardOnly, adLockOptimistic
rsa.AddNew
rsa!IdReferencia = Text1.Text
rsa!FechaGraba = Format(CDate(txtfecha.Text), "YYYY/MM/DD")
rsa!IdOfiOrigen = cmdgrupo2.BoundText
rsa!DescriOficina = cmdgrupo2.Text
rsa!Observaciones = txtobserva.Text
rsa!IdUserRecibe = dtrespo.BoundText
rsa!NombreRecibe = dtrespo.Text
rsa!IdUserEntrega = dtrespo2.BoundText
rsa!NombreEntrega = dtrespo2.Text
rsa!IdUserAutoriza = dtrespo3.BoundText
rsa!NombreAutoriza = dtrespo3.Text
rsa!idactivobaja = idactivo
rsa.Update
PushButton2.Enabled = False
End Sub

Private Sub PushButton3_Click()
FrmOficinas.Show
End Sub

Private Sub PushButton4_Click()
isotromodulo = 0
Unload Me
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
