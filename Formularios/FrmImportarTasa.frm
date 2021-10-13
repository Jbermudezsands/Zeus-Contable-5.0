VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmImportarTasa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tasa de Cambios"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   6210
   Begin TrueOleDBGrid80.TDBGrid TDBGridTasas 
      Bindings        =   "FrmImportarTasa.frx":0000
      Height          =   7695
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   13573
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "USxCS"
      Columns(0).DataField=   "USxCS"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "FECHA"
      Columns(1).DataField=   "FECHA"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0"
      _StyleDefs(40)  =   "Named:id=34:Heading"
      _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=34,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=35:Footing"
      _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=36:Selected"
      _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=37:Caption"
      _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(49)  =   "Named:id=38:HighlightRow"
      _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=39:EvenRow"
      _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=40:OddRow"
      _StyleDefs(54)  =   ":id=40,.parent=33"
      _StyleDefs(55)  =   "Named:id=41:RecordSelector"
      _StyleDefs(56)  =   ":id=41,.parent=34"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc AdoTasas 
      Height          =   375
      Left            =   1440
      Top             =   8880
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
      Caption         =   "AdoTasas"
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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "tipo_cambio$"
      Top             =   8880
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "Abrir Archivo"
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmImportarTasa.frx":001B
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "Guardar "
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmImportarTasa.frx":1B6D
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   2760
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "Salir      "
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmImportarTasa.frx":36BF
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "xls"
   End
   Begin XtremeSuiteControls.ProgressBar Progress 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc AdoRegistros 
      Height          =   375
      Left            =   1200
      Top             =   8760
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
      Caption         =   "AdoRegistros"
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
   Begin MSAdodcLib.Adodc AdoImporta 
      Height          =   375
      Left            =   1080
      Top             =   8880
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
      Caption         =   "AdoImporta"
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
Attribute VB_Name = "FrmImportarTasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

 Me.TDBGridTasas.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGridTasas.OddRowStyle.BackColor = &H80000005
 Me.TDBGridTasas.AlternatingRowStyle = True

With Me.AdoTasas
  Me.AdoTasas.ConnectionString = Conexion
  Me.AdoTasas.RecordSource = "Tasas"
  Me.AdoTasas.Refresh
End With

With Me.AdoRegistros
   .ConnectionString = Conexion
   .RecordSource = "SELECT Identificador, CodigoCuenta AS USxCS, DescripcionCuenta AS FECHA FROM  RegistrosCuentas"
   .Refresh
End With

With Me.AdoImporta
  Me.AdoImporta.ConnectionString = Conexion
End With



        Me.AdoRegistros.Refresh
        Do While Not Me.AdoRegistros.Recordset.EOF
          Me.AdoRegistros.Recordset.Delete
         Me.AdoRegistros.Recordset.MoveNext
        Loop

End Sub

Private Sub PushButton1_Click()
Dim TasaCambio As Double, FechaTasa As String, cadena As String
On Error GoTo TipoErrs
Dim Directorio As String, i As Double
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Me.CommonDialog1.Filter = "Documentos de Texto|*.txt"

Me.CommonDialog1.ShowOpen
Directorio = Me.CommonDialog1.FileName

        Me.AdoRegistros.Refresh
        Do While Not Me.AdoRegistros.Recordset.EOF
          Me.AdoRegistros.Recordset.Delete
         Me.AdoRegistros.Recordset.MoveNext
        Loop
    
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '    ***************************ABRO EL ARCHIVO DE TASAS***********************************************************************
        '    **************************************************************************************************************************
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
        i = 0
        Open Directorio For Input As #1
        
        Me.AdoImporta.RecordSource = "SELECT Identificador, CodigoCuenta AS USxCS, CodigoDepartamento AS FECHA From RegistrosCuentas"
        Me.AdoImporta.Refresh
        
        
        
        
        While Not EOF(1)
        Salir = False
        Line Input #1, cadena
        
         If i <> 0 Then
       
            Identificador = Mid("A", 1, 1)
            FechaTasa = Mid(cadena, 1, 10)
            TasaCambio = Mid(cadena, 11, 18)
            
            
            
             Me.AdoRegistros.Recordset.AddNew
                AdoRegistros.Recordset("Identificador") = Identificador
                AdoRegistros.Recordset("USxCS") = Format(TasaCambio, "##,##0.0000")
                AdoRegistros.Recordset("FECHA") = Format(CDate(FechaTasa), "dd/mm/yyyy")
            Me.AdoRegistros.Recordset.Update
            
         End If
           i = i + 1
        Wend
        Close #1


'Me.Data1.DatabaseName = Archivo
'Me.Data1.Refresh

   

Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub

Private Sub PushButton2_Click()
Dim Registros As Integer, i As Integer, Contador As Integer
Dim Fecha As Date, TasaCambio As Double

On Error GoTo TipoErrs

Registros = Me.AdoRegistros.Recordset.RecordCount
Me.Progress.Visible = True

With Progress
   .Min = 0
   .Value = 0
   .Max = Registros
   i = 1
   Contador = 0
   
      Me.AdoRegistros.Refresh
      Do While Not Me.AdoRegistros.Recordset.EOF
      
         Fecha = Me.AdoRegistros.Recordset("FECHA")
         TasaCambio = Me.AdoRegistros.Recordset("USxCS")
       
         Me.AdoTasas.RecordSource = "SELECT * From Tasas WHERE (FechaTasas = CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102))"
         Me.AdoTasas.Refresh
         If Me.AdoTasas.Recordset.EOF Then
            Me.AdoTasas.Recordset.AddNew
             Me.AdoTasas.Recordset("FechaTasas") = Fecha
              Me.AdoTasas.Recordset("MontoCordobas") = TasaCambio
            Me.AdoTasas.Recordset.Update
         Else
            Me.AdoTasas.Recordset("MontoCordobas") = TasaCambio
            Me.AdoTasas.Recordset.Update
            
         End If
         DoEvents
   
         .Value = i
         i = i + 1
         Me.AdoRegistros.Recordset.MoveNext
      Loop
   
   
End With

MsgBox "Se Agregaron con Exito, las Tasas", vbInformation, "Zeus Facturacion"
Exit Sub
TipoErrs:
 MsgBox err.Description


End Sub

Private Sub PushButton3_Click()
Unload Me
End Sub
