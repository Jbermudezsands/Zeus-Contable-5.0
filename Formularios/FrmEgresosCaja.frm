VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmEgresosCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Egresos de Caja"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   12105
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   5160
      TabIndex        =   20
      Top             =   3960
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   77332481
      CurrentDate     =   41294
   End
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   375
      Left            =   960
      Top             =   8280
      Width           =   4335
      _ExtentX        =   7646
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
      Caption         =   "AdoCaja"
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   11895
      _Version        =   786432
      _ExtentX        =   20981
      _ExtentY        =   2990
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "FrmEgresosCaja.frx":0000
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmEgresosCaja.frx":006A
         TabIndex        =   18
         Top             =   1200
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmEgresosCaja.frx":00EA
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmEgresosCaja.frx":015E
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox CmbMoneda 
         Height          =   315
         ItemData        =   "FrmEgresosCaja.frx":01C8
         Left            =   3480
         List            =   "FrmEgresosCaja.frx":01D2
         TabIndex        =   14
         Text            =   "Córdobas"
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8520
         Picture         =   "FrmEgresosCaja.frx":01E9
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CmdBuscaCuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2880
         Picture         =   "FrmEgresosCaja.frx":0337
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox TxtNombreBanco 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   4335
      End
      Begin MSDataListLib.DataCombo DBCodigo 
         Bindings        =   "FrmEgresosCaja.frx":0485
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker TxtFecha 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   77332481
         CurrentDate     =   38918
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmEgresosCaja.frx":049B
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin MSDataListLib.DataCombo DBEmpleado 
         Bindings        =   "FrmEgresosCaja.frx":0503
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   5040
         OleObjectBlob   =   "FrmEgresosCaja.frx":051E
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblSaldo 
         Height          =   375
         Left            =   7080
         OleObjectBlob   =   "FrmEgresosCaja.frx":05A2
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   7080
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5C1A1&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   12135
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Registro de Egresos de Caja"
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
         Left            =   3360
         TabIndex        =   1
         Top             =   360
         Width           =   4320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   12120
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image2 
         Height          =   840
         Left            =   480
         Picture         =   "FrmEgresosCaja.frx":0600
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1080
      End
   End
   Begin MSComDlg.CommonDialog CDRuta 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrueOleDBGrid80.TDBGrid DBGTransacciones 
      Bindings        =   "FrmEgresosCaja.frx":C642
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5318
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).Caption=   "Detalles de  Movimientos"
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
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
      AllowAddNew     =   -1  'True
      Appearance      =   3
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      PictureCurrentRow(0)=   "bHQAAO4BAABCTe4BAAAAAAAANgAAACgAAAAOAAAACgAAAAEAGAAAAAAAuAEAAAAAAAAAAAAAAAAA"
      PictureCurrentRow(1)=   "AAAAAADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAMbHxgAAAP//"
      PictureCurrentRow(2)=   "/////////////////////////////////////////8bHxgAAxsfGAAAAhIaExsfGxsfGxsfGxsfG"
      PictureCurrentRow(3)=   "xsfGxsfGxsfGxsfGxsfG////xsfGAADGx8YAAACEhoTGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bG"
      PictureCurrentRow(4)=   "x8b////Gx8YAAMbHxgAAAISGhMbHxsbHxsbHxsbHxsbHxsbHxsbHxsbHxsbHxv///8bHxgAAxsfG"
      PictureCurrentRow(5)=   "AAAAhIaExsfGxsfGxsfGxsfGxsfGxsfGxsfGxsfGxsfG////xsfGAADGx8YAAACEhoTGx8bGx8bG"
      PictureCurrentRow(6)=   "x8bGx8bGx8bGx8bGx8bGx8bGx8b////Gx8YAAMbHxgAAAISGhISGhISGhISGhISGhISGhISGhISG"
      PictureCurrentRow(7)=   "hISGhISGhP///8bHxgAAxsfGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAxsfG"
      PictureCurrentRow(8)=   "AADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAA=="
      PictureCurrentRow.vt=   9
      PictureModifiedRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      PictureModifiedRow(0)=   "bHQAAO4BAABCTe4BAAAAAAAANgAAACgAAAAOAAAACgAAAAEAGAAAAAAAuAEAAAAAAAAAAAAAAAAA"
      PictureModifiedRow(1)=   "AAAAAADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAMbHxgAAAP//"
      PictureModifiedRow(2)=   "/////////////////////////////////////////8bHxgAAxsfGAAAAhIaEAP//AP//AP//AP//"
      PictureModifiedRow(3)=   "AP//AP//AP//AP//AP//////xsfGAADGx8YAAACEhoQA//8A//8A//8A//8A//8A//8A//8A//8A"
      PictureModifiedRow(4)=   "///////Gx8YAAMbHxgAAAISGhAD//wD//wD//wD//wD//wD//wD//wD//wD//////8bHxgAAxsfG"
      PictureModifiedRow(5)=   "AAAAhIaEAP//AP//AP//AP//AP//AP//AP//AP//AP//////xsfGAADGx8YAAACEhoQA//8A//8A"
      PictureModifiedRow(6)=   "//8A//8A//8A//8A//8A//8A///////Gx8YAAMbHxgAAAISGhISGhISGhISGhISGhISGhISGhISG"
      PictureModifiedRow(7)=   "hISGhISGhP///8bHxgAAxsfGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAxsfG"
      PictureModifiedRow(8)=   "AADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAA=="
      PictureModifiedRow.vt=   9
      PictureAddnewRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      PictureAddnewRow(0)=   "bHQAAO4BAABCTe4BAAAAAAAANgAAACgAAAAOAAAACgAAAAEAGAAAAAAAuAEAAAAAAAAAAAAAAAAA"
      PictureAddnewRow(1)=   "AAAAAADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAMbHxgAAAAAA"
      PictureAddnewRow(2)=   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMbHxgAAxsfG////hIaEhIaEhIaEhIaEhIaE"
      PictureAddnewRow(3)=   "hIaEhIaEhIaEhIaEhIaEAAAAxsfGAADGx8b///8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP+E"
      PictureAddnewRow(4)=   "hoQAAADGx8YAAMbHxv///wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/4SGhAAAAMbHxgAAxsfG"
      PictureAddnewRow(5)=   "////AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/hIaEAAAAxsfGAADGx8b///8AAP8AAP8AAP8A"
      PictureAddnewRow(6)=   "AP8AAP8AAP8AAP8AAP8AAP+EhoQAAADGx8YAAMbHxv///wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA"
      PictureAddnewRow(7)=   "/wAA/4SGhAAAAMbHxgAAxsfG////////////////////////////////////////////AAAAxsfG"
      PictureAddnewRow(8)=   "AADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAA=="
      PictureAddnewRow.vt=   9
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(8)   =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(9)   =   ":id=4,.fontname=MS Sans Serif"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HBFD6DD&,.fgcolor=&H800000&"
      _StyleDefs(24)  =   ":id=22,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(25)  =   ":id=22,.fontname=Lucida Calligraphy"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HBFD6DD&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(27)  =   ":id=14,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(28)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   960
      Top             =   8640
      Width           =   4335
      _ExtentX        =   7646
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
      Caption         =   "AdoConsulta"
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
   Begin MSAdodcLib.Adodc AdoSaldos 
      Height          =   375
      Left            =   1080
      Top             =   9000
      Width           =   4335
      _ExtentX        =   7646
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
      Caption         =   "AdoSaldos"
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
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   5400
      Top             =   9000
      Width           =   4335
      _ExtentX        =   7646
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
      Caption         =   "AdoEmpleados"
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
   Begin MSAdodcLib.Adodc AdoDetalleEgreso 
      Height          =   375
      Left            =   720
      Top             =   7920
      Width           =   4335
      _ExtentX        =   7646
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
      Caption         =   "AdoDetalleEgreso"
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
Attribute VB_Name = "FrmEgresosCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBCodigo_Change()
  Dim Debito As Double, Credito As Double, Total1 As Double, TipoCuenta As String

    Me.AdoConsulta.RecordSource = "SELECT  * From Cuentas WHERE  (CodCuentas = '" & Me.DBCodigo.Text & "')"
    Me.AdoConsulta.Refresh
    If Not Me.AdoConsulta.Recordset.EOF Then
      Me.TxtNombreBanco.Text = Me.AdoConsulta.Recordset("DescripcionCuentas")
      TipoCuenta = Me.AdoConsulta.Recordset("TipoCuenta")
      
          Me.AdoSaldos.RecordSource = "SELECT CodCuentas, MAX(DescripcionMovimiento) AS DescripcionMovimiento, MAX(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(ROUND(Debito * TCambio, 3)) AS MDebito, SUM(ROUND(TCambio * Credito, 3)) AS MCredito From Transacciones GROUP BY CodCuentas HAVING  (CodCuentas = '" & Me.DBCodigo.Text & "')"
          Me.AdoSaldos.Refresh
          
          Do While Not Me.AdoSaldos.Recordset.EOF
           If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
            If Not IsNull(Me.AdoSaldos.Recordset("MDebito")) Then
             Debito = Format(Me.AdoSaldos.Recordset("MDebito"), "##,##0.00")
            End If
            If Not IsNull(Me.AdoSaldos.Recordset("MCredito")) Then
             Credito = Format(Me.AdoSaldos.Recordset("MCredito"), "##,##0.00")
            End If
            Total1 = Debito - Credito + Total1
            Debito = 0
            Credito = 0
           Else
            If Not IsNull(Me.AdoSaldos.Recordset("MDebito")) Then
             Debito = Format(Me.AdoSaldos.Recordset("MDebito"), "##,##0.00")
            End If
            If Not IsNull(Me.AdoSaldos.Recordset("MCredito")) Then
             Credito = Format(Me.AdoSaldos.Recordset("MCredito"), "##,##0.00")
            End If
            Total1 = Credito - Debito + Total1
            Debito = 0
            Credito = 0
           End If
           
           Me.AdoSaldos.Recordset.MoveNext
          Loop
          
          Me.LblSaldo.Caption = Format(Total1, "##,##0.00")
    
    End If
    
    
        Me.AdoConsulta.RecordSource = "SELECT * From IndiceEgreso WHERE (CodigoCaja = '" & Me.DBCodigo.Text & "')"
        Me.AdoConsulta.Refresh
        If Not Me.AdoConsulta.Recordset.EOF Then
          Me.AdoConsulta.Recordset.MoveLast
          Me.DBEmpleado.Text = Me.AdoConsulta.Recordset("ResponsableFondo")
        End If
End Sub

Private Sub Form_Load()
 Me.TxtFecha.Value = Now
 MDIPrimero.Skin1.ApplySkin hWnd
 
    With Me.AdoCaja
       .ConnectionString = Conexion
    End With
    
    With Me.AdoEmpleados
       .ConnectionString = Conexion
    End With
    
    With Me.AdoConsulta
       .ConnectionString = Conexion
    End With
    
    With Me.AdoSaldos
       .ConnectionString = Conexion
    End With

    With Me.AdoDetalleEgreso
       .ConnectionString = Conexion
    End With
    
    Me.AdoCaja.RecordSource = "SELECT  * From Cuentas WHERE (TipoCuenta = 'Caja')"
    Me.AdoCaja.Refresh
    Me.DBCodigo.ListField = "CodCuentas"
    
    Me.AdoEmpleados.RecordSource = "SELECT  NombreEncargado, CodEncargado From Encargado"
    Me.AdoEmpleados.Refresh
    Me.DBEmpleado.ListField = "NombreEncargado"
    
    Me.AdoDetalleEgreso.RecordSource = "SELECT  CodCuentas, NombreCuenta, Concepto, ReciboNumero, FacturaNumero, FechaGasto, MontoGasto, NumeroEgreso From DetalleEgreso Where (NumeroEgreso = -1000000)"
    Me.AdoDetalleEgreso.Refresh
    Me.DBGTransacciones.Columns(0).Button = True
    Me.DBGTransacciones.Columns(1).Locked = True
    Me.DBGTransacciones.Columns(1).Locked = True
    Me.DBGTransacciones.Columns(0).Width = 1500
    Me.DBGTransacciones.Columns(3).Caption = "ReciboNo"
    Me.DBGTransacciones.Columns(3).Width = 1000
    Me.DBGTransacciones.Columns(4).Caption = "FacturaNo"
    Me.DBGTransacciones.Columns(4).Width = 1000
End Sub

