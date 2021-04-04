VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmResponsablesAreas 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro y Control de Responsables de Areas o Departamentos"
   ClientHeight    =   6960
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10845
   Icon            =   "FrmResponsablesAreas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10845
   Begin VB.Frame Frame6 
      BackColor       =   &H80000003&
      ForeColor       =   &H00404000&
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.TextBox txtcedula 
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
         Left            =   7320
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtcargo 
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
         Left            =   7320
         MaxLength       =   200
         TabIndex        =   14
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtmail 
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
         Left            =   7320
         MaxLength       =   200
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
      Begin MSMask.MaskEdBox msktel 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H80000003&
         Caption         =   "Responsables de Areas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   10035
         Begin TrueOleDBGrid80.TDBGrid DataGrid2 
            Bindings        =   "FrmResponsablesAreas.frx":058A
            Height          =   2775
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   4895
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
            Splits(0).Caption=   "Responsables de Areas"
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
         Begin MSAdodcLib.Adodc Adodc3 
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
         Begin MSAdodcLib.Adodc Adodc4 
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
         Begin MSAdodcLib.Adodc adorespon 
            Height          =   330
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
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
      Begin VB.TextBox txtnombrecompleto 
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
         MaxLength       =   200
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtfecha 
         Height          =   345
         Left            =   1680
         TabIndex        =   2
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   17104897
         CurrentDate     =   38651
      End
      Begin MSDataListLib.DataCombo cmdgrupo2 
         Bindings        =   "FrmResponsablesAreas.frx":05A2
         DataField       =   "Idreg"
         Height          =   360
         Left            =   1680
         TabIndex        =   7
         Top             =   840
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
         Left            =   5280
         TabIndex        =   8
         ToolTipText     =   "Filtrar Oficinas "
         Top             =   840
         Width           =   375
         _Version        =   786432
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
         Left            =   4200
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cedula *"
         Height          =   195
         Left            =   5880
         TabIndex        =   18
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo Actual  *"
         Height          =   195
         Left            =   5880
         TabIndex        =   13
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correo electronico *"
         Height          =   195
         Left            =   5880
         TabIndex        =   11
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono *"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oficina de trabajo *"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Completo *"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Registro:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   1350
      End
   End
   Begin XtremeSuiteControls.PushButton btnguarda 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   6120
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Guardar"
      ForeColor       =   0
      BackColor       =   -2147483633
      Enabled         =   0   'False
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   6120
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Nuevo"
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   375
      Left            =   9240
      TabIndex        =   17
      Top             =   6120
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cerrar"
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Campos Obligatorios"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   6600
      Width           =   1545
   End
End
Attribute VB_Name = "FrmResponsablesAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim errorcedu As Integer
Dim idreg As Integer

Private Sub btnguarda_Click()
If txtnombrecompleto.Text = "" Or cmdgrupo2.Text = "" Or msktel.Text = "" Or txtmail.Text = "" Or txtcargo.Text = "" Or txtcedula.Text = "" Then
    MsgBox ("Informacion incompleta, por favor verifique"), vbInformation
    Exit Sub
Else
    guardarinfo
    limpiar
End If

End Sub
Private Sub guardarinfo()
If idreg <> 0 Then
    SQL = "select * from ResponsablesAreas where idreg=" & idreg & " "
Else
    SQL = "select * from ResponsablesAreas"
End If
adorespon.ConnectionString = Conexion
adorespon.RecordSource = SQL
adorespon.Refresh

If idreg = 0 Then
    adorespon.Recordset.AddNew
Else
End If
adorespon.Recordset!NombreResponsable = Trim(txtnombrecompleto.Text)
adorespon.Recordset!Area = Trim(cmdgrupo2.Text)
adorespon.Recordset!Telefono = msktel.Text
adorespon.Recordset!Email = txtmail.Text
adorespon.Recordset!Cargo = txtcargo.Text
adorespon.Recordset!fechareg = Format(Now, "YYYY/MM/DD")
adorespon.Recordset!IdAreaTrabajo = cmdgrupo2.BoundText
adorespon.Recordset!cedula = txtcedula.Text
adorespon.Recordset.Update
cargarresponsables
End Sub

Private Sub DataGrid2_Click()
ubicadatos
End Sub
Private Sub ubicadatos()
idreg = adorespon.Recordset!No
txtnombrecompleto.Text = adorespon.Recordset!Nombre_Completo
cmdgrupo2.BoundText = adorespon.Recordset!IdAreaTrabajo
msktel.Mask = "#####-###"
msktel.Mask = adorespon.Recordset!Telefono
txtmail.Text = adorespon.Recordset!Email
txtcargo.Text = adorespon.Recordset!Cargo
txtcedula.Text = adorespon.Recordset!cedula
End Sub

Private Sub DataGrid2_DblClick()
ubicadatos
btnguarda.Enabled = True
btnguarda.SetFocus
End Sub

Private Sub Form_Activate()
limpiar
cargaoficina
cargarresponsables
End Sub
Private Sub cargarresponsables()
adorespon.ConnectionString = Conexion
adorespon.CommandTimeout = 0
adorespon.RecordSource = "select IdReg as No, NombreResponsable as Nombre_Completo, Area as Area_de_Trabajo, Telefono, Email, Cargo, fechareg,IdAreaTrabajo,cedula from ResponsablesAreas"
adorespon.Refresh
End Sub

Private Sub Form_Load()
cargaoficina
msktel.Mask = "#####-###"

 Me.DataGrid2.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DataGrid2.OddRowStyle.BackColor = &H80000005
 Me.DataGrid2.AlternatingRowStyle = True
End Sub
Private Sub cargaoficina()
With Me.ofic
    .ConnectionString = Conexion
 End With
    CargaADODCConta "Oficinas", ofic, "1", cmdgrupo2.Name, "Trim", Conexion, Me, "order by Idreg"
End Sub

Private Sub msktel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtmail.SetFocus
End If
End Sub

Private Sub PushButton1_Click()
idreg = 0
limpiar
txtnombrecompleto.SetFocus
End Sub
Private Sub limpiar()
txtnombrecompleto.Text = ""
cmdgrupo2.Text = ""
'msktel.Text = " "
msktel.Mask = ""
msktel.Mask = "#####-###"
txtmail.Text = ""
txtcargo.Text = ""
txtcedula.Text = ""
End Sub

Private Sub PushButton2_Click()
Unload Me
End Sub

Private Sub PushButton3_Click()
FrmOficinas.Show vbModal
End Sub

Private Sub txtcargo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyAscii = 13 Then
    txtcedula.SetFocus
End If
End Sub

Private Sub txtcedula_Change()
If EsCedulaValida(txtcedula.Text) = False Then
    errorcedu = True
    btnguarda.Enabled = False
Else
    btnguarda.Enabled = True
    btnguarda.SetFocus
End If
End Sub

Private Sub txtmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtcargo.SetFocus
End If
End Sub

Private Sub txtnombrecompleto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdgrupo2.SetFocus
End If
End Sub

