VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmAgregarActivoFijo 
   Caption         =   "Agregando Activo Fijo"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17565
   Icon            =   "FrmAgregarActivoFijo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   17565
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11520
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   12120
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   12720
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   17415
      _Version        =   786432
      _ExtentX        =   30718
      _ExtentY        =   13996
      _StockProps     =   68
      Appearance      =   9
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   2
      Item(0).Caption =   "Datos Generales"
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "Label49"
      Item(0).Control(1)=   "Label64"
      Item(0).Control(2)=   "Frame6"
      Item(0).Control(3)=   "Frame7"
      Item(0).Control(4)=   "Frame4"
      Item(0).Control(5)=   "Frame9"
      Item(0).Control(6)=   "Frame5"
      Item(0).Control(7)=   "Check4"
      Item(0).Control(8)=   "Check3"
      Item(0).Control(9)=   "Check2"
      Item(1).Caption =   "Vehiculo"
      Item(1).ControlCount=   26
      Item(1).Control(0)=   "Text9"
      Item(1).Control(1)=   "Frame1"
      Item(1).Control(2)=   "Check1"
      Item(1).Control(3)=   "Text10"
      Item(1).Control(4)=   "Command2"
      Item(1).Control(5)=   "Command1"
      Item(1).Control(6)=   "vehi"
      Item(1).Control(7)=   "vh"
      Item(1).Control(8)=   "combustible"
      Item(1).Control(9)=   "combus"
      Item(1).Control(10)=   "grupodc"
      Item(1).Control(11)=   "grupo"
      Item(1).Control(12)=   "DataCombo3"
      Item(1).Control(13)=   "conduc"
      Item(1).Control(14)=   "DTPicker7"
      Item(1).Control(15)=   "Label20"
      Item(1).Control(16)=   "Label19"
      Item(1).Control(17)=   "Label12"
      Item(1).Control(18)=   "Label11"
      Item(1).Control(19)=   "Label10"
      Item(1).Control(20)=   "Label9"
      Item(1).Control(21)=   "Label8"
      Item(1).Control(22)=   "Label7"
      Item(1).Control(23)=   "Frame3"
      Item(1).Control(24)=   "Frame8"
      Item(1).Control(25)=   "Frame2"
      Begin VB.TextBox Text9 
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
         Left            =   -68080
         MaxLength       =   20
         TabIndex        =   107
         Top             =   2400
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0073DCFF&
         Caption         =   "Informacion de Compra / arrendamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -64600
         TabIndex        =   94
         Top             =   600
         Visible         =   0   'False
         Width           =   5415
         Begin VB.TextBox Text26 
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
            TabIndex        =   99
            Top             =   1320
            Width           =   1815
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
            TabIndex        =   98
            Top             =   2040
            Width           =   3255
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   20
            TabIndex        =   97
            Text            =   "0.00"
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H0073DCFF&
            Caption         =   "Vehiculo Alguilado"
            Height          =   375
            Left            =   2040
            TabIndex        =   96
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0073DCFF&
            Caption         =   "Vehiculo Propio"
            Height          =   375
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   300
            Left            =   2040
            TabIndex        =   100
            Top             =   600
            Width           =   1860
            _ExtentX        =   3281
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
            Format          =   17104897
            CurrentDate     =   38651
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   2040
            TabIndex        =   101
            Top             =   1680
            Width           =   1860
            _ExtentX        =   3281
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
            Format          =   17104897
            CurrentDate     =   38651
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota sobre la compra:"
            Height          =   195
            Left            =   120
            TabIndex        =   106
            Top             =   2040
            Width           =   1560
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Garantia caduca el:"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   1680
            Width           =   1395
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comprado o alguilado a:"
            Height          =   195
            Left            =   120
            TabIndex        =   104
            Top             =   1320
            Width           =   1710
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kilometraje en la compra:"
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   960
            Width           =   1770
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de adquisicion:"
            Height          =   195
            Left            =   120
            TabIndex        =   102
            Top             =   600
            Width           =   1560
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0073DCFF&
         Caption         =   "Vehiculo Inactivo"
         Height          =   195
         Left            =   -68080
         TabIndex        =   93
         Top             =   3960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   -68080
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   92
         Top             =   3120
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         Height          =   375
         Left            =   -65200
         TabIndex        =   91
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   375
         Left            =   -65200
         TabIndex        =   90
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0073DCFF&
         Caption         =   "Permiso del Vehiculo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -69880
         TabIndex        =   78
         Top             =   4320
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox Text16 
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
            Left            =   1440
            MaxLength       =   100
            TabIndex        =   81
            Top             =   240
            Width           =   3615
         End
         Begin VB.TextBox Text17 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1530
            Left            =   1440
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   80
            Top             =   600
            Width           =   3615
         End
         Begin VB.TextBox Text18 
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
            MaxLength       =   20
            TabIndex        =   79
            Top             =   3000
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   300
            Left            =   1440
            TabIndex        =   82
            Top             =   2160
            Width           =   2820
            _ExtentX        =   4974
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
            Format          =   17104897
            CurrentDate     =   38651
         End
         Begin MSComCtl2.DTPicker DTPicker6 
            Height          =   300
            Left            =   1440
            TabIndex        =   83
            Top             =   2520
            Width           =   2820
            _ExtentX        =   4974
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
            Format          =   17104897
            CurrentDate     =   38651
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia:"
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota:"
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   600
            Width           =   390
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicia:"
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   2160
            Width           =   915
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expira en:"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   2520
            Width           =   705
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mostrar alerta "
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   3000
            Width           =   1005
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "dias antes del vencimiento "
            Height          =   195
            Left            =   1800
            TabIndex        =   84
            Top             =   3000
            Width           =   1920
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H0073DCFF&
         Caption         =   "Imagen 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -58720
         TabIndex        =   75
         Top             =   600
         Visible         =   0   'False
         Width           =   3615
         Begin XtremeSuiteControls.PushButton PushButton7 
            Height          =   375
            Left            =   3120
            TabIndex        =   76
            Top             =   1680
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   79
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmAgregarActivoFijo.frx":058A
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton PushButton8 
            Height          =   375
            Left            =   3120
            TabIndex        =   77
            Top             =   2160
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   79
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmAgregarActivoFijo.frx":0B24
            ImageAlignment  =   0
         End
         Begin VB.Image foto2 
            BorderStyle     =   1  'Fixed Single
            Height          =   2415
            Left            =   120
            Picture         =   "FrmAgregarActivoFijo.frx":10BE
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E8AA8C&
         Caption         =   "Imagen 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   13440
         TabIndex        =   72
         Top             =   840
         Width           =   3615
         Begin XtremeSuiteControls.PushButton PushButton5 
            Height          =   375
            Left            =   3120
            TabIndex        =   73
            Top             =   1680
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   79
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmAgregarActivoFijo.frx":5175
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   375
            Left            =   3120
            TabIndex        =   74
            Top             =   2160
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   79
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmAgregarActivoFijo.frx":570F
            ImageAlignment  =   0
         End
         Begin VB.Image foto 
            BorderStyle     =   1  'Fixed Single
            Height          =   2415
            Left            =   120
            Picture         =   "FrmAgregarActivoFijo.frx":5CA9
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0073DCFF&
         Caption         =   "Seguros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -64480
         TabIndex        =   56
         Top             =   4320
         Visible         =   0   'False
         Width           =   4935
         Begin VB.TextBox Text19 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   1440
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   61
            Top             =   2040
            Width           =   3375
         End
         Begin VB.TextBox Text15 
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
            MaxLength       =   20
            TabIndex        =   60
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox Text14 
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
            Left            =   1440
            MaxLength       =   100
            TabIndex        =   59
            Top             =   960
            Width           =   3375
         End
         Begin VB.TextBox Text13 
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
            Left            =   1440
            MaxLength       =   100
            TabIndex        =   58
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox Text12 
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
            Left            =   1440
            MaxLength       =   100
            TabIndex        =   57
            Top             =   240
            Width           =   3375
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   300
            Left            =   1440
            TabIndex        =   62
            Top             =   1320
            Width           =   3300
            _ExtentX        =   5821
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
            Format          =   17104897
            CurrentDate     =   38651
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   300
            Left            =   1440
            TabIndex        =   63
            Top             =   1680
            Width           =   3300
            _ExtentX        =   5821
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
            Format          =   17104897
            CurrentDate     =   38651
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota:"
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   2040
            Width           =   390
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "dias antes del vencimiento "
            Height          =   195
            Left            =   1800
            TabIndex        =   70
            Top             =   3000
            Width           =   1920
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mostrar alerta "
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   3000
            Width           =   1005
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expira en:"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicia:"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia:"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   960
            Width           =   825
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compañia de seg."
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Asegurador:"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E8AA8C&
         Caption         =   "Imagen 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   13440
         TabIndex        =   53
         Top             =   3960
         Width           =   3615
         Begin XtremeSuiteControls.PushButton PushButton4 
            Height          =   375
            Left            =   3120
            TabIndex        =   54
            Top             =   1680
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   79
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmAgregarActivoFijo.frx":9D60
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton PushButton6 
            Height          =   375
            Left            =   3120
            TabIndex        =   55
            Top             =   2160
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   79
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmAgregarActivoFijo.frx":A2FA
            ImageAlignment  =   0
         End
         Begin VB.Image foto1 
            BorderStyle     =   1  'Fixed Single
            Height          =   2415
            Left            =   120
            Picture         =   "FrmAgregarActivoFijo.frx":A894
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E8AA8C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E8AA8C&
         Height          =   4575
         Left            =   5760
         TabIndex        =   30
         Top             =   2160
         Width           =   7335
         Begin VB.TextBox Text20 
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
            Left            =   5280
            MaxLength       =   20
            TabIndex        =   128
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox Text7 
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
            Left            =   5280
            MaxLength       =   20
            TabIndex        =   127
            Top             =   3240
            Width           =   1215
         End
         Begin VB.TextBox Text23 
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
            Left            =   5280
            MaxLength       =   20
            TabIndex        =   126
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton Command6 
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
            Left            =   6600
            Picture         =   "FrmAgregarActivoFijo.frx":E94B
            Style           =   1  'Graphical
            TabIndex        =   125
            Top             =   3240
            Width           =   375
         End
         Begin VB.CommandButton Command5 
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
            Left            =   6600
            Picture         =   "FrmAgregarActivoFijo.frx":EA99
            Style           =   1  'Graphical
            TabIndex        =   124
            Top             =   3720
            Width           =   375
         End
         Begin VB.CommandButton Command4 
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
            Left            =   6600
            Picture         =   "FrmAgregarActivoFijo.frx":EBE7
            Style           =   1  'Graphical
            TabIndex        =   123
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox Text25 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   37
            Text            =   "0.00"
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   36
            Text            =   "0.00"
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text22 
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
            Left            =   5280
            MaxLength       =   20
            TabIndex        =   35
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox Text21 
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
            Left            =   5280
            MaxLength       =   20
            TabIndex        =   34
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox TxtValorEstMeses 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   33
            Text            =   "0.00"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox TxtValorRescate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   32
            Text            =   "0.00"
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox TxtDepAcumulada 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   2280
            Visible         =   0   'False
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTPicker8 
            Height          =   300
            Left            =   5280
            TabIndex        =   38
            Top             =   360
            Width           =   1500
            _ExtentX        =   2646
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
            Format          =   17104897
            CurrentDate     =   38651
         End
         Begin MSComCtl2.DTPicker TxtFechaUltDep 
            Height          =   285
            Left            =   5280
            TabIndex        =   39
            Top             =   1800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   37992
         End
         Begin MSDataListLib.DataCombo DBGrupos 
            Height          =   315
            Left            =   1320
            TabIndex        =   40
            Top             =   2760
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker TxtFechaBaja 
            Height          =   285
            Left            =   5280
            TabIndex        =   41
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   37992
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Contable:"
            Height          =   195
            Left            =   3960
            TabIndex        =   131
            Top             =   2760
            Width           =   1230
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Gastos:"
            Height          =   195
            Left            =   3960
            TabIndex        =   130
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Deprec.:"
            Height          =   195
            Left            =   3960
            TabIndex        =   129
            Top             =   3720
            Width           =   1170
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IVA:"
            Height          =   195
            Left            =   840
            TabIndex        =   52
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F. Compra:"
            Height          =   195
            Left            =   4320
            TabIndex        =   51
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Factura:"
            Height          =   195
            Left            =   4440
            TabIndex        =   50
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia:"
            Height          =   195
            Left            =   4200
            TabIndex        =   49
            Top             =   960
            Width           =   825
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Original"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Ultima Dep"
            Height          =   255
            Left            =   3720
            TabIndex        =   47
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Estimado en Meses"
            Height          =   375
            Left            =   240
            TabIndex        =   46
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Rescate"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label63 
            BackStyle       =   0  'Transparent
            Caption         =   "Dep Acumulada"
            Height          =   375
            Left            =   240
            TabIndex        =   44
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label62 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Baja"
            Height          =   255
            Left            =   3840
            TabIndex        =   43
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo de Cuentas"
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   2640
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E8AA8C&
         Height          =   5895
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   5415
         Begin VB.TextBox DBCodigo 
            Height          =   285
            Left            =   1800
            TabIndex        =   132
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox TxtMarbete 
            Height          =   285
            Left            =   1800
            TabIndex        =   17
            Top             =   4080
            Width           =   2175
         End
         Begin VB.TextBox TxtLocalizacion 
            Height          =   285
            Left            =   1800
            TabIndex        =   16
            Top             =   3720
            Width           =   2175
         End
         Begin VB.TextBox Text24 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   15
            Top             =   720
            Width           =   3375
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
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   14
            Top             =   1560
            Width           =   1815
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
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   13
            Top             =   1920
            Width           =   1815
         End
         Begin VB.TextBox Text4 
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
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   12
            Top             =   3000
            Width           =   1815
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
            Height          =   330
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   11
            Top             =   2640
            Width           =   1815
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
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   10
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox Text5 
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
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   9
            Top             =   3360
            Width           =   3495
         End
         Begin MSDataListLib.DataCombo DBEncargado 
            Height          =   315
            Left            =   1800
            TabIndex        =   18
            Top             =   4440
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "Empleado Encag:"
            Height          =   255
            Left            =   480
            TabIndex        =   29
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Numero Marbete:"
            Height          =   255
            Left            =   480
            TabIndex        =   28
            Top             =   4080
            Width           =   1335
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "Localizacion:"
            Height          =   255
            Left            =   480
            TabIndex        =   27
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
            Height          =   255
            Left            =   600
            TabIndex        =   26
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label50 
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo Activo"
            Height          =   255
            Left            =   480
            TabIndex        =   25
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad #"
            Height          =   195
            Left            =   840
            TabIndex        =   24
            Top             =   1560
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marca:"
            Height          =   195
            Left            =   960
            TabIndex        =   23
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color:"
            Height          =   195
            Left            =   1080
            TabIndex        =   22
            Top             =   3120
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año:"
            Height          =   195
            Left            =   1200
            TabIndex        =   21
            Top             =   2760
            Width           =   330
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo:"
            Height          =   195
            Left            =   960
            TabIndex        =   20
            Top             =   2400
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero de serie (VIN):"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   3360
            Width           =   1605
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E8AA8C&
         Height          =   855
         Left            =   5760
         TabIndex        =   5
         Top             =   840
         Width           =   3135
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   615
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Agregar"
            ForeColor       =   0
            Appearance      =   6
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   615
            Left            =   1680
            TabIndex        =   7
            Top             =   120
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Salir"
            ForeColor       =   0
            Appearance      =   6
            ImageAlignment  =   0
         End
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00E8AA8C&
         Caption         =   "Activo dado de Alta"
         Enabled         =   0   'False
         Height          =   195
         Left            =   9000
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E8AA8C&
         Caption         =   "Activo Trasladado"
         Enabled         =   0   'False
         Height          =   195
         Left            =   9000
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E8AA8C&
         Caption         =   "Activo Inactivo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   9000
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo vehi 
         Bindings        =   "FrmAgregarActivoFijo.frx":ED35
         DataField       =   "idvh"
         Height          =   360
         Left            =   -68080
         TabIndex        =   108
         Top             =   960
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483640
         ListField       =   "DescriCpcion"
         BoundColumn     =   "idvh"
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
      Begin MSAdodcLib.Adodc vh 
         Height          =   330
         Left            =   -66400
         Top             =   960
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
      Begin MSDataListLib.DataCombo combustible 
         Bindings        =   "FrmAgregarActivoFijo.frx":ED47
         DataField       =   "idcombus"
         Height          =   360
         Left            =   -68080
         TabIndex        =   109
         Top             =   1320
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483640
         ListField       =   "descripcombus"
         BoundColumn     =   "idcombus"
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
      Begin MSAdodcLib.Adodc combus 
         Height          =   330
         Left            =   -66400
         Top             =   1320
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
      Begin MSDataListLib.DataCombo grupodc 
         Bindings        =   "FrmAgregarActivoFijo.frx":ED5D
         DataField       =   "IdSede"
         Height          =   360
         Left            =   -68080
         TabIndex        =   110
         Top             =   1680
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483640
         ListField       =   "Descripcion"
         BoundColumn     =   "IdSede"
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
      Begin MSAdodcLib.Adodc grupo 
         Height          =   330
         Left            =   -66400
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
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "FrmAgregarActivoFijo.frx":ED72
         DataField       =   "CodEncargado"
         Height          =   360
         Left            =   -68080
         TabIndex        =   111
         Top             =   2040
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483640
         ListField       =   "NombreEncargado"
         BoundColumn     =   "CodEncargado"
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
         Left            =   -66400
         Top             =   2040
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
      Begin MSComCtl2.DTPicker DTPicker7 
         Height          =   300
         Left            =   -68080
         TabIndex        =   112
         Top             =   2760
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
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
         Format          =   17104897
         CurrentDate     =   38651
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Datos Generales del Activo"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   122
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label64 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Datos Depreciacion"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   121
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de renovacion:"
         Height          =   195
         Left            =   -69760
         TabIndex        =   120
         Top             =   2760
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa:"
         Height          =   195
         Left            =   -69760
         TabIndex        =   119
         Top             =   2400
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota:"
         Height          =   195
         Left            =   -69640
         TabIndex        =   118
         Top             =   3240
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conductor:"
         Height          =   195
         Left            =   -69760
         TabIndex        =   117
         Top             =   2040
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   -69760
         TabIndex        =   116
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Combustible"
         Height          =   195
         Left            =   -69760
         TabIndex        =   115
         Top             =   1320
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Informacion si el Activo es un Vehiculo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69760
         TabIndex        =   114
         Top             =   480
         Visible         =   0   'False
         Width           =   3990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de vehiculo:"
         Height          =   195
         Left            =   -69760
         TabIndex        =   113
         Top             =   960
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin MSAdodcLib.Adodc DtaCuentas 
      Height          =   375
      Left            =   3360
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "DtaCuentas"
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
   Begin MSAdodcLib.Adodc DtaEncargado 
      Height          =   375
      Left            =   3360
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "DtaEncargado"
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
   Begin MSAdodcLib.Adodc DtaActivoFijo 
      Height          =   375
      Left            =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "DtaActivoFijo"
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
   Begin MSAdodcLib.Adodc DtaGrupoCuentas 
      Height          =   330
      Left            =   3360
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "DtaGrupoCuentas"
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
   Begin MSAdodcLib.Adodc DtaNacceso 
      Height          =   330
      Left            =   120
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "DtaNacceso"
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
   Begin MSAdodcLib.Adodc DtaBusca 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "DtaBusca"
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
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmAgregarActivoFijo.frx":ED88
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   10590
   End
End
Attribute VB_Name = "FrmAgregarActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isvh, iscombu As Integer
Dim hl As String
Public actualiza As Integer
Dim ruta, Ruta1, ruta2 As String
Public idregAF As Integer

Private Sub combustible_Click(Area As Integer)
combustible.BackColor = vb3DLight
End Sub

Private Sub combustible_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grupodc.SetFocus
End If
End Sub

Private Sub combustible_LostFocus()
combustible.BackColor = vbWhite
End Sub

Private Sub Command1_Click()
isvh = 1
iscombu = 0
llamaconsulta 1
End Sub
Private Sub guardavh()
Set rsa = Nothing
SQL = "select * from dbo.CataVH"
rsa.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
rsa.AddNew
rsa!descriCpcion = Trim(hl)
rsa.Update
cvh
End Sub
Private Sub guardacombu()
Set rsa = Nothing
SQL = "select * from dbo.CatCombus"
rsa.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
rsa.AddNew
rsa!DescripCombus = Trim(hl)
rsa.Update
ccombus
End Sub
Private Sub llamaconsulta(opcion As Integer)
If opcion = 1 Then
    hl = InputBox("Digite Descripcion del Vehiculo:", "Sistema Activo Fijo")
    If Len(hl) > 0 Then
        guardavh
    End If
End If

If opcion = 2 Then
    hl = InputBox("Digite Descripcion del Combustible:", "Sistema Activo Fijo")
    If Len(hl) > 0 Then
        guardacombu
    End If
End If

End Sub

Private Sub Command2_Click()
isvh = 0
iscombu = 1
llamaconsulta 2
End Sub

Private Sub Command4_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.Text20.Text = FrmConsulta.Cuenta
End Sub

Private Sub Command5_Click()
QueProducto = "Cuenta"
FrmConsulta.Show 1
Me.Text23.Text = FrmConsulta.Cuenta
End Sub

Private Sub Command6_Click()
QueProducto = "Cuenta"
FrmConsulta.Show 1
Me.Text7.Text = FrmConsulta.Cuenta
End Sub

Private Sub DataCombo3_Click(Area As Integer)
DataCombo3.BackColor = vb3DLight
End Sub

Private Sub DataCombo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text9.SetFocus
End If
End Sub

Private Sub DataCombo3_LostFocus()
DataCombo3.BackColor = vbWhite
End Sub

Private Sub DataCombo4_Click(Area As Integer)
DataCombo4.BackColor = vb3DLight
End Sub

Private Sub DataCombo4_LostFocus()
DataCombo4.BackColor = vbWhite
End Sub



Private Sub Form_Load()
limpiar
hl = ""
cvh
ccombus
'CargaADODC "_Sede", grupo, "1", grupodc.Name, "Trim", conexion, Me, "order by Descripcion"
'CargaADODC "Empleado", conduc, "", DataCombo3.Name, "Trim", conexion, Me, "ORDER BY CodEncargado"
With Me.conduc
  .ConnectionString = Conexion
  .RecordSource = "SELECT Encargado.* From Encargado ORDER BY CodEncargado"
  .Refresh
End With


If actualiza <> 0 Then
    cargardatosAF
End If


With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Accesos"
   .Refresh
End With

With Me.DtaGrupoCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from GrupoCuentas"
   .Refresh
   .ConnectionString = Conexion
   .RecordSource = "Select * from Cuentas"
   .Refresh
End With
LlenarDataCombos DtaGrupoCuentas, DBGrupos, "DescripcionGrupo", "CodGrupo"

With Me.DtaBusca
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaActivoFijo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from CatalogoActivoFijo"
   .Refresh
End With



'LlenarDataCombos DtaActivoFijo, Me.DBCodigo, "idReg", "CodCuenta"

With Me.DtaCuentas
   '.DatabaseName = Ruta
End With


With Me.DtaEncargado
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Encargado"
   .Refresh
End With
LlenarDataCombos DtaEncargado, DBEncargado, "NombreEncargado", "CodEncargado"


'
'Me.BackColor = RGB(216, 228, 248)
'Me.Frame1.BackColor = RGB(216, 228, 248)
'Me.Frame2.BackColor = RGB(216, 228, 248)
'Me.Frame3.BackColor = RGB(216, 228, 248)
'Me.Frame4.BackColor = RGB(216, 228, 248)
'Me.Frame5.BackColor = RGB(216, 228, 248)
'Me.Frame6.BackColor = RGB(216, 228, 248)
'Me.Frame7.BackColor = RGB(216, 228, 248)
'Me.Frame8.BackColor = RGB(216, 228, 248)
'Me.Option1.BackColor = RGB(216, 228, 248)
'Me.Option2.BackColor = RGB(216, 228, 248)
'Me.Check1.BackColor = RGB(216, 228, 248)
'Me.Check2.BackColor = RGB(216, 228, 248)
'Me.Check3.BackColor = RGB(216, 228, 248)
'Me.Check4.BackColor = RGB(216, 228, 248)

End Sub
Private Sub cargardatosAF()
Set rsa = Nothing
If actualiza = 1 Then
    SQL = "select * from dbo.CatalogoActivoFijo where idreg=" & idregAF & ""
End If
rsa.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic

txttipopla.Text = rsa!Unidad
Text1.Text = rsa!marca
Text2.Text = rsa!modelo
Text3.Text = rsa!Año
Text4.Text = rsa!Color
Text5.Text = rsa!Serie
Text24.Text = rsa!DescripcionAF
Text20.Text = rsa!CNTACONTABLE
Text7.Text = rsa!CuentaGastos
Text23.Text = rsa!CuentaDepreciacion


    With Me.DtaEncargado
       '.DatabaseName = Ruta
       .ConnectionString = Conexion
       .RecordSource = "Select * from Encargado"
       .Refresh
    End With

   '/////Busco el Dodigo del Encargado/////////////////
  Criterio = "CodEncargado='" & rsa!CodEncargado & "'"
  If Me.DtaEncargado.Recordset.RecordCount <> 0 Then DtaEncargado.Recordset.MoveFirst
  Me.DtaEncargado.Recordset.Find (Criterio)
  If Not DtaEncargado.Recordset.EOF Then
    Me.DBEncargado.Text = DtaEncargado.Recordset!NombreEncargado
  End If



  If Not IsNull(rsa!Localizacion) Then
     Me.DBCodigo.Text = rsa!CodCuenta
  End If
  Me.TxtLocalizacion.Text = rsa!Localizacion
  If Not IsNull(rsa!NumeroMarbete) Then
    Me.TxtMarbete.Text = rsa!NumeroMarbete
  End If
  Me.TxtFechaUltDep.Value = rsa!FechaUltimaDepre
  Me.TxtValorEstMeses.Text = rsa!ValorEstimadoMeses
  Me.TxtValorRescate.Text = rsa!ValorRescate
  If Not IsNull(rsa!FechaBaja) Then
   Me.TxtFechaBaja.Value = rsa!FechaBaja
  End If

If IsNull(rsa!tipovehiculo) Or rsa!tipovehiculo = "" Then
Else
    vehi.BoundText = rsa!tipovehiculo
End If
vehi.Text = rsa!descriptipoveh
If IsNull(rsa!tipocombus) Then
    combustible.Text = ""
Else
combustible.BoundText = rsa!tipocombus
End If

If IsNull(rsa!grupo) Then
    grupodc.Text = ""
Else
grupodc.BoundText = rsa!grupo
End If


grupodc.Text = rsa!descrigrupo
If IsNull(rsa!codconductor) Then
    DataCombo3.Text = ""
Else
    DataCombo3.BoundText = rsa!codconductor
End If
DataCombo3.Text = rsa!nombreconduc
Text9.Text = rsa!placa
DTPicker7.Value = Format(rsa!frenovacion, "DD/MM/YYYY")

If IsNull(rsa!Nota) Or rsa!Nota = "" Then
    Text10.Text = ""
Else
    Text10.Text = rsa!Nota
End If
If rsa!isvehipropio = 1 Then
Option1.Value = True
Else
Option1.Value = False
End If
DTPicker2.Value = Format(rsa!fadquicisionvh, "DD/MM/YYYY")

If IsNull(rsa!kilomcompravh) Then
    Text6.Text = ""
Else
    Text6.Text = rsa!kilomcompravh
End If
Text26.Text = rsa!compradooalqui
If IsNull(rsa!costovh) Then
    Text11.Text = ""
Else
    Text11.Text = rsa!costovh
End If
If IsNull(rsa!ivavh) Then
    Text25.Text = ""
Else
    Text25.Text = rsa!ivavh
End If
DTPicker1.Value = Format(rsa!garantiacaduvh, "DD/MM/YYYY")

If IsNull(rsa!notacompravh) Then
    Text8.Text = ""
Else
    Text8.Text = rsa!notacompravh
End If
If Not IsNull(rsa!Aseguradorvh) Then
  Text12.Text = rsa!Aseguradorvh
End If
If Not IsNull(rsa!compasegvh) Then
  Text13.Text = rsa!compasegvh
End If
If Not IsNull(rsa!referencia) Then
 Text14.Text = rsa!referencia
End If
DTPicker3.Value = Format(rsa!finiasevh, "DD/MM/YYYY")
DTPicker4.Value = Format(rsa!ffinasevh, "DD/MM/YYYY")

If IsNull(rsa!notaasevh) Then
    Text19.Text = ""
Else
    Text19.Text = rsa!notaasevh
End If

If Not IsNull(rsa!perrefvh) Then
  Text16.Text = rsa!perrefvh
End If

If IsNull(rsa!notapervh) Then
    Text17.Text = ""
Else
    Text17.Text = rsa!notapervh
End If
DTPicker5.Value = Format(rsa!finiper, "DD/MM/YYYY")
DTPicker6.Value = Format(rsa!ffinper, "DD/MM/YYYY")

If rsa!alarmaseguro = 0 Then
Text15.Text = "0"
Else
Text15.Text = rsa!alarmaseguro
End If

If rsa!alarmapermiso = 0 Then
Text18.Text = "0"
Else
Text18.Text = rsa!alarmapermiso
End If
Text20.Text = rsa!CNTACONTABLE
Text21.Text = rsa!refegeneral
Text22.Text = rsa!Factura
DTPicker8.Value = Format(rsa!fcompragen, "DD/MM/YYYY")

If IsNull(rsa!costogen) Then
    Text11.Text = ""
Else
    Text11.Text = rsa!costogen
End If

If IsNull(rsa!ivagen) Then
    Text25.Text = ""
Else
    Text25.Text = rsa!ivagen
End If

If rsa!dadobaja = 1 Or rsa!dadobaja = True Then
Check2.Value = 1
Else
Check2.Value = 0
End If

If rsa!DatoAlta = 1 Or rsa!DatoAlta = True Then
Check4.Value = 1
Else
Check4.Value = 0
End If


If rsa!Trasladado = 1 Or rsa!Trasladado = True Then
Check3.Value = 1
Else
Check3.Value = 0
End If
If IsNull(rsa!dirfoto) Or rsa!dirfoto = "" Then
    ruta = ""
    If Dir(ruta) <> "" Then
      foto.Picture = LoadPicture(ruta)
    End If
Else
    ruta = rsa!dirfoto
    If Dir(ruta) <> "" Then
     foto.Picture = LoadPicture(ruta)
    End If
End If

If IsNull(rsa!dirfoto1) Or rsa!dirfoto1 = "" Then
    Ruta1 = ""
    If Dir(Ruta1) <> "" Then
     foto1.Picture = LoadPicture(Ruta1)
    End If
Else
    Ruta1 = rsa!dirfoto1
    foto1.Picture = LoadPicture(Ruta1)
End If

If IsNull(rsa!dirfoto2) Or rsa!dirfoto2 = "" Then
    ruta2 = ""
    If Dir(ruta2) <> "" Then
     foto2.Picture = LoadPicture(ruta2)
    End If
Else
    ruta2 = rsa!dirfoto2
    If Dir(ruta2) <> "" Then
      foto2.Picture = LoadPicture(ruta2)
    End If
End If
End Sub
Private Sub limpiar()
'txttipopla.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text24.Text = ""
Text20.Text = ""
Text7.Text = ""
Text23.Text = ""
vehi.Text = ""
combustible.Text = ""
grupodc.Text = ""
grupodc.Text = ""
DataCombo3.BoundText = ""
DataCombo3.Text = ""
Text9.Text = ""
DTPicker7.Value = Now
Text10.Text = ""
Option1.Value = False
DTPicker2.Value = Now
Text6.Text = ""
Text26.Text = ""
Text11.Text = ""
Text25.Text = ""
DTPicker1.Value = Now
Text8.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
DTPicker3.Value = Now
DTPicker4.Value = Now
Text19.Text = ""
Text16.Text = ""
Text17.Text = ""
DTPicker5.Value = Now
DTPicker6.Value = Now
Text15.Text = ""
Text18.Text = ""
Text20.Text = ""
Text21.Text = ""
Text22.Text = ""
DTPicker8.Value = Now
Text11.Text = ""
Text25.Text = ""
Check2.Value = 0
Check4.Value = 0
Check3.Value = 0

End Sub

Private Sub ccombus()
CargaADODC "CatCombus", combus, "", combustible.Name, "Trim", Conexion, Me, "order by idcombus"
End Sub
Private Sub cvh()
CargaADODC "CataVH", vh, "", vehi.Name, "Trim", Conexion, Me, "order by idvh"
End Sub

Private Sub grupodc_Click(Area As Integer)
grupodc.BackColor = vb3DLight
End Sub

Private Sub grupodc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DataCombo3.SetFocus
End If
End Sub

Private Sub grupodc_LostFocus()
grupodc.BackColor = vbWhite
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub PushButton1_Click()
On Error GoTo TipoErrs
'la ruta de la imagen en el disco

'ruta inicial del CommonDialog
CommonDialog1.InitDir = "C:\Documents and Settings\Administrador\Mis documentos\Mis imágenes\"

'Titulo del CommonDialog
CommonDialog1.DialogTitle = "Seleccione el archivo gif"

'Extensión del CommonDialog
CommonDialog1.Filter = "Archivos Graficos (*.bmp;*.gif;*.jpg)"

'Abrimos el CommonDialog
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then
   'No se ha seleccionado ningún archivo
   MsgBox "No se ha seleccionado ningún archivo", vbInformation
Else
  'Mostramos la ruta archivo seleccionado
  ruta = CommonDialog1.FileName
  foto.Picture = LoadPicture(CommonDialog1.FileName)
End If

TipoErrs:
 MsgBox err.Description
End Sub

Private Sub PushButton2_Click()
Set rsa = Nothing
If actualiza = 0 Then
    SQL = "select * from dbo.CatalogoActivoFijo"
Else
    SQL = "select * from dbo.CatalogoActivoFijo WHERE idreg=" & idregAF & ""
End If


rsa.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic

If actualiza = 0 Then
    rsa.AddNew
End If


   '/////Busco el Dodigo del Encargado/////////////////
  Criterio = "NombreEncargado='" & Me.DBEncargado.Text & "'"
  If DtaEncargado.Recordset.RecordCount <> 0 Then DtaEncargado.Recordset.MoveFirst
  Me.DtaEncargado.Recordset.Find (Criterio)
  If Not DtaEncargado.Recordset.EOF Then
    CodEncargado = DtaEncargado.Recordset!CodEncargado
  Else
    MsgBox "Seleccione de la Lista de Encargado", vbCritical, "Sistema Contable"
    Exit Sub
  End If
 
   rsa!CodEncargado = CodEncargado


If DBCodigo.Text <> "" Then
  rsa!CodCuenta = Me.DBCodigo.Text
End If


If TxtLocalizacion.Text <> "" Then
  rsa!Localizacion = Me.TxtLocalizacion.Text
End If

If TxtFechaUltDep.Value <> "" Then
  rsa!FechaUltimaDepre = Me.TxtFechaUltDep.Value
End If

If TxtValorEstMeses.Text <> "" Then
  rsa!ValorEstimadoMeses = Me.TxtValorEstMeses.Text
End If

If TxtValorRescate.Text <> "" Then
  rsa!ValorRescate = Me.TxtValorRescate.Text
End If


If TxtMarbete.Text <> "" Then
  rsa!NumeroMarbete = Me.TxtMarbete.Text
End If

If txttipopla.Text <> "" Then
  rsa!Unidad = txttipopla.Text
Else
  rsa!Unidad = 1
End If

If Text1.Text <> "" Then
  rsa!marca = Text1.Text
End If
If Text2.Text <> "" Then
 rsa!modelo = Text2.Text
End If

If Text3.Text = "" Then
    rsa!Año = 0
Else
    rsa!Año = Text3.Text
End If

If Text4.Text <> "" Then
  rsa!Color = Text4.Text
End If

If Text5.Text <> "" Then
 rsa!Serie = Text5.Text
End If
If Text24.Text <> "" Then
  rsa!DescripcionAF = Text24.Text
End If

If Text20.Text <> "" Then
 rsa!CNTACONTABLE = Text20.Text
End If

If Text7.Text <> "" Then
  rsa!CuentaGastos = Text7.Text
End If

If Text23.Text <> "" Then
 rsa!CuentaDepreciacion = Text23.Text
End If

If vehi.BoundText = "" Then

Else
    rsa!tipovehiculo = vehi.BoundText
End If
rsa!descriptipoveh = vehi.Text
If combustible.BoundText = "" Then

Else
    rsa!tipocombus = combustible.BoundText
End If
rsa!descriptipocombus = combustible.Text
If grupodc.BoundText = "" Then

Else
    rsa!grupo = grupodc.BoundText
End If
rsa!descrigrupo = grupodc.Text
If DataCombo3.BoundText = "" Then

Else
    rsa!codconductor = DataCombo3.BoundText
End If
rsa!nombreconduc = DataCombo3.Text
rsa!placa = Text9.Text
rsa!frenovacion = Format(DTPicker7, "DD/MM/YYYY")
If Text10.Text = "" Then

Else
    rsa!Nota = Text10.Text
End If
If Option1.Value = True Then
    rsa!isvehipropio = 1
Else
    rsa!isvehipropio = 0
End If
rsa!fadquicisionvh = Format(DTPicker2, "DD/MM/YYYY")
If Text6.Text = "" Then

Else
    rsa!kilomcompravh = Text6.Text
End If
rsa!compradooalqui = Text26.Text
If Text11.Text = "" Then

Else
    rsa!costovh = Text11.Text
End If
If Text25.Text = "" Then
Else
    rsa!ivavh = Text25.Text
End If
rsa!garantiacaduvh = Format(DTPicker1, "DD/MM/YYYY")
If Text8.Text = "" Then
Else
    rsa!notacompravh = Text8.Text
End If
If Text12.Text <> "" Then
rsa!Aseguradorvh = Text12.Text
End If
If Text13.Text <> "" Then
  rsa!compasegvh = Text13.Text
End If
If Text14.Text <> "" Then
  rsa!referencia = Text14.Text
End If
rsa!finiasevh = Format(DTPicker3, "DD/MM/YYYY")
rsa!ffinasevh = Format(DTPicker4, "DD/MM/YYYY")
If Text19.Text = "" Then
Else
    rsa!notaasevh = Text19.Text
End If
rsa!perrefvh = Text16.Text
If Text17.Text = "" Then
Else
    rsa!notapervh = Text17.Text
End If

rsa!finiper = Format(DTPicker5, "DD/MM/YYYY")
rsa!ffinper = Format(DTPicker6, "DD/MM/YYYY")
If Text15.Text = "" Then
    rsa!alarmaseguro = 0
Else
    rsa!alarmaseguro = Text15.Text
End If

If Text18.Text = "" Then
    rsa!alarmapermiso = 0
Else
    rsa!alarmapermiso = Text18.Text
End If
rsa!CNTACONTABLE = Text20.Text
rsa!refegeneral = Text21.Text
rsa!Factura = Text22.Text
rsa!fcompragen = Format(DTPicker8, "DD/MM/YYYY")
If Text11.Text = "" Then
Else
    rsa!costogen = Text11.Text
End If
If Text25.Text = "" Then
Else
    rsa!ivagen = Text25.Text
End If
'If Check2.Value = 1 Then
'    rsa!dadobaja = 1
'Else
'    rsa!dadobaja = 0
'End If
'
'If Check4.Value = 1 Then
'    rsa!DatoAlta = 1
'Else
'    rsa!DatoAlta = 0
'End If
'
'
'If Check3.Value = 1 Then
'    rsa!trasladado = 1
'Else
'    rsa!trasladado = 1
'End If
If Len(vehi.Text) > 0 Then
    rsa!isvh = 1
Else
    rsa!isvh = 0
End If
If ruta <> "" Then
  rsa!dirfoto = ruta
End If
If Ruta1 <> "" Then
  rsa!dirfoto1 = Ruta1
End If
If ruta2 <> "" Then
  rsa!dirfoto2 = ruta2
End If
rsa.Update
limpiar
End Sub

Private Sub PushButton3_Click()
actualiza = 0
Unload Me
End Sub

Private Sub PushButton6_Click()
'la ruta de la imagen en el disco

'ruta inicial del CommonDialog
CommonDialog2.InitDir = "C:\Documents and Settings\Administrador\Mis documentos\Mis imágenes\"

'Titulo del CommonDialog
CommonDialog2.DialogTitle = "Seleccione el archivo gif"

'Extensión del CommonDialog
CommonDialog2.Filter = "Archivos Graficos (*.bmp;*.gif;*.jpg)"

'Abrimos el CommonDialog
CommonDialog2.ShowOpen

If CommonDialog2.FileName = "" Then
   'No se ha seleccionado ningún archivo
   MsgBox "No se ha seleccionado ningún archivo", vbInformation
Else
  'Mostramos la ruta archivo seleccionado
  Ruta1 = CommonDialog2.FileName
  foto1.Picture = LoadPicture(CommonDialog2.FileName)
End If
End Sub

Private Sub PushButton8_Click()
'la ruta de la imagen en el disco

'ruta inicial del CommonDialog
CommonDialog3.InitDir = "C:\Documents and Settings\Administrador\Mis documentos\Mis imágenes\"

'Titulo del CommonDialog
CommonDialog3.DialogTitle = "Seleccione el archivo gif"

'Extensión del CommonDialog
CommonDialog3.Filter = "Archivos Graficos (*.bmp;*.gif;*.jpg)"

'Abrimos el CommonDialog
CommonDialog3.ShowOpen

If CommonDialog3.FileName = "" Then
   'No se ha seleccionado ningún archivo
   MsgBox "No se ha seleccionado ningún archivo", vbInformation
Else
  'Mostramos la ruta archivo seleccionado
  ruta2 = CommonDialog3.FileName
  foto2.Picture = LoadPicture(CommonDialog3.FileName)
End If
End Sub

Private Sub Text1_Click()
Text1.BackColor = vb3DLight
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = vbWhite
End Sub

Private Sub Text10_Click()
Text10.BackColor = vb3DLight
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text6.SetFocus
End If
End Sub

Private Sub Text10_LostFocus()
Text10.BackColor = vbWhite
End Sub

Private Sub Text11_Click()
Text11.BackColor = vb3DLight
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text25.SetFocus
End If
End Sub

Private Sub Text11_LostFocus()
Text11.BackColor = vbWhite
End Sub

Private Sub Text12_Click()
Text12.BackColor = vb3DLight
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text13.SetFocus
End If
End Sub

Private Sub Text12_LostFocus()
Text12.BackColor = vbWhite
End Sub

Private Sub Text13_Click()
Text13.BackColor = vb3DLight
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text14.SetFocus
End If
End Sub

Private Sub Text13_LostFocus()
Text13.BackColor = vbWhite
End Sub

Private Sub Text14_Click()
Text14.BackColor = vb3DLight
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text19.SetFocus
End If
End Sub

Private Sub Text14_LostFocus()
Text14.BackColor = vbWhite
End Sub

Private Sub Text15_Click()
Text15.BackColor = vb3DLight
End Sub

Private Sub Text15_LostFocus()
Text15.BackColor = vbWhite
End Sub

Private Sub Text16_Click()
Text16.BackColor = vb3DLight
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text17.SetFocus
End If
End Sub

Private Sub Text16_LostFocus()
Text16.BackColor = vbWhite
End Sub

Private Sub Text17_Click()
Text17.BackColor = vb3DLight
End Sub

Private Sub Text17_LostFocus()
Text17.BackColor = vbWhite
End Sub

Private Sub Text18_Click()
Text18.BackColor = vb3DLight
End Sub

Private Sub Text18_LostFocus()
Text18.BackColor = vbWhite
End Sub

Private Sub Text19_Click()
Text19.BackColor = vb3DLight
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text16.SetFocus
End If
End Sub

Private Sub Text19_LostFocus()
Text19.BackColor = vbWhite
End Sub

Private Sub Text2_Click()
Text2.BackColor = vb3DLight
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
End Sub

Private Sub Text2_LostFocus()
Text2.BackColor = vbWhite
End Sub

Private Sub Text20_Click()
Text20.BackColor = vb3DLight
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text7.SetFocus
End If
End Sub

Private Sub Text20_LostFocus()
Text20.BackColor = vbWhite
End Sub

Private Sub Text21_Click()
Text21.BackColor = vb3DLight
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text22.SetFocus
End If
End Sub

Private Sub Text21_LostFocus()
Text21.BackColor = vbWhite
End Sub

Private Sub Text22_Click()
Text22.BackColor = vb3DLight
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    vehi.SetFocus
End If
End Sub

Private Sub Text22_LostFocus()
Text22.BackColor = vbWhite
End Sub

Private Sub Text23_Click()
Text23.BackColor = vb3DLight
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text11.SetFocus
End If
End Sub

Private Sub Text23_LostFocus()
Text23.BackColor = vbWhite
End Sub

Private Sub Text24_Click()
Text24.BackColor = vb3DLight
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttipopla.SetFocus
End If
End Sub

Private Sub Text24_LostFocus()
Text24.BackColor = vbWhite
End Sub

Private Sub Text25_Click()
Text25.BackColor = vb3DLight
End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text21.SetFocus
End If
End Sub

Private Sub Text25_LostFocus()
Text25.BackColor = vbWhite
End Sub

Private Sub Text26_Click()
Text26.BackColor = vb3DLight
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text8.SetFocus
End If
End Sub

Private Sub Text26_LostFocus()
Text26.BackColor = vbWhite
End Sub

Private Sub Text3_Click()
Text3.BackColor = vb3DLight
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = vbWhite
End Sub

Private Sub Text4_Click()
Text4.BackColor = vb3DLight
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5.SetFocus
End If
End Sub

Private Sub Text4_LostFocus()
Text4.BackColor = vbWhite
End Sub

Private Sub Text5_Click()
Text5.BackColor = vb3DLight
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text20.SetFocus
End If
End Sub

Private Sub Text5_LostFocus()
Text5.BackColor = vbWhite
End Sub

Private Sub Text6_Click()
Text6.BackColor = vb3DLight
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text26.SetFocus
End If
End Sub

Private Sub Text6_LostFocus()
Text6.BackColor = vbWhite
End Sub

Private Sub Text7_Click()
Text7.BackColor = vb3DLight
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text23.SetFocus
End If
End Sub

Private Sub Text7_LostFocus()
Text7.BackColor = vbWhite
End Sub

Private Sub Text8_Click()
Text8.BackColor = vb3DLight
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text12.SetFocus
End If
End Sub

Private Sub Text8_LostFocus()
Text8.BackColor = vbWhite
End Sub

Private Sub Text9_Click()
Text9.BackColor = vb3DLight
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text10.SetFocus
End If
End Sub

Private Sub Text9_LostFocus()
Text9.BackColor = vbWhite
End Sub

Private Sub txttipopla_Click()
txttipopla.BackColor = vb3DLight
End Sub

Private Sub txttipopla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
End If
End Sub

Private Sub txttipopla_LostFocus()
txttipopla.BackColor = vbWhite
End Sub

Private Sub vehi_Click(Area As Integer)
vehi.BackColor = vb3DLight
End Sub

Private Sub vehi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    combustible.SetFocus
End If
End Sub

Private Sub vehi_LostFocus()
vehi.BackColor = vbWhite
End Sub
