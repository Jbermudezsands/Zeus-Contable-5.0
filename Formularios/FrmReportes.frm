VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmReportes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin SmartButtonProject.SmartButton BtnExcel 
      Height          =   855
      Left            =   120
      TabIndex        =   230
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Exportar"
      Picture         =   "FrmReportes.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureLayout   =   7
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   5325
      Left            =   120
      TabIndex        =   47
      Top             =   1200
      Visible         =   0   'False
      Width           =   9480
      Begin VB.PictureBox picTV 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         ScaleHeight     =   855
         ScaleWidth      =   7455
         TabIndex        =   48
         Top             =   1440
         Width           =   7455
      End
      Begin VB.Label Lb9 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando........."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   64
         Top             =   2520
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.Label Lb1 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   63
         Top             =   2520
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.Label Lb2 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando.."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   62
         Top             =   2520
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb3 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   61
         Top             =   2520
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb4 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   60
         Top             =   2520
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb5 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando....."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   59
         Top             =   2520
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.Label Lb6 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando......"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   58
         Top             =   2520
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb7 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando......."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   57
         Top             =   2520
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb8 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando........"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   56
         Top             =   2520
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb0 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   55
         Top             =   2520
         Width           =   6735
      End
      Begin VB.Label Lb10 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando.........."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   54
         Top             =   2520
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Lb11 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando..........."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   53
         Top             =   2520
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Lb12 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando............"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   52
         Top             =   2520
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Lb13 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando............."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   51
         Top             =   2520
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Lb14 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando.............."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   50
         Top             =   2520
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Lb15 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando..............."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   49
         Top             =   2520
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Image Img2 
         Height          =   480
         Left            =   960
         Picture         =   "FrmReportes.frx":183C
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image img1 
         Height          =   480
         Left            =   1920
         Picture         =   "FrmReportes.frx":BCC3
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc AdoConsultas 
      Height          =   375
      Left            =   1680
      Top             =   9720
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
      Caption         =   "AdoConsultas"
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
   Begin VB.TextBox TxtTipoReporte 
      Height          =   375
      Left            =   1560
      TabIndex        =   202
      Top             =   8640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5295
      Left            =   120
      TabIndex        =   73
      Top             =   1200
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9340
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   13225935
      TabCaption(0)   =   "Reportes"
      TabPicture(0)   =   "FrmReportes.frx":1966E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblTransaccion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblTitulo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblNivel2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LblNivel3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LblNivel4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LblNivel5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LblNivel6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LblNivel7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame9"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CmbMoneda"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtTransaccion"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CmbReportes"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "ChkBalanza"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "frameBalanza"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "CmbNivel2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ChkSinNiveles"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CmbNivel3"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "CmbNivel4"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "CmbNivel5"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "CmbNivel6"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "CmbNivel7"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "FrmDepartamento"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "ChkQuitarMovimiento"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "ChkMostrarMovxMes"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Configuracion Reportes"
      TabPicture(1)   =   "FrmReportes.frx":1968A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).Control(1)=   "Frame11"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Configuracion Reportes "
      TabPicture(2)   =   "FrmReportes.frx":196A6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).Control(1)=   "Frame13"
      Tab(2).ControlCount=   2
      Begin VB.CheckBox ChkMostrarMovxMes 
         Caption         =   "Mostrar mov x mes"
         Height          =   375
         Left            =   960
         TabIndex        =   229
         Top             =   4320
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CheckBox ChkQuitarMovimiento 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quitar Movimientos de Cierre"
         Height          =   255
         Left            =   4080
         TabIndex        =   228
         Top             =   4080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Frame FrmDepartamento 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Departamento"
         Height          =   1215
         Left            =   9960
         TabIndex        =   221
         Top             =   3720
         Width           =   4575
         Begin VB.TextBox TxtDptoHasta 
            Height          =   285
            Left            =   1320
            TabIndex        =   227
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox TxtDptoDesde 
            Height          =   285
            Left            =   1320
            TabIndex        =   224
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton Command3 
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
            Left            =   3960
            Picture         =   "FrmReportes.frx":196C2
            Style           =   1  'Graphical
            TabIndex        =   223
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton Command2 
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
            Left            =   3960
            Picture         =   "FrmReportes.frx":19810
            Style           =   1  'Graphical
            TabIndex        =   222
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label62 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
            Height          =   255
            Left            =   240
            TabIndex        =   226
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
            Height          =   255
            Left            =   240
            TabIndex        =   225
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.ComboBox CmbNivel7 
         Height          =   315
         ItemData        =   "FrmReportes.frx":1995E
         Left            =   10560
         List            =   "FrmReportes.frx":1999E
         TabIndex        =   215
         Text            =   "2"
         Top             =   3000
         Width           =   615
      End
      Begin VB.ComboBox CmbNivel6 
         Height          =   315
         ItemData        =   "FrmReportes.frx":199E9
         Left            =   10560
         List            =   "FrmReportes.frx":19A29
         TabIndex        =   214
         Text            =   "2"
         Top             =   2520
         Width           =   615
      End
      Begin VB.ComboBox CmbNivel5 
         Height          =   315
         ItemData        =   "FrmReportes.frx":19A74
         Left            =   10560
         List            =   "FrmReportes.frx":19AB4
         TabIndex        =   213
         Text            =   "2"
         Top             =   2040
         Width           =   615
      End
      Begin VB.ComboBox CmbNivel4 
         Height          =   315
         ItemData        =   "FrmReportes.frx":19AFF
         Left            =   10560
         List            =   "FrmReportes.frx":19B3F
         TabIndex        =   212
         Text            =   "3"
         Top             =   1560
         Width           =   615
      End
      Begin VB.ComboBox CmbNivel3 
         Height          =   315
         ItemData        =   "FrmReportes.frx":19B8A
         Left            =   10560
         List            =   "FrmReportes.frx":19BCA
         TabIndex        =   211
         Text            =   "3"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox ChkSinNiveles 
         Caption         =   "Movimientos sin Niveles"
         Height          =   255
         Left            =   6360
         TabIndex        =   210
         Top             =   4440
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox CmbNivel2 
         Height          =   315
         ItemData        =   "FrmReportes.frx":19C15
         Left            =   5280
         List            =   "FrmReportes.frx":19C55
         TabIndex        =   208
         Text            =   "3"
         Top             =   4440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E0E0E0&
         Height          =   4695
         Left            =   -70560
         TabIndex        =   178
         Top             =   360
         Width           =   4335
         Begin VB.TextBox TxtOtrosGastos 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   207
            Text            =   "Rebajas y Devoluciones S/Ventas"
            Top             =   3360
            Width           =   1935
         End
         Begin VB.TextBox TxtComisioneGanadas 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   197
            Text            =   "Comisiones e Intereses Ganados"
            Top             =   2400
            Width           =   1935
         End
         Begin VB.TextBox TxtComisionesPagadas 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   196
            Text            =   "Comisiones e Intereses Pagados"
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox TxtEnergiaElectrica 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   187
            Text            =   "Energia Electrica y Agua Potable"
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox TxtSueldosAdministrativos 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   186
            Text            =   "Sueldos Administrativos"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox TxtPropaganda 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   180
            Text            =   "Propaganda"
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox TxtSueldosyComisiones 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   179
            Text            =   "Sueldos y Comisiones"
            Top             =   480
            Width           =   1935
         End
         Begin XtremeSuiteControls.CheckBox ChkComisiones 
            Height          =   255
            Left            =   3480
            TabIndex        =   181
            Top             =   480
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkPropaganda 
            Height          =   255
            Left            =   3480
            TabIndex        =   182
            Top             =   840
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkSueldosAdmon 
            Height          =   255
            Left            =   3480
            TabIndex        =   188
            Top             =   1440
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkEnergiaElectrica 
            Height          =   255
            Left            =   3480
            TabIndex        =   189
            Top             =   1800
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkComisionesGanadas 
            Height          =   255
            Left            =   3480
            TabIndex        =   198
            Top             =   2400
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkComisionesPagadas 
            Height          =   255
            Left            =   3480
            TabIndex        =   199
            Top             =   2760
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkOtrosGastos 
            Height          =   255
            Left            =   3480
            TabIndex        =   205
            Top             =   3360
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label59 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Otros Gastos"
            Height          =   255
            Left            =   120
            TabIndex        =   201
            Top             =   3360
            Width           =   1455
         End
         Begin VB.Label Label58 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Otros Gastos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   200
            Top             =   3120
            Width           =   2535
         End
         Begin VB.Label Label57 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ingresos y Gatos Financieros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   195
            Top             =   2160
            Width           =   2535
         End
         Begin VB.Label Label56 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Comisiones Gan"
            Height          =   255
            Left            =   120
            TabIndex        =   194
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label40 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Comisiones Pag"
            Height          =   255
            Left            =   120
            TabIndex        =   193
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label55 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Energia y Agua"
            Height          =   255
            Left            =   120
            TabIndex        =   192
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label54 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sueldos"
            Height          =   255
            Left            =   120
            TabIndex        =   191
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label44 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Gastos de Administracion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   190
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label41 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Propaganda"
            Height          =   255
            Left            =   120
            TabIndex        =   185
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label42 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sdos y Com"
            Height          =   255
            Left            =   120
            TabIndex        =   184
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label43 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Gastos de Ventas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   183
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0E0E0&
         Height          =   4695
         Left            =   -74880
         TabIndex        =   154
         Top             =   360
         Width           =   4335
         Begin VB.TextBox TxtOtrosIngresos 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   206
            Text            =   "Otros  Ingresos"
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox TxtRebajasyDevolucionesVentas 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   175
            Text            =   "Rebajas y Devoluciones S/Ventas"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox TxtIngresoVentas 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   161
            Text            =   "Ingresos por Ventas"
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox TxtServiciosVentas 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   159
            Text            =   "Ventas de Servicios"
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox TxtComisionVentas 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   158
            Text            =   "Ingresos por Comisiones"
            Top             =   1095
            Width           =   1935
         End
         Begin VB.TextBox TxtCostodeVentas 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   157
            Text            =   "Costos de Ventas"
            Top             =   2400
            Width           =   1935
         End
         Begin VB.TextBox TxtCostodeProduccion 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   156
            Text            =   "Costos de Produccion"
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox TxtCostosGeneralesdeProduccion 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   155
            Text            =   "Costos Generales de Produccion"
            Top             =   3120
            Width           =   1935
         End
         Begin XtremeSuiteControls.CheckBox ChkVentas 
            Height          =   375
            Left            =   3480
            TabIndex        =   160
            Top             =   285
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkVentasServicios 
            Height          =   255
            Left            =   3480
            TabIndex        =   162
            Top             =   705
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkIngresoComisiones 
            Height          =   255
            Left            =   3480
            TabIndex        =   163
            Top             =   1170
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkCostoVentas 
            Height          =   255
            Left            =   3480
            TabIndex        =   164
            Top             =   2400
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkCostoProduccion 
            Height          =   255
            Left            =   3480
            TabIndex        =   165
            Top             =   2760
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkCostosGeneralesProduccion 
            Height          =   255
            Left            =   3480
            TabIndex        =   166
            Top             =   3120
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkRebajas 
            Height          =   255
            Left            =   3480
            TabIndex        =   176
            Top             =   1440
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkOtrosIngresos 
            Height          =   255
            Left            =   3480
            TabIndex        =   204
            Top             =   1800
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label60 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Otros Ingresos"
            Height          =   255
            Left            =   120
            TabIndex        =   203
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label49 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Reb y Dev"
            Height          =   255
            Left            =   120
            TabIndex        =   177
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label53 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ingresos/Ventas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   174
            Top             =   120
            Width           =   2535
         End
         Begin VB.Label Label52 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ingresos/Ventas"
            Height          =   255
            Left            =   120
            TabIndex        =   173
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label51 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vtas Servicios"
            Height          =   255
            Left            =   120
            TabIndex        =   172
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label50 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ing/Comisiones"
            Height          =   255
            Left            =   120
            TabIndex        =   171
            Top             =   1110
            Width           =   1215
         End
         Begin VB.Label Label48 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Costo de Ventas"
            Height          =   255
            Left            =   120
            TabIndex        =   170
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label47 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ctos Producc"
            Height          =   255
            Left            =   120
            TabIndex        =   169
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label46 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ctos Grales Prod"
            Height          =   255
            Left            =   120
            TabIndex        =   168
            Top             =   3165
            Width           =   1575
         End
         Begin VB.Label Label45 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Costos "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   167
            Top             =   2160
            Width           =   2535
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   4695
         Left            =   -70680
         TabIndex        =   136
         Top             =   360
         Width           =   4215
         Begin VB.TextBox TxtOtrasCtasCapital 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   42
            Text            =   "Otras Ctas de Capital"
            Top             =   4290
            Width           =   2175
         End
         Begin VB.TextBox TxtProveedores 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   22
            Text            =   "Proveedores"
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox TxtImpuestosxPagar 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   24
            Text            =   "Impuestos x Pagar"
            Top             =   670
            Width           =   2175
         End
         Begin VB.TextBox TxtDocumentosxPagar 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   26
            Text            =   "Documentos x Pagar"
            Top             =   980
            Width           =   2175
         End
         Begin VB.TextBox TxtCobrosAnticipados 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   28
            Text            =   "Cobros Anticipados"
            Top             =   1290
            Width           =   2175
         End
         Begin VB.TextBox TxtPasivosAcumulados 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   30
            Text            =   "Pasivos Acumulados"
            Top             =   1600
            Width           =   2175
         End
         Begin VB.TextBox TxtCtasxPagarLP 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   32
            Text            =   "Cuentas x Pagar LP"
            Top             =   2230
            Width           =   2175
         End
         Begin VB.TextBox TxtDocumentosxPagarLP 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   34
            Text            =   "Documentos x Pagar LP"
            Top             =   2540
            Width           =   2175
         End
         Begin VB.TextBox TxtOtrosPasivos 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   36
            Text            =   "Otros Pasivos"
            Top             =   3160
            Width           =   2175
         End
         Begin VB.TextBox TxtAccionesComunes 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   38
            Text            =   "Acciones Comunes"
            Top             =   3670
            Width           =   2175
         End
         Begin VB.TextBox TxtUtilidadAcumulada 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   40
            Text            =   "Utilidad Acumulada"
            Top             =   3960
            Width           =   2175
         End
         Begin XtremeSuiteControls.CheckBox ChkProveedores 
            Height          =   375
            Left            =   3360
            TabIndex        =   23
            Top             =   280
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkImpuestosxPagar 
            Height          =   255
            Left            =   3360
            TabIndex        =   25
            Top             =   700
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkDocumentosxPagar 
            Height          =   255
            Left            =   3360
            TabIndex        =   27
            Top             =   1050
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkCobrosAnticipados 
            Height          =   255
            Left            =   3360
            TabIndex        =   29
            Top             =   1360
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkPasivosAcum 
            Height          =   255
            Left            =   3360
            TabIndex        =   31
            Top             =   1680
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkCuentasxPagarLP 
            Height          =   255
            Left            =   3360
            TabIndex        =   33
            Top             =   2270
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkDocumentosxPagLP 
            Height          =   255
            Left            =   3360
            TabIndex        =   35
            Top             =   2600
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkOtrosPasivos 
            Height          =   255
            Left            =   3360
            TabIndex        =   37
            Top             =   3220
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkAccionesComunes 
            Height          =   255
            Left            =   3360
            TabIndex        =   39
            Top             =   3720
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkUtilidadAcumulada 
            Height          =   255
            Left            =   3360
            TabIndex        =   41
            Top             =   4020
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkOtrasCtasCapital 
            Height          =   255
            Left            =   3360
            TabIndex        =   43
            Top             =   4340
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label29 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pasivo Diferido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   151
            Top             =   2880
            Width           =   2535
         End
         Begin VB.Label Label39 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Otros Cap"
            Height          =   255
            Left            =   120
            TabIndex        =   150
            Top             =   4350
            Width           =   975
         End
         Begin VB.Label Label38 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pasivo Circulante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   149
            Top             =   120
            Width           =   2535
         End
         Begin VB.Label Label37 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Proveedor"
            Height          =   255
            Left            =   120
            TabIndex        =   148
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Imp x Pag"
            Height          =   255
            Left            =   120
            TabIndex        =   147
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label35 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Doc x Pag"
            Height          =   255
            Left            =   120
            TabIndex        =   146
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label34 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cob Antic"
            Height          =   255
            Left            =   120
            TabIndex        =   145
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label33 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pasivo Acum"
            Height          =   255
            Left            =   120
            TabIndex        =   144
            Top             =   1640
            Width           =   975
         End
         Begin VB.Label Label32 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pagos  LP"
            Height          =   255
            Left            =   120
            TabIndex        =   143
            Top             =   2250
            Width           =   855
         End
         Begin VB.Label Label31 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Doc LP"
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   2560
            Width           =   855
         End
         Begin VB.Label Label30 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pasivo Fijo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   1920
            Width           =   2535
         End
         Begin VB.Label Label28 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Capital"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   3480
            Width           =   2535
         End
         Begin VB.Label Label27 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Otros Pasivo"
            Height          =   255
            Left            =   120
            TabIndex        =   139
            Top             =   3200
            Width           =   975
         End
         Begin VB.Label Label26 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acciones C."
            Height          =   255
            Left            =   120
            TabIndex        =   138
            Top             =   3720
            Width           =   975
         End
         Begin VB.Label Label25 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Util Acum"
            Height          =   255
            Left            =   120
            TabIndex        =   137
            Top             =   4005
            Width           =   975
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Height          =   4695
         Left            =   -74880
         TabIndex        =   121
         Top             =   360
         Width           =   4215
         Begin VB.TextBox TxtOtrosActivos 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   20
            Text            =   "Otros Activos"
            Top             =   3960
            Width           =   2175
         End
         Begin VB.TextBox TxtPagosAnticipados 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   18
            Text            =   "Pagos Anticipados"
            Top             =   3670
            Width           =   2175
         End
         Begin VB.TextBox TxtPapeleria 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   16
            Text            =   "Papeleria y Utiles de Oficina"
            Top             =   3360
            Width           =   2175
         End
         Begin VB.TextBox TxtDepreciacion 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   14
            Text            =   "Depreciacion Acumulada"
            Top             =   2850
            Width           =   2175
         End
         Begin VB.TextBox TxtEquipoRodante 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   12
            Text            =   "Equipo Rodante"
            Top             =   2540
            Width           =   2175
         End
         Begin VB.TextBox TxtMobiliario 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   10
            Text            =   "Mobiliario y Equipos"
            Top             =   2230
            Width           =   2175
         End
         Begin VB.TextBox TxtTerreno 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   8
            Text            =   "Terreno y Edificio"
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox TxtInventario 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   6
            Text            =   "Inventario"
            Top             =   1290
            Width           =   2175
         End
         Begin VB.TextBox TxtCtasxCobrar 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "Cuentas x Cobrar"
            Top             =   980
            Width           =   2175
         End
         Begin VB.TextBox TxtBanco 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "Banco"
            Top             =   670
            Width           =   2175
         End
         Begin XtremeSuiteControls.CheckBox ChkCaja 
            Height          =   375
            Left            =   3360
            TabIndex        =   1
            Top             =   280
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.TextBox TxtCaja 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   0
            Text            =   "Caja"
            Top             =   360
            Width           =   2175
         End
         Begin XtremeSuiteControls.CheckBox ChkBanco 
            Height          =   255
            Left            =   3360
            TabIndex        =   3
            Top             =   700
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkCtasxCob 
            Height          =   255
            Left            =   3360
            TabIndex        =   5
            Top             =   1050
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkInventario 
            Height          =   255
            Left            =   3360
            TabIndex        =   7
            Top             =   1360
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkTerreno 
            Height          =   255
            Left            =   3360
            TabIndex        =   9
            Top             =   1920
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkMobiliario 
            Height          =   255
            Left            =   3360
            TabIndex        =   11
            Top             =   2270
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkEquipoRodante 
            Height          =   255
            Left            =   3360
            TabIndex        =   13
            Top             =   2600
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkDepreciacionAcum 
            Height          =   255
            Left            =   3360
            TabIndex        =   15
            Top             =   2900
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkPapeleria 
            Height          =   255
            Left            =   3360
            TabIndex        =   17
            Top             =   3360
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkPagosAnticipados 
            Height          =   255
            Left            =   3360
            TabIndex        =   19
            Top             =   3720
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkOtrosActivos 
            Height          =   255
            Left            =   3360
            TabIndex        =   21
            Top             =   4020
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexo"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label24 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Otros Activos"
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   4005
            Width           =   975
         End
         Begin VB.Label Label23 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pago Antic"
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   3720
            Width           =   975
         End
         Begin VB.Label Label22 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Papel y Util"
            Height          =   255
            Left            =   120
            TabIndex        =   133
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Label21 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Activo Diferido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   132
            Top             =   3120
            Width           =   2535
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dep Acum"
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label19 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Activo Fijo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   1600
            Width           =   2535
         End
         Begin VB.Label Label18 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eq Rodante"
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   2560
            Width           =   855
         End
         Begin VB.Label Label17 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mob y Equi"
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   2250
            Width           =   855
         End
         Begin VB.Label Label16 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Terreno y Ed"
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label15 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Inventario"
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ctas x Cob"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Banco"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Caja"
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Activo Circulante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Frame frameBalanza 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1080
         TabIndex        =   115
         Top             =   3840
         Visible         =   0   'False
         Width           =   2535
         Begin VB.OptionButton OptColumna 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Columna"
            Height          =   195
            Left            =   1440
            TabIndex        =   117
            Top             =   200
            Width           =   975
         End
         Begin VB.OptionButton OptTradicional 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tradicional"
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   200
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CheckBox ChkBalanza 
         Caption         =   "Solo Historial Periodo Selecionado"
         Height          =   255
         Left            =   4080
         TabIndex        =   114
         Top             =   4080
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ListBox CmbReportes 
         Height          =   3180
         ItemData        =   "FrmReportes.frx":19CA0
         Left            =   120
         List            =   "FrmReportes.frx":19CA2
         TabIndex        =   113
         Top             =   840
         Width           =   3735
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   4080
         TabIndex        =   108
         Top             =   720
         Width           =   4455
         Begin MSComCtl2.DTPicker DTFecha2 
            Height          =   285
            Left            =   3000
            TabIndex        =   109
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   202178561
            CurrentDate     =   37837
         End
         Begin MSComCtl2.DTPicker DTFecha1 
            Height          =   285
            Left            =   720
            TabIndex        =   110
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   202178561
            CurrentDate     =   37837
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Fin"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   111
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   4080
         TabIndex        =   104
         Top             =   720
         Visible         =   0   'False
         Width           =   4455
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2003"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2004"
            Height          =   255
            Left            =   1680
            TabIndex        =   106
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2005"
            Height          =   255
            Left            =   3120
            TabIndex        =   105
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox TxtTransaccion 
         Height          =   285
         Left            =   8280
         TabIndex        =   103
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   4080
         TabIndex        =   99
         Top             =   1560
         Visible         =   0   'False
         Width           =   4455
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Por Codigo"
            Height          =   255
            Left            =   1440
            TabIndex        =   102
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Por Grupo"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cuentas de Mayor"
            Height          =   255
            Left            =   2640
            TabIndex        =   100
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   1455
         Left            =   3960
         TabIndex        =   88
         Top             =   720
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0E0&
            Height          =   735
            Left            =   2040
            TabIndex        =   92
            Top             =   120
            Width           =   2415
            Begin VB.OptionButton Option6 
               BackColor       =   &H00E0E0E0&
               Caption         =   "2000"
               Height          =   255
               Left            =   1560
               TabIndex        =   95
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton Option7 
               BackColor       =   &H00E0E0E0&
               Caption         =   "2000"
               Height          =   255
               Left            =   840
               TabIndex        =   94
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton Option8 
               BackColor       =   &H00E0E0E0&
               Caption         =   "2000"
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   240
               Value           =   -1  'True
               Width           =   735
            End
         End
         Begin VB.ComboBox CmbIni 
            Height          =   315
            ItemData        =   "FrmReportes.frx":19CA4
            Left            =   1320
            List            =   "FrmReportes.frx":19CCC
            TabIndex        =   91
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox CmbFin 
            Height          =   315
            ItemData        =   "FrmReportes.frx":19CF7
            Left            =   1320
            List            =   "FrmReportes.frx":19D1F
            TabIndex        =   90
            Text            =   "1"
            Top             =   600
            Width           =   615
         End
         Begin VB.ComboBox CmbNivel 
            Height          =   315
            Left            =   1320
            TabIndex        =   89
            Text            =   "3"
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Periodo Hasta:"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Periodo Desde:"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nivel"
            Height          =   255
            Left            =   720
            TabIndex        =   96
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.ComboBox CmbMoneda 
         Height          =   315
         ItemData        =   "FrmReportes.frx":19D4A
         Left            =   6000
         List            =   "FrmReportes.frx":19D54
         TabIndex        =   87
         Text            =   "Córdobas"
         Top             =   3720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clasificado"
         Height          =   1455
         Left            =   3960
         TabIndex        =   77
         Top             =   2160
         Visible         =   0   'False
         Width           =   4575
         Begin VB.TextBox TxtHasta 
            Height          =   285
            Left            =   840
            TabIndex        =   78
            Top             =   600
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.TextBox TxtDesde 
            Height          =   285
            Left            =   840
            TabIndex        =   79
            Top             =   240
            Visible         =   0   'False
            Width           =   3015
         End
         Begin TrueOleDBList80.TDBCombo DBCodigoHasta 
            Height          =   315
            Left            =   840
            TabIndex        =   232
            Top             =   600
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   10583
            _EDITHEIGHT     =   556
            _GAPHEIGHT      =   53
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
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            Appearance      =   1
            BorderStyle     =   1
            ComboStyle      =   0
            AutoCompletion  =   0   'False
            LimitToList     =   0   'False
            ColumnHeaders   =   -1  'True
            ColumnFooters   =   0   'False
            DataMode        =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            Caption         =   ""
            EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            LayoutName      =   ""
            LayoutFileName  =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            AutoSize        =   -1  'True
            ListField       =   ""
            BoundColumn     =   ""
            IntegralHeight  =   0   'False
            CellTipsWidth   =   0
            CellTipsDelay   =   1000
            AutoDropdown    =   0   'False
            RowTracking     =   -1  'True
            RightToLeft     =   0   'False
            RowMember       =   ""
            MouseIcon       =   0
            MouseIcon.vt    =   3
            MousePointer    =   0
            MatchEntryTimeout=   2000
            OLEDragMode     =   0
            OLEDropMode     =   0
            AnimateWindow   =   0
            AnimateWindowDirection=   0
            AnimateWindowTime=   200
            AnimateWindowClose=   0
            DropdownPosition=   0
            Locked          =   0   'False
            ScrollTrack     =   0   'False
            RowDividerColor =   15790320
            RowSubDividerColor=   15790320
            AddItemSeparator=   ";"
            _PropDict       =   $"FrmReportes.frx":19D6B
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
         Begin TrueOleDBList80.TDBCombo DBCodigo 
            Height          =   315
            Left            =   840
            TabIndex        =   231
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   10583
            _EDITHEIGHT     =   556
            _GAPHEIGHT      =   53
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
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            Appearance      =   1
            BorderStyle     =   1
            ComboStyle      =   0
            AutoCompletion  =   0   'False
            LimitToList     =   0   'False
            ColumnHeaders   =   -1  'True
            ColumnFooters   =   0   'False
            DataMode        =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            Caption         =   ""
            EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            LayoutName      =   ""
            LayoutFileName  =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            AutoSize        =   -1  'True
            ListField       =   ""
            BoundColumn     =   ""
            IntegralHeight  =   0   'False
            CellTipsWidth   =   0
            CellTipsDelay   =   1000
            AutoDropdown    =   0   'False
            RowTracking     =   -1  'True
            RightToLeft     =   0   'False
            RowMember       =   ""
            MouseIcon       =   0
            MouseIcon.vt    =   3
            MousePointer    =   0
            MatchEntryTimeout=   2000
            OLEDragMode     =   0
            OLEDropMode     =   0
            AnimateWindow   =   0
            AnimateWindowDirection=   0
            AnimateWindowTime=   200
            AnimateWindowClose=   0
            DropdownPosition=   0
            Locked          =   0   'False
            ScrollTrack     =   0   'False
            RowDividerColor =   15790320
            RowSubDividerColor=   15790320
            AddItemSeparator=   ";"
            _PropDict       =   $"FrmReportes.frx":19E15
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
         Begin VB.CommandButton CmdBuscarEmpleado 
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
            Left            =   3960
            Picture         =   "FrmReportes.frx":19EBF
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   240
            Width           =   375
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
            Left            =   3960
            Picture         =   "FrmReportes.frx":1A00D
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   600
            Width           =   375
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   80
            Top             =   960
            Width           =   2895
            Begin VB.OptionButton Option11 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Detallado"
               Height          =   195
               Left            =   1560
               TabIndex        =   82
               Top             =   200
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton Option10 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Resumen"
               Height          =   195
               Left            =   120
               TabIndex        =   81
               Top             =   200
               Width           =   1095
            End
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   3960
         TabIndex        =   74
         Top             =   2160
         Visible         =   0   'False
         Width           =   2295
         Begin VB.OptionButton OptPeriodo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Periodo"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   200
            Width           =   855
         End
         Begin VB.OptionButton OptAcumulado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acumulado"
            Height          =   195
            Left            =   1080
            TabIndex        =   75
            Top             =   200
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label LblNivel7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nivel Gastos"
         Height          =   255
         Left            =   9480
         TabIndex        =   220
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label LblNivel6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nivel Costos"
         Height          =   255
         Left            =   9480
         TabIndex        =   219
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label LblNivel5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nivel Ingresos"
         Height          =   255
         Left            =   9480
         TabIndex        =   218
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label LblNivel4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nivel Capital"
         Height          =   255
         Left            =   9480
         TabIndex        =   217
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label LblNivel3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nivel Pasivo"
         Height          =   255
         Left            =   9480
         TabIndex        =   216
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label LblNivel2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nivel Activos"
         Height          =   255
         Left            =   4200
         TabIndex        =   209
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione el Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   120
         Top             =   480
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label LblTransaccion 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaccion No."
         Height          =   255
         Left            =   4080
         TabIndex        =   119
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo de Moneda:"
         Height          =   255
         Left            =   4200
         TabIndex        =   118
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   9600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5C1A1&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   -120
      ScaleHeight     =   1095
      ScaleWidth      =   9975
      TabIndex        =   71
      Top             =   -120
      Width           =   9975
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   480
         X2              =   9960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reportes Generales"
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
         Left            =   2280
         TabIndex        =   72
         Top             =   360
         Width           =   4320
      End
   End
   Begin SmartButtonProject.SmartButton CmdVerReporte3 
      Height          =   855
      Left            =   120
      TabIndex        =   70
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Ver Reporte"
      Picture         =   "FrmReportes.frx":1A15B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureLayout   =   7
   End
   Begin VB.TextBox TxtKeyGrupoHasta 
      Height          =   375
      Left            =   6600
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   10200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TxtKeyGrupoDesde 
      Height          =   375
      Left            =   6720
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc AdoHistorial 
      Height          =   375
      Left            =   0
      Top             =   9960
      Width           =   4095
      _ExtentX        =   7223
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
      Caption         =   "AdoHistorial"
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
   Begin SmartButtonProject.SmartButton CmdVerReporte2 
      Height          =   855
      Left            =   120
      TabIndex        =   67
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Ver Reporte"
      Picture         =   "FrmReportes.frx":1B46D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureLayout   =   7
   End
   Begin VB.CheckBox ChkExportar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exportar sin Link"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   66
      Top             =   6600
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc DtaCuentas 
      Height          =   375
      Left            =   0
      Top             =   9960
      Width           =   3975
      _ExtentX        =   7011
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
   Begin MSAdodcLib.Adodc DtaHistorial 
      Height          =   375
      Left            =   4680
      Top             =   9360
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "DtaHistorial"
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
   Begin MSAdodcLib.Adodc DtaConsulta2 
      Height          =   375
      Left            =   3960
      Top             =   9600
      Width           =   3855
      _ExtentX        =   6800
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
      Caption         =   "DtaConsulta2"
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
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   3480
      Top             =   9120
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "DtaConsulta"
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
   Begin MSAdodcLib.Adodc DtaGrupos 
      Height          =   375
      Left            =   120
      Top             =   9960
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "DtaGrupos"
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
   Begin MSAdodcLib.Adodc DtaPeriodos 
      Height          =   375
      Left            =   3120
      Top             =   9120
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "DtaPeriodos"
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
   Begin MSAdodcLib.Adodc DtaBancos 
      Height          =   375
      Left            =   0
      Top             =   9960
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "DtaBancos"
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
   Begin MSAdodcLib.Adodc DtaElimina 
      Height          =   375
      Left            =   480
      Top             =   9960
      Width           =   3855
      _ExtentX        =   6800
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
      Caption         =   "DtaElimina"
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
   Begin MSAdodcLib.Adodc DtaDatosEmpresa 
      Height          =   375
      Left            =   720
      Top             =   9360
      Width           =   3855
      _ExtentX        =   6800
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
      Caption         =   "DtaDatosEmpresa"
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
   Begin MSAdodcLib.Adodc DtaTasas 
      Height          =   375
      Left            =   3480
      Top             =   9000
      Width           =   3735
      _ExtentX        =   6588
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
      Caption         =   "DtaTasas"
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
   Begin MSAdodcLib.Adodc DtaTasas2 
      Height          =   375
      Left            =   5280
      Top             =   9720
      Width           =   3735
      _ExtentX        =   6588
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
      Caption         =   "DtaTasas2"
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
   Begin MSAdodcLib.Adodc DtaReportes 
      Height          =   375
      Left            =   5040
      Top             =   9840
      Width           =   3855
      _ExtentX        =   6800
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
      Caption         =   "DtaReportes"
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
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   8280
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   6600
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin SmartButtonProject.SmartButton CmdSalir 
      Height          =   855
      Left            =   7800
      TabIndex        =   46
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Salir"
      Picture         =   "FrmReportes.frx":1C77F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureLayout   =   7
   End
   Begin MSComDlg.CommonDialog CDRuta 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SmartButtonProject.SmartButton CmdVerReporte 
      Height          =   855
      Left            =   120
      TabIndex        =   44
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Ver Reporte"
      Picture         =   "FrmReportes.frx":1DA91
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureLayout   =   7
   End
   Begin XtremeSuiteControls.ProgressBar osProgress2 
      Height          =   255
      Left            =   4080
      TabIndex        =   153
      Top             =   7320
      Visible         =   0   'False
      Width           =   3615
      _Version        =   786432
      _ExtentX        =   6376
      _ExtentY        =   450
      _StockProps     =   93
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ProgressBar osProgress1 
      Height          =   375
      Left            =   1440
      TabIndex        =   152
      Top             =   6840
      Visible         =   0   'False
      Width           =   6255
      _Version        =   786432
      _ExtentX        =   11033
      _ExtentY        =   661
      _StockProps     =   93
      Scrolling       =   1
      Appearance      =   6
   End
   Begin VB.Label lblProgreso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1320
      TabIndex        =   65
      Top             =   7680
      Width           =   3645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   45
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "FrmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Public CodDesde As String, CodHasta As String
Dim Tape As New clsTape
Dim HayUtiBruta As Boolean
Private WithEvents CmbBeneficiario As TDBCombo
Attribute CmbBeneficiario.VB_VarHelpID = -1
Private LblBeneficiario As Label
Private CodigoBeneficiarioSeleccionado As String

Private Sub EscribirSubtotalGrupo(ByVal Fila As Long, _
                                  ByVal SIni As Double, ByVal SDeb As Double, _
                                  ByVal SCre As Double, ByVal SFin As Double)
    With objExcel.ActiveSheet
        .Range(.Cells(Fila, 2), .Cells(Fila, 6)).Value = _
            Array("TOTAL GRUPO", SIni, SDeb, SCre, SFin)
        .Range(.Cells(Fila, 2), .Cells(Fila, 6)).Font.Bold = True
        .Range(.Cells(Fila, 3), .Cells(Fila, 6)).Borders(8).LineStyle = 1   ' línea superior
    End With
End Sub


Private Sub CmbBeneficiario_Change()

'    MsgBox "Código seleccionado: " & CmbBeneficiario.Text

End Sub

Private Sub CmbBeneficiario_ItemChange()

    If Trim(CmbBeneficiario.Text) = "" Or _
       UCase(Trim(CmbBeneficiario.Text)) = "TODOS" Then

        CodigoBeneficiarioSeleccionado = ""

    Else

        CodigoBeneficiarioSeleccionado = Trim(CmbBeneficiario.Text)

    End If

End Sub

'==========================================================
' AUXILIAR DE CUENTAS
'   Detallado    : True = movimientos, False = totalizado por cuenta
'   AgruparBenef : True = subgrupos por beneficiario dentro de la cuenta
'==========================================================
Private Sub GenerarAuxiliarDeCuentas(ByVal Detallado As Boolean, _
                                     Optional ByVal AgruparBenef As Boolean = False)

    Dim Sql As String
    Dim Benef As String
    Dim FiltroBenefMayor As String
    Dim GrupoDesde As String
    Dim GrupoHasta As String

    Dim rptAux As ArepAuxiliar
    Dim rptMayor As ArepAuxiliarMayor
    Dim fPreview As New FrmPreview

    '==========================================================
    ' VARIABLES GLOBALES QUE USAN LOS REPORTES
    '==========================================================
    NumFecha1 = Me.DTFecha1.Value
    NumFecha2 = Me.DTFecha2.Value
    Moneda = IIf(Me.CmbMoneda.Text = "Córdobas", "Cordobas", "Dolares")
    AuxiliarDetallado = Detallado
    ActivarBeneficiarioReporte = AgruparBenef

    '==========================================================
    ' BENEFICIARIO (vacío o "TODOS" = sin filtro)
    '==========================================================
    Benef = ""
    If ActivarBeneficiario = 1 Or AgruparBenef Then
        Benef = Trim$(CmbBeneficiario.Text)
    End If

    '==========================================================
    ' OPTION 4 - AUXILIAR POR CUENTA
    '==========================================================
    If Me.Option4.Value = True Then

        CodDesde = Me.DBCodigo.Text
        CodHasta = Me.DBCodigoHasta.Text

        Sql = SqlAuxiliarCuentas( _
                Me.DTFecha1.Value, _
                Me.DTFecha2.Value, _
                Me.CmbMoneda.Text, _
                Me.DBCodigo.Text, _
                Me.DBCodigoHasta.Text, _
                Benef, _
                Detallado, _
                AgruparBenef)

        Set rptAux = New ArepAuxiliar
        rptAux.LblRangoFecha.Caption = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
        rptAux.LblFecha.Caption = Format(Now, "dd/mm/yyyy")
        rptAux.DataControl1.ConnectionString = ConexionReporte
        rptAux.DataControl1.Source = Sql

        fPreview.RunReport rptAux
        fPreview.Show 1

    '==========================================================
    ' OPTION 5 - GRUPO / CUENTA DE MAYOR
    '==========================================================
    ElseIf Me.Option5.Value = True Then

        If Me.TxtDesde.Text <> "" Then GrupoDesde = Me.TxtKeyGrupoDesde.Text
        If Me.TxtHasta.Text <> "" Then GrupoHasta = Me.TxtKeyGrupoHasta.Text

        CodDesde = GrupoDesde
        CodHasta = GrupoHasta

        '--- Filtro de beneficiario
        If Benef <> "" And UCase$(Benef) <> "TODOS" Then
            FiltroBenefMayor = "AND Transacciones.Cod_Tercero = '" & Replace(Benef, "'", "''") & "' "
        End If

        Sql = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, "
        Sql = Sql & "Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, "
        Sql = Sql & "Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, "
        Sql = Sql & "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, "
        Sql = Sql & "Transacciones.FechaTransaccion, Cuentas.TipoCuenta, Transacciones.TCambio AS Expr1, "
        Sql = Sql & "Transacciones.TCambio * Transacciones.Debito AS Debito, "
        Sql = Sql & "Transacciones.TCambio * Transacciones.Credito AS Credito, "
        Sql = Sql & "Grupos.KeyGrupo, Grupos.DescripcionGrupo "
        Sql = Sql & "FROM Cuentas "
        Sql = Sql & "INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas "
        Sql = Sql & "INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo "
        Sql = Sql & "WHERE Transacciones.FechaTransaccion BETWEEN '" & Format(Me.DTFecha1.Value, "yyyymmdd") & "' "
        Sql = Sql & "AND '" & Format(Me.DTFecha2.Value, "yyyymmdd") & "' "

        If GrupoDesde <> "" Then Sql = Sql & "AND Grupos.KeyGrupo >= '" & Replace(GrupoDesde, "'", "''") & "' "
        If GrupoHasta <> "" Then Sql = Sql & "AND Grupos.KeyGrupo <= '" & Replace(GrupoHasta, "'", "''") & "' "

        Sql = Sql & FiltroBenefMayor
        Sql = Sql & "ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"

        Set rptMayor = New ArepAuxiliarMayor
        rptMayor.LblRangoFecha.Caption = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
        rptMayor.LblFecha.Caption = Format(Now, "dd/mm/yyyy")
        rptMayor.LblCodigo.Caption = "Mis Code:" & Me.DBCodigo.Text
        rptMayor.LblRango.Caption = "Filtrado Desde: " & _
            IIf(GrupoDesde = "", "Inicio", GrupoDesde) & " Hasta " & IIf(GrupoHasta = "", "Final", GrupoHasta)

        If RutaLogo <> "" Then
            If Dir(RutaLogo) <> "" Then rptMayor.Logo.Picture = LoadPicture(RutaLogo)
        End If

        rptMayor.LblEmpresa.Caption = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
        rptMayor.LblEmpresa1.Caption = Me.DtaDatosEmpresa.Recordset("Direccion")
        rptMayor.LblEmpresa2.Caption = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
        rptMayor.LblFechaImpreso.Caption = Format(Now, "dd/mm/yyyy")
        rptMayor.DataControl1.ConnectionString = ConexionReporte
        rptMayor.DataControl1.Source = Sql

        fPreview.RunReport rptMayor
        fPreview.Show 1

    End If

End Sub

Sub HayUtilidadBruta()
 HayUtiBruta = True
 
 
        With FrmReportes.DtaReportes
            .RecordSource = "select * from reportes where descripcion like '%" & MDIPrimero.AdoConfiguracion.Recordset!OtrosINgresos & "%' order by orden"
            .Refresh
            If .Recordset.RecordCount = 0 Then HayUtiBruta = False
            .RecordSource = "select * from reportes where descripcion like '%" & MDIPrimero.AdoConfiguracion.Recordset!Costos & "%' order by orden"
            .Refresh
            If .Recordset.RecordCount = 0 Then HayUtiBruta = False
            .RecordSource = "select * from reportes where descripcion like '%" & MDIPrimero.AdoConfiguracion.Recordset!Ingresos & "%' order by orden"
            .Refresh
            If .Recordset.RecordCount = 0 Then HayUtiBruta = False
            .RecordSource = "select * from reportes where descripcion like '%" & MDIPrimero.AdoConfiguracion.Recordset!CostosOperativos & "%' order by orden"
            .Refresh
            If .Recordset.RecordCount = 0 Then HayUtiBruta = False
        End With
        
End Sub

Private Sub BtnExcel_Click()

Dim Utilidad As Double, Utilidad2 As Double, Utilidad3 As Double, RegTCostosOper As Integer
Dim Decrementador As Integer, TotalActivoCirculante As Double, TotalActivoFijo As Double, TotalActivoDiferido As Double
Dim TotalPasivoCirculante As Double, TotalPasivoFijo As Double, TotalPasivoDiferido As Double, TotalCapitalSocial As Double
Dim RegInicioCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro donde comienzan los costos operativos
Dim RegTotalCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro de total de costos operativos
Dim Totalingresos As Double, TotalCostoVentas As Double, TotalGastosAdmon As Double, TotalGastos As Double
Dim TotalGastoVentas As Double, TotalIngresosFinancieros As Double, TotalOtrosIngresos As Double, TotalOtrosGastos As Double
Dim TotalUtilidadBruta As Double, TotalImpuestos As Double, TotalUtilidadNeta As Double, Fecha1 As String, Fecha2 As String
Dim TotalCompras As Double, TotalInventarioInicial As Double, TotalInventarioFinal As Double
Dim TotalAcarreo As Double, TotalRebajaVentas As Double, TotalDisponible As Double, TotalGastosR As Double, TotalCosto As Double
Dim TotalSalidas As Double, TotalGastoOperacion As Double, TotalPasivo As Double, TotalCapital As Double
Dim TotalCostos As Double, ListaActivos As Variant, TotalInventario As Double, TotalCuentaxCobrar As Double
Dim TotalCuentasxPagar As Double, TotalActivos As Double, UtilidadBrutas As Double, UtilidadNetas As Double
Dim UltimoOrden As Integer, RegIngresos  As Integer, PrimReg As Integer, UltReg As Integer
Dim Fechas1 As String, Fechas2 As String, Orden As Integer, Sql As String, i As Double
Dim ListaMeses As Variant, CantRegistros As Double, ComboIni As Double, ComboFin As Double, TotstoFijo As Double, TotalGastoFijo As Double
Dim mes As Double, J As Double, TotalDebito As Double, TotalCredito As Double, CodigoCuentaDesde As String, CodigoCuentaHasta As String
Dim rpt As Object, TotalCostoFijo As Double, Ajuste As String, SqlString As String, Cont As Double, KeyGrupo As String
Dim fPreview As New FrmPreview, V As Double, H As Double, DescripcionGrupo As String
Dim Rs As New ADODB.Recordset
Dim Parche As ADODB.Connection


'On Error GoTo TipoErrs
Me.Frame6.Visible = True
Me.CmdVerReporte2.Enabled = False
Me.CmdSalir.Enabled = False
SaldoIni = 0
SaldoFin = 0
Total1 = 0
TotalCuenta = 0

FechaIni = Me.DTFecha1.Value
FechaFin = Me.DTFecha2.Value




Select Case Me.CmbReportes.Text

Case "TOTAL AUXILIAR DE CUENTAS"

    Dim sqlTotAux As String
    Dim rsTotAux As ADODB.Recordset
    Dim filaXls As Long
    Dim grupoActual As String, grupoFila As String
    Dim debAnt As Double, creAnt As Double, debPer As Double, crePer As Double
    Dim saldoIniCta As Double, saldoFinCta As Double
    Dim subIni As Double, subDeb As Double, subCre As Double, subFin As Double
    Dim ctaDesdeTot As String, ctaHastaTot As String
    Dim grpDesdeTot As String, grpHastaTot As String
    Dim benefTot As String
    Const SIN_GRUPO As String = "<<INICIO>>"

    '------------------------------------------------------
    ' Parámetros según la opción
    '------------------------------------------------------
    If Me.Option4.Value = True Then
        ctaDesdeTot = Me.DBCodigo.Text
        ctaHastaTot = Me.DBCodigoHasta.Text
    ElseIf Me.Option5.Value = True Then
        If Me.TxtDesde.Text <> "" Then grpDesdeTot = Me.TxtKeyGrupoDesde.Text
        If Me.TxtHasta.Text <> "" Then grpHastaTot = Me.TxtKeyGrupoHasta.Text
    End If

    If ActivarBeneficiario = 1 Then benefTot = CmbBeneficiario.Text

    '------------------------------------------------------
    ' Consulta (resumida, sin agrupar por beneficiario)
    '------------------------------------------------------
    Me.lblProgreso.Caption = "Consultando saldos..."
    Me.lblProgreso.AutoSize = True
    DoEvents

    sqlTotAux = SqlAuxiliarCuentas( _
                    Me.DTFecha1.Value, Me.DTFecha2.Value, Me.CmbMoneda.Text, _
                    ctaDesdeTot, ctaHastaTot, benefTot, _
                    False, False, _
                    grpDesdeTot, grpHastaTot)

    Set rsTotAux = New ADODB.Recordset
    rsTotAux.CursorLocation = adUseClient
    rsTotAux.Open sqlTotAux, Conexion, adOpenStatic, adLockReadOnly

    If rsTotAux.EOF Then

        Me.lblProgreso.Caption = ""
        MsgBox "No hay movimientos para los filtros seleccionados.", vbInformation, "Zeus Contable"

    Else

        Call Inicio_Excel

        '--------------------------------------------------
        ' Encabezados
        '--------------------------------------------------
        With objExcel.ActiveSheet

            .Cells(1, 1) = NombreEmpresa
            .Range("A1:J1").Merge
            .Range("A1:J1").HorizontalAlignment = xlHAlignCenter
            .Cells(1, 1).Font.Size = 20
            .Cells(1, 1).Font.Bold = True

            .Cells(2, 1) = "RUC " & RUC

            .Cells(3, 1) = "EXPORTACION DESDE " & Me.DTFecha1.Value & " HASTA " & Me.DTFecha2.Value
            .Range("A3:J3").Merge
            .Range("A3:J3").HorizontalAlignment = xlHAlignCenter
            .Cells(3, 1).Font.Size = 14
            .Cells(3, 1).Font.Bold = True

            .Cells(4, 1) = "Moneda Reporte: " & Me.CmbMoneda.Text
            .Cells(4, 3) = "Antes del " & Me.DTFecha1.Value
            .Cells(4, 6) = "Al " & Me.DTFecha2.Value

            .Cells(5, 1) = "Codigo Cuenta"
            .Cells(5, 2) = "Nombre Cuenta"
            .Cells(5, 3) = "Saldo Inicial"
            .Cells(5, 4) = "Movimiento"
            .Cells(5, 6) = "Saldo Final"
            .Range("D5:E5").Merge
            .Range("D5:E5").HorizontalAlignment = xlHAlignCenter
            .Cells(6, 4) = "Debe"
            .Cells(6, 5) = "Haber"
            .Range("D6:E6").HorizontalAlignment = xlHAlignCenter
            .Range("A5:F6").Font.Bold = True

            .Columns("A").ColumnWidth = 24.43
            .Columns("B").ColumnWidth = 40
            .Columns("C").ColumnWidth = 19
            .Columns("D").ColumnWidth = 16
            .Columns("E").ColumnWidth = 16
            .Columns("F").ColumnWidth = 19
            .Columns("A").NumberFormat = "@"
            .Columns("C:F").NumberFormat = "##,##0.00"

        End With

        '--------------------------------------------------
        ' Detalle por cuenta
        '--------------------------------------------------
        Me.lblProgreso.Caption = "Exportando a Excel..."
        Me.osProgress1.Visible = True
        Me.osProgress1.Min = 0
        Me.osProgress1.Max = IIf(rsTotAux.RecordCount > 0, rsTotAux.RecordCount, 1)
        Me.osProgress1.Value = 0

        filaXls = 7
        grupoActual = SIN_GRUPO

        Do While Not rsTotAux.EOF

            grupoFila = rsTotAux("DescripcionGrupo") & ""

            '--- Cambio de grupo: cierro el anterior y abro el nuevo
            If grupoFila <> grupoActual Then

                If grupoActual <> SIN_GRUPO Then
                    EscribirSubtotalGrupo filaXls, subIni, subDeb, subCre, subFin
                    filaXls = filaXls + 2
                End If

                objExcel.ActiveSheet.Cells(filaXls, 1) = grupoFila
                objExcel.ActiveSheet.Cells(filaXls, 1).Font.Bold = True
                filaXls = filaXls + 1

                subIni = 0: subDeb = 0: subCre = 0: subFin = 0
                grupoActual = grupoFila

            End If

            '--- Saldos de la cuenta
            debAnt = NzD(rsTotAux("DebitoAnterior"))
            creAnt = NzD(rsTotAux("CreditoAnterior"))
            debPer = NzD(rsTotAux("Debito"))
            crePer = NzD(rsTotAux("Credito"))

            If EsCuentaDeudora(rsTotAux("TipoCuenta") & "") Then
                saldoIniCta = debAnt - creAnt
                saldoFinCta = saldoIniCta + debPer - crePer
            Else
                saldoIniCta = creAnt - debAnt
                saldoFinCta = saldoIniCta + crePer - debPer
            End If

            '--- Una sola escritura por fila (más rápido que celda por celda)
            objExcel.ActiveSheet.Range( _
                objExcel.ActiveSheet.Cells(filaXls, 1), _
                objExcel.ActiveSheet.Cells(filaXls, 6)).Value = _
                Array(rsTotAux("CodCuentas") & "", _
                      rsTotAux("DescripcionCuentas") & "", _
                      saldoIniCta, debPer, crePer, saldoFinCta)

            subIni = subIni + saldoIniCta
            subDeb = subDeb + debPer
            subCre = subCre + crePer
            subFin = subFin + saldoFinCta

            filaXls = filaXls + 1
            Me.osProgress1.Value = Me.osProgress1.Value + 1
            DoEvents

            rsTotAux.MoveNext
        Loop

        '--- Subtotal del último grupo
        EscribirSubtotalGrupo filaXls, subIni, subDeb, subCre, subFin

        Me.lblProgreso.Caption = ""
        Me.osProgress1.Visible = False
        MsgBox "Proceso Terminado!!!", vbExclamation, "Zeus Contable"

    End If

    rsTotAux.Close
    Set rsTotAux = Nothing



        Case "REPORTE PBI"
        
        
    
        
    NumFecha1 = Me.DTFecha1.Value
    NumFecha2 = Me.DTFecha2.Value
    

      
      If Me.Option4.Value = True Then
      
      
            If Me.DBCodigo.Text = "" Then
                 Me.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
                 Me.DtaConsulta.Refresh
                 If Not Me.DtaConsulta.Recordset.EOF Then
                   Me.DtaConsulta.Recordset.MoveFirst
                   CodigoCuentaDesde = Me.DtaConsulta.Recordset("CodCuentas")
                End If
            Else
              CodigoCuentaDesde = Me.DBCodigo.Text
            End If
            
            If Me.DBCodigoHasta.Text = "" Then
                 Me.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
                 Me.DtaConsulta.Refresh
                 If Not Me.DtaConsulta.Recordset.EOF Then
                   Me.DtaConsulta.Recordset.MoveLast
                   CodigoCuentaHasta = Me.DtaConsulta.Recordset("CodCuentas")
                End If
            Else
               CodigoCuentaHasta = Me.DBCodigoHasta.Text
            End If
            
            
            CodDesde = CodigoCuentaDesde
            CodHasta = CodigoCuentaHasta
            
             If Me.CmbMoneda.Text = "Córdobas" Then
             
               Moneda = "Cordobas"
'               Sql = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Transacciones.NPeriodo) AS NPeriodo, MAX(Transacciones.NTransaccion) AS NTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.VoucherNo) AS VoucherNo, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.Clave) AS Clave, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END) AS Debito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END) AS Credito, MAX(Transacciones.FacturaNo) AS FacturaNo, MAX(Transacciones.ChequeNo) AS ChequeNo, MAX(Transacciones.Fuente) AS Fuente, MAX(Cuentas.TipoCuenta) AS TipoCuenta, SUM(Transacciones.Debito + Transacciones.Credito) AS Saldo FROM  Cuentas INNER JOIN " & _
'                     "Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas  " & _
'                     "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) AND (IndiceTransaccion.Ajuste <> 'Dólares') GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (SUM(Transacciones.Debito + Transacciones.Credito) <> 0) ORDER BY Transacciones.CodCuentas"
             
               Sql = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END AS Debito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END AS Credito, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, Cuentas.TipoCuenta, Transacciones.Debito + Transacciones.Credito AS Saldo FROM  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND  " & _
                     "Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas  " & _
                     "WHERE  (Transacciones.FechaTransaccion Between CONVERT(DATETIME, '" & Format(Me.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) AND (IndiceTransaccion.Ajuste <> 'Dólares') AND (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (Transacciones.Debito + Transacciones.Credito <> 0) ORDER BY Transacciones.CodCuentas"
             
             
             Else
               Moneda = "Dolares"
'               Sql = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Transacciones.NPeriodo) AS NPeriodo, MAX(Transacciones.NTransaccion) AS NTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.VoucherNo) AS VoucherNo, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.Clave) AS Clave, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END) AS Debito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END) AS Credito, MAX(Transacciones.FacturaNo) AS FacturaNo, MAX(Transacciones.ChequeNo) AS ChequeNo, MAX(Transacciones.Fuente) AS Fuente, MAX(Cuentas.TipoCuenta) AS TipoCuenta, SUM(Transacciones.Debito + Transacciones.Credito) AS Saldo FROM  Cuentas INNER JOIN " & _
'                     "Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas  " & _
                     "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) AND (IndiceTransaccion.Ajuste <> 'Córdobas') GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (SUM(Transacciones.Debito + Transacciones.Credito) <> 0) ORDER BY Transacciones.CodCuentas"
            
             
              Sql = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END AS Debito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END AS Credito, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, Cuentas.TipoCuenta, Transacciones.Debito + Transacciones.Credito AS Saldo FROM  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND  " & _
                     "Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas  " & _
                     "WHERE  (Transacciones.FechaTransaccion Between CONVERT(DATETIME, '" & Format(Me.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) AND (IndiceTransaccion.Ajuste <> 'Córdobas') AND (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (Transacciones.Debito + Transacciones.Credito <> 0) ORDER BY Transacciones.CodCuentas"
             
             End If
            



     
       ElseIf Me.Option5.Value = True Then
       
      
            If Me.TxtDesde.Text = "" Then
               Me.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
               Me.DtaConsulta.Refresh
               If Not Me.DtaConsulta.Recordset.EOF Then
                 Me.DtaConsulta.Recordset.MoveFirst
                 CodigoCuentaDesde = Me.DtaConsulta.Recordset("KeyGrupo")
               End If
            Else
                CodigoCuentaDesde = Me.TxtKeyGrupoDesde.Text
            End If
               
            If Me.TxtHasta.Text = "" Then
               Me.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
               Me.DtaConsulta.Refresh
               If Not Me.DtaConsulta.Recordset.EOF Then
                 Me.DtaConsulta.Recordset.MoveLast
                 CodigoCuentaHasta = Me.DtaConsulta.Recordset("KeyGrupo")
               End If
            Else
               CodigoCuentaHasta = Me.TxtKeyGrupoHasta.Text
            End If
       
             Sql = "SELECT  Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.FacturaNo,Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Cuentas.TipoCuenta, Transacciones.TCambio AS Expr1, Transacciones.TCambio * Transacciones.Debito AS Debito, Transacciones.TCambio * Transacciones.Credito AS Credito, Grupos.KeyGrupo, Grupos.DescripcionGrupo " & _
                    "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo " & _
                    "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(Me.DTFecha1, "yyyymmdd") & "' AND '" & Format(Me.DTFecha2, "yyyymmdd") & "') AND (Grupos.KeyGrupo BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') " & _
                    "ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion "
       
'             ArepAuxiliarMayor.DataControl1.Source = Sql

       
       End If
     
                Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel

            V = 5
            H = 0
            i = 1

          '///////////////////////////////////////////////////////////////////////////////////////
          '////////////////////ENCABEZADOS//////////////////////////////////////////////////////
          '///////////////////////////////////////////////////////////////////////////////////
            objExcel.ActiveSheet.Cells(1, 1) = NombreEmpresa
            objExcel.ActiveSheet.Range("A1:J1").Merge
            objExcel.ActiveSheet.Range("A1", "J1").HorizontalAlignment = xlHAlignCenter
            With objExcel.ActiveSheet.Cells(1, 1)
                  .Font.Size = 20             ' tamaño de letra
                  .Font.Bold = True           ' Fuente en negrita
            End With
            objExcel.ActiveSheet.Cells(2, 1) = "RUC " & RUC
            objExcel.ActiveSheet.Cells(3, 1) = "EXPORTACION DESDE " & Me.DTFecha1.Value & " HASTA " & Me.DTFecha1.Value
            objExcel.ActiveSheet.Range("A3:J3").Merge
            objExcel.ActiveSheet.Range("A3", "J3").HorizontalAlignment = xlHAlignCenter
            With objExcel.ActiveSheet.Cells(3, 1)
                  .Font.Size = 14             ' tamaño de letra
                  .Font.Bold = True           ' Fuente en negrita
            End With


            objExcel.ActiveSheet.Columns("A").ColumnWidth = 11
            objExcel.ActiveSheet.Columns("B").ColumnWidth = 11
            objExcel.ActiveSheet.Columns("C").ColumnWidth = 11
            objExcel.ActiveSheet.Columns("D").ColumnWidth = 11
            objExcel.ActiveSheet.Columns("E").ColumnWidth = 11
            objExcel.ActiveSheet.Columns("F").ColumnWidth = 16
            objExcel.ActiveSheet.Columns("G").ColumnWidth = 16
             objExcel.ActiveSheet.Columns("J").ColumnWidth = 16
             objExcel.ActiveSheet.Columns("K").ColumnWidth = 42
             objExcel.ActiveSheet.Columns("L").ColumnWidth = 17
             objExcel.ActiveSheet.Columns("M").ColumnWidth = 14
             objExcel.ActiveSheet.Columns("N").ColumnWidth = 93
            
            objExcel.ActiveSheet.Cells(4, 1) = "Numero Registro: "
            objExcel.ActiveSheet.Cells(4, 2) = "Fecha Registro "
            objExcel.ActiveSheet.Cells(4, 3) = "Año"
            objExcel.ActiveSheet.Cells(4, 4) = "Moneda"
            objExcel.ActiveSheet.Cells(4, 5) = "Tipo Cambio"
            objExcel.ActiveSheet.Cells(4, 6) = "Tipo Movimiento"
            objExcel.ActiveSheet.Cells(4, 7) = "DEBE"
            objExcel.ActiveSheet.Cells(4, 8) = "HABER"
            objExcel.ActiveSheet.Cells(4, 9) = "SALDO"
            objExcel.ActiveSheet.Cells(4, 10) = "COD_CUENTA"
            objExcel.ActiveSheet.Cells(4, 11) = "DESCRIPCION DE LA CUENTA"
            objExcel.ActiveSheet.Cells(4, 12) = "TIPO CUENTA"
            objExcel.ActiveSheet.Cells(4, 13) = "DESCRIPCION CLIENTE O PROVEEDOR"
            objExcel.ActiveSheet.Cells(4, 14) = "CAUSAL DE REGISTRO"
            objExcel.ActiveSheet.Cells(4, 15) = "FacturaNo"
            objExcel.ActiveSheet.Cells(4, 16) = "ChequeNo"
            objExcel.ActiveSheet.Cells(4, 17) = "NumeroMovimiento"


            objExcel.ActiveSheet.Columns("D").NumberFormat = "##,##0.00"
            objExcel.ActiveSheet.Columns("E").NumberFormat = "##,##0.00"
            objExcel.ActiveSheet.Columns("F").NumberFormat = "##,##0.00"
            objExcel.ActiveSheet.Columns("G").NumberFormat = "##,##0.00"
            
            objExcel.ActiveSheet.Cells(4, 1).Font.Size = 10             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 1).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 2).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 2).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 3).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 3).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 4).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 4).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 5).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 5).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 6).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 6).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 7).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 7).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 8).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 8).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 9).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 9).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 10).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 10).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 11).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 11).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 12).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 12).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 13).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 13).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 14).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 14).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 15).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 15).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 16).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 16).Font.Bold = True           ' Fuente en negrita
            objExcel.ActiveSheet.Cells(4, 17).Font.Size = 11             ' tamaño de letra
            objExcel.ActiveSheet.Cells(4, 17).Font.Bold = True           ' Fuente en negrita
            
            With objExcel.ActiveSheet.Cells(3, 1)
                  .Font.Size = 14             ' tamaño de letra
                  .Font.Bold = True           ' Fuente en negrita
            End With
            
            Me.DtaConsulta.RecordSource = Sql
            Me.DtaConsulta.Refresh
            
            
            Me.osProgress1.Min = 0
            Me.DtaConsulta.Recordset.MoveLast
            Me.osProgress1.Max = Me.DtaConsulta.Recordset.RecordCount
            Me.DtaConsulta.Recordset.MoveFirst
            Me.osProgress1.Value = 0
            Do While Not Me.DtaConsulta.Recordset.EOF
                DoEvents
                objExcel.ActiveSheet.Cells(V, 1) = Me.DtaConsulta.Recordset("NTransaccion")
                objExcel.ActiveSheet.Cells(V, 2) = Format(Me.DtaConsulta.Recordset("FechaTransaccion"), "dd/mm/yyyy")
                objExcel.ActiveSheet.Cells(V, 3) = Year(Me.DtaConsulta.Recordset("FechaTransaccion"))
                objExcel.ActiveSheet.Cells(V, 4) = Me.CmbMoneda.Text
                objExcel.ActiveSheet.Cells(V, 5) = BuscaTasaCambio(Me.DtaConsulta.Recordset("FechaTransaccion"))
                objExcel.ActiveSheet.Cells(V, 6) = Me.DtaConsulta.Recordset("Clave")
                objExcel.ActiveSheet.Cells(V, 7) = Me.DtaConsulta.Recordset("Debito")
                objExcel.ActiveSheet.Cells(V, 8) = Me.DtaConsulta.Recordset("Credito")
                objExcel.ActiveSheet.Cells(V, 9) = Me.DtaConsulta.Recordset("Saldo")
                objExcel.ActiveSheet.Cells(V, 10) = Me.DtaConsulta.Recordset("CodCuentas")
                objExcel.ActiveSheet.Cells(V, 11) = Me.DtaConsulta.Recordset("DescripcionCuentas")
                objExcel.ActiveSheet.Cells(V, 12) = Me.DtaConsulta.Recordset("TipoCuenta")
                objExcel.ActiveSheet.Cells(V, 13) = Me.DtaConsulta.Recordset("Fuente")
                objExcel.ActiveSheet.Cells(V, 14) = Me.DtaConsulta.Recordset("DescripcionMovimiento")
                objExcel.ActiveSheet.Cells(V, 15) = Me.DtaConsulta.Recordset("FacturaNo")
                objExcel.ActiveSheet.Cells(V, 16) = Me.DtaConsulta.Recordset("ChequeNo")
                objExcel.ActiveSheet.Cells(V, 17) = Me.DtaConsulta.Recordset("NumeroMovimiento")
                
                
                
                DoEvents
                V = V + 1
                Me.osProgress1.Value = Me.osProgress1.Value + 1
                Me.DtaConsulta.Recordset.MoveNext
            Loop
            
            
         MsgBox "Proceso Terminado!!!", vbExclamation, "Zeus Contable"
 
        Case "BALANZA DE COMPROBACION"
        
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////ESTA OPCION CALCULA LA BALANZA POR CODIGO //////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
         
         If Me.CmbMoneda.Text = "Córdobas" Then
            Moneda = "Cordobas"
         Else
            Moneda = "Dolares"
         End If
         
         
           
            
         If Option4.Value = True Then
            Me.DtaReportes.Refresh
            Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
            Me.lblProgreso.AutoSize = True
            Me.osProgress1.Visible = True
            Me.osProgress1.Value = 0
            Me.osProgress1.Min = 0
            Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
            Do While Not Me.DtaReportes.Recordset.EOF
             Me.DtaReportes.Recordset.Delete
                Me.DtaReportes.Recordset.MoveNext
                Me.osProgress1.Value = Me.osProgress1.Value + 1
            Loop
            
            Me.lblProgreso.Caption = ""
            Me.osProgress1.Visible = False
            SaldoReportes ("BalanzaCodigo")
            Me.DtaReportes.Refresh
            'aqui voy a poner el parche para que salga los movimientos de debito y credito en periodo

            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open
        
            Parche.Execute "update reportes set debe2=0,haber2=0"
            Parche.Execute "update reportes " & _
                "set debe2=mdebito,haber2=mcredito " & _
                "from (" & ConsultaTotalesMovimientos & ") as Temporal " & _
                "where temporal.codcuentas=reportes.codcuentas"
                
            Parche.Execute "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2"
                
            
            Parche.Execute "Update [Reportes] SET [Debe3] =  CASE WHEN (Debe1 + Debe2) > (Haber1 + Haber2) THEN (Debe1 + Debe2) - (Haber1 + Haber2) ELSE 0 END ,[Haber3] =  CASE WHEN (Haber1 + Haber2) > (Debe1 + Debe2) THEN (Haber1 + Haber2) - (Debe1 + Debe2) ELSE 0 END"
               
               
 
            
            
            
            If Me.OptTradicional.Value = True Then
                ArepBalanza.LblMoneda.Caption = Me.CmbMoneda.Text
                
                ArepBalanza.Logo.Picture = LoadPicture(RutaLogo)
                
                ArepBalanza.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                ArepBalanza.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                ArepBalanza.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                ArepBalanza.LblFechaFin = Me.DTFecha2.Value
                ArepBalanza.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
                ArepBalanza.LblFechaIni = Me.DTFecha1.Value
                ArepBalanza.DataControl1.ConnectionString = ConexionReporte
                ArepBalanza.DataControl1.Source = "SELECT * From Reportes ORDER BY CodCuentas"
                ArepBalanza.DataControl1.Refresh
                
                QUIEN = "BalanzaCodigo"
                ArepBalanza.Show 1
        '         fPreview.arv.ReportSource = ArepBalanza
        '         fPreview.Show 1
            End If
            Me.lblProgreso.Caption = ""
            Me.osProgress1.Visible = False
            
            
            
         ElseIf Option5.Value = True Then
           '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
           '//////////////////////EN ESTA OPCION CALCULO LA BALANZA POR GRUPOS//////////////////////////////////
           '////////////////////////////////////////////////////////////////////////////////////////////////////////////////
         
         
        
         
            Me.DtaReportes.Refresh
            
            Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
            Me.lblProgreso.AutoSize = True
            Me.osProgress1.Value = 0
            Me.osProgress1.Min = 0
            Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
            
            
            
            Do While Not Me.DtaReportes.Recordset.EOF
             Me.osProgress1.Value = Me.osProgress1.Value + 1
             Me.DtaReportes.Recordset.Delete
             Me.DtaReportes.Recordset.MoveNext
            Loop
            CreaEstructura ("Balanza")
            SaldoReportes ("Balanza")
            EliminaRegistroCero ("Balanza")
            
            '-----------------------BORRO TODAS LAS CUENTAS QUE NO SUMAN NINGUN VALOR ------------------
            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open
            Parche.Execute "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)"
            
            

            
            
            Me.DtaConsulta.RecordSource = "select sum(Debe1) as SumaDebe1,sum(debe2) as SumaDebe2,sum(debe3)as SumaDebe3,sum(haber1) as SumaHaber1,sum(haber2) as SumaHaber2,sum(haber3) as SumaHaber3 from reportes where descripcion not like '%total%'"
            Me.DtaConsulta.Refresh
            If Not DtaConsulta.Recordset.EOF Then
''        '      ArepBalanza.LblDebe1 = Format(Me.DtaConsulta.Recordset("SumaDebe1"), "##,##0.00")
''        '      ArepBalanza.LblHaber1 = Format(Me.DtaConsulta.Recordset("SumaHaber1"), "##,##0.00")
''        '      ArepBalanza.LblDebe2 = Format(Me.DtaConsulta.Recordset("SumaDebe2"), "##,##0.00")
''        '      ArepBalanza.LblHaber2 = Format(Me.DtaConsulta.Recordset("SumaHaber2"), "##,##0.00")
''        '      ArepBalanza.LblDebe3 = Format(Me.DtaConsulta.Recordset("SumaDebe3"), "##,##0.00")
''        '      ArepBalanza.LblHaber3 = Format(Me.DtaConsulta.Recordset("SumaHaber3"), "##,##0.00")
''              ArepBalanza.LblDebe3 = Me.DtaConsulta.Recordset("SumaDebe3")
''              ArepBalanza.LblHaber3 = Me.DtaConsulta.Recordset("SumaHaber3")
            
            End If
            
            
            Parche.Execute "Update [Reportes] SET [Debe3] =  CASE WHEN (Debe1 + Debe2) > (Haber1 + Haber2) THEN (Debe1 + Debe2) - (Haber1 + Haber2) ELSE 0 END ,[Haber3] =  CASE WHEN (Haber1 + Haber2) > (Debe1 + Debe2) THEN (Haber1 + Haber2) - (Debe1 + Debe2) ELSE 0 END"
            
            Dim Resultado As String
              If Me.OptTradicional.Value = True Then
                Me.DtaReportes.Refresh

                Sql = "SELECT * From Reportes ORDER BY Orden"
                QUIEN = "Balanza"
        '        ArepBalanza.Show 1
        
        
            Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel

            V = 6
            H = 0
            i = 1

          '///////////////////////////////////////////////////////////////////////////////////////
          '////////////////////ENCABEZADOS//////////////////////////////////////////////////////
          '///////////////////////////////////////////////////////////////////////////////////
            objExcel.ActiveSheet.Cells(1, 1) = NombreEmpresa
            objExcel.ActiveSheet.Range("A1:J1").Merge
            objExcel.ActiveSheet.Range("A1", "J1").HorizontalAlignment = xlHAlignCenter
            With objExcel.ActiveSheet.Cells(1, 1)
                  .Font.Size = 20             ' tamaño de letra
                  .Font.Bold = True           ' Fuente en negrita
            End With
            objExcel.ActiveSheet.Cells(2, 1) = "RUC " & RUC
            objExcel.ActiveSheet.Cells(3, 1) = "EXPORTACION DESDE " & Me.DTFecha1.Value & " HASTA " & Me.DTFecha1.Value
            objExcel.ActiveSheet.Range("A3:J3").Merge
            objExcel.ActiveSheet.Range("A3", "J3").HorizontalAlignment = xlHAlignCenter
            With objExcel.ActiveSheet.Cells(3, 1)
                  .Font.Size = 14             ' tamaño de letra
                  .Font.Bold = True           ' Fuente en negrita
            End With

            objExcel.ActiveSheet.Cells(4, 1) = "Moneda Reporte: " & Me.CmbMoneda.Text
            objExcel.ActiveSheet.Cells(4, 2) = "Saldo Antes del " & Me.DTFecha1.Value
            objExcel.ActiveSheet.Cells(4, 6) = "Saldo Al " & Me.DTFecha2.Value
            objExcel.ActiveSheet.Cells(5, 1) = "Codigo - Descripcion"
            objExcel.ActiveSheet.Columns("A").ColumnWidth = 60
            objExcel.ActiveSheet.Columns("B").ColumnWidth = 16
            objExcel.ActiveSheet.Columns("C").ColumnWidth = 16
            objExcel.ActiveSheet.Columns("D").ColumnWidth = 16
            objExcel.ActiveSheet.Columns("E").ColumnWidth = 16
            objExcel.ActiveSheet.Columns("F").ColumnWidth = 16
            objExcel.ActiveSheet.Columns("G").ColumnWidth = 16
            
            objExcel.ActiveSheet.Cells(5, 2) = "Debe"
            objExcel.ActiveSheet.Cells(5, 3) = "Haber"
            objExcel.ActiveSheet.Cells(5, 4) = "Debe"
            objExcel.ActiveSheet.Cells(5, 5) = "Haber"
            objExcel.ActiveSheet.Cells(5, 6) = "Debe"
            objExcel.ActiveSheet.Cells(5, 7) = "Haber"
            objExcel.ActiveSheet.Columns("B").NumberFormat = "##,##0.00"
            objExcel.ActiveSheet.Columns("C").NumberFormat = "##,##0.00"
            objExcel.ActiveSheet.Columns("D").NumberFormat = "##,##0.00"
            objExcel.ActiveSheet.Columns("E").NumberFormat = "##,##0.00"
            objExcel.ActiveSheet.Columns("F").NumberFormat = "##,##0.00"
            objExcel.ActiveSheet.Columns("G").NumberFormat = "##,##0.00"
            
            MDIPrimero.AdoConsulta.RecordSource = Sql
            MDIPrimero.AdoConsulta.Refresh
            
            Me.osProgress1.Min = 0
            MDIPrimero.AdoConsulta.Recordset.MoveLast
            Me.osProgress1.Max = MDIPrimero.AdoConsulta.Recordset.RecordCount
            MDIPrimero.AdoConsulta.Recordset.MoveFirst
            Me.osProgress1.Value = 0
            Do While Not MDIPrimero.AdoConsulta.Recordset.EOF
                
                objExcel.ActiveSheet.Cells(V, 1) = MDIPrimero.AdoConsulta.Recordset("Descripcion")
                objExcel.ActiveSheet.Cells(V, 2) = MDIPrimero.AdoConsulta.Recordset("Debe1")
                objExcel.ActiveSheet.Cells(V, 3) = MDIPrimero.AdoConsulta.Recordset("Haber1")
                objExcel.ActiveSheet.Cells(V, 4) = MDIPrimero.AdoConsulta.Recordset("Debe2")
                objExcel.ActiveSheet.Cells(V, 5) = MDIPrimero.AdoConsulta.Recordset("Haber2")
                objExcel.ActiveSheet.Cells(V, 6) = MDIPrimero.AdoConsulta.Recordset("Debe2")
                objExcel.ActiveSheet.Cells(V, 7) = MDIPrimero.AdoConsulta.Recordset("Haber2")
            
                V = V + 1
                Me.osProgress1.Value = Me.osProgress1.Value + 1
                MDIPrimero.AdoConsulta.Recordset.MoveNext
            Loop
            
            '////////////////////////////////////////////////////////////////////
            '/////////////TOTALIZO LA BALANZA /////////////////
            '////////////////////////////////////////////////////////////////
             FrmReportes.DtaConsulta.RecordSource = "SELECT SUM(Debe1) AS Debe1, SUM(Haber1) AS Haber1, SUM(Debe2) AS Debe2, SUM(Haber2) AS Haber2, SUM(Debe3) AS Debe3, SUM(Haber3) AS Haber3 From Reportes WHERE (Nivel = 1) AND (Descripcion LIKE N'%Total%')"
             FrmReportes.DtaConsulta.Refresh
      
            If Not FrmReportes.DtaConsulta.Recordset.EOF Then
              V = V + 1

                 objExcel.ActiveSheet.Cells(V, 1) = "Totales"

               If Not IsNull(FrmReportes.DtaConsulta.Recordset("Debe1")) Then
                 objExcel.ActiveSheet.Cells(V, 2) = FrmReportes.DtaConsulta.Recordset("Debe1")
               End If
               
               If Not IsNull(FrmReportes.DtaConsulta.Recordset("Haber1")) Then
                 objExcel.ActiveSheet.Cells(V, 3) = FrmReportes.DtaConsulta.Recordset("Haber1")
               End If
               
               If Not IsNull(FrmReportes.DtaConsulta.Recordset("Debe2")) Then
                 objExcel.ActiveSheet.Cells(V, 4) = FrmReportes.DtaConsulta.Recordset("Debe2")
               End If
               
               If Not IsNull(FrmReportes.DtaConsulta.Recordset("Haber2")) Then
                 objExcel.ActiveSheet.Cells(V, 5) = FrmReportes.DtaConsulta.Recordset("Haber2")
               End If
               
               If Not IsNull(FrmReportes.DtaConsulta.Recordset("Debe3")) Then
                 objExcel.ActiveSheet.Cells(V, 6) = FrmReportes.DtaConsulta.Recordset("Debe3")
               End If
               
               If Not IsNull(FrmReportes.DtaConsulta.Recordset("Haber3")) Then
                objExcel.ActiveSheet.Cells(V, 7) = FrmReportes.DtaConsulta.Recordset("Haber3")
               End If
               
             objExcel.ActiveSheet.Cells(V, 1).Font.Size = 10             ' tamaño de letra
             objExcel.ActiveSheet.Cells(V, 1).Font.Bold = True           ' Fuente en negrita
             objExcel.ActiveSheet.Cells(V, 2).Font.Size = 10             ' tamaño de letra
             objExcel.ActiveSheet.Cells(V, 2).Font.Bold = True
             objExcel.ActiveSheet.Cells(V, 3).Font.Size = 10             ' tamaño de letra
             objExcel.ActiveSheet.Cells(V, 3).Font.Bold = True
             objExcel.ActiveSheet.Cells(V, 4).Font.Size = 10             ' tamaño de letra
             objExcel.ActiveSheet.Cells(V, 4).Font.Bold = True
             objExcel.ActiveSheet.Cells(V, 5).Font.Size = 10             ' tamaño de letra
             objExcel.ActiveSheet.Cells(V, 5).Font.Bold = True
             objExcel.ActiveSheet.Cells(V, 6).Font.Size = 10             ' tamaño de letra
             objExcel.ActiveSheet.Cells(V, 6).Font.Bold = True
             objExcel.ActiveSheet.Cells(V, 7).Font.Size = 10             ' tamaño de letra
             objExcel.ActiveSheet.Cells(V, 7).Font.Bold = True
             
             
            End If
            
        
'
'                    Set rpt = New ArepBalanza
'                    rpt.DataControl1.ConnectionString = ConexionReporte
'                    rpt.DataControl1.Source = Sql
'                    fPreview.RunReport rpt
'                    fPreview.Show 1
                    
        
              ElseIf Me.OptColumna.Value = True Then
                Me.DtaReportes.Refresh
                ArepBalanzaColumna.LblMoneda.Caption = Me.CmbMoneda.Text
                ArepBalanzaColumna.Logo.Picture = LoadPicture(RutaLogo)
                ArepBalanzaColumna.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                ArepBalanzaColumna.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                ArepBalanzaColumna.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                ArepBalanzaColumna.LblFechaFin = Me.DTFecha2.Value
                ArepBalanzaColumna.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
                ArepBalanzaColumna.LblFechaIni = Me.DTFecha1.Value
                ArepBalanzaColumna.DataControl1.ConnectionString = ConexionReporte
                ArepBalanzaColumna.DataControl1.Source = "SELECT * From Reportes ORDER BY Orden"
                ArepBalanzaColumna.Field8.Visible = False
                ArepBalanzaColumna.Field9.Visible = False
                ArepBalanzaColumna.Field10.Visible = False
                ArepBalanzaColumna.Field11.Visible = False
                ArepBalanzaColumna.FldTDebe3.Visible = False
                ArepBalanzaColumna.FldTHaber3.Visible = False
                ArepBalanzaColumna.Show 1
             
              End If
                 
             Me.lblProgreso.Caption = ""
            Me.osProgress1.Visible = False
            
         ElseIf Option9.Value = True Then
         
            '/////////////////////////////////////////////////////////////////////////////////////////
            '//////////////EN ESTA OPCION CALCULO POR CUENTA DE MAYOR/////////////////////////////////////
            '//////////////////////////////////////////////////////////////////////////////////////////
            
        Parche.Execute "Update [Reportes] SET [Debe3] =  CASE WHEN (Debe1 + Debe2) > (Haber1 + Haber2) THEN (Debe1 + Debe2) - (Haber1 + Haber2) ELSE 0 END ,[Haber3] =  CASE WHEN (Haber1 + Haber2) > (Debe1 + Debe2) THEN (Haber1 + Haber2) - (Debe1 + Debe2) ELSE 0 END"
         
        '     SQL = "SELECT Max(Transacciones.CodCuentas) AS CodCuentas, Max(Transacciones.NumeroMovimiento) AS NumeroMovimiento, Max(Transacciones.NombreCuenta) AS NombreCuenta, Avg(Transacciones.TCambio) AS TCambio, Sum(Transacciones.TCambio*Transacciones.Debito) AS Debito, Sum(Transacciones.TCambio*Transacciones.Credito) AS Credito, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta " & _
        '    "FROM Grupos INNER JOIN (Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas) ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
        '    "WHERE     (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102)) " & _
        '    "GROUP BY Cuentas.KeyGrupo, Cuentas.DescripcionGrupo " & _
        '    "ORDER BY Cuentas.KeyGrupo "
        '    "WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) " & _

            Sql = "SELECT MAX(Transacciones.CodCuentas) AS CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, SUM(Transacciones.Debito) AS Debito,SUM(Transacciones.Credito) AS Credito, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas GROUP BY Cuentas.KeyGrupo, Cuentas.DescripcionGrupo " & _
                  "HAVING (MAX(Transacciones.FechaTransaccion) <= CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) ORDER BY Cuentas.KeyGrupo"
        
        
            If Me.OptTradicional.Value = True Then
                ArepBalanzaMayor.DataControl1.ConnectionString = ConexionReporte
                ArepBalanzaMayor.DataControl1.Source = Sql
                ArepBalanzaMayor.Logo.Picture = LoadPicture(RutaLogo)
                ArepBalanzaMayor.LblEmpresa.Caption = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                ArepBalanzaMayor.LblEmpresa1.Caption = Me.DtaDatosEmpresa.Recordset("Direccion")
                ArepBalanzaMayor.LblEmpresa2.Caption = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                ArepBalanzaMayor.LblFechaIni.Caption = FechaIni
                ArepBalanzaMayor.LblFechaFin.Caption = FechaFin
                ArepBalanzaMayor.LblFechaImpreso.Caption = "Impreso: " & Format(Now, "dd/mm/yyyy")
                ArepBalanzaMayor.Show 1
        '         fPreview.arv.ReportSource = ArepBalanzaMayor
        '         fPreview.Show 1
            Else
                ArepBalanzaMayorColumna.DataControl1.ConnectionString = ConexionReporte
                ArepBalanzaMayorColumna.DataControl1.Source = Sql
                ArepBalanzaMayorColumna.Logo.Picture = LoadPicture(RutaLogo)
                ArepBalanzaMayorColumna.LblEmpresa.Caption = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                ArepBalanzaMayorColumna.LblEmpresa1.Caption = Me.DtaDatosEmpresa.Recordset("Direccion")
                ArepBalanzaMayorColumna.LblEmpresa2.Caption = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                ArepBalanzaMayorColumna.LblFechaIni.Caption = FechaIni
                ArepBalanzaMayorColumna.LblFechaFin.Caption = FechaFin
                ArepBalanzaMayorColumna.LblFechaImpreso.Caption = "Impreso: " & Format(Now, "dd/mm/yyyy")
                ArepBalanzaMayorColumna.Show 1
        
            End If
         End If
         
         
End Select



Me.Frame6.Visible = False
Me.CmdVerReporte2.Enabled = True
Me.CmdSalir.Enabled = True
Me.CmdVerReporte3.Visible = False
Me.CmdVerReporte2.Visible = False
Me.CmdVerReporte.Visible = True

End Sub

Private Sub ChkExportar_Click()
' If Me.ChkExportar.Value = 1 Then
'   Me.CommonDialog1.ShowSave
'  RutaArchivo = ""
'  RutaArchivo = Me.CommonDialog1.FileName + ".xls"
'
' End If
End Sub

Private Sub ChkMostrarMovxMes_Click()
 If Me.ChkMostrarMovxMes.Value = 1 Then
   Me.ChkSinNiveles.Value = 0
   Me.ChkSinNiveles.Enabled = False
 Else
   Me.ChkSinNiveles.Enabled = True
 End If
End Sub

Private Sub ChkSinNiveles_Click()
 If Me.ChkSinNiveles.Value = 1 Then
   Me.CmbNivel2.Enabled = False
   Me.LblNivel2.Enabled = False
   Me.CmbNivel3.Enabled = False
   Me.LblNivel3.Enabled = False
   Me.CmbNivel4.Enabled = False
   Me.LblNivel4.Enabled = False
   Me.CmbNivel5.Enabled = False
   Me.LblNivel5.Enabled = False
   Me.CmbNivel6.Enabled = False
   Me.LblNivel6.Enabled = False
   Me.CmbNivel7.Enabled = False
   Me.LblNivel7.Enabled = False
 Else
   Me.CmbNivel2.Enabled = True
   Me.LblNivel2.Enabled = True
   Me.CmbNivel3.Enabled = True
   Me.LblNivel3.Enabled = True
   Me.CmbNivel4.Enabled = True
   Me.LblNivel4.Enabled = True
   Me.CmbNivel5.Enabled = True
   Me.LblNivel5.Enabled = True
   Me.CmbNivel6.Enabled = True
   Me.LblNivel6.Enabled = True
   Me.CmbNivel7.Enabled = True
   Me.LblNivel7.Enabled = True
 End If
End Sub

Private Sub CmbReportes_Click()
Dim AÑO1 As String, AÑO2 As String, AÑO3 As String
Dim IngresosVentas, ServiciosVentas, ComisionVentas As Double
Dim CostosProduccion, CostosGeneralesProduccion As Double
 
 Me.FrmDepartamento.Left = 9360
 Me.FrmDepartamento.Caption = "Departamento"
 Me.CmbNivel = 3
 Me.ChkSinNiveles.Visible = False
 Me.Frame1.Visible = True
 Me.ChkMostrarMovxMes.Visible = False
 Me.CmbNivel2.Visible = False
 Me.LblNivel2.Visible = False
 Me.CmbNivel3.Visible = False
 Me.LblNivel3.Visible = False
 Me.CmbNivel4.Visible = False
 Me.LblNivel4.Visible = False
 Me.CmbNivel5.Visible = False
 Me.LblNivel5.Visible = False
 Me.CmbNivel6.Visible = False
 Me.LblNivel6.Visible = False
 Me.CmbNivel7.Visible = False
 Me.LblNivel7.Visible = False
 Me.LblTransaccion.Visible = False
 Me.TxtTransaccion.Visible = False
 Me.Frame3.Visible = False
 Me.Frame7.Visible = False
 Me.CmdVerReporte.Visible = True
 Me.CmdVerReporte2.Visible = False
 Me.CmdVerReporte3.Visible = False
         Me.CmbMoneda.Visible = False
         Me.Label3.Visible = False
         Me.ChkBalanza.Visible = False
         Me.Frame9.Visible = False
         Me.Frame8.Enabled = False
         Me.Frame1.Visible = True
         Me.Frame4.Visible = False
         Me.CmbNivel.Visible = True
         Me.Label6.Visible = True
         Me.frameBalanza.Visible = False
         Me.SSTab.TabVisible(1) = False
         Me.SSTab.TabVisible(2) = False
         Me.ChkQuitarMovimiento.Visible = False
         
             CmbBeneficiario.Visible = False
    LblBeneficiario.Visible = False

         
   Select Case Me.CmbReportes.Text
     Case "RETENCIONES X BENEFICIARIO DETALLE"
     
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = False
         Me.CmdVerReporte3.Visible = True


      Me.Label3.Visible = True
      Me.CmbMoneda.Visible = True
      Me.Label2.Visible = False
      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Label2.Visible = True
       Me.Frame7.Visible = True
      Me.Frame2.Visible = False
'      Me.Frame3.Visible = True

        Frame7.Caption = "Realizar Filtrado por Codigo"
        Me.TxtDesde.Visible = False
        Me.TxtHasta.Visible = False
        Me.DBCodigo.Visible = True
        Me.DBCodigoHasta.Visible = True
        
        CargarCuentas Me.DBCodigo, True
        CargarCuentas Me.DBCodigoHasta, True
     
     
     Case "RETENCIONES X BENEFICIARIO RESUMIDO"
     
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = False
         Me.CmdVerReporte3.Visible = True


      Me.Label3.Visible = True
      Me.CmbMoneda.Visible = True
      Me.Label2.Visible = False
      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Label2.Visible = True
       Me.Frame7.Visible = True
      Me.Frame2.Visible = False
'      Me.Frame3.Visible = True

        Frame7.Caption = "Realizar Filtrado por Codigo"
        Me.TxtDesde.Visible = False
        Me.TxtHasta.Visible = False
        Me.DBCodigo.Visible = True
        Me.DBCodigoHasta.Visible = True
        
        CargarCuentas Me.DBCodigo, True
        CargarCuentas Me.DBCodigoHasta, True
     
     
     Case "RETENCIONES X PAGAR"
     
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = False
         Me.CmdVerReporte3.Visible = True


      Me.Label3.Visible = True
      Me.CmbMoneda.Visible = True
      Me.Label2.Visible = False
      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Label2.Visible = True
       Me.Frame7.Visible = True
      Me.Frame2.Visible = False
'      Me.Frame3.Visible = True

        Frame7.Caption = "Realizar Filtrado por Codigo"
        Me.TxtDesde.Visible = False
        Me.TxtHasta.Visible = False
        Me.DBCodigo.Visible = True
        Me.DBCodigoHasta.Visible = True
        
        CargarCuentas Me.DBCodigo, True
        CargarCuentas Me.DBCodigoHasta, True
     
     
     Case "PRESUPUESTO DEL MES"
      Me.CmbNivel = 3
      Me.Frame4.Visible = True
      Me.Frame1.Visible = False
      Me.CmbMoneda.Visible = True
      Me.Label3.Visible = True
               Me.ChkSinNiveles.Visible = True
         Me.ChkSinNiveles.Caption = "Expresado en Dolares"
      
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
         Me.Option6.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop
   
   
       Case "REPORTE PBI"
      Me.Label3.Visible = True
      Me.CmbMoneda.Visible = True
      Me.Label2.Visible = False
      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Label2.Visible = True
      Me.Frame7.Visible = False
      Me.Frame2.Visible = False
      Me.Frame3.Visible = False
         
       Case "COMPROBANTE DE DIARIO"
       
         Me.FrmDepartamento.Left = 3960
         Me.FrmDepartamento.top = 2160
         Me.FrmDepartamento.Caption = "Fuente"
      
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = False
         Me.CmdVerReporte3.Visible = True
        
              Me.Label4.Visible = True
              Me.DTFecha1.Visible = True
              Me.Label2.Visible = True
              Me.Label3.Visible = True
              Me.CmbMoneda.Visible = True
'              Me.Frame7.Visible = True
'              Me.Frame2.Visible = False
'              Me.Frame3.Visible = True
       Case "LISTA CUENTAS X PAGAR"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = False
         Me.CmdVerReporte3.Visible = True
        
              Me.Label4.Visible = True
              Me.DTFecha1.Visible = True
              Me.Label2.Visible = True
              Me.Frame7.Visible = True
              Me.Frame2.Visible = False
              Me.Frame3.Visible = True
       Case "LISTA CUENTAS X COBRAR"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = False
         Me.CmdVerReporte3.Visible = True
        
              Me.Label4.Visible = True
              Me.DTFecha1.Visible = True
              Me.Label2.Visible = True
              Me.Frame7.Visible = True
              Me.Frame2.Visible = False
              Me.Frame3.Visible = True
         
        Case "RETENCIONES EN LA FUENTE I.R X COBRAR"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.Frame1.Visible = False
         Me.Frame4.Visible = True
         Me.CmbFin.Visible = False
         Me.CmbNivel.Visible = False
         Me.Label7.Visible = False
         Me.Label6.Visible = False
            Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
            Me.DtaConsulta.Refresh
            Do While Not DtaConsulta.Recordset.EOF
             If AÑO1 = "" Then
              AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option8.Caption = AÑO1
             ElseIf AÑO2 = "" Then
              AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option7.Caption = AÑO2
             Else
               AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
               Me.Option6.Caption = AÑO3
             End If
              
              Me.DtaConsulta.Recordset.MoveNext
            Loop
 
              
        Case "RETENCIONES EN LA FUENTE I.R X PAGAR"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.Frame1.Visible = False
         Me.Frame4.Visible = True
         Me.CmbFin.Visible = False
         Me.CmbNivel.Visible = False
         Me.Label7.Visible = False
         Me.Label6.Visible = False
            Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
            Me.DtaConsulta.Refresh
            Do While Not DtaConsulta.Recordset.EOF
             If AÑO1 = "" Then
              AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option8.Caption = AÑO1
             ElseIf AÑO2 = "" Then
              AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option7.Caption = AÑO2
             Else
               AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
               Me.Option6.Caption = AÑO3
             End If
              
              Me.DtaConsulta.Recordset.MoveNext
            Loop
            
       Case "RETENCIONES EN LA FUENTE I.R X COBRAR"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.Frame1.Visible = False
         Me.Frame4.Visible = True
         Me.CmbFin.Visible = False
         Me.CmbNivel.Visible = False
         Me.Label7.Visible = False
         Me.Label6.Visible = False
            Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
            Me.DtaConsulta.Refresh
            Do While Not DtaConsulta.Recordset.EOF
             If AÑO1 = "" Then
              AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option8.Caption = AÑO1
             ElseIf AÑO2 = "" Then
              AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option7.Caption = AÑO2
             Else
               AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
               Me.Option6.Caption = AÑO3
             End If
              
              Me.DtaConsulta.Recordset.MoveNext
            Loop
       
       Case "ANEXO FISCAL IVA CLIENTES"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.Frame1.Visible = False
         Me.Frame4.Visible = True
         Me.CmbFin.Visible = False
         Me.CmbNivel.Visible = False
         Me.Label7.Visible = False
         Me.Label6.Visible = False
            Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
            Me.DtaConsulta.Refresh
            Do While Not DtaConsulta.Recordset.EOF
             If AÑO1 = "" Then
              AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option8.Caption = AÑO1
             ElseIf AÑO2 = "" Then
              AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option7.Caption = AÑO2
             Else
               AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
               Me.Option6.Caption = AÑO3
             End If
              
              Me.DtaConsulta.Recordset.MoveNext
            Loop
            
      Case "ANEXO FISCAL IVA PROVEEDOR"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.Frame1.Visible = False
         Me.Frame4.Visible = True
         Me.CmbFin.Visible = False
         Me.CmbNivel.Visible = False
         Me.Label7.Visible = False
         Me.Label6.Visible = False
            Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
            Me.DtaConsulta.Refresh
            Do While Not DtaConsulta.Recordset.EOF
             If AÑO1 = "" Then
              AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option8.Caption = AÑO1
             ElseIf AÑO2 = "" Then
              AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option7.Caption = AÑO2
             Else
               AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
               Me.Option6.Caption = AÑO3
             End If
              
              Me.DtaConsulta.Recordset.MoveNext
            Loop
         
      Case "PUNTO DE EQUILIBRIO"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = False
         Me.CmdVerReporte3.Visible = True
         Me.Frame1.Visible = False
         Me.Frame9.Visible = True
         Me.Frame4.Visible = True
      Case "COMPARATIVO INGRESOS VRS GASTOS"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.Frame1.Visible = False
         Me.Frame9.Visible = True
         Me.Frame4.Visible = True
      Case "COMPARATIVO UTILIDADES"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.Frame1.Visible = False
         Me.Frame9.Visible = True
         Me.Frame4.Visible = True
      Case "CATALOGO RESUMEN"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.Frame1.Visible = False
         Me.Frame4.Visible = True
      Case "RAZONES FINANCIERAS"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.Frame1.Visible = False
         Me.Frame4.Visible = True
      Case "BALANCE GENERAL TRADICIONAL"
         Me.CmbNivel = 3
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.CmbMoneda.Visible = True
         Me.Label3.Visible = True
         Me.Frame4.Visible = True
         Me.Frame1.Visible = False
         Me.ChkQuitarMovimiento.Visible = True

         
      Case "ESTADO DE RESULTADO TRADICIONAL"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.CmbMoneda.Visible = True
         Me.Label3.Visible = True
         Me.Frame9.Visible = True
         Me.Frame4.Visible = True
         Me.Frame1.Visible = False
         Me.ChkQuitarMovimiento.Visible = True
      Case "ESTADO DE RESULTADO RESUMEN 2"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.CmbMoneda.Visible = True
         Me.Label3.Visible = True
         Me.Frame9.Visible = True
         Me.Frame4.Visible = True
         Me.CmbNivel.Visible = False
         Me.Label6.Visible = False
         Me.Frame1.Visible = False
         Me.ChkQuitarMovimiento.Visible = True
      Case "ESTADO DE RESULTADO RESUMEN"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.CmbMoneda.Visible = True
         Me.Label3.Visible = True
         Me.Frame9.Visible = True
         Me.Frame4.Visible = True
         Me.CmbNivel.Visible = False
         Me.Label6.Visible = False
         Me.Frame1.Visible = False
         Me.ChkQuitarMovimiento.Visible = True
      Case "ESTADO DE RESULTADO RESUMEN ANEXOS"
         Me.SSTab.TabVisible(2) = True
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmdVerReporte3.Visible = False
         Me.CmbMoneda.Visible = True
         Me.Label3.Visible = True
         Me.Frame9.Visible = True
         Me.Frame4.Visible = True
         Me.CmbNivel.Visible = False
         Me.Label6.Visible = False
         Me.Frame1.Visible = False
         Me.ChkQuitarMovimiento.Visible = True
     Case "BALANCE GENERAL RESUMEN"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.Frame1.Visible = False
         Me.CmbMoneda.Visible = True
         Me.Label3.Visible = True
         Me.Frame4.Visible = True
         Me.CmbNivel.Visible = False
         Me.Label6.Visible = False
         
     Case "BALANCE GENERAL RESUMEN ANEXOS"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmbMoneda.Visible = True
         Me.Label3.Visible = True
         Me.Frame4.Visible = True
         Me.Frame1.Visible = False
         Me.CmbNivel.Visible = False
         Me.Label6.Visible = False
         Me.SSTab.TabVisible(1) = True
    Case "LIBRO DIARIO"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmbMoneda.Visible = True
         Me.Label3.Visible = True
         Me.CmbNivel2.Visible = True
         Me.LblNivel2.Visible = True
         Me.CmbNivel2.top = 1600
         Me.LblNivel2.top = 1600
         Me.ChkSinNiveles.top = 1600
         Me.ChkSinNiveles.Left = 6360
         Me.CmbNivel2.Left = 5280
         Me.LblNivel2.Left = 4200
         
         Me.CmbNivel3.Visible = True
         Me.LblNivel3.Visible = True
         Me.CmbNivel3.top = 1950
         Me.LblNivel3.top = 1950
         Me.CmbNivel3.Left = 5280
         Me.LblNivel3.Left = 4200
         Me.CmbNivel4.Visible = True
         Me.LblNivel4.Visible = True
         Me.CmbNivel4.top = 2300
         Me.LblNivel4.top = 2300
         Me.CmbNivel4.Left = 5280
         Me.LblNivel4.Left = 4200
         Me.CmbNivel5.Visible = True
         Me.LblNivel5.Visible = True
         Me.CmbNivel5.top = 2650
         Me.LblNivel5.top = 2650
         Me.CmbNivel5.Left = 5280
         Me.LblNivel5.Left = 4200
         Me.CmbNivel6.Visible = True
         Me.LblNivel6.Visible = True
         Me.CmbNivel6.top = 3000
         Me.LblNivel6.top = 3000
         Me.CmbNivel6.Left = 5280
         Me.LblNivel6.Left = 4200
         Me.CmbNivel7.Visible = True
         Me.LblNivel7.Visible = True
         Me.CmbNivel7.top = 3350
         Me.LblNivel7.top = 3350
         Me.CmbNivel7.Left = 5280
         Me.LblNivel7.Left = 4200
'
         
         Me.ChkSinNiveles.Visible = True
         
    Case "LIBRO MAYOR"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
         Me.CmbMoneda.Visible = True
         Me.Label3.Visible = True
         Me.CmbNivel2.Visible = True
         Me.LblNivel2.Visible = True
         Me.CmbNivel2.top = 1600
         Me.LblNivel2.top = 1600
         Me.ChkSinNiveles.top = 1600
         Me.ChkSinNiveles.Left = 6360
         Me.CmbNivel2.Left = 5280
         Me.LblNivel2.Left = 4200
         
         Me.CmbNivel3.Visible = True
         Me.LblNivel3.Visible = True
         Me.CmbNivel3.top = 1950
         Me.LblNivel3.top = 1950
         Me.CmbNivel3.Left = 5280
         Me.LblNivel3.Left = 4200
         Me.CmbNivel4.Visible = True
         Me.LblNivel4.Visible = True
         Me.CmbNivel4.top = 2300
         Me.LblNivel4.top = 2300
         Me.CmbNivel4.Left = 5280
         Me.LblNivel4.Left = 4200
         Me.CmbNivel5.Visible = True
         Me.LblNivel5.Visible = True
         Me.CmbNivel5.top = 2650
         Me.LblNivel5.top = 2650
         Me.CmbNivel5.Left = 5280
         Me.LblNivel5.Left = 4200
         Me.CmbNivel6.Visible = True
         Me.LblNivel6.Visible = True
         Me.CmbNivel6.top = 3000
         Me.LblNivel6.top = 3000
         Me.CmbNivel6.Left = 5280
         Me.LblNivel6.Left = 4200
         Me.CmbNivel7.Visible = True
         Me.LblNivel7.Visible = True
         Me.CmbNivel7.top = 3350
         Me.LblNivel7.top = 3350
         Me.CmbNivel7.Left = 5280
         Me.LblNivel7.Left = 4200
'
         
         Me.ChkSinNiveles.Visible = True
         
         Me.ChkMostrarMovxMes.Visible = True
         Me.ChkMostrarMovxMes.top = 1900
         Me.ChkMostrarMovxMes.Left = 6360
         Me.ChkMostrarMovxMes.Visible = True
         Me.ChkMostrarMovxMes.Caption = "Mostrar Mov x Mes"
         
         
    Case "DETALLE DIARIO MAYOR"
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = True
    Case "ESTRUCTURA DE CUENTAS"
         Me.Frame2.Visible = False
      Me.Frame1.Visible = False
   Case "CONTROL DE BANCOS"
             Me.Label2.Visible = True
             Me.Frame7.Visible = True
             Me.Frame7.top = 1440
             Me.CmbMoneda.Visible = True
             Me.Label3.Visible = True
   Case "COMPROBANTE DE PAGO"
       Me.LblTransaccion.Visible = True
      Me.TxtTransaccion.Visible = True
      Me.Label2.Visible = False
      Me.Frame7.Visible = False
   Case "LISTADO CUENTAS"
         Me.Frame2.Visible = False
      Me.Frame1.Visible = False
   Case "TARJETA ACTIVO FIJO"
      Me.Frame2.Visible = False
      Me.Frame1.Visible = False
   Case "LISTA ACTIVO FIJO"
      Me.Frame2.Visible = False
      Me.Frame1.Visible = False
      
   Case "TARJETA EMPLEADOS"
     Me.Frame2.Visible = False
     Me.Frame1.Visible = False
   Case "GRUPOS DE CUENTAS"
     Me.Frame2.Visible = False
     Me.Frame1.Visible = False
   Case "TASAS DE CAMBIO"
      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Frame2.Visible = False
      Me.Frame1.Visible = True
   Case "LISTA DE USUARIOS"
      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Frame2.Visible = False
      Me.Frame1.Visible = False
   Case "TARJETA CONTRATISTA"
      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Frame2.Visible = False
      Me.Frame1.Visible = False
   Case "REGISTRO DE MOVIMIENTOS"
      Me.Frame2.Visible = False
      Me.Frame7.Visible = False
  Case "AUXILIAR x GRUPO"
      Me.Label3.Visible = False
      Me.CmbMoneda.Visible = False
      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Frame7.Visible = False
      Me.Frame2.Visible = False
   Case "TOTAL AUXILIAR DE CUENTAS"

      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Label2.Visible = True
      Me.Frame7.Visible = True
      Me.Frame2.Visible = False
      Me.Frame3.Visible = True
      
   Case "AUXILIAR CUENTAS X BENEFICIARIO"
      Me.Label3.Visible = True
      Me.CmbMoneda.Visible = True
      Me.Label2.Visible = False
      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Label2.Visible = True
      Me.Frame7.Visible = True
      Me.Frame2.Visible = False
      Me.Frame3.Visible = True
      CmbBeneficiario.Visible = True
      LblBeneficiario.Visible = True
      
         Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = False
         Me.CmdVerReporte3.Visible = True
      
   Case "AUXILIAR DE CUENTAS"
      Me.Label3.Visible = True
      Me.CmbMoneda.Visible = True
      Me.Label2.Visible = False
      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Label2.Visible = True
      Me.Frame7.Visible = True
      Me.Frame2.Visible = False
      Me.Frame3.Visible = True
'      CmbBeneficiario.Visible = True
'      LblBeneficiario.Visible = True
    

   Case "BALANZA DE COMPROBACION"
      Me.Label2.Visible = False
      Me.Label4.Visible = True
      Me.DTFecha1.Visible = True
      Me.Frame3.Visible = True
      Me.Frame7.Visible = True
      Me.Frame8.Enabled = True
      Me.Frame2.Visible = False
      Me.Label3.Visible = True
      Me.CmbMoneda.Visible = True
      Me.ChkBalanza.Visible = True
      Me.Option4.Visible = True
      Me.Option9.Visible = False
      Me.Option11.Value = True
      
'      Me.frameBalanza.Visible = True
'      Me.frameBalanza.top = 2160
'      Me.frameBalanza.Left = 3960
'      Me.Frame9.Visible = True
'      Me.Frame9.top = 2760
'      Me.Frame9.Left = 6360
      Me.ChkQuitarMovimiento.Visible = True
   Case "CUENTAS X COBRAR"
      Me.Label3.Visible = True
      Me.CmbMoneda.Visible = True
      Me.Label2.Visible = True
      Me.Frame7.Visible = True
      Me.Label4.Visible = False
      Me.Frame2.Visible = False
   Case "CUENTAS X PAGAR"
      Me.Label3.Visible = True
      Me.Label2.Visible = True
      Me.Frame7.Visible = True
      Me.CmbMoneda.Visible = True
      Me.Label4.Visible = False
      Me.Frame2.Visible = False
   
   Case "PRESUPUESTO ANUAL"
      Me.Frame1.Visible = False
      Me.Label2.Visible = False
         Me.Frame7.Visible = False
      Me.Frame2.Visible = True
    Me.LblTitulo.Caption = "Seleccione el Periodo"
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option1.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option2.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
         Me.Option3.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop
   Case "BALANCE GENERAL"
      Me.CmbNivel = 3
      Me.Frame4.Visible = True
      Me.Frame1.Visible = False
      Me.CmbMoneda.Visible = True
      Me.Label3.Visible = True
               Me.ChkSinNiveles.Visible = True
         Me.ChkSinNiveles.Caption = "Expresado en Dolares"
      
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
         Me.Option6.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop
      
   Case "BALANCE COMPARATIVO"
      Me.CmbNivel = 3
      Me.Frame4.Visible = True
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
         Me.Option6.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop
   
    Case "BALANCE GENERAL"

      Me.Frame4.Visible = True
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
         Me.Option6.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop
      
    Case "BALANCE ACUMULADO"
      Me.Frame4.Visible = True
      Me.Frame1.Visible = False
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
         Me.Option6.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop
      
    Case "BALANCE HISTORICO"
      Me.Frame4.Visible = True
      Me.Frame1.Visible = False
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
         Me.Option6.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop

    Case "ESTADO DE RESULTADO"
'      Me.FrmDepartamento.Left = 3960
'      Me.FrmDepartamento.top = 2160
      Me.CmbMoneda.Visible = True
      Me.Label3.Visible = True
      Me.ChkSinNiveles.Visible = True
      Me.ChkSinNiveles.Caption = "Expresado en Dolares"

      
      
      Me.ChkQuitarMovimiento.Visible = True
      Me.Frame4.Visible = True
      Me.Frame1.Visible = False
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
         Me.Option6.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop
      
    Case "ESTADO DE RESULTADO DPTO"
    

      Me.CmdVerReporte.Visible = False
      Me.CmdVerReporte2.Visible = False
      Me.CmdVerReporte3.Visible = True
         
      Me.FrmDepartamento.Left = 3960
      Me.FrmDepartamento.top = 2160
      Me.ChkQuitarMovimiento.Visible = True
      Me.Frame4.Visible = True
      Me.Frame1.Visible = False
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
         Me.Option6.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop
      
    Case "RESULTADO ACUMULADO"

      Me.ChkQuitarMovimiento.Visible = True
      Me.Frame4.Visible = True
      Me.Frame1.Visible = False
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
         Me.Option6.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop
      
    Case "RESULTADO HISTORICO"

      Me.ChkQuitarMovimiento.Visible = True
      Me.Frame4.Visible = True
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Me.Frame1.Visible = False
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
        Me.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
         Me.Option6.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop
 End Select

End Sub

Private Sub CmdBuscarEmpleado_Click()
If Me.Option4.Value = True Then
 QueProducto = "CuentaReportes"
 FrmConsulta.Show 1
Else
 QUIEN = "CuentasReportes"
 FrmGrupoLista.Show 1

End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub


Private Sub CmdVerReporte_Click()
Dim Fechas1 As String, Fechas2 As String, Orden As Integer, Sql As String
Dim UltimoOrden As Integer, RegIngresos  As Integer, PrimReg As Integer, UltReg As Integer
Dim Utilidad As Double, Utilidad2 As Double, Utilidad3 As Double, RegTCostosOper As Integer
Dim Decrementador As Integer
Dim RegInicioCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro donde comienzan los costos operativos
Dim RegTotalCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro de total de costos operativos
Dim Rs As New ADODB.Recordset, Fecha2 As String
Dim CodigoCuentaDesde As String, CodigoCuentaHasta As String
Dim rpt As Object
Dim fPreview As New FrmPreview





'On Error GoTo TipoErrs
Me.Frame6.Visible = True
Me.CmdVerReporte.Enabled = False
Me.CmdSalir.Enabled = False
SaldoIni = 0
SaldoFin = 0
Total1 = 0
TotalCuenta = 0

FechaIni = Me.DTFecha1.Value
FechaFin = Me.DTFecha2.Value

Select Case Me.CmbReportes.Text




Case "CUENTAS X COBRAR"


If Me.DBCodigo.Text = "" Then
'SQL = "SELECT     Transacciones.CodCuentas AS CodCuentas, Transacciones.DescripcionMovimiento AS DescripcionMovimiento, " & _
'     "Transacciones.FechaTransaccion AS FechaTransaccion, Transacciones.FechaVence AS FechaVence, " & _
'     "Transacciones.NumeroMovimiento AS NumeroMovimiento, Cuentas.TipoCuenta AS TipoCuenta, Transacciones.FacturaNo AS FacturaNo, " & _
'     "Transacciones.Debito * Transacciones.TCambio AS Debito, Transacciones.Credito * Transacciones.TCambio AS Credito, " & _
'     "(Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio AS Saldo, Transacciones.ChequeNo, Cuentas.DescripcionCuentas, " & _
'     "Transacciones.TCambio " & _
'     "FROM         Transacciones INNER JOIN " & _
'     "Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
'     "WHERE  ((Transacciones.FechaTransaccion) Between '" & Format(Me.DTFecha1.Value, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "') AND (Cuentas.TipoCuenta = N'Cuentas x Cobrar') " & _
'     "ORDER BY Transacciones.FacturaNo "

Fecha2 = Format(FechaFin, "yyyy-mm-dd")
Sql = "SELECT Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.FechaVence) AS FechaVence, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Transacciones.FacturaNo) AS FacturaNo, SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) AS Saldo, AVG(Transacciones.TCambio) AS TCambio, Cuentas.DescripcionCuentas  " & _
      "FROM   Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas GROUP BY Transacciones.CodCuentas, Cuentas.DescripcionCuentas  " & _
      "HAVING (MAX(Transacciones.FechaTransaccion) < CONVERT(DATETIME, '" & Fecha2 & "', 102)) AND (MAX(Cuentas.TipoCuenta) = 'Cuentas x Cobrar') AND (SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) <> 0) AND (MAX(Transacciones.FacturaNo) <> N'-') AND (MAX(Transacciones.FacturaNo) <> N'.') ORDER BY MAX(Transacciones.FacturaNo)"

Else
CodigoCuenta = Me.DBCodigo.Text

Fecha2 = Format(FechaFin, "yyyy-mm-dd")


'SQL = "SELECT Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.FechaVence) AS FechaVence, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Transacciones.FacturaNo) AS FacturaNo, SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) AS Saldo, AVG(Transacciones.TCambio) AS TCambio, Cuentas.DescripcionCuentas  " & _
'      "FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas GROUP BY Transacciones.CodCuentas, Cuentas.DescripcionCuentas  " & _
'      "HAVING (MAX(Transacciones.FechaTransaccion) < CONVERT(DATETIME, '" & Fecha2 & "', 102)) AND (MAX(Cuentas.TipoCuenta) = 'Cuentas x Cobrar') AND (SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) <> 0) AND (MAX(Transacciones.FacturaNo) <> N'-') AND (MAX(Transacciones.FacturaNo) <> N'.') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY MAX(Transacciones.FacturaNo)"
Sql = "SELECT Transacciones.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) AS Saldo, SUM(Transacciones.TCambio) AS Expr5 " & _
      "FROM  Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
      "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) AND (Cuentas.TipoCuenta = 'Cuentas x Cobrar') GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuenta & "' AND '" & Me.DBCodigoHasta.Text & "')"

End If
  

  EstadoCuenta.AdoEstadoCuenta.ConnectionString = ConexionReporte
  EstadoCuenta.AdoEstadoCuenta.Source = Sql
  EstadoCuenta.LblFecha1.Caption = Me.DTFecha1.Value
  EstadoCuenta.LblFecha.Caption = Me.DTFecha2.Value

   EstadoCuenta.Logo.Picture = LoadPicture(RutaLogo)

  EstadoCuenta.LblEmpresa.Caption = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
  EstadoCuenta.LblEmpresa1.Caption = Me.DtaDatosEmpresa.Recordset("Direccion")
  EstadoCuenta.LblEmpresa2.Caption = "RUC " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
  
  
'  fPreview.arv.ReportSource = EstadoCuenta
'  fPreview.Show 1

     Set rpt = New EstadoCuenta
     rpt.AdoEstadoCuenta.ConnectionString = ConexionReporte
     rpt.AdoEstadoCuenta.Source = Sql
     fPreview.RunReport rpt
     fPreview.Show 1

Case "ESTRUCTURA DE CUENTAS"
    
    Ejecutar.Execute "delete from reportes"
    
    Me.DtaReportes.Refresh
'    Do While Not Me.DtaReportes.Recordset.EOF
'     Me.DtaReportes.Recordset.Delete
'     Me.DtaReportes.Recordset.MoveNext
'    Loop
  EstructuraCatalogo ("Catalogo")
    Me.DtaReportes.Refresh
    ArepCatalogo.Logo.Picture = LoadPicture(RutaLogo)
    ArepCatalogo.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepCatalogo.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepCatalogo.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepCatalogo.LblImpreso = Format(Now, "dd/mm/yyyy")

    ArepCatalogo.DataControl1.Source = "SELECT * From Reportes ORDER BY Orden"
    ArepCatalogo.DataControl1.ConnectionString = ConexionReporte
    '    ArepCatalogo.Show 1

     Set rpt = New ArepCatalogo
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = "SELECT * From Reportes ORDER BY Orden"
     fPreview.RunReport rpt
     fPreview.Show 1
     
'     Dim rpt As Object
'Dim fPreview As New FrmPreview
'
'     Set rpt = New ArepTransacciones
'     rpt.DataControl1.ConnectionString = ConexionReporte
'     rpt.DataControl1.Source = SQL
'     fPreview.RunReport rpt
'
'
'     fPreview.Show 1


Case "ESTADO DE RESULTADO"

    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    Me.ChkQuitarMovimiento.Visible = True
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
       Me.DtaConsulta.Recordset.MoveLast
       i = Me.DtaConsulta.Recordset.RecordCount
       Me.DtaConsulta.Recordset.MoveFirst
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
    
    
    
    
    Me.DtaReportes.Refresh
    
    FrmReportes.lblProgreso.Caption = "Eliminando Datos del Reporte Anterior"
    FrmReportes.osProgress1.Visible = True
    FrmReportes.osProgress1.Value = 0
    FrmReportes.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
    Me.DtaReportes.Refresh
'    Do While Not Me.DtaReportes.Recordset.EOF
'     FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
'
'     Me.DtaReportes.Recordset.Delete
'     Me.DtaReportes.Refresh
'     Me.DtaReportes.Recordset.MoveNext
'    Loop
    
    Rs.Open "DELETE FROM Reportes", Conexion

    CreaEstructura ("Resultado")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.DtaConsulta.Recordset.MoveLast
      Orden = Me.DtaConsulta.Recordset("Orden")
    End If
    
    Orden = Orden + 1
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Costos y Gastos"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "CG"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    Orden = Orden + 1
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Resultado Periodo"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "RP"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    SaldoReportes ("Resultado")
'    SaldoReportes ("UtilidadResultado")
'    SaldoReportesAcumulado ("UtilidadResultado")
'    '------------------------------------------------------------------------------
'    '----------------IGUALO LOS SALDOS DEL RESULTADO ACUMULADO --------------------
'    '------------------------------------------------------------------------------
    Rs.Open "UPDATE [Reportes] Set [Debe1] = [Debe2] ,[Haber1] = [Haber2] WHERE (KeyGrupo = 'RP')", Conexion
    
    

'    SaldoReportesAcumulado ("Resultado")
    SaldoReportes ("UtilidadResultado")
    
    If Me.CmbNivel.Text = 0 Then
     EliminaRegistroCero ("Resultado")
      Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
    Else
     EliminaRegistroCero ("Nivel")
     Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
     Ejecutar.Execute "DELETE FROM Reportes WHERE (Descripcion NOT LIKE N'%Total%') AND (Nivel = " & Me.CmbNivel.Text & ")"
    End If
    
    If Me.ChkSinNiveles.Value = 1 Then
     ConvertirReporte (FechaFin)
    End If
    
    Me.DtaReportes.RecordSource = "Reportes"
    Me.DtaReportes.Refresh
    ArepResultado.LblMoneda.Caption = "Expresado en " & Me.CmbMoneda.Text
    If Dir(RutaLogo) <> "" Then
     ArepResultado.Logo.Picture = LoadPicture(RutaLogo)
    End If
    ArepResultado.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepResultado.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepResultado.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepResultado.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepResultado.LblFechaFin = FechaFin

    ArepResultado.LblFechaIni = FechaIni
     
    Utilidadbruta
    
'    ArepResultado.DataControl1.ConnectionString = ConexionReporte
'    ArepResultado.DataControl1.Source = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
'    ArepResultado.Show 1
    
'     fPreview.arv.ReportSource = ArepResultado
'     fPreview.Show 1

     Set rpt = New ArepResultado
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
     fPreview.RunReport rpt
     fPreview.Show 1

    
    
    FrmReportes.lblProgreso.Caption = ""
    FrmReportes.osProgress1.Visible = False



Case "RESULTADO ACUMULADO"
    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    Me.ChkQuitarMovimiento.Visible = True
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
       Me.DtaConsulta.Recordset.MoveLast
       i = Me.DtaConsulta.Recordset.RecordCount
       Me.DtaConsulta.Recordset.MoveFirst
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
    
    
    
    Me.DtaReportes.Refresh
    

    Ejecutar.Execute "delete from reportes"
    CreaEstructura ("Resultado")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.DtaConsulta.Recordset.MoveLast
      Orden = Me.DtaConsulta.Recordset("Orden")
    End If
    
    Orden = Orden + 1
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Costos y Gastos"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "CG"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    Orden = Orden + 1
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Resultado Periodo"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "RP"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    SaldoReportesAcumulado ("Resultado")
    SaldoReportesAcumulado ("UtilidadResultado")
    
    If Me.CmbNivel.Text = 0 Then
     EliminaRegistroCero ("Resultado")
      Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
    Else
     EliminaRegistroCero ("Nivel")
     Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
     Ejecutar.Execute "DELETE FROM Reportes WHERE (Descripcion NOT LIKE N'%Total%') AND (Nivel = " & Me.CmbNivel.Text & ")"
    End If
    
'    If Me.CmbNivel.Text = 0 Then
'     EliminaRegistroCero ("Resultado")
'    Else
'     EliminaRegistroCero ("Nivel")
'    End If

    Me.DtaReportes.Refresh
    ArepResultadoAcumulado.LblMoneda.Caption = "Expresado en " & Me.CmbMoneda.Text
    ArepResultadoAcumulado.Logo.Picture = LoadPicture(RutaLogo)
    ArepResultadoAcumulado.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepResultadoAcumulado.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepResultadoAcumulado.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepResultadoAcumulado.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepResultadoAcumulado.LblFechaFin = FechaFin
    ArepResultadoAcumulado.LblFechaIni = FechaIni
    ArepResultadoAcumulado.LblAcumulado.Caption = "SALDO HASTA " & FechaFin
    ArepResultadoAcumulado.LblActual.Caption = "ACTIVIDAD PERIODO"
    ArepResultadoAcumulado.LblBalance.Caption = "ESTADO DE RESULTADO ACUMULADO"
    
    
    Utilidadbruta
    
    ArepResultadoAcumulado.DataControl1.ConnectionString = ConexionReporte
    ArepResultadoAcumulado.DataControl1.Source = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    
'    ArepResultadoAcumulado.Show 1
    
'     fPreview.arv.ReportSource = ArepResultadoAcumulado
'     fPreview.Show 1

     Set rpt = New ArepResultadoAcumulado
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
     fPreview.RunReport rpt
     fPreview.Show 1
     
    FrmReportes.lblProgreso.Caption = ""
    FrmReportes.osProgress1.Visible = False
    

Case "RESULTADO HISTORICO"
    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    Me.ChkQuitarMovimiento.Visible = True
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
       Me.DtaConsulta.Recordset.MoveLast
       i = Me.DtaConsulta.Recordset.RecordCount
       Me.DtaConsulta.Recordset.MoveFirst
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
    
    Me.DtaReportes.Refresh
    
    

    Ejecutar.Execute "delete from reportes"
    Me.DtaReportes.Refresh
    CreaEstructura ("Resultado")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.DtaConsulta.Recordset.MoveLast
      Orden = Me.DtaConsulta.Recordset("Orden")
    End If
    
    Orden = Orden + 1
    
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Costos y Gastos"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "CG"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    Orden = Orden + 1
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Resultado Periodo"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "RP"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    SaldoReportesAcumulado ("Resultado")
    SaldoReportesAcumulado ("UtilidadResultado")
    
'    If Me.CmbNivel.Text = 0 Then
'     EliminaRegistroCero ("Resultado")
'    Else
'     EliminaRegistroCero ("Nivel")
'    End If

    If Me.CmbNivel.Text = 0 Then
     EliminaRegistroCero ("Resultado")
      Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
    Else
     EliminaRegistroCero ("Nivel")
     Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
     Ejecutar.Execute "DELETE FROM Reportes WHERE (Descripcion NOT LIKE N'%Total%') AND (Nivel = " & Me.CmbNivel.Text & ")"
    End If

    Me.DtaReportes.Refresh
    ArepResultadoHistorico.LblMoneda.Caption = "Expresado en " & Me.CmbMoneda.Text
    ArepResultadoHistorico.Logo.Picture = LoadPicture(RutaLogo)
    ArepResultadoHistorico.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepResultadoHistorico.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepResultadoHistorico.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepResultadoHistorico.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepResultadoHistorico.LblFechaFin = FechaFin
    ArepResultadoHistorico.LblFechaIni = FechaIni
    ArepResultadoHistorico.LblAcumulado.Caption = "SALDO HASTA " & FechaFin
    ArepResultadoHistorico.LblActual.Caption = "ACTIVIDAD PERIODO"
    ArepResultadoHistorico.LblAnterior.Caption = "SALDO ANTES " & FechaIni
    ArepResultadoHistorico.LblBalance.Caption = "ESTADO DE RESULTADO HISTORICO"
    
    Utilidadbruta
    
    ArepResultadoHistorico.DataControl1.ConnectionString = ConexionReporte
    ArepResultadoHistorico.DataControl1.Source = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    Sql = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
   
     Set rpt = New ArepResultadoHistorico
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = Sql
     fPreview.RunReport rpt
     fPreview.Show 1
   
'     fPreview.arv.ReportSource = ArepResultadoHistorico
'     fPreview.Show 1
    FrmReportes.lblProgreso.Caption = ""
    FrmReportes.osProgress1.Visible = False



Case "BALANCE GENERAL"

'''''''''''''''''''''''''INICIO EL BALANCE GENERAL ''''''''''''''''''''''''''''''''''''''''''''''''
 
    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
       If Not Me.DtaConsulta.Recordset.EOF Then
         Me.DtaConsulta.Recordset.MoveLast
         i = Me.DtaConsulta.Recordset.RecordCount
         Me.DtaConsulta.Recordset.MoveFirst
       End If
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
    
    
    Ejecutar.Execute "delete from reportes"
'    Me.DtaReportes.Refresh
'    Do While Not Me.DtaReportes.Recordset.EOF
'     Me.DtaReportes.Recordset.Delete
'     Me.DtaReportes.Recordset.MoveNext
'    Loop
    CreaEstructura ("Balance")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.DtaConsulta.Recordset.MoveLast
      Orden = Me.DtaConsulta.Recordset("Orden") + 1
    End If
    
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Pasivo Màs Capital"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "PC"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update

'   SaldoReportes ("Utilidad")
'   SaldoReportes ("Balance")

    SaldoReportesAcumulado ("Utilidad")
    SaldoReportesAcumulado ("Balance")
'    SaldoReportesAcumulado ("UtilidadAnterior")
    

   
    If Me.CmbNivel.Text = 0 Then
     EliminaRegistroCero ("Balance")
    Else
     EliminaRegistroCero ("Nivel")
     Ejecutar.Execute "DELETE FROM Reportes WHERE (Descripcion NOT LIKE N'%Total%') AND (Nivel = " & Me.CmbNivel.Text & ")"
    End If
    
    AjusteDiferencial ("Balance")
    
'    '))))))))))PARCHE TEMPORAL PANAM ///////////////////////
'    Dim UtilidadAnterior As Double
'    If Me.CmbMoneda.Text = "Dólares" Then
'        UtilidadAnterior = -332.76
'        MDIPrimero.AdoConsulta.RecordSource = "SELECT  Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden, CodCuentas, Ubicacion From Reportes WHERE (Ubicacion LIKE N'%Resultado Periodo%')"
'        MDIPrimero.AdoConsulta.Refresh
'        If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
'            MDIPrimero.AdoConsulta.Recordset("Debe1") = MDIPrimero.AdoConsulta.Recordset("Debe1") - UtilidadAnterior
'            MDIPrimero.AdoConsulta.Recordset("Debe3") = MDIPrimero.AdoConsulta.Recordset("Debe3") - UtilidadAnterior
'            MDIPrimero.AdoConsulta.Recordset.Update
'        End If
'
'        MDIPrimero.AdoConsulta.RecordSource = "SELECT  Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden, CodCuentas, Ubicacion From Reportes WHERE  (CodCuentas = N'5600')"
'        MDIPrimero.AdoConsulta.Refresh
'        If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
'            MDIPrimero.AdoConsulta.Recordset("Debe1") = MDIPrimero.AdoConsulta.Recordset("Debe1") + UtilidadAnterior
'            MDIPrimero.AdoConsulta.Recordset("Debe3") = MDIPrimero.AdoConsulta.Recordset("Debe3") + UtilidadAnterior
'            MDIPrimero.AdoConsulta.Recordset.Update
'        End If
'     End If
    
    
    
    
'    If Me.ChkSinNiveles.Value = 1 Then
'     ConvertirReporte (FechaFin)
'    End If
    
    
    Me.DtaReportes.Refresh
    ArepBalance.LblMoneda.Caption = "Expresado en " & Me.CmbMoneda.Text
    If Dir(RutaLogo) <> "" Then
      ArepBalance.Logo.Picture = LoadPicture(RutaLogo)
    End If
    ArepBalance.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepBalance.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepBalance.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepBalance.LblFechaFin = FechaFin
    ArepBalance.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepBalance.LblFechaIni = FechaIni
    ArepBalance.DataControl1.ConnectionString = ConexionReporte
    ArepBalance.DataControl1.Source = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
'    ArepBalance.Show 1
    
     Set rpt = New ArepBalance
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
     fPreview.RunReport rpt


     fPreview.Show 1
    
    
    
'    SQL = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    


'Dim rpt As Object
'Dim fPreview As New FrmPreview
'
'      SQL = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
'     Set rpt = New ArepBalance
'     rpt.DataControl1.ConnectionString = ConexionReporte
'     rpt.DataControl1.Source = SQL
'     fPreview.RunReport rpt
'
'
'     fPreview.Show 1


    
Case "BALANCE ACUMULADO"
    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
      
       If Me.DtaConsulta.Recordset.EOF Then
         MsgBox "Error en el Periodo", vbCritical, "Zeus Contable"
         Exit Sub
       End If
      
       Me.DtaConsulta.Recordset.MoveLast
       i = Me.DtaConsulta.Recordset.RecordCount
       Me.DtaConsulta.Recordset.MoveFirst
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
    
    
    
    Me.DtaReportes.Refresh
    Do While Not Me.DtaReportes.Recordset.EOF
     Me.DtaReportes.Recordset.Delete
     Me.DtaReportes.Recordset.MoveNext
    Loop
    CreaEstructura ("Balance")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.DtaConsulta.Recordset.MoveLast
      Orden = Me.DtaConsulta.Recordset("Orden") + 1
    End If
    
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Pasivo Màs Capital"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "PC"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
       
    FrmReportes.DtaReportes.Recordset.Update



    SaldoReportesAcumulado ("Utilidad")
    SaldoReportesAcumulado ("Balance")
    
    If Me.CmbNivel.Text = 0 Then
     EliminaRegistroCero ("Balance")
    Else
     EliminaRegistroCero ("Nivel")
     Ejecutar.Execute "DELETE FROM Reportes WHERE (Descripcion NOT LIKE N'%Total%') AND (Nivel = " & Me.CmbNivel.Text & ")"
    End If
    
    AjusteDiferencial ("Balance")
    
    Me.DtaReportes.Refresh
    ArepBalanceComparativo.LblMoneda.Caption = "Expresado en " & Me.CmbMoneda.Text
    
    If Dir(RutaLogo) = "" Then
      MsgBox "sin logo tipo", vbCritical, "Zeus facturacion"
      Exit Sub
    End If
    
    ArepBalanceComparativo.Logo.Picture = LoadPicture(RutaLogo)
    ArepBalanceComparativo.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepBalanceComparativo.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepBalanceComparativo.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepBalanceComparativo.LblFechaFin = FechaFin
    ArepBalanceComparativo.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepBalanceComparativo.LblFechaIni = FechaIni
    ArepBalanceComparativo.DataControl1.ConnectionString = ConexionReporte
    Sql = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, Debe1 + Debe2 AS TotalDebe, Haber1 + Haber2 AS TotalHaber, KeyGrupo,KeyGrupoSuperior , KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    ArepBalanceComparativo.LblBalance.Caption = "BALANCE GENERAL ACUMULADO"
    ArepBalanceComparativo.LblActual.Caption = "ACTIVIDAD PERIODO"
    ArepBalanceComparativo.LblAcumulado.Caption = "SALDO HASTA " & FechaFin
'    ArepBalanceComparativo.Show 1
'     fPreview.arv.ReportSource = ArepBalanceComparativo
'     fPreview.Show 1

     Set rpt = New ArepBalanceComparativo
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = Sql
     fPreview.RunReport rpt
     fPreview.Show 1
    

 Case "BALANCE HISTORICO"


 
    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
      If Me.DtaConsulta.Recordset.RecordCount = 0 Then
        MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
        Exit Sub
      End If
       Me.DtaConsulta.Recordset.MoveLast
       i = Me.DtaConsulta.Recordset.RecordCount
       Me.DtaConsulta.Recordset.MoveFirst
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
    
    
    
    Me.DtaReportes.Refresh
    Do While Not Me.DtaReportes.Recordset.EOF
     Me.DtaReportes.Recordset.Delete
     Me.DtaReportes.Recordset.MoveNext
    Loop
    CreaEstructura ("Balance")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.DtaConsulta.Recordset.MoveLast
      Orden = Me.DtaConsulta.Recordset("Orden") + 1
    End If
    
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Pasivo Màs Capital"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "PC"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update



    SaldoReportesAcumulado ("Utilidad")
    SaldoReportesAcumulado ("Balance")
    
    If Me.CmbNivel.Text = 0 Then
     EliminaRegistroCero ("Balance")
    Else
     EliminaRegistroCero ("Nivel")
    End If
    
    Me.DtaReportes.Refresh
    ArepBalanceHistorico.LblMoneda.Caption = "Expresado en " & Me.CmbMoneda.Text
    If Dir(RutaLogo) <> "" Then
    ArepBalanceHistorico.Logo.Picture = LoadPicture(RutaLogo)
    End If
    ArepBalanceHistorico.LblBalance.Caption = "BALANCE GENERAL HISTORICO"
    ArepBalanceHistorico.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepBalanceHistorico.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepBalanceHistorico.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepBalanceHistorico.LblFechaFin = FechaFin
    ArepBalanceHistorico.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepBalanceHistorico.LblFechaIni = FechaIni
    ArepBalanceHistorico.LblSaldoAntes.Caption = "SALDO ANTES  " & FechaIni
    ArepBalanceHistorico.LblSaldoPeriodo.Caption = "ACTIVIDAD PERIODO"
    ArepBalanceHistorico.LblSaldoAcumulado.Caption = "SALDO HASTA  " & FechaFin
    ArepBalanceHistorico.DataControl1.ConnectionString = ConexionReporte
    ArepBalanceHistorico.DataControl1.Source = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, Debe1 + Debe2 AS TotalDebe, Haber1 + Haber2 AS TotalHaber, KeyGrupo,KeyGrupoSuperior , KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"

    Sql = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, Debe1 + Debe2 AS TotalDebe, Haber1 + Haber2 AS TotalHaber, KeyGrupo,KeyGrupoSuperior , KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
'Dim rpt As Object
'Dim fPreview As New FrmPreview

     Set rpt = New ArepBalanceHistorico
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = Sql
     fPreview.RunReport rpt
     fPreview.Show 1

''    ArepBalanceHistorico.Show 1
'     fPreview.arv.ReportSource = ArepBalanceHistorico
'     fPreview.Show 1
    
Case "BALANZA DE COMPROBACION"

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////ESTA OPCION CALCULA LA BALANZA POR CODIGO //////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 
 If Me.CmbMoneda.Text = "Córdobas" Then
    Moneda = "Cordobas"
 Else
    Moneda = "Dolares"
 End If
 
 
 If Option4.Value = True Then
    Me.DtaReportes.Refresh
    Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
    Me.lblProgreso.AutoSize = True
    Me.osProgress1.Visible = True
    Me.osProgress1.Value = 0
    Me.osProgress1.Min = 0
    Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
    Do While Not Me.DtaReportes.Recordset.EOF
     Me.DtaReportes.Recordset.Delete
        Me.DtaReportes.Recordset.MoveNext
        Me.osProgress1.Value = Me.osProgress1.Value + 1
    Loop
    
    Me.lblProgreso.Caption = ""
    Me.osProgress1.Visible = False
    SaldoReportes ("BalanzaCodigo")
    Me.DtaReportes.Refresh
    'aqui voy a poner el parche para que salga los movimientos de debito y credito en periodo
    Dim Parche As ADODB.Connection
    Set Parche = New ADODB.Connection
    Parche.ConnectionString = Conexion
    Parche.Open

    Parche.Execute "update reportes set debe2=0,haber2=0"
    Parche.Execute "update reportes " & _
        "set debe2=mdebito,haber2=mcredito " & _
        "from (" & ConsultaTotalesMovimientos & ") as Temporal " & _
        "where temporal.codcuentas=reportes.codcuentas"
        
    Parche.Execute "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2"
        
    
    Parche.Execute "Update [Reportes] SET [Debe3] =  CASE WHEN (Debe1 + Debe2) > (Haber1 + Haber2) THEN (Debe1 + Debe2) - (Haber1 + Haber2) ELSE 0 END ,[Haber3] =  CASE WHEN (Haber1 + Haber2) > (Debe1 + Debe2) THEN (Haber1 + Haber2) - (Debe1 + Debe2) ELSE 0 END"
       
       
    '/////////////////////////parche temporal emtrides ////////////////
'     Parche.Execute "UPDATE [dbo].[Reportes] Set [Debe1] = dbo.Reportes.Debe1 - 1.42 ,[Debe3] = dbo.Reportes.Debe3  - 1.42 WHERE dbo.Reportes.CodCuentas = '10020201'"
'     Parche.Execute "UPDATE [dbo].[Reportes] Set [Debe1] = dbo.Reportes.Debe1 + 0.26 ,[Debe3] = dbo.Reportes.Debe3  + 0.26 WHERE dbo.Reportes.CodCuentas = '100101'"
       
    
    If Me.OptTradicional.Value = True Then
        ArepBalanza.LblMoneda.Caption = Me.CmbMoneda.Text
        
        ArepBalanza.Logo.Picture = LoadPicture(RutaLogo)
        
        ArepBalanza.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
        ArepBalanza.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
        ArepBalanza.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
        ArepBalanza.LblFechaFin = Me.DTFecha2.Value
        ArepBalanza.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
        ArepBalanza.LblFechaIni = Me.DTFecha1.Value
        ArepBalanza.DataControl1.ConnectionString = ConexionReporte
        ArepBalanza.DataControl1.Source = "SELECT * From Reportes ORDER BY CodCuentas"
        ArepBalanza.DataControl1.Refresh
        
        QUIEN = "BalanzaCodigo"
        ArepBalanza.Show 1
'         fPreview.arv.ReportSource = ArepBalanza
'         fPreview.Show 1
    End If
    Me.lblProgreso.Caption = ""
    Me.osProgress1.Visible = False
    
    
    
 ElseIf Option5.Value = True Then
   '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   '//////////////////////EN ESTA OPCION CALCULO LA BALANZA POR GRUPOS//////////////////////////////////
   '////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 

 
    Me.DtaReportes.Refresh
    
    Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
    Me.lblProgreso.AutoSize = True
    Me.osProgress1.Value = 0
    Me.osProgress1.Min = 0
    Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
    
    
    
    Do While Not Me.DtaReportes.Recordset.EOF
     Me.osProgress1.Value = Me.osProgress1.Value + 1
     Me.DtaReportes.Recordset.Delete
     Me.DtaReportes.Recordset.MoveNext
    Loop
    CreaEstructura ("Balanza")
    SaldoReportes ("Balanza")
    EliminaRegistroCero ("Balanza")
    
    '-----------------------BORRO TODAS LAS CUENTAS QUE NO SUMAN NINGUN VALOR ------------------
    Set Parche = New ADODB.Connection
    Parche.ConnectionString = Conexion
    Parche.Open
    Parche.Execute "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe2 + Haber2 + Debe3 + Haber3 = 0)"
    
    
    Me.DtaConsulta.RecordSource = "select sum(Debe1) as SumaDebe1,sum(debe2) as SumaDebe2,sum(debe3)as SumaDebe3,sum(haber1) as SumaHaber1,sum(haber2) as SumaHaber2,sum(haber3) as SumaHaber3 from reportes where descripcion not like '%total%'"
    Me.DtaConsulta.Refresh
    If Not DtaConsulta.Recordset.EOF Then
'      ArepBalanza.LblDebe1 = Format(Me.DtaConsulta.Recordset("SumaDebe1"), "##,##0.00")
'      ArepBalanza.LblHaber1 = Format(Me.DtaConsulta.Recordset("SumaHaber1"), "##,##0.00")
'      ArepBalanza.LblDebe2 = Format(Me.DtaConsulta.Recordset("SumaDebe2"), "##,##0.00")
'      ArepBalanza.LblHaber2 = Format(Me.DtaConsulta.Recordset("SumaHaber2"), "##,##0.00")
'      ArepBalanza.LblDebe3 = Format(Me.DtaConsulta.Recordset("SumaDebe3"), "##,##0.00")
'      ArepBalanza.LblHaber3 = Format(Me.DtaConsulta.Recordset("SumaHaber3"), "##,##0.00")
      ArepBalanza.LblDebe3 = Me.DtaConsulta.Recordset("SumaDebe3")
      ArepBalanza.LblHaber3 = Me.DtaConsulta.Recordset("SumaHaber3")
    
    End If
    
    
    Parche.Execute "Update [Reportes] SET [Debe3] =  CASE WHEN (Debe1 + Debe2) > (Haber1 + Haber2) THEN (Debe1 + Debe2) - (Haber1 + Haber2) ELSE 0 END ,[Haber3] =  CASE WHEN (Haber1 + Haber2) > (Debe1 + Debe2) THEN (Haber1 + Haber2) - (Debe1 + Debe2) ELSE 0 END"
    
    
'     If Me.DTFecha1.Value >= "01/01/2023" Then
'
'    '/////////////////////////parche temporal emtrides ////////////////
'     Parche.Execute "UPDATE [dbo].[Reportes] Set [Debe1] = dbo.Reportes.Debe1 - 1.42 ,[Debe3] = dbo.Reportes.Debe3  - 1.42 WHERE dbo.Reportes.CodCuentas = '10020201'"
'
'     Parche.Execute "UPDATE [dbo].[Reportes] Set [Debe1] = dbo.Reportes.Debe1 + 0.26 ,[Debe3] = dbo.Reportes.Debe3  + 0.26 WHERE dbo.Reportes.CodCuentas = '100101'"
'
'     Parche.Execute "UPDATE [dbo].[Reportes] Set [Debe1] = dbo.Reportes.Debe1 + 0.26 ,[Debe3] = dbo.Reportes.Debe3  + 0.26 WHERE (KeyGrupo = 'A010001') AND (Descripcion LIKE '%Total%')"
'     Parche.Execute "UPDATE [dbo].[Reportes] Set [Debe1] = dbo.Reportes.Debe1 - 1.42 ,[Debe3] = dbo.Reportes.Debe3  - 1.42 WHERE (KeyGrupo = 'A010002') AND (Descripcion LIKE '%Total%')"
'     End If
   
       
    
    Dim Resultado As String
      If Me.OptTradicional.Value = True Then
        Me.DtaReportes.Refresh
        ArepBalanza.LblMoneda.Caption = Me.CmbMoneda.Text

'          ArepBalanza.Logo.Picture = LoadPicture(RutaLogo)
'
'        ArepBalanza.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
'        ArepBalanza.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
'        ArepBalanza.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
'        ArepBalanza.LblFechaFin = Me.DTFecha2.Value
'        ArepBalanza.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
'        ArepBalanza.LblFechaIni = Me.DTFecha1.Value
'        ArepBalanza.DataControl1.ConnectionString = ConexionReporte
'        ArepBalanza.DataControl1.Source = "SELECT * From Reportes ORDER BY Orden"
'        ArepBalanza.Field8.Visible = False
'        ArepBalanza.Field9.Visible = False
'        ArepBalanza.Field10.Visible = False
'        ArepBalanza.Field11.Visible = False
'        ArepBalanza.FldTDebe3.Visible = False
'        ArepBalanza.FldTHaber3.Visible = False
        Sql = "SELECT * From Reportes ORDER BY Orden"
        QUIEN = "Balanza"
'        ArepBalanza.Show 1

            Set rpt = New ArepBalanza
            rpt.DataControl1.ConnectionString = ConexionReporte
            rpt.DataControl1.Source = Sql
            fPreview.RunReport rpt
            fPreview.Show 1
            

      ElseIf Me.OptColumna.Value = True Then
        Me.DtaReportes.Refresh
        ArepBalanzaColumna.LblMoneda.Caption = Me.CmbMoneda.Text
        ArepBalanzaColumna.Logo.Picture = LoadPicture(RutaLogo)
        ArepBalanzaColumna.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
        ArepBalanzaColumna.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
        ArepBalanzaColumna.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
        ArepBalanzaColumna.LblFechaFin = Me.DTFecha2.Value
        ArepBalanzaColumna.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
        ArepBalanzaColumna.LblFechaIni = Me.DTFecha1.Value
        ArepBalanzaColumna.DataControl1.ConnectionString = ConexionReporte
        ArepBalanzaColumna.DataControl1.Source = "SELECT * From Reportes ORDER BY Orden"
        ArepBalanzaColumna.Field8.Visible = False
        ArepBalanzaColumna.Field9.Visible = False
        ArepBalanzaColumna.Field10.Visible = False
        ArepBalanzaColumna.Field11.Visible = False
        ArepBalanzaColumna.FldTDebe3.Visible = False
        ArepBalanzaColumna.FldTHaber3.Visible = False
        ArepBalanzaColumna.Show 1
     
      End If
         
     Me.lblProgreso.Caption = ""
    Me.osProgress1.Visible = False
    
 ElseIf Option9.Value = True Then
 
    '/////////////////////////////////////////////////////////////////////////////////////////
    '//////////////EN ESTA OPCION CALCULO POR CUENTA DE MAYOR/////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////
    
Parche.Execute "Update [Reportes] SET [Debe3] =  CASE WHEN (Debe1 + Debe2) > (Haber1 + Haber2) THEN (Debe1 + Debe2) - (Haber1 + Haber2) ELSE 0 END ,[Haber3] =  CASE WHEN (Haber1 + Haber2) > (Debe1 + Debe2) THEN (Haber1 + Haber2) - (Debe1 + Debe2) ELSE 0 END"
 
'     SQL = "SELECT Max(Transacciones.CodCuentas) AS CodCuentas, Max(Transacciones.NumeroMovimiento) AS NumeroMovimiento, Max(Transacciones.NombreCuenta) AS NombreCuenta, Avg(Transacciones.TCambio) AS TCambio, Sum(Transacciones.TCambio*Transacciones.Debito) AS Debito, Sum(Transacciones.TCambio*Transacciones.Credito) AS Credito, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta " & _
'    "FROM Grupos INNER JOIN (Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas) ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
'    "WHERE     (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102)) " & _
'    "GROUP BY Cuentas.KeyGrupo, Cuentas.DescripcionGrupo " & _
'    "ORDER BY Cuentas.KeyGrupo "
'    "WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) " & _

    Sql = "SELECT MAX(Transacciones.CodCuentas) AS CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, SUM(Transacciones.Debito) AS Debito,SUM(Transacciones.Credito) AS Credito, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas GROUP BY Cuentas.KeyGrupo, Cuentas.DescripcionGrupo " & _
          "HAVING (MAX(Transacciones.FechaTransaccion) <= CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) ORDER BY Cuentas.KeyGrupo"


    If Me.OptTradicional.Value = True Then
        ArepBalanzaMayor.DataControl1.ConnectionString = ConexionReporte
        ArepBalanzaMayor.DataControl1.Source = Sql
        ArepBalanzaMayor.Logo.Picture = LoadPicture(RutaLogo)
        ArepBalanzaMayor.LblEmpresa.Caption = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
        ArepBalanzaMayor.LblEmpresa1.Caption = Me.DtaDatosEmpresa.Recordset("Direccion")
        ArepBalanzaMayor.LblEmpresa2.Caption = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
        ArepBalanzaMayor.LblFechaIni.Caption = FechaIni
        ArepBalanzaMayor.LblFechaFin.Caption = FechaFin
        ArepBalanzaMayor.LblFechaImpreso.Caption = "Impreso: " & Format(Now, "dd/mm/yyyy")
        ArepBalanzaMayor.Show 1
'         fPreview.arv.ReportSource = ArepBalanzaMayor
'         fPreview.Show 1
    Else
        ArepBalanzaMayorColumna.DataControl1.ConnectionString = ConexionReporte
        ArepBalanzaMayorColumna.DataControl1.Source = Sql
        ArepBalanzaMayorColumna.Logo.Picture = LoadPicture(RutaLogo)
        ArepBalanzaMayorColumna.LblEmpresa.Caption = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
        ArepBalanzaMayorColumna.LblEmpresa1.Caption = Me.DtaDatosEmpresa.Recordset("Direccion")
        ArepBalanzaMayorColumna.LblEmpresa2.Caption = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
        ArepBalanzaMayorColumna.LblFechaIni.Caption = FechaIni
        ArepBalanzaMayorColumna.LblFechaFin.Caption = FechaFin
        ArepBalanzaMayorColumna.LblFechaImpreso.Caption = "Impreso: " & Format(Now, "dd/mm/yyyy")
        ArepBalanzaMayorColumna.Show 1

    End If
 End If
    

Case "AUXILIAR x GRUPO"
    ArepTotalAuxiliarGrupo.LblFecha = Format(Now, "dd/mm/yyyy")
    NumFecha1 = Me.DTFecha1.Value
    NumFecha2 = Me.DTFecha2.Value
    
    
    ArepTotalAuxiliarGrupo.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
    ArepTotalAuxiliarGrupo.DataControl1.ConnectionString = ConexionReporte
    
    If Me.DBCodigo.Text = "" Then
'     DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ")) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas"
     Me.DtaConsulta.RecordSource = "SELECT  Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas FROM  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE  (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas"
     DtaConsulta.Refresh
     If DtaConsulta.Recordset.EOF Then
      ArepTotalAuxiliarGrupo.DataControl1.Source = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, GrupoCuentas.CodGrupo, GrupoCuentas.DescripcionGrupo FROM GrupoCuentas INNER JOIN (Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas) ON GrupoCuentas.CodGrupo = Cuentas.CodGrupo GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, GrupoCuentas.CodGrupo, GrupoCuentas.DescripcionGrupo ORDER BY GrupoCuentas.CodGrupo, Cuentas.CodCuentas"
      ArepTotalAuxiliarGrupo.Field18.Visible = False
      ArepTotalAuxiliarGrupo.Field19.Visible = False
     Else
'     ArepTotalAuxiliarGrupo.DataControl1.Source = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, GrupoCuentas.CodGrupo, GrupoCuentas.DescripcionGrupo FROM GrupoCuentas INNER JOIN (Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas) ON GrupoCuentas.CodGrupo = Cuentas.CodGrupo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ")) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, GrupoCuentas.CodGrupo, GrupoCuentas.DescripcionGrupo ORDER BY  GrupoCuentas.CodGrupo,Cuentas.CodCuentas"
     ArepTotalAuxiliarGrupo.DataControl1.Source = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas, GrupoCuentas.CodGrupo , GrupoCuentas.DescripcionGrupo FROM GrupoCuentas INNER JOIN  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON GrupoCuentas.CodGrupo = Cuentas.CodGrupo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, GrupoCuentas.CodGrupo, GrupoCuentas.DescripcionGrupo ORDER BY GrupoCuentas.CodGrupo, Cuentas.CodCuentas"
     End If
    Else
'     DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ")) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas"
     Me.DtaConsulta.RecordSource = "SELECT  Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas FROM  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE  (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas"
     DtaConsulta.Refresh
     If DtaConsulta.Recordset.EOF Then
'      ArepTotalAuxiliarGrupo.DataControl1.Source = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas"
      ArepTotalAuxiliarGrupo.DataControl1.Source = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas, GrupoCuentas.CodGrupo , GrupoCuentas.DescripcionGrupo FROM GrupoCuentas INNER JOIN  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON GrupoCuentas.CodGrupo = Cuentas.CodGrupo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, GrupoCuentas.CodGrupo, GrupoCuentas.DescripcionGrupo ORDER BY GrupoCuentas.CodGrupo, Cuentas.CodCuentas"

      ArepTotalAuxiliarGrupo.Field18.Visible = False
      ArepTotalAuxiliarGrupo.Field19.Visible = False
     Else
     ArepTotalAuxiliarGrupo.DataControl1.Source = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ")) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas Having (((Cuentas.CodCuentas) = '" & Me.DBCodigo.Text & "'))"
     'ArepTotalAuxiliar.DataControl1.Source = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas HAVING (((Cuentas.CodCuentas)='" & Me.DBCodigo.Text & "'))"
     ArepTotalAuxiliarGrupo.LblCodigo.Caption = "Mis Code:" & Me.DBCodigo.Text
     End If
    End If
    ArepTotalAuxiliarGrupo.Logo.Picture = LoadPicture(RutaLogo)
    ArepTotalAuxiliarGrupo.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepTotalAuxiliarGrupo.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepTotalAuxiliarGrupo.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
'    ArepTotalAuxiliarGrupo.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepTotalAuxiliarGrupo.Show 1
'      fPreview.arv.ReportSource = ArepTotalAuxiliar
'     fPreview.Show 1
 
 Case "TOTAL AUXILIAR DE CUENTAS"

    
    
    If Me.Option4.Value = True Then
    
     ArepTotalAuxiliar.LblFecha = Format(Now, "dd/mm/yyyy")
            NumFecha1 = Me.DTFecha1.Value
            NumFecha2 = Me.DTFecha2.Value
            
            Fechas1 = Format(Me.DTFecha1.Value, "yyyy-mm-dd")
            Fechas2 = Format(Me.DTFecha2.Value, "yyyy-mm-dd")
            
            
            ArepTotalAuxiliar.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
            ArepTotalAuxiliar.DataControl1.ConnectionString = ConexionReporte
    
    Me.DtaReportes.Refresh
    Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
    Me.lblProgreso.AutoSize = True
    Rs.Open "DELETE FROM Reportes", Conexion
'    Me.osProgress1.Visible = True
'    Me.osProgress1.Value = 0
'    Me.osProgress1.Min = 0
'    Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
'
'
'    Do While Not Me.DtaReportes.Recordset.EOF
'     Me.DtaReportes.Recordset.Delete
'        Me.DtaReportes.Recordset.MoveNext
'        Me.osProgress1.Value = Me.osProgress1.Value + 1
'    Loop
    
    Me.lblProgreso.Caption = ""
    Me.osProgress1.Visible = False
    SaldoReportes ("BalanzaCodigo")
    Me.DtaReportes.Refresh
    'aqui voy a poner el parche para que salga los movimientos de debito y credito en periodo
'    Dim Parche As ADODB.Connection
            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open
            
            
        
            Parche.Execute "update reportes set debe2=0,haber2=0"
            Parche.Execute "update reportes " & _
                "set debe2=mdebito,haber2=mcredito " & _
                "from (" & ConsultaTotalesMovimientos & ") as Temporal " & _
                "where temporal.codcuentas=reportes.codcuentas"
                
            Parche.Execute "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2"
     
                    ArepTotalAuxiliar.Logo.Picture = LoadPicture(RutaLogo)
                    ArepTotalAuxiliar.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                    ArepTotalAuxiliar.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                    ArepTotalAuxiliar.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                    ArepTotalAuxiliar.DataControl1.Source = "SELECT  Reportes.Descripcion, Reportes.Debe1 + Reportes.Haber1 AS SaldoInicial, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3 - Reportes.Haber3 AS SaldoFinal, Reportes.Orden , Cuentas.CodCuentas, Cuentas.DescripcionCuentas FROM  Reportes INNER JOIN  Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas ORDER BY Reportes.Orden"
'                    ArepTotalAuxiliar.Show 1
                    fPreview.arv.ReportSource = ArepTotalAuxiliar
                    fPreview.Show 1

     
     ElseIf Me.Option5.Value = True Then
     
     
            ArepAuxiliarGrupo.LblFecha = Format(Now, "dd/mm/yyyy")
            NumFecha1 = Me.DTFecha1.Value
            NumFecha2 = Me.DTFecha2.Value
            
            Fechas1 = Format(Me.DTFecha1.Value, "yyyy-mm-dd")
            Fechas2 = Format(Me.DTFecha2.Value, "yyyy-mm-dd")
            
            
           ArepAuxiliarGrupo.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
     
        '///////////////////////ESTA OPCION ES TOTAL AUXILIAR POR GRUPO //////////////////////////////

            Me.DtaReportes.Refresh
            
            Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
            Me.lblProgreso.AutoSize = True
            Rs.Open "DELETE FROM Reportes", Conexion
            
'            Me.osProgress1.Value = 0
'            Me.osProgress1.Min = 0
'            Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
'            Do While Not Me.DtaReportes.Recordset.EOF
'             Me.osProgress1.Value = Me.osProgress1.Value + 1
'             Me.DtaReportes.Recordset.Delete
'             Me.DtaReportes.Recordset.MoveNext
'            Loop

'
'            CreaEstructura ("Balanza")
'            SaldoReportes ("Balanza")
'            EliminaRegistroCero ("Balanza")
            SaldoReportes ("BalanzaCodigo")

            Me.DtaReportes.Refresh
                    
'            ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 5))  AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 5) - ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Cuentas.CodCuentas"
            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open

            Parche.Execute "update reportes set debe2=0,haber2=0"
            Parche.Execute "update reportes " & _
                "set debe2=mdebito,haber2=mcredito " & _
                "from (" & ConsultaTotalesMovimientos & ") as Temporal " & _
                "where temporal.codcuentas=reportes.codcuentas"
        
            Parche.Execute "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2"
                    'aqui voy a poner el parche para que salga los movimientos de debito y credito en periodo
                '    Dim Parche As ADODB.Connection
     
                    ArepAuxiliarGrupo.LblFecha = Format(Now, "dd/mm/yyyy")
                    ArepAuxiliarGrupo.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
                    ArepAuxiliarGrupo.DataControlGrupo.ConnectionString = Conexion
                    ArepAuxiliarGrupo.Logo.Picture = LoadPicture(RutaLogo)
                    ArepAuxiliarGrupo.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                    ArepAuxiliarGrupo.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                    ArepAuxiliarGrupo.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
'                    ArepAuxiliarGrupo.LblContador.Caption = Me.DtaDatosEmpresa.Recordset("Elaborado_Por")
'                    ArepAuxiliarGrupo.LblGerenteFinanciero.Caption = Me.DtaDatosEmpresa.Recordset("Revisado_Por")
'                    ArepAuxiliarGrupo.LblGerenteGeneral.Caption = Me.DtaDatosEmpresa.Recordset("Autorizado_Por")
'                    ArepTotalAuxiliarGrupo.Field17.Visible = False
'                    ArepTotalAuxiliarGrupo.Field22.Visible = False
'                    ArepTotalAuxiliarGrupo.Field23.Visible = True
                    If Me.TxtDesde.Text = "" And Me.TxtHasta.Text = "" Then
'                      ArepAuxiliarGrupo.DataControlGrupo.Source = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3 - Reportes.Haber3 AS Debe3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta, Reportes.Nivel, Reportes.Orden, Reportes.CodCuentas, Reportes.Ubicacion, Reportes.CodDepartamento, Reportes.Debe4, Reportes.Haber4, Reportes.Debe5, Reportes.Haber5, Reportes.Debe6, Reportes.Haber6, Grupos.DescripcionGrupo, Reportes.Debe1 + Reportes.Haber1 + Reportes.Debe2 + Reportes.Haber2 + Reportes.Debe3 AS Saldo, Grupos.KeyGrupoSuperior, Grupos_1.DescripcionGrupo AS DescripcionGrupoSuperior FROM Reportes INNER JOIN Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo INNER JOIN Grupos AS Grupos_1 ON Grupos.KeyGrupoSuperior = Grupos_1.KeyGrupo " & _
'                                                                  "WHERE (Reportes.CodCuentas Is Not Null) And (Reportes.Debe1 + Reportes.Haber1 + Reportes.Debe2 + Reportes.Haber2 + Reportes.Debe3 <> 0) ORDER BY Reportes.CodCuentas, Reportes.Orden"
                    
                       ArepAuxiliarGrupo.DataControlGrupo.Source = "SELECT Reportes.Descripcion, CASE WHEN Cuentas.TipoCuenta = 'Activo Fijo' OR Cuentas.TipoCuenta = 'Otros Activos' OR Cuentas.TipoCuenta = 'Caja' OR Cuentas.TipoCuenta = 'Cuentas x Cobrar' OR  Cuentas.TipoCuenta = 'Bancos' OR Cuentas.TipoCuenta = 'Costos' OR Cuentas.TipoCuenta = 'Gastos' OR  Cuentas.TipoCuenta = 'Papeleria - Utiles' OR Cuentas.TipoCuenta = 'Inventario' THEN Reportes.Debe1 - Reportes.Haber1 ELSE Reportes.Haber1 - Reportes.Debe1 END AS SaldoInicial, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, CASE WHEN Cuentas.TipoCuenta = 'Activo Fijo' OR  Cuentas.TipoCuenta = 'Otros Activos' OR Cuentas.TipoCuenta = 'Caja' OR  Cuentas.TipoCuenta = 'Cuentas x Cobrar' OR  Cuentas.TipoCuenta = 'Bancos' OR  Cuentas.TipoCuenta = 'Costos' OR Cuentas.TipoCuenta = 'Gastos' OR  Cuentas.TipoCuenta = 'Papeleria - Utiles' OR " & _
                                                                   "Cuentas.TipoCuenta = 'Inventario' THEN Reportes.Debe3 - Reportes.Haber3 ELSE Reportes.Haber3 - Reportes.Debe3 END AS SaldoFinal, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta, Reportes.Nivel, Reportes.Orden, Reportes.CodCuentas, Reportes.Ubicacion, Reportes.CodDepartamento, Reportes.Debe4, Reportes.Haber4, Reportes.Debe5, Reportes.Haber5, Reportes.Debe6, Reportes.Haber6, Grupos.DescripcionGrupo, Reportes.Debe1 + Reportes.Haber1 + Reportes.Debe2 + Reportes.Haber2 + Reportes.Debe3 AS Saldo, Grupos.KeyGrupoSuperior AS Expr1, Grupos_1.DescripcionGrupo AS DescripcionGrupoSuperior, Cuentas.TipoCuenta FROM  Reportes INNER JOIN Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo INNER JOIN Grupos AS Grupos_1 ON Grupos.KeyGrupoSuperior = Grupos_1.KeyGrupo " & _
                                                                   "Where(Reportes.CodCuentas Is Not Null) And (Reportes.Debe1 + Reportes.Haber1 + Reportes.Debe2 + Reportes.Haber2 + Reportes.Debe3 <> 0) ORDER BY Reportes.CodCuentas, Reportes.Orden"
                    Else
                       ArepAuxiliarGrupo.DataControlGrupo.Source = "SELECT Reportes.Descripcion, CASE WHEN Cuentas.TipoCuenta = 'Activo Fijo' OR Cuentas.TipoCuenta = 'Otros Activos' OR Cuentas.TipoCuenta = 'Caja' OR Cuentas.TipoCuenta = 'Cuentas x Cobrar' OR  Cuentas.TipoCuenta = 'Bancos' OR Cuentas.TipoCuenta = 'Costos' OR Cuentas.TipoCuenta = 'Gastos' OR  Cuentas.TipoCuenta = 'Papeleria - Utiles' OR Cuentas.TipoCuenta = 'Inventario' THEN Reportes.Debe1 - Reportes.Haber1 ELSE Reportes.Haber1 - Reportes.Debe1 END AS SaldoInicial, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, CASE WHEN Cuentas.TipoCuenta = 'Activo Fijo' OR  Cuentas.TipoCuenta = 'Otros Activos' OR Cuentas.TipoCuenta = 'Caja' OR  Cuentas.TipoCuenta = 'Cuentas x Cobrar' OR  Cuentas.TipoCuenta = 'Bancos' OR  Cuentas.TipoCuenta = 'Costos' OR Cuentas.TipoCuenta = 'Gastos' OR  Cuentas.TipoCuenta = 'Papeleria - Utiles' OR " & _
                                                                   "Cuentas.TipoCuenta = 'Inventario' THEN Reportes.Debe3 - Reportes.Haber3 ELSE Reportes.Haber3 - Reportes.Debe3 END AS SaldoFinal, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta, Reportes.Nivel, Reportes.Orden, Reportes.CodCuentas, Reportes.Ubicacion, Reportes.CodDepartamento, Reportes.Debe4, Reportes.Haber4, Reportes.Debe5, Reportes.Haber5, Reportes.Debe6, Reportes.Haber6, Grupos.DescripcionGrupo, Reportes.Debe1 + Reportes.Haber1 + Reportes.Debe2 + Reportes.Haber2 + Reportes.Debe3 AS Saldo, Grupos.KeyGrupoSuperior AS Expr1, Grupos_1.DescripcionGrupo AS DescripcionGrupoSuperior, Cuentas.TipoCuenta FROM  Reportes INNER JOIN Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo INNER JOIN Grupos AS Grupos_1 ON Grupos.KeyGrupoSuperior = Grupos_1.KeyGrupo " & _
                                                                   "Where(Reportes.CodCuentas Is Not Null) And (Reportes.Debe1 + Reportes.Haber1 + Reportes.Debe2 + Reportes.Haber2 + Reportes.Debe3 <> 0) AND (Grupos.KeyGrupo BETWEEN '" & Me.TxtKeyGrupoDesde.Text & "' AND '" & Me.TxtKeyGrupoHasta.Text & "') ORDER BY Reportes.CodCuentas, Reportes.Orden"
'                      ArepAuxiliarGrupo.DataControlGrupo.Source = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3 - Reportes.Haber3 AS Debe3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta, Reportes.Nivel, Reportes.Orden, Reportes.CodCuentas, Reportes.Ubicacion, Reportes.CodDepartamento, Reportes.Debe4, Reportes.Haber4, Reportes.Debe5, Reportes.Haber5, Reportes.Debe6, Reportes.Haber6, Grupos.DescripcionGrupo, Reportes.Debe1 + Reportes.Haber1 + Reportes.Debe2 + Reportes.Haber2 + Reportes.Debe3 AS Saldo, Grupos.KeyGrupoSuperior, Grupos_1.DescripcionGrupo AS DescripcionGrupoSuperior FROM Reportes INNER JOIN Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo INNER JOIN Grupos AS Grupos_1 ON Grupos.KeyGrupoSuperior = Grupos_1.KeyGrupo " & _
                                                                  "WHERE (Reportes.CodCuentas Is Not Null) And (Reportes.Debe1 + Reportes.Haber1 + Reportes.Debe2 + Reportes.Haber2 + Reportes.Debe3 <> 0) AND (Grupos.DescripcionGrupo BETWEEN '" & Me.TxtDesde.Text & "' AND '" & Me.TxtHasta.Text & "') ORDER BY Reportes.CodCuentas, Reportes.Orden"
'                      ArepAuxiliarGrupo.DataControlGrupo.Source = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3 - Reportes.Haber3 As Debe3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta, Reportes.Nivel, Reportes.Orden, Reportes.CodCuentas, Reportes.Ubicacion, Reportes.CodDepartamento, Reportes.Debe4, Reportes.Haber4, Reportes.Debe5, Reportes.Haber5, Reportes.Debe6, Reportes.Haber6, Grupos.DescripcionGrupo, Reportes.Debe1 + Reportes.Haber1 + Reportes.Debe2 + Reportes.Haber2 + Reportes.Debe3 AS Saldo FROM Reportes INNER JOIN Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo  WHERE (Reportes.CodCuentas IS NOT NULL) AND (Reportes.Debe1 + Reportes.Haber1 + Reportes.Debe2 + Reportes.Haber2 + Reportes.Debe3 <> 0) AND (Grupos.DescripcionGrupo BETWEEN '" & Me.TxtDesde.Text & "' AND '" & Me.TxtHasta.Text & "') ORDER BY Reportes.CodCuentas, Reportes.Orden"
                    End If
                    
'                    ArepAuxiliarGrupo.Show 1
                    
                    ArepAuxiliarGrupo.LblFecha = Format(Now, "dd/mm/yyyy")
                    ArepAuxiliarGrupo.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
                    ArepAuxiliarGrupo.DataControlGrupo.ConnectionString = Conexion
                    ArepAuxiliarGrupo.Logo.Picture = LoadPicture(RutaLogo)
                    ArepAuxiliarGrupo.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                    ArepAuxiliarGrupo.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                    ArepAuxiliarGrupo.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                    


                    fPreview.arv.ReportSource = ArepAuxiliarGrupo
                    fPreview.Show 1
                    
                    
'                    fPreview.arv.ReportSource = ArepTotalAuxiliarG
'                    fPreview.Show 1

     
     End If


  Case "COMPROBANTE DE PAGO"
  '///////imprimo el reporte/////
'///////imprimo el reporte/////
 Debito = 0
 Credito = 0
 TotalDebito = 0
 TotalCredito = 0
      NumFecha1 = Me.DTFecha1.Value
      NumFecha2 = Me.DTFecha2.Value
      If Not Me.TxtTransaccion = "" Then
       NMovimiento = Val(Me.TxtTransaccion.Text)
       NumeroTransaccion = NMovimiento
       Me.DtaConsulta.RecordSource = "SELECT Transacciones.ChequeNo,Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, TCambio*Debito AS MDebito, TCambio*Credito AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion) Between '" & Format(Me.DTFecha1, "yyyymmdd") & "' And '" & Format(Me.DTFecha2, "yyyymmdd") & "') AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
       Me.DtaConsulta.Refresh
      Else
       Exit Sub
      End If
      Do While Not Me.DtaConsulta.Recordset.EOF
      If Not IsNull(Me.DtaConsulta.Recordset("Debito")) Then
       Debito = Me.DtaConsulta.Recordset("Debito")
      End If
      If Not IsNull(Me.DtaConsulta.Recordset("Credito")) Then
       Credito = Me.DtaConsulta.Recordset("Credito")
      End If
 
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       Me.DtaConsulta.Recordset.MoveNext
      Loop
  
       Me.DtaConsulta.RecordSource = "SELECT Transacciones.ChequeNo, Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, TCambio*Debito AS MDebito, TCambio*Credito AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.VoucherNo, Transacciones.ChequeNo From Transacciones WHERE (((Transacciones.FechaTransaccion) Between '" & Format(Me.DTFecha1, "yyyymmdd") & "' And '" & Format(Me.DTFecha2, "yyyymmdd") & "') AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & ") AND ((Transacciones.ChequeNo) Is Not Null))"
       Me.DtaConsulta.Refresh
       If Not DtaConsulta.Recordset.EOF Then
        If Not IsNull(Me.DtaConsulta.Recordset("ChequeNo")) Then
            If Not IsNumeric(Me.DtaConsulta.Recordset("ChequeNo")) Then
                MsgBox "Cheque con No inválido: " & Me.DtaConsulta.Recordset("ChequeNo"), vbInformation
            Else
                CkNo = Me.DtaConsulta.Recordset("ChequeNo")
            End If
        End If
       End If
 
 ArepCheque.LblChequeNo.Caption = CkNo
 ArepCheque.Field15.Visible = False
 Sql = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between '" & Format(Me.DTFecha1, "yyyymmdd") & "' And '" & Format(Me.DTFecha2, "yyyymmdd") & "') AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
' ArepCheque.Show 1

            Set rpt = New ArepCheque
            rpt.DtaCheque.ConnectionString = ConexionReporte
            rpt.DtaCheque.Source = Sql
            fPreview.RunReport rpt
            fPreview.Show 1
            
  
                    

   Case "PRESUPUESTO ANUAL"
   
   Dim SqlString As String
   
   If Me.Option1.Value = True Then
'///////Busco fecha Inicial///////////////////
    DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1))"
    Me.DtaConsulta.Refresh
    If Not DtaConsulta.Recordset.EOF Then
      Fecha1 = Me.DtaConsulta.Recordset("FechaPeriodo")
      Tabla = 1
    End If
'////////////Busco fecha final para la Opcion 1//////////////////
    DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 12) And ((Periodos.NumeroTabla) = 1))"
    Me.DtaConsulta.Refresh
    If Not DtaConsulta.Recordset.EOF Then
      Fecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
    End If
  End If
  
    If Me.Option2.Value = True Then
'///////Busco fecha Inicial///////////////////
    DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 2))"
    Me.DtaConsulta.Refresh
    If Not DtaConsulta.Recordset.EOF Then
      Fecha1 = Me.DtaConsulta.Recordset("FechaPeriodo")
      Tabla = 2
    End If
'////////////Busco fecha final para la Opcion 1//////////////////
    DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 12) And ((Periodos.NumeroTabla) = 2))"
    Me.DtaConsulta.Refresh
    If Not DtaConsulta.Recordset.EOF Then
      Fecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
    End If
   End If
   
    If Me.Option3.Value = True Then
'///////Busco fecha Inicial///////////////////
    DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 3))"
    Me.DtaConsulta.Refresh
    If Not DtaConsulta.Recordset.EOF Then
      Fecha1 = Me.DtaConsulta.Recordset("FechaPeriodo")
      Tabla = 3
    End If
'////////////Busco fecha final para la Opcion 1//////////////////
    DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 12) And ((Periodos.NumeroTabla) = 3))"
    Me.DtaConsulta.Refresh
    If Not DtaConsulta.Recordset.EOF Then
      Fecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
    End If
 End If
    NumFecha1 = Fecha1
'    NumFecha2 = Fecha2
    
     ArepPresupuestoGeneral.DataControl1.ConnectionString = ConexionReporte
     ArepPresupuestoGeneral.Label13.Caption = "PRESUPUESTO PARA EL AÑO " & Year(Fecha1)
     SqlString = "SELECT  DescripcionGrupo, KeyGrupo,  ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 1)),0) As Enero , ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 2)),0) As Febrero, ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 3)),0) As Marzo, ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN  Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 4)),0) As Abril, ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 5)),0) As Mayo, ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 6)),0) As Junio, ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 7)),0) As Julio, ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 8)),0) As Agosto, ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 9)),0) As Septiembre, ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo  " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 10)),0) As Octubre, ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo  " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 11)),0) As Noviembre, ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo  " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 12)),0) As Diciembre, ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 1)),0) + ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 2)),0) + ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 3)),0) + ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN  Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 4)),0) + ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 5)),0) + ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 6)),0) + ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 7)),0) + ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 8)),0) + ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 9)),0) + ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo  " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 10)),0) + ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo  " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 11)),0) + ISNULL((SELECT Presupuesto.MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo  " & _
                 "WHERE (Presupuesto.CodCuenta =  dbo.EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & Tabla & ") AND (MONTH(Periodos.FechaPeriodo) = 12)),0) As Acumulado From EstructuraPresupuesto ORDER BY KeyGrupo"
'
'     ArepPresupuestoGeneral.DataControl1.ConnectionString = ConexionReporte
'     ArepPresupuestoGeneral.Label13.Caption = "PRESUPUESTO PARA EL AÑO " & Year(Fecha1)
'    ArepPresupuesto.DataControl1.Source = SqlString
 
            Set rpt = New ArepPresupuestoGeneral
            rpt.DataControl1.ConnectionString = ConexionReporte
            rpt.DataControl1.Source = SqlString
            fPreview.RunReport rpt
            fPreview.Show 1
 
 
  Case "PRESUPUESTO DEL MES"
  
   Dim DescripcionPresupuesto As String, MontoPresupuesto As Double, MontoReal As Double, TipoCuenta As String
   Dim MontoDebito As Double, MontoCredito As Double, PorcientoPresupuesto As Double
   Dim KeyGrupo As String
   
    Ejecutar.Execute "delete from reportes"
   
    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
       If Not Me.DtaConsulta.Recordset.EOF Then
         Me.DtaConsulta.Recordset.MoveLast
         i = Me.DtaConsulta.Recordset.RecordCount
         Me.DtaConsulta.Recordset.MoveFirst
       End If
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
      
      
      '///////////////////////////////////////RECORRO LAS UNIDADES PRESUPUESTARIAS //////
       Sql = "SELECT DescripcionGrupo, KeyGrupo, ISNULL ((SELECT  SUM(Presupuesto.MontoPresupuestado)As MontoPresupuestado FROM Presupuesto INNER JOIN Periodos ON Presupuesto.NPeriodo = Periodos.NPeriodo WHERE (Presupuesto.CodCuenta = EstructuraPresupuesto.KeyGrupo) AND (Periodos.NumeroTabla = " & NumeroTabla & ") AND (MONTH(Periodos.FechaPeriodo) BETWEEN  " & NumeroPeriodo1 & " AND " & NumeroPeriodo2 & ")), 0) AS Mes From EstructuraPresupuesto ORDER BY KeyGrupo"
       Me.AdoConsultas.RecordSource = Sql
       Me.AdoConsultas.Refresh
       Me.osProgress1.Visible = True
       Me.osProgress1.Min = 0
       Me.osProgress1.Value = 0
       If Not Me.AdoConsultas.Recordset.EOF Then
            Me.AdoConsultas.Recordset.MoveLast
            Me.osProgress1.Max = Me.AdoConsultas.Recordset.RecordCount
            Me.AdoConsultas.Recordset.MoveFirst
        Else
            Me.osProgress1.Max = 0
        End If
       
       Do While Not AdoConsultas.Recordset.EOF
         DescripcionPresupuesto = Me.AdoConsultas.Recordset("DescripcionGrupo")
         MontoPresupuesto = Me.AdoConsultas.Recordset("Mes")
         KeyGrupo = Me.AdoConsultas.Recordset("KeyGrupo")
         
         Me.lblProgreso.Caption = DescripcionPresupuesto
         DoEvents
         
         If KeyGrupo = "A0103" Then
           KeyGrupo = "A0103"
         End If
         
         '///////////////////////////CONSULTO EL MONTO REAL PARA ESTE GRUPO //////////////////////////////
         SqlString = "SELECT  SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * ISNULL(Tasas.MontoCordobas, 1) END) AS Debito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * ISNULL(Tasas.MontoCordobas, 1) END) AS Credito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / ISNULL(Tasas.MontoCordobas, 0) END) AS MDebito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / ISNULL(Tasas.MontoCordobas, 0) END) AS MCredito, Transacciones.KeyPresupuesto,  Cuentas.TipoCuenta FROM  Transacciones INNER JOIN Periodos ON Transacciones.NPeriodo = Periodos.NPeriodo INNER JOIN  IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN " & _
                     "Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas WHERE (Periodos.NumeroTabla = " & NumeroTabla & ") AND (Transacciones.FacturaNo <> '-') AND (Month(Periodos.FechaPeriodo) BETWEEN  " & NumeroPeriodo1 & " AND " & NumeroPeriodo2 & " ) GROUP BY Transacciones.KeyPresupuesto, Cuentas.TipoCuenta  HAVING  (Transacciones.KeyPresupuesto = '" & KeyGrupo & "') "
        
         MDIPrimero.AdoConsulta.RecordSource = SqlString
         MDIPrimero.AdoConsulta.Refresh
         MontoReal = 0
         Me.osProgress2.Min = 0
         Me.osProgress2.Value = 0
         Me.osProgress2.Visible = True
        If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
            MDIPrimero.AdoConsulta.Recordset.MoveLast
            Me.osProgress2.Max = MDIPrimero.AdoConsulta.Recordset.RecordCount
            MDIPrimero.AdoConsulta.Recordset.MoveFirst
        Else
            Me.osProgress2.Max = 0
        End If
        
         
         '/////////////////////////////////////RECORRO LOS REGISTROS PARA ESTE GRUPO ///////////////////
         MontoReal = 0
       
         Do While Not MDIPrimero.AdoConsulta.Recordset.EOF
             Resultado = 0
             If Me.CmbMoneda.Text = "Córdobas" Then
               MontoDebito = MDIPrimero.AdoConsulta.Recordset("Debito")
               MontoCredito = MDIPrimero.AdoConsulta.Recordset("Credito")
             ElseIf Me.CmbMoneda.Text = "Dólares" Then
               MontoDebito = MDIPrimero.AdoConsulta.Recordset("MDebito")
               MontoCredito = MDIPrimero.AdoConsulta.Recordset("MCredito")
             End If
            
            '
            
             TipoCuenta = MDIPrimero.AdoConsulta.Recordset("TipoCuenta")
             'If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
             If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                Resultado = MontoDebito - MontoCredito
             Else
                Resultado = MontoCredito - MontoDebito
             End If
             
             'If TipoCuenta = "Ingresos - Ventas" Then
         
             MontoReal = Abs(MontoReal + Resultado)
            
             Me.osProgress2.Value = Me.osProgress2.Value + 1
             MDIPrimero.AdoConsulta.Recordset.MoveNext
         Loop
         
         PorcientoPresupuesto = 0
         If MontoPresupuesto <> 0 Then
           PorcientoPresupuesto = MontoReal / MontoPresupuesto
         End If
         
         Resultado = MontoPresupuesto - MontoReal
       
       
         '/////////////////////////////////////////////////////////////////////////////////
         '///////////////////////AGREGO LOS REGISTROS PARA REPORTES DEL GRUPO///////////////////////
         '//////////////////////////////////////////////////////////////////////////////////////
           Me.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & KeyGrupo & "'))"
           Me.DtaConsulta.Refresh
           If Me.DtaConsulta.Recordset.EOF Then
             Me.DtaConsulta.Recordset.AddNew
               Me.DtaConsulta.Recordset("Descripcion") = DescripcionPresupuesto
               Me.DtaConsulta.Recordset("Debe1") = MontoReal
               Me.DtaConsulta.Recordset("Haber1") = MontoPresupuesto
               Me.DtaConsulta.Recordset("Debe2") = PorcientoPresupuesto
               Me.DtaConsulta.Recordset("Haber2") = Resultado
               Me.DtaConsulta.Recordset("KeyGrupo") = KeyGrupo
             Me.DtaConsulta.Recordset.Update
           
           End If
       
       
       
         Me.osProgress1.Value = Me.osProgress1.Value + 1
         Me.AdoConsultas.Recordset.MoveNext
       Loop
       
      
     ArepPresupuestoArea.FechaIni = FechaIni
     ArepPresupuestoArea.FechaFin = FechaFin
     ArepPresupuestoArea.DataControl1.ConnectionString = ConexionReporte
     ArepPresupuestoArea.Label13.Caption = "PRESUPUESTO PARA EL AÑO " & Year(Fecha1)
     Sql = "SELECT Reportes.Descripcion, Reportes.Debe1 As Real, Reportes.Haber1 As Presupuesto, Reportes.Debe2 As Porciento, Reportes.Haber2 As Resultado, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes ORDER BY Descripcion"
     ArepPresupuestoArea.DataControl1.Source = Sql
'     ArepPresupuestoArea.Logo.Picture = LoadPicture(RutaLogo)
'     ArepPresupuestoArea.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
'     ArepPresupuestoArea.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
'     ArepPresupuestoArea.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
'     ArepPresupuestoArea.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
            Set rpt = New ArepPresupuestoArea
            rpt.DataControl1.ConnectionString = ConexionReporte
            rpt.DataControl1.Source = Sql
            rpt.FechaIni = FechaIni
            rpt.FechaFin = FechaFin
            fPreview.RunReport rpt
            fPreview.Show 1
            
            

Case "AUXILIAR DE CUENTAS"

    GenerarAuxiliarDeCuentas True, False
'    ActivarBeneficiarioReporte = False
''    GenerarAuxiliarDeCuentas
'
'    GenerarAuxiliarDeCuentas True     ' detallado
''    GenerarAuxiliarDeCuentas False    ' totalizado


    

  Case "REGISTRO DE MOVIMIENTOS"
    ArepTransacciones.LblFecha = Format(Now, "dd/mm/yyyy")
    ArepTransacciones.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
    NumFecha1 = Me.DTFecha1.Value
    NumFecha2 = Me.DTFecha2.Value
    ArepTransacciones.DataControl1.ConnectionString = ConexionReporte
    Sql = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FacturaNo , Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
          "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.NumeroMovimiento, Transacciones.NTransaccion"

     Set rpt = New ArepTransacciones
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = Sql
     fPreview.RunReport rpt


     fPreview.Show 1


   Case "TARJETA CONTRATISTA"
    ArepContratista.DataControl1.ConnectionString = ConexionReporte
    'ArepContratista.DataControl1.Source = "SELECT Usuarios.CodUsuario, Usuarios.NombreUsuario, Usuarios.Nivel FROM Usuarios"

    ArepContratista.Logo.Picture = LoadPicture(RutaLogo)
    ArepContratista.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepContratista.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepContratista.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepContratista.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
'    ArepContratista.Show 1
           fPreview.arv.ReportSource = ArepContratista
           fPreview.Show 1

   Case "LISTA DE USUARIOS"
    ArepUsuarios.DataControl1.ConnectionString = ConexionReporte
    ArepUsuarios.DataControl1.Source = "SELECT Usuarios.CodUsuario, Usuarios.NombreUsuario, Usuarios.Nivel FROM Usuarios"

    ArepUsuarios.Logo.Picture = LoadPicture(RutaLogo)
    ArepUsuarios.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepUsuarios.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepUsuarios.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepUsuarios.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
'    ArepUsuarios.Show 1
           fPreview.arv.ReportSource = ArepUsuarios
           fPreview.Show 1


   Case "TASAS DE CAMBIO"
    NumFecha1 = Me.DTFecha1.Value
    NumFecha2 = Me.DTFecha2.Value
    ArepTasaCambio.DataControl1.ConnectionString = ConexionReporte
    ArepTasaCambio.DataControl1.Source = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas WHERE (((Tasas.FechaTasas) Between " & NumFecha1 & " And " & NumFecha2 & ")) ORDER BY Tasas.FechaTasas;"
    ArepTasaCambio.Logo.Picture = LoadPicture(RutaLogo)
    ArepTasaCambio.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepTasaCambio.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepTasaCambio.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepTasaCambio.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepTasaCambio.Show 1

   Case "GRUPOS DE CUENTAS"
    ArepGrupos.DataControl1.ConnectionString = ConexionReporte
    ArepGrupos.DataControl1.Source = "SELECT GrupoCuentas.CodGrupo, GrupoCuentas.DescripcionGrupo FROM GrupoCuentas"
    ArepGrupos.Logo.Picture = LoadPicture(RutaLogo)
    ArepGrupos.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
     ArepGrupos.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepGrupos.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepGrupos.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
'    ArepGrupos.Show 1
           fPreview.arv.ReportSource = ArepGrupos
           fPreview.Show 1
   
   Case "TARJETA EMPLEADOS"
     ArepEncargado.DataControl1.ConnectionString = ConexionReporte
      ArepEncargado.DataControl1.Source = "SELECT Encargado.CodEncargado, Encargado.NombreEncargado, Encargado.Direccion, Encargado.Telefono, Encargado.Cargo, Encargado.Email, Encargado.CP, Encargado.Fax, Encargado.FechaContratacion From Encargado"
      ArepEncargado.Logo.Picture = LoadPicture(RutaLogo)
      ArepEncargado.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
      ArepEncargado.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
      ArepEncargado.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
      ArepEncargado.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
'      ArepEncargado.Show 1
           fPreview.arv.ReportSource = ArepEncargado
           fPreview.Show 1

      Case "LISTA ACTIVO FIJO"
      ArepListadoAF.DataControl1.ConnectionString = ConexionReporte
      ArepListadoAF.DataControl1.Source = "SELECT CodCuenta, DescripcionAF, marca, modelo, fcompragen, costogen, ivagen, ValorRescate, DepreciacionAcumulada From CatalogoActivoFijo Order by CodCuenta"
      ArepListadoAF.Logo.Picture = LoadPicture(RutaLogo)
      ArepListadoAF.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
      ArepListadoAF.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
      ArepListadoAF.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
      ArepListadoAF.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
'      ArepActivoFijo.Show 1
           fPreview.arv.ReportSource = ArepListadoAF
           fPreview.Show 1
           
   Case "TARJETA ACTIVO FIJO"
      ArepActivoFijo.DataControl1.ConnectionString = ConexionReporte
      'ArepActivoFijo.DataControl1.Source = "SELECT ActivoFijo.CodCuenta, ActivoFijo.Localizacion, ActivoFijo.NumeroMarbete, ActivoFijo.FechaCompra, ActivoFijo.FechaCompra, ActivoFijo.ValorOriginal, ActivoFijo.FechaUltimaDepre, ActivoFijo.ValorEstimadoMeses, ActivoFijo.ValorRescate, ActivoFijo.FechaBaja, ActivoFijo.DepreciacionAcumulada, ActivoFijo.DescripcionActivo, ActivoFijo.NumeroSerie, ActivoFijo.Marca, Encargado.NombreEncargado FROM Encargado INNER JOIN ActivoFijo ON Encargado.CodEncargado = ActivoFijo.CodEncargado"
      ArepActivoFijo.DataControl1.Source = "SELECT CatalogoActivoFijo.CodCuenta, CatalogoActivoFijo.DescripcionAF, CatalogoActivoFijo.marca, CatalogoActivoFijo.modelo, CatalogoActivoFijo.fcompragen, CatalogoActivoFijo.costogen, CatalogoActivoFijo.ivagen, CatalogoActivoFijo.ValorRescate, CatalogoActivoFijo.DepreciacionAcumulada, CatalogoActivoFijo.Serie, CatalogoActivoFijo.NumeroMarbete, CatalogoActivoFijo.ValorEstimadoMeses, CatalogoActivoFijo.FechaUltimaDepre, CatalogoActivoFijo.Localizacion , CatalogoActivoFijo.FechaBaja, Encargado.NombreEncargado FROM  CatalogoActivoFijo INNER JOIN Encargado ON CatalogoActivoFijo.CodEncargado = Encargado.CodEncargado ORDER BY CatalogoActivoFijo.CodCuenta"
      ArepActivoFijo.Logo.Picture = LoadPicture(RutaLogo)
      ArepActivoFijo.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
      ArepActivoFijo.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
      ArepActivoFijo.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
      ArepActivoFijo.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
'      ArepActivoFijo.Show 1
           fPreview.arv.ReportSource = ArepActivoFijo
           fPreview.Show 1
      
  
    Case "LISTADO CUENTAS"

       ArepCuentas.DataControl1.ConnectionString = ConexionReporte
       ArepCuentas.Logo.Picture = LoadPicture(RutaLogo)
       ArepCuentas.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
       ArepCuentas.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
       ArepCuentas.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
       ArepCuentas.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
'       ArepCuentas.Show 1
           fPreview.arv.ReportSource = ArepCuentas
           fPreview.Show 1
'InputBox "", "", ArepCuentas.DataControl1.Source

    Case "CONTROL DE BANCOS"
      
      CodigoCuenta = Me.DBCodigo.Text
     
      NumFecha1 = Me.DTFecha1
      NumFecha2 = Me.DTFecha2
       'Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
        Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(Me.DTFecha1, "yyyymmdd") & "' And '" & Format(Me.DTFecha2, "yyyymmdd") & "'))"
       'Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, Transacciones.Debito, Transacciones.Credito, 5000+[Transacciones]![Debito]-[Transacciones]![Credito] AS Balance, Transacciones.NPeriodo From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.CodCuentas)='" & CodigoCuenta & "'))"
       Me.DtaConsulta.Refresh
  If Not DtaConsulta.Recordset.EOF Then
         Periodo1 = Me.DtaConsulta.Recordset("NPeriodo")
         Periodo1 = Periodo1 - 1
        NumFecha1 = Me.DTFecha1.Value
        NumFecha2 = Me.DTFecha2.Value
'///////////////Busco el Acumulado de la cuenta hasta la ultima fecha Seleccionada////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
       
         If Me.CmbMoneda.Text = "Córdobas" Then
'            Me.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(ROUND(Debito*TCambio,2)) AS MDebito, Sum(ROUND(TCambio*Credito,2)) AS MCredito From Transacciones Where (((Transacciones.FechaTransaccion) < '" & Format(Me.DTFecha1, "yyyymmdd") & "')) GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CodigoCuenta & "'))"
            Me.DtaHistorial.RecordSource = "SELECT Cuentas.CodCuentas, Transacciones.FechaTransaccion, ROUND(SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5)),5) AS MDebito, ROUND(SUM(ROUND(Transacciones.TCambio * Transacciones.Credito,5)),5) AS MCredito, ROUND(SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5) - ROUND(Transacciones.TCambio * Transacciones.Credito,5)),5) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Tasas.MontoCordobas) AS MontoCordobas, MAX(Tasas.MontoLibras) AS MontoLibras, MAX(Transacciones.NTransaccion) AS NTransaccion FROM  Tasas INNER JOIN  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Tasas.FechaTasas = Transacciones.FechaTasas GROUP BY Cuentas.CodCuentas, Transacciones.FechaTransaccion  " & _
                                           "HAVING (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.FechaTransaccion < CONVERT(DATETIME,'" & Format(FechaIni, "yyyymmdd") & "', 102)) ORDER BY Cuentas.CodCuentas"
            Me.DtaHistorial.Refresh
         Else
'            Me.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas),2)) AS MDebito,SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito,2)) As MCredito FROM Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion < '" & Format(Me.DTFecha1, "yyyymmdd") & "') GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & CodigoCuenta & "')"
            Me.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, ROUND(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas), 5) AS MDebito, ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 5) As MCredito FROM  Transacciones INNER JOIN  Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                           "WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "')"
            Me.DtaHistorial.Refresh
         End If
         
         TotalDebito = 0
         TotalCredito = 0
         SaldoIni = 0
            Do While Not Me.DtaHistorial.Recordset.EOF
                     
                         If Not IsNull(DtaHistorial.Recordset("MDebito")) Then
                             Debito = DtaHistorial.Recordset("MDebito")
                             Debito = TRUNC(Debito, 3)
                             Debito = Format(Debito, "##,##0.00")
                             TotalDebito = Debito + TotalDebito
                        End If
                        If Not IsNull(DtaHistorial.Recordset("MCredito")) Then
                             Credito = DtaHistorial.Recordset("MCredito")
                             Credito = TRUNC(Credito, 3)
                             Credito = Format(Credito, "##,##0.00")
                             TotalCredito = Credito + TotalCredito
                        End If
                     
                 Me.DtaHistorial.Recordset.MoveNext
             Loop

          SaldoIni = TotalDebito - TotalCredito
         
'         If Not DtaHistorial.Recordset.EOF Then
'           If Not IsNull(Me.DtaHistorial.Recordset("MDebito")) Then
'            Debito = Me.DtaHistorial.Recordset("MDebito")
'           End If
'           If Not IsNull(Me.DtaHistorial.Recordset("MCredito")) Then
'             Credito = Me.DtaHistorial.Recordset("MCredito")
'           End If
'           Total = Debito - Credito
'           SaldoIni = Total
'         Else
'          SaldoIni = 0
'         End If
          
           
  Else
'///////////////Busco el Acumulado de la cuenta hasta la ultima fecha Seleccionada////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
         If Me.CmbMoneda.Text = "Córdobas" Then
'          Me.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones WHERE (FechaTransaccion < CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CodigoCuenta & "'))"
'          Me.DtaHistorial.Refresh
            Me.DtaHistorial.RecordSource = "SELECT Cuentas.CodCuentas, Transacciones.FechaTransaccion, ROUND(SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5)),5) AS MDebito, ROUND(SUM(ROUND(Transacciones.TCambio * Transacciones.Credito,5)),5) AS MCredito, ROUND(SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5) - ROUND(Transacciones.TCambio * Transacciones.Credito,5)),5) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Tasas.MontoCordobas) AS MontoCordobas, MAX(Tasas.MontoLibras) AS MontoLibras, MAX(Transacciones.NTransaccion) AS NTransaccion FROM  Tasas INNER JOIN  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Tasas.FechaTasas = Transacciones.FechaTasas GROUP BY Cuentas.CodCuentas, Transacciones.FechaTransaccion  " & _
                                           "HAVING (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.FechaTransaccion < CONVERT(DATETIME,'" & Format(FechaIni, "yyyymmdd") & "', 102)) ORDER BY Cuentas.CodCuentas"
            Me.DtaHistorial.Refresh
         Else
'            Me.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas), 2)) AS MDebito, SUM(Round(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) As MCredito FROM Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion < '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "') GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "')"
            Me.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, ROUND(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas), 5) AS MDebito, ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 5) As MCredito FROM  Transacciones INNER JOIN  Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                           "WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "')"
            Me.DtaHistorial.Refresh
         End If


         TotalDebito = 0
         TotalCredito = 0
         SaldoIni = 0
            Do While Not Me.DtaHistorial.Recordset.EOF
                     
                         If Not IsNull(DtaHistorial.Recordset("MDebito")) Then
                             Debito = DtaHistorial.Recordset("MDebito")
                             Debito = TRUNC(Debito, 3)
                             Debito = Format(Debito, "##,##0.00")
                             TotalDebito = Debito + TotalDebito
                        End If
                        If Not IsNull(DtaHistorial.Recordset("MCredito")) Then
                             Credito = DtaHistorial.Recordset("MCredito")
                             Credito = TRUNC(Credito, 3)
                             Credito = Format(Credito, "##,##0.00")
                             TotalCredito = Credito + TotalCredito
                        End If
                     
                 Me.DtaHistorial.Recordset.MoveNext
             Loop

          SaldoIni = TotalDebito - TotalCredito


  End If
  

       Me.AdoConsultas.RecordSource = "SELECT  * From Conciliacion WHERE  (FechaConciliacion <= CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) AND (CodCuenta = '" & Me.DBCodigo.Text & "') AND (Activo = 1)"
       Me.AdoConsultas.Refresh
        If Not Me.AdoConsultas.Recordset.EOF Then
            ArepBank.LblSaldoEstadoCuenta.Caption = Format(Me.AdoConsultas.Recordset("SaldoEstadoCuenta"), "##,##0.00")
        Else
            ArepBank.LblSaldoEstadoCuenta.Caption = "0.00"
        End If
  
       
       ArepBank.DtaBanco.ConnectionString = ConexionReporte
       ArepBank.LblBalance = Format(SaldoIni, "##,##0.00")
       ArepBank.LblPaidIn.Caption = Format(SaldoIni, "##,##0.00")
       ArepBank.LblTipo.Caption = "CUENTA DE BANCO: " & Me.DBCodigo
       ArepBank.LblFecha1 = Me.DTFecha1
       ArepBank.LblFecha2 = Me.DTFecha2
       ArepBank.DtaBanco.ConnectionString = ConexionReporte
       
       SaldoIniBank = Format(SaldoIni, "##,##0.00")
       CodigoBancoBank = Me.DBCodigo.Text
       FechaIniBank = Me.DTFecha1.Value
       FechaFinBank = Me.DTFecha2.Value
       
       
       If Me.CmbMoneda.Text = "Córdobas" Then
       'ArepBank.DtaBanco.Source
         Sql = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, TCambio*Debito AS Debito, TCambio*Credito AS Credito, " & SaldoIni & "+Transacciones.Debito-Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento,Beneficiario From Transacciones WHERE (((Transacciones.FechaTransaccion) Between '" & Format(Me.DTFecha1, "yyyymmdd") & "' And '" & Format(Me.DTFecha2, "yyyymmdd") & "') AND ((Transacciones.CodCuentas)='" & CodigoCuenta & "')) ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
       Else
       
       'ArepBank.DtaBanco.Source
         Sql = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo,Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito,2) AS Debito, ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2) AS Credito,ROUND((2750.03 + Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)) - Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas), 2) AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento , Transacciones.Beneficiario FROM Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(Me.DTFecha1, "yyyymmdd") & "' And '" & Format(Me.DTFecha2, "yyyymmdd") & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
       End If
       ArepBank.Logo.Picture = LoadPicture(RutaLogo)
       ArepBank.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
       ArepBank.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
       ArepBank.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
       ArepBank.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
       ArepBank.LblMoneda.Caption = Me.CmbMoneda.Text
'       ArepBank.Show 1


        Set rpt = New ArepBank
        rpt.DtaBanco.ConnectionString = ConexionReporte
        rpt.DtaBanco.Source = Sql
        fPreview.RunReport rpt
         fPreview.Show 1

           
           

    
End Select
Me.Frame6.Visible = False
Me.CmdVerReporte.Enabled = True
Me.CmdSalir.Enabled = True
Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub



Private Sub CmdVerReporte2_Click()
Dim Fechas1 As String, Fechas2 As String, Orden As Integer, Sql As String, i As Double
Dim UltimoOrden As Integer, RegIngresos  As Integer, PrimReg As Integer, UltReg As Integer
Dim Utilidad As Double, Utilidad2 As Double, Utilidad3 As Double, RegTCostosOper As Integer
Dim Decrementador As Integer, TotalActivoCirculante As Double, TotalActivoFijo As Double, TotalActivoDiferido As Double
Dim TotalPasivoCirculante As Double, TotalPasivoFijo As Double, TotalPasivoDiferido As Double, TotalCapitalSocial As Double
Dim RegInicioCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro donde comienzan los costos operativos
Dim RegTotalCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro de total de costos operativos
Dim Totalingresos As Double, TotalCostoVentas As Double, TotalGastosAdmon As Double, TotalGastos As Double
Dim TotalGastoVentas As Double, TotalIngresosFinancieros As Double, TotalOtrosIngresos As Double, TotalOtrosGastos As Double
Dim TotalUtilidadBruta As Double, TotalImpuestos As Double, TotalUtilidadNeta As Double, Fecha1 As String, Fecha2 As String
Dim TotalCompras As Double, TotalInventarioInicial As Double, TotalInventarioFinal As Double
Dim TotalAcarreo As Double, TotalRebajaVentas As Double, TotalDisponible As Double, TotalGastosR As Double, TotalCosto As Double
Dim TotalSalidas As Double, TotalGastoOperacion As Double, TotalPasivo As Double, TotalCapital As Double
Dim TotalCostos As Double, ListaActivos As Variant, TotalInventario As Double, TotalCuentaxCobrar As Double
Dim TotalCuentasxPagar As Double, TotalActivos As Double, UtilidadBrutas As Double, UtilidadNetas As Double
Dim ListaMeses As Variant, CantRegistros As Double, ComboIni As Double, ComboFin As Double, TotalCostoFijo As Double, TotalGastoFijo As Double
Dim mes As Double, R As Variant
Dim rpt As Object
Dim fPreview As New FrmPreview
Dim Rs As New ADODB.Recordset
Dim IngresosVentas, ServiciosVentas, ComisionVentas As Double
Dim CostosProduccion, CostosGeneralesProduccion As Double
Dim Parche As ADODB.Connection


On Error GoTo TipoErrs
Me.Frame6.Visible = True
Me.CmdVerReporte2.Enabled = False
Me.CmdSalir.Enabled = False
SaldoIni = 0
SaldoFin = 0
Total1 = 0
TotalCuenta = 0

FechaIni = Me.DTFecha1.Value
FechaFin = Me.DTFecha2.Value




Select Case Me.CmbReportes.Text

Case "RETENCIONES EN LA FUENTE I.R X COBRAR"
            NumeroPeriodo1 = Me.CmbIni.Text
            NumeroPeriodo2 = Me.CmbFin.Text
            
            If Me.Option8 = True Then
             NumeroTabla = 1
            ElseIf Me.Option7 = True Then
              NumeroTabla = 2
            ElseIf Me.Option6 = True Then
              NumeroTabla = 3
            End If
            
              Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) = " & NumeroPeriodo1 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
              Me.DtaConsulta.Refresh
              If Me.DtaConsulta.Recordset.RecordCount = 0 Then
                MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                Exit Sub
              End If
               Me.DtaConsulta.Recordset.MoveLast
               i = Me.DtaConsulta.Recordset.RecordCount
               Me.DtaConsulta.Recordset.MoveFirst
              Do While Not DtaConsulta.Recordset.EOF
        
        
                If i = 1 Then
                  FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                  FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                Else
        
                 If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
                  FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                 ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
                  FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                 End If
                End If
                Me.DtaConsulta.Recordset.MoveNext
              Loop
              
              Fecha1 = Format(FechaIni, "yyyy-mm-dd")
              Fecha2 = Format(FechaFin, "yyyy-mm-dd")
             
              
              ListaMeses = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")


   ArepRetencionesIR.DataControl1.ConnectionString = ConexionReporte
   ArepRetencionesIR.LblRUC.Caption = MDIPrimero.AdoConfiguracion.Recordset("NumeroRuc")
   ArepRetencionesIR.LblRazonSocial.Caption = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa")
   ArepRetencionesIR.LblDireccion.Caption = MDIPrimero.AdoConfiguracion.Recordset("Direccion")
   ArepRetencionesIR.LblPeriodo.Caption = ListaMeses(NumeroPeriodo1 - 1) & " " & Year(FechaIni)
   ArepRetencionesIR.LblTelefono.Caption = MDIPrimero.AdoConfiguracion.Recordset("Telefono")
   ArepRetencionesIR.LblBalance.Caption = Me.CmbReportes.Text
   Sql = "SELECT Cuentas.RUC, Cuentas.Cedula, Cuentas.Apellido1, Cuentas.Apellido2, Cuentas.Nombre1 + ' ' + Cuentas.Nombre2 AS Nombres,Transacciones.DescripcionMovimiento, Transacciones.TCambio * Transacciones.Debito + Transacciones.TCambio * Transacciones.Credito AS Monto, Transacciones.CodCuentaProveedor, Cuentas.CausaRetencion, Cuentas.DescRetencion, Transacciones.TipoFactura, Cuentas.TipoCuenta, Cuentas.CodCuentas , Transacciones.FacturaNo, Transacciones.FechaTransaccion  " & _
         "FROM  Transacciones INNER JOIN  Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
         "WHERE (Cuentas.CausaRetencion = 1) AND (Cuentas.TipoCuenta = 'Cuentas x Cobrar') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) "
         ArepRetencionesIR.DataControl1.Source = Sql
         
         
     Set rpt = New ArepRetencionesIR
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = Sql
     fPreview.RunReport rpt

     fPreview.Show 1


Case "RETENCIONES EN LA FUENTE I.R X PAGAR"
            NumeroPeriodo1 = Me.CmbIni.Text
            NumeroPeriodo2 = Me.CmbFin.Text
            
            If Me.Option8 = True Then
             NumeroTabla = 1
            ElseIf Me.Option7 = True Then
              NumeroTabla = 2
            ElseIf Me.Option6 = True Then
              NumeroTabla = 3
            End If
            
              Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) = " & NumeroPeriodo1 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
              Me.DtaConsulta.Refresh
              If Me.DtaConsulta.Recordset.RecordCount = 0 Then
                MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                Exit Sub
              End If
               Me.DtaConsulta.Recordset.MoveLast
               i = Me.DtaConsulta.Recordset.RecordCount
               Me.DtaConsulta.Recordset.MoveFirst
              Do While Not DtaConsulta.Recordset.EOF
        
        
                If i = 1 Then
                  FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                  FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                Else
        
                 If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
                  FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                 ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
                  FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                 End If
                End If
                Me.DtaConsulta.Recordset.MoveNext
              Loop
              
              Fecha1 = Format(FechaIni, "yyyy-mm-dd")
              Fecha2 = Format(FechaFin, "yyyy-mm-dd")
             
              
              ListaMeses = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")

   ArepRetencionesIR.DataControl1.ConnectionString = ConexionReporte
   ArepRetencionesIR.LblRUC.Caption = MDIPrimero.AdoConfiguracion.Recordset("NumeroRuc")
   ArepRetencionesIR.LblRazonSocial.Caption = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa")
   ArepRetencionesIR.LblDireccion.Caption = MDIPrimero.AdoConfiguracion.Recordset("Direccion")
   ArepRetencionesIR.LblPeriodo.Caption = ListaMeses(NumeroPeriodo1 - 1) & " " & Year(FechaIni)
   ArepRetencionesIR.LblTelefono.Caption = MDIPrimero.AdoConfiguracion.Recordset("Telefono")
   ArepRetencionesIR.LblBalance.Caption = Me.CmbReportes.Text
   Sql = "SELECT Cuentas.RUC, Cuentas.Cedula, Cuentas.Apellido1, Cuentas.Apellido2, Cuentas.Nombre1 + ' ' + Cuentas.Nombre2 AS Nombres,Transacciones.DescripcionMovimiento, Transacciones.TCambio * Transacciones.Debito + Transacciones.TCambio * Transacciones.Credito AS Monto, Transacciones.CodCuentaProveedor, Cuentas.CausaRetencion, Cuentas.DescRetencion, Transacciones.TipoFactura, Cuentas.TipoCuenta, Cuentas.CodCuentas , Transacciones.FacturaNo, Transacciones.FechaTransaccion  " & _
         "FROM  Transacciones INNER JOIN  Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
         "WHERE (Cuentas.CausaRetencion = 1) AND (Cuentas.TipoCuenta = 'Cuentas x Pagar') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) "
         ArepRetencionesIR.DataControl1.Source = Sql
         
     Set rpt = New ArepRetencionesIR
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = Sql
     fPreview.RunReport rpt

     fPreview.Show 1

Case "ANEXO FISCAL IVA CLIENTES"
            NumeroPeriodo1 = Me.CmbIni.Text
            NumeroPeriodo2 = Me.CmbFin.Text
            
            If Me.Option8 = True Then
             NumeroTabla = 1
            ElseIf Me.Option7 = True Then
              NumeroTabla = 2
            ElseIf Me.Option6 = True Then
              NumeroTabla = 3
            End If
            
              Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) = " & NumeroPeriodo1 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
              Me.DtaConsulta.Refresh
              If Me.DtaConsulta.Recordset.RecordCount = 0 Then
                MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                Exit Sub
              End If
               Me.DtaConsulta.Recordset.MoveLast
               i = Me.DtaConsulta.Recordset.RecordCount
               Me.DtaConsulta.Recordset.MoveFirst
              Do While Not DtaConsulta.Recordset.EOF
        
        
                If i = 1 Then
                  FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                  FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                Else
        
                 If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
                  FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                 ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
                  FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                 End If
                End If
                Me.DtaConsulta.Recordset.MoveNext
              Loop
              
              Fecha1 = Format(FechaIni, "yyyy-mm-dd")
              Fecha2 = Format(FechaFin, "yyyy-mm-dd")
             
              
              ListaMeses = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")
              
              ArepCreditoProveedor.DataControl1.ConnectionString = ConexionReporte
              ArepCreditoProveedor.LblBalance.Caption = "ANEXO AL CREDITO FISCAL DECLARACION I.V.A FACTURA VENTA"
              ArepCreditoProveedor.LblProveedor.Caption = "CLIENTES"
              ArepCreditoProveedor.LblAño.Caption = Year(FechaIni)
              ArepCreditoProveedor.LblMes.Caption = ListaMeses(NumeroPeriodo1 - 1)
              ArepCreditoProveedor.LblFechaImpreso.Caption = Format(Now, "Long Date")
              ArepCreditoProveedor.LblRUC.Caption = "RUC: " & MDIPrimero.AdoConfiguracion.Recordset("NumeroRuc")
              ArepCreditoProveedor.LblCompañia.Caption = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa")


 Sql = "SELECT Transacciones.FechaTransaccion, Transacciones.FacturaNo, Cuentas.RUC,Cuentas.Nombre1 + ' ' + Cuentas.Nombre2 + ' ' + Cuentas.Apellido1 + ' ' + Cuentas.Apellido2 AS Nombres, Transacciones.DescripcionMovimiento, Transacciones.TCambio * Transacciones.Debito + Transacciones.TCambio * Transacciones.Credito AS SubTotal, Transacciones.TipoFactura, Cuentas.TipoCuenta, Cuentas.CausaIva, Cuentas.CausaRetencion, Cuentas.DescRetencion, Cuentas.CodCuentas, Transacciones.CodCuentaProveedor " & _
       "FROM Transacciones INNER JOIN  Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
       "WHERE (Transacciones.TipoFactura = 'FacturaVenta') AND (Cuentas.CausaIva <> 1) AND (Cuentas.CausaRetencion <> 1) AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
       "ORDER BY Transacciones.FechaTransaccion"
       
 ArepCreditoProveedor.DataControl1.Source = Sql


     Set rpt = New ArepCreditoProveedor
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = Sql
     fPreview.RunReport rpt
     fPreview.Show 1

Case "ANEXO FISCAL IVA PROVEEDOR"
            NumeroPeriodo1 = Me.CmbIni.Text
            NumeroPeriodo2 = Me.CmbFin.Text
            
            If Me.Option8 = True Then
             NumeroTabla = 1
            ElseIf Me.Option7 = True Then
              NumeroTabla = 2
            ElseIf Me.Option6 = True Then
              NumeroTabla = 3
            End If
            
              Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) = " & NumeroPeriodo1 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
              Me.DtaConsulta.Refresh
              If Me.DtaConsulta.Recordset.RecordCount = 0 Then
                MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                Exit Sub
              End If
               Me.DtaConsulta.Recordset.MoveLast
               i = Me.DtaConsulta.Recordset.RecordCount
               Me.DtaConsulta.Recordset.MoveFirst
              Do While Not DtaConsulta.Recordset.EOF
        
        
                If i = 1 Then
                  FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                  FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                Else
        
                 If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
                  FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                 ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
                  FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                 End If
                End If
                Me.DtaConsulta.Recordset.MoveNext
              Loop
              
              Fecha1 = Format(FechaIni, "yyyy-mm-dd")
              Fecha2 = Format(FechaFin, "yyyy-mm-dd")
             
              
              ListaMeses = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")
              
              ArepCreditoProveedor.DataControl1.ConnectionString = ConexionReporte
              ArepCreditoProveedor.LblBalance.Caption = "ANEXO AL CREDITO FISCAL DECLARACION I.V.A FACTURA COMPRA"
              ArepCreditoProveedor.LblProveedor.Caption = "PROVEEDOR"
              ArepCreditoProveedor.LblAño.Caption = Year(FechaIni)
              ArepCreditoProveedor.LblMes.Caption = ListaMeses(NumeroPeriodo1 - 1)
              ArepCreditoProveedor.LblFechaImpreso.Caption = Format(Now, "Long Date")
              ArepCreditoProveedor.LblRUC.Caption = "RUC: " & MDIPrimero.AdoConfiguracion.Recordset("NumeroRuc")
              ArepCreditoProveedor.LblCompañia.Caption = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa")


' SQL = "SELECT Transacciones.FechaTransaccion, Transacciones.FacturaNo, Cuentas.RUC,Cuentas.Nombre1 + ' ' + Cuentas.Nombre2 + ' ' + Cuentas.Apellido1 + ' ' + Cuentas.Apellido2 AS Nombres, Transacciones.DescripcionMovimiento, Transacciones.TCambio * Transacciones.Debito + Transacciones.TCambio * Transacciones.Credito AS SubTotal, Transacciones.TipoFactura, Cuentas.TipoCuenta, Cuentas.CausaIva, Cuentas.CausaRetencion, Cuentas.DescRetencion, Cuentas.CodCuentas, Transacciones.CodCuentaProveedor " & _
'       "FROM Transacciones INNER JOIN  Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
'       "WHERE (Transacciones.TipoFactura = 'FacturaCompra') AND (Cuentas.CausaIva <> 1) AND (Cuentas.CausaRetencion <> 1) AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
'       "ORDER BY Transacciones.FechaTransaccion"

Sql = "SELECT  Transacciones.FechaTransaccion, Transacciones.FacturaNo, Cuentas.RUC,Cuentas.Nombre1 + ' ' + Cuentas.Nombre2 + ' ' + Cuentas.Apellido1 + ' ' + Cuentas.Apellido2 AS Nombres, Transacciones.DescripcionMovimiento, Transacciones.TCambio * Transacciones.Debito + Transacciones.TCambio * Transacciones.Credito AS SubTotal, Transacciones.TipoFactura, Cuentas.TipoCuenta, Cuentas.CausaIva, Cuentas.CausaRetencion, Cuentas.DescRetencion, Cuentas.CodCuentas, Transacciones.CodCuentaProveedor " & _
      "FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
      "WHERE (Cuentas.CausaIva = 1) AND (Cuentas.CausaRetencion <> 1) AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "',102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) AND (Cuentas.TipoCuenta = 'Cuentas x Pagar')  ORDER BY Transacciones.TipoFactura DESC, Transacciones.FechaTransaccion"
       
 ArepCreditoProveedor.DataControl1.Source = Sql

     Set rpt = New ArepCreditoProveedor
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = Sql
     fPreview.RunReport rpt
     fPreview.Show 1


Case "COMPARATIVO INGRESOS VRS GASTOS"
     
    ListaMeses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    
     With ArepGraficosIngresos.Chart2D.ChartGroups(1).Data
      If Me.CmbFin.Text = Me.CmbIni.Text Then
         CantRegistros = 1
         ComboIni = Val(Me.CmbIni.Text)
         ComboFin = Val(Me.CmbFin.Text)
      Else
         CantRegistros = Val(Me.CmbFin.Text) - Val(Me.CmbIni.Text) + 1
         ComboIni = Val(Me.CmbIni.Text)
         ComboFin = Val(Me.CmbFin.Text)
      End If
        

        Me.osProgress1.Visible = True
        Me.osProgress1.Value = 0
        Me.osProgress1.Max = CantRegistros
        
         For i = 1 To CantRegistros
   
        Fecha2 = Format(FechaPeriodo(ComboIni), "yyyy-mm-dd")
        Fecha1 = Format(FechaPeriodoIni(ComboIni), "yyyy-mm-dd")
        If i = 1 Then
        
         If Me.OptAcumulado.Value = True Then
            ArepGraficosIngresos.Chart2D.Header.Text = "COMPARATIVO INGRESOS VRS COSTOS Y GASTOS - " & Year(Fecha1)
'
         Else
             ArepGraficosIngresos.Chart2D.Header.Text = "COMPARATIVO INGRESOS VRS COSTOS Y GASTOS - " & Year(Fecha1)
         End If
        End If
             If Me.OptAcumulado.Value = True Then
                '////////////////////////TOTAL INGRESOS//////////////////////////////////////////////////////////
                Totalingresos = 0
                Totalingresos = SaldosRazonesCreditos(Fecha2, "Ingresos - Ventas")
                
                '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalCosto = 0
                TotalCosto = SaldosRazonesDebitos(Fecha2, "Costos")

                 '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalGastos = 0
                TotalGastos = SaldosRazonesDebitos(Fecha2, "Gastos")
                
             ElseIf Me.OptPeriodo.Value = True Then
                '////////////////////////TOTAL INGRESOS//////////////////////////////////////////////////////////
                Totalingresos = 0
                Totalingresos = SaldoPeriodoCredito(Fecha1, Fecha2, "Ingresos - Ventas")

             
                '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalCosto = 0
                TotalCosto = SaldoPeriodoDebito(Fecha1, Fecha2, "Costos")

                 '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalGastos = 0
                TotalGastos = SaldoPeriodoDebito(Fecha1, Fecha2, "Gastos")

             
             End If
             
                '//////////////////////////SUMO LAS UTILIDADES////////////////////////////////////////////////////////
                UtilidadBrutas = 0
                UtilidadNetas = 0
                UtilidadBrutas = Totalingresos - TotalCosto
                UtilidadNetas = (Totalingresos - TotalCosto - TotalGastos) / 1000
                

         
            ArepGraficosIngresos.Chart2D.ChartGroups(1).Data.NumPoints(1) = CantRegistros
            ArepGraficosIngresos.Chart2D.ChartGroups(1).Data.NumPoints(2) = CantRegistros
            ArepGraficosIngresos.Chart2D.ChartGroups(1).Data.x(1, i) = i
            ArepGraficosIngresos.Chart2D.ChartGroups(1).Data.y(1, i) = TotalCosto + TotalGastos
            ArepGraficosIngresos.Chart2D.ChartGroups(1).Data.y(2, i) = Totalingresos
         
         ArepGraficosIngresos.MSFlexGrid.cols = CantRegistros + 1
         ArepGraficosIngresos.MSFlexGrid.ColWidth(i) = 1500
         ArepGraficosIngresos.MSFlexGrid.col = i
         ArepGraficosIngresos.MSFlexGrid.Text = Format(Totalingresos, "##,##0.00")
         ArepGraficosIngresos.MSFlexGrid.TextMatrix(0, i) = ListaMeses(ComboIni - 1)
         
         ArepGraficosIngresos.MSFlexGrid1.cols = CantRegistros + 1
         ArepGraficosIngresos.MSFlexGrid1.ColWidth(i) = 1500
         ArepGraficosIngresos.MSFlexGrid1.col = i
         ArepGraficosIngresos.MSFlexGrid1.Text = Format(TotalCosto + TotalGastos, "##,##0.00")
         ArepGraficosIngresos.MSFlexGrid1.TextMatrix(0, i) = ListaMeses(ComboIni - 1)
         
         Me.osProgress1.Value = Me.osProgress1.Value + 1
         DoEvents
         ComboIni = ComboIni + 1
         Next
    End With
          

        With ArepGraficosIngresos.Chart2D.ChartGroups(1).SeriesLabels
        .Add "Costos y Gastos"
        .Add "Ingresos"
        End With
                ArepGraficosIngresos.LblMoneda.Caption = "Expresado en " & Me.CmbMoneda.Text
                ArepGraficosIngresos.Logo.Picture = LoadPicture(RutaLogo)
                ArepGraficosIngresos.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                ArepGraficosIngresos.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                ArepGraficosIngresos.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                ArepGraficosIngresos.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
                ArepGraficosIngresos.LblFechaFin = FechaFin
                ArepGraficosIngresos.LblFechaIni = FechaIni
    
         ArepGraficosIngresos.Show 1
'                 fPreview.arv.ReportSource = ArepGraficosIngresos
'                 fPreview.Show 1
     
'     Set rpt = New ArepGraficosIngresos
''     rpt.DataControl1.ConnectionString = ConexionReporte
''     rpt.DataControl1.Source = SQL
'     fPreview.RunReport rpt
'     fPreview.Show 1

Case "COMPARATIVO UTILIDADES"
     
    
    ListaMeses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    
    With ArepGraficosUtilidades.MSChartUtilidades.DataGrid
      If Me.CmbFin.Text = Me.CmbIni.Text Then
         CantRegistros = 1
         ComboIni = Val(Me.CmbIni.Text)
         ComboFin = Val(Me.CmbFin.Text)
      Else
         CantRegistros = Val(Me.CmbFin.Text) - Val(Me.CmbIni.Text) + 1
         ComboIni = Val(Me.CmbIni.Text)
         ComboFin = Val(Me.CmbFin.Text)
      End If
        
        
        .RowCount = CantRegistros
        ArepGraficosUtilidades.MSChartIngresos.RowCount = CantRegistros
        ArepGraficosUtilidades.MSChartGastos.RowCount = CantRegistros
        
        Me.osProgress1.Visible = True
        Me.osProgress1.Value = 0
        Me.osProgress1.Max = CantRegistros
        
         For i = 1 To CantRegistros
       
        Fecha2 = Format(FechaPeriodo(ComboIni), "yyyy-mm-dd")
        Fecha1 = Format(FechaPeriodoIni(ComboIni), "yyyy-mm-dd")
        If i = 1 Then
        
         If Me.OptAcumulado.Value = True Then
            ArepGraficosUtilidades.MSChartUtilidades.TitleText = "UTILIDADES ACUMULADAS - " & Year(Fecha1)
            ArepGraficosUtilidades.MSChartIngresos.TitleText = "INGREOS ACUMULADOS - " & Year(Fecha1)
           ArepGraficosUtilidades.MSChartGastos.TitleText = "COSTOS Y GASTOS ACUMULADOS - " & Year(Fecha1)
           
         Else
            ArepGraficosUtilidades.MSChartUtilidades.TitleText = "UTILIDADES DEL PERIODO - " & Year(Fecha1)
            ArepGraficosUtilidades.MSChartIngresos.TitleText = "INGREOS DEL PERIODO - " & Year(Fecha1)
            ArepGraficosUtilidades.MSChartGastos.TitleText = "COSTOS Y GASTOS DEL PERIODO - " & Year(Fecha1)
            
         End If
        End If
             If Me.OptAcumulado.Value = True Then
                '////////////////////////TOTAL INGRESOS//////////////////////////////////////////////////////////
                Totalingresos = 0
                Totalingresos = SaldosRazonesCreditos(Fecha2, "Ingresos - Ventas")
                
                '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalCosto = 0
                TotalCosto = SaldosRazonesDebitos(Fecha2, "Costos")

                 '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalGastos = 0
                TotalGastos = SaldosRazonesDebitos(Fecha2, "Gastos")
                
             ElseIf Me.OptPeriodo.Value = True Then
                '////////////////////////TOTAL INGRESOS//////////////////////////////////////////////////////////
                Totalingresos = 0
                Totalingresos = SaldoPeriodoCredito(Fecha1, Fecha2, "Ingresos - Ventas")

             
                '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalCosto = 0
                TotalCosto = SaldoPeriodoDebito(Fecha1, Fecha2, "Costos")

                 '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalGastos = 0
                TotalGastos = SaldoPeriodoDebito(Fecha1, Fecha2, "Gastos")

             
             End If
             
                '//////////////////////////SUMO LAS UTILIDADES////////////////////////////////////////////////////////
                UtilidadBrutas = 0
                UtilidadNetas = 0
                UtilidadBrutas = Totalingresos - TotalCosto
                UtilidadNetas = (Totalingresos - TotalCosto - TotalGastos) / 1000

          
         .SetData i, 1, UtilidadNetas, 0
         .RowLabel(i, 1) = ListaMeses(ComboIni - 1) & " " & Format((UtilidadNetas), "##,##0.00")
         ArepGraficosUtilidades.MSChartIngresos.DataGrid.RowLabel(i, 1) = ListaMeses(ComboIni - 1) & " " & Format(Totalingresos, "##,##0.00")
         ArepGraficosUtilidades.MSChartIngresos.DataGrid.SetData i, 1, Totalingresos, 0
         
         ArepGraficosUtilidades.MSChartGastos.DataGrid.RowLabel(i, 1) = ListaMeses(ComboIni - 1) & " " & Format(TotalGastos + TotalCosto, "##,##0.00")
         ArepGraficosUtilidades.MSChartGastos.DataGrid.SetData i, 1, TotalGastos + TotalCosto, 0
         
         
         Me.osProgress1.Value = Me.osProgress1.Value + 1
         DoEvents
         ComboIni = ComboIni + 1
         Next
    End With
  
                ArepGraficosUtilidades.LblMoneda.Caption = "Expresado en " & Me.CmbMoneda.Text
                ArepGraficosUtilidades.Logo.Picture = LoadPicture(RutaLogo)
                ArepGraficosUtilidades.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                ArepGraficosUtilidades.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                ArepGraficosUtilidades.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                ArepGraficosUtilidades.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
                ArepGraficosUtilidades.LblFechaFin = FechaFin
                ArepGraficosUtilidades.LblFechaIni = FechaIni
                    ArepGraficosUtilidades.Show 1
                
'                 fPreview.arv.ReportSource = ArepGraficosUtilidades
'                 fPreview.Show 1

           
Case "RAZONES FINANCIERAS"

                    Set rpt = New ArepRazonesFinancieras
                    fPreview.RunReport rpt
                    fPreview.Show 1

Case "CATALOGO RESUMEN"

                    Set rpt = New ArepCatalogoResumen
                    fPreview.RunReport rpt
                    fPreview.Show 1
  
 
Case "BALANCE GENERAL TRADICIONAL"
            NumeroPeriodo1 = Me.CmbIni.Text
            NumeroPeriodo2 = Me.CmbFin.Text
            
            If Me.Option8 = True Then
             NumeroTabla = 1
            ElseIf Me.Option7 = True Then
              NumeroTabla = 2
            ElseIf Me.Option6 = True Then
              NumeroTabla = 3
            End If
            
              Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
              Me.DtaConsulta.Refresh
              If Me.DtaConsulta.Recordset.RecordCount = 0 Then
                MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                Exit Sub
              End If
               Me.DtaConsulta.Recordset.MoveLast
               i = Me.DtaConsulta.Recordset.RecordCount
               Me.DtaConsulta.Recordset.MoveFirst
              Do While Not DtaConsulta.Recordset.EOF
        
        
                If i = 1 Then
                  FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                  FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                Else
        
                 If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
                  FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                 ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
                  FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                 End If
                End If
                Me.DtaConsulta.Recordset.MoveNext
              Loop
              
              Fecha1 = Format(FechaIni, "yyyy-mm-dd")
              Fecha2 = Format(FechaFin, "yyyy-mm-dd")
              
              
              ListaActivos = Array("Caja", "Bancos", "Inventario", "Cuentas x Cobrar", "Papeleria - Utiles", "Otros Activos")
              
               
              '//////////////////////////////////////////////////////////////////////////////
              '/////////////TOTAL ACTIVO CIRCULANTES//////////////////////////////////////////////////
              TotalActivoCirculante = 0
              i = 0
              For i = 0 To 5
              Me.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
                                            "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                            "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta " & _
                                            "HAVING (Cuentas.TipoCuenta = '" & ListaActivos(i) & "') AND (SUM(Transacciones.Debito * Transacciones.TCambio) + SUM(Transacciones.TCambio * Transacciones.Credito) <> 0)"
'                                            "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta " & _

              Me.DtaConsulta.Refresh
              If Not Me.DtaConsulta.Recordset.EOF Then
               TotalActivoCirculante = Me.DtaConsulta.Recordset("Total") + TotalActivoCirculante
'               ArepBalanceTradicional.LblActivoCirculante.Caption = Format(TotalActivoCirculante, "##,##0.00")
               RTotalActivoCirculante = TotalActivoCirculante
              End If
              
              Next
              
              '//////////////////////////////////////////////////////////////////////////////
              '/////////////TOTAL ACTIVO fijo//////////////////////////////////////////////////
              Me.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio)-SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
                                            "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
                                            "WHERE    (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                                            "GROUP BY Cuentas.TipoCuenta HAVING (Cuentas.TipoCuenta = 'Activo Fijo') AND (SUM(Transacciones.Debito * Transacciones.TCambio) + SUM(Transacciones.TCambio * Transacciones.Credito) <> 0)"
                                            
              Me.DtaConsulta.Refresh
              If Not Me.DtaConsulta.Recordset.EOF Then
               TotalActivoFijo = Me.DtaConsulta.Recordset("Total")
'               ArepBalanceTradicional.LblTotalActivoFijo.Caption = Format(TotalActivoFijo, "##,##0.00")
               RTotalActivoFijo = TotalActivoFijo
              End If
              
'              ArepBalanceTradicional.LblTotalActivos.Caption = Format(TotalActivoCirculante + TotalActivoFijo, "##,##0.00")
              
              '//////////////////////////////////////////////////////////////////////////////
              '/////////////TOTAL PASIVO//////////////////////////////////////////////////
              
              i = 0
              TotalPasivo = 0
              ListaActivos = Array("Cuentas x Pagar", "Otros Pasivos", "Pasivo")
              For i = 0 To 2
              Me.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Credito * Transacciones.TCambio) - SUM(Transacciones.Debito * Transacciones.TCambio) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
                                            "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
                                            "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                                            "GROUP BY Cuentas.TipoCuenta HAVING (Cuentas.TipoCuenta = '" & ListaActivos(i) & "') AND (SUM(Transacciones.Debito * Transacciones.TCambio) + SUM(Transacciones.TCambio * Transacciones.Credito) <> 0)"
                                            
              Me.DtaConsulta.Refresh
              If Not Me.DtaConsulta.Recordset.EOF Then
               TotalPasivo = Me.DtaConsulta.Recordset("Total") + TotalPasivo
              End If
              Next
              
              RTotalPasivo = TotalPasivo
'              ArepBalanceTradicional.LblTotalPasivo.Caption = Format(TotalPasivo, "##,##0.00")
              
              '//////////////////////////////////////////////////////////////////////////////
              '/////////////TOTAL CAPITAL//////////////////////////////////////////////////
              Me.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Credito * Transacciones.TCambio) - SUM(Transacciones.Debito * Transacciones.TCambio) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
                                            "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
                                            "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                                            "GROUP BY Cuentas.TipoCuenta HAVING (Cuentas.TipoCuenta = 'Capital') AND (SUM(Transacciones.Debito * Transacciones.TCambio) + SUM(Transacciones.TCambio * Transacciones.Credito) <> 0)"
                                            
'

              Me.DtaConsulta.Refresh
              If Not Me.DtaConsulta.Recordset.EOF Then
               TotalCapital = Me.DtaConsulta.Recordset("Total")
              End If
              
              RTotalCapital = TotalCapital
'              ArepBalanceTradicional.LblTotalCapital.Caption = Format(TotalCapital, "##,##0.00")
              '//////////////////////////////////////////////////////////////////////////////
              '/////////////TOTAL INGRESOS//////////////////////////////////////////////////
              
              
              Me.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Credito * Transacciones.TCambio) - SUM(Transacciones.Debito * Transacciones.TCambio) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
                                            "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
                                            "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                                            "GROUP BY Cuentas.TipoCuenta HAVING (Cuentas.TipoCuenta = 'Ingresos - Ventas') AND (SUM(Transacciones.Credito * Transacciones.TCambio) + SUM(Transacciones.Debito * Transacciones.TCambio) <> 0) ORDER BY MAX(Cuentas.CodCuentas)"
                                            
              Me.DtaConsulta.Refresh
              If Not Me.DtaConsulta.Recordset.EOF Then
               Totalingresos = Me.DtaConsulta.Recordset("Total")
              End If
              
              
              '//////////////////////////////////////////////////////////////////////////////
              '/////////////TOTAL COStOS//////////////////////////////////////////////////
              i = 0
              ListaActivos = Array("Costos", "Gastos")
              For i = 0 To 1
                  Me.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Credito * Transacciones.TCambio) - SUM(Transacciones.Debito * Transacciones.TCambio) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
                                                "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
                                                "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                                                "GROUP BY Cuentas.TipoCuenta HAVING (Cuentas.TipoCuenta = '" & ListaActivos(i) & "') AND (SUM(Transacciones.Credito * Transacciones.TCambio) + SUM(Transacciones.Debito * Transacciones.TCambio) <> 0)"
    
                  
                                                  
                  Me.DtaConsulta.Refresh
                  If Not Me.DtaConsulta.Recordset.EOF Then
                   TotalCostos = Abs(Me.DtaConsulta.Recordset("Total")) + TotalCostos
                  End If
              
              Next
              Utilidad = Totalingresos - TotalCostos
              RUtilidad = Utilidad
              
'             ArepBalanceTradicional.LblTotalPasivomasCapital.Caption = Format(TotalCapital + TotalPasivo + Utilidad, "##,##0.00")
'             ArepBalanceTradicional.LblResultadoPeriodo.Caption = Format(Utilidad, "##,##0.00")
             
             
                ArepBalanceTradicional.LblMoneda.Caption = "Expresado en " & Me.CmbMoneda.Text
                ArepBalanceTradicional.Logo.Picture = LoadPicture(RutaLogo)
                ArepBalanceTradicional.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                ArepBalanceTradicional.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                ArepBalanceTradicional.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                ArepBalanceTradicional.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
                ArepBalanceTradicional.LblFechaFin = FechaFin
                ArepBalanceTradicional.LblFechaIni = FechaIni

'            ArepBalanceTradicional.LblFecha2.Caption = Format(FechaFin, "yyyy-mm-dd")
'            ArepBalanceTradicional.Show 1
                 
                        Set rpt = New ArepBalanceTradicional
'                        rpt.DataControl1.ConnectionString = ConexionReporte
'                        rpt.DataControl1.Source = SQL
                        fPreview.RunReport rpt
                        fPreview.Show 1
                




Case "ESTADO DE RESULTADO TRADICIONAL"



    SaldoReportes ("UtilidadResultado")

     Set rpt = New ArepResultadoTradicional
     fPreview.RunReport rpt
     fPreview.Show 1

Case "ESTADO DE RESULTADO RESUMEN ANEXOS"

    NumeroPeriodo1 = FrmReportes.CmbIni.Text
    NumeroPeriodo2 = FrmReportes.CmbFin.Text
    
    If FrmReportes.Option8 = True Then
     NumeroTabla = 1
    ElseIf FrmReportes.Option7 = True Then
      NumeroTabla = 2
    ElseIf FrmReportes.Option6 = True Then
      NumeroTabla = 3
    End If
    
      FrmReportes.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      FrmReportes.DtaConsulta.Refresh
       FrmReportes.DtaConsulta.Recordset.MoveLast
       i = FrmReportes.DtaConsulta.Recordset.RecordCount
       FrmReportes.DtaConsulta.Recordset.MoveFirst
      Do While Not FrmReportes.DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(FrmReportes.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = FrmReportes.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(FrmReportes.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = FrmReportes.DtaConsulta.Recordset("Periodo") Then
          FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        FrmReportes.DtaConsulta.Recordset.MoveNext
      Loop
    
    FrmReportes.DtaReportes.Refresh
    
    

    Ejecutar.Execute "delete from reportes"
    FrmReportes.DtaReportes.Refresh
    CreaEstructura ("Resultado")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not FrmReportes.DtaConsulta.Recordset.EOF Then
      FrmReportes.DtaConsulta.Recordset.MoveLast
      Orden = FrmReportes.DtaConsulta.Recordset("Orden")
    End If
    
    Orden = Orden + 1
    
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Costos y Gastos"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "CG"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    Orden = Orden + 1
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Resultado Periodo"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "RP"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    SaldoReportesAcumulado ("Resultado")
'    SaldoReportesAcumulado ("UtilidadResultado")
    SaldoReportes ("UtilidadResultado")
    
    If FrmReportes.CmbNivel.Text = 0 Then
     EliminaRegistroCero ("Resultado")
    Else
     EliminaRegistroCero ("Nivel")
    End If
    
    Utilidadbruta
    
    ActualizaConfiguracionReporteResultado
    
    Me.TxtTipoReporte.Text = "ESTADO DE RESULTADO RESUMEN ANEXOS"
    
     Set rpt = New ArepResultadoResumen2
     fPreview.RunReport rpt
     fPreview.Show 1
     
     '//////////////////AHORA IMPRIMO LOS ANEXOS//////////////////////
     Set rpt = New ArepAnexosResultado
     fPreview.RunReport rpt
     fPreview.Show 1
     

     
     

Case "ESTADO DE RESULTADO RESUMEN"
    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
       Me.DtaConsulta.Recordset.MoveLast
       i = Me.DtaConsulta.Recordset.RecordCount
       Me.DtaConsulta.Recordset.MoveFirst
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
    
    Me.DtaReportes.Refresh
    
    

    Ejecutar.Execute "delete from reportes"
    Me.DtaReportes.Refresh
    CreaEstructura ("Resultado")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.DtaConsulta.Recordset.MoveLast
      Orden = Me.DtaConsulta.Recordset("Orden")
    End If
    
    Orden = Orden + 1
    
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Costos y Gastos"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "CG"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    Orden = Orden + 1
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Resultado Periodo"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "RP"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    SaldoReportesAcumulado ("Resultado")
'    SaldoReportesAcumulado ("UtilidadResultado")
    SaldoReportes ("UtilidadResultado")
    
    If Me.CmbNivel.Text = 0 Then
     EliminaRegistroCero ("Resultado")
      Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
    Else
     EliminaRegistroCero ("Nivel")
     Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
     Ejecutar.Execute "DELETE FROM Reportes WHERE (Descripcion NOT LIKE N'%Total%') AND (Nivel = " & Me.CmbNivel.Text & ")"
    End If
    
    If Me.CmbNivel.Text = 0 Then
     EliminaRegistroCero ("Resultado")
      Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
    Else
     EliminaRegistroCero ("Nivel")
     Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
     Ejecutar.Execute "DELETE FROM Reportes WHERE (Descripcion NOT LIKE N'%Total%') AND (Nivel = " & Me.CmbNivel.Text & ")"
    End If
    
    
    Utilidadbruta

    Me.TxtTipoReporte.Text = "ESTADO DE RESULTADO RESUMEN"

     Set rpt = New ArepResultadoResumen
     fPreview.RunReport rpt
     fPreview.Show 1
     
Case "ESTADO DE RESULTADO RESUMEN 2"
    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
       Me.DtaConsulta.Recordset.MoveLast
       i = Me.DtaConsulta.Recordset.RecordCount
       Me.DtaConsulta.Recordset.MoveFirst
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
    
    Me.DtaReportes.Refresh
    
    

    Ejecutar.Execute "delete from reportes"
    Me.DtaReportes.Refresh
    CreaEstructura ("Resultado")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.DtaConsulta.Recordset.MoveLast
      Orden = Me.DtaConsulta.Recordset("Orden")
    End If
    
    Orden = Orden + 1
    
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Costos y Gastos"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "CG"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    Orden = Orden + 1
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Resultado Periodo"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "RP"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    If FrmReportes.OptAcumulado.Value = True Then
      SaldoReportesAcumulado ("Resultado")
      SaldoReportesAcumulado ("UtilidadResultado")
    Else
'      SaldoReportes ("Resultado")
'      SaldoReportes ("UtilidadResultado")
      SaldoReportesAcumulado ("Resultado")
      SaldoReportesAcumulado ("UtilidadResultado")
    End If
    
'    If Me.CmbNivel.Text = 0 Then
'     EliminaRegistroCero ("Resultado")
'    Else
'     EliminaRegistroCero ("Nivel")
'    End If
    
    Utilidadbruta
    
    Me.DtaReportes.Refresh

    Me.TxtTipoReporte.Text = "ESTADO DE RESULTADO RESUMEN 2"

     Set rpt = New ArepResultadoResumen3
     fPreview.RunReport rpt
     fPreview.Show 1
     

Case "BALANCE GENERAL RESUMEN"
    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    
    
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
       If Not Me.DtaConsulta.Recordset.EOF Then
         Me.DtaConsulta.Recordset.MoveLast
         i = Me.DtaConsulta.Recordset.RecordCount
         Me.DtaConsulta.Recordset.MoveFirst
       End If
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
    
    
    Ejecutar.Execute "delete from reportes"
'    Me.DtaReportes.Refresh
'    Do While Not Me.DtaReportes.Recordset.EOF
'     Me.DtaReportes.Recordset.Delete
'     Me.DtaReportes.Recordset.MoveNext
'    Loop
    CreaEstructura ("Balance")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.DtaConsulta.Recordset.MoveLast
      Orden = Me.DtaConsulta.Recordset("Orden") + 1
    End If
    
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Pasivo Màs Capital"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "PC"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    
 

'   SaldoReportes ("Utilidad")
'   SaldoReportes ("Balance")

    SaldoReportesAcumulado ("Utilidad")
    SaldoReportesAcumulado ("Balance")
   
'    If Me.CmbNivel.Text = 0 Then
'     EliminaRegistroCero ("Balance")
'    Else
'     EliminaRegistroCero ("Nivel")
'    End If

     AjusteDiferencial ("Balance")
    
     Me.DtaReportes.Refresh
       
     Me.TxtTipoReporte.Text = "BALANCE GENERAL RESUMEN"
     Me.DtaReportes.Refresh
     Set rpt = New ArepBalancePersonalizado
     fPreview.RunReport rpt
     fPreview.Show 1
     
           
Case "BALANCE GENERAL RESUMEN ANEXOS"


    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    
    
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
       If Not Me.DtaConsulta.Recordset.EOF Then
         Me.DtaConsulta.Recordset.MoveLast
         i = Me.DtaConsulta.Recordset.RecordCount
         Me.DtaConsulta.Recordset.MoveFirst
       End If
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
    
    
    Ejecutar.Execute "delete from reportes"

    CreaEstructura ("Balance")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.DtaConsulta.Recordset.MoveLast
      Orden = Me.DtaConsulta.Recordset("Orden") + 1
    End If
    
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Pasivo Màs Capital"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "PC"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update

'   SaldoReportes ("Utilidad")
'   SaldoReportes ("Balance")

    SaldoReportesAcumulado ("Utilidad")
    SaldoReportesAcumulado ("Balance")
   
'    If Me.CmbNivel.Text = 0 Then
'     EliminaRegistroCero ("Balance")
'    Else
'     EliminaRegistroCero ("Nivel")
'    End If

    AjusteDiferencial ("Balance")
    
    ConvertirReporte (FechaFin)
 
    Me.DtaReportes.Refresh
 
'///////////FUNCION REPORTE
     ActualizaConfiguracionReporte
'     R = ReporteResumenAnexos(FechaIni, FechaFin)

     Me.TxtTipoReporte.Text = "BALANCE GENERAL RESUMEN ANEXOS"
     
     Me.DtaReportes.Refresh
     Set rpt = New ArepBalancePersonalizado
     fPreview.RunReport rpt
     fPreview.Show 1
    
'     ArepBalancePersonalizado.Show 1
     
     '////////////AHORA IMPRIMO LOS ANEXOS/////////////////////////////
     'R = AnexosReporteResumen(FechaIni, FechaFin)
     Set rpt = New ArepAnexosBalances
     fPreview.RunReport rpt
     fPreview.Show 1
     


Case "LIBRO MAYOR"


    If Me.ChkSinNiveles.Value = 1 Then
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '//////////////////////////////////////HAGO LA CONSULTA PARA LOS SALDOS INICIALES ///////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            Me.DtaReportes.Refresh
            Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
            Me.lblProgreso.AutoSize = True
            Me.osProgress1.Visible = True
            Me.osProgress1.Value = 0
            Me.osProgress1.Min = 0
            Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
'            Do While Not Me.DtaReportes.Recordset.EOF
'             Me.DtaReportes.Recordset.Delete
'                Me.DtaReportes.Recordset.MoveNext
'                Me.osProgress1.Value = Me.osProgress1.Value + 1
'            Loop
            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open
            Parche.Execute "DELETE FROM Reportes"
            Parche.Close
            
            Me.lblProgreso.Caption = ""
            Me.osProgress1.Visible = False
            SaldoReportes ("BalanzaCodigo")
            Me.DtaReportes.Refresh
            
            'aqui voy a poner el parche para que salga los movimientos de debito y credito en periodo
            
            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open
        
            Parche.Execute "update reportes set debe2=0,haber2=0"
            Parche.Execute "update reportes " & _
                "set debe2=mdebito,haber2=mcredito " & _
                "from (" & ConsultaTotalesMovimientos & ") as Temporal " & _
                "where temporal.codcuentas=reportes.codcuentas"
                
            Parche.Execute "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2"
'            SQL = "SELECT  MAX(Reportes.CodCuentas) AS CodCuentas, MAX(Cuentas.DescripcionCuentas) AS NombreCuenta, Cuentas.KeyGrupo, SUM(Reportes.Debe2) AS Debito, SUM(Reportes.Haber2) AS Credito, Cuentas.DescripcionGrupo FROM  Reportes INNER JOIN Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas GROUP BY Cuentas.KeyGrupo, Cuentas.DescripcionGrupo ORDER BY MAX(Reportes.CodCuentas) "
            Sql = "SELECT  MAX(Reportes.CodCuentas) AS CodCuentas, MAX(Cuentas.DescripcionCuentas) AS NombreCuenta, Cuentas.KeyGrupo, SUM(Reportes.Debe2) AS Debe2, SUM(Reportes.Haber2) AS Haber2, Cuentas.DescripcionGrupo AS Descripcion, SUM(Reportes.Debe1) AS Debe1, SUM(Reportes.Haber1) AS Haber1, SUM(Reportes.Debe3) AS Debe3, SUM(Reportes.Haber3) AS Haber3 FROM Reportes INNER JOIN Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas GROUP BY Cuentas.KeyGrupo, Cuentas.DescripcionGrupo ORDER BY CodCuentas"
       
       
            ArepLibroMayorSNiveles.DataControl1.Source = Sql
            ArepLibroMayorSNiveles.Logo.Picture = LoadPicture(RutaLogo)
            ArepLibroMayorSNiveles.LblEmpresa.Caption = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
            ArepLibroMayorSNiveles.LblEmpresa1.Caption = Me.DtaDatosEmpresa.Recordset("Direccion")
            ArepLibroMayorSNiveles.LblEmpresa2.Caption = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
            ArepLibroMayorSNiveles.LblFechaIni.Caption = FechaIni
            ArepLibroMayorSNiveles.LblFechaFin.Caption = FechaFin
            ArepLibroMayorSNiveles.LblFechaImpreso.Caption = "Impreso: " & Format(Now, "dd/mm/yyyy")
            ArepLibroMayorSNiveles.DataControl1.ConnectionString = ConexionReporte
            
            ArepLibroMayorSNiveles.Show 1
       
       
       Else

                '----------------------------

            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open
            Parche.Execute "DELETE FROM Reportes"
            Parche.Close
        
           
            If Me.ChkMostrarMovxMes.Value = 1 Then
             CreaEstructura ("BalanzaLibroMayor")
             SaldoReportes ("BalanzaLibroMayor")
            Else
             CreaEstructura ("Balanza")
             SaldoReportes ("Balanza")
            End If
            
            EliminaRegistroCero ("Libro")
            
            '-----------------------BORRO TODAS LAS CUENTAS QUE NO SUMAN NINGUN VALOR ------------------
            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open
            'Parche.Execute "DELETE FROM Reportes Where (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)"
            Parche.Execute "DELETE FROM Reportes Where (Debe2 + Debe3 - Haber2 - Haber3 = 0)"
'            Parche.Execute "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)"
        
        
'            Sql = "SELECT  * From Reportes WHERE (Nivel = " & Me.CmbNivel2.Text & ") AND (Descripcion LIKE N'%Total%') AND (Debe2 + Debe3 - Haber2 - Haber3 <> 0) ORDER BY Orden"
'
'            ArepLibroMayor.DataControl1.Source = Sql
            ArepLibroMayor.Logo.Picture = LoadPicture(RutaLogo)
            ArepLibroMayor.LblEmpresa.Caption = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
            ArepLibroMayor.LblEmpresa1.Caption = Me.DtaDatosEmpresa.Recordset("Direccion")
            ArepLibroMayor.LblEmpresa2.Caption = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
            ArepLibroMayor.LblFechaIni.Caption = FechaIni
            ArepLibroMayor.LblFechaFin.Caption = FechaFin
            ArepLibroMayor.LblFechaImpreso.Caption = "Impreso: " & Format(Now, "dd/mm/yyyy")
        '    ArepLibroMayor.DataControl1.ConnectionString = ConexionReporte
            
            ArepLibroMayor.Show 1


       End If
       



Case "LIBRO DIARIO"
    NumFecha1 = FechaIni
    NumFecha2 = FechaFin
    
    If Me.ChkSinNiveles.Value = 1 Then
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '//////////////////////////////////////HAGO LA CONSULTA PARA LOS SALDOS INICIALES ///////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            Me.DtaReportes.Refresh
            Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
            Me.lblProgreso.AutoSize = True
            Me.osProgress1.Visible = True
            Me.osProgress1.Value = 0
            Me.osProgress1.Min = 0
            Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
'            Do While Not Me.DtaReportes.Recordset.EOF
'             Me.DtaReportes.Recordset.Delete
'                Me.DtaReportes.Recordset.MoveNext
'                Me.osProgress1.Value = Me.osProgress1.Value + 1
'            Loop
'        SQL = "SELECT Max(Transacciones.CodCuentas) AS CodCuentas, Max(Transacciones.NumeroMovimiento) AS NumeroMovimiento, Max(Transacciones.NombreCuenta) AS NombreCuenta, Avg(Transacciones.TCambio) AS TCambio, Sum(Transacciones.TCambio*Transacciones.Debito) AS Debito, Sum(Transacciones.TCambio*Transacciones.Credito) AS Credito, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta " & _
'        "FROM Grupos INNER JOIN (Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas) ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
'        "WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) " & _
'        "GROUP BY Cuentas.KeyGrupo, Cuentas.DescripcionGrupo " & _
'        "ORDER BY Cuentas.KeyGrupo "
            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open
            Parche.Execute "DELETE FROM Reportes"
            Parche.Close
            
            Me.lblProgreso.Caption = ""
            Me.osProgress1.Visible = False
            SaldoReportes ("BalanzaCodigo")
            Me.DtaReportes.Refresh
            
            'aqui voy a poner el parche para que salga los movimientos de debito y credito en periodo
'
            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open

            Parche.Execute "update reportes set debe2=0,haber2=0"
            Parche.Execute "update reportes " & _
                "set debe2=mdebito,haber2=mcredito " & _
                "from (" & ConsultaTotalesMovimientos & ") as Temporal " & _
                "where temporal.codcuentas=reportes.codcuentas"

            Parche.Execute "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2"
''            SQL = "SELECT  MAX(Reportes.CodCuentas) AS CodCuentas, MAX(Cuentas.DescripcionCuentas) AS NombreCuenta, Cuentas.KeyGrupo, SUM(Reportes.Debe2) AS Debito, SUM(Reportes.Haber2) AS Credito, Cuentas.DescripcionGrupo FROM  Reportes INNER JOIN Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas GROUP BY Cuentas.KeyGrupo, Cuentas.DescripcionGrupo ORDER BY MAX(Reportes.CodCuentas) "
            Sql = "SELECT  MAX(Reportes.CodCuentas) AS CodCuentas, MAX(Cuentas.DescripcionCuentas) AS NombreCuenta, Cuentas.KeyGrupo, SUM(Reportes.Debe2) AS Debe2, SUM(Reportes.Haber2) AS Haber2, Cuentas.DescripcionGrupo AS Descripcion, SUM(Reportes.Debe1) AS Debe1, SUM(Reportes.Haber1) AS Haber1, SUM(Reportes.Debe3) AS Debe3, SUM(Reportes.Haber3) AS Haber3 FROM Reportes INNER JOIN Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas GROUP BY Cuentas.KeyGrupo, Cuentas.DescripcionGrupo ORDER BY CodCuentas"
    
    
        ArepLibroDiarioSNiveles.DataControl1.Source = Sql
        ArepLibroDiarioSNiveles.Logo.Picture = LoadPicture(RutaLogo)
        ArepLibroDiarioSNiveles.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
        ArepLibroDiarioSNiveles.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
        ArepLibroDiarioSNiveles.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
        ArepLibroDiarioSNiveles.LblFechaImpreso.Caption = "Impreso: " & Format(Now, "dd/mm/yyyy")
        ArepLibroDiarioSNiveles.LblTipoMoneda.Caption = "Tipo Moneda: " & Me.CmbMoneda.Text
        ArepLibroDiarioSNiveles.DataControl1.ConnectionString = ConexionReporte
        ArepLibroDiarioSNiveles.LblFechaIni.Caption = FechaIni
        ArepLibroDiarioSNiveles.LblFechaFin.Caption = FechaFin
       
        
        ArepLibroDiarioSNiveles.Show 1
    '           fPreview.arv.ReportSource = ArepLibroDiario
    '           fPreview.Show 1
    
    Else
    
    
            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open
            Parche.Execute "DELETE FROM Reportes"
            Parche.Close
        
                
            CreaEstructura ("Balanza")
            SaldoReportes ("Balanza")
            EliminaRegistroCero ("Libro")
            
            AjusteDiferencial ("Libro")
    
            '-----------------------BORRO TODAS LAS CUENTAS QUE NO SUMAN NINGUN VALOR ------------------
            Set Parche = New ADODB.Connection
            Parche.ConnectionString = Conexion
            Parche.Open
            Parche.Execute "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe2 + Debe3 - Haber2  - Haber3 = 0)"
            'Parche.Execute "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)"
        
        
'            SQL = "SELECT  * From Reportes WHERE (Nivel = " & Me.CmbNivel2.Text & ") AND (Descripcion LIKE N'%Total%') AND (Debe2 + Haber2 <> 0) ORDER BY Orden"
             Sql = "SELECT  * From Reportes WHERE (Nivel = '-100') AND (Descripcion LIKE N'%Total%') AND (Debe2 + Haber2 <> 0) ORDER BY Orden"
    
'    ArepLibroDiario.DataControl1.Source = SQL
    ArepLibroDiario.Logo.Picture = LoadPicture(RutaLogo)
    ArepLibroDiario.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepLibroDiario.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepLibroDiario.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepLibroDiario.LblFechaImpreso.Caption = "Impreso: " & Format(Now, "dd/mm/yyyy")
    ArepLibroDiario.LblTipoMoneda.Caption = "Tipo Moneda: " & Me.CmbMoneda.Text
'    ArepLibroDiario.DataControl1.ConnectionString = ConexionReporte
    ArepLibroDiario.LblFechaIni.Caption = FechaIni
    ArepLibroDiario.LblFechaFin.Caption = FechaFin
   
    
    ArepLibroDiario.Show 1
'           fPreview.arv.ReportSource = ArepLibroDiario
'           fPreview.Show 1
    
    End If
    


 Case "DETALLE DIARIO MAYOR"
    NumFecha1 = FechaIni
    NumFecha2 = FechaFin
    Sql = "SELECT MAX(Transacciones.CodCuentas) AS CodCuentas, Transacciones.FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, Avg(Transacciones.TCambio) AS TCambio, Sum(Transacciones.TCambio*Transacciones.Debito) AS Debito, Sum(Transacciones.TCambio*Transacciones.Credito) AS Credito, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo " & _
    "FROM (Grupos INNER JOIN Cuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo) INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
    "GROUP BY Transacciones.FechaTransaccion, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo " & _
    "HAVING (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) " & _
    "ORDER BY Cuentas.KeyGrupo, Transacciones.FechaTransaccion "

    ArepLibroDiarioMayor.DataControl1.Source = Sql
    ArepLibroDiarioMayor.Logo.Picture = LoadPicture(RutaLogo)
    ArepLibroDiarioMayor.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepLibroDiarioMayor.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepLibroDiarioMayor.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepLibroDiarioMayor.LblFechaImpreso.Caption = "Impreso: " & Format(Now, "dd/mm/yyyy")
    ArepLibroDiarioMayor.LblFechaIni.Caption = FechaIni
    ArepLibroDiarioMayor.LblFechaFin.Caption = FechaFin
    ArepLibroDiarioMayor.DataControl1.ConnectionString = ConexionReporte
    
'    ArepLibroDiarioMayor.Show 1
           fPreview.arv.ReportSource = ArepLibroDiarioMayor
           fPreview.Show 1
    
End Select

Me.Frame6.Visible = False
Me.CmdVerReporte2.Enabled = True
Me.CmdSalir.Enabled = True
Me.CmdVerReporte2.Visible = False
Me.CmdVerReporte.Visible = True
Exit Sub
TipoErrs:
  MsgBox err.Description

End Sub

Private Sub CmdVerReporte3_Click()
Dim Utilidad As Double, Utilidad2 As Double, Utilidad3 As Double, RegTCostosOper As Integer
Dim Decrementador As Integer, TotalActivoCirculante As Double, TotalActivoFijo As Double, TotalActivoDiferido As Double
Dim TotalPasivoCirculante As Double, TotalPasivoFijo As Double, TotalPasivoDiferido As Double, TotalCapitalSocial As Double
Dim RegInicioCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro donde comienzan los costos operativos
Dim RegTotalCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro de total de costos operativos
Dim Totalingresos As Double, TotalCostoVentas As Double, TotalGastosAdmon As Double, TotalGastos As Double
Dim TotalGastoVentas As Double, TotalIngresosFinancieros As Double, TotalOtrosIngresos As Double, TotalOtrosGastos As Double
Dim TotalUtilidadBruta As Double, TotalImpuestos As Double, TotalUtilidadNeta As Double, Fecha1 As String, Fecha2 As String
Dim TotalCompras As Double, TotalInventarioInicial As Double, TotalInventarioFinal As Double
Dim TotalAcarreo As Double, TotalRebajaVentas As Double, TotalDisponible As Double, TotalGastosR As Double, TotalCosto As Double
Dim TotalSalidas As Double, TotalGastoOperacion As Double, TotalPasivo As Double, TotalCapital As Double
Dim TotalCostos As Double, ListaActivos As Variant, TotalInventario As Double, TotalCuentaxCobrar As Double
Dim TotalCuentasxPagar As Double, TotalActivos As Double, UtilidadBrutas As Double, UtilidadNetas As Double
Dim UltimoOrden As Integer, RegIngresos  As Integer, PrimReg As Integer, UltReg As Integer
Dim Fechas1 As String, Fechas2 As String, Orden As Integer, Sql As String, i As Double
Dim ListaMeses As Variant, CantRegistros As Double, ComboIni As Double, ComboFin As Double, TotstoFijo As Double, TotalGastoFijo As Double
Dim mes As Double, J As Double, TotalDebito As Double, TotalCredito As Double
Dim rpt As Object, TotalCostoFijo As Double, Ajuste As String
Dim fPreview As New FrmPreview
Dim Rs As New ADODB.Recordset
Dim Parche As ADODB.Connection


'On Error GoTo TipoErrs
Me.Frame6.Visible = True
Me.CmdVerReporte2.Enabled = False
Me.CmdSalir.Enabled = False
SaldoIni = 0
SaldoFin = 0
Total1 = 0
TotalCuenta = 0

FechaIni = Me.DTFecha1.Value
FechaFin = Me.DTFecha2.Value




Select Case Me.CmbReportes.Text

Case "RETENCIONES X BENEFICIARIO DETALLE"

   ReporteRetencionesBeneficiario False
 "RETENCIONES X BENEFICIARIO RESUMIDO"

   ReporteRetencionesBeneficiario True

Case "RETENCIONES X PAGAR"


   ReporteRetenciones


Case "AUXILIAR CUENTAS X BENEFICIARIO"


   GenerarAuxiliarDeCuentas True, True

Case "ESTADO DE RESULTADO DPTO"
    NumeroPeriodo1 = Me.CmbIni.Text
    NumeroPeriodo2 = Me.CmbFin.Text
    
    
    
             Me.CmdVerReporte.Visible = False
         Me.CmdVerReporte2.Visible = False
         Me.CmdVerReporte3.Visible = True
         Me.CmbMoneda.Visible = True
         Me.Label3.Visible = True
         Me.Frame9.Visible = True
         Me.Frame4.Visible = True
         Me.Frame1.Visible = False
    Me.ChkQuitarMovimiento.Visible = True
    
    If Me.Option8 = True Then
     NumeroTabla = 1
    ElseIf Me.Option7 = True Then
      NumeroTabla = 2
    ElseIf Me.Option6 = True Then
      NumeroTabla = 3
    End If
    
      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      Me.DtaConsulta.Refresh
       Me.DtaConsulta.Recordset.MoveLast
       i = Me.DtaConsulta.Recordset.RecordCount
       Me.DtaConsulta.Recordset.MoveFirst
      Do While Not DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        Me.DtaConsulta.Recordset.MoveNext
      Loop
    
    
    
    
    Me.DtaReportes.Refresh
    
    FrmReportes.lblProgreso.Caption = "Eliminando Datos del Reporte Anterior"
    FrmReportes.osProgress1.Visible = True
    FrmReportes.osProgress1.Value = 0
    FrmReportes.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
    Me.DtaReportes.Refresh
'    Do While Not Me.DtaReportes.Recordset.EOF
'     FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
'
'     Me.DtaReportes.Recordset.Delete
'     Me.DtaReportes.Refresh
'     Me.DtaReportes.Recordset.MoveNext
'    Loop
    
    Rs.Open "DELETE FROM Reportes", Conexion

    CreaEstructuraDpto ("Resultado")
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
    FrmReportes.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.DtaConsulta.Recordset.MoveLast
      Orden = Me.DtaConsulta.Recordset("Orden")
    End If
    
    Orden = Orden + 1
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Total Costos y Gastos"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "CG"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    Orden = Orden + 1
    FrmReportes.DtaReportes.Recordset.AddNew
       FrmReportes.DtaReportes.Recordset("Descripcion") = " Resultado Periodo"
       FrmReportes.DtaReportes.Recordset("KeyGrupo") = "RP"
       FrmReportes.DtaReportes.Recordset("Nivel") = 1
       FrmReportes.DtaReportes.Recordset("Orden") = Orden
    FrmReportes.DtaReportes.Recordset.Update
    
    SaldoReportesDpto ("Resultado")
    SaldoReportesDpto ("UtilidadResultado")
    
    If Me.CmbNivel.Text = 0 Then
     EliminaRegistroCero ("Resultado")
      Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
    Else
     EliminaRegistroCeroDpto ("Nivel")
     Rs.Open "DELETE FROM Reportes Where (Not (CodCuentas Is Null)) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0)", Conexion
     Ejecutar.Execute "DELETE FROM Reportes WHERE (Descripcion NOT LIKE N'%Total%') AND (Nivel = " & Me.CmbNivel.Text & ")"
    End If
    
    Me.DtaReportes.RecordSource = "Reportes"
    Me.DtaReportes.Refresh
    ArepResultado.LblMoneda.Caption = "Expresado en " & Me.CmbMoneda.Text
    If Dir(RutaLogo) <> "" Then
     ArepResultado.Logo.Picture = LoadPicture(RutaLogo)
    End If
    ArepResultado.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepResultado.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
    ArepResultado.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepResultado.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepResultado.LblFechaFin = FechaFin

    ArepResultado.LblFechaIni = FechaIni
     
'    Utilidadbruta
    
    ArepResultado.DataControl1.ConnectionString = ConexionReporte
    ArepResultado.DataControl1.Source = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
'    ArepResultado.Show 1
    
'     fPreview.arv.ReportSource = ArepResultado
'     fPreview.Show 1

     Set rpt = New ArepResultado
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
     fPreview.RunReport rpt
     fPreview.Show 1

    
    
    FrmReportes.lblProgreso.Caption = ""
    FrmReportes.osProgress1.Visible = False

Case "COMPROBANTE DE DIARIO"
'    ArepComprobanteDiario.LblFecha = Format(Now, "dd/mm/yyyy")
'    ArepComprobanteDiario.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
    NumFecha1 = Me.DTFecha1.Value
    NumFecha2 = Me.DTFecha2.Value
    ArepComprobanteDiario.DataControl1.ConnectionString = ConexionReporte
    
    If Me.CmbMoneda.Text = "Córdobas" Then
         Moneda = "Cordobas"
    Else
         Moneda = "Dolares"
    End If
    
   If Me.TxtDptoDesde.Text = "" And Me.TxtDptoHasta.Text = "" Then
        Sql = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END) AS Debito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END) AS DebitoD, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END) AS Credito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END) AS CreditoD, Transacciones.FechaTransaccion, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN " & _
              "Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
              "GROUP BY Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.FechaTransaccion, IndiceTransaccion.TipoMoneda, Transacciones.NumeroMovimiento, Tasas.MontoCordobas , Cuentas.DescripcionGrupo HAVING (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
    Else

        Sql = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END) AS Debito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END) AS DebitoD, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END) AS Credito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END) AS CreditoD, Transacciones.FechaTransaccion, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN " & _
              "Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
              "WHERE  (IndiceTransaccion.Fuente BETWEEN '" & Me.TxtDptoDesde.Text & "' AND '" & Me.TxtDptoHasta.Text & "') GROUP BY Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.FechaTransaccion, IndiceTransaccion.TipoMoneda, Transacciones.NumeroMovimiento, Tasas.MontoCordobas, Cuentas.DescripcionGrupo HAVING (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"

    End If

     Set rpt = New ArepComprobanteDiario
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = Sql
     fPreview.RunReport rpt


     fPreview.Show 1


Case "LISTA CUENTAS X PAGAR"
    ArepTotalAuxiliar.LblFecha = Format(Now, "dd/mm/yyyy")
    NumFecha1 = Me.DTFecha1.Value
    NumFecha2 = Me.DTFecha2.Value
    
    Fechas1 = Format(Me.DTFecha1.Value, "yyyy-mm-dd")
    Fechas2 = Format(Me.DTFecha2.Value, "yyyy-mm-dd")
    
    
    ArepTotalAuxiliar.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
    ArepTotalAuxiliar.LblTotal.Caption = "REPORTE DE CUENTAS X PAGAR"
    ArepTotalAuxiliar.DataControl1.ConnectionString = ConexionReporte
    
    
    If Me.Option4.Value = True Then
    
    
    Me.DtaReportes.Refresh
    Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
    Me.lblProgreso.AutoSize = True
    Me.osProgress1.Visible = True
    Me.osProgress1.Value = 0
    Me.osProgress1.Min = 0
    Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
    
    Do While Not Me.DtaReportes.Recordset.EOF
     Me.DtaReportes.Recordset.Delete
        Me.DtaReportes.Recordset.MoveNext
        Me.osProgress1.Value = Me.osProgress1.Value + 1
    Loop
    
    Me.lblProgreso.Caption = ""
    Me.osProgress1.Visible = False
    Me.AdoConsultas.RecordSource = "SELECT  * From Cuentas WHERE (TipoCuenta = 'Cuentas x Pagar') ORDER BY CodCuentas"
    Me.AdoConsultas.Refresh
    If Not Me.AdoConsultas.Recordset.EOF Then
        Me.AdoConsultas.Refresh
'          Me.TxtDesde.Text = Me.AdoConsultas.Recordset("CodCuentas")
          If Me.DBCodigo.Text = "" Then
           Me.DBCodigo.Text = Me.AdoConsultas.Recordset("CodCuentas")
          End If
          
        Me.AdoConsultas.Recordset.MoveLast

'          Me.TxtHasta.Text = Me.AdoConsultas.Recordset("CodCuentas")
         If Me.DBCodigoHasta.Text = "" Then
          Me.DBCodigoHasta.Text = Me.AdoConsultas.Recordset("CodCuentas")
         End If
    End If
    
    SaldoReportes ("SaldoCuentas")
    Me.DtaReportes.Refresh
    'aqui voy a poner el parche para que salga los movimientos de debito y credito en periodo
'    Dim Parche As ADODB.Connection
    Set Parche = New ADODB.Connection
    Parche.ConnectionString = Conexion
    Parche.Open

    Parche.Execute "update reportes set debe2=0,haber2=0"
    Parche.Execute "update reportes " & _
        "set debe2=mdebito,haber2=mcredito " & _
        "from (" & ConsultaTotalesMovimientos & ") as Temporal " & _
        "where temporal.codcuentas=reportes.codcuentas"
        
    Parche.Execute "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2"
     
                    ArepTotalAuxiliar.Logo.Picture = LoadPicture(RutaLogo)
                    ArepTotalAuxiliar.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                    ArepTotalAuxiliar.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                    ArepTotalAuxiliar.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                    ArepTotalAuxiliar.DataControl1.Source = "SELECT  Reportes.Descripcion, Reportes.Debe1 + Reportes.Haber1 AS SaldoInicial, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3 - Reportes.Haber3 AS SaldoFinal, Reportes.Orden , Cuentas.CodCuentas, Cuentas.DescripcionCuentas FROM  Reportes INNER JOIN  Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas ORDER BY Reportes.Orden"
                    ArepTotalAuxiliar.Show 1
'                    fPreview.arv.ReportSource = ArepTotalAuxiliar
'                    fPreview.Show 1

     
     ElseIf Me.Option5.Value = True Then

            Me.DtaReportes.Refresh
            
            Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
            Me.lblProgreso.AutoSize = True
            Me.osProgress1.Value = 0
            Me.osProgress1.Min = 0
            Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
            Do While Not Me.DtaReportes.Recordset.EOF
             Me.osProgress1.Value = Me.osProgress1.Value + 1
             Me.DtaReportes.Recordset.Delete
             Me.DtaReportes.Recordset.MoveNext
            Loop
            CreaEstructura ("Balanza")
            SaldoReportes ("Balanza")
            EliminaRegistroCero ("Balanza")

                    Me.DtaReportes.Refresh
                    'aqui voy a poner el parche para que salga los movimientos de debito y credito en periodo
                '    Dim Parche As ADODB.Connection
     
                    ArepTotalAuxiliar.LblFecha = Format(Now, "dd/mm/yyyy")
                    ArepTotalAuxiliar.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
                    ArepTotalAuxiliar.DataControl1.ConnectionString = ConexionReporte
                    ArepTotalAuxiliar.Logo.Picture = LoadPicture(RutaLogo)
                    ArepTotalAuxiliar.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                    ArepTotalAuxiliar.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                    ArepTotalAuxiliar.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                    ArepTotalAuxiliar.Field17.Visible = False
                    ArepTotalAuxiliar.Field22.Visible = False
                    ArepTotalAuxiliar.Field23.Visible = True
                    ArepTotalAuxiliar.DataControl1.Source = "SELECT Descripcion, Debe1 + Haber1 AS SaldoInicial, Debe2, Haber2, Debe3 - Haber3 AS SaldoFinal, Orden,CodCuentas,KeyGrupo From Reportes ORDER BY Orden"
                    ArepTotalAuxiliar.Show 1
'                    fPreview.arv.ReportSource = ArepTotalAuxiliarG
'                    fPreview.Show 1

     
     End If

Case "LISTA CUENTAS X COBRAR"
    ArepTotalAuxiliar.LblFecha = Format(Now, "dd/mm/yyyy")
    NumFecha1 = Me.DTFecha1.Value
    NumFecha2 = Me.DTFecha2.Value
    
    Fechas1 = Format(Me.DTFecha1.Value, "yyyy-mm-dd")
    Fechas2 = Format(Me.DTFecha2.Value, "yyyy-mm-dd")
    
    
    ArepTotalAuxiliar.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
    ArepTotalAuxiliar.LblTotal.Caption = "REPORTE DE CUENTAS X COBRAR"
    ArepTotalAuxiliar.DataControl1.ConnectionString = ConexionReporte
    
    
    If Me.Option4.Value = True Then
    
    
    Me.DtaReportes.Refresh
    Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
    Me.lblProgreso.AutoSize = True
    Me.osProgress1.Visible = True
    Me.osProgress1.Value = 0
    Me.osProgress1.Min = 0
    Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
    
    Do While Not Me.DtaReportes.Recordset.EOF
     Me.DtaReportes.Recordset.Delete
        Me.DtaReportes.Recordset.MoveNext
        Me.osProgress1.Value = Me.osProgress1.Value + 1
    Loop
    
    Me.lblProgreso.Caption = ""
    Me.osProgress1.Visible = False
    Me.AdoConsultas.RecordSource = "SELECT  * From Cuentas WHERE (TipoCuenta = 'Cuentas x Cobrar') ORDER BY CodCuentas"
    Me.AdoConsultas.Refresh
    If Not Me.AdoConsultas.Recordset.EOF Then
        Me.AdoConsultas.Refresh
          Me.TxtDesde.Text = Me.AdoConsultas.Recordset("CodCuentas")
         FrmReportes.DBCodigo.Text = Me.AdoConsultas.Recordset("CodCuentas")

        Me.AdoConsultas.Recordset.MoveLast

          Me.TxtHasta.Text = Me.AdoConsultas.Recordset("CodCuentas")
          FrmReportes.DBCodigoHasta.Text = Me.AdoConsultas.Recordset("CodCuentas")

    
    End If
    
    SaldoReportes ("SaldoCuentas")
    Me.DtaReportes.Refresh
    'aqui voy a poner el parche para que salga los movimientos de debito y credito en periodo
'    Dim Parche As ADODB.Connection
    Set Parche = New ADODB.Connection
    Parche.ConnectionString = Conexion
    Parche.Open

    Parche.Execute "update reportes set debe2=0,haber2=0"
    Parche.Execute "update reportes " & _
        "set debe2=mdebito,haber2=mcredito " & _
        "from (" & ConsultaTotalesMovimientos & ") as Temporal " & _
        "where temporal.codcuentas=reportes.codcuentas"
        
    Parche.Execute "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2"
     
                    ArepTotalAuxiliar.Logo.Picture = LoadPicture(RutaLogo)
                    ArepTotalAuxiliar.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                    ArepTotalAuxiliar.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                    ArepTotalAuxiliar.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                    ArepTotalAuxiliar.DataControl1.Source = "SELECT  Reportes.Descripcion, Reportes.Debe1 + Reportes.Haber1 AS SaldoInicial, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3 - Reportes.Haber3 AS SaldoFinal, Reportes.Orden , Cuentas.CodCuentas, Cuentas.DescripcionCuentas FROM  Reportes INNER JOIN  Cuentas ON Reportes.CodCuentas = Cuentas.CodCuentas ORDER BY Reportes.Orden"
                    ArepTotalAuxiliar.Show 1
'                    fPreview.arv.ReportSource = ArepTotalAuxiliar
'                    fPreview.Show 1

     
     ElseIf Me.Option5.Value = True Then

            Me.DtaReportes.Refresh
            
            Me.lblProgreso.Caption = "Limpiando registros del Reporte Anterior..."
            Me.lblProgreso.AutoSize = True
            Me.osProgress1.Value = 0
            Me.osProgress1.Min = 0
            Me.osProgress1.Max = Me.DtaReportes.Recordset.RecordCount
            Do While Not Me.DtaReportes.Recordset.EOF
             Me.osProgress1.Value = Me.osProgress1.Value + 1
             Me.DtaReportes.Recordset.Delete
             Me.DtaReportes.Recordset.MoveNext
            Loop
            CreaEstructura ("Balanza")
            SaldoReportes ("Balanza")
            EliminaRegistroCero ("Balanza")

                    Me.DtaReportes.Refresh
                    'aqui voy a poner el parche para que salga los movimientos de debito y credito en periodo
                '    Dim Parche As ADODB.Connection
     
                    ArepTotalAuxiliar.LblFecha = Format(Now, "dd/mm/yyyy")
                    ArepTotalAuxiliar.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
                    ArepTotalAuxiliar.DataControl1.ConnectionString = ConexionReporte
                    ArepTotalAuxiliar.Logo.Picture = LoadPicture(RutaLogo)
                    ArepTotalAuxiliar.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                    ArepTotalAuxiliar.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                    ArepTotalAuxiliar.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                    ArepTotalAuxiliar.Field17.Visible = False
                    ArepTotalAuxiliar.Field22.Visible = False
                    ArepTotalAuxiliar.Field23.Visible = True
                    ArepTotalAuxiliar.DataControl1.Source = "SELECT Descripcion, Debe1 + Haber1 AS SaldoInicial, Debe2, Haber2, Debe3 - Haber3 AS SaldoFinal, Orden,CodCuentas,KeyGrupo From Reportes ORDER BY Orden"
                    ArepTotalAuxiliar.Show 1
'                    fPreview.arv.ReportSource = ArepTotalAuxiliarG
'                    fPreview.Show 1

     
     End If




Case "PUNTO DE EQUILIBRIO"
     
    
    ListaMeses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    
    
    With ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data
      If Me.CmbFin.Text = Me.CmbIni.Text Then
         CantRegistros = 1
         ComboIni = Val(Me.CmbIni.Text)
         ComboFin = Val(Me.CmbFin.Text)
      Else
         CantRegistros = Val(Me.CmbFin.Text) - Val(Me.CmbIni.Text) + 1
         ComboIni = Val(Me.CmbIni.Text)
         ComboFin = Val(Me.CmbFin.Text)
      End If
        
        Me.osProgress1.Visible = True
        Me.osProgress1.Value = 0
        Me.osProgress1.Max = CantRegistros
        
      J = 0
      For i = 1 To CantRegistros
        Fecha2 = Format(FechaPeriodo(ComboIni), "yyyy-mm-dd")
        Fecha1 = Format(FechaPeriodoIni(ComboIni), "yyyy-mm-dd")
        If i = 1 Then
        
         If Me.OptAcumulado.Value = True Then
            ArepPuntoEquilibrio.Chart2D.Header.Text = "REPORTE PUNTO DE EQUILIBRIO - " & Year(Fecha1)
'
         Else
             ArepPuntoEquilibrio.Chart2D.Header.Text = "REPORTE PUNTO DE EQUILIBRIO - " & Year(Fecha1)
         End If
        
        
        End If
             If Me.OptAcumulado.Value = True Then
                '////////////////////////TOTAL INGRESOS//////////////////////////////////////////////////////////
                Totalingresos = 0
                Totalingresos = SaldosRazonesCreditos(Fecha2, "Ingresos - Ventas")
                
                '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalCosto = 0
                TotalCosto = SaldosRazonesDebitos(Fecha2, "Costos")
                TotalCostoFijo = SaldosRazonesDebitosFijo(Fecha2, "Costos", "Fijo")

                 '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalGastos = 0
                TotalGastos = SaldosRazonesDebitos(Fecha2, "Gastos")
                TotalGastoFijo = SaldosRazonesDebitosFijo(Fecha2, "Gastos", "Fijo")
                
             ElseIf Me.OptPeriodo.Value = True Then
                '////////////////////////TOTAL INGRESOS//////////////////////////////////////////////////////////
                Totalingresos = 0
                Totalingresos = SaldoPeriodoCredito(Fecha1, Fecha2, "Ingresos - Ventas")

             
                '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalCosto = 0
                TotalCosto = SaldoPeriodoDebito(Fecha1, Fecha2, "Costos")
                TotalCostoFijo = SaldoPeriodoDebitoFijo(Fecha1, Fecha2, "Costos", "Fijo")

                 '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalGastos = 0
                TotalGastos = SaldoPeriodoDebito(Fecha1, Fecha2, "Gastos")
                TotalGastoFijo = SaldoPeriodoDebitoFijo(Fecha1, Fecha2, "Gastos", "Fijo")
             
             End If
             
            With ArepPuntoEquilibrio.Chart2D.ChartArea.Axes("X")
                .Min = 0

            End With
            With ArepPuntoEquilibrio.Chart2D.ChartArea.Axes("Y")
                .Min = 0

            End With

         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.NumPoints(1) = CantRegistros + 1
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.NumPoints(2) = CantRegistros + 1
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.NumPoints(3) = CantRegistros + 1
         
         If i = 1 Then
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.x(1, 1) = 0
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.y(1, 1) = TotalCostoFijo + TotalGastoFijo
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.x(2, 1) = 0
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.y(2, 1) = 0
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.x(3, 1) = 0
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.y(3, 1) = TotalCostoFijo + TotalGastoFijo
         
         
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.x(1, 2) = 1
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.y(1, 2) = TotalCostoFijo + TotalGastoFijo
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.x(2, 2) = 1
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.y(2, 2) = Totalingresos
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.x(3, 2) = 1
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.y(3, 2) = TotalCosto + TotalGastos
         
         J = 2
         Else
         
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.x(1, J) = J - 1
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.y(1, J) = TotalCostoFijo + TotalGastoFijo
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.x(2, J) = J - 1
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.y(2, J) = Totalingresos
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.x(3, J) = J - 1
         ArepPuntoEquilibrio.Chart2D.ChartGroups(1).Data.y(3, J) = TotalCosto + TotalGastos
         End If
         ArepPuntoEquilibrio.MSFlexGrid.cols = CantRegistros + 1
         ArepPuntoEquilibrio.MSFlexGrid.ColWidth(i) = 1500
         ArepPuntoEquilibrio.MSFlexGrid.col = i
         ArepPuntoEquilibrio.MSFlexGrid.Text = Format(Totalingresos, "##,##0.00")
         ArepPuntoEquilibrio.MSFlexGrid.TextMatrix(0, i) = ListaMeses(ComboIni - 1)
         
         ArepPuntoEquilibrio.MSFlexGrid1.cols = CantRegistros + 1
         ArepPuntoEquilibrio.MSFlexGrid1.ColWidth(i) = 1500
         ArepPuntoEquilibrio.MSFlexGrid1.col = i
         ArepPuntoEquilibrio.MSFlexGrid1.Text = Format(TotalCostoFijo + TotalGastoFijo, "##,##0.00")
         ArepPuntoEquilibrio.MSFlexGrid1.TextMatrix(0, i) = ListaMeses(ComboIni - 1)
         
         ArepPuntoEquilibrio.MSFlexGrid2.cols = CantRegistros + 1
         ArepPuntoEquilibrio.MSFlexGrid2.ColWidth(i) = 1500
         ArepPuntoEquilibrio.MSFlexGrid2.col = i
         ArepPuntoEquilibrio.MSFlexGrid2.Text = Format((TotalCosto + TotalGastos) - (TotalCostoFijo + TotalGastoFijo), "##,##0.00")
         ArepPuntoEquilibrio.MSFlexGrid2.TextMatrix(0, i) = ListaMeses(ComboIni - 1)
         
         Me.osProgress1.Value = Me.osProgress1.Value + 1
         DoEvents
         ComboIni = ComboIni + 1
         J = J + 1
         Next
    End With


        With ArepPuntoEquilibrio.Chart2D.ChartGroups(1).SeriesLabels
        .Add "Costos Fijos y Gastos Fijos"
        .Add "Ingresos"
        .Add "Costos Variables y Gastos Variables"
        End With
        
          
                ArepPuntoEquilibrio.LblMoneda.Caption = "Expresado en " & Me.CmbMoneda.Text
                ArepPuntoEquilibrio.Logo.Picture = LoadPicture(RutaLogo)
                ArepPuntoEquilibrio.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
                ArepPuntoEquilibrio.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
                ArepPuntoEquilibrio.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
                ArepPuntoEquilibrio.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
                ArepPuntoEquilibrio.LblFechaFin = FechaFin
                ArepPuntoEquilibrio.LblFechaIni = FechaIni
                ArepPuntoEquilibrio.Show 1

      
     Set rpt = New ArepPuntoEquilibrio
'     rpt.DataControl1.ConnectionString = ConexionReporte
'     rpt.DataControl1.Source = SQL
'     fPreview.RunReport rpt
'
'     fPreview.Show 1
'
End Select

Me.Frame6.Visible = False
Me.CmdVerReporte2.Enabled = True
Me.CmdSalir.Enabled = True
Me.CmdVerReporte3.Visible = False
Me.CmdVerReporte2.Visible = False
Me.CmdVerReporte.Visible = True

Exit Sub
TipoErrs:
 MsgBox err.Description


End Sub

Private Sub Command1_Click()
If Me.Option4.Value = True Then
 QueProducto = "CuentaReportes2"
 FrmConsulta.Show 1
Else
 QUIEN = "CuentasReportes2"
 FrmGrupoLista.Show 1
End If
End Sub

Private Sub Command2_Click()

 If Me.FrmDepartamento.Caption = "Fuente" Then
   QueProducto = "Fuente"
   FrmConsulta.Show 1
 Else
   QueProducto = "Departamento"
   FrmConsulta.Show 1
 End If
   
   Me.TxtDptoDesde.Text = FrmConsulta.Codigo
   
End Sub

Private Sub Command3_Click()

   If Me.FrmDepartamento.Caption = "Fuente" Then
    QueProducto = "Fuente"
    FrmConsulta.Show 1
   Else
    QueProducto = "Departamento"
    FrmConsulta.Show 1
   End If
   
   Me.TxtDptoHasta.Text = FrmConsulta.Codigo
End Sub

'Option Explicit

'Dim Tape As New clsTape
Private Sub Form_Load()

'==========================================================
' CREAR CONTROLES DINÁMICOS DE BENEFICIARIO
'==========================================================

SSTab.Tab = 0

Set LblBeneficiario = Me.Controls.Add( _
                        "VB.Label", _
                        "LblBeneficiarioDinamico", _
                        SSTab)

With LblBeneficiario
    .Caption = "Beneficiario:"
    .Left = 4200
    .top = 4200
    .Width = 1200
    .Height = 255
    .Visible = False
End With


Set CmbBeneficiario = Me.Controls.Add( _
                        "TrueOleDBList80.TDBCombo", _
                        "CmbBeneficiarioDinamico", _
                        SSTab)

With CmbBeneficiario
    .Left = 6000
    .top = 4200
    .Width = 2295
    .Height = 315
    .Visible = False
    .DropdownWidth = 5000
End With


    CargarCuentas Me.DBCodigo
    CargarCuentas Me.DBCodigoHasta

If ActivarBeneficiario = 1 Then

'    CmbBeneficiario.Visible = True
'    LblBeneficiario.Visible = True

    CargarBeneficiarios CmbBeneficiario
    CmbBeneficiario.Columns(2).Visible = False
    
  
    

Else

    CmbBeneficiario.Visible = False
    LblBeneficiario.Visible = False
    


End If


Me.CmdVerReporte2.Visible = False
Me.Frame3.Visible = False
Me.Timer1.Enabled = True
Me.Timer1.Interval = Tape.Speed



Me.SSTab.TabVisible(1) = False
Me.SSTab.TabVisible(2) = False
Me.Picture1.BackColor = RGB(161, 193, 245)
Me.ChkQuitarMovimiento.BackColor = RGB(239, 235, 222)

With Me.DtaTasas2

   .ConnectionString = Conexion
End With

With Me.AdoHistorial
   .ConnectionString = Conexion
End With

With Me.AdoConsultas
   .ConnectionString = Conexion
End With

With Me.DtaDatosEmpresa
   .ConnectionString = Conexion
   .RecordSource = "DatosEmpresa"
   .Refresh
End With

With Me.DtaGrupos

   .ConnectionString = Conexion
   .RecordSource = "Grupos"
End With

With Me.DtaElimina

   .ConnectionString = Conexion
End With

With Me.DtaConsulta2

   .ConnectionString = Conexion
End With

With Me.DtaTasas

   .ConnectionString = Conexion
End With


With Me.DtaReportes
   .ConnectionString = Conexion
   .RecordSource = "Reportes"
End With

With Me.DtaBancos

   .ConnectionString = Conexion
End With

With Me.DtaConsulta

   .ConnectionString = Conexion
End With

With Me.DtaHistorial

   .ConnectionString = Conexion
End With

With Me.DtaCuentas

   .ConnectionString = Conexion
   .RecordSource = "Select * from cuentas"
   .Refresh
End With

With Me.DtaPeriodos

   .ConnectionString = Conexion
End With


Me.DTFecha1 = Format(Now, "dd/mm/yyyy")
Me.DTFecha2 = Format(Now, "dd/mm/yyyy")

ConfiguracionReportesBalance
Me.BtnExcel.Visible = False

Select Case QUIEN
 Case "ExportarExcel"
  Me.CmbReportes.AddItem ("BALANZA DE COMPROBACION")
  Me.CmbReportes.AddItem ("REPORTE PBI")
  Me.CmbReportes.AddItem ("TOTAL AUXILIAR DE CUENTAS")
  Me.BtnExcel.Visible = True
  
 
 Case "ReporteCxC"
   Me.CmbReportes.AddItem ("LISTA CUENTAS X COBRAR")
   Me.CmbReportes.AddItem ("LISTA CUENTAS X PAGAR")
   Me.CmbReportes.AddItem ("CUENTAS X COBRAR")
   
 Case "ReportesTerceros"
  Me.CmbReportes.AddItem ("AUXILIAR CUENTAS X BENEFICIARIO")
  Me.CmbReportes.AddItem ("RETENCIONES X PAGAR")
  Me.CmbReportes.AddItem ("RETENCIONES X BENEFICIARIO DETALLE")
  Me.CmbReportes.AddItem ("RETENCIONES X BENEFICIARIO RESUMIDO")
 
'  If ActivarBeneficiario = 1 Then
   
'  End If
   
 Case "ReporteMovimientos"
  Me.Frame7.Visible = False
  Me.CmbReportes.AddItem ("PRESUPUESTO ANUAL")
  Me.CmbReportes.AddItem ("PRESUPUESTO DEL MES")
  Me.CmbReportes.AddItem ("REGISTRO DE MOVIMIENTOS")
  Me.CmbReportes.AddItem ("RETENCIONES X PAGAR")
  Me.CmbReportes.AddItem ("AUXILIAR DE CUENTAS")
  Me.CmbReportes.AddItem ("TOTAL AUXILIAR DE CUENTAS")
  Me.CmbReportes.AddItem ("AUXILIAR x GRUPO")
  Me.CmbReportes.AddItem ("BALANZA DE COMPROBACION")
  

  
'  Me.CmbReportes.AddItem ("ANEXO FISCAL IVA PROVEEDOR")
'  Me.CmbReportes.AddItem ("ANEXO FISCAL IVA CLIENTES")
'  Me.CmbReportes.AddItem ("RETENCIONES EN LA FUENTE I.R X COBRAR")
'  Me.CmbReportes.AddItem ("RETENCIONES EN LA FUENTE I.R X PAGAR")
'  Me.CmbReportes.AddItem ("CUENTAS X COBRAR")
'  Me.CmbReportes.AddItem ("CUENTAS X PAGAR")
  Me.CmbReportes.AddItem ("COMPROBANTE DE DIARIO")
  Me.CmbReportes.AddItem ("LIBRO DIARIO")
  Me.CmbReportes.AddItem ("LIBRO MAYOR")
  Me.CmbReportes.AddItem ("DETALLE DIARIO MAYOR")
  Me.DtaBancos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta From Cuentas ORDER BY Cuentas.CodCuentas"
  Me.DtaBancos.Refresh
  
  






'  Me.DBCodigo.RowSource = Me.DtaBancos
'  Me.DBCodigo.ListField = "CodCuentas"
'  Me.DBCodigoHasta.ListField = "CodCuentas"
'  Me.Label2.Visible = False
'  Me.Label1.Visible = False
'  Me.LblTransaccion.Visible = False
  
 Case "ReporteGenerales"
 Me.Frame7.Visible = False
 Me.Label2.Visible = False
 Me.CmbReportes.AddItem ("ESTRUCTURA DE CUENTAS")
 Me.CmbReportes.AddItem ("CATALOGO RESUMEN")
 Me.CmbReportes.AddItem ("LISTADO CUENTAS")
 Me.CmbReportes.AddItem ("TARJETA ACTIVO FIJO")
 Me.CmbReportes.AddItem ("LISTA ACTIVO FIJO")
 Me.CmbReportes.AddItem ("TARJETA EMPLEADOS")
 Me.CmbReportes.AddItem ("GRUPOS DE CUENTAS")
 Me.CmbReportes.AddItem ("TASAS DE CAMBIO")
 Me.CmbReportes.AddItem ("LISTA DE USUARIOS")
 Me.CmbReportes.AddItem ("TARJETA CONTRATISTA")

 Case "ReporteBancos"
  Me.CmbReportes.AddItem ("CONTROL DE BANCOS")
  Me.CmbReportes.AddItem ("COMPROBANTE DE PAGO")
  Me.DtaBancos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta From Cuentas Where (((Cuentas.TipoCuenta) = 'Bancos' Or (Cuentas.TipoCuenta) = 'Bancos' Or (Cuentas.TipoCuenta) = 'Caja')) ORDER BY Cuentas.CodCuentas"
  Me.DtaBancos.Refresh
  Me.DBCodigo.ListField = "CodCuentas"
 
 Case "EstadosFinancieros"
  Me.CmbReportes.AddItem ("BALANCE GENERAL")
  Me.CmbReportes.AddItem ("BALANCE ACUMULADO")
  Me.CmbReportes.AddItem ("BALANCE HISTORICO")
  Me.CmbReportes.AddItem ("BALANCE GENERAL RESUMEN")
  Me.CmbReportes.AddItem ("BALANCE GENERAL RESUMEN ANEXOS")
'  Me.CmbReportes.AddItem ("BALANCE GENERAL TRADICIONAL")
  Me.CmbReportes.AddItem ("ESTADO DE RESULTADO")
  Me.CmbReportes.AddItem ("ESTADO DE RESULTADO DPTO")
  Me.CmbReportes.AddItem ("RESULTADO ACUMULADO")
  Me.CmbReportes.AddItem ("RESULTADO HISTORICO")
  Me.CmbReportes.AddItem ("ESTADO DE RESULTADO RESUMEN")
  Me.CmbReportes.AddItem ("ESTADO DE RESULTADO RESUMEN 2")
  Me.CmbReportes.AddItem ("ESTADO DE RESULTADO RESUMEN ANEXOS")
  Me.CmbReportes.AddItem ("ESTADO DE RESULTADO TRADICIONAL")
  Me.Frame7.Visible = False
  Me.Label2.Visible = False
  Me.LblTransaccion.Visible = False
 
 Case "Analisis Financieros"
  Me.CmbReportes.AddItem ("RAZONES FINANCIERAS")
  Me.CmbReportes.AddItem ("COMPARATIVO UTILIDADES")
  Me.CmbReportes.AddItem ("COMPARATIVO INGRESOS VRS GASTOS")
  Me.CmbReportes.AddItem ("PUNTO DE EQUILIBRIO")
 
 End Select
End Sub

Private Sub SmartButton1_Click()

End Sub


Private Sub Option4_Click()
Frame7.Caption = "Realizar Filtrado por Codigo"
Me.TxtDesde.Visible = False
Me.TxtHasta.Visible = False
Me.DBCodigo.Visible = True
Me.DBCodigoHasta.Visible = True

    CargarCuentas Me.DBCodigo
    CargarCuentas Me.DBCodigoHasta

End Sub

Private Sub Option5_Click()
    
    Me.Frame7.Caption = "Realizar Filtrado por Grupo"
    Me.DBCodigo.Visible = False
    Me.DBCodigoHasta.Visible = False
    Me.TxtDesde.Visible = True
    Me.TxtHasta.Visible = True
End Sub

Private Sub Option9_Click()
Me.Frame7.Caption = "Realizar el Filtrado por Cuenta de Mayor"
Me.DBCodigo.Visible = False
Me.DBCodigoHasta.Visible = False
Me.TxtDesde.Visible = True
Me.TxtHasta.Visible = True
End Sub

Private Sub Timer1_Timer()
On Error GoTo TipoErrs
Dim intWidth As Integer
Dim intLeft As Integer      'Posición izquierda
Dim objImage As Control     'Control Image
Dim objImage1 As Control
Randomize
'Dim intLeft As Integer      'Posición izquierda
    'Dim objImage As Control     'Control Image
    Randomize   ' Inicializa el generaor de números aleatorios.


    ' Obtiene la anchura de la presentación
    intWidth = picTV.Width
    'Llama al método de la clase Tape
    ' para reproducir la cinta.
    Tape.Animate intWidth
    
    ' Obtiene la propiedad Left a partir de la clase
   intLeft = Tape.Left

If img1.Visible = True Then
        img1.Visible = False
        Set objImage = Img2
    Else
        img1.Visible = True
        Set objImage = img1
    End If
    
 If Lb0.Visible = True Then
   Lb1.Visible = True
   Lb0.Visible = False
   
 ElseIf Lb1.Visible = True Then
    Lb1.Visible = False
    Lb2.Visible = True
 ElseIf Lb2.Visible = True Then
    Lb2.Visible = False
    Lb3.Visible = True
ElseIf Lb3.Visible = True Then
    Lb3.Visible = False
    Lb4.Visible = True
ElseIf Lb4.Visible = True Then
    Lb4.Visible = False
    Lb5.Visible = True
  ElseIf Lb5.Visible = True Then
    Lb5.Visible = False
    Lb6.Visible = True
  ElseIf Lb6.Visible = True Then
    Lb6.Visible = False
    Lb7.Visible = True
  ElseIf Lb7.Visible = True Then
    Lb7.Visible = False
    Lb8.Visible = True
  ElseIf Lb8.Visible = True Then
    Lb8.Visible = False
    Lb9.Visible = True
  ElseIf Lb9.Visible = True Then
    Lb9.Visible = False
    Lb10.Visible = True
  ElseIf Lb10.Visible = True Then
    Lb10.Visible = False
    Lb11.Visible = True
  ElseIf Lb11.Visible = True Then
    Lb11.Visible = False
    Lb12.Visible = True
  ElseIf Lb12.Visible = True Then
    Lb12.Visible = False
    Lb13.Visible = True
  ElseIf Lb13.Visible = True Then
    Lb13.Visible = False
    Lb14.Visible = True
  ElseIf Lb14.Visible = True Then
    Lb14.Visible = False
    Lb15.Visible = True
  ElseIf Lb15.Visible = True Then
    Lb15.Visible = False
    Lb0.Visible = True
    
 End If

' Borra la presentación
    picTV.Cls
    ' Muestra la nueva imagen en la nueva posición
    picTV.PaintPicture objImage.Picture, intLeft, 100, 800, 800
 Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub

Function ExtraerCodigo(cadena As String) As String
    Dim Caracter As String
    For i = 1 To Len(cadena) 'busca algun caracter que no sea numero ni punto
        Caracter = Mid(cadena, i, 1)
        If Asc(Caracter) < vbKey0 Or Asc(Caracter) > vbKey9 Then
            If Asc(Caracter) <> vbKeyDecimal Then
                Exit For
            End If
        Else
            
        End If
    Next i
    
End Function

Private Sub GenerarAuxiliarCuentasXBeneficiario()

    Dim Sql As String
    Dim CodigoCuentaDesde As String
    Dim CodigoCuentaHasta As String
    Dim CodigoBeneficiario As String

    Dim rpt As Object
    Dim fPreview As New FrmPreview
    


        If Me.CmbMoneda.Text = "Córdobas" Then
            Moneda = "Cordobas"
        Else
            Moneda = "Dolares"
        End If


    '==========================================================
    ' CUENTA DESDE
    '==========================================================

    If Me.DBCodigo.Text = "" Then

        Me.DtaConsulta.RecordSource = _
            "SELECT Cuentas.* FROM Cuentas ORDER BY CodCuentas"

        Me.DtaConsulta.Refresh

        If Not Me.DtaConsulta.Recordset.EOF Then

            Me.DtaConsulta.Recordset.MoveFirst

            CodigoCuentaDesde = _
                Me.DtaConsulta.Recordset("CodCuentas")

        End If

    Else

        CodigoCuentaDesde = Me.DBCodigo.Text

    End If


    '==========================================================
    ' CUENTA HASTA
    '==========================================================

    If Me.DBCodigoHasta.Text = "" Then

        Me.DtaConsulta.RecordSource = _
            "SELECT Cuentas.* FROM Cuentas ORDER BY CodCuentas"

        Me.DtaConsulta.Refresh

        If Not Me.DtaConsulta.Recordset.EOF Then

            Me.DtaConsulta.Recordset.MoveLast

            CodigoCuentaHasta = _
                Me.DtaConsulta.Recordset("CodCuentas")

        End If

    Else

        CodigoCuentaHasta = Me.DBCodigoHasta.Text

    End If


    '==========================================================
    ' BENEFICIARIO
    '==========================================================

    CodigoBeneficiario = Trim$(CmbBeneficiario.Text)


    '==========================================================
    ' GENERAR SQL
    '==========================================================

    Sql = ObtenerSqlAuxiliarPorBeneficiario( _
                Me.DTFecha1.Value, _
                Me.DTFecha2.Value, _
                CodigoCuentaDesde, _
                CodigoCuentaHasta, _
                "", _
                "", _
                Me.CmbMoneda.Text, _
                CodigoBeneficiario)

MsgBox right$(Sql, 500)
    '==========================================================
    ' EJECUTAR REPORTE
    '==========================================================

    Set rpt = New ArepAuxiliar

    rpt.DataControl1.ConnectionString = ConexionReporte
    rpt.DataControl1.Source = Sql

    fPreview.RunReport rpt
    fPreview.Show 1

End Sub

Private Sub ReporteRetencionesBeneficiario(ByVal Resumido As Boolean)

    Dim CodigoCuentaDesde As String
    Dim CodigoCuentaHasta As String
    Dim Rs As ADODB.Recordset

    Dim rpt As Object
    Dim fPreview As New FrmPreview

    On Error GoTo ErrReporte

    If UCase$(Left$(Trim$(Me.CmbMoneda.Text), 1)) = "C" Then
        Moneda = "Cordobas"
    Else
        Moneda = "Dolares"
    End If

    '--- Validar fechas
    If Me.DTFecha1.Value > Me.DTFecha2.Value Then
        MsgBox "La fecha inicial no puede ser mayor que la fecha final.", _
               vbExclamation, "Retenciones por Beneficiario"
        Exit Sub
    End If

    '--- Rango de cuentas (vacio = sin limite)
    CodigoCuentaDesde = Trim$(Me.DBCodigo.Text)
    CodigoCuentaHasta = Trim$(Me.DBCodigoHasta.Text)

    '--- Datos: Contabilidad (Conexion) + Beneficiario (ConexionFacturacion)
    Set Rs = ObtenerRsRetencionesBeneficiario(Me.DTFecha1.Value, Me.DTFecha2.Value, _
                                              CodigoCuentaDesde, CodigoCuentaHasta, Resumido)

    If Rs.RecordCount = 0 Then
        MsgBox "No hay retenciones en el rango seleccionado.", _
               vbInformation, "Retenciones por Beneficiario"
        Exit Sub
    End If

    '--- Reporte
    If Resumido Then
        Set rpt = New ArepRetencionesBeneficiarioResumido
    Else
        Set rpt = New ArepRetencionesBeneficiario
    End If

    rpt.FechaDesde = Me.DTFecha1.Value
    rpt.FechaHasta = Me.DTFecha2.Value
    rpt.CuentaDesde = CodigoCuentaDesde
    rpt.CuentaHasta = CodigoCuentaHasta

    Set rpt.DataControl1.Recordset = Rs

    fPreview.RunReport rpt
    fPreview.Show 1

    Exit Sub

ErrReporte:
    MsgBox "Error al generar el reporte de retenciones por beneficiario." & vbCrLf & _
           "Numero: " & err.Number & vbCrLf & _
           "Descripcion: " & err.Description, _
           vbExclamation, "Retenciones por Beneficiario"

End Sub

Private Sub ReporteRetenciones()

    Dim Sql As String
    Dim CodigoCuentaDesde As String
    Dim CodigoCuentaHasta As String

    Dim rpt As ArepRetenciones
    Dim fPreview As New FrmPreview

    On Error GoTo ErrReporte

    If Me.CmbMoneda.Text = "Córdobas" Then
        Moneda = "Cordobas"
    Else
        Moneda = "Dolares"
    End If

    '--- Validar fechas
    If Me.DTFecha1.Value > Me.DTFecha2.Value Then
        MsgBox "La fecha inicial no puede ser mayor que la fecha final.", _
               vbExclamation, "Reporte de Retenciones"
        Exit Sub
    End If

    '--- Rango de cuentas (vacío = sin límite)
    CodigoCuentaDesde = Trim$(Me.DBCodigo.Text)
    CodigoCuentaHasta = Trim$(Me.DBCodigoHasta.Text)

    '--- SQL
    Sql = ObtenerSqlRetenciones(Me.DTFecha1.Value, Me.DTFecha2.Value, _
                                CodigoCuentaDesde, CodigoCuentaHasta)

    '--- Reporte
    Set rpt = New ArepRetenciones

    rpt.FechaDesde = Me.DTFecha1.Value
    rpt.FechaHasta = Me.DTFecha2.Value
    rpt.CuentaDesde = CodigoCuentaDesde
    rpt.CuentaHasta = CodigoCuentaHasta

    rpt.DataControl1.ConnectionString = ConexionReporte
    rpt.DataControl1.Source = Sql

    fPreview.RunReport rpt
    fPreview.Show 1

    Exit Sub

ErrReporte:
    MsgBox "Error al generar el reporte de retenciones." & vbCrLf & _
           "Número: " & err.Number & vbCrLf & _
           "Descripción: " & err.Description, _
           vbExclamation, "Reporte de Retenciones"

End Sub

Public Function ObtenerSqlRetenciones( _
    ByVal FechaDesde As Date, _
    ByVal FechaHasta As Date, _
    ByVal CuentaDesde As String, _
    ByVal CuentaHasta As String) As String

    Dim Sql As String

    Sql = ""
    Sql = Sql & "SELECT "
    Sql = Sql & "T.FechaTransaccion, T.NPeriodo, T.NumeroMovimiento, T.NTransaccion, "
    Sql = Sql & "T.CodCuentas, C.DescripcionCuentas AS NombreCuenta, C.DescRetencion, "
    Sql = Sql & "T.Cod_Tercero, T.Beneficiario, "
    Sql = Sql & "T.FacturaNo, T.NumNomina, T.MontoBase, "
    Sql = Sql & "T.DescripcionMovimiento, T.VoucherNo, T.ChequeNo, T.Fuente, "
    Sql = Sql & "ISNULL(T.Debito,0) AS Debito, "
    Sql = Sql & "ISNULL(T.Credito,0) AS Credito, "
    Sql = Sql & "(ISNULL(T.Credito,0) - ISNULL(T.Debito,0)) AS Retenido "

    Sql = Sql & "FROM Transacciones AS T "
    Sql = Sql & "INNER JOIN Cuentas AS C ON T.CodCuentas = C.CodCuentas "

    Sql = Sql & "WHERE ISNULL(C.CausaRetencion,0) = 1 "

    If Len(Trim$(CuentaDesde)) > 0 Then
        Sql = Sql & "AND T.CodCuentas >= '" & Replace(Trim$(CuentaDesde), "'", "''") & "' "
    End If

    If Len(Trim$(CuentaHasta)) > 0 Then
        Sql = Sql & "AND T.CodCuentas <= '" & Replace(Trim$(CuentaHasta), "'", "''") & "' "
    End If

    Sql = Sql & "AND T.FechaTransaccion >= '" & Format$(FechaDesde, "yyyymmdd") & "' "
    Sql = Sql & "AND T.FechaTransaccion < '" & Format$(DateAdd("d", 1, FechaHasta), "yyyymmdd") & "' "

    Sql = Sql & "ORDER BY T.CodCuentas, T.FechaTransaccion, T.NumeroMovimiento, T.NTransaccion"

    ObtenerSqlRetenciones = Sql

End Function
