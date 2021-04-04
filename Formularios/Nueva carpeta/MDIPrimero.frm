VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Begin VB.MDIForm MDIPrimero 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Contable"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "MDIPrimero.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPrimero.frx":628A
   WindowState     =   2  'Maximized
   Begin XtremeTaskPanel.TaskPanel TaskPanel1 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11880
      _Version        =   786432
      _ExtentX        =   20955
      _ExtentY        =   53
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
      Align           =   3  'Align Left
      Height          =   4470
      Left            =   0
      TabIndex        =   19
      Top             =   900
      Width           =   2220
      _Version        =   786432
      _ExtentX        =   3916
      _ExtentY        =   7885
      _StockProps     =   64
      Animation       =   1
      ItemLayout      =   3
      HotTrackStyle   =   1
   End
   Begin MSAdodcLib.Adodc AdoConsulta 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6795
      Visible         =   0   'False
      Width           =   11880
      _ExtentX        =   20955
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
   Begin ACTIVESKINLibCtl.Skin SkinBlanco 
      Left            =   4440
      OleObjectBlob   =   "MDIPrimero.frx":23450C
      Top             =   6000
   End
   Begin MSComctlLib.ImageList ImageList12 
      Left            =   4560
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":234740
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":237D4A
            Key             =   "Empleados"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":23C50C
            Key             =   "Auxiliar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":23E29E
            Key             =   "Tasas"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":23EF78
            Key             =   "Cuentas"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":23F852
            Key             =   "Usuarios"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":24012C
            Key             =   "Contratistas"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":240A06
            Key             =   "ActivoFijo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2412E0
            Key             =   "Transacciones"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":241FBA
            Key             =   "Cheques"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":242C94
            Key             =   "CalcularDepreciacion"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":24356E
            Key             =   "Presupuesto"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":244248
            Key             =   "Periodos"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":245FDA
            Key             =   "GrupoCuentas"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2468B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":24718E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   0
      ScaleHeight     =   305.455
      ScaleMode       =   0  'User
      ScaleWidth      =   11850
      TabIndex        =   1
      Top             =   30
      Width           =   11880
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   6720
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin SmartButtonProject.SmartButton CmdMovimiento 
         Height          =   690
         Left            =   7800
         TabIndex        =   2
         ToolTipText     =   "Modulo de Activo Fijo"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Activo Fijo"
         Picture         =   "MDIPrimero.frx":247A68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdEmpleado 
         Height          =   690
         Left            =   4920
         TabIndex        =   3
         ToolTipText     =   "Registro de Empleados"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Empleados"
         Picture         =   "MDIPrimero.frx":247FF2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdActivar 
         Height          =   690
         Left            =   5880
         TabIndex        =   4
         ToolTipText     =   "Modulo de Cheques"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Cheques"
         Picture         =   "MDIPrimero.frx":2485B6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton Cmd13vo 
         Height          =   690
         Left            =   9720
         TabIndex        =   5
         ToolTipText     =   "Tasas de Cambios"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Tasas"
         Picture         =   "MDIPrimero.frx":248BB4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdCalcular 
         Height          =   690
         Left            =   6840
         TabIndex        =   6
         ToolTipText     =   "Modulo de Presupuesto"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Presup."
         Picture         =   "MDIPrimero.frx":249130
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdDespido 
         Height          =   690
         Left            =   8760
         TabIndex        =   7
         ToolTipText     =   "Modulo de Transacciones"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Transacc."
         Picture         =   "MDIPrimero.frx":2497EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdSubsidio 
         Height          =   690
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Cuentas Contables del Sistema"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Cuentas"
         Picture         =   "MDIPrimero.frx":249E14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdUsuario 
         Height          =   690
         Left            =   11640
         TabIndex        =   9
         ToolTipText     =   "Registro de Usuarios del Sistema"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Usuarios"
         Picture         =   "MDIPrimero.frx":24A4A5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdSalir 
         Height          =   690
         Left            =   14280
         TabIndex        =   10
         ToolTipText     =   "Boton de Salir del sistema"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Salir"
         Picture         =   "MDIPrimero.frx":24A9FD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdAdelanto 
         Height          =   690
         Left            =   3000
         TabIndex        =   11
         ToolTipText     =   "Contactos o Contratistas"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Contratistas"
         Picture         =   "MDIPrimero.frx":24E1A7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdInss 
         Height          =   690
         Left            =   3960
         TabIndex        =   12
         ToolTipText     =   "Tabla de Periodos"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Periodos"
         Picture         =   "MDIPrimero.frx":24E7A7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdGrupo 
         Height          =   690
         Left            =   10680
         TabIndex        =   13
         ToolTipText     =   "Grupo de Cuentas"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Grupo Cta"
         Picture         =   "MDIPrimero.frx":24EDEC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdMayor 
         Height          =   690
         Left            =   1080
         TabIndex        =   14
         ToolTipText     =   "Boton de Salir del sistema"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Ctas Mayor"
         Picture         =   "MDIPrimero.frx":24F376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdAuxiliar 
         Height          =   690
         Left            =   2040
         TabIndex        =   15
         ToolTipText     =   "Contactos o Contratistas"
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Auxilar Ctas"
         Picture         =   "MDIPrimero.frx":24F9EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdRespaldar 
         Height          =   690
         Left            =   12600
         TabIndex        =   16
         ToolTipText     =   "Realizar Respaldos"
         Top             =   45
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Respaldar"
         Picture         =   "MDIPrimero.frx":250032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdConfiguracion 
         Height          =   690
         Left            =   13440
         TabIndex        =   17
         ToolTipText     =   "Configuración General"
         Top             =   45
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1217
         ForeColor       =   8388608
         Caption         =   "Config."
         Picture         =   "MDIPrimero.frx":2505E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
   End
   Begin MSAdodcLib.Adodc DtaNacceso 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   979
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
   Begin MSAdodcLib.Adodc DtaTasas 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      Top             =   5745
      Visible         =   0   'False
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   873
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
   Begin ComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   7125
      WhatsThisHelpID =   1
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   714
      SimpleText      =   "Programa Bajo Licencia de Juan"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Picture         =   "MDIPrimero.frx":250C85
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7937
            MinWidth        =   7937
            Text            =   "Licencia: Juan"
            TextSave        =   "Licencia: Juan"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1393
            MinWidth        =   1393
            TextSave        =   "NÚM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "08:52 a.m."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "MDIPrimero.frx":250F9F
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2512B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2548C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":259085
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":25AE17
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":25BAF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":25C3CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":25CCA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":25D57F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":25DE59
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":25EB33
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":25F80D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2600E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":260DC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":262B53
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":26342D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":263D07
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2645E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2652BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":265AE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":26630F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2160
      OleObjectBlob   =   "MDIPrimero.frx":266FE9
      Top             =   3240
   End
   Begin MSAdodcLib.Adodc AdoConfiguracion 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   5370
      Visible         =   0   'False
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "AdoConfiguracion"
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5160
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   45
      ImageHeight     =   45
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   33
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C2816
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C30A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C3A3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C439A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C4B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C551D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C5D5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C6646
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C6FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C7796
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C7EC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C8679
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C8E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4C9806
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4CA137
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4CAA79
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4CB3E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4CBE06
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4CC649
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4CCDE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4CD79F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4CE0F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4CE9E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4CF1DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4CFBBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4D0567
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4D0DCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4D17A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4D21BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4D2988
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4D44DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4D602C
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":4D7B7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PopupControl PopupControl1 
      Left            =   6360
      Top             =   3240
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      VisualTheme     =   4
   End
End
Attribute VB_Name = "MDIPrimero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim arrPanes(1 To 6) As frmPane


Private Sub Cmd13vo_Click()
frmTasa2.Show 1
End Sub

Private Sub CmdActivar_Click()
FrmCheque.Show
End Sub

Private Sub CmdAdelanto_Click()
FrmContactos.Show
End Sub

Private Sub CmdAuxiliar_Click()
FrmAuxiliarCuentas.Show 1
End Sub

Private Sub CmdCalcular_Click()
FrmPresupuesto.Show 1
End Sub

Private Sub CmdConfiguracion_Click()
    FrmConfiguracion.Show
End Sub

Private Sub CmdDespido_Click()
FrmTransacciones.Show
End Sub

Private Sub CmdEmpleado_Click()
FrmEmpleados.Show
End Sub

Private Sub CmdGrupo_Click()
FrmGrupo.Show
End Sub

Private Sub CmdInss_Click()
FrmPeriodos.Show 1
End Sub

Private Sub CmdMayor_Click()
FrmCuentasMayor.Show
End Sub

Private Sub CmdMovimiento_Click()
FrmActivoFijo.Show
End Sub

Private Sub CmdRespaldar_Click()
    FrmRespaldar.Show vbModal
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSubsidio_Click()
FrmCuentas.Show
End Sub

Private Sub CmdUsuario_Click()
FrmUsuarios.Show
End Sub

Private Sub CmdUsurio_Click()

End Sub

Private Sub Command1_Click()
    FrmReporteComprobantes.Show
End Sub

Private Sub MDIForm_Activate()
AdoConfiguracion.ConnectionString = Conexion
AdoConfiguracion.RecordSource = "SELECT * FROM DatosEmpresa"
AdoConfiguracion.Refresh
End Sub

Sub CreateTaskPanel()


    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    Set Group = wndTaskPanel.Groups.Add(100, "Cuentas")
    Group.Tooltip = "cuentas del sistema Contable"
    Group.Special = True
    Group.Items.Add 1, "Cuentas", xtpTaskItemTypeLink, 2
    Group.Items.Add 2, "Grupo de Cuentas", xtpTaskItemTypeLink, 4
    Group.Items.Add 3, "Cuentas de Mayor", xtpTaskItemTypeLink, 3
    Group.Items.Add 4, "Auxiliar de Cuentas", xtpTaskItemTypeLink, 1
    Group.Items.Add 5, "Reporte Diario", xtpTaskItemTypeLink, 5
    
    Set Group = wndTaskPanel.Groups.Add(100, "Catalogos")
    Group.Tooltip = "Catalogo del sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 6, "Empleados", xtpTaskItemTypeLink, 6
    Group.Items.Add 7, "Contratistas", xtpTaskItemTypeLink, 7
    Group.Items.Add 8, "Periodo", xtpTaskItemTypeLink, 8
    Group.Items.Add 9, "Activo Fijo", xtpTaskItemTypeLink, 9
    Group.Items.Add 10, "Usuarios", xtpTaskItemTypeLink, 10
    Group.Items.Add 11, "Tasas de Cambio", xtpTaskItemTypeLink, 11
    Group.Items.Add 12, "Departamento", xtpTaskItemTypeLink, 12
    Group.Items.Add 13, "Niveles de Acceso", xtpTaskItemTypeLink, 13
    
    
    Set Group = wndTaskPanel.Groups.Add(100, "Procesos")
    Group.Tooltip = "Procesos del Sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 13, "Transacciones", xtpTaskItemTypeLink, 14
    Group.Items.Add 13, "Cheques", xtpTaskItemTypeLink, 15
    Group.Items.Add 13, "Calcular Depreciacion", xtpTaskItemTypeLink, 16
    Group.Items.Add 13, "Presupuesto", xtpTaskItemTypeLink, 17
    Group.Items.Add 13, "Prorrateo", xtpTaskItemTypeLink, 32
    
    Set Group = wndTaskPanel.Groups.Add(100, "Opciones")
    Group.Tooltip = "Procesos del Sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 13, "Calculadora", xtpTaskItemTypeLink, 18
    Group.Items.Add 13, "Informacion de Usuarios", xtpTaskItemTypeLink, 19
    Group.Items.Add 13, "Configuracion", xtpTaskItemTypeLink, 28
    Group.Items.Add 13, "Respaldar", xtpTaskItemTypeLink, 29
    Group.Items.Add 13, "Compañia", xtpTaskItemTypeLink, 33
    
    Set Group = wndTaskPanel.Groups.Add(100, "Reportes")
    Group.Tooltip = "Procesos del Sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 13, "Reportes Generales", xtpTaskItemTypeLink, 20
    Group.Items.Add 13, "Reportes de Movimientos", xtpTaskItemTypeLink, 21
    Group.Items.Add 13, "Reportes de Bancos", xtpTaskItemTypeLink, 22
    Group.Items.Add 13, "Estados Financieros", xtpTaskItemTypeLink, 23
    Group.Items.Add 13, "Analisis Financieros", xtpTaskItemTypeLink, 31
    
    Set Group = wndTaskPanel.Groups.Add(100, "Ayuda")
    Group.Tooltip = "Procesos del Sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 13, "Configuracion de Cheques", xtpTaskItemTypeLink, 24
    Group.Items.Add 13, "Importar Transacciones", xtpTaskItemTypeLink, 25
    Group.Items.Add 13, "Importar Cuentas", xtpTaskItemTypeLink, 26
    Group.Items.Add 13, "Auditor", xtpTaskItemTypeLink, 27
    
     
    wndTaskPanel.SetImageList Me.ImageList2
End Sub

Private Sub MDIForm_Load()
On Error GoTo TipoErrs




Set Ejecutar = New ADODB.Connection
Ejecutar.ConnectionString = Conexion
Ejecutar.Open

Dim SqlSuspenciones As String, TipoAcceso As String
Dim VerificaTasa As Boolean
Dim Entrar As Boolean
Dim FechaIni As Date
Dim FechaFin As Date
Dim Encontrado As Boolean
Dim Fecha As String
Dim NumFecha As Long

Tasa = True

'With Me.DtaPassword
'    .ConnectionString = Conexion
'    .RecordSource = "Usuarios"
'    .Refresh
'End With

With Me.AdoConsulta
   .ConnectionString = Conexion
End With

With Me.DtaTasas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.AdoConfiguracion
   .ConnectionString = Conexion
   .RecordSource = "SELECT * FROM DatosEmpresa"
   .Refresh
End With

With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

MDIPrimero.Picture = LoadPicture(RutaIconos + "\Zw.bmp")
'AdoConfiguracion.ConnectionString = Conexion
'AdoConfiguracion.RecordSource = "SELECT * FROM DatosEmpresa"
'AdoConfiguracion.Refresh

'
'Me.Picture1.BackColor = RGB(236, 233, 216)
'Me.CmdAuxiliar.BackColor = RGB(236, 233, 216)
'Me.CmdMayor.BackColor = RGB(236, 233, 216)
'Me.CmdActivar.BackColor = RGB(236, 233, 216)
'Me.CmdEmpleado.BackColor = RGB(236, 233, 216)
'Me.CmdInss.BackColor = RGB(236, 233, 216)
'Me.CmdSalir.BackColor = RGB(236, 233, 216)
'Me.Cmd13vo.BackColor = RGB(236, 233, 216)
'Me.CmdAdelanto.BackColor = RGB(236, 233, 216)
'Me.CmdCalcular.BackColor = RGB(236, 233, 216)
'Me.CmdDespido.BackColor = RGB(236, 233, 216)
'Me.CmdMovimiento.BackColor = RGB(236, 233, 216)
'Me.CmdSubsidio.BackColor = RGB(236, 233, 216)
'Me.CmdUsuario.BackColor = RGB(236, 233, 216)
'Me.CmdGrupo.BackColor = RGB(236, 233, 216)
'Me.CmdRespaldar.BackColor = RGB(236, 233, 216)
'Me.CmdConfiguracion.BackColor = RGB(236, 233, 216)
'Me.SmartMenuXP1.BackColor = RGB(236, 233, 216)
'Me.SSListBarVentas.CaptionBackColor = RGB(236, 233, 216)

'//////////BARRA COLOR AZUL/////////////////////

Me.Picture1.BackColor = RGB(173, 199, 236)
Me.CmdAuxiliar.BackColor = RGB(173, 199, 236)
Me.CmdMayor.BackColor = RGB(173, 199, 236)
Me.CmdActivar.BackColor = RGB(173, 199, 236)
Me.CmdEmpleado.BackColor = RGB(173, 199, 236)
Me.CmdInss.BackColor = RGB(173, 199, 236)
Me.CmdSalir.BackColor = RGB(173, 199, 236)
Me.Cmd13vo.BackColor = RGB(173, 199, 236)
Me.CmdAdelanto.BackColor = RGB(173, 199, 236)
Me.CmdCalcular.BackColor = RGB(173, 199, 236)
Me.CmdDespido.BackColor = RGB(173, 199, 236)
Me.CmdMovimiento.BackColor = RGB(173, 199, 236)
Me.CmdSubsidio.BackColor = RGB(173, 199, 236)
Me.CmdUsuario.BackColor = RGB(173, 199, 236)
Me.CmdGrupo.BackColor = RGB(173, 199, 236)
Me.CmdRespaldar.BackColor = RGB(173, 199, 236)
Me.CmdConfiguracion.BackColor = RGB(173, 199, 236)
'Me.SmartMenuXP1.BackColor = RGB(173, 199, 236)

'Me.CmdCuentas.BackColor = RGB(173, 199, 236)
'Me.CmdCtasMayor.BackColor = RGB(173, 199, 236)



'With SmartMenuXP1.MenuItems
' '1
'    .Add 0, "Archivo", , "&Archivo"
'    .Add "Archivo", "Niveles", , "&Niveles", , vbCtrlMask, vbKeyN
'    .Add "Niveles", "Editar Niveles", , "&Editar Niveles", , vbCtrlMask, vbKeyN
'    .Add "Archivo", "Cuentas", , "&Cuentas", , vbCtrlMask, vbKeyK
'    .Add "Archivo", "CuentasMayor", , "&Cuentas de Mayor", , vbCtrlMask, vbKeyC
'    .Add "Archivo", "AuxiliarCuentas", , "&Auxiliar de Cuentas", , vbCtrlMask, vbKeyA
'    .Add "Archivo", "GrupoCuentas", , "&Grupo de Cuentas", , vbCtrlMask, vbKeyG
'    .Add "Archivo", "Contratistas", , "&Contratistas", , vbCtrlMask, vbKeyZ
'    .Add "Archivo", "Empleado", , "&Empleados", , vbCtrlMask, vbKeyX
'    .Add "Archivo", "Activo Fijo", , "&Activo Fijo", , vbCtrlMask, vbKeyF
'    .Add "Archivo", , smiSeparator
'    .Add "Archivo", "Periodos", , "&Periodos", , vbCtrlMask, vbKeyW
'    .Add "Archivo", , smiSeparator
'    .Add "Archivo", "Salir", , "&Salir", , vbCtrlMask, vbKeyL
'
''15
'    .Add 0, "Proceso", , "&Proceso"
'    .Add "Proceso", "Transacciones", , "&Transacciones", , vbCtrlMask, vbKeyT
'    .Add "Proceso", "Cheques", , "&Cheques", , vbCtrlMask, vbKeyH
'    .Add "Proceso", "Calcular Depreciacion", , "&Calcular Depreciacion", , vbCtrlMask, vbKeyD
'    .Add "Proceso", , smiSeparator
'    .Add "Proceso", "Presupuesto", , "&Presupuesto", , vbCtrlMask, vbKeyP
'    .Add "Proceso", , smiSeparator
'    '.Add "Proceso", "Adelantos y Justificacion", , "Adelantos y Justificacion", , vbCtrlMask, vbKeyA
'    '.Add "Proceso", "Auxiliar de Cuentas", , "Auxiliar de Cuentas", , vbCtrlMask, vbKeyX
'    '.Add "Proceso", , smiSeparator
'    '.Add "Proceso", "Conciliacion Bancaria", , "Conciliacion Bancaria", , vbCtrlMask, vbKeyR
' '22
'    .Add 0, "Opciones", , "&Opciones"
'    .Add "Opciones", "Usuarios", , "&Usuarios", , vbCtrlMask, vbKeyU
'    .Add "Opciones", "Moneda", , "&Tasa de Cambio", , vbCtrlMask, vbKeyR
'    .Add "Opciones", "Calculadora", , "&Calculadora", , vbCtrlMask, vbKeyC
'    .Add "Opciones", "Informacion", , "&Informacion Usuarios", , vbCtrlMask, vbKeyI
'    .Add "Opciones", , smiSeparator
'    .Add "Opciones", "BarraH", smiCheckBox, "&Barra de Herramientas", , vbCtrlMask, vbKeyB
'    .Add "Opciones", "BarraE", smiCheckBox, "Barra de &Estado", , vbCtrlMask, vbKeyE
' '33
'    .Add 0, "Reportes", , "R&eportes"
'    .Add "Reportes", "ReporteGenerales", , "R&eportes Generales", , vbCtrlMask, vbKeyE
'    .Add "Reportes", "ReporteMovimientos", , "Re&portes Movimientos", , vbCtrlMask, vbKeyV
'    .Add "Reportes", "ReporteBancos", , "Re&portes Bancos", , vbCtrlMask, vbKeyN
'    .Add "Reportes", "EstadosFinancieros", , "Estados Financieros", , vbCtrlMask, vbKeyV
'  '38
'    .Add 0, "Ventanas", , "&Ventanas"
'    .Add "Ventanas", "Cascadas", , "&Cascadas", , vbCtrlMask, vbKeyD
'    .Add "Ventanas", "Mosaicos", , "&Mosaicos", , vbCtrlMask, vbKeyM
'    .Add "Ventanas", "Organizar", , "&Organizar Iconos", , vbCtrlMask, vbKeyO
'  '42
'    .Add 0, "Ayuda", , "&Ayuda"
'    .Add "Ayuda", "ConfCheque", , "&Configuracion Cheque", , vbCtrlMask, vbKeyT
'    .Add "Ayuda", "ImportarTransacciones", , "I&mportar Transacciones", , vbCtrlMask, vbKey
'    .Add "Ayuda", "ImportarCuentas", , "&Importar Cuentas", , vbCtrlMask, vbKeyS
'    .Add "Ayuda", "Sobre", , "Auditor", , vbCtrlMask, vbKey
'
   
    
'End With

Dim Rutas As String
'Rutas = App.Path & "\conta.ico"
'Set Item = PopupControl1.AddItem(7, 20, 12, 27, "")
'    Item.SetIcon LoadIcon(Rutas, 48, 48), xtpPopupItemIconNormal
'    Item.ID = IDCLOSE
'    Item.Button = True


Dim NombreEmpresa As String, RUC As String
If Not Me.AdoConfiguracion.Recordset.EOF Then
 NombreEmpresa = Me.AdoConfiguracion.Recordset("NombreEmpresa")
 RUC = Me.AdoConfiguracion.Recordset("NumeroRuc")
End If
    
    
Set Item = PopupControl1.AddItem(20, 15, 270, 45, NombreEmpresa)
Item.TextColor = RGB(0, 61, 178)
Item.Bold = True
Set Item = PopupControl1.AddItem(20, 29, 400, 100, "R.U.C :" & RUC)
Item.TextColor = RGB(0, 61, 178)
Item.Bold = True
Set Item = PopupControl1.AddItem(60, 60, 400, 100, "Bienvenido: " & NombreUsuario)
    Item.Bold = True
    PopupControl1.VisualTheme = xtpPopupThemeOffice2003
    PopupControl1.SetSize 300, 110
    Me.PopupControl1.Show
    Me.PopupControl1.Show

    

'//////bUSCO LOS PERMISOS/////////////////////////////

If Not CodigoUsuario = 0 Then

         Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Niveles'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(3) = False
         End If
         Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(4) = False
           Me.CmdSubsidio.Enabled = False
         End If
         Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(5) = False
           Me.CmdGrupo.Enabled = False
         End If
         Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(6) = False
           Me.CmdAdelanto.Enabled = False
         End If
         Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(7) = False
           Me.CmdEmpleado.Enabled = False
         End If
                  Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(8) = False
           Me.CmdMovimiento.Enabled = False
         End If
         
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Periodos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(10) = False
           Me.CmdInss.Enabled = False
         End If

          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(14) = False
           Me.CmdDespido.Enabled = False
         End If
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(15) = False
           Me.CmdActivar.Enabled = False
         End If
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Depreciacion'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(16) = False
         End If
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(18) = False
           Me.CmdCalcular.Enabled = False
         End If
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(21) = False
           Me.CmdUsuario.Enabled = False
         End If
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(22) = False
           Me.Cmd13vo.Enabled = False
         End If
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Generales'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(29) = False
         End If
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Movimientos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(30) = False
         End If
        
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(31) = False
         End If
End If

''///Archivo//////
'Set SmartMenuXP1.MenuItems.Picture(2) = LoadPicture(RutaIconos + "\Niveles.bmp")
'Set SmartMenuXP1.MenuItems.Picture(2) = Me.ImageList1.ListImages(19).ExtractIcon
'Set SmartMenuXP1.MenuItems.Picture(3) = LoadPicture(RutaIconos + "\Editar.bmp")
'Set SmartMenuXP1.MenuItems.Picture(3) = Me.ImageList1.ListImages(20).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(4) = LoadPicture(RutaIconos + "\Cuentas.bmp")
'Set SmartMenuXP1.MenuItems.Picture(4) = Me.ImageList1.ListImages(5).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(5) = LoadPicture(RutaIconos + "\CuentaMayorM.bmp")
'Set SmartMenuXP1.MenuItems.Picture(5) = Me.ImageList1.ListImages(15).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(6) = LoadPicture(RutaIconos + "\AuxiliarCuentasM.bmp")
'Set SmartMenuXP1.MenuItems.Picture(6) = Me.ImageList1.ListImages(3).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(7) = LoadPicture(RutaIconos + "\GrupoC.bmp")
'Set SmartMenuXP1.MenuItems.Picture(7) = Me.ImageList1.ListImages(14).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(8) = LoadPicture(RutaIconos + "\Contratistas.bmp")
'Set SmartMenuXP1.MenuItems.Picture(8) = Me.ImageList1.ListImages(7).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(9) = LoadPicture(RutaIconos + "\Empleado.bmp")
'Set SmartMenuXP1.MenuItems.Picture(9) = Me.ImageList1.ListImages(2).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(10) = LoadPicture(RutaIconos + "\ActivoFijo2.bmp")
'Set SmartMenuXP1.MenuItems.Picture(10) = Me.ImageList1.ListImages(8).ExtractIcon
'
''Set SmartMenuXP1.MenuItems.Picture(12) = LoadPicture(RutaIconos + "\Periodo.bmp")
'Set SmartMenuXP1.MenuItems.Picture(12) = Me.ImageList1.ListImages(13).ExtractIcon
'
'
''jonathan
''Set SmartMenuXP1.MenuItems.Picture(14) = LoadPicture(RutaIconos + "\Salir.bmp")
'
'Set SmartMenuXP1.MenuItems.Picture(14) = ImageList1.ListImages(1).ExtractIcon
''/////Proceso/////
''Set SmartMenuXP1.MenuItems.Picture(16) = LoadPicture(RutaIconos + "\Transacciones.bmp")
'Set SmartMenuXP1.MenuItems.Picture(16) = Me.ImageList1.ListImages(9).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(17) = LoadPicture(RutaIconos + "\ChequesM.bmp")
'Set SmartMenuXP1.MenuItems.Picture(17) = Me.ImageList1.ListImages(10).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(18) = LoadPicture(RutaIconos + "\CalcularDep.bmp")
'Set SmartMenuXP1.MenuItems.Picture(18) = Me.ImageList1.ListImages(11).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(20) = LoadPicture(RutaIconos + "\Presupuesto.bmp")
'Set SmartMenuXP1.MenuItems.Picture(20) = Me.ImageList1.ListImages(12).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(19) = LoadPicture(RutaIconos + "\Adelantos.bmp")
''Set SmartMenuXP1.MenuItems.Picture(20) = LoadPicture(RutaIconos + "\Auxiliar.bmp")
''Set SmartMenuXP1.MenuItems.Picture(22) = LoadPicture(RutaIconos + "\Conciliacion.bmp")
'
''Opciones
''Set SmartMenuXP1.MenuItems.Picture(23) = LoadPicture(RutaIconos + "\Usuarios.bmp")
'Set SmartMenuXP1.MenuItems.Picture(23) = Me.ImageList1.ListImages(6).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(24) = LoadPicture(RutaIconos + "\Tasas.bmp")
'Set SmartMenuXP1.MenuItems.Picture(24) = Me.ImageList1.ListImages(4).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(25) = LoadPicture(RutaIconos + "\Calculadora.bmp")
'Set SmartMenuXP1.MenuItems.Picture(25) = Me.ImageList1.ListImages(16).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(26) = LoadPicture(RutaIconos + "\InformacionUsuario.bmp")
'Set SmartMenuXP1.MenuItems.Picture(26) = Me.ImageList1.ListImages(18).ExtractIcon
'Me.SmartMenuXP1.MenuItems.Value(28) = smiChecked
'Me.SmartMenuXP1.MenuItems.Value(29) = smiChecked
'
''Reportes
'Set SmartMenuXP1.MenuItems.Picture(31) = LoadPicture(RutaIconos + "\Reporte1.bmp")
'Set SmartMenuXP1.MenuItems.Picture(32) = LoadPicture(RutaIconos + "\Reporte2.bmp")
'Set SmartMenuXP1.MenuItems.Picture(33) = LoadPicture(RutaIconos + "\Reporte3.bmp")
'Set SmartMenuXP1.MenuItems.Picture(34) = LoadPicture(RutaIconos + "\Reporte4.bmp")
'
''Ventanas
'Set SmartMenuXP1.MenuItems.Picture(36) = LoadPicture(RutaIconos + "\Cascada.bmp")
'Set SmartMenuXP1.MenuItems.Picture(37) = LoadPicture(RutaIconos + "\Mosaico.bmp")
'Set SmartMenuXP1.MenuItems.Picture(38) = LoadPicture(RutaIconos + "\Organizar.bmp")
'
''Ayudas
'
''Set SmartMenuXP1.MenuItems.Picture(39) = LoadPicture(RutaIconos + "\Ayuda1.bmp")
'Set SmartMenuXP1.MenuItems.Picture(39) = Me.ImageList1.ListImages(17).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(40) = LoadPicture(RutaIconos + "\Ayuda.bmp")
'Set SmartMenuXP1.MenuItems.Picture(40) = Me.ImageList1.ListImages(17).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(41) = LoadPicture(RutaIconos + "\Ayuda.bmp")
'Set SmartMenuXP1.MenuItems.Picture(41) = Me.ImageList1.ListImages(17).ExtractIcon
'
''Set SmartMenuXP1.MenuItems.Picture(42) = LoadPicture(RutaIconos + "\Ayuda.bmp")
'Set SmartMenuXP1.MenuItems.Picture(42) = Me.ImageList1.ListImages(17).ExtractIcon
''Set SmartMenuXP1.MenuItems.Picture(43) = LoadPicture(RutaIconos + "\Ayuda.bmp")
'Set SmartMenuXP1.MenuItems.Picture(43) = Me.ImageList1.ListImages(17).ExtractIcon
'
''////////Registro los datos de la Compañia////////////////////
''DtaEmpresa.Refresh
''Titulo = DtaEmpresa.Recordset.nombreempresa
''Subtitulo = DtaEmpresa.Recordset.Direccion + " RUC: " + DtaEmpresa.Recordset.numeroruc
''RutaLogo = DtaEmpresa.Recordset.RutaLogo
''StatusBar2.Panels(2) = "Licencia: " + Titulo


Fecha = Format(Now, "yyyy/mm/dd")
'NumFecha = Fecha
'DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = " & NumFecha & "))ORDER BY Tasas.FechaTasas"
Me.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE(FechaTasas = CONVERT(DATETIME, '" & Fecha & "', 102)) ORDER BY FechaTasas"
Me.DtaTasas.Refresh

If Not DtaTasas.Recordset.EOF Then
Fecha = Format(DtaTasas.Recordset("FechaTasas"), "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset("MontoCordobas")
   ' MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset("MontoLibras"), "##,##0.00")
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00")
End If

If Not Encontrado Then
  MsgBox "La Tasa de Hoy no ha sido grabada"
  Cancel = 100
  frmTasa2.Show 1
End If

CreateTaskPanel

DoEvents


          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
'           Me.SmartMenuXP1.MenuItems.Enabled(31) = False
         End If

Exit Sub
TipoErrs:
MsgBox err, vbCritical, "SISTEMA DFID"
MsgBox err.Description
End Sub


Private Sub SmartMenuXP1_Click(ByVal ID As Long)
On Error GoTo TipoErrs
Dim AÑO1 As String, AÑO2 As String, AÑO3 As String


With SmartMenuXP1.MenuItems
        Select Case .Key(ID)
    'archivo
          Case "Cuentas"
             FrmCuentas.Show
          Case "Editar Niveles"
              FrmEditarNiveles.Show 1
          Case "GrupoCuentas"
             FrmGrupo.Show
          Case "Contratistas"
            FrmContactos.Show
          Case "Empleado"
            FrmEmpleados.Show
          Case "CuentasMayor"
           FrmCuentasMayor.Show
          Case "AuxiliarCuentas"
             FrmAuxiliarCuentas.Show 1
          Case "Activo Fijo"
            FrmActivoFijo.Show
          Case "Periodos"
            FrmPeriodos.Show 1
          Case "Salir"
            Unload Me
             
      'procesos
        Case "Transacciones"
          FrmTransacciones.Show
        Case "Cheques"
          FrmCheque.Show
       
       Case "Calcular Depreciacion"
          FrmCalcularDepreciacion.Show 1
       Case "Presupuesto"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Presupuesto'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmPresupuesto.Show 1
       Case "Adelantos y Justificacion"
       
       Case "Auxiliar de Cuentas"
    
       Case "Conciliacion Bancaria"
    'Opciones
    
    Case "Usuarios"
         MDIPrimero.MousePointer = 11
         FrmUsuarios.Show
         MDIPrimero.MousePointer = 0
         Case "Moneda"
         MDIPrimero.MousePointer = 11
         frmTasa2.Show
          MDIPrimero.MousePointer = 0
         Case "Calculadora"
          Directorio = Directorio & App.Path & "\Calc.exe"
          Directorio = Shell(Directorio)
          MDIPrimero.MousePointer = 0
         Case "Informacion"
          FrmInforme.Show 1
   'Reportes
 
          
    Case "ReporteGenerales"
        QUIEN = "ReporteGenerales"
      FrmReportes.Show 1
    
    Case "ReporteMovimientos"
      QUIEN = "ReporteMovimientos"
      FrmReportes.Show 1
    
    Case "ReporteBancos"
     QUIEN = "ReporteBancos"
      FrmReportes.Show 1
     Case "EstadosFinancieros"
     QUIEN = "EstadosFinancieros"
      FrmReportes.CmbNivel.AddItem ("0")
      FrmReportes.CmbNivel.AddItem ("1")
      FrmReportes.CmbNivel.AddItem ("2")
      FrmReportes.CmbNivel.AddItem ("3")
      FrmReportes.CmbNivel.AddItem ("4")
      FrmReportes.CmbNivel.AddItem ("5")
      FrmReportes.CmbNivel.AddItem ("6")
      FrmReportes.CmbNivel.AddItem ("7")
      FrmReportes.CmbNivel.AddItem ("8")
      FrmReportes.CmbNivel.AddItem ("9")
      FrmReportes.CmbNivel.AddItem ("10")
      FrmReportes.CmbNivel.AddItem ("11")
      FrmReportes.CmbNivel.AddItem ("12")
      FrmReportes.CmbNivel.AddItem ("13")
      FrmReportes.CmbNivel.AddItem ("14")
      FrmReportes.CmbNivel.AddItem ("15")
      FrmReportes.CmbNivel.AddItem ("16")
      FrmReportes.CmbNivel.AddItem ("17")
      FrmReportes.CmbNivel.AddItem ("18")
      FrmReportes.CmbNivel.AddItem ("19")
      FrmReportes.CmbNivel.AddItem ("20")
      FrmReportes.Label3.Visible = True
      FrmReportes.CmbMoneda.Visible = True
      FrmReportes.Frame1.Visible = False
      FrmReportes.Frame4.Visible = True
      FrmReportes.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      FrmReportes.DtaConsulta.Refresh
      Do While Not FrmReportes.DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
        FrmReportes.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
        FrmReportes.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
         FrmReportes.Option6.Caption = AÑO3
       End If
        
        FrmReportes.DtaConsulta.Recordset.MoveNext
      Loop
      FrmReportes.Show 1
    
    Case "BarraH"
         If Me.Picture1.Visible = True Then
'          Me.SmartMenuXP1.MenuItems.Value(ID) = smiUnchecked
          Me.Picture1.Visible = False
         Else
'          Me.SmartMenuXP1.MenuItems.Value(ID) = smiChecked
          Me.Picture1.Visible = True
         End If
                   
         
         Case "BarraE"
          If StatusBar2.Visible = True Then
'            Me.SmartMenuXP1.MenuItems.Value(ID) = smiUnchecked
            StatusBar2.Visible = False
          Else
              StatusBar2.Visible = True
'              Me.SmartMenuXP1.MenuItems.Value(ID) = smiChecked
          End If
  'Ventana.
        Case "Cascadas"
          MDIPrimero.Arrange vbCascade
        Case "Mosaicos"
         MDIPrimero.Arrange vbTileHorizontal
        Case "Organizar"
         MDIPrimero.Arrange vbArrangeIcons
             
             
   'Ayuda
         Case "ConfCheque"
           FrmConfiguraCheque.Show 1
         Case "ImportarTransacciones"
          FrmImporta.Show 1
        Case "ImportarCuentas"
          FrmImportarCuentas.Show 1
         Case "Sobre"
          FrmAuditor.Show 1
             
        End Select
End With

Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub SSListBarVentas_ListItemClick(ByVal ItemClicked As Listbar.SSListItem)
    Select Case ItemClicked.Key
        Case "Cuentas"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cuentas'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmCuentas.Show
        Case "Grupo de Cuentas"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Grupo Cuentas'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmGrupo.Show
        Case "Cuentas de Mayor"
            FrmCuentasMayor.Show
        Case "Auxiliar de Cuentas"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Contratistas'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmAuxiliarCuentas.Show 1
        Case "Contratistas"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Contratistas'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmContactos.Show
        Case "Periodos"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Periodos'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmPeriodos.Show 1
        Case "Empleados"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Empleados'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmEmpleados.Show
        Case "Cheques"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cheques'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmCheque.Show
        Case "Presupuesto"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Presupuesto'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmPresupuesto.Show 1
        Case "Activo Fijo"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Activo Fijo'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmActivoFijo.Show
        Case "Transacciones"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Transacciones'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmTransacciones.Show
        Case "Tasas"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tasa Cambio'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            frmTasa2.Show 1
        Case "Usuarios"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Usuarios'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmUsuarios.Show
        Case "Calcular Depreciación"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Depreciacion'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmCalcularDepreciacion.Show 1
        Case "Comprobantes"
            FrmReporteComprobantes.Show

    End Select
End Sub

Private Sub wndTaskPanel_GroupExpanding(ByVal Group As XtremeTaskPanel.ITaskPanelGroup, ByVal Expanding As Boolean, Cancel As Boolean)
 If Expanding = True Then
  Select Case Group.Caption
    Case "Cuentas"
'              wndTaskPanel.Groups(1).Expanded = True
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
              wndTaskPanel.Groups(6).Expanded = False
    Case "Catalogos"
              wndTaskPanel.Groups(1).Expanded = False
'              wndTaskPanel.Groups(2).Expanded = True
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
              wndTaskPanel.Groups(6).Expanded = False
    Case "Procesos"
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
'              wndTaskPanel.Groups(3).Expanded = True
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
              wndTaskPanel.Groups(6).Expanded = False
    
    Case "Opciones"
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
'              wndTaskPanel.Groups(4).Expanded = True
              wndTaskPanel.Groups(5).Expanded = False
              wndTaskPanel.Groups(6).Expanded = False
    
    Case "Reportes"
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
'              wndTaskPanel.Groups(5).Expanded = True
              wndTaskPanel.Groups(6).Expanded = False
    
    Case "Ayuda"
    
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
'              wndTaskPanel.Groups(6).Expanded = True
    
  
  End Select
 End If

End Sub

Private Sub wndTaskPanel_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
    
    Select Case Item.Caption
        Case "Compañia"
           FrmCompañia.Show 1
        Case "Calcular Depreciacion"
           FrmCalcularDepreciacion.Show 1
        Case "Cuentas"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cuentas'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            

            FrmCuentas.Show
            

              
            
        Case "Grupo de Cuentas"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Grupo Cuentas'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If

            FrmGrupo.Show
            
        Case "Cuentas de Mayor"

            FrmCuentasMayor.Show
            
        Case "Auxiliar de Cuentas"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Contratistas'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If

            FrmAuxiliarCuentas.Show 1
        Case "Niveles de Acceso"
              FrmEditarNiveles.Show 1
        Case "Contratistas"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Contratistas'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmContactos.Show
        Case "Periodo"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Periodos'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmPeriodos.Show 1
        Case "Empleados"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Empleados'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmEmpleados.Show
        Case "Cheques"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cheques'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmCheque.Show
        Case "Presupuesto"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Presupuesto'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmPresupuesto.Show 1
        Case "Activo Fijo"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Activo Fijo'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmActivoFijo.Show
        Case "Transacciones"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Transacciones'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmTransacciones.Show
        Case "Tasas de Cambio"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tasa Cambio'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            frmTasa2.Show 1
        Case "Usuarios"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Usuarios'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmUsuarios.Show
        Case "Calcular Depreciación"
            Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Depreciacion'))"
            Me.DtaNacceso.Refresh
            If Me.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmCalcularDepreciacion.Show 1
        Case "Prorrateo"
            FrmProrrateo.Show
        Case "Calculadora"
          Directorio = Directorio & App.Path & "\Calc.exe"
          Directorio = Shell(Directorio)
          MDIPrimero.MousePointer = 0
        Case "Informacion de Usuarios"
          FrmInforme.Show 1
          
    Case "Reportes Generales"
        QUIEN = "ReporteGenerales"
      FrmReportes.Show 1
    
    Case "Reportes de Movimientos"
      QUIEN = "ReporteMovimientos"
      FrmReportes.Show 1
    
    Case "Reportes de Bancos"
     QUIEN = "ReporteBancos"
      FrmReportes.Show 1
      
     Case "Estados Financieros"
     QUIEN = "EstadosFinancieros"
      FrmReportes.CmbNivel.AddItem ("0")
      FrmReportes.CmbNivel.AddItem ("1")
      FrmReportes.CmbNivel.AddItem ("2")
      FrmReportes.CmbNivel.AddItem ("3")
      FrmReportes.CmbNivel.AddItem ("4")
      FrmReportes.CmbNivel.AddItem ("5")
      FrmReportes.CmbNivel.AddItem ("6")
      FrmReportes.CmbNivel.AddItem ("7")
      FrmReportes.CmbNivel.AddItem ("8")
      FrmReportes.CmbNivel.AddItem ("9")
      FrmReportes.CmbNivel.AddItem ("10")
      FrmReportes.CmbNivel.AddItem ("11")
      FrmReportes.CmbNivel.AddItem ("12")
      FrmReportes.CmbNivel.AddItem ("13")
      FrmReportes.CmbNivel.AddItem ("14")
      FrmReportes.CmbNivel.AddItem ("15")
      FrmReportes.CmbNivel.AddItem ("16")
      FrmReportes.CmbNivel.AddItem ("17")
      FrmReportes.CmbNivel.AddItem ("18")
      FrmReportes.CmbNivel.AddItem ("19")
      FrmReportes.CmbNivel.AddItem ("20")
      FrmReportes.Label3.Visible = True
      FrmReportes.CmbMoneda.Visible = True
      FrmReportes.Frame1.Visible = False
      FrmReportes.Frame4.Visible = True
      FrmReportes.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      FrmReportes.DtaConsulta.Refresh
      Do While Not FrmReportes.DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
        FrmReportes.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
        FrmReportes.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
         FrmReportes.Option6.Caption = AÑO3
       End If
        
        FrmReportes.DtaConsulta.Recordset.MoveNext
      Loop
      FrmReportes.Show 1
    
    Case "Analisis Financieros"
     QUIEN = "Analisis Financieros"
      FrmReportes.CmbMoneda.Visible = True
      FrmReportes.Frame1.Visible = False
      FrmReportes.Frame4.Visible = True
      FrmReportes.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      FrmReportes.DtaConsulta.Refresh
      Do While Not FrmReportes.DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
        FrmReportes.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
        FrmReportes.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
         FrmReportes.Option6.Caption = AÑO3
       End If
        
        FrmReportes.DtaConsulta.Recordset.MoveNext
      Loop
     FrmReportes.Show 1
      
   Case "Configuracion de Cheques"
           FrmConfiguraCheque.Show 1
   Case "Importar Transacciones"
          FrmImporta.Show 1
   Case "Importar Cuentas"
          FrmImportarCuentas.Show 1
   Case "Auditor"
          FrmAuditor.Show 1
   Case "Respaldar"
           FrmRespaldar.Show vbModal
   Case "Configuracion"
            FrmConfiguracion.Show

    End Select
End Sub


