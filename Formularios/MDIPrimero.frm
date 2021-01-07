VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.CommandBars.v12.0.0.Demo.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.DockingPane.v12.0.0.Demo.ocx"
Begin VB.MDIForm MDIPrimero 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema de Polizas"
   ClientHeight    =   8370
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   Icon            =   "MDIPrimero.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPrimero.frx":57E2
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11880
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   6720
         TabIndex        =   1
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
         Picture         =   "MDIPrimero.frx":4DCEB
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
         Picture         =   "MDIPrimero.frx":4E275
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
         Picture         =   "MDIPrimero.frx":4E839
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
         Picture         =   "MDIPrimero.frx":4EE37
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
         Picture         =   "MDIPrimero.frx":4F3B3
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
         Picture         =   "MDIPrimero.frx":4FA6F
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
         Picture         =   "MDIPrimero.frx":50097
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
         Picture         =   "MDIPrimero.frx":50728
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
         Picture         =   "MDIPrimero.frx":50C80
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
         Picture         =   "MDIPrimero.frx":5442A
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
         Picture         =   "MDIPrimero.frx":54A2A
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
         Picture         =   "MDIPrimero.frx":5506F
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
         Picture         =   "MDIPrimero.frx":555F9
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
         Picture         =   "MDIPrimero.frx":55C6F
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
         Picture         =   "MDIPrimero.frx":562B5
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
         Picture         =   "MDIPrimero.frx":56863
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
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   30
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons2 
      Left            =   3000
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":56F08
            Key             =   ""
            Object.Tag             =   "119"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":574A2
            Key             =   ""
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":57A3C
            Key             =   ""
            Object.Tag             =   "128"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":57FD6
            Key             =   ""
            Object.Tag             =   "115"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":58370
            Key             =   ""
            Object.Tag             =   "130"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":5890A
            Key             =   ""
            Object.Tag             =   "160"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":58EA4
            Key             =   ""
            Object.Tag             =   "116"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":5943E
            Key             =   ""
            Object.Tag             =   "300"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":599D8
            Key             =   ""
            Object.Tag             =   "118"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":59F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":5C2F4
            Key             =   ""
            Object.Tag             =   "204"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":5C88E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":5EC10
            Key             =   ""
            Object.Tag             =   "129"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":5F1AA
            Key             =   ""
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":5F744
            Key             =   ""
            Object.Tag             =   "205"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":5FCDE
            Key             =   ""
            Object.Tag             =   "1331"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":60278
            Key             =   ""
            Object.Tag             =   "1311"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":61F82
            Key             =   ""
            Object.Tag             =   "131"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":6251C
            Key             =   ""
            Object.Tag             =   "0"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":62AB6
            Key             =   ""
            Object.Tag             =   "133"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":63050
            Key             =   ""
            Object.Tag             =   "132"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":635EA
            Key             =   ""
            Object.Tag             =   "134"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":63B84
            Key             =   ""
            Object.Tag             =   "140111"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":6411E
            Key             =   ""
            Object.Tag             =   "139"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":646B8
            Key             =   ""
            Object.Tag             =   "320"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":64C52
            Key             =   ""
            Object.Tag             =   "136"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   60
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":651EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":652FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":65450
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":65562
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":65674
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":65786
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoConsulta 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      Top             =   6090
      Visible         =   0   'False
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   794
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
   Begin MSAdodcLib.Adodc DtaNacceso 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      Top             =   6540
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
      Top             =   7095
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
   Begin MSAdodcLib.Adodc AdoConfiguracion 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   7590
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
   Begin ComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   18
      Top             =   7965
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
            Picture         =   "MDIPrimero.frx":65898
            TextSave        =   ""
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
            TextSave        =   ""
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
            Object.Width           =   1393
            MinWidth        =   1393
            TextSave        =   "NÚM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "10:53 a.m."
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
      MouseIcon       =   "MDIPrimero.frx":65BB2
   End
   Begin ACTIVESKINLibCtl.Skin SkinBlanco 
      Left            =   7440
      OleObjectBlob   =   "MDIPrimero.frx":65ECC
      Top             =   4560
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5160
      OleObjectBlob   =   "MDIPrimero.frx":66100
      Top             =   1800
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4080
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   45
      ImageHeight     =   45
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   34
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C192D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C21B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C2B53
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C34B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C3CB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C4634
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C4E72
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C575D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C60C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C68AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C6FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C7790
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C7F29
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C891D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C924E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2C9B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2CA4FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2CAF1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2CB760
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2CBEFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2CC8B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2CD20D
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2CDAFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2CE2F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2CECD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2CF67E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2CFEE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D08C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D12D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D1A9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D35F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D5143
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D6C95
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D87E7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoConsultaFacturacion 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   794
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
      Caption         =   "AdoConsultaFacturacion"
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
   Begin XtremeSuiteControls.PopupControl PopupControl1 
      Left            =   6240
      Top             =   3840
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      VisualTheme     =   4
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   2400
      Top             =   3600
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   2
      VisualTheme     =   2
   End
   Begin XtremeSuiteControls.PopupControl PopupControl 
      Left            =   5580
      Top             =   3870
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      Animation       =   2
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   5190
      Top             =   3840
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      ScaleMode       =   1
   End
End
Attribute VB_Name = "MDIPrimero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim arrPanes(1 To 2) As frmPane
 'Variable de tipo Ipicturedisp para cargar la imagen


Private Sub DockingPaneManager_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If (Item.Id) = 1 Then
        Set arrPanes(Item.Id) = New frmPane
        Item.Handle = arrPanes(Item.Id).hWnd
    End If

 
    
End Sub

Public Function RibbonBar() As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
End Function



Private Sub CommandBars_GetClientBordersWidth(Left As Long, top As Long, right As Long, bottom As Long)
    
    If Me.StatusBar2.Visible Then
        bottom = StatusBar2.Height
    End If
'

   If Me.Picture1.Visible Then
      top = Me.Picture1.Height
   End If


End Sub


Private Sub MDIForm_Activate()
'AdoConfiguracion.ConnectionString = Conexion
'AdoConfiguracion.RecordSource = "SELECT * FROM DatosEmpresa"
'AdoConfiguracion.Refresh
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
  Dim Directorio As String, RutaUpdate As String
  Dim AÑO1 As String, AÑO2 As String, AÑO3 As String

      

      Select Case Control.Id
        Case 1300: Unload Me
        Case 1700: FrmCuentas.Show
        Case 1701: FrmCuentasMayor.Show
        Case 1702: FrmAuxiliarCuentas.Show 1
        Case 1703: FrmPeriodos.Show 1
        Case 1704: FrmContactos.Show
        Case 1705: FrmEmpleados.Show
        Case 1706: FrmUsuarios.Show
        Case 1707: FrmAgregarActivoFijo.Show 1  'FrmActivoFijo.Show
        Case 1708: frmTasa2.Show 1
        Case 1709: FrmAuditor.Show
        Case 1710
          Directorio = App.Path & "\Calc.exe"
          Directorio = Shell(Directorio)
          MDIPrimero.MousePointer = 0
        Case 1711: FrmCheque.Show
        Case 1712: FrmPresupuesto.Show 1
        Case 1713: FrmRespaldar.Show 1
        Case 1714: FrmCompañia.Show 1
        Case 1715: FrmConfiguracion.Show
        Case 1716: FrmProrrateo.Show
        Case 1717: FrmTransacciones.Show
        Case 1718
                QUIEN = "ReporteGenerales"
                FrmReportes.Label10.Caption = "Reportes Generales"
                FrmReportes.Image2.Picture = LoadPicture(App.Path & "\Imagenes\ReporteGeneral.bmp")
                FrmReportes.Show 1
        Case 1719
              QUIEN = "ReporteMovimientos"
              FrmReportes.Label10.Caption = "Reporte de Movimientos"
              FrmReportes.Image2.Picture = LoadPicture(App.Path & "\Imagenes\ReporteMovimientos.bmp")
              FrmReportes.Show 1
        Case 1720
             QUIEN = "ReporteBancos"
             FrmReportes.Label10.Caption = "Reporte de Bancos"
             FrmReportes.Image2.Picture = LoadPicture(App.Path & "\Imagenes\ReportesBancos.bmp")
             FrmReportes.Show 1
        Case 1721
             QUIEN = "EstadosFinancieros"
             FrmReportes.Label10.Caption = "Estados Financieros"
             FrmReportes.Image2.Picture = LoadPicture(App.Path & "\Imagenes\EstadosFinancieros.bmp")
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
         Case 1722
                 QUIEN = "Analisis Financieros"
                 FrmReportes.Label10.Caption = "Analisis Financieros"
                 FrmReportes.Image2.Picture = LoadPicture(App.Path & "\Imagenes\AnalisisFinanciero2.bmp")
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
                
        Case 1723
              
                  If MDIPrimero.AdoConfiguracion.Recordset.RecordCount > 0 Then
                     If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion) Then
                        FrmConexiones.TxtConexionString.Text = MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion
          
                     Else
                        FrmConexiones.TxtConexionString.Text = ""
                     End If
                  End If
     
  
              FrmConexiones.Option2.Value = True
              FrmConexiones.Image2 = LoadPicture(App.Path & "\Imagenes\ConexionFacturacion.bmp")
              FrmConexiones.lbltitulo.Caption = "Conexion Sistema Facturacion"
              FrmConexiones.Show
       
       Case 1724
             FrmContabilizaFacturacion.Show 1
       
       Case 1725
              If MDIPrimero.AdoConfiguracion.Recordset.RecordCount > 0 Then
                 If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!ConexionNomina) Then
                       FrmConexiones.TxtConexionString.Text = MDIPrimero.AdoConfiguracion.Recordset!ConexionNomina
                 Else
                       FrmConexiones.TxtConexionString.Text = ""
                 End If
              End If
              
              FrmConexiones.Option1.Value = True
              FrmConexiones.Image2 = LoadPicture(App.Path & "\Imagenes\ConexionContabilidad.bmp")
              FrmConexiones.lbltitulo.Caption = "Conexion Sistema Nominas"
              FrmConexiones.Show
              
              
       Case 1726
             FrmContabilizaNomina.Show 1
       Case 1727
             FrmAltaBienes.Show
       Case 1728
             FrmOficinas.Show
       Case 1730
             FrmtrasladoActivos.Show
       Case 1731
         FrmBajaBienes.Show
       Case 1732
         FrmCalcularDepreciacion.Show 1
       Case 1733
         FrmResponsablesAreas.Show
       Case 1734
         FrmMantenimientoActivos.Show 1
       Case 1735
          '--------------------BUSCO SI NO ES NULO ------------------------
          If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!Ruta_Actualizar) Then
             If Dir(MDIPrimero.AdoConfiguracion.Recordset!Ruta_Actualizar) <> "" Then
             
                Open App.Path + "\RutaUpdate.dll" For Output As #1
                  Print #1, ""
                Close #1
                
                Open App.Path + "\RutaUpdate.dll" For Output As #1
                     RutaUpdate = MDIPrimero.AdoConfiguracion.Recordset!Ruta_Actualizar
                     Print #1, RutaUpdate
                Close #1
             
                Directorio = App.Path & "\Actualizar.exe"
                Directorio = Shell(Directorio, vbNormalFocus)
                Unload Me
             Else
                MsgBox "NO existe el Archivo en la Ruta Indicada", vbCritical, "Zeus Contable"
             End If
          Else
            MsgBox "Seleccione una Ruta de Actualizacion", vbCritical, "Zeus Contable"
          End If
       Case 1736
              QUIEN = "ReporteCxC"
              FrmReportes.Label10.Caption = "Reportes CxC y CxP"
              FrmReportes.Image2.Picture = LoadPicture(App.Path & "\Imagenes\ReporteMovimientos.bmp")
              FrmReportes.Show 1
        
       Case 1737
              FrmEgresos.Show
              
       Case 1738
              FrmSolicitudPagoLista.Show
              
       Case 1739: FrmListaChequeReimpresion.Show
       Case 1740: FrmCheque.Show
       Case 1741: FrmSolicitudPagoLista.Show
       Case 1742: FrmEstructuraPresupuesto.Show


        End Select
End Sub



Private Sub MDIForm_Load()



DoEvents

On Error GoTo TipoErrs


MDIPrimero.Picture = LoadPicture(App.Path + "\Imagenes\Zw.jpg")



Set Ejecutar = New ADODB.Connection
Ejecutar.ConnectionString = Conexion
Ejecutar.Open

Dim SqlSuspenciones As String, TipoAcceso As String
Dim VerificaTasa As Boolean, Valor As Double, Cadena2 As String
Dim Entrar As Boolean
Dim FechaIni As Date
Dim FechaFin As Date
Dim Encontrado As Boolean
Dim Fecha As String
Dim NumFecha As Long
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Tasa = True


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

        Unidad = App.Path + "\"
        RutaFoto = App.Path + "\fotos\"
        RutaLogo = App.Path + "\Imagenes\logo.jpg"
        RutaIconos = App.Path + "\Imagenes"
        



Dim Rutas As String

Dim NombreEmpresa As String, RUC As String
If Not Me.AdoConfiguracion.Recordset.EOF Then
 NombreEmpresa = Me.AdoConfiguracion.Recordset("NombreEmpresa")
 RUC = Me.AdoConfiguracion.Recordset("NumeroRuc")
 RutaLogo = Me.AdoConfiguracion.Recordset("DireccionLogo")
' Valor = Me.AdoConfiguracion.Recordset("Valor") + 1
'
' Me.AdoConfiguracion.Recordset("valor") = Valor
' Me.AdoConfiguracion.Recordset.Update
End If
    
    
    
'    If Valor >= 30 Then
'     Unload Me
'    End If
'
Set Item = PopupControl1.AddItem(50, 15, 270, 45, NombreEmpresa)
Item.TextColor = RGB(0, 61, 178)
Item.Bold = True

Set Item = PopupControl1.AddItem(12, 20, 12, 27, "")
Item.SetIcon LoadIcon("Imagenes\Imagen.ico", 32, 32), xtpPopupItemIconNormal

Set Item = PopupControl1.AddItem(50, 29, 400, 100, "R.U.C :" & RUC)
Item.TextColor = RGB(0, 61, 178)
Item.Bold = True

Set Item = PopupControl1.AddItem(60, 60, 400, 100, "Bienvenido: " & NombreUsuario)
    Item.Bold = True
    PopupControl1.VisualTheme = xtpPopupThemeOffice2003
    PopupControl1.SetSize 300, 110
    Me.PopupControl1.Show
    Me.PopupControl1.Show

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

Fecha = Format(Now, "yyyy/mm/dd")
Me.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE(FechaTasas = CONVERT(DATETIME, '" & Fecha & "', 102)) ORDER BY FechaTasas"
Me.DtaTasas.Refresh

If Not DtaTasas.Recordset.EOF Then
Fecha = Format(DtaTasas.Recordset("FechaTasas"), "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset("MontoCordobas")
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00")
End If

If Not Encontrado Then
  MsgBox "La Tasa de Hoy no ha sido grabada"
  Cancel = 100
  frmTasa2.Show 1
End If


DoEvents



CargarInterfaz

CreateRibbonBar

RibbonBar.EnableFrameTheme


          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE ((Accesos.CodUsuario)= " & CodigoUsuario & ") "
          Me.DtaNacceso.Refresh
         Do While Not Me.DtaNacceso.Recordset.EOF

                         Select Case Me.DtaNacceso.Recordset("AccesoModulo")
                           
                           Case "Presupuesto"
                        
                           Case "Editar Niveles"

                                 
                           Case "Cuentas"
                            
                           Case "Grupo Cuentas"
                             
                             
                           Case "Contratistas"
                             
                               
                           Case "Empleados"
                             
                           Case "Activo Fijo"
                           
                                
                           Case "Periodos"
                                 
                                      
                           Case "Transacciones"
                           
                           Case "Cheques"
                           
                               
                           Case "Calcular Depreciacion"

                            
                           Case "Usuarios"

                           
                           Case "Tasa de Cambio"
                           
                                 
                           Case "Reportes Movimientos"
                                 
                                
                           Case "Reportes Bancos"
                            
                            
                           End Select




           Me.DtaNacceso.Recordset.MoveNext
         Loop
         
         

Me.Caption = "Licencia para:  " & NombreEmpresa & "   RUC:  " & RUC


AdoConfiguracion.ConnectionString = Conexion
AdoConfiguracion.RecordSource = "SELECT * FROM DatosEmpresa"
AdoConfiguracion.Refresh


                  If MDIPrimero.AdoConfiguracion.Recordset.RecordCount > 0 Then
                     If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion) Then
                        ConexionFacturacion = MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion
          
                     Else
                        ConexionFacturacion = ""
                     End If
                  End If

'FrmListaUsuario.CmdSalir.Value = True

'Guardando configuración
'    MDIPrimero.AdoConfiguracion.Refresh
'
'    MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa = "3F DINAMARCA"
'    MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa1 = "3F DINAMARCA"
'    MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa2 = "3F DINAMARCA"
'    MDIPrimero.AdoConfiguracion.Recordset!numerorUC = "-"
'
'
'MDIPrimero.AdoConfiguracion.Recordset.Update


'If CDate(Now) > CDate("06/01/2015") Then
'  rs.Open "DELETE FROM Periodos", Conexion
'End If

    MDIPrimero.AdoConsulta.ConnectionString = Conexion
    MDIPrimero.AdoConsulta.RecordSource = "DatosEmpresa"
    MDIPrimero.AdoConsulta.Refresh
    
    If Not IsNull(MDIPrimero.AdoConsulta.Recordset("Valor")) Then
     If Not MDIPrimero.AdoConsulta.Recordset("Valor") = "" Then
      Cadena2 = MDIPrimero.AdoConsulta.Recordset("Valor")
     End If
    End If

Cadena2 = Decrypt(Cadena2)

'Cadena = Encrypt("Siempre")

'¯°±ž  ABC0
'Á×ÓÛÞàÓ  Siempre

'If Cadena2 <> "Siempre" Then
' If CDbl(Mid(Cadena2, 4, 9)) > 2400 Then
'  MsgBox "Licencia Demo ha Caducado!!!!", vbCritical, " Demos"
'  Unload Me
' End If
'End If

FechaIngreso = Now

Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 2 Then
'    PopupMenu MnuMenu, 2, X
'End If
End Sub



Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Demo (Now)

    KillProcess ("ZeusContabilidad.exe")
    End


End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
'        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
'        SaveSetting App.Title, "Settings", "MainTop", Me.top
'        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
'        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
     
End Sub

Public Sub CargarInterfaz()
 
    CommandBarsGlobalSettings.App = App
'    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
'    Me.top = GetSetting(App.Title, "Settings", "MainTop", 1000)
'    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
'    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
'
      
    Dim Workspace  As TabWorkspace
    Set Workspace = CommandBars.ShowTabWorkspace(True)
    Workspace.ThemedBackColor = False
    Workspace.PaintManager.ShowIcons = False
    
    Dim Pane1 As Pane
    Set Pane1 = DockingPaneManager.CreatePane(1, 154, 120, DockLeftOf, Nothing)
    Pane1.Title = "Navegador"
    Pane1.Options = PaneNoCloseable
    Pane1.Select
    
  
    CommandBars.Options.KeyboardCuesShow = xtpKeyboardCuesShowWindowsDefault
        
    CommandBars.EnableCustomization True

    DockingPaneManager.SetCommandBars CommandBars
    DockingPaneManager.ImageList = Me.imlPaneIcons
End Sub



Private Sub MnuSalir_Click()
End
End Sub

Private Sub CreateRibbonBar()

    Dim TabView As RibbonTab
    Dim TabHome As RibbonTab
    Dim TabCatalogo As RibbonTab
    Dim TabEdit As RibbonTab
    Dim TabPrintPreview As RibbonTab
    Dim GroupFile As RibbonGroup
    Dim GroupClipboard As RibbonGroup
    Dim GroupEditing As RibbonGroup
    Dim GroupShowHide As RibbonGroup
    Dim GroupDocumentViews As RibbonGroup
    Dim GroupWindow As RibbonGroup
    Dim GroupPrint As RibbonGroup
    Dim GroupPageSetup As RibbonGroup
    Dim GroupZoom As RibbonGroup
    Dim GroupPreview As RibbonGroup
    Dim ControlCuentas As CommandBarButton
    Dim ControlPrint As CommandBarPopup
    Dim Control As CommandBarControl
    Dim ControlPaste As CommandBarPopup
    Dim ControlSelect As CommandBarPopup
    Dim ControlPopup As CommandBarPopup
    Dim ControlMargins As CommandBarPopup
    Dim ControlOrientation As CommandBarPopup
    Dim ControlSize As CommandBarPopup
    Dim ControlFile As CommandBarPopup
    Dim ControlAbout As CommandBarControl
    Dim Item As CommandBarControl

    Dim RibbonBar As RibbonBar
    CommandBars.Options.UseSharedImageList = False
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Icono.png", 1200, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Salir.png", 1300, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Cuentas.png", 1700, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\CtasMayor.png", 1701, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\AuxiliarCuentas.png", 1702, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Periodo.png", 1703, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Contratista.png", 1704, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Empleados.png", 1705, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Usuarios.png", 1706, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ActivoFijo.png", 1707, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\TasaCambio.png", 1708, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Auditor.png", 1709, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Calculadora.png", 1710, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Cheques.png", 1711, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Presupuesto.png", 1712, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Respaldar2.png", 1713, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Compañia.png", 1714, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Configuracion.png", 1715, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Prorrateo.png", 1716, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Transacciones.png", 1717, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ReporteGeneral.png", 1718, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ReporteMovimientos.png", 1719, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ReportesBancos.png", 1720, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\EstadosFinancieros.png", 1721, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\AnalisisFinanciero2.png", 1722, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ConexionFacturacion.png", 1723, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ContabilizarFacturacion.png", 1724, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ConexionContabilidad.png", 1725, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ContabilizarNomina.png", 1726, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\TipoActivo.png", 1727, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\DirectorioActivoFijo.png", 1728, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\CodigoActivo.png", 1729, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\TrasladoActivo.png", 1730, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\BajaActivos.png", 1731, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\CalcularDepreciacion.png", 1732, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\SoporteTecnico.png", 1733, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Programar.png", 1734, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Update.png", 1735, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Money2.png", 1736, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Caja.png", 1737, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\SolicitudPago.png", 1738, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Listado.png", 1739, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Cheques.png", 1740, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\SolicitudPago.png", 1741, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\EstructuraPresupuesto.png", 1742, xtpImageNormal
    
    
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////////CREO EL RIBBON Y LE CARGO LA IMAGEN//////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Set RibbonBar = CommandBars.AddRibbonBar("The Ribbon")
    RibbonBar.EnableDocking xtpFlagStretched
    
    Set ControlFile = RibbonBar.AddSystemButton()
    ControlFile.IconId = 1200
           Set Control = ControlFile.CommandBar.Controls.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1300, "S&alir", False, False)
    Control.BeginGroup = True
    ControlFile.CommandBar.SetIconSize 35, 35

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////////CREO LOS TABS//////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(0, "&Accesos")
    TabHome.Id = 130
        Set GroupFile = TabHome.Groups.AddGroup("Cuentas", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1700, "&Cuentas", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1701, "&Cuentas de Mayor", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1702, "&Auxiliar de Cuentas", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow

        Set GroupFile = TabHome.Groups.AddGroup("Catalogo", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1704, "&Contratistas", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1703, "&Periodo", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1708, "&Tasa Cambio", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1706, "&Usuarios", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     
     Set GroupFile = TabHome.Groups.AddGroup("Procesos", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1717, "&Transacciones", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1711, "&Registro de Cheques", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1712, "&Presupuestos", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1716, "&Prorrateo Cuentas", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1737, "&Egresos Efectivo", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1738, "&Solicitud de Pago", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     
    
     Set GroupFile = TabHome.Groups.AddGroup("Opciones", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1710, "&Calculadora", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1715, "&Configuracion", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1713, "&Respaldar", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1714, "&Compañia", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     
     Set GroupFile = TabHome.Groups.AddGroup("Ayuda", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1709, "&Auditor", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1735, "&Actualizar", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow

    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE ACTIVO FIJO//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(1, "&Activo Fijo")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Catalogos", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1707, "&Registro de Activo Fijo", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1705, "&Empleados", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1728, "&Oficinas", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1729, "&Codigos Activo Fijo", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set GroupFile = TabHome.Groups.AddGroup("Procesos", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1727, "&Alta de Activos", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1731, "&Bajas de Activos", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1730, "&Traslado de Activo Fijo", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1732, "&Calcular Depreciacion", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1733, "&Responsable Areas", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1734, "&Programar Mantenimientos", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     
     
         '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE BANCOS//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(2, "&Finanzas")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Chequera", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1739, "&Listado Cheques", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1740, "&Registro de Cheques", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1741, "&Solicitud de Pagos", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
      Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1716, "&Prorrateo Cuentas", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1737, "&Egresos Efectivo", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set GroupFile = TabHome.Groups.AddGroup("Presupuesto", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1742, "&Estructura Presupuesto", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1712, "&Presupuestos", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow

    
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE REPORTES//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(3, "&Reportes")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Basicos", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1718, "Reportes &Generales", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set GroupFile = TabHome.Groups.AddGroup("Reportes de Transacciones", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1736, "Reportes &CxC y CxP", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1719, "Reportes &Movimientos", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1720, "Reportes &Bancos", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set GroupFile = TabHome.Groups.AddGroup("Reportes Financieros", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1721, "&Estados Financieros", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1722, "&Analisis Financieros", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE CONTABILIZAR//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(4, "&Contabilizar")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Sistema Facturacion", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1723, "Conexion Facturacion", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1724, "Contabilizar Facturacion", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set GroupFile = TabHome.Groups.AddGroup("Sistema Nominas", 1)
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1725, "Conexion Nominas", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
     Set Item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1726, "Contabilizar Nomina", False, False)
     Item.Style = xtpButtonIconAndCaptionBelow
    
    RibbonBar.QuickAccessControls.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_FILE_SAVE, "Zeus Contable 6.26", False, False


End Sub


Function LoadIcon(Path As String, cx As Long, cy As Long) As Long
    LoadIcon = LoadImage(App.hInstance, App.Path + "\" + Path, 1, cx, cy, 16)
End Function


