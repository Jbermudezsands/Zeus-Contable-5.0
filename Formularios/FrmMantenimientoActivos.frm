VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmMantenimientoActivos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Mantenimientos de Activos"
   ClientHeight    =   9435
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   8175
      Left            =   7560
      TabIndex        =   6
      Top             =   0
      Width           =   12615
      Begin TabDlg.SSTab SSTab1 
         Height          =   7815
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   13785
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Informacion General"
         TabPicture(0)   =   "FrmMantenimientoActivos.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label40"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label8"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label9"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label10"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label11"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label12"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label19"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Frame5"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Frame6"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Frame7"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Text9"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Text24"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Text10"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Text16"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Text17"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Text18"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Text27"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Check1"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Frame9"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Check4"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Check3"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Check5"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).ControlCount=   22
         TabCaption(1)   =   "Programacion de Mantenimientos"
         TabPicture(1)   =   "FrmMantenimientoActivos.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame11"
         Tab(1).Control(1)=   "Command4"
         Tab(1).Control(2)=   "Command3"
         Tab(1).Control(3)=   "Command2"
         Tab(1).Control(4)=   "Frame10"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Orden de Trabajo"
         TabPicture(2)   =   "FrmMantenimientoActivos.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Command9"
         Tab(2).Control(1)=   "Command8"
         Tab(2).Control(2)=   "Command7"
         Tab(2).Control(3)=   "Command6"
         Tab(2).Control(4)=   "Command5"
         Tab(2).Control(5)=   "Command1"
         Tab(2).Control(6)=   "Frame13"
         Tab(2).Control(7)=   "Frame12"
         Tab(2).ControlCount=   8
         TabCaption(3)   =   "Piezas"
         TabPicture(3)   =   "FrmMantenimientoActivos.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         Begin VB.CommandButton Command9 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -63360
            Picture         =   "FrmMantenimientoActivos.frx":0070
            Style           =   1  'Graphical
            TabIndex        =   89
            ToolTipText     =   "Agregar nueva programacion"
            Top             =   5520
            Width           =   585
         End
         Begin VB.CommandButton Command8 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -63360
            Picture         =   "FrmMantenimientoActivos.frx":05FA
            Style           =   1  'Graphical
            TabIndex        =   88
            ToolTipText     =   "Eliminar programacion "
            Top             =   6960
            Width           =   585
         End
         Begin VB.CommandButton Command7 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -63360
            Picture         =   "FrmMantenimientoActivos.frx":0B84
            Style           =   1  'Graphical
            TabIndex        =   87
            ToolTipText     =   "Editar programacion"
            Top             =   6240
            Width           =   585
         End
         Begin VB.CommandButton Command6 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -63360
            Picture         =   "FrmMantenimientoActivos.frx":110E
            Style           =   1  'Graphical
            TabIndex        =   86
            ToolTipText     =   "Agregar nueva programacion"
            Top             =   2280
            Width           =   585
         End
         Begin VB.CommandButton Command5 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -63360
            Picture         =   "FrmMantenimientoActivos.frx":1698
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Eliminar programacion "
            Top             =   3720
            Width           =   585
         End
         Begin VB.CommandButton Command1 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -63360
            Picture         =   "FrmMantenimientoActivos.frx":1C22
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Editar programacion"
            Top             =   3000
            Width           =   585
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00B3BFAC&
            Caption         =   "Gastos (Piezas, materiales, mano de obra, etc.)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   -74880
            TabIndex        =   82
            Top             =   4680
            Width           =   11445
            Begin MSAdodcLib.Adodc Adodc2 
               Height          =   330
               Left            =   0
               Top             =   2880
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
            Begin MSDataGridLib.DataGrid DataGrid3 
               Bindings        =   "FrmMantenimientoActivos.frx":21AC
               Height          =   2655
               Left            =   60
               TabIndex        =   83
               Top             =   240
               Width           =   11295
               _ExtentX        =   19923
               _ExtentY        =   4683
               _Version        =   393216
               AllowUpdate     =   -1  'True
               HeadLines       =   2
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
                  Size            =   2
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00B3BFAC&
            Caption         =   "Ordenes de trabajo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3975
            Left            =   -74880
            TabIndex        =   80
            Top             =   480
            Width           =   11445
            Begin MSAdodcLib.Adodc Adodc1 
               Height          =   330
               Left            =   0
               Top             =   3600
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
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "FrmMantenimientoActivos.frx":21C1
               Height          =   3615
               Left            =   60
               TabIndex        =   81
               Top             =   240
               Width           =   11295
               _ExtentX        =   19923
               _ExtentY        =   6376
               _Version        =   393216
               AllowUpdate     =   -1  'True
               HeadLines       =   2
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
                  Size            =   2
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H80000016&
            Height          =   855
            Left            =   -74760
            TabIndex        =   77
            Top             =   6840
            Width           =   7935
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
               TabIndex        =   78
               ToolTipText     =   "Filtre por fecha de Proximo servicio, o por Nombre del Servicio"
               Top             =   150
               Width           =   4545
            End
            Begin VB.Label Label20 
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
               TabIndex        =   79
               Top             =   240
               Width           =   1080
            End
         End
         Begin VB.CommandButton Command4 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -65280
            Picture         =   "FrmMantenimientoActivos.frx":21D6
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Ver Listado de Programaciones"
            Top             =   6960
            Width           =   1185
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -63960
            Picture         =   "FrmMantenimientoActivos.frx":2760
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Eliminar programacion "
            Top             =   6960
            Width           =   1185
         End
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -66600
            Picture         =   "FrmMantenimientoActivos.frx":2CEA
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Agregar nueva programacion"
            Top             =   6960
            Width           =   1185
         End
         Begin VB.Frame Frame10 
            Height          =   6255
            Left            =   -74880
            TabIndex        =   71
            Top             =   480
            Width           =   12135
            Begin VB.Frame Rechazo 
               BackColor       =   &H00B3BFAC&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   6135
               Left            =   0
               TabIndex        =   72
               Top             =   0
               Width           =   12165
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
               Begin MSDataGridLib.DataGrid DataGrid7 
                  Bindings        =   "FrmMantenimientoActivos.frx":3274
                  Height          =   5775
                  Left            =   60
                  TabIndex        =   73
                  Top             =   240
                  Width           =   12015
                  _ExtentX        =   21193
                  _ExtentY        =   10186
                  _Version        =   393216
                  AllowUpdate     =   -1  'True
                  HeadLines       =   2
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
                     Size            =   2
                     BeginProperty Column00 
                     EndProperty
                     BeginProperty Column01 
                     EndProperty
                  EndProperty
               End
            End
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Activo Inactivo"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5760
            TabIndex        =   70
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Activo Trasladado"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5760
            TabIndex        =   69
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Activo dado de Alta"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5760
            TabIndex        =   68
            Top             =   840
            Width           =   1935
         End
         Begin VB.Frame Frame9 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   7800
            TabIndex        =   67
            Top             =   480
            Width           =   4335
            Begin VB.Image Image1 
               BorderStyle     =   1  'Fixed Single
               Height          =   2655
               Left            =   120
               Picture         =   "FrmMantenimientoActivos.frx":3289
               Stretch         =   -1  'True
               Top             =   240
               Width           =   3975
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Es Vehiculo?"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   3960
            Width           =   1575
         End
         Begin VB.TextBox Text27 
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
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   58
            Top             =   3240
            Width           =   1815
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   57
            Top             =   3600
            Width           =   1815
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
            Height          =   330
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   56
            Top             =   2880
            Width           =   1815
         End
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
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   55
            Top             =   3240
            Width           =   1815
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
            Height          =   330
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   54
            Top             =   3600
            Width           =   2895
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
            Height          =   330
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   53
            Top             =   2520
            Width           =   3375
         End
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   52
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Frame Frame7 
            Caption         =   "Info. Principal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   240
            TabIndex        =   35
            Top             =   480
            Width           =   5295
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
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   42
               Top             =   240
               Width           =   1215
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
               Left            =   3720
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   41
               Top             =   960
               Width           =   1455
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
               Left            =   3720
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   40
               Top             =   1320
               Width           =   1455
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
               Left            =   3240
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   39
               Text            =   "0.00"
               Top             =   600
               Width           =   735
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
               Left            =   4440
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   38
               Text            =   "0.00"
               Top             =   600
               Width           =   735
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
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   37
               Top             =   600
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
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   36
               Top             =   960
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker DTPicker8 
               Height          =   300
               Left            =   3600
               TabIndex        =   43
               Top             =   240
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   529
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
               Format          =   76021761
               CurrentDate     =   38651
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cuenta Contable:"
               Height          =   195
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   1230
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Referencia:"
               Height          =   195
               Left            =   2760
               TabIndex        =   50
               Top             =   960
               Width           =   825
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Factura:"
               Height          =   195
               Left            =   2760
               TabIndex        =   49
               Top             =   1320
               Width           =   585
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Costo:"
               Height          =   195
               Left            =   2760
               TabIndex        =   48
               Top             =   600
               Width           =   450
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "F. Compra:"
               Height          =   195
               Left            =   2760
               TabIndex        =   47
               Top             =   240
               Width           =   765
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "IVA:"
               Height          =   195
               Left            =   4080
               TabIndex        =   46
               Top             =   600
               Width           =   540
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cuenta Gastos:"
               Height          =   195
               Left            =   120
               TabIndex        =   45
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cuenta Deprec.:"
               Height          =   195
               Left            =   120
               TabIndex        =   44
               Top             =   960
               Width           =   1170
            End
         End
         Begin VB.Frame Frame6 
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
            Left            =   120
            TabIndex        =   22
            Top             =   4320
            Width           =   5415
            Begin VB.OptionButton Option1 
               Caption         =   "Vehiculo Propio"
               Height          =   375
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Vehiculo Alguilado"
               Height          =   375
               Left            =   2040
               TabIndex        =   26
               Top             =   240
               Width           =   1815
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
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   25
               Text            =   "0.00"
               Top             =   960
               Width           =   1695
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
               Locked          =   -1  'True
               MaxLength       =   300
               MultiLine       =   -1  'True
               TabIndex        =   24
               Top             =   2040
               Width           =   3255
            End
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
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   23
               Top             =   1320
               Width           =   1815
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   300
               Left            =   2040
               TabIndex        =   28
               Top             =   600
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   529
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
               Format          =   76021761
               CurrentDate     =   38651
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   300
               Left            =   2040
               TabIndex        =   29
               Top             =   1680
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   529
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
               Format          =   76021761
               CurrentDate     =   38651
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de adquisicion:"
               Height          =   195
               Left            =   120
               TabIndex        =   34
               Top             =   600
               Width           =   1560
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kilometraje en la compra:"
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   960
               Width           =   1770
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Comprado o alguilado a:"
               Height          =   195
               Left            =   120
               TabIndex        =   32
               Top             =   1320
               Width           =   1710
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Garantia caduca el:"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   1680
               Width           =   1395
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nota sobre la compra:"
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   2040
               Width           =   1560
            End
         End
         Begin VB.Frame Frame5 
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
            Left            =   6120
            TabIndex        =   9
            Top             =   4320
            Width           =   5055
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
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   13
               Top             =   240
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
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   12
               Top             =   600
               Width           =   3375
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
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   11
               Top             =   960
               Width           =   3375
            End
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
               Height          =   1290
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   300
               MultiLine       =   -1  'True
               TabIndex        =   10
               Top             =   2040
               Width           =   3375
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   300
               Left            =   1440
               TabIndex        =   14
               Top             =   1320
               Width           =   3300
               _ExtentX        =   5821
               _ExtentY        =   529
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
               Format          =   76021761
               CurrentDate     =   38651
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   300
               Left            =   1440
               TabIndex        =   15
               Top             =   1680
               Width           =   3300
               _ExtentX        =   5821
               _ExtentY        =   529
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
               Format          =   76021761
               CurrentDate     =   38651
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Asegurador:"
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Compañia de seg."
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Top             =   600
               Width           =   1275
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Referencia:"
               Height          =   195
               Left            =   120
               TabIndex        =   19
               Top             =   960
               Width           =   825
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Inicia:"
               Height          =   195
               Left            =   120
               TabIndex        =   18
               Top             =   1320
               Width           =   915
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Expira en:"
               Height          =   195
               Left            =   120
               TabIndex        =   17
               Top             =   1680
               Width           =   705
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nota:"
               Height          =   195
               Left            =   120
               TabIndex        =   16
               Top             =   2040
               Width           =   390
            End
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año:"
            Height          =   195
            Left            =   4080
            TabIndex        =   65
            Top             =   3000
            Width           =   330
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color:]"
            Height          =   195
            Left            =   4080
            TabIndex        =   64
            Top             =   3240
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad #"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   2880
            Width           =   660
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marca:"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   3600
            Width           =   570
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero de serie (VIN):"
            Height          =   195
            Left            =   4080
            TabIndex        =   60
            Top             =   3600
            Width           =   1605
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrip. del Activo Fijo"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   2520
            Width           =   1620
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   15600
      TabIndex        =   2
      Top             =   8160
      Width           =   4575
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   360
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
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   615
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Editar"
         ForeColor       =   0
         Appearance      =   6
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   615
         Left            =   3000
         TabIndex        =   5
         Top             =   360
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
   Begin VB.Frame Frame1 
      Caption         =   " Control de Secciones de campo por Finca"
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
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
         Height          =   8055
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7395
         Begin VB.CheckBox Check2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ver solo Vehiculos"
            Height          =   255
            Left            =   5280
            TabIndex        =   7
            Top             =   0
            Width           =   1935
         End
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   330
            Left            =   60
            Top             =   6600
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
         Begin TrueOleDBGrid80.TDBGrid DataGrid2 
            Bindings        =   "FrmMantenimientoActivos.frx":7340
            Height          =   7695
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   13573
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
            Splits(0).Caption=   "Listado de Activos"
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
            AllowUpdate     =   0   'False
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
      End
      Begin VB.Image foto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2415
         Left            =   0
         Picture         =   "FrmMantenimientoActivos.frx":7351
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2895
      End
   End
End
Attribute VB_Name = "FrmMantenimientoActivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idactivo As Integer
Dim recibe As String
Dim idMantoAF As Integer
Dim idOrdenTrabajo As Integer
Dim afdescri As String
Dim isdobleclic As Integer

Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset, rsConexion As New ADODB.Recordset
Private SQL As String
Private modal As Boolean
Private getVal As Boolean
Private Id As Integer

Private Sub Check2_Click()
If Check2.Value = 1 Then
    cargarcatalogoAF
Else
    cargarcatalogoAF
End If
End Sub

Private Sub Command2_Click()
If idactivo = 0 Then
    FrmMaestraServicios.ServicioalAF = False
    FrmMaestraServicios.activocod = 0
    MsgBox ("Seleccione un Activo primeramente al cual le asignara el mantenimiento"), vbInformation
    Exit Sub
Else
    FrmMaestraServicios.activocod = idactivo
    FrmMaestraServicios.ServicioalAF = True
    FrmMaestraServicios.Show vbModal
End If
End Sub

Private Sub Command4_Click()
If idactivo = 0 Then
    FrmMaestraServicios.ServicioalAF = False
    FrmMaestraServicios.activocod = 0
    MsgBox ("Seleccione un Activo primeramente al cual le asignara el mantenimiento"), vbInformation
    Exit Sub
Else
    FrmMaestraServicios.activocod = idactivo
    FrmMaestraServicios.ServicioalAF = False
    FrmMaestraServicios.Show vbModal
End If

End Sub

Private Sub Command6_Click()
If idactivo = 0 Or afdescri = "" Then
    FrmOrdenTrabajo.isactualiza = 0
    MsgBox ("Seleccione un Activo para agregar la Orden de Trabajo"), vbInformation
Else
    FrmOrdenTrabajo.isactualiza = 0
    FrmOrdenTrabajo.txttipopla.Text = afdescri
    FrmOrdenTrabajo.idAF = idactivo
    FrmOrdenTrabajo.Show vbModal
End If
End Sub

Private Sub DataGrid1_Click()
isdobleclic = 0
queidorden
End Sub
Private Sub queidorden()
If Adodc1.Recordset.RecordCount = 0 Then
    idOrdenTrabajo = 0
Else
    idOrdenTrabajo = Adodc1.Recordset!No
    If isdobleclic = 1 Then
        FrmOrdenTrabajo.isactualiza = 1
        FrmOrdenTrabajo.idAF = idactivo
        FrmOrdenTrabajo.idordentra = idOrdenTrabajo
        FrmOrdenTrabajo.txttipopla.Text = afdescri
        
        FrmOrdenTrabajo.Show vbModal
    End If
    
End If
End Sub

Private Sub DataGrid1_DblClick()
isdobleclic = 1
queidorden
End Sub

Private Sub DataGrid2_Click()
'If rs.RecordCount = 0 Then
'
'Else
'    veridAF
'    cargardatosAF
'    cargamanttos 1
'    cargarordenestrabajo
'End If
End Sub
Public Sub cargarordenestrabajo()
Adodc1.ConnectionString = Conexion
Adodc1.CommandTimeout = 0
 Adodc1.RecordSource = "select idreg as No, Fcreado as Creado_el, frequeireOrden as Requerido_el, Reportadopor as Reportado_por , proveeresponsable as Proveedor_Responsable, Descripcion as Descripcion, case when Estado='P' then 'Pendiente' else case when  Estado='EC' THEN 'En Curso' else case when  Estado='C' then 'Cerrado' end end end  as Estado, Nota from dbo.ControlOrdenTrabajo where IdActivo=" & idactivo & "  "
' If Adodc1.Recordset.EOF = True Then
'
' Else
    Adodc1.Refresh
'End If
End Sub
Public Sub cargamanttos(opcion As Integer)
Adoreg.ConnectionString = Conexion
Adoreg.CommandTimeout = 0
If opcion = 1 Then
    Adoreg.RecordSource = "select idreg as No, Descripcion, repetircada as Frecuencia, tiporepeticion as Ejecutar_Frecuencia_Cada,ultimoservicio as Ultimo_Servicio,proximomanto as Proximo_Servicio from  dbo.MantenimientoPorActivo where IdActivo=" & idactivo & " "
End If

If opcion = 2 Then
    Adoreg.RecordSource = "select idreg as No, Descripcion, repetircada as Frecuencia, tiporepeticion as Ejecutar_Frecuencia_Cada,ultimoservicio as Ultimo_Servicio,proximomanto as Proximo_Servicio from  dbo.MantenimientoPorActivo where IdActivo=" & idactivo & " and  Descripcion LIKE '" & Trim(txtfiltrorapido.Text) & "%' or proximomanto LIKE '" & Trim(txtfiltrorapido.Text) & "%'"
End If

Adoreg.Refresh
End Sub
Public Function datosalta() As String
Set rsa3 = Nothing
SQL = "select NombreEntrega, nombrerecibe from altadebienes where IdActivoAlta=" & idactivo & ""
rsa3.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
datosalta = rsa3!NombreEntrega
recibe = rsa3!NombreRecibe
End Function
Private Sub cargardatosAF()
Set rsa = Nothing
SQL = "select * from dbo.CatalogoActivoFijo where idreg=" & idactivo & ""
rsa.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
Text20.Text = rsa!CNTACONTABLE
Text7.Text = rsa!CuentaGastos
Text23.Text = rsa!CuentaDepreciacion
DTPicker8.Value = Format(rsa!fcompragen, "DD/MM/YYYY")
'If IsNull(rsa!costovh) Then
'    Text11.Text = ""
'Else
'    Text11.Text = rsa!costovh
'End If
'
'If IsNull(rsa!ivavh) Then
'    Text25.Text = ""
'Else
'    Text25.Text = rsa!ivavh
'End If
Text21.Text = rsa!refegeneral
Text22.Text = rsa!Factura
Text24.Text = rsa!DescripcionAF
Text9.Text = rsa!Unidad
Text27.Text = rsa!marca
Text18.Text = rsa!modelo
Text17.Text = rsa!Año
Text16.Text = rsa!Color
Text10.Text = rsa!Serie
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
If rsa!tipovehiculo <> 0 Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If
If rsa!dadobaja = 1 Or rsa!dadobaja = True Then
    Check5.Value = 1
Else
    Check5.Value = 0
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
    Image1.Picture = LoadPicture(ruta)
Else
    ruta = rsa!dirfoto
    If Dir(ruta) <> "" Then
     Image1.Picture = LoadPicture(ruta)
    End If
End If

End Sub
Private Sub veridAF()
If rs.RecordCount = 0 Then
    idactivo = 0
     afdescri = ""
Else
    idactivo = rs!No
     afdescri = rs!Activo_Fijo
End If
End Sub

Private Sub DataGrid2_DblClick()
veridAF
If idactivo = 0 Then
    FrmAgregarActivoFijo.actualiza = 0
    MsgBox ("No hay registro que mostrar"), vbInformation
     afdescri = ""
Else
    FrmAgregarActivoFijo.actualiza = 1
    FrmAgregarActivoFijo.idregAF = idactivo
    FrmAgregarActivoFijo.Show vbModal
    afdescri = rs!Activo_Fijo
End If
End Sub

Private Sub DataGrid2_FilterChange()
On Error GoTo TipoErrs:

    Dim col As TrueOleDBGrid80.Column
    Dim cols As TrueOleDBGrid80.Columns
    
    'On Error GoTo errHandler
    On Error Resume Next
    Set cols = Me.DataGrid2.Columns
    Dim c As Integer
    
    c = DataGrid2.col
    DataGrid2.HoldFields
    SQL = rs.Filter
    rs.Filter = getFilter(col, cols)
    
    
    DataGrid2.col = c
    DataGrid2.EditActive = True


Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub
Private Function getFilter(col As TrueOleDBGrid80.Column, cols As TrueOleDBGrid80.Columns) As String
'Creates the SQL statement in adodc1.recordset.filter
'and only filters text currently. It must be modified to
'filter other data types.
Dim tmp As String
Dim n As Integer
Dim X As Integer


For Each col In cols
    If Trim(col.FilterText) <> "" Then
        n = n + 1
        If n > 1 Then tmp = tmp & " AND "
        Select Case rs.Fields(X).Type
        Case adVarWChar, adVarChar: tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        Case adInteger, adNumeric: tmp = tmp & "[" & col.DataField & "] = " & col.FilterText
        Case adDBTimeStamp: tmp = tmp & "[" & col.DataField & "] = #" & col.FilterText & "#"
        End Select
    End If
    X = X + 1
Next col
getFilter = tmp

End Function

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rs.RecordCount = 0 Then

Else
    veridAF
    cargardatosAF
    cargamanttos 1
    cargarordenestrabajo
End If
End Sub

Private Sub DataGrid7_Click()
datosMantoAF
End Sub
Private Sub datosMantoAF()
If Adoreg.Recordset.RecordCount = 0 Then
    idMantoAF = 0
Else
    idMantoAF = Adoreg.Recordset!No
End If
End Sub

Private Sub DataGrid7_DblClick()
datosMantoAF
If idMantoAF = 0 Then
    FrmServicios.ismantoAF = 0
    MsgBox ("No hay informacion que mostrar"), vbind
    Exit Sub
Else
    FrmServicios.idservicio = idMantoAF
    FrmServicios.ismantoAF = 1
    FrmServicios.Show vbModal
End If
End Sub


Private Sub Form_Activate()
isdobleclic = 0
End Sub

Private Sub Form_Load()


isdobleclic = 0
cargarcatalogoAF

 Me.DataGrid2.DataSource = rs
 Me.DataGrid2.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DataGrid2.OddRowStyle.BackColor = &H80000005
 Me.DataGrid2.AlternatingRowStyle = True
 Me.BackColor = RGB(216, 228, 248)
End Sub
Private Sub cargarcatalogoAF()
  Dim sqlconsulta As String
    Adodc3.ConnectionString = Conexion
    Adodc3.CommandTimeout = 0
    
    If cnx.State = adStateClosed Then
        cnx.ConnectionString = Conexion
        cnx.Open
    End If
    If Check2.Value = 1 Then
    
'        Adodc3.RecordSource = "Select idReg as No, unidad as Unidad, DescripcionAF AS Activo_Fijo,marca,modelo,año,Serie from dbo.CatalogoActivoFijo WHERE (DatoAlta = 1) AND (isvh = 1)"
        sqlconsulta = "Select idReg as No, unidad as Unidad, DescripcionAF AS Activo_Fijo,marca,modelo,año,Serie from dbo.CatalogoActivoFijo WHERE (DatoAlta = 1) AND (isvh = 1)"
    Else
'        Adodc3.RecordSource = "Select idReg as No, unidad as Unidad, DescripcionAF AS Activo_Fijo, Serie from dbo.CatalogoActivoFijo WHERE(DatoAlta = 1)"
        sqlconsulta = "Select idReg as No, unidad as Unidad, DescripcionAF AS Activo_Fijo, Serie from dbo.CatalogoActivoFijo WHERE(DatoAlta = 1)"
    End If
    With rs
         If Not rs.State = adStateClosed Then
          .Close
         End If
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
    End With
    Me.DataGrid2.DataSource = rs
'    Adodc3.Refresh
End Sub

Private Sub Frame14_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub PushButton1_Click()
veridAF
If idactivo = 0 Then
    FrmAgregarActivoFijo.actualiza = 0
    MsgBox ("No hay registro que mostrar"), vbInformation
     afdescri = ""
Else
    FrmAgregarActivoFijo.actualiza = 1
    FrmAgregarActivoFijo.idregAF = idactivo
    FrmAgregarActivoFijo.Show vbModal
    afdescri = rs!Activo_Fijo
End If
End Sub

Private Sub PushButton2_Click()
FrmAgregarActivoFijo.actualiza = 0
FrmAgregarActivoFijo.Show vbModal
End Sub

Private Sub PushButton3_Click()
Unload Me
End Sub

Private Sub txtfiltrorapido_Change()
cargamanttos 2
End Sub
