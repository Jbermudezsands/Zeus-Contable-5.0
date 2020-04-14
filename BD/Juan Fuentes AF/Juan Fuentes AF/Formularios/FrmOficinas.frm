VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.1#0"; "Codejock.Controls.v12.1.1.ocx"
Begin VB.Form FrmOficinas 
   BackColor       =   &H80000003&
   Caption         =   "Registro de Oficinas para Alta y Asignacion de Activos Fijos"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6450
   Icon            =   "FrmOficinas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H80000003&
      ForeColor       =   &H00404000&
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame7 
         BackColor       =   &H80000003&
         Height          =   2775
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   5655
         Begin VB.Frame Frame8 
            BackColor       =   &H80000003&
            Caption         =   "Descripcion de Oficinas"
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
            Height          =   2535
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   5595
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
            Begin MSDataGridLib.DataGrid DataGrid2 
               Bindings        =   "FrmOficinas.frx":058A
               Height          =   2295
               Left            =   120
               TabIndex        =   6
               ToolTipText     =   "Doble Clic para actualizar registro"
               Top             =   240
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   4048
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
            Begin MSAdodcLib.Adodc adoofic 
               Height          =   330
               Left            =   0
               Top             =   2400
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
      End
      Begin VB.TextBox txtespecie 
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
         MaxLength       =   70
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtfecha 
         Height          =   345
         Left            =   4200
         TabIndex        =   7
         Top             =   360
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
         Format          =   108789761
         CurrentDate     =   38651
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Registro:"
         Height          =   195
         Left            =   4200
         TabIndex        =   9
         Top             =   120
         Width           =   1350
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   840
      End
   End
   Begin XtremeSuiteControls.PushButton btnguarda 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
      _Version        =   786433
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
      Left            =   1920
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
      _Version        =   786433
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
      Left            =   4800
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
      _Version        =   786433
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cerrar"
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
End
Attribute VB_Name = "FrmOficinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idreg As Integer
Private Sub btnguarda_Click()
If idreg <> 0 Then
    sql = "select * from oficinas where idreg= " & idreg & " "
Else
    sql = "select * from oficinas"
End If
adoofic.ConnectionString = ConexionContable
adoofic.RecordSource = sql
adoofic.Refresh

If idreg = 0 Then
    adoofic.Recordset.AddNew
Else
End If
adoofic.Recordset!descripcion = txtespecie.Text
adoofic.Recordset!fechareg = Format(Now, "YYYY/MM/DD")
adoofic.Recordset.Update
cargaoficce
End Sub

Private Sub DataGrid2_Click()
ubicadatos
End Sub
Private Sub ubicadatos()
idreg = adoofic.Recordset!no
txtespecie.Text = adoofic.Recordset!descripcion
dtfecha.Value = Format((adoofic.Recordset!Fecha_Registro), "DD/MM/YYYY")
End Sub

Private Sub DataGrid2_DblClick()
ubicadatos
btnguarda.Enabled = True
End Sub

Private Sub Form_Load()
cargaoficce
End Sub
Private Sub cargaoficce()
    adoofic.ConnectionString = ConexionContable
    adoofic.CommandTimeout = 0
    adoofic.RecordSource = "select Idreg as No, Descripcion , FechaReg as Fecha_Registro from oficinas order by idreg"
    adoofic.Refresh
End Sub

Private Sub PushButton1_Click()
idreg = 0
txtespecie.Text = ""
txtespecie.SetFocus
End Sub

Private Sub PushButton2_Click()
Unload Me
End Sub

Private Sub txtespecie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     btnguarda.Enabled = True
     btnguarda.SetFocus
     
End If
End Sub
