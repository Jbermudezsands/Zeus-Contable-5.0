VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.1#0"; "Codejock.Controls.v12.1.1.ocx"
Begin VB.Form FrmResponsablesAreas 
   BackColor       =   &H80000003&
   Caption         =   "Registro y Control de Responsables de Areas o Departamentos"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10845
   Icon            =   "FrmResponsablesAreas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H80000003&
      ForeColor       =   &H00404000&
      Height          =   7215
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
         TabIndex        =   20
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
         TabIndex        =   15
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
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
      Begin MSMask.MaskEdBox msktel 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
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
         Height          =   4335
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   10035
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
            Bindings        =   "FrmResponsablesAreas.frx":058A
            Height          =   3975
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Doble Clic para actualizar registro"
            Top             =   240
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   7011
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
         Format          =   95420417
         CurrentDate     =   38651
      End
      Begin MSDataListLib.DataCombo cmdgrupo2 
         Bindings        =   "FrmResponsablesAreas.frx":05A2
         DataField       =   "Idreg"
         Height          =   360
         Left            =   1680
         TabIndex        =   8
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
         TabIndex        =   9
         ToolTipText     =   "Filtrar Oficinas "
         Top             =   840
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
         TabIndex        =   19
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo Actual  *"
         Height          =   195
         Left            =   5880
         TabIndex        =   14
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correo electronico *"
         Height          =   195
         Left            =   5880
         TabIndex        =   12
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono *"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oficina de trabajo *"
         Height          =   195
         Left            =   240
         TabIndex        =   7
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
      TabIndex        =   16
      Top             =   7440
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
      Left            =   1680
      TabIndex        =   17
      Top             =   7440
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
      Left            =   9240
      TabIndex        =   18
      Top             =   7440
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Campos Obligatorios"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   7920
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
    sql = "select * from ResponsablesAreas where idreg=" & idreg & " "
Else
    sql = "select * from ResponsablesAreas"
End If
adorespon.ConnectionString = ConexionContable
adorespon.RecordSource = sql
adorespon.Refresh

If idreg = 0 Then
    adorespon.Recordset.AddNew
Else
End If
adorespon.Recordset!NombreResponsable = Trim(txtnombrecompleto.Text)
adorespon.Recordset!Area = Trim(cmdgrupo2.Text)
adorespon.Recordset!Telefono = msktel.Text
adorespon.Recordset!Email = txtmail.Text
adorespon.Recordset!cargo = txtcargo.Text
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
txtcargo.Text = adorespon.Recordset!cargo
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
adorespon.ConnectionString = ConexionContable
adorespon.CommandTimeout = 0
adorespon.RecordSource = "select IdReg as No, NombreResponsable as Nombre_Completo, Area as Area_de_Trabajo, Telefono, Email, Cargo, fechareg,IdAreaTrabajo,cedula from ResponsablesAreas"
adorespon.Refresh
End Sub

Private Sub Form_Load()
cargaoficina
msktel.Mask = "#####-###"
End Sub
Private Sub cargaoficina()
With Me.ofic
    .ConnectionString = Conexion
 End With
    CargaADODCConta "Oficinas", ofic, "1", cmdgrupo2.Name, "Trim", ConexionContable, Me, "order by Idreg"
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

