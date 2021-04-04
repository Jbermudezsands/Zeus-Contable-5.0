VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmRespaldar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Respaldar Base de Datos"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   HelpContextID   =   120000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4305
      ScaleWidth      =   6705
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   6735
      Begin VB.PictureBox picTV 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   240
         ScaleHeight     =   855
         ScaleWidth      =   5895
         TabIndex        =   4
         Top             =   240
         Width           =   5895
      End
      Begin XtremeSuiteControls.ProgressBar Barra2 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2760
         Width           =   6135
         _Version        =   786432
         _ExtentX        =   10821
         _ExtentY        =   661
         _StockProps     =   93
         Scrolling       =   2
         Appearance      =   6
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
         Left            =   360
         TabIndex        =   20
         Top             =   1560
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
         Left            =   360
         TabIndex        =   19
         Top             =   1560
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
         Left            =   360
         TabIndex        =   18
         Top             =   1560
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
         Left            =   360
         TabIndex        =   17
         Top             =   1560
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
         Left            =   360
         TabIndex        =   16
         Top             =   1560
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
         Left            =   360
         TabIndex        =   15
         Top             =   1560
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
         Left            =   360
         TabIndex        =   14
         Top             =   1560
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
         Left            =   360
         TabIndex        =   13
         Top             =   1560
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
         Left            =   360
         TabIndex        =   12
         Top             =   1560
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
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   5415
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
         Left            =   360
         TabIndex        =   10
         Top             =   1560
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
         Left            =   360
         TabIndex        =   9
         Top             =   1560
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
         Left            =   360
         TabIndex        =   8
         Top             =   1560
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
         Left            =   360
         TabIndex        =   7
         Top             =   1560
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
         Left            =   360
         TabIndex        =   6
         Top             =   1560
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
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Image Img2 
         Height          =   480
         Left            =   480
         Picture         =   "FrmRes.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   495
      End
      Begin VB.Image img1 
         Height          =   480
         Left            =   1440
         Picture         =   "FrmRes.frx":A487
         Stretch         =   -1  'True
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   5040
   End
   Begin MSAdodcLib.Adodc AdoPassword 
      Height          =   330
      Left            =   2280
      Top             =   600
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
      Caption         =   "AdoPassword"
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6975
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Image Image2 
         Height          =   645
         Left            =   480
         Picture         =   "FrmRes.frx":17E32
         Top             =   120
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   6960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Top             =   120
         Width           =   645
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Respaldar/Resturar"
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
         TabIndex        =   1
         Top             =   240
         Width           =   2520
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informacion"
      Height          =   3735
      Left            =   240
      TabIndex        =   21
      Top             =   1320
      Width           =   6495
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo de Respaldo"
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   6135
         Begin VB.OptionButton OptDiferencial 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Diferencial"
            Height          =   255
            Left            =   2640
            TabIndex        =   29
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton OptCompleto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Completo"
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdcerrar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   26
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdBackup 
         Caption         =   "Respaldar"
         Height          =   375
         Left            =   3720
         TabIndex        =   25
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtruta 
         Height          =   885
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2160
         Width           =   6135
      End
      Begin VB.TextBox txtbd 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "SistemaVentas"
         Top             =   600
         Width           =   6135
      End
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Examinar"
         Height          =   375
         Left            =   5160
         TabIndex        =   22
         Top             =   1800
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4320
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Data File Name:"
         Filter          =   "*.bkp"
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmRes.frx":194A0
         Top             =   3120
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmRes.frx":274CCD
         TabIndex        =   30
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmRes.frx":274D4D
         TabIndex        =   31
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmRespaldar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tape As New clsTape

Dim ConexionBackup As New ADODB.Connection
Dim rstBD As ADODB.Recordset
Dim Contraseña As String



Sub Seleccionar(TextBox As TextBox)
    TextBox.SetFocus
    TextBox.SelStart = 0
    TextBox.SelLength = Len(TextBox)
End Sub

Private Sub cmdBackup_Click()
On Error GoTo error
Dim Longitud As Double
Dim Directorio As String


Me.Picture2.Visible = True
DoEvents


If txtbd.Text = "" Then
    MsgBox "Debe ingresar la base de datos a respaldar", vbExclamation
    Exit Sub
End If



 Directorio = ""
 Longitud = Len(Me.CommonDialog1.FileName)
 Directorio = Mid(Me.CommonDialog1.FileName, 1, Longitud - 4)
 
 
If Me.OptCompleto.Value Then
 Me.txtruta.Text = Directorio & " Full " & Format(Now, "dd-mm-yyyy") & ".bkp"
'    Me.txtruta.Text = App.Path & "\Respaldos\" & Me.txtbd.Text & "Full" & Format(Now, "ddmmyyyy-hh-mm-ss") & ".bkp"
Else
'    Me.txtruta.Text = App.Path & "\Respaldos\" & Me.txtbd.Text & "Dif" & Format(Now, "ddmmmyyyy-hh-mm-ss") & ".bkp"
 Me.txtruta.Text = Directorio & "Dif" & Format(Now, "dd-mm-yyyy") & ".bkp"
End If

If txtruta.Text = "" Then
    MsgBox "Debe indicar la ruta donde guardara el respaldo", vbExclamation
    txtruta.SetFocus
Else
    DoEvents
    If OptCompleto.Value Then
        ConexionBackup.Execute "Backup DATABASE [" + txtbd.Text + "] TO DISK='" & txtruta & "'"
    ElseIf OptDiferencial.Value Then
        ConexionBackup.Execute "Backup DATABASE [" + txtbd.Text + "] TO DISK='" & txtruta & "' with DIFFERENTIAL"
    End If
    Debug.Print Me.txtruta.Text
    MsgBox "Base de datos Respaldada con exito", vbInformation
    Unload Me
End If

error:
If err.Number <> 0 Then
    MsgBox err.Description '"Ha ocurrido un error al momento de intentar realizar el respaldo", vbInformation
End If
'Exit Sub

End Sub

Private Sub cmdBrowse_Click()
 On Error GoTo errHandler:
     Me.CommonDialog1.FileName = Me.txtbd.Text
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "All Files (*.*)|*.*|Backup Files (*.bak)|*.bak"
    CommonDialog1.DefaultExt = "bak"
    CommonDialog1.DialogTitle = "Nombre del Respaldo"
    Me.CommonDialog1.ShowSave
    txtruta.Text = CommonDialog1.FileName
'    CommonDialog1.Action = 0
 

    
    Exit Sub
    
errHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdExaminarRestaurar_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdRestaurar_Click()
'On Error GoTo error
'Dim Longitud As Double
'Dim Directorio As String
'
'Dim NextLine As String, Cadena As Variant
'Dim posicion As Long, Posicion2 As Long
'Dim Servidor As String, Catalogo As String
'Dim ConexionBackupSTR1 As String
'Dim StrCn As String
'Dim BaseDatos As String
'Dim ConexionMaster As String
'
'Me.Picture2.Visible = True
'DoEvents
'
'
'If txtbd.Text = "" Then
'    MsgBox "Debe ingresar la base de datos a respaldar", vbExclamation
'    Exit Sub
'End If
'
'
'
' Directorio = ""
' Longitud = Len(Me.CommonDialog1.FileName)
' Directorio = Me.CommonDialog1.FileName
'
'ConexionBackup.Close
'
'ConexionBackupSTR1 = Conexion
'posicion = InStr(1, ConexionBackupSTR1, "Source")
'Servidor = Mid(ConexionBackupSTR1, posicion + 7, Len(ConexionBackupSTR1) - posicion + 6)
'
'Posicion2 = InStr(1, ConexionBackupSTR1, "Data")
'posicion = InStr(1, ConexionBackupSTR1, "Catalog")
'Catalogo = Mid(ConexionBackupSTR1, posicion + 8, Posicion2 - posicion - 9)
'
'
'StrCn = "Driver=SQL SERVER;UID=ADMINISTRADOR;SERVER=" & Servidor & ";DATABASE=MASTER;TRUSTED_CONNECTION=YES; APP=VIRUS;WID=SISTEMAS"
'StrCn = Conexion 'temporal
'ConexionBackup.ConnectionString = StrCn
'Debug.Print StrCn
'On Error Resume Next
'ConexionBackup.Open
'Set rstBD = New ADODB.Recordset
'
'rstBD.Open "select name from sysdatabases", ConexionBackup, adOpenDynamic, adLockOptimistic
'
'
'If Me.TxtBdRestaura.Text = "" Then
'    MsgBox "Debe indicar la ruta donde guardara el respaldo", vbExclamation
'    txtruta.SetFocus
'Else
'    DoEvents
'    If OptCompleto.Value Then
'        ConexionBackup.Execute "RESTORE DATABASE [" + Me.TxtBdRestaura.Text + "] FROM DISK='" & Me.TxtRutaRestaurar.Text & "' WITH REPLACE"
'    ElseIf OptDiferencial.Value Then
'        ConexionBackup.Execute "RESTORE DATABASE [" + Me.TxtBdRestaura.Text + "] FROM DISK='" & Me.TxtRutaRestaurar.Text & "' WITH REPLACE"
'    End If
'    Debug.Print Me.txtruta.Text
'    MsgBox "Base de datos Restaurada con exito", vbInformation
'    Unload Me
'End If
'
'error:
'If err.Number <> 0 Then
'    MsgBox err.Description '"Ha ocurrido un error al momento de intentar realizar el respaldo", vbInformation
'End If
''Exit Sub

End Sub



Private Sub Form_Load()
Skin1.ApplySkin Me.hWnd
Dim NextLine As String, Cadena As Variant
Dim posicion As Long
Dim Servidor As String
Dim ConexionBackupSTR1 As String
Dim StrCn As String
Dim BaseDatos As String

Me.Picture2.BackColor = RGB(219, 226, 242)
Me.picTV.BackColor = RGB(219, 226, 242)

Me.Timer1.Enabled = True
Me.Timer1.Interval = Tape.Speed

'Open App.Path + "\SysInfo.dll" For Input As #1
'i = 1
' Do Until EOF(1)
'    Line Input #1, NextLine
'    If i = 1 Then
'        ConexionBackupSTR1 = Trim(NextLine)
'    Else
'        ConexionBackupSTR2 = Trim(NextLine)
'    End If
'    i = i + 1
' Loop
'Close #1
ConexionBackupSTR1 = Conexion
posicion = InStr(1, ConexionBackupSTR1, "Source")
Servidor = Mid(ConexionBackupSTR1, posicion + 7, Len(ConexionBackupSTR1) - posicion + 6)

Cadena = UCase(Conexion)

'Debug.Print Conexion
Dim inicioBD As Integer, FinBD As Integer
inicioBD = InStr(UCase(Conexion), UCase("Initial Catalog="))
inicioBD = inicioBD + 16

FinBD = InStr(UCase(Conexion), UCase(";Data Source="))

If FinBD <> 0 Then
  Me.txtbd.Text = Mid(Conexion, inicioBD, FinBD - inicioBD)
'  Me.TxtBdRestaura.Text = Mid(Conexion, inicioBD, FinBD - inicioBD)
Else
 FinBD = Len(UCase(Conexion)) + 1
 Me.txtbd.Text = Mid(Conexion, inicioBD, FinBD - inicioBD)
' Me.TxtBdRestaura.Text = Mid(Conexion, inicioBD, FinBD - inicioBD)
End If

StrCn = "Driver=SQL SERVER;UID=ADMINISTRADOR;SERVER=" & Servidor & ";DATABASE=MASTER;TRUSTED_CONNECTION=YES; APP=VIRUS;WID=SISTEMAS"
StrCn = Conexion 'temporal
ConexionBackup.ConnectionString = StrCn
Debug.Print StrCn
On Error Resume Next
ConexionBackup.Open
Set rstBD = New ADODB.Recordset
rstBD.Open "select name from sysdatabases", ConexionBackup, adOpenDynamic, adLockOptimistic

'Me.AdoPassword.ConnectionString = Conexion
'Me.AdoPassword.RecordSource = "select * from configuracion"
'Me.AdoPassword.Refresh
'Contraseña = AdoPassword.Recordset!PasswordBackup

Me.txtruta.Text = App.Path & "\Respaldos\" & Me.txtbd.Text & Format(Date, "ddmmyyyy") & ".bkp"
Me.txtruta.Locked = True
Me.OptDiferencial.Value = 0
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then txtcon.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ConexionBackup = Nothing
    Set rstBD = Nothing
End Sub

Private Sub Timer1_Timer()
On Error GoTo TipoErrs
Dim intWidth As Integer
Dim intLeft As Integer      'Posición izquierda
Dim objImage As Control     'Control Image
Dim objImage1 As Control

DoEvents
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
    
    DoEvents
    
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
 
 DoEvents

' Borra la presentación
    picTV.Cls
    ' Muestra la nueva imagen en la nueva posición
    picTV.PaintPicture objImage.Picture, intLeft, 100, 800, 800
 Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub
