VERSION 5.00
Begin VB.Form FrmInstala 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "INSTALACION ZEUS NOMINAS"
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11580
   Icon            =   "FrmInstala.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdManualUsuario 
      BackColor       =   &H00FFFFFF&
      Caption         =   "VER MANUAL DE USUARIO"
      Height          =   1455
      Left            =   9240
      Picture         =   "FrmInstala.frx":74F2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Cmdframework 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INSTALAR FRAMEWORK"
      Height          =   1455
      Left            =   9240
      Picture         =   "FrmInstala.frx":7AD4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&CANCELAR INSTALACION"
      Height          =   1575
      Left            =   9240
      Picture         =   "FrmInstala.frx":CA4E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton CmdInstalarZeus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INSTALACION  ZEUS"
      Height          =   1455
      Left            =   9240
      Picture         =   "FrmInstala.frx":FC54
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton CmdSQL2005 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INSTALAR SQL  2005"
      Height          =   1455
      Left            =   9240
      Picture         =   "FrmInstala.frx":12E5A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   7575
      Left            =   0
      Picture         =   "FrmInstala.frx":1E69C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9240
   End
End
Attribute VB_Name = "FrmInstala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Cmdframework_Click()
On Error GoTo TipoErrs

 Dim Directorio As String
   Directorio = App.Path & "\dotnetfx.exe"
   Directorio = Shell(Directorio)
   
   Exit Sub
TipoErrs:
MsgBox Err.Description
End Sub

Private Sub CmdPacioli_Click()

End Sub

Private Sub CmdInstalarZeus_Click()
On Error GoTo TipoErrs

 Dim Directorio As String
   Directorio = App.Path & "\InstaldorCompleto.exe"
   Directorio = Shell(Directorio)
'   Directorio = App.Path & "\Debug\setup.exe"
'   Directorio = Shell(Directorio)
   Exit Sub
TipoErrs:
MsgBox Err.Description
End Sub

Private Sub CmdManualUsuario_Click()
On Error GoTo TipoErrs

Dim res As Long
Dim ArchivoPDF As String
ArchivoPDF = App.Path & "\Manual de Usuarios.pdf"
res = ShellExecute(FrmInstala.hwnd, "open", ArchivoPDF, "", "", 1)

Exit Sub
TipoErrs:
MsgBox Err.Description

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSQL2005_Click()
On Error GoTo TipoErrs

 Dim Directorio As String
   Directorio = App.Path & "\SQLEXPR_ADV_ESN.EXE"
   Directorio = Shell(Directorio)
   
   Exit Sub
TipoErrs:
MsgBox Err.Description
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
'Me.BackColor = RGB(236, 233, 216)
'Me.CmdPacioli.BackColor = RGB(236, 233, 216)
'Me.CmdEnlace.BackColor = RGB(236, 233, 216)
'Me.CmdSalir.BackColor = RGB(236, 233, 216)
End Sub

