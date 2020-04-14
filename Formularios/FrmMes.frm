VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMes 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SmartButton1 
      Caption         =   "CANCELAR"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   64094209
      CurrentDate     =   38066
   End
End
Attribute VB_Name = "FrmMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
Me.MonthView1.Value = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
 FrmJustificacion.DBGMovimiento.Columns(4).Text = Me.MonthView1.Value
 Unload Me
End Sub

Private Sub SmartButton1_Click()
Unload Me
End Sub
