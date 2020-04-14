VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmConvertir 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   2280
      OleObjectBlob   =   "FrmConvertir.frx":0000
      TabIndex        =   9
      Top             =   240
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmConvertir.frx":0060
      TabIndex        =   8
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton SmartButton2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton SmartButton1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox CmbA 
      Height          =   315
      ItemData        =   "FrmConvertir.frx":00C2
      Left            =   2760
      List            =   "FrmConvertir.frx":00CF
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox CmbDe 
      Height          =   315
      ItemData        =   "FrmConvertir.frx":00EE
      Left            =   600
      List            =   "FrmConvertir.frx":00FB
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conversion de Moneda"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4095
      Begin VB.TextBox LblTotal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox TxtTasa 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox TxtMonto 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNombre 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmConvertir.frx":011A
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmConvertir.frx":0194
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmConvertir.frx":020E
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmConvertir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbA_Click()
Me.Label3.Caption = "Monto " & Me.CmbA.Text
End Sub

Private Sub CmbDe_Click()
  Me.LblNombre.Caption = "Monto " & Me.CmbDe
End Sub

Private Sub Form_Load()
Me.TxtTasa.BackColor = RGB(207, 207, 207)
MDIPrimero.Skin1.ApplySkin hWnd
End Sub

Private Sub SmartButton1_Click()
 Select Case Indice
  Case 1
     FrmTransacciones.DBGTransacciones.Columns(8).Text = Me.LblTotal
  Case 2
     FrmTransacciones.DBGTransacciones.Columns(9).Text = Me.LblTotal
  
 End Select
End Sub

Private Sub SmartButton2_Click()
Unload Me
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TxtMonto_Change()
Dim Tasa As Double, Monto As Double
If Me.TxtMonto.Text = "" Then
Me.LblTotal.Text = ""

End If
 If IsNumeric(TxtMonto.Text) Then
  If Me.CmbDe.Text = "Córdobas" And Me.CmbA.Text = "Dólares" Then
   Tasa = Me.TxtTasa
   Monto = Me.TxtMonto.Text
   Total = Tasa * Monto
   Me.LblTotal.Text = Format(Total, "##,##0.00")
  End If
  If Me.CmbDe.Text = "Dólares" And Me.CmbA.Text = "Córdobas" Then
   Tasa = Me.TxtTasa
   Monto = Me.TxtMonto.Text
   
   Total = (1 / Tasa) * Monto
   Me.LblTotal.Text = Format(Total, "##,##0.00")
  End If

  If Me.CmbDe.Text = "Libras" And Me.CmbA.Text = "Córdobas" Then
   Tasa = Me.TxtTasa
   Monto = Me.TxtMonto.Text
   
   Total = (1 / Tasa) * Monto
   Me.LblTotal.Text = Format(Total, "##,##0.00")
  End If
  
  If Me.CmbDe.Text = "Córdobas" And Me.CmbA.Text = "Dólares" Then
   Tasa = Me.TxtTasa
   Monto = Me.TxtMonto.Text
   
   Total = Tasa * Monto
   Me.LblTotal.Text = Format(Total, "##,##0.00")
  End If
  
   If Me.CmbDe.Text = "Córdobas" And Me.CmbA.Text = "Libras" Then
   Tasa = Me.TxtTasa
   Monto = Me.TxtMonto.Text
   
   Total = Tasa * Monto
   Me.LblTotal.Text = Format(Total, "##,##0.00")
  End If
  
   If Me.CmbDe.Text = "Dólares" And Me.CmbA.Text = "Libras" Then
   Tasa = Me.TxtTasa
   Monto = Me.TxtMonto.Text
   
   Total = (1 / Tasa) * Monto
   Me.LblTotal.Text = Format(Total, "##,##0.00")
  End If
  
   
   If Me.CmbDe.Text = "Libras" And Me.CmbA.Text = "Dólares" Then
   Tasa = Me.TxtTasa
   Monto = Me.TxtMonto.Text
   
   Total = Tasa * Monto
   Me.LblTotal.Text = Format(Total, "##,##0.00")
  End If
  

 Else
   'MsgBox "Debe digitar solo numeros", vbCritical, "Sistema Contable"
   'Me.LblTotal.Caption = ""
 End If
End Sub

Private Sub TxtTasa_Change()
Dim Tasa As Double, Monto As Double

 If IsNumeric(TxtMonto.Text) Then
   Tasa = Me.TxtTasa
   Monto = Me.TxtMonto.Text
   Total = Tasa * Monto
   Me.LblTotal.Text = Format(Total, "##,##0.00")
 Else
  ' MsgBox "Debe digitar solo numeros", vbCritical, "Sistema Contable"
  ' Me.LblTotal.Caption = ""
 End If
End Sub
