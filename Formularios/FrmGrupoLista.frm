VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmGrupoLista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupo de Cuentas"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   Icon            =   "FrmGrupoLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton SmartButton1 
      Caption         =   "Pegar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc DtaCuentas 
      Height          =   375
      Left            =   1680
      Top             =   7920
      Width           =   2895
      _ExtentX        =   5106
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
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   1560
      Top             =   7800
      Width           =   2895
      _ExtentX        =   5106
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
      Left            =   1680
      Top             =   7800
      Width           =   3375
      _ExtentX        =   5953
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
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9128
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGrupoLista.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGrupoLista.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGrupoLista.frx":0BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGrupoLista.frx":1000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmGrupoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
Dim NodX As Node
Dim Relatives As String, RelationsShips As String
Dim LLave As String, Texto As String, Imagen1 As Integer
Dim Imagen2 As Integer

With Me.DtaGrupos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Grupos"
   .Refresh
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

 Me.DtaGrupos.Refresh
 Do While Not Me.DtaGrupos.Recordset.EOF
   If Not IsNull(Me.DtaGrupos.Recordset("KeyGrupoSuperior")) Then
    Relatives = Me.DtaGrupos.Recordset("KeyGrupoSuperior")
   Else
     Relatives = ""
   End If
   If Not IsNull(Me.DtaGrupos.Recordset("Child")) Then
     RelationsShips = Me.DtaGrupos.Recordset("Child")
   Else
     RelationsShips = ""
   End If
   LLave = Me.DtaGrupos.Recordset("KeyGrupo")
   Texto = Me.DtaGrupos.Recordset("DescripcionGrupo")
   Imagen1 = Me.DtaGrupos.Recordset("Imagen1")
   Imagen2 = Me.DtaGrupos.Recordset("Imagen2")
   If Relatives = "" And RelationsShips = "" Then
   Set NodX = Me.TreeView1.Nodes.Add(, , LLave, Texto, Imagen1, Imagen2)
   Else
   Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, LLave, Texto, Imagen1, Imagen2)
   End If
   
  Me.DtaGrupos.Recordset.MoveNext
 Loop

'Set NodX = Me.TreeView1.Nodes.Add(, , "A", "1.Activos", 4, 3)
'Set NodX = Me.TreeView1.Nodes.Add(, , "B", "2.Pasivo", 4, 3)
'Set NodX = Me.TreeView1.Nodes.Add(, , "C", "3.Capital", 4, 3)
'Set NodX = Me.TreeView1.Nodes.Add(, , "D", "4.Ingresos", 4, 3)
'Set NodX = Me.TreeView1.Nodes.Add(, , "O", "5.Costos", 4, 3)
'Set NodX = Me.TreeView1.Nodes.Add(, , "G", "6.Gastos", 4, 3)
'Set NodX = Me.TreeView1.Nodes.Add("A", 4, "A0100", "Activo Circulante", 2, 1)
'Set NodX = Me.TreeView1.Nodes.Add("B", 4, "P0100", "Pasivo Circulante", 2, 1)
'Set NodX = Me.TreeView1.Nodes.Add("B", 4, "P0200", "Pasivo Fijo", 2, 1)
KeyPrincipal = "A"
Me.TreeView1.Nodes(Me.TreeView1.Nodes.Count).EnsureVisible
NodoBase = True
Me.DtaCuentas.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.KeyGrupo) = '" & KeyPrincipal & "')) ORDER BY Cuentas.CodCuentas"
Me.DtaCuentas.Refresh


End Sub

Private Sub SmartButton1_Click()
 Dim TipoCuenta As String, KeyTipo As String
 
 If QUIEN = "MoverGrupos" Then
  FrmMoverGrupos.TxtDescripcionGrupo.Text = Me.TreeView1.SelectedItem
   FrmMoverGrupos.TxtKeyGrupo = KeyPrincipal
  Unload Me
  Exit Sub
 ElseIf QUIEN = "CuentasReportes" Then
   FrmReportes.TxtDesde.Text = Me.TreeView1.SelectedItem
   FrmReportes.TxtKeyGrupoDesde.Text = KeyPrincipal
  Unload Me
  Exit Sub
  
 ElseIf QUIEN = "CuentasReportes2" Then
   FrmReportes.TxtHasta.Text = Me.TreeView1.SelectedItem
   FrmReportes.TxtKeyGrupoHasta.Text = KeyPrincipal
  Unload Me
  Exit Sub
 
End If

 
 
 KeyTipo = Mid(KeyPrincipal, 1, 1)
 KeyGrupoCuenta = KeyPrincipal
 TipoCuenta = FrmCuentas.CmbTipo
 If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
  TipoCuenta = "A"
 ElseIf TipoCuenta = "Otros Pasivos" Or TipoCuenta = "Cuentas x Pagar" Or TipoCuenta = "Pasivo" Then
  TipoCuenta = "B"
 ElseIf TipoCuenta = "Capital" Then
  TipoCuenta = "C"
 ElseIf TipoCuenta = "Costos" Then
  TipoCuenta = "G"
 ElseIf TipoCuenta = "Gastos" Then
  TipoCuenta = "O"
 ElseIf TipoCuenta = "Ingresos - Ventas" Then
  TipoCuenta = "D"
 ElseIf TipoCuenta = "Cuentas de Orden" Then
  TipoCuenta = "P"
 End If
   
 
 If KeyTipo = TipoCuenta Then
  FrmCuentas.TxtDescripcionGrupo.Text = Me.TreeView1.SelectedItem
 Else
  MsgBox "Ha Seleccionado,un Grupo Distinto a la Naturaleza de la cuenta", vbCritical, "Sistema Contable"
  Exit Sub
 End If

 Unload Me
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim numero As Integer
  Dim Cadena1 As String, Cadena2 As String
  KeyPadre = ""
  KeyHijo = ""
  KeyNodoUltimo = ""
  KeyPrincipal = Node.Key
 Me.DtaCuentas.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.KeyGrupo) = '" & KeyPrincipal & "')) ORDER BY Cuentas.CodCuentas"
 Me.DtaCuentas.Refresh


 

If Len(KeyPrincipal) = 1 Then
    NodoBase = True
Else
    NodoBase = False
    KeyPadre = Node.Parent.Key
End If
End Sub
