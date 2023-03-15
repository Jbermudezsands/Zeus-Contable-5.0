VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmCreaNodos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                  Nuevo Grupo"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin MSAdodcLib.Adodc DtaGrupos 
      Height          =   375
      Left            =   240
      Top             =   2760
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.TextBox TxtCodigo 
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Siguiente Nivel"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mismo Nivel"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCreaNodos.frx":0000
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmCreaNodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGrabar_Click()
Dim Cadena1 As String, Cadena2 As String
Dim KeyHijo2 As String, Cadena3 As String
Dim LongitudNodo As Integer

If Me.Option1.Value = False Then
   '///////////////////////Esta intruccion crea un nodo al siguiente Nivel//////////////////////

    
     
   '//////Si el grupo no tiene hijo recojo los datos para agregarlos//////
     Me.DtaGrupos.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo, Grupos.Imagen1, Grupos.Imagen2 From Grupos Where (((Grupos.KeyGrupoSuperior) = '" & KeyPrincipal & "')) ORDER BY Grupos.KeyGrupo"
     Me.DtaGrupos.Refresh
   If Not DtaGrupos.Recordset.EOF Then
      Me.DtaGrupos.Recordset.MoveLast
      KeyNodoUltimo = Me.DtaGrupos.Recordset("KeyGrupo")
      KeyPadre = Me.DtaGrupos.Recordset("KeyGrupoSuperior")
     LongitudNodo = Len(KeyNodoUltimo)
     Cadena1 = Val((Mid(KeyNodoUltimo, LongitudNodo - 1, LongitudNodo))) + 1
     Cadena2 = Mid(KeyNodoUltimo, 1, LongitudNodo - 2)
     If Len(Cadena2) = 1 Then
      Cadena2 = Cadena2 & "001"
      
     End If
     DescripcionNodo = Me.TxtCodigo
     If Len(Cadena1) = 1 Then
       KeyHijo = Cadena2 & "00" & Cadena1
     Else
       KeyHijo = Cadena2 & Cadena1
     End If
   Else
     LongitudNodo = Len(KeyPrincipal)
     'Cadena1 = Val((Mid(KeyPrincipal, LongitudNodo - 1, LongitudNodo)))
     'Cadena2 = Mid(KeyNodoUltimo, 1, LongitudNodo - 2)
     If LongitudNodo = 1 Then
      Cadena2 = KeyPrincipal & "0100"
      KeyHijo = Cadena2
      KeyPadre = KeyPrincipal
     Else
      Cadena2 = KeyPrincipal & "00"
      KeyHijo = Cadena2
      KeyPadre = KeyPrincipal
     End If
   
   End If


  
      DescripcionNodo = Me.TxtCodigo
         
         
      Set NodX = FrmCuentasMayor.TreeView1.Nodes.Add(KeyPadre, 4, KeyHijo, DescripcionNodo, 2, 1)
      Me.DtaGrupos.Recordset.AddNew
       Me.DtaGrupos.Recordset("KeyGrupo") = KeyHijo
       Me.DtaGrupos.Recordset("KeyGrupoSuperior") = KeyPadre
       Me.DtaGrupos.Recordset("Child") = 4
       Me.DtaGrupos.Recordset("DescripcionGrupo") = DescripcionNodo
       Me.DtaGrupos.Recordset("Imagen1") = 2
       Me.DtaGrupos.Recordset("Imagen2") = 1
       
      Me.DtaGrupos.Recordset.Update
   
 Else
    '///////////////////////Esta intruccion crea un nodo al mismo Nivel//////////////////////
    '//////BUSCO LA LLAVE DEL PADRE////////////////////////////////
    '/////Para el Nivel Seleccionado///////////////////////////////////
    Me.DtaGrupos.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupo) = '" & KeyPrincipal & "'))"
    Me.DtaGrupos.Refresh
    If Not DtaGrupos.Recordset.EOF Then
     KeyPadre = Me.DtaGrupos.Recordset("KeyGrupoSuperior")
    End If
'/////////////////Busco todos los Hijos del Padre///////////////////////////////////////////////////
    Me.DtaGrupos.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo, Grupos.Imagen1, Grupos.Imagen2 From Grupos Where (((Grupos.KeyGrupoSuperior) = '" & KeyPadre & "')) ORDER BY Grupos.KeyGrupo"
    Me.DtaGrupos.Refresh
    If Not DtaGrupos.Recordset.EOF Then
     Me.DtaGrupos.Recordset.MoveLast

     KeyNodoUltimo = Me.DtaGrupos.Recordset("KeyGrupo")
     LongitudNodo = Len(KeyNodoUltimo)
     Cadena1 = Val(Mid(KeyNodoUltimo, LongitudNodo - 1, LongitudNodo)) + 1
     Cadena2 = Mid(KeyNodoUltimo, 1, LongitudNodo - 2)
     Cadena3 = Mid(KeyNodoUltimo, 1, LongitudNodo - 2)
     DescripcionNodo = Me.TxtCodigo
     
     If Len(Cadena2) = 1 Then
      Cadena2 = Cadena2 & "001"
      
     End If
     DescripcionNodo = Me.TxtCodigo
     If Len(Cadena1) = 1 Then
       KeyHijo = Cadena2 & "00" & Cadena1
     Else
       KeyHijo = Cadena2 & Cadena1
     End If
            
     'If Len(Cadena1) = 1 Then
      ' KeyHijo = Cadena3 & "0" & Cadena1
     'Else
      'KeyHijo = Cadena3 & Cadena1
     'End If
    End If
      DescripcionNodo = Me.TxtCodigo
         
      Set NodX = FrmCuentasMayor.TreeView1.Nodes.Add(KeyPadre, 4, KeyHijo, DescripcionNodo, 2, 1)
      Me.DtaGrupos.Recordset.AddNew
       Me.DtaGrupos.Recordset("KeyGrupo") = KeyHijo
       Me.DtaGrupos.Recordset("KeyGrupoSuperior") = KeyPadre
       Me.DtaGrupos.Recordset("Child") = 4
       Me.DtaGrupos.Recordset("DescripcionGrupo") = DescripcionNodo
       Me.DtaGrupos.Recordset("Imagen1") = 2
       Me.DtaGrupos.Recordset("Imagen2") = 1
       
      Me.DtaGrupos.Recordset.Update

 
 
 End If

Unload Me
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
With Me.DtaGrupos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With
End Sub

Private Sub Form_Load()
    MDIPrimero.Skin1.ApplySkin hWnd
End Sub
