VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmEditarNiveles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Niveles"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   6000
      ScaleHeight     =   2055
      ScaleWidth      =   1935
      TabIndex        =   18
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Image Image2 
         Height          =   2055
         Left            =   0
         Picture         =   "FrmEditaNiveles.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2040
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   6000
      ScaleHeight     =   2055
      ScaleWidth      =   1935
      TabIndex        =   17
      Top             =   3600
      Width           =   1935
      Begin VB.Image Image1 
         Height          =   2055
         Left            =   0
         Picture         =   "FrmEditaNiveles.frx":30042
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2040
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8880
      Top             =   3840
   End
   Begin MSDataListLib.DataList DBLNEmpleado 
      Bindings        =   "FrmEditaNiveles.frx":60084
      Height          =   2010
      Left            =   6360
      TabIndex        =   15
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3545
      _Version        =   393216
      ListField       =   "nombreusuario"
   End
   Begin MSAdodcLib.Adodc DtaNacceso2 
      Height          =   330
      Left            =   120
      Top             =   7320
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "DtaNacceso2"
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
      Height          =   330
      Left            =   2760
      Top             =   7320
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSAdodcLib.Adodc DtaPasword2 
      Height          =   330
      Left            =   2760
      Top             =   7800
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "DtaPasword2"
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
   Begin MSAdodcLib.Adodc DtaPasword 
      Height          =   330
      Left            =   2760
      Top             =   8160
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
      Caption         =   "DtaPasword"
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
      Height          =   330
      Left            =   240
      Top             =   8160
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10800
      TabIndex        =   13
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   5760
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "FrmEditaNiveles.frx":6009D
      TabIndex        =   11
      Top             =   2880
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   6360
      OleObjectBlob   =   "FrmEditaNiveles.frx":60125
      TabIndex        =   10
      Top             =   120
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   10920
      OleObjectBlob   =   "FrmEditaNiveles.frx":601BB
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtCodPasword 
      DataSource      =   "DtaPasword"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Text            =   "TxtCodPasword"
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   7920
      TabIndex        =   7
      Top             =   2760
      Width           =   3975
      Begin ACTIVESKINLibCtl.SkinLabel LblNombre 
         Height          =   285
         Left            =   120
         OleObjectBlob   =   "FrmEditaNiveles.frx":60241
         TabIndex        =   14
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.TextBox TxtNivel 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "TxtNivel"
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Permisos"
      Height          =   2175
      Left            =   9720
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      Begin VB.CheckBox ChEliminar 
         Caption         =   "Eliminar Datos"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox ChGrabar 
         Caption         =   "Grabar Datos"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox ChLeer 
         Caption         =   "Leer Datos"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox ChAbrir 
         Caption         =   "Abrir o Ejecutar"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.ListBox ListAcceso 
      Height          =   1620
      ItemData        =   "FrmEditaNiveles.frx":6029F
      Left            =   8160
      List            =   "FrmEditaNiveles.frx":602A1
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   -480
      Top             =   9360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   41
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":602A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":6C2F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":6DE47
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":79E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":85EEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":87A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":93A8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":955E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":97133
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":98C85
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":9A7D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":9C329
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":9DE7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":9F9CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":A151F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":C8031
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":F8083
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":1280D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":158127
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":164179
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":165CCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":168D3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":174D8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":1768E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":178433
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":179F85
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":185FD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":187B29
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":193B7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":1C3BCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":1CFC1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":1DBC71
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":1DD7C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":1E9815
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":1F5867
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":2018B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":20D90B
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":21995D
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":2259AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":231A01
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditaNiveles.frx":23DA53
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6135
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   10821
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList3"
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
End
Attribute VB_Name = "FrmEditarNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Seleccion As String


Private Sub CmdAceptar_Click()
On Error GoTo TipoErrs
'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
     If NivelAcceso < Val(TxtNivel.Text) Then
        MsgBox "Imposible Modificar su Nivel es Inferior", vbExclamation, "Sistema de Nominas"
        Exit Sub
      End If
      
  Select Case Seleccion 'Me.ListAcceso.Text
        Case "Presupuesto"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Presupuesto"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '////////////Verifico Grabar /////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Presupuesto"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
        '////////////Verifico Leer /////////////////////
         
         If Me.ChLeer.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Presupuesto"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
     '////////////Verifico Borrar /////////////////////
         
         If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Presupuesto"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
  
  
  
  
      Case "Editar Niveles"
      
      '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Niveles'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Niveles"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Niveles'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
       '////////////Verifico Grabar Niveles//////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Niveles'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Niveles"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Niveles'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
      Case "Cuentas"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Cuentas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cuentas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '////////////Verifico Grabar /////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Cuentas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Cuentas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
        '////////////Verifico Leer /////////////////////
         
         If Me.ChLeer.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Cuentas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Cuentas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
     '////////////Verifico Borrar /////////////////////
         
         If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Cuentas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Cuentas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
      
      Case "Grupo Cuentas"
      '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Grupo Cuentas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '////////////Verifico Grabar /////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Grupo Cuentas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
        '////////////Verifico Leer /////////////////////
         
         If Me.ChLeer.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Grupo Cuentas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
     '////////////Verifico Borrar /////////////////////
         
         If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Grupo Cuentas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
      
      
         
      Case "Contratistas"
            '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Contratistas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Contratistas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '////////////Verifico Grabar /////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Contratistas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Contratistas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
        '////////////Verifico Leer /////////////////////
         
         If Me.ChLeer.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Contratistas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Contratistas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
     '////////////Verifico Borrar /////////////////////
         
         If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Contratistas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Contratistas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
      
        
      Case "Empleados"
                 '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Empleados"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Empleados'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '////////////Verifico Grabar /////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Empleados"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Empleados'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
        '////////////Verifico Leer /////////////////////
         
         If Me.ChLeer.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Empleados"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Empleados'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
     '////////////Verifico Borrar /////////////////////
         
         If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Empleados"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Empleados'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
      
    Case "Activo Fijo"
    
          '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Activo Fijo"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '////////////Verifico Grabar /////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Activo Fijo"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
        '////////////Verifico Leer /////////////////////
         
         If Me.ChLeer.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Activo Fijo"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
     '////////////Verifico Borrar /////////////////////
         
         If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Activo Fijo"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
         
    Case "Periodos"
    
          '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Periodos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Periodos"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Periodos'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '////////////Verifico Grabar /////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Periodos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Periodos"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Periodos'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
        '////////////Verifico Leer /////////////////////
         
         If Me.ChLeer.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Periodos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Periodos"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Periodos'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
               
    Case "Transacciones"
               '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Transacciones"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Transacciones'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '////////////Verifico Grabar /////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Transacciones"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
        '////////////Verifico Leer /////////////////////
         
         If Me.ChLeer.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Transacciones"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Transacciones'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
     '////////////Verifico Borrar /////////////////////
         
         If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Transacciones"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Transacciones'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
    Case "Cheques"
              '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Cheques"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cheques'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '////////////Verifico Grabar /////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Cheques"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Cheques'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
        '////////////Verifico Leer /////////////////////
         
         If Me.ChLeer.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Cheques"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Cheques'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
     '////////////Verifico Borrar /////////////////////
         
         If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Cheques"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Cheques'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
    
        
    Case "Calcular Depreciacion"
               '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Depreciacion'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Depreciacion"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Depreciacion'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
     
    Case "Usuarios"
             '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Usuarios"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Usuarios'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '////////////Verifico Grabar /////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Usuarios"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
        '////////////Verifico Leer /////////////////////
         
         If Me.ChLeer.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Usuarios"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Usuarios'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
     '////////////Verifico Borrar /////////////////////
         
         If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Usuarios"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
    
    Case "Tasa de Cambio"
               '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tasa Cambio"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '////////////Verifico Grabar /////////////////////
         
         If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tasa Cambio"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
        '////////////Verifico Leer /////////////////////
         
         If Me.ChLeer.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Tasa Cambio"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
          
     '////////////Verifico Borrar /////////////////////
         
         If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tasa Cambio"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
       
        End If
    Case "Reportes Generales"
          '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Generales'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Reportes Generales"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Generales'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
 
    
          
    Case "Reportes Movimientos"
                   '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Movimientos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Reportes Movimientos"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Movimientos'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
         
         
    Case "Reportes Bancos"
                  '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Reportes Bancos"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
     
    End Select
      
Exit Sub
TipoErrs:
 ControlErrores
Unload Me
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub
Private Sub DBLNEmpleado_Click()
On Error GoTo TipoErrs
LblNombre.Caption = DBLNEmpleado.Text
DtaPasword2.Refresh
      Do While Not DtaPasword2.Recordset.EOF
       If DtaPasword2.Recordset!NombreUsuario = DBLNEmpleado.Text Then
         CodUsuario = DtaPasword2.Recordset!CodUsuario
         TxtNivel.Text = DtaPasword2.Recordset!Nivel
         Exit Do
       End If
        DtaPasword2.Recordset.MoveNext
      Loop
 
 DtaNacceso.Refresh
      Do While Not DtaNacceso.Recordset.EOF
         If DtaNacceso.Recordset("CodUsuario") = TxtCodPasword.Text Then
           
           
           
           
           Exit Sub
         End If
        DtaNacceso.Recordset.MoveNext
      Loop
      
      
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub Form_Activate()
On Error GoTo TipoErrs

Dim NodX As Node
Dim Relatives As String, RelationsShips As String
Dim LLave As String, Texto As String, Imagen1 As Integer
Dim Imagen2 As Integer, SPantallas As Variant, i As Double, Textos As Variant

 
If Not CodigoUsuario = 0 Then
 Me.DtaConsulta.ConnectionString = Conexion
 Me.DtaConsulta.RecordSource = "select * from accesos"
 Me.DtaConsulta.Refresh
 
 Me.DtaPasword.ConnectionString = Conexion
 Me.DtaPasword.RecordSource = "select * from usuarios"
 Me.DtaPasword.Refresh
 
 Me.DtaPasword2.ConnectionString = Conexion
 Me.DtaPasword2.RecordSource = "select * from usuarios"
 Me.DtaPasword2.Refresh
 
 Me.DtaNacceso.ConnectionString = Conexion
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Niveles'))"
 Me.DtaNacceso.Refresh
 
 Me.DtaNacceso2.ConnectionString = Conexion
 Me.DtaNacceso2.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Niveles'))"
 Me.DtaNacceso2.Refresh
 
 
   Me.CmdAceptar.Enabled = True
 If Me.DtaNacceso.Recordset.EOF Then
'   Me.CmdAceptar.Enabled = False

 End If
 
Else

  Me.CmdAceptar.Enabled = False

End If


SPantallas = Array("A", "B", "C", "D", "E", "F", "G")
Textos = Array("Cuentas", "Catalogo", "Procesos", "Opciones", "Ayuda", "Reportes", "Contabilizar")


For i = 0 To 6
    LLave = SPantallas(i)
    
    
     

    
    '///////////////////////////////////////////////////////////////////////////////////
    '///////////////////AGREGO LOS SUBNIVELES////////////////////////////////////////////

    
   Select Case i
     Case 0
       Set NodX = Me.TreeView1.Nodes.Add(, , LLave, Textos(i), 1, 1)
       
           Relatives = LLave
           RelationsShips = 4
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Cuentas", "Cuentas", 5, 5)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Grupo", "Grupo Cuentas", 6, 6)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "CuentaMayor", "Cuentas de Mayor", 7, 7)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Auxiliar", "Auxiliar de Cuentas", 8, 8)
'       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "ReporteDiario", "Reporte Diario", 1, 1)
      Case 1
       Set NodX = Me.TreeView1.Nodes.Add(, , LLave, Textos(i), 2, 2)
            Relatives = LLave
            RelationsShips = 4
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Empleados", "Empleados", 9, 9)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Contratista", "Contratista", 10, 10)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Periodos", "Periodos", 11, 11)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "ActivoFijo", "Activo Fijo", 12, 12)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Usuarios", "Usuarios", 13, 13)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Tasa", "Tasa de Cambio", 14, 14)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Departamento", "Departamento", 15, 15)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Nivel Usuarios", "Nivel de Usuarios", 16, 16)
      Case 2
       Set NodX = Me.TreeView1.Nodes.Add(, , LLave, Textos(i), 3, 3)
            Relatives = LLave
            RelationsShips = 4
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Transacciones", "Transacciones", 20, 20)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Cheques", "Cheques", 21, 21)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Depreciacion", "Calcular Depreciacion", 22, 22)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Presupuesto", "Presupuesto", 23, 23)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Prorrateo", "Prorrateo", 24, 24)

      Case 3
       Set NodX = Me.TreeView1.Nodes.Add(, , LLave, Textos(i), 4, 4)
            Relatives = LLave
            RelationsShips = 4
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Configuracion", "Configuracion", 25, 25)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Respaldar", "Respaldar", 26, 26)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Compañia", "Compañia", 27, 27)
       
      Case 4
       Set NodX = Me.TreeView1.Nodes.Add(, , LLave, Textos(i), 17, 17)
            Relatives = LLave
            RelationsShips = 4
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Configuracion", "Configuracion de Cheques", 28, 28)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "ImportarTransacciones", "Importar Transacciones", 29, 29)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Exportar Transacciones", "Exportar Transacciones", 30, 30)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "ImportarCuentas", "Importar Cuentas", 31, 31)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "Auditor", "Auditor", 32, 32)
       
      Case 5
       Set NodX = Me.TreeView1.Nodes.Add(, , LLave, Textos(i), 18, 18)
             Relatives = LLave
             RelationsShips = 4
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "ReportesGenerales", "Reportes Generales", 33, 33)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "ReportesMovimientos", "Reportes de Movimientos", 34, 34)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "ReportesBancos", "Reporte de Bancos", 35, 35)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "EstadosFinancieros", "Estados Financieros", 36, 36)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "AnalisisFinancieros", "Analisis Financieros", 37, 37)
       
      Case 6
       Set NodX = Me.TreeView1.Nodes.Add(, , LLave, Textos(i), 19, 19)
             Relatives = LLave
             RelationsShips = 4
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "ConexionFacturacion", "Conexion Facturacion", 38, 38)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "ContabilizarFacturacion", "Contabilizar Facturacion", 39, 39)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "ConexionNomina", "Conexion Nominas", 40, 40)
       Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, i & "ContabilizarNomina", "Contabilizar Nomina", 41, 41)

   End Select


Next

Me.TreeView1.Nodes(Me.TreeView1.Nodes.Count).EnsureVisible
NodoBase = True
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd



'With Me.DtaNacceso2
'
'   .ConnectionString = Conexion
'End With
'
'With Me.DtaConsulta
'
'   .ConnectionString = Conexion
'End With
'
'With Me.DtaNacceso
'
'   .ConnectionString = Conexion
'End With
'
'With Me.DtaPasword
'
'   .ConnectionString = Conexion
'End With
'
'With Me.DtaPasword2
'
'   .ConnectionString = Conexion
'End With

Me.ListAcceso.AddItem ("Editar Niveles")
Me.ListAcceso.AddItem ("Cuentas")
Me.ListAcceso.AddItem ("Grupo Cuentas")
Me.ListAcceso.AddItem ("Contratistas")
Me.ListAcceso.AddItem ("Empleados")
Me.ListAcceso.AddItem ("Activo Fijo")
Me.ListAcceso.AddItem ("Periodos")
Me.ListAcceso.AddItem ("Transacciones")
Me.ListAcceso.AddItem ("Presupuesto")
Me.ListAcceso.AddItem ("Cheques")
Me.ListAcceso.AddItem ("Calcular Depreciacion")
Me.ListAcceso.AddItem ("Usuarios")
Me.ListAcceso.AddItem ("Tasa de Cambio")
Me.ListAcceso.AddItem ("Reportes Generales")
Me.ListAcceso.AddItem ("Reportes Movimientos")
Me.ListAcceso.AddItem ("Reportes Bancos")


End Sub

Private Sub ListAcceso_Click()
On Error GoTo TipoErrs
    'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
    
    Select Case Me.ListAcceso.Text
      
            Case "Presupuesto"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
    
      Case "Editar Niveles"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
     '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Niveles'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Niveles'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
          
          
          
          
      Case "Cuentas"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
      
      Case "Grupo Cuentas"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
                  '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
      Case "Contratistas"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
                  '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
      Case "Empleados"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
      
                           '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
      
      Case "Activo Fijo"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Periodos"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Periodos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Periodos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Periodos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
   
          
    Case "Transacciones"
         Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Cheques"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Calcular Depreciacion"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Depreciacion'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
   

          
    Case "Usuarios"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Tasa de Cambio"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
    Case "Reportes Generales"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Generales'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Reportes Generales'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Reportes Generales'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Reportes Generales'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Reportes Movimientos"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Movimientos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Reportes Movimientos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Reportes Movimientos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Reportes Movimientos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Reportes Bancos"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
    
    End Select
    
      DtaNacceso.Refresh
      Do While Not DtaNacceso.Recordset.EOF
         If DtaNacceso.Recordset("CodUsuario") = TxtCodPasword.Text Then
          
           Exit Sub
         End If
        DtaNacceso.Recordset.MoveNext
      Loop
     
Exit Sub
TipoErrs:
ControlErrores
Unload Me
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()


  If Me.Picture1.Visible = True Then
    Me.Picture1.Visible = False
    Me.Picture2.Visible = True
  Else
    Me.Picture2.Visible = False
    Me.Picture1.Visible = True
    
  End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 Dim Opciones As String
  KeyPrincipal = Node.Key
  
  Opciones = Mid(KeyPrincipal, 2, Len(KeyPrincipal) - 1)
  
  Seleccion = Opciones
  
      Select Case Opciones
      
          Case "Presupuesto"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Presupuesto'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
    
      Case "Editar Niveles"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
     '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Niveles'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Niveles'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
          
          
          
          
      Case "Cuentas"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
      
      Case "Grupo"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
                  '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Grupo Cuentas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
      Case "Contratistas"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
                  '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Contratistas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
      Case "Empleados"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
      
                           '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
      
      Case "Activo Fijo"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Activo Fijo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Periodos"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Periodos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Periodos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Periodos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
   
          
    Case "Transacciones"
         Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Transacciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Cheques"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Cheques'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Calcular Depreciacion"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Depreciacion'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
   

          
    Case "Usuarios"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Tasa de Cambio"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True
          Me.ChLeer.Enabled = True
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tasa Cambio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
    Case "Reportes Generales"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Generales'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Reportes Generales'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Reportes Generales'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Reportes Generales'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Reportes Movimientos"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Movimientos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Reportes Movimientos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Reportes Movimientos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Reportes Movimientos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
          
    Case "Reportes Bancos"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False
          Me.ChLeer.Enabled = False
          
                            '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
     '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
          '///////Chek leer////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Leer Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChLeer.Value = 0
         Else
          Me.ChLeer.Value = 1
         End If
     '///////Chek borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Reportes Bancos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
    
    End Select
    
      DtaNacceso.Refresh
      Do While Not DtaNacceso.Recordset.EOF
         If DtaNacceso.Recordset("CodUsuario") = TxtCodPasword.Text Then
          
           Exit Sub
         End If
        DtaNacceso.Recordset.MoveNext
      Loop
End Sub
