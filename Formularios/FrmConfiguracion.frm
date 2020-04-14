VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmConfiguracion 
   BorderStyle     =   0  'None
   Caption         =   "Configuración del Sistema"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7200
   HelpContextID   =   90000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales de la Empesa"
      TabPicture(0)   =   "FrmConfiguracion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Grupos de Cuentas para Utilidad Bruta"
      TabPicture(1)   =   "FrmConfiguracion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   6855
         Begin VB.TextBox TxtI 
            Height          =   285
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   28
            Top             =   1800
            Width           =   3135
         End
         Begin VB.TextBox TxtC 
            Height          =   285
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   27
            Top             =   2280
            Width           =   3135
         End
         Begin VB.TextBox TxtOI 
            Height          =   405
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   26
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox TxtCO 
            Height          =   405
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   24
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox TxtIO 
            Height          =   405
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   21
            Top             =   240
            Width           =   3135
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmConfiguracion.frx":0038
            TabIndex        =   22
            Top             =   240
            Width           =   3375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmConfiguracion.frx":00E4
            TabIndex        =   23
            Top             =   720
            Width           =   3495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmConfiguracion.frx":018C
            TabIndex        =   25
            Top             =   1200
            Width           =   3495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmConfiguracion.frx":0242
            TabIndex        =   29
            Top             =   1800
            Width           =   3255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmConfiguracion.frx":02D8
            TabIndex        =   30
            Top             =   2280
            Width           =   3495
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   6855
         Begin VB.CommandButton Command1 
            Height          =   375
            Left            =   6360
            Picture         =   "FrmConfiguracion.frx":036A
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox TxtRutaExe 
            Height          =   375
            Left            =   2880
            TabIndex        =   31
            Top             =   2760
            Width           =   3495
         End
         Begin VB.TextBox TxtTelefono 
            Height          =   375
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   19
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox TxtRucEmpresa 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   18
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox TxtDireccionEmpresa 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   17
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox TxtNombreEmpresa2 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   16
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox TxtNombreEmpresa1 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   11
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox TxtRutaLogo 
            Height          =   375
            Left            =   2880
            TabIndex        =   9
            Top             =   2280
            Width           =   3495
         End
         Begin VB.PictureBox ImgLogo2 
            AutoSize        =   -1  'True
            Height          =   2055
            Left            =   120
            ScaleHeight     =   1995
            ScaleWidth      =   2235
            TabIndex        =   6
            Top             =   240
            Width           =   2295
            Begin VB.Image ImgLogo 
               Height          =   2055
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2295
            End
         End
         Begin VB.CommandButton CmdBuscarLogo 
            Height          =   375
            Left            =   6360
            Picture         =   "FrmConfiguracion.frx":0820
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2280
            Width           =   375
         End
         Begin VB.TextBox TxtNombreEmpresa 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   4
            Top             =   120
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmConfiguracion.frx":0CD6
            TabIndex        =   7
            Top             =   2400
            Width           =   2535
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "FrmConfiguracion.frx":0D72
            TabIndex        =   8
            Top             =   120
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "FrmConfiguracion.frx":0DEC
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "FrmConfiguracion.frx":0E68
            TabIndex        =   12
            Top             =   600
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "FrmConfiguracion.frx":0EE4
            TabIndex        =   13
            Top             =   840
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "FrmConfiguracion.frx":0F54
            TabIndex        =   14
            Top             =   1080
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "FrmConfiguracion.frx":0FC6
            TabIndex        =   15
            Top             =   1320
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmConfiguracion.frx":1036
            TabIndex        =   33
            Top             =   2880
            Width           =   2535
         End
      End
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc AdoConfiguracion 
      Height          =   375
      Left            =   120
      Top             =   5280
      Visible         =   0   'False
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
      CommandType     =   1
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
      Caption         =   "AdoConfiguracion"
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
   Begin MSComDlg.CommonDialog CMRutaFoto 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   256
   End
End
Attribute VB_Name = "FrmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBuscarLogo_Click()
Dim retval
Dim OpenFileName As String
    On Error Resume Next
    ' Set the commom dialog properties we need
    If Me.TxtRutaLogo.Text <> "" Then
       CMRutaFoto.InitDir = Me.TxtRutaLogo.Text
    End If
    CMRutaFoto.FileName = ""
    ' We will load BMP, JPG, and TIF files
    CMRutaFoto.Filter = "Image Files |*.bmp;*.gif;*.jpg;*.png;*.tif|All files |*.*"
    ' Display common dialog box
    CMRutaFoto.ShowOpen
    Me.TxtRutaLogo.Text = CMRutaFoto.FileName
    
    Me.imgLogo.Picture = LoadPicture(Me.TxtRutaLogo.Text)
End Sub

Private Sub CmdCerrar_Click()
Unload Me
End Sub


Private Sub CmdGuardar_Click()
On Error GoTo TipoErr
'verifico que las numeraciones no esten siendo utilizadas

'Guardando configuración
    MDIPrimero.AdoConfiguracion.Refresh
    MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo = Me.TxtRutaLogo.Text
    
    MDIPrimero.AdoConfiguracion.Recordset!Telefono = Me.TxtTelefono.Text
    MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa = Me.TxtNombreEmpresa.Text
    MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa1 = Me.TxtNombreEmpresa1.Text
    MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa2 = Me.TxtNombreEmpresa2.Text
    MDIPrimero.AdoConfiguracion.Recordset!Direccion = Me.TxtDireccionEmpresa.Text
    MDIPrimero.AdoConfiguracion.Recordset!numerorUC = Me.TxtRucEmpresa.Text
    MDIPrimero.AdoConfiguracion.Recordset!Ruta_Actualizar = Me.TxtRutaExe.Text
 

    MDIPrimero.AdoConfiguracion.Recordset!IngresosOperativos = Me.TxtIO.Text
    MDIPrimero.AdoConfiguracion.Recordset!CostosOperativos = Me.TxtCO.Text
    MDIPrimero.AdoConfiguracion.Recordset!OtrosINgresos = Me.TxtOI.Text
    MDIPrimero.AdoConfiguracion.Recordset!Ingresos = Me.TxtI.Text
    MDIPrimero.AdoConfiguracion.Recordset!Costos = Me.TxtC.Text

MDIPrimero.AdoConfiguracion.Recordset.Update

MDIPrimero.AdoConfiguracion.Refresh


MsgBox "Configuración Guardada con éxito", vbInformation
Unload Me
Exit Sub
TipoErr:
  MsgBox err.Description
End Sub


Private Sub Command1_Click()
Dim retval
Dim OpenFileName As String
    On Error Resume Next
    ' Set the commom dialog properties we need
    If Me.TxtRutaLogo.Text <> "" Then
       CMRutaFoto.InitDir = Me.TxtRutaLogo.Text
    End If
    CMRutaFoto.FileName = ""
    ' We will load BMP, JPG, and TIF files
    CMRutaFoto.Filter = "Acualizacion |*.exe"
    ' Display common dialog box
    CMRutaFoto.ShowOpen
    Me.TxtRutaExe.Text = CMRutaFoto.FileName
End Sub

Private Sub Form_Load()
On Error GoTo TipoErr
MDIPrimero.Skin1.ApplySkin Me.hWnd

If MDIPrimero.AdoConfiguracion.Recordset.RecordCount > 0 Then
      
  
   
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa) Then
       Me.TxtNombreEmpresa.Text = MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa
    End If
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa1) Then
       Me.TxtNombreEmpresa1.Text = MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa1
    End If
    
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa2) Then
       Me.TxtNombreEmpresa2.Text = MDIPrimero.AdoConfiguracion.Recordset!NombreEmpresa2
    End If
    
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!Direccion) Then
       Me.TxtDireccionEmpresa.Text = MDIPrimero.AdoConfiguracion.Recordset!Direccion
    End If
    
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!numerorUC) Then
       Me.TxtRucEmpresa.Text = MDIPrimero.AdoConfiguracion.Recordset!numerorUC
    End If
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!Telefono) Then
       Me.TxtTelefono.Text = MDIPrimero.AdoConfiguracion.Recordset!Telefono
    End If
    
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!IngresosOperativos) Then
     Me.TxtIO.Text = MDIPrimero.AdoConfiguracion.Recordset!IngresosOperativos
    End If
    
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!CostosOperativos) Then
     Me.TxtCO.Text = MDIPrimero.AdoConfiguracion.Recordset!CostosOperativos
    End If
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!OtrosINgresos) Then
     Me.TxtOI.Text = MDIPrimero.AdoConfiguracion.Recordset!OtrosINgresos
    End If
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!Ingresos) Then
     Me.TxtI.Text = MDIPrimero.AdoConfiguracion.Recordset!Ingresos
    End If
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!Costos) Then
     Me.TxtC.Text = MDIPrimero.AdoConfiguracion.Recordset!Costos
    End If
    
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!Ruta_Actualizar) Then
       If Dir(MDIPrimero.AdoConfiguracion.Recordset!Ruta_Actualizar) <> "" Then
        Me.TxtRutaExe.Text = MDIPrimero.AdoConfiguracion.Recordset!Ruta_Actualizar
       End If
    End If
    
    If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) Then
       If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.imgLogo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
        Me.TxtRutaLogo.Text = MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo
       End If
    End If
End If
Exit Sub
TipoErr:
  MsgBox err.Description
End Sub


