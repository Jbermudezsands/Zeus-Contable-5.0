VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios."
   ClientHeight    =   3570
   ClientLeft      =   3150
   ClientTop       =   2955
   ClientWidth     =   4035
   HelpContextID   =   17
   Icon            =   "FrmUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4035
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   4095
      TabIndex        =   12
      Top             =   0
      Width           =   4095
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuarios"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   1320
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Top             =   120
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   6720
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   480
         Picture         =   "FrmUsuarios.frx":27A2
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.TextBox TxtConfirma 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   9
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmUsuarios.frx":2ECE
      TabIndex        =   11
      Top             =   1920
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmUsuarios.frx":2F3C
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmUsuarios.frx":2FA8
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmUsuarios.frx":3020
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DBCnombre 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaUsuarios 
      Height          =   375
      Left            =   240
      Top             =   4680
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "DtaUsuarios"
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
   Begin MSAdodcLib.Adodc DtaNacceso2 
      Height          =   375
      Left            =   240
      Top             =   4320
      Width           =   2655
      _ExtentX        =   4683
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
   Begin MSAdodcLib.Adodc DtaUsuarios2 
      Height          =   375
      Left            =   240
      Top             =   4320
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "DtaUsuarios2"
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
      Height          =   375
      Left            =   360
      Top             =   4440
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.Frame Frame1 
      Caption         =   "Botones de Comando"
      ForeColor       =   &H00008000&
      Height          =   735
      Left            =   120
      MouseIcon       =   "FrmUsuarios.frx":30A0
      TabIndex        =   3
      Top             =   2760
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Height          =   375
         Left            =   2640
         Picture         =   "FrmUsuarios.frx":34E2
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   1440
         Picture         =   "FrmUsuarios.frx":3630
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   120
         Picture         =   "FrmUsuarios.frx":377E
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Txtpasword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   9
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox TxtAcceso 
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
   End
End
Attribute VB_Name = "FrmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Permiso = True
Dim Usuario As Integer

   
   
   
     If Not IsNumeric(TxtAcceso.Text) Then
      MsgBox "El Nivel no es Correcto", vbExclamation, "Sistema de Facturacion"
        TxtAcceso.Text = ""
        TxtAcceso.SetFocus
        Exit Sub
     End If
     
     If Val(Nivel2) > Val(NivelAcceso) Then
        MsgBox "No Puede Cambiar el Nivel", vbExclamation, "Sistema de Facturacion"
        Exit Sub
     End If
     
     If (Txtpasword.Text) <> (TxtConfirma.Text) Then
       MsgBox "Confime Nuemvamente", vbInformation, "Sistema de Facturacion"
       TxtConfirma.Text = ""
       TxtConfirma.SetFocus
       Exit Sub
     End If

DtaUsuarios.Refresh
     Do While Not DtaUsuarios.Recordset.EOF
        If DtaUsuarios.Recordset("NombreUsuario") = DBCnombre.Text Then
             If DtaUsuarios.Recordset("Nivel") < Val(TxtAcceso.Text) Then
               MsgBox "No Puede Ponerse un Nivel Superior", vbExclamation, "Sistema de Facturacion"
               Exit Sub
             End If
             'DtaUsuarios.Recordset.Edit
             DtaUsuarios.Recordset("Nivel") = TxtAcceso.Text
             DtaUsuarios.Recordset("Clave") = Txtpasword.Text
             DtaUsuarios.Recordset.Update
             DBCnombre.Text = ""
             Permiso = True
             Exit Sub
        End If
       DtaUsuarios.Recordset.MoveNext
      Loop
             
             If NivelAcceso < Val(TxtAcceso.Text) Then
               MsgBox "No Puede Crear un Nivel Superior", vbExclamation, "Sistema de Facturacion"
               Exit Sub
             End If
                       
             
             
             DtaUsuarios.Recordset.AddNew
             DtaUsuarios.Recordset("NombreUsuario") = DBCnombre.Text
             DtaUsuarios.Recordset("Nivel") = TxtAcceso.Text
             DtaUsuarios.Recordset("Clave") = Txtpasword.Text
             DtaUsuarios.Recordset.Update
             CodUsuario = DtaUsuarios.Recordset("CodUsuario")
             
             DBCnombre.Text = ""
             Permiso = True
             
           Me.DtaNacceso.Recordset.AddNew
'             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Niveles"
           Me.DtaNacceso.Recordset.Update
             
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Niveles"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Cuentas"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Cuentas"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Cuentas"
           Me.DtaNacceso.Recordset.Update
       
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Cuentas"
           Me.DtaNacceso.Recordset.Update
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Grupo Cuentas"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Grupo Cuentas"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Grupo Cuentas"
           Me.DtaNacceso.Recordset.Update
       
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Grupo Cuentas"
           Me.DtaNacceso.Recordset.Update
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Contratistas"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Contratistas"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Contratistas"
           Me.DtaNacceso.Recordset.Update
       
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Contratistas"
           Me.DtaNacceso.Recordset.Update
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Empleados"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Empleados"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Empleados"
           Me.DtaNacceso.Recordset.Update
       
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Empleados"
           Me.DtaNacceso.Recordset.Update
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Activo Fijo"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Activo Fijo"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Activo Fijo"
           Me.DtaNacceso.Recordset.Update
       
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Activo Fijo"
           Me.DtaNacceso.Recordset.Update
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Periodos"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Periodos"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Periodos"
           Me.DtaNacceso.Recordset.Update
       
         Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Transacciones"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Transacciones"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Transacciones"
           Me.DtaNacceso.Recordset.Update
       
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Transacciones"
           Me.DtaNacceso.Recordset.Update
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Cheques"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Cheques"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Cheques"
           Me.DtaNacceso.Recordset.Update
       
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Cheques"
           Me.DtaNacceso.Recordset.Update
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Depreciacion"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Usuarios"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Usuarios"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Usuarios"
           Me.DtaNacceso.Recordset.Update
       
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Usuarios"
           Me.DtaNacceso.Recordset.Update
         
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tasa Cambio"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tasa Cambio"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Tasa Cambio"
           Me.DtaNacceso.Recordset.Update
       
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tasa Cambio"
           Me.DtaNacceso.Recordset.Update
           
            Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Presupuesto"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Presupuesto"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Leer Presupuesto"
           Me.DtaNacceso.Recordset.Update
       
         
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Presupuesto"
           Me.DtaNacceso.Recordset.Update
         
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Reportes Generales"
           Me.DtaNacceso.Recordset.Update
              
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Reportes Movimientos"
           Me.DtaNacceso.Recordset.Update
           
          Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Reportes Bancos"
           Me.DtaNacceso.Recordset.Update
       Me.DBCnombre.Refresh
           
Exit Sub
'TipoErrs:
'ControlErrores
'Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo TipoErrs
 Dim Respuesta, Rsp
'Elimino el registro activo en la pantalla
  Set Rsp = DtaUsuarios.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando al Usuario: " & DBCnombre.Text)
   If Respuesta = 6 Then
     DtaUsuarios2.Refresh
       Do While Not DtaUsuarios2.Recordset.EOF
        If DtaUsuarios2.Recordset("NombreUsuario") = DBCnombre.Text Then
           Rsp.Delete
           If Me.DtaUsuarios.Recordset.RecordCount > 0 Then
                DtaUsuarios.Recordset.MovePrevious
                Me.DBCnombre.Refresh
                DBCnombre.Text = DtaUsuarios.Recordset("NombreUsuario")
                Permiso = True
            Else
                DBCnombre.Text = ""
            End If
           Exit Sub
        End If
       DtaUsuarios2.Recordset.MoveNext
      Loop
    Else
   Exit Sub
   End If
MsgBox "No se puede Eliminar,No Existe este Registro", vbCritical, "Sistema de Facturacion"
Exit Sub
TipoErrs:
If err = 3021 Then
 DtaUsuarios.Refresh
 DBCnombre.Text = ""
 Permiso = True
Exit Sub
Else
'ControlErrores
 Unload Me
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
Private Sub DBCNombre_Change()
DtaUsuarios.Refresh
If Not Me.DBCnombre.Text = "" Then
   Do While Not DtaUsuarios.Recordset.EOF
        If DtaUsuarios.Recordset("NombreUsuario") = DBCnombre.Text Then
           Nivel2 = DtaUsuarios.Recordset("Nivel")
           If NivelAcceso >= Nivel2 Then
           TxtAcceso.Text = DtaUsuarios.Recordset("Nivel")
           Txtpasword.Text = DtaUsuarios.Recordset("Clave")
           Permiso = True
           Exit Sub
           Else
           MsgBox "No Puede Cambiar el Nivel", vbExclamation, "Sistema de Facturacion"
           DBCnombre.Text = ""
           Exit Sub
           End If
        End If
       DtaUsuarios.Recordset.MoveNext
      Loop
 End If
TxtAcceso.Text = ""
Txtpasword.Text = ""
TxtConfirma.Text = ""
Permiso = False
End Sub

Private Sub DBCNombre_DblClick(Area As Integer)
DtaUsuarios.Refresh
Do While Not DtaUsuarios.Recordset.EOF
        If DtaUsuarios.Recordset("NombreUsuario") = DBCnombre.Text Then
           TxtAcceso.Text = DtaUsuarios.Recordset("NivelAcceso")
           Txtpasword.Text = DtaUsuarios.Recordset("Pasword")
           Exit Sub
        End If
       DtaUsuarios.Recordset.MoveNext
      Loop
TxtAcceso.Text = ""
Txtpasword.Text = ""
End Sub

Private Sub DBCnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 TxtAcceso.SetFocus
End If
End Sub

Private Sub Form_Activate()
On Error GoTo TipoErrs
 
If Not CodigoUsuario = 0 Then
    Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Usuarios'))"
    Me.DtaNacceso.Refresh
    If Me.DtaNacceso.Recordset.EOF Then
      Me.Command1.Enabled = False
    End If
    Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Usuarios'))"
    Me.DtaNacceso.Refresh
    If Me.DtaNacceso.Recordset.EOF Then
      Me.Command2.Enabled = False
    
    End If
End If
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
Permiso = True
With Me.DtaNacceso2
   .ConnectionString = Conexion
End With

With Me.DtaUsuarios
   .ConnectionString = Conexion
   .RecordSource = "select * from usuarios"
   .Refresh
End With
LlenarDataCombos Me.DtaUsuarios, Me.DBCnombre, "NombreUsuario", "CodUsuario"
With Me.DtaUsuarios2
   .ConnectionString = Conexion
   .RecordSource = "select * from usuarios"
   .Refresh
End With

With Me.DtaNacceso
   .ConnectionString = Conexion
   .RecordSource = "Accesos"
   .Refresh
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Dim Respuesta As Integer
 If Permiso = False Then
   Respuesta = MsgBox("Desea Guardar los Cambios?", vbYesNo, "Usuario: " & DBCnombre.Text)
    If Respuesta = 6 Then
      Command1.Value = True
    End If
 End If
End Sub

Private Sub TxtAcceso_Change()
If Val(TxtAcceso.Text) > 100 Then
 MsgBox "No deber ser mayor de 100", vbCritical, "Sistema de Facturacion"
 TxtAcceso.Text = ""
 Exit Sub
End If
Permiso = False
End Sub

Private Sub TxtAcceso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Txtpasword.SetFocus
End If
End Sub

Private Sub TxtConfirma_Change()
Permiso = False
End Sub

Private Sub TxtConfirma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Command1.Enabled Then Command1.SetFocus
End If
End Sub


Private Sub Txtpasword_Change()
Permiso = False
End Sub

Private Sub Txtpasword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtConfirma.SetFocus
End If
End Sub
