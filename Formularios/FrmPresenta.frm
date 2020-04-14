VERSION 5.00
Object = "{AF8CD3F4-666F-11D1-940D-000021A73813}#5.0#0"; "osProgress.ocx"
Begin VB.Form FrmPresenta 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3075
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmPresenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DtaElimina 
      Caption         =   "DtaElimina"
      Connect         =   "Access"
      DatabaseName    =   "C:\Enlace\Enlace.jbh"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Data DtaConsulta 
      Caption         =   "DtaConsulta"
      Connect         =   "Access"
      DatabaseName    =   "C:\Enlace\Enlace.jbh"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Height          =   2850
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7200
      Begin Progress.osProgress Barra 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   6855
         _ExtentX        =   6694
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox picTV 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   6735
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label Label15 
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
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Label14 
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
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Label13 
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
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Label12 
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
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Label11 
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
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Label10 
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
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Label0 
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
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   6735
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label7 
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
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label5 
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
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Image Img2 
         Height          =   480
         Left            =   1440
         Picture         =   "FrmPresenta.frx":000C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   600
      End
      Begin VB.Image img1 
         Height          =   465
         Left            =   480
         Picture         =   "FrmPresenta.frx":0CD6
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmPresenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Tape As New clsTape


Private Sub Form_Activate()
On Error GoTo TipoErrs
Dim Maximo As Integer, i As Integer
Dim Transaccion As String, Diferencia As Double, Ajuste As Double


FrmExportar.DtaRegistros.Recordset.MoveLast
Maximo = FrmExportar.DtaRegistros.Recordset.RecordCount
FrmExportar.DtaRegistros.Recordset.MoveFirst
With Barra
   .Min = 0
   .Value = 0
   .Max = Maximo
   i = 0
 Do While Not FrmExportar.DtaRegistros.Recordset.EOF
 '///////Almaceno las variables en caso de realizar un ajuste/////////
      TipoMovimiento = FrmExportar.DtaRegistros.Recordset.TipoMovimiento
      Fuente = FrmExportar.DtaRegistros.Recordset.Fuente
      Consecutivo = FrmExportar.DtaRegistros.Recordset.IdRegistros
                 Fecha = FrmExportar.DtaRegistros.Recordset.Fecha
                 NTransaccion = FrmExportar.DtaRegistros.Recordset.NTransaccion
                 CodCuenta = FrmExportar.DtaRegistros.Recordset.CodCuenta
                 CodDepartamento = FrmExportar.DtaRegistros.Recordset.CodDepartamento
                 CodAcciones = FrmExportar.DtaRegistros.Recordset.CodAcciones
                 ClaveProyecto = FrmExportar.DtaRegistros.Recordset.ClaveProyecto
                 NFactura = FrmExportar.DtaRegistros.Recordset.FacturaNumero
                 ReferenciaCh = FrmExportar.DtaRegistros.Recordset.RefCheque
                 Descripcion = FrmExportar.DtaRegistros.Recordset.Descripcion
                 FechaDescuento = FrmExportar.DtaRegistros.Recordset.FechaDescuento
                 FechaVencimiento = FrmExportar.DtaRegistros.Recordset.FechaVencimiento
                 ImporteDescuento = FrmExportar.DtaRegistros.Recordset.ImporteDescuento
                 ValorUnit = FrmExportar.DtaRegistros.Recordset.ValorUnitario
                 TipoTransaccion = FrmExportar.DtaRegistros.Recordset.TipoTransaccion
  Transaccion = FrmExportar.DtaRegistros.Recordset.NTransaccion
  Me.DtaConsulta.RecordSource = "SELECT Registros.NTransaccion, Sum(Registros.DebitoDolar) AS SumaDeDebitoDolar, Sum(Registros.CreditoDolar) AS SumaDeCreditoDolar, Sum([Registros]![DebitoDolar]-[Registros]![CreditoDolar]) AS Diferencia From Registros GROUP BY Registros.NTransaccion Having (((Registros.NTransaccion) = '" & Transaccion & "'))"
  Me.DtaConsulta.Refresh
  Diferencia = DtaConsulta.Recordset.Diferencia
  
  If Diferencia < 0 Then
   '/////Menor que Cero es se ajusta en el Debito/////////
   DtaConsulta.RecordSource = "SELECT Registros.NTransaccion, Registros.DebitoDolar, Registros.CreditoDolar From Registros Where (((Registros.NTransaccion) = '" & Transaccion & "') And ((Registros.DebitoDolar) <> 0)) ORDER BY Registros.NTransaccion"
   DtaConsulta.Refresh
  
   If Not DtaConsulta.Recordset.EOF Then
    Ajuste = DtaConsulta.Recordset.DebitoDolar
    Ajuste = Ajuste + Abs(Diferencia)
   Else
    If FechaDescuento = "" Then
       FechaDescuento = "        "
    End If
    If FechaVencimiento = "" Then
     FechaVencimiento = "        "
    End If
    FrmExportar.DtaVerifica.Recordset.AddNew
     FrmExportar.DtaVerifica.Recordset.TipoMovimiento = "03"
     FrmExportar.DtaVerifica.Recordset.Fuente = Fuente
     FrmExportar.DtaVerifica.Recordset.IdRegistros = Consecutivo
     FrmExportar.DtaVerifica.Recordset.Fecha = Fecha
     FrmExportar.DtaVerifica.Recordset.CodCuenta = CodCuenta
     FrmExportar.DtaVerifica.Recordset.CodDepartamento = CodDepartamento
     FrmExportar.DtaVerifica.Recordset.CodAcciones = CodAcciones
     FrmExportar.DtaVerifica.Recordset.ClaveProyecto = ClaveProyecto
     FrmExportar.DtaVerifica.Recordset.FacturaNumero = NFactura
     FrmExportar.DtaVerifica.Recordset.RefCheque = ReferenciaCh
     FrmExportar.DtaVerifica.Recordset.Descripcion = "Ajuste del Descuadre,Enlace Pacioli"
     FrmExportar.DtaVerifica.Recordset.FechaDescuento = FechaDescuento
     FrmExportar.DtaVerifica.Recordset.FechaVencimiento = FechaVencimiento
     FrmExportar.DtaVerifica.Recordset.ImporteDescuento = ImporteDescuento
     FrmExportar.DtaVerifica.Recordset.ValorUnitario = ValorUnit
     FrmExportar.DtaVerifica.Recordset.TipoTransaccion = TipoTransaccion
     FrmExportar.DtaVerifica.Recordset.NTransaccion = Transaccion
     FrmExportar.DtaVerifica.Recordset.DebitoDolar = Abs(Diferencia)
    FrmExportar.DtaVerifica.Recordset.Update
    Diferencia = 0
   End If
   
    '////Si la diferencia es mayor que 1 Agrego un ajuste a una cuenta///////////
    If Abs(Diferencia) < 1 Then
     If Not DtaConsulta.Recordset.EOF Then
      DtaConsulta.Recordset.Edit
       DtaConsulta.Recordset.DebitoDolar = Ajuste
      DtaConsulta.Recordset.Update
     End If
    Else
     If FechaDescuento = "" Then
        FechaDescuento = "        "
     End If
     If FechaVencimiento = "" Then
       FechaVencimiento = "        "
     End If
     FrmExportar.DtaVerifica.Recordset.AddNew
      FrmExportar.DtaVerifica.Recordset.TipoMovimiento = "03"
      FrmExportar.DtaVerifica.Recordset.Fuente = Fuente
      FrmExportar.DtaVerifica.Recordset.IdRegistros = Consecutivo
      FrmExportar.DtaVerifica.Recordset.Fecha = Fecha
      FrmExportar.DtaVerifica.Recordset.CodCuenta = CodCuenta
      FrmExportar.DtaVerifica.Recordset.CodDepartamento = CodDepartamento
      FrmExportar.DtaVerifica.Recordset.CodAcciones = CodAcciones
      FrmExportar.DtaVerifica.Recordset.ClaveProyecto = ClaveProyecto
      FrmExportar.DtaVerifica.Recordset.FacturaNumero = NFactura
      FrmExportar.DtaVerifica.Recordset.RefCheque = ReferenciaCh
      FrmExportar.DtaVerifica.Recordset.Descripcion = "Ajuste del Descuadre,Enlace Pacioli"
      FrmExportar.DtaVerifica.Recordset.FechaDescuento = FechaDescuento
      FrmExportar.DtaVerifica.Recordset.FechaVencimiento = FechaVencimiento
      FrmExportar.DtaVerifica.Recordset.ImporteDescuento = ImporteDescuento
      FrmExportar.DtaVerifica.Recordset.ValorUnitario = ValorUnit
      FrmExportar.DtaVerifica.Recordset.TipoTransaccion = TipoTransaccion
      FrmExportar.DtaVerifica.Recordset.NTransaccion = Transaccion
      FrmExportar.DtaVerifica.Recordset.DebitoDolar = Abs(Diferencia)
     FrmExportar.DtaVerifica.Recordset.Update
    End If
  ElseIf Diferencia > 0 Then
   '////Mayor que Cero Ajuste al Credito/////////////////
   DtaConsulta.RecordSource = "SELECT Registros.NTransaccion, Registros.DebitoDolar, Registros.CreditoDolar From Registros Where (((Registros.NTransaccion) = '" & Transaccion & "') And ((Registros.CreditoDolar) <> 0)) ORDER BY Registros.NTransaccion"
   DtaConsulta.Refresh
  
    If Not DtaConsulta.Recordset.EOF Then
     Ajuste = DtaConsulta.Recordset.CreditoDolar
     Ajuste = Ajuste + Diferencia
    Else
     If FechaDescuento = "" Then
        FechaDescuento = "        "
     End If
     If FechaVencimiento = "" Then
       FechaVencimiento = "        "
     End If
     FrmExportar.DtaVerifica.Recordset.AddNew
     FrmExportar.DtaVerifica.Recordset.TipoMovimiento = "07"
     FrmExportar.DtaVerifica.Recordset.Fuente = Fuente
     FrmExportar.DtaVerifica.Recordset.IdRegistros = Consecutivo
     FrmExportar.DtaVerifica.Recordset.Fecha = Fecha
     FrmExportar.DtaVerifica.Recordset.CodCuenta = CodCuenta
     FrmExportar.DtaVerifica.Recordset.CodDepartamento = CodDepartamento
     FrmExportar.DtaVerifica.Recordset.CodAcciones = CodAcciones
     FrmExportar.DtaVerifica.Recordset.ClaveProyecto = ClaveProyecto
     FrmExportar.DtaVerifica.Recordset.FacturaNumero = NFactura
     FrmExportar.DtaVerifica.Recordset.RefCheque = ReferenciaCh
     FrmExportar.DtaVerifica.Recordset.Descripcion = "Ajuste del Descuadre,Enlace Pacioli"
     FrmExportar.DtaVerifica.Recordset.FechaDescuento = FechaDescuento
     FrmExportar.DtaVerifica.Recordset.FechaVencimiento = FechaVencimiento
     FrmExportar.DtaVerifica.Recordset.ImporteDescuento = ImporteDescuento
     FrmExportar.DtaVerifica.Recordset.ValorUnitario = ValorUnit
     FrmExportar.DtaVerifica.Recordset.TipoTransaccion = TipoTransaccion
     FrmExportar.DtaVerifica.Recordset.NTransaccion = Transaccion
     FrmExportar.DtaVerifica.Recordset.CreditoDolar = Abs(Diferencia)
    FrmExportar.DtaVerifica.Recordset.Update
    Diferencia = 0
    End If
  
  
    If Diferencia < 1 Then
     If Not DtaConsulta.Recordset.EOF Then
      DtaConsulta.Recordset.Edit
       DtaConsulta.Recordset.CreditoDolar = Ajuste
      DtaConsulta.Recordset.Update
     End If
    Else
    If FechaDescuento = "" Then
      FechaDescuento = "        "
    End If
    If FechaVencimiento = "" Then
      FechaVencimiento = "        "
    End If
    FrmExportar.DtaVerifica.Recordset.AddNew
     FrmExportar.DtaVerifica.Recordset.TipoMovimiento = "07"
     FrmExportar.DtaVerifica.Recordset.Fuente = Fuente
     FrmExportar.DtaVerifica.Recordset.IdRegistros = Consecutivo
     FrmExportar.DtaVerifica.Recordset.Fecha = Fecha
     FrmExportar.DtaVerifica.Recordset.CodCuenta = CodCuenta
     FrmExportar.DtaVerifica.Recordset.CodDepartamento = CodDepartamento
     FrmExportar.DtaVerifica.Recordset.CodAcciones = CodAcciones
     FrmExportar.DtaVerifica.Recordset.ClaveProyecto = ClaveProyecto
     FrmExportar.DtaVerifica.Recordset.FacturaNumero = NFactura
     FrmExportar.DtaVerifica.Recordset.RefCheque = ReferenciaCh
     FrmExportar.DtaVerifica.Recordset.Descripcion = "Ajuste del Descuadre,Enlace Pacioli"
     FrmExportar.DtaVerifica.Recordset.FechaDescuento = FechaDescuento
     FrmExportar.DtaVerifica.Recordset.FechaVencimiento = FechaVencimiento
     FrmExportar.DtaVerifica.Recordset.ImporteDescuento = ImporteDescuento
     FrmExportar.DtaVerifica.Recordset.ValorUnitario = ValorUnit
     FrmExportar.DtaVerifica.Recordset.TipoTransaccion = TipoTransaccion
     FrmExportar.DtaVerifica.Recordset.NTransaccion = Transaccion
     FrmExportar.DtaVerifica.Recordset.CreditoDolar = Abs(Diferencia)
    FrmExportar.DtaVerifica.Recordset.Update
   End If
  End If
  
  i = i + 1
  If Maximo < i Then
   i = Maximo
  End If
  Me.Caption = "Procesando:  " & i & " de " & Maximo & " Registros "
      .Value = i
      DoEvents
      
    FrmExportar.DtaRegistros.Recordset.MoveNext
     
 Loop
 
End With

If Ultimo = True Then
 FrmExportar.DtaRegistros.RecordSource = "SELECT Registros.Control, Registros.IdRegistros, Registros.Fecha, Registros.NTransaccion, Registros.Fuente, Registros.CodCuenta, Registros.Descripcion, Registros.CodDepartamento, Registros.CodAcciones, Registros.ClaveProyecto, Registros.FacturaNumero, Registros.TipoMovimiento, Registros.RefCheque, Registros.FechaDescuento, Registros.FechaVencimiento, Registros.ImporteTransaccionDebito, Registros.ImporteTransaccionCredito, Registros.ImporteDescuento, Registros.ValorUnitario, Registros.TipoTransaccion, Registros.DebitoDolar, Registros.CreditoDolar From Registros Where (((Registros.IdRegistros) = " & Consecutivo & ")) ORDER BY Registros.Fecha, Registros.NTransaccion, Registros.CodCuenta"
 FrmExportar.DtaRegistros.Refresh
ElseIf Buscado = True Then
 FrmExportar.DtaRegistros.RecordSource = "SELECT Registros.Control, Registros.IdRegistros, Registros.Fecha, Registros.NTransaccion, Registros.Fuente, Registros.CodCuenta, Registros.Descripcion, Registros.CodDepartamento, Registros.CodAcciones, Registros.ClaveProyecto, Registros.FacturaNumero, Registros.TipoMovimiento, Registros.RefCheque, Registros.FechaDescuento, Registros.FechaVencimiento, Registros.ImporteTransaccionDebito, Registros.ImporteTransaccionCredito, Registros.ImporteDescuento, Registros.ValorUnitario, Registros.TipoTransaccion, Registros.DebitoDolar, Registros.CreditoDolar From Registros WHERE (((Registros.Fecha) Between " & NumFecha1 & " And " & NumFecha2 & "))ORDER BY Registros.Fecha, Registros.NTransaccion, Registros.CodCuenta"
 FrmExportar.DtaRegistros.Refresh
End If

Unload Me

 Exit Sub
TipoErrs:
 MsgBox Err.Description
 
 Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs
Dim Maximo As Integer, i As Integer
picTV.BackColor = Me.BackColor
Frame1.BackColor = Me.BackColor
Ruta = App.Path + "\Enlace.jbh"
Me.Timer1.Enabled = True
Timer1.Interval = Tape.Speed
FrmExportar.DtaRegistros.DatabaseName = Ruta
FrmExportar.DtaRegistros.Connect = ConexioN
FrmExportar.DtaVerifica.Connect = ConexioN
Me.DtaConsulta.DatabaseName = Ruta
FrmExportar.DtaVerifica.DatabaseName = Ruta
Me.DtaConsulta.Connect = ConexioN
Me.DtaElimina.DatabaseName = Ruta
Me.DtaElimina.Connect = ConexioN
 Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub
Private Sub Timer1_Timer()
On Error GoTo TipoErrs
Dim intWidth As Integer
Dim intLeft As Integer      'Posición izquierda
Dim objImage As Control     'Control Image
Dim objImage1 As Control
Randomize
'Dim intLeft As Integer      'Posición izquierda
    'Dim objImage As Control     'Control Image
    Randomize   ' Inicializa el generador de números aleatorios.


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
    
 If Label0.Visible = True Then
   Label1.Visible = True
   Label0.Visible = False
   
 ElseIf Label1.Visible = True Then
    Label1.Visible = False
    Label2.Visible = True
 ElseIf Label2.Visible = True Then
    Label2.Visible = False
    Label3.Visible = True
ElseIf Label3.Visible = True Then
    Label3.Visible = False
    Label4.Visible = True
ElseIf Label4.Visible = True Then
    Label4.Visible = False
    Label5.Visible = True
  ElseIf Label5.Visible = True Then
    Label5.Visible = False
    Label6.Visible = True
  ElseIf Label6.Visible = True Then
    Label6.Visible = False
    Label7.Visible = True
  ElseIf Label7.Visible = True Then
    Label7.Visible = False
    Label8.Visible = True
  ElseIf Label8.Visible = True Then
    Label8.Visible = False
    Label9.Visible = True
  ElseIf Label9.Visible = True Then
    Label9.Visible = False
    Label10.Visible = True
  ElseIf Label10.Visible = True Then
    Label10.Visible = False
    Label11.Visible = True
  ElseIf Label11.Visible = True Then
    Label11.Visible = False
    Label12.Visible = True
  ElseIf Label12.Visible = True Then
    Label12.Visible = False
    Label13.Visible = True
  ElseIf Label13.Visible = True Then
    Label13.Visible = False
    Label14.Visible = True
  ElseIf Label14.Visible = True Then
    Label14.Visible = False
    Label15.Visible = True
  ElseIf Label15.Visible = True Then
    Label15.Visible = False
    Label0.Visible = True
    
 End If

' Borra la presentación
    picTV.Cls
    ' Muestra la nueva imagen en la nueva posición
    picTV.PaintPicture objImage.Picture, intLeft, 100, 500, 500
 Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub
