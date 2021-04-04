VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Begin VB.Form frmPane 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
      Align           =   3  'Align Left
      Height          =   6405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      _Version        =   786432
      _ExtentX        =   2646
      _ExtentY        =   11298
      _StockProps     =   64
      Animation       =   1
      ItemLayout      =   3
      HotTrackStyle   =   1
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5040
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   45
      ImageHeight     =   45
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   37
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":088C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":1226
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":1B84
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":2388
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":2D07
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":3545
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":3E30
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":4796
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":4F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":56AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":5E63
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":65FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":6FF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":7921
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":8263
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":142B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":14CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":15519
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":15CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":1666F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":16FC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":178B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":180AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":18A8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":48ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":54B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":5550A
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":55F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":61F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":63AC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":65616
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":67168
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":68CBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":74D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":7685E
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPane.frx":783B0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    CreateTaskPanel
  End Sub
Private Sub wndTaskPanel_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
    Dim Directorio As String
    Dim AÑO1 As String, AÑO2 As String, AÑO3 As String
    
    Select Case Item.Caption
        Case "Consolidacion"
           FrmConsolidacion.Show
        Case "Cambiar Fecha"
           FrmFecha.Show
        Case "Departamento"
           FrmGrupo.Show
        Case "Reporte Diario"
           FrmReporteComprobantes.Show
        Case "Compañia"
           FrmCompañia.Show 1
        Case "Calcular Depreciacion"
           FrmCalcularDepreciacion.Show 1
        Case "Cuentas"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cuentas'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            

            FrmCuentas.Show
            

              
            
        Case "Grupo de Cuentas"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Grupo Cuentas'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If

            FrmGrupo.Show
            
        Case "Cuentas de Mayor"

            FrmCuentasMayor.Show
            
        Case "Auxiliar de Cuentas"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Contratistas'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If

            FrmAuxiliarCuentas.Show 1
        Case "Niveles de Acceso"
              FrmEditarNiveles.Show 1
        Case "Contratistas"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Contratistas'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmContactos.Show
        Case "Periodo"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Periodos'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmPeriodos.Show 1
        Case "Empleados"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Empleados'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmEmpleados.Show
        Case "Cheques"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Cheques'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmCheque.Show
        Case "Presupuesto"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Presupuesto'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmPresupuesto.Show 1
        Case "Activo Fijo"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Activo Fijo'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmActivoFijo.Show
        Case "Transacciones"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Transacciones'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmTransacciones.Show
        Case "Tasas de Cambio"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tasa Cambio'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            frmTasa2.Show 1
        Case "Usuarios"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Usuarios'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmUsuarios.Show
        Case "Calcular Depreciación"
            MDIPrimero.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Depreciacion'))"
            MDIPrimero.DtaNacceso.Refresh
            If MDIPrimero.DtaNacceso.Recordset.EOF And CodigoUsuario <> 0 Then
                Exit Sub
            End If
            FrmCalcularDepreciacion.Show 1
        Case "Prorrateo"
            FrmProrrateo.Show
        Case "Calculadora"
          Directorio = App.Path & "\Calc.exe"
          Directorio = Shell(Directorio)
          MDIPrimero.MousePointer = 0
        Case "Informacion de Usuarios"
          FrmInforme.Show 1
          
    Case "Reportes Generales"
        QUIEN = "ReporteGenerales"
      FrmReportes.Show 1
    
    Case "Reportes de Movimientos"
      QUIEN = "ReporteMovimientos"
      FrmReportes.Show 1
    
    Case "Reportes de Bancos"
     QUIEN = "ReporteBancos"
      FrmReportes.Show 1
      
     Case "Estados Financieros"
     QUIEN = "EstadosFinancieros"
      FrmReportes.CmbNivel.AddItem ("0")
      FrmReportes.CmbNivel.AddItem ("1")
      FrmReportes.CmbNivel.AddItem ("2")
      FrmReportes.CmbNivel.AddItem ("3")
      FrmReportes.CmbNivel.AddItem ("4")
      FrmReportes.CmbNivel.AddItem ("5")
      FrmReportes.CmbNivel.AddItem ("6")
      FrmReportes.CmbNivel.AddItem ("7")
      FrmReportes.CmbNivel.AddItem ("8")
      FrmReportes.CmbNivel.AddItem ("9")
      FrmReportes.CmbNivel.AddItem ("10")
      FrmReportes.CmbNivel.AddItem ("11")
      FrmReportes.CmbNivel.AddItem ("12")
      FrmReportes.CmbNivel.AddItem ("13")
      FrmReportes.CmbNivel.AddItem ("14")
      FrmReportes.CmbNivel.AddItem ("15")
      FrmReportes.CmbNivel.AddItem ("16")
      FrmReportes.CmbNivel.AddItem ("17")
      FrmReportes.CmbNivel.AddItem ("18")
      FrmReportes.CmbNivel.AddItem ("19")
      FrmReportes.CmbNivel.AddItem ("20")
      FrmReportes.Label3.Visible = True
      FrmReportes.CmbMoneda.Visible = True
      FrmReportes.Frame1.Visible = False
      FrmReportes.Frame4.Visible = True
      FrmReportes.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      FrmReportes.DtaConsulta.Refresh
      Do While Not FrmReportes.DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
        FrmReportes.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
        FrmReportes.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
         FrmReportes.Option6.Caption = AÑO3
       End If
        
        FrmReportes.DtaConsulta.Recordset.MoveNext
      Loop
      FrmReportes.Show 1
    
    Case "Analisis Financieros"
     QUIEN = "Analisis Financieros"
      FrmReportes.CmbMoneda.Visible = True
      FrmReportes.Frame1.Visible = False
      FrmReportes.Frame4.Visible = True
      FrmReportes.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      FrmReportes.DtaConsulta.Refresh
      Do While Not FrmReportes.DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
        FrmReportes.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
        FrmReportes.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
         FrmReportes.Option6.Caption = AÑO3
       End If
        
        FrmReportes.DtaConsulta.Recordset.MoveNext
      Loop
     FrmReportes.Show 1
      
   Case "Configuracion de Cheques"
           FrmConfiguraCheque.Show 1
   Case "Importar Transacciones"
          FrmImporta.Show 1
   Case "Importar Cuentas"
          FrmImportarCuentas.Show 1
   Case "Auditor"
          FrmAuditor.Show 1
   Case "Respaldar"
           FrmRespaldar.Show vbModal
   Case "Configuracion"
            FrmConfiguracion.Show
   Case "Exportar Transacciones"
          FrmExportacion.Show
   Case "Importar Tasas"
          FrmImportarTasa.Show
    End Select
End Sub


Private Sub wndTaskPanel_GroupExpanding(ByVal Group As XtremeTaskPanel.ITaskPanelGroup, ByVal Expanding As Boolean, Cancel As Boolean)
 If Expanding = True Then
  Select Case Group.Caption
    Case "Cuentas"
'              wndTaskPanel.Groups(1).Expanded = True
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
'              wndTaskPanel.Groups(6).Expanded = False
    Case "Catalogos"
              wndTaskPanel.Groups(1).Expanded = False
'              wndTaskPanel.Groups(2).Expanded = True
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
'              wndTaskPanel.Groups(6).Expanded = False
    Case "Procesos"
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
'              wndTaskPanel.Groups(3).Expanded = True
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
'              wndTaskPanel.Groups(6).Expanded = False
    
    Case "Opciones"
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
'              wndTaskPanel.Groups(4).Expanded = True
              wndTaskPanel.Groups(5).Expanded = False
'              wndTaskPanel.Groups(6).Expanded = False
    
    Case "Reportes"
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
'              wndTaskPanel.Groups(5).Expanded = True
'              wndTaskPanel.Groups(6).Expanded = False
    
    Case "Ayuda"
    
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
'              wndTaskPanel.Groups(6).Expanded = True
    
  
  End Select
 End If

End Sub


Sub CreateTaskPanel()


    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    Set Group = wndTaskPanel.Groups.Add(100, "Cuentas")
    Group.Tooltip = "cuentas del sistema Contable"
    Group.Special = True
    Group.Items.Add 1, "Cuentas", xtpTaskItemTypeLink, 2
    Group.Items.Add 2, "Grupo de Cuentas", xtpTaskItemTypeLink, 4
    Group.Items.Add 3, "Cuentas de Mayor", xtpTaskItemTypeLink, 3
    Group.Items.Add 4, "Auxiliar de Cuentas", xtpTaskItemTypeLink, 1
    Group.Items.Add 5, "Reporte Diario", xtpTaskItemTypeLink, 5
    
    Set Group = wndTaskPanel.Groups.Add(100, "Catalogos")
    Group.Tooltip = "Catalogo del sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 6, "Empleados", xtpTaskItemTypeLink, 6
    Group.Items.Add 7, "Contratistas", xtpTaskItemTypeLink, 7
    Group.Items.Add 8, "Periodo", xtpTaskItemTypeLink, 8
    Group.Items.Add 9, "Activo Fijo", xtpTaskItemTypeLink, 9
    Group.Items.Add 10, "Usuarios", xtpTaskItemTypeLink, 10
    Group.Items.Add 11, "Tasas de Cambio", xtpTaskItemTypeLink, 11
    Group.Items.Add 12, "Departamento", xtpTaskItemTypeLink, 12
    Group.Items.Add 13, "Niveles de Acceso", xtpTaskItemTypeLink, 13
    
    
    Set Group = wndTaskPanel.Groups.Add(100, "Procesos")
    Group.Tooltip = "Procesos del Sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 13, "Transacciones", xtpTaskItemTypeLink, 14
    Group.Items.Add 13, "Cheques", xtpTaskItemTypeLink, 15
    Group.Items.Add 13, "Calcular Depreciacion", xtpTaskItemTypeLink, 16
    Group.Items.Add 13, "Presupuesto", xtpTaskItemTypeLink, 17
    Group.Items.Add 13, "Prorrateo", xtpTaskItemTypeLink, 32
    
    Set Group = wndTaskPanel.Groups.Add(100, "Opciones")
    Group.Tooltip = "Procesos del Sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 13, "Calculadora", xtpTaskItemTypeLink, 18
    Group.Items.Add 13, "Informacion de Usuarios", xtpTaskItemTypeLink, 19
    Group.Items.Add 13, "Configuracion", xtpTaskItemTypeLink, 28
    Group.Items.Add 13, "Respaldar", xtpTaskItemTypeLink, 29
    Group.Items.Add 13, "Compañia", xtpTaskItemTypeLink, 33
    Group.Items.Add 13, "Cambiar Fecha", xtpTaskItemTypeLink, 35
    Group.Items.Add 13, "Consolidacion", xtpTaskItemTypeLink, 35
   
    
'    Set Group = wndTaskPanel.Groups.Add(100, "Reportes")
'    Group.ToolTip = "Procesos del Sistema Contable"
'    Group.Special = True
'    Group.Expanded = False
'    Group.Items.Add 13, "Reportes Generales", xtpTaskItemTypeLink, 20
'    Group.Items.Add 13, "Reportes de Movimientos", xtpTaskItemTypeLink, 21
'    Group.Items.Add 13, "Reportes de Bancos", xtpTaskItemTypeLink, 22
'    Group.Items.Add 13, "Estados Financieros", xtpTaskItemTypeLink, 23
'    Group.Items.Add 13, "Analisis Financieros", xtpTaskItemTypeLink, 31
    
    Set Group = wndTaskPanel.Groups.Add(100, "Ayuda")
    Group.Tooltip = "Procesos del Sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 13, "Configuracion de Cheques", xtpTaskItemTypeLink, 24
    Group.Items.Add 13, "Importar Transacciones", xtpTaskItemTypeLink, 37
    Group.Items.Add 13, "Importar Cuentas", xtpTaskItemTypeLink, 26
    Group.Items.Add 13, "Importar Tasas", xtpTaskItemTypeLink, 30
    Group.Items.Add 13, "Exportar Transacciones", xtpTaskItemTypeLink, 36
    Group.Items.Add 13, "Auditor", xtpTaskItemTypeLink, 27
    
     
    wndTaskPanel.SetImageList Me.ImageList2
End Sub

Function CreateToolboxGroup(Caption As String) As TaskPanelGroup
'    Dim Folder As TaskPanelGroup, Pointer As TaskPanelGroupItem
'
'    Set Folder = wndToolBox.Groups.Add(0, Caption)
'    Folder.IconIndex = 100
'
'    Set CreateToolboxGroup = Folder
End Function

Function CreateToolboxItem(Group As TaskPanelGroup, Caption As String, IconIndex As Long) As TaskPanelGroup
     Group.Items.Add IconIndex + 1, Caption, xtpTaskItemTypeLink, IconIndex
   
End Function






Private Sub Form_Resize()
    On Error Resume Next
    Me.wndTaskPanel.Move Me.ScaleLeft + 1, Me.ScaleTop + 1, Me.ScaleWidth - 2, Me.ScaleHeight - 2

    
End Sub

