VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepFacturas 
   Caption         =   "Reporte de Facturas"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepFacturas.dsx":0000
End
Attribute VB_Name = "ArepFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_ReportStart()
Dim ConexionFacturacion As String, FacturaNumero As String

              FacturaNumero = NumeroFact

                     If MDIPrimero.AdoConfiguracion.Recordset.RecordCount > 0 Then
                         If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion) Then
                            ConexionFacturacion = MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion
                         Else
                            ConexionFacturacion = ""
                         End If
                     End If


                SQL = "SELECT  * FROM  Facturas INNER JOIN Bodegas ON Facturas.Cod_Bodega = Bodegas.Cod_Bodega " & _
                      "WHERE (Facturas.Numero_Factura = '" & FacturaNumero & "') AND (Facturas.Tipo_Factura = 'Factura')"
              
                MDIPrimero.AdoConsultaFacturacion.ConnectionString = ConexionFacturacion
                MDIPrimero.AdoConsultaFacturacion.RecordSource = SQL
                MDIPrimero.AdoConsultaFacturacion.Refresh
                If Not MDIPrimero.AdoConsultaFacturacion.Recordset.EOF Then
                  Me.LblSubTotal.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("SubTotal"), "##,##0.00")
                  Me.LblIVA.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("IVA"), "##,##0.00")
                  Me.LblTotal.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("SubTotal") + MDIPrimero.AdoConsultaFacturacion.Recordset("IVA"), "##,##0.00")
                  Me.LblBodega.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Cod_Bodega") & " " & MDIPrimero.AdoConsultaFacturacion.Recordset("Nombre_Bodega")
'                  Me.LblObservaciones.Caption = MDIPimero.AdoConsultaFacturacion.Recordset("Observaciones")
                  Me.LblCodigoCliente.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Cod_Cliente")
                  Me.LblNombreCliente.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Nombre_Cliente")
                  Me.LblNuestraRef.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("MonedaFactura")
                End If
                
                   Me.Logo.Picture = LoadPicture(RutaLogo)
                
                  Me.LblEmpresa.Caption = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
                  Me.LblEmpresa1.Caption = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
                  Me.LblEmpresa2.Caption = "RUC " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
End Sub

