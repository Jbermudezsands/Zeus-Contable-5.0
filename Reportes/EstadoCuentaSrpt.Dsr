VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} EstadoCuentaSrpt 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20955
   _ExtentY        =   19420
   SectionData     =   "EstadoCuentaSrpt.dsx":0000
End
Attribute VB_Name = "EstadoCuentaSrpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FacturaNo As String
Private Sub ActiveReport_FetchData(EOF As Boolean)
    If Not EOF Then
    'Gets the current records SupplierID
      FacturaNo = Me.AdoEstadoCuenta.Recordset.Fields("FacturaNo")
    End If
End Sub
 Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
 Dim FacturaNo As String, FechaFactura As Date, SQL As String
 Dim rpt As Object, fPreview As New FrmPreview
'Check to see if an email link or web page has been selected

FacturaNo = Link
If InStr(1, Link, "htm", vbTextCompare) = 0 And InStr(1, Link, "mailto", vbTextCompare) = 0 Then


    FacturaNo = Link
    
    FechaFactura = Me.Field6.Text

    
    
    SQL = "SELECT  Productos.*, Detalle_Facturas.* FROM Detalle_Facturas INNER JOIN Productos ON Detalle_Facturas.Cod_Producto = Productos.Cod_Productos  " & _
          "WHERE  (Detalle_Facturas.Numero_Factura = '" & FacturaNo & "') AND (Detalle_Facturas.Tipo_Factura = 'Factura') AND (Detalle_Facturas.Fecha_Factura = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-mm-dd") & "', 102))"
    
     
      ArepFacturas.DataControl1.ConnectionString = ConexionFacturacion
      ArepFacturas.DataControl1.Source = SQL
    
       ArepFacturas.Logo.Picture = LoadPicture(RutaLogo)
    
      ArepFacturas.LblEmpresa.Caption = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
      ArepFacturas.LblEmpresa1.Caption = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
      ArepFacturas.LblEmpresa2.Caption = "RUC " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
      
      ArepFacturas.Show
      
   
'         Set rpt = New ArepFacturas
'         rpt.DataControl1.ConnectionString = ConexionFacturacion
'         rpt.DataControl1.Source = SQL
'         fPreview.RunReport rpt
'         fPreview.Show 1


End If
End Sub



Private Sub Detail_Format()
Dim FechaVence As String

FechaVence = Format(Me.Field7.Text, "dd/mm/yyyy")
If FechaVence = "01/01/1900" Then
  Me.Field7.Text = Me.Field6.Text
End If

'Me.Field5.Hyperlink = FacturaNo
End Sub
