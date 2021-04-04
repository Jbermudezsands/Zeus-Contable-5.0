VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepSptDepositos 
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepSptDepositos.dsx":0000
End
Attribute VB_Name = "ArepSptDepositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Detail_Format()
 TotalConDebito = Val(Me.Field5.Text) + TotalConDebito
 TotalConCredito = Val(Me.Field6.Text) + TotalConCredito
End Sub

Private Sub PageHeader_Format()
TotalConDebito = 0
TotalConCredito = 0
End Sub
