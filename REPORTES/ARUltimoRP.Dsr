VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARUltimoRP 
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   Icon            =   "ARUltimoRP.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARUltimoRP.dsx":0442
End
Attribute VB_Name = "ARUltimoRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ImporteTotal As Currency
Private Sub Detail_Format()
On Error GoTo merror

With RDODataControl1.Resultset
     ImporteTotal = CCur(ImporteTotal) + CCur(.rdoColumns("importecobro"))
End With

Exit Sub

merror:
tratarerrores "Error en ARUltimaImportacionRapipago"
End Sub

Private Sub PageFooter_Format()
FieldPagina.Text = Me.pageNumber
End Sub

Private Sub ReportFooter_Format()
FieldTotal.Text = Format(ImporteTotal, "0.00")
End Sub
