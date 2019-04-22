VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARClientesLista 
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "ARClientesLista.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARClientesLista.dsx":0442
End
Attribute VB_Name = "ARClientesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ImporteTotal As Currency

Private Sub Detail_Format()
Dim Saldo As Currency
On Error GoTo merror

With RDODataControl1.Resultset
     Saldo = ObtenerSaldoCliente(.rdoColumns("idcliente"))
     ImporteTotal = CCur(ImporteTotal) + CCur(Saldo)
End With
FieldSaldo.Text = Format(Saldo, "0.00")

Exit Sub
merror:
tratarerrores "Error en reporte ARClientesLista2"
End Sub

Private Sub PageFooter_Format()
FieldPagina.Text = Me.pageNumber
End Sub

Private Sub ReportFooter_Format()
FieldTotal.Text = Format(ImporteTotal, "0.00")
End Sub
