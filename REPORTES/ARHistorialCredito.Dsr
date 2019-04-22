VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARHistorialCredito 
   Caption         =   "Historial de credito"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20370
   Icon            =   "ARHistorialCredito.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   16642
   SectionData     =   "ARHistorialCredito.dsx":0442
End
Attribute VB_Name = "ARHistorialCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'usado para totales
Dim ImporteTotal As Currency
Dim TotalCobrado As Currency
Dim IdCredito As Long
Private Sub Detail_Format()
Dim SaldoCuota As Currency
Dim ImporteCobrado As Currency
On Error GoTo merror

ImporteCobrado = 0
With RDODataControl1.Resultset
     
    FieldObs.Text = "Manual"
    If .rdoColumns("pagofacil") Then
       FieldObs.Text = "PagoFacil"
    End If
    
    If .rdoColumns("rapipago") Then
       FieldObs.Text = "RapiPago"
    End If
    
    TotalCobrado = CCur(TotalCobrado) + CCur(.rdoColumns("importecobrado"))
    
    IdCredito = .rdoColumns("idcredito")
End With

Exit Sub
merror:
tratarerrores "Error en ARResumenCreditoDetail"
End Sub
Private Sub PageFooter_Format()
FieldPagina.Text = Me.pageNumber
End Sub

Private Sub ReportFooter_Format()
FieldCobrado.Text = Format(TotalCobrado, "0.00")
FieldSaldoTotal.Text = Format(ObtenerSaldoCredito(IdCredito, Date), "0.00")
End Sub
