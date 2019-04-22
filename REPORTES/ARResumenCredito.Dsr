VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARResumenCredito 
   Caption         =   "Resumen de credito"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   Icon            =   "ARResumenCredito.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARResumenCredito.dsx":0442
End
Attribute VB_Name = "ARResumenCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'usado para totales
Dim ImporteTotal As Currency
Dim ImporteCobrado As Currency
Private Sub Detail_Format()
Dim SaldoCuota As Currency
On Error GoTo merror

ImporteCobrado = 0
With RDODataControl1.Resultset
     
    FieldObs.Text = ""

    If .rdoColumns("cuotacomodin") Then
       FieldObs.Text = "COMODIN"
    End If
    
    If .rdoColumns("exceptuada") Then
       FieldObs.Text = "EXCEPTUADA"
    End If
    
    SaldoCuota = ObtenerSaldoCuotaOKK(.rdoColumns("idcredito"), .rdoColumns("numcuota"), .rdoColumns("fechavencimiento1"), .rdoColumns("fechavencimiento2"), .rdoColumns("exceptuada"), FieldFecha.Text)
    FieldSaldo.Text = Format(SaldoCuota, "0.00")
    ImporteTotal = CCur(ImporteTotal) + CCur(SaldoCuota)
    FieldCuota.Text = CStr(.rdoColumns("numcuota")) & "/" & CStr(.rdoColumns("numcuotas"))

If CCur(ImporteCobrado) = 0 Then
   ImporteCobrado = ObtenerCobrosCredito(.rdoColumns("idcredito"))
End If

End With

Exit Sub
merror:
tratarerrores "Error en ARResumenCreditoDetail"
End Sub
Private Sub PageFooter_Format()
FieldPagina.Text = Me.pageNumber
End Sub

Private Sub ReportFooter_Format()
FieldCobrado.Text = Format(ImporteCobrado, "0.00")
FieldSaldoTotal.Text = Format(ImporteTotal, "0.00")
End Sub
