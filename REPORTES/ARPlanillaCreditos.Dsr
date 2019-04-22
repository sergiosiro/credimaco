VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARPlanillaCreditos 
   Caption         =   "Resumen de creditos"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   Icon            =   "ARPlanillaCreditos.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARPlanillaCreditos.dsx":0442
End
Attribute VB_Name = "ARPlanillaCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Total As Currency

Private Sub Detail_Format()
Dim SaldoCredito As Currency
Dim SaldoCapital As Currency
Dim Fecha As Date
Dim CondicionZ As String
On Error GoTo merror

With RDODataControl1.Resultset
     
     SaldoCredito = ObtenerSaldoCredito(.rdoColumns("idcredito"), FieldFechaActual.Text)
     Total = CCur(Total) + CCur(SaldoCredito)
     FieldSaldoCredito.Text = Format(SaldoCredito, "0.00")
     FieldCuotasCobradas.Text = ObtenerCuotasCobradas(.rdoColumns("idcredito"))
     FieldCuotasVencidas.Text = ObtenerCuotasVencidas(.rdoColumns("idcredito"), FieldFechaActual.Text)
     FieldCuotasPendientes.Text = ObtenerCuotasPendientes(.rdoColumns("idcredito"), FieldFechaActual.Text)
     Fecha = ObtenerUltimaFechaCobro(.rdoColumns("idcredito"))
     If CDate(Fecha) = Date + 1 Then
        FieldUltimaFechaCobro.Text = ""
     Else
        FieldUltimaFechaCobro.Text = Fecha
     End If
End With

Exit Sub
merror:
tratarerrores "Error en reporte ARPlanillaCreditos-1"
End Sub
Private Sub PageFooter_Format()
Field26.Text = Me.pageNumber
End Sub
Private Sub ReportFooter_Format()
FieldTotal.Text = Format(Total, "0.00")
End Sub
