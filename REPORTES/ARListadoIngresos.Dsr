VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARListadoIngresos 
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   Icon            =   "ARListadoIngresos.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARListadoIngresos.dsx":0442
End
Attribute VB_Name = "ARListadoIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalGral As Currency
Dim TotalGral2 As Currency
Dim TotalIvaInteresCobrado As Currency
Dim TotalIvaSegurosCobrado As Currency
Dim TotalIvaOtorGastosCobrado As Currency
Dim TotalMoraCobrada As Currency
Dim TotalIvaMoraCobrada As Currency
Private Sub Detail_Format()
On Error GoTo merror

With RDODataControl1.Resultset
     FieldCuota.Text = CStr(.rdoColumns("numcuota")) & "/" & CStr(.rdoColumns("numcuotas"))
     TotalGral = CCur(TotalGral) + CCur(.rdoColumns("importecobrado"))
     TotalIvaInteresCobrado = TotalIvaInteresCobrado + CCur(.rdoColumns("ivainterescobrado"))
     TotalIvaSegurosCobrado = TotalIvaSegurosCobrado + CCur(.rdoColumns("ivaseguroscobrado"))
     TotalIvaOtorGastosCobrado = TotalIvaOtorGastosCobrado + CCur(.rdoColumns("ivaotorgastoscobrado"))
     TotalMoraCobrada = CCur(TotalMoraCobrada) + CCur(.rdoColumns("moracobrada"))
     TotalIvaMoraCobrada = CCur(TotalIvaMoraCobrada) + CCur(.rdoColumns("ivamoracobrada"))
End With

Exit Sub
merror:
tratarerrores "Error en ARListadoIngresos-1-1"
End Sub
Private Sub PageFooter_Format()
Field26.Text = Me.pageNumber
End Sub

Private Sub ReportFooter_Format()
'al finalizar cada grupo reinicia las variables
FieldTotalGral.Text = Format(TotalGral, "0.00")
TotalGral2 = CCur(TotalGral) - CCur(TotalIvaInteresCobrado) - CCur(TotalIvaSegurosCobrado) - CCur(TotalIvaOtorGastosCobrado) - CCur(TotalMoraCobrada) - CCur(TotalIvaMoraCobrada)
FieldTotal2.Text = Format(TotalGral2, "0.00")
TotalGral = 0
TotalGral2 = 0
End Sub
