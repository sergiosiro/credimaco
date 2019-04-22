VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARCartaRec 
   Caption         =   "Carta Reclamo"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "ARCartaRec.dsx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARCartaRec.dsx":0442
End
Attribute VB_Name = "ARCartaRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mTotal As Currency

Private Sub Detail_Format()
Dim SaldoCuota As Currency
On Error GoTo merror

'va acumulando el total actualizado por cada cliente
With RDODataControl1.Resultset
     FieldNumCuota.Text = CStr(.rdoColumns("numcuota")) & "/" & CStr(.rdoColumns("numcuotas"))
     'obtengo el total adeudado de un cliente
     SaldoCuota = ObtenerSaldoCuotaOKK(.rdoColumns("idcredito"), .rdoColumns("numcuota"), .rdoColumns("fechavencimiento1"), .rdoColumns("fechavencimiento2"), .rdoColumns("exceptuada"), FieldFecha.Text)
     FieldImporteActualizado.Text = Format(SaldoCuota, "0.00")
     mTotal = CCur(mTotal) + CCur(SaldoCuota)
End With
      
Exit Sub
merror:
tratarerrores "Error en reporte CartaReclamo1"
End Sub

Private Sub GroupHeader1_BeforePrint()
'actualizo el total adeudado por cada cliente
On Error GoTo merror

rtf.ReplaceField "deuda", Format(mTotal, "0.00")
mTotal = 0

Exit Sub
merror:
tratarerrores "Error en reporte CartaReclamo2"
End Sub
Private Sub GroupHeader1_Format()
'cargo los campos personalizados del rich text
On Error GoTo merror

With RDODataControl1.Resultset
     rtf.ReplaceField "titular", UCase(IIf(IsNull(.rdoColumns("cliente")), "", .rdoColumns("cliente")))
     rtf.ReplaceField "domicilio", IIf(IsNull(.rdoColumns("domicilio")), "", .rdoColumns("domicilio"))
     rtf.ReplaceField "cp", IIf(IsNull(.rdoColumns("codigopostal")), "", .rdoColumns("codigopostal"))
     rtf.ReplaceField "localidad", IIf(IsNull(.rdoColumns("localidad")), "", .rdoColumns("localidad"))
     rtf.ReplaceField "provincia", IIf(IsNull(.rdoColumns("provincia")), "", .rdoColumns("provincia"))
End With

Exit Sub
merror:
tratarerrores "Error en reporte CartaReclamo3"
End Sub
Private Sub PageFooter_Format()
Field2.Text = Me.pageNumber
End Sub

