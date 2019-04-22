VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARCuotasCreditoUnif 
   Caption         =   "Imprimir cuotas de credito1"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "ARCuotasCreditoUnif.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARCuotasCreditoUnif.dsx":0442
End
Attribute VB_Name = "ARCuotasCreditoUnif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
'se ejecuta una sola vez
On Error GoTo merror

'si hay segundo vencimiento muestro los campos
If VG_APLICARSEGUNDOVENCIMIENTO Then
   FrameVencimiento2.Visible = True
   FrameVencimiento2Bis.Visible = True
End If

Exit Sub
merror:
tratarerrores "Error en reporte ARCuotasCreditos2-1"
End Sub
Private Sub Detail_Format()
'se ejecuta para cada factura
Dim SaldoCuota As Currency
Dim SaldoCuota1erVenc As Currency
Dim Importe1erVenc As Currency
Dim ImporteParcial As Currency
Dim Fecha As Date
Dim ImporteMora As Currency
Dim IvaMora As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim IvaACobrarDevuelto  As Currency
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

'por defecto los campos de actualizacion no se muestran
Call OcultarCamposActualizacion

'la mora y los recargos van juntos
FieldImporteVto1.Text = 0
FieldImporteVto1Bis.Text = 0
FieldImporteVto2.Text = 0
FieldImporteVto2Bis.Text = 0
FieldImporteFinal.Text = 0
FieldImporteFinalBis.Text = 0
FieldVencimientoFinal.Text = Date
FieldVencimientoFinalBis.Text = Date

Fecha = CDate(FieldFecha.Text)
'actualizo los importes si hubo cambios

With RDODataControl1.Resultset


     FieldImporteVto1.Text = CCur(.rdoColumns("importevencimiento1"))
     
     'saldo de credimaco
     
     'el importe 2 es el 1 mas recargo de 2 vto
     FieldImporteVto2.Text = CCur(.rdoColumns("importevencimiento2"))
     
     'muestro o oculto las leyendas de refinanciada o comodin
     Call Carteles
     
     'si no esta cobrada
     If IsNull(.rdoColumns("fechacobro")) Then
         LabelImporteFinal.Caption = "Imp.actualizado $:"
         LabelVencimientoFinal.Caption = "A la fecha:"
         LabelImporteFinalBis.Caption = "Imp.actualizado $:"
         LabelVencimientoFinalBis.Caption = "A la fecha:"
        
         SaldoCuota = .rdoColumns("SaldoCuota")
               
         If CDate(FieldFecha.Text) > CDate(.rdoColumns("fechavencimiento2")) Or _
            .rdoColumns("PagadoParcial") > 0 Then
            Call MostrarCamposActualizacion
         End If
    Else
         'la cuota esta cobrada
         LabelImporteFinal.Caption = "Imp.cobrado $:"
         LabelImporteFinalBis.Caption = "Imp.cobrado $:"
                       
         SaldoCuota = CCur(.rdoColumns("importecobrado"))
         Fecha = CDate(.rdoColumns("fechacobro"))
         Call MostrarCamposActualizacion
     End If
     
     'formateo la salida
     FieldImporteVto1.Text = Format(FieldImporteVto1.Text, "0.00")
     FieldImporteVto1Bis.Text = Format(FieldImporteVto1.Text, "0.00")
     FieldImporteVto2.Text = Format(FieldImporteVto2.Text, "0.00")
     FieldImporteVto2Bis.Text = Format(FieldImporteVto2.Text, "0.00")
     FieldImporteFinal.Text = Format(SaldoCuota, "0.00")
     FieldImporteFinalBis.Text = Format(SaldoCuota, "0.00")
     FieldVencimientoFinal.Text = CDate(Fecha)
     FieldVencimientoFinalBis.Text = CDate(Fecha)
     
End With
sql = "select count(*) as cantidad from cuotastemp"

Set rec = cnSQL.OpenResultset(sql)

If rec.EOF Then
    f_totaldoc.Text = "0"
Else
    f_totaldoc.Text = rec.rdoColumns("cantidad")
End If
Field16.Text = f_totaldoc.Text

Exit Sub
merror:
tratarerrores "Error en reporte ARCuotasCredito2-2"
End Sub
Private Sub MostrarCamposActualizacion()
'si hubo cambios mora,descuentos,recargos en cobradas o no cobradas
'muestra los detalles de actualizacion
On Error GoTo merror

LabelImporteFinal.Visible = True
LabelImporteFinalBis.Visible = True
LabelVencimientoFinal.Visible = True
LabelVencimientoFinalBis.Visible = True

FieldImporteFinal.Visible = True
FieldImporteFinalBis.Visible = True
FieldVencimientoFinal.Visible = True
FieldVencimientoFinalBis.Visible = True

Exit Sub
merror:
tratarerrores "Error en reporte ARCuotasCreditos2-3"
End Sub
Private Sub OcultarCamposActualizacion()
'si no hay actualizacion oculta los campos
On Error GoTo merror

LabelImporteFinal.Visible = False
LabelImporteFinalBis.Visible = False
LabelVencimientoFinal.Visible = False
LabelVencimientoFinalBis.Visible = False

FieldImporteFinal.Visible = False
FieldImporteFinalBis.Visible = False
FieldVencimientoFinal.Visible = False
FieldVencimientoFinalBis.Visible = False

Exit Sub
merror:
tratarerrores "Error en reporte ARCuotasCreditos2-4"
End Sub
Private Sub Carteles()
On Error GoTo merror

With RDODataControl1.Resultset
     'si esta refinanciada muestro el cartel
     If Not IsNull(.rdoColumns("fecharefinanciacion")) Then
        LabelRefinanciada.Visible = True
        LabelRefinanciadaBis.Visible = True
     Else
        LabelRefinanciada.Visible = False
        LabelRefinanciadaBis.Visible = False
     End If
     
     'si es comodin muestro el cartel
     If .rdoColumns("cuotacomodin") Then
        LabelComodin.Visible = True
        LabelComodinBis.Visible = True
     Else
        LabelComodin.Visible = False
        LabelComodinBis.Visible = False
     End If
End With

Exit Sub
merror:
tratarerrores "Error en reporte Carteles-ARCuotasCredito2-5"
End Sub
