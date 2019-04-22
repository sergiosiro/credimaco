VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARCuotasCredito2 
   Caption         =   "Imprimir cuotas de credito1"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "ARCuotasCredito2.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARCuotasCredito2.dsx":0442
End
Attribute VB_Name = "ARCuotasCredito2"
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

     FieldImporteVto1.Text = CCur(.rdoColumns("importetotal"))
     
     ImporteParcial = ObtenerImporteParcialX(.rdoColumns("idcredito"), .rdoColumns("numcuota"))
     
     'saldo de credimaco
     SaldoCuota = ObtenerSaldoCuotaX(.rdoColumns("idcredito"), .rdoColumns("numcuota"), CDate(FieldFecha.Text), SaldoCuota1erVenc)
     Importe1erVenc = ObtenerImporte1erVenc(.rdoColumns("idcredito"), .rdoColumns("numcuota"))
     
     'el importe 2 es el 1 mas recargo de 2 vto
     FieldImporteVto2.Text = CCur(FieldImporteVto1.Text) + CCur(.rdoColumns("importerecargovencimiento2"))
     
     'muestro o oculto las leyendas de refinanciada o comodin
     Call Carteles
     
     'si no esta cobrada
     If IsNull(.rdoColumns("fechacobro")) Then
         LabelImporteFinal.Caption = "Imp.actualizado $:"
         LabelVencimientoFinal.Caption = "A la fecha:"
         LabelImporteFinalBis.Caption = "Imp.actualizado $:"
         LabelVencimientoFinalBis.Caption = "A la fecha:"
           
         'si no esta refinanciada y no es comodin
         If IsNull(.rdoColumns("fecharefinanciacion")) And Not (.rdoColumns("cuotacomodin")) Then
            'si estoy en mora actualizo
            If CDate(FieldFecha.Text) > CDate(.rdoColumns("fechavencimiento2")) Then
               'calculo la mora de forma habitual
               'puedo pasarle el campo [exceptuada]
               ImporteMora = CalculoMoraPendiente(.rdoColumns("idcredito"), .rdoColumns("numcuota"), .rdoColumns("exceptuada"), CCur(FieldImporteVto1.Text), .rdoColumns("fechavencimiento1"), CDate(FieldFecha.Text), IvaACobrarDevuelto)
               '''''''********ImporteMora = CalcularInteresMoraZZ(.rdoColumns("exceptuada"), SaldoCalculoMora, FechaCalculoMora, CDate(FieldFecha.Text))
               IvaMora = 0
               If VG_APLICARIMPUESTOS Then
                  If VG_IMPUESTOSCREDIMACO Then
                     IvaMora = IvaACobrarDevuelto
                  End If
               End If
               '''''''********SoloMoraCobrada = ObtenerMoraCobrada(.rdoColumns("idcredito"), .rdoColumns("numcuota"))
               '''''''********SoloIvaMoraCobrada = ObtenerIvaMoraCobrada(.rdoColumns("idcredito"), .rdoColumns("numcuota"))
               '''''''********If CCur(ImporteMora) <= CCur(SoloMoraCobrada) Then
               '''''''********   ImporteMora = 0
               '''''''********Else
                  'si es mayor la mora es solo la diferencia
               '''''''********   ImporteMora = CCur(ImporteMora) - CCur(SoloMoraCobrada)
               '''''''********End If
               '''''''********If CCur(IvaMora) <= CCur(SoloIvaMoraCobrada) Then
               '''''''********   IvaMora = 0
               '''''''********Else
                  'si es mayor la mora es solo la diferencia
               '''''''********   IvaMora = CCur(IvaMora) - CCur(SoloIvaMoraCobrada)
               '''''''********End If
               SaldoCuota = CCur(SaldoCuota1erVenc) + CCur(ImporteMora) + CCur(IvaMora)
               
               Call MostrarCamposActualizacion
            End If
            
            If CCur(ImporteParcial) > 0 Then
               'debo actualiza los importes
               Call MostrarCamposActualizacion
            End If
         
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
