VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARCuotasCredito3 
   Caption         =   "Imprimir cuotas de credito"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "ARCuotasCredito3.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARCuotasCredito3.dsx":0442
End
Attribute VB_Name = "ARCuotasCredito3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
'se ejecuta una sola vez
On Error GoTo merror

If VG_APLICARSEGUNDOVENCIMIENTO Then
   FrameVencimiento2.Visible = True
   FrameVencimiento2Bis.Visible = True
   FrameVencimiento2Tri.Visible = True
End If

Exit Sub
merror:
tratarerrores "Error en reporte ARCuotasCreditos3-1"
End Sub
Private Sub Detail_Format()
'se ejecuta para cada factura
Dim SaldoCuota As Currency
Dim ImporteMora As Currency
Dim IvaMora As Currency
Dim ImporteParcial As Currency
Dim Fecha As Date
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim SaldoCuota1erVenc As Currency
Dim IvaACobrarDevuelto  As Currency
On Error GoTo merror

'por defecto los campos de actualizacion no se muestran
Call OcultarCamposActualizacion

ImporteMora = 0

'por defecto no hay mora
FieldImporteVto1.Text = 0
FieldImporteVto1Bis.Text = 0
FieldImporteVto1Tri.Text = 0
FieldImporteVto2.Text = 0
FieldImporteVto2Bis.Text = 0
FieldImporteVto2Tri.Text = 0
FieldImporteFinal.Text = 0
FieldImporteFinalBis.Text = 0
FieldImporteFinalTri.Text = 0
FieldVencimientoFinal.Text = Date
FieldVencimientoFinalBis.Text = Date
FieldVencimientoFinalTri.Text = Date

Fecha = FieldFecha.Text

'actualizo los importes si hubo cambios
With RDODataControl1.Resultset
     FieldImporteVto1.Text = CCur(.rdoColumns("importetotal"))
     
     ImporteParcial = ObtenerImporteParcialX(.rdoColumns("idcredito"), .rdoColumns("numcuota"))
   
     'saldo de credimaco
     SaldoCuota = ObtenerSaldoCuotaX(.rdoColumns("idcredito"), .rdoColumns("numcuota"), CDate(FieldFecha.Text), SaldoCuota1erVenc)
     
     'por defecto el importe 2 es el 1 mas recargo de 2 vto
     FieldImporteVto2.Text = CCur(FieldImporteVto1.Text) + CCur(.rdoColumns("importerecargovencimiento2"))
     
     Call Carteles
     
     'si no esta cobrada
     If IsNull(.rdoColumns("fechacobro")) Then
        LabelImporteFinal.Caption = "Imp.actualizado:"
        LabelImporteFinalBis.Caption = "Imp.actualizado:"
        LabelImporteFinalTri.Caption = "Imp.actualizado:"
        LabelVencimientoFinal.Caption = "A la fecha:"
        LabelVencimientoFinalBis.Caption = "A la fecha:"
        LabelVencimientoFinalTri.Caption = "A la fecha:"

        'si no esta refinanciada y no es comodin
        If IsNull(.rdoColumns("fecharefinanciacion")) And Not (.rdoColumns("cuotacomodin")) Then
           'si estoy en mora actualizo (vale si aplico o no 2 vto)
           If CDate(FieldFecha.Text) > CDate(.rdoColumns("fechavencimiento2")) Then
              'calculo la mora de forma habitual
              'puedo pasarle el campo [exceptuada]
              ImporteMora = CalculoMoraPendiente(.rdoColumns("idcredito"), .rdoColumns("numcuota"), .rdoColumns("exceptuada"), SaldoCuota, .rdoColumns("fechavencimiento1"), CDate(FieldFecha.Text), IvaACobrarDevuelto)
              '''''''********ImporteMora = CalcularInteresMoraZZ(.rdoColumns("exceptuada"), SaldoCuota, .rdoColumns("fechavencimiento2"), CDate(FieldFecha.Text))
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
        'esta cobrada
        LabelImporteFinal.Caption = "Imp.cobrado:"
        LabelImporteFinalBis.Caption = "Imp.cobrado:"
        LabelImporteFinalTri.Caption = "Imp.cobrado:"
        LabelVencimientoFinal.Caption = "Fecha de cobro:"
        LabelVencimientoFinalBis.Caption = "Fecha de cobro:"
        LabelVencimientoFinalTri.Caption = "Fecha de cobro:"
        
        SaldoCuota = CCur(.rdoColumns("importecobrado"))
        Fecha = CDate(.rdoColumns("fechacobro"))
        Call MostrarCamposActualizacion
     End If
     
     'formateo la salida
     FieldImporteVto1.Text = Format(FieldImporteVto1.Text, "0.00")
     FieldImporteVto1Bis.Text = Format(FieldImporteVto1.Text, "0.00")
     FieldImporteVto1Tri.Text = Format(FieldImporteVto1.Text, "0.00")
     
     FieldImporteVto2.Text = Format(FieldImporteVto2.Text, "0.00")
     FieldImporteVto2Bis.Text = Format(FieldImporteVto2.Text, "0.00")
     FieldImporteVto2Tri.Text = Format(FieldImporteVto2.Text, "0.00")
            
     FieldImporteFinal.Text = Format(SaldoCuota, "0.00")
     FieldImporteFinalBis.Text = Format(SaldoCuota, "0.00")
     FieldImporteFinalTri.Text = Format(SaldoCuota, "0.00")
     
     FieldVencimientoFinal.Text = CDate(Fecha)
     FieldVencimientoFinalBis.Text = CDate(Fecha)
     FieldVencimientoFinalTri.Text = CDate(Fecha)
     
End With

Exit Sub
merror:
tratarerrores "Error en reporte ARCuotasCreditos2-2"
End Sub
Private Sub MostrarCamposActualizacion()
'si hubo cambios mora,descuentos,recargos en cobradas o no cobradas
'muestra los detalles de actualizacion
On Error GoTo merror

LabelPesos1.Visible = True
LabelPesos2.Visible = True
LabelPesos3.Visible = True

LabelImporteFinal.Visible = True
LabelImporteFinalBis.Visible = True
LabelImporteFinalTri.Visible = True

LabelVencimientoFinal.Visible = True
LabelVencimientoFinalBis.Visible = True
LabelVencimientoFinalTri.Visible = True

FieldImporteFinal.Visible = True
FieldImporteFinalBis.Visible = True
FieldImporteFinalTri.Visible = True

FieldVencimientoFinal.Visible = True
FieldVencimientoFinalBis.Visible = True
FieldVencimientoFinalTri.Visible = True

Exit Sub
merror:
tratarerrores "Error en reporte ARCuotasCreditos3-3"
End Sub
Private Sub OcultarCamposActualizacion()
'si no hay actualizacion oculta los campos
On Error GoTo merror

LabelPesos1.Visible = False
LabelPesos2.Visible = False
LabelPesos3.Visible = False

LabelImporteFinal.Visible = False
LabelImporteFinalBis.Visible = False
LabelImporteFinalTri.Visible = False

LabelVencimientoFinal.Visible = False
LabelVencimientoFinalBis.Visible = False
LabelVencimientoFinalTri.Visible = False

FieldImporteFinal.Visible = False
FieldImporteFinalBis.Visible = False
FieldImporteFinalTri.Visible = False

FieldVencimientoFinal.Visible = False
FieldVencimientoFinalBis.Visible = False
FieldVencimientoFinalTri.Visible = False

Exit Sub
merror:
tratarerrores "Error en reporte ARCuotasCreditos3-4"
End Sub
Private Sub Carteles()
On Error GoTo merror

With RDODataControl1.Resultset
     'si esta refinanciada muestro el cartel
     If Not IsNull(.rdoColumns("fecharefinanciacion")) Then
        LabelRefinanciada.Visible = True
        LabelRefinanciadaBis.Visible = True
        LabelRefinanciadaTri.Visible = True
     Else
        LabelRefinanciada.Visible = False
        LabelRefinanciadaBis.Visible = False
        LabelRefinanciadaTri.Visible = False
     End If
     
     'si es comodin muestro el cartel
     If .rdoColumns("cuotacomodin") Then
        LabelComodin.Visible = True
        LabelComodinBis.Visible = True
        LabelComodinTri.Visible = True
     Else
        LabelComodin.Visible = False
        LabelComodinBis.Visible = False
        LabelComodinTri.Visible = False
     End If
End With

Exit Sub
merror:
tratarerrores "Error en reporte Carteles-ARCuotasCredito1-5"
End Sub

