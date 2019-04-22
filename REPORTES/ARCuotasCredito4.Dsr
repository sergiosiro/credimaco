VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARCuotasCredito4 
   Caption         =   "Imprimir cuotas de credito"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "ARCuotasCredito4.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARCuotasCredito4.dsx":0442
End
Attribute VB_Name = "ARCuotasCredito4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
'se ejecuta una sola vez
On Error GoTo merror

'si muestro los recuadros los habilito visibles
If VG_MOSTRARRECUADROS Then
   ShapeCuadro1.Visible = True
   ShapeCuadro2.Visible = True
   LineCuadro1.Visible = True
   LineCuadro2.Visible = True
   LineCuadro3.Visible = True
   LineCuadro4.Visible = True
   LineCuadro5.Visible = True
   LineCuadro6.Visible = True
   LineCuadro7.Visible = True
   LabelTitulo1.Visible = True
   LabelTitulo2.Visible = True
   LabelTitulo3.Visible = True
   LabelTitulo4.Visible = True
   LabelTitulo5.Visible = True
   LabelTitulo6.Visible = True
   LabelTitulo7.Visible = True
   LabelTitulo8.Visible = True
   LabelTitulo9.Visible = True
End If

Exit Sub
merror:
tratarerrores "Error en reporte ARCuotasCreditos4-1"
End Sub
Private Sub Detail_Format()
'se ejecuta para cada factura
Dim ImporteMora As Currency
Dim SaldoCuota As Currency
Dim SaldoCuota1erVenc As Currency
Dim ImporteParcial As Currency
Dim Fecha As Date
Dim IvaMora As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim IvaACobrarDevuelto  As Currency
On Error GoTo merror

ImporteMora = 0

'por defecto no hay mora
FieldImporteVto1.Text = 0
FieldImporteVto2.Text = 0

Fecha = CDate(FieldFecha.Text)
'actualizo los importes si hubo cambios

With RDODataControl1.Resultset
     'campo fecha arriba es el vencimiento de la cuota este cobrada o no
     FieldDia.Text = Format(Day(.rdoColumns("fechavencimiento1")), "00")
     FieldMes.Text = Format(Month(.rdoColumns("fechavencimiento1")), "00")
     FieldAno.Text = Format(Year(.rdoColumns("fechavencimiento1")), "0000")

     FieldImporteVto1.Text = CCur(.rdoColumns("importetotal"))
     
     ImporteParcial = ObtenerImporteParcialX(.rdoColumns("idcredito"), .rdoColumns("numcuota"))
     SaldoCuota = ObtenerSaldoCuotaX(.rdoColumns("idcredito"), .rdoColumns("numcuota"), CDate(FieldFecha.Text), SaldoCuota1erVenc)
     'por defecto el importe 2 es el 1 mas recargo de 2 vto
     FieldImporteVto2.Text = CCur(FieldImporteVto1.Text) + CCur(.rdoColumns("importerecargovencimiento2"))
     
     'si no esta cobrada
     If IsNull(.rdoColumns("fechacobro")) Then
        'si no esta refinanciada y no es comodin
        If IsNull(.rdoColumns("fecharefinanciacion")) And Not (.rdoColumns("cuotacomodin")) Then
           'si estoy en mora actualizo (vale si aplico o no 2 vto)
           If CDate(FieldFecha.Text) > CDate(.rdoColumns("fechavencimiento2")) Then
              'calculo la mora de forma habitual
              'puedo pasarle el campo [exceptuada]
              ImporteMora = CalculoMoraPendiente(.rdoColumns("idcredito"), .rdoColumns("numcuota"), .rdoColumns("exceptuada"), SaldoCuota1erVenc, .rdoColumns("fechavencimiento1"), CDate(FieldFecha.Text), IvaACobrarDevuelto)
              '''''''********ImporteMora = CalcularInteresMoraZZ(.rdoColumns("exceptuada"), SaldoCuota1erVenc, .rdoColumns("fechavencimiento1"), CDate(FieldFecha.Text))
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
           End If
        End If
     Else
        'esta cobrada
        SaldoCuota = CCur(.rdoColumns("importecobrado"))
        Fecha = CDate(.rdoColumns("fechacobro"))
        ImporteMora = CCur(.rdoColumns("importemora"))
        IvaMora = CCur(.rdoColumns("ivamora"))
     End If
     
     'si hay mora muestro los campos mora e iva mora
     If CCur(ImporteMora) > 0 Then
        If VG_IMPRIMIRMORAIVA Then
           FieldMora1.Visible = True
           FieldMora2.Visible = True
           
           FieldIvaMora1.Visible = True
           FieldIvaMora2.Visible = True
           
           FieldImporteMora.Visible = True
           FieldImporteMoraBis.Visible = True
           
           FieldIvaMora.Visible = True
           FieldIvaMoraBis.Visible = True
        End If
     Else
        FieldMora1.Visible = False
        FieldMora2.Visible = False
        FieldIvaMora1.Visible = False
        FieldIvaMora2.Visible = False
        FieldImporteMora.Visible = False
        FieldImporteMoraBis.Visible = False
        FieldIvaMora.Visible = False
        FieldIvaMoraBis.Visible = False
     End If
     
     If CCur(.rdoColumns("importerefinanciacion")) > 0 Then
        FieldRefin1.Visible = True
        FieldRefin2.Visible = True
        FieldRecRefin.Visible = True
        FieldRecRefinBis.Visible = True
        
     Else
        FieldRefin1.Visible = False
        FieldRefin2.Visible = False
        FieldRecRefin.Visible = False
        FieldRecRefinBis.Visible = False
     End If
     
     
     'formateo la salida
     FieldImporteVto1.Text = Format(FieldImporteVto1.Text, "0.00")
     FieldImporteVto2.Text = Format(FieldImporteVto2.Text, "0.00")
     FieldImporteMora.Text = Format(ImporteMora, "0.00")
     FieldImporteMoraBis.Text = Format(ImporteMora, "0.00")
     FieldIvaMora.Text = Format(IvaMora, "0.00")
     FieldIvaMoraBis.Text = Format(IvaMora, "0.00")
End With

Exit Sub
merror:
tratarerrores "Error en reporte ARCuotasCreditos4-2"
End Sub

