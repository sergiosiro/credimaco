VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARListadoCuotas 
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   Icon            =   "ARListadoCuotas.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARListadoCuotas.dsx":0442
End
Attribute VB_Name = "ARListadoCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalCapital As Currency
Dim TotalInteres As Currency
Dim TotalGastos As Currency
Dim TotalSeguros As Currency
Dim TotalImpuestos As Currency
Dim TotalSaldo As Currency
Dim SubtotalPagina As Currency
Private Sub Detail_Format()
Dim ImporteMora As Currency
Dim ImporteParcial As Currency
Dim SaldoCuota As Currency
Dim IvaMora As Currency
Dim RecargoCuota As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim SaldoCuota1erVenc As Currency
On Error GoTo merror

SaldoCuota = 0

With RDODataControl1.Resultset
     Field1.ForeColor = vbBlack
     FieldCuota.ForeColor = vbBlack
     
     'si no esta refinanciada y no es comodin
     If IsNull(.rdoColumns("fecharefinanciacion")) And Not (.rdoColumns("cuotacomodin")) Then
        
        SaldoCuota = 0
        SaldoCuota = ObtenerSaldoCuotaX(.rdoColumns("idcredito"), .rdoColumns("NumCuota"), CDate(FieldFecha.Text), SaldoCuota1erVenc)
        
        'si esta pendiente
        If IsNull(.rdoColumns("fechacobro")) Then
           'si esta entre 1 y 2 vto no hace nada(no le descuenta el recargo al 2 vto)
           'si hay mora
           If CDate(FieldFecha.Text) > CDate(.rdoColumns("fechavencimiento2")) Then
              ImporteMora = CalcularInteresMoraZZ(.rdoColumns("exceptuada"), SaldoCuota, .rdoColumns("fechavencimiento2"), CDate(FieldFecha.Text), .rdoColumns("fechavencimiento1"))
              IvaMora = 0
              If VG_APLICARIMPUESTOS Then
                 If VG_IMPUESTOSCREDIMACO Then
                    'calculo el iva de la mora
                    IvaMora = CCur(VG_PORCENTAJEIVA * ImporteMora / 100)
                 End If
              End If
              SoloMoraCobrada = ObtenerMoraCobrada(.rdoColumns("idcredito"), .rdoColumns("numcuota"))
              SoloIvaMoraCobrada = ObtenerIvaMoraCobrada(.rdoColumns("idcredito"), .rdoColumns("numcuota"))
              If CCur(ImporteMora) <= CCur(SoloMoraCobrada) Then
                 ImporteMora = 0
              Else
                 'si es mayor la mora es solo la diferencia
                 ImporteMora = CCur(ImporteMora) - CCur(SoloMoraCobrada)
              End If
              If CCur(IvaMora) <= CCur(SoloIvaMoraCobrada) Then
                 IvaMora = 0
              Else
                 'si es mayor la mora es solo la diferencia
                 IvaMora = CCur(IvaMora) - CCur(SoloIvaMoraCobrada)
              End If
              SaldoCuota = CCur(SaldoCuota1erVenc) + CCur(ImporteMora) + CCur(IvaMora)
          End If
        Else
           'si esta cobrada el saldo es cero
           SaldoCuota = 0
        End If
        
        'solo actualiza saldos de gastos etc si es cuota ok
        TotalCapital = CCur(TotalCapital) + CCur(.rdoColumns("importeamortizacion"))
        TotalInteres = CCur(TotalInteres) + CCur(.rdoColumns("importeinteres"))
        TotalGastos = CCur(TotalGastos) + CCur(.rdoColumns("importegastos"))
        TotalSeguros = CCur(TotalSeguros) + CCur(.rdoColumns("importeseguros"))
        TotalImpuestos = CCur(TotalImpuestos) + CCur(.rdoColumns("importeimpuestos"))
     Else
        'si es comodin pongo en verde
        If .rdoColumns("cuotacomodin") Then
           Field1.ForeColor = &H8000&
           FieldCuota.ForeColor = &H8000&
        End If
        'si esta refin pongo en morado
        If Not IsNull(.rdoColumns("fecharefinanciacion")) Then
           Field1.ForeColor = &H800080
           FieldCuota.ForeColor = &H800080
        End If
     End If
     'los dos integrados en una
     FieldGastos.Text = CCur(.rdoColumns("importegastos")) + CCur(.rdoColumns("otorgamiento"))
     FieldGastos.Text = Format(FieldGastos.Text, "0.00")
     
     FieldCuota.Text = CStr(.rdoColumns("numcuota")) & "/" & CStr(.rdoColumns("numcuotas"))
     FieldCliente.Text = .rdoColumns("cliente")
     'actualizo totales
     TotalSaldo = CCur(TotalSaldo) + CCur(SaldoCuota)
     
     'formateo los campos
     FieldSaldo.Text = Format(SaldoCuota, "0.00")
End With

Exit Sub
merror:
tratarerrores "Error en reporte ARListaCuotasCredito-1"
End Sub
Private Sub PageFooter_Format()
FieldPagina.Text = Me.pageNumber
End Sub
Private Sub ReportFooter_Format()
FieldCapitalTotal.Text = Format(TotalCapital, "0.00")
FieldInteresTotal.Text = Format(TotalInteres, "0.00")
FieldSegurosTotal.Text = Format(TotalSeguros, "0.00")
FieldIvaTotal.Text = Format(TotalImpuestos, "0.00")
FieldGastosTotal.Text = Format(TotalGastos, "0.00")
FieldTotal.Text = Format(TotalSaldo, "0.00")
End Sub
