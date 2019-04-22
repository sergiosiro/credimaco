Attribute VB_Name = "ModuloFunciones"
'***FUNCIONES VARIAS QUE SE USAN EN TODO EL SISTEMA
'****ESTAS FUNCIONES FUERON PROGRAMADAS PARA LA VERSION
'PERSONALIZADA DE CREDIMACO

Public Sub RefrescarOpcionesSistema()
'carga las variables globales con los parametros grabados en pantalla de opciones
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

sql = "select * from configuracionsistema"

Set rec = cnSQL.OpenResultset(sql)

If rec.EOF Then Exit Sub

'datos de la empresa
VG_EMPRESA = rec.rdoColumns("empresa") & vbNullString
VG_CIUDAD = rec.rdoColumns("ciudad") & vbNullString
VG_CUIT = rec.rdoColumns("cuit") & vbNullString
VG_IVA = rec.rdoColumns("iva") & vbNullString
VG_INGRESOSBRUTOS = rec.rdoColumns("ingresosbrutos") & vbNullString
VG_DOMICILIO = rec.rdoColumns("domicilio") & vbNullString
VG_TELEFONO = rec.rdoColumns("telefono") & vbNullString
VG_EMAIL = rec.rdoColumns("email") & vbNullString
VG_WEBSITE = rec.rdoColumns("website") & vbNullString
VG_HORARIOATENCION = rec.rdoColumns("horarioatencion") & vbNullString
VG_LUGARESPAGO = rec.rdoColumns("lugarespago") & vbNullString
 
'requisitos de credito
VG_GARANTE = rec.rdoColumns("garante")
VG_CLIENTESIMULTANEO = rec.rdoColumns("clientesimultaneo")
VG_CLIENTEJUDICIAL = rec.rdoColumns("clientejudicial")
VG_PAGARCUOTASDESORDENADAS = rec.rdoColumns("pagarcuotasdesordenadas")
VG_FINALIZARAUTOMATICAMENTE = rec.rdoColumns("finalizarautomaticamente")
VG_APLICARCOBROSPARCIALES = rec.rdoColumns("aplicarcobrosparciales")
VG_APLICARVENCIMIENTOSABADOS = rec.rdoColumns("aplicarvencimientosabados")
VG_EDAD = rec.rdoColumns("EdadMaxCredito")
VG_CANT_DIAS = rec.rdoColumns("cantdiasvenc")
VG_FECHALIMITEINGRESO = rec.rdoColumns("fechalimite")
VG_DIAS_MORA = rec.rdoColumns("cantdiasmora")
VG_ANTIGUEDAD_MORA = rec.rdoColumns("antiguedadmora")
VG_MONTO_MORA = rec.rdoColumns("montominimomora")
VG_TIEMPO_LOGOUT = rec.rdoColumns("tiempologout")

'redondeos
VG_REDONDEAR = rec.rdoColumns("redondearcuotas")

'Si confirma la impresion del recibo=factura credimaco luego de un cobro
VG_APLICARRECIBOS = rec.rdoColumns("aplicarrecibo")

'si acepta creditos diferidos
VG_CREDITOSDIFERIDOS = rec.rdoColumns("creditosdiferidos")
'si acepta cobros diferidos
VG_COBROSDIFERIDOS = rec.rdoColumns("cobrosdiferidos")

'tasas de interes
VG_TASAMORA = CDbl(rec.rdoColumns("tasamora"))
'tasa para moras menores a 60 dias
VG_TASAFINANCIACION = CDbl(rec.rdoColumns("tasafinanciacion"))
VG_TASAREFINANCIACION = CDbl(rec.rdoColumns("tasarefinanciacion"))
VG_APLICARTASAREFINANCIACION = rec.rdoColumns("aplicartasarefinanciacion")
'nueva tasa para moras mayores a 60 dias
VG_TASAMORA2 = CDbl(rec.rdoColumns("tasamora2"))

'gastos administrativos
VG_APLICARGASTOS = rec.rdoColumns("aplicargastos")
VG_NOAPLICARGASTOSREFINANCIACION = rec.rdoColumns("noaplicargastosrefinanciacion")
VG_APLICARGASTOSCUOTA1 = rec.rdoColumns("aplicargastoscuota1")
VG_APLICARGASTOSCUOTA2 = rec.rdoColumns("aplicargastoscuota2")
VG_IMPORTEGASTOS = CCur(rec.rdoColumns("importegastos"))
VG_IMPORTEGASTOSFIJOS = CCur(rec.rdoColumns("importegastosfijos"))
VG_PORCCAPNOINT = rec.rdoColumns("PorcentajeCapitalyNoInt")
VG_PORCFUNNOCAP = rec.rdoColumns("PorcentajefuncNoCapital")
VG_PORCCAPINT = rec.rdoColumns("PorcentajeCapitalInteres")

'seguros
VG_APLICARSEGURO = rec.rdoColumns("aplicarseguro")
VG_IMPORTESEGURO = CCur(rec.rdoColumns("importeseguro"))
VG_SEGUROFIJO = CCur(rec.rdoColumns("importesegurosfijos"))
VG_ALICUOTASEGUROS = CDbl(rec.rdoColumns("alicuotaseguros"))
VG_NOAPLICARSEGUROSREFINANCIACION = rec.rdoColumns("noaplicarsegurosrefinanciacion")
VG_APLICARSEGUROSCUOTA1 = rec.rdoColumns("aplicarseguroscuota1")

'impuestos
VG_APLICARIMPUESTOS = rec.rdoColumns("aplicarimpuestos")
VG_NOAPLICARIMPUESTOSREFINANCIACION = rec.rdoColumns("noaplicarimpuestosrefinanciacion")
VG_APLICARIMPUESTOSCUOTA1 = rec.rdoColumns("aplicarimpuestoscuota1")
VG_APLICARIMPUESTOSCUOTA2 = rec.rdoColumns("aplicarimpuestoscuota2")
VG_IMPORTEIMPUESTOS = CCur(rec.rdoColumns("importeimpuestos"))
VG_IMPUESTOSFIJOS = CCur(rec.rdoColumns("importeimpuestosfijos"))
VG_IMPUESTOSCREDIMACO = rec.rdoColumns("impuestoscredimaco")
VG_PORCENTAJEIVA = CDbl(rec.rdoColumns("porcentajeiva"))

'vencimiento
VG_APLICARSEGUNDOVENCIMIENTO = rec.rdoColumns("aplicarsegundovencimiento")
VG_VENCIMIENTO2IMPORTE = CCur(rec.rdoColumns("vencimiento2importe"))
VG_VENCIMIENTO2PORCENTAJE = CDbl(rec.rdoColumns("vencimiento2porcentaje"))
VG_APLICARVENCIMIENTO2MORA = rec.rdoColumns("aplicarvencimiento2mora")
VG_DIASVENCIMIENTOFINANCIACION = rec.rdoColumns("diasvencimientofinanciacion")
VG_DIASVENCIMIENTOREFINANCIACION = rec.rdoColumns("diasvencimientorefinanciacion")

'extension del archivo 1032
VG_CODIGOAUTOMATICO = CLng(rec.rdoColumns("codigoautomatico"))
'numero de empresa ante rapipago 753
VG_NUMEMPRESA = rec.rdoColumns("num1")

'libre deuda
VG_TEXTOLIBREDEUDA1 = rec.rdoColumns("textolibredeuda1") & vbNullString
VG_TEXTOLIBREDEUDA2 = rec.rdoColumns("textolibredeuda2") & vbNullString

'carta reclamo
VG_TEXTOCARTARECLAMO1 = rec.rdoColumns("textocartareclamo1") & vbNullString
VG_TEXTOCARTARECLAMO2 = rec.rdoColumns("textocartareclamo2") & vbNullString

'pagare
VG_TEXTOACUERDOMUTUO1 = rec.rdoColumns("textoacuerdomutuo1") & vbNullString
VG_TEXTOACUERDOMUTUO2 = rec.rdoColumns("textoacuerdomutuo2") & vbNullString
VG_TEXTOACUERDOMUTUO3 = rec.rdoColumns("textoacuerdomutuo3") & vbNullString
VG_TEXTOACUERDOMUTUO4 = rec.rdoColumns("textoacuerdomutuo4") & vbNullString
VG_TEXTOACUERDOMUTUO5 = rec.rdoColumns("textoacuerdomutuo5") & vbNullString
VG_TEXTOACUERDOMUTUO6 = rec.rdoColumns("textoacuerdomutuo6") & vbNullString
VG_TEXTOACUERDOMUTUO7 = rec.rdoColumns("textoacuerdomutuo7") & vbNullString
VG_TEXTOACUERDOMUTUO8 = rec.rdoColumns("textoacuerdomutuo8") & vbNullString
VG_TEXTOACUERDOMUTUO9 = rec.rdoColumns("textoacuerdomutuo9") & vbNullString
VG_TEXTOACUERDOMUTUO10 = rec.rdoColumns("textoacuerdomutuo10") & vbNullString

'Impresion
VG_MODELOFACTURA1 = rec.rdoColumns("modelofactura1")
VG_MODELOFACTURA2 = rec.rdoColumns("modelofactura2")
VG_MODELOFACTURA3 = rec.rdoColumns("modelofactura3")
VG_MODELOFACTURA4 = rec.rdoColumns("modelofactura4")
VG_ULTIMONUMRECIBO = CLng(rec.rdoColumns("ultimonumrecibo"))

'varias para impresion de factura
VG_TOP = CLng(rec.rdoColumns("margentop"))
VG_LEFT = CLng(rec.rdoColumns("margenleft"))
VG_BOTOM = CLng(rec.rdoColumns("margenbotom"))
VG_MOSTRARRECUADROS = rec.rdoColumns("mostrarrecuadros")
VG_NUMCOPIAS = CLng(rec.rdoColumns("numcopias"))
VG_IMPRIMIRMORAIVA = rec.rdoColumns("imprimirmoraiva")

'Gastos de otorgamiento
VG_APLICAROTORGAMIENTO = rec.rdoColumns("aplicarotorgamiento")
VG_APLICAROTORGAMIENTOCUOTA1 = rec.rdoColumns("aplicarotorgamientocuota1")
VG_IMPORTEOTORGAMIENTO = CCur(rec.rdoColumns("importeotorgamiento"))
VG_OTORCAPNOINT = rec.rdoColumns("OtorCapNoInt")
VG_OTORINTNOCAP = rec.rdoColumns("OtorIntNoCap")
VG_OTORCAPMASINT = rec.rdoColumns("OtorCapmasInt")
VG_NOAPLICAROTREFIN = rec.rdoColumns("noaplicarotorrefin")

VG_INICIORP = rec.rdoColumns("iniciorp")

Exit Sub
merror:
tratarerrores "Error en procedimiento RefrescarOpcionesSistema"
End Sub
Public Sub ImprimirFacturaCredimaco(ByVal Condicion As String, ByVal Fecha As Date)
'imprime las facturas de credimaco cobradas
'usada en pantalla de cobros multiples
'usada en cobros individuales
'usada en ingresos
Dim sql As String
Dim rec As rdoResultset
'para que el sistema permita usar los reportes debe instalar previamente
'el active reports
Dim Mreporte As New ARCuotasCredito4
On Error GoTo merror

'obtengo todos los datos de las cuotas con datos de cliente y creditos
sql = "select creditos.idcredito as midcredito,creditos.numcuotas," & _
      "creditos.importesellados as mimpsellado," & _
      "clientes.numlegajo,clientes.domicilio,clientes.telefono," & _
      "clientes.apellido + ' ' + clientes.nombre as cliente," & _
      "clientes.tipoiva,clientes.cuil,localidades.nombre as localidad," & _
      "cuotas.*,cuotas.importevencimiento1 as importetotal,cuotas.logic1 as exceptuada," & _
      "ingresos.numrecibo,ingresos.fechacobro as fechacobro2," & _
      "ingresos.importecobrado as importecobrado2 " & _
      "from localidades inner join (clientes inner join " & _
      "(creditos inner join (cuotas inner join ingresos " & _
      "on cuotas.idcredito=ingresos.idcredito and " & _
      "cuotas.numcuota=ingresos.numcuota) " & _
      "on creditos.idcredito=cuotas.idcredito) " & _
      "on clientes.idcliente=creditos.idcliente) " & _
      "on localidades.idlocalidad=clientes.idlocalidad " & _
      "where " & Condicion


Set rec = cnSQL.OpenResultset(sql)

Mreporte.RDODataControl1.Resultset = rec
   
Mreporte.PageSettings.LeftMargin = VG_LEFT
Mreporte.PageSettings.TopMargin = VG_TOP
Mreporte.PageSettings.TopMargin = VG_BOTOM
Mreporte.FieldFecha.Text = Fecha
   
Mreporte.Show vbModal

Exit Sub
merror:
tratarerrores "Error imprimiendo la factura"
End Sub
Public Sub ImprimirResumenCredito(ByVal IdCredito As Long, ByVal Fecha As Date)
'imprime un resumen del credito
'usada en registrar credito
'usada en consultar creditos
Dim sql As String
Dim rec As rdoResultset
Dim Mreporte As New ARResumenCredito
On Error GoTo merror

sql = "select clientes.numlegajo,cuotas.*,cuotas.logic1 as exceptuada," & _
      "(cuotas.importevencimiento1) as importetotal,creditos.numcuotas," & _
      "creditos.fechacredito,creditos.importeafinanciar," & _
      "creditos.tasa,creditos.formula,creditos.codprestamo," & _
      "clientes.domicilio,localidades.nombre as localidad," & _
      "localidades.codigopostal as cp," & _
      "clientes.apellido + ', ' + clientes.nombre as cliente " & _
      "from localidades inner join (clientes inner join " & _
      "(creditos inner join cuotas on creditos.idcredito=cuotas.idcredito) " & _
      "on clientes.idcliente=creditos.idcliente) " & _
      "on localidades.idlocalidad=clientes.idlocalidad " & _
      "where cuotas.idcredito=" & CLng(IdCredito) & _
      " order by cuotas.numcuota"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Mreporte.RDODataControl1.Resultset = rec
   Mreporte.Caption = "Resumen de credito"
   
   'si imprimo los datos de empresa
   Mreporte.LabelEmpresa = VG_EMPRESA & vbNullString
   Mreporte.LabelTitulo = "Resumen de credito"
   Mreporte.LabelLe = rec.rdoColumns("numlegajo") & vbNullString
   Mreporte.LabelCli = rec.rdoColumns("cliente") & vbNullString
   Mreporte.LabelCiu = rec.rdoColumns("localidad") & vbNullString
   Mreporte.LabelCPo = rec.rdoColumns("cp") & vbNullString
   Mreporte.LabelDom = rec.rdoColumns("domicilio") & vbNullString
   Mreporte.LabelNumCredito = Format(CStr(NumCredito), "000000")
   Mreporte.LabelNumCuotas = Format(CStr(rec.rdoColumns("numcuotas")), "000")
   Mreporte.LabelFechaCredito = CStr(rec.rdoColumns("fechacredito"))
   Mreporte.LabelTasa = CStr(CDbl(rec.rdoColumns("Tasa")))
   Mreporte.LabelCapital = Format(CStr(rec.rdoColumns("importeafinanciar")), "0.00")
   Mreporte.FieldFecha.Text = Fecha

   Mreporte.Show vbModal
Else
   MsgE "No hay creditos para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo el resumen del credito"
End Sub
Public Function ExisteFacturaCredimaco(ByVal NumRecibo As Long) As Boolean
'verifica si ya existe el numero de factura de credimaco en ingresos
'usada en cobros multiples
'usada en cobros individuales
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteFacturaCredimaco = False

sql = "select numrecibo " & _
      "from ingresos " & _
      "where numrecibo='" & CLng(NumRecibo) & "'"
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numrecibo")) Then
      ExisteFacturaCredimaco = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteFactura"
End Function
Public Function ExisteCodPrestamo(ByVal CodPrestamo As String) As Boolean
'verifica si codigo de prestamo ya existe
'usada en registrar creditos
'usada en refinanciar creditos
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteCodPrestamo = False

sql = "select codprestamo " & _
      "from creditos " & _
      "where codprestamo='" & CStr(CodPrestamo) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("codprestamo")) Then
      ExisteCodPrestamo = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteCodPrestamo"
End Function
Public Function ObtenerDiasMoraMaximo(ByVal Cliente As Long, ByVal FechaAlta As Date) As Integer
'Winik 1/11/17
Dim sql As String
Dim sql1 As String
Dim rec As rdoResultset
Dim rec1 As rdoResultset
Dim Max_Cant_Dias_Mora As Integer
Dim Dias_Mora As Integer
Dim Meses_Mora As Integer
Dim Monto_Mora As Currency
Dim FechaMaxAdmitida As Date
Dim HayRefinanciadoMoroso As Boolean
Dim CreditoVigente As Boolean

On Error GoTo merror

Max_Cant_Dias_Mora = 0
HayRefinanciadoMoroso = False

sql = "select * from creditos where IDCliente=" & Cliente

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Do While Not rec.EOF
      If Mid$(rec("codprestamo"), 7, 1) = "M" Then
        HayRefinanciadoMoroso = True
      End If
      If IsNull(rec("fechafinalizacion")) And IsNull(rec("fechabloqueo")) Then
          CreditoVigente = True
      Else
          CreditoVigente = False
      End If
      sql1 = "select * from cuotas  where IDCredito = " & rec("IdCredito")
      Set rec1 = cnSQL.OpenResultset(sql1)
      
      If Not rec1.EOF Then
        Do While Not rec1.EOF
            If rec1("CuotaComodin") = False Then
                If IsNull(rec1("FechaRefinanciacion")) Then
                    Dias_Mora = 0
                    Meses_Mora = DateDiff("m", rec1("fechavencimiento1"), FechaAlta)
                    If Abs(Meses_Mora) <= VG_ANTIGUEDAD_MORA Then
                        If CreditoVigente Then
                            Monto_Mora = ObtenerSaldoCuotaConPunitorios(rec1("idCredito"), rec1("NumCuota"), FechaAlta)
                        Else
                            FechaMaxAdmitida = DateAdd("d", VG_DIAS_MORA, rec1("fechavencimiento1"))
                            Monto_Mora = rec1("importevencimiento1") - ObtenerMontoMoraPagada(rec1("idCredito"), rec1("NumCuota"), FechaMaxAdmitida)
                        End If
                        If Monto_Mora > VG_MONTO_MORA Then
                            If Not IsNull(rec1("FechaCobro")) Then
                                Dias_Mora = DateDiff("d", rec1("fechavencimiento1"), rec1("FechaCobro"))
                            Else
                                Dias_Mora = DateDiff("d", rec1("fechavencimiento1"), FechaAlta)
                            End If
                        End If
                    End If
                    If Dias_Mora > Max_Cant_Dias_Mora Then
                        Max_Cant_Dias_Mora = Dias_Mora
                    End If
                End If
            End If
        rec1.MoveNext
        Loop
      End If
      rec.MoveNext
   Loop
End If

If HayRefinanciadoMoroso Then
    If VG_DIAS_MORA = 9999 Then
        ObtenerDiasMoraMaximo = 0
    Else
        ObtenerDiasMoraMaximo = 30000
    End If
Else
    ObtenerDiasMoraMaximo = Max_Cant_Dias_Mora
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerDiasMoraMaximo"
End Function

Public Function ObtenerMontoMoraPagada(IdCredito As Long, NumCuota As Long, Vencimiento1 As Date) As Currency

Dim sql As String
Dim rec As rdoResultset
Dim Monto As Currency

On Error GoTo merror

Monto = 0

sql = "select sum(importecobrado) as importecobrado from ingresos where idcredito=" & IdCredito & " and numcuota = " & NumCuota & " and fechacobro <= '" & ConvertirFechaSql(Vencimiento1, "DD/MM/YYYY") & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If IsNull(rec("importecobrado")) Then
       Monto = 0
   Else
       Monto = rec("importecobrado")
   End If
End If
ObtenerMontoMoraPagada = Monto

Exit Function
merror:
tratarerrores "Error en funcion ObtenerMontoMoraPagada"
End Function

Public Function CuotaCobrada(ByVal IdCredito As Long, ByVal NumCuota As Long) As Boolean
'verifica si una cuota esta cobrada
'usada en cobros masivos
'usada en cobros individuales
'usada en consultarcreditos
'usada en cobros parciales
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

CuotaCobrada = False

sql = "select fechacobro from cuotas " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("fechacobro")) Then
      CuotaCobrada = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CuotaCobrada"
End Function
Public Function CuotaRefinanciada(ByVal IdCredito As Long, ByVal NumCuota As Long) As Boolean
'verifica si una factura esta refinanciada
'usada en cobros masivos
'usada en cobros individuales
'usada en consultarcreditos
'usada en cobros parciales
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

CuotaRefinanciada = False

sql = "select fecharefinanciacion " & _
      "from cuotas " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("fecharefinanciacion")) Then
      CuotaRefinanciada = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CuotaRefinanciada"
End Function
Public Function CuotaEsComodin(ByVal IdCredito As Long, NumCuota As Long) As Boolean
'chequea si una factura es cuota comodin
'usada en cobros masivos
'usada en cobros individuales
'usada en consultar
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

CuotaEsComodin = False

sql = "select numcuota,cuotacomodin " & _
      "from cuotas " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numcuota")) Then
      If rec.rdoColumns("cuotacomodin") Then
         CuotaEsComodin = True
      End If
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CuotaEsComodin"
End Function
Public Function ObtenerComisionCobrador(ByVal IdCobrador As Long, ByVal ImporteFactura As Currency) As Currency
'calcula lo que le corresponde al cobrador
Dim sql As String
Dim rec As rdoResultset
Dim Comision As Currency
Dim ImporteCobrador As Currency
Dim PorcentajeCobrador As Double
On Error GoTo merror

ObtenerComisionCobrador = 0

sql = "select idcobrador,importecomision,porcentajecomision " & _
      "from cobradores " & _
      "where idcobrador='" & CLng(IdCobrador) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcobrador")) Then
      ImporteCobrador = CCur(rec.rdoColumns("importecomision"))
      PorcentajeCobrador = CDbl(rec.rdoColumns("porcentajecomision"))
      
      If CCur(ImporteCobrador) > 0 Then
         ObtenerComisionCobrador = CCur(ImporteCobrador)
         Exit Function
      End If
      
      If CDbl(PorcentajeCobrador) > 0 Then
         Comision = CDbl(ImporteFactura * PorcentajeCobrador / 100)
         ObtenerComisionCobrador = CCur(Comision)
         Exit Function
      End If
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerComisionCobrador"
End Function
Public Sub CreditoComodin(ByVal IdCredito As Long, ByVal marca As Long)
'marca o desmarca un credito que usa cuota comodin
Dim sql As String
On Error GoTo merror

sql = "update creditos set aplicarcuotacomodin='" & CLng(marca) & "'" & _
      " where creditos.idcredito='" & CLng(IdCredito) & "'"
   
cnSQL.Execute sql
   
Exit Sub
merror:
tratarerrores "Error en procedimiento CreditoComodin"
End Sub
Public Function HayCuotasImpagasPrevias(ByVal IdCredito As Long, ByVal NumCuota As Long) As Boolean
'se usa para evitar cobrar cuotas salteadas
'verifica si hay cuotas impagas anteriores a una determinada cuota
'usada en cobrar cuotas individuales
'falta aun implementarla en cobros masivos
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

HayCuotasImpagasPrevias = False
   
sql = "select cuotas.numcuota " & _
      "from creditos inner join cuotas on " & _
      "creditos.idcredito=cuotas.idcredito " & _
      "where creditos.idcredito='" & CLng(IdCredito) & "'" & _
      "and creditos.fechabloqueo Is Null " & _
      "and creditos.fechafinalizacion is Null " & _
      "and cuotas.fecharefinanciacion is Null " & _
      "and cuotas.fechacobro is Null " & _
      "and cuotas.cuotacomodin = 'False' " & _
      "and cuotas.numcuota < '" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numcuota")) Then
      HayCuotasImpagasPrevias = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion HayCuotasImpagasPrevias"
End Function
Public Function ExisteCredito(ByVal IdCredito As Long) As Boolean
'verifica si un credito existe
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteCredito = False

sql = "select idcredito " & _
      "from creditos " & _
      "where idcredito=" & CLng(IdCredito)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      ExisteCredito = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteCredito"
End Function
Public Function ExisteCuota(ByVal IdCredito As Long, ByVal NumCuota As Long) As Boolean
'verifica si la cuota existe
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteCuota = False

sql = "select numcuota " & _
      "from cuotas " & _
      "where idcredito='" & CLng(IdCredito) & "' and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numcuota")) Then
      ExisteCuota = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteCuota"
End Function
Public Function ExistePlan(ByVal IdPlan As Long) As Boolean
'verifica si un plan existe
'usada en registrar creditos
'usada en refinanciar creditos
'usada en abm planes
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExistePlan = False

sql = "select idplan " & _
      "from planes " & _
      "where idplan=" & CLng(IdPlan)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idplan")) Then
      ExistePlan = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExistePlan"
End Function
Public Function ExisteFactura(ByVal NumFactura As Long) As Boolean
'verifica si una factura (comprobante) existe en tabla cuotas
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteFactura = False

sql = "select numfactura " & _
      "from cuotas " & _
      "where numfactura='" & CLng(NumFactura) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numfactura")) Then
      ExisteFactura = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteFactura"
End Function
Public Function ExisteCliente(ByVal IdCliente As Long) As Boolean
'verifica si el cliente existe
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

ExisteCliente = False

sql = "select idcliente from clientes " & _
      "where idcliente=" & CLng(IdCliente)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcliente")) Then
      ExisteCliente = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteCliente"
End Function
Public Function ExisteCobrador(ByVal IdCobrador As Long) As Boolean
'verifica si existe un cobrador
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteCobrador = False

sql = "select idcobrador " & _
      "from cobradores " & _
      "where idcobrador=" & CLng(IdCobrador)
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcobrador")) Then
      ExisteCobrador = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteCobrador"
End Function

Public Function NombreCobrador(ByVal IdCobrador As Long) As String
'devuelve el nombre de un cobrador
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

NombreCobrador = ""

sql = "select nombre, apellido " & _
      "from cobradores " & _
      "where idcobrador=" & CLng(IdCobrador)
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
    NombreCobrador = rec.rdoColumns("apellido") & " " & rec.rdoColumns("nombre")
End If

Exit Function
merror:
tratarerrores "Error en funcion NombreCobrador"
End Function

Public Function ExisteBanco(ByVal IdBanco As Long) As Boolean
'verifica si un banco existe
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteBanco = False

sql = "select idbanco " & _
      "from bancos " & _
      "where idbanco=" & CLng(IdBanco)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idbanco")) Then
      ExisteBanco = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteBanco"
End Function
Public Function ExisteComercio(ByVal IdComercio As Long) As Boolean
'verifica si un comercio existe
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteComercio = False

sql = "select idcomercio " & _
      "from comercios " & _
      "where idcomercio=" & CLng(IdComercio)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcomercio")) Then
      ExisteComercio = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteComercio"
End Function
Public Function ExisteVendedor(ByVal IdVendedor As Long) As Boolean
'verifica si un vendedor existe
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteVendedor = False

sql = "select idvendedor " & _
      "from vendedores " & _
      "where idvendedor=" & CLng(IdVendedor)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idvendedor")) Then
      ExisteVendedor = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteVendedor"
End Function
Public Function ExisteUsuario(ByVal idusuario As Long) As Boolean
'verifica si un usuario existe
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteUsuario = False

sql = "select idusuario " & _
      "from usuarios " & _
      "where idusuario=" & CLng(idusuario)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idusuario")) Then
      ExisteUsuario = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteUsuario"
End Function
Public Function ExisteTipoUsuario(ByVal Idtipousuario As Long) As Boolean
'verifica si un tipo de usuario existe
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteTipoUsuario = False

sql = "select idtipousuario " & _
      "from tipousuario " & _
      "where idtipousuario=" & CLng(Idtipousuario)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idtipousuario")) Then
      ExisteTipoUsuario = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteTipoUsuario"
End Function
Public Function ExisteLocalidad(ByVal IdLocalidad As Long) As Boolean
'verifica si una localidad existe
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteLocalidad = False

sql = "select idlocalidad " & _
      "from localidades " & _
      "where idlocalidad=" & CLng(IdLocalidad)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idlocalidad")) Then
      ExisteLocalidad = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteLocalidad"
End Function
Public Function ExisteProvincia(ByVal IdProvincia As Long) As Boolean
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteProvincia = False

sql = "select idprovincia " & _
      "from provincias " & _
      "where idprovincia=" & CLng(IdProvincia)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idprovincia")) Then
      ExisteProvincia = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteProvincia"
End Function
Public Function ExisteEstudio(ByVal IdEstudio As Long) As Boolean
'verifica si un estudio existe
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteEstudio = False

sql = "select idestudio " & _
      "from estudios " & _
      "where idestudio=" & CLng(IdEstudio)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idestudio")) Then
      ExisteEstudio = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteEstudio"
End Function

Public Function CreditoBloqueado(ByVal NumFactura As Long) As Boolean
'chequea si un credito esta bloqueado en base a una factura
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

CreditoBloqueado = False

sql = "select creditos.idcredito,creditos.fechabloqueo " & _
      "from creditos inner join cuotas " & _
      "on creditos.idcredito=cuotas.idcredito " & _
      "where cuotas.numfactura='" & CLng(NumFactura) & "'"
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      If Not IsNull(rec.rdoColumns("fechabloqueo")) Then
         CreditoBloqueado = True
      End If
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CreditoBloqueado"
End Function
Public Function CreditoBloqueado1(ByVal IdCredito As Long) As Boolean
'chequea si un credito esta bloqueado
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

CreditoBloqueado1 = False

sql = "select idcredito,fechabloqueo " & _
      "from creditos " & _
      "where idcredito=" & CLng(IdCredito)
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      If Not IsNull(rec.rdoColumns("fechabloqueo")) Then
         CreditoBloqueado1 = True
      End If
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CreditoBloqueado1"
End Function
Public Function CreditoFinalizado(ByVal IdCredito As Long) As Boolean
'chequea si un credito esta finalizado
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

CreditoFinalizado = False

sql = "select idcredito,fechafinalizacion " & _
      "from creditos " & _
      "where idcredito='" & CLng(IdCredito) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      If Not IsNull(rec.rdoColumns("fechafinalizacion")) Then
         CreditoFinalizado = True
      End If
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CreditoFinalizado"
End Function
Public Function FacturaRefinanciada(ByVal NumFactura As Long) As Boolean
'chequea si una cuota esta esta refinanciada en base a su nro de factura
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

FacturaRefinanciada = False

sql = "select numfactura,fecharefinanciacion " & _
      "from cuotas " & _
      "where numfactura='" & CLng(NumFactura) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numfactura")) Then
      If Not IsNull(rec.rdoColumns("fecharefinanciacion")) Then
         FacturaRefinanciada = True
      End If
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion FacturaRefinanciada"
End Function
Public Function FacturaEsComodin(ByVal NumFactura As Long) As Boolean
'chequea si una cuota es comodin en base a su factura
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

FacturaEsComodin = False

sql = "select numfactura,cuotacomodin " & _
      "from cuotas " & _
      "where numfactura='" & CLng(NumFactura) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numfactura")) Then
      If rec.rdoColumns("cuotacomodin") Then
         FacturaEsComodin = True
      End If
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion FacturaEsComodin"
End Function
Public Function CantidadCuotasComodin(ByVal IdCredito As Long) As Long
'devuelve la cantidad de cuotas comodin de un credito
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

CantidadCuotasComodin = 0

sql = "select count(cuotas.cuotacomodin) as cantidad " & _
      "from cuotas " & _
      "where idcredito=" & CLng(IdCredito) & " " & _
      "and cuotacomodin=1"
    
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("cantidad")) Then
      CantidadCuotasComodin = CLng(rec.rdoColumns("cantidad"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CantidadCuotasComodin"
End Function
Public Function CreditoTieneCobros(ByVal IdCredito As Long) As Boolean
'verifica si un credito tiene cuotas cobradas
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

CreditoTieneCobros = False

sql = "select idcredito " & _
      "from cuotas " & _
      "where idcredito='" & CLng(IdCredito) & _
      "' and fechacobro is not Null"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      CreditoTieneCobros = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CreditoTienecobros"
End Function
Public Function CreditoTieneRefinanciadas(ByVal IdCredito As Long) As Boolean
'cuenta si un credito tiene facturas refinanciadas
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

CreditoTieneRefinanciadas = False

sql = "select idcredito " & _
      "from cuotas " & _
      "where idcredito='" & CLng(IdCredito) & _
      "' and fecharefinanciacion is not Null"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      CreditoTieneRefinanciadas = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CreditoTieneRefinanciadas"
End Function
Public Sub FinalizarCredito(ByVal IdCredito As Long, ByVal Fecha As Date)
'al pagar la ultima cuota finaliza el credito
Dim sql As String
On Error GoTo merror

If CuotasImpagas(IdCredito) = 0 Then
   sql = "update creditos set fechafinalizacion='" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "' " & _
         "where creditos.idcredito=" & CLng(IdCredito)
   cnSQL.Execute sql
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento FinalizarCredito"
End Sub
Public Function CuotasImpagas(ByVal IdCredito As Long) As Long
'chequean cuantas cuotas impagas quedan de un credito vigente
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

CuotasImpagas = 0

sql = "select count(cuotas.numfactura) as cantidad " & _
      "from creditos inner join cuotas on creditos.idcredito=cuotas.idcredito " & _
      "where creditos.idcredito=" & CLng(IdCredito) & " " & _
      "and cuotas.fechacobro is null " & _
      "and cuotas.fecharefinanciacion is null " & _
      "and cuotas.cuotacomodin = 0"
            
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull("cantidad") Then
      CuotasImpagas = CLng(rec.rdoColumns("cantidad"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CuotasImpagas"
End Function
Public Function CalcularRecargoRefinanciacion(ByVal ImporteARefinanciar As Currency) As Currency
'calcula la comision por refinanciacion
'usada en refinanciarcreditos
Dim Porcentaje As Double
Dim Resultado As Currency
On Error GoTo merror

Resultado = 0

If VG_APLICARTASAREFINANCIACION Then
   If CDbl(VG_TASAREFINANCIACION) > 0 Then
      'lo pongo en tanto por uno
      Porcentaje = CDbl(VG_TASAREFINANCIACION / 100)
      Resultado = CCur(ImporteARefinanciar * Porcentaje)
   End If
End If

CalcularRecargoRefinanciacion = CCur(Resultado)

Exit Function
merror:
tratarerrores "Error en funcion CalcularRecargoRefinanciacion"
End Function
Public Function RequisitosGeneralesClienteOk(ByVal IdCliente As Long) As Boolean
'Verifica si el cliente es simultaneo o es judicial
On Error GoTo merror

RequisitosGeneralesClienteOk = True

'si no permito clientes simultaneos
If Not VG_CLIENTESIMULTANEO Then
   'verifico si el cliente tiene un credito pendiente anterior
   If TieneCreditoVigente(IdCliente) Then
      RequisitosGeneralesClienteOk = False
      MsgE "El cliente ya tiene un credito vigente..."
      Exit Function
   End If
End If

'si no permito clientes titulares de creditos bloqueados
If Not VG_CLIENTEJUDICIAL Then
   If TieneCreditoJudicial(IdCliente) Then
      RequisitosGeneralesClienteOk = False
      MsgE "El cliente es titular de un credito bloqueado"
      Exit Function
   End If
End If
   
Exit Function
merror:
tratarerrores "Error en funcion RequisitosGeneralesClienteOk"
End Function
Public Function TieneCreditoVigente(ByVal IdCliente As Long) As Boolean
'verifica si un cliente tiene un credito pendiente vigente
'usada en la funcion RequisitosGeneralesClienteOk
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

TieneCreditoVigente = False

sql = "select idcredito " & _
      "from creditos " & _
      "where idcliente=" & CLng(IdCliente) & _
      " and fechafinalizacion is null and fechabloqueo is null"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      TieneCreditoVigente = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion TieneCreditoVigente"
End Function
Public Function TieneCreditoJudicial(ByVal IdCliente As Long) As Boolean
'verifica si un cliente tiene un credito bloqueado a su nombre
'usada en la funcion RequisitosGeneralesClienteOk
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

TieneCreditoJudicial = False

sql = "select idcredito " & _
      "from creditos " & _
      "where idcliente=" & CLng(IdCliente) & _
      " and not fechabloqueo is null"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      TieneCreditoJudicial = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion TieneCreditoJudicial"
End Function
Public Function RequisitosBasicosClienteOk(ByVal IdCliente As Long, ByVal ImporteSolicitado As Currency) As Boolean
'verifica requisitos de un cliente..debe comparar con los requisitos exigidos para ese tipo de cliente
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

RequisitosBasicosClienteOk = True

'obtengo los datos del cliente y de su tipo
sql = "select idcliente,creditomaximo " & _
      "from clientes " & _
      "where idcliente=" & CLng(IdCliente)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If IsNull(rec.rdoColumns("idcliente")) Then
      RequisitosBasicosClienteOk = False
      Exit Function
   End If
   
   'chequeo el importe maximo de credito
   If (CCur(ImporteSolicitado) > CCur(rec.rdoColumns("creditomaximo"))) Then
      RequisitosBasicosClienteOk = False
      MsgE "El importe PTF es superior al permitido para este cliente"
      Exit Function
   End If
Else
   RequisitosBasicosClienteOk = False
   MsgE "El cliente no existe en la base de datos"
End If

Exit Function
merror:
tratarerrores "Error en funcion RequisitosBasicosClienteOk"
End Function
Public Function CalcularImporteVencimiento2(ByVal importevto1 As Currency, ByVal Vencimiento1 As Date, ByVal Vencimiento2 As Date) As Currency
'usada en registrar creditos y refinanciar creditos
'le agregue una tercera variante que es la mora de credimaco
'calcula el importe del segundo vencimiento en base al primer vencimiento
'y a las opciones predeterminadas del sistema
Dim Importe As Currency
Dim DiasDiferencia As Long
Dim Importeporcentajediario As Currency
Dim Importeporcentaje As Currency
On Error GoTo merror

'por defecto devuelve el importe del 1º vto
Importe = CCur(importevto1)

If VG_APLICARSEGUNDOVENCIMIENTO Then
   If CDate(Vencimiento2) > CDate(Vencimiento1) Then
      DiasDiferencia = DateDiff("d", Vencimiento1, Vencimiento2)
      
      'si aplico un recargo fijo sin importar la diferencia de dias
      If CCur(VG_VENCIMIENTO2IMPORTE) > 0 Then
         Importe = CCur(importevto1 + CCur(VG_VENCIMIENTO2IMPORTE / 10))
      End If
      
      'si aplico un recargo de porcentaje sobre el importe del vto1
      If CDbl(VG_VENCIMIENTO2PORCENTAJE) > 0 Then
         'calculo el porcentaje sobre el importe del vto1
         Importeporcentaje = CCur(importevto1 * CDbl(VG_VENCIMIENTO2PORCENTAJE / 100))
         'le sumo ese importe
         Importe = CCur(importevto1) + CCur(Importeporcentaje)
      End If
      
      'si aplico la formula de credimaco
      If VG_APLICARVENCIMIENTO2MORA Then
         Importe = (CCur(importevto1) + CalcularInteresMoraZZ(False, importevto1, Vencimiento1, Vencimiento2, Vencimiento1))
      End If
      
   End If
End If

CalcularImporteVencimiento2 = CCur(Importe)

Exit Function
merror:
tratarerrores "Error en funcion CalcularImporteVencimiento2"
End Function

Public Function ObtenerUltimaFechaCobroParcialCuota(ByVal IdCredito As Long, ByVal NumCuota As Long) As Date
'obtiene la ultima fecha de cobro parcial de un credito
Dim rec As rdoResultset
Dim sql As String
Dim Fecha As Date
On Error GoTo merror

ObtenerUltimaFechaCobroParcialCuota = CDate("1900/01/01")

sql = "ObtenerUltimaFechaCobroParcialCuota " & IdCredito & "," & NumCuota
           
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("fechacobro")) Then
      ObtenerUltimaFechaCobroParcialCuota = CDate(rec.rdoColumns("fechacobro"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerUltimaFechaCobroParcialCuota"
End Function

Public Function CalculoMoraPendiente(ByVal IdCredito As Long, ByVal NumCuota As Long, ByVal Exceptuada As Boolean, ByVal Saldo1erVto As Currency, ByVal Fecha1erVenc As Date, ByVal FechaCalculo As Date, ByRef IvaACobrarDevuelto As Currency) As Currency
'obtiene la ultima fecha de cobro parcial de un credito
Dim rec As rdoResultset
Dim sql As String
Dim nTotalMora As Currency
Dim nIva As Currency
Dim nMora As Currency
Dim nSaldo As Currency
Dim nDif As Currency
Dim dFecha As Date
Dim dFechaCobro As Date
Dim IvaACobrar As Currency
Dim UltimoInteres As Currency
On Error GoTo merror

sql = "ObtenerFechasCredito " & IdCredito
           
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
    If Not IsNull(rec.rdoColumns("FechaFinalizacion")) Then
        If rec.rdoColumns("FechaFinalizacion") < FechaCalculo Then
            FechaCalculo = rec.rdoColumns("FechaFinalizacion")
        End If
    Else
        If Not IsNull(rec.rdoColumns("FechaBloqueo")) Then
            If rec.rdoColumns("FechaBloqueo") < FechaCalculo Then
                FechaCalculo = rec.rdoColumns("FechaBloqueo")
            End If
        End If
    End If
End If

rec.Close

sql = "ObtenerCobrosParciales " & IdCredito & "," & NumCuota
           
Set rec = cnSQL.OpenResultset(sql)

nTotalMora = 0
IvaACobrar = 0
If Not rec.EOF Then
    nSaldo = Saldo1erVto
    dFecha = Fecha1erVenc
    Do While Not rec.EOF
        nMora = CalcularInteresMoraZZ(Exceptuada, nSaldo, dFecha, rec.rdoColumns("FechaCobro"), Fecha1erVenc)
        nIva = nMora * CCur(VG_PORCENTAJEIVA) / 100
        If nMora + nTotalMora > rec.rdoColumns("MoraCobrada") Then
            nTotalMora = nTotalMora + nMora - rec.rdoColumns("MoraCobrada")
            If nIva > rec.rdoColumns("IvaMoraCobrada") Then
                IvaACobrar = IvaACobrar + (nIva - rec.rdoColumns("IvaMoraCobrada"))
            End If
        End If
        If rec.rdoColumns("ImporteCobrado") > (rec.rdoColumns("MoraCobrada") + rec.rdoColumns("IvaMoraCobrada")) Then
            nSaldo = nSaldo - (rec.rdoColumns("ImporteCobrado") - rec.rdoColumns("MoraCobrada") - rec.rdoColumns("IvaMoraCobrada"))
        Else
            nDif = nMora - rec.rdoColumns("MoraCobrada")
            nIva = nDif * CCur(VG_PORCENTAJEIVA) / 100
            If nIva >= rec.rdoColumns("IvaMoraCobrada") Then
                nSaldo = nSaldo + nDif + nIva - rec.rdoColumns("IvaMoraCobrada")
            Else
                nSaldo = nSaldo + nDif
            End If
        End If
        dFecha = rec.rdoColumns("FechaCobro")
        If dFecha < Fecha1erVenc Then
            dFecha = Fecha1erVenc
        End If
        rec.MoveNext
    Loop
    UltimoInteres = CalcularInteresMoraZZ(Exceptuada, nSaldo, dFecha, FechaCalculo, Fecha1erVenc)
    nTotalMora = nTotalMora + UltimoInteres
    IvaACobrar = IvaACobrar + CCur(VG_PORCENTAJEIVA) * UltimoInteres / 100
Else
    nTotalMora = CalcularInteresMoraZZ(Exceptuada, Saldo1erVto, Fecha1erVenc, FechaCalculo, Fecha1erVenc)
    IvaACobrar = CCur(VG_PORCENTAJEIVA) * nTotalMora / 100
End If

IvaACobrarDevuelto = 0

If VG_APLICARIMPUESTOS Then
    If VG_IMPUESTOSCREDIMACO Then
        IvaACobrarDevuelto = IvaACobrar
    End If
End If

CalculoMoraPendiente = nTotalMora

Exit Function
merror:
tratarerrores "Error en funcion CalculoMoraPendiente"
End Function


Public Function CalcularInteresMoraZZ(ByVal Exceptuada As Boolean, ByVal Importe As Currency, ByVal Vencimiento As Date, ByVal Hoy As Date, ByVal PrimerVencimiento As Date) As Currency
'tiene en cuenta las exceptuadas por el campo logic1
Dim Interes As Currency
Dim Dias As Long
Dim Tasadiaria As Double
Dim Diferencia As Double
Dim Calculo As Boolean
Dim DiasDif As Long
Dim DiasDesdePrimerVencimiento As Long
On Error GoTo merror

CalcularInteresMoraZZ = 0

'solo calcula la mora si la cuota no esta exceptuada
If Not Exceptuada Then
   If CCur(Importe) > 0 Then
      If (Vencimiento < Hoy) Then
         DiasDesdePrimerVencimiento = DateDiff("d", PrimerVencimiento, Hoy)
         Dias = DateDiff("d", Vencimiento, Hoy)
         'si la mora es inferior a 60 dias aplico la tasa anual
         If DiasDesdePrimerVencimiento < 60 Then
            Tasadiaria = CDbl(VG_TASAMORA / 365)
         Else
            'si la mora es mayor o igual a 60 dias de atraso
            'aplico una tasa mensual especial
            Tasadiaria = CDbl(VG_TASAMORA2 / 30)
         End If
         Diferencia = CDbl(Dias / 100)
         'MW 1/10/2014
         Interes = CCur(Importe * Diferencia * Tasadiaria)
         CalcularInteresMoraZZ = CCur(Interes)
      End If
   End If
End If 'de exceptuada

Exit Function
merror:
tratarerrores "Error en funcion CalcularInteresMoraZZ"
End Function
Public Function ObtenerSaldoCliente(ByVal IdCliente As Long) As Currency
'obtiene el saldo total de un cliente siempre de creditos vigentes del mismo
'usada en registrarclientes-imprimir resumen de clientes y exportar clientes
Dim sql As String
Dim rec As rdoResultset
Dim SaldoCuota As Currency
Dim ImporteTotal As Currency
Dim ImporteMora As Currency
Dim IvaMora As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim SaldoCuota1erVenc As Currency
Dim Importe1erVenc As Currency
Dim IvaACobrarDevuelto  As Currency
On Error GoTo merror

ObtenerSaldoCliente = 0

sql = "select creditos.idcredito,cuotas.numcuota,cuotas.logic1 as exceptuada," & _
      "cuotas.fechavencimiento1,cuotas.fechavencimiento2," & _
      "(cuotas.importevencimiento1) as importe1,cuotas.importevencimiento2 as importe2 " & _
      "from creditos inner join cuotas on creditos.idcredito=cuotas.idcredito " & _
      "where creditos.idcliente=" & CLng(IdCliente) & _
      " and cuotas.cuotacomodin=0 " & _
      "and cuotas.fecharefinanciacion is null " & _
      "and cuotas.fechacobro is null " & _
      "and creditos.fechafinalizacion is null " & _
      "and creditos.fechabloqueo is null"

Set rec = cnSQL.OpenResultset(sql)

ImporteTotal = 0

If Not rec.EOF Then
   Do While Not rec.EOF
      'si hay cobros parciales en esa cuota
      SaldoCuota = ObtenerSaldoCuotaX(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), Date, SaldoCuota1erVenc)
      ImporteMora = 0
      IvaMora = 0
      'si hay mora
      If Date > CDate(rec.rdoColumns("fechavencimiento2")) Then
             'calculo la mora de manera habitual
             'puedo poner el campo [exceptuada]
             Importe1erVenc = ObtenerImporte1erVenc(IdCredito, NumCuota)
             ImporteMora = CalculoMoraPendiente(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("exceptuada"), Importe1erVenc, rec.rdoColumns("fechavencimiento1"), Date, IvaACobrarDevuelto)
             '''''''********ImporteMora = CalcularInteresMoraZZ(rec.rdoColumns("exceptuada"), SaldoCuota, rec.rdoColumns("fechavencimiento2"), Date)
             If VG_APLICARIMPUESTOS Then
                If VG_IMPUESTOSCREDIMACO Then
                   'ahora le incluyo el iva de mora
                   IvaMora = IvaACobrarDevuelto
                End If
             End If
             SoloMoraCobrada = ObtenerMoraCobrada(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
             SoloIvaMoraCobrada = ObtenerIvaMoraCobrada(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
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
      
      ImporteTotal = CCur(ImporteTotal) + CCur(SaldoCuota)
      rec.MoveNext
   Loop
End If

ObtenerSaldoCliente = CCur(ImporteTotal)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSaldoCliente"
End Function
Public Function ObtenerSaldoCredito(ByVal IdCredito As Long, ByVal Fecha As Date) As Currency
'obtiene el saldo de un credito
'usada en cobrarcreditos
'usada en consultarcreditos
'usada en exportarcreditos
'usada en planilla de creditos
Dim sql As String
Dim rec As rdoResultset
Dim SaldoCuota As Currency
Dim Total As Currency
Dim Exceptuada As Boolean
Dim FechaVencimiento1 As Date
Dim FechaVencimiento2 As Date
Dim NumCuota As Long
On Error GoTo merror

ObtenerSaldoCredito = 0

'obtiene todas las cuotas pendientes de un credito
'trae de todos tipo de credito vigente-bloqueado-finalizado

sql = "ObtenerSaldoCredito " & CLng(IdCredito)
Set rec = cnSQL.OpenResultset(sql)

Total = 0

If Not rec.EOF Then

   Do While Not rec.EOF

      NumCuota = rec.rdoColumns("numcuota")

      FechaVencimiento1 = rec.rdoColumns("fechavencimiento1")

      FechaVencimiento2 = rec.rdoColumns("fechavencimiento2")

      Exceptuada = rec.rdoColumns("exceptuada")

      SaldoCuota = ObtenerSaldoCuotaOKK(IdCredito, NumCuota, FechaVencimiento1, FechaVencimiento2, Exceptuada, Fecha)
      
      Total = CCur(Total) + CCur(SaldoCuota)
      
      rec.MoveNext
      
   Loop
   
End If


ObtenerSaldoCredito = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSaldoCredito: '" & IdCredito & "' " & Err.Number & " " & Err.Description
End Function

Public Function ObtenerEstadoCreditoFinalizado(ByVal IdCredito As Long) As String
'obtiene el saldo de un credito
'usada en cobrarcreditos
'usada en consultarcreditos
'usada en exportarcreditos
'usada en planilla de creditos
Dim sql As String
Dim rec As rdoResultset
Dim Estado As String
On Error GoTo merror

Estado = "Finalizado"

sql = "ObtenerEstadoCreditoFinalizado " & CLng(IdCredito)
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
    If Not IsNull(rec("estado")) Then
        Estado = rec("estado")
    End If
End If


ObtenerEstadoCreditoFinalizado = Estado

Exit Function
merror:
tratarerrores "Error en funcion ObtenerEstadoCreditoFinalizado: '" & IdCredito & "' " & Err.Number & " " & Err.Description
End Function

Public Function ObtenerUltimoLegajo() As Long
'obtiene el ultimo numero de cliente
'usada en registrarclientes
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerUltimoLegajo = 0

'obtengo el mas reciente grabado
sql = "select numlegajo " & _
      "from clientes " & _
      "order by idcliente desc"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numlegajo")) Then
      If IsNumeric(rec.rdoColumns("numlegajo")) Then
         ObtenerUltimoLegajo = CLng(rec.rdoColumns("numlegajo"))
      End If
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerUltimoLegajo"
End Function
Public Function ObtenerCuotasCobradas(ByVal IdCredito As Long) As Long
'chequean cuantas cuotas fueron cobradas de un credito
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

ObtenerCuotasCobradas = 0
        
sql = "ObtenerCuotasCobradas  " & CLng(IdCredito)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull("cantidad") Then
      ObtenerCuotasCobradas = CLng(rec.rdoColumns("cantidad"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCuotasCobradas"
End Function
Public Function ObtenerUltimaFechaCobro(ByVal IdCredito As Long) As Date
'obtiene la ultima fecha de cobro de un credito
'usada para la planilla de creditos y exportar
Dim rec As rdoResultset
Dim sql As String
Dim Fecha As Date
On Error GoTo merror

'otra alternativa es que devuelva un espacio en blanco
'si no hay ultima fecha de cobro

ObtenerUltimaFechaCobro = Date + 1

sql = "ObtenerUltimaFechaCobro " & CLng(IdCredito)
           
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("fechacobro")) Then
      ObtenerUltimaFechaCobro = CDate(rec.rdoColumns("fechacobro"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerUltimaFechaCobro"
End Function
Public Function ObtenerUltimaFechaCobroParcial(ByVal IdCredito As Long) As Date
'obtiene la ultima fecha de cobro parcial de un credito
Dim rec As rdoResultset
Dim sql As String
Dim Fecha As Date
On Error GoTo merror

ObtenerUltimaFechaCobroParcial = Date + 1

sql = "ObtenerUltimaFechaCobroParcial " & CLng(IdCredito)
           
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("fechacobro")) Then
      ObtenerUltimaFechaCobroParcial = CDate(rec.rdoColumns("fechacobro"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerUltimaFechaCobroParcial"
End Function


Public Function ObtenerSaldoCuotaX(ByVal IdCredito As Long, ByVal NumCuota As Long, ByVal Fecha As Date, ByRef Saldo1erVenc As Currency) As Currency
'nueva de credimaco..calcula en base a lo que aun resta cobrar
'de todos los items que la componen
Dim rec As rdoResultset
Dim sql As String
Dim SaldoCuota As Currency
Dim ImporteCobrado As Currency
On Error GoTo merror

ObtenerSaldoCuotaX = 0

'ahora obtengo los valores originales de la cuota
sql = "ObtenerSaldoCuotaX " & CLng(IdCredito) & "," & CLng(NumCuota)
Set rec = cnSQL.OpenResultset(sql)

SaldoCuota = 0

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numcuota")) Then
      'si no esta cobrada
      If Not rec.rdoColumns("exceptuada") Then
         If IsNull(rec.rdoColumns("fechacobro")) Then
            'primero obtengo lo cobrado hasta ahora sin mora
            ImporteCobrado = ObtenerImporteParcialZ(IdCredito, NumCuota)
            
            'si esta pendiente
            'si estoy al primer vencimiento
            Saldo1erVenc = CCur(rec.rdoColumns("importevencimiento1")) - CCur(ImporteCobrado)
            If CDate(Fecha) <= CDate(rec.rdoColumns("fechavencimiento1")) Then
               SaldoCuota = CCur(rec.rdoColumns("importevencimiento1")) - CCur(ImporteCobrado)
            Else
               'si estoy despues del primer vencimiento
               SaldoCuota = CCur(rec.rdoColumns("importevencimiento2")) - CCur(ImporteCobrado)
            End If
         End If
      Else
        ImporteCobrado = ObtenerImporteParcialZ(IdCredito, NumCuota)
        Saldo1erVenc = CCur(rec.rdoColumns("importevencimiento1")) - CCur(ImporteCobrado)
        SaldoCuota = CCur(rec.rdoColumns("importevencimiento1")) - CCur(ImporteCobrado)
      End If
   End If
End If

ObtenerSaldoCuotaX = CCur(SaldoCuota)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSaldoCuotaX"
End Function

Public Function ObtenerSaldoCuotaConPunitorios(ByVal IdCredito As Long, ByVal NumCuota As Long, ByVal Fecha As Date) As Currency
Dim SaldoCuota1erVenc As Currency
Dim SaldoCuota  As Currency
Dim Importe1erVenc As Currency
Dim ImporteActualizado  As Currency
Dim ImporteMora  As Currency
Dim IvaMora As Currency
Dim DiasMora As Integer
Dim IvaACobrarDevuelto As Currency
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror
   sql = "SELECT * From cuotas " & _
         "WHERE idcredito = " & IdCredito & " AND numcuota = " & NumCuota
    
   Set rec = cnSQL.OpenResultset(sql)
    
   If rec.EOF Then
       ObtenerSaldoCuotaConPunitorios = 0
       rec.Close
       Exit Function
   End If

   SaldoCuota = ObtenerSaldoCuotaX(IdCredito, NumCuota, Fecha, SaldoCuota1erVenc)
   Importe1erVenc = ObtenerImporte1erVenc(IdCredito, NumCuota)
   ImporteActualizado = 0
   ImporteMora = 0
   IvaMora = 0
   DiasMora = 0
   If Fecha > CDate(rec.rdoColumns("fechavencimiento1")) Then
        DiasMora = Fecha - CDate(rec.rdoColumns("fechavencimiento1"))
        If SaldoCuota = 0 Then
            DiasMora = 0
        End If
   End If
   'si no esta cobrada actualizo el importe si es necesario
   If IsNull(rec.rdoColumns("fechacobro")) And Not rec.rdoColumns("cuotacomodin") And IsNull(rec.rdoColumns("fecharefinanciacion")) Then
      'esto funciona para ambos vencimientos (si hay un solo vto ambos son iguales)
      If Fecha > CDate(rec.rdoColumns("fechavencimiento2")) Then
         'esto es para mostrar en la columna correspondiente
         'calculo la mora en forma habitual
         'puedo pasarle el campo [exceptuada]
         ImporteMora = CalculoMoraPendiente(IdCredito, NumCuota, rec.rdoColumns("logic1"), Importe1erVenc, rec.rdoColumns("fechavencimiento1"), Fecha, IvaACobrarDevuelto)
         If VG_APLICARIMPUESTOS Then
            If VG_IMPUESTOSCREDIMACO Then
               'calculo el iva de la mora
               IvaMora = IvaACobrarDevuelto
            End If
         End If
         
         ImporteActualizado = CCur(SaldoCuota1erVenc) + CCur(ImporteMora) + CCur(IvaMora)
      Else
        ImporteActualizado = CCur(SaldoCuota)
      End If
   Else
    ImporteActualizado = CCur(SaldoCuota)
   End If
   rec.Close
   ObtenerSaldoCuotaConPunitorios = ImporteActualizado

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSaldoCuotaConPunitorios"
End Function




Public Function ObtenerSaldoVencido(ByVal IdCredito As Long, ByVal Fecha As Date) As Currency
'obtiene el importe total en mora (todo concepto mas mora) de un credito
'solo de cuotas en mora
Dim sql As String
Dim rec As rdoResultset
Dim Suma As Currency
Dim SaldoCuota As Currency
Dim NumCuota As Long
Dim Exceptuada As Boolean
Dim FechaVencimiento1 As Date
Dim FechaVencimiento2 As Date
On Error GoTo merror

ObtenerSaldoVencido = 0

sql = "ObtenerSaldoVencido " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"

'sql = "select creditos.idcredito,cuotas.numcuota," & _
'      "cuotas.fechavencimiento1,cuotas.fechavencimiento2,cuotas.logic1 as exceptuada " & _
'      "from creditos inner join cuotas " & _
'      "on creditos.idcredito=cuotas.idcredito " & _
'      "where creditos.idcredito='" & CLng(IdCredito) & "' " & _
'      "and cuotas.fechacobro is Null " & _
'      "and cuotas.cuotacomodin = 0" & _
'      "and cuotas.logic1 = 0 " & _
'      "and cuotas.fecharefinanciacion is Null " & _
'      "and cuotas.fechavencimiento1<'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "' " & _
'      "and creditos.fechafinalizacion is Null"
      
Set rec = cnSQL.OpenResultset(sql)

Suma = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      NumCuota = CLng(rec.rdoColumns("numcuota"))
      FechaVencimiento1 = rec.rdoColumns("fechavencimiento1")
      FechaVencimiento2 = rec.rdoColumns("fechavencimiento2")
      Exceptuada = rec.rdoColumns("exceptuada")
      'este saldo es total de la cuota
      SaldoCuota = ObtenerSaldoCuotaOKK(IdCredito, NumCuota, FechaVencimiento1, FechaVencimiento2, Exceptuada, Fecha)
      Suma = CCur(Suma) + CCur(SaldoCuota)
      rec.MoveNext
   Loop
End If

ObtenerSaldoVencido = CCur(Suma)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSaldoVencido"
End Function

'Public Function ObtenerSaldoCuotaX1erVenc(ByVal IdCredito As Long, ByVal NumCuota As Long, ByVal Fecha As Date) As Currency
'Dim rec As rdoResultset
'Dim sql As String
'Dim SaldoCuota As Currency
'Dim ImporteCobrado As Currency
'On Error GoTo merror
'
'ObtenerSaldoCuotaX1erVenc = 0
'sql = "ObtenerSaldoCuotaX " & CLng(IdCredito) & "," & CLng(NumCuota)
'Set rec = cnSQL.OpenResultset(sql)
'
'SaldoCuota = 0
'
'If Not rec.EOF Then
'   If Not IsNull(rec.rdoColumns("numcuota")) Then
'      If Not rec.rdoColumns("exceptuada") Then
'         If IsNull(rec.rdoColumns("fechacobro")) Then
'            ImporteCobrado = ObtenerImporteParcialZ(IdCredito, NumCuota)
'
'            SaldoCuota = CCur(rec.rdoColumns("importevencimiento1")) - CCur(ImporteCobrado)
'
'         End If
'      End If
'   End If
'End If
'
'ObtenerSaldoCuotaX1erVenc = CCur(SaldoCuota)
'
'Exit Function
'merror:
'tratarerrores "Error en funcion ObtenerSaldoCuotaX1erVenc"
'End Function

Public Function ObtenerImporte1erVenc(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
Dim rec As rdoResultset
Dim sql As String
Dim ImporteCuota As Currency
Dim ImporteCobrado As Currency
On Error GoTo merror

ObtenerImporte1erVenc = 0

'ahora obtengo los valores originales de la cuota

sql = "ObtenerImporte1erVenc " & CLng(IdCredito) & "," & CLng(NumCuota)
Set rec = cnSQL.OpenResultset(sql)

ImporteCuota = 0

If Not rec.EOF Then
   ImporteCuota = CCur(rec.rdoColumns("importevencimiento1"))
End If

ObtenerImporte1erVenc = CCur(ImporteCuota)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerImporte1erVenc"
End Function
Public Function ObtenerImporteParcial(ByVal IdCredito As Long) As Currency
'calcula el parcial importecobrado de un credito
Dim rec As rdoResultset
Dim sql As String
Dim ImporteCobrado1 As Currency
On Error GoTo merror

ObtenerImporteParcial = 0

'obtengo lo cobrado hasta la fecha de esa cuota..sacado de ingresos
sql = "select numcuota,importecobrado " & _
      "from ingresos " & _
      "where idcredito='" & CLng(IdCredito) & "'"
      '"and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

ImporteCobrado1 = 0

If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         'sumo todo el importe cobrado
         ImporteCobrado1 = CCur(ImporteCobrado1) + CCur(rec.rdoColumns("importecobrado"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerImporteParcial = CCur(ImporteCobrado1)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerImporteParcial"
End Function


Public Function ObtenerImporteParcialX(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'calcula el importecobrado de una cuota desde ingresos
Dim rec As rdoResultset
Dim sql As String
Dim ImporteCobrado1 As Currency
On Error GoTo merror

ObtenerImporteParcialX = 0

'obtengo lo cobrado hasta la fecha de esa cuota..sacado de ingresos
sql = "select numcuota,importecobrado " & _
      "from ingresos " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

ImporteCobrado1 = 0

If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         'sumo todo el importe cobrado
         ImporteCobrado1 = CCur(ImporteCobrado1) + CCur(rec.rdoColumns("importecobrado"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerImporteParcialX = CCur(ImporteCobrado1)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerImporteParcialX"
End Function
Public Function ObtenerImporteParcialZ(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'devuelve lo cobrado de los items sin incluir la mora e ivamora cobradas
Dim rec As rdoResultset
Dim sql As String
Dim ImporteCobrado As Currency
On Error GoTo merror

ObtenerImporteParcialZ = 0

'obtengo lo cobrado hasta la fecha de esa cuota..sacado de ingresos
sql = "ObtenerImporteParcialZ " & CLng(IdCredito) & "," & CLng(NumCuota)

Set rec = cnSQL.OpenResultset(sql)

ImporteCobrado = 0

If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         'sumo todo el importe cobrado
         ImporteCobrado = CCur(ImporteCobrado) + CCur(rec.rdoColumns("importecobrado"))
         'le resto la mora e iva mora
         ImporteCobrado = CCur(ImporteCobrado) - CCur(rec.rdoColumns("moracobrada")) - CCur(rec.rdoColumns("ivamoracobrada"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerImporteParcialZ = CCur(ImporteCobrado)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerImporteParcialZ"
End Function
Public Function ObtenerPorcentajeSellados(ByVal IdProvincia As Long) As Double
'obtiene el porcentaje de sellado de una provincia
'usada en registrarcreditos
'usada en refinanciarcreditos
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerPorcentajeSellados = 0

sql = "select porcentajesellados " & _
      "from provincias " & _
      "where idprovincia=" & CLng(IdProvincia)
    
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("porcentajesellados")) Then
      ObtenerPorcentajeSellados = CDbl(rec.rdoColumns("porcentajesellados"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerPorcentajeSellados"
End Function
Public Function ObtenerImporteSellados(ByVal Porcentaje As Double, ByVal Importe As Currency) As Currency
'calcula el importe de sellados
'usada en registrar creditos
'usada en refinanciarcreditos
On Error GoTo merror

ObtenerImporteSellados = Porcentaje * Importe / 100

Exit Function
merror:
tratarerrores "Error en funcion ObtenerImporteSellados"
End Function
Public Function ObtenerIdCredito(ByVal IdCliente As Long, ByVal NumComprobante As Long) As Long
'obtiene el credito de una factura
'usada en importar rapipago
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerIdCredito = 0

sql = "select creditos.idcredito " & _
      "from creditos inner join cuotas on creditos.idcredito=cuotas.idcredito " & _
      "where creditos.idcliente='" & CLng(IdCliente) & "' and cuotas.numfactura='" & CLng(NumComprobante) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      ObtenerIdCredito = CLng(rec.rdoColumns("idcredito"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerIdCredito"
End Function
Public Function ObtenerNumCuota(ByVal IdCliente As Long, ByVal NumComprobante As Long) As Long
'obtiene la cuota de una factura
'usada en importar rapipago
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerNumCuota = 0

sql = "select cuotas.numcuota " & _
      "from creditos inner join cuotas on creditos.idcredito=cuotas.idcredito " & _
      "where creditos.idcliente='" & CLng(IdCliente) & "' and cuotas.numfactura='" & CLng(NumComprobante) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numcuota")) Then
      ObtenerNumCuota = CLng(rec.rdoColumns("numcuota"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerNumCuota"
End Function
Public Function ObtenerCodPrestamo(ByVal IdCliente As Long, ByVal IdCredito As Long) As String
'obtiene el codigo de prestamo de un credito
'usada en importar rapipago
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerCodPrestamo = ""

sql = "select idcredito,codprestamo " & _
      "from creditos " & _
      "where idcliente='" & CLng(IdCliente) & "' and idcredito='" & CLng(IdCredito) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      ObtenerCodPrestamo = rec.rdoColumns("codprestamo")
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCodPrestamo"
End Function

Public Function ObtenerCodPrestamoConDocumento(ByVal NroDocumento As String) As String
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerCodPrestamoConDocumento = ""

sql = "SELECT Creditos.CodPrestamo " & _
      "FROM Creditos, Clientes, Cuotas " & _
      "WHERE NumDocumento = '" & Val(NroDocumento) & "' and Clientes.IdCliente = Creditos.IdCliente and " & _
      "FechaFinalizacion IS NULL and FechaBloqueo IS NULL and Creditos.FechaRefinanciacion IS NULL and " & _
      "Cuotas.IdCredito = Creditos.IdCredito and FechaCobro IS NULL and cuotas.cuotacomodin = 0" & _
      "ORDER BY Cuotas.FechaVencimiento1, Creditos.fechacredito"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   ObtenerCodPrestamoConDocumento = rec.rdoColumns("CodPrestamo")
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCodPrestamoConDocumento"
End Function

Public Function ObtenerCodPrestamoConDocumentoInclFinalizados(ByVal NroDocumento As String) As String
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerCodPrestamoConDocumentoInclFinalizados = ""

sql = "SELECT Creditos.CodPrestamo " & _
      "FROM Creditos, Clientes " & _
      "WHERE NumDocumento = '" & Val(NroDocumento) & "' and Clientes.IdCliente = Creditos.IdCliente " & _
      "ORDER BY Creditos.fechacredito desc"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   ObtenerCodPrestamoConDocumentoInclFinalizados = rec.rdoColumns("CodPrestamo")
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCodPrestamoConDocumentoInclFinalizados"
End Function


Public Function ObtenerCliente(ByVal IdCliente As Long) As String
'obtiene el nombre del cliente
'usada en importar rapipago
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerCliente = ""

sql = "select apellido + ' ' + nombre as cliente " & _
      "from clientes " & _
      "where idcliente='" & CLng(IdCliente) & "'"
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   ObtenerCliente = CStr(rec.rdoColumns("cliente"))
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCliente"
End Function
Public Function TieneCobrosParciales(ByVal IdCredito As Long, ByVal NumCuota As Long) As Boolean
'chequea si una cuota tiene cobros parciales
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

TieneCobrosParciales = False

sql = "select numcuota,cobrosparciales " & _
      "from cuotas " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numcuota")) Then
      If rec.rdoColumns("cobrosparciales") Then
         TieneCobrosParciales = True
      End If
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion TieneCobrosParciales"
End Function
Public Function ObtenerCapitalCobrado(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene el capital cobrado de una cuota desde tabla ingresos
'usada en cobro de cuotas, consulta de cobros, importar ambos e importar rapipago
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerCapitalCobrado = 0

'sql = "select numcuota,capitalcobrado " & _
'      "from ingresos " & _
'      "where idcredito='" & CLng(IdCredito) & "' " & _
'      "and numcuota='" & CLng(NumCuota) & "'"
sql = "ObtenerCapitalCobrado " & CLng(IdCredito) & "," & CLng(NumCuota)

Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("capitalcobrado"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerCapitalCobrado = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCapitalCobrado"
End Function
Public Function ObtenerInteresCobrado(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene el interes cobrado de una cuota desde ingresos
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerInteresCobrado = 0

sql = "select numcuota,interescobrado " & _
      "from ingresos " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("interescobrado"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerInteresCobrado = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerInteresCobrado"
End Function
Public Function ObtenerGastosCobrados(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene el gastocobrado de una cuota desde ingresos
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerGastosCobrados = 0

sql = "select numcuota,gastoscobrados " & _
      "from ingresos " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("gastoscobrados"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerGastosCobrados = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerGastosCobrados"
End Function
Public Function ObtenerSegurosCobrados(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene el segurocobrado de una cuota desde ingresos
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerSegurosCobrados = 0

sql = "select numcuota,seguroscobrados " & _
      "from ingresos " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("seguroscobrados"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerSegurosCobrados = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSegurosCobrados"
End Function
Public Function ObtenerOtorgamientoCobrado(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene el otorgamientocobrado de una cuota desde ingresos
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerOtorgamientoCobrado = 0

sql = "select numcuota,otorgamientocobrado " & _
      "from ingresos " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("otorgamientocobrado"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerOtorgamientoCobrado = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerOtorgamientoCobrado"
End Function
Public Function ObtenerVencimiento2Cobrado(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene el vencimiento2cobrado de una cuota desde ingresos
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerVencimiento2Cobrado = 0

sql = "select numcuota,vencimiento2cobrado " & _
      "from ingresos " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("vencimiento2cobrado"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerVencimiento2Cobrado = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerVencimiento2Cobrado"
End Function
Public Function ObtenerRefinCobrado(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene el item refincobrado (comision de refinanciacion) de una cuota desde ingresos
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerRefinCobrado = 0

sql = "select numcuota,refincobrado " & _
      "from ingresos " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("refincobrado"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerRefinCobrado = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerRefinCobrado"
End Function
Public Function ObtenerIvaInteresCobrado(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene el iva interescobrado de una cuota desde ingresos
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerIvaInteresCobrado = 0

sql = "select numcuota,ivainterescobrado " & _
      "from ingresos " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("ivainterescobrado"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerIvaInteresCobrado = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerIvaInteresCobrado"
End Function
Public Function ObtenerIvaSegurosCobrado(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene el iva seguros cobrado de una cuota desde ingresos
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerIvaSegurosCobrado = 0

sql = "select numcuota,ivaseguroscobrado " & _
      "from ingresos " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("ivaseguroscobrado"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerIvaSegurosCobrado = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerIvaSegurosCobrado"
End Function
Public Function ObtenerIvaOtorGastosCobrado(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene el ivaotorgastos cobrado de una cuota desde ingresos
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerIvaOtorGastosCobrado = 0

sql = "select numcuota,ivaotorgastoscobrado " & _
      "from ingresos " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"

Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("ivaotorgastoscobrado"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerIvaOtorGastosCobrado = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerIvaOtorGastosCobrado"
End Function
Public Function ObtenerMoraCobrada(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene la mora cobrada de una cuota desde ingresos
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerMoraCobrada = 0

sql = "ObtenerMoraCobrada " & CLng(IdCredito) & "," & CLng(NumCuota)
Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("moracobrada"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerMoraCobrada = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerMoraCobrada"
End Function
Public Function ObtenerIvaMoraCobrada(ByVal IdCredito As Long, ByVal NumCuota As Long) As Currency
'obtiene el iva mora cobrado de una cuota desde ingresos
Dim sql As String
Dim rec As rdoResultset
Dim Total As Currency
On Error GoTo merror

ObtenerIvaMoraCobrada = 0

sql = "ObtenerIVAMoraCobrada " & CLng(IdCredito) & "," & CLng(NumCuota)

Set rec = cnSQL.OpenResultset(sql)

Total = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      If Not IsNull(rec.rdoColumns("numcuota")) Then
         Total = CCur(Total) + CCur(rec.rdoColumns("ivamoracobrada"))
      End If
      rec.MoveNext
   Loop
End If

ObtenerIvaMoraCobrada = CCur(Total)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerIvaMoraCobrada"
End Function

Public Function ObtenerNuevoCupon() As Long
Dim rec As rdoResultset
Dim sql As String
Dim nNuevoCupon As Long

On Error GoTo merror

nNuevoCupon = 0

Set rec = cnSQL.OpenResultset("select ultimocupon from configuracionsistema")

If Not rec.EOF Then
    nNuevoCupon = rec.rdoColumns("ultimocupon") + 1
End If

ObtenerNuevoCupon = nNuevoCupon

Exit Function
merror:
ObtenerNuevoCupon = 0
tratarerrores "Error en funcion ObtenerNuevoCupon"
End Function

Public Function UltimoId(ByVal campo As String, ByVal tabla As String) As Long
'devuelve el maximo valor de un campo de una tabla
'usada para ir creando nuevos id de cada tabla en los abm
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

UltimoId = 0

sql = "SELECT MAX(" & campo & ") AS Num FROM " & tabla

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("num")) Then
      UltimoId = CLng(rec.rdoColumns("num"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion UltimoId"
End Function
'***********************FUNCIONES DE FECHAS**************************************

Public Function FormatearFecha(ByVal Fecha As Date) As String
'recibe una fecha y devuelve..por ej:lunes 4 de noviembre de 2006
Dim Dia As String
Dim Mes As String
Dim Año As String
Dim nombredia As String
On Error GoTo merror

nombredia = CStr(Format(Fecha, "dddd")) 'lunes
Dia = CStr(Day(Fecha))
Mes = CStr(Format(Fecha, "mmmm"))
Año = CStr(Format(Fecha, "yyyy"))

FormatearFecha = nombredia + " " + Dia + " de " + Mes + " de " + Año

Exit Function
merror:
tratarerrores "Error en funcion FormatearFecha"
End Function
Public Function ArmarFecha(ByVal diaprioridad As Long, ByVal Fecha As Date, ByVal Frecuencia As Integer) As Date
'recibe una fecha a la que tiene que incrementarle un mes,o dos meses
'usada para generar las fechas de vencimiento de cuotas
Dim NuevaFecha As String
Dim Dia As Long
Dim Mes As Long
Dim Año As Long
Dim incremento As Long
On Error GoTo merror

Dia = diaprioridad
Mes = Month(Fecha) '11
Año = Year(Fecha) '2006

'si es mensual
If Frecuencia = 1 Then
   incremento = 1
End If

'si es bimestral
If Frecuencia = 2 Then
   incremento = 2
End If

'si cambio de año
If (Mes + incremento) > 12 Then
   Mes = (Mes + incremento) - 12
   Año = Año + 1
Else
   Mes = Mes + incremento
End If

'verificar que pasa con la configuracion de fechas en ingles
'compongo como string el nuevo vencimiento----20-2-2005
NuevaFecha = CStr(Dia) & "/" & CStr(Mes) & "/" & CStr(Año)

'verifico de devolver una fecha valida por ej no poner 30-febrero
Do While Not IsDate(NuevaFecha)
   Dia = Dia - 1
   NuevaFecha = CStr(Dia) & "/" & CStr(Mes) & "/" & CStr(Año)
Loop

ArmarFecha = CDate(NuevaFecha)

Exit Function
merror:
tratarerrores "Error en funcion ArmarFecha"
End Function
Public Function ObtenerFechaVencimiento(ByVal Fecha As Date, ByVal Dias As Long) As Date
'busca un dia habil de vencimiento de cuotas
Dim Mes As Long
Dim Fechatemp1 As Date
Dim Fechatemp2 As Date
Dim Diferencia1 As Long
Dim Diferencia2 As Long
On Error GoTo merror

'tendria que devolver la misma si no es feriado ni fin de semana
'y ahorrarse los ciclos posteriores ganado tiempo
If Not EsFeriado(Fecha) And Not EsSabado(Fecha) And Not EsDomingo(Fecha) Then
   ObtenerFechaVencimiento = CDate(Fecha)
   'la fecha era habil y la devuelvo inmediatamente
   Exit Function
End If

'si continuo hacia abajo es porque la fecha no es habil y debo buscar un dia habil
Fechatemp1 = Fecha
Fechatemp2 = Fecha

'primero si es feriado o fin de semana busca hacia adelante
Do While EsFeriado(Fechatemp1) Or EsSabado(Fechatemp1) Or EsDomingo(Fechatemp1)
   Fechatemp1 = Fechatemp1 + 1
   'si es fin de mes puede llegar a pasarse al siguiente
Loop

'si la diferencia de vencimientos es por 1 dia solo llevo hacia adelante
If Dias = 1 Then
   ObtenerFechaVencimiento = Fechatemp1
   Exit Function
End If

'ahora pruebo buscando hacia atras si es feriado o fin de semana
Do While EsFeriado(Fechatemp2) Or EsSabado(Fechatemp2) Or EsDomingo(Fechatemp2)
   Fechatemp2 = Fechatemp2 - 1
   'si es principio de mes puede llegar a pasarse al mes anterior
Loop

'tomo el mes de la fecha original
Mes = Month(Fecha)

'si las dos fechas son del mismo mes debo devolver la mas cercana
If Month(Fechatemp1) = Mes And Month(Fechatemp2) = Mes Then
   Diferencia1 = DateDiff("d", Fecha, Fechatemp1)
   Diferencia2 = DateDiff("d", Fechatemp2, Fecha)
   If CLng(Diferencia1) <= CLng(Diferencia2) Then
      'si la diferencia es igual devuelve la que va hacia adelante
      ObtenerFechaVencimiento = CDate(Fechatemp1)
   Else
      ObtenerFechaVencimiento = CDate(Fechatemp2)
   End If
Else
   'si la fecha 1 se paso de mes devuelvo la segunda
   If Month(Fechatemp1) <> Mes Then
      ObtenerFechaVencimiento = CDate(Fechatemp2)
   End If
   
   'si la fecha 2 se paso de mes devuelvo la primera
   If Month(Fechatemp2) <> Mes Then
      ObtenerFechaVencimiento = CDate(Fechatemp1)
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerFechaVencimiento"
End Function
Public Function EsFeriado(ByVal Fecha As Date) As Boolean
'abre la tabla de feriados y chequea
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

EsFeriado = False

sql = "select fecha " & _
      "from feriados " & _
      "where fecha='" & ConvertirFechaSql(Fecha, "DD/MM/YYYY") & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("fecha")) Then
      EsFeriado = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion EsFeriado"
End Function
Public Function EsSabado(ByVal Fecha As Date) As Boolean
On Error GoTo merror

EsSabado = False
'si no aplico vencimiento los sabados verifico si es sabado
If Not VG_APLICARVENCIMIENTOSABADOS Then
   If Weekday(Fecha) = 7 Then
      EsSabado = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion EsSabado"
End Function
Public Function EsDomingo(ByVal Fecha As Date) As Boolean
On Error GoTo merror

EsDomingo = False

If Weekday(Fecha) = 1 Then
   EsDomingo = True
End If

Exit Function
merror:
tratarerrores "Error en funcion EsDomingo"
End Function
Public Function CrearFechaNormal(ByVal FechaJuliana As String) As Date
'convierte una fecha juliana en normal
'usada en importacion de rapipago
Dim fecha1 As Date
Dim fecha2 As Date
Dim CadAño As String
Dim Año As Long
Dim CadDias As String
Dim Dias As Long
On Error GoTo merror

'obtengo el año desde la fecha juliana que son los primeros dos digitos
CadAño = Mid(FechaJuliana, 1, 2)
'formo el 2008 por ejemplo
CadAño = "20" + Format(CadAño, "00")

Año = CLng(CadAño)

'armo la fecha inico del año
fecha1 = CDate("01/01/" + CadAño)

'obtengo los dias de la juliana
CadDias = Mid(FechaJuliana, 3, 3)
Dias = CLng(CadDias)
'armo la nueva fecha normal sumandole los dias
fecha2 = CDate(fecha1) + Dias - 1

CrearFechaNormal = CDate(fecha2)

Exit Function
merror:
tratarerrores "Error en funcion CrearFechaNormal"
End Function
Public Function GenerarCodigoBarras(ByVal EmpServicio As Integer, ByVal NumCliente As Long, ByVal NumCta As Long, ByVal Monto1Venc As Currency, ByVal Fecha1venc As Date, ByVal MontoRecargo As Currency, ByVal DiasRecargo As Long) As String
'esta funcion es la nueva de RAPIPAGO
'Esta funcion se encarga de armar el codigo de barras en formato string
' 1º : Numero de Empresa          - 3 posiciones (Este # es asignado al comercio)
' 2º : Numero de cliente          - 8 posiciones (completar con 0 a la izq)(cliente titular de la factura)
' 3º : Numero de comprobante      - 11 posiciones
' 4º : Importe al primer vto      - 8 posiciones (compleatr con 0 a la izq.)(6 ent y 2 dec)
' 5º : Fecha 1º vto juliano       - 5 posiciones AAJJJ (AA ES EL AÑO JJJ ES EL Nº DE DIA DESDE EL PRIMER DIA DEL AÑO)
' 6º : Recargo despues del 1º vto - 6 posiciones (4 ENT Y 2 DEC)
' 7º : Dias al 2º vto             - 2 posiciones
' 8º : Digito verificador         - 1 posicion
'(*) ancho total del codigo 44 incluyendo el verificador

Dim Cadena As String
Dim CadenaFinal As String
Dim Empresa As String
Dim ImporteEntero1 As String
Dim ImporteDecimal1 As String
Dim ImporteEntero2 As String
Dim ImporteDecimal2 As String
Dim Comprobante As String
Dim Ndecimal As Double
Dim fecha1 As String
Dim vardecimal As Long
Dim AA, JJJ As String
'este se usa para el calculo del digito verificador
Dim matriz(43) As Integer
Dim Secuencia(43) As Integer
'nuevos
Dim Cliente As String
Dim CadDias As String
Dim I As Long
Dim Suma As Currency
Dim Division As Currency
Dim Digito As String
Dim Entero As Long
On Error GoTo merror

'EMPRESA:obtengo el numero de empresa
Cadena = CStr(EmpServicio)
Empresa = Trim(Format(Cadena, "000"))

'CLIENTE:obtengo el numero de cliente (esto es nuevo)
Cadena = CStr(NumCliente)
Cliente = Trim(Format(Cadena, "00000000"))

'COMPROBANTE:obtengo el numero de factura ahora de 11 digitos (antes era de 15)
Cadena = CStr(NumCta)
Comprobante = Format(Cadena, "00000000000")

'IMPORTE 1º VTO:obtengo el importe 1er vencimiento
ImporteEntero1 = Trim(Format(Int(Monto1Venc), "000000"))

'obtengo la parte decimal..por jemplo 0,56
Ndecimal = CCur(Monto1Venc) - Int(Monto1Venc)
'aca la multiplica por 100 para obtener solo los ultimos dos digitos sin la coma y el cero
vardecimal = (Format(Ndecimal, "0.00")) * 100
Cadena = CStr(vardecimal)
ImporteDecimal1 = Trim(Format(Cadena, "00"))

'FECHA 1º VTO:calculo de la fecha de vencimento
'saco el año en formato juliano
AA = Format(Right(Year(Fecha1venc), 2), "00")
'saco ahora la fecha juliana que es la diferencia entre la fecha parametro y el 1ro de enero
fecha1 = "01/01/" & Year(Fecha1venc)
JJJ = Abs(Format(DateDiff("d", fecha1, Fecha1venc), "000"))
JJJ = JJJ + 1
fecha1 = AA + Format(JJJ, "000")

'RECARGO AL 2º VTO:obtengo el importe 1er vencimiento (menos digitos)
ImporteEntero2 = Trim(Format(Int(MontoRecargo), "0000"))
Ndecimal = CCur(MontoRecargo) - Int(MontoRecargo)
vardecimal = (Format(Ndecimal, "0.00")) * 100
Cadena = CStr(vardecimal)
ImporteDecimal2 = Trim(Format(Cadena, "00"))

'DIAS 2º VTO:
Cadena = CStr(DiasRecargo)
CadDias = Trim(Format(Cadena, "00"))

'empresa + cliente + comprobante + importe 1º vto + fecha 1º vto + recargo + dias
CadenaFinal = Empresa + Cliente + Comprobante + ImporteEntero1 + ImporteDecimal1 + fecha1 + ImporteEntero2 + ImporteDecimal2 + CadDias

'aca inicia el tema del digito verificador
For I = 1 To 43
    cadenx = Left(CadenaFinal, I)
    
    'esta linea saca solo el digito mas reciente y lo va cargando en el vector
    'al vector lo va armando de izquierda a derecha
    cadenx = Mid(cadenx, I, 1)
    matriz(I) = cadenx
Next I

'calculo del digito verificador
Secuencia(1) = 1
Secuencia(2) = 3
Secuencia(3) = 5

X = 4
Do While X <= 43
    Secuencia(X) = 7
    X = X + 1
    Secuencia(X) = 9
    X = X + 1
    Secuencia(X) = 3
    X = X + 1
    Secuencia(X) = 5
    X = X + 1
Loop

For I = 1 To 43
    'multiplico el numero de la secuencia por cada elemento de
    'la serie a verificar
    matriz(I) = matriz(I) * Secuencia(I)
    Suma = Suma + matriz(I)
Next I

Division = Suma / 2
Entero = Int(Division)

'si esto esta demas el calculo seria para cumplir con rapipago
Cadena = CStr(Entero)
'el DV es el primero de la derecha
Digito = Trim(Right(Cadena, 1))

CadenaFinal = CadenaFinal + Digito
'fin codigo verificador

GenerarCodigoBarras = CadenaFinal

Exit Function
merror:
tratarerrores "Error en funcion GenerarCodigoBarras"
End Function
Public Function VerificarSeleccionLista(ByVal lv As ListView) As Boolean
'verifica si hay un elemento seleccionado en un listview
Dim I As Long
On Error GoTo merror

VerificarSeleccionLista = False

If lv.ListItems.Count = 0 Then Exit Function
       
For I = 1 To CLng(lv.ListItems.Count())
    If lv.ListItems(I).Selected = True Then
       VerificarSeleccionLista = True
       Exit Function
    End If
Next I

Exit Function
merror:
tratarerrores "Error en funcion VerificarSeleccionLista"
End Function
Public Function HayFilasChequeadas(ByVal lv As ListView) As Boolean
'devuelve si hay alguna fila marcada en un listview
Dim Filas As Long
Dim I As Long
On Error GoTo merror

HayFilasChequeadas = False

Filas = lv.ListItems.Count

If Filas = 0 Then Exit Function

For I = 1 To Filas
    If lv.ListItems.Item(I).Checked Then
       HayFilasChequeadas = True
       Exit Function
    End If
Next I

Exit Function
merror:
tratarerrores "Error en funcion HayFilasChequeadas"
End Function
Public Function HayMasChequeadas(ByVal lv As ListView) As Boolean
'devuelve si hay filas marcadas en un listview
'usada en importaciones (importar ambos)
'aca importa la cantidad de selecciones que haya
Dim Filas As Long
Dim I As Long
Dim Contador As Long
On Error GoTo merror

HayMasChequeadas = False

Filas = lv.ListItems.Count

If Filas = 0 Then Exit Function

Contador = 0

For I = 1 To Filas
    If lv.ListItems.Item(I).Checked Then
       Contador = Contador + 1
    End If
Next I

If Contador > 1 Then
   HayMasChequeadas = True
End If

Exit Function
merror:
tratarerrores "Error en funcion HayMasChequeadas"
End Function
Public Function Encriptar(ByVal password As String) As String
'encripta una clave de acceso de usuarios de la base de access del sistema
Dim Origen As Integer
Dim rvalue As String
Dim Pos As Integer
Dim place As Integer
On Error GoTo merror

password = Trim(password)

Origen = 105
rvalue = ""

Pos = 1
While Pos <= Len(password)
      place = (Asc(Mid(password, Pos, 1)) + 2550 + Origen - Pos) Mod 255
      rvalue = rvalue + Chr(place)
      Pos = Pos + 1
Wend

Encriptar = rvalue
    
Exit Function
merror:
tratarerrores "Error en funcion Encriptar"
End Function
Public Function Desencriptar(ByVal password As String) As String
'desencripta una clave de acceso de usuarios de la base de access del sistema
Dim Origen As Integer
Dim rvalue As String
Dim Pos As Integer
Dim place As Integer
On Error GoTo merror

Origen = 105
rvalue = ""
    
Pos = 1
While Pos <= Len(password)
      place = (Asc(Mid(password, Pos, 1)) + 2550 - Origen + Pos) Mod 255
      rvalue = rvalue + Chr(place)
      Pos = Pos + 1
Wend

Desencriptar = rvalue

Exit Function
merror:
tratarerrores "Error en funcion Desencriptar"
End Function
Public Sub tratarerrores(Optional Texto As String)
'esta funcion reacciona ante una colgada emitiendo un mensaje
'de esta forma el programa no se cuelga y continua funcionando
Dim Msj As String
Dim MiError As Error
Dim sAux As String

On Error Resume Next
    
If Not IsEmpty(Texto) Then Msj = "ORIGEN: " + Texto + "." + vbCr
    
For Each MiError In Errors
    sAux = sAux + "ERROR:  " + CStr(MiError.Number) + " = " + MiError.Description + "." + vbCr
Next
    
If sAux <> "" Then
   Msj = Msj + sAux
   MsgBox Msj, , "Mensaje de error", Err.HelpFile, Err.HelpContext
Else
   MsgBox Msj, , "Mensaje de error", Err.HelpFile, Err.HelpContext
End If
   
Err.Clear
    
On Error GoTo 0

End Sub
Public Sub CenterForm(Formu As Form)
'La funcion centra un formulario dentro de un MDI
Dim Top As Integer
On Error GoTo merror
    
Top = (Screen.Height - Formu.Height) / 2

If Formu.MDIChild Then
   Top = Top * 0.4
End If

Formu.Move (Screen.Width - Formu.Width) / 2, Top
   
Exit Sub
merror:
tratarerrores "Error en procedimiento CenterForm"
End Sub

'****************************FUNCIONES DE MENSAJES VARIOS**********************

Public Function MsgP(ByVal Mensaje As String) As Boolean
On Error GoTo merror

'simbolo de stop amarillo con ruido y dos botones(si/no)
If MsgBox(Mensaje, vbYesNo + vbQuestion + vbDefaultButton2, "Mensaje de confirmacion") = vbYes Then
   MsgP = True
Else
   MsgP = False
End If

Exit Function
merror:
tratarerrores "Error en funcion MsgP"
End Function
Public Sub MsgI(ByVal Mensaje As String)
On Error GoTo merror

'signo de exclamacion (!) con ruido
If MsgBox(Mensaje, vbOKOnly + vbInformation, "Mensaje de informacion") Then
   'nada
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento MsgI"
End Sub
Public Sub MsgE(ByVal Mensaje As String)
On Error GoTo merror

'signo de pregunta (?) sin ruido
If MsgBox(Mensaje, vbOKOnly + vbExclamation, "Mensaje de alerta") Then
   'nada
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento MsgE"
End Sub
Public Function SonPesos(ByVal tyCantidad As Currency) As String
'Muestra un importe en letras
Dim lyCantidad As Currency
Dim lyCentavos As Currency
Dim lnDigito As Byte
Dim lnPrimerDigito As Byte
Dim lnSegundoDigito As Byte
Dim lnTercerDigito As Byte
Dim lcBloque As String
Dim lnNumeroBloques As Byte
Dim lnBloqueCero
Dim centavos As Currency
On Error GoTo merror

    lyCantidad = Int(tyCantidad)
    lyCentavos = (tyCantidad - lyCantidad) * 100
    
    laUnidades = Array("UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")
    laDecenas = Array("DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
    laCentenas = Array("CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
    lnNumeroBloques = 1
    Do
    lnPrimerDigito = 0
    lnSegundoDigito = 0
    lnTercerDigito = 0
    lcBloque = ""
    lnBloqueCero = 0
    For I = 1 To 3
    lnDigito = lyCantidad Mod 10
    If lnDigito <> 0 Then
    Select Case I
    Case 1
    lcBloque = " " & laUnidades(lnDigito - 1)
    lnPrimerDigito = lnDigito
    Case 2
    If lnDigito <= 2 Then
    lcBloque = " " & laUnidades((lnDigito * 10) + lnPrimerDigito - 1)
    Else
    lcBloque = " " & laDecenas(lnDigito - 1) & IIf(lnPrimerDigito <> 0, " Y", Null) & lcBloque
    End If
    lnSegundoDigito = lnDigito
    Case 3
    lcBloque = " " & IIf(lnDigito = 1 And lnPrimerDigito = 0 And lnSegundoDigito = 0, "CIEN", laCentenas(lnDigito - 1)) & lcBloque
    lnTercerDigito = lnDigito
    End Select
    Else
    lnBloqueCero = lnBloqueCero + 1
    End If
    lyCantidad = Int(lyCantidad / 10)
    If lyCantidad = 0 Then
    Exit For
    End If
    Next I
    Select Case lnNumeroBloques
    Case 1
    SonPesos = lcBloque
    Case 2
    SonPesos = lcBloque & IIf(lnBloqueCero = 3, Null, " MIL") & SonPesos
    Case 3
    SonPesos = lcBloque & IIf(lnPrimerDigito = 1 And lnSegundoDigito = 0 And lnTercerDigito = 0, " MILLON", " MILLONES") & SonPesos
    End Select
    lnNumeroBloques = lnNumeroBloques + 1
    Loop Until lyCantidad = 0
    SonPesos = Trim(SonPesos)
    
    If lyCentavos >= 1 Then
        SonPesos = SonPesos & " PESOS, CON " & Format(str(lyCentavos), "00") & " CENTAVOS"
    End If

Exit Function
merror:
tratarerrores "Error en funcion SonPesos"
End Function
'****************FUNCIONES DE CARGAR COMBOS******************
Public Sub CargarComboCobradores(ByVal tabla As String, cb As ComboBox, ByVal activos As Boolean, ByVal OrdenNombre As Boolean)
'carga un combo ordenado por el campo predeterminado
Dim rec As rdoResultset
Dim sql As String
Dim Condicion As String
Dim Orden As String
On Error GoTo merror


'Solo activos
If activos Then
   Condicion = " where activo = 'True'"
Else
   'activos y no activos
   Condicion = " where activo = 'True' or activo = 'False'"
End If

If OrdenNombre Then
   Orden = " order by apellido,nombre,predeterminada"
Else
   Orden = " order by predeterminada desc,apellido,nombre"
End If

sql = "select * from " & tabla & Condicion & Orden

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Do While Not rec.EOF
      cb.AddItem rec.rdoColumns(2) & " " & rec.rdoColumns(1)
      cb.ItemData(cb.NewIndex) = rec.rdoColumns(0) '!id
      rec.MoveNext
   Loop
End If


Exit Sub
merror:
tratarerrores "Error en procedimiento CargarComboCobradores"
End Sub
Public Sub CargarComboWhere(ByVal tabla As String, cb As ComboBox, cb1 As ComboBox)
'carga un combo ordenado por el campo predeterminado
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror



sql = "select * from comercios WHERE idprovincia = '" & CLng(cb1.ItemData(cb1.ListIndex)) & "'"

'Limpio el combo comercios
    
        Do While cb.ListCount <> 0
            cb.RemoveItem (0)
        Loop
    
Set rec = cnSQL.OpenResultset(sql)
If Not rec.EOF Then
   Do While Not rec.EOF
      cb.AddItem rec.rdoColumns(1) '!Descripción
      cb.ItemData(cb.NewIndex) = rec.rdoColumns(0) '!id
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento CargarCombowhere"
End Sub
Public Sub CargarCombo2(ByVal tabla As String, cb As ComboBox)
'carga un combo ordenado por el campo predeterminado
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

sql = "select * from " & tabla & " order by nombre Asc"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Do While Not rec.EOF
      cb.AddItem rec.rdoColumns(1) '!Descripción
      cb.ItemData(cb.NewIndex) = rec.rdoColumns(0) '!id
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento CargarCombo2"
End Sub
Public Sub CargarComboUsuarios(ByVal tabla As String, cb As ComboBox)
'carga un combo ordenado por el campo predeterminado
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

sql = "select * from " & tabla & " order by predeterminada desc,usuario"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Do While Not rec.EOF
      cb.AddItem rec.rdoColumns(1) '!Descripción
      cb.ItemData(cb.NewIndex) = rec.rdoColumns(0) '!id
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento CargarComboUsuarios"
End Sub
Public Sub CargarComboProvincias(ByVal tabla As String, cb As ComboBox)
'carga un combo ordenado por el campo predeterminado
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

sql = "select * from " & tabla & " order by nombre Asc"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Do While Not rec.EOF
      cb.AddItem rec.rdoColumns(1) '!Descripción
      cb.ItemData(cb.NewIndex) = rec.rdoColumns(0) '!id
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento CargarComboProvincias"
End Sub

Public Sub CargarComboComercios(cb As ComboBox, IdProvincia As Long)
'carga un combo ordenado por el campo predeterminado
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

sql = "select * from comercios where IDProvincia = " & IdProvincia & " order by nombre Asc"

Set rec = cnSQL.OpenResultset(sql)
cb.Clear
If Not rec.EOF Then
   Do While Not rec.EOF
      cb.AddItem rec.rdoColumns(1) '!Descripción
      cb.ItemData(cb.NewIndex) = rec.rdoColumns(0) '!id
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento CargarComboComercios"
End Sub

Public Sub CargarComboPlanes(cb As ComboBox)
'CARGA SOLO LOS ACTIVOS
'carga un combo ordenado por el campo predeterminado
'usada en registrar creditos
'usada en refinanciar creditos
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

sql = "select * " & _
      "from planes " & _
      "where predeterminada=0 " & _
      "order by cantcuotas"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Do While Not rec.EOF
      cb.AddItem rec.rdoColumns(1) '!Descripción
      cb.ItemData(cb.NewIndex) = rec.rdoColumns(0) '!id
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento CargarComboPlanes"
End Sub
Public Sub LimpiarCampos(frm As Form)
'limpia los campos de una pantalla
Dim ctl As Control
On Error GoTo merror



For Each ctl In frm.Controls
    If TypeOf ctl Is TextBox And ctl.Tag = "" Then ctl.Text = ""
    If TypeOf ctl Is TextBox And ctl.Tag = "N" Then ctl.Text = 0
    If TypeOf ctl Is TextBox And ctl.Tag = "NA" Then ctl.Text = "0,00"
    If TypeOf ctl Is ComboBox And ctl.Tag = "" Then ctl.ListIndex = -1
    If TypeOf ctl Is ListBox And ctl.Tag = "" Then ctl.Clear
    If TypeOf ctl Is CheckBox And ctl.Tag = "" Then ctl.Value = 0
    If TypeOf ctl Is ListView And ctl.Tag = "BORRAR" Then ctl.ListItems.Clear
    If TypeOf ctl Is DTPicker And ctl.Tag = "" Then ctl.Value = Date
Next ctl




Exit Sub
merror:
tratarerrores "Error en procedimiento LimpiarCampos"
End Sub
Public Function ReemplazarEnter(frm As String)
'reemplaza en observaciones de la solapa el fin de linea y comienzo

Dim I As Long
Dim Cadena As String
On Error GoTo merror
 

       If InStr(frm, Chr$(13)) Then
          Cadena = Trim(frm)
          For I = 1 To Len(Cadena)

              If Mid(Cadena, I, 1) = Chr$(13) Then

                 Mid(Cadena, I, 1) = " "

              End If

              If Mid(Cadena, I, 1) = Chr$(10) Then

                 Mid(Cadena, I, 1) = " "

              End If

          Next I

          ReemplazarEnter = Trim(Cadena)

       Else

         ReemplazarEnter = Trim(frm)

       End If
 

Exit Function

merror:

tratarerrores "Error en procedimiento ReemplazarEnter"


End Function
Public Sub ReemplazarComillas(frm As Form)
'reemplaza en los campos Text las comillas no permitidas como por ej: Dell'era
Dim ctl As Control
Dim I As Long
Dim Cadena As String
On Error GoTo merror

For Each ctl In frm.Controls
    If TypeOf ctl Is TextBox And ctl.Tag = "" Then
       If InStr(ctl.Text, "'") Then
          Cadena = Trim(ctl.Text)
          For I = 1 To Len(Cadena)
              If Mid(Cadena, I, 1) = "'" Then
                 Mid(Cadena, I, 1) = " "
              End If
          Next I
          ctl.Text = Trim(Cadena)
       End If
    End If
Next ctl

Exit Sub
merror:
tratarerrores "Error en procedimiento ReemplazarComillas"
End Sub
Public Function ExisteTabla(ByVal tabla As String) As Boolean
'para crear tablas nuevas por codigo al cargar el sistema
'revisa que exista una tabla de la base de datos
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

ExisteTabla = True

   
sql = "SELECT Count(*) FROM " & tabla
        
Set rec = cnSQL.OpenResultset(sql)
    

Exit Function

merror:
ExisteTabla = False
End Function
Public Sub LimpiarTabla(ByVal tabla As String)
'usada para vaciar tablas temporales
'usada en consultardeudores
'usada en importacionrapipago
Dim sql As String
On Error GoTo merror

sql = "delete from " & tabla

cnSQL.Execute (sql)

Exit Sub
merror:
tratarerrores "Error en procedimiento LimpiarTabla: " & Err.Number & " (" & Err.Description & ")"

End Sub

Public Sub LimpiarTablaSP(ByVal tabla As String)
'usada para vaciar tablas temporales
'usada en consultardeudores
'usada en importacionrapipago
Dim sql As String
On Error GoTo merror

sql = "LimpiarTabla " & tabla

cnSQL.Execute (sql)

Exit Sub
merror:
tratarerrores "Error en procedimiento LimpiarTablaSP: " & Err.Number & " (" & Err.Description & ")"

End Sub

Public Function ObtenerCredito(ByVal CodPrestamo As String) As Long
'obtiene el idcredito de un numero de prestamo
'usada en importar ambos
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerCredito = 0

sql = "select idcredito " & _
      "from creditos " & _
      "where codprestamo='" & CStr(CodPrestamo) & "'"
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      ObtenerCredito = CLng(rec.rdoColumns("idcredito"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCredito"
End Function

Public Function ObtenerNombreCliente(ByVal IdCliente As Long) As String
'obtiene el idcredito de un numero de prestamo
'usada en importar ambos
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerNombreCliente = ""

sql = "select Apellido, Nombre " & _
      "from Clientes " & _
      "where IdCliente=" & IdCliente
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
    ObtenerNombreCliente = Trim$(UCase$(rec.rdoColumns("Apellido"))) & ", " & Trim$(UCase$(rec.rdoColumns("Nombre")))
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCredito"
End Function

Public Function ObtenerNumCliente(ByVal IdCredito As Long) As Long
'usado en importarambos
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerNumCliente = 0

sql = "select idcliente " & _
      "from creditos " & _
      "where idcredito='" & CLng(IdCredito) & "'"
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   ObtenerNumCliente = CLng(rec.rdoColumns("idcliente"))
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerNumCliente"
End Function
Public Function ExisteCarpeta() As Boolean
'verifica si una carpeta existe o no..si no existe la crea
'carpeta de exportacion de planillas excel
On Error GoTo merror

ExisteCarpeta = False

MkDir "c:\ExportacionExcel"

ExisteCarpeta = True

Exit Function
merror:
ExisteCarpeta = True
End Function
Public Sub ColorBlanco(frm As Form)
'pone en blanco los campos al editarlos(abms)
Dim ctl As Control
Dim I As Long
On Error GoTo merror

For Each ctl In frm.Controls
    If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
       ctl.BackColor = vbWhite
    End If
Next ctl

Exit Sub
merror:
tratarerrores "Error en procedimiento ColorBlanco"
End Sub
Public Sub ColorCyan(frm As Form)
'pone en cyan los campos al no editarlos(abm)
Dim ctl As Control
Dim I As Long
On Error GoTo merror

For Each ctl In frm.Controls
    If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
       ctl.BackColor = &HFFFFC0
    End If
Next ctl

Exit Sub
merror:
tratarerrores "Error en procedimiento ColorCyan"
End Sub
Public Function ObtenerUltVto1Cred(ByVal IdCredito As Long) As Date
'obtiene el primer vencimiento de la ultima cuota de un credito
'usada para cambiar la fecha de vto de una cuota
Dim sql As String
Dim rec As rdoResultset
Dim Vencimiento1 As Date
On Error GoTo merror

sql = "select fechavencimiento1 " & _
      "from cuotas " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "order by numcuota desc"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Vencimiento1 = CDate(rec.rdoColumns("fechavencimiento1"))
End If

ObtenerUltVto1Cred = Vencimiento1

Exit Function
merror:
tratarerrores "Error en funcion ObtenerUltVto1Cred"
End Function
Public Function ObtenerUltVto2Cred(ByVal IdCredito As Long) As Date
'obtiene el segundo vencimiento de la ultima cuota de un credito
Dim sql As String
Dim rec As rdoResultset
Dim Vencimiento2 As Date
On Error GoTo merror

sql = "select fechavencimiento2 " & _
      "from cuotas " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "order by numcuota desc"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Vencimiento2 = CDate(rec.rdoColumns("fechavencimiento2"))
End If

ObtenerUltVto2Cred = Vencimiento2

Exit Function
merror:
tratarerrores "Error en funcion ObtenerUltVto2Cred"
End Function
Public Function EstaExceptuada(ByVal IdCredito As Long, ByVal NumCuota As Long) As Boolean
'verifica si una cuota esta exceptuada de cobro de mora
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

EstaExceptuada = False

sql = "select idcredito,numcuota,logic1 " & _
      "from cuotas where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      If rec.rdoColumns("logic1") Then
         EstaExceptuada = True
      End If
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion EstaExceptuada"
End Function
Public Function ObtenerSaldoCuotaOKK(ByVal IdCredito As Long, ByVal NumCuota As Long, ByVal FechaVencimiento1 As Date, ByVal FechaVencimiento2 As Date, ByVal Exceptuada As Boolean, ByVal Fecha As Date) As Currency
'calcula el saldo de una cuota individual
Dim ImporteMora As Currency
Dim IvaMora As Currency
Dim Suma As Currency
Dim ImporteParcial As Currency
Dim SaldoCuota1erVenc As Currency
Dim Importe1erVenc As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim SaldoASumar As Currency
Dim SaldoCuota As Currency
Dim IvaACobrarDevuelto  As Currency
On Error GoTo merror

ObtenerSaldoCuotaOKK = 0

ImporteMora = 0
IvaMora = 0
Suma = 0
SaldoASumar = 0

'este saldo es sin mora
SaldoCuota = ObtenerSaldoCuotaX(IdCredito, NumCuota, Fecha, SaldoCuota1erVenc)
Importe1erVenc = ObtenerImporte1erVenc(IdCredito, NumCuota)
If CDate(Fecha) > CDate(FechaVencimiento2) Then
   'calculo la mora de la forma habitual
   'puedo pasarle el campo [exceptuada]
   ImporteMora = CalculoMoraPendiente(IdCredito, NumCuota, Exceptuada, Importe1erVenc, FechaVencimiento1, Fecha, IvaACobrarDevuelto)
   '''''''********ImporteMora = CalcularInteresMoraZZ(Exceptuada, Importe1erVenc, FechaVencimiento1, Fecha)
   IvaMora = 0
   If VG_APLICARIMPUESTOS Then
      If VG_IMPUESTOSCREDIMACO Then
         IvaMora = IvaACobrarDevuelto
      End If
   End If
   '''''''********SoloMoraCobrada = ObtenerMoraCobrada(IdCredito, NumCuota)
   '''''''********SoloIvaMoraCobrada = ObtenerIvaMoraCobrada(IdCredito, NumCuota)
   '''''''********If ImporteMora <= SoloMoraCobrada Then
   '''''''********   ImporteMora = 0
   '''''''********Else
      'si es mayor la mora es solo la diferencia
   '''''''********   ImporteMora = ImporteMora - SoloMoraCobrada
   '''''''********End If
   '''''''********If IvaMora <= SoloIvaMoraCobrada Then
   '''''''********   IvaMora = 0
   '''''''********Else
      'si es mayor la mora es solo la diferencia
   '''''''********   IvaMora = IvaMora - SoloIvaMoraCobrada
   '''''''********End If
   SaldoASumar = SaldoCuota1erVenc
Else
   SaldoASumar = SaldoCuota
End If
Suma = SaldoASumar + ImporteMora + IvaMora
      
ObtenerSaldoCuotaOKK = Suma

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSaldoCuotaOKK"
End Function
Public Function ObtenerCuotasVencidas(ByVal IdCredito As Long, ByVal Fecha As Date) As Long
'obtiene el total de cuotas en mora de un credito a una fecha determinada
'se usa en consultar deudores
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerCuotasVencidas = 0

sql = "ObtenerCuotasVencidas " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("cantidad")) Then
      ObtenerCuotasVencidas = rec.rdoColumns("cantidad")
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCuotasVencidas"
End Function
Public Function ObtenerCuotasPendientes(ByVal IdCredito As Long, ByVal Fecha As Date) As Long
'chequea cuantas cuotas pendientes de un credito al dia de la fecha
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

ObtenerCuotasPendientes = 0

sql = "ObtenerCuotasPendientes " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
            
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull("cantidad") Then
      ObtenerCuotasPendientes = CLng(rec.rdoColumns("cantidad"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCuotasPendientes"
End Function
Public Function ObtenerProvinciaCredito(ByVal IdCredito As Long) As String
'obtiene la provincia del credito
'usada en consultardeudores
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerProvinciaCredito = ""

sql = "ObtenerProvinciaCredito " & CLng(IdCredito)
'sql = "select provincias.nombre as provincia " & _
'      "from provincias inner join creditos on provincias.idprovincia=creditos.idprovincia " & _
'      "where creditos.idcredito='" & CLng(IdCredito) & "'"
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   ObtenerProvinciaCredito = CStr(rec.rdoColumns("provincia")) & vbNullString
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerProvinciaCredito"
End Function
Public Function ObtenerCobrosCredito(ByVal IdCredito As Long) As Currency
'obtiene el total cobrado por un credito por todo concepto(items)
'para consultar creditos y exportacion compleja
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerCobrosCredito = 0

'no discrimina finalizados ni bloqueados
sql = "ObtenerCobrosCredito " & CLng(IdCredito)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("suma")) Then
      ObtenerCobrosCredito = CCur(rec.rdoColumns("suma"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCobrosCredito"
End Function

Public Function ObtenerTotalCredito(ByVal IdCredito As Long) As Double
Dim sql             As String
Dim rec             As rdoResultset
Dim nTotalCredito   As Double
On Error GoTo merror

nTotalCredito = 0

sql = "ObtenerTotalCredito " & CLng(IdCredito)

Set rec = cnSQL.OpenResultset(sql)

Do While Not rec.EOF
    nTotalCredito = nTotalCredito + rec.rdoColumns("ImporteVencimiento1")
    rec.MoveNext
Loop
   
ObtenerTotalCredito = nTotalCredito

Exit Function
merror:
tratarerrores "Error en funcion ObtenerTotalCredito"
End Function


Public Function ConvertirFechaSql(ByVal cFecha As String, ByVal cFormato As String) As String

    
        Select Case cFormato
        Case "DD/MM/YYYY"
        If cAmbiente = "DESA" Then
         ConvertirFechaSql = Mid$(cFecha, 7, 4) & "/" & Mid$(cFecha, 4, 2) & "/" & Mid$(cFecha, 1, 2)
        ElseIf cAmbiente = "DESAW" Then
         ConvertirFechaSql = Mid$(cFecha, 7, 4) & "/" & Mid$(cFecha, 1, 2) & "/" & Mid$(cFecha, 4, 2)
        ElseIf cAmbiente = "DESAN" Then
         ConvertirFechaSql = Mid$(cFecha, 7, 4) & "/" & Mid$(cFecha, 4, 2) & "/" & Mid$(cFecha, 1, 2)
        ElseIf cAmbiente = "TEST" Then
         ConvertirFechaSql = Mid$(cFecha, 7, 4) & "/" & Mid$(cFecha, 4, 2) & "/" & Mid$(cFecha, 1, 2)
        Else
         ConvertirFechaSql = Mid$(cFecha, 7, 4) & "/" & Mid$(cFecha, 4, 2) & "/" & Mid$(cFecha, 1, 2)
        End If
        End Select
     
    
End Function

Public Function ConvertirDblSql(ByVal nNumero As Double) As String

    Dim cNumero     As String
    Dim cNumeroAux  As String
    Dim J           As Integer
    
    cNumeroAux = ""
    cNumero = CStr(nNumero)
    For J = 1 To Len(cNumero)
        If Mid$(cNumero, J, 1) = "," Then
            cNumeroAux = cNumeroAux & "."
        Else
            cNumeroAux = cNumeroAux & Mid$(cNumero, J, 1)
        End If
    Next
    ConvertirDblSql = cNumeroAux
    
End Function

Public Function NuevoCodPrestamo(ByVal IdCliente As Long, cTipoAlta As String) As String
    Dim rec         As rdoResultset
    Dim sql         As String
    Dim cSecuencia  As String
    Dim nSec        As Long
    Dim Numlegajo   As Long
    On Error GoTo merror

    cSecuencia = ""
    
    sql = "SELECT NumLegajo from Clientes WHERE IdCliente = " & IdCliente
    
    Set rec = cnSQL.OpenResultset(sql)
    Numlegajo = CLng(rec.rdoColumns("NumLegajo"))
        
    sql = "SELECT COUNT(*) as CantCreditos FROM Creditos WHERE IdCliente = " & IdCliente
    
    If cTipoAlta = "R" Or cTipoAlta = "C" Then
        sql = sql & " AND logic1 = 0"
    Else
        sql = sql & " AND logic1 = 1"
    End If
          
    Set rec = cnSQL.OpenResultset(sql)
    
    If cTipoAlta = "R" Or cTipoAlta = "C" Then
        nSec = rec.rdoColumns("CantCreditos")
        If nSec > 0 Then
            cSecuencia = cTipoAlta & Format$(nSec, "000")
        End If
        If cTipoAlta = "C" And nSec = 0 Then
            cSecuencia = cTipoAlta
        End If
    Else
        nSec = rec.rdoColumns("CantCreditos") + 1
        cSecuencia = cTipoAlta & Format$(nSec, "000")
    End If
    
    NuevoCodPrestamo = Format$(Numlegajo, "000000") & cSecuencia
    
    Exit Function
merror:
tratarerrores "Error en funcion NuevoCodPrestamo"
End Function

Public Function ObtenerDigitoCUIT(ByVal cPrefijo As String, ByVal cDocumento As String, ByRef bDosDigitos As Boolean) As String
    
    Dim nPosicion       As Integer
    Dim nMultiplicador  As Integer
    Dim nAcumulado      As Integer
    Dim nOnceMenos      As Integer
    Dim nDigito         As Integer
    Dim cCUIT10         As String
    
    nPosicion = 10
    nMultiplicador = 2
    nAcumulado = 0
    cCUIT10 = cPrefijo & cDocumento
    
    While nPosicion > 0
        nAcumulado = nAcumulado + Val(Mid$(cCUIT10, nPosicion, 1)) * nMultiplicador
        nMultiplicador = nMultiplicador + 1
        If nMultiplicador > 7 Then
            nMultiplicador = 2
        End If
        nPosicion = nPosicion - 1
    Wend
    
    nOnceMenos = 11 - (nAcumulado Mod 11)
    nDigito = nOnceMenos
    
    bDosDigitos = False
    If nDigito = 10 Then
        bDosDigitos = True
    End If
    
    If nDigito = 11 Then nDigito = 0
    If nDigito = 10 Then nDigito = 9
    
    ObtenerDigitoCUIT = nDigito
    
End Function

Public Function ObtenerCUIT(ByVal cDocumento As String, ByVal cSexo As String) As String

    Dim cPrefijo    As String
    Dim cDigito     As String
    Dim bDosDigitos As Boolean
    
    ObtenerCUIT = ""
    
    Select Case cSexo
    Case "M"
        cPrefijo = "20"
    Case "F"
        cPrefijo = "27"
    Case Else
        Exit Function
    End Select
    
    cDigito = ObtenerDigitoCUIT(cPrefijo, cDocumento, bDosDigitos)
    
    If bDosDigitos Then
        cPrefijo = "23"
        cDigito = ObtenerDigitoCUIT(cPrefijo, cDocumento, bDosDigitos)
    End If
    
    ObtenerCUIT = cPrefijo & cDocumento & cDigito
    
End Function

Public Sub RefreshTimer()
    minutosLogout = 0
    MDIPrincipal.tmrLogout.Enabled = False
    MDIPrincipal.tmrLogout.Enabled = True
End Sub
