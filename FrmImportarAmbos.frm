VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImportarAmbos 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Cobros RapiPago y PagoFacil (desde planillas Excel)"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "FrmImportarAmbos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdImprimirUltimaImportacion 
      Caption         =   "Imprimir ultima importacion"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      ToolTipText     =   "Imprime la ultima importacion realizada de RapiPago"
      Top             =   7680
      Width           =   2295
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   7680
      Width           =   2295
   End
   Begin VB.CommandButton CmdImportar 
      Caption         =   "Importar Cobros"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Importa los cobros de los archivos seleccionados"
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Contenido del archivo seleccionado:"
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   0
      TabIndex        =   6
      Top             =   4320
      Width           =   8295
      Begin VB.TextBox TxtContador2 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2760
         Width           =   735
      End
      Begin MSComctlLib.ListView Lv2 
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Esta lista muestra las facturas del archivo seleccionado en la lista de arriba"
         Top             =   240
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img2"
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Archivo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Prestamo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "FechaCobro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Imp.Cobrado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Negocio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Recibo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "(*) Esta lista muestra el contenido de las filas del archivo seleccionado en la lista de arriba."
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2160
         TabIndex        =   12
         Top             =   2760
         Width           =   5295
      End
      Begin VB.Label Label3 
         Caption         =   "Cant.Facturas:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones:"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   8295
      Begin VB.TextBox TxtMensaje 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox ComboOpciones 
         Height          =   315
         ItemData        =   "FrmImportarAmbos.frx":058A
         Left            =   120
         List            =   "FrmImportarAmbos.frx":0594
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "(*) Este importador usa siempre la Cta Cte"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lista de archivos:"
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton CmdImprimirContenidos 
         Caption         =   "Imprimir contenidos"
         Height          =   255
         Left            =   6360
         TabIndex        =   15
         ToolTipText     =   "Imprime el contenido del archivo seleccionado en la lista superior"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton CmdVerContenidos 
         Caption         =   "Ver contenido"
         Height          =   255
         Left            =   4440
         TabIndex        =   14
         ToolTipText     =   "Muestra en la lista de abajo el cotenido del archivo seleccionado"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox TxtContador 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3120
         Width           =   735
      End
      Begin MSComctlLib.ListView lv 
         Height          =   2655
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre del archivo"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Archivo Solo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha Importacion"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Usuario"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Total Archivos:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Cobra comenzando con las mas antiguas"
         Top             =   3120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmImportarAmbos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***IMPORTA COBROS DE PAGOFACIL Y RAPIPAGO DESDE ARCHIVOS DE EXCEL
'UBICADOS EN LA CARPETA C:\PAGOFACIL-RAPIPAGO. COBRA CON CTA CTE CUBRIENDO CUOTAS
'EN ASCENDENTE

Private Sub Form_Load()
Call RefreshTimer
TxtContador.Text = 0
ComboOpciones.ListIndex = 0
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
Unload Me
End Sub
Private Sub CargarListaArchivos()
'carga el listview con archivos de c:\rapipago-pagofacil
Dim sql As String
Dim rec As rdoResultset
Dim Archivo As String
Dim Nitem As ListItem
Dim Carpeta As String
Dim Cant As Long
Dim FechaImportacion As String
Dim Usuario As String
On Error GoTo merror

'valido la carpeta
Carpeta = "c:\pagofacil-rapipago\"
Archivo = Dir(Carpeta)
If Trim(Archivo) = "" Then
   MsgE "La carpeta " & Carpeta & " no existe, debe crearla y cargarle las planillas excel de cobros"
   Exit Sub
End If

'busca archivos xls
Archivo = Dir(Carpeta & "*.xls")

lv.ListItems.Clear

Do While Trim(Archivo) <> ""
   Usuario = ""
   FechaImportacion = ObtenerFechaImportacion(Archivo)
   'si existe la tabla historica de archivos procesados
   If ExisteHistorico(Archivo) Then
      If ComboOpciones.Text = "Procesados" Then
         Usuario = ObtenerUsuarioImportacion(Archivo)
         Set Nitem = lv.ListItems.Add(, , Carpeta & Archivo)
         Nitem.SubItems(1) = Archivo
         Nitem.SubItems(2) = FechaImportacion
         Nitem.SubItems(3) = Usuario
      End If
   Else
      'aun no fue procesado
      If ComboOpciones.Text = "No procesados" Then
         Set Nitem = lv.ListItems.Add(, , Carpeta & Archivo)
         Nitem.SubItems(1) = Archivo
         Nitem.SubItems(2) = FechaImportacion
         Nitem.SubItems(3) = Usuario
      End If
   End If
   'obtiene el siguiente archivo
   Archivo = Dir
Loop

TxtContador.Text = lv.ListItems.Count

If lv.ListItems.Count > 0 Then
   If ComboOpciones.Text = "No procesados" Then
      CmdImportar.Enabled = True
   Else
      CmdImportar.Enabled = False
   End If
Else
   CmdImportar.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento CargarListaArchivos"
End Sub
Private Function ObtenerUsuarioImportacion(ByVal Archivo As String) As String
'obtiene el usuario que importo ese archivo
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerUsuarioImportacion = ""

sql = "select usuario from pagofacilhistorico2 " & _
      "where nombrearchivo='" & CStr(Archivo) & "'"
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("usuario")) Then
      ObtenerUsuarioImportacion = CStr(rec.rdoColumns("usuario"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerUsuarioImportacion"
End Function
Private Function ObtenerFechaImportacion(ByVal Archivo As String) As String
'obtiene la fecha de importacion de un archivo del historico
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerFechaImportacion = ""

sql = "select fechaproceso " & _
      "from pagofacilhistorico2 " & _
      "where nombrearchivo='" & CStr(Archivo) & "'"
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   ObtenerFechaImportacion = CStr(rec.rdoColumns("fechaproceso"))
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerFechaImportacion"
End Function
Private Sub CmdImportar_Click()
Dim Mensaje As String
On Error GoTo merror
Call RefreshTimer

CmdImportar.Enabled = False
Me.MousePointer = vbHourglass

Mensaje = "No hubo importacion"

If datosok() Then
   If MsgP("¿Confirma la importacion de los archivos seleccionados?") Then
      'inicio de transaccion
      cnSQL.BeginTrans
      
      'vacio la tabla temporal
      Call LimpiarTabla("pagofacil2")
      Call RefreshTimer

      'si agregue facturas a la tabla temporal
      If AgregarRegistrosArchivos() Then
         'cobro usando la cuenta corriente
         Call CobrarCuotasCTACTE
         Mensaje = "Se finalizo la importacion correctamente"
      Else
         Mensaje = "No hubo importacion...verifique que los archivos no estan importados"
      End If
      
      'fin de transaccion
      cnSQL.CommitTrans
      
      MsgI Mensaje
   End If
End If

'actualizo la lista de archivos
Call CargarListaArchivos

CmdImportar.Enabled = True
Me.MousePointer = vbDefault

Exit Sub
merror:
tratarerrores "Error en boton Importar"
End Sub
Private Function AgregarRegistrosArchivos() As Boolean
'solo agrega a la tabla temporal las facturas de los archivos seleccionados
Dim Filas As Long
Dim I As Long
Dim Detalle As String
Dim ArchivoSolo As String
Dim Mensaje As String
Dim CantReg As Long
Dim cont As Long
Dim Agregue As Boolean
On Error GoTo merror

Agregue = False

Filas = lv.ListItems.Count

cont = 0
Mensaje = ""
'recorro la lista de archivos
For I = 1 To Filas
    'si el archivo esta seleccionado
    If lv.ListItems.Item(I).Checked Then
       ArchivoSolo = lv.ListItems.Item(I).SubItems(1)
       
       'si el archivo no esta procesado de antes
       If Not ExisteHistorico(ArchivoSolo) Then
          'agrego los registros de ese archivo a la tabla pagofacil2
          Call AgregarFacturas(ArchivoSolo)
          'paso el archivo al historico
          Call RegistrarHistorico(ArchivoSolo)
          Agregue = True
       End If
    End If
Next I

AgregarRegistrosArchivos = Agregue

Exit Function
merror:
tratarerrores "Error en funcion AgregarRegistrosArchivos"
End Function
Private Function ExisteHistorico(ByVal Archivo As String) As Boolean
'verifica si un archivo fue procesado con anterioridad
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

ExisteHistorico = False

Archivo = UCase(Trim(Archivo))

sql = "select nombrearchivo " & _
      "from pagofacilhistorico2 " & _
      "where nombrearchivo='" & CStr(Archivo) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("nombrearchivo")) Then
      ExisteHistorico = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteHistorico"
End Function
Private Sub RegistrarHistorico(ByVal Archi As String)
'registra en el historico el archivo recien procesado
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

'verifico si ya esta registrado
If ExisteHistorico(Archi) Then Exit Sub

sql = "insert into pagofacilhistorico2 (nombrearchivo,fechaproceso,usuario) " & _
      "values ('" & CStr(Archi) & "','" & ConvertirFechaSql(Mid$(Now, 1, 10), "DD/MM/YYYY") & "','" & CStr(VG_USUARIOLOGIN) & "')"

cnSQL.Execute sql

Exit Sub
merror:
tratarerrores "Error en procedimiento RegistrarHistorico"
End Sub
Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Orden As Integer
  
If lv.ListItems.Count > 1 Then
   lv.SortKey = ColumnHeader.Index - 1
   Orden = lv.SortKey
   lv.SortOrder = Abs(Not lv.SortOrder = 1)
   lv.Sorted = True
End If

End Sub
Private Sub ComboOpciones_Click()
Call CargarListaArchivos
End Sub
Private Sub lv_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Lv2.ListItems.Clear
TxtContador2.Text = 0
End Sub
Private Sub CmdImprimirContenidos_Click()
Dim I As Long
Dim Archivo As String
On Error GoTo merror
Call RefreshTimer

'imprime el contenido de un archivo seleccionado
Me.MousePointer = vbHourglass
CmdImprimirContenidos.Enabled = False

If DatosOk2() Then
   Archivo = ""
   'tomo el archivo seleccionado
   For I = 1 To lv.ListItems.Count
       If lv.ListItems.Item(I).Checked Then
          Archivo = lv.ListItems.Item(I).SubItems(1)
          Exit For
       End If
   Next I
   'la transaccion esta dentro
   If TraspasoDatos(Archivo) Then
      Call ImprimirContenidos
   End If
End If

CmdImprimirContenidos.Enabled = True
Me.MousePointer = vbDefault
Call RefreshTimer

Exit Sub
merror:
tratarerrores "Error en boton CmdImprimirContenidos"
End Sub
Private Sub ImprimirContenidos()
'imprime el contenido de un archivo seleccionado
Dim sql As String
Dim rec As rdoResultset
Dim Mreporte As New ARContenidoAmbos
On Error GoTo merror

sql = "select * from pagofaciltemp2 " & _
      "order by nombrearchivo,codprestamo"
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Mreporte.RDODataControl1.Resultset = rec
   Mreporte.Show vbModal
Else
   MsgE "No hay datos para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error en funcion ImprimirContenidoArchivos"
End Sub
Private Sub CmdImprimirUltimaImportacion_Click()
On Error GoTo merror
Call RefreshTimer

'imprime la ultima importacion realizada
CmdImprimirUltimaImportacion.Enabled = False
Me.MousePointer = vbHourglass

Call ImprimirUltimaImportacion

Me.MousePointer = vbDefault
CmdImprimirUltimaImportacion.Enabled = True
Call RefreshTimer

Exit Sub
merror:
tratarerrores "Error en boton CmdImprimirUltimaImportacion"
End Sub
Private Sub ImprimirUltimaImportacion()
'imprime la lista de cobros parciales
Dim sql As String
Dim rec As rdoResultset
Dim Archivo As String
Dim Mreporte As New ARUltimoAmbos
Dim Titulo As String
On Error GoTo merror

'si imprimo todos los cobros parciales de la cuota
Titulo = "Ultimos cobros importados de RapiPago Y PagoFacil"

'pasarle el cambio este a mariano!!!
sql = "SELECT * " & _
      "from pagofacil2 " ' & _
'      "order by idcredito,fechacobro"
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Mreporte.RDODataControl1.Resultset = rec
   Mreporte.Caption = "Imprimir ultimos cobros importados de RapiPago Y PagoFacil"
   Mreporte.LabelTitulo = Titulo
   Mreporte.Show vbModal
Else
   MsgE "No hay importaciones previas para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo ultima importacion"
End Sub
Private Function CobrarCuotasCTACTE() As Long
'toma los registros de la tabla rapipago y los procesa
Dim sql As String
Dim rec As rdoResultset
Dim rec2 As rdoResultset
Dim Rec3 As rdoResultset
Dim CapitalRestante As Currency
Dim InteresRestante As Currency
Dim GastoRestante As Currency
Dim SeguroRestante As Currency
Dim OtorgamientoRestante As Currency
Dim IvaInteresRestante As Currency
Dim IvaSeguroRestante As Currency
Dim IvaOtorGastoRestante As Currency
Dim Vencimiento2Restante As Currency
Dim RefinRestante As Currency
Dim CapitalCobrado As Currency
Dim InteresCobrado As Currency
Dim GastosCobrados As Currency
Dim SegurosCobrados As Currency
Dim OtorgamientoCobrado As Currency
Dim MoraCobrada As Currency
Dim IvaMoraCobrada As Currency
Dim IvaInteresCobrado As Currency
Dim IvaSegurosCobrado As Currency
Dim IvaOtorGastosCobrado As Currency
Dim SaldoCuota As Currency
Dim SaldoCuota1erVenc As Currency
Dim CapitalCuota As Currency
Dim InteresCuota As Currency
Dim Vencimiento2Cuota As Currency
Dim RefinCuota As Currency
Dim GastosCuota As Currency
Dim SegurosCuota As Currency
Dim OtorgamientoCuota As Currency
Dim IvaInteresCuota As Currency
Dim IvaSegurosCuota As Currency
Dim IvaOtorGastosCuota As Currency
Dim Vencimiento1 As Date
Dim Vencimiento2 As Date
Dim ImporteParcial As Currency
Dim CodPrestamo As String
Dim IdCredito As Long
Dim NumCuota As Long
Dim ImporteRealCobrado As Currency
Dim ImporteMora As Currency
Dim IvaMora As Currency
Dim IdIngreso As Long
Dim IngresosOk As Boolean
Dim ImporteRemanente As Currency
Dim Cubre As Boolean
Dim IdExcedenteCliente As Long
Dim Importe2 As Currency
Dim IdCliente As Long
Dim ArchiSolo As String
Dim Observaciones As String
Dim DetalleExcedente As String
Dim IdPagoFacil As Long
Dim ImporteTotalRestante As Currency
Dim ImporteGralRestante As Currency
Dim ImporteIngresos As Currency
Dim I As Long
Dim Diferencia As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim FueRapipago As Long
Dim FuePagoFacil As Long
Dim CreditoOk As Boolean
Dim CuotaOk As Boolean
Dim DetalleNew As String
Dim IvaACobrarDevuelto  As Currency
Dim bPrimero As Boolean
Dim nImporteActualizar As Currency
Dim nIdPagoAnterior As Long
Dim IdPagoFacilNuevo As Long
Dim ClienteAnterior As String
Dim NumReciboAnterior As String
Dim ArchivoAnterior As String
Dim Origen As String
Dim FechaImputacion As Date
Dim ImporteTotalRestanteOriginal As Currency
Dim SecuenciaIngreso As Long
Dim MontoCobradoSecuencia As Currency

On Error GoTo merror

'obtengo las facturas ordenadas credito
sql = "select * from pagofacil2"
Set rec = cnSQL.OpenResultset(sql, rdOpenStatic)
Call RefreshTimer

ImporteTotalRestante = 0
ImporteGralRestante = 0
I = 1
If Not rec.EOF Then
    Do While Not rec.EOF
       'obtengo el id del registro
       IdPagoFacil = rec.rdoColumns("idpagofacil")
       'me fijo cual es el credito
       IdCredito = CLng(rec.rdoColumns("idcredito"))
       
       CreditoOk = True
       
       bPrimero = True
       
       'si todo esta ok sigo adelante
       IdCliente = ObtenerClientePrestamo(rec.rdoColumns("codprestamo"))
       
       FuePagoFacil = 1
       FueRapipago = 0
       
         'tomo el importe cobrado
         ImporteCobro = CCur(rec.rdoColumns("importecobro"))
         FechaCobro = CDate(rec.rdoColumns("fechacobro"))
         'ahora le asigno el importecobrado en esta cuota del rp
         'este es el que se ira restando al cubrir cuotas
         ImporteTotalRestante = CCur(ImporteCobro)
         'tomo los datos de cada factura
         IdPagoFacil = CLng(rec.rdoColumns("idpagofacil"))
         DetalleExcedente = ""
         CodPrestamo = rec.rdoColumns("codprestamo")
         ArchiSolo = rec.rdoColumns("nombrearchivo")
         Origen = rec.rdoColumns("origen")
         FechaImputacion = rec.rdoColumns("fechaimportacion")
         
         SecuenciaIngreso = UltimoId("secuencia", "ingresos") + 1
         MontoCobradoSecuencia = ImporteTotalRestante

         
         'tomo todas las cuotas pendientes de ese credito ordenadas en ascendente
         'IMPORTANTE: aca ordeno por vencimiento
         'para que desplace las que les cambiamos el vencimiento hacia el final
'Se modifica por Cobranzas PMC
'         sql = "select cuotas.*,cuotas.logic1 as exceptuada " & _
'               "from creditos inner join cuotas on creditos.idcredito=cuotas.idcredito " & _
'               "where creditos.idcredito=" & CLng(IdCredito) & " " & _
'               "and cuotas.fechacobro is Null " & _
'               "and cuotas.fecharefinanciacion is Null " & _
'               "and cuotas.cuotacomodin = 0 " & _
'               "and creditos.fechafinalizacion is Null " & _
'               "and creditos.fechabloqueo is Null " & _
'               "order by cuotas.fechavencimiento1"
         sql = "select creditos.codprestamo,cuotas.*,cuotas.logic1 as exceptuada " & _
               "from creditos, cuotas " & _
               "where creditos.idcliente=" & CLng(IdCliente) & " " & _
               "and creditos.idcredito = cuotas.idcredito " & _
               "and cuotas.fechacobro is Null " & _
               "and cuotas.fecharefinanciacion is Null " & _
               "and cuotas.cuotacomodin = 0 " & _
               "and creditos.fechafinalizacion is Null " & _
               "and creditos.fechabloqueo is Null " & _
               "order by cuotas.fechavencimiento1,creditos.fechacredito"
         
         Set rec2 = cnSQL.OpenResultset(sql)
         Call RefreshTimer

       
         'si hay cuotas pendientes de ese credito
         If Not rec2.EOF Then
            'recorro la lista de cuotas saldando lo que encuentre sin pagar
            Do While Not rec2.EOF
               'valida que la cuota no sea comodin ni refinanciada
               NumCuota = CLng(rec2.rdoColumns("numcuota"))
               IdCredito = CLng(rec2.rdoColumns("IdCredito"))
               CodPrestamo = rec2.rdoColumns("codprestamo")
               CuotaOk = True
               
               'solo entra si la cuota esta pendiente
               If CuotaOk Then
                  'si queda resto aun para seguir cubriendo cuotas
                  ImporteTotalRestanteOriginal = CCur(ImporteTotalRestante)
                  If CCur(ImporteTotalRestante) > 0 Then
                     ImporteVencimiento1 = CCur(rec2.rdoColumns("importevencimiento1"))
                     ImporteVencimiento2 = CCur(rec2.rdoColumns("importevencimiento2"))
                     Vencimiento1 = CDate(rec2.rdoColumns("fechavencimiento1"))
                     Vencimiento2 = CDate(rec2.rdoColumns("fechavencimiento2"))
                     NumComprobante = CLng(rec2.rdoColumns("numfactura"))
                     'obtengo los campos originales
                     CapitalCuota = CCur(rec2.rdoColumns("importeamortizacion"))
                     InteresCuota = CCur(rec2.rdoColumns("importeinteres"))
                     Vencimiento2Cuota = CCur(rec2.rdoColumns("importerecargovencimiento2"))
                     RefinCuota = CCur(rec2.rdoColumns("importerefinanciacion"))
                     GastosCuota = CCur(rec2.rdoColumns("importegastos"))
                     SegurosCuota = CCur(rec2.rdoColumns("importeseguros"))
                     OtorgamientoCuota = CCur(rec2.rdoColumns("otorgamiento"))
                     IvaInteresCuota = CCur(rec2.rdoColumns("ivainteres"))
                     IvaSegurosCuota = CCur(rec2.rdoColumns("ivaseguros"))
                     IvaOtorGastosCuota = CCur(rec2.rdoColumns("ivaotorgamientogastos"))
                     'obtengo los restantes
                     CapitalRestante = CCur(CapitalCuota) - ObtenerCapitalCobrado(IdCredito, NumCuota)
                     InteresRestante = CCur(InteresCuota) - ObtenerInteresCobrado(IdCredito, NumCuota)
                     Vencimiento2Restante = CCur(Vencimiento2Cuota) - ObtenerVencimiento2Cobrado(IdCredito, NumCuota)
                     RefinRestante = CCur(RefinCuota) - ObtenerRefinCobrado(IdCredito, NumCuota)
                     GastoRestante = CCur(GastosCuota) - ObtenerGastosCobrados(IdCredito, NumCuota)
                     SeguroRestante = CCur(SegurosCuota) - ObtenerSegurosCobrados(IdCredito, NumCuota)
                     OtorgamientoRestante = CCur(OtorgamientoCuota) - ObtenerOtorgamientoCobrado(IdCredito, NumCuota)
                     IvaInteresRestante = CCur(IvaInteresCuota) - ObtenerIvaInteresCobrado(IdCredito, NumCuota)
                     IvaSeguroRestante = CCur(IvaSegurosCuota) - ObtenerIvaSegurosCobrado(IdCredito, NumCuota)
                     IvaOtorGastoRestante = CCur(IvaOtorGastosCuota) - ObtenerIvaOtorGastosCobrado(IdCredito, NumCuota)
                     'veo si tiene cobros parciales de antes
                     ImporteParcial = ObtenerImporteParcialX(IdCredito, NumCuota)
                     Cubre = False
                     SaldoCuota = 0
                     ImporteRemanente = 0
                     ImporteMora = 0
                     IvaMora = 0
                     SaldoCuota = ObtenerSaldoCuotaX(IdCredito, NumCuota, FechaCobro, SaldoCuota1erVenc)
                     'si pago al dia antes del 1º vto
                     If CDate(FechaCobro) <= CDate(Vencimiento1) Then
                        Vencimiento2Cuota = 0
                        'si cubro el importe del vencimiento1
                        If CCur(ImporteTotalRestante) + CCur(ImporteParcial) >= CCur(ImporteVencimiento1) Then
                           Cubre = True
                        End If
                     End If 'fin al 1 vto
          
                     'si pago entre el 1º y 2º vencimiento uso el recargo
                     If CDate(FechaCobro) > CDate(Vencimiento1) And CDate(FechaCobro) <= CDate(Vencimiento2) Then
                        'si cubre el importe del segundo vencimiento
                        If (CCur(ImporteTotalRestante) + CCur(ImporteParcial)) >= CCur(ImporteVencimiento2) Then
                           Cubre = True
                        End If
                     End If 'fin entre 1 2 vto
          
                     'si pago despues del 2º vto
                     If CDate(FechaCobro) > CDate(Vencimiento2) Then
                        'pago despues del 2º vto entonces debo calcular mora
                        'calculo la mora entre la fecha de 2º vto y la fecha de cobro
                        'calculo la mora de forma habitual
                        'puedo pasarle el campo [exceptuada]
                        ImporteMora = CalculoMoraPendiente(rec2.rdoColumns("idcredito"), rec2.rdoColumns("numcuota"), rec2.rdoColumns("exceptuada"), ImporteVencimiento1, Vencimiento1, FechaCobro, IvaACobrarDevuelto)
                        '''''''********ImporteMora = CalcularInteresMoraZZ(rec2.rdoColumns("exceptuada"), SaldoCuota, Vencimiento2, FechaCobro)
                        If VG_APLICARIMPUESTOS Then
                           If VG_IMPUESTOSCREDIMACO Then
                              'calculo el iva de la mora..uso la variable global
                              IvaMora = IvaACobrarDevuelto
                           End If
                        End If
                        '''''''********SoloMoraCobrada = ObtenerMoraCobrada(IdCredito, NumCuota)
                        '''''''********SoloIvaMoraCobrada = ObtenerIvaMoraCobrada(IdCredito, NumCuota)
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
                        SaldoCuota = Round(CCur(SaldoCuota1erVenc) + CCur(ImporteMora) + CCur(IvaMora), 2)
                        If CCur(ImporteTotalRestante) >= CCur(SaldoCuota) Then
                           Cubre = True
                        End If
                     End If 'fin con mora
                     'registro cobro escalonado por items
                     IngresosOk = False
                     'si cubro la cuota, el importe real cobrado es lo que le quedaba de saldo
                     If Cubre Then
                        ImporteRealCobrado = CCur(SaldoCuota)
                        'si tenia cobros parciales
                        If CCur(ImporteParcial) > 0 Then
                           'traigo el importe total cobrado incluyendo este
                           ImporteRealCobrado = CCur(ImporteParcial) + CCur(SaldoCuota)
                        End If 'fin si tenia cobros parciales
                        sql = "update cuotas set rapipago=" & CLng(FueRapipago) & ",pagofacil=" & CLng(FuePagoFacil) & "," & _
                              "fechacobro='" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & "'," & _
                              "importecobrado=" & ConvertirDblSql(CCur(ImporteRealCobrado)) & ",origen='" & Origen & _
                              "' where idcredito=" & CLng(IdCredito) & " and numcuota=" & CLng(NumCuota)
                              cnSQL.Execute sql
                        'estos son para compatibilizar los ingresos
                        'solo grabo las diferencias que es lo realmente cobrado
                        CapitalCobrado = CCur(CapitalRestante)
                        InteresCobrado = CCur(InteresRestante)
                        Vencimiento2Cobrado = CCur(Vencimiento2Restante)
                        RefinCobrado = CCur(RefinRestante)
                        GastosCobrados = CCur(GastoRestante)
                        SegurosCobrados = CCur(SeguroRestante)
                        OtorgamientoCobrado = CCur(OtorgamientoRestante)
                        IvaInteresCobrado = CCur(IvaInteresRestante)
                        IvaSegurosCobrado = CCur(IvaSeguroRestante)
                        IvaOtorGastosCobrado = CCur(IvaOtorGastoRestante)
                        MoraCobrada = CCur(ImporteMora)
                        IvaMoraCobrada = CCur(IvaMora)
                        'debo descontar el importetotalreal para seguir cobrando otras
                        'cuotas que sigan en el proximo ciclo
                        ImporteTotalRestante = CCur(ImporteTotalRestante) - CCur(SaldoCuota)
                        ImporteIngresos = SaldoCuota
                        IngresosOk = True
                    Else ' else si cubre
                        'no cubre ES UN COBRO PARCIAL DE CUOTA
                        ImporteRealCobrado = CCur(ImporteTotalRestante)
                        ImporteIngresos = CCur(ImporteRealCobrado)
                        'si queda resto
                        IvaMoraCobrada = 0
                        If CCur(ImporteRealCobrado) > 0 Then
                           'si hay IVA mora intento cubrirla
                           If CCur(IvaMora) > 0 Then
                              If CCur(ImporteRealCobrado) >= CCur(IvaMora) Then
                                 ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(IvaMora)
                                 IvaMoraCobrada = CCur(IvaMora)
                              Else
                                 IvaMoraCobrada = CCur(ImporteRealCobrado)
                                 ImporteRealCobrado = 0
                              End If
                           End If
                        End If
                        'aca cubro por items y si importa el orden
                        MoraCobrada = 0
                        If CCur(ImporteRealCobrado) > 0 Then
                           If CCur(ImporteMora) > 0 Then
                              If CCur(ImporteRealCobrado) >= CCur(ImporteMora) Then
                                 ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(ImporteMora)
                                 MoraCobrada = CCur(ImporteMora)
                              Else
                                 MoraCobrada = CCur(ImporteRealCobrado)
                                 ImporteRealCobrado = 0
                              End If
                           End If
                        End If
                        Vencimiento2Cobrado = 0
                        If CDate(FechaCobro) > CDate(Vencimiento1) Then
                           If CCur(ImporteRealCobrado) > 0 Then
                              'si sigue intento cubrir
                              If CCur(Vencimiento2Restante) > 0 Then
                                 If CCur(ImporteRealCobrado) >= CCur(Vencimiento2Restante) Then
                                    'saldo la parte de vto2
                                    ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(Vencimiento2Restante)
                                    Vencimiento2Cobrado = CCur(Vencimiento2Restante)
                                 Else
                                    Vencimiento2Cobrado = CCur(ImporteRealCobrado)
                                    ImporteRealCobrado = 0
                                 End If
                              End If
                           End If
                        End If
                        'recargo por refin
                        'si queda aun resto
                        RefinCobrado = 0
                        If CCur(ImporteRealCobrado) > 0 Then
                           'si sigue intento cubrr
                           If CCur(RefinRestante) > 0 Then
                              If CCur(ImporteRealCobrado) >= CCur(RefinRestante) Then
                                 ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(RefinRestante)
                                 RefinCobrado = CCur(RefinRestante)
                              Else
                                 RefinCobrado = CCur(ImporteRealCobrado)
                                 ImporteRealCobrado = 0
                              End If
                           End If
                        End If
                        'si queda aun resto
                        IvaOtorGastosCobrado = 0
                        If CCur(ImporteRealCobrado) > 0 Then
                           'si sigue intento cubrr el ivaotorgastos
                           If CCur(IvaOtorGastoRestante) > 0 Then
                              If CCur(ImporteRealCobrado) >= CCur(IvaOtorGastoRestante) Then
                                 ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(IvaOtorGastoRestante)
                                 IvaOtorGastosCobrado = CCur(IvaOtorGastoRestante)
                              Else
                                 IvaOtorGastosCobrado = CCur(ImporteRealCobrado)
                                 ImporteRealCobrado = 0
                              End If
                           End If
                        End If
                        'si queda aun resto
                        IvaSegurosCobrado = 0
                        If CCur(ImporteRealCobrado) > 0 Then
                           'si sigue intento cubrr el iva seguros
                           If CCur(IvaSeguroRestante) > 0 Then
                              If CCur(ImporteRealCobrado) >= CCur(IvaSeguroRestante) Then
                                 ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(IvaSeguroRestante)
                                 IvaSegurosCobrado = CCur(IvaSeguroRestante)
                              Else
                                 IvaSegurosCobrado = CCur(ImporteRealCobrado)
                                 ImporteRealCobrado = 0
                              End If
                           End If
                        End If
                     'si queda aun resto
                     IvaInteresCobrado = 0
                     If CCur(ImporteRealCobrado) > 0 Then
                        'si sigue intento cubrr el interes
                        If CCur(IvaInteresRestante) > 0 Then
                           If CCur(ImporteRealCobrado) >= CCur(IvaInteresRestante) Then
                              ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(IvaInteresRestante)
                              IvaInteresCobrado = CCur(IvaInteresRestante)
                           Else
                              IvaInteresCobrado = CCur(ImporteRealCobrado)
                              ImporteRealCobrado = 0
                           End If
                        End If
                     End If
                     'si queda aun resto
                      OtorgamientoCobrado = 0
                      If CCur(ImporteRealCobrado) > 0 Then
                         'si sigue intento cubrr el otorgamiento
                         If CCur(OtorgamientoRestante) > 0 Then
                            If CCur(ImporteRealCobrado) >= CCur(OtorgamientoRestante) Then
                               ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(OtorgamientoRestante)
                               OtorgamientoCobrado = CCur(OtorgamientoRestante)
                            Else
                               OtorgamientoCobrado = CCur(ImporteRealCobrado)
                               ImporteRealCobrado = 0
                            End If
                        End If
                      End If
                      'si queda aun resto
                      GastosCobrados = 0
                      If CCur(ImporteRealCobrado) > 0 Then
                         'si sigue intento cubrr los gastos SI ES QUE AUN NO ESTAN CUBIERTOS!!!
                         If CCur(GastoRestante) > 0 Then
                            If CCur(ImporteRealCobrado) >= CCur(GastoRestante) Then
                               ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(GastoRestante)
                               GastosCobrados = CCur(GastoRestante)
                            Else
                               GastosCobrados = CCur(ImporteRealCobrado)
                               ImporteRealCobrado = 0
                            End If
                         End If
                      End If
                      'si queda aun resto
                      SegurosCobrados = 0
                      If CCur(ImporteRealCobrado) > 0 Then
                         'si sigue intento cubrr el seguro
                         If CCur(SeguroRestante) > 0 Then
                            If CCur(ImporteRealCobrado) >= CCur(SeguroRestante) Then
                               ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(SeguroRestante)
                               SegurosCobrados = CCur(SeguroRestante)
                            Else
                               SegurosCobrados = CCur(ImporteRealCobrado)
                               ImporteRealCobrado = 0
                            End If
                         End If
                      End If
                      'si queda aun resto
                      InteresCobrado = 0
                      If CCur(ImporteRealCobrado) > 0 Then
                         'si sigue intento cubrr el interes
                         If CCur(InteresRestante) > 0 Then
                            If CCur(ImporteRealCobrado) >= CCur(InteresRestante) Then
                               ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(InteresRestante)
                               InteresCobrado = CCur(InteresRestante)
                            Else
                               InteresCobrado = CCur(ImporteRealCobrado)
                               ImporteRealCobrado = 0
                            End If
                         End If
                      End If
                      'si queda aun resto
                      CapitalCobrado = 0
                      If CCur(ImporteRealCobrado) > 0 Then
                         'si sigue intento cubrr el iva interes
                         If CCur(CapitalRestante) > 0 Then
                            If CCur(ImporteRealCobrado) >= CCur(CapitalRestante) Then
                               ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(CapitalRestante)
                               CapitalCobrado = CCur(CapitalRestante)
                            Else
                               CapitalCobrado = CCur(ImporteRealCobrado)
                               ImporteRealCobrado = 0
                            End If
                         End If
                      End If
                      
                      'registro el cobro parcial
                      sql = "update cuotas set rapipago=" & CLng(FueRapipago) & "," & _
                            "pagofacil=" & CLng(FuePagoFacil) & "," & _
                            "cobrosparciales=1,origen='" & Origen & "' " & _
                            "where idcredito=" & CLng(IdCredito) & " and numcuota=" & CLng(NumCuota)
                      cnSQL.Execute sql
                      IngresosOk = True
                      'resto el importetotalrestante
                      'no queda mas resto para seguir cobrando otras cuotas
                      ImporteTotalRestante = 0
                  End If 'fin si cubre
                
                  'ahora grabo los items cobrados
                  If IngresosOk Then
                     IdIngreso = UltimoId("idingreso", "ingresos") + 1
                     
                     sql = "insert into ingresos (idingreso,idcredito,numcuota," & _
                           "fechacobro,importecobrado,codprestamo,numcomprobante," & _
                           "capitalcobrado,interescobrado,vencimiento2cobrado,refincobrado," & _
                           "gastoscobrados,seguroscobrados,otorgamientocobrado,ivainterescobrado," & _
                           "ivaseguroscobrado,ivaotorgastoscobrado,moracobrada,ivamoracobrada,usuario,pagofacil,rapipago,origen,fechaimputacion,secuencia,montocobradosecuencia) " & _
                           "values(" & CLng(IdIngreso) & "," & CLng(IdCredito) & "," & CLng(NumCuota) & _
                           ",'" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & "'," & ConvertirDblSql(CCur(ImporteIngresos)) & _
                           ",'" & CStr(CodPrestamo) & "'," & CLng(NumComprobante) & _
                           "," & ConvertirDblSql(CCur(CapitalCobrado)) & "," & ConvertirDblSql(CCur(InteresCobrado)) & _
                           "," & ConvertirDblSql(CCur(Vencimiento2Cobrado)) & "," & ConvertirDblSql(CCur(RefinCobrado)) & _
                           "," & ConvertirDblSql(CCur(GastosCobrados)) & "," & ConvertirDblSql(CCur(SegurosCobrados)) & _
                           "," & ConvertirDblSql(CCur(OtorgamientoCobrado)) & "," & ConvertirDblSql(CCur(IvaInteresCobrado)) & _
                           "," & ConvertirDblSql(CCur(IvaSegurosCobrado)) & "," & ConvertirDblSql(CCur(IvaOtorGastosCobrado)) & _
                           "," & ConvertirDblSql(CCur(MoraCobrada)) & "," & ConvertirDblSql(CCur(IvaMoraCobrada)) & _
                           ",'" & CStr(VG_USUARIOLOGIN) & "'," & CLng(FuePagoFacil) & "," & CLng(FueRapipago) & ",'" & Origen & "','" & ConvertirFechaSql(FechaImputacion, "DD/MM/YYYY") & "'," & SecuenciaIngreso & "," & ConvertirDblSql(MontoCobradoSecuencia) & ")"
                     cnSQL.Execute sql
                                          
                     If bPrimero Then
                        bPrimero = False
                        nImporteActualizar = CCur(ImporteIngresos)
                        nIdPagoAnterior = IdPagoFacil
                        ClienteAnterior = rec.rdoColumns("Cliente")
                        NumReciboAnterior = rec.rdoColumns("Recibo")
                        ArchivoAnterior = rec.rdoColumns("NombreArchivo")
                     Else
                        
                        sql = "update pagofacil2 set importecobro = " & ConvertirDblSql(nImporteActualizar) & _
                              "where idpagofacil = " & nIdPagoAnterior
                        cnSQL.Execute sql
                        
                        IdPagoFacilNuevo = UltimoId("idpagofacil", "pagofacil2") + 1
                        sql = "insert into pagofacil2(idpagofacil,cliente,idcredito,codprestamo," & _
                              "fechacobro,importecobro,nombrearchivo,fechaimportacion," & _
                              "negocio,recibo) " & _
                              "values(" & CLng(IdPagoFacilNuevo) & ",'" & ClienteAnterior & "'," & _
                              CLng(IdCredito) & ",'" & CStr(CodPrestamo) & "'," & _
                              "'" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & "'," & ConvertirDblSql(CCur(ImporteTotalRestanteOriginal)) & "," & _
                              "'" & CStr(ArchivoAnterior) & "',GetDate()," & _
                              "'PMC','" & NumReciboAnterior & "')"
                        cnSQL.Execute sql
                        
                        nImporteActualizar = CCur(ImporteIngresos)
                        nIdPagoAnterior = IdPagoFacilNuevo
                     End If
                     
                     If VG_FINALIZARAUTOMATICAMENTE Then
                        'si es la ultima cuota finalizo el credito
                        Call FinalizarCredito(IdCredito, Date)
                     End If
                  
                  End If 'fin ingesosok
                End If 'si importetotalrestante
              
              End If  'si la cuotaok
              
              rec2.MoveNext
              Call RefreshTimer

           Loop 'del rec2
           
           'me fijo si quedo algo del credito recien procesado
           If CCur(ImporteTotalRestante) > 0 Then
              Observaciones = "Excedente Cta Cte"
              IdExcedenteCliente = UltimoId("idexcedentecliente", "excedentesclientes") + 1
              'en esta importacion no hay numcuota
              NumCuota = 0
              sql = "insert into excedentesclientes (idexcedentecliente,idcliente,idcredito," & _
                    "codprestamo,numcuota,fechacobro,importecobro,rapipago,pagofacil,archivorp,observaciones,origen,fechaimputacion) " & _
                    "values(" & CLng(IdExcedenteCliente) & "," & CLng(IdCliente) & _
                    "," & CLng(IdCredito) & ",'" & CStr(CodPrestamo) & "'," & CLng(NumCuota) & _
                    ",'" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & "'," & ConvertirDblSql(CCur(ImporteTotalRestante)) & _
                    "," & CLng(FueRapipago) & "," & CLng(FuePagoFacil) & ",'" & CStr(ArchiSolo) & "','" & CStr(Observaciones) & "','" & Origen & "','" & ConvertirFechaSql(FechaImputacion, "DD/MM/YYYY") & "')"
              cnSQL.Execute sql
              DetalleExcedente = "Excedente de cta cte"
              'agrego el detalle de excedente a rapipago
              sql = "update pagofacil2 set excedentes='" & CStr(DetalleExcedente) & "' " & _
                    "where idpagofacil=" & CLng(IdPagoFacil)
              cnSQL.Execute sql
           End If
       Else 'si rec2 es eof es porque no hay cuotas en el sql
           
           '***
           If CCur(ImporteTotalRestante) > 0 Then
              'en esta importacion no hay numcuota(solo hay idcredito)
              NumCuota = 0
           
              DetalleNew = "Excedente de cta cte"
              If Not ExisteCredito(IdCredito) Then
                 DetalleNew = "Excedente, credito no existe"
              End If
              If CreditoBloqueado1(IdCredito) Then
                 DetalleNew = "Excedente, credito bloqueado"
              End If
              If CreditoFinalizado(IdCredito) Then
                 DetalleNew = "Excedente, credito finalizado"
              End If
       
              IdExcedenteCliente = UltimoId("idexcedentecliente", "excedentesclientes") + 1
              sql = "insert into excedentesclientes (idexcedentecliente,idcliente,idcredito," & _
                    "codprestamo,numcuota,fechacobro,importecobro,rapipago,pagofacil,archivorp,observaciones,origen,fechaimputacion) " & _
                    "values(" & CLng(IdExcedenteCliente) & "," & CLng(IdCliente) & _
                    "," & CLng(IdCredito) & ",'" & CStr(CodPrestamo) & "'," & CLng(NumCuota) & _
                    ",'" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & "'," & ConvertirDblSql(CCur(ImporteTotalRestante)) & _
                    "," & CLng(FueRapipago) & "," & CLng(FuePagoFacil) & ",'" & CStr(ArchiSolo) & "','" & CStr(DetalleNew) & "','" & Origen & "','" & ConvertirFechaSql(FechaImputacion, "DD/MM/YYYY") & "')"
              cnSQL.Execute sql
             
              'agrego el detalle de excedente a rapipago
              sql = "update pagofacil2 set excedentes='" & CStr(DetalleNew) & "' " & _
                    "where idpagofacil=" & CLng(IdPagoFacil)
                    cnSQL.Execute sql
           End If
       End If  'si el credito no tiene cuotas pendientes
    
  I = I + 1
  rec.MoveNext
  Call RefreshTimer

Loop 'del while

End If ' si no es eof del primer if

Exit Function
merror:
tratarerrores "Error en funcion CobrarCuotasCtaCte"
End Function
Private Sub CmdVerContenidos_Click()
'permite ver el contenido de un archivo seleccionado
Dim I As Long
Dim Archivo As String
Dim Archi As String
On Error GoTo merror
Call RefreshTimer

CmdVerContenidos.Enabled = False
Me.MousePointer = vbHourglass

Lv2.ListItems.Clear

'si todo esta ok
If DatosOk2() Then
   Archivo = ""
   'tomo el archivo seleccionado
   For I = 1 To lv.ListItems.Count
       If lv.ListItems.Item(I).Checked Then
          Archivo = lv.ListItems.Item(I).SubItems(1)
          Exit For
       End If
   Next I
   
   'traspaso datos de la planilla excel seleccionada a la tabla
   If TraspasoDatos(Archivo) Then
      'muestro el contenido en la segunda lista de la pantalla
      Call MostrarContenidos
   End If
End If

CmdVerContenidos.Enabled = True
Me.MousePointer = vbDefault
Call RefreshTimer

Exit Sub
merror:
tratarerrores "Error en boton VerContenidos"
End Sub
Private Sub AgregarFacturas(ByVal Archivo As String)
'viene del importador viejo de pagofacil y saca datos de excel
'a tabla access carga las facturas a una tabla para que despues
'las procese la otra funcion
Dim MiExcel As Excel.APPLICATION ' Aplicacion Excel
Dim MiLibro As Excel.Workbook ' Un libro de Excel
Dim MiHoja As Excel.Worksheet ' Una hoja de Excel
Dim sql As String
Dim CodPrestamo As String
Dim IdCredito As Long
Dim NumCuota As Long
Dim IdCliente As Long
Dim NumComprobante As Long
Dim ImporteCobro As Currency
Dim FechaCobro As Date
Dim Archi As String
Dim I As Integer
Dim Cadena As String
Dim Ok As Boolean
Dim CadCuota As String
Dim Pos As Long
Dim Concepto As String
Dim NumRecibo As String
Dim IdPagoFacil As Long
Dim Cliente As String
Dim Negocio As String
Dim FechaImportacion As String
Dim CadPrestamo As String
Dim CadCliente As String
Dim CadFechaCobro As String
Dim CadImporteCobro As String
Dim CadNroDocumento As String
Dim CadOrigen As String
Dim Origen As String
On Error GoTo merror

'verifico si ya existe
Archi = Dir("C:\PAGOFACIL-RAPIPAGO\" & Archivo)

Archi = UCase(Archi)

If Trim(Archi) = "" Then
   MsgE "La planilla " & Archivo & " no existe..debe estar dentro de la carpeta en C:\PAGOFACIL-RAPIPAGO\"
   Exit Sub
End If

Set MiExcel = New Excel.APPLICATION
'oculto la aplicacion excel para que sea todo el proceso en background
MiExcel.Visible = False

'abro un libro existente
Set MiLibro = MiExcel.Workbooks.Open("C:\PAGOFACIL-RAPIPAGO\" & Archivo)
 
'asigno la primera hoja por defecto
Set MiHoja = MiLibro.Worksheets(1)
  
'me paro en el principio del primer credito
MiHoja.Range("B3").Activate
Cadena = Trim(CStr(MiHoja.Range("A2").Value))

I = 2

If Trim(Cadena) = "" Then
   MsgE "El primer prestamo de la planilla esta en blanco...revise la planilla (Celda A2)"
   Exit Sub
End If

FechaImportacion = CStr(Date)
'recorro la planilla tomando cada factura
Do While Cadena <> ""
   Ok = True
  
   'valido el credito ahora es codprestamo
   'si es blanco
   CodPrestamo = ""
   IdCredito = 0
   CadNroDocumento = Trim$(MiHoja.Range("A" + Trim(str(I))).Value)
   
   If CadNroDocumento = "" Then
      IdCredito = 0
      Ok = False
   Else
       If Val(CadNroDocumento) = CadNroDocumento Then
            CadPrestamo = ObtenerCodPrestamoConDocumento(CadNroDocumento)
            If CadPrestamo = "" Then
                CadPrestamo = ObtenerCodPrestamoConDocumentoInclFinalizados(CadNroDocumento)
                If CadPrestamo = "" Then
                    IdCredito = 0
                    Ok = False
                End If
            End If
       Else
           IdCredito = 0
           Ok = False
       End If
   End If
   
   CadPrestamo = UCase(Trim(CadPrestamo))
   If CadPrestamo = "" Then
      IdCredito = 0
      Ok = False
   Else
      'hay un dato
      'obtengo el codigo de prestamo similar a FG002527
      CodPrestamo = CadPrestamo
      'obtengo el idcredito de ese prestamo
      IdCredito = ObtenerCredito(CodPrestamo)
   End If
   
   IdCliente = ObtenerNumCliente(IdCredito)
    
   IdCredito = ObtenerCredito(CodPrestamo)
   
   'obtengo el cliente
   CadCliente = ObtenerNombreCliente(IdCliente)
   CadCliente = UCase(Trim(CadCliente))
   If CadCliente = "" Then
      Cliente = ""
      Ok = False
   Else
      Cliente = CadCliente
   End If
   
   'obtengo la fecha de cobro
   'valido la fecha
   CadFechaCobro = MiHoja.Range("D" + Trim(str(I))).Value
   CadFechaCobro = Trim(CadFechaCobro)
    
   'si no es date
   If Not IsDate(CadFechaCobro) Then
      FechaCobro = Date + 1
      Ok = False
   End If
    
   'obtengo el importe cobrado
   CadImporteCobro = MiHoja.Range("E" + Trim(str(I))).Value
   CadImporteCobro = Trim(CadImporteCobro)
   If CadImporteCobro = "" Then
      ImporteCobro = 0
      Ok = False
   End If
  
   'si no es numerico
   If Not IsNumeric(CadImporteCobro) Then
      ImporteCobro = 0
      Ok = False
   Else
      'es numerico
      ImporteCobro = CCur(CadImporteCobro)
   End If
   
   If CCur(ImporteCobro) > 0 Then
      ImporteCobro = CCur(ImporteCobro)
   Else
      ImporteCobro = 0
      Ok = False
   End If
  
   CadOrigen = MiHoja.Range("H" + Trim(str(I))).Value
    If CadOrigen = "PC" Or CadOrigen = "HB" Or CadOrigen = "S1" Then
        Origen = "PMC"
    Else
        Origen = "ANTICIPO"
    End If
    
   NumRecibo = MiHoja.Range("I" + Trim(str(I))).Value
   
   'si todos los datos de la factura estan ok los grabo en la tabla
   If Ok Then
      IdPagoFacil = UltimoId("idpagofacil", "pagofacil2") + 1
      'agrego la factura al archivo de pagofacil
      sql = "insert into pagofacil2(idpagofacil,cliente,idcredito,codprestamo," & _
            "fechacobro,importecobro,nombrearchivo,fechaimportacion," & _
            "negocio,recibo,origen) " & _
            "values(" & CLng(IdPagoFacil) & ",'" & CStr(Cliente) & "'," & _
            CLng(IdCredito) & ",'" & CStr(CodPrestamo) & "'," & _
            "'" & ConvertirFechaSql(CadFechaCobro, "DD/MM/YYYY") & "'," & ConvertirDblSql(CCur(ImporteCobro)) & "," & _
            "'" & CStr(Archivo) & "',GetDate()," & _
            "'PMC','" & NumRecibo & "','" & Origen & "')"
      cnSQL.Execute sql
      Call RefreshTimer

   End If
  
   I = I + 1
   'saco otra cadena siguiente
   Cadena = Trim(CStr(MiHoja.Range("A" + Trim(str(I))).Value))
Loop

'cierra el libro actual
MiLibro.Close
  
'cierro excel
MiExcel.Quit

Set MiExcel = Nothing

Exit Sub
merror:
tratarerrores "Error cargando registros de excel a PagoFacil-RaPiPago"
End Sub
Private Function TraspasoDatos(ByVal Archivo As String) As Boolean
'traspasa datos de una planilla seleccionada a la tabla
'luego se puede usar la tabla cargada para
'imprimir contenidos
'mostrar contenidos
'incluso para cobrar
Dim sql As String
Dim rec As rdoResultset
Dim MiExcel As Excel.APPLICATION ' Aplicacion Excel
Dim MiLibro As Excel.Workbook ' Un libro de Excel
Dim MiHoja As Excel.Worksheet ' Una hoja de Excel
Dim Archi As String
Dim HuboProcesos As Boolean
Dim I As Long
Dim Cadena As String
Dim IdCredito As Long
Dim CodPrestamo As String
Dim NumRecibo As String
Dim Cliente As String
Dim ImporteCobro As Currency
Dim FechaCobro As Date
Dim FechaImportacion As String
Dim IdCliente As Long
Dim Ok As Boolean
Dim CadPrestamo As String
Dim CadCliente As String
Dim CadFechaCobro As String
Dim CadImporteCobro As String
Dim CadNumRecibo As String
Dim CadNroDocumento  As String
Dim Mreporte As New ARContenidoAmbos
On Error GoTo merror

TraspasoDatos = False

'verifico si el archivo existe
Archi = Dir("C:\PAGOFACIL-RAPIPAGO\" & Archivo)

If Trim(Archi) = "" Then
   MsgE "La planilla " & Archivo & " no existe..debe estar en la carpeta C:\PAGOFACIL-RAPIPAGO\"
   Exit Function
End If

Set MiExcel = New Excel.APPLICATION
'oculto la aplicacion excel para que sea todo el proceso en background
MiExcel.Visible = False

'abro un libro existente
Set MiLibro = MiExcel.Workbooks.Open("C:\PAGOFACIL-RAPIPAGO\" & Archivo)
 
'asigno la primera hoja por defecto
Set MiHoja = MiLibro.Worksheets(1)
    
'me paro en el principio del primer credito
MiHoja.Range("B3").Activate

Cadena = Trim(CStr(MiHoja.Range("A2").Value))

I = 2
If Trim(Cadena) = "" Then
   MsgI "El primer prestamo de la planilla esta en blanco...revise la planilla (Celda A2)"
   Exit Function
End If

'inicio transaccion
cnSQL.BeginTrans

'primero blanqueo la tabla temporal
Call LimpiarTabla("pagofaciltemp2")

FechaImportacion = ""

HuboProcesos = False
'recorro una sola planilla tomando cada factura
Do While Cadena <> ""
   'supongo que los datos de cada factura estan correctos
   Ok = True
  
   'valido el credito ahora es codprestamo
   CodPrestamo = ""
   IdCredito = 0
   
   CadNroDocumento = Trim$(MiHoja.Range("A" + Trim(str(I))).Value)
   
   If CadNroDocumento = "" Then
      IdCredito = 0
      Ok = False
   Else
       If Val(CadNroDocumento) = CadNroDocumento Then
            CadPrestamo = ObtenerCodPrestamoConDocumento(CadNroDocumento)
            If CadPrestamo = "" Then
                CadPrestamo = ObtenerCodPrestamoConDocumentoInclFinalizados(CadNroDocumento)
                If CadPrestamo = "" Then
                    IdCredito = 0
                    Ok = False
                End If
            End If
       Else
           IdCredito = 0
           Ok = False
       End If
   End If
   
   CadPrestamo = UCase(Trim(CadPrestamo))
   If CadPrestamo = "" Then
      IdCredito = 0
      Ok = False
   Else
      'hay un dato
      'obtengo el codigo de prestamo similar a FG002527
      CodPrestamo = CadPrestamo
      'obtengo el idcredito de ese prestamo
      IdCredito = ObtenerCredito(CodPrestamo)
      If IdCredito > 0 Then
         FechaImportacion = CStr(Date)
      Else
         FechaImportacion = ""
      End If
   End If
   
   IdCliente = ObtenerNumCliente(IdCredito)
    
   'obtengo el cliente
   CadCliente = ObtenerNombreCliente(IdCliente)
   CadCliente = UCase(Trim(CadCliente))
   If CadCliente = "" Then
      Cliente = ""
      Ok = False
   Else
      Cliente = CadCliente
   End If
     
   'valido la fecha de cobro
   CadFechaCobro = MiHoja.Range("D" + Trim(str(I))).Value
   CadFechaCobro = Trim(CadFechaCobro)
    
   'si no es date
   If Not IsDate(CadFechaCobro) Then
      FechaCobro = Date + 1
      Ok = False
   End If
    
   'obtengo el importe cobrado
   CadImporteCobro = MiHoja.Range("E" + Trim(str(I))).Value
   CadImporteCobro = Trim(CadImporteCobro)
   
   If CadImporteCobro = "" Then
      ImporteCobro = 0
      Ok = False
   End If
   
   'si no es numerico
   If Not IsNumeric(CadImporteCobro) Then
      ImporteCobro = 0
      Ok = False
   Else
      'es numerico
      ImporteCobro = CCur(CadImporteCobro)
   End If
   
   'esto es por si ingresaron un negativo o cero
   If CCur(ImporteCobro) > 0 Then
      ImporteCobro = CCur(ImporteCobro)
   Else
      ImporteCobro = 0
      Ok = False
   End If
  
   'obtengo el recibo
   CadNumRecibo = MiHoja.Range("I" + Trim(str(I))).Value
   
   'si todos los datos de la factura estan ok los grabo en la tabla
   If Ok Then
      'agrego la factura al archivo de pagofacil
      sql = "insert into pagofaciltemp2(nombrearchivo,negocio,cliente,codprestamo," & _
            "fechacobro,importecobro,recibo) " & _
            "values('" & CStr(Archivo) & "','PMC'," & _
            "'" & CStr(Cliente) & "','" & CStr(CodPrestamo) & "'," & _
            "'" & ConvertirFechaSql(CadFechaCobro, "DD/MM/YYYY") & "'," & ConvertirDblSql(CCur(ImporteCobro)) & "," & _
            "'" & CStr(CadNumRecibo) & "')"
      cnSQL.Execute sql
      HuboProcesos = True
   End If
  
   I = I + 1
   'saco otra cadena siguiente
   Cadena = Trim(CStr(MiHoja.Range("A" + Trim(str(I))).Value))
Loop

'fin de transaccion
cnSQL.CommitTrans

'cierra el libro actual
MiLibro.Close
  
'cierro excel
MiExcel.Quit

Set MiExcel = Nothing

If HuboProcesos Then
   TraspasoDatos = True
End If

Exit Function
merror:
tratarerrores "Error en funcion TraspasoDatos"
End Function
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If lv.ListItems.Count = 0 Then
   datosok = False
   MsgE "No hay archivos en la lista"
   Exit Function
End If

If Not HayFilasChequeadas(lv) Then
   datosok = False
   MsgE "Debe seleccionar archivos en la lista"
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosOk"
End Function
Private Function DatosOk2() As Boolean
DatosOk2 = True

'si la lista esta vacia no hago nada
If lv.ListItems.Count() = 0 Then
   MsgE "No hay archivos en la lista"
   DatosOk2 = False
   Exit Function
End If

'si no hay items tildados no hago nada
If Not HayFilasChequeadas(lv) Then
   MsgE "Debe marcar archivos"
   DatosOk2 = False
   Exit Function
End If

'solo admito una sola
If HayMasChequeadas(lv) Then
   MsgE "Debe seleccionar un solo archivo"
   DatosOk2 = False
   Exit Function
End If

End Function
Private Sub MostrarContenidos()
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

sql = "select * from pagofaciltemp2 order by codprestamo"
Set rec = cnSQL.OpenResultset(sql)

Lv2.ListItems.Clear
If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = Lv2.ListItems.Add(, , Format(rec.rdoColumns("nombrearchivo"), "0000"))
      Nitem.SubItems(1) = rec.rdoColumns("cliente") & vbNullString
      Nitem.SubItems(2) = rec.rdoColumns("codprestamo") & vbNullString
      Nitem.SubItems(3) = rec.rdoColumns("fechacobro") & vbNullString
      Nitem.SubItems(4) = rec.rdoColumns("importecobro") & vbNullString
      Nitem.SubItems(5) = rec.rdoColumns("negocio") & vbNullString
      Nitem.SubItems(6) = rec.rdoColumns("recibo") & vbNullString
        
      rec.MoveNext
   Loop
   TxtContador2.Text = Lv2.ListItems.Count
End If

Exit Sub
merror:
tratarerrores "Error en funcion MostrarContenidos"
End Sub
Private Function ObtenerClientePrestamo(ByVal CodPrestamo As String) As Long
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerClientePrestamo = 0

sql = "select idcliente from creditos where codprestamo='" & CStr(CodPrestamo) & "'"
Set rec = cnSQL.OpenResultset(sql)
If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcliente")) Then
      ObtenerClientePrestamo = rec.rdoColumns("idcliente")
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerClientePrestamo"
End Function
