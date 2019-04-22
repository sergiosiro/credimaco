VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImportarRapiPago 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Cobros - RapiPago"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "FrmImportarRapiPago.frx":0000
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
      Left            =   4080
      TabIndex        =   18
      ToolTipText     =   "Imprime la ultima importacion realizada de RapiPago"
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir contenidos"
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      ToolTipText     =   "Imprime el contenido de los archivos seleccionados en la lista superior"
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton CmdImportar 
      Caption         =   "Importar Cobros"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "Importa los cobros de los archivos seleccionados"
      Top             =   7680
      Width           =   1935
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
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2760
         Width           =   735
      End
      Begin MSComctlLib.ListView Lv2 
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Esta lista muestra las facturas del archivo seleccionado en la lista de arriba"
         Top             =   240
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   4471
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha Cobro"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Imp.Cobrado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CodigoBarras"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Empresa"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente Nº"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cupon Nº"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Importe 1º Vto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Fecha 1º Vto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Imp.Recargo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Dias al 2º Vto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Importe 2º Vto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Fecha 2º Vto"
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
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   3300
      Width           =   8295
      Begin VB.TextBox TxtInicio 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox ComboOpciones 
         Height          =   315
         ItemData        =   "FrmImportarRapiPago.frx":0442
         Left            =   120
         List            =   "FrmImportarRapiPago.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label LabelMensaje 
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5160
         TabIndex        =   19
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "(Recomendada 76)"
         Height          =   255
         Left            =   3360
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Posicion de inicio de lectura:"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lista de archivos:"
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8295
      Begin VB.TextBox TxtContador 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2880
         Width           =   735
      End
      Begin MSComctlLib.ListView lv 
         Height          =   2535
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4471
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha del archivo"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre del archivo"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Archivo Solo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha Importacion"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Usuario"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "(*) Para importar cobros debe marcar archivos en la lista."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   2880
         Width           =   6015
      End
      Begin VB.Label Label1 
         Caption         =   "Total Archivos:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Cobra comenzando con las mas antiguas"
         Top             =   2880
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmImportarRapiPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***IMPORTA COBROS DE RAPIPAGO TOMANDO ARCHIVOS UBICADOS EN LA CARPETA
'C:\RAPIPAGO. COBRA CON CTA CTE CUBRIENDO CUOTAS IMPAGAS EN ASCENDENTE

Private Sub Form_Load()
Call RefreshTimer
TxtContador.Text = 0
ComboOpciones.ListIndex = 0
'inicio de lectura dentro del archivo
TxtInicio.Text = VG_INICIORP
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
Unload Me
End Sub
Private Sub CargarListaArchivos()
'carga el listview con archivos de c:\rapipago
Dim sql As String
Dim rec As rdoResultset
Dim Archivo As String
Dim Fecha As Date
Dim Nitem As ListItem
Dim Carpeta As String
Dim Cant As Long
Dim FechaImportacion As String
Dim Usuario As String
On Error GoTo merror

Carpeta = "c:\rapipago\"

Archivo = Dir("C:\RAPIPAGO\")

If Trim(Archivo) = "" Then
   MsgE "La carpeta RAPIPAGO no existe, debe crearla y cargarle los archivos enviados por RAPIPAGO"
   Exit Sub
End If

'busca archivos de la forma RP*.1032
Archivo = Dir("c:\rapipago\RP*." & CStr(VG_CODIGOAUTOMATICO))
lv.ListItems.Clear

Do While Trim(Archivo) <> ""
   Usuario = ""
   'obtengo la fecha del archivo
    Fecha = ObtenerFechaArchivoRP(Archivo)
   FechaImportacion = ObtenerFechaImportacion(Archivo)
   'si existe en la tabla historica de archivos procesados
   If ExisteRPHistorico(Archivo) Then
      If ComboOpciones.Text = "Procesados" Then
         Usuario = ObtenerUsuarioProceso(Archivo)
         Set Nitem = lv.ListItems.Add(, , Fecha)
         Nitem.SubItems(1) = Carpeta & Archivo
         Nitem.SubItems(2) = Archivo
         Nitem.SubItems(3) = FechaImportacion
         Nitem.SubItems(4) = Usuario
      End If
   Else
      'si no esta en el historico aun no fue procesado
      If ComboOpciones.Text = "No procesados" Then
         Set Nitem = lv.ListItems.Add(, , Fecha)
         Nitem.SubItems(1) = Carpeta & Archivo
         Nitem.SubItems(2) = Archivo
         Nitem.SubItems(3) = FechaImportacion
         Nitem.SubItems(4) = Usuario
      End If
   End If
   'sigue cargando el siguiente archivo de esa carpeta
   Archivo = Dir
Loop


'busca archivos de la forma RP*.1032
Archivo = Dir("c:\rapipago\SF*." & CStr(VG_CODIGOAUTOMATICO))

Do While Trim(Archivo) <> ""
   Usuario = ""
   'obtengo la fecha del archivo
    Fecha = ObtenerFechaArchivoRP(Archivo)
   FechaImportacion = ObtenerFechaImportacion(Archivo)
   'si existe en la tabla historica de archivos procesados
   If ExisteRPHistorico(Archivo) Then
      If ComboOpciones.Text = "Procesados" Then
         Usuario = ObtenerUsuarioProceso(Archivo)
         Set Nitem = lv.ListItems.Add(, , Fecha)
         Nitem.SubItems(1) = Carpeta & Archivo
         Nitem.SubItems(2) = Archivo
         Nitem.SubItems(3) = FechaImportacion
         Nitem.SubItems(4) = Usuario
      End If
   Else
      'si no esta en el historico aun no fue procesado
      If ComboOpciones.Text = "No procesados" Then
         Set Nitem = lv.ListItems.Add(, , Fecha)
         Nitem.SubItems(1) = Carpeta & Archivo
         Nitem.SubItems(2) = Archivo
         Nitem.SubItems(3) = FechaImportacion
         Nitem.SubItems(4) = Usuario
      End If
   End If
   'sigue cargando el siguiente archivo de esa carpeta
   Archivo = Dir
Loop



Archivo = Dir("c:\rapipago\*.CSV")

Do While Trim(Archivo) <> ""
   Usuario = ""
   'obtengo la fecha del archivo
   Fecha = ObtenerFechaArchivoLINK(Archivo)
   FechaImportacion = ObtenerFechaImportacion(Archivo)
   'si existe en la tabla historica de archivos procesados
   If ExisteRPHistorico(Archivo) Then
      If ComboOpciones.Text = "Procesados" Then
         Usuario = ObtenerUsuarioProceso(Archivo)
         Set Nitem = lv.ListItems.Add(, , Fecha)
         Nitem.SubItems(1) = Carpeta & Archivo
         Nitem.SubItems(2) = Archivo
         Nitem.SubItems(3) = FechaImportacion
         Nitem.SubItems(4) = Usuario
      End If
   Else
      'si no esta en el historico aun no fue procesado
      If ComboOpciones.Text = "No procesados" Then
         Set Nitem = lv.ListItems.Add(, , Fecha)
         Nitem.SubItems(1) = Carpeta & Archivo
         Nitem.SubItems(2) = Archivo
         Nitem.SubItems(3) = FechaImportacion
         Nitem.SubItems(4) = Usuario
      End If
   End If
   'sigue cargando el siguiente archivo de esa carpeta
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
Private Sub CmdImportar_Click()
'inicia el proceso de importacion
Dim Mensaje As String
Dim bAgregar As Boolean
On Error GoTo merror
Call RefreshTimer

LabelMensaje.Caption = ""

'mientras importa deshabilita el boton
CmdImportar.Enabled = False
Me.MousePointer = vbHourglass

Mensaje = "No hubo importacion"

If datosok() Then
   If MsgP("¿Confirma la importacion de los archivos seleccionados?") Then
      'inicio de transaccion
      cnSQL.BeginTrans
      
      'vacio la tabla temporal
      Call LimpiarTabla("rapipago")
      Call RefreshTimer

      'si agregue registros a la tabla temporal
      bAgregar = AgregarRegistrosArchivos()
      If bAgregar Then
         'cobro usando la cuenta corriente
         Call CobrarCuotasCTACTE
         
         'actualizo el ultimo inicio usado en el proceso de cobro
         sql = "update configuracionsistema " & _
               "set iniciorp=" & CLng(TxtInicio.Text)
         cnSQL.Execute sql
         
         Mensaje = "Se finalizo la importacion correctamente"
      Else
         Mensaje = "No hubo importacion..verifique que los archivos no esten procesados desde antes"
      End If
      
      'fin de transaccion
      cnSQL.CommitTrans
      
      MsgI Mensaje
   End If
End If

'actualizo la lista de archivos RP
Call CargarListaArchivos

CmdImportar.Enabled = True
Me.MousePointer = vbDefault

Exit Sub
merror:
tratarerrores "Error en boton Importar"
End Sub
Private Function AgregarRegistrosArchivos() As Boolean
'agrega a la tabla temporal (rapipago) las facturas de los archivos rp seleccionados
Dim Filas As Long
Dim I As Long
Dim Detalle As String
Dim Archivo As String
Dim ArchivoSolo As String
Dim Mensaje As String
Dim CantReg As Long
Dim cont As Long
Dim Agregue As Boolean
Dim TipoArchivo As String
On Error GoTo merror

Agregue = False
cont = 0
CantReg = 0
Mensaje = ""

Filas = lv.ListItems.Count

'recorro la lista de archivos
For I = 1 To Filas
    'si el archivo esta seleccionado
    If lv.ListItems.Item(I).Checked Then
       Archivo = lv.ListItems.Item(I).SubItems(1)
       ArchivoSolo = lv.ListItems.Item(I).SubItems(2)
       
       TipoArchivo = DeterminarTipoArchivo(Archivo)
       
       Select Case TipoArchivo
       Case "RP-PF"
           'si el archivo es de credimaco
           If VerificarArchivoRP(Archivo) Then
              'si el archivo no esta procesado de antes
              If Not ExisteRPHistorico(ArchivoSolo) Then
                 'obtengo la cantidad de facturas de cada archivo rp
                 CantReg = ObtenerCantRegistros(Archivo)
                 If CantReg > 0 Then
                    'agrego los registros de ese archivo a la tabla rapipago
                    Call RefreshTimer
    
                    Call AgregarFacturas(Archivo, ArchivoSolo, CantReg)
                    
                    'paso el archivo al historico
                    Call RegistrarRPHistorico(ArchivoSolo, TipoArchivo)
                    Agregue = True
                 End If
              End If
           End If
        Case "LINK"
            If Not ExisteRPHistorico(ArchivoSolo) Then
                'agrego los registros de ese archivo a la tabla rapipago
                Call RefreshTimer
    
                If AgregarFacturasLINK(Archivo, ArchivoSolo) Then
                    Agregue = True
                End If
                'paso el archivo al historico
                Call RegistrarRPHistorico(ArchivoSolo, TipoArchivo)
            End If
        End Select
    End If
Next I

AgregarRegistrosArchivos = Agregue

Exit Function
merror:
tratarerrores "Error en procedimiento AgregarRegistrosArchivos"
End Function

Private Sub AgregarFacturas(ByVal Archi As String, ByVal ArchiSolo As String, ByVal CantReg As Long)
'solo pasa las facturas de un archivo a la tabla rapipago
Dim sql As String
Dim rec As rdoResultset
Dim nFic As Integer
Dim TamFic As Long
Dim Scontenido As String
Dim Inicio As Long
Dim cont As Long
Dim I As Long
Dim logic1 As Integer
Dim Registro As String
Dim CadFechaCobro As String
Dim CadImporteCobro As String
Dim CodigoBarras As String
Dim Año As String
Dim Mes As String
Dim Dia As String
Dim Fecha As String
Dim FechaCobro As Date
Dim Entero As String
Dim Decim As String
Dim cDNI As String
Dim CadPrestamo As String
Dim Decim1 As Currency
Dim ImporteCobro As Currency
Dim Comprobante As String
Dim numcomprobante As Long
Dim ImporteVencimiento1 As Currency
Dim ImporteVencimiento2 As Currency
Dim FechaJuliana As String
Dim FechaVencimiento1 As Date
Dim FechaVencimiento2 As Date
Dim ImporteRecargo As Currency
Dim CadDias As String
Dim DiasVencimiento2 As Long
Dim CadCliente As String
Dim IdCliente As Long
Dim CadEmpresa As String
Dim Header As String
Dim Origen As String
Dim NumEmpresa As Long
Dim IdCredito As Long
Dim NumCuota As Long
Dim CodPrestamo As String
Dim IdRapiPago As Long
Dim Ok As Boolean
On Error GoTo merror

'leo todo
If Len(Dir$(Archi)) Then
    nFic = FreeFile
    Open Archi For Input As nFic
    TamFic = LOF(nFic)
    Scontenido = Input$(TamFic, nFic)
    Close nFic
End If

cont = 0

Header = Mid(Scontenido, 1, 72)
If Mid$(Header, 49, 8) = "RAPIPAGO" Then
    Origen = "RP"
Else
    Origen = "PF"
End If

'debo iniciar la lectura en el primer registro 76
Inicio = CLng(TxtInicio.Text)

For I = 1 To CantReg
    'saco el registro completo
    Registro = (Mid(Scontenido, Inicio, 80))
    
    'saco los 3 campos principales
    CadFechaCobro = Mid(Registro, 1, 8)
    CadImporteCobro = Mid(Registro, 9, 15)
    CodigoBarras = Mid(Registro, 24, 44)
        
    'armo la fecha de cobro
    Año = Mid(CadFechaCobro, 1, 4)
    Mes = Mid(CadFechaCobro, 5, 2)
    Dia = Mid(CadFechaCobro, 7, 2)
    Fecha = Dia + "/" + Mes + "/" + Año
    If IsDate(Fecha) Then
       FechaCobro = CDate(Fecha)
       Ok = True
    Else
       'la fecha es incorrecta por alguna causa
       Ok = False
       LabelMensaje.Caption = "El formato del archivo de RapiPago:" & ArchivoSolo & " es incorrecto...(por favor reviselo)"
    End If
    
    If Ok Then
       'armo el importe cobrado
       'saco la parte entera del importe
       Entero = Mid(CadImporteCobro, 1, 13)
       'obtengo los decimales
       Decim = Mid(CadImporteCobro, 14, 2)
       Decim1 = CCur(Decim) / 100
       ImporteCobro = CCur(Entero) + CCur(Decim1)
    
       CadEmpresa = Mid(CodigoBarras, 1, 3)
       NumEmpresa = CLng(CadEmpresa)
       CadCliente = Mid(CodigoBarras, 4, 8)
       IdCliente = CLng(CadCliente)
    
       'ahora saco el resto de los componentes desde el codigo de barras
       'obtengo el numero de comprobante desde 14 para sacar solo un long
       Comprobante = Mid(CodigoBarras, 14, 9)
       numcomprobante = CLng(Comprobante)
    
       'obtengo el importe al primer vto
       Entero = Mid(CodigoBarras, 23, 6)
       'obtengo los decimales
       Decim = Mid(CodigoBarras, 29, 2)
       Decim1 = CCur(Decim) / 100
       ImporteVencimiento1 = CCur(Entero) + CCur(Decim1)
    
       'obtengo la fecha de vencimiento original de la factura
       'esta en formato juliano dentro del codigo de barras
       FechaJuliana = Mid(CodigoBarras, 31, 5)
    
       'reconstruyo la fecha de vto1 normal a partir de la juliana
       FechaVencimiento1 = CrearFechaNormal(FechaJuliana)
    
       'obtengo el recargo
       'la parte entera
       Entero = Mid(CodigoBarras, 36, 4)
       'la parte decimal
       Decim = Mid(CodigoBarras, 40, 2)
       'recontruyo el importe de recargo
       Decim1 = CCur(Decim) / 100
       ImporteRecargo = CCur(Entero) + CCur(Decim1)
    
       ImporteVencimiento2 = CCur(ImporteVencimiento1) + CCur(ImporteRecargo)
    
       'obtengo los dias al segundo vencimiento
       CadDias = Mid(CodigoBarras, 42, 2)
       DiasVencimiento2 = CLng(CadDias)
    
       FechaVencimiento2 = FechaVencimiento1 + DiasVencimiento2
    
       IdCredito = ObtenerIdCredito(IdCliente, numcomprobante)
       NumCuota = ObtenerNumCuota(IdCliente, numcomprobante)
       CodPrestamo = ObtenerCodPrestamo(IdCliente, IdCredito)
       logic1 = 0
       
       'proceso cliente con dni
        If Mid(CadCliente, 1, 2) <> "00" Then
            cDNI = CadCliente
            CadPrestamo = ObtenerCodPrestamoConDocumento(cDNI)
            If CadPrestamo = "" Then
                CadPrestamo = ObtenerCodPrestamoConDocumentoInclFinalizados(cDNI)
            End If
            CadPrestamo = UCase(Trim(CadPrestamo))
            CadCliente = ""
            IdCredito = 0
            IdCliente = 0
            If CadPrestamo <> "" Then
                IdCredito = ObtenerCredito(CadPrestamo)
                IdCliente = ObtenerNumCliente(IdCredito)
                CadCliente = UCase$(Trim$(ObtenerNombreCliente(IdCliente)))
            End If
            numcomprobante = Val(cDNI)
            logic1 = 1
        End If
       
       'agrego los datos de la factura a la tabla rapipago
       IdRapiPago = UltimoId("idrapipago", "rapipago") + 1
    
       sql = "insert into rapipago(idrapipago,numempresa,idcliente," & _
             "numcomprobante,codprestamo,idcredito,numcuota,fechacobro," & _
             "importecobro,codigobarras,importevencimiento1,fechavencimiento1," & _
             "importevencimiento2,fechavencimiento2,importerecargo,diasvencimiento2," & _
             "archivorp,fechaimportacion,origen,logic1) " & _
             "values(" & CLng(IdRapiPago) & "," & CLng(NumEmpresa) & _
             "," & CLng(IdCliente) & "," & CLng(numcomprobante) & _
             ",'" & CStr(CodPrestamo) & "'," & CLng(IdCredito) & _
             "," & CLng(NumCuota) & ",'" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & _
             "'," & ConvertirDblSql(CCur(ImporteCobro)) & ",'" & CStr(CodigoBarras) & _
             "'," & ConvertirDblSql(CCur(ImporteVencimiento1)) & ",'" & ConvertirFechaSql(CDate(FechaVencimiento1), "DD/MM/YYYY") & _
             "'," & ConvertirDblSql(CCur(ImporteVencimiento2)) & ",'" & ConvertirFechaSql(CDate(FechaVencimiento2), "DD/MM/YYYY") & _
             "'," & ConvertirDblSql(CCur(ImporteRecargo)) & "," & CLng(DiasVencimiento2) & _
             ",'" & CStr(ArchiSolo) & "','" & ConvertirFechaSql(Mid$(Now, 1, 10), "DD/MM/YYYY") & "','" & Origen & "'," & logic1 & ")"
      cnSQL.Execute (sql)
      Call RefreshTimer

    End If 'is ok
    
    'avanzo en el archivo de texto
    If Trim$(Mid$(Registro, 68, 6)) = "" Then
        Inicio = Inicio + 75
    Else
        Inicio = Inicio + 79
    End If
Next I

Exit Sub
merror:
tratarerrores "Error en funcion AgregarRegistros"
End Sub

Private Function AgregarFacturasLINK(ByVal Archivo As String, ByVal ArchivoSolo As String) As Boolean
Dim bTitulo As Boolean
Dim bAgregue As Boolean
Dim Scontenido As String
Dim nFic As Integer
Dim cDNI As String
Dim cImporte As String
Dim cFecha As String
Dim Campos() As String
Dim Cliente As String
Dim CadPrestamo As String
Dim IdCredito As Long
Dim IdCliente As Long
Dim IdRapiPago As Long
Dim numcomprobante As Long
Dim ImporteCobrado As Currency
Dim CadCliente As String
Dim FechaArchivo As Date
Dim Mreporte As New ARContenidoRP

On Error GoTo merror
Call RefreshTimer

bAgregue = False
'inicio transaccion
cnSQL.BeginTrans
FechaArchivo = ObtenerFechaArchivoLINK(ArchivoSolo)
nFic = FreeFile
Open Archivo For Input As #nFic
bTitulo = True
While Not EOF(nFic)
    Line Input #nFic, Scontenido
    If bTitulo Then
        bTitulo = False
    Else
        Campos = Split(Scontenido, ",")
        cDNI = DepuraString(Campos(1))
        cImporte = DepuraString(Campos(2))
        ImporteCobrado = CCur(cImporte) / 100
        cFecha = DepuraString(Campos(3))
        CadPrestamo = ObtenerCodPrestamoConDocumento(cDNI)
        If CadPrestamo = "" Then
            CadPrestamo = ObtenerCodPrestamoConDocumentoInclFinalizados(cDNI)
        End If
        CadPrestamo = UCase(Trim(CadPrestamo))
        CadCliente = ""
        IdCredito = 0
        IdCliente = 0
        If CadPrestamo <> "" Then
            IdCredito = ObtenerCredito(CadPrestamo)
            IdCliente = ObtenerNumCliente(IdCredito)
            CadCliente = UCase$(Trim$(ObtenerNombreCliente(IdCliente)))
        End If
        Origen = "RL"
        numcomprobante = Val(cDNI)
        IdRapiPago = UltimoId("idrapipago", "rapipago") + 1
        sql = "insert into rapipago(idrapipago,numempresa,idcliente," & _
              "numcomprobante,codprestamo,idcredito,numcuota,fechacobro," & _
              "importecobro,codigobarras,importevencimiento1,fechavencimiento1," & _
              "importevencimiento2,fechavencimiento2,importerecargo,diasvencimiento2," & _
              "archivorp,fechaimportacion,origen,logic1) " & _
              "values(" & CLng(IdRapiPago) & ",0" & _
              "," & CLng(IdCliente) & "," & numcomprobante & _
              ",'" & CStr(CadPrestamo) & "'," & CLng(IdCredito) & _
              ",0,'" & Mid$(cFecha, 1, 4) & "/" & Mid$(cFecha, 5, 2) & "/" & Mid$(cFecha, 7, 2) & _
              "'," & ConvertirDblSql(ImporteCobrado) & ",''" & _
              ",0,'1900/01/01',0,'1900/01/01',0,0" & _
              ",'" & CStr(ArchivoSolo) & "','" & ConvertirFechaSql(Mid$(Now, 1, 10), "DD/MM/YYYY") & "','" & Origen & "',1)"
        cnSQL.Execute (sql)
        bAgregue = True
        Call RefreshTimer
    End If
Wend
Close #nFic
'fin de transaccion
cnSQL.CommitTrans
Call RefreshTimer
AgregarFacturasLINK = bAgregue
Exit Function
merror:
tratarerrores "Error en funcion AgregarFacturasLINK"
End Function

Private Function CobrarCuotasCTACTE() As Long
'toma los registros de la tabla rapipago y los procesa
Dim sql As String
Dim rec As rdoResultset
Dim rec2 As rdoResultset
'estos son para cobros por items
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
Dim ImporteTotalRestanteOriginal As Currency

'para cobros por items parciales
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

'para restar con los demas en cobros totales y ultimo parcial
Dim SaldoCuota As Currency
Dim SaldoCuota1erVenc As Currency
Dim Importe1erVenc As Currency

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
Dim Cubre As Boolean
Dim IdExcedenteCliente As Long
Dim Importe2 As Currency
Dim IdCliente As Long
Dim ArchiSolo As String
Dim Observaciones As String
Dim Origen As String
Dim DetalleExcedente As String
Dim IdRapiPago As Long
Dim IdRapiPagoNuevo As Long
Dim nImporteActualizar As Currency
Dim nIdPagoAnterior As Long
Dim ImporteTotalRestante As Currency
Dim ImporteGralRestante As Currency
Dim ImporteIngresos As Currency
Dim I As Long
Dim Diferencia As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim CuotaOk As Boolean
Dim DetalleNew As String
Dim logic1 As Boolean
Dim bPrimero As Boolean
Dim ClienteAnterior As Long
Dim NumComprobanteAnterior As Long
Dim ArchivoAnterior As String
Dim OrigenAnterior As String

Dim numcomprobante As Long
Dim ImporteCobro As Currency
Dim FechaCobro As Date
Dim FechaImportacion As Date
Dim ImporteVencimiento1 As Currency
Dim ImporteVencimiento2 As Currency
Dim IvaACobrarDevuelto  As Currency

Dim SecuenciaIngreso As Long
Dim MontoCobradoSecuencia As Currency

On Error GoTo merror

sql = "select * from rapipago " & _
      "order by idcliente,idcredito,numcuota"
      
Set rec = cnSQL.OpenResultset(sql)
Call RefreshTimer

ImporteTotalRestante = 0
ImporteGralRestante = 0
I = 1

If Not rec.EOF Then
    Do While Not rec.EOF
       'estos los saca de rapipago con rec
       'tomo los datos de cada factura
       IdRapiPago = CLng(rec.rdoColumns("idrapipago"))
       IdCliente = CLng(rec.rdoColumns("idcliente"))
       IdCredito = CLng(rec.rdoColumns("idcredito"))
       NumCuota = CLng(rec.rdoColumns("numcuota"))
       
       numcomprobante = CLng(rec.rdoColumns("numcomprobante"))
       ImporteCobro = CCur(rec.rdoColumns("importecobro"))
       FechaCobro = CDate(rec.rdoColumns("fechacobro"))
       
       ArchiSolo = rec.rdoColumns("archivorp")
       CodPrestamo = rec.rdoColumns("codprestamo")
       
       Origen = rec.rdoColumns("origen")
       FechaImportacion = rec.rdoColumns("FechaImportacion")
       
       logic1 = rec.rdoColumns("logic1")
       bPrimero = True
       'si la cuota esta cobrada...cubre otras pendientes del mismo credito
       'si la cuota es comodin...cubre otras pendientes del mismo credito
       'si la cuota esta refinanciada...cubre otras pendientes del mismo credito
       'si no hay pendientes deja todo el importe como excedente
       'que pasa si el credito esta finalizado??..importa igual
       
       'A PARTIR DE AHORA SI EL CREDITO DE LA CUOTA ESTA BLOQUEADO o finalizado LO PASA A EXCEDENTE
                    
       'ahora le asigno el importecobrado en esta cuota del rp
       'este es el que se ira restando al cubrir cuotas
       ImporteTotalRestante = CCur(ImporteCobro)
       DetalleExcedente = ""
       
       SecuenciaIngreso = UltimoId("secuencia", "ingresos") + 1
       MontoCobradoSecuencia = ImporteTotalRestante

       
       'tomo todas las cuotas pendientes de ese credito ordenadas en ascendente
       'IMPORTANTE: Aca cambio el ordenamiento de las cuotas que va a ir cubriendo
       'con cta cte....donde dejara para el ultimo las que vencen al final
       'y si hay alguna que se le cambio el vto quedara sola para el final
       'antes ordenaba solo por cuotas.numcuota...ahora le cambie a cuotas.fechavencimiento1
          
       'NUEVO 2010 le puse la condicion de no bloqueada..si esta bloqueada
       'no carga cuotas pendientes de ese credito y salta directametne
       'a excedentes
       'no debe estar finalizado ni bloqueado sino salta abajo
'       sql = "select cuotas.*,cuotas.logic1 as exceptuada " & _
'             "from creditos inner join cuotas " & _
'             "on creditos.idcredito=cuotas.idcredito " & _
'             "where creditos.idcliente=" & CLng(IdCliente) & " " & _
'             "and cuotas.idcredito=" & CLng(IdCredito) & " " & _
'             "and cuotas.fechacobro is Null " & _
'             "and cuotas.fecharefinanciacion is Null " & _
'             "and cuotas.cuotacomodin = 0 " & _
'             "and creditos.fechabloqueo is Null " & _
'             "and creditos.fechafinalizacion is Null " & _
'             "order by cuotas.fechavencimiento1"

       sql = "select creditos.codprestamo,cuotas.*,cuotas.logic1 as exceptuada " & _
             "from creditos, cuotas " & _
             "where creditos.idcliente = " & CLng(IdCliente) & " " & _
             "and creditos.idcredito = cuotas.idcredito " & _
             "and cuotas.fechacobro is Null " & _
             "and cuotas.fecharefinanciacion is Null " & _
             "and cuotas.cuotacomodin = 0 " & _
             "and creditos.fechabloqueo is Null " & _
             "and creditos.fechafinalizacion is Null " & _
             "order by cuotas.fechavencimiento1,creditos.fechacredito"

       Set rec2 = cnSQL.OpenResultset(sql)
       Call RefreshTimer

       If Not rec2.EOF Then
             'recorro la lista de cuotas saldando lo que encuentre sin pagar
             Do While Not rec2.EOF
                'si queda resto aun para seguir cubriendo cuotas
                
                If CCur(ImporteTotalRestante) > 0 Then
                    ImporteTotalRestanteOriginal = CCur(ImporteTotalRestante)
                   'esta puede llegar a ser otra cuota o la misma inicial
                   NumCuota = CLng(rec2.rdoColumns("numcuota"))
                   IdCredito = CLng(rec2.rdoColumns("IdCredito"))
                   CodPrestamo = rec2.rdoColumns("codprestamo")
                   ImporteVencimiento1 = CCur(rec2.rdoColumns("importevencimiento1"))
                   ImporteVencimiento2 = CCur(rec2.rdoColumns("importevencimiento2"))
                   Vencimiento1 = CDate(rec2.rdoColumns("fechavencimiento1"))
                   Vencimiento2 = CDate(rec2.rdoColumns("fechavencimiento2"))
                   
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
                   
                   Cubre = False
                   IngresosOk = False
                   SaldoCuota = 0
                   ImporteMora = 0
                   IvaMora = 0
                   'veo si tiene cobros parciales de antes
                   ImporteParcial = ObtenerImporteParcialX(IdCredito, NumCuota)
                   SaldoCuota = ObtenerSaldoCuotaX(IdCredito, NumCuota, FechaCobro, SaldoCuota1erVenc)
                   Importe1erVenc = ObtenerImporte1erVenc(IdCredito, NumCuota)
                   'si pago al dia antes del 1º vto
                   If CDate(FechaCobro) <= CDate(Vencimiento1) Then
                      Vencimiento2Cuota = 0
                      'si cubro el importe del vencimiento1
                      If (CCur(ImporteTotalRestante) + CCur(ImporteParcial)) >= CCur(ImporteVencimiento1) Then
                         Cubre = True
                      End If
                   End If 'fin al 1 vto
                   'si pago entre el 1º y 2º vencimiento uso el recargo
                   If (CDate(FechaCobro) > CDate(Vencimiento1)) And (CDate(FechaCobro) <= CDate(Vencimiento2)) Then
                      'si cubre el importe del segundo vencimiento
                      If (CCur(ImporteTotalRestante) + CCur(ImporteParcial)) >= CCur(ImporteVencimiento2) Then
                         Cubre = True
                      End If
                   End If 'fin entre 1 2 vto
                   'si pago despues del 2º vto
                   If CDate(FechaCobro) > CDate(Vencimiento2) Then
                      'pago despues del 2º vto entonces debo calcular mora
                      'calculo la mora entre la fecha de 2º vto y la fecha de cobro de rapipago
                      'calculo la mora de forma habitual
                      'puedo pasarle el campo [exceptuada]
                      ImporteMora = CalculoMoraPendiente(rec2.rdoColumns("idcredito"), rec2.rdoColumns("numcuota"), rec2.rdoColumns("exceptuada"), Importe1erVenc, Vencimiento1, FechaCobro, IvaACobrarDevuelto)
                      '''''''********ImporteMora = CalcularInteresMoraZZ(rec2.rdoColumns("exceptuada"), Importe1erVenc, Vencimiento1, FechaCobro)
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
                      SaldoCuota = CCur(SaldoCuota1erVenc) + CCur(ImporteMora) + CCur(IvaMora)
                                        
                      If CCur(ImporteTotalRestante) >= CCur(SaldoCuota) Then
                         Cubre = True
                      End If
                   End If 'fin 2º vto con mora
             
                   'si cubro la cuota, el importe real cobrado es lo que le quedaba de saldo
                   If Cubre Then
                      ImporteRealCobrado = CCur(SaldoCuota)
                
                      'si tenia cobros parciales
                      If CCur(ImporteParcial) > 0 Then
                         'traigo el importe total cobrado incluyendo este
                         'aca le cambie y saque el saldo que lo incrementaba
                         ImporteRealCobrado = CCur(ImporteRealCobrado) + CCur(ImporteParcial)
                      End If 'fin si tenia cobros parciales
                      'registro el cobro total
                      sql = "update cuotas set rapipago='true'," & _
                            "fechacobro='" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & "'," & _
                            "importecobrado='" & ConvertirDblSql(CCur(ImporteRealCobrado)) & "',origen='" & Origen & _
                            "' where idcredito='" & CLng(IdCredito) & "' and numcuota='" & CLng(NumCuota) & "'"
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
                      ImporteIngresos = CCur(SaldoCuota)
                      IngresosOk = True
                 
                   Else ' else del si cubre
                      'no cubre ES UN COBRO PARCIAL DE CUOTA
                      'lo cobrado parcial es todo el importecobrado de la factura del archivo rp
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
                         'si sigue intento cubrir
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
                         'si sigue intento cubrir el ivaotorgastos
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
                         'si sigue intento cubrir el iva seguros
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
                         'si sigue intento cubrir el interes
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
                      'si sigue intento cubrir el otorgamiento
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
                     'si sigue intento cubrir los gastos SI ES QUE AUN NO ESTAN CUBIERTOS!!!
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
                    'si sigue intento cubrir el seguro
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
                    'si sigue intento cubrir el interes
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
                     'si sigue intento cubrir el iva interes
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
                 sql = "update cuotas set rapipago=1, origen='" & Origen & "'," & _
                       "cobrosparciales=1 " & _
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
                     "fechacobro,importecobrado,codprestamo,rapipago,numcomprobante," & _
                     "capitalcobrado,interescobrado,vencimiento2cobrado,refincobrado," & _
                     "gastoscobrados,seguroscobrados,otorgamientocobrado,ivainterescobrado," & _
                     "ivaseguroscobrado,ivaotorgastoscobrado,moracobrada,ivamoracobrada,usuario,origen,fechaimputacion,secuencia,montocobradosecuencia) " & _
                     "values(" & CLng(IdIngreso) & "," & CLng(IdCredito) & "," & CLng(NumCuota) & _
                     ",'" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & "'," & ConvertirDblSql(CCur(ImporteIngresos)) & _
                     ",'" & CStr(CodPrestamo) & "',1," & CLng(numcomprobante) & _
                     "," & ConvertirDblSql(CCur(CapitalCobrado)) & "," & ConvertirDblSql(CCur(InteresCobrado)) & _
                     "," & ConvertirDblSql(CCur(Vencimiento2Cobrado)) & "," & ConvertirDblSql(CCur(RefinCobrado)) & _
                     "," & ConvertirDblSql(CCur(GastosCobrados)) & "," & ConvertirDblSql(CCur(SegurosCobrados)) & _
                     "," & ConvertirDblSql(CCur(OtorgamientoCobrado)) & "," & ConvertirDblSql(CCur(IvaInteresCobrado)) & _
                     "," & ConvertirDblSql(CCur(IvaSegurosCobrado)) & "," & ConvertirDblSql(CCur(IvaOtorGastosCobrado)) & _
                     "," & ConvertirDblSql(CCur(MoraCobrada)) & "," & ConvertirDblSql(CCur(IvaMoraCobrada)) & _
                     ",'" & CStr(VG_USUARIOLOGIN) & "','" & Origen & "','" & ConvertirFechaSql(FechaImportacion, "DD/MM/YYYY") & "'," & SecuenciaIngreso & "," & ConvertirDblSql(MontoCobradoSecuencia) & ")"
               cnSQL.Execute sql
               
               If logic1 Then
                     If bPrimero Then
                        bPrimero = False
                        nImporteActualizar = CCur(ImporteIngresos)
                        nIdPagoAnterior = IdRapiPago
                        ClienteAnterior = rec.rdoColumns("idCliente")
                        NumComprobanteAnterior = rec.rdoColumns("numcomprobante")
                        ArchivoAnterior = rec.rdoColumns("archivorp")
                        OrigenAnterior = rec.rdoColumns("origen")
                     Else
                        
                        sql = "update rapipago set importecobro = " & ConvertirDblSql(nImporteActualizar) & _
                              "where idrapipago = " & nIdPagoAnterior
                        cnSQL.Execute sql

                        IdRapiPagoNuevo = UltimoId("idrapipago", "rapipago") + 1
                        sql = "insert into rapipago(idrapipago,idcliente,idcredito,codprestamo," & _
                              "fechacobro,importecobro,archivorp,fechaimportacion," & _
                              "origen,numcomprobante) " & _
                              "values(" & CLng(IdRapiPagoNuevo) & ",'" & ClienteAnterior & "'," & _
                              CLng(IdCredito) & ",'" & CStr(CodPrestamo) & "'," & _
                              "'" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & "'," & ConvertirDblSql(CCur(ImporteTotalRestanteOriginal)) & "," & _
                              "'" & CStr(ArchivoAnterior) & "',GetDate()," & _
                              "'" & OrigenAnterior & "','" & NumComprobanteAnterior & "')"
                        cnSQL.Execute sql
                        
                        nImporteActualizar = CCur(ImporteIngresos)
                        nIdPagoAnterior = IdRapiPagoNuevo
                     End If
               End If
               
               If VG_FINALIZARAUTOMATICAMENTE Then
                  'si es la ultima cuota finalizo el credito
                  Call FinalizarCredito(IdCredito, Date)
               End If
            End If 'fin ingresosok
                      
          End If 'si importetotalrestante
          
          rec2.MoveNext
          Call RefreshTimer

       Loop  'fin del segundo ciclo interno que recorre cuotas
       
       'aca termino de recorrer la lista de cuotas de una cuota/credito de rapipago
       'cuidado que el excedente debe ser por cliente!!!!
       'me fijo si quedo algo de la vuelta anterior
       If CCur(ImporteTotalRestante) > 0 Then
          Observaciones = "Excedente de cta cte"
          IdExcedenteCliente = UltimoId("idexcedentecliente", "excedentesclientes") + 1
          'es un excedente de la cuota original de rapipago de arriba
          NumCuota = CLng(rec.rdoColumns("numcuota"))
                
          sql = "insert into excedentesclientes (idexcedentecliente,idcliente,idcredito," & _
                "codprestamo,numcuota,fechacobro,importecobro,rapipago,archivorp,observaciones,origen,fechaimputacion) " & _
                "values('" & CLng(IdExcedenteCliente) & "','" & CLng(IdCliente) & _
                "','" & CLng(IdCredito) & "','" & CStr(CodPrestamo) & "','" & CLng(NumCuota) & _
                "','" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & "','" & ConvertirDblSql(CCur(ImporteTotalRestante)) & _
                "','True','" & CStr(ArchiSolo) & "','" & CStr(Observaciones) & "','" & Origen & "','" & ConvertirFechaSql(FechaImportacion, "DD/MM/YYYY") & "')"
          cnSQL.Execute sql
          DetalleExcedente = "Excedente de cta cte"
          'agrego el detalle de excedente a rapipago
          sql = "update rapipago set excedentes='" & CStr(DetalleExcedente) & "' " & _
                "where idrapipago='" & CLng(IdRapiPago) & "'"
          cnSQL.Execute sql
       End If
    
    Else 'del primer IF NOT REC2.EOF que contiene al primer while
       'esto es el creditode la factura original de rp no tiene cuotas pendientes
       'O ESTA EL CREDITO BLOQUEADO
       If CCur(ImporteTotalRestante) > 0 Then
          Observaciones = "Excedente de cta cte"
          DetalleExcedente = "Excedentes"
          
          If CreditoBloqueado1(IdCredito) Then
             Observaciones = "Excedente Cta Cte, credito bloqueado"
             DetalleExcedente = "Excedentes, credito bloqueado"
          End If
          
          If CreditoFinalizado(IdCredito) Then
             Observaciones = "Excedente Cta Cte, credito finalizado"
             DetalleExcedente = "Excedentes, credito finalizado"
          End If
          
          IdExcedenteCliente = UltimoId("idexcedentecliente", "excedentesclientes") + 1
          sql = "insert into excedentesclientes (idexcedentecliente,idcliente,idcredito," & _
                "codprestamo,numcuota,fechacobro,importecobro,rapipago,archivorp,observaciones,origen,fechaimputacion) " & _
                "values(" & CLng(IdExcedenteCliente) & "," & CLng(IdCliente) & _
                "," & CLng(IdCredito) & ",'" & CStr(CodPrestamo) & "'," & CLng(NumCuota) & _
                ",'" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & "'," & ConvertirDblSql(CCur(ImporteTotalRestante)) & _
                ",1,'" & CStr(ArchiSolo) & "','" & CStr(Observaciones) & "','" & Origen & "','" & ConvertirFechaSql(FechaImportacion, "DD/MM/YYYY") & "')"
          cnSQL.Execute sql
          
          'agrego el detalle de excedente a rapipago
          sql = "update rapipago set excedentes='" & CStr(DetalleExcedente) & "' " & _
                "where idrapipago='" & CLng(IdRapiPago) & "'"
          cnSQL.Execute sql
       
       End If
    End If
    
  I = I + 1
  rec.MoveNext
  Call RefreshTimer

 
 Loop 'del 1º while

End If ' si no es eof del primer if

Exit Function
merror:
tratarerrores "Error en funcion CobrarCuotasCtaCte"
End Function

Private Function VerificarArchivoRP(ByVal Archi As String) As Boolean
'verifica si el archivo pertenece a credimaco,num 753
Dim Cadena As Long
On Error GoTo merror

VerificarArchivoRP = False

'tomo los ultimos 4 digitos de la extension

Cadena = Trim(Right(Archi, 4))

'debe ser igual al numero 1032
If Val(Cadena) = VG_CODIGOAUTOMATICO Then
   VerificarArchivoRP = True
End If

Exit Function
merror:
tratarerrores "Error en funcion VerificarArchivoRP"
End Function
Private Function ExisteRPHistorico(ByVal Archivo As String) As Boolean
'verifica si un archivo fue procesado con anterioridad
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

ExisteRPHistorico = False

Archivo = UCase(Trim(Archivo))

sql = "select nombrearchivo " & _
      "from rapipagohistorico " & _
      "where nombrearchivo='" & CStr(Archivo) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("nombrearchivo")) Then
      ExisteRPHistorico = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteRPHistorico"
End Function
Private Function ObtenerUsuarioProceso(ByVal Archivo As String) As String
'obtiene el usuario que proceso ese archivo
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerUsuarioProceso = ""

sql = "select usuario " & _
      "from rapipagohistorico " & _
      "where nombrearchivo='" & CStr(Archivo) & "'"
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("usuario")) Then
      ObtenerUsuarioProceso = CStr(rec.rdoColumns("usuario"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerUsuarioProceso"
End Function
Private Function ObtenerFechaImportacion(ByVal Archivo As String) As String
'obtiene la fecha de importacion de cada archivo del historico
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerFechaImportacion = ""

sql = "select fechaproceso " & _
      "from rapipagohistorico " & _
      "where nombrearchivo='" & CStr(Archivo) & "'"
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   ObtenerFechaImportacion = CStr(rec.rdoColumns("fechaproceso"))
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerFechaImportacion"
End Function
Private Function ObtenerCantRegistros(ByVal Archi As String) As Long
'devuelve la cantidad de registros (facturas) del archivo RP
'tomando el contador del final del mismo archivo
Dim nFic As Integer
Dim TamFic As Long
Dim Scontenido As String
Dim Pos As Long
Dim CantRegistros As Long
On Error GoTo merror

If Len(Dir$(Archi)) Then
    nFic = FreeFile
    Open Archi For Input As nFic
    TamFic = LOF(nFic)
    Scontenido = Input$(TamFic, nFic)
    Close nFic
End If

'busco donde comienza el trailer
Pos = InStr(Scontenido, "99999999")
CantRegistros = CLng(Mid(Scontenido, Pos + 8, 8))
ObtenerCantRegistros = CLng(CantRegistros)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCantRegistros"
End Function
Private Sub RegistrarRPHistorico(ByVal Archi As String, ByVal Tipo As String)
'registra en el historico el archivo recien procesado
Dim sql As String
Dim rec As rdoResultset
Dim Fecha As Date
On Error GoTo merror

Select Case Tipo
Case "RP-PF"
    Fecha = ObtenerFechaArchivoRP(Archi)
Case "LINK"
    Fecha = ObtenerFechaArchivoLINK(Archi)
End Select
'verifico si ya esta registrado
If ExisteRPHistorico(Archi) Then Exit Sub

sql = "insert into rapipagohistorico (nombrearchivo,fechaarchivo,fechaproceso,usuario) " & _
      "values ('" & CStr(Archi) & "','" & ConvertirFechaSql((CDate(Fecha)), "DD/MM/YYYY") & "','" & ConvertirFechaSql(Mid$(Now, 1, 10), "DD/MM/YYYY") & "','" & CStr(VG_USUARIOLOGIN) & "')"

cnSQL.Execute sql

Exit Sub
merror:
tratarerrores "Error en procedimiento RegistrarRPHistorico"
End Sub
Private Function ObtenerFechaArchivoRP(ByVal Archi As String) As Date
'obtiene la fecha desde el nombre del archivo pf
Dim Dia As Long
Dim Mes As Long
Dim Año As Long
Dim Fecha As Date
On Error GoTo merror

'ahora cambio a dia,mes,año
Dia = CLng(Mid(Archi, 3, 2))
Mes = CLng(Mid(Archi, 5, 2))
Año = CLng(Mid(Archi, 7, 2))
Fecha = CDate(CStr(Dia) & "/" & CStr(Mes) & "/" & CStr(Año))

ObtenerFechaArchivoRP = Fecha

Exit Function
merror:
tratarerrores "Error en funcion ObtenerFechaArchivoRP"
End Function

Private Function ObtenerFechaArchivoLINK(ByVal Archi As String) As Date
'obtiene la fecha desde el nombre del archivo pf
Dim Dia As Long
Dim Mes As Long
Dim Año As Long
Dim Fecha As Date
On Error GoTo merror

'ahora cambio a dia,mes,año
Dia = CLng(Mid(Archi, 7, 2))
Mes = CLng(Mid(Archi, 5, 2))
Año = CLng(Format(Date, "yy"))
Fecha = CDate(CStr(Dia) & "/" & CStr(Mes) & "/" & CStr(Año))

ObtenerFechaArchivoLINK = Fecha

Exit Function
merror:
tratarerrores "Error en funcion ObtenerFechaArchivoLINK"
End Function

Private Sub cmdimprimir_Click()
    Call MostrarContenido
End Sub
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

If Trim(TxtInicio.Text) = "" Then
   datosok = False
   MsgE "Debe indicar la posicion de inicio de lectura de registros (recomendadas 58 y 75)"
   TxtInicio.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtInicio.Text) Then
   datosok = False
   MsgE "La posicion de inicio de lectura de registros debe ser numerica"
   TxtInicio.SetFocus
   Exit Function
End If
If Val(TxtInicio.Text) <= 0 Then
   datosok = False
   MsgE "La posicion de inicio debe ser mayor a cero"
   TxtInicio.SetFocus
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosOk"
End Function

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
Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim TipoArchivo As String
'muestra en la lista de abajo el contenido del archivo seleccionado arriba
If lv.ListItems.Count = 0 Then Exit Sub
If Not VerificarSeleccionLista(lv) Then Exit Sub
If Trim(TxtInicio.Text) = "" Then Exit Sub
If Not IsNumeric(TxtInicio.Text) Then Exit Sub
If Val(TxtInicio.Text) <= 0 Then Exit Sub

TipoArchivo = DeterminarTipoArchivo(lv.SelectedItem.SubItems(1))

Select Case TipoArchivo
Case "RP-PF"
    Call MostrarFacturasRapiPago(lv.SelectedItem.SubItems(1))
Case "LINK"
    Call MostrarFacturasLINK(lv.SelectedItem.SubItems(1))
End Select
End Sub

Private Sub MostrarFacturasRapiPago(ByVal Archi As String)
'muestra en la segunda lista las facturas de un archivo de rapipago seleccionado arriba
'solo muestra por pantalla (no registra datos en tablas temporales)
Dim I As Long
Dim nFic As Integer
Dim TamFic As Long
Dim Scontenido As String
Dim CantReg As Long
Dim Registro As String
Dim FechaCobro As String
Dim ImporteCobrado As String
Dim CodigoBarras As String
Dim Dia As String
Dim Mes As String
Dim Año As String
Dim Fecha As Date
Dim Entero As String
Dim Decim As String
Dim Importe As Currency
Dim Decim1 As Currency
Dim FechaJuliana As String
Dim CadDigito As String
Dim CadDias As String
Dim cDNI As String
Dim CadPrestamo As String
Dim cont As Long
Dim ImporteVencimiento1 As Currency
Dim ImporteVencimiento2 As Currency
Dim ImporteRecargo As Currency
Dim Dias As Long
Dim FechaVencimiento2 As Date
Dim CadeFecha As String
Dim Ok As Boolean

On Error GoTo merror

Lv2.ListItems.Clear

CantReg = ObtenerCantRegistros(Archi)
TxtContador2.Text = CantReg

If Len(Dir$(Archi)) Then
    nFic = FreeFile
    Open Archi For Input As nFic
    TamFic = LOF(nFic)
    Scontenido = Input$(TamFic, nFic)
    Close nFic
End If

cont = 0

Inicio = CLng(TxtInicio.Text)

For I = 1 To CantReg
    Registro = (Mid(Scontenido, Inicio, 80))
    
    FechaCobro = Mid(Registro, 1, 8)
    
    ImporteCobro = Mid(Registro, 9, 15)
    'aca solo leo 44 digitos del codigo de barras aunque traiga 50
    CodigoBarras = Mid(Registro, 24, 44)
    Año = Mid(FechaCobro, 1, 4)
    Mes = Mid(FechaCobro, 5, 2)
    Dia = Mid(FechaCobro, 7, 2)
    
    CadeFecha = Dia + "/" + Mes + "/" + Año
    If IsDate(CadeFecha) Then
       Fecha = CDate(CadeFecha)
       Ok = True
    Else
       Ok = False
    End If
    
    If Ok Then
       Entero = Mid(ImporteCobro, 1, 13)
       Decim = Mid(ImporteCobro, 14, 2)
       Decim1 = CCur(Decim) / 100
       Importe = CCur(Entero) + CCur(Decim1)
    
       Set Nitem = Lv2.ListItems.Add(, , Fecha)
       Nitem.SubItems(1) = Format(Importe, "0.00") & vbNullString
       Nitem.SubItems(2) = CodigoBarras
    
       'empresa
       Nitem.SubItems(3) = Mid(CodigoBarras, 1, 3)
    
       'cliente
       Nitem.SubItems(4) = Mid(CodigoBarras, 4, 8)

       'comprobante
       Nitem.SubItems(5) = Mid(CodigoBarras, 14, 9)
    
       'importe al primer vencimiento
       Entero = Mid(CodigoBarras, 23, 6)
       Decim = Mid(CodigoBarras, 29, 2)
       Decim1 = CCur(Decim) / 100
       ImporteVencimiento1 = CCur(Entero) + CCur(Decim1)
       Nitem.SubItems(6) = Format(ImporteVencimiento1, "0.00") & vbNullString
    
       'fecha de primer vencimiento
       FechaJuliana = Mid(CodigoBarras, 31, 5)
       FechaVencimiento1 = CrearFechaNormal(FechaJuliana)
       Nitem.SubItems(7) = FechaVencimiento1
    
      'importe de recargo
      Entero = Mid(CodigoBarras, 36, 4)
      Decim = Mid(CodigoBarras, 40, 2)
      Decim1 = CCur(Decim) / 100
      ImporteRecargo = CCur(Entero) + CCur(Decim1)
      Nitem.SubItems(8) = Format(ImporteRecargo, "0.00") & vbNullString

      'dias al 2º vto
      CadDias = Mid(CodigoBarras, 42, 2)
      Dias = CLng(CadDias)
      Nitem.SubItems(9) = Format(CadDias, "00")
    
      'importe al segundo vencimiento
      ImporteVencimiento2 = CCur(ImporteVencimiento1) + CCur(ImporteRecargo)
    
      Nitem.SubItems(10) = Format(ImporteVencimiento2, "0.00")
      FechaVencimiento2 = FechaVencimiento1 + Dias
      Nitem.SubItems(11) = FechaVencimiento2
    End If 'if ok
    
    'avanzo al siguiente registro
    If Trim$(Mid$(Registro, 68, 6)) = "" Then
        Inicio = Inicio + 75
    Else
        Inicio = Inicio + 79
    End If
Next I

Exit Sub
merror:
tratarerrores "Error en funcion MostrarFacturasRapipago"
End Sub
Private Sub MostrarFacturasLINK(ByVal Archi As String)
'muestra en la segunda lista las facturas de un archivo de red link seleccionado arriba
'solo muestra por pantalla (no registra datos en tablas temporales)
Dim nFic As Integer
Dim CantReg As Long
Dim Scontenido As String
Dim bTitulo As Boolean
Dim cDNI As String
Dim cImporte As String
Dim cFecha As String
Dim Ok As Boolean
Dim Campos() As String


On Error GoTo merror

Lv2.ListItems.Clear

nFic = FreeFile
Open Archi For Input As #nFic

CantReg = 0
bTitulo = True
While Not EOF(nFic)
    Line Input #nFic, Scontenido
    If bTitulo Then
        bTitulo = False
    Else
        Campos = Split(Scontenido, ",")
        cDNI = DepuraString(Campos(1))
        cImporte = DepuraString(Campos(2))
        cFecha = DepuraString(Campos(3))
        Set Nitem = Lv2.ListItems.Add(, , CDate(Mid$(cFecha, 7, 2) & "/" & Mid$(cFecha, 5, 2) & "/" & Mid$(cFecha, 1, 4)))
        Nitem.SubItems(1) = Format(CCur(cImporte) / 100, "0.00") & vbNullString
        Nitem.SubItems(4) = CLng(cDNI) & vbNullString
        CantReg = CantReg + 1
    End If
Wend
TxtContador2.Text = CantReg
Close #nFic

Exit Sub
merror:
tratarerrores "Error en funcion MostrarFacturasLINK"
End Sub
Private Sub MostrarContenido()
'imprime el contenido de los archivos seleccionados en la lista superior
'primero recorre cada archivo selecionado y lo va grabandoi en un
'archivo temporal..luego saca con un sql y se lo pasa a un reporte que solo
'imprime el contenido
Dim Archivo As String
Dim ArchivoSolo As String
Dim TipoArchivo As String
Dim HuboProcesos As Boolean
Dim ProcesoOK As Boolean
Dim Mreporte As New ARContenidoRP
On Error GoTo merror
Call RefreshTimer

If lv.ListItems.Count = 0 Then Exit Sub
If Not HayFilasChequeadas(lv) Then
   MsgE "Debe seleccionar archivos"
   Exit Sub
End If

'inicio transaccion
cnSQL.BeginTrans

Call LimpiarTabla("rapipagotemp")
HuboProcesos = False

For I = 1 To lv.ListItems.Count
    If lv.ListItems.Item(I).Checked Then
       Archivo = lv.ListItems.Item(I).SubItems(1)
       ArchivoSolo = lv.ListItems.Item(I).SubItems(2)
       
       TipoArchivo = DeterminarTipoArchivo(Archivo)
       
       Select Case TipoArchivo
       Case "RP-PF"
            ProcesoOK = MostrarContenidoRP(Archivo, ArchivoSolo)
       Case "LINK"
            ProcesoOK = MostrarContenidoLINK(Archivo, ArchivoSolo)
       End Select
        
       If ProcesoOK Then
            HuboProcesos = True
       End If
    End If ' si esta marcada la fila del lv
Next I

'fin de transaccion
cnSQL.CommitTrans

If HuboProcesos Then
   sql = "select * from rapipagotemp order by archivo,numcomprobante"
   Set rec = cnSQL.OpenResultset(sql)
   If Not rec.EOF Then
      Mreporte.RDODataControl1.Resultset = rec
      Mreporte.LabelTitulo = "Contenido de archivos de RapiPago"
      Mreporte.Show vbModal
   End If
Else
   MsgE "No hay datos para imprimir"
End If

Call RefreshTimer

Exit Sub
merror:
tratarerrores "Error en funcion MostrarContenido"

End Sub

Public Function MostrarContenidoRP(ByVal Archivo As String, ByVal ArchivoSolo As String) As Boolean
'imprime el contenido de los archivos seleccionados en la lista superior
'primero recorre cada archivo selecionado y lo va grabandoi en un
'archivo temporal..luego saca con un sql y se lo pasa a un reporte que solo
'imprime el contenido
Dim sql As String
Dim rec As rdoResultset
Dim FechaArchivo As Date
Dim CantReg As Long
Dim I As Long
Dim J As Long
Dim Inicio As Long
Dim Scontenido As String
Dim nFic As Integer
Dim Header As String
Dim Registro As String
Dim CadFechaCobro As String
Dim CadImporteCobro As String
Dim FechaCobro As Date
Dim ImporteCobrado As Currency
Dim CodigoBarras As String
Dim Año As String
Dim Mes As String
Dim Dia As String
Dim Entero As String
Dim Decim As String
Dim Decim1 As Currency
Dim CadCliente As String
Dim CadEmpresa As String
Dim CadComprobante As String
Dim cDNI As String
Dim CadPrestamo As String
Dim NumEmpresa As Long
Dim NumCliente As Long
Dim numcomprobante As Long
Dim ImporteVencimiento1 As Currency
Dim FechaJuliana As String
Dim FechaVencimiento1 As Date
Dim ImporteRecargo As Currency
Dim CadDias As String
Dim CadDigito As String
Dim DiasVencimiento2 As Long
Dim DigitoVerificador As Long
Dim HuboProcesos As Boolean
Dim NumCuota As Long
Dim Cliente As String
On Error GoTo merror

       Call RefreshTimer
       HuboProcesos = False
       CantReg = ObtenerCantRegistros(Archivo)
       FechaArchivo = ObtenerFechaArchivoRP(ArchivoSolo)
       'abro el archivo correspondiente
       If Len(Dir$(Archivo)) Then
          nFic = FreeFile
          Open Archivo For Input As nFic
          TamFic = LOF(nFic)
          Scontenido = Input$(TamFic, nFic)
          Close nFic
       End If
                 
       cont = 0
       
       Inicio = CLng(TxtInicio.Text)

       'recorro cada archivo
       For J = 1 To CantReg
           Registro = (Mid(Scontenido, Inicio, 80))
           CadFechaCobro = Mid(Registro, 1, 8)
           CadImporteCobro = Mid(Registro, 9, 15)
           CodigoBarras = Mid(Registro, 24, 44)
    
           Año = Mid(CadFechaCobro, 1, 4)
           Mes = Mid(CadFechaCobro, 5, 2)
           Dia = Mid(CadFechaCobro, 7, 2)
           FechaCobro = CDate(Dia + "/" + Mes + "/" + Año)
    
           Entero = Mid(CadImporteCobro, 1, 13)
           Decim = Mid(CadImporteCobro, 14, 2)
           Decim1 = CCur(Decim) / 100
           ImporteCobrado = CCur(Entero) + CCur(Decim1)
           
           'ahora saco los demas campos del codigo de barras
           'numero de empresa
           CadEmpresa = Mid(CodigoBarras, 1, 3)
           NumEmpresa = CLng(CadEmpresa)
           
           CadCliente = Mid(CodigoBarras, 4, 8)
           NumCliente = CLng(CadCliente)
           
           'debe ser del 12,11 pero por overflow lo corro y achico
           CadComprobante = Mid(CodigoBarras, 14, 9)
           numcomprobante = CLng(CadComprobante)
           
           'importe al primer vencimiento
           Entero = Mid(CodigoBarras, 23, 6)
           Decim = Mid(CodigoBarras, 29, 2)
           Decim1 = CCur(Decim) / 100
           ImporteVencimiento1 = CCur(Entero) + CCur(Decim1)
           
           'fecha de primer vencimiento
           FechaJuliana = Mid(CodigoBarras, 31, 5)
           FechaVencimiento1 = CrearFechaNormal(FechaJuliana)
           
           'importe de recargo
           Entero = Mid(CodigoBarras, 36, 4)
           Decim = Mid(CodigoBarras, 40, 2)
           Decim1 = CCur(Decim) / 100
           ImporteRecargo = CCur(Entero) + CCur(Decim1)
           
           'dias al 2º vto
           CadDias = Mid(CodigoBarras, 42, 2)
           DiasVencimiento2 = CLng(CadDias)
           
           'digito verificarod
           CadDigito = Mid(CodigoBarras, 44, 1)
           DigitoVerificador = CLng(CadDigito)
           
           'obtengo el nombre del cliente
           Cliente = ObtenerCliente(NumCliente)
           NumCuota = ObtenerNumCuota(NumCliente, numcomprobante)
           
            If Mid(CadCliente, 1, 2) <> "00" Then
                cDNI = CadCliente
                CadPrestamo = ObtenerCodPrestamoConDocumento(cDNI)
                If CadPrestamo = "" Then
                    CadPrestamo = ObtenerCodPrestamoConDocumentoInclFinalizados(cDNI)
                End If
                CadPrestamo = UCase(Trim(CadPrestamo))
                CadCliente = ""
                IdCredito = 0
                IdCliente = 0
                If CadPrestamo <> "" Then
                    IdCredito = ObtenerCredito(CadPrestamo)
                    IdCliente = ObtenerNumCliente(IdCredito)
                    CadCliente = UCase$(Trim$(ObtenerNombreCliente(IdCliente)))
                End If
                NumCliente = IdCliente
                Cliente = ObtenerCliente(NumCliente)
                numcomprobante = Val(cDNI)
            End If
           
           
           'agrego en la tabla temporal cada uno de los campos del cobro
           sql = "insert into rapipagotemp (archivo,fechaarchivo," & _
                 "numcuota,fechacobro,importecobrado,codigobarras,numempresa," & _
                 "numcliente,cliente,numcomprobante,importevencimiento1," & _
                 "fechavencimiento1,importerecargo,diasvencimiento2," & _
                 "digitoverificador) " & _
                 "values ('" & CStr(ArchivoSolo) & "','" & ConvertirFechaSql(CDate(FechaArchivo), "DD/MM/YYYY") & _
                 "'," & CLng(NumCuota) & ",'" & ConvertirFechaSql(CDate(FechaCobro), "DD/MM/YYYY") & _
                 "'," & ConvertirDblSql(CCur(ImporteCobrado)) & ",'" & CStr(CodigoBarras) & _
                 "'," & CLng(NumEmpresa) & "," & CLng(NumCliente) & _
                 ",'" & CStr(Cliente) & "'," & CLng(numcomprobante) & _
                 "," & ConvertirDblSql(CCur(ImporteVencimiento1)) & _
                 ",'" & ConvertirFechaSql(CDate(FechaVencimiento1), "DD/MM/YYYY") & _
                 "'," & ConvertirDblSql(CCur(ImporteRecargo)) & _
                 "," & CLng(DiasVencimiento2) & "," & CLng(DigitoVerificador) & ")"
           cnSQL.Execute sql
           
           'luego estoy en condiciones de pasarselo al reporte y que lo imprima
           HuboProcesos = True
           
           'incremento la posicion de lectura dentro del archivo
           If Trim$(Mid$(Registro, 68, 6)) = "" Then
                Inicio = Inicio + 75
           Else
                Inicio = Inicio + 79
           End If
       Next J

       Call RefreshTimer
       
       MostrarContenidoRP = HuboProcesos

Exit Function
merror:
tratarerrores "Error en funcion MostrarContenidoRP"

End Function

Public Function MostrarContenidoLINK(ByVal Archivo As String, ByVal ArchivoSolo As String) As Boolean
Dim bTitulo As Boolean
Dim Scontenido As String
Dim nFic As Integer
Dim cDNI As String
Dim cImporte As String
Dim cFecha As String
Dim Campos() As String
Dim Cliente As String
Dim CadPrestamo As String
Dim IdCredito As Long
Dim IdCliente As Long
Dim numcomprobante As Long
Dim ImporteCobrado As Currency
Dim CadCliente As String
Dim FechaArchivo As Date
Dim Mreporte As New ARContenidoRP

        On Error GoTo merror
        Call RefreshTimer
        HuboProcesos = False

        FechaArchivo = ObtenerFechaArchivoLINK(ArchivoSolo)
        nFic = FreeFile
        Open Archivo For Input As #nFic
        bTitulo = True
        While Not EOF(nFic)
            Line Input #nFic, Scontenido
            If bTitulo Then
                bTitulo = False
            Else
                Campos = Split(Scontenido, ",")
                cDNI = DepuraString(Campos(1))
                cImporte = DepuraString(Campos(2))
                ImporteCobrado = CCur(cImporte) / 100
                cFecha = DepuraString(Campos(3))
                CadPrestamo = ObtenerCodPrestamoConDocumento(cDNI)
                If CadPrestamo = "" Then
                    CadPrestamo = ObtenerCodPrestamoConDocumentoInclFinalizados(cDNI)
                End If
                CadPrestamo = UCase(Trim(CadPrestamo))
                CadCliente = ""
                IdCredito = 0
                IdCliente = 0
                If CadPrestamo <> "" Then
                    IdCredito = ObtenerCredito(CadPrestamo)
                    IdCliente = ObtenerNumCliente(IdCredito)
                    CadCliente = UCase$(Trim$(ObtenerNombreCliente(IdCliente)))
                End If
                numcomprobante = Val(cDNI)
                'agrego en la tabla temporal cada uno de los campos del cobro
                sql = "insert into rapipagotemp (archivo,fechaarchivo," & _
                        "fechacobro,importecobrado," & _
                        "numcliente,cliente,numcomprobante)" & _
                        "values ('" & CStr(ArchivoSolo) & "','" & cFecha & _
                        "','" & Mid$(cFecha, 1, 4) & "/" & Mid$(cFecha, 5, 2) & "/" & Mid$(cFecha, 7, 2) & _
                        "'," & ConvertirDblSql(ImporteCobrado) & _
                        "," & CLng(IdCliente) & _
                        ",'" & CStr(CadCliente) & "'," & numcomprobante & ")"
                cnSQL.Execute sql
                HuboProcesos = True
                Call RefreshTimer
            End If
        Wend
        Close #nFic
        
        Call RefreshTimer
        MostrarContenidoLINK = HuboProcesos
Exit Function
merror:
tratarerrores "Error en funcion MostrarContenidoLINK"
End Function

Private Function DepuraString(str) As String
Dim aux As String
Dim J As Integer

str = Trim$(str)
aux = ""
For J = 1 To Len(str)
    If Asc(Mid$(str, J, 1)) <> 61 And Asc(Mid$(str, J, 1)) <> 34 Then
        aux = aux + Mid$(str, J, 1)
    End If
Next
DepuraString = Trim$(aux)
End Function

Private Function DeterminarTipoArchivo(ByVal cArchivo As String) As String
Dim nFic As Integer
Dim Scontenido As String
Dim cPrimerosOcho As String
Dim cTipoArchivo As String

On Error GoTo merror

cTipoArchivo = ""
nFic = FreeFile
Open cArchivo For Input As #nFic

If Not EOF(nFic) Then
    Line Input #nFic, Scontenido
    cPrimerosOcho = Mid$(Scontenido, 1, 8)
    Select Case UCase$(cPrimerosOcho)
    Case "00000000"
        cTipoArchivo = "RP-PF"
    Case "CONCEPTO"
        cTipoArchivo = "LINK"
    End Select
End If
Close #nFic
DeterminarTipoArchivo = cTipoArchivo
Exit Function
merror:
tratarerrores "Error en funcion DeterminarTipoArchivo"

End Function

Private Sub CmdImprimirUltimaImportacion_Click()
Call RefreshTimer
Call ImprimirUltimaImportacionRP
Call RefreshTimer

End Sub
Private Sub ImprimirUltimaImportacionRP()
'imprime la lista facturas que fueron importadas recientemente
'OJO que si aparece una factura no significa que esa este realmente cobrada
'porque cobra con cta cte y si era una muy alta ultima pudo tal vez comenzar
'a cobrar desde las mas viejas y no alcanzar a cubrir esa
Dim sql As String
Dim rec As rdoResultset
Dim Archivo As String
Dim Mreporte As New ARUltimoRP
Dim Titulo As String
On Error GoTo merror

'si imprimo todos los cobros parciales de la cuota
Titulo = "Ultimos cobros importados de RapiPago"

sql = "SELECT clientes.apellido + ' ' + clientes.nombre as cliente," & _
      "rapipago.* " & _
      "from clientes inner join rapipago " & _
      "on clientes.idcliente=rapipago.idcliente " & _
      "order by rapipago.idcliente,rapipago.idcredito,rapipago.numcuota,rapipago.fechacobro"
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Mreporte.RDODataControl1.Resultset = rec
   Mreporte.Caption = "Imprimir ultimos cobros importados de RapiPago"
   Mreporte.LabelTitulo = Titulo
   Mreporte.Show vbModal
Else
   MsgE "No hay importaciones previas para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo ultima importacion RapiPago"
End Sub
