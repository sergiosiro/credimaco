VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCreditosConsultarDeudores 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultar Deudores de Cuotas en Mora"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   HelpContextID   =   19
   Icon            =   "FrmCreditosConsultarDeudores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CheckBloqueados 
      Caption         =   "Incluir cuotas de creditos bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      ToolTipText     =   "Incluye en pantalla, listados y exportaciones a los creditos bloqueados"
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton CmdExportarVeraz 
      Caption         =   "Exportar Veraz"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   35
      ToolTipText     =   "Exporta la planilla de veraz"
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Frame Frame6 
      Caption         =   "Lista de cuotas en mora:"
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   120
      TabIndex        =   33
      Top             =   5160
      Width           =   8895
      Begin MSComctlLib.ListView lvcuotas 
         Height          =   2055
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Muestra las cuotas en mora del credito seleccionado en la lista superior"
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº Credito"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cod.Prestamo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nº cuota"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vencimiento"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DiasMora"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Imp.Mora"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Saldo"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame FrameMenor 
      Caption         =   "Importe Menor (Opcional):"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1920
      TabIndex        =   28
      Top             =   1080
      Width           =   7095
      Begin VB.TextBox TxtImporteMenor 
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   29
         ToolTipText     =   "Importe de comparacion de la columna MENOR de la planilla excel"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Usado para las columnas MENOR de la exportacion de deudores."
         Height          =   375
         Left            =   3960
         TabIndex        =   40
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Descartar Saldos menores a:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "$"
         Height          =   255
         Left            =   3600
         TabIndex        =   30
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.CommandButton CmdBuscarDeudas 
      Height          =   255
      Left            =   4080
      Picture         =   "FrmCreditosConsultarDeudores.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Busca las deudas del cliente seleccionado"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Fecha de consulta:"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   1695
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   55705601
         CurrentDate     =   39964
      End
   End
   Begin VB.CommandButton CmdExportarDeudores 
      Caption         =   "Exportar Deudores"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      ToolTipText     =   "Exporta la lista de deudores con columnas de totales"
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton CmdLibreDeuda 
      Caption         =   "Libre deuda"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      ToolTipText     =   "Cierra la pantalla"
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton CmdImprimirCarta 
      Caption         =   "&Carta Reclamo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Imprime la carta reclamo del credito seleccionado en la lista "
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Opciones de busqueda:"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8895
      Begin VB.ComboBox ComboOpciones 
         Height          =   315
         ItemData        =   "FrmCreditosConsultarDeudores.frx":058C
         Left            =   120
         List            =   "FrmCreditosConsultarDeudores.frx":059F
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione el tipo de consulta"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente:"
         Height          =   735
         Left            =   2520
         TabIndex        =   18
         Top             =   120
         Width           =   6255
         Begin VB.CommandButton CmdBuscarCliente 
            Height          =   375
            Left            =   5640
            Picture         =   "FrmCreditosConsultarDeudores.frx":060B
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Permite seleccionar al cliente de una lista"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox TxtCliente 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   8
            ToolTipText     =   "Nombre del cliente a consultar"
            Top             =   240
            Width           =   5415
         End
      End
      Begin VB.Frame FrameCuotas 
         Caption         =   "Que adeudan mas de:"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   2520
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   6255
         Begin VB.TextBox TxtCuotas 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   2
            Tag             =   "no"
            ToolTipText     =   "Debe ingresar la cantidad de cuotas adeudadas"
            Top             =   240
            Width           =   705
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   840
            TabIndex        =   16
            Top             =   240
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "TxtCuotas"
            BuddyDispid     =   196629
            OrigLeft        =   2400
            OrigTop         =   240
            OrigRight       =   2655
            OrigBottom      =   525
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "cuotas vencidas por credito"
            Height          =   255
            Left            =   1200
            TabIndex        =   17
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame FrameImporte 
         Caption         =   "Importe adeudado:"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   2520
         TabIndex        =   31
         Top             =   120
         Width           =   6255
         Begin VB.TextBox TxtImporteAdeudado 
            Height          =   285
            Left            =   1920
            MaxLength       =   7
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "Por credito"
            Height          =   255
            Left            =   3600
            TabIndex        =   42
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Que adeudan mas de $:"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame FrameDiasMora 
         Caption         =   "Ingrese el rango de maximos de dias de mora:"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   2520
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   6255
         Begin VB.TextBox TxtDiasMora2 
            Height          =   285
            Left            =   2040
            MaxLength       =   5
            TabIndex        =   23
            ToolTipText     =   "Fin del rango de maximos"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtDiasMora 
            Height          =   285
            Left            =   720
            MaxLength       =   5
            TabIndex        =   20
            ToolTipText     =   "Inicio del rango de maximos"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   1560
            TabIndex        =   22
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "(*)Muestra los creditos cuyos maximos de mora estan en el rango."
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   3480
            TabIndex        =   27
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Frame FrameMayorA 
         Caption         =   "Mayor a N dias de mora:"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   2520
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   6255
         Begin VB.TextBox TxtDiasMoraNew 
            Height          =   285
            Left            =   120
            MaxLength       =   7
            TabIndex        =   37
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Muestra los creditos cuyo maximo de mora es superior o igual a esta cantidad."
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1920
            TabIndex        =   43
            Top             =   240
            Width           =   4215
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lista de deudores con creditos en mora:"
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   8895
      Begin VB.TextBox TxtContador1 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "N"
         ToolTipText     =   "Cantidad total de deudores"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox TxtImporteTotal 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   10
         Tag             =   "N"
         ToolTipText     =   "Importe total de las deudas"
         Top             =   2520
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvdeudores 
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Muestra la lista de creditos en mora"
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº Cliente"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Telefono"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nº Credito"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cod.Prestamo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Capital Orig."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "TotalCuotas"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "CuotasMora"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "1º VtoMoroso"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "MoraMaxima"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "Total deudores:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Importe total $:"
         Height          =   255
         Left            =   6120
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   1920
      TabIndex        =   38
      Top             =   1080
      Width           =   7095
   End
End
Attribute VB_Name = "FrmCreditosConsultarDeudores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***AQUI SE CONSULTAN DEUDORES POR DISTINTOS CRITERIOS.ALGUNOS COMPLEJOS

Public IdCliente2 As Long
Public Numlegajocliente2 As Long
Public NumDni As Long
Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call RefrescarOpcionesSistema

Call LimpiarCampos(Me)
ComboOpciones.ListIndex = 0
TxtDiasMora.Text = 1
TxtDiasMora2.Text = 1
TxtDiasMoraNew.Text = 1
TxtImporteAdeudado.Text = ""
DTPicker1.Value = Date

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de consulta de deudores"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
IdCliente2 = 0
Numlegajocliente2 = 0
NumDni = 0
Unload Me
End Sub
Private Sub ComboOpciones_Click()
lvdeudores.ListItems.Clear
lvcuotas.ListItems.Clear

TxtContador1.Text = 0
TxtImporteTotal.Text = 0

CmdImprimirCarta.Enabled = False
CmdLibreDeuda.Enabled = False
CmdExportarDeudores.Enabled = False
CmdExportarVeraz.Enabled = False

FrameCliente.Visible = False
FrameCuotas.Visible = False
FrameDiasMora.Visible = False
FrameImporte.Visible = False
FrameMayorA.Visible = False

If ComboOpciones.Text = "Por cliente" Then
   FrameCliente.Visible = True
   CmdLibreDeuda.Enabled = True
End If

If ComboOpciones.Text = "Por cuotas adeudadas" Then
   FrameCuotas.Visible = True
   TxtCuotas.Text = 1
End If

If ComboOpciones.Text = "Por importe adeudado" Then
   FrameImporte.Visible = True
End If

If ComboOpciones.Text = "Por Maximo dias de mora" Then
   FrameMayorA.Visible = True
End If

If ComboOpciones.Text = "Por rango de Maximos" Then
   FrameDiasMora.Visible = True
End If


End Sub
Private Sub CmdBuscarCliente_Click()
'carga la pantalla de seleccion de clientes
FrmClientesAbm.FormularioPadre = "CONSULTARDEUDORES"
Call CenterForm(FrmClientesAbm)
FrmClientesAbm.Show
End Sub
Private Sub CmdBuscarDeudas_Click()
Dim bHayDeudores        As Boolean

'este boton inicia la busqueda de cuotas
On Error GoTo merror
Call RefreshTimer

CmdBuscarDeudas.Enabled = False

lvdeudores.ListItems.Clear
lvcuotas.ListItems.Clear

If Not datosok() Then
   CmdBuscarDeudas.Enabled = True
   Exit Sub
End If

bHayDeudores = CrearTemporal()
If bHayDeudores Then
   Call CargarLista
   Call CargarCuotasCreditos
Else
   MsgE "No hay deudas segun el criterio indicado"
End If

Call SetearEntorno

CmdBuscarDeudas.Enabled = True

Exit Sub
merror:
tratarerrores "Error en boton BuscarDeudas"
End Sub

Private Sub TxtCliente_Change()
Call CmdBuscarDeudas_Click
End Sub
Private Function CrearTemporal() As Boolean
'genera la tabla temporal con los deudores
'1 paso:genero todos los que tienen creditos
'2-paso:filtro los que tienen deudas(no incluyo los que estan al dia)
'3-paso:agrego estos deudores a una tabla temporal
'paso4-:saco los deudores de esta tabla tempoprao
'paso 5:se los paso a un reporte
Dim sql As String
Dim rec As rdoResultset
Dim CondicionCliente As String
Dim IdCredito As Long
Dim CuotasVencidas As Long
Dim SaldoVencido As Currency
Dim MoraVencida As Currency
Dim SaldoCredito As Currency
Dim SaldoOriginal As Currency
Dim ProvCredito As String
Dim IdCliente As Long
Dim Cliente As String
Dim FechaCredito As Date
Dim Numlegajo As String
Dim Telefono As String
Dim NumCuotas As Long
Dim NumDocumento As String
Dim CodPrestamo As String
Dim Domicilio As String
Dim Localidad As String
Dim Provincia As String
Dim CodigoPostal As String
Dim Sexo As String
Dim Nacionalidad As String
Dim FechaNacimiento As String
Dim Comercio As String
Dim CapitalOriginal As Currency
Dim OkGrabarDeudor As Boolean
Dim MaxDiasMora As Long
Dim ImporteMenor As Currency
Dim NumCuota As Long
Dim VencimientoMoroso As Date
Dim CondicionBloqueados As String
Dim CondicionZ As String
Dim bSeguir As Boolean
Dim cnSQL2 As rdoConnection

On Error GoTo merror

CrearTemporal = False

Call LimpiarTablaSP("templistadeudores")

If ComboOpciones.Text = "Por cliente" Then
   CondicionCliente = CLng(IdCliente2)
Else
   'todos los clientes
   CondicionCliente = 0
End If

If CheckBloqueados.Value = 1 Then
   CondicionBloqueados = 1
Else
   CondicionBloqueados = 0
End If

sql = "CrearTemporalDeudores " & CondicionCliente & "," & CondicionBloqueados

'sql = "select localidades.nombre as localidad,localidades.codigopostal," & _
'      "provincias.nombre as provincia," & _
'      "clientes.idcliente,clientes.numlegajo,clientes.telefono," & _
'      "clientes.domicilio,clientes.fechanacimiento," & _
'      "clientes.apellido + ', ' + clientes.nombre as cliente," & _
'      "clientes.cad1 as sexo,clientes.numdocumento,clientes.nacionalidad," & _
'      "creditos.idcredito,creditos.codprestamo,creditos.fechacredito," & _
'      "creditos.numcuotas,creditos.motivobloqueo as comercio,creditos.importeafinanciar " & _
'      "from provincias inner join (localidades inner join " & _
'      "(clientes inner join creditos on clientes.idcliente=creditos.idcliente) " & _
'      "on localidades.idlocalidad=clientes.idlocalidad) " & _
'      "on provincias.idprovincia=localidades.idprovincia " & _
'      "where " & CondicionCliente & _
'      " and creditos.fechafinalizacion is Null and " & _
'      CondicionBloqueados & _
'      " order by creditos.codprestamo"

Set cnSQL2 = enSQL.OpenConnection("", rdDriverNoPrompt, False, cConexion)
Set rec = cnSQL2.OpenResultset(sql)
'borro siempre aunque no encuentre clientes con deudas
'porque sino las demas funciones pueden encontrar datos
'de consultas previas

ImporteMenor = 0
If MenorOk() Then
    ImporteMenor = TxtImporteMenor.Text
End If

If Not rec.EOF Then

   'recorro la lista agregando los deudores
   Do While Not rec.EOF
   
      Call RefreshTimer
      
      cnSQL.BeginTrans

      IdCredito = rec.rdoColumns("idcredito")
      bSeguir = True

      'verifico si el cliente tiene cuotas en mora
      CuotasVencidas = ObtenerCuotasVencidas(IdCredito, DTPicker1.Value)

      If (CuotasVencidas) > 0 Then

        If ComboOpciones.Text = "Por cuotas adeudadas" Then
            If CuotasVencidas < CLng(TxtCuotas.Text) Then
               bSeguir = False
            End If
         End If
      
         If bSeguir Then
             
            VencimientoMoroso = "2099-01-01"
            MaxDiasMora = ObtenerMaxDiasMora(IdCredito, DTPicker1.Value, ImporteMenor, VencimientoMoroso)

            If ComboOpciones.Text = "Por Maximo dias de mora" Then
                'trae los dias de mora de la cuota mas vieja en mora
                If MaxDiasMora < CLng(TxtDiasMoraNew.Text) Then
                    bSeguir = False
                End If
            End If
         
            'rango de máximos es la más usada
            If ComboOpciones.Text = "Por rango de Maximos" Then
                If MaxDiasMora < CLng(TxtDiasMora.Text) Or MaxDiasMora > CLng(TxtDiasMora2.Text) Then
                    bSeguir = False
                End If
            End If
             
             
            If bSeguir Then
            
            'inicializo en 0 el saldo vencido para grabarlo en la tabla de salida (no se usa ni en la exportacion
            'ni en la consulta)
             SaldoVencido = 0
             If ComboOpciones.Text = "Por importe adeudado" Then
                'saldo de las cuotas en mora(todos los items mas mora)(optimizada ok)
                 SaldoVencido = ObtenerSaldoVencido(IdCredito, DTPicker1.Value)
                 If CCur(SaldoVencido) < CCur(TxtImporteAdeudado.Text) Then
                    bSeguir = False
                 End If
             End If
             
             If bSeguir Then
         
                 'la mora sola de las cuotas en mora(mora e ivamora)
                 'se pasa el calculo de la mora vencida a la exportacion
                 'MoraVencida = ObtenerMoraVencida(IdCredito, DTPicker1.Value)
                 MoraVencida = 0
    
                 'saldo total del credito(optimizada ok)
                 'trae de vigentes,finalizados,bloqueados etc por eso no requiere la condicionbloqueados
'este campo no se usa ni en la pantalla ni en la exportación, asi que lo dejo en cero
'                 SaldoCredito = ObtenerSaldoCredito(IdCredito, DTPicker1.Value)
                 SaldoCredito = 0
'este campo no se usa ni en la pantalla ni en la exportación, asi que lo dejo en cero
'                 SaldoOriginal = CCur(SaldoVencido) - CCur(MoraVencida)
                 SaldoOriginal = 0
                 'no requiere la condicionbloqueados
                 'paso este cálculo a ObtenerMaxDiasMora
                 'VencimientoMoroso = ObtenerVtoMoroso(rec.rdoColumns("idcredito"), DTPicker1.Value, ImporteMenor)
             
                 'provincia del credito
                 ProvCredito = ObtenerProvinciaCredito(IdCredito)
                 IdCliente = rec.rdoColumns("idcliente")
                 Cliente = rec.rdoColumns("cliente") & vbNullString
                 FechaCredito = CDate(rec.rdoColumns("fechacredito"))
                 Numlegajo = rec.rdoColumns("numlegajo") & vbNullString
                 Telefono = rec.rdoColumns("telefono") & vbNullString
                 NumCuotas = rec.rdoColumns("numcuotas")
                 NumDocumento = rec.rdoColumns("numdocumento") & vbNullString
                 CodPrestamo = rec.rdoColumns("codprestamo") & vbNullString
                 Domicilio = rec.rdoColumns("domicilio") & vbNullString
'este campo no se usa ni en la pantalla ni en la exportación, asi que lo dejo en blancos
'                 Localidad = rec.rdoColumns("localidad") & vbNullString
                 Localidad = ""
'este campo no se usa ni en la pantalla ni en la exportación, asi que lo dejo en blancos
                 'esta es la provincia del cliente...no del credito
                 'Provincia = rec.rdoColumns("provincia") & vbNullString
                 Provincia = ""
'este campo no se usa ni en la pantalla ni en la exportación, asi que lo dejo en blancos
                 'CodigoPostal = rec.rdoColumns("codigopostal") & vbNullString
                 CodigoPostal = ""
                 Sexo = rec.rdoColumns("sexo") & vbNullString
                 Nacionalidad = rec.rdoColumns("nacionalidad") & vbNullString
                 FechaNacimiento = rec.rdoColumns("fechanacimiento")
                 Comercio = rec.rdoColumns("comercio") & vbNullString
                 CapitalOriginal = rec.rdoColumns("importeafinanciar")
                      
                'creo la tabla de deudores (no de cuotas) sino clientes/creditos
                sql = "InsertarTemplistadeudores " & CLng(IdCliente) & ",'" & CStr(Numlegajo) & "'," & _
                      CLng(IdCredito) & ",'" & CStr(CodPrestamo) & "'," & _
                      "'" & ConvertirFechaSql(CDate(FechaCredito), "DD/MM/YYYY") & "','" & CStr(Cliente) & "'," & _
                      CLng(CuotasVencidas) & "," & CLng(MaxDiasMora) & ",'" & CStr(Telefono) & "'," & _
                      ConvertirDblSql(CCur(SaldoOriginal)) & "," & ConvertirDblSql(CCur(MoraVencida)) & "," & _
                      ConvertirDblSql(CCur(SaldoVencido)) & "," & ConvertirDblSql(CCur(SaldoCredito)) & "," & _
                      CLng(NumCuotas) & "," & _
                      "'" & CStr(NumDocumento) & "','" & CStr(Domicilio) & "','" & CStr(Localidad) & "','" & _
                      CStr(Provincia) & "','" & CStr(CodigoPostal) & "','" & CStr(Sexo) & "','" & CStr(Nacionalidad) & _
                      "','" & CStr(FechaNacimiento) & "','" & CStr(Comercio) & "'," & ConvertirDblSql(CCur(CapitalOriginal)) & _
                      ",'" & CStr(ProvCredito) & "','" & ConvertirFechaSql(CDate(VencimientoMoroso), "DD/MM/YYYY") & "'"
    '            sql = "insert into templistadeudores (idcliente,numlegajo," & _
    '                  "idcredito,codprestamo,fechacredito,cliente,cuotasenmora,maxdiasmora," & _
    '                  "telefono,deudareal,importemora,saldoenmora,saldocredito," & _
    '                  "diasmora,numcuotas,numdocumento,domicilio,localidad," & _
    '                  "provincia,codigopostal,sexo,nacionalidad,fechanacimiento,comercio,capital,provcredito,vencimiento) " & _
    '                  "values (" & CLng(IdCliente) & ",'" & CStr(Numlegajo) & "'," & _
    '                  CLng(IdCredito) & ",'" & CStr(CodPrestamo) & "'," & _
    '                  "'" & ConvertirFechaSql(CDate(FechaCredito), "DD/MM/YYYY") & "','" & CStr(Cliente) & "'," & _
    '                  CLng(CuotasVencidas) & "," & CLng(MaxDiasMora) & ",'" & CStr(Telefono) & "'," & _
    '                  ConvertirDblSql(CCur(SaldoOriginal)) & "," & ConvertirDblSql(CCur(MoraVencida)) & "," & _
    '                  ConvertirDblSql(CCur(SaldoVencido)) & "," & ConvertirDblSql(CCur(SaldoCredito)) & ",0," & _
    '                  CLng(NumCuotas) & "," & _
    '                  "'" & CStr(NumDocumento) & "','" & CStr(Domicilio) & "','" & CStr(Localidad) & "','" & CStr(Provincia) & "','" & CStr(CodigoPostal) & "','" & CStr(Sexo) & "','" & CStr(Nacionalidad) & "','" & CStr(FechaNacimiento) & "','" & CStr(Comercio) & "'," & ConvertirDblSql(CCur(CapitalOriginal)) & ",'" & CStr(ProvCredito) & "','" & ConvertirFechaSql(CDate(VencimientoMoroso), "DD/MM/YYYY") & "')"
                cnSQL.Execute (sql)
                cnSQL.CommitTrans
                CrearTemporal = True
             End If
            End If
         End If
      End If
      rec.MoveNext
   Loop
End If

cnSQL2.Close

Exit Function
merror:
tratarerrores "Error en funcion CrearTemporal"
End Function
Private Function ObtenerMoraVencida(ByVal IdCredito As Long, ByVal Fecha As Date) As Currency
'obtiene el importe en mora de un credito(sin base)
Dim sql As String
Dim rec As rdoResultset
Dim ImporteMora As Currency
Dim Suma As Currency
Dim ImporteParcial As Currency
Dim SaldoCuota As Currency
Dim FechaTramo As Date
Dim MoraCobrada As Currency
Dim IvaMora As Currency
Dim NumCuota As Long
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim IvaACobrarDevuelto  As Currency
Dim Importe1erVenc  As Currency
On Error GoTo merror

ObtenerMoraVencida = 0

sql = "ObtenerMoraVencida " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"

'sql = "select creditos.idcredito," & _
'      "cuotas.numcuota,cuotas.fechavencimiento1," & _
'      "cuotas.logic1 as exceptuada " & _
'      "from creditos inner join cuotas " & _
'      "on creditos.idcredito=cuotas.idcredito " & _
'      "where creditos.idcredito=" & CLng(IdCredito) & _
'      " and cuotas.fechacobro is Null " & _
'      "and cuotas.cuotacomodin = 0 " & _
'      "and cuotas.logic1 = 0 " & _
'      "and cuotas.fecharefinanciacion is Null " & _
'      "and cuotas.fechavencimiento1<'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "' " & _
'      "and creditos.fechafinalizacion is Null"
     
Set rec = cnSQL.OpenResultset(sql)

ImporteMora = 0
Suma = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      NumCuota = CLng(rec.rdoColumns("numcuota"))
    
      'este saldo es sin mora
'      SaldoCuota2 = ObtenerSaldoCuotaX(IdCredito, NumCuota, Fecha, SaldoCuota)
            
      'si hay mora
      ImporteMora = 0
      IvaMora = 0
      If Not rec.rdoColumns("exceptuada") Then
         'calculo la mora de la forma habitual
         'puedo pasarle el campo [exceptuada]
         Importe1erVenc = ObtenerImporte1erVenc(rec.rdoColumns("idcredito"), rec.rdoColumns("NumCuota"))
         ImporteMora = CalculoMoraPendiente(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("exceptuada"), Importe1erVenc, rec.rdoColumns("fechavencimiento1"), Fecha, IvaACobrarDevuelto)
         '''''''********ImporteMora = CalcularInteresMoraZZ(rec.rdoColumns("exceptuada"), SaldoCuota, rec.rdoColumns("fechavencimiento1"), Fecha)
         IvaMora = 0
         If VG_APLICARIMPUESTOS Then
            If VG_IMPUESTOSCREDIMACO Then
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
      End If

      Suma = CCur(Suma) + CCur(ImporteMora) + CCur(IvaMora)
      
      rec.MoveNext
      
   Loop
End If

ObtenerMoraVencida = CCur(Suma)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerMoraVencida"
End Function
Private Function ObtenerDiasMoraCredito(ByVal IdCredito As Long, ByVal Fecha As Date) As Currency
'solo la mora de las cuotas vencidas
Dim sql As String
Dim rec As rdoResultset
Dim Suma As Long
Dim Dias As Long
On Error GoTo merror

ObtenerDiasMoraCredito = 0

sql = "select creditos.idcredito," & _
      "cuotas.fechavencimiento1 " & _
      "from creditos inner join cuotas " & _
      "on creditos.idcredito=cuotas.idcredito " & _
      "where creditos.idcredito=" & CLng(IdCredito) & " " & _
      "and cuotas.fechacobro is null " & _
      "and cuotas.cuotacomodin  = 0 " & _
      "and cuotas.logic1 = 0 " & _
      "and cuotas.fecharefinanciacion is Null " & _
      "and cuotas.fechavencimiento1<'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "' " & _
      "and creditos.fechafinalizacion is Null"

Set rec = cnSQL.OpenResultset(sql)

Suma = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      'obtengo los dias de mora de cada cuota
      Dias = CDate(Fecha) - CDate(rec.rdoColumns("fechavencimiento1"))
      Suma = Suma + Dias
      
      rec.MoveNext
   Loop
End If

ObtenerDiasMoraCredito = CCur(Suma)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerDiasMoraCredito"
End Function
Private Function ObtenerMaxDiasMora(ByVal IdCredito As Long, ByVal Fecha As Date, ByVal ImporteMenor As Currency, ByRef VencimientoMoroso As Date) As Long
'obtiene la maxima cantidad de dias de mora de un credito entre sus cuotas
Dim rec As rdoResultset
Dim sql As String
Dim Maximo As Long
On Error GoTo merror

ObtenerMaxDiasMora = 0

Maximo = 0

'si hay importe menor debo buscar el maximo de dias de mora
'pero solo de las cuotas cuyo saldo supera el importe menor
sql = "ObtenerMaxDiasMora " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
'sql = "select datediff(dd,cuotas.fechavencimiento1,'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "') as diasmora," & _
'      "cuotas.idcredito,cuotas.numcuota,cuotas.fechavencimiento1,cuotas.fechavencimiento2,cuotas.logic1 as exceptuada " & _
'      "from cuotas " & _
'      "where idcredito=" & CLng(IdCredito) & " " & _
'      "and fechacobro is Null " & _
'      "and cuotacomodin = 0 " & _
'      "and logic1 = 0 " & _
'      "and fecharefinanciacion is Null " & _
'      "and cuotas.fechavencimiento1<'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
          
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Do While Not rec.EOF
      SaldoCuota = ObtenerSaldoCuotaOKK(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("fechavencimiento1"), rec.rdoColumns("fechavencimiento2"), rec.rdoColumns("exceptuada"), Fecha)
      'si la cuota supera el saldo busco entre esas
      If CCur(SaldoCuota) > CCur(ImporteMenor) Then
         'busco un maximo
         If rec.rdoColumns("diasmora") > Maximo Then
            Maximo = rec.rdoColumns("diasmora")
         End If
         If rec.rdoColumns("fechavencimiento1") < VencimientoMoroso Then
            VencimientoMoroso = rec.rdoColumns("fechavencimiento1")
         End If
      End If
      rec.MoveNext
   Loop
End If

ObtenerMaxDiasMora = Maximo

Exit Function
merror:
tratarerrores "Error en funcion ObtenerMaxDiasMora"
End Function
Private Function ObtenerSoloMoraCuota(ByVal IdCredito As Long, ByVal NumCuota As Long, ByVal FechaVencimiento1 As Date, ByVal Exceptuada As Boolean, ByVal Fecha As Date) As Currency
'obtiene solo la mora de una cuota
'para mostrar en la lista de cuotas
Dim ImporteMora As Currency
Dim Suma As Currency
Dim IvaMora As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim IvaACobrarDevuelto  As Currency
On Error GoTo merror

ObtenerSoloMoraCuota = 0
      
ImporteMora = 0
IvaMora = 0
Suma = 0

            
'si hay mora
If Not Exceptuada Then
   'calculo la mora de la forma habitual
   'puedo pasarle el campo [exceptuada]
   Importe1erVenc = ObtenerImporte1erVenc(IdCredito, NumCuota)
   ImporteMora = CalculoMoraPendiente(IdCredito, NumCuota, Exceptuada, Importe1erVenc, FechaVencimiento1, Fecha, IvaACobrarDevuelto)
   '''''''********ImporteMora = CalcularInteresMoraZZ(Exceptuada, SaldoCuota, FechaVencimiento1, Fecha)
   'falta el iva mora
   If VG_APLICARIMPUESTOS Then
      If VG_IMPUESTOSCREDIMACO Then
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
End If

Suma = CCur(ImporteMora) + CCur(IvaMora)

ObtenerSoloMoraCuota = CCur(Suma)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSoloMoraCuota"
End Function
Private Sub CargarLista()
'carga la lista de deudores de cuotas
Dim sql As String
Dim rec As rdoResultset
Dim Nitem As ListItem
Dim Total As Currency
Dim cont As Long
On Error GoTo merror

If Not datosok() Then Exit Sub

sql = "SeleccionarTemplistadeudores"

Set rec = cnSQL.OpenResultset(sql)
Call RefreshTimer

Total = 0
cont = 0
If Not rec.EOF Then
      Do While Not rec.EOF

         Set Nitem = lvdeudores.ListItems.Add(, , rec.rdoColumns("numlegajo"))
         Nitem.SubItems(1) = rec.rdoColumns("cliente") & vbNullString
         Nitem.SubItems(2) = rec.rdoColumns("telefono") & vbNullString
         Nitem.SubItems(3) = Format(rec.rdoColumns("idcredito"), "000000") & vbNullString
         Nitem.SubItems(4) = Format(rec.rdoColumns("codprestamo"), "000000") & vbNullString
         Nitem.SubItems(5) = Format(rec.rdoColumns("capital"), "0.00") & vbNullString
         Nitem.SubItems(6) = Format(rec.rdoColumns("numcuotas"), "000") & vbNullString
         Nitem.SubItems(7) = Format(rec.rdoColumns("cuotasenmora"), "000") & vbNullString
         Nitem.SubItems(8) = rec.rdoColumns("vencimiento")
         'mora mas antigua
         Nitem.SubItems(9) = Format(rec.rdoColumns("maxdiasmora"), "0000") & vbNullString
         'se dejan de mostrar en pantalla los campos: deudareal, importemora, saldoenmora
         'Nitem.SubItems(10) = Format(rec.rdoColumns("deudareal"), "0.00")
         'Nitem.SubItems(11) = Format(rec.rdoColumns("importemora"), "0.00")
         'Nitem.SubItems(12) = Format(rec.rdoColumns("saldoenmora"), "0.00")
 
         Total = Total + CCur(rec.rdoColumns("saldoenmora"))
         cont = cont + 1
         rec.MoveNext
         Call RefreshTimer
      Loop
Else
   MsgE "No hay deudores segun el criterio indicado"
End If

TxtContador1.Text = cont
TxtImporteTotal.Text = Format(Total, "0.00")

Exit Sub
merror:
tratarerrores "Error en funcion CargarLista"
End Sub
Private Sub lvdeudores_ItemClick(ByVal Item As MSComctlLib.ListItem)
'al clickerar la lista de deudores muestra las cuotas abajo
If lvdeudores.ListItems.Count = 0 Then Exit Sub
Call CargarCuotasCreditos
End Sub
Private Sub CargarCuotasCreditos()
'carga las cuotas del credito indicado
'tambien sirve para la carta
Dim rec As rdoResultset
Dim Nitem As ListItem
Dim ImporteMora As Currency
Dim IvaMora As Currency
Dim ImporteActualizado As Currency
Dim ImporteVencimiento2 As Currency
Dim ImporteTotal As Currency
Dim TotalMora As Currency
Dim TotalParcial As Currency
Dim ImporteParcial As Currency
Dim SaldoCuota As Currency
Dim ImporteCobrado As Currency
Dim Cad As String
Dim I As Long
Dim Vencimiento As Date
Dim InteresCuota As Currency
Dim RecargoVto2 As Currency
Dim MoraCobrada As Currency
Dim IdCredito As Long
Dim NumCuota As Long
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim Suma As Currency
Dim SaldoActual As Currency
Dim CondicionBloqueados As Integer
On Error GoTo merror

lvcuotas.ListItems.Clear

If lvdeudores.ListItems.Count() = 0 Then Exit Sub
IdCredito = lvdeudores.SelectedItem.SubItems(3)

If CheckBloqueados.Value = 1 Then
   CondicionBloqueados = 1
Else
   CondicionBloqueados = 0
End If
'esta trae solo las cuotas en mora del credito indicado
Set rec = CargarRecCuotasMora(IdCredito, DTPicker1.Value, CondicionBloqueados)

I = 1
Do While Not rec.EOF
   ImporteMora = 0
     
   Set Nitem = lvcuotas.ListItems.Add(, , Format(rec.rdoColumns("idcredito"), "000000"))
   Nitem.SubItems(1) = rec.rdoColumns("codprestamo") & vbNullString
   Nitem.SubItems(2) = Format(rec.rdoColumns("numcuota"), "00") & vbNullString
   Nitem.SubItems(3) = rec.rdoColumns("fechavencimiento1") & vbNullString
   DiasMora = DTPicker1.Value - rec.rdoColumns("fechavencimiento1")
   Nitem.SubItems(4) = Format(DiasMora, "0000")
   Nitem.SubItems(5) = Format(rec.rdoColumns("importevencimiento1"), "0.00")
   ImporteMora = ObtenerSoloMoraCuota(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("fechavencimiento1"), rec.rdoColumns("exceptuada"), DTPicker1.Value)
   Nitem.SubItems(6) = Format(ImporteMora, "0.00")
   SaldoActual = ObtenerSaldoCuotaOKK(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("fechavencimiento1"), rec.rdoColumns("fechavencimiento2"), rec.rdoColumns("exceptuada"), DTPicker1.Value)
   Nitem.SubItems(7) = Format(SaldoActual, "0.00")
            
   rec.MoveNext
   Call RefreshTimer
   
   I = I + 1
Loop

Exit Sub
merror:
tratarerrores "Error cargando cuotas de creditos"
End Sub
Private Function CargarRecCuotasMora(ByVal IdCredito As Long, ByVal Fecha As Date, ByVal CondicionBloqueados As Integer) As rdoResultset
'tiene que atraer las cuotas en mora de un credito a la fecha
Dim sql As String
On Error GoTo merror
  
sql = "CargarRecCuotasMora " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'," & CondicionBloqueados
'sql = "select provincias.nombre as provincia,localidades.nombre as localidad," & _
'      "localidades.codigopostal,clientes.domicilio,clientes.telefono," & _
'      "clientes.numlegajo,clientes.apellido + ', ' + clientes.nombre as cliente," & _
'      "creditos.idcredito,creditos.numcuotas,creditos.codprestamo,cuotas.numcuota,cuotas.cobrosparciales," & _
'      "cuotas.fechavencimiento1,cuotas.fechavencimiento2," & _
'      "cuotas.fechacobro,cuotas.importecobrado,cuotas.logic1 as exceptuada," & _
'      "cuotas.importevencimiento1,cuotas.importevencimiento2 " & _
'      "from provincias inner join (localidades inner join " & _
'      "(clientes inner join (creditos inner join cuotas on " & _
'      "creditos.idcredito=cuotas.idcredito) " & _
'      "on clientes.idcliente=creditos.idcliente) on " & _
'      "localidades.idlocalidad=clientes.idlocalidad) on " & _
'      "provincias.idprovincia=localidades.idprovincia " & _
'      "where cuotas.idcredito=" & CLng(IdCredito) & " " & _
'      "and cuotas.fechacobro is Null " & _
'      "and cuotas.cuotacomodin = 0 " & _
'      "and cuotas.logic1 = 0 " & _
'      "and cuotas.fecharefinanciacion is Null " & _
'      "and cuotas.fechavencimiento1<'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "' " & _
'      "and creditos.fechafinalizacion is Null " & _
'      "and " & CondicionBloqueados & _
'      " order by cuotas.numcuota"
                  
Set CargarRecCuotasMora = cnSQL.OpenResultset(sql)

Exit Function
merror:
tratarerrores "Error cargando el registro de cuotas en mora"
End Function
Private Function ObtenerVtoMoroso(ByVal IdCredito As Long, ByVal Fecha As Date, ByVal ImporteMenor As Currency) As Date
'devuelve el vencimiento moroso mas antiguo de un credito
'no de la cuota mas vieja que aparece en pantalla
Dim sql As String
Dim rec As rdoResultset
Dim SaldoCuota As Currency
Dim Minimo As Date

Dim auxauxcredito As Long

On Error GoTo merror

ObtenerVtoMoroso = Date

'TRAE EL VENCIMIENTO MAS VIEJO DEL CREDITO
'aunque esa cuota no aparezca listada
sql = "ObtenerVtoMoroso " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
'sql = "select cuotas.idcredito,cuotas.numcuota," & _
'      "cuotas.fechavencimiento1,cuotas.fechavencimiento2," & _
'      "cuotas.logic1 as exceptuada " & _
'      "from cuotas " & _
'      "where idcredito=" & CLng(IdCredito) & " " & _
'      "and fechacobro is Null " & _
'      "and cuotacomodin = 0 " & _
'      "and cuotas.logic1 = 0 " & _
'      "and fecharefinanciacion is Null " & _
'      "and cuotas.fechavencimiento1<'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
      
Set rec = cnSQL.OpenResultset(sql)

'parto desde dat para encontrar menores si o si
'o podria poner el vencimiento maximo del credito indicado

If Not rec.EOF Then
    auxauxcredito = IdCredito
    Minimo = ObtenerMaxVtoCredito(rec.rdoColumns("idcredito"), Fecha)
    
   Do While Not rec.EOF
   
      SaldoCuota = ObtenerSaldoCuotaOKK(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("fechavencimiento1"), rec.rdoColumns("fechavencimiento2"), rec.rdoColumns("exceptuada"), Fecha)
      
      'si la cuota supera el saldo busco entre esas
      If CCur(SaldoCuota) > CCur(ImporteMenor) Then
      
         If rec.rdoColumns("fechavencimiento1") < Minimo Then
         
            Minimo = rec.rdoColumns("fechavencimiento1")
         End If
      End If
    
      rec.MoveNext
   Loop
   
   ObtenerVtoMoroso = Minimo

End If
Exit Function
merror:
tratarerrores "Error en funcion ObtenerVencimientoMoroso " & auxauxcredito
End Function
Private Function ObtenerMaxVtoCredito(ByVal IdCredito As Long, ByVal Fecha As Date) As Date
'el maximo vencimiento de un credito
'teoricamente de la cuota mas reciente en mora
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

'parto desde este
ObtenerMaxVtoCredito = Date

sql = "ObtenerMaxVtoCredito " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
'sql = "select max(cuotas.fechavencimiento1) as maximo " & _
'      "from cuotas " & _
'      "where cuotas.idcredito=" & CLng(IdCredito) & " " & _
'      "and fechacobro is Null " & _
'      "and cuotacomodin = 0 " & _
'      "and cuotas.logic1  = 0" & _
'      "and fecharefinanciacion is Null " & _
'      "and cuotas.fechavencimiento1<'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"

Set rec = cnSQL.OpenResultset(sql)
If Not rec.EOF Then
   ObtenerMaxVtoCredito = rec.rdoColumns("maximo")
End If


Exit Function
merror:
tratarerrores "Error en funcion ObtenerMaxVtoCredito"
End Function
Private Function ObtenerSaldoCapVencido(ByVal IdCredito As Long, ByVal Fecha As Date) As Currency
'devuelve el saldo solo del ITEM CAPITAL vencido
'tiene en cienta el IMPORTE MENOR ingresado opcionalmente
Dim sql As String
Dim rec As rdoResultset
Dim rec2 As rdoResultset
Dim CapitalCobrado As Currency
On Error GoTo merror

ObtenerSaldoCapVencido = 0

'obtengo la suma del capital original de las cuotas en mora vencidas
'sql = "select sum(cuotas.importeamortizacion) as sumacapital " & _
'      "from cuotas " & _
'      "where idcredito='" & CLng(IdCredito) & "' " & _
'      "and fechacobro is Null " & _
'      "and cuotacomodin = 'False' " & _
'      "and cuotas.logic1 = 'False' " & _
'      "and fecharefinanciacion is Null " & _
'      "and cuotas.fechavencimiento1<'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
sql = "ObtenerSaldoCapVencido " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("sumacapital")) Then
      'busco el capital cobrado hasta la fecha de ese credito
      'de las cuotas vencidas
      
      'esto no iria mas si es solo de cuotas de pantalla
      'sql = "select sum(ingresos.capitalcobrado) as sumacobrado " & _
      '      "from cuotas inner join ingresos " & _
      '      "on cuotas.idcredito=ingresos.idcredito and cuotas.numcuota=ingresos.numcuota " & _
      '      "where ingresos.idcredito='" & CLng(IdCredito) & "' " & _
      '      "and cuotas.fechacobro is Null " & _
      '      "and cuotas.cuotacomodin = 'False' " & _
      '      "and cuotas.logic1 = 'False' " & _
      '      "and cuotas.fecharefinanciacion is null " & _
      '      "and cuotas.fechavencimiento1< '" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "' "
      sql = "ObtenerSaldoCapVencidoCobrado " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
        
      Set rec2 = cnSQL.OpenResultset(sql)
      CapitalCobrado = 0
      If Not rec2.EOF Then
         If Not IsNull(rec2.rdoColumns("sumacobrado")) Then
            CapitalCobrado = CCur(rec2.rdoColumns("sumacobrado"))
         End If
      End If
      ObtenerSaldoCapVencido = CCur(rec.rdoColumns("sumacapital")) - CCur(CapitalCobrado)
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSaldoCapitalVencido"
End Function
Private Function ObtenerSaldoCapAVencer(ByVal IdCredito As Long, ByVal Fecha As Date) As Currency
'devuelve el capital de las cuotas a vencer impagas
'tiene en cuenta los cobros parciales??
'si es asi deberia restarle todos los cobros parciales de capital
'de la tabla de ingresos
Dim sql As String
Dim rec As rdoResultset
Dim rec2 As rdoResultset
Dim CapitalCobrado As Currency
On Error GoTo merror

ObtenerSaldoCapAVencer = 0

'TRAE CUOTAS DE TODO EL CREDITO
'sql = "select sum(cuotas.importeamortizacion) as sumacapital " & _
'      "from cuotas " & _
'      "where idcredito='" & CLng(IdCredito) & "' " & _
'      "and fechacobro is Null " & _
'      "and cuotacomodin = 'False' " & _
'      "and cuotas.logic1 = 'False' " & _
'      "and fecharefinanciacion is Null " & _
'      "and cuotas.fechavencimiento1>='" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "' "
sql = "ObtenerSaldoCapAVencer " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   
   If Not IsNull(rec.rdoColumns("sumacapital")) Then
      'busco el capital cobrado hasta la fecha de ese credito
      'de las cuotas a vencer
      
'      sql = "select sum(ingresos.capitalcobrado) as sumacobrado " & _
'            "from cuotas inner join ingresos " & _
'            "on cuotas.idcredito=ingresos.idcredito and cuotas.numcuota=ingresos.numcuota " & _
'            "where ingresos.idcredito='" & CLng(IdCredito) & "' " & _
'            "and cuotas.fechacobro is Null " & _
'            "and cuotas.cuotacomodin = 'False' " & _
'            "and cuotas.logic1 = 'False' " & _
'            "and cuotas.fecharefinanciacion is Null " & _
'            "and cuotas.fechavencimiento1>='" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "' "
      sql = "ObtenerSaldoCapAVencerCobrado " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
      Set rec2 = cnSQL.OpenResultset(sql)
      
      CapitalCobrado = 0
      If Not rec2.EOF Then
         If Not IsNull(rec2.rdoColumns("sumacobrado")) Then
            CapitalCobrado = CCur(rec2.rdoColumns("sumacobrado"))
         End If
      End If
      ObtenerSaldoCapAVencer = CCur(rec.rdoColumns("sumacapital")) - CCur(CapitalCobrado)
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSaldoCapitalAVencer"
End Function
Private Function ObtenerSaldoTotalVencido(ByVal IdCredito As Long, ByVal Fecha As Date) As Currency
'calcula el saldo gral de un credito teniendo en cta solo las cuotas vencidas
Dim sql As String
Dim rec As rdoResultset
Dim SaldoCuota As Currency
Dim ImporteTotal As Currency
On Error GoTo merror

ObtenerSaldoTotalVencido = 0

'trae solo las cuotas vencidas e impagas a la fecha de hoy
'sql = "select creditos.idcredito,cuotas.numcuota,cuotas.fechavencimiento1,cuotas.fechavencimiento2,cuotas.logic1 as exceptuada " & _
'      "from creditos inner join cuotas on creditos.idcredito=cuotas.idcredito " & _
'      "where creditos.idcredito='" & CLng(IdCredito) & "' " & _
'      "and cuotas.fechacobro is Null " & _
'      "and cuotas.cuotacomodin = 'False' " & _
'      "and cuotas.logic1 = 'False' " & _
'      "and cuotas.fecharefinanciacion is Null " & _
'      "and cuotas.fechavencimiento1<'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
     
sql = "ObtenerSaldoTotalVencido " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
Set rec = cnSQL.OpenResultset(sql)

ImporteTotal = 0

If Not rec.EOF Then
   Do While Not rec.EOF
      'si hay cobros parciales en esa cuota
      SaldoCuota = ObtenerSaldoCuotaOKK(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("fechavencimiento1"), rec.rdoColumns("fechavencimiento2"), rec.rdoColumns("exceptuada"), Fecha)
      
      ImporteTotal = CCur(ImporteTotal) + CCur(SaldoCuota)
      rec.MoveNext
   Loop
End If

ObtenerSaldoTotalVencido = CCur(ImporteTotal)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSaldoTotalVencido"
End Function
Private Function ObtenerSaldoTotalAVencer(ByVal IdCredito As Long, ByVal Fecha As Date) As Currency
'calcula el saldo de un credito teniendo en cta solo las cuotas A VENCER
Dim sql As String
Dim rec As rdoResultset
Dim SaldoCuota As Currency
Dim ImporteTotal As Currency
On Error GoTo merror

ObtenerSaldoTotalAVencer = 0

'trae solo las cuotas a vencer e impagas al dia de hoy
'sql = "select creditos.idcredito,cuotas.numcuota,cuotas.fechavencimiento1,cuotas.fechavencimiento2,cuotas.logic1 as exceptuada " & _
'      "from creditos inner join cuotas on creditos.idcredito=cuotas.idcredito " & _
'      "where creditos.idcredito='" & CLng(IdCredito) & "' " & _
'      "and cuotas.fechacobro is Null " & _
'      "and cuotas.cuotacomodin = 'False' " & _
'      "and cuotas.logic1 = 'False' " & _
'      "and cuotas.fecharefinanciacion is Null " & _
'      "and cuotas.fechavencimiento1>='" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
sql = "ObtenerSaldoTotalAVencer " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
Set rec = cnSQL.OpenResultset(sql)

ImporteTotal = 0

If Not rec.EOF Then
   Do While Not rec.EOF
      'si hay cobros parciales en esa cuota
      SaldoCuota = ObtenerSaldoCuotaOKK(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("fechavencimiento1"), rec.rdoColumns("fechavencimiento2"), rec.rdoColumns("exceptuada"), Fecha)
      ImporteTotal = CCur(ImporteTotal) + CCur(SaldoCuota)
      rec.MoveNext
   Loop
End If

ObtenerSaldoTotalAVencer = CCur(ImporteTotal)

Exit Function
merror:
tratarerrores "Error en funcion ObtenerSaldoTotalAVencer"
End Function
'*****************************************************************
Private Function CapitalVencidoMenor(ByVal IdCredito As Long, ByVal Fecha As Date, ByVal ImporteMenor As Currency) As Currency
'devuelve el importe a restar de capital VENCIDO
'de las cuotas cuyo saldo es inferior al IMPORTE MENOR
Dim sql As String
Dim rec As rdoResultset
Dim SaldoCuota As Currency
Dim Capital As Currency
Dim CapitalCobrado As Currency
Dim SaldoCapital As Currency
On Error GoTo merror

CapitalVencidoMenor = 0

'traigo las vencidas y a vencer (todas las adeudadas)
'sql = "select cuotas.idcredito,cuotas.numcuota," & _
'      "cuotas.importeamortizacion,cuotas.fechavencimiento1,cuotas.fechavencimiento2,cuotas.logic1 as exceptuada " & _
'      "from cuotas " & _
'      "where idcredito='" & CLng(IdCredito) & "' " & _
'      "and fechacobro is Null " & _
'      "and cuotacomodin ='False' " & _
'      "and cuotas.logic1 = 'False' " & _
'      "and fecharefinanciacion is Null " & _
'      "and fechavencimiento1<'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"

sql = "CapitalVencidoMenor " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
Set rec = cnSQL.OpenResultset(sql)

SaldoCapital = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      'obtengo el saldo total de la cuota
      SaldoCuota = ObtenerSaldoCuotaOKK(IdCredito, rec.rdoColumns("numcuota"), rec.rdoColumns("fechavencimiento1"), rec.rdoColumns("fechavencimiento2"), rec.rdoColumns("exceptuada"), Fecha)
      If CCur(SaldoCuota) < CCur(ImporteMenor) Then
         'aca debe sumar el saldo de capital de esa cuota
         'no el capital original
         Capital = rec.rdoColumns("importeamortizacion")
         CapitalCobrado = ObtenerCapitalCobrado(IdCredito, rec.rdoColumns("numcuota"))
         
         SaldoCapital = CCur(Capital) - CCur(CapitalCobrado)
      End If
      rec.MoveNext
   Loop
End If

CapitalVencidoMenor = SaldoCapital

Exit Function
merror:
tratarerrores "Error en funcion CapitalVencidoMenor"
End Function
Private Function CapitalAVencerMenor(ByVal IdCredito As Long, ByVal Fecha As Date, ByVal ImporteMenor As Currency) As Currency
'devuelve el importe a restar de capital
'de las cuotas cuyo saldo es inferior al IMPORTE MENOR
Dim sql As String
Dim rec As rdoResultset
Dim SaldoCuota As Currency
Dim Capital As Currency
Dim CapitalCobrado As Currency
Dim SaldoCapital As Currency
On Error GoTo merror

CapitalAVencerMenor = 0

'traigo las vencidas y a vencer (todas las adeudadas)
'sql = "select cuotas.idcredito,cuotas.numcuota," & _
'      "cuotas.importeamortizacion,cuotas.fechavencimiento1,cuotas.fechavencimiento2,cuotas.logic1 as exceptuada " & _
'      "from cuotas " & _
'      "where idcredito='" & CLng(IdCredito) & "' " & _
'      "and fechacobro is Null " & _
'      "and cuotacomodin = 'False' " & _
'      "and cuotas.logic1 = 'False' " & _
'      "and fecharefinanciacion is Null " & _
'      "and fechavencimiento1>='" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
sql = "CapitalAVencerMenor " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"

Set rec = cnSQL.OpenResultset(sql)

Suma = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      'obtengo el saldo total de la cuota
      SaldoCuota = ObtenerSaldoCuotaOKK(IdCredito, rec.rdoColumns("numcuota"), rec.rdoColumns("fechavencimiento1"), rec.rdoColumns("fechavencimiento2"), rec.rdoColumns("exceptuada"), Fecha)
      If CCur(SaldoCuota) < CCur(ImporteMenor) Then
         Capital = rec.rdoColumns("importeamortizacion")
         CapitalCobrado = ObtenerCapitalCobrado(IdCredito, rec.rdoColumns("numcuota"))
         SaldoCapital = CCur(Capital) - CCur(CapitalCobrado)
      End If
      rec.MoveNext
   Loop
End If

CapitalAVencerMenor = SaldoCapital

Exit Function
merror:
tratarerrores "Error en funcion CapitalAVencerMenor"
End Function
Private Function SaldoVencidoMenor(ByVal IdCredito As Long, ByVal Fecha As Date, ByVal ImporteMenor As Currency) As Currency
'devuelve el importe a restar de saldo total VENCIDO
'de las cuotas cuyo saldo es inferior al IMPORTE MENOR
Dim sql As String
Dim rec As rdoResultset
Dim SaldoCuota As Currency
Dim Suma As Currency
On Error GoTo merror

SaldoVencidoMenor = 0

'traigo las vencidas
'sql = "select cuotas.idcredito,cuotas.numcuota," & _
'      "cuotas.importeamortizacion,cuotas.fechavencimiento1,cuotas.fechavencimiento2,cuotas.logic1 as exceptuada " & _
'      "from cuotas " & _
'      "where idcredito='" & CLng(IdCredito) & "' " & _
'      "and fechacobro is Null " & _
'      "and cuotacomodin = 'False' " & _
'      "and cuotas.logic1 = 'False' " & _
'      "and fecharefinanciacion is Null " & _
'      "and fechavencimiento1<'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
sql = "SaldoVencidoMenor " & CLng(IdCredito) & ",'" & ConvertirFechaSql(CDate(Fecha), "DD/MM/YYYY") & "'"
Set rec = cnSQL.OpenResultset(sql)

Suma = 0
If Not rec.EOF Then
   Do While Not rec.EOF
      'obtengo el saldo total de la cuota
      SaldoCuota = ObtenerSaldoCuotaOKK(IdCredito, rec.rdoColumns("numcuota"), rec.rdoColumns("fechavencimiento1"), rec.rdoColumns("fechavencimiento2"), rec.rdoColumns("exceptuada"), Fecha)
      If CCur(SaldoCuota) < CCur(ImporteMenor) Then
         Suma = CCur(Suma) + CCur(SaldoCuota)
      End If
      rec.MoveNext
   Loop
End If

SaldoVencidoMenor = Suma

Exit Function
merror:
tratarerrores "Error en funcion SaldoVencidoMenor"
End Function
Private Sub CmdExportarDeudores_Click()
Call RefreshTimer
CmdExportarDeudores.Enabled = False
Me.MousePointer = vbHourglass
'Call ExportarDeudores
Call ExportarDeudoresTXT
Me.MousePointer = vbDefault
CmdExportarDeudores.Enabled = True
End Sub
Private Sub ExportarDeudoresTXT()
'EXPORTACION COMPLEJA
Dim MiExcel As Excel.APPLICATION ' Aplicacion Excel
Dim MiLibro As Excel.Workbook ' Un libro de Excel
Dim MiHoja As Excel.Worksheet ' Una hoja de Excel
Dim Filas As Long
Dim FilaTitulos As Long
Dim Archivo As String
Dim Archi As String
Dim rec As rdoResultset
Dim rec2 As rdoResultset
Dim sql As String
Dim SaldoCredito As Currency
Dim SaldoCapital As Currency
Dim CuotasCobradas As Long
Dim Fecha As Date
Dim SaldoCuota As Currency
Dim ImporteParcial As Currency
Dim Mensaje As String
Dim ImporteActualizado As Currency
Dim Dia As String
Dim Mes As String
Dim Ano As String
Dim MoraTotal As Currency
Dim IdCliente As Long
Dim IdCredito As Long
Dim Provincia As String
Dim KOriginal As Currency
Dim TotalCobrado As Currency
Dim CuotasAVencer As Long
Dim CuotasVencidas As Long
Dim KVencido As Currency
Dim KAVencer As Currency
Dim SaldoTotalVencido As Currency
Dim SaldoTotalAVencer As Currency
Dim Menor1 As Currency
Dim Menor2 As Currency
Dim Mcad As String
Dim MiCadFecha As String
Dim ProvCredito As String
Dim MiNew As Currency
Dim MoraVencida  As Currency
Dim cnSQL2 As rdoConnection

On Error GoTo merror

If lvdeudores.ListItems.Count() = 0 Then Exit Sub

Dia = Format(CStr(Day(CDate(DTPicker1.Value))), "00")
Mes = Format(CStr(Month(CDate(DTPicker1.Value))), "00")
Ano = Format(CStr(Year(CDate(DTPicker1.Value))), "0000")

Archi = "Deudores"

Archi = Archi + "-" + Dia + "-" + Mes + "-" + Ano + ".xls"


If Not MsgP("¿Confirma la exportacion de deudores hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & "?") Then Exit Sub

'inicio transaccion

'si no existe la crea
If Not ExisteCarpeta() Then Exit Sub

'verifico si ya existe en la ubicacion
Archivo = Dir("c:\exportacionexcel\" & Archi)

If Trim(Archivo) <> "" Then
  'si ya existe lo borro para que despues no haya errores con la pantallita
  'que despliega el excel
   Kill ("c:\exportacionexcel\" & Archi)
End If

Open "c:\exportacionexcel\" & Archi For Output As #1

Print #1, "Listado de deudores a la fecha:" & CStr(Date)
Print #1, "Prestamo"; Chr$(9); "Cliente"; Chr$(9); "Nro.Cliente"; Chr$(9); "Dni"; Chr$(9); "Provincia"; Chr$(9); "Comercio"; Chr$(9); "K.Original"; Chr$(9); "Total.Cobrado"; Chr$(9); "Total Cuotas"; Chr$(9); "Cuotas.Vencidas"; Chr$(9); "Cuotas.A.Vencer"; Chr$(9); "Primer.Vto.Moroso"; Chr$(9); "Analisis"; Chr$(9); "Max.Dias.Mora"; Chr$(9); "Saldo.Capital.Vencido"; Chr$(9); "Menores"; Chr$(9); "Saldo.Capital.A.Vencer"; Chr$(9); "Saldo.Capital.A.Vencer"; Chr$(9); "Saldo.Total.Capital"; Chr$(9); "Saldo.Total.Vencido"; Chr$(9); "Saldo.Menor"; Chr$(9); "Saldo.Total.A.Vencer"; Chr$(9); "Saldo.Total.A.Vencer"; Chr$(9); "Honorarios"; Chr$(9); "Saldo.Total.A.Refinanciar"; Chr$(9); "Importe Mora"


'obtengo la lista de creditos de pantalla

Set cnSQL2 = enSQL.OpenConnection("", rdDriverNoPrompt, False, cConexion)
sql = "SeleccionarTemplistadeudoresOrden"
Set rec = cnSQL2.OpenResultset(sql)

MoraTotal = 0

If Not rec.EOF Then
   Do While Not rec.EOF

         
      'total cobrado del credito (incluye todas las cuotas en mora y a vencer??)
      'porque puede haber cobros de cuotas en mora y cobros de cuotas a vencer
      'incluye cobros parciales
      TotalCobrado = ObtenerCobrosCredito(rec.rdoColumns("idcredito"))
         
      'cuotas a vencer mas adelante
      CuotasAVencer = ObtenerCuotasPendientes(rec.rdoColumns("idcredito"), DTPicker1.Value)
               
      'fecha de analisis (hoy)
      StrHoy = DTPicker1.Value
      
      'suma de saldo capital vencido
      'de todas las cuotascuotas vencidas en mora
       KVencido = ObtenerSaldoCapVencido(rec.rdoColumns("idcredito"), DTPicker1.Value)
      'ahora le resto el importe MENOR de pantalla
      MiNew = 0
      If MenorOk() Then
         'aca le resto el saldo de capital de las cuotas menores
         MiNew = CapitalVencidoMenor(rec.rdoColumns("idcredito"), DTPicker1.Value, TxtImporteMenor.Text)
         KVencido = KVencido - Mnew
      End If
      'MiHoja.Cells(Filas, 16).Value = KVencido
         
      'esta columna debe sumar los capitales de las cuotas
      'que tienen saldo menor al importe MENOR de pantalla
      Menor1 = 0
      If MenorOk() Then
            Menor1 = MiNew
      End If
      'MiHoja.Cells(Filas, 17).Value = Menor1
      
      'suma saldo capital a vencer de cuotas que vencen mas adelante OK
      KAVencer = ObtenerSaldoCapAVencer(rec.rdoColumns("idcredito"), DTPicker1.Value)
      'resto capital a vencer de las cuotas que tienen saldos menores
      'al IMPORTE MENOR
      If MenorOk() Then
            KAVencer = KAVencer - CapitalAVencerMenor(rec.rdoColumns("idcredito"), DTPicker1.Value, TxtImporteMenor.Text)
      End If
      'capital a vencer
      'MiHoja.Cells(Filas, 18).Value = KAVencer
      
      'saldo total capital
      StrSaldoTotCapital = KVencido + Menor1 + KAVencer
      
      'saldo total capital
      'MiHoja.Cells(Filas, 20).Value = KVencido + Menor1 + KAVencer
      
       'Winik
      'saldo total vencido incluyendo todos los items(columna ultima derecha de pantalla)
      SaldoTotalVencido = ObtenerSaldoTotalVencido(rec.rdoColumns("idcredito"), DTPicker1.Value)
      If MenorOk() Then
         SaldoTotalVencido = SaldoTotalVencido - SaldoVencidoMenor(rec.rdoColumns("idcredito"), DTPicker1.Value, TxtImporteMenor.Text)
      End If
      'MiHoja.Cells(Filas, 21).Value = SaldoTotalVencido

      Menor2 = 0
      If MenorOk() Then
         Menor2 = SaldoVencidoMenor(rec.rdoColumns("idcredito"), DTPicker1.Value, TxtImporteMenor.Text)
      End If
      'MiHoja.Cells(Filas, 22).Value = Menor2
      
      SaldoTotalAVencer = ObtenerSaldoTotalAVencer(rec.rdoColumns("idcredito"), DTPicker1.Value)
      'MiHoja.Cells(Filas, 23).Value = SaldoTotalAVencer
            
      StrSaldoTotRef = SaldoTotalVencido + Menor2 + SaldoTotalAVencer
      MoraVencida = ObtenerMoraVencida(rec.rdoColumns("idcredito"), DTPicker1.Value)
      StrImporteMora = MoraVencida
           
      
      Print #1, rec.rdoColumns("codprestamo") & vbNullString; Chr$(9); rec.rdoColumns("cliente") & vbNullString; Chr$(9); rec.rdoColumns("numlegajo") & vbNullString; Chr$(9); rec.rdoColumns("numdocumento") & vbNullString; Chr$(9); rec.rdoColumns("provcredito") & vbNullString; Chr$(9); rec.rdoColumns("comercio") & vbNullString; Chr$(9); rec.rdoColumns("capital"); Chr$(9); TotalCobrado & vbNullString; Chr$(9); rec.rdoColumns("numcuotas") & vbNullString; Chr$(9); rec.rdoColumns("cuotasenmora") & vbNullString; Chr$(9); CuotasAVencer & vbNullString; Chr$(9); CDate(rec.rdoColumns("vencimiento")) & vbNullString; Chr$(9); StrHoy & vbNullString; Chr$(9); rec.rdoColumns("maxdiasmora") & vbNullString; Chr$(9); KVencido & vbNullString; Chr$(9); Menor1 & vbNullString; Chr$(9); KAVencer & vbNullString; Chr$(9); " " & vbNullString; Chr$(9); StrSaldoTotCapital & vbNullString; Chr$(9); SaldoTotalVencido & vbNullString; Chr$(9); Menor2 & vbNullString; Chr$(9); _
      SaldoTotalAVencer & vbNullString; Chr$(9); " " & vbNullString; Chr$(9); " " & vbNullString; Chr$(9); StrSaldoTotRef & vbNullString; Chr$(9); StrImporteMora & vbNullString
            
      rec.MoveNext
   Loop
   
   Close #1
      
   Mensaje = "Se exporto la lista de deudores a la planilla C:\ExportacionExcel\" & Archi
Else
   Mensaje = "No hay datos para exportar"
End If
   
'fin de transaccion
cnSQL2.Close

MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error Exportando el listado de deudores...verifique que los archivos de Excel esten cerrados"
End Sub
Private Sub ExportarDeudores()
'EXPORTACION COMPLEJA
Dim MiExcel As Excel.APPLICATION ' Aplicacion Excel
Dim MiLibro As Excel.Workbook ' Un libro de Excel
Dim MiHoja As Excel.Worksheet ' Una hoja de Excel
Dim Filas As Long
Dim FilaTitulos As Long
Dim Archivo As String
Dim Archi As String
Dim rec As rdoResultset
Dim rec2 As rdoResultset
Dim sql As String
Dim SaldoCredito As Currency
Dim SaldoCapital As Currency
Dim CuotasCobradas As Long
Dim Fecha As Date
Dim SaldoCuota As Currency
Dim ImporteParcial As Currency
Dim Mensaje As String
Dim ImporteActualizado As Currency
Dim Dia As String
Dim Mes As String
Dim Ano As String
Dim MoraTotal As Currency
Dim IdCliente As Long
Dim IdCredito As Long
Dim Provincia As String
Dim KOriginal As Currency
Dim TotalCobrado As Currency
Dim CuotasAVencer As Long
Dim CuotasVencidas As Long
Dim KVencido As Currency
Dim KAVencer As Currency
Dim SaldoTotalVencido As Currency
Dim SaldoTotalAVencer As Currency
Dim Menor1 As Currency
Dim Menor2 As Currency
Dim Mcad As String
Dim MiCadFecha As String
Dim ProvCredito As String
Dim MiNew As Currency
On Error GoTo merror

If lvdeudores.ListItems.Count() = 0 Then Exit Sub

Dia = Format(CStr(Day(CDate(DTPicker1.Value))), "00")
Mes = Format(CStr(Month(CDate(DTPicker1.Value))), "00")
Ano = Format(CStr(Year(CDate(DTPicker1.Value))), "0000")

Archi = "Deudores"

Archi = Archi + "-" + Dia + "-" + Mes + "-" + Ano + ".xls"


If Not MsgP("¿Confirma la exportacion de deudores hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & "?") Then Exit Sub

'inicio transaccion
cnSQL.BeginTrans

'si no existe la crea
If Not ExisteCarpeta() Then Exit Sub

'verifico si ya existe en la ubicacion
Archivo = Dir("c:\exportacionexcel\" & Archi)

If Trim(Archivo) <> "" Then
  'si ya existe lo borro para que despues no haya errores con la pantallita
  'que despliega el excel
   Kill ("c:\exportacionexcel\" & Archi)
End If

Set MiExcel = New Excel.APPLICATION

Set MiLibro = MiExcel.Workbooks.Add
 
'asigno la primera hoja por defecto
Set MiHoja = MiLibro.Worksheets(1)
  
'la aplicacion no esta a la vista
MiExcel.Visible = False

'titulo principal
MiHoja.Cells(1, 1).Value = "Listado de deudores a la fecha:" & CStr(CDate(DTPicker1.Value))

'establezco el ancho de todas las columnas
MiHoja.Columns.ColumnWidth = 17

FilaTitulos = 2

MiHoja.Cells(FilaTitulos, 1).Value = "Prestamo"
MiHoja.Cells(FilaTitulos, 2).Value = "Cliente"
MiHoja.Cells(FilaTitulos, 3).Value = "Nro.Cliente"
MiHoja.Cells(FilaTitulos, 4).Value = "Dni"
MiHoja.Cells(FilaTitulos, 5).Value = "Provincia"
MiHoja.Cells(FilaTitulos, 6).Value = "Comercio"
MiHoja.Cells(FilaTitulos, 7).Value = "K.Original"
MiHoja.Cells(FilaTitulos, 8).Value = "Total.Cobrado"
MiHoja.Cells(FilaTitulos, 9).Value = "Total Cuotas"
MiHoja.Cells(FilaTitulos, 10).Value = "Cuotas.Vencidas"
MiHoja.Cells(FilaTitulos, 11).Value = "Cuotas.A.Vencer"
MiHoja.Cells(FilaTitulos, 12).Value = "Primer.Vto.Moroso"
MiHoja.Cells(FilaTitulos, 13).Value = "Analisis"
MiHoja.Cells(FilaTitulos, 14).Value = "Dias.Mora"
MiHoja.Cells(FilaTitulos, 15).Value = "Max.Dias.Mora"
MiHoja.Cells(FilaTitulos, 16).Value = "Saldo.Capital.Vencido"
MiHoja.Cells(FilaTitulos, 17).Value = "Menores"
MiHoja.Cells(FilaTitulos, 18).Value = "Saldo.Capital.A.Vencer"
MiHoja.Cells(FilaTitulos, 19).Value = "Saldo.Capital.A.Vencer"
MiHoja.Cells(FilaTitulos, 20).Value = "Saldo.Total.Capital"
MiHoja.Cells(FilaTitulos, 21).Value = "Saldo.Total.Vencido"
MiHoja.Cells(FilaTitulos, 22).Value = "Saldo.Menor"
MiHoja.Cells(FilaTitulos, 23).Value = "Saldo.Total.A.Vencer"
MiHoja.Cells(FilaTitulos, 24).Value = "Saldo.Total.A.Vencer"
MiHoja.Cells(FilaTitulos, 25).Value = "Honorarios"
MiHoja.Cells(FilaTitulos, 26).Value = "Saldo.Total.A.Refinanciar"
MiHoja.Cells(FilaTitulos, 27).Value = "Importe Mora"
'pongo los titulos en negritas
MiHoja.Range("a1:AA2").Font.Bold = True

'obtengo la lista de creditos de pantalla
sql = "select * from templistadeudores order by codprestamo"

Set rec = cnSQL.OpenResultset(sql)

MoraTotal = 0

'comienzo en la fila 3 porque las anteriores son de titulos
Filas = 3
If Not rec.EOF Then
   Do While Not rec.EOF
      'grabo cada usuario en la planilla
      MiHoja.Cells(Filas, 1).Value = rec.rdoColumns("codprestamo") & vbNullString
      MiHoja.Cells(Filas, 2).Value = rec.rdoColumns("cliente") & vbNullString
      MiHoja.Cells(Filas, 3).Value = rec.rdoColumns("numlegajo") & vbNullString
      MiHoja.Cells(Filas, 4).Value = rec.rdoColumns("numdocumento") & vbNullString
      MiHoja.Cells(Filas, 5).Value = rec.rdoColumns("provcredito") & vbNullString
      'comercio que no lo tengo
      MiHoja.Cells(Filas, 6).Value = rec.rdoColumns("comercio") & vbNullString
         
      'capital original
      KOriginal = rec.rdoColumns("capital")
      MiHoja.Cells(Filas, 7).Value = KOriginal
         
      'total cobrado del credito (incluye todas las cuotas en mora y a vencer??)
      'porque puede haber cobros de cuotas en mora y cobros de cuotas a vencer
      'incluye cobros parciales
      TotalCobrado = ObtenerCobrosCredito(rec.rdoColumns("idcredito"))
      MiHoja.Cells(Filas, 8).Value = TotalCobrado
         
      'cantidad de cuotas del credito
      MiHoja.Cells(Filas, 9).Value = rec.rdoColumns("numcuotas")
      
      'cuotas vencidas en mora
      CuotasVencidas = rec.rdoColumns("cuotasenmora")
      MiHoja.Cells(Filas, 10).Value = CuotasVencidas
         
      'cuotas a vencer mas adelante
      CuotasAVencer = ObtenerCuotasPendientes(rec.rdoColumns("idcredito"), DTPicker1.Value)
      MiHoja.Cells(Filas, 11).Value = CuotasAVencer
         
      'vencimiento mas viejo moroso del credito
      'debe descartar las de IMPORTE MENOR
      MiHoja.Cells(Filas, 12).Value = CDate(rec.rdoColumns("vencimiento"))
               
      'fecha de analisis (hoy)
      MiHoja.Cells(Filas, 13).Value = DTPicker1.Value
      
      'total dias de mora del credito
      MiHoja.Cells(Filas, 14).Value = rec.rdoColumns("diasmora")
 
      MiHoja.Cells(Filas, 15).Value = rec.rdoColumns("maxdiasmora")
 
      'suma de saldo capital vencido
      'de todas las cuotascuotas vencidas en mora
       KVencido = ObtenerSaldoCapVencido(rec.rdoColumns("idcredito"), DTPicker1.Value)
      'ahora le resto el importe MENOR de pantalla
      MiNew = 0
      If MenorOk() Then
         'aca le resto el saldo de capital de las cuotas menores
         MiNew = CapitalVencidoMenor(rec.rdoColumns("idcredito"), DTPicker1.Value, TxtImporteMenor.Text)
         KVencido = KVencido - Mnew
      End If
      MiHoja.Cells(Filas, 16).Value = KVencido
         
      'esta columna debe sumar los capitales de las cuotas
      'que tienen saldo menor al importe MENOR de pantalla
      Menor1 = 0
      If MenorOk() Then
            Menor1 = MiNew
      End If
      MiHoja.Cells(Filas, 17).Value = Menor1
      
      'suma saldo capital a vencer de cuotas que vencen mas adelante OK
      KAVencer = ObtenerSaldoCapAVencer(rec.rdoColumns("idcredito"), DTPicker1.Value)
      'resto capital a vencer de las cuotas que tienen saldos menores
      'al IMPORTE MENOR
      If MenorOk() Then
            KAVencer = KAVencer - CapitalAVencerMenor(rec.rdoColumns("idcredito"), DTPicker1.Value, TxtImporteMenor.Text)
      End If
      'capital a vencer
      MiHoja.Cells(Filas, 18).Value = KAVencer
      
      'saldo total capital
      MiHoja.Cells(Filas, 20).Value = KVencido + Menor1 + KAVencer
         
      'saldo total vencido incluyendo todos los items(columna ultima derecha de pantalla)
      SaldoTotalVencido = ObtenerSaldoTotalVencido(rec.rdoColumns("idcredito"), DTPicker1.Value)
      If MenorOk() Then
         SaldoTotalVencido = SaldoTotalVencido - SaldoVencidoMenor(rec.rdoColumns("idcredito"), DTPicker1.Value, TxtImporteMenor.Text)
      End If
      MiHoja.Cells(Filas, 21).Value = SaldoTotalVencido

      Menor2 = 0
      If MenorOk() Then
         Menor2 = SaldoVencidoMenor(rec.rdoColumns("idcredito"), DTPicker1.Value, TxtImporteMenor.Text)
      End If
      MiHoja.Cells(Filas, 22).Value = Menor2
      
      SaldoTotalAVencer = ObtenerSaldoTotalAVencer(rec.rdoColumns("idcredito"), DTPicker1.Value)
      MiHoja.Cells(Filas, 23).Value = SaldoTotalAVencer
            
      MiHoja.Cells(Filas, 26).Value = SaldoTotalVencido + Menor2 + SaldoTotalAVencer
      MiHoja.Cells(Filas, 27).Value = CCur(rec.rdoColumns("importemora"))
                 
      Filas = Filas + 1
      rec.MoveNext
   Loop
    
   'grabo los cambios
   MiLibro.SaveAs ("c:\ExportacionExcel\" & Archi)
   'cierro el libro
   MiLibro.Close
   'salgo de excel
   MiExcel.Quit
   Set MiExcel = Nothing
   Mensaje = "Se exporto la lista de deudores a la planilla C:\ExportacionExcel\" & Archi
Else
   'grabo los cambios en una tabla falsa
   MiLibro.SaveAs ("c:\ExportacionExcel\temporal.xls")
   'cierro el libro
   MiLibro.Close
   'salgo de excel sin grabar los cambios en el archivo adecuado..uso otro
   MiExcel.Quit
   Set MiExcel = Nothing
   
   'borro la planilla temporal
   Kill ("c:\ExportacionExcel\temporal.xls")
  
   Mensaje = "No hay datos para exportar"
End If
   
'fin de transaccion
cnSQL.CommitTrans

MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error Exportando el listado de deudores...verifique que los archivos de Excel esten cerrados"
End Sub
Private Function MenorOk() As Boolean
'valida el importe menor opcional
MenorOk = True

If Trim(TxtImporteMenor.Text) = "" Then
   MenorOk = False
   Exit Function
End If

If Not IsNumeric(TxtImporteMenor.Text) Then
   MenorOk = False
   Exit Function
End If

If CCur(TxtImporteMenor.Text) <= 0 Then
   MenorOk = False
   Exit Function
End If

End Function
Private Sub CmdExportarVeraz_Click()
Call RefreshTimer
CmdExportarVeraz.Enabled = False
Me.MousePointer = vbHourglass
Call ExportarVeraz
Me.MousePointer = vbDefault
CmdExportarVeraz.Enabled = True
End Sub
Private Sub ExportarVeraz()
'exporta a planilla excel
Dim MiExcel As Excel.APPLICATION ' Aplicacion Excel
Dim MiLibro As Excel.Workbook ' Un libro de Excel
Dim MiHoja As Excel.Worksheet ' Una hoja de Excel
Dim Filas As Long
Dim FilaTitulos As Long
Dim Archivo As String
Dim Archi As String
Dim rec As rdoResultset
Dim rec2 As rdoResultset
Dim sql As String
Dim SaldoCredito As Currency
Dim SaldoCapital As Currency
Dim CuotasCobradas As Long
Dim Fecha As Date
Dim SaldoCuota As Currency
Dim ImporteParcial As Currency
Dim Mensaje As String
Dim ImporteActualizado As Currency
Dim Dia As String
Dim Mes As String
Dim Ano As String
Dim MoraTotal As Currency
Dim IdCliente As Long
Dim IdCredito As Long
Dim Provincia As String
Dim KOriginal As Currency
Dim TotalCobrado As Currency
Dim CuotasAVencer As Long
Dim CuotasVencidas As Long
Dim KVencido As Currency
Dim KAVencer As Currency
Dim SaldoTotalVencido As Currency
Dim SaldoTotalAVencer As Currency
Dim Menor1 As Currency
Dim Menor2 As Currency
Dim Mcad As String
Dim MiCadFecha As String
Dim SaldoEnMora As Currency
Dim VtoMorosoMin As Date
Dim I As Long
On Error GoTo merror

If lvdeudores.ListItems.Count() = 0 Then Exit Sub

Dia = Format(CStr(Day(CDate(DTPicker1.Value))), "00")
Mes = Format(CStr(Month(CDate(DTPicker1.Value))), "00")
Ano = Format(CStr(Year(CDate(DTPicker1.Value))), "0000")

Archi = "Veraz"

Archi = Archi + "-" + Dia + "-" + Mes + "-" + Ano + ".xls"


If Not MsgP("¿Confirma la exportacion de deudores hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & "?") Then Exit Sub

'inicio transaccion
cnSQL.BeginTrans

'si no existe la crea
If Not ExisteCarpeta() Then Exit Sub

'verifico si ya existe en la ubicacion
Archivo = Dir("c:\exportacionexcel\" & Archi)

If Trim(Archivo) <> "" Then
  'si ya existe lo borro para que despues no haya errores con la pantallita
  'que despliega el excel
   Kill ("c:\exportacionexcel\" & Archi)
End If

Set MiExcel = New Excel.APPLICATION

Set MiLibro = MiExcel.Workbooks.Add
 
'asigno la primera hoja por defecto
Set MiHoja = MiLibro.Worksheets(1)
  
'la aplicacion no esta a la vista
MiExcel.Visible = False

'titulo principal
MiHoja.Cells(1, 1).Value = "Listado veraz a la fecha:" & CStr(CDate(DTPicker1.Value))

'establezco el ancho de todas las columnas
MiHoja.Columns.ColumnWidth = 17

FilaTitulos = 2

MiHoja.Cells(FilaTitulos, 1).Value = "NOMBRE_COMPLETO"
MiHoja.Cells(FilaTitulos, 2).Value = "DNI_CUIL"
MiHoja.Cells(FilaTitulos, 3).Value = "MOR_SUCURSAL"
MiHoja.Cells(FilaTitulos, 4).Value = "C_2_TIPO_OPERACION"
MiHoja.Cells(FilaTitulos, 5).Value = "C_4_CALIDAD"
MiHoja.Cells(FilaTitulos, 6).Value = "C_5_NRO_OPERACION"
MiHoja.Cells(FilaTitulos, 7).Value = "C_6_MARCA_TARJETA"
MiHoja.Cells(FilaTitulos, 8).Value = "C_7_IMPORTE"
MiHoja.Cells(FilaTitulos, 9).Value = "C_8_ACREEDOR"
MiHoja.Cells(FilaTitulos, 10).Value = "ANT_FECHA"
MiHoja.Cells(FilaTitulos, 11).Value = "SEXO"
MiHoja.Cells(FilaTitulos, 12).Value = "CALLE"
MiHoja.Cells(FilaTitulos, 13).Value = "NUMERO"
MiHoja.Cells(FilaTitulos, 14).Value = "PISO"
MiHoja.Cells(FilaTitulos, 15).Value = "LOCALIDAD"
MiHoja.Cells(FilaTitulos, 16).Value = "PROV"
MiHoja.Cells(FilaTitulos, 17).Value = "CP"
MiHoja.Cells(FilaTitulos, 18).Value = "CDI_PASAPORTE"
MiHoja.Cells(FilaTitulos, 19).Value = "PROV_DOC2"
MiHoja.Cells(FilaTitulos, 20).Value = "NACIONALIDAD"
MiHoja.Cells(FilaTitulos, 21).Value = "TELEFONO"
MiHoja.Cells(FilaTitulos, 22).Value = "FECHA_NAC"
MiHoja.Cells(FilaTitulos, 23).Value = "MOR_CLIENTE"
MiHoja.Cells(FilaTitulos, 24).Value = "MOR_EST_CIVIL"

'pongo los titulos en negritas
MiHoja.Range("a1:y2").Font.Bold = True

'obtengo la lista de creditos de pantalla
sql = "select * from templistadeudores order by codprestamo"

Set rec = cnSQL.OpenResultset(sql)

MoraTotal = 0

'comienzo en la fila 3 porque las anteriores son de titulos
Filas = 3
If Not rec.EOF Then
   Do While Not rec.EOF
      'grabo cada usuario en la planilla
      MiHoja.Cells(Filas, 1).Value = rec.rdoColumns("cliente") & vbNullString
      MiHoja.Cells(Filas, 2).Value = rec.rdoColumns("numdocumento") & vbNullString
      'este va en cero
      MiHoja.Cells(Filas, 3).Value = 0
      MiHoja.Cells(Filas, 4).Value = "PP"
      MiHoja.Cells(Filas, 5).Value = "TIT"
      MiHoja.Cells(Filas, 6).Value = rec.rdoColumns("codprestamo") & vbNullString
      'este va vacio marcatarjeta
      MiHoja.Cells(Filas, 7).Value = ""
      
      'se puede reemplazar
      SaldoEnMora = rec.rdoColumns("saldoenmora")
      MiHoja.Cells(Filas, 8).Value = Format(Int(SaldoEnMora), "000000")
      
      MiHoja.Cells(Filas, 9).Value = "CREDIMACO S.A. (PRESTAMO DE LA CASA - GRUPO BERCOMAT)"
      'el vencimiento moroso va convertido a un formato
      VtoMorosoMin = CDate(rec.rdoColumns("vencimiento"))
      MiHoja.Cells(Filas, 10).Value = CStr(Year(CDate(VtoMorosoMin))) + Format(CStr(Month(CDate(VtoMorosoMin))), "00") + Format(CStr(Day(CDate(VtoMorosoMin))), "00")
      
      'sexo que para los clientes esta en blanco pero los nuevos tendran sexo
      'o se podran cambiar los viejos
      MiHoja.Cells(Filas, 11).Value = rec.rdoColumns("sexo") & vbNullString
         
      'dom
      MiHoja.Cells(Filas, 12).Value = rec.rdoColumns("domicilio") & vbNullString
      
      'loc
      MiHoja.Cells(Filas, 15).Value = rec.rdoColumns("localidad") & vbNullString
      'prov
      MiHoja.Cells(Filas, 16).Value = rec.rdoColumns("provincia") & vbNullString
      'cp
      MiHoja.Cells(Filas, 17).Value = rec.rdoColumns("codigopostal") & vbNullString
                 
      '18 y 19 nada
      MiHoja.Cells(Filas, 20).Value = rec.rdoColumns("nacionalidad") & vbNullString
      MiHoja.Cells(Filas, 21).Value = rec.rdoColumns("telefono") & vbNullString
      'esta fecha es solo texto para que no haya conflicto de alguna fecha mal grabada etc
      'anteriores exportaciones dse colganab con fechas de nacim
      If IsDate(rec.rdoColumns("fechanacimiento")) Then
         MiHoja.Cells(Filas, 22).Value = CDate(rec.rdoColumns("fechanacimiento"))
      Else
         MiHoja.Cells(Filas, 22).Value = rec.rdoColumns("fechanacimiento") & vbNullString
      End If
      MiHoja.Cells(Filas, 23).Value = rec.rdoColumns("numlegajo") & vbNullString
                       
      Filas = Filas + 1
      rec.MoveNext
   Loop
   
   'grabo los cambios
   MiLibro.SaveAs ("c:\ExportacionExcel\" & Archi)
   'cierro el libro
   MiLibro.Close
   'salgo de excel
   MiExcel.Quit
   Set MiExcel = Nothing
   
   Mensaje = "Se exporto la lista de VERAZ a la planilla C:\ExportacionExcel\" & Archi
Else
   'grabo los cambios en una tabla falsa
   MiLibro.SaveAs ("c:\ExportacionExcel\temporal.xls")
   'cierro el libro
   MiLibro.Close
   'salgo de excel sin grabar los cambios en el archivo adecuado..uso otro
   MiExcel.Quit
   Set MiExcel = Nothing
   'borro la planilla temporal
   Kill ("c:\ExportacionExcel\temporal.xls")
  
   Mensaje = "No hay datos para exportar"
End If
   
'fin de transaccion
cnSQL.CommitTrans

MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error Exportando el Listado de Veraz (verifique que el excel este cerrado)"
End Sub
Private Sub CmdImprimirCarta_Click()
Call RefreshTimer
CmdImprimirCarta.Enabled = False
If datosok() Then
   Call ImprimirCartaReclamo
End If
CmdImprimirCarta.Enabled = True
End Sub
Private Sub ImprimirCartaReclamo()
'debe recibir una lista de cuotas
Dim rec As rdoResultset
Dim Mreporte As New ARCartaRec
Dim Fechacompleta As String
Dim IdCredito As Long
Dim CondicionBloqueados As Integer
On Error GoTo merror

If lvdeudores.ListItems.Count = 0 Then Exit Sub
           
IdCredito = lvdeudores.SelectedItem.SubItems(3)

If CheckBloqueados.Value = 1 Then
   CondicionBloqueados = 1
Else
   CondicionBloqueados = 0
End If
Set rec = CargarRecCuotasMora(IdCredito, DTPicker1.Value, CondicionBloqueados)

If Not rec.EOF Then
   With Mreporte
        .RDODataControl1.Resultset = rec
        .rtf.ReplaceField "parrafo1", VG_TEXTOCARTARECLAMO1
        .rtf.ReplaceField "parrafo2", VG_TEXTOCARTARECLAMO2
        .rtf.ReplaceField "fecha", CStr(CDate(DTPicker1.Value))
        .GroupHeader1.DataField = "idcredito"
        .LabelEncabezado.Caption = VG_EMPRESA & vbNullString
        Fechacompleta = FormatearFecha(CDate(DTPicker1.Value))
        .LabelLugarFecha.Caption = VG_CIUDAD + " " + Fechacompleta
        .FieldFecha.Text = DTPicker1.Value
        
        .Show (vbModal)
    End With
Else
    MsgI "No hay datos para imprimir la carta reclamo"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo la carta reclamo"
End Sub
Private Sub CmdLibreDeuda_Click()
'solo imprime el libre deuda por cliente
Call RefreshTimer
If ComboOpciones.Text = "Por cliente" Then
   CmdLibreDeuda.Enabled = False
   
   If VG_EMITELIBREDEUDA Then
      If ComboOpciones.Text = "Por cliente" Then
         'si hay un cliente
         If Trim(TxtCliente.Text) = "" Then
            MsgE "Debe seleccionar un cliente"
         Else
            If Not TieneCreditoBloqueado(IdCliente2) Then
               Call ImprimirLibreDeuda2
            Else
               MsgE "El cliente tiene creditos bloqueados"
            End If
         End If
      End If
   Else
      MsgE "El usuario actual no tiene permiso para emitir libre deuda"
   End If

   CmdLibreDeuda.Enabled = True
End If
End Sub
Private Sub ImprimirLibreDeuda2()
'imprime el libre deuda
Dim Mreporte As New ARLibreDeudaCreditos
Dim Archivo As String
On Error GoTo merror

'si hay datos no imprime el libre deuda
If lvdeudores.ListItems.Count > 0 Then Exit Sub

'si imprimo datos de la empresa
Mreporte.LabelEmpresa = VG_EMPRESA & vbNullString
Mreporte.RtfCuotas.Visible = True
Mreporte.RtfCuotas.ReplaceField "parrafo1", VG_TEXTOLIBREDEUDA1
Mreporte.RtfCuotas.ReplaceField "parrafo2", VG_TEXTOLIBREDEUDA2
Mreporte.RtfCreditos.Visible = False
Mreporte.LabelSubtitulo = "Libre deuda de cuotas hasta la fecha " & CStr(CDate(DTPicker1.Value))
Mreporte.LabelLegajo = Numlegajocliente2 & vbNullString
Mreporte.LabelCliente = UCase(TxtCliente.Text) & vbNullString
Mreporte.LabelDni = NumDni
Mreporte.LabelLugarFecha = VG_CIUDAD & " " & FormatearFecha(CDate(DTPicker1.Value))
Mreporte.Show vbModal
  
Exit Sub
merror:
tratarerrores "Error imprimiendo libre deuda"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If ComboOpciones.Text = "Por cliente" Then
   If Trim(TxtCliente.Text) = "" Then
      datosok = False
      MsgE "Debe seleccionar el cliente"
      TxtCliente.SetFocus
      Exit Function
   End If
End If

If ComboOpciones.Text = "Por cuotas adeudadas" Then
   If Trim(TxtCuotas.Text) = "" Then
      datosok = False
      MsgE "Debe ingresar la cantidad de cuotas adeudadas"
      TxtCuotas.SetFocus
      Exit Function
   End If
   If Not IsNumeric(TxtCuotas.Text) Then
      datosok = False
      MsgE "La cantidad de cuotas debe ser numerica"
      TxtCuotas.SetFocus
      Exit Function
   End If
   
   If CLng(TxtCuotas.Text) <= 0 Then
      datosok = False
      MsgE "La cantidad de cuotas debe ser mayor a cero"
      TxtCuotas.SetFocus
      Exit Function
   End If
End If

If ComboOpciones.Text = "Por rango de Maximos" Then
   If Trim(TxtDiasMora.Text) = "" Then
      datosok = False
      MsgE "Debe ingresar el rango de maximo de dias de mora"
      TxtDiasMora.SetFocus
      Exit Function
   End If
   If Not IsNumeric(TxtDiasMora.Text) Then
      datosok = False
      MsgE "El rango de maximos de dias de mora debe ser numerico"
      TxtDiasMora.SetFocus
      Exit Function
   End If
   
   If CLng(TxtDiasMora.Text) <= 0 Then
      datosok = False
      MsgE "El rango de maximos de dias de mora debe ser mayor a cero"
      TxtDiasMora.SetFocus
      Exit Function
   End If
   
   If Trim(TxtDiasMora2.Text) = "" Then
      datosok = False
      MsgE "Debe ingresar el rango de maximo de dias de mora"
      TxtDiasMora2.SetFocus
      Exit Function
   End If
   If Not IsNumeric(TxtDiasMora2.Text) Then
      datosok = False
      MsgE "El rango de maximos de dias de mora debe ser numerico"
      TxtDiasMora2.SetFocus
      Exit Function
   End If
   
   If CLng(TxtDiasMora2.Text) <= 0 Then
      datosok = False
      MsgE "El rango de maximos de dias de mora debe ser mayor a cero"
      TxtDiasMora2.SetFocus
      Exit Function
   End If
End If

If ComboOpciones.Text = "Por importe adeudado" Then
   If Trim(TxtImporteAdeudado.Text) = "" Then
      datosok = False
      MsgE "Debe ingresar el importe adeudado"
      TxtImporteAdeudado.SetFocus
      Exit Function
   End If
   If Not IsNumeric(TxtImporteAdeudado.Text) Then
      datosok = False
      MsgE "El importe debe ser numerico"
      TxtImporteAdeudado.SetFocus
      Exit Function
   End If
   If CCur(TxtImporteAdeudado.Text) <= 0 Then
      datosok = False
      MsgE "El importe debe ser mayor a cero"
      TxtImporteAdeudado.SetFocus
      Exit Function
   End If
End If

If ComboOpciones.Text = "Por Maximo dias de mora" Then
   If Trim(TxtDiasMoraNew.Text) = "" Then
      datosok = False
      MsgE "Debe ingresar el maximo de dias de mora"
      TxtDiasMoraNew.SetFocus
      Exit Function
   End If
   If Not IsNumeric(TxtDiasMoraNew.Text) Then
      datosok = False
      MsgE "El maximo de dias de mora debe ser numerico"
      TxtDiasMoraNew.SetFocus
      Exit Function
   End If
   If CLng(TxtDiasMoraNew.Text) <= 0 Then
      datosok = False
      MsgE "El maximo de dias de mora debe ser mayor a cero"
      TxtDiasMoraNew.SetFocus
      Exit Function
   End If
   
   'este es opcional
   If Not IsNumeric(TxtImporteMenor.Text) Then
      TxtImporteMenor.Text = ""
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-ConsultarDeudores"
End Function
Private Sub TxtCuotas_Change()
lvdeudores.ListItems.Clear
lvcuotas.ListItems.Clear

TxtContador1.Text = 0
TxtImporteTotal.Text = 0

Call SetearEntorno

End Sub
Private Sub SetearEntorno()
On Error GoTo merror

If lvdeudores.ListItems.Count = 0 Then
   CmdImprimirCarta.Enabled = False
   CmdExportarDeudores.Enabled = False
   CmdExportarVeraz.Enabled = False
   TxtImporteTotal.Text = 0
   TxtContador1.Text = 0
   CmdLibreDeuda.Enabled = True
Else
   'si el tipo de usuario puede emitir la carta reclamo
   If VG_EMITECARTARECLAMO Then
      CmdImprimirCarta.Enabled = True
   End If
   
   CmdLibreDeuda.Enabled = False
   
   If VG_EXPORTA Then
      CmdExportarDeudores.Enabled = True
      CmdExportarVeraz.Enabled = True
   End If
   lvdeudores.SetFocus
End If

Exit Sub
merror:
tratarerrores "Error seteando el entorno-ConsultarDeudores"
End Sub
Private Sub TxtImporteAdeudado_Change()
lvdeudores.ListItems.Clear
lvcuotas.ListItems.Clear

TxtContador1.Text = 0
TxtImporteTotal.Text = 0

Call SetearEntorno

End Sub
Private Function TieneCreditoBloqueado(ByVal IdCliente As Long) As Boolean
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

TieneCreditoBloqueado = False

sql = "select idcliente from creditos where creditos.fechabloqueo is not Null " & _
      "and idcliente='" & CLng(IdCliente) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcliente")) Then
      TieneCreditoBloqueado = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion TieneCreditoBloqueado"
End Function
