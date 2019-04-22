VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H00FF0000&
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9285
   HelpContextID   =   2
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrLogout 
      Interval        =   60000
      Left            =   480
      Top             =   1320
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PictureToolBar 
      Align           =   1  'Align Top
      Height          =   560
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9225
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      Begin VB.TextBox TxtUsuarioActual 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9480
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
      Begin VB.Image Image13 
         Enabled         =   0   'False
         Height          =   360
         Left            =   6720
         Picture         =   "MDIPrincipal.frx":030A
         Stretch         =   -1  'True
         ToolTipText     =   "Importa de PagoFacil y RapiPago simultaneamente (Desde Excel)"
         Top             =   120
         Width           =   480
      End
      Begin VB.Image Image12 
         Enabled         =   0   'False
         Height          =   495
         Left            =   5760
         Picture         =   "MDIPrincipal.frx":074C
         Stretch         =   -1  'True
         ToolTipText     =   "Importar cobros de RapiPago"
         Top             =   0
         Width           =   825
      End
      Begin VB.Image Image11 
         Enabled         =   0   'False
         Height          =   480
         Left            =   7440
         Picture         =   "MDIPrincipal.frx":17F6
         ToolTipText     =   "Consultar cobros registrados"
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image8 
         Enabled         =   0   'False
         Height          =   480
         Left            =   3240
         Picture         =   "MDIPrincipal.frx":1B00
         ToolTipText     =   "Cobrar cuotas multiples"
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image10 
         Enabled         =   0   'False
         Height          =   480
         Left            =   4440
         Picture         =   "MDIPrincipal.frx":1E0A
         ToolTipText     =   "Administrar cobradores"
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image7 
         Enabled         =   0   'False
         Height          =   480
         Left            =   5160
         Picture         =   "MDIPrincipal.frx":2114
         ToolTipText     =   "Refinanciar creditos"
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image6 
         Enabled         =   0   'False
         Height          =   480
         Left            =   3840
         Picture         =   "MDIPrincipal.frx":2556
         ToolTipText     =   "Consultar deudores"
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image5 
         Enabled         =   0   'False
         Height          =   480
         Left            =   2520
         Picture         =   "MDIPrincipal.frx":2998
         ToolTipText     =   "Cobrar cuotas"
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image4 
         Enabled         =   0   'False
         Height          =   480
         Left            =   1920
         Picture         =   "MDIPrincipal.frx":2CA2
         ToolTipText     =   "Consultar creditos"
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image3 
         Enabled         =   0   'False
         Height          =   480
         Left            =   1320
         Picture         =   "MDIPrincipal.frx":9A64
         ToolTipText     =   "Registrar creditos"
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image2 
         Enabled         =   0   'False
         Height          =   480
         Left            =   720
         Picture         =   "MDIPrincipal.frx":9EA6
         ToolTipText     =   "Registrar clientes"
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   480
         Left            =   120
         Picture         =   "MDIPrincipal.frx":A1B0
         ToolTipText     =   "Personalizar opciones del sistema"
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Menu files 
      Caption         =   "Archivo"
      Begin VB.Menu Opciones 
         Caption         =   "Configurar opciones del sistema"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu Backup 
         Caption         =   "Mantenimiento de la base de datos"
         Enabled         =   0   'False
      End
      Begin VB.Menu mtiposusuarios 
         Caption         =   "Tipos de usuarios"
         Enabled         =   0   'False
      End
      Begin VB.Menu AbmUsuarios 
         Caption         =   "Usuarios"
         Enabled         =   0   'False
      End
      Begin VB.Menu mimportar 
         Caption         =   "Importar"
         Visible         =   0   'False
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir del sistema"
      End
   End
   Begin VB.Menu actualiz 
      Caption         =   "Actualizaciones"
      Begin VB.Menu AbmClientes 
         Caption         =   "Clientes"
         Enabled         =   0   'False
      End
      Begin VB.Menu mplanes 
         Caption         =   "Planes de creditos"
         Enabled         =   0   'False
      End
      Begin VB.Menu mcomercios 
         Caption         =   "Comercios"
         Enabled         =   0   'False
      End
      Begin VB.Menu mabmvendedores 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu mcobradores 
         Caption         =   "Cobradores"
         Enabled         =   0   'False
      End
      Begin VB.Menu Abmbancos 
         Caption         =   "Bancos"
      End
      Begin VB.Menu AbmLocalidades 
         Caption         =   "Ciudades"
         Enabled         =   0   'False
      End
      Begin VB.Menu AbmProvincias 
         Caption         =   "Provincias"
         Enabled         =   0   'False
      End
      Begin VB.Menu AbmFeriados 
         Caption         =   "Dias no habiles para vencimientos"
         Enabled         =   0   'False
      End
      Begin VB.Menu AbmEstudios 
         Caption         =   "Estudios"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu moperaciones 
      Caption         =   "Operaciones"
      Begin VB.Menu RegistrarCreditos 
         Caption         =   "Registrar nuevo credito"
         Enabled         =   0   'False
      End
      Begin VB.Menu ConsultarCreditos 
         Caption         =   "Consultar creditos / Imprimir cuotas"
         Enabled         =   0   'False
      End
      Begin VB.Menu CobrarCreditos 
         Caption         =   "Cobrar cuotas por factura"
         Enabled         =   0   'False
      End
      Begin VB.Menu mmasiva 
         Caption         =   "Cobrar cuotas multiples"
         Enabled         =   0   'False
      End
      Begin VB.Menu ConsultarDeudores 
         Caption         =   "Consultar deudores (carta reclamo)"
         Enabled         =   0   'False
      End
      Begin VB.Menu RefinanciarCreditos 
         Caption         =   "Refinanciar creditos"
         Enabled         =   0   'False
      End
      Begin VB.Menu madmcobradores 
         Caption         =   "Administrar cobradores"
         Enabled         =   0   'False
      End
      Begin VB.Menu mingresos 
         Caption         =   "Consultar cobros realizados"
         Enabled         =   0   'False
      End
      Begin VB.Menu mimportarpagofacil 
         Caption         =   "Importar PagoFacil"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mimportarrapipago 
         Caption         =   "Importar RapiPago"
         Enabled         =   0   'False
      End
      Begin VB.Menu mimportarambos 
         Caption         =   "Importar desde RapiPago y PagoFacil"
         Enabled         =   0   'False
      End
      Begin VB.Menu mExportarVeraz 
         Caption         =   "Exportar Archivos Para Veraz"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mayuda 
      Caption         =   "Ayuda (F1)"
      Begin VB.Menu AcercaDe 
         Caption         =   "Acerca de"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHistorial 
      Caption         =   "Historial"
      Visible         =   0   'False
      Begin VB.Menu mnuHistorialPagos 
         Caption         =   "Historial de Pagos"
      End
      Begin VB.Menu mnuHistorialCuotas 
         Caption         =   "Historial de Cuotas"
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***PANTALLA PRINCIPAL DEL SISTEMA


Private Sub MDIForm_Activate()
    If cAmbiente = "TEST" Or cAmbiente = "TESR" Then
        Me.BackColor = 100
    End If
'    If cAmbiente = "TESR" Then
'        Me.Picture = LoadPicture("c:\sergio\test\jpg\river.jpg")
'    End If
End Sub

Private Sub MDIForm_Load()
On Error GoTo merror

Me.Caption = App.ProductName

Me.StatusBar1.Panels(1).Text = "Nombre de usuario: " & UCase(VG_USUARIOLOGIN)

'cargo las variables globales con los permisos del tipo de usuario
'correspondiente al usuario que esta ingresando
Call CargarPermisosTipoUsuario(VG_IDTIPOUSUARIOLOGIN)
   
'si es administrador habilito todo incluyendo el abm de usuarios y tipos de usuarios
If VG_IDTIPOUSUARIOLOGIN = 1 Then
   MDIPrincipal.AbmUsuarios.Enabled = True
   MDIPrincipal.mtiposusuarios.Enabled = True
   MDIPrincipal.AbmFeriados.Enabled = True
   MDIPrincipal.AbmLocalidades.Enabled = True
   MDIPrincipal.AbmEstudios.Enabled = True
   MDIPrincipal.AbmProvincias.Enabled = True
   MDIPrincipal.mcobradores.Enabled = True
   MDIPrincipal.RegistrarCreditos.Enabled = True
   MDIPrincipal.Image3.Enabled = True
   MDIPrincipal.RefinanciarCreditos.Enabled = True
   MDIPrincipal.Image7.Enabled = True
   MDIPrincipal.CobrarCreditos.Enabled = True
   MDIPrincipal.Image5.Enabled = True
   MDIPrincipal.ConsultarCreditos.Enabled = True
   MDIPrincipal.Image4.Enabled = True
   MDIPrincipal.AbmClientes.Enabled = True
   MDIPrincipal.Image2.Enabled = True
   MDIPrincipal.madmcobradores.Enabled = True
   MDIPrincipal.Image10.Enabled = True
   MDIPrincipal.ConsultarDeudores.Enabled = True
   MDIPrincipal.Image6.Enabled = True
   MDIPrincipal.Backup.Enabled = False
   MDIPrincipal.Opciones.Enabled = True
   MDIPrincipal.Image1.Enabled = True
   MDIPrincipal.Image8.Enabled = True
   MDIPrincipal.mmasiva.Enabled = True
   MDIPrincipal.mingresos.Enabled = True
   MDIPrincipal.Image11.Enabled = True
   MDIPrincipal.mplanes.Enabled = True
   MDIPrincipal.mcomercios.Enabled = True
   MDIPrincipal.mimportarrapipago.Enabled = True
   MDIPrincipal.mimportarpagofacil.Enabled = True
   MDIPrincipal.Image12.Enabled = True
   MDIPrincipal.mimportarambos.Enabled = True
   MDIPrincipal.mExportarVeraz.Enabled = True
   MDIPrincipal.Image13.Enabled = True
   MDIPrincipal.AcercaDe.Enabled = True
Else
   'si no es administrador voy habilitando de acuerdo al tipo
   'de usuario (no habilito las pantallas de usuarios ni tipo de usuarios)
   MDIPrincipal.AbmUsuarios.Enabled = False
   MDIPrincipal.mtiposusuarios.Enabled = False
   
   'si el tipo de usuario actualiza pantallas varias
   If VG_ACTUALIZA Then
      MDIPrincipal.AbmFeriados.Enabled = True
      MDIPrincipal.AbmLocalidades.Enabled = True
      MDIPrincipal.AbmProvincias.Enabled = True
      MDIPrincipal.AbmEstudios.Enabled = True
      MDIPrincipal.mcomercios.Enabled = True
   End If
     
   'si administra planes
   If VG_ADMPLANES Then
      MDIPrincipal.mplanes.Enabled = True
   End If
      
   'si el tipo de usuario registra creditos
   If VG_REGISTRA Then
      MDIPrincipal.RegistrarCreditos.Enabled = True
      MDIPrincipal.Image3.Enabled = True
   End If
      
   'si el tipo de usuario refinancia creditos
   If VG_REFINANCIA Then
      MDIPrincipal.RefinanciarCreditos.Enabled = True
      MDIPrincipal.Image7.Enabled = True
   End If
      
   'si el tipo de usuario cobra
   If VG_COBRA Then
      MDIPrincipal.CobrarCreditos.Enabled = True
      MDIPrincipal.mmasiva.Enabled = True
      MDIPrincipal.Image5.Enabled = True
      MDIPrincipal.Image8.Enabled = True
   End If
   'la anulacion de cobros se maneja dentro de la pantalla cobros
      
   'si consulta creditos
   If VG_CONSULTA Then
      MDIPrincipal.ConsultarCreditos.Enabled = True
      MDIPrincipal.Image4.Enabled = True
   End If
   'si borra, bloquea, finaliza, comodin se maneja dentro de pantalla de consulta
   
   'si administra clientes
   If VG_ADMCLIENTES Then
      MDIPrincipal.AbmClientes.Enabled = True
      MDIPrincipal.Image2.Enabled = True
   End If
      
   'si administra cobradores
   If VG_ADMCOBRADORES Then
      MDIPrincipal.madmcobradores.Enabled = True
      MDIPrincipal.Image10.Enabled = True
      MDIPrincipal.mcobradores.Enabled = True
   End If
      
   'si consulta deudores
   If VG_CONSULTADEUDORES Then
      MDIPrincipal.ConsultarDeudores.Enabled = True
      MDIPrincipal.Image6.Enabled = True
   End If
   'la carta reclamo y libre deuda se maneja en la pantalla de deudores
   
   If VG_EFECTUABACKUP Then
      MDIPrincipal.Backup.Enabled = False
   End If
      
   'si actualiza opciones
   If VG_ACTUALIZAOPCIONES Then
      MDIPrincipal.Opciones.Enabled = True
      MDIPrincipal.Image1.Enabled = True
   End If
      
   'si importa puede usar todos los importadores del sistema
   If VG_IMPORTA Then
      MDIPrincipal.mimportarrapipago.Enabled = True
      MDIPrincipal.mimportarpagofacil.Enabled = True
      MDIPrincipal.Image12.Enabled = True
      MDIPrincipal.mimportarambos.Enabled = True
      MDIPrincipal.mExportarVeraz.Enabled = True
      MDIPrincipal.Image13.Enabled = True
   End If
   
   If VG_CONSULTAINGRESOS Then
      MDIPrincipal.mingresos.Enabled = True
      MDIPrincipal.Image11.Enabled = True
   End If
   
   If VG_ACERCADE Then
    MDIPrincipal.AcercaDe.Enabled = True
   End If
      
   
   'los usuarios comunes no pueden manejar pantallas de usuarios
   MDIPrincipal.AbmUsuarios.Enabled = False
   MDIPrincipal.mtiposusuarios.Enabled = False
End If
minutosLogout = 0
Exit Sub
merror:
tratarerrores "Error cargando la pantalla principal de soft de creditos"
End Sub
Private Sub CargarPermisosTipoUsuario(ByVal Idtipousuario As Long)
'carga los permisos del tipo de usuario actual
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

sql = "select * from tipousuario " & _
      "where idtipousuario=" & CLng(Idtipousuario)
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   VG_ACTUALIZA = rec.rdoColumns("actualiza")
   VG_REGISTRA = rec.rdoColumns("registra")
   VG_REFINANCIA = rec.rdoColumns("refinancia")
   VG_COBRA = rec.rdoColumns("cobra")
   VG_ANULA = rec.rdoColumns("anula")
   VG_CONSULTA = rec.rdoColumns("consulta")
   VG_ADMCREDITOS = rec.rdoColumns("admcreditos")
   VG_ADMCOBRADORES = rec.rdoColumns("admcobradores")
   VG_IMPRIMECUOTAS = rec.rdoColumns("imprimecuotas")
   VG_EFECTUABACKUP = rec.rdoColumns("efectuabackup")
   VG_EMITECARTARECLAMO = rec.rdoColumns("emitecartareclamo")
   VG_EMITELIBREDEUDA = rec.rdoColumns("emitelibredeuda")
   VG_ADMCLIENTES = rec.rdoColumns("admclientes")
   VG_ACTUALIZAOPCIONES = rec.rdoColumns("actualizaopciones")
   VG_CONSULTADEUDORES = rec.rdoColumns("consultadeudores")
   VG_EXPORTA = rec.rdoColumns("adminversores")
   VG_IMPORTA = rec.rdoColumns("admcomodines")
   VG_CONSULTAINGRESOS = rec.rdoColumns("consultaingresos")
   VG_ACERCADE = rec.rdoColumns("acercade")
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento CargarPermisosTipoUsuario"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'al cerrar la pantalla principal termina de cerrar correctamente el sistema
 Call CLOSE_MODULE
 End
End Sub
Private Sub abmclientes_Click()
Call CenterForm(FrmClientesAbm)
FrmClientesAbm.Show
End Sub
Private Sub mcobradores_Click()
Call CenterForm(FrmCobradoresAbm)
FrmCobradoresAbm.Show
End Sub
Private Sub madmcobradores_Click()
Call CenterForm(FrmAdministrarCobradores)
FrmAdministrarCobradores.Show
End Sub
Private Sub abmusuarios_Click()
Call CenterForm(FrmUsuariosAbm)
FrmUsuariosAbm.Show
End Sub
Private Sub backup_Click()
Call CenterForm(FrmBackup)
FrmBackup.Show
End Sub
Private Sub cobrarcreditos_Click()
Call CenterForm(FrmCobrarCreditos)
FrmCobrarCreditos.Show
End Sub
Private Sub consultardeudores_Click()
Call CenterForm(FrmCreditosConsultarDeudores)
FrmCreditosConsultarDeudores.Show
End Sub
Private Sub acercade_Click()
Call CenterForm(FrmAcercaDe)
FrmAcercaDe.Show
End Sub
Private Sub abmlocalidades_Click()
Call CenterForm(FrmLocalidadesAbm)
FrmLocalidadesAbm.Show
End Sub
Private Sub abmEstudios_Click()
Call CenterForm(FrmEstudiosAbm)
FrmEstudiosAbm.Show
End Sub
Private Sub abmferiados_Click()
Call CenterForm(FrmFeriadosAbm)
FrmFeriadosAbm.Show
End Sub
Private Sub consultarcreditos_Click()
Call CenterForm(FrmConsultarCreditos)
FrmConsultarCreditos.Show
End Sub

Private Sub mExportarVeraz_Click()
Call CenterForm(frmExportarVeraz)
frmExportarVeraz.Show
End Sub

Private Sub mimportarambos_Click()
Call CenterForm(FrmImportarAmbos)
FrmImportarAmbos.Show
End Sub

Private Sub mmasiva_Click()
Call CenterForm(FrmCobrosMasivos)
FrmCobrosMasivos.Show
End Sub

Private Sub mnuHistorialCuotas_Click()
'imprime el historial de cobros del crdito selecionado
Call RefreshTimer

If FrmConsultarCreditos.lvCreditos.ListItems.Count = 0 Then Exit Sub

Call FrmConsultarCreditos.ImprimirHistorialCredito(FrmConsultarCreditos.lvCreditos.SelectedItem.SubItems(1))
End Sub



Private Sub mnuHistorialPagos_Click()
'imprime el historial de cobros del crdito selecionado
Call RefreshTimer

If FrmConsultarCreditos.lvCreditos.ListItems.Count = 0 Then Exit Sub

Call FrmConsultarCreditos.ImprimirHistorialPagosCliente(FrmConsultarCreditos.lvCreditos.SelectedItem.SubItems(1))

End Sub

Private Sub mplanes_Click()
Call CenterForm(FrmPlanesAbm)
FrmPlanesAbm.Show
End Sub
Private Sub mimportarrapipago_Click()
Call CenterForm(FrmImportarRapiPago)
FrmImportarRapiPago.Show
End Sub
Private Sub mtiposusuarios_Click()
Call CenterForm(FrmTiposUsuariosAbm)
FrmTiposUsuariosAbm.Show
End Sub
Private Sub opciones_Click()
Call CenterForm(FrmOpciones)
FrmOpciones.Show
End Sub
Private Sub abmprovincias_Click()
Call CenterForm(FrmProvinciasAbm)
FrmProvinciasAbm.Show
End Sub
Private Sub refinanciarcreditos_Click()
Call CenterForm(FrmReFinanciarCreditos)
FrmReFinanciarCreditos.Show
End Sub
Private Sub registrarcreditos_Click()
Call CenterForm(FrmRegistrarCreditos)
FrmRegistrarCreditos.Show
End Sub
Private Sub salir_Click()
Unload Me
End Sub
Private Sub Image1_Click()
Call CenterForm(FrmOpciones)
FrmOpciones.Show
End Sub
Private Sub Image2_Click()
Call CenterForm(FrmClientesAbm)
FrmClientesAbm.Show
End Sub
Private Sub Image3_Click()
Call CenterForm(FrmRegistrarCreditos)
FrmRegistrarCreditos.Show
End Sub
Private Sub Image4_Click()
Call CenterForm(FrmConsultarCreditos)
FrmConsultarCreditos.Show
End Sub
Private Sub Image5_Click()
Call CenterForm(FrmCobrarCreditos)
FrmCobrarCreditos.Show
End Sub
Private Sub Image6_Click()
Call CenterForm(FrmCreditosConsultarDeudores)
FrmCreditosConsultarDeudores.Show
End Sub
Private Sub Image7_Click()
Call CenterForm(FrmReFinanciarCreditos)
FrmReFinanciarCreditos.Show
End Sub
Private Sub Image10_Click()
Call CenterForm(FrmAdministrarCobradores)
FrmAdministrarCobradores.Show
End Sub
Private Sub Image8_Click()
Call CenterForm(FrmCobrosMasivos)
FrmCobrosMasivos.Show
End Sub
Private Sub Image11_Click()
Call CenterForm(FrmConsultarIngresos)
FrmConsultarIngresos.Show
End Sub
Private Sub mingresos_Click()
Call CenterForm(FrmConsultarIngresos)
FrmConsultarIngresos.Show
End Sub
Private Sub mcomercios_Click()
Call CenterForm(FrmComerciosAbm)
FrmComerciosAbm.Show
End Sub
Private Sub Image12_Click()
Call CenterForm(FrmImportarRapiPago)
FrmImportarRapiPago.Show
End Sub
Private Sub Image13_Click()
Call CenterForm(FrmImportarAmbos)
FrmImportarAmbos.Show
End Sub
Private Sub mabmvendedores_Click()
Call CenterForm(FrmVendedoresAbm)
FrmVendedoresAbm.Show
End Sub
Private Sub Abmbancos_Click()
Call CenterForm(FrmBancosAbm)
FrmBancosAbm.Show
End Sub

Private Sub Timer1_Timer()
    MsgBox "hola"
End Sub

Private Sub tmrLogout_Timer()
    minutosLogout = minutosLogout + 1
    If minutosLogout = VG_TIEMPO_LOGOUT Then
        End
    End If
End Sub
