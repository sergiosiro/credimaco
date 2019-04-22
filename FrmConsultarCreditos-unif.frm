VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmConsultarCreditos 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultar Creditos "
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   HelpContextID   =   17
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUnificarCupones 
      Caption         =   "Unificar Cupones"
      Height          =   300
      Left            =   6240
      TabIndex        =   102
      ToolTipText     =   "Imprime las cuotas con codigo de barras"
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton CmdSellos 
      Caption         =   "Sellos"
      Height          =   300
      Left            =   9240
      TabIndex        =   101
      ToolTipText     =   "Imprime Sellos"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame Frame8 
      Height          =   1095
      Left            =   8400
      TabIndex        =   92
      Top             =   6960
      Width           =   2415
      Begin VB.CommandButton cmdcerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   360
         TabIndex        =   93
         ToolTipText     =   "Cierra la pantalla"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones de busqueda:"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   10695
      Begin VB.ComboBox ComboOpciones 
         Height          =   315
         ItemData        =   "FrmConsultarCreditos.frx":0000
         Left            =   120
         List            =   "FrmConsultarCreditos.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Selecciona que tipo de creditos consultar"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame FrameCliente 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   4200
         TabIndex        =   40
         Top             =   120
         Width           =   6255
         Begin VB.CommandButton CmdBuscarCliente 
            Height          =   350
            Left            =   5640
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Permite seleccionar al cliente de una lista"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox TxtCliente 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   20
            ToolTipText     =   "Cliente titular de los creditos a consultar"
            Top             =   240
            Width           =   5415
         End
         Begin VB.Label Label24 
            Caption         =   "Cliente:"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame FrameConsultaMasiva 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   2280
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton CmdBuscar 
            Height          =   255
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar creditos en el periodo indicado"
            Top             =   120
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   255
            Left            =   600
            TabIndex        =   3
            ToolTipText     =   "Fecha final de la consulta"
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   24641537
            CurrentDate     =   39049
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Left            =   600
            TabIndex        =   2
            ToolTipText     =   "Fecha inicial de la consulta"
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   24641537
            CurrentDate     =   39049
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Lista de creditos:"
      ForeColor       =   &H00FF0000&
      Height          =   2820
      Left            =   120
      TabIndex        =   33
      Top             =   560
      Width           =   10695
      Begin VB.CommandButton CmdHistorial 
         Caption         =   "Historial credito"
         Enabled         =   0   'False
         Height          =   255
         Left            =   7440
         TabIndex        =   97
         ToolTipText     =   "Imprime el historial de cobros del credito seleccionado"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton CmdAnularRefin 
         Caption         =   "Anular Refinanc."
         Enabled         =   0   'False
         Height          =   255
         Left            =   7440
         TabIndex        =   96
         ToolTipText     =   "Anula la refinanciacion seleccionada en la lista"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton CmdFinalizar 
         Caption         =   "&Finalizar credito"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7440
         TabIndex        =   91
         ToolTipText     =   "Finaliza/Restaura el credito seleccionado en la lista superior"
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton CmdBloquear 
         Caption         =   "Bloqu&ear credito"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7440
         TabIndex        =   90
         ToolTipText     =   "Bloquea/desbloquea el credito seleccionado en la lista superior "
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "&Borrar credito"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7440
         TabIndex        =   89
         ToolTipText     =   "Borra el credito seleccionado en la lista superior"
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox CheckSoloRefinanciados 
         Caption         =   "Solo Refinanciados"
         Height          =   255
         Left            =   6960
         TabIndex        =   81
         ToolTipText     =   "Muestra solo los creditos generados por refinanciacion"
         Top             =   2420
         Width           =   1695
      End
      Begin VB.ComboBox ComboProvincias 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   79
         ToolTipText     =   "Lista de provincias"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox CheckProvincia 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   9120
         TabIndex        =   78
         ToolTipText     =   "Filtra los creditos por provincia de sellados"
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox CheckFinalizados 
         Caption         =   "Incluir finalizados"
         Height          =   255
         Left            =   5160
         TabIndex        =   71
         ToolTipText     =   "Incluye los creditos finalizados"
         Top             =   2420
         Width           =   1575
      End
      Begin VB.CheckBox CheckBloqueados 
         Caption         =   "Incluir bloqueados"
         Height          =   255
         Left            =   3360
         TabIndex        =   70
         ToolTipText     =   "Incluye los creditos bloqueados"
         Top             =   2420
         Width           =   1695
      End
      Begin VB.CommandButton CmdExportarCreditos 
         Caption         =   "Exportar creditos"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9120
         TabIndex        =   11
         ToolTipText     =   "Exporta la lista de creditos a una planilla Excel"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton CmdPlanilla 
         Caption         =   "Planilla creditos"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9120
         TabIndex        =   10
         ToolTipText     =   "Resumen de creditos"
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Refrescar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7440
         TabIndex        =   6
         ToolTipText     =   "Actualiza el saldo total"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtTotal 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   22
         Tag             =   "N"
         ToolTipText     =   "Saldo total de los creditos de la lista"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton CmdImprimirResumen 
         Caption         =   "Resumen"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9120
         TabIndex        =   9
         ToolTipText     =   "Imprime el resumen del credito seleccionado "
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton CmdImprimirLista 
         Caption         =   "&Imprimir creditos"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9120
         TabIndex        =   8
         ToolTipText     =   "Imprime la lista de creditos de la lista superior"
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdImprimirMutuo 
         Caption         =   "&Mutuo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9120
         TabIndex        =   7
         ToolTipText     =   "Imprime el contrato del credito seleccionado"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtContador1 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   21
         Tag             =   "N"
         ToolTipText     =   "Nº total de creditos de la lista"
         Top             =   2400
         Width           =   615
      End
      Begin MSComctlLib.ListView lvCreditos 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Lista de creditos"
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3836
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
         NumItems        =   16
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod.Prestamo"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Credito Nº"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha Credito"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cuotas"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Capital"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Total"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Fecha finalizacion"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Fecha bloqueo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ID Cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Importe Cuota"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Primer Vto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Observaciones"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Comercio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Vendedor"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label23 
         Caption         =   "Saldo Gral $:"
         Height          =   255
         Left            =   840
         TabIndex        =   63
         ToolTipText     =   "Saldo general de los creditos de la lista"
         Top             =   2400
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cuotas del credito:"
      ForeColor       =   &H00FF0000&
      Height          =   3645
      Left            =   120
      TabIndex        =   34
      Top             =   3360
      Width           =   10695
      Begin VB.CommandButton CmdPMC 
         Caption         =   "Exportar PMC"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7680
         TabIndex        =   100
         ToolTipText     =   "Exporta los saldos de las cuotas"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txttotalEXP 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   98
         Tag             =   "N"
         ToolTipText     =   "Total del credito"
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton CmdRestituirVto 
         Caption         =   "Restituir Vto"
         Enabled         =   0   'False
         Height          =   255
         Left            =   7680
         TabIndex        =   94
         ToolTipText     =   "Restituye el vencimiento original de la cuota seleccionada"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton CmdImprimirCuotas 
         Caption         =   "Imprimir cuo&tas"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9120
         TabIndex        =   88
         ToolTipText     =   "Imprime las cuotas del credito seleccionado"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton CmdImprimirCupones 
         Caption         =   "Imprimir cupones"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9120
         TabIndex        =   87
         ToolTipText     =   "Imprime las cuotas con codigo de barras"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdLista 
         Caption         =   "Cuotas x mes"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9120
         TabIndex        =   86
         ToolTipText     =   "Imprime cuotas agrupadas por meses"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton CmdListado 
         Caption         =   "Listado"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9120
         TabIndex        =   85
         ToolTipText     =   "Imprime el listado de cuotas seleccionadas"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton CmdExportarSaldos 
         Caption         =   "Exportar Saldos"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7680
         TabIndex        =   84
         ToolTipText     =   "Exporta los saldos de las cuotas"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton CmdExceptuar 
         Caption         =   "Exceptuar mora"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7680
         TabIndex        =   83
         ToolTipText     =   "Exceptua la mora de la cuota seleccionada"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton CmdCambiarVto 
         Caption         =   "Cambiar Vto"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7680
         TabIndex        =   82
         ToolTipText     =   "Cambia el vencimiento de la cuota seleccionada"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton CmdExportarCuotas 
         Caption         =   "Exportar cuotas"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9120
         TabIndex        =   64
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox TxtTotalMora 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   29
         Tag             =   "N"
         ToolTipText     =   "Importe total de interes por mora"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox TxtTotalImpuestos 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   28
         Tag             =   "N"
         ToolTipText     =   "Importe total de impuestos"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TxtTotalSeguros 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   27
         Tag             =   "N"
         ToolTipText     =   "Importe total de seguros"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox TxtTotalGastos 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   26
         Tag             =   "N"
         ToolTipText     =   "Importe total de gastos"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TxtTotalInteres 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   25
         Tag             =   "N"
         ToolTipText     =   "Importe total de interes"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox TxtTotalCapital 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   24
         Tag             =   "N"
         ToolTipText     =   "Importe total de capital"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TxtTotalCobrado 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   30
         Tag             =   "N"
         ToolTipText     =   "Importe total cobrado"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox TxtImporteTotal 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   31
         Tag             =   "N"
         ToolTipText     =   "Saldo pendiente del credito seleccionado"
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00800080&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   135
         Left            =   5640
         TabIndex        =   47
         ToolTipText     =   "Financiada (fuera de vigencia)"
         Top             =   3360
         Width           =   135
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   135
         Left            =   4680
         TabIndex        =   46
         ToolTipText     =   "Vencida en mora"
         Top             =   3360
         Width           =   135
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   135
         Left            =   3600
         TabIndex        =   45
         ToolTipText     =   "Pendiente al dia"
         Top             =   3360
         Width           =   135
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   135
         Left            =   2640
         TabIndex        =   44
         ToolTipText     =   "Cobrada o comodin(si tiene la letra ""C"")"
         Top             =   3360
         Width           =   135
      End
      Begin VB.CommandButton CmdComodin 
         Caption         =   "C&uota Comodin"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Marca/Desmarca la cuota seleccionada como comodin"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtContador2 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   23
         Tag             =   "N"
         ToolTipText     =   "Nº de cuotas del credito seleccionado en la lista superior"
         Top             =   3000
         Width           =   495
      End
      Begin MSComctlLib.ListView lvcuotas 
         Height          =   2475
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Lista de cuotas del credito seleccionado en la lista superior"
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4366
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   29
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Observaciones"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cod.Prestamo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Credito"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cuota"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cupon Nº"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Capital"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Interes"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "1º vto"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Imp.1º vto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "2º vto"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Imp.2º vto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Saldo"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Fecha cobro"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Imp.cobrado"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Gastos"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Otorgamiento"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Rec.2Vto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Rec.Refin."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Seguros"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Iva Interes"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Iva Seguros"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Iva Ot/Gastos"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "DiasMora"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Mora"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "IvaMora"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "Imp.Parcial"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "CodigoBarras"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "PF"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "RP"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label Label28 
         Caption         =   "Total Prestamo   $:"
         Height          =   255
         Left            =   6720
         TabIndex        =   99
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "Parcial"
         Height          =   255
         Left            =   360
         TabIndex        =   77
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "Exceptuada"
         Height          =   255
         Left            =   1440
         TabIndex        =   76
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "(E)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   75
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "Mora:"
         Height          =   255
         Left            =   4920
         TabIndex        =   62
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Impuestos:"
         Height          =   255
         Left            =   4920
         TabIndex        =   61
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Seguros:"
         Height          =   255
         Left            =   2880
         TabIndex        =   60
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Gastos:"
         Height          =   255
         Left            =   2880
         TabIndex        =   59
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Interes:"
         Height          =   255
         Left            =   840
         TabIndex        =   58
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Capital:"
         Height          =   255
         Left            =   840
         TabIndex        =   57
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cobrado $:"
         Height          =   255
         Left            =   7200
         TabIndex        =   56
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Saldo     $:"
         Height          =   255
         Left            =   7200
         TabIndex        =   55
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "[*]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   52
         ToolTipText     =   "Si tiene la letra ""C"" es cuota comodin"
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Refinanc."
         Height          =   255
         Left            =   5880
         TabIndex        =   51
         ToolTipText     =   "Financiada (fuera de vigencia)"
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Vencida"
         Height          =   255
         Left            =   4920
         TabIndex        =   50
         ToolTipText     =   "Vencida en mora"
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Pendiente"
         Height          =   255
         Left            =   3840
         TabIndex        =   49
         ToolTipText     =   "Pendiente al dia"
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Cobrada"
         Height          =   255
         Left            =   2880
         TabIndex        =   48
         ToolTipText     =   "Cobrada o comodin (si tiene la letra ""C"")"
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Cuotas:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2760
         Width           =   615
      End
   End
   Begin VB.Frame FrameFiltro 
      Caption         =   "Filtrar cuotas por:"
      ForeColor       =   &H00FF0000&
      Height          =   1125
      Left            =   120
      TabIndex        =   35
      Top             =   6960
      Width           =   8295
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   255
         Left            =   6480
         TabIndex        =   80
         ToolTipText     =   "Fecha de la consulta"
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   39882
      End
      Begin VB.ComboBox ComboOrden 
         Height          =   315
         ItemData        =   "FrmConsultarCreditos.frx":0069
         Left            =   120
         List            =   "FrmConsultarCreditos.frx":0076
         Style           =   2  'Dropdown List
         TabIndex        =   72
         ToolTipText     =   "Tipo de ordenamiento de los listados"
         Top             =   770
         Width           =   2055
      End
      Begin VB.CheckBox CheckRefinanciadas 
         Caption         =   "Ver refinanc."
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Tag             =   "no"
         ToolTipText     =   "Muestra las cuotas refinanciadas (sin vigencia) del credito seleccionado"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox CheckCobradas 
         Caption         =   "Ver cobradas"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Tag             =   "no"
         ToolTipText     =   "Muestra las cuotas cobradas del credito seleccionado"
         Top             =   600
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox CheckComodin 
         Caption         =   "Ver comodines"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Tag             =   "no"
         ToolTipText     =   "Muestra las cuotas comodin del credito seleccionado"
         Top             =   120
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox CheckTodosCreditos 
         Caption         =   "Todos los creditos"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         ToolTipText     =   "Muestra cuotas de todos los creditos o solo del credito seleccionado"
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox ComboOpcionesCuotas 
         Height          =   315
         ItemData        =   "FrmConsultarCreditos.frx":00C5
         Left            =   120
         List            =   "FrmConsultarCreditos.frx":00E1
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Seleccione el tipo de filtro"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame FrameFechas 
         Caption         =   "Periodo:"
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   4080
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   255
            Left            =   600
            TabIndex        =   18
            ToolTipText     =   "Fecha inicial"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   24641537
            CurrentDate     =   39259
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   255
            Left            =   600
            TabIndex        =   19
            ToolTipText     =   "Fecha final"
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   24641537
            CurrentDate     =   39259
         End
         Begin VB.Label Label16 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label15 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame FrameCupon 
         Height          =   1095
         Left            =   2280
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton CmdBuscarCupon 
            Caption         =   "Buscar"
            Height          =   255
            Left            =   1920
            TabIndex        =   69
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox TxtNumCupon 
            Height          =   285
            Left            =   240
            MaxLength       =   9
            TabIndex        =   67
            ToolTipText     =   "Nº de cupon (o comprobante)"
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Cupon Nº:"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame FrameMensajeMora 
         Height          =   1095
         Left            =   4080
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
         Begin VB.Label Label25 
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de consulta:"
         Height          =   255
         Left            =   6480
         TabIndex        =   95
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Orden:"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   540
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmConsultarCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AQUI SE CONSULTAN CREDITOS Y CUOTAS SEGUN DISTINTOS CRITERIOS Y FILTROS

Public IdCliente As Long

Private Sub CmdPMC_Click()
CmdPMC.Enabled = False
Me.MousePointer = vbHourglass
Call ExportarPMC
Me.MousePointer = vbDefault
CmdPMC.Enabled = True
End Sub

Private Sub CmdSellos_Click()

If Trim(ComboProvincias.Text) = "" Then
   MsgE "Debe ingresar una provincia"
Else
   Call GenerarSellosTXT(ComboProvincias.Text)
End If

End Sub
Function GenerarSellosTXT(Provincia As String)
    Dim rec                 As rdoResultset
    Dim ndia                As Integer
    Dim nMes                As Integer
    Dim nAnio               As Integer
    Dim Dia                 As Integer
    Dim Mes                 As Integer
    Dim Anio                As Integer
    Dim I                   As Integer
    Dim nCantRegDetalle     As Integer
    Dim cNombreArchivo      As String
    Dim cArchivoCompleto    As String
    Dim cHeader             As String
    Dim cDetalle            As String
    Dim cTrailer            As String
    Dim cPeriodo            As String
    Dim sql                 As String
    Dim Nombreprovincia     As String
    Dim FechaCredito        As String
    Dim StrCuil             As String
    Dim StrCliente          As String
    Dim Quincena            As String
    Dim Periodo             As String
    Dim impRetenido         As Currency
    Dim impCredito          As Currency
    Dim IdCredito           As Long
    Dim alicuota            As Single
    Dim StrDni              As String
    Dim StrImportetotal     As String
    Dim impCreditoEnt       As String
    Dim impCreditoDec       As String
    Dim impRetenidoEnt      As String
    Dim impRetenidoDec      As String
    Dim impCreditoStr       As String
    Dim impRetenidoStr      As String
    Dim nSec                As Integer
    On Error GoTo merror

    GenerarSellosTXT = False

            
    ndia = Format(CStr(Day(Date)), "00")
    nMes = Format(CStr(Month(Date)), "00")
    nAnio = Format(CStr(Year(Date)), "0000")
    cPeriodo = Format(nAnio, "0000") & Format(nMes, "00") & Format(ndia, "00")
    
    Nombreprovincia = "SELLOS" & Provincia
    cNombreArchivo = Nombreprovincia & cPeriodo & ".txt"
    
    If Provincia = "CHACO" Then
        cNombreArchivo = "RETAV.txt"
    End If
        
    If Provincia = "FORMOSA" Then
        cNombreArchivo = Nombreprovincia & cPeriodo & ".csv"
    End If

    If Not ExisteCarpeta() Then Exit Function

    cArchivoCompleto = "c:\ExportacionExcel\" & cNombreArchivo

    If Trim(Dir(cArchivoCompleto)) <> "" Then
        If Not MsgP("¿El archivo " & cArchivoCompleto & " ya existe. Desea borrarlo?") Then Exit Function
        Kill (cArchivoCompleto)
    End If
       
    Open cArchivoCompleto For Output As #1
            
            'MISIONES
            If Provincia = "MISIONES" Then
             alicuota = 1
             For I = 1 To lvCreditos.ListItems.Count
                Dia = Format(CStr(Day(lvCreditos.ListItems(I).SubItems(3))), "00")
                Mes = Format(CStr(Month(lvCreditos.ListItems(I).SubItems(3))), "00")
                Año = Format(CStr(Year(lvCreditos.ListItems(I).SubItems(3))), "0000")
                FechaCredito = Format(CStr(Dia), "00") & "-" & Format(CStr(Mes), "00") & "-" & Format(CStr(Año), "0000")
                
                'obtengo cuil del cliente
                sql = "select cuil,nombre,apellido from clientes where idcliente='" & CLng(lvCreditos.ListItems(I).SubItems(9)) & "'"
                           
                Set rec = cnSQL.OpenResultset(sql)
                
                If Not rec.EOF Then
                    StrCuil = rec("cuil")
                    StrCliente = Trim$(rec("apellido")) & " " & Trim$(rec("nombre"))
                Else
                    StrCuil = ""
                    StrCliente = ""
                End If
                
                IdCredito = CLng(lvCreditos.ListItems(I).SubItems(1))
                impCredito = ObtenerTotalCredito(IdCredito)
                impRetenido = alicuota * impCredito / 100
                
                impCreditoEnt = CStr(Fix(impCredito))
                impCreditoDec = Format(Fix((impCredito - Fix(impCredito)) * 100), "00")
                    
                impRetenidoEnt = CStr(Fix(impRetenido))
                impRetenidoDec = Format(Fix((impRetenido - Fix(impRetenido)) * 100), "00")
                
                cDetalle = lvCreditos.ListItems(I).SubItems(9) & "," & "0101," & StrCliente & "," & FechaCredito & "," & impCreditoEnt & "." & impCreditoDec & ",0.0100," & impRetenidoEnt & "." & impRetenidoDec & ",0," & StrCuil & "," & lvCreditos.ListItems.Item(I) & "," & ","
                
                Print #1, cDetalle
            Next I
            End If
            If Provincia = "ENTRE RIOS" Then
                alicuota = 1
                If ndia <= 15 Then
                    Quincena = (nMes * 2) - 1
                Else
                    Quincena = nMes * 2
                End If
                Periodo = Format(Quincena, "00") & "-" & nAnio
                cHeader = Periodo & ";502;" & "30-71029451-4;01;" & "502"
                Print #1, cHeader
                For I = 1 To lvCreditos.ListItems.Count
                    Dia = Format(CStr(Day(lvCreditos.ListItems(I).SubItems(3))), "00")
                    Mes = Format(CStr(Month(lvCreditos.ListItems(I).SubItems(3))), "00")
                    Año = Format(CStr(Year(lvCreditos.ListItems(I).SubItems(3))), "0000")
                    FechaCredito = Format(CStr(Dia), "00") & "-" & Format(CStr(Mes), "00") & "-" & Format(CStr(Año), "0000")
                    
                    IdCredito = CLng(lvCreditos.ListItems(I).SubItems(1))
                    impCredito = ObtenerTotalCredito(IdCredito)
                    impCreditoStr = Trim$(Format(impCredito, "#########0.00"))
                    impRetenidoStr = Trim$(Format(alicuota * impCredito / 100, "#########0.00"))
                    
                    cDetalle = "1912;" & lvCreditos.ListItems.Item(I) & ";" & FechaCredito & ";$;" & impCreditoStr & ";" & "10;" & impRetenidoStr
                  Print #1, cDetalle
                Next I
            End If
            
            If Provincia = "CHACO" Then
                             
                alicuota = 1
                For I = 1 To lvCreditos.ListItems.Count
                    Dia = Format(CStr(Day(lvCreditos.ListItems(I).SubItems(3))), "00")
                    Mes = Format(CStr(Month(lvCreditos.ListItems(I).SubItems(3))), "00")
                    Año = Format(CStr(Year(lvCreditos.ListItems(I).SubItems(3))), "0000")
                    FechaCredito = Format(CStr(Año), "0000") & Format(CStr(Mes), "00") & Format(CStr(Dia), "00")
                    
                    'obtengo doc del cliente
                    sql = "select numdocumento,nombre,apellido  from clientes where idcliente='" & CLng(lvCreditos.ListItems(I).SubItems(9)) & "'"
                           
                    Set rec = cnSQL.OpenResultset(sql)
                
                    If Not rec.EOF Then
                        StrDni = rec("numdocumento")
                        StrCliente = Trim$(Mid$(Trim$(rec("apellido")) & " " & Trim$(rec("nombre")), 1, 25))
                    Else
                        StrDni = ""
                        StrCliente = ""
                    End If
                    
                    IdCredito = CLng(lvCreditos.ListItems(I).SubItems(1))
                    impCredito = ObtenerTotalCredito(IdCredito)
                    impRetenido = alicuota * impCredito / 100
                    
                    cDetalle = "30710294514" & Space(20 - Len(lvCreditos.ListItems(I))) & lvCreditos.ListItems(I) & FechaCredito & " Art.15-23" & Space(49) & "-1" & Format(StrDni, "00000000000") & Space(25 - Len(StrCliente)) & StrCliente & Format(impCredito * 100, "00000000000") & "1" & "000100" & Format(impRetenido * 100, "00000000000")
                    Print #1, cDetalle
                Next I
            End If
            
            If Provincia = "FORMOSA" Then
                alicuota = 1
                For I = 1 To lvCreditos.ListItems.Count
                    Dia = Format(CStr(Day(lvCreditos.ListItems(I).SubItems(3))), "00")
                    Mes = Format(CStr(Month(lvCreditos.ListItems(I).SubItems(3))), "00")
                    Año = Format(CStr(Year(lvCreditos.ListItems(I).SubItems(3))), "0000")
                    FechaCredito = Format(CStr(Dia), "00") & "/" & Format(CStr(Mes), "00") & "/" & Format(CStr(Año), "0000")
                    
                    'obtengo doc del cliente
                    sql = "select numdocumento from clientes where idcliente='" & CLng(lvCreditos.ListItems(I).SubItems(9)) & "'"
                           
                    Set rec = cnSQL.OpenResultset(sql)
                
                    If Not rec.EOF Then
                        StrDni = rec("numdocumento")
                    Else
                        StrDni = ""
                    End If
                    
                    IdCredito = CLng(lvCreditos.ListItems(I).SubItems(1))
                    impCredito = ObtenerTotalCredito(IdCredito)
                    impRetenido = alicuota * impCredito / 100
                    
                    impCreditoEnt = CStr(Fix(impCredito))
                    impCreditoDec = Format(Fix((impCredito - Fix(impCredito)) * 100), "00")
                    
                    impRetenidoEnt = CStr(Fix(impRetenido))
                    impRetenidoDec = Format(Fix((impRetenido - Fix(impRetenido)) * 100), "00")
                    
                    cDetalle = StrDni & ";" & lvCreditos.ListItems(I).SubItems(2) & ";" & FechaCredito & ";" & lvCreditos.ListItems.Item(I) & ";" & "PAGARE" & ";" & "1.00" & ";" & impCreditoEnt & "." & impCreditoDec & ";" & impRetenidoEnt & "." & impRetenidoDec
                  Print #1, cDetalle
                Next I
            End If
            
            If Provincia = "CORRIENTES" Then
                nSec = 0
                alicuota = 1
                If Day(lvCreditos.ListItems(1).SubItems(3)) > 15 Then
                    nSec = 1000
                End If
                For I = 1 To lvCreditos.ListItems.Count
                    Dia = Format(CStr(Day(lvCreditos.ListItems(I).SubItems(3))), "00")
                    Mes = Format(CStr(Month(lvCreditos.ListItems(I).SubItems(3))), "00")
                    Año = Format(CStr(Year(lvCreditos.ListItems(I).SubItems(3))), "0000")
                    FechaCredito = Format(CStr(Dia), "00") & "/" & Format(CStr(Mes), "00") & "/" & Format(CStr(Año), "0000")
                    
                    'obtengo dni del cliente
                    sql = "select numdocumento from clientes where idcliente='" & CLng(lvCreditos.ListItems(I).SubItems(9)) & "'"
                    Set rec = cnSQL.OpenResultset(sql)
                
                    If Not rec.EOF Then
                        StrDni = rec("numdocumento")
                    Else
                        StrDni = ""
                    End If
                    rec.Close
                    
                    IdCredito = CLng(lvCreditos.ListItems(I).SubItems(1))
                    impCredito = ObtenerTotalCredito(IdCredito)
                    impRetenido = alicuota * impCredito / 100
                    
                    nSec = nSec + 1
                    cDetalle = nAnio & Format(nSec, "000000") & " " & Format(CStr(Año), "0000") & Format(CStr(Mes), "00") & " 30710294514 " & "00 " & FechaCredito & "    513" & "   1 " & Space(12 - Len(Format(impCredito, "########0.00"))) & Format(impCredito, "########0.00") & "      1 " & Space(12 - Len(Format(impRetenido, "########0.00"))) & Format(impRetenido, "########0.00") & " " & Format(StrDni, "00000000")
                  Print #1, cDetalle
                Next I
            End If
       
         
   
    Close #1
    
   MensajeMsg = "Se generó con éxito el archivo de siguiente archivo de sellos: " & cArchivoCompleto
   
   MsgI MensajeMsg
    
    GenerarSellosTXT = True

Exit Function
merror:
tratarerrores "Error exportando Sellos: " & Err.Number & " " & Err.Description
End Function

Private Sub cmdUnificarCupones_Click()
'imprime las cuponeras unificadas (para todos los créditos vigentes de un cliente)

cmdUnificarCupones.Enabled = False
If DatosImpresionCuotasOk() Then
   'imprimo los comprobantes con codigo de barras
   Call ImprimirCuotas(3, DTPicker5.Value)
End If
cmdUnificarCupones.Enabled = True

End Sub

Private Sub Form_Load()
On Error GoTo merror

Call RefrescarOpcionesSistema

Call LimpiarCampos(Me)

ComboOpciones.ListIndex = 0
ComboOpcionesCuotas.ListIndex = 0

If Not VG_APLICARSEGUNDOVENCIMIENTO Then
   lvcuotas.ColumnHeaders.Item(10).Width = 0
   lvcuotas.ColumnHeaders.Item(11).Width = 0
End If

'establezco el primer ordenamiento de las cuotas..por prestamo y vencimiento
ComboOrden.ListIndex = 0

Call CargarComboProvincias("provincias", ComboProvincias)

'este es para la consulta y exportacion de saldos
DTPicker5.Value = Date

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de consulta de creditos"
End Sub
Private Sub CmdCerrar_Click()
IdCliente = 0
Unload Me
End Sub
Private Sub CheckBloqueados_Click()
TxtTotal.Text = 0
Call ActualizarListas
End Sub
Private Sub CheckFinalizados_Click()
TxtTotal.Text = 0
Call ActualizarListas
End Sub
Private Sub CmdBuscarCupon_Click()
'filtra las cuotas por cupon
Call CargarCuotasCreditos
End Sub
Private Sub ComboOrden_Click()
'si cambia el orden tambien cambia el orden en pantalla
If ComboOrden.Text = "" Then Exit Sub
If ComboOrden.ListIndex = 1 Then
   CmdLista.Enabled = True
Else
   CmdLista.Enabled = False
End If
Call CargarCuotasCreditos
End Sub
Private Sub CmdBuscarCliente_Click()
'permite seleccionar un cliente
FrmClientesAbm.FormularioPadre = "CONSULTARCREDITOS"
Call CenterForm(FrmClientesAbm)
FrmClientesAbm.Show
End Sub
Private Sub CmdActualizar_Click()
Call ActualizarListas
End Sub
Private Sub CmdImprimirResumen_Click()
'imprime el estado de cuenta del credito seleccionado

If lvCreditos.ListItems.Count = 0 Then Exit Sub
Call ImprimirResumenCredito(lvCreditos.SelectedItem.SubItems(1), CDate(DTPicker5.Value))

End Sub
Private Sub CheckTodosCreditos_Click()
'filtra las cuotas de todos los creditos o solo del seleccionado
Call CargarCuotasCreditos
End Sub
Private Sub CheckCobradas_Click()
'incluye las cuotas cobradas
Call CargarCuotasCreditos
End Sub
Private Sub CheckComodin_Click()
'incluye las cuotas comodin
Call CargarCuotasCreditos
End Sub
Private Sub CheckRefinanciadas_Click()
'incluye las cuotas refinanciadas
Call CargarCuotasCreditos
End Sub
Private Sub DTPicker3_Change()
'si cambia la fecha del periodo de consulta
Call CargarCuotasCreditos
End Sub
Private Sub DTPicker4_Change()
'si cambia la fecha del periodo de consulta
Call CargarCuotasCreditos
End Sub
Private Sub ComboOpcionesCuotas_Click()
'opciones de filtro
FrameMensajeMora.Visible = False
FrameCupon.Visible = False
ComboOrden.Enabled = True

CheckComodin.Enabled = False
CheckCobradas.Enabled = False
CheckRefinanciadas.Enabled = False

If ComboOpcionesCuotas.Text = "Todas" Then
   FrameFechas.Visible = False
   CheckComodin.Enabled = True
   CheckCobradas.Enabled = True
   CheckRefinanciadas.Enabled = True
End If

If ComboOpcionesCuotas.Text = "Todas por fechas" Then
   FrameFechas.Visible = True
   CheckComodin.Enabled = True
   CheckCobradas.Enabled = True
   CheckRefinanciadas.Enabled = True
End If

If ComboOpcionesCuotas.Text = "Pendientes" Then
   FrameFechas.Visible = True
   CheckComodin.Value = 0
   CheckComodin.Enabled = False
   CheckCobradas.Value = 0
   CheckCobradas.Enabled = False
   CheckRefinanciadas.Value = 0
   CheckRefinanciadas.Enabled = False
End If
   
If ComboOpcionesCuotas.Text = "Cobradas" Then
   FrameFechas.Visible = True
   CheckComodin.Value = 0
   CheckComodin.Enabled = False
   CheckCobradas.Value = 1
   CheckCobradas.Enabled = False
   CheckRefinanciadas.Value = 0
   CheckRefinanciadas.Enabled = False
End If
   
'se entiende que es en mora al dia de la fecha
If ComboOpcionesCuotas.Text = "En mora" Then
   FrameFechas.Visible = False
   FrameMensajeMora.Visible = True
   CheckComodin.Value = 0
   CheckComodin.Enabled = False
   CheckCobradas.Value = 0
   CheckCobradas.Enabled = False
   CheckRefinanciadas.Value = 0
   CheckRefinanciadas.Enabled = False
End If

If ComboOpcionesCuotas.Text = "Financiadas" Then
   FrameFechas.Visible = True
   CheckComodin.Value = 0
   CheckComodin.Enabled = False
   CheckCobradas.Value = 0
   CheckCobradas.Enabled = False
   CheckRefinanciadas.Value = 1
   CheckRefinanciadas.Enabled = False
End If
   
  
If ComboOpcionesCuotas.Text = "Por cupon" Then
   FrameFechas.Visible = False
   FrameCupon.Visible = True
End If

If ComboOpcionesCuotas.Text = "Cobradas parcialmente" Then
   FrameFechas.Visible = True
   ComboOrden.ListIndex = 0
   ComboOrden.Enabled = False
   CheckComodin.Value = 0
   CheckComodin.Enabled = False
   CheckCobradas.Value = 0
   CheckCobradas.Enabled = False
   CheckRefinanciadas.Value = 0
   CheckRefinanciadas.Enabled = False
End If

Call CargarCuotasCreditos

End Sub
Private Sub cmdborrar_Click()
CmdBorrar.Enabled = False
Call BorrarCredito
If lvCreditos.ListItems.Count > 0 Then
   CmdBorrar.Enabled = True
End If
End Sub
Private Sub CmdBloquear_Click()
CmdBloquear.Enabled = False
Call BloquearCredito
If lvCreditos.ListItems.Count > 0 Then
   CmdBloquear.Enabled = True
End If
End Sub
Private Sub CargarCreditos()
'carga la lista de creditos
Dim rec As rdoResultset
Dim Nitem As ListItem
Dim ImporteTotal As Currency
On Error GoTo merror

lvCreditos.ListItems.Clear
lvcuotas.ListItems.Clear
TxtContador1.Text = 0
TxtContador2.Text = 0
txttotalEXP.Text = 0

If ComboOpciones.Text = "Por cliente" Then
   If Trim(TxtCliente.Text) = "" Then Exit Sub
End If

Set rec = CargarRecCreditos()

ImporteTotal = 0

If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lvCreditos.ListItems.Add(, , rec.rdoColumns("codprestamo"))
      Nitem.SubItems(1) = Format(rec.rdoColumns("idcredito"), "000000") & vbNullString
      Nitem.SubItems(2) = rec.rdoColumns("cliente") & vbNullString
      Nitem.SubItems(3) = rec.rdoColumns("fechacredito") & vbNullString
      Nitem.SubItems(4) = Format(rec.rdoColumns("numcuotas"), "00") & vbNullString
      Nitem.SubItems(5) = Format(rec.rdoColumns("importeafinanciar"), "0.00") & vbNullString
      Nitem.SubItems(6) = Format(rec.rdoColumns("importetotal"), "0.00") & vbNullString
      Nitem.SubItems(7) = rec.rdoColumns("fechafinalizacion") & vbNullString
      Nitem.SubItems(8) = rec.rdoColumns("fechabloqueo") & vbNullString
      Nitem.SubItems(9) = rec.rdoColumns("codcliente") & vbNullString
      Nitem.SubItems(10) = rec.rdoColumns("importecuota") & vbNullString
      Nitem.SubItems(11) = rec.rdoColumns("fechavencimiento1") & vbNullString
      
      'esto es para indicar si el credito es una refinanciacion
      If rec.rdoColumns("logic1") Then
         Nitem.SubItems(12) = "REFINANCIACION"
      Else
         Nitem.SubItems(12) = "CREDITO NORMAL"
      End If
      Nitem.SubItems(13) = rec.rdoColumns("cad1") & vbNullString
      Nitem.SubItems(14) = Trim(rec.rdoColumns("motivobloqueo") & vbNullString)
      Nitem.SubItems(15) = Trim(rec.rdoColumns("cad2") & vbNullString)
      
      ImporteTotal = CCur(ImporteTotal) + ObtenerSaldoCredito(rec.rdoColumns("idcredito"), DTPicker5.Value)
      rec.MoveNext
   Loop
End If

TxtContador1.Text = lvCreditos.ListItems.Count
TxtTotal.Text = Format(ImporteTotal, "0.00")
rec.Close
Exit Sub
merror:
tratarerrores "Error cargando la lista de creditos"
End Sub
Private Function CargarRecCreditos() As rdoResultset
'carga el registro de creditos
Dim sql As String
Dim Sqlbase As String
Dim Sqlcondicion As String
Dim IdProvincia As Long
On Error GoTo merror

'"(creditos.importefinanciado + creditos.importegastos + creditos.importeseguros + creditos.ivainteres + creditos.ivaseguros + creditos.ivaotgastos + creditos.ivamora + creditos.importerefinanciacion + creditos.importesellados + creditos.importeotorgamiento) as total," & _

Sqlbase = "SELECT creditos.*,planes.nombre as 'plan'," & _
          "(creditos.importefinanciado + creditos.importegastos + creditos.importeseguros + creditos.ivainteres + creditos.ivaseguros + creditos.ivaotgastos + creditos.importerefinanciacion + creditos.importeotorgamiento) as total," & _
          "clientes.idcliente as codcliente,clientes.numdocumento,clientes.fechanacimiento as nacimiento," & _
          "clientes.numlegajo,clientes.apellidogarante + ', ' + clientes.nombregarante as garante," & _
          "clientes.apellido + ', ' + clientes.nombre as cliente,clientes.domicilio," & _
          "localidades.nombre as localidad,provincias.nombre as provincia " & _
          "FROM provincias inner join(localidades inner join (clientes inner join (planes inner join creditos on planes.idplan=creditos.idplan) on clientes.idcliente=creditos.idcliente) on localidades.idlocalidad=clientes.idlocalidad) on provincias.idprovincia=creditos.idprovincia "

Sqlcondicion = ""

'carga solo los creditos vigentes de un cliente
If ComboOpciones.Text = "Por cliente" Then
   Sqlcondicion = "where creditos.idcliente=" & CLng(IdCliente) & _
                  " and creditos.fechacredito>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "'" & _
                  " and creditos.fechacredito<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "'"
   
   'si no incluyo bloqueados
   If CheckBloqueados.Value = 0 Then
      Sqlcondicion = Sqlcondicion & " and creditos.fechabloqueo is null"
   End If
   
   'si no incluyo finalizados
   If CheckFinalizados.Value = 0 Then
      Sqlcondicion = Sqlcondicion & " and creditos.fechafinalizacion is  null"
   End If
   If CheckSoloRefinanciados.Value = 1 Then
      'si muestro solo los refinanciados de ese cliente
      'Sqlcondicion = Sqlcondicion & " and creditos.logic1 is not Null"
      Sqlcondicion = Sqlcondicion & " and creditos.logic1 = 1"
      
   End If
End If

'carga todos los creditos vigentes incluyendo los refinanciados
If ComboOpciones.Text = "Creditos vigentes" Then
   Sqlcondicion = "where creditos.fechafinalizacion is null " & _
   "and creditos.fechabloqueo is null"
   
   If CheckSoloRefinanciados.Value = 1 Then
      'si muestro solo los refinanciados de ese cliente
      'Sqlcondicion = Sqlcondicion & " and creditos.logic1 is not Null"
      Sqlcondicion = Sqlcondicion & " and creditos.logic1 = 1"
   End If
End If

'carga todos los creditos finalizados en un periodo
If ComboOpciones.Text = "Creditos finalizados" Then
   Sqlcondicion = "where creditos.fechafinalizacion is not Null " & _
                  "and creditos.fechafinalizacion>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
                  "and creditos.fechafinalizacion<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "'"
                  
                  
   If CheckSoloRefinanciados.Value = 1 Then
      'si muestro solo los refinanciados de ese cliente
      'Sqlcondicion = Sqlcondicion & " and creditos.logic1 is not Null"
      Sqlcondicion = Sqlcondicion & " and creditos.logic1 = 1"
   End If

End If

'carga todos los creditos bloqueados en un periodo
If ComboOpciones.Text = "Creditos bloqueados" Then
   Sqlcondicion = "where creditos.fechabloqueo is not Null " & _
                  "and creditos.fechabloqueo>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
                  "and creditos.fechabloqueo<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "'"
   
   If CheckSoloRefinanciados.Value = 1 Then
      'si muestro solo los refinanciados de ese cliente
      'Sqlcondicion = Sqlcondicion & " and creditos.logic1 is not Null"
      Sqlcondicion = Sqlcondicion & " and creditos.logic1 = 1"
   End If

End If

'carga todos los creditos en un periodo
If ComboOpciones.Text = "Todos" Then
   Sqlcondicion = "where creditos.fechacredito>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
                  "and creditos.fechacredito<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "'"
                  
   'si no incluyo bloqueados
   If CheckBloqueados.Value = 0 Then
      Sqlcondicion = Sqlcondicion & " and creditos.fechabloqueo is Null"
   End If
   
   'si no inlcuyo finalizados
   If CheckFinalizados.Value = 0 Then
      Sqlcondicion = Sqlcondicion & " and creditos.fechafinalizacion is Null"
   End If
   
   If CheckSoloRefinanciados.Value = 1 Then
      'si muestro solo los refinanciados de ese cliente
      'Sqlcondicion = Sqlcondicion & " and creditos.logic1 is not Null"
      Sqlcondicion = Sqlcondicion & " and creditos.logic1 = 1"
   End If
End If

'nuevo filtro de 2009 por provincia de sellados
If CheckProvincia.Value = 1 Then
   If ComboProvincias.Text <> "" Then
      IdProvincia = CLng(ComboProvincias.ItemData(ComboProvincias.ListIndex))
      Sqlcondicion = Sqlcondicion & " and creditos.idprovincia='" & CLng(IdProvincia) & "'"
   End If
End If

'antes ordenaba por idcredito
sql = Sqlbase & Sqlcondicion & " order by creditos.codprestamo"

Set CargarRecCreditos = cnSQL.OpenResultset(sql)

Exit Function
merror:
tratarerrores "Error cargando el registro de creditos"
End Function
Private Function DatosCuponOk() As Boolean
'valida el filtro por cupon
On Error GoTo merror

DatosCuponOk = True

If ComboOpcionesCuotas.Text = "Por cupon" Then
   If Trim(TxtNumCupon.Text) = "" Then
      DatosCuponOk = False
      TxtNumCupon.SetFocus
      Exit Function
   End If
   If Not IsNumeric(TxtNumCupon.Text) Then
      DatosCuponOk = False
      TxtNumCupon.SetFocus
      Exit Function
   End If
   If CLng(TxtNumCupon.Text) <= 0 Then
      DatosCuponOk = False
      TxtNumCupon.SetFocus
      Exit Function
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosCuponOk"
End Function

Private Function CargarCuotasPend() As Boolean
Dim sql As String
Dim CodigoBarras As String
Dim Periodo As String
Dim nMesActual As Integer
Dim nAnioActual As Integer
Dim nMesActualLeido As Integer
Dim nAnioActualLeido As Integer
Dim IdCredito As Long
Dim NumFactura As Long
Dim NumCuota As Long
Dim ImporteVencimiento1 As Currency
Dim ImporteVencimiento2 As Currency
Dim ImporteRecargoVencimiento2 As Currency
Dim SaldoCuota As Currency
Dim SaldoCuotaFinal As Currency
Dim ImporteParcial As Currency
Dim SaldoCuota1erVenc As Currency
Dim Importe1erVenc As Currency
Dim ImporteMora As Currency
Dim IvaACobrarDevuelto As Currency
Dim IvaMora As Currency
Dim ImporteCobrado As Currency
Dim Vencimiento1 As Date
Dim Vencimiento2 As Date
Dim FechaCobro As Date
Dim Cobrada As Boolean
Dim rec As rdoResultset
On Error GoTo merror

    CargarCuotasPend = False
    IdCliente = CLng(lvCreditos.SelectedItem.SubItems(9))
    IdCredito = CLng(lvCreditos.SelectedItem.SubItems(1))
    
    sql = "select cuotas.* from cuotas, creditos " & _
          "where cuotas.idcredito = creditos.idcredito " & _
          "  and creditos.idCliente = " & IdCliente & _
          "  and creditos.fechabloqueo IS NULL and creditos.fechafinalizacion IS NULL and creditos.fecharefinanciacion IS NULL " & _
          "  and cuotas.logic1=0 " & _
          "order by cuotas.fechavencimiento1"
        
    Set rec = cnSQL.OpenResultset(sql)
    
    If rec.EOF Then
        MsgE " Sin creditos vigentes"
        Exit Function
    End If
    
    sql = "delete from cuotastemp"
        
    cnSQL.Execute sql
 
    NumCuota = 0
    
    Do While Not rec.EOF
        nMesActual = Month(rec.rdoColumns("fechavencimiento1"))
        nAnioActual = Year(rec.rdoColumns("fechavencimiento1"))
        nMesActualLeido = Month(rec.rdoColumns("fechavencimiento1"))
        nAnioActualLeido = Year(rec.rdoColumns("fechavencimiento1"))
        NumFactura = rec.rdoColumns("numfactura")
        ImporteVencimiento1 = 0
        ImporteVencimiento2 = 0
        SaldoCuota = 0
        ImporteCobrado = 0
        SaldoCuotaFinal = 0
        Vencimiento1 = CDate("2200/01/01")
        FechaCobro = CDate("1900/01/01")
        Cobrada = True
        Do While nMesActual = nMesActualLeido And _
                 nAnioActual = nAnioActualLeido And _
                 Not rec.EOF
            ImporteVencimiento1 = ImporteVencimiento1 + rec.rdoColumns("importevencimiento1")
            ImporteVencimiento2 = ImporteVencimiento2 + rec.rdoColumns("importevencimiento2")
            If rec.rdoColumns("fechavencimiento1") < Vencimiento1 Then
                Vencimiento1 = rec.rdoColumns("fechavencimiento1")
                Vencimiento2 = rec.rdoColumns("fechavencimiento2")
                CodigoBarras = rec.rdoColumns("codigobarras")
                Periodo = rec.rdoColumns("Periodo")
            End If
            
            ImporteCobrado = ImporteCobrado + rec.rdoColumns("importecobrado")
            
            ImporteParcial = ObtenerImporteParcialX(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
     
            'saldo de credimaco
            SaldoCuota = ObtenerSaldoCuotaX(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), CDate(DTPicker5.Value), SaldoCuota1erVenc)
            Importe1erVenc = ObtenerImporte1erVenc(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
            
            'si no esta cobrada
            If IsNull(rec.rdoColumns("fechacobro")) Then
                'si no esta refinanciada y no es comodin
                If IsNull(rec.rdoColumns("fecharefinanciacion")) And Not (rec.rdoColumns("cuotacomodin")) Then
                    'si estoy en mora actualizo
                    If CDate(DTPicker5.Value) > CDate(rec.rdoColumns("fechavencimiento2")) Then
                       'calculo la mora de forma habitual
                       'puedo pasarle el campo [exceptuada]
                       ImporteMora = CalculoMoraPendiente(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), False, rec.rdoColumns("importevencimiento1"), rec.rdoColumns("fechavencimiento1"), CDate(DTPicker5.Value), IvaACobrarDevuelto)
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
                       
                    End If
                End If
            End If
         
            SaldoCuotaFinal = SaldoCuotaFinal + SaldoCuota
            
            If IsNull(rec.rdoColumns("FechaCobro")) Then
                Cobrada = False
            Else
                If rec.rdoColumns("FechaCobro") > FechaCobro Then
                    FechaCobro = rec.rdoColumns("FechaCobro")
                End If
            End If
            rec.MoveNext
            If Not rec.EOF Then
                nMesActualLeido = Month(rec.rdoColumns("fechavencimiento1"))
                nAnioActualLeido = Year(rec.rdoColumns("fechavencimiento1"))
            End If
        Loop
        NumCuota = NumCuota + 1
        ImporteRecargoVencimiento2 = ImporteVencimiento2 - ImporteVencimiento1
        
        sql = "insert into cuotastemp (numfactura,idcredito,numcuota,importecuota," & _
          "fechavencimiento1,fechavencimiento2,importegastos,importeseguros," & _
          "importeimpuestos,importerecargovencimiento2,codigobarras," & _
          "importeamortizacion,importeinteres,periodo,importevencimiento1,importevencimiento2,otorgamiento,ivainteres,ivaseguros,ivaotorgamientogastos,saldocuota,importecobrado,pagadoparcial) " & _
          "values(" & NumFactura & "," & IdCredito & _
          "," & NumCuota & ",0" & _
          ",'" & ConvertirFechaSql(Vencimiento1, "DD/MM/YYYY") & "','" & ConvertirFechaSql(Vencimiento2, "DD/MM/YYYY") & _
          "',0,0,0," & ConvertirDblSql(ImporteRecargoVencimiento2) & ",'" & CodigoBarras & _
          "',0,0,'" & Periodo & "'," & ConvertirDblSql(ImporteVencimiento1) & _
          "," & ConvertirDblSql(ImporteVencimiento2) & ",0,0,0,0," & ConvertirDblSql(SaldoCuotaFinal) & "," & ConvertirDblSql(ImporteCobrado) & _
          "," & ConvertirDblSql(ImporteParcial) & ")"
        
        cnSQL.Execute sql
        
        If Cobrada Then
            sql = "update cuotastemp set fechacobro='" & ConvertirFechaSql(FechaCobro, "DD/MM/YYYY") & "' where idcredito=" & IdCredito & " and numcuota=" & NumCuota
            cnSQL.Execute sql
        End If
        
        CargarCuotasPend = True
    Loop
    
Exit Function
merror:
tratarerrores "Error en funcion DatosCuponOk"
End Function


Private Sub CargarCuotasCreditos()
'carga la lista de cuotas
Dim rec As rdoResultset
Dim Nitem As ListItem
Dim ImporteMora As Currency
Dim ImporteActualizado As Currency
Dim ImporteVencimiento2 As Currency
Dim ImporteTotal As Currency
Dim TotalCapital As Currency
Dim TotalInteres As Currency
Dim TotalGastos As Currency
Dim TotalSeguros As Currency
Dim TotalImpuestos As Currency
Dim TotalMora As Currency
Dim TotalParcial As Currency
Dim ImporteParcial As Currency
Dim SaldoCuota As Currency
Dim SaldoCuota1erVenc As Currency
Dim Importe1erVenc As Currency
Dim ImporteCobrado As Currency
Dim ImporteTotalCobrado As Currency
Dim Cad As String
Dim I As Long
Dim IvaMora As Currency
Dim DiasMora As Long
Dim RecargoCuota As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim IvaACobrarDevuelto  As Currency
Dim totExp As Currency
Dim nIdCreditoAnt As Long
On Error GoTo merror

lvcuotas.ListItems.Clear
TxtContador2.Text = 0
nIdCreditoAnt = 0
totExp = 0

If Not VerificarSeleccionLista(lvCreditos) Then Exit Sub

If Not DatosCuponOk() Then Exit Sub

Set rec = CargarRecCuotasCreditos()

I = 1
Do While Not rec.EOF
   'pongo nueva descripcion de cuotas
   Cad = ""
  
   If IsNull(rec.rdoColumns("fechacobro")) Then
      Cad = "Pendiente"
   End If
   
   If Not IsNull(rec.rdoColumns("fechacobro")) Then
      Cad = "Cobrada"
   End If
   
   If Not IsNull(rec.rdoColumns("fecharefinanciacion")) Then
      Cad = "Refinanciada"
   End If
   
   If rec.rdoColumns("cuotacomodin") Then
      Cad = "Comodin"
   End If
   
   If rec.rdoColumns("cobrosparciales") Then
      Cad = Cad & "(*)"
   End If
   
   If rec.rdoColumns("exceptuada") Then
      Cad = Cad & "(E)"
   End If
      
   Set Nitem = lvcuotas.ListItems.Add(, , Cad)
     
   Nitem.SubItems(1) = rec.rdoColumns("codprestamo") & vbNullString
   Nitem.SubItems(2) = Format(rec.rdoColumns("idcredito"), "000000") & vbNullString
   Nitem.SubItems(3) = Format(rec.rdoColumns("numcuota"), "00") & vbNullString
   Nitem.SubItems(4) = Format(rec.rdoColumns("numfactura"), "0000000") & vbNullString
   Nitem.SubItems(5) = Format(rec.rdoColumns("importeamortizacion"), "0.00") & vbNullString
   Nitem.SubItems(6) = Format(rec.rdoColumns("importeinteres"), "0.00") & vbNullString
   Nitem.SubItems(7) = rec.rdoColumns("fechavencimiento1") & vbNullString
   'este es el importe del vto1
   Nitem.SubItems(8) = Format(rec.rdoColumns("importetotal"), "0.00") & vbNullString
   Nitem.SubItems(9) = rec.rdoColumns("fechavencimiento2") & vbNullString
   ImporteVencimiento2 = CCur(CCur(rec.rdoColumns("importetotal")) + CCur(rec.rdoColumns("importerecargovencimiento2")))
   
   ImporteParcial = ObtenerImporteParcialX(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
   Nitem.SubItems(10) = Format(rec.rdoColumns("ImporteVencimiento2"), "0.00") & vbNullString
     
   'por defecto pongo en azul
   lvcuotas.ListItems.Item(I).ForeColor = &HFF0000
   lvcuotas.ListItems.Item(I).ListSubItems(1).ForeColor = &HFF0000
   lvcuotas.ListItems.Item(I).ListSubItems(2).ForeColor = &HFF0000
   lvcuotas.ListItems.Item(I).ListSubItems(3).ForeColor = &HFF0000
   lvcuotas.ListItems.Item(I).ListSubItems(4).ForeColor = &HFF0000
   lvcuotas.ListItems.Item(I).ListSubItems(5).ForeColor = &HFF0000
     
   SaldoCuota = ObtenerSaldoCuotaX(rec.rdoColumns("idcredito"), rec.rdoColumns("NumCuota"), CDate(DTPicker5.Value), SaldoCuota1erVenc)
   Importe1erVenc = ObtenerImporte1erVenc(rec.rdoColumns("idcredito"), rec.rdoColumns("NumCuota"))
   ImporteActualizado = 0
   ImporteMora = 0
   IvaMora = 0
   DiasMora = 0
   If CDate(DTPicker5.Value) > CDate(rec.rdoColumns("fechavencimiento1")) Then
        DiasMora = CDate(CDate(DTPicker5.Value)) - CDate(rec.rdoColumns("fechavencimiento1"))
        If SaldoCuota = 0 Then
            DiasMora = 0
        End If
   End If
   'si no esta cobrada actualizo el importe si es necesario
   If IsNull(rec.rdoColumns("fechacobro")) And Not rec.rdoColumns("cuotacomodin") And IsNull(rec.rdoColumns("fecharefinanciacion")) Then
      'esto funciona para ambos vencimientos (si hay un solo vto ambos son iguales)
      If CDate(DTPicker5.Value) > CDate(rec.rdoColumns("fechavencimiento2")) Then
         'esto es para mostrar en la columna correspondiente
         'calculo la mora en forma habitual
         'puedo pasarle el campo [exceptuada]
         ImporteMora = CalculoMoraPendiente(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("exceptuada"), Importe1erVenc, rec.rdoColumns("fechavencimiento1"), CDate(DTPicker5.Value), IvaACobrarDevuelto)
         '''''''********ImporteMora = MoraPendiente + CalcularInteresMoraZZ(rec.rdoColumns("exceptuada"), SaldoCalculoMora, FechaCalculoMora, CDate(DTPicker5.Value))
         If VG_APLICARIMPUESTOS Then
            If VG_IMPUESTOSCREDIMACO Then
               'calculo el iva de la mora
               IvaMora = IvaACobrarDevuelto
            End If
         End If
         '''''''********SoloMoraCobrada = ObtenerMoraCobrada(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
         '''''''********SoloIvaMoraCobrada = ObtenerIvaMoraCobrada(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
         '''''''********If CCur(ImporteMora) <= CCur(SoloMoraCobrada) Then
         '''''''********   ImporteMora = 0
         '''''''********Else
         '''''''********   'si es mayor la mora es solo la diferencia
         '''''''********   ImporteMora = CCur(ImporteMora) - CCur(SoloMoraCobrada)
         '''''''********End If
         '''''''********If CCur(IvaMora) <= CCur(SoloIvaMoraCobrada) Then
         '''''''********   IvaMora = 0
         '''''''********Else
         '''''''********   'si es mayor la mora es solo la diferencia
         '''''''********   IvaMora = CCur(IvaMora) - CCur(SoloIvaMoraCobrada)
         '''''''********End If
         
         ImporteActualizado = CCur(SaldoCuota1erVenc) + CCur(ImporteMora) + CCur(IvaMora)
         
         'pongo en rojo sin cobrar en mora
         lvcuotas.ListItems.Item(I).ForeColor = vbRed
         lvcuotas.ListItems.Item(I).ListSubItems(1).ForeColor = vbRed
         lvcuotas.ListItems.Item(I).ListSubItems(2).ForeColor = vbRed
         lvcuotas.ListItems.Item(I).ListSubItems(3).ForeColor = vbRed
         lvcuotas.ListItems.Item(I).ListSubItems(4).ForeColor = vbRed
         lvcuotas.ListItems.Item(I).ListSubItems(5).ForeColor = vbRed
         
      Else
         'no hay mora...
         ImporteActualizado = CCur(SaldoCuota)
      End If
   Else
      'si esta refinanciada la pongo en morado
      If Not IsNull(rec.rdoColumns("fecharefinanciacion")) Then
         lvcuotas.ListItems.Item(I).ForeColor = &H800080
         lvcuotas.ListItems.Item(I).ListSubItems(1).ForeColor = &H800080
         lvcuotas.ListItems.Item(I).ListSubItems(2).ForeColor = &H800080
         lvcuotas.ListItems.Item(I).ListSubItems(3).ForeColor = &H800080
         lvcuotas.ListItems.Item(I).ListSubItems(4).ForeColor = &H800080
         lvcuotas.ListItems.Item(I).ListSubItems(5).ForeColor = &H800080
      Else
         'esta cobrada o es comodin la pongo en verde
         lvcuotas.ListItems.Item(I).ForeColor = &H8000&
         lvcuotas.ListItems.Item(I).ListSubItems(1).ForeColor = &H8000&
         lvcuotas.ListItems.Item(I).ListSubItems(2).ForeColor = &H8000&
         lvcuotas.ListItems.Item(I).ListSubItems(3).ForeColor = &H8000&
         lvcuotas.ListItems.Item(I).ListSubItems(4).ForeColor = &H8000&
         lvcuotas.ListItems.Item(I).ListSubItems(5).ForeColor = &H8000&
      End If
   End If
         
   If ImporteActualizado > 0 Then
        Nitem.SubItems(11) = Format(ImporteActualizado, "0.00") & vbNullString
   Else
        Nitem.SubItems(11) = Format(0, "0.00") & vbNullString
   End If
   
   Nitem.SubItems(12) = rec.rdoColumns("fechacobro") & vbNullString
   
   If Not IsNull(rec.rdoColumns("fechacobro")) Then
      ImporteCobrado = CCur(rec.rdoColumns("importecobrado"))
   Else
      'trae lo cobrado hasta la fecha sin incluir mora e iva mora cobrada
      ImporteCobrado = CCur(ImporteParcial)
   End If
   
  ' If IsNull(rec.rdoColumns("fechacobro")) And IsNull(Format(ImporteCobrado, "0.00")) Then
  '  TotExp = TotExp + ImporteActualizado
  '  txttotalEXP = Format(TotExp, "0.00")
  ' Else
  '  TotExp = TotExp + ImporteActualizado + ImporteCobrado
  '  txttotalEXP = Format(TotExp, "0.00")
  ' End If
   
   '''''''TotExp = TotExp + Format(rec.rdoColumns("importetotal"), "0.00")
   'txttotalEXP = Format(TotExp, "0.00")
   
   If nIdCreditoAnt <> rec.rdoColumns("idcredito") Then
        totExp = totExp + ObtenerTotalCredito(rec.rdoColumns("idcredito"))
        nIdCreditoAnt = rec.rdoColumns("idcredito")
        txttotalEXP = Format(totExp, "0.00")
    End If

   
   Nitem.SubItems(13) = Format(ImporteCobrado, "0.00") & vbNullString
      
   Nitem.SubItems(14) = Format(rec.rdoColumns("importegastos"), "0.00") & vbNullString
   Nitem.SubItems(15) = Format(rec.rdoColumns("otorgamiento"), "0.00") & vbNullString
   Nitem.SubItems(16) = Format(rec.rdoColumns("importerecargovencimiento2"), "0.00") & vbNullString
   Nitem.SubItems(17) = Format(rec.rdoColumns("importerefinanciacion"), "0.00") & vbNullString
   Nitem.SubItems(18) = Format(rec.rdoColumns("importeseguros"), "0.00") & vbNullString
   Nitem.SubItems(19) = Format(rec.rdoColumns("ivainteres"), "0.00") & vbNullString
   Nitem.SubItems(20) = Format(rec.rdoColumns("ivaseguros"), "0.00") & vbNullString
   Nitem.SubItems(21) = Format(rec.rdoColumns("ivaotorgamientogastos"), "0.00") & vbNullString
   Nitem.SubItems(22) = Format(DiasMora, "000") & vbNullString
   Nitem.SubItems(23) = Format(ImporteMora, "0.00") & vbNullString
   Nitem.SubItems(24) = Format(IvaMora, "0.00") & vbNullString
   Nitem.SubItems(25) = Format(ImporteParcial, "0.00") & vbNullString
   Nitem.SubItems(26) = rec.rdoColumns("codigobarras") & vbNullString
   
   If rec.rdoColumns("pagofacil") Then
      Nitem.SubItems(27) = "SI"
   Else
      Nitem.SubItems(27) = " "
   End If
   If rec.rdoColumns("rapipago") Then
      Nitem.SubItems(28) = "SI"
   Else
      Nitem.SubItems(28) = " "
   End If
   
   'sumo campos totales
   TotalCapital = CCur(TotalCapital) + CCur(rec.rdoColumns("importeamortizacion"))
   TotalInteres = CCur(TotalInteres) + CCur(rec.rdoColumns("importeinteres"))
   TotalGastos = CCur(TotalGastos) + CCur(rec.rdoColumns("importegastos")) + CCur(rec.rdoColumns("otorgamiento"))
   TotalSeguros = CCur(TotalSeguros) + CCur(rec.rdoColumns("importeseguros"))
   TotalImpuestos = CCur(TotalImpuestos) + CCur(rec.rdoColumns("importeimpuestos"))
   TotalParcial = CCur(TotalParcial) + CCur(ImporteParcial)
   TotalMora = CCur(TotalMora) + CCur(ImporteMora)
   ImporteTotal = CCur(ImporteTotal) + Round(CCur(ImporteActualizado), 2)
   
   If CCur(rec.rdoColumns("importecobrado")) > 0 Then
      ImporteTotalCobrado = CCur(ImporteTotalCobrado) + CCur(rec.rdoColumns("importecobrado"))
   Else
      ImporteTotalCobrado = CCur(ImporteTotalCobrado) + CCur(ImporteParcial)
   End If
   
   'pongo todas en bold
   lvcuotas.ListItems.Item(I).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(1).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(2).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(3).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(4).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(5).Bold = True
     
      
   rec.MoveNext
   
   I = I + 1
Loop

TxtContador2.Text = lvcuotas.ListItems.Count
TxtTotalCapital.Text = Format(TotalCapital, "0.00")
TxtTotalInteres.Text = Format(TotalInteres, "0.00")
TxtTotalGastos.Text = Format(TotalGastos, "0.00")
TxtTotalSeguros.Text = Format(TotalSeguros, "0.00")
TxtTotalImpuestos.Text = Format(TotalImpuestos, "0.00")
TxtTotalMora.Text = Format(TotalMora, "0.00")

'estos 2 no los actualiza correctamente cuando esta seleccionada la
'casilla TODOS LOS CREDITOS...los deja siempre igual

TxtImporteTotal.Text = Format(ImporteTotal, "0.00")
TxtTotalCobrado.Text = Format(ImporteTotalCobrado, "0.00")
rec.Close
Exit Sub
merror:
tratarerrores "Error cargando cuotas de creditos"
End Sub
Private Function CargarRecCuotasCreditos() As rdoResultset
'carga el registro de cuotas
Dim sql As String
Dim IdCredito As Long
Dim CondicionFiltro As String
Dim CondicionCliente As String
Dim CondicionTipoCreditos As String
Dim CondicionRangoCreditos As String
Dim CondicionOrden As String
Dim CondicionProvincia As String
Dim IdProvincia As Long
On Error GoTo merror
  
IdCredito = CLng(lvCreditos.SelectedItem.SubItems(1))

'por ahora traigo de un solo credito por vez
If CheckTodosCreditos.Value = 0 Then
   CondicionFiltro = "cuotas.idcredito='" & CLng(IdCredito) & "'"
Else
   '***
   CondicionFiltro = "1=1"
End If

'por defecto todas las provincias
CondicionProvincia = " and 1=1"
If CheckProvincia.Value = 1 Then
   'si hay una provincia seleccionada
   If ComboProvincias.Text <> "" Then
      IdProvincia = CLng(ComboProvincias.ItemData(ComboProvincias.ListIndex))
      CondicionProvincia = " and creditos.idprovincia='" & CLng(IdProvincia) & "'"
   End If
End If
   
If ComboOpcionesCuotas.Text = "Todas" Then
   'traigo las cuotas de los creditos actuales en el rango seleccionado
   'cobradas y pendientes
   CondicionFiltro = CondicionFiltro
End If
   
If ComboOpcionesCuotas.Text = "Todas por fechas" Then
   'traigo las cuotas de los creditos actuales en el rango seleccionado
   'cobradas y pendientes
   CondicionFiltro = CondicionFiltro & _
   " and cuotas.fechavencimiento1>= '" & ConvertirFechaSql(CDate(DTPicker3.Value), "DD/MM/YYYY") & "' " & _
   "and cuotas.fechavencimiento1<='" & ConvertirFechaSql(CDate(DTPicker4.Value), "DD/MM/YYYY") & "'"
End If
   
If ComboOpcionesCuotas.Text = "Cobradas" Then
   CondicionFiltro = CondicionFiltro & _
   " and cuotas.fechacobro is not Null " & _
   "and cuotas.fechacobro>='" & ConvertirFechaSql(CDate(DTPicker3.Value), "DD/MM/YYYY") & "' " & _
   "and cuotas.fechacobro<='" & ConvertirFechaSql(CDate(DTPicker4.Value), "DD/MM/YYYY") & "'"
End If
   
If ComboOpcionesCuotas.Text = "Cobradas parcialmente" Then
   'esta es nueva..le falta que respete el rango(antes tomaba las fechas de cobros parciales parciales)..obligaba a usar distinct..despues problemas con los ordenamientos
   'por los cobros parciales cobrados..puede en cambio
   'solo tener en cuenta la fecha de vencimiento1 o 2
   CondicionFiltro = CondicionFiltro & _
   " and cuotas.fechacobro is Null " & _
   "and cuotas.cobrosparciales = 'True' " & _
   "and cuotas.fechavencimiento1>='" & ConvertirFechaSql(CDate(DTPicker3.Value), "DD/MM/YYYY") & "' " & _
   "and cuotas.fechavencimiento1<='" & ConvertirFechaSql(CDate(DTPicker4.Value), "DD/MM/YYYY") & "'"
End If
   
'solo pendientes de creditos vigentes
If ComboOpcionesCuotas.Text = "Pendientes" Then
   CondicionFiltro = CondicionFiltro & _
   " and cuotas.fechacobro is null " & _
   "and cuotas.fechavencimiento1>='" & ConvertirFechaSql(CDate(DTPicker3.Value), "DD/MM/YYYY") & "' " & _
   "and cuotas.fechavencimiento1<='" & ConvertirFechaSql(CDate(DTPicker4.Value), "DD/MM/YYYY") & "'"
End If
   
'en mora de creditos vigentes
'en mora al dia de la fecha lo cambie por el dtpicker5
If ComboOpcionesCuotas.Text = "En mora" Then
   CondicionFiltro = CondicionFiltro & _
   " and cuotas.fechacobro is Null " & _
   "and cuotas.fechavencimiento2 < '" & ConvertirFechaSql(CDate(DTPicker5.Value), "DD/MM/YYYY") & "'"
End If
   
If ComboOpcionesCuotas.Text = "Financiadas" Then
   CondicionFiltro = CondicionFiltro & " and cuotas.fecharefinanciacion is not Null and cuotas.fecharefinanciacion>='" & ConvertirFechaSql(CDate(DTPicker3.Value), "DD/MM/YYYY") & "' and cuotas.fecharefinanciacion<='" & ConvertirFechaSql(CDate(DTPicker4.Value), "DD/MM/YYYY") & "'"
End If

'SI es por cupon hago un filtro
If ComboOpcionesCuotas.Text = "Por cupon" Then
   CondicionFiltro = "cuotas.numfactura='" & CLng(TxtNumCupon.Text) & "'"
   CondicionOrden = " order by cuotas.numfactura"
   CondicionCliente = " and 1=1"
   CondicionTipoCreditos = " and 1=1"
   CondicionRangoCreditos = " and 1=1"
Else
   'solo tiene en cuenta los check si no es por cupon
   'si imprimo cuotas comodin
   
   If CheckComodin.Value = 1 Then
      CondicionFiltro = CondicionFiltro
   Else
      CondicionFiltro = CondicionFiltro & " and cuotas.cuotacomodin = 'False'"
   End If
   
   'si imprimo cuotas cobradas
   If CheckCobradas.Value = 1 Then
      CondicionFiltro = CondicionFiltro
   Else
      CondicionFiltro = CondicionFiltro & " and cuotas.fechacobro is Null"
   End If

   'si imprimo cuotas refinanciadas
   If CheckRefinanciadas.Value = 1 Then
      CondicionFiltro = CondicionFiltro
   Else
      CondicionFiltro = CondicionFiltro & " and cuotas.fecharefinanciacion is Null"
   End If
     
   'AHORA USARE CONDICIONTIPOCREDITOS
   'si no incluyo creditos bloqueados
   CondicionTipoCreditos = " and 1=1"
   If CheckBloqueados.Value = 0 Then
      CondicionTipoCreditos = CondicionTipoCreditos & " and creditos.fechabloqueo is null"
   Else
      If ComboOpciones.Text = "Creditos bloqueados" Then
         CondicionTipoCreditos = CondicionTipoCreditos & _
         " and creditos.fechabloqueo is not Null"
      End If
   End If
   'si no incluyo finalizados
   If CheckFinalizados.Value = 0 Then
      'trae solo los no finalizados
      CondicionTipoCreditos = CondicionTipoCreditos & " and creditos.fechafinalizacion is null"
   Else
      'en este caso trae solo los finalizados
      If ComboOpciones.Text = "Creditos finalizados" Then
         CondicionTipoCreditos = CondicionTipoCreditos & _
         " and creditos.fechafinalizacion is not Null"
     End If
   End If
   
   'aca excluye por si solo a los que no son refinanciados sin importar lo
   'que haya seleccionado arriba
   If CheckSoloRefinanciados.Value = 1 Then
      CondicionTipoCreditos = CondicionTipoCreditos & " and creditos.logic1 = 1 "
   End If
   
    
   CondicionCliente = " and 1=1"
   'si estoy por cliente y hay uno seleccionado
   If ComboOpciones.Text = "Por cliente" And IdCliente > 0 Then
      CondicionCliente = " and creditos.idcliente='" & CLng(IdCliente) & "'"
   End If

   CondicionRangoCreditos = " and 1=1"
   If ComboOpciones.Text <> "Creditos vigentes" Then
      If ComboOpciones.Text = "Creditos bloqueados" Then
         CondicionRangoCreditos = " and creditos.fechabloqueo>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' and creditos.fechabloqueo<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "'"
      End If
      If ComboOpciones.Text = "Creditos finalizados" Then
         CondicionRangoCreditos = " and creditos.fechafinalizacion>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' and creditos.fechafinalizacion<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "'"
      End If
      If ComboOpciones.Text = "Por cliente" Or ComboOpciones.Text = "Todos" Then
         CondicionRangoCreditos = " and creditos.fechacredito>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' and creditos.fechacredito<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "'"
      End If
   End If

   'solo tiene en cuenta el orden si no es por cupon
   'ordeno por prestamo,cuota y vencimiento (la cuota esta implicita con el vto)
   If ComboOrden.ListIndex = 0 Then
      CondicionOrden = " order by creditos.codprestamo,cuotas.fechavencimiento1"
   End If
   'si ordeno por vencimiento,prestamo(la cuota esta implicita con el vto)
   If ComboOrden.ListIndex = 1 Then
      CondicionOrden = " order by cuotas.fechavencimiento1,creditos.codprestamo"
   End If
   
   'solo dejo ordenar por cliente si no es parciales porque se cuelga
   'por el distinct
   If ComboOpcionesCuotas.Text <> "Cobradas parcialmente" Then
      'ordeno por cliente,credito y cuota
      If ComboOrden.ListIndex = 2 Then
         CondicionOrden = " order by clientes.apellido + clientes.nombre,creditos.idcredito,cuotas.numcuota"
      End If
   Else
      'si son parciales los ordena por el primer criterio
      ComboOrden.ListIndex = 0
      CondicionOrden = " order by creditos.idcredito,cuotas.numcuota"
   End If
   
End If

sql = "select provincias.nombre as provincia,localidades.nombre as localidad," & _
      "localidades.codigopostal,clientes.domicilio,clientes.telefono,clientes.tipoiva,clientes.cuil," & _
      "clientes.numdocumento,clientes.fechanacimiento as nacimiento,clientes.numlegajo,clientes.apellido + ', ' + clientes.nombre as cliente,creditos.numcuotas," & _
      "creditos.idcredito,creditos.codprestamo,creditos.motivobloqueo,creditos.fechacredito,cuotas.periodo,cuotas.numcuota,cuotas.numfactura,cuotas.cobrosparciales," & _
      "cuotas.fechavencimiento1,cuotas.fechavencimiento2,cuotas.importeamortizacion,cuotas.importeinteres," & _
      "cuotas.importecuota,cuotas.importerecargovencimiento2,cuotas.fecharefinanciacion,cuotas.otorgamiento,cuotas.formacobro," & _
      "cuotas.importegastos,cuotas.importeseguros,cuotas.importeimpuestos,cuotas.importeparcial,cuotas.ivainteres,cuotas.ivaseguros,cuotas.ivaotorgamientogastos,cuotas.ivamora," & _
      "cuotas.importedescuentos,cuotas.importerecargos,cuotas.importemora,cuotas.importerefinanciacion,cuotas.importerenovacion," & _
      "cuotas.fechacobro,cuotas.importecobrado,cuotas.codigobarras,cuotas.cuotacomodin,cuotas.pagofacil,cuotas.rapipago,cuotas.logic1 as exceptuada,cuotas.logic2 as vtocambiado," & _
      "(cuotas.importevencimiento1) as importetotal,cuotas.importevencimiento2 " & _
      "from provincias inner join (localidades inner join " & _
      "(clientes inner join (creditos inner join cuotas on " & _
      "creditos.idcredito=cuotas.idcredito) on " & _
      "clientes.idcliente=creditos.idcliente) on " & _
      "localidades.idlocalidad=clientes.idlocalidad) on " & _
      "provincias.idprovincia=creditos.idprovincia " & _
      "where " & CondicionFiltro & CondicionCliente & CondicionTipoCreditos & CondicionRangoCreditos & CondicionProvincia & CondicionOrden

Set CargarRecCuotasCreditos = cnSQL.OpenResultset(sql)

Exit Function
merror:
tratarerrores "Error cargando el registro de cuotas"
End Function
Private Function CargarRecCuotasTempCreditos() As rdoResultset
'carga el registro de cuotas
Dim sql As String
Dim IdCredito As Long
Dim CondicionFiltro As String
Dim CondicionCliente As String
Dim CondicionTipoCreditos As String
Dim CondicionRangoCreditos As String
Dim CondicionOrden As String
Dim CondicionProvincia As String
Dim IdProvincia As Long
On Error GoTo merror
  
IdCredito = CLng(lvCreditos.SelectedItem.SubItems(1))


sql = "select provincias.nombre as provincia,localidades.nombre as localidad," & _
      "localidades.codigopostal,clientes.domicilio,clientes.telefono,clientes.tipoiva,clientes.cuil," & _
      "clientes.numdocumento,clientes.fechanacimiento as nacimiento,clientes.numlegajo,clientes.apellido + ', ' + clientes.nombre as cliente,creditos.numcuotas," & _
      "creditos.idcredito,creditos.codprestamo,creditos.motivobloqueo,creditos.fechacredito,cuotastemp.periodo,cuotastemp.numcuota,cuotastemp.numfactura,cuotastemp.cobrosparciales," & _
      "cuotastemp.fechavencimiento1,cuotastemp.fechavencimiento2,cuotastemp.importeamortizacion,cuotastemp.importeinteres," & _
      "cuotastemp.importecuota,cuotastemp.importerecargovencimiento2,cuotastemp.fecharefinanciacion,cuotastemp.otorgamiento,cuotastemp.formacobro," & _
      "cuotastemp.importegastos,cuotastemp.importeseguros,cuotastemp.importeimpuestos,cuotastemp.importeparcial,cuotastemp.ivainteres,cuotastemp.ivaseguros,cuotastemp.ivaotorgamientogastos,cuotastemp.ivamora," & _
      "cuotastemp.importedescuentos,cuotastemp.importerecargos,cuotastemp.importemora,cuotastemp.importerefinanciacion,cuotastemp.importerenovacion," & _
      "cuotastemp.fechacobro,cuotastemp.importecobrado,cuotastemp.codigobarras,cuotastemp.cuotacomodin,cuotastemp.pagofacil,cuotastemp.rapipago,cuotastemp.logic1 as exceptuada,cuotastemp.logic2 as vtocambiado," & _
      "cuotastemp.importevencimiento1,cuotastemp.importevencimiento2,cuotastemp.saldocuota,cuotastemp.pagadoparcial " & _
      "from provincias inner join (localidades inner join " & _
      "(clientes inner join (creditos inner join cuotastemp on " & _
      "creditos.idcredito=cuotastemp.idcredito) on " & _
      "clientes.idcliente=creditos.idcliente) on " & _
      "localidades.idlocalidad=clientes.idlocalidad) on " & _
      "provincias.idprovincia=creditos.idprovincia " & _
      "where creditos.idcredito = " & IdCredito

Set CargarRecCuotasTempCreditos = cnSQL.OpenResultset(sql)

Exit Function
merror:
tratarerrores "Error cargando el registro de cuotas"
End Function

Private Function PuedoBorrarCredito(ByVal IdCredito As Long) As Boolean
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

PuedoBorrarCredito = True

'reviso en cuotas
sql = "select idcredito " & _
      "from cuotas " & _
      "where idcredito='" & CLng(IdCredito) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcredito")) Then
      PuedoBorrarCredito = False
      Exit Function
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion PuedoBorrarCredito"
End Function
Private Sub BorrarCredito()
'no permito borrar creditos vigentes..deben estar bloqueados o finalizados
Dim sql As String
Dim IdCredito As Long
On Error GoTo merror

If Not VerificarSeleccionLista(lvCreditos) Then
   MsgE "No hay creditos seleccionados"
   Exit Sub
End If

IdCredito = CLng(lvCreditos.SelectedItem.SubItems(1))
   
If Not CreditoFinalizado(IdCredito) Then
   MsgE "El credito no esta finalizado, no se puede borrar"
   Exit Sub
End If
   
If Not PuedoBorrarCredito(IdCredito) Then
   If Not MsgP("El credito tiene registros asociados...¿Lo borra igual?") Then Exit Sub
End If

If CuotasImpagas(IdCredito) > 0 Then
   If Not MsgP("El credito seleccionado tiene cuotas pendientes...¿Lo borra igual?") Then Exit Sub
Else
   If Not MsgP("¿Confirma el borrado del credito seleccionado?") Then Exit Sub
End If

'otras validaciones
If Not ExisteCredito(IdCredito) Then
   MsgE " El credito no existe"
   Exit Sub
End If

'inicio de la transaccion
cnSQL.BeginTrans
   
'borro el credito
sql = "delete from creditos WHERE (creditos.idcredito)='" & CLng(IdCredito) & "'"
cnSQL.Execute sql

'borro sus cuotas
sql = "delete from cuotas WHERE (cuotas.idcredito)='" & CLng(IdCredito) & "'"
cnSQL.Execute sql

'borro las liquidaciones a cobradores de ese credito
sql = "delete from cobradorespagos WHERE (idcredito)='" & CLng(IdCredito) & "'"
cnSQL.Execute sql

'borro SUS INGRESOS
sql = "delete from ingresos WHERE (idcredito)='" & CLng(IdCredito) & "'"
cnSQL.Execute sql

'borro sus excedentes
sql = "delete from excedentesclientes WHERE (idcredito)='" & CLng(IdCredito) & "'"
cnSQL.Execute sql

'fin de la transaccion
cnSQL.CommitTrans

Call ActualizarListas

MsgI ("El credito fue borrado exitosamente!")

Exit Sub
merror:
tratarerrores "Error borrando creditos"
End Sub
Private Sub BloquearCredito()
'bloquea o desbloquea un credito..solo puede bloquear los vigentes no los finalizados
Dim sql As String
Dim IdCredito As Long
Dim Cadena As String
Dim Mensaje As String
On Error GoTo merror

Cadena = ""
Mensaje = ""

If Not VerificarSeleccionLista(lvCreditos) Then Exit Sub
   
IdCredito = CLng(lvCreditos.SelectedItem.SubItems(1))

If CmdBloquear.Caption = "Bloqu&ear credito" Then
   Cadena = "bloqueando creditos"
Else
   Cadena = "Desbloqueando creditos"
End If
   
'si ya esta bloqueado
If CreditoBloqueado1(IdCredito) Then
   If Not MsgP("El credito esta bloqueado..¿lo desbloquea?") Then
      Exit Sub
   Else
      'lo desbloqueo
      'otras validaciones
      If Not ExisteCredito(IdCredito) Then
         MsgE "El credito no existe"
         Exit Sub
      End If
      
      'desbloqueo de nuevo
      'inicio de transaccion
      cnSQL.BeginTrans
      
      sql = "update creditos set fechabloqueo=null " & _
      "where idcredito='" & CLng(IdCredito) & "'"
      cnSQL.Execute sql
      
      'fin de transaccion
      cnSQL.CommitTrans
      
      Call ActualizarListas
      MsgI "El credito fue desbloqueado"
      Exit Sub
   End If
End If
'si paso es porque no estaba bloqueado

'no dejo bloquear si esta finalizado
If CreditoFinalizado(IdCredito) Then
   MsgE "El credito esta finalizado no se puede bloquear"
   Exit Sub
End If
   
If Not MsgP("¿Confirma el bloqueo del credito seleccionado?") Then Exit Sub
  
'otras validaciones
If Not ExisteCredito(IdCredito) Then
   MsgE "El credito no existe"
   Exit Sub
End If
   
'inicio de la transaccion
cnSQL.BeginTrans
   
sql = "UPDATE creditos SET creditos.fechabloqueo= '" & ConvertirFechaSql(Mid$(Now, 1, 10), "DD/MM/YYYY") & "'" & _
      "WHERE (creditos.idcredito)='" & CLng(IdCredito) & "'"
 
cnSQL.Execute sql

'fin de la transaccion
cnSQL.CommitTrans
   
MsgI "El credito fue bloqueado exitosamente"

'actualizo las listas
Call ActualizarListas

Exit Sub
merror:
tratarerrores "Error " & Cadena
End Sub
Private Sub CmdFinalizar_Click()
CmdFinalizar.Enabled = False
Call FinalizarCredito
If lvCreditos.ListItems.Count > 0 Then
   CmdFinalizar.Enabled = True
End If
End Sub
Private Sub FinalizarCredito()
Dim sql As String
Dim IdCredito As Long
On Error GoTo merror

If Not VerificarSeleccionLista(lvCreditos) Then Exit Sub

IdCredito = CLng(lvCreditos.SelectedItem.SubItems(1))

If CreditoFinalizado(IdCredito) Then
   If CreditoTieneRefinanciadas(IdCredito) Then
      MsgE "El credito tiene cuotas refinanciadas..no se puede poner en vigencia nuevamente"
      Exit Sub
   End If
      
   If Not MsgP("El credito esta finalizado..¿lo pone en vigencia nuevamente?") Then
      Exit Sub
   Else
      
      'otras validaciones
      If Not ExisteCredito(IdCredito) Then
         MsgE "El credito no existe"
         Exit Sub
      End If
   
      'ponerlo en vigencia de nuevo
      'inicio de transaccion
      cnSQL.BeginTrans
      
      sql = "update creditos set fechafinalizacion=null " & _
      "where idcredito='" & CLng(IdCredito) & "'"
      cnSQL.Execute sql
      
      'fin de transaccion
      cnSQL.CommitTrans
      
      Call ActualizarListas
      MsgI "El credito fue puesto en vigencia nuevamente"
      Exit Sub
   End If
End If

If CreditoBloqueado1(IdCredito) Then
   MsgE "El credito esta bloqueado..no se puede finalizar"
   Exit Sub
End If
   
If CuotasImpagas(IdCredito) > 0 Then
   If Not MsgP("El credito seleccionado tiene cuotas pendientes...¿Lo finaliza igual?") Then Exit Sub
Else
   If Not MsgP("¿Confirma la finalizacion del credito seleccionado?") Then Exit Sub
End If

'otras validaciones
If Not ExisteCredito(IdCredito) Then
   MsgE "El credito no existe"
   Exit Sub
End If
   
'inicio de transaccion
cnSQL.BeginTrans

'aca grabo con fecha de finalizacion igual al dia de la fecha
sql = "update creditos set fechafinalizacion='" & ConvertirFechaSql(Mid$(Now, 1, 10), "DD/MM/YYYY") & "' " & _
      "where creditos.idcredito='" & CLng(IdCredito) & "'"
cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

Call ActualizarListas

MsgI "El credito fue finalizado"

Exit Sub
merror:
tratarerrores "Error finalizando credito"
End Sub
Private Sub CmdComodin_Click()
'marca o desmarca una cuota comodin
Dim sql As String
Dim IdCredito As Long
Dim NumCuota As Long
Dim DesmarcarComodin As Boolean
On Error GoTo merror

If Not VerificarSeleccionLista(lvCreditos) Then Exit Sub
   
If Not VerificarSeleccionLista(lvcuotas) Then Exit Sub
   
IdCredito = CLng(lvcuotas.SelectedItem.SubItems(2))
NumCuota = CLng(lvcuotas.SelectedItem.SubItems(3))


If CuotaCobrada(IdCredito, NumCuota) Then
   MsgE "La cuota ya esta cobrada"
   Exit Sub
End If

'no permito comodin en cuotas que tienen cobros parciales
If ObtenerImporteParcialX(IdCredito, NumCuota) > 0 Then
   MsgE "La cuota tiene cobros parciales...no se puede cambiar a comodin"
   Exit Sub
End If
   
'no permite comodin en creditos bloqueados,finalizados o cuotas cobradas
If CreditoFinalizado(IdCredito) Then
   MsgE "El credito al que pertenece la cuota esta finalizado"
   Exit Sub
End If

If CreditoBloqueado1(IdCredito) Then
   MsgE "El credito al que pertenece la cuota esta bloqueado"
   Exit Sub
End If
   
'valido si es vigente
If CuotaRefinanciada(IdCredito, NumCuota) Then
   MsgE "La cuota esta refinanciada (no esta vigente)"
   Exit Sub
End If

DesmarcarComodin = False

If CuotaEsComodin(IdCredito, NumCuota) Then
   If MsgP("La cuota ya esta registrada como comodin. ¿La desmarca?") Then
      'desmarcar
      DesmarcarComodin = True
   Else
      Exit Sub
   End If
End If
      
'agrega el comodin
If Not DesmarcarComodin Then
   If CantidadCuotasComodin(IdCredito) > 0 Then
      MsgE "El credito seleccionado ya tiene una cuota comodin"
      Exit Sub
   End If
      
   If Not MsgP("¿Confirma la cuota seleccionada como cuota comodin?") Then Exit Sub
   
   'otras validaciones
   If Not ExisteCredito(IdCredito) Then
      MsgE "El credito no existe"
      Exit Sub
   End If
   
   'no permite comodin en creditos bloqueados,finalizados o cuotas cobradas
   If CreditoFinalizado(IdCredito) Then
      MsgE "El credito al que pertenece la cuota esta finalizado"
      Exit Sub
   End If
   
   If CreditoBloqueado1(IdCredito) Then
      MsgE "El credito esta bloqueado"
      Exit Sub
   End If
   
   If CantidadCuotasComodin(IdCredito) > 0 Then
      MsgE "El credito seleccionado ya tiene una cuota comodin"
      Exit Sub
   End If
   
   'no permito comodin en cuotas que tienen cobros parciales
   If ObtenerImporteParcialX(IdCredito, NumCuota) > 0 Then
      MsgE "La cuota tiene cobros parciales...no se puede cambiar a comodin"
      Exit Sub
   End If
   
   If CuotaCobrada(IdCredito, NumCuota) Then
      MsgE "La cuota ya esta cobrada"
      Exit Sub
   End If

   'valido si es vigente
   If CuotaRefinanciada(IdCredito, NumCuota) Then
      MsgE "La cuota esta refinanciada (no esta vigente)"
      Exit Sub
   End If

   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update cuotas set cuotacomodin= 1 " & _
         "where cuotas.idcredito=" & CLng(IdCredito) & " and cuotas.numcuota=" & CLng(NumCuota)
   cnSQL.Execute sql
   
   'marca el uso de comodin en el credito
   Call CreditoComodin(IdCredito, 1)
   
   'fin transaccion
   cnSQL.CommitTrans
   
   'pongo la cuota en verde
   lvcuotas.SelectedItem.ForeColor = &H8000&
   lvcuotas.SelectedItem.ListSubItems(1).ForeColor = &H8000&
   lvcuotas.SelectedItem.ListSubItems(2).ForeColor = &H8000&
   lvcuotas.SelectedItem.ListSubItems(3).ForeColor = &H8000&
   lvcuotas.SelectedItem.ListSubItems(4).ForeColor = &H8000&
   lvcuotas.SelectedItem.ListSubItems(5).ForeColor = &H8000&
   
   MsgI "Se registro la cuota comodin exitosamente"
Else
   'saco el comodin solo si es un credito vigente o bloqueado
   'si es finalizado no
   
   'otras validaciones
   If Not ExisteCredito(IdCredito) Then
      MsgE "El credito no existe"
      Exit Sub
   End If
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update cuotas set cuotacomodin=0 " & _
         "where cuotas.idcredito=" & CLng(IdCredito) & " and cuotas.numcuota=" & CLng(NumCuota)
   cnSQL.Execute sql
   
   'desmarco el uso de comodin en el credito
   Call CreditoComodin(IdCredito, 0)
   
   'fin de transaccion
   cnSQL.CommitTrans
   
   'si es menor al segundo vencimiento
   'pongo la cuota en azul si esta al dia
   If CDate(DTPicker5.Value) <= CDate(lvcuotas.SelectedItem.ListSubItems(8)) Then
      lvcuotas.SelectedItem.ForeColor = &HFF0000
      lvcuotas.SelectedItem.ListSubItems(1).ForeColor = &HFF0000
      lvcuotas.SelectedItem.ListSubItems(2).ForeColor = &HFF0000
      lvcuotas.SelectedItem.ListSubItems(3).ForeColor = &HFF0000
      lvcuotas.SelectedItem.ListSubItems(4).ForeColor = &HFF0000
      lvcuotas.SelectedItem.ListSubItems(5).ForeColor = &HFF0000
   Else
      'pongo la cuota en rojo si la cuota esta en mora
      lvcuotas.SelectedItem.ForeColor = vbRed
      lvcuotas.SelectedItem.ListSubItems(1).ForeColor = vbRed
      lvcuotas.SelectedItem.ListSubItems(2).ForeColor = vbRed
      lvcuotas.SelectedItem.ListSubItems(3).ForeColor = vbRed
      lvcuotas.SelectedItem.ListSubItems(4).ForeColor = vbRed
      lvcuotas.SelectedItem.ListSubItems(5).ForeColor = vbRed
   End If
   
   MsgI "Se desmarco la cuota comodin exitosamente"
End If

'actualizo las cuotas
Call CmdActualizar_Click

Exit Sub
merror:
tratarerrores "Error registrando cuota comodin"
End Sub
Private Function DatosImpresionCreditosOk() As Boolean
DatosImpresionCreditosOk = True

If lvCreditos.ListItems.Count() = 0 Then
   DatosImpresionCreditosOk = False
   MsgE "No hay creditos para imprimir"
   Exit Function
End If

End Function
Private Sub CmdLista_Click()
'imprime lista de cuotas por meses
If Not DatosCuponOk() Then Exit Sub

CmdImprimirLista.Enabled = False
If DatosImpresionCuotasOk() Then
   'que imprima en listado
   Call ImprimirListaCuotas
End If
CmdImprimirLista.Enabled = True
End Sub
Private Sub cmdimprimirlista_Click()
'imprime la lista de creditos
If DatosImpresionCreditosOk() Then
   Call ImprimirCreditos
End If
End Sub
Private Sub ImprimirCreditos()
'imprime la lista de creditos
Dim rec As rdoResultset
Dim Archivo As String
Dim Mreporte As New ARCreditosNuevo
On Error GoTo merror

Set rec = CargarRecCreditos()

If Not rec.EOF Then
   Mreporte.RDODataControl1.Resultset = rec
   
   Mreporte.Caption = "Imprimir la lista de creditos"
   
   'si imprimo los datos de empresa
   Mreporte.LabelEmpresa = VG_EMPRESA & vbNullString
   
   If ComboOpciones.Text = "Por cliente" Then
      Mreporte.LabelTitulo = "Lista de creditos vigentes del cliente " & UCase(TxtCliente.Text) & " a la fecha: " & CStr(CDate(DTPicker5.Value))
   End If
   If ComboOpciones.Text = "Creditos vigentes" Then
      Mreporte.LabelTitulo = "Lista de creditos vigentes al: " & CStr(CDate(DTPicker5.Value))
   End If
   If ComboOpciones.Text = "Creditos finalizados" Then
      Mreporte.LabelTitulo = "Lista de creditos finalizados entre el: " & CStr(DTPicker1.Value) & " y el " & CStr(DTPicker2.Value)
   End If
   If ComboOpciones.Text = "Creditos bloqueados" Then
      Mreporte.LabelTitulo = "Lista de creditos bloqueados entre el: " & CStr(DTPicker1.Value) & " y el " & CStr(DTPicker2.Value)
   End If
   
   Mreporte.Show vbModal
Else
   MsgE "No hay creditos para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo la lista de creditos"
End Sub
Private Sub cmdimprimircuotas_Click()
'imprime las cuotas de un credito
If Not DatosCuponOk() Then Exit Sub
CmdImprimirCuotas.Enabled = False
If DatosImpresionCuotasOk() Then
   'imprimo las facturas
   Call ImprimirCuotas(1, DTPicker5.Value)
End If
CmdImprimirCuotas.Enabled = True
End Sub
Private Sub CmdImprimirCupones_Click()
'imprime las cuotas de un credito
If Not DatosCuponOk() Then Exit Sub

CmdImprimirCupones.Enabled = False
If DatosImpresionCuotasOk() Then
   'imprimo los comprobantes con codigo de barras
   Call ImprimirCuotas(2, DTPicker5.Value)
End If
CmdImprimirCupones.Enabled = True
End Sub
Private Function DatosImpresionCuotasOk() As Boolean
Dim CantCuotas As Long
On Error GoTo merror

DatosImpresionCuotasOk = True

If Not VerificarSeleccionLista(lvCreditos) Then
   DatosImpresionCuotasOk = False
   MsgE "No hay creditos seleccionados"
   Exit Function
End If

If lvcuotas.ListItems.Count() = 0 Then
   DatosImpresionCuotasOk = False
   MsgE "No hay cuotas para imprimir"
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosImpresionCuotasOk-ConsultarCreditos"
End Function
Private Sub ImprimirCuotas(ByVal Indicador As Long, ByVal Fecha As Date)
'imprime las cuotas de un documento seleccionado
'si el indicador es 1 imprime factura..sino imprime cupones con codigo de barras
Dim rec As rdoResultset
Dim Archivo As String
On Error GoTo merror

If Indicador = 3 Then
    If CargarCuotasPend() Then
        Set rec = CargarRecCuotasTempCreditos()
    End If
Else
    Set rec = CargarRecCuotasCreditos()
End If

If Not rec.EOF Then
   'si imprimo facturas
   If Indicador = 1 Then
      If VG_MODELOFACTURA1 Then
         Dim Mreporte1 As New ARCuotasCredito4
         Mreporte1.RDODataControl1.Resultset = rec
         Mreporte1.Caption = "Imprimir cuotas de creditos"
         'poner lo de las copias parametrizadas
         Mreporte1.FieldFecha.Text = Fecha
         Mreporte1.Printer.Copies = VG_NUMCOPIAS
         Mreporte1.PageSettings.LeftMargin = VG_LEFT
         Mreporte1.PageSettings.TopMargin = VG_TOP
         Mreporte1.PageSettings.TopMargin = VG_BOTOM
         Mreporte1.Show vbModal
      End If
   Else
      If Indicador = 3 Then
          Dim Mreporte5 As New ARCuotasCreditoUnif
          Mreporte5.RDODataControl1.Resultset = rec
          Mreporte5.Caption = "Imprimir cuotas de creditos unificados"
          'si imprimo los datos de empresa
          Mreporte5.LabelEmpresa = VG_EMPRESA & vbNullString
          'cargo datos de la empresa en los labels
          Mreporte5.LabelCuit = VG_CUIT
          Mreporte5.LabelIngresosBrutos = VG_INGRESOSBRUTOS
          Mreporte5.LabelIva = VG_IVA
          Mreporte5.LabelHorarioAtencion = VG_HORARIOATENCION
          Mreporte5.LabelLugaresPago = VG_LUGARESPAGO
          
          'esto ahora es parametrizado
          'poner lo de las copias parametrizadas
          Mreporte5.FieldFecha.Text = Fecha
         
          Mreporte5.Printer.Copies = VG_NUMCOPIAS
          Mreporte5.Show vbModal
      Else
      'sino imprime cupones viejos de codigo de barras
       If VG_MODELOFACTURA2 Then
          Dim Mreporte2 As New ARCuotasCredito
          Mreporte2.RDODataControl1.Resultset = rec
          Mreporte2.Caption = "Imprimir cuotas de creditos"
          'si imprimo los datos de empresa
          Mreporte2.LabelEmpresa = VG_EMPRESA & vbNullString
          'cargo datos de la empresa en los labels
          Mreporte2.LabelCuit = VG_CUIT
          Mreporte2.LabelIngresosBrutos = VG_INGRESOSBRUTOS
          Mreporte2.LabelIva = VG_IVA
          Mreporte2.LabelHorarioAtencion = VG_HORARIOATENCION
          Mreporte2.LabelLugaresPago = VG_LUGARESPAGO
          'esto ahora es parametrizado
          'poner lo de las copias parametrizadas
          Mreporte2.FieldFecha.Text = Fecha
          
          Mreporte2.Printer.Copies = VG_NUMCOPIAS
          Mreporte2.Show vbModal
       End If
    
       If VG_MODELOFACTURA3 Then
          Dim Mreporte3 As New ARCuotasCredito2
          Mreporte3.RDODataControl1.Resultset = rec
          Mreporte3.Caption = "Imprimir cuotas de creditos"
          'si imprimo los datos de empresa
          Mreporte3.LabelEmpresa = VG_EMPRESA & vbNullString
          'cargo datos de la empresa en los labels
          Mreporte3.LabelCuit = VG_CUIT
          Mreporte3.LabelIngresosBrutos = VG_INGRESOSBRUTOS
          Mreporte3.LabelIva = VG_IVA
          Mreporte3.LabelHorarioAtencion = VG_HORARIOATENCION
          Mreporte3.LabelLugaresPago = VG_LUGARESPAGO
          
          'esto ahora es parametrizado
          'poner lo de las copias parametrizadas
          Mreporte3.FieldFecha.Text = Fecha
         
          Mreporte3.Printer.Copies = VG_NUMCOPIAS
          Mreporte3.Show vbModal
       End If
    
       If VG_MODELOFACTURA4 Then
          Dim mreporte4 As New ARCuotasCredito3
          mreporte4.RDODataControl1.Resultset = rec
          mreporte4.Caption = "Imprimir cuotas de creditos"
          'si imprimo los datos de empresa
          mreporte4.LabelEmpresa = VG_EMPRESA & vbNullString
          'cargo datos de la empresa en los labels
          mreporte4.LabelCuit = VG_CUIT
          mreporte4.LabelIngresosBrutos = VG_INGRESOSBRUTOS
          mreporte4.LabelIva = VG_IVA
          mreporte4.LabelHorarioAtencion = VG_HORARIOATENCION
          mreporte4.LabelLugaresPago = VG_LUGARESPAGO
             
          'esto ahora es parametrizado
          'poner lo de las copias parametrizadas
          mreporte4.FieldFecha.Text = Fecha
         
          mreporte4.Printer.Copies = VG_NUMCOPIAS
          mreporte4.Show vbModal
       End If
     End If
   End If 'del indicador
   
Else
   MsgE "No hay cuotas para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo cuotas de credito"
End Sub
Private Sub ImprimirListaCuotas()
'imprime las cuotas de un documento seleccionado
Dim rec As rdoResultset
Dim Archivo As String
Dim Mreporte1 As New ARListaCuotasCredito
On Error GoTo merror

Set rec = CargarRecCuotasCreditos()

If Not rec.EOF Then
   Mreporte1.RDODataControl1.Resultset = rec
   Mreporte1.Caption = "Imprimir cuotas de creditos"
   If ComboOpcionesCuotas.Text = "Sin filtro" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value) & " agrupado por meses"
   End If
   If ComboOpcionesCuotas.Text = "Cobradas" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas cobradas desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value) & " agrupado por meses"
   End If
   If ComboOpcionesCuotas.Text = "Cobradas parcialmente" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas cobradas parcialmente desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value) & " agrupado por meses"
   End If
   
   If ComboOpcionesCuotas.Text = "Pendientes" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas pendientes desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value) & " agrupado por meses"
   End If
   If ComboOpcionesCuotas.Text = "En mora" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas en mora desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value) & " agrupado por meses"
   End If
   If ComboOpcionesCuotas.Text = "Refinanciadas" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas refinanciadas desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value) & " agrupado por meses"
   End If
   
   Mreporte1.FieldFecha = DTPicker5.Value
      
   'si imprimo los datos de empresa
    Mreporte1.LabelEmpresa = VG_EMPRESA & vbNullString
    'la agrupacion del reporte
    Mreporte1.GroupHeader1.DataField = "periodo"
    
   Mreporte1.Show vbModal
Else
   MsgE "No hay cuotas para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo lista de cuotas de credito"
End Sub
Private Sub ComboOpciones_Click()
TxtTotal.Text = 0

If ComboOpciones.Text = "Por cliente" Then
   FrameConsultaMasiva.Visible = True
   FrameCliente.Visible = True
   CmdBloquear.Caption = "Bloqu&ear credito"
   CmdFinalizar.Caption = "&Finalizar credito"
   CheckBloqueados.Value = 0
   CheckFinalizados.Value = 0
   CheckBloqueados.Enabled = True
   CheckFinalizados.Enabled = True
End If

If ComboOpciones.Text = "Creditos vigentes" Then
   FrameConsultaMasiva.Visible = False
   FrameCliente.Visible = False
   CmdBloquear.Caption = "Bloqu&ear credito"
   CmdFinalizar.Caption = "&Finalizar credito"
   CheckBloqueados.Value = 0
   CheckFinalizados.Value = 0
   CheckBloqueados.Enabled = False
   CheckFinalizados.Enabled = False
End If

If ComboOpciones.Text = "Creditos finalizados" Then
   FrameConsultaMasiva.Visible = True
   FrameCliente.Visible = False
   CmdFinalizar.Caption = "&Restaurar credito"
   CheckBloqueados.Value = 0
   CheckFinalizados.Value = 1
   CheckBloqueados.Enabled = False
   CheckFinalizados.Enabled = False
End If

If ComboOpciones.Text = "Creditos bloqueados" Then
   FrameConsultaMasiva.Visible = True
   FrameCliente.Visible = False
   CmdBloquear.Caption = "&Desbloquear credito"
   CmdFinalizar.Caption = "&Finalizar credito"
   CheckBloqueados.Value = 1
   CheckFinalizados.Value = 0
   CheckBloqueados.Enabled = False
   CheckFinalizados.Enabled = False
End If

If ComboOpciones.Text = "Todos" Then
   FrameConsultaMasiva.Visible = True
   FrameCliente.Visible = False
   CmdBloquear.Caption = "Bloqu&ear credito"
   CmdFinalizar.Caption = "&Finalizar credito"
   CheckBloqueados.Value = 0
   CheckFinalizados.Value = 0
   CheckBloqueados.Enabled = True
   CheckFinalizados.Enabled = True
End If

Call ActualizarListas
End Sub

Private Sub lvcreditos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Orden As Long
On Error GoTo merror
  
If lvCreditos.ListItems.Count > 1 Then
   lvCreditos.SortKey = ColumnHeader.Index - 1
   Orden = lvCreditos.SortKey
   lvCreditos.SortOrder = Abs(Not lvCreditos.SortOrder = 1)
   lvCreditos.Sorted = True
End If

Exit Sub
merror:
tratarerrores "Error ordenando creditos"
End Sub
Private Sub lvcreditos_ItemClick(ByVal Item As MSComctlLib.ListItem)
Call CargarCuotasCreditos
End Sub
Private Sub SetearEntorno()
On Error GoTo merror

If lvCreditos.ListItems.Count = 0 Then
   CmdBorrar.Enabled = False
   CmdBloquear.Enabled = False
   CmdFinalizar.Enabled = False
   CmdImprimirMutuo.Enabled = False
   CmdImprimirResumen.Enabled = False
   CmdActualizar.Enabled = False
   CmdExportarCreditos.Enabled = False
  
   CmdImprimirLista.Enabled = False
   CmdImprimirCuotas.Enabled = False
   FrameFiltro.Enabled = False
   CmdComodin.Enabled = False
   CmdPlanilla.Enabled = False
   CmdExportarCuotas.Enabled = False
   CmdExceptuar.Enabled = False
   CmdCambiarVto.Enabled = False
   CmdRestituirVto.Enabled = False
   CmdExportarSaldos.Enabled = False
   CmdPMC.Enabled = False
   CmdListado.Enabled = False
   CmdSellos.Enabled = False
   CmdLista.Enabled = False
   CmdImprimirCupones.Enabled = False
   CmdAnularRefin.Enabled = False
   CmdHistorial.Enabled = False
   
   TxtTotalCapital.Text = 0
   TxtTotalInteres.Text = 0
   TxtTotalGastos.Text = 0
   TxtTotalSeguros.Text = 0
   TxtTotalImpuestos.Text = 0
   TxtTotalMora.Text = 0
   TxtTotalCobrado.Text = 0
   TxtImporteTotal.Text = 0
Else
   If VG_EXPORTA Then
      CmdExportarCreditos.Enabled = True
      CmdExportarCuotas.Enabled = True
      CmdExportarSaldos.Enabled = True
      CmdPMC.Enabled = True
   End If
   
   'si el tipo de usuario puede borrar,bloquear etc
   If VG_ADMCREDITOS Then
      CmdBorrar.Enabled = True
      CmdBloquear.Enabled = True
      CmdFinalizar.Enabled = True
      CmdPlanilla.Enabled = True
      CmdImprimirMutuo.Enabled = True
      CmdImprimirResumen.Enabled = True
      CmdComodin.Enabled = True
   End If
   
   'si el tipo de usuario puede imprimir cuotas
   If VG_IMPRIMECUOTAS Then
      FrameFiltro.Enabled = True
      CmdImprimirMutuo.Enabled = True
      CmdImprimirResumen.Enabled = True
      CmdPlanilla.Enabled = True
      CmdImprimirLista.Enabled = True
      CmdImprimirCuotas.Enabled = True
      CmdListado.Enabled = True
      CmdSellos.Enabled = True
      CmdLista.Enabled = True
      CmdImprimirCupones.Enabled = True
   End If
   
   CmdActualizar.Enabled = True
   
   If VG_REFINANCIA Then
      CmdExceptuar.Enabled = True
      CmdCambiarVto.Enabled = True
      CmdRestituirVto.Enabled = True
   Else
      CmdExceptuar.Enabled = False
      CmdCambiarVto.Enabled = False
      CmdRestituirVto.Enabled = False
   End If
   
   If VG_REFINANCIA Then
      CmdAnularRefin.Enabled = True
   End If
   
   CmdHistorial.Enabled = True
   
End If

Exit Sub
merror:
tratarerrores "Error seteando entorno-ConsultarCreditos"
End Sub
Private Sub lvcuotas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Orden As Long
On Error GoTo merror
  
If lvcuotas.ListItems.Count > 1 Then
   lvcuotas.SortKey = ColumnHeader.Index - 1
   Orden = lvcuotas.SortKey
   lvcuotas.SortOrder = Abs(Not lvcuotas.SortOrder = 1)
   lvcuotas.Sorted = True
End If

Exit Sub
merror:
tratarerrores "Error ordenando cuotas"
End Sub
Private Sub lvcuotas_DblClick()
'muestra los cobros parciales de las cuotas
Dim IdCredito As Long
Dim CodPrestamo As String
Dim NumCuota As Long
Dim Saldo As Currency
Dim Cliente As Long
Dim NumFactura As Long
On Error GoTo merror

If Not VerificarSeleccionLista(lvCreditos) Then Exit Sub
If Not VerificarSeleccionLista(lvcuotas) Then Exit Sub

CodPrestamo = CStr(lvcuotas.SelectedItem.SubItems(1))
IdCredito = CLng(lvcuotas.SelectedItem.SubItems(2))
Cliente = CLng(lvCreditos.SelectedItem.SubItems(9))
NumCuota = CLng(lvcuotas.SelectedItem.SubItems(3))
Saldo = CCur(lvcuotas.SelectedItem.SubItems(11))
NumFactura = CLng(lvcuotas.SelectedItem.SubItems(4))

If TieneCobrosParciales(IdCredito, NumCuota) Then
   'si la cuota tiene cobros parciales las muestra
   If ObtenerImporteParcialX(IdCredito, NumCuota) > 0 Then
      FrmCobrosParciales.xnumcredito = IdCredito
      FrmCobrosParciales.xcodprestamo = CodPrestamo
      FrmCobrosParciales.xnumcuota = NumCuota
      FrmCobrosParciales.ximporteactualizado = Saldo
      FrmCobrosParciales.xidcliente = Cliente
      FrmCobrosParciales.xfactura = NumFactura
   
      Call CenterForm(FrmCobrosParciales)
      FrmCobrosParciales.Show vbModal
   
     Call CmdActualizar_Click
  End If
End If

Exit Sub
merror:
tratarerrores "Error seleccionando cobros parciales de cuotas"
End Sub
Private Sub LvCuotas_KeyDown(KeyCode As Integer, Shift As Integer)
'si apretaron enter en una cuota
If KeyCode = vbKeyReturn Then
   Call lvcuotas_DblClick
End If
End Sub
Private Sub ActualizarListas()
Call CargarCreditos
Call CargarCuotasCreditos
Call SetearEntorno
If lvCreditos.ListItems.Count > 0 Then
   lvCreditos.SetFocus
End If
End Sub
Private Sub TxtCliente_Change()
'si cambia el cliente
Call ActualizarListas
End Sub
Private Sub CmdBuscar_Click()
'buscar por rango de fechas
Call ActualizarListas
End Sub
Private Sub DTPicker1_Change()
'si cambia la fecha inicial del rango
Call ActualizarListas
End Sub
Private Sub DTPicker2_Change()
'si cambia la fecha final del rango
Call ActualizarListas
End Sub
Private Sub CmdImprimirMutuo_Click()
'imprime el contrato
Dim ImporteCuota As Currency
Dim IdCredito As Long
Dim IdCliente As Long
Dim ImporteTotal As Currency
Dim NumCuotas As Long
Dim Vencimiento As Date
Dim AltaCredito As Date
On Error GoTo merror

If lvCreditos.ListItems.Count = 0 Then Exit Sub
If lvcuotas.ListItems.Count = 0 Then Exit Sub

IdCliente = CLng(lvCreditos.SelectedItem.SubItems(9))
IdCredito = CLng(lvCreditos.SelectedItem.SubItems(1))
ImporteTotal = CCur(lvCreditos.SelectedItem.SubItems(6))
NumCuotas = CLng(lvCreditos.SelectedItem.SubItems(4))

ImporteCuota = CCur(lvCreditos.SelectedItem.SubItems(10))
Vencimiento = CDate(lvCreditos.SelectedItem.SubItems(11))
AltaCredito = CDate(lvCreditos.SelectedItem.SubItems(3))

Call ImprimirAcuerdoMutuo(IdCliente, IdCredito, ImporteTotal, NumCuotas, ImporteCuota, Vencimiento, AltaCredito)

Exit Sub
merror:
tratarerrores "Error imprimiendo acuerdo mutuo"
End Sub
Private Sub ImprimirAcuerdoMutuo(ByVal IdCliente As Long, ByVal IdCredito As Long, ByVal ImporteTotal As Currency, CantidadCuotas As Long, ByVal ImporteCuota As Currency, ByVal VencimientoCuota As Date, ByVal FechaAlta As Date)
'imprime el mutuo acuerdo de partes usada en registrar creditos y consultar
Dim sql As String
Dim rec As rdoResultset
Dim rec2 As rdoResultset
Dim Mreporte As New AcuerdoMutuo
Dim Fechacompleta As String
Dim Archivo As String
Dim Cliente As String
Dim documentocliente As String
Dim cuilcliente As String
Dim nacionalidadcliente As String
Dim profesioncliente As String
Dim domiciliocliente As String
Dim ciudadcliente As String
Dim garante As String
Dim documentogarante As String
Dim cuilgarante As String
Dim nacionalidadgarante As String
Dim profesiongarante As String
Dim domiciliogarante As String
On Error GoTo merror

'obtengo los datos del cliente
sql = "select *,localidades.nombre as localidad, clientes.nombre as Nombre1 " & _
      "from localidades inner join clientes on localidades.idlocalidad=clientes.idlocalidad " & _
      "where idcliente='" & CLng(IdCliente) & "'"
           
Set rec2 = cnSQL.OpenResultset(sql)

If Not rec2.EOF Then
   Cliente = rec2.rdoColumns("apellido") & " " & rec2.rdoColumns("nombre1") & vbNullString
   documentocliente = rec2.rdoColumns("numdocumento") & vbNullString
   cuilcliente = rec2.rdoColumns("cuil") & vbNullString
   profesioncliente = rec2.rdoColumns("profesion") & vbNullString
   nacionalidadcliente = rec2.rdoColumns("nacionalidad") & vbNullString
   domiciliocliente = rec2.rdoColumns("domicilio") & vbNullString
   ciudadcliente = rec2.rdoColumns("localidad") & vbNullString
   garante = rec2.rdoColumns("apellidogarante") & rec2.rdoColumns("nombregarante") & vbNullString
   documentogarante = rec2.rdoColumns("documentogarante") & vbNullString
   cuilgarante = rec2.rdoColumns("cuitgarante") & vbNullString
   profesiongarante = rec2.rdoColumns("profesiongarante") & vbNullString
   nacionalidadgarante = rec2.rdoColumns("nacionalidadgarante") & vbNullString
   domiciliogarante = rec2.rdoColumns("domiciliogarante") & vbNullString
End If

'obtengo los datos de las cuota
Set rec = CargarRecCuotas(IdCredito)

If Not rec.EOF Then
   With Mreporte
        .RDODataControl1.Resultset = rec
        'aca establezco los campos iguales para todos
        .rtf.ReplaceField "PESOS", Format(CCur(ImporteTotal), "0.00")
        .rtf.ReplaceField "CUOTAS", CLng(CantidadCuotas)
        .rtf.ReplaceField "VENCIMIENTO", CDate(VencimientoCuota)
        .rtf.ReplaceField "DOMICILIO", domiciliocliente
        .rtf.ReplaceField "CIUDAD", ciudadcliente
        
        'completo los parrafos
        .rtf.ReplaceField "PARRAFO1", VG_TEXTOACUERDOMUTUO1
        .rtf.ReplaceField "PARRAFO2", VG_TEXTOACUERDOMUTUO2
        .rtf.ReplaceField "PARRAFO3", VG_TEXTOACUERDOMUTUO3
        .rtf.ReplaceField "PARRAFO4", VG_TEXTOACUERDOMUTUO4
        .rtf.ReplaceField "PARRAFO5", VG_TEXTOACUERDOMUTUO5 & VG_TEXTOACUERDOMUTUO6 & VG_TEXTOACUERDOMUTUO7
        .rtf.ReplaceField "PARRAFO6", VG_TEXTOACUERDOMUTUO8 & VG_TEXTOACUERDOMUTUO9 & VG_TEXTOACUERDOMUTUO10
        
        'la agrupacion del reporte
        .GroupHeader1.DataField = "idcredito"
        
        'datos del cliente
        .FieldCliente.Text = Cliente
        .FieldDocumento.Text = documentocliente
        .FieldNacionalidad.Text = nacionalidadcliente
        .FieldCuil.Text = cuilcliente
        .FieldDomicilio.Text = domiciliocliente
        .FieldProfesion.Text = profesioncliente
        
        'de garante
        .FieldGarante.Text = garante
        .FieldDocumentoGarante.Text = documentogarante
        .FieldNacionalidadGarante.Text = nacionalidadgarante
        .FieldCuilGarante.Text = cuilgarante
        .FieldDomicilioGarante.Text = domiciliogarante
        .FieldProfesionGarante.Text = profesiongarante
        
        .LabelEncabezado.Caption = VG_EMPRESA & vbNullString
        'establezco la fecha en letras
        'Fechacompleta = FormatearFecha(CDate(DTPicker5.Value))
        Fechacompleta = FormatearFecha(CDate(FechaAlta))
        
        Provincia = ObtenerProvinciaCredito(IdCredito)
        .LabelLugarFecha.Caption = Provincia + " " + Fechacompleta
        .Show (vbModal)
    End With
Else
    MsgI "No hay datos para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo el mutuo acuerdo"
End Sub
Private Function CargarRecCuotas(ByVal IdCredito As Long) As rdoResultset
'usado para el mutuo acuerdo en pantallas de registrar creditos y consultar
Dim sql As String
Dim Condicion As String
Dim DiferenciaVencimientos As Long
On Error GoTo merror

Condicion = "creditos.idcredito='" & CLng(IdCredito) & "'"
   

 
'"(cuotas.importevencimiento1) as importetotal,'" & CDate(DTPicker5.Value) & "' - cdate(cuotas.fechavencimiento1) as totaldiasmora," & _

'consulta detallada impresa en carta reclamo o lista
sql = "SELECT creditos.idcredito,creditos.idcliente,creditos.numcuotas,creditos.codprestamo," & _
      "clientes.apellido + ' ' + clientes.nombre as cliente," & _
      "clientes.domicilio,localidades.codigopostal," & _
      "localidades.nombre as localidad,provincias.nombre as provincia," & _
      "cuotas.fechavencimiento2,cuotas.numcuota,cuotas.cobrosparciales,cuotas.numfactura," & _
      "(cuotas.importevencimiento1) as importetotal,'" & ConvertirFechaSql(CDate(DTPicker5.Value), "DD/MM/YYYY") & "' - cuotas.fechavencimiento1 as totaldiasmora," & _
      "1 as totalcuotasadeudadas " & _
      "FROM provincias inner join (localidades inner join " & _
      "(clientes inner join (creditos inner join cuotas " & _
      "on creditos.idcredito=cuotas.idcredito) " & _
      "on clientes.idcliente=creditos.idcliente) " & _
      "on localidades.idlocalidad=clientes.idlocalidad) " & _
      "on provincias.idprovincia=localidades.idprovincia " & _
      "where " & Condicion & " order by creditos.idcredito,cuotas.numcuota"
     
Set CargarRecCuotas = cnSQL.OpenResultset(sql)

Exit Function
merror:
tratarerrores "Error cargando las cuotas del credito"
End Function
Private Sub CmdListado_Click()
'imprime lista de cuotas ordenada por fecha de vencimiento
If Not DatosCuponOk() Then Exit Sub

CmdListado.Enabled = False
If DatosImpresionCuotasOk() Then
   'que imprima en listado
   Call ImprimirListadoCuotas
End If
CmdListado.Enabled = True
End Sub
Private Sub ImprimirListadoCuotas()
'imprime las cuotas de ordenadas por fecha de vencimiento
Dim rec As rdoResultset
Dim Archivo As String
Dim Mreporte1 As New ARListadoCuotas
On Error GoTo merror

Set rec = CargarRecCuotasCreditos()

If Not rec.EOF Then
   Mreporte1.RDODataControl1.Resultset = rec
   Mreporte1.Caption = "Imprimir listado de cuotas de creditos"
   If ComboOpcionesCuotas.Text = "Sin filtro" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value)
   End If
   If ComboOpcionesCuotas.Text = "Cobradas" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas cobradas desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value)
   End If
   
   If ComboOpcionesCuotas.Text = "Cobradas parcialmente" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas cobradas parcialmente desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value)
   End If
   
   If ComboOpcionesCuotas.Text = "Pendientes" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas pendientes desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value)
   End If
   If ComboOpcionesCuotas.Text = "En mora" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas en mora desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value)
   End If
   If ComboOpcionesCuotas.Text = "Refinanciadas" Then
      Mreporte1.LabelTitulo.Caption = "Lista de cuotas refinanciadas desde el:" & CStr(DTPicker3.Value) & " al " & CStr(DTPicker4.Value)
   End If
   
   Mreporte1.FieldFecha = DTPicker5.Value
      
   'si imprimo los datos de empresa
   Mreporte1.LabelEmpresa = VG_EMPRESA & vbNullString
   'la agrupacion del reporte
   'Mreporte1.GroupHeader1.DataField = "periodo"
   Mreporte1.Show vbModal
Else
   MsgE "No hay cuotas para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo el listado de cuotas de creditos"
End Sub
Private Sub CmdPlanilla_Click()
'imprime la lista de creditos
If DatosImpresionCreditosOk() Then
   Call ImprimirResumenCreditos
End If
End Sub
Private Sub ImprimirResumenCreditos()
'imprime la lista de creditos
Dim rec As rdoResultset
Dim Archivo As String
Dim Mreporte As New ARPlanillaCreditos
On Error GoTo merror

Set rec = CargarRecCreditos()

If Not rec.EOF Then
   Mreporte.RDODataControl1.Resultset = rec
   Mreporte.Caption = "Imprimir el resumen de creditos"
   Mreporte.FieldFechaActual.Text = DTPicker5.Value
   
   'si imprimo los datos de empresa
   Mreporte.LabelEmpresa = VG_EMPRESA & vbNullString
   
   If ComboOpciones.Text = "Por cliente" Then
      Mreporte.LabelTitulo = "Lista de creditos vigentes del cliente " & UCase(TxtCliente.Text) & " a la fecha: " & CStr(CDate(DTPicker5.Value))
   End If
   If ComboOpciones.Text = "Creditos vigentes" Then
      Mreporte.LabelTitulo = "Lista de creditos vigentes al: " & CStr(CDate(DTPicker5.Value))
   End If
   If ComboOpciones.Text = "Creditos finalizados" Then
      Mreporte.LabelTitulo = "Lista de creditos finalizados entre el: " & CStr(DTPicker1.Value) & " y el " & CStr(DTPicker2.Value)
   End If
   If ComboOpciones.Text = "Creditos bloqueados" Then
      Mreporte.LabelTitulo = "Lista de creditos bloqueados entre el: " & CStr(DTPicker1.Value) & " y el " & CStr(DTPicker2.Value)
   End If
   
   Mreporte.Show vbModal
Else
   MsgE "No hay creditos para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo el resumen de creditos"
End Sub
Private Sub CmdExportarCreditos_Click()
CmdExportarCreditos.Enabled = False
Me.MousePointer = vbHourglass
Call ExportarCreditos
Me.MousePointer = vbDefault
CmdExportarCreditos.Enabled = True
End Sub
Private Sub ExportarCreditos()
Dim MiExcel As Excel.APPLICATION ' Aplicacion Excel
Dim MiLibro As Excel.Workbook ' Un libro de Excel
Dim MiHoja As Excel.Worksheet ' Una hoja de Excel
Dim Filas As Long
Dim FilaTitulos As Long
Dim Archivo As String
Dim Archi As String
Dim rec As rdoResultset
Dim sql As String
Dim SaldoCredito As Currency
Dim SaldoCapital As Currency
Dim CuotasCobradas As Long
Dim CuotasVencidas As Long
Dim CuotasPendientes As Long
Dim Fecha As Date
Dim FechaPagoParcial As Date
Dim Mensaje As String
Dim Dia As String
Dim Mes As String
Dim Ano As String
Dim CondicionZ As String
On Error GoTo merror

If lvCreditos.ListItems.Count() = 0 Then Exit Sub

Dia = Format(CStr(Day(CDate(DTPicker5.Value))), "00")
Mes = Format(CStr(Month(CDate(DTPicker5.Value))), "00")
Ano = Format(CStr(Year(CDate(DTPicker5.Value))), "0000")

Archi = "Creditos"

Archi = Archi + "-" + Dia + "-" + Mes + "-" + Ano + ".xls"

If Not MsgP("¿Confirma la exportacion del listado de Creditos hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & "?") Then Exit Sub

'inicio transaccion
cnSQL.BeginTrans

'estas lineas estaban arriba pero las puse aca para que entren en la transaccion
'si no existe la carpeta la crea
If Not ExisteCarpeta() Then Exit Sub

'verifico si ya existe en la ubicacion
Archivo = Dir("c:\exportacionexcel\" & Archi)

If Trim(Archivo) <> "" Then
  'si ya existe la planilla lo borro para que despues no haya errores con la pantallita
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
MiHoja.Cells(1, 1).Value = "Lista de " & ComboOpciones.Text & " a la fecha:" & CStr(CDate(DTPicker5.Value))

'establezco el ancho de todas las columnas
MiHoja.Columns.ColumnWidth = 17

FilaTitulos = 2

MiHoja.Cells(FilaTitulos, 1).Value = "Nro.Cliente"

MiHoja.Cells(FilaTitulos, 2).Value = "DNI"
MiHoja.Cells(FilaTitulos, 3).Value = "Fecha.Nac."

'pongo los titulos de las columnas..puede ser texto o textbox
MiHoja.Cells(FilaTitulos, 4).Value = "Nro.Prestamo"
MiHoja.Cells(FilaTitulos, 5).Value = "Fecha alta"
MiHoja.Cells(FilaTitulos, 6).Value = "Provincia"
MiHoja.Cells(FilaTitulos, 7).Value = "Comercio"
MiHoja.Cells(FilaTitulos, 8).Value = "Nombre y Apellido"
MiHoja.Cells(FilaTitulos, 9).Value = "Capital Otorgado"
MiHoja.Cells(FilaTitulos, 10).Value = "Tasa"
MiHoja.Cells(FilaTitulos, 11).Value = "Cantidad Cuotas"
MiHoja.Cells(FilaTitulos, 12).Value = "Cuota Mensual (*)"
MiHoja.Cells(FilaTitulos, 13).Value = "Total Prestamo"
MiHoja.Cells(FilaTitulos, 14).Value = "1º Vto"
MiHoja.Cells(FilaTitulos, 15).Value = "Saldo Prestamo"
MiHoja.Cells(FilaTitulos, 16).Value = "Cuotas Cobradas"
MiHoja.Cells(FilaTitulos, 17).Value = "Cuotas Vencidas"
MiHoja.Cells(FilaTitulos, 18).Value = "Cuotas Pendientes"
MiHoja.Cells(FilaTitulos, 19).Value = "Fecha Ultimo Pago"
MiHoja.Cells(FilaTitulos, 20).Value = "Total cobrado"
MiHoja.Cells(FilaTitulos, 21).Value = "IVA Exento"
MiHoja.Cells(FilaTitulos, 22).Value = "Observaciones"
MiHoja.Cells(FilaTitulos, 23).Value = "Vendedor"
MiHoja.Cells(FilaTitulos, 24).Value = "Usuario"


'pongo los titulos en negritas
MiHoja.Range("a1:x2").Font.Bold = True

'cargo el registro de creditos segun las condiciones de la pantalla
Set rec = CargarRecCreditos()

'comienzo en la fila 3 porque las anteriores son de titulos
Filas = 3
If Not rec.EOF Then
   Do While Not rec.EOF
      MiHoja.Cells(Filas, 1).Value = rec.rdoColumns("numlegajo")
      
      MiHoja.Cells(Filas, 2).Value = rec.rdoColumns("numdocumento")
      MiHoja.Cells(Filas, 3).Value = CDate(rec.rdoColumns("nacimiento"))
      
      'grabo cada usuario en la planilla
      MiHoja.Cells(Filas, 4).Value = rec.rdoColumns("codprestamo")
      
      'nuevos alta y provincia
      MiHoja.Cells(Filas, 5).Value = CDate(rec.rdoColumns("fechacredito"))
      MiHoja.Cells(Filas, 6).Value = rec.rdoColumns("provincia")
      
      MiHoja.Cells(Filas, 7).Value = rec.rdoColumns("motivobloqueo") & vbNullString
      
      MiHoja.Cells(Filas, 8).Value = rec.rdoColumns("cliente")
      MiHoja.Cells(Filas, 9).Value = Format(rec.rdoColumns("importeafinanciar"), "$0.00")
      MiHoja.Cells(Filas, 10).Value = rec.rdoColumns("tasa")
      MiHoja.Cells(Filas, 11).Value = rec.rdoColumns("numcuotas")
      MiHoja.Cells(Filas, 12).Value = Format(rec.rdoColumns("importecuota"), "$0.00")
      
      MiHoja.Cells(Filas, 13).Value = Format(ObtenerTotalCredito(rec.rdoColumns("idcredito")), "$0.00")
      
      'MiHoja.Cells(Filas, 13).Value = Format(rec.rdoColumns("total"), "$0.00")
      
      MiHoja.Cells(Filas, 14).Value = CDate(rec.rdoColumns("fechavencimiento1"))
      
      'cargo mas datos
      SaldoCredito = ObtenerSaldoCredito(rec.rdoColumns("idcredito"), DTPicker5.Value)
      MiHoja.Cells(Filas, 15).Value = SaldoCredito
      
      CuotasCobradas = ObtenerCuotasCobradas(rec.rdoColumns("idcredito"))
      MiHoja.Cells(Filas, 16).Value = CuotasCobradas
      
      CuotasVencidas = ObtenerCuotasVencidas(rec.rdoColumns("idcredito"), DTPicker5.Value)
      
      MiHoja.Cells(Filas, 17).Value = CuotasVencidas
      
      CuotasPendientes = ObtenerCuotasPendientes(rec.rdoColumns("idcredito"), DTPicker5.Value)
      MiHoja.Cells(Filas, 18).Value = CuotasPendientes
      
      Fecha = ObtenerUltimaFechaCobro(rec.rdoColumns("idcredito"))
      FechaPagoParcial = ObtenerUltimaFechaCobroParcial(rec.rdoColumns("idcredito"))
      If CDate(Fecha) = Date + 1 Then
         If CDate(FechaPagoParcial) = Date + 1 Then
            MiHoja.Cells(Filas, 19).Value = ""
         Else
            MiHoja.Cells(Filas, 19).Value = CDate(FechaPagoParcial) & " (*)"
         End If
      Else
         If CDate(FechaPagoParcial) = Date + 1 Then
            MiHoja.Cells(Filas, 19).Value = "'" & CDate(Fecha)
         Else
            If CDate(FechaPagoParcial) > CDate(Fecha) Then
                MiHoja.Cells(Filas, 19).Value = CDate(FechaPagoParcial) & "(*)"
            Else
                MiHoja.Cells(Filas, 19).Value = "'" & CDate(Fecha)
            End If
         End If
      End If
      
      ImporteParcial = ObtenerImporteParcial(rec.rdoColumns("idcredito"))
      
      MiHoja.Cells(Filas, 20).Value = ImporteParcial
      
      If Not IsNull(rec.rdoColumns("codprestamo")) Then
        If Mid(rec.rdoColumns("codprestamo"), 7, 1) = "C" Then
          MiHoja.Cells(Filas, 21).Value = "SI"
        Else
          MiHoja.Cells(Filas, 21).Value = "NO"
        End If
      End If
      
      MiHoja.Cells(Filas, 22).Value = rec.rdoColumns("observaciones") + "."
      
      'nuevo vendedor
      MiHoja.Cells(Filas, 23).Value = rec.rdoColumns("cad2") & "." & vbNullString
      
      MiHoja.Cells(Filas, 24).Value = rec.rdoColumns("cad1") & vbNullString

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
   
   Mensaje = "Se exporto el listado de creditos a la planilla C:\ExportacionExcel\" & Archi
Else
   'grabo los cambios en una tabla falsa
   MiLibro.SaveAs ("c:\ExportacionExcel\temporal.xls")
   'cierro el libro
   MiLibro.Close
   'salgo de excel sin grabar los cambios en el archivo adecuado..uso otro
   MiExcel.Quit
   'borro la tabla falsa
   Set MiExcel = Nothing
   
   Kill ("c:\ExportacionExcel\temporal.xls")
  
   Mensaje = "No hay datos para exportar"
End If
   
'finalizo la transaccion
cnSQL.CommitTrans

MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error Exportando el listado de creditos...verifique que los archivos de Excel esten cerrados"
End Sub
Private Sub CmdExportarCuotas_Click()
If Not DatosCuponOk() Then Exit Sub

CmdExportarCuotas.Enabled = False
Me.MousePointer = vbHourglass
Call ExportarCuotas
Me.MousePointer = vbDefault
CmdExportarCuotas.Enabled = True
End Sub
Private Sub ExportarCuotas()
Dim MiExcel As Excel.APPLICATION ' Aplicacion Excel
Dim MiLibro As Excel.Workbook ' Un libro de Excel
Dim MiHoja As Excel.Worksheet ' Una hoja de Excel
Dim Filas As Long
Dim FilaTitulos As Long
Dim Archivo As String
Dim Archi As String
Dim rec As rdoResultset
Dim sql As String
Dim SaldoCredito As Currency
Dim SaldoCapital As Currency
Dim CuotasCobradas As Long
Dim CuotasVencidas As Long
Dim CuotasPendientes As Long
Dim Fecha As Date
Dim SaldoCuota As Currency
Dim SaldoCuota1erVenc As Currency
Dim Importe1erVenc As Currency
Dim ImporteParcial As Currency
Dim Mensaje As String
Dim Cadena As String
Dim TotalIva As Currency
Dim TotalCuota As Currency
Dim TotalCobrado As Currency
Dim Dia As String
Dim Mes As String
Dim Ano As String
Dim RecargoCuota As Currency
Dim ImporteMora As Currency
Dim IvaMora As Currency
Dim IvaMora2 As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim IvaACobrarDevuelto  As Currency
On Error GoTo merror

If lvCreditos.ListItems.Count() = 0 Then Exit Sub

If lvcuotas.ListItems.Count() = 0 Then Exit Sub

Dia = Format(CStr(Day(CDate(DTPicker5.Value))), "00")
Mes = Format(CStr(Month(CDate(DTPicker5.Value))), "00")
Ano = Format(CStr(Year(CDate(DTPicker5.Value))), "0000")

If ComboOpcionesCuotas.Text = "Todas" Then
   Archi = "Cuotas"
End If

If ComboOpcionesCuotas.Text = "Todas por fechas" Then
   Archi = "Cuotas"
End If

If ComboOpcionesCuotas.Text = "Pendientes" Then
   Archi = "CuotasPendientes"
End If

If ComboOpcionesCuotas.Text = "Cobradas" Then
   Archi = "CuotasCobradas"
End If

If ComboOpcionesCuotas.Text = "Cobradas parcialmente" Then
   Archi = "CuotasParciales"
End If

If ComboOpcionesCuotas.Text = "En mora" Then
   Archi = "CuotasEnMora"
End If

If ComboOpcionesCuotas.Text = "Financiadas" Then
   Archi = "CuotasFinanciadas"
End If

If ComboOpcionesCuotas.Text = "Por cupon" Then
   Archi = "CuotasPorCupon"
End If

Archi = Archi + "-" + Dia + "-" + Mes + "-" + Ano + ".xls"

If Not MsgP("¿Confirma la exportacion de las cuotas hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & "?") Then Exit Sub

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

If ComboOpcionesCuotas.Text = "Todas" Or ComboOpcionesCuotas.Text = "Todas por fechas" Then
   Cadena = ""
Else
   Cadena = ComboOpcionesCuotas.Text
End If

'titulo principal
MiHoja.Cells(1, 1).Value = "Listado de cuotas " & Cadena & " a la fecha:" & CStr(CDate(DTPicker5.Value))

'establezco el ancho de todas las columnas
MiHoja.Columns.ColumnWidth = 17

FilaTitulos = 2

'pongo los titulos de las columnas..puede ser texto o textbox
 MiHoja.Cells(FilaTitulos, 1).Value = "Prestamo"
MiHoja.Cells(FilaTitulos, 2).Value = "Cuota"
MiHoja.Cells(FilaTitulos, 3).Value = "Cliente"
MiHoja.Cells(FilaTitulos, 4).Value = "DNI"
MiHoja.Cells(FilaTitulos, 5).Value = "Fecha.Nac"
MiHoja.Cells(FilaTitulos, 6).Value = "Provincia"
MiHoja.Cells(FilaTitulos, 7).Value = "Comercio"
MiHoja.Cells(FilaTitulos, 8).Value = "Vencimiento"
MiHoja.Cells(FilaTitulos, 9).Value = "Capital"
MiHoja.Cells(FilaTitulos, 10).Value = "Interes"
MiHoja.Cells(FilaTitulos, 11).Value = "Recargo.Refin."
MiHoja.Cells(FilaTitulos, 12).Value = "Rec.Vto2"
MiHoja.Cells(FilaTitulos, 13).Value = "Otorgamiento"
MiHoja.Cells(FilaTitulos, 14).Value = "Gastos"
MiHoja.Cells(FilaTitulos, 15).Value = "Seguros"
MiHoja.Cells(FilaTitulos, 16).Value = "IVA Interes"
MiHoja.Cells(FilaTitulos, 17).Value = "IVA Seguros"
MiHoja.Cells(FilaTitulos, 18).Value = "IVA OtorGastos"
MiHoja.Cells(FilaTitulos, 19).Value = "Total IVA"
MiHoja.Cells(FilaTitulos, 20).Value = "Saldo"
MiHoja.Cells(FilaTitulos, 21).Value = "Total Cobrado"
MiHoja.Cells(FilaTitulos, 22).Value = "Total cuota"
MiHoja.Cells(FilaTitulos, 23).Value = "IVA Exento"
MiHoja.Cells(FilaTitulos, 24).Value = "Fecha Alta Credito"
MiHoja.Cells(FilaTitulos, 25).Value = "Nro.Cupón"


'pongo los titulos en negritas
MiHoja.Range("a1:y2").Font.Bold = True
'cargo el registro de creditos segun las condiciones de la pantalla
Set rec = CargarRecCuotasCreditos()

'comienzo en la fila 3 porque las anteriores son de titulos
Filas = 3
If Not rec.EOF Then
   Do While Not rec.EOF
      'grabo cada usuario en la planilla
      MiHoja.Cells(Filas, 1).Value = rec.rdoColumns("codprestamo")
      MiHoja.Cells(Filas, 2).Value = CStr(rec.rdoColumns("numcuota")) & " de " & CStr(rec.rdoColumns("numcuotas"))
      MiHoja.Cells(Filas, 3).Value = rec.rdoColumns("cliente")
      MiHoja.Cells(Filas, 4).Value = rec.rdoColumns("numdocumento")
      MiHoja.Cells(Filas, 5).Value = rec.rdoColumns("nacimiento")
      MiHoja.Cells(Filas, 6).Value = rec.rdoColumns("provincia")
      MiHoja.Cells(Filas, 7).Value = rec.rdoColumns("motivobloqueo") & vbNullString
      MiHoja.Cells(Filas, 8).Value = CDate(rec.rdoColumns("fechavencimiento1"))
      MiHoja.Cells(Filas, 9).Value = Format(rec.rdoColumns("importeamortizacion"), "$0.00")
      MiHoja.Cells(Filas, 10).Value = Format(rec.rdoColumns("importeinteres"), "$0.00")
      MiHoja.Cells(Filas, 11).Value = Format(rec.rdoColumns("importerefinanciacion"), "$0.00")
      MiHoja.Cells(Filas, 12).Value = Format(rec.rdoColumns("importerecargovencimiento2"), "$0.00")
      MiHoja.Cells(Filas, 13).Value = Format(rec.rdoColumns("otorgamiento"), "$0.00")
      MiHoja.Cells(Filas, 14).Value = Format(rec.rdoColumns("importegastos"), "$0.00")
      MiHoja.Cells(Filas, 15).Value = Format(rec.rdoColumns("importeseguros"), "$0.00")
      'incluye el ivamora
      MiHoja.Cells(Filas, 16).Value = Format(rec.rdoColumns("ivainteres"), "$0.00")
      MiHoja.Cells(Filas, 17).Value = Format(rec.rdoColumns("ivaseguros"), "$0.00")
      MiHoja.Cells(Filas, 18).Value = Format(rec.rdoColumns("ivaotorgamientogastos"), "$0.00")
      
      SaldoCuota = 0
      TotalIva = 0
      ImporteMora = 0
      IvaMora = 0
      IvaMora2 = 0
      'NUEVO CALCULO DE SALDO SEGUN CREDIMACO(no inclye mora ni iva mora)
      SaldoCuota = ObtenerSaldoCuotaX(rec.rdoColumns("idcredito"), rec.rdoColumns("NumCuota"), CDate(DTPicker5.Value), SaldoCuota1erVenc)
      Importe1erVenc = ObtenerImporte1erVenc(rec.rdoColumns("idcredito"), rec.rdoColumns("NumCuota"))
      'si esta pendiente
      If IsNull(rec.rdoColumns("fechacobro")) Then
         'si hay mora
         If CDate(DTPicker5.Value) > CDate(rec.rdoColumns("fechavencimiento2")) Then
            'ver que pasa si hay mora(ya trae el recargo al 2 vto)
            'calculo la mora de forma habitual
            'le puedo psar el campo [exceptuada]
            ImporteMora = CalculoMoraPendiente(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("exceptuada"), Importe1erVenc, rec.rdoColumns("fechavencimiento1"), CDate(DTPicker5.Value), IvaACobrarDevuelto)
            '''''''********ImporteMora = CalcularInteresMoraZZ(rec.rdoColumns("exceptuada"), Importe1erVenc, rec.rdoColumns("fechavencimiento1"), CDate(DTPicker5.Value))
            IvaMora = 0
            If VG_APLICARIMPUESTOS Then
               If VG_IMPUESTOSCREDIMACO Then
                  'calculo el iva de la mora
                  IvaMora = IvaACobrarDevuelto
               End If
            End If
            '''''''********SoloMoraCobrada = ObtenerMoraCobrada(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
            '''''''********SoloIvaMoraCobrada = ObtenerIvaMoraCobrada(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
            '''''''********If CCur(ImporteMora) <= CCur(SoloMoraCobrada) Then
            '''''''********   ImporteMora = 0
            '''''''********Else
               'si es mayor la mora es solo la diferencia
            '''''''********   ImporteMora = CCur(ImporteMora) - CCur(SoloMoraCobrada)
            '''''''********End If
            '''''''********If CCur(IvaMora) <= CCur(SoloIvaMoraCobrada) Then
            '''''''********   IvaMora = 0
            '''''''********Else
            '''''''********   'si es mayor la mora es solo la diferencia
            '''''''********   IvaMora = CCur(IvaMora) - CCur(SoloIvaMoraCobrada)
            '''''''********End If
            
            SaldoCuota = CCur(SaldoCuota1erVenc) + CCur(ImporteMora) + CCur(IvaMora)
         End If
         
         'el iva mora se calcula
         IvaMora2 = CCur(IvaMora)
         
      Else
         SaldoCuota = 0
         'el ivamora es el cobrado
         IvaMora2 = 0
      End If
   
      TotalIva = CCur(rec.rdoColumns("ivainteres")) + CCur(rec.rdoColumns("ivaseguros")) + CCur(rec.rdoColumns("ivaotorgamientogastos"))
      MiHoja.Cells(Filas, 19).Value = Format(TotalIva, "$0.00")
      
      If rec.rdoColumns("CuotaComodin") Then
          MiHoja.Cells(Filas, 20).Value = Format(0, "$0.00")
          MiHoja.Cells(Filas, 24).Value = "Cuota Comodin"
      Else
          MiHoja.Cells(Filas, 20).Value = Format(SaldoCuota, "$0.00")
          MiHoja.Cells(Filas, 24).Value = ""
      End If
            
      'Capital+ Interes +Gastos Mensuales+ Cargo Otorgamiento+ Seguro de Vida+ Recargo Refinanciación+ IVA Interes +IVA Recargo Refinanciación + IVA Cargo Otorgamiento + IVA Seguro de Vida. + IVA Gastos Mensuales
      TotalCuota = CCur(rec.rdoColumns("importeamortizacion")) + CCur(rec.rdoColumns("importeinteres")) + CCur(rec.rdoColumns("importegastos") + CCur(rec.rdoColumns("otorgamiento")) + CCur(rec.rdoColumns("importeseguros")) + CCur(rec.rdoColumns("importerefinanciacion")) + CCur(rec.rdoColumns("ivainteres")) + CCur(rec.rdoColumns("ivaotorgamientogastos")) + CCur(rec.rdoColumns("ivaseguros")))
                           
      'Importe pagado parcial de la cuota
      ImporteParcial = ObtenerImporteParcialX(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
      MiHoja.Cells(Filas, 21).Value = Format(CCur(ImporteParcial), "$0.00")
                           
                     
      MiHoja.Cells(Filas, 22).Value = Format(TotalCuota, "$0.00")
      
      
      If Not IsNull(rec.rdoColumns("codprestamo")) Then
        If Mid(rec.rdoColumns("codprestamo"), 7, 1) = "C" Then
          MiHoja.Cells(Filas, 23).Value = "SI"
        Else
          MiHoja.Cells(Filas, 23).Value = "NO"
        End If
      End If
      
      
      If Not IsNull(rec.rdoColumns("fechacredito")) Then
          MiHoja.Cells(Filas, 24).Value = CDate(rec.rdoColumns("fechacredito"))
      End If
      
      MiHoja.Cells(Filas, 25).Value = Format(rec.rdoColumns("numfactura"), "0000000")
      
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
   
   Mensaje = "Se exporto el listado de cuotas a la planilla C:\ExportacionExcel\" & Archi
Else
   'grabo los cambios en una tabla falsa
   MiLibro.SaveAs ("c:\ExportacionExcel\temporal.xls")
   'cierro el libro
   MiLibro.Close
   'salgo de excel sin grabar los cambios en el archivo adecuado..uso otro
   MiExcel.Quit
   'borro la tabla falsa
   Set MiExcel = Nothing
   
   Kill ("c:\ExportacionExcel\temporal.xls")
  
   Mensaje = "No hay datos para exportar"
End If
   
MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error Exportando el listado de cuotas...verifique que los archivos de Excel esten cerrados"
End Sub
Private Sub TxtNumCupon_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Call CmdBuscarCupon_Click
End If
End Sub
'*************************************************************
Private Sub CmdExceptuar_Click()
'boton de exceptuar la mora de las cuotas
Dim sql As String
Dim rec As rdoResultset
Dim NumCuota As Long
Dim IdCredito As Long
Dim Mensaje As String
On Error GoTo merror

If Not VerificarSeleccionLista(lvCreditos) Then Exit Sub
   
If Not VerificarSeleccionLista(lvcuotas) Then Exit Sub
   
IdCredito = CLng(lvcuotas.SelectedItem.SubItems(2))
NumCuota = CLng(lvcuotas.SelectedItem.SubItems(3))

If CuotaCobrada(IdCredito, NumCuota) Then
   MsgE "La cuota ya esta cobrada, no se puede exceptuar"
   Exit Sub
End If

'no permito comodin en cuotas que tienen cobros parciales
If ObtenerImporteParcialX(IdCredito, NumCuota) > 0 Then
   MsgE "La cuota tiene cobros parciales...no se puede exceptuar"
   Exit Sub
End If
   
'no permite comodin en creditos bloqueados,finalizados o cuotas cobradas
If CreditoFinalizado(IdCredito) Then
   MsgE "El credito al que pertenece la cuota esta finalizado"
   Exit Sub
End If

If CreditoBloqueado1(IdCredito) Then
   MsgE "El credito al que pertenece la cuota esta bloqueado"
   Exit Sub
End If
   
'valido si es vigente
If CuotaRefinanciada(IdCredito, NumCuota) Then
   MsgE "La cuota esta refinanciada (no esta vigente)"
   Exit Sub
End If

'inicio transaccion
cnSQL.BeginTrans

If EstaExceptuada(IdCredito, NumCuota) Then
   If Not MsgP("La cuota ya esta exceptuada..la desmarca?") Then Exit Sub
   'desmarco
   sql = "update cuotas set logic1= 'False' " & _
         "where idcredito='" & CLng(IdCredito) & "' " & _
         "and numcuota='" & CLng(NumCuota) & "'"
   cnSQL.Execute sql
   
   Mensaje = "La cuota fue desmarcada exitosamente"
Else
   If Not MsgP("Exceptua la cuota seleccionada?") Then Exit Sub
   
   'marco
   sql = "update cuotas set logic1= 'True' " & _
         "where idcredito='" & CLng(IdCredito) & "' " & _
         "and numcuota='" & CLng(NumCuota) & "'"
   cnSQL.Execute sql
   Mensaje = "La cuota fue exceptuada de mora"
End If

'fin de transaccion
cnSQL.CommitTrans

'actualizo las cuotas
Call CmdActualizar_Click
MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error exceptuando mora de cuotas"
End Sub
Private Sub CheckProvincia_Click()
If CheckProvincia.Value = 1 Then
   ComboProvincias.Enabled = True
   ComboProvincias.BackColor = vbWhite
Else
   ComboProvincias.Enabled = False
   ComboProvincias.ListIndex = -1
   ComboProvincias.BackColor = &HFFFFC0
End If
End Sub
Private Sub ComboProvincias_Click()
TxtTotal.Text = 0
Call ActualizarListas
End Sub
Private Sub CmdExportarSaldos_Click()
CmdExportarSaldos.Enabled = False
Me.MousePointer = vbHourglass
Call ExportarSaldos
Me.MousePointer = vbDefault
CmdExportarSaldos.Enabled = True
End Sub
Private Sub ExportarPMC()
Dim Archivo As String
Dim Archi As String
Dim rec1 As rdoResultset
Dim sql As String
Dim Fecha As Date
Dim FechaPagoParcial As Date
Dim MensajeMsg As String
Dim Dia As String
Dim Mes As String
Dim Ano As String
Dim Total As Currency
Dim Imp1 As Long
Dim Imp2 As Long
Dim CantReg As Long
Dim Lin As String, n As Integer
Dim IdClienteActual As Long
Dim IdCreditoActual As Long
Dim IdClienteLeido As Long
Dim IdCreditoLeido As Long
Dim NroDocumentoActual As String
Dim NumDoc As String
Dim FechaProceso As Date
Dim bTienePorVencerCredito As Boolean
Dim bTienePorVencerCliente As Boolean
Dim ImporteVencimiento1 As Currency
Dim ImporteVencimiento2 As Currency
Dim FechaVencimiento1 As Date
Dim FechaVencimiento2 As Date
Dim SaldoVencido1 As Currency
Dim SaldoVencido2 As Currency
Dim FechaVto1SaldosVencidos As Date
Dim FechaVto2SaldosVencidos As Date
Dim ImporteVencimiento1PMC As Currency
Dim ImporteVencimiento2PMC As Currency
Dim SaldoVencimiento1 As Currency
Dim SaldoVencimiento2 As Currency
Dim SaldoCredito1 As Currency
Dim SaldoCredito2 As Currency
Dim FechaVencimiento1PMC As Date
Dim FechaVencimiento2PMC As Date
Dim CuotaActual As Integer
Dim CuotaActualPMC As Integer
Dim Mensaje As String
Dim MensajePMC As String

On Error GoTo merror

If lvCreditos.ListItems.Count() = 0 Then Exit Sub

FechaProceso = CDate(DTPicker5.Value)
FechaVto1SaldosVencidos = DateAdd("d", 10, FechaProceso)
FechaVto2SaldosVencidos = DateAdd("d", 25, FechaProceso)
Dia = Format(CStr(Day(FechaProceso)), "00")
Mes = Format(CStr(Month(FechaProceso)), "00")
Ano = Format(CStr(Year(FechaProceso)), "0000")

Archi = "fac1816."

Archi = Archi + Dia + Mes + Mid$(Ano, 3, 2)

If Not MsgP("¿Confirma la exportacion del listado de Pago Mis Cuentas hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & "?") Then Exit Sub

'estas lineas estaban arriba pero las puse aca para que entren en la transaccion
'si no existe la carpeta la crea
If Not ExisteCarpeta() Then Exit Sub

'verifico si ya existe en la ubicacion
Archivo = Dir("c:\exportacionexcel\" & Archi)

If Trim(Archivo) <> "" Then
  'si ya existe el archivo lo borro para que despues no haya errores
  If Not MsgP("¿El archivo C:\EXPORTACIONEXCEL\" & Archi & " ya existe...desea Borrarlo?") Then Exit Sub
   Kill ("c:\exportacionexcel\" & Archi)
  
End If
    
    cnSQL.BeginTrans
    
    sql = "VaciarPagoMisCuentasTemp"
    cnSQL.Execute sql
   
    For I = 1 To lvcuotas.ListItems.Count
        If lvcuotas.ListItems(I).SubItems(11) > 0 Then
            SaldoVencimiento1 = ObtenerSaldoCuotaOKK(lvcuotas.ListItems(I).SubItems(2), lvcuotas.ListItems(I).SubItems(3), lvcuotas.ListItems(I).SubItems(7), lvcuotas.ListItems(I).SubItems(9), False, FechaVto1SaldosVencidos)
            SaldoVencimiento2 = ObtenerSaldoCuotaOKK(lvcuotas.ListItems(I).SubItems(2), lvcuotas.ListItems(I).SubItems(3), lvcuotas.ListItems(I).SubItems(7), lvcuotas.ListItems(I).SubItems(9), False, FechaVto2SaldosVencidos)
            sql = "InsertarPagoMisCuentasTemp " & lvcuotas.ListItems(I).SubItems(2) & "," & lvcuotas.ListItems(I).SubItems(3) & ",'" & lvcuotas.ListItems(I).SubItems(1) & "'," & lvcuotas.ListItems(I).SubItems(4) & ",'" & ConvertirFechaSql(lvcuotas.ListItems(I).SubItems(7), "DD/MM/YYYY") & "','" & ConvertirFechaSql(lvcuotas.ListItems(I).SubItems(9), "DD/MM/YYYY") & "'," & ConvertirDblSql(lvcuotas.ListItems(I).SubItems(8)) & "," & ConvertirDblSql(lvcuotas.ListItems(I).SubItems(10)) & "," & ConvertirDblSql(lvcuotas.ListItems(I).SubItems(11)) & "," & ConvertirDblSql(SaldoVencimiento1) & "," & ConvertirDblSql(SaldoVencimiento2)
            cnSQL.Execute sql
        End If
    Next I

   Open "c:\exportacionexcel\" & Archi For Output As #1
   Print #1, "04001816" + Ano + Mes + Dia + "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"
    
   cnSQL.CommitTrans
              
   sql = "SELECT * From PagoMisCuentasTemp ORDER BY IdCliente, IdCredito, NumCuota"
      
   Set rec1 = cnSQL.OpenResultset(sql)
   
   If Not rec1.EOF Then
       Do While Not rec1.EOF
       
           IdClienteActual = rec1.rdoColumns("IdCliente")
           IdClienteLeido = rec1.rdoColumns("IdCliente")
           NroDocumentoActual = rec1.rdoColumns("NroDocumento")
           bTienePorVencerCliente = False
           ImporteVencimiento1 = 0
           ImporteVencimiento2 = 0
           FechaVencimiento1 = CDate("2099/12/31")
           FechaVencimiento2 = CDate("2099/12/31")
           SaldoVencido1 = 0
           SaldoVencido2 = 0
           
           Do While IdClienteLeido = IdClienteActual
           
               IdCreditoActual = rec1.rdoColumns("IdCredito")
               IdCreditoLeido = rec1.rdoColumns("IdCredito")
               bTienePorVencerCredito = False
               
               Do While IdCreditoLeido = IdCreditoActual And _
                        IdClienteLeido = IdClienteActual
                   
                   If rec1.rdoColumns("FechaVencimiento2") > FechaProceso And Not bTienePorVencerCredito Then
                   
                       bTienePorVencerCredito = True
                       bTienePorVencerCliente = True
                       
                       If rec1.rdoColumns("NumCuota") = rec1.rdoColumns("CantCuotas") Then
                           If rec1.rdoColumns("FechaVencimiento1") > FechaProceso Then
                               'SaldoCredito1 = ObtenerSaldoCredito(IdCreditoActual, rec1.rdoColumns("FechaVencimiento1"))
                               'If SaldoCredito1 < rec1.rdoColumns("ImporteVencimiento1") Then
                               '     ImporteVencimiento1 = ImporteVencimiento1 + SaldoCredito1
                               'Else
                               '     ImporteVencimiento1 = ImporteVencimiento1 + rec1.rdoColumns("ImporteVencimiento1")
                               'End If
                               ImporteVencimiento1 = ImporteVencimiento1 + ObtenerSaldoCuotaOKK(IdCreditoActual, rec1.rdoColumns("NumCuota"), rec1.rdoColumns("FechaVencimiento1"), rec1.rdoColumns("FechaVencimiento2"), False, rec1.rdoColumns("FechaVencimiento1"))
                           Else
                               'SaldoCredito2 = ObtenerSaldoCredito(IdCreditoActual, rec1.rdoColumns("FechaVencimiento2"))
                               'If SaldoCredito2 < rec1.rdoColumns("ImporteVencimiento2") Then
                               '     ImporteVencimiento1 = ImporteVencimiento1 + SaldoCredito2
                               'Else
                               '     ImporteVencimiento1 = ImporteVencimiento1 + rec1.rdoColumns("ImporteVencimiento2")
                               'End If
                               ImporteVencimiento1 = ImporteVencimiento1 + ObtenerSaldoCuotaOKK(IdCreditoActual, rec1.rdoColumns("NumCuota"), rec1.rdoColumns("FechaVencimiento1"), rec1.rdoColumns("FechaVencimiento2"), False, rec1.rdoColumns("FechaVencimiento2"))
                           End If
                           'SaldoCredito2 = ObtenerSaldoCredito(IdCreditoActual, rec1.rdoColumns("FechaVencimiento2"))
                           'If SaldoCredito2 < rec1.rdoColumns("ImporteVencimiento2") Then
                           '     ImporteVencimiento2 = ImporteVencimiento2 + SaldoCredito2
                           'Else
                           '     ImporteVencimiento2 = ImporteVencimiento2 + rec1.rdoColumns("ImporteVencimiento2")
                           'End If
                           ImporteVencimiento2 = ImporteVencimiento2 + ObtenerSaldoCuotaOKK(IdCreditoActual, rec1.rdoColumns("NumCuota"), rec1.rdoColumns("FechaVencimiento1"), rec1.rdoColumns("FechaVencimiento2"), False, rec1.rdoColumns("FechaVencimiento2"))
                       Else
                           ImporteVencimiento1 = ImporteVencimiento1 + rec1.rdoColumns("ImporteVencimiento1")
                           ImporteVencimiento2 = ImporteVencimiento2 + rec1.rdoColumns("ImporteVencimiento2")
                       End If
                       
                       If rec1.rdoColumns("FechaVencimiento1") < FechaVencimiento1 Then
                           FechaVencimiento1 = rec1.rdoColumns("FechaVencimiento1")
                           CuotaActual = rec1.rdoColumns("NumCuota")
                           Mensaje = Format(rec1.rdoColumns("NumFactura"), "00000000") + Space(5)
                       End If
                       
                       If rec1.rdoColumns("FechaVencimiento2") < FechaVencimiento2 Then
                           FechaVencimiento2 = rec1.rdoColumns("FechaVencimiento2")
                       End If
                       
                   Else
                   
                    If Not bTienePorVencerCredito Then
                        SaldoVencido1 = SaldoVencido1 + rec1.rdoColumns("SaldoVencimiento1")
                        SaldoVencido2 = SaldoVencido2 + rec1.rdoColumns("SaldoVencimiento2")
                    End If
                   
                   End If
                   
                   rec1.MoveNext
                   
                   If rec1.EOF Then
                       IdCreditoLeido = 99999999
                       IdClienteLeido = 99999999
                   Else
                       IdCreditoLeido = rec1.rdoColumns("IdCredito")
                       IdClienteLeido = rec1.rdoColumns("IdCliente")
                   End If
               
               Loop
           
               If Not bTienePorVencerCliente Then
                   'SaldoVencido1 = SaldoVencido1 + ObtenerSaldoCredito(IdCreditoActual, FechaVto1SaldosVencidos)
                   'SaldoVencido2 = SaldoVencido2 + ObtenerSaldoCredito(IdCreditoActual, FechaVto2SaldosVencidos)
                   If SaldoVencido2 < SaldoVencido1 Then
                       SaldoVencido2 = SaldoVencido1
                   End If
               End If
               
           Loop
           
           If bTienePorVencerCliente Then
               
               ImporteVencimiento1PMC = ImporteVencimiento1
               ImporteVencimiento2PMC = ImporteVencimiento2
               FechaVencimiento1PMC = FechaVencimiento1
               FechaVencimiento2PMC = FechaVencimiento2
               CuotaActualPMC = CuotaActual
               MensajePMC = Mensaje
           
           Else
           
               ImporteVencimiento1PMC = SaldoVencido1
               ImporteVencimiento2PMC = SaldoVencido2
               FechaVencimiento1PMC = FechaVto1SaldosVencidos
               FechaVencimiento2PMC = FechaVto2SaldosVencidos
               CuotaActualPMC = 99
               MensajePMC = "SALDO VENCIDO"
               
           End If
           
           NumDoc = Trim$(NroDocumentoActual) & Space(19 - Len(Trim$(NroDocumentoActual)))
           
           Dia1 = Format(CStr(Day(FechaVencimiento1PMC)), "00")
           Mes1 = Format(CStr(Month(FechaVencimiento1PMC)), "00")
           Ano1 = Format(CStr(Year(FechaVencimiento1PMC)), "0000")
            
           Dia2 = Format(CStr(Day(FechaVencimiento2PMC)), "00")
           Mes2 = Format(CStr(Month(FechaVencimiento2PMC)), "00")
           Ano2 = Format(CStr(Year(FechaVencimiento2PMC)), "0000")
           
           Imp1 = Int(Round(ImporteVencimiento1PMC, 2) * 100)
           Imp2 = Int(Round(ImporteVencimiento2PMC, 2) * 100)
            
           CantReg = CantReg + 1
           Total = Total + Imp1
           
           Lin = "5" + NumDoc + MensajePMC + Space(7) + "0" + Ano1 + Mes1 + Dia1 + Format(Imp1, "00000000000") + Ano2 + Mes2 + Dia2 + Format(Imp2, "00000000000") + "00000000" + "00000000000" + "0000000000000000000" + NumDoc + MensajePMC + Space(27) + MensajePMC + Space(2) + Space(60) + "00000000000000000000000000000"
           Print #1, Lin
            
       Loop
   End If
          
   Print #1, "94001816" + Ano + Mes + Dia & Format$(CantReg, "0000000") & "0000000" & Format$(Total, "00000000000") & "00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"
   
   Close #1
      
   MensajeMsg = "Se exporto el listado de Pagos Mis Cuentas a la planilla C:\ExportacionExcel\" & Archi

   
   MsgI MensajeMsg

Exit Sub
merror:
tratarerrores "Error Exportando el listado de Pago Mis Cuentas...verifique que los archivos esten cerrados " & Err.Number & " " & Err.Description
End Sub

Private Sub ExportarSaldos()
Dim MiExcel As Excel.APPLICATION ' Aplicacion Excel
Dim MiLibro As Excel.Workbook ' Un libro de Excel
Dim MiHoja As Excel.Worksheet ' Una hoja de Excel
Dim Filas As Long
Dim FilaTitulos As Long
Dim Archivo As String
Dim Archi As String
Dim rec As rdoResultset
Dim sql As String
Dim SaldoCredito As Currency
Dim SaldoCapital As Currency
Dim CuotasCobradas As Long
Dim CuotasVencidas As Long
Dim CuotasPendientes As Long
Dim Fecha As Date
Dim SaldoCuota As Currency
Dim SaldoCuota1erVenc As Currency
Dim Importe1erVenc As Currency
Dim ImporteParcial As Currency
Dim Mensaje As String
Dim Cadena As String
Dim TotalIva As Currency
Dim Dia As String
Dim Mes As String
Dim Ano As String
Dim RecargoCuota As Currency
Dim ImporteMora As Currency
Dim IvaMora As Currency
Dim IvaMora2 As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
'nuevas 2009
Dim CapitalRestante As Currency
Dim InteresRestante As Currency
Dim GastoRestante As Currency
Dim OtorgamientoRestante As Currency
Dim SeguroRestante As Currency
Dim IvaInteresRestante As Currency
Dim IvaSeguroRestante As Currency
Dim IvaOtorGastoRestante As Currency
Dim Vencimiento2Restante As Currency
Dim RefinRestante As Currency
Dim DiasMora As Long
Dim IvaACobrarDevuelto As Currency
On Error GoTo merror

If lvcuotas.ListItems.Count() = 0 Then Exit Sub

Dia = Format(CStr(Day(CDate(DTPicker5.Value))), "00")
Mes = Format(CStr(Month(CDate(DTPicker5.Value))), "00")
Ano = Format(CStr(Year(CDate(DTPicker5.Value))), "0000")

Archi = "Saldos"

Archi = Archi + "-" + Dia + "-" + Mes + "-" + Ano + ".xls"

If Not MsgP("¿Confirma la exportacion de saldos de cuotas hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & "?") Then Exit Sub

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
MiHoja.Cells(1, 1).Value = "Listado de saldos de cuotas a la fecha:" & CStr(CDate(DTPicker5.Value))

'establezco el ancho de todas las columnas
MiHoja.Columns.ColumnWidth = 17

FilaTitulos = 2

'pongo los titulos de las columnas..puede ser texto o textbox
 MiHoja.Cells(FilaTitulos, 1).Value = "Prestamo"
MiHoja.Cells(FilaTitulos, 2).Value = "Cuota"
MiHoja.Cells(FilaTitulos, 3).Value = "Cliente"
MiHoja.Cells(FilaTitulos, 4).Value = "Nro.Cliente"
MiHoja.Cells(FilaTitulos, 5).Value = "DNI"
MiHoja.Cells(FilaTitulos, 6).Value = "Provincia"
MiHoja.Cells(FilaTitulos, 7).Value = "Comercio"
MiHoja.Cells(FilaTitulos, 8).Value = "Vencimiento1"
MiHoja.Cells(FilaTitulos, 9).Value = "Vencimiento2"
MiHoja.Cells(FilaTitulos, 10).Value = "Recargo.Refin."
MiHoja.Cells(FilaTitulos, 11).Value = "Saldo.Capital"
MiHoja.Cells(FilaTitulos, 12).Value = "Saldo.Interes"
MiHoja.Cells(FilaTitulos, 13).Value = "Saldo.Gastos"
MiHoja.Cells(FilaTitulos, 14).Value = "Saldo.Otorg."
MiHoja.Cells(FilaTitulos, 15).Value = "Saldo.Seg."
MiHoja.Cells(FilaTitulos, 16).Value = "Saldo.Iva.Int."
MiHoja.Cells(FilaTitulos, 17).Value = "Saldo.Iva.Seg."
MiHoja.Cells(FilaTitulos, 18).Value = "Saldo.Iva.Ot."
MiHoja.Cells(FilaTitulos, 19).Value = "Saldo.Rec.2ºvto"
MiHoja.Cells(FilaTitulos, 20).Value = "Saldo.Refin."
'revizar esta porque cambia si esta cobrada o impaga
'si esta impaga es algo solo calculado que se muestra en pantalla
'si esta cobrada es el ivamoracobrado

MiHoja.Cells(FilaTitulos, 21).Value = "Dias.Mora"
MiHoja.Cells(FilaTitulos, 22).Value = "Saldo"
MiHoja.Cells(FilaTitulos, 23).Value = "Observaciones"

'pongo los titulos en negritas
MiHoja.Range("a1:w2").Font.Bold = True
'cargo el registro de creditos segun las condiciones de la pantalla
Set rec = CargarRecCuotasCreditos()

'comienzo en la fila 3 porque las anteriores son de titulos
Filas = 3
If Not rec.EOF Then
   Do While Not rec.EOF
      'grabo cada usuario en la planilla
      MiHoja.Cells(Filas, 1).Value = rec.rdoColumns("codprestamo")
      MiHoja.Cells(Filas, 2).Value = CStr(rec.rdoColumns("numcuota")) & " de " & CStr(rec.rdoColumns("numcuotas"))
      MiHoja.Cells(Filas, 3).Value = rec.rdoColumns("cliente")
      MiHoja.Cells(Filas, 4).Value = rec.rdoColumns("numlegajo")
      MiHoja.Cells(Filas, 5).Value = rec.rdoColumns("numdocumento")
      MiHoja.Cells(Filas, 6).Value = rec.rdoColumns("provincia")
      MiHoja.Cells(Filas, 7).Value = rec.rdoColumns("motivobloqueo") & vbNullString
      MiHoja.Cells(Filas, 8).Value = CDate(rec.rdoColumns("fechavencimiento1"))
      MiHoja.Cells(Filas, 9).Value = CDate(rec.rdoColumns("fechavencimiento2"))
            
      SaldoCuota = 0
      TotalIva = 0
      ImporteMora = 0
      IvaMora = 0
      IvaMora2 = 0
      'NUEVO CALCULO DE SALDO SEGUN CREDIMACO(no inclye mora ni iva mora)
      SaldoCuota = ObtenerSaldoCuotaX(rec.rdoColumns("idcredito"), rec.rdoColumns("NumCuota"), CDate(DTPicker5.Value), SaldoCuota1erVenc)
      Importe1erVenc = ObtenerImporte1erVenc(rec.rdoColumns("idcredito"), rec.rdoColumns("NumCuota"))
      DiasMora = 0
      If CDate(DTPicker5.Value) > CDate(rec.rdoColumns("fechavencimiento1")) Then
         DiasMora = CDate(DTPicker5.Value) - CDate(rec.rdoColumns("fechavencimiento1"))
      End If
   
      'si esta pendiente
      If IsNull(rec.rdoColumns("fechacobro")) Then
         'si hay mora
         If CDate(DTPicker5.Value) > CDate(rec.rdoColumns("fechavencimiento2")) Then
            'ver que pasa si hay mora(ya trae el recargo al 2 vto)
            'calculo la mora de forma habitual
            'le puedo psar el campo [exceptuada]
            ImporteMora = CalculoMoraPendiente(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("exceptuada"), Importe1erVenc, rec.rdoColumns("fechavencimiento1"), CDate(DTPicker5.Value), IvaACobrarDevuelto)
            '''''''********ImporteMora = CalcularInteresMoraZZ(rec.rdoColumns("exceptuada"), Importe1erVenc, rec.rdoColumns("fechavencimiento1"), CDate(DTPicker5.Value))
            IvaMora = 0
            If VG_APLICARIMPUESTOS Then
               If VG_IMPUESTOSCREDIMACO Then
                  'calculo el iva de la mora
                  IvaMora = IvaACobrarDevuelto
               End If
            End If
            '''''''********SoloMoraCobrada = ObtenerMoraCobrada(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
            '''''''********SoloIvaMoraCobrada = ObtenerIvaMoraCobrada(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
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
         
         'el iva mora se calcula
         IvaMora2 = CCur(IvaMora)
         
      Else
         'esta cobrada
         SaldoCuota = 0
         'el ivamora es el cobrado
         IvaMora2 = 0
         DiasMora = 0
      End If
   
      
      MiHoja.Cells(Filas, 22).Value = Format(SaldoCuota, "$0.00")
                   
      MiHoja.Cells(Filas, 10).Value = Format(rec.rdoColumns("importerefinanciacion"), "0.00") & vbNullString
      
      'desde aca MUESTRA LO QUE RESTA DE CADA ITEM
      
      CapitalRestante = CCur(rec.rdoColumns("importeamortizacion")) - ObtenerCapitalCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
      MiHoja.Cells(Filas, 11).Value = Format(CapitalRestante, "0.00") & vbNullString
       
      InteresRestante = CCur(rec.rdoColumns("importeinteres")) - ObtenerInteresCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
      MiHoja.Cells(Filas, 12).Value = Format(InteresRestante, "0.00") & vbNullString
      
      GastoRestante = CCur(rec.rdoColumns("importegastos")) - ObtenerGastosCobrados(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
      MiHoja.Cells(Filas, 13).Value = Format(GastoRestante, "0.00") & vbNullString
      
      OtorgamientoRestante = CCur(rec.rdoColumns("otorgamiento")) - ObtenerOtorgamientoCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
      MiHoja.Cells(Filas, 14).Value = Format(OtorgamientoRestante, "0.00") & vbNullString
      
      SeguroRestante = CCur(rec.rdoColumns("importeseguros")) - ObtenerSegurosCobrados(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
      MiHoja.Cells(Filas, 15).Value = Format(SeguroRestante, "0.00") & vbNullString
    
      IvaInteresRestante = CCur(rec.rdoColumns("ivainteres")) - ObtenerIvaInteresCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
      MiHoja.Cells(Filas, 16).Value = Format(IvaInteresRestante, "0.00") & vbNullString
    
      IvaSeguroRestante = CCur(rec.rdoColumns("ivaseguros")) - ObtenerIvaSegurosCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
      MiHoja.Cells(Filas, 17).Value = Format(IvaSeguroRestante, "0.00") & vbNullString
    
      IvaOtorGastoRestante = CCur(rec.rdoColumns("ivaotorgamientogastos")) - ObtenerIvaOtorGastosCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
      MiHoja.Cells(Filas, 18).Value = Format(IvaOtorGastoRestante, "0.00") & vbNullString
   
      'esta solo va si se cubrio el item en parte o totalmente
      Vencimiento2Restante = 0
      'si esta cobrada
      If Not IsNull(rec.rdoColumns("fechacobro")) Then
         If CDate(rec.rdoColumns("fechacobro")) > CDate(rec.rdoColumns("fechavencimiento1")) Then
            Vencimiento2Restante = CCur(rec.rdoColumns("importerecargovencimiento2")) - ObtenerVencimiento2Cobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
         End If
      Else
         'esta impaga
         If CDate(DTPicker1.Value) > CDate(rec.rdoColumns("fechavencimiento1")) Then
            Vencimiento2Restante = CCur(rec.rdoColumns("importerecargovencimiento2")) - ObtenerVencimiento2Cobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
         End If
      End If
      
      MiHoja.Cells(Filas, 19).Value = Format(Vencimiento2Restante, "0.00") & vbNullString
   
      RefinRestante = CCur(rec.rdoColumns("importerefinanciacion")) - ObtenerRefinCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
      MiHoja.Cells(Filas, 20).Value = Format(RefinRestante, "0.00") & vbNullString
   
      MiHoja.Cells(Filas, 21).Value = Format(DiasMora, "00000") & vbNullString
      
      'el campo original es logic1
      If rec.rdoColumns("exceptuada") Then
         MiHoja.Cells(Filas, 23).Value = "Exceptuada" & vbNullString
      End If
      'el campo original es logic1
      If rec.rdoColumns("vtocambiado") Then
         MiHoja.Cells(Filas, 23).Value = "CambioVto" & vbNullString
      End If
      
      If rec.rdoColumns("CuotaComodin") Then
         MiHoja.Cells(Filas, 11).Value = Format(0, "0.00") & vbNullString
         MiHoja.Cells(Filas, 12).Value = Format(0, "0.00") & vbNullString
         MiHoja.Cells(Filas, 13).Value = Format(0, "0.00") & vbNullString
         MiHoja.Cells(Filas, 14).Value = Format(0, "0.00") & vbNullString
         MiHoja.Cells(Filas, 15).Value = Format(0, "0.00") & vbNullString
         MiHoja.Cells(Filas, 16).Value = Format(0, "0.00") & vbNullString
         MiHoja.Cells(Filas, 17).Value = Format(0, "0.00") & vbNullString
         MiHoja.Cells(Filas, 18).Value = Format(0, "0.00") & vbNullString
         MiHoja.Cells(Filas, 19).Value = Format(0, "0.00") & vbNullString
         MiHoja.Cells(Filas, 20).Value = Format(0, "0.00") & vbNullString
         MiHoja.Cells(Filas, 21).Value = Format(0, "0") & vbNullString
         MiHoja.Cells(Filas, 22).Value = Format(0, "0.00") & vbNullString
         MiHoja.Cells(Filas, 23).Value = "Cuota Comodin" & vbNullString
      End If
      
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
   
   Mensaje = "Se exporto el listado de saldos de cuotas a la planilla C:\ExportacionExcel\" & Archi
Else
   'grabo los cambios en una tabla falsa
   MiLibro.SaveAs ("c:\ExportacionExcel\temporal.xls")
   'cierro el libro
   MiLibro.Close
   'salgo de excel sin grabar los cambios en el archivo adecuado..uso otro
   MiExcel.Quit
   'borro la tabla falsa
   Set MiExcel = Nothing
   
   Kill ("c:\ExportacionExcel\temporal.xls")
  
   Mensaje = "No hay datos para exportar"
End If
   
'fin de transaccion
cnSQL.CommitTrans

MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error Exportando el listado saldos de cuotas...verifique que los archivos de Excel esten cerrados"
End Sub
Public Function CargarCreditos2()
'esta se llama desde otra pantalla, por eso es public
Call ActualizarListas
End Function
Private Sub CheckSoloRefinanciados_Click()
TxtTotal.Text = 0
Call ActualizarListas
End Sub
Private Sub CmdAnularRefin_Click()
'si deseo anular un plan refinanciado
Dim IdCredito As Long
Dim NumCuota As Long
Dim Obs As String
Dim sql As String
Dim rec As rdoResultset
Dim CreditoOld As Long
Dim Cuota As Long
On Error GoTo merror

If lvCreditos.ListItems.Count = 0 Then Exit Sub

Obs = CStr(lvCreditos.SelectedItem.SubItems(12))

If Obs = "REFINANCIACION" Then
   IdCredito = CLng(lvCreditos.SelectedItem.SubItems(1))
   
   If CreditoTieneCobros(IdCredito) Then
      MsgE "El plan refinanciado tiene cobros, no se puede anular"
   Else
      If Not MsgP("Confirma la anulacion del plan refinanciado?") Then Exit Sub
      'sigo adelante con la cancelacion
      'busco todas las cuotas que tengan ese numero de credito actual refinanciado
      'en el campo de asociacion nuevo que agregare
      'y voy limpiando ese campo en esas cuotas liberandolas
      
      'inicio transaccion
      cnSQL.BeginTrans
      
      'como pongo en vigencia los creditos de esas cuotas asociadas?
      'obtengo todas las cuotas asociadas
      sql = "select idcredito,numcuota " & _
            "from cuotas " & _
            "where num1='" & CLng(IdCredito) & "' " & _
            "order by idcredito,numcuota"
      
      Set rec = cnSQL.OpenResultset(sql)
      If Not rec.EOF Then
         'recorro las cuotas asociadas
         Do While Not rec.EOF
            CreditoOld = rec.rdoColumns("idcredito")
            If CreditoFinalizado(CreditoOld) Then
               Call PonerEnVigenciaCredito(CreditoOld)
            End If
            rec.MoveNext
         Loop
      End If
                 
      'libero las cuotas refinanciadas
      sql = "update cuotas set num1=0,fecharefinanciacion=null " & _
            "where num1='" & CLng(IdCredito) & "'"
      
      cnSQL.Execute (sql)
      
      'ahora borro el plan refinanciado que estoy anulando
      'primero borro las cuotas del actual plan refinanciado
      sql = "delete from cuotas where idcredito='" & CLng(IdCredito) & "'"
      cnSQL.Execute sql
      
      'ahora borro el credito pla refinanciado actual
      sql = "delete from creditos where idcredito='" & CLng(IdCredito) & "'"
      cnSQL.Execute sql
      
      'fin de transaccion
      cnSQL.CommitTrans
      
     
      MsgI "La anulacion del plan refinanciado fue realizada con exito"
      'actualizo las listas
      Call ActualizarListas
     
   End If
Else
   MsgE "El credito seleccionado no es una refinanciacion"
End If

Exit Sub
merror:
tratarerrores "Error anulando plan de refinanciacion"
End Sub
Private Sub PonerEnVigenciaCredito(ByVal IdCredito As Long)
Dim sql As String
Dim rec As rdoResultset

sql = "update creditos set fechafinalizacion=null where idcredito='" & CLng(IdCredito) & "'"
cnSQL.Execute sql

End Sub
Private Sub CmdCambiarVto_Click()
'boton de exceptuar la mora de las cuotas
Dim sql As String
Dim rec As rdoResultset
Dim NumCuota As Long
Dim IdCredito As Long
Dim Mensaje As String
Dim Vencimiento1 As Date
Dim Vencimiento2 As Date
Dim Vto1Actual As String
Dim Vto2Actual As String
On Error GoTo merror

If Not VerificarSeleccionLista(lvCreditos) Then Exit Sub
   
If Not VerificarSeleccionLista(lvcuotas) Then Exit Sub
   
IdCredito = CLng(lvcuotas.SelectedItem.SubItems(2))
NumCuota = CLng(lvcuotas.SelectedItem.SubItems(3))
'obtengo los vencimientos originales
'aca estas 2 pueden devolver una fecha o una cadena vacia

'obtengo los vencimientos actuales de la cuota
Vto1Actual = CDate(lvcuotas.SelectedItem.SubItems(7))
Vto2Actual = CDate(lvcuotas.SelectedItem.SubItems(9))

'obtengo los vencimientos de la ultima cuota de ese credito
'estas son 2 fechas siempre
Vencimiento1 = ObtenerUltVto1Cred(IdCredito)
Vencimiento2 = ObtenerUltVto2Cred(IdCredito)

If CuotaCobrada(IdCredito, NumCuota) Then
   MsgE "La cuota ya esta cobrada, no se le puede cambiar el vencimiento"
   Exit Sub
End If

'no permito comodin en cuotas que tienen cobros parciales
If ObtenerImporteParcialX(IdCredito, NumCuota) > 0 Then
   MsgE "La cuota tiene cobros parciales...no se le puede cambiar el vencimiento"
   Exit Sub
End If
   
'no permite comodin en creditos bloqueados,finalizados o cuotas cobradas
If CreditoFinalizado(IdCredito) Then
   MsgE "El credito al que pertenece la cuota esta finalizado"
   Exit Sub
End If

If CreditoBloqueado1(IdCredito) Then
   MsgE "El credito al que pertenece la cuota esta bloqueado"
   Exit Sub
End If
   
'valido si es vigente
If CuotaRefinanciada(IdCredito, NumCuota) Then
   MsgE "La cuota esta refinanciada (no esta vigente)"
   Exit Sub
End If

'si las fechas actuales y ultimas son iguales no cambio nada
If CDate(Vto1Actual) = CDate(Vencimiento1) Or CDate(Vto2Actual) = CDate(Vencimiento2) Then
   MsgE "La cuota ya tiene el vencimiento cambiado desde antes"
   Exit Sub
End If

If Not MsgP("Le cambia el vencimiento a la cuota seleccionada? (¿poniendole el vto de la ultima cuota?") Then Exit Sub
   
'inicio transaccion
cnSQL.BeginTrans

'marco
sql = "update cuotas set logic2='True',fechavencimiento1='" & ConvertirFechaSql(CDate(Vencimiento1), "DD/MM/YYYY") & _
      "',fechavencimiento2='" & ConvertirFechaSql(CDate(Vencimiento2), "DD/MM/YYYY") & _
      "',fech1='" & ConvertirFechaSql(CDate(Vto1Actual), "DD/MM/YYYY") & "',fech2='" & ConvertirFechaSql(CDate(Vto2Actual), "DD/MM/YYYY") & "' " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"
cnSQL.Execute sql
Mensaje = "Se le modifico el vencimiento a la cuota"
   
'fin de transaccion
cnSQL.CommitTrans

'actualizo las cuotas
Call CmdActualizar_Click
MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error cambiando el vencimiento de cuota"
End Sub
Private Sub CmdHistorial_Click()
'imprime el historial de cobros del crdito selecionado

If lvCreditos.ListItems.Count = 0 Then Exit Sub

Call ImprimirHistorialCredito(lvCreditos.SelectedItem.SubItems(1))

End Sub
Private Sub ImprimirHistorialCredito(ByVal IdCredito As Long)
'imprime un resumen del credito
'usada en registrar credito
'usada en consultar creditos
Dim sql As String
Dim rec As rdoResultset
Dim Mreporte As New ARHistorialCredito
On Error GoTo merror

'debe obtener el historial de cobros del credito
'sacandolos de ingresos

sql = "select cuotas.fechavencimiento1,cuotas.importevencimiento1,ingresos.idcredito,ingresos.numcuota,ingresos.codprestamo,ingresos.fechacobro,ingresos.importecobrado," & _
"ingresos.numcomprobante,ingresos.numrecibo,ingresos.pagofacil,ingresos.rapipago,ingresos.concepto," & _
"ingresos.capitalcobrado,ingresos.interescobrado,ingresos.gastoscobrados,ingresos.seguroscobrados,ingresos.vencimiento2cobrado," & _
"ingresos.refincobrado,ingresos.ivainterescobrado,ingresos.ivaseguroscobrado,ingresos.ivaotorgastoscobrado,ingresos.otorgamientocobrado," & _
"ingresos.moracobrada,ingresos.ivamoracobrada,ingresos.descuentos,ingresos.recargos,ingresos.usuario,ingresos.ivaotorgamiento,ingresos.ivagastos " & _
"from cuotas inner join ingresos on (cuotas.idcredito=ingresos.idcredito and cuotas.numcuota=ingresos.numcuota) " & _
"where ingresos.idcredito=" & CLng(IdCredito) & _
" order by ingresos.idcredito,ingresos.numcuota,ingresos.fechacobro"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Mreporte.RDODataControl1.Resultset = rec
   Mreporte.Caption = "Historial de cobros de un credito"
   
   'si imprimo los datos de empresa
   Mreporte.LabelEmpresa = VG_EMPRESA & vbNullString
   
   Mreporte.LabelTitulo = "Historial de cobros del credito:"
   
   Mreporte.LabelCli = ObtenerClienteCredito(IdCredito)
   
   Mreporte.LabelLe = ObtenerLegajoCliente(IdCredito)
     
   Mreporte.LabelNumCredito = Format(CStr(IdCredito), "000000")
   Mreporte.LabelNumCuotas = ObtenerCuotasCredito(IdCredito)
   
   Mreporte.LabelFechaCredito = ObtenerFechaCredito(IdCredito)
   Mreporte.LabelTasa = ObtenerTasaCredito(IdCredito)
   Mreporte.LabelCapital = ObtenerCapitalCredito(IdCredito)
        
   Mreporte.Show vbModal
Else
   MsgE "El credito seleccionado no tiene historial de cobros"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo el historial de creditos"
End Sub
Private Function ObtenerClienteCredito(ByVal IdCredito As Long) As String
Dim sql As String
Dim rec As rdoResultset
Dim Cliente As String
On Error GoTo merror

sql = "select clientes.apellido + ' ' + clientes.nombre as cliente " & _
    "from clientes inner join creditos on clientes.idcliente=creditos.idcliente " & _
    "where creditos.idcredito=" & CLng(IdCredito)
Set rec = cnSQL.OpenResultset(sql)
If Not rec.EOF Then
   Cliente = rec.rdoColumns("cliente")
Else
   Cliente = ""
End If

ObtenerClienteCredito = Cliente

Exit Function
merror:
tratarerrores " Error en funcion ObtenerClienteCredito"
End Function
Private Function ObtenerLegajoCliente(ByVal IdCredito As Long) As String
Dim sql As String
Dim rec As rdoResultset
Dim Legajo As String
On Error GoTo merror

sql = "select clientes.numlegajo " & _
      "from clientes inner join creditos on clientes.idcliente=creditos.idcliente " & _
      "where creditos.idcredito=" & CLng(IdCredito)
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Legajo = rec.rdoColumns("numlegajo")
Else
   Legajo = ""
End If

ObtenerLegajoCliente = Legajo

Exit Function
merror:
tratarerrores "Error en funcion ObtenerLegajoCliente"
End Function
Private Function ObtenerFechaCredito(ByVal IdCredito As Long) As String
Dim sql As String
Dim rec As rdoResultset
Dim Fecha As String
On Error GoTo merror

sql = "select creditos.fechacredito " & _
      "from creditos " & _
      "where creditos.idcredito=" & CLng(IdCredito)
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Fecha = CStr(rec.rdoColumns("fechacredito"))
Else
   Fecha = ""
End If

ObtenerFechaCredito = Fecha

Exit Function
merror:
tratarerrores "Error en funcion ObtenerFechaCredito"
End Function
Private Function ObtenerCapitalCredito(ByVal IdCredito As Long) As String
Dim sql As String
Dim rec As rdoResultset
Dim Capital As Currency
On Error GoTo merror

sql = "select creditos.importeafinanciar " & _
      "from creditos " & _
      "where creditos.idcredito=" & CLng(IdCredito)
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Capital = CCur(rec.rdoColumns("importeafinanciar"))
Else
   Capital = 0
End If

ObtenerCapitalCredito = Capital

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCapitalCredito"
End Function
Private Function ObtenerTasaCredito(ByVal IdCredito As Long) As Double
Dim sql As String
Dim rec As rdoResultset
Dim TasaCredito As Double
On Error GoTo merror

sql = "select creditos.tasa " & _
      "from creditos " & _
      "where creditos.idcredito=" & CLng(IdCredito)
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   TasaCredito = rec.rdoColumns("tasa")
Else
   TasaCredito = 0
End If

ObtenerTasaCredito = TasaCredito

Exit Function
merror:
tratarerrores "Error en funcion ObtenerTasaCredito"
End Function
Private Function ObtenerCuotasCredito(ByVal IdCredito As Long) As Long
Dim sql As String
Dim rec As rdoResultset
Dim Cuotas As String
On Error GoTo merror

sql = "select creditos.numcuotas " & _
    "from creditos " & _
    "where creditos.idcredito=" & CLng(IdCredito)
Set rec = cnSQL.OpenResultset(sql)
If Not rec.EOF Then
   Cuotas = rec.rdoColumns("numcuotas")
Else
   Cuotas = 0
End If

ObtenerCuotasCredito = Cuotas

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCuotasCredito"
End Function
Private Sub CmdRestituirVto_Click()
'boton de exceptuar la mora de las cuotas
Dim sql As String
Dim rec As rdoResultset
Dim NumCuota As Long
Dim IdCredito As Long
Dim Mensaje As String
Dim Vto1Actual As Date
Dim Vto2Actual As Date
Dim Vto1Original As String
Dim Vto2Original As String
On Error GoTo merror

If Not VerificarSeleccionLista(lvCreditos) Then Exit Sub
   
If Not VerificarSeleccionLista(lvcuotas) Then Exit Sub
   
IdCredito = CLng(lvcuotas.SelectedItem.SubItems(2))
NumCuota = CLng(lvcuotas.SelectedItem.SubItems(3))
'obtengo los vencimientos originales
'estas pueden devolver cadenas o fechas
Vto1Original = ObtenerVto1Original(IdCredito, NumCuota)
Vto2Original = ObtenerVto2Original(IdCredito, NumCuota)

'obtengo los vencimientos actuales de la cuota
Vto1Actual = CDate(lvcuotas.SelectedItem.SubItems(7))
Vto2Actual = CDate(lvcuotas.SelectedItem.SubItems(9))

If CuotaCobrada(IdCredito, NumCuota) Then
   MsgE "La cuota ya esta cobrada, no se le puede cambiar el vencimiento"
   Exit Sub
End If

'no permito comodin en cuotas que tienen cobros parciales
If ObtenerImporteParcialX(IdCredito, NumCuota) > 0 Then
   MsgE "La cuota tiene cobros parciales...no se le puede cambiar el vencimiento"
   Exit Sub
End If
   
'no permite comodin en creditos bloqueados,finalizados o cuotas cobradas
If CreditoFinalizado(IdCredito) Then
   MsgE "El credito al que pertenece la cuota esta finalizado"
   Exit Sub
End If

If CreditoBloqueado1(IdCredito) Then
   MsgE "El credito al que pertenece la cuota esta bloqueado"
   Exit Sub
End If
   
'valido si es vigente
If CuotaRefinanciada(IdCredito, NumCuota) Then
   MsgE "La cuota esta refinanciada (no esta vigente)"
   Exit Sub
End If

'solo permito restituir si las fechas originales y las de la ultima cuota
'son iguales (que significa que fueron cambiadas antes)

'si las fechas guardadas temporalmente son vacias no hubo cambio antes
If Trim(Vto1Original) = "" Or Trim(Vto2Original) = "" Then
   MsgE "La cuota no tiene el vencimiento cambiado (no se puede restituir)"
   Exit Sub
End If

'que solo restituya una vez
'si lo guardado es igual al vencimiento actual de la cuota no debo dejarlo
If IsDate(Vto1Original) And IsDate(Vto2Original) Then
   If CDate(Vto1Original) = CDate(Vto1Actual) And CDate(Vto2Original) = CDate(Vto2Actual) Then
      MsgE "La cuota ya fue restituida antes"
      Exit Sub
   End If
End If

If Not MsgP("Le restituye el vencimiento original a la cuota seleccionada?") Then Exit Sub
   
'inicio transaccion
cnSQL.BeginTrans

'marco
sql = "update cuotas set logic2='False',fechavencimiento1='" & ConvertirFechaSql(CDate(Vto1Original), "DD/MM/YYYY") & _
      "',fechavencimiento2='" & ConvertirFechaSql(CDate(Vto2Original), "DD/MM/YYYY") & "' " & _
      "where idcredito='" & CLng(IdCredito) & "' " & _
      "and numcuota='" & CLng(NumCuota) & "'"
cnSQL.Execute sql
   
Mensaje = "Se restituyo el vencimiento original de la cuota"
   
'fin de transaccion
cnSQL.CommitTrans

'actualizo las cuotas
Call CmdActualizar_Click
MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error restituyendo el vencimiento de cuota"
End Sub
Private Function ObtenerVto1Original(ByVal IdCredito As Long, NumCuota As Long) As String
'obtengo la fecha de 1 vto original que esta guardada en una fecha termporal
'puede ocurrir que no haya nada
Dim sql As String
Dim rec As rdoResultset
Dim Fecha As String
On Error GoTo merror

Fecha = ""

sql = "select fech1 " & _
      "from cuotas " & _
      "where idcredito=" & CLng(IdCredito) & " and numcuota=" & CLng(NumCuota)
Set rec = cnSQL.OpenResultset(sql)
If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("fech1")) Then
      Fecha = CStr(rec.rdoColumns("fech1"))
   End If
End If

ObtenerVto1Original = Fecha

Exit Function
merror:
tratarerrores "Error en funcion ObtenerVto1Original"
End Function
Private Function ObtenerVto2Original(ByVal IdCredito As Long, NumCuota As Long) As String
'obtengo la fecha de 1 vto original que esta guardada en una fecha temporal
'puede ocurrir que no haya nada
Dim sql As String
Dim rec As rdoResultset
Dim Fecha As String
On Error GoTo merror

Fecha = ""

sql = "select fech2 " & _
      "from cuotas " & _
      "where idcredito=" & CLng(IdCredito) & " and numcuota=" & CLng(NumCuota)
Set rec = cnSQL.OpenResultset(sql)
If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("fech2")) Then
      Fecha = CStr(rec.rdoColumns("fech2"))
   End If
End If

ObtenerVto2Original = Fecha

Exit Function
merror:
tratarerrores "Error en funcion ObtenerVto2Original"
End Function
Private Sub DTPicker5_Change()
    Call ActualizarListas
End Sub



