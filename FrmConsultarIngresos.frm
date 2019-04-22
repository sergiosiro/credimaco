VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultarIngresos 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultar cobros por periodo"
   ClientHeight    =   8160
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   9465
   ClipControls    =   0   'False
   HelpContextID   =   34
   Icon            =   "FrmConsultarIngresos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5632.178
   ScaleMode       =   0  'User
   ScaleWidth      =   8888.125
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExcedentes 
      Caption         =   "Excedentes"
      Height          =   375
      Left            =   4680
      TabIndex        =   57
      ToolTipText     =   "Exporta la lista de Iva a una planilla Excel del disco"
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton CmdVer 
      Caption         =   "Ver"
      Height          =   255
      Left            =   4320
      TabIndex        =   54
      ToolTipText     =   "Ejecuta la consulta"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton CmdExportarIva 
      Caption         =   "Exportar Iva"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      ToolTipText     =   "Exporta la lista de Iva a una planilla Excel del disco"
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton CmdFactura 
      Caption         =   "Imprimir cuotas"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Imprime la factura del cobro seleccionado"
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Totales cobrados por items:"
      ForeColor       =   &H00FF0000&
      Height          =   2340
      Left            =   120
      TabIndex        =   27
      Top             =   5280
      Width           =   9255
      Begin VB.TextBox TxtTotal2 
         BackColor       =   &H0080FF80&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   53
         ToolTipText     =   "Importe total cobrado sin iva ni mora"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalVencimiento2 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   48
         ToolTipText     =   "Recargo al 2º vto cobrado"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtTotalOtorgamiento 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   20
         ToolTipText     =   "Otorgamiento total cobrado"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TxtTotalIvaMora 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "Iva mora total cobrado"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalIvaOtorGastos 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "Iva otor.Gastos total cobrado"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalIvaSeguros 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Iva Seguro total cobrado"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalRefin 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Recargo por refinanciacion total cobrado"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TxtTotalRecargos 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   22
         ToolTipText     =   "Recargo total efectuado"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalDescuentos 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "Descuento total efectuado"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalMora 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Mora total cobrada"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalIvaInteres 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Iva interes total cobrado"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalSeguros 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Seguro total cobrado"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalGastos 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Gasto total cobrado"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalInteres 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Interes total cobrado"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtTotalCapital 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Capital total cobrado"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtTotalGral 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "Importe total cobrado"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Total sin iva ni mora:"
         Height          =   255
         Left            =   6000
         TabIndex        =   52
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Rec.2º Vto.Cbr:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Otorg.Cobr:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "IvaMoraCobrada:"
         Height          =   255
         Left            =   6000
         TabIndex        =   46
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Iva Ot/Gast.Cobr:"
         Height          =   255
         Left            =   6000
         TabIndex        =   45
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Iva Seg.Cobr:"
         Height          =   255
         Left            =   6000
         TabIndex        =   44
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Rec.Refin.Cbr:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Recargos:"
         Height          =   255
         Left            =   2880
         TabIndex        =   38
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Descuentos:"
         Height          =   255
         Left            =   2880
         TabIndex        =   37
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Mora Cobrada:"
         Height          =   255
         Left            =   2880
         TabIndex        =   36
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Iva Interes.Cobr:"
         Height          =   255
         Left            =   6000
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Seguros.Cobr:"
         Height          =   255
         Left            =   2880
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Gastos.Cobr:"
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Interes cobrado:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Capital cobrado:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Total cobrado:"
         Height          =   255
         Left            =   6000
         TabIndex        =   30
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir listado"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Imprime el listado de cuotas cobradas"
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lista de cuotas:"
      ForeColor       =   &H00FF0000&
      Height          =   3405
      Left            =   120
      TabIndex        =   26
      Top             =   1800
      Width           =   9255
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   135
         Left            =   3360
         TabIndex        =   41
         ToolTipText     =   "Pendiente al dia"
         Top             =   3120
         Width           =   135
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   135
         Left            =   1680
         TabIndex        =   40
         ToolTipText     =   "Cobrada o comodin(si tiene la letra ""C"")"
         Top             =   3120
         Width           =   135
      End
      Begin MSComctlLib.ListView lvcuotas 
         Height          =   2835
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Lista de cobros efectuados en el periodo indicado"
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5001
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
         NumItems        =   31
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Obs."
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Prestamo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Credito"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cuota"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Factura"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Cupon Nº"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "1º Vto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Imp.1º Vto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "2º Vto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Imp.2º Vto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Capital.Cobr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Interes.Cobr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Rec.2Vto.Cobr."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Refin.Cobr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Gastos.Cobr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Seguros.Cobr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Otorg.Cobr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Iva.Int.Cobr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "IvaSeg.Cobr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Iva.Ot/Gast.Cobr"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Mora.Cobr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Iva.Mora.Cobr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Descuentos"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Recargos"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "Fecha.Cobro"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Total.Cobrado"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "Usuario"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "Origen"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "Fecha Imputación"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Cobrador"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label14 
         Caption         =   "Cuotas aun pendientes"
         Height          =   255
         Left            =   3600
         TabIndex        =   43
         ToolTipText     =   "El color azul indica los cobros parciales de cuotas aun pendientes"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Cuotas cobradas"
         Height          =   255
         Left            =   1920
         TabIndex        =   42
         ToolTipText     =   "El color verde indica las cuotas que ya fueron cobradas en su totalidad"
         Top             =   3120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de consulta:"
      ForeColor       =   &H00FF0000&
      Height          =   1290
      Left            =   120
      TabIndex        =   25
      Top             =   72
      Width           =   9255
      Begin VB.CheckBox CheckFiltroFechaImputacion 
         Caption         =   "Filtrar por fecha de imputación"
         Height          =   255
         Left            =   2280
         TabIndex        =   58
         ToolTipText     =   "No incluye en la"
         Top             =   960
         Width           =   3135
      End
      Begin VB.CheckBox CheckProvincia 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   5640
         TabIndex        =   56
         Tag             =   "no"
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox ComboProvincias 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Tag             =   "no"
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox ComboOrden 
         Height          =   315
         ItemData        =   "FrmConsultarIngresos.frx":030A
         Left            =   6360
         List            =   "FrmConsultarIngresos.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Tag             =   "no"
         ToolTipText     =   "Ordenamiento de las cuotas"
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox CheckIncluirParciales 
         Caption         =   "No incluir cuotas cobradas parcialmente"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox CheckIncluirBloqueados 
         Caption         =   "No incluir cuotas de creditos bloqueados"
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   480
         Width           =   3255
      End
      Begin VB.CheckBox CheckIncluirFinalizados 
         Caption         =   "No incluir cuotas de creditos finalizados"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         ToolTipText     =   "No incluye en la"
         Top             =   240
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   720
         TabIndex        =   1
         ToolTipText     =   "Seleccione la fecha de fin de consulta"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   55705601
         CurrentDate     =   39434
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   720
         TabIndex        =   0
         ToolTipText     =   "Seleccione la fecha de inicio de la consulta"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   55705601
         CurrentDate     =   39434
      End
      Begin VB.Label Label15 
         Caption         =   "Orden:"
         Height          =   255
         Left            =   5640
         TabIndex        =   51
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   7920
      TabIndex        =   10
      ToolTipText     =   "Cierra la pantalla"
      Top             =   7680
      Width           =   1500
   End
End
Attribute VB_Name = "FrmConsultarIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***AQUI SE CONSULTAN LOS COBROS DE CUOTAS REGISTRADOS EN UN PERIODO DE FECHAS

Private Sub cmdExcedentes_Click()
Call RefreshTimer
cmdExcedentes.Enabled = False
Me.MousePointer = vbHourglass
Call ExportarExcedentes
Me.MousePointer = vbDefault
cmdExcedentes.Enabled = True
End Sub

Private Sub Form_Load()
Call RefreshTimer
DTPicker1.Value = Date
DTPicker2.Value = Date
ComboOrden.ListIndex = 0
Call CargarComboProvincias("provincias", comboprovincias)
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
Unload Me
End Sub
Private Sub CmdVer_Click()
Call RefreshTimer

If CDate(DTPicker1.Value) > CDate(DTPicker2.Value) Then
   MsgE "La fecha de inicio del periodo debe ser menor que la fecha de fin del periodo"
   Exit Sub
End If

Call CargarCuotasCreditos

If lvcuotas.ListItems.Count() > 0 Then
   CmdImprimir.Enabled = True
   CmdFactura.Enabled = True
   If VG_EXPORTA Then
      CmdExportarIva.Enabled = True
   Else
      CmdExportarIva.Enabled = False
   End If
Else
   CmdImprimir.Enabled = False
   CmdFactura.Enabled = False
   CmdExportarIva.Enabled = False
End If

End Sub
Private Sub CargarCuotasCreditos()
'carga los cobros de cuotas del periodo seleccionado
Dim rec As rdoResultset
Dim Nitem As ListItem
Dim Cad1 As String
Dim Cad2 As String
Dim I As Long
Dim IdCredito As Long
Dim NumCuota As Long
Dim ImporteCobrado As Currency
Dim ImporteVencimiento2 As Currency
Dim TotalCapital As Currency
Dim TotalInteres As Currency
Dim TotalVencimiento2 As Currency
Dim TotalRefin As Currency
Dim TotalGastos As Currency
Dim TotalSeguros As Currency
Dim TotalOtorgamiento As Currency
Dim TotalIvaInteres As Currency
Dim TotalIvaSeguros As Currency
Dim TotalIvaOtorGastos As Currency
Dim TotalMora As Currency
Dim TotalIvaMora As Currency
Dim TotalDescuentos As Currency
Dim TotalRecargos As Currency
Dim TotalGral As Currency
On Error GoTo merror

lvcuotas.ListItems.Clear

Set rec = CargarRecCuotasCreditos()

I = 1
Do While Not rec.EOF
   Cad1 = ""
   Cad2 = ""
   
   '***
   'si tiene cobros parciales y no esta saldada aun
   If rec.rdoColumns("cobrosparciales") Then
      Cad1 = "CP"
   End If
   
   'si es comodin
   If rec.rdoColumns("cuotacomodin") Then
      Cad1 = "*"
   End If
   
   'si esta cobrada
   If Not IsNull(rec.rdoColumns("fechacobro2")) Then
      Cad1 = "C"
   End If
   
   'si esta refinanciada
   If Not IsNull(rec.rdoColumns("fecharefinanciacion")) Then
      Cad1 = "R"
   End If
   
   'si no esta cobrada y no refinanciada y no comodin y fechavencimiento >= hoy
   If IsNull(rec.rdoColumns("fechacobro2")) And IsNull(rec.rdoColumns("fecharefinanciacion")) And Not (rec.rdoColumns("cuotacomodin")) And CDate(rec.rdoColumns("fechavencimiento2")) >= Date Then
      Cad1 = "P"
   End If
   
   'si no esta cobrada y no refinanciada y no comodin y fechavencimiento >= hoy
   If rec.rdoColumns("cobrosparciales") And IsNull(rec.rdoColumns("fechacobro2")) And IsNull(rec.rdoColumns("fecharefinanciacion")) And Not (rec.rdoColumns("cuotacomodin")) And CDate(rec.rdoColumns("fechavencimiento2")) >= Date Then
      Cad1 = "P*"
   End If
   
   'si no esta cobrada, y no refinanciada y no comodin y fechavencimiento<hoy
   If IsNull(rec.rdoColumns("fechacobro")) And IsNull(rec.rdoColumns("fecharefinanciacion")) And Not (rec.rdoColumns("cuotacomodin")) And CDate(rec.rdoColumns("fechavencimiento2")) < Date Then
      'esta en mora
      Cad1 = "M"
   End If
   
   'si no esta cobrada, y no refinanciada y no comodin y fechavencimiento<hoy
   If rec.rdoColumns("cobrosparciales") And IsNull(rec.rdoColumns("fechacobro")) And IsNull(rec.rdoColumns("fecharefinanciacion")) And Not (rec.rdoColumns("cuotacomodin")) And CDate(rec.rdoColumns("fechavencimiento2")) < Date Then
      'esta en mora
      Cad1 = "M*"
   End If
      
   'esto es para indicar que esta cobrada y con cobros parciales
   'si tiene cobros parciales y esta cobrada
   If rec.rdoColumns("cobrosparciales") And Not IsNull(rec.rdoColumns("fechacobro2")) Then
      Cad1 = "C*"
   End If
   
   Set Nitem = lvcuotas.ListItems.Add(, , Cad1)
   IdCredito = CLng(rec.rdoColumns("idcredito"))
   NumCuota = CLng(rec.rdoColumns("numcuota"))
   
   Nitem.SubItems(1) = rec.rdoColumns("cliente") & vbNullString
   Nitem.SubItems(2) = rec.rdoColumns("codprestamo") & vbNullString
   Nitem.SubItems(3) = Format(rec.rdoColumns("idcredito"), "000000") & vbNullString
   Nitem.SubItems(4) = Format(rec.rdoColumns("numcuota"), "00") & vbNullString
   
   'este es el numero de factura que se imprime y ya viene preimpreso
   Nitem.SubItems(5) = rec.rdoColumns("numrecibo") & vbNullString
   
   Nitem.SubItems(6) = Format(rec.rdoColumns("numfactura"), "0000000") & vbNullString
   Nitem.SubItems(7) = rec.rdoColumns("fechavencimiento1") & vbNullString
   Nitem.SubItems(8) = rec.rdoColumns("importetotal") & vbNullString
   Nitem.SubItems(9) = rec.rdoColumns("fechavencimiento2") & vbNullString
   ImporteVencimiento2 = rec.rdoColumns("importetotal") + rec.rdoColumns("importerecargovencimiento2")
   Nitem.SubItems(10) = Format(ImporteVencimiento2, "0.00") & vbNullString
   Nitem.SubItems(11) = Format(rec.rdoColumns("capitalcobrado"), "0.00") & vbNullString
   Nitem.SubItems(12) = Format(rec.rdoColumns("interescobrado"), "0.00") & vbNullString
   Nitem.SubItems(13) = Format(rec.rdoColumns("vencimiento2cobrado"), "0.00") & vbNullString
   Nitem.SubItems(14) = Format(rec.rdoColumns("refincobrado"), "0.00") & vbNullString
   Nitem.SubItems(15) = Format(rec.rdoColumns("gastoscobrados"), "0.00") & vbNullString
   Nitem.SubItems(16) = Format(rec.rdoColumns("seguroscobrados"), "0.00") & vbNullString
   Nitem.SubItems(17) = Format(rec.rdoColumns("otorgamientocobrado"), "0.00") & vbNullString
   Nitem.SubItems(18) = Format(rec.rdoColumns("ivainterescobrado"), "0.00") & vbNullString
   Nitem.SubItems(19) = Format(rec.rdoColumns("ivaseguroscobrado"), "0.00") & vbNullString
   Nitem.SubItems(20) = Format(rec.rdoColumns("ivaotorgastoscobrado"), "0.00") & vbNullString
   Nitem.SubItems(21) = Format(rec.rdoColumns("moracobrada"), "0.00") & vbNullString
   Nitem.SubItems(22) = Format(rec.rdoColumns("ivamoracobrada"), "0.00") & vbNullString
   Nitem.SubItems(23) = Format(rec.rdoColumns("descuentos"), "0.00") & vbNullString
   Nitem.SubItems(24) = Format(rec.rdoColumns("recargos"), "0.00") & vbNullString
   Nitem.SubItems(25) = rec.rdoColumns("fechacobro") & vbNullString
   ImporteCobrado = CCur(rec.rdoColumns("importecobrado2"))
   Nitem.SubItems(26) = Format(ImporteCobrado, "0.00") & vbNullString
   
   TotalCapital = CCur(TotalCapital) + CCur(rec.rdoColumns("capitalcobrado"))
   TotalInteres = CCur(TotalInteres) + CCur(rec.rdoColumns("interescobrado"))
   TotalVencimiento2 = CCur(TotalVencimiento2) + CCur(rec.rdoColumns("vencimiento2cobrado"))
   TotalRefin = CCur(TotalRefin) + CCur(rec.rdoColumns("refincobrado"))
   TotalGastos = CCur(TotalGastos) + CCur(rec.rdoColumns("gastoscobrados"))
   TotalSeguros = CCur(TotalSeguros) + CCur(rec.rdoColumns("seguroscobrados"))
   TotalOtorgamiento = CCur(TotalOtorgamiento) + CCur(rec.rdoColumns("otorgamientocobrado"))
   TotalIvaInteres = CCur(TotalIvaInteres) + CCur(rec.rdoColumns("ivainterescobrado"))
   TotalIvaSeguros = CCur(TotalIvaSeguros) + CCur(rec.rdoColumns("ivaseguroscobrado"))
   TotalIvaOtorGastos = CCur(TotalIvaOtorGastos) + CCur(rec.rdoColumns("ivaotorgastoscobrado"))
   TotalMora = CCur(TotalMora) + CCur(rec.rdoColumns("moracobrada"))
   TotalIvaMora = CCur(TotalIvaMora) + CCur(rec.rdoColumns("ivamoracobrada"))
   TotalDescuentos = CCur(TotalDescuentos) + CCur(rec.rdoColumns("descuentos"))
   TotalRecargos = CCur(TotalRecargos) + CCur(rec.rdoColumns("recargos"))
   TotalGral = CCur(TotalGral) + CCur(rec.rdoColumns("importecobrado2"))

   'If rec.rdoColumns("pagofacil") Then
   '   Nitem.SubItems(27) = "SI"
   'Else
   '   Nitem.SubItems(27) = " "
   'End If
  '
  ' If rec.rdoColumns("rapipago") Then
  '    Nitem.SubItems(28) = "SI"
  ' Else
  '    Nitem.SubItems(28) = " "
  ' End If
   
   'indico el usuario que registro el cobro
   Nitem.SubItems(27) = rec.rdoColumns("usuario") & vbNullString
   
   
    Nitem.SubItems(28) = ""
    If Not IsNull(rec.rdoColumns("origenING")) Then
        Nitem.SubItems(28) = rec.rdoColumns("origenING")
    Else
        If rec.rdoColumns("pagofacil") Then
           Nitem.SubItems(28) = "PMC/ANTICIPO"
        End If
        If rec.rdoColumns("rapipago") Then
           Nitem.SubItems(28) = "PF/RP"
        End If
    End If
    Nitem.SubItems(29) = ""
    If Not IsNull(rec.rdoColumns("fechaimputacionING")) Then
        Nitem.SubItems(29) = Format(rec.rdoColumns("fechaimputacionING"), "dd/mm/yyyy")
    End If
    Nitem.SubItems(30) = ""
    If Not IsNull(rec.rdoColumns("idcobradorING")) Then
        Nitem.SubItems(30) = NombreCobrador(rec.rdoColumns("idcobradorING"))
    End If

   
   'si esta cobrada en su totalidad
   If CuotaCobrada(IdCredito, NumCuota) Then
      'como estan cobradas las pongo en verde
      lvcuotas.ListItems.Item(I).ForeColor = &H8000&
      lvcuotas.ListItems.Item(I).ListSubItems(1).ForeColor = &H8000&
      lvcuotas.ListItems.Item(I).ListSubItems(2).ForeColor = &H8000&
      lvcuotas.ListItems.Item(I).ListSubItems(3).ForeColor = &H8000&
      lvcuotas.ListItems.Item(I).ListSubItems(4).ForeColor = &H8000&
      lvcuotas.ListItems.Item(I).ListSubItems(5).ForeColor = &H8000&
   Else
      'es parcial
      lvcuotas.ListItems.Item(I).ForeColor = &H8000&
      lvcuotas.ListItems.Item(I).ListSubItems(1).ForeColor = &HFF0000
      lvcuotas.ListItems.Item(I).ListSubItems(2).ForeColor = &HFF0000
      lvcuotas.ListItems.Item(I).ListSubItems(3).ForeColor = &HFF0000
      lvcuotas.ListItems.Item(I).ListSubItems(4).ForeColor = &HFF0000
      lvcuotas.ListItems.Item(I).ListSubItems(5).ForeColor = &HFF0000
   End If
   'pongo todas en bold
   lvcuotas.ListItems.Item(I).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(1).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(2).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(3).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(4).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(5).Bold = True
     
   I = I + 1
   rec.MoveNext
Loop

'muestro totales
TxtTotalCapital.Text = Format(TotalCapital, "0.00")
TxtTotalInteres.Text = Format(TotalInteres, "0.00")
TxtTotalVencimiento2.Text = Format(TotalVencimiento2, "0.00")
TxtTotalRefin.Text = Format(TotalRefin, "0.00")
TxtTotalGastos.Text = Format(TotalGastos, "0.00")
TxtTotalSeguros.Text = Format(TotalSeguros, "0.00")
TxtTotalOtorgamiento.Text = Format(TotalOtorgamiento, "0.00")
TxtTotalIvaInteres.Text = Format(TotalIvaInteres, "0.00")
TxtTotalIvaSeguros.Text = Format(TotalIvaSeguros, "0.00")
TxtTotalIvaOtorGastos.Text = Format(TotalIvaOtorGastos, "0.00")
TxtTotalMora.Text = Format(TotalMora, "0.00")
TxtTotalIvaMora.Text = Format(TotalIvaMora, "0.00")
TxtTotalDescuentos.Text = Format(TotalDescuentos, "0.00")
TxtTotalRecargos.Text = Format(TotalRecargos, "0.00")
TxtTotalGral.Text = Format(TotalGral, "0.00")
TxtTotal2.Text = Format(TotalGral - TotalIvaInteres - TotalIvaSeguros - TotalIvaOtorGastos - TotalMora - TotalIvaMora, "0.00")

Exit Sub
merror:
tratarerrores "Error cargando la lista de cuotas cobradas"
End Sub
Private Function CargarRecCuotasCreditos() As rdoResultset
'carga las cuotas cobradas desde ingresos..cada cuota trae los items
'que realmente se cobraron en esa oportunidad
Dim sql As String
Dim CondicionParcial As String
Dim CondicionBloqueadas As String
Dim CondicionFinalizadas As String
Dim CondicionOrden As String
Dim CondicionProvincia As String
Dim IdProvincia As Long
On Error GoTo merror

If CheckIncluirFinalizados.Value = 0 Then
   CondicionFinalizadas = " and 1=1"
Else
   CondicionFinalizadas = " and creditos.fechafinalizacion is Null"
End If

If CheckIncluirBloqueados.Value = 0 Then
   CondicionBloqueadas = " and 1=1"
Else
   CondicionBloqueadas = " and creditos.fechabloqueo is Null"
End If
  
If CheckIncluirParciales.Value = 0 Then
   CondicionParcial = "1=1"
Else
   CondicionParcial = " cuotas.cobrosparciales = 'False'"
End If

If ComboOrden.ListIndex = 0 Then
   'dar mas opciones con una lista
   CondicionOrden = "order by clientes.apellido + ' ' + clientes.nombre,ingresos.codprestamo,ingresos.numcuota,ingresos.fechacobro"
End If

If ComboOrden.ListIndex = 1 Then
   'dar mas opciones con una lista
   CondicionOrden = "order by ingresos.fechacobro,clientes.apellido + ' ' + clientes.nombre,ingresos.codprestamo,ingresos.numcuota"
End If

CondicionProvincia = " and 1=1"
If comboprovincias.Text <> "" Then
   IdProvincia = CLng(comboprovincias.ItemData(comboprovincias.ListIndex))
   CondicionProvincia = " and creditos.idprovincia='" & CLng(IdProvincia) & "'"
End If

sql = "select provincias.nombre as provincia," & _
      "clientes.domicilio,clientes.telefono,clientes.numdocumento," & _
      "clientes.numlegajo,clientes.apellido + ', ' + clientes.nombre as cliente," & _
      "creditos.idcredito as idcredito,creditos.codprestamo,creditos.numcuotas,cuotas.fechacobro as fechacobro2," & _
      "ingresos.fechacobro as fechacobro,cuotas.*,cuotas.importevencimiento1 as importetotal," & _
      "ingresos.importecobrado as importecobrado2,ingresos.numrecibo as numrecibo,ingresos.pagofacil as pagofacil,ingresos.rapipago as rapipago,ingresos.numcomprobante," & _
      "ingresos.capitalcobrado,ingresos.interescobrado,ingresos.vencimiento2cobrado,ingresos.refincobrado,ingresos.gastoscobrados,ingresos.seguroscobrados," & _
      "ingresos.otorgamientocobrado,ingresos.ivainterescobrado,ingresos.ivaseguroscobrado,ingresos.ivaotorgastoscobrado,ingresos.moracobrada,ingresos.ivamoracobrada," & _
      "ingresos.descuentos,ingresos.recargos,ingresos.usuario, ingresos.origen as origenING, ingresos.fechaimputacion as fechaimputacionING, ingresos.idcobrador as idcobradorING " & _
      "from provincias inner join " & _
      "(clientes inner join (creditos inner join (cuotas inner join ingresos on cuotas.idcredito=ingresos.idcredito and cuotas.numcuota=ingresos.numcuota) on " & _
      "creditos.idcredito=cuotas.idcredito) on " & _
      "clientes.idcliente=creditos.idcliente) on " & _
      "provincias.idprovincia=creditos.idprovincia " & _
      "where " & CondicionParcial & CondicionFinalizadas & CondicionBloqueadas & CondicionProvincia
      
If CheckFiltroFechaImputacion.Value Then
    sql = sql & " and (ingresos.fechaimputacion>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
      "and ingresos.fechaimputacion<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "' )" & CondicionOrden
Else
    sql = sql & " and (ingresos.fechacobro>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
      "and ingresos.fechacobro<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "' )" & CondicionOrden
End If

'Set rec = cnSQL.OpenResultset(sql)
Set CargarRecCuotasCreditos = cnSQL.OpenResultset(sql)
Call RefreshTimer

Exit Function
merror:
tratarerrores "Error cargando el registro de cuotas"
End Function
Private Function CargarRecExcedentes() As rdoResultset
'carga las cuotas cobradas desde ingresos..cada cuota trae los items
'que realmente se cobraron en esa oportunidad
Dim sql As String
On Error GoTo merror

sql = "select excedentesclientes.idcliente,clientes.numlegajo,clientes.apellido + ', ' + clientes.nombre as cliente," & _
      "clientes.numdocumento, creditos.codprestamo, excedentesclientes.numcuota, excedentesclientes.idcredito, " & _
      "excedentesclientes.fechacobro, excedentesclientes.importecobro, excedentesclientes.observaciones," & _
      "excedentesclientes.origen, excedentesclientes.fechaimputacion, excedentesclientes.rapipago, excedentesclientes.pagofacil," & _
      "provincias.nombre as provincia, creditos.motivobloqueo as comercio " & _
      "from excedentesclientes, clientes, provincias, creditos "
      
If CheckFiltroFechaImputacion.Value Then
      sql = sql & "where excedentesclientes.fechaimputacion>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
      "and excedentesclientes.fechaimputacion<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "' "
Else
      sql = sql & "where excedentesclientes.fechacobro >='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
      "and excedentesclientes.fechacobro <='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "' "
End If

sql = sql & "and excedentesclientes.idcliente = clientes.idcliente " & _
      "and excedentesclientes.idcredito = creditos.idcredito " & _
      "and excedentesclientes.fechaproceso IS NULL " & _
      "and provincias.idprovincia=creditos.idprovincia "
      
      
sql = sql & " UNION "


sql = sql & "select excedentesclientes.idcliente,clientes.numlegajo,clientes.apellido + ', ' + clientes.nombre as cliente," & _
      "clientes.numdocumento, '' as codprestamo, excedentesclientes.numcuota, excedentesclientes.idcredito, " & _
      "excedentesclientes.fechacobro, excedentesclientes.importecobro, 'Excedente sin crédito asociado' as observaciones," & _
      "excedentesclientes.origen, excedentesclientes.fechaimputacion, excedentesclientes.rapipago, excedentesclientes.pagofacil," & _
      "'' as provincia, '' as comercio " & _
      "from excedentesclientes, clientes "
      
If CheckFiltroFechaImputacion.Value Then
      sql = sql & "where excedentesclientes.fechaimputacion>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
      "and excedentesclientes.fechaimputacion<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "' "
Else
      sql = sql & "where excedentesclientes.fechacobro >='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
      "and excedentesclientes.fechacobro <='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "' "
End If

sql = sql & "and excedentesclientes.idcliente = clientes.idcliente " & _
      "and excedentesclientes.idcredito = 0 " & _
      "and excedentesclientes.fechaproceso IS NULL "

      
sql = sql & " UNION "


sql = sql & "select excedentesclientes.idcliente,'' as numlegajo,'' as cliente," & _
      "'' as numdocumento, '' as codprestamo, excedentesclientes.numcuota, excedentesclientes.idcredito, " & _
      "excedentesclientes.fechacobro, excedentesclientes.importecobro, 'Excedente sin cliente asociado' as observaciones," & _
      "excedentesclientes.origen, excedentesclientes.fechaimputacion, excedentesclientes.rapipago, excedentesclientes.pagofacil," & _
      "'' as provincia, '' as comercio " & _
      "from excedentesclientes "
      
If CheckFiltroFechaImputacion.Value Then
      sql = sql & "where excedentesclientes.fechaimputacion>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
      "and excedentesclientes.fechaimputacion<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "' "
Else
      sql = sql & "where excedentesclientes.fechacobro >='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
      "and excedentesclientes.fechacobro <='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "' "
End If

sql = sql & "and excedentesclientes.idcliente not in (select idcliente from clientes) " & _
      "and excedentesclientes.idcredito = 0 " & _
      "and excedentesclientes.fechaproceso IS NULL "

'Set rec = cnSQL.OpenResultset(sql)
Set CargarRecExcedentes = cnSQL.OpenResultset(sql)

Exit Function
merror:
tratarerrores "Error cargando el registro de excedentes"
End Function

Public Function ObtenerFactura(ByVal IdCredito As Long, NumCuota As Long) As Long

Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerFactura = 0

sql = "select numfactura " & _
      "from cuotas " & _
      "where idcredito=" & IdCredito & " and numcuota =" & NumCuota

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("numfactura")) Then
      ObtenerFactura = rec.rdoColumns("numfactura")
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerFactura"
End Function


Private Sub cmdimprimir_Click()
'imprime la lista de cobros
Call RefreshTimer
CmdImprimir.Enabled = False
If lvcuotas.ListItems.Count > 0 Then
   Call ImprimirListadoCuotas2
End If
CmdImprimir.Enabled = True
End Sub
Private Sub ImprimirListadoCuotas2()
'imprime las cuotas cobradas
Dim rec As rdoResultset
Dim Archivo As String
Dim Mreporte1 As New ARListadoIngresos
On Error GoTo merror

Set rec = CargarRecCuotasCreditos()

If Not rec.EOF Then
   Mreporte1.RDODataControl1.Resultset = rec
   Mreporte1.Caption = "Imprimir listado de cuotas cobradas"
   Mreporte1.LabelTitulo.Caption = "Lista de cuotas cobradas desde el:" & CStr(DTPicker1.Value) & " al " & CStr(DTPicker2.Value)
   
   'imprimo los datos de empresa
   Mreporte1.LabelEmpresa = VG_EMPRESA & vbNullString
   
   'cargo los totales al final del reporte
   Mreporte1.FieldTotalCapital.Text = TxtTotalCapital.Text
   Mreporte1.FieldTotalInteres.Text = TxtTotalInteres.Text
   Mreporte1.FieldTotalGastos.Text = TxtTotalGastos.Text
   Mreporte1.FieldTotalSeguros.Text = TxtTotalSeguros.Text
   Mreporte1.FieldTotalOtorgamiento.Text = TxtTotalOtorgamiento.Text
   Mreporte1.FieldTotalRefin.Text = TxtTotalRefin.Text
   Mreporte1.FieldTotalVencimiento2.Text = TxtTotalVencimiento2.Text
   Mreporte1.FieldTotalIvaInteres.Text = TxtTotalIvaInteres.Text
   Mreporte1.FieldTotalIvaSeguros.Text = TxtTotalIvaSeguros.Text
   Mreporte1.FieldTotalIvaOtGastos.Text = TxtTotalIvaOtorGastos.Text
   Mreporte1.FieldTotalMora.Text = TxtTotalMora.Text
   Mreporte1.FieldTotalIvaMora.Text = TxtTotalIvaMora.Text
   Mreporte1.FieldTotalDesc.Text = TxtTotalDescuentos.Text
   Mreporte1.FieldTotalRec.Text = TxtTotalRecargos.Text
   Mreporte1.FieldTotalGral.Text = TxtTotalGral.Text
   Mreporte1.Show vbModal
Else
   MsgE "No hay cuotas cobradas para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo el listado de cuotas cobradas"
End Sub
Private Sub CmdLimpiar_Click()
Call RefreshTimer
Call LimpiarCampos(Me)
lvcuotas.ListItems.Clear
End Sub
Private Sub CmdFactura_Click()
Dim rec As rdoResultset
Dim Archivo As String
On Error GoTo merror
Call RefreshTimer

Set rec = CargarRecCuotasCreditos()

If Not rec.EOF Then
   Dim Mreporte1 As New ARCuotasCredito4
   Mreporte1.RDODataControl1.Resultset = rec
   Mreporte1.Caption = "Imprimir cuotas de creditos"
   Mreporte1.FieldFecha.Text = Date
   Mreporte1.Printer.Copies = VG_NUMCOPIAS
   Mreporte1.PageSettings.LeftMargin = VG_LEFT
   Mreporte1.PageSettings.TopMargin = VG_TOP
   Mreporte1.PageSettings.TopMargin = VG_BOTOM
   Mreporte1.Show vbModal
End If

Exit Sub
merror:
tratarerrores "Error en boton imprimir facturas"
End Sub
Private Sub ExportarExcedentes()
'exporta a una planilla de excel y la guarda en la carpeta C:\EXPORTACIONEXCEL
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
Dim ImporteParcial As Currency
Dim Mensaje As String
Dim MensajePagoFacil As String
Dim MensajeRapiPago As String
Dim Dia As String
Dim Mes As String
Dim Ano As String
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim NumFactura As Long
Dim Origen As String
Dim FechaImputacion As String
On Error GoTo merror

Dia = Format(CStr(Day(Date)), "00")
Mes = Format(CStr(Month(Date)), "00")
Ano = Format(CStr(Year(Date)), "0000")

FechaDesde = CDate(DTPicker1.Value)
FechaHasta = CDate(DTPicker2.Value)

Archi = "Excedentes"

Archi = Archi + "-" + Dia + "-" + Mes + "-" + Ano + ".xls"

If Not MsgP("¿Confirma la exportacion del listado de Excedentes hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & "?") Then Exit Sub


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
MiHoja.Cells(1, 1).Value = "Listado de Excdedentes entre las fechas:" & Format(FechaDesde, "dd/mm/yyyy") & " y " & Format(FechaHasta, "dd/mm/yyyy")

'establezco el ancho de todas las columnas
MiHoja.Columns.ColumnWidth = 17

FilaTitulos = 2

'pongo los titulos de las columnas..puede ser texto o textbox
MiHoja.Cells(FilaTitulos, 1).Value = "Nro.Cliente"
MiHoja.Cells(FilaTitulos, 2).Value = "Cliente"
MiHoja.Cells(FilaTitulos, 3).Value = "DNI"
MiHoja.Cells(FilaTitulos, 4).Value = "Nº Prestamo"
MiHoja.Cells(FilaTitulos, 5).Value = "Cupon Nº"
MiHoja.Cells(FilaTitulos, 6).Value = "Cuota"
MiHoja.Cells(FilaTitulos, 7).Value = "Importe"
MiHoja.Cells(FilaTitulos, 8).Value = "Fecha"
MiHoja.Cells(FilaTitulos, 9).Value = "Provincia"
MiHoja.Cells(FilaTitulos, 10).Value = "Comercio"
MiHoja.Cells(FilaTitulos, 11).Value = "Tipo"
MiHoja.Cells(FilaTitulos, 12).Value = "Origen"
MiHoja.Cells(FilaTitulos, 13).Value = "Fecha Imputación"


'pongo los titulos en negritas
MiHoja.Range("a1:m2").Font.Bold = True

'cargo el registro de creditos segun las condiciones de la pantalla
Set rec = CargarRecExcedentes()

'comienzo en la fila 3 porque las anteriores son de titulos
Filas = 3
If Not rec.EOF Then
   Do While Not rec.EOF
      
      If rec.rdoColumns("numcuota") > 0 Then
          NumFactura = ObtenerFactura(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
      Else
          NumFactura = 0
      End If
      
      MiHoja.Cells(Filas, 1).Value = rec.rdoColumns("numlegajo")
      MiHoja.Cells(Filas, 2).Value = rec.rdoColumns("cliente")
      MiHoja.Cells(Filas, 3).Value = rec.rdoColumns("numdocumento")
      MiHoja.Cells(Filas, 4).Value = rec.rdoColumns("codprestamo")
      MiHoja.Cells(Filas, 5).Value = NumFactura
      MiHoja.Cells(Filas, 6).Value = rec.rdoColumns("numcuota")
      MiHoja.Cells(Filas, 7).Value = Format(rec.rdoColumns("importecobro"), "$0.00")
      MiHoja.Cells(Filas, 8).Value = CDate(rec.rdoColumns("fechacobro"))
      MiHoja.Cells(Filas, 9).Value = rec.rdoColumns("provincia")
      MiHoja.Cells(Filas, 10).Value = rec.rdoColumns("comercio")
      MiHoja.Cells(Filas, 11).Value = rec.rdoColumns("observaciones")
      Origen = ""
      If Not IsNull(rec.rdoColumns("origen")) Then
          Origen = rec.rdoColumns("origen")
      Else
          If rec.rdoColumns("pagofacil") Then
            Origen = "PMC/ANTICIPO"
          End If
          If rec.rdoColumns("rapipago") Then
            Origen = "PF/RP"
          End If
      End If
      FechaImputacion = ""
      If Not IsNull(rec.rdoColumns("fechaimputacion")) Then
          FechaImputacion = Format(rec.rdoColumns("fechaimputacion"), "dd/mm/yyyy")
      End If
      MiHoja.Cells(Filas, 12).Value = Origen
      MiHoja.Cells(Filas, 13).Value = FechaImputacion
      
      Filas = Filas + 1
             
      rec.MoveNext
   Loop
   
   'grabo los cambios
   MiLibro.SaveAs ("c:\ExportacionExcel\" & Archi)
   'cierro el libro de excel
   MiLibro.Close
   'salgo de excel
   MiExcel.Quit
   Set MiExcel = Nothing
   
   Mensaje = "Se exporto el listado de Excedentes a la planilla C:\ExportacionExcel\" & Archi
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
   
MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error Exportando el listado Excedentes...verifique que Excel y otros archivos Excel esten cerrados"
End Sub

Private Sub CmdExportarIva_Click()
Call RefreshTimer
CmdExportarIva.Enabled = False
Me.MousePointer = vbHourglass
'Call ExportarIva
Call ExportarIvaTXT
Me.MousePointer = vbDefault
CmdExportarIva.Enabled = True
End Sub
Private Sub ExportarIvaTXT()
'exporta a una planilla de excel y la guarda en la carpeta C:\EXPORTACIONEXCEL
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
Dim ImporteParcial As Currency
Dim Mensaje As String
Dim MensajePagoFacil As String
Dim MensajeRapiPago As String
Dim Dia As String
Dim Mes As String
Dim Ano As String
Dim IdCredito As Long
Dim Str19, Str21 As String
Dim Origen As String
Dim FechaImputacion As String
Dim Cobrador As String
Dim NumCuota As Long
On Error GoTo merror

If lvcuotas.ListItems.Count() = 0 Then Exit Sub

Dia = Format(CStr(Day(Date)), "00")
Mes = Format(CStr(Month(Date)), "00")
Ano = Format(CStr(Year(Date)), "0000")

Archi = "ListadoIva"

Archi = Archi + "-" + Dia + "-" + Mes + "-" + Ano + ".xls"

If Not MsgP("¿Confirma la exportacion del listado Iva hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & "?") Then Exit Sub

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

Open "c:\exportacionexcel\" & Archi For Output As #1

'Titulos
Print #1, "Listado de Iva a la fecha:" & CStr(Date)
Print #1, "Nro.Cliente"; Chr$(9); "Cliente"; Chr$(9); "DNI"; Chr$(9); "Nº Prestamo"; Chr$(9); "Factura Nº"; Chr$(9); "Cupon Nº"; Chr$(9); "Capital"; Chr$(9); "Interes"; Chr$(9); "Rec.2º Vto"; Chr$(9); "Recargo.Refin."; Chr$(9); "Gastos"; Chr$(9); "Otorgamiento"; Chr$(9); "Seguros"; Chr$(9); "IVA Interes"; Chr$(9); "IVA Seguros"; Chr$(9); "IVA Ot/Gastos"; Chr$(9); "Mora"; Chr$(9); "IVA Mora"; Chr$(9); "Vencimiento"; Chr$(9); "Descuento"; Chr$(9); "Fecha Cobro"; Chr$(9); "Recargo Cobrado"; Chr$(9); "Total cobrado"; Chr$(9); "Cuota"; Chr$(9); "Provincia"; Chr$(9); "Origen"; Chr$(9); "Fecha Imputación"; Chr$(9); "Cobrador"

'cargo el registro de creditos segun las condiciones de la pantalla
Set rec = CargarRecCuotasCreditos()

If Not rec.EOF Then
   Do While Not rec.EOF
      
      'tomo el dato de una cuota
      IdCredito = CLng(rec.rdoColumns("idcredito"))
      NumCuota = CLng(rec.rdoColumns("numcuota"))
      
      If IsNull(rec.rdoColumns("fechavencimiento1")) Then
        Str19 = rec.rdoColumns("fechavencimiento1")
      Else
        Str19 = CDate(rec.rdoColumns("fechavencimiento1"))
      End If
      
      If IsNull(rec.rdoColumns("fechacobro")) Then
        Str21 = rec.rdoColumns("fechacobro")
      Else
        Str21 = CDate(rec.rdoColumns("fechacobro"))
      End If
      
                         
        Origen = ""
        If Not IsNull(rec.rdoColumns("origenING")) Then
            Origen = rec.rdoColumns("origenING")
        Else
            If rec.rdoColumns("pagofacil") Then
               Origen = "PMC/ANTICIPO"
            End If
            If rec.rdoColumns("rapipago") Then
               Origen = "PF/RP"
            End If
        End If
        FechaImputacion = ""
        If Not IsNull(rec.rdoColumns("fechaimputacionING")) Then
            FechaImputacion = Format(rec.rdoColumns("fechaimputacionING"), "dd/mm/yyyy")
        End If
        Cobrador = ""
        If Not IsNull(rec.rdoColumns("idcobradorING")) Then
            Cobrador = NombreCobrador(rec.rdoColumns("idcobradorING"))
        End If
                                      
       Print #1, rec.rdoColumns("numlegajo") & vbNullString; Chr$(9); rec.rdoColumns("cliente") & vbNullString; Chr$(9); rec.rdoColumns("numdocumento") & vbNullString; Chr$(9); rec.rdoColumns("codprestamo"); Chr$(9); rec.rdoColumns("numrecibo") & vbNullString; Chr$(9); rec.rdoColumns("numfactura") & vbNullString; Chr$(9); Format(rec.rdoColumns("capitalcobrado"), "$0.00"); Chr$(9); Format(rec.rdoColumns("interescobrado"), "$0.00") & vbNullString; Chr$(9); Format(rec.rdoColumns("vencimiento2cobrado"), "$0.00") & vbNullString; Chr$(9); Format(rec.rdoColumns("refincobrado"), "$0.00"); Chr$(9); Format(rec.rdoColumns("gastoscobrados"), "$0.00"); Chr$(9); Format(rec.rdoColumns("otorgamientocobrado"), "$0.00"); Chr$(9); Format(rec.rdoColumns("seguroscobrados"), "$0.00"); Chr$(9); Format(rec.rdoColumns("ivainterescobrado"), "$0.00"); Chr$(9); Format(rec.rdoColumns("ivaseguroscobrado"), "$0.00"); _
       Chr$(9); Format(rec.rdoColumns("ivaotorgastoscobrado"), "$0.00") & vbNullString; Chr$(9); Format(rec.rdoColumns("moracobrada"), "$0.00") & vbNullString _
       ; Chr$(9); Format(rec.rdoColumns("ivamoracobrada"), "$0.00"); Chr$(9); Str19; Chr$(9); Format(rec.rdoColumns("descuentos"), "$0.00"); Chr$(9); Str21 & vbNullString; Chr$(9); Format(rec.rdoColumns("recargos"), "$0.00"); Chr$(9); Format(rec.rdoColumns("importecobrado2"), "$0.00"); Chr$(9); Format(rec.rdoColumns("numcuota"), "00") & " de " & Format(rec.rdoColumns("NumCuotas"), "00"); Chr$(9); rec.rdoColumns("provincia"); _
       Chr$(9); Origen & vbNullString; Chr$(9); FechaImputacion & vbNullString; ; Chr$(9); Cobrador & vbNullString
      
                   
      rec.MoveNext
   Loop
   
   Close #1
   
   Mensaje = "Se exporto el listado de Iva a la planilla C:\ExportacionExcel\" & Archi
Else
   Mensaje = "No hay datos para exportar"
End If
   
'fin de transaccion
cnSQL.CommitTrans

MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error Exportando el listado Iva...verifique que Excel y otros archivos Excel esten cerrados"
End Sub
Private Sub ExportarIva()
'exporta a una planilla de excel y la guarda en la carpeta C:\EXPORTACIONEXCEL
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
Dim ImporteParcial As Currency
Dim Mensaje As String
Dim MensajePagoFacil As String
Dim MensajeRapiPago As String
Dim Dia As String
Dim Mes As String
Dim Ano As String
Dim IdCredito As Long
Dim NumCuota As Long
On Error GoTo merror

If lvcuotas.ListItems.Count() = 0 Then Exit Sub

Dia = Format(CStr(Day(Date)), "00")
Mes = Format(CStr(Month(Date)), "00")
Ano = Format(CStr(Year(Date)), "0000")

Archi = "ListadoIva"

Archi = Archi + "-" + Dia + "-" + Mes + "-" + Ano + ".xls"

If Not MsgP("¿Confirma la exportacion del listado Iva hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & "?") Then Exit Sub

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
MiHoja.Cells(1, 1).Value = "Listado de Iva a la fecha:" & CStr(Date)

'establezco el ancho de todas las columnas
MiHoja.Columns.ColumnWidth = 17

FilaTitulos = 2

'pongo los titulos de las columnas..puede ser texto o textbox
MiHoja.Cells(FilaTitulos, 1).Value = "Nro.Cliente"
MiHoja.Cells(FilaTitulos, 2).Value = "Cliente"
MiHoja.Cells(FilaTitulos, 3).Value = "DNI"
MiHoja.Cells(FilaTitulos, 4).Value = "Nº Prestamo"
MiHoja.Cells(FilaTitulos, 5).Value = "Factura Nº"
MiHoja.Cells(FilaTitulos, 6).Value = "Cupon Nº"
MiHoja.Cells(FilaTitulos, 7).Value = "Capital"
MiHoja.Cells(FilaTitulos, 8).Value = "Interes"
MiHoja.Cells(FilaTitulos, 9).Value = "Rec.2º Vto"
MiHoja.Cells(FilaTitulos, 10).Value = "Recargo.Refin."
MiHoja.Cells(FilaTitulos, 11).Value = "Gastos"
MiHoja.Cells(FilaTitulos, 12).Value = "Otorgamiento"
MiHoja.Cells(FilaTitulos, 13).Value = "Seguros"
MiHoja.Cells(FilaTitulos, 14).Value = "IVA Interes"
MiHoja.Cells(FilaTitulos, 15).Value = "IVA Seguros"
MiHoja.Cells(FilaTitulos, 16).Value = "IVA Ot/Gastos"
MiHoja.Cells(FilaTitulos, 17).Value = "Mora"
MiHoja.Cells(FilaTitulos, 18).Value = "IVA Mora"
MiHoja.Cells(FilaTitulos, 19).Value = "Vencimiento"
MiHoja.Cells(FilaTitulos, 20).Value = "Descuento"
MiHoja.Cells(FilaTitulos, 21).Value = "Fecha Cobro"
MiHoja.Cells(FilaTitulos, 22).Value = "Recargo Cobrado"
MiHoja.Cells(FilaTitulos, 23).Value = "Total cobrado"
MiHoja.Cells(FilaTitulos, 24).Value = "PagoFacil"
MiHoja.Cells(FilaTitulos, 25).Value = "RapiPago"
MiHoja.Cells(FilaTitulos, 26).Value = "Cuota"
MiHoja.Cells(FilaTitulos, 27).Value = "Provincia"



'pongo los titulos en negritas
MiHoja.Range("a1:AA2").Font.Bold = True

'cargo el registro de creditos segun las condiciones de la pantalla
Set rec = CargarRecCuotasCreditos()

'comienzo en la fila 3 porque las anteriores son de titulos
Filas = 3
If Not rec.EOF Then
   Do While Not rec.EOF
      
      'tomo el dato de una cuota
      IdCredito = CLng(rec.rdoColumns("idcredito"))
      NumCuota = CLng(rec.rdoColumns("numcuota"))
      
      MiHoja.Cells(Filas, 1).Value = rec.rdoColumns("numlegajo")
      MiHoja.Cells(Filas, 2).Value = rec.rdoColumns("cliente")
      MiHoja.Cells(Filas, 3).Value = rec.rdoColumns("numdocumento")
      MiHoja.Cells(Filas, 4).Value = rec.rdoColumns("codprestamo")
      MiHoja.Cells(Filas, 5).Value = rec.rdoColumns("numrecibo")
      MiHoja.Cells(Filas, 6).Value = rec.rdoColumns("numfactura")
      MiHoja.Cells(Filas, 7).Value = Format(rec.rdoColumns("capitalcobrado"), "$0.00")
      MiHoja.Cells(Filas, 8).Value = Format(rec.rdoColumns("interescobrado"), "$0.00")
      MiHoja.Cells(Filas, 9).Value = Format(rec.rdoColumns("vencimiento2cobrado"), "$0.00")
      MiHoja.Cells(Filas, 10).Value = Format(rec.rdoColumns("refincobrado"), "$0.00")
      MiHoja.Cells(Filas, 11).Value = Format(rec.rdoColumns("gastoscobrados"), "$0.00")
      MiHoja.Cells(Filas, 12).Value = Format(rec.rdoColumns("otorgamientocobrado"), "$0.00")
      MiHoja.Cells(Filas, 13).Value = Format(rec.rdoColumns("seguroscobrados"), "$0.00")
      MiHoja.Cells(Filas, 14).Value = Format(rec.rdoColumns("ivainterescobrado"), "$0.00")
      MiHoja.Cells(Filas, 15).Value = Format(rec.rdoColumns("ivaseguroscobrado"), "$0.00")
      MiHoja.Cells(Filas, 16).Value = Format(rec.rdoColumns("ivaotorgastoscobrado"), "$0.00")
      MiHoja.Cells(Filas, 17).Value = Format(rec.rdoColumns("moracobrada"), "$0.00")
      MiHoja.Cells(Filas, 18).Value = Format(rec.rdoColumns("ivamoracobrada"), "$0.00")
      
      If IsNull(rec.rdoColumns("fechavencimiento1")) Then
        MiHoja.Cells(Filas, 19).Value = rec.rdoColumns("fechavencimiento1")
      Else
        MiHoja.Cells(Filas, 19).Value = CDate(rec.rdoColumns("fechavencimiento1"))
      End If
      MiHoja.Cells(Filas, 20).Value = Format(rec.rdoColumns("descuentos"), "$0.00")
      If IsNull(rec.rdoColumns("fechacobro")) Then
        MiHoja.Cells(Filas, 21).Value = rec.rdoColumns("fechacobro")
      Else
        MiHoja.Cells(Filas, 21).Value = CDate(rec.rdoColumns("fechacobro"))
      End If
      
      MiHoja.Cells(Filas, 22).Value = Format(rec.rdoColumns("recargos"), "$0.00")
      
      MiHoja.Cells(Filas, 23).Value = Format(rec.rdoColumns("importecobrado2"), "$0.00")
      
      If rec.rdoColumns("pagofacil") Then
         MensajePagoFacil = "SI"
      Else
         MensajePagoFacil = "NO"
      End If
      
      MiHoja.Cells(Filas, 24).Value = MensajePagoFacil & vbNullString
      
      If rec.rdoColumns("rapipago") Then
         MensajeRapiPago = "SI"
      Else
         MensajeRapiPago = "NO"
      End If
      
      MiHoja.Cells(Filas, 25).Value = MensajeRapiPago & vbNullString
      
      MiHoja.Cells(Filas, 26).Value = Format(rec.rdoColumns("numcuota"), "00") & " de " & Format(rec.rdoColumns("NumCuotas"), "00")
      MiHoja.Cells(Filas, 27).Value = rec.rdoColumns("provincia")
      Filas = Filas + 1
             
      rec.MoveNext
   Loop
   
   'grabo los cambios
   MiLibro.SaveAs ("c:\ExportacionExcel\" & Archi)
   'cierro el libro de excel
   MiLibro.Close
   'salgo de excel
   MiExcel.Quit
   Set MiExcel = Nothing
   
   Mensaje = "Se exporto el listado de Iva a la planilla C:\ExportacionExcel\" & Archi
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
tratarerrores "Error Exportando el listado Iva...verifique que Excel y otros archivos Excel esten cerrados"
End Sub
Private Sub CheckProvincia_Click()
If CheckProvincia.Value = 1 Then
   comboprovincias.Enabled = True
   comboprovincias.BackColor = vbWhite
Else
   comboprovincias.Enabled = False
   comboprovincias.ListIndex = -1
   comboprovincias.BackColor = &HFFFFC0
End If
End Sub
Private Sub ComboProvincias_Click()
Call CmdVer_Click
End Sub

