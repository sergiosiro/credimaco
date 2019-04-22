VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReFinanciarCreditos 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Refinanciar Creditos"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   HelpContextID   =   20
   Icon            =   "FrmReFinanciarCreditos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox ComboComercios 
      Height          =   315
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox TxtComercio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      MaxLength       =   255
      TabIndex        =   68
      ToolTipText     =   "Comercio del credito"
      Top             =   3530
      Width           =   2295
   End
   Begin VB.TextBox TxtTotalSellados 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   61
      Tag             =   "N"
      ToolTipText     =   "Importe de sellados del nuevo plan"
      Top             =   3530
      Width           =   1215
   End
   Begin VB.ComboBox ComboProvincias 
      Height          =   315
      ItemData        =   "FrmReFinanciarCreditos.frx":0442
      Left            =   840
      List            =   "FrmReFinanciarCreditos.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   60
      ToolTipText     =   "Provincia de los sellados"
      Top             =   3510
      Width           =   1695
   End
   Begin VB.Frame FrameTasas 
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   9255
      Begin VB.TextBox TxtRecargo 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         MaxLength       =   7
         TabIndex        =   67
         Tag             =   "N"
         ToolTipText     =   "Importe de recargo"
         Top             =   400
         Width           =   1095
      End
      Begin VB.TextBox TxtDescuento 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         MaxLength       =   7
         TabIndex        =   66
         Tag             =   "N"
         ToolTipText     =   "Importe de descuento"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CheckBox CheckRecargos 
         Caption         =   "Recargos"
         Height          =   255
         Left            =   2400
         TabIndex        =   65
         ToolTipText     =   "Aplica recargos a las cuotas seleccionadas"
         Top             =   400
         Width           =   1095
      End
      Begin VB.CheckBox CheckDescuentos 
         Caption         =   "Descuentos"
         Height          =   255
         Left            =   2400
         TabIndex        =   64
         ToolTipText     =   "Aplica descuentos a las cuotas seleccionadas"
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox TxtSubtotal 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   58
         Tag             =   "N"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtTasa2 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   32
         Tag             =   "N"
         ToolTipText     =   "Tasa TEM del plan seleccionado"
         Top             =   360
         Width           =   720
      End
      Begin VB.TextBox TxtImporteAFinanciar 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   8
         Tag             =   "N"
         ToolTipText     =   "Importe a refinanciar"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtCuotasARefinanciar 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "N"
         ToolTipText     =   "Nº de cuotas seleccionadas para refinanciar"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtTasaRefinanciacion 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   10
         Tag             =   "no"
         ToolTipText     =   "Tasa de comision (recargo) por refinanciacion"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TxtTasaFinanciacion 
         Height          =   285
         Left            =   6360
         MaxLength       =   5
         TabIndex        =   9
         Tag             =   "no"
         ToolTipText     =   "Tasa TNA del plan seleccionado"
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label14 
         Caption         =   "Subtotal"
         Height          =   255
         Left            =   4920
         TabIndex        =   59
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "TEM:"
         Height          =   255
         Left            =   7200
         TabIndex        =   29
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "Imp.a refinanciar $"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "C.Selecc. "
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   735
      End
      Begin VB.Label LabelComision 
         Caption         =   "Comis.Refin.%:"
         Height          =   255
         Left            =   8040
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "TNA:"
         Height          =   255
         Left            =   6360
         TabIndex        =   20
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame FrameDeudas 
      Caption         =   "Cuotas adeudadas:"
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   9255
      Begin VB.CheckBox ChTodas 
         Caption         =   "Seleccionar todas"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   71
         Top             =   240
         Width           =   1575
      End
      Begin MSComctlLib.ListView LvDeudas 
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Tag             =   "BORRAR"
         ToolTipText     =   "Lista de cuotas adeudadas por el cliente"
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod.Prestamo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Credito Nº"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Factura Nº"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cuota Nº"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "1º Vto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Importe 1º Vto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "2º Vto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Importe 2º Vto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Saldo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Importe Mora"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      ToolTipText     =   "Limpia la pantalla"
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      ToolTipText     =   "Cierra la pantalla"
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Frame FrameCuotas 
      Caption         =   "Nuevas Cuotas Refinanciadas"
      ForeColor       =   &H00FF0000&
      Height          =   3315
      Left            =   120
      TabIndex        =   17
      Top             =   4680
      Width           =   9255
      Begin VB.CheckBox CHKRenov 
         Caption         =   "RENOVACION REFINANCIADA"
         Height          =   195
         Left            =   2640
         TabIndex        =   73
         Top             =   3000
         Width           =   3015
      End
      Begin VB.CheckBox ChkDDJJ 
         Caption         =   "CERTIFICACION"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   72
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox TxtImporteRecargo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   56
         Tag             =   "N"
         ToolTipText     =   "Importe de recargo por refinanciacion"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox TxtObservaciones 
         Height          =   285
         Left            =   3840
         MaxLength       =   100
         TabIndex        =   54
         ToolTipText     =   "Observaciones del nuevo plan"
         Top             =   2640
         Width           =   5295
      End
      Begin VB.TextBox TxtTotalIvaGastos 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   52
         Tag             =   "N"
         ToolTipText     =   "Importe total de iva sobre otorgamiento y gastos"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox TxtTotalIvaSeguros 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   50
         Tag             =   "N"
         ToolTipText     =   "Importe total de iva seguros"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox TxtTotalOtorgamiento 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   48
         Tag             =   "N"
         ToolTipText     =   "Importe total de otorgamiento"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox TxtImporteTotal 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   47
         Tag             =   "N"
         ToolTipText     =   "Importe total del nuevo credito"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox TxtTotalIvaInteres 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   45
         Tag             =   "N"
         ToolTipText     =   "Importe total de iva interes"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox TxtTotalSeguros 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   43
         Tag             =   "N"
         ToolTipText     =   "Importe total de seguros"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox TxtTotalGastos 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   41
         Tag             =   "N"
         ToolTipText     =   "Importe total de gastos"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox TxtTotalInteres 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   39
         Tag             =   "N"
         ToolTipText     =   "Importe total de interes del nuevo credito"
         Top             =   2280
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvcuotas 
         Height          =   1785
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Lista de cuotas del nuevo credito"
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3149
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
            Text            =   "Cuota"
            Object.Width           =   1060
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Capital"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Interes"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "IVA s/interes"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cuota s/gastos"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Seguros"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "IVA s/seguro"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Otorgamiento"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Gastos"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "IVA s/otorg."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Importe 1º Vto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "1º Vto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Imp.2º Vto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "2º Vto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Recargo Refin."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "IvaRecRefin"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "Total Refin:"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   2640
         TabIndex        =   55
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "Iva Ot/Gastos"
         Height          =   255
         Left            =   6600
         TabIndex        =   53
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Iva seguros:"
         Height          =   255
         Left            =   5520
         TabIndex        =   51
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Otorgamiento:"
         Height          =   255
         Left            =   3360
         TabIndex        =   49
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Importe Total  $:"
         Height          =   255
         Left            =   7680
         TabIndex        =   46
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Iva Interes:"
         Height          =   255
         Left            =   4440
         TabIndex        =   44
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Seguro:"
         Height          =   255
         Left            =   2280
         TabIndex        =   42
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Gastos:"
         Height          =   255
         Left            =   1200
         TabIndex        =   40
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Total Interes:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2040
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "&Registrar Refinanciacion"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Registra el nuevo plan"
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Frame FrameNuevoPLan 
      Caption         =   "Datos del nuevo plan:"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox TxtTotalPtf 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   36
         Tag             =   "N"
         ToolTipText     =   "Importe PTF del nuevo credito"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtCodPrestamo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   35
         ToolTipText     =   "Codigo del prestamo"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtCantidadCuotas 
         Height          =   285
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   30
         Tag             =   "N"
         ToolTipText     =   "Cantidad de cuotas del nuevo plan refinanciado"
         Top             =   480
         Width           =   600
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   3600
         TabIndex        =   0
         ToolTipText     =   "Vencimiento de la primer cuota del nuevo plan"
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54132737
         CurrentDate     =   39081
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   255
         Left            =   5040
         TabIndex        =   1
         ToolTipText     =   "2º vencimiento de la primer cuota del nuevo plan"
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54132737
         CurrentDate     =   39081
      End
      Begin VB.TextBox TxtImporteFinanciado 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   12
         Tag             =   "N"
         ToolTipText     =   "Importe Refinanciado"
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Fecha de refinanciacion"
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54132737
         CurrentDate     =   39081
      End
      Begin VB.TextBox TxtNumCredito 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   27
         Tag             =   "N"
         ToolTipText     =   "Nº de credito nuevo"
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "P.T.F:"
         Height          =   255
         Left            =   6720
         TabIndex        =   37
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Cod.Prestamo:"
         Height          =   255
         Left            =   2280
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Cuotas"
         Height          =   255
         Left            =   1560
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha Refin:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LabelVencimiento2 
         Caption         =   "Fecha 2º Vto:"
         Height          =   255
         Left            =   5040
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Imp.Refin:"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha 1º Vto:"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Credito Nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame FrameCliente 
      Caption         =   "Seleccione un cliente:"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   50
      Width           =   9255
      Begin VB.CommandButton CmdSeleccionar 
         Height          =   375
         Left            =   8280
         Picture         =   "FrmReFinanciarCreditos.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Permite seleccionar al cliente de una lista"
         Top             =   160
         Width           =   615
      End
      Begin VB.TextBox TxtCliente 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "no"
         ToolTipText     =   "Cliente al cual le refinanciaremos un credito"
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Label Label15 
      Caption         =   "Comercio:"
      Height          =   255
      Left            =   5040
      TabIndex        =   69
      Top             =   3555
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "Sellados:"
      Height          =   255
      Left            =   2760
      TabIndex        =   63
      Top             =   3530
      Width           =   735
   End
   Begin VB.Label Label24 
      Caption         =   "Provincia:"
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   3530
      Width           =   735
   End
End
Attribute VB_Name = "FrmReFinanciarCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***AQUI SE REFINANCIAN DEUDAS Y SE GENERA UN NUEVO CREDITO
'***EN LAS REFINANCIACIONES NO SE USAN PLANES Y LAS CUOTAS SON
'SIEMPRE MENSUALES
Public IdCliente As Long

Private Sub ChkDDJJ_Click()
'si cambio la condicion de certificacion
Call CalcularImportes
End Sub

Private Sub ChTodas_Click()
    Dim I  As Integer
    
    TxtImporteAFinanciar.Text = 0
    TxtCuotasARefinanciar.Text = 0
    
    For I = 1 To LvDeudas.ListItems.Count
        LvDeudas.ListItems.Item(I).Checked = ChTodas.Value
        If ChTodas.Value Then
            Call LvDeudas_ItemCheck(LvDeudas.ListItems.Item(I))
        End If
    Next
    
 End Sub

Private Sub ComboComercios_Click()
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

BandComercio = True
sql = "select * from comercios WHERE idcomercio = '" & CLng(ComboComercios.ItemData(ComboComercios.ListIndex)) & "'"

Set rec = cnSQL.OpenResultset(sql)
If Not rec.EOF Then
For I = 0 To ComboProvincias.ListCount - 1
      If ComboProvincias.ItemData(I) = rec.rdoColumns("idprovincia") Then
       
       ComboProvincias.ListIndex = I
       BandComercio = False
       Exit For
      End If
   Next I
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento cargandoprovincia"
End Sub

Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call RefrescarOpcionesSistema

Call LimpiarCampos(Me)

IdCliente = 0

Call CargarComboProvincias("provincias", ComboProvincias)

TxtTasaRefinanciacion.Text = VG_TASAREFINANCIACION & vbNullString

TxtNumCredito.Text = Format(UltimoId("idcredito", "creditos") + 1, "000000")

'en credimaco no hay segundo vencimiento
If VG_APLICARSEGUNDOVENCIMIENTO Then
   LabelVencimiento2.Visible = True
   DTPicker3.Visible = True
End If

If VG_APLICARTASAREFINANCIACION Then
   LabelComision.Visible = True
   TxtTasaRefinanciacion.Visible = True
End If

Call CargarCombo2("comercios", ComboComercios)

Exit Sub
merror:
tratarerrores "Error cargando la pantalla RefinanciarCreditos"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
IdCliente = 0
Unload Me
End Sub
Private Sub cmdnuevo_Click()
Call RefreshTimer
CmdGenerar.Enabled = True
IdCliente = 0
TxtCliente.Text = ""
Call LimpiarCampos(Me)
TxtNumCredito.Text = Format(UltimoId("idcredito", "creditos") + 1, "000000")
End Sub
Private Sub CmdSeleccionar_Click()
FrmClientesAbm.FormularioPadre = "REFINANCIARCREDITOS"
Call CenterForm(FrmClientesAbm)
FrmClientesAbm.Show
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(TxtCliente.Text) = "" Then
   datosok = False
   MsgE "Debe seleccionar un cliente"
   CmdSeleccionar.SetFocus
   Exit Function
End If

'varifico si tiene deudas
If (LvDeudas.ListItems.Count = 0) Then
   datosok = False
   MsgE "El cliente no adeuda cuotas"
   Exit Function
End If

'si hay deudas marcadas
If Not HayFilasChequeadas(LvDeudas) Then
   datosok = False
   MsgE "Debe marcar las cuotas a refinanciar en la lista superior"
   LvDeudas.SetFocus
   Exit Function
End If

'verifico si selecciono cuotas de distintos creditos
If CHKRenov.Value = False Then
    If HayCreditosDistintos(LvDeudas) Then
        datosok = False
        MsgE "Debe seleccionar cuotas de un mismo credito (no se pueden refinanciar creditos distintos)"
        LvDeudas.SetFocus
        Exit Function
    End If
End If
   
'If Trim(TxtCodPrestamo.Text) = "" Then
'   datosok = False
'   MsgE "Debe ingresar el codigo de prestamo"
'   TxtCodPrestamo.SetFocus
'   Exit Function
'End If

'valido los datos de la refinanciacion
If Trim(TxtImporteAFinanciar.Text) = "" Then
   datosok = False
   MsgE "Falta el importe a refinanciar..debe seleccionar cuotas"
   Exit Function
End If

'valido el numero de credito porque ahora lo pueden ingresar a mano
If Trim(TxtNumCredito.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el numero de credito"
   TxtNumCredito.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtNumCredito.Text) Then
   datosok = False
   MsgE "El numero de credito debe ser numerico"
   TxtNumCredito.SetFocus
   Exit Function
End If
If CLng(TxtNumCredito.Text) <= 0 Then
   datosok = False
   MsgE "El numero de credito debe ser mayor a cero"
   TxtNumCredito.SetFocus
   Exit Function
End If

'valido la tasa (antes era TNA)
If Trim(TxtTasaFinanciacion.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar la tasa de financiacion"
   TxtTasaFinanciacion.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtTasaFinanciacion.Text) Then
   datosok = False
   MsgE "La tasa de financiacion debe ser numerica"
   TxtTasaFinanciacion.SetFocus
   Exit Function
End If
If CDbl(TxtTasaFinanciacion.Text) < 0 Then
   datosok = False
   MsgE "La tasa de financiacion debe ser mayor o igual a cero"
   TxtTasaFinanciacion.SetFocus
   Exit Function
End If

'valido la cantidad de cuotas
If Trim(TxtCantidadCuotas.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar la cantidad de cuotas del nuevo plan"
   TxtCantidadCuotas.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtCantidadCuotas.Text) Then
   datosok = False
   MsgE "La cantidad de cuotas debe ser numerica"
   TxtCantidadCuotas.SetFocus
   Exit Function
End If
If CLng(TxtCantidadCuotas.Text) <= 0 Then
   datosok = False
   MsgE "La cantidad de cuotas debe ser mayor a cero"
   TxtCantidadCuotas.SetFocus
   Exit Function
End If

'esto es para sellados
If ComboProvincias.Text = "" Then
   datosok = False
   MsgE "Debe seleccionar la provincia"
   ComboProvincias.SetFocus
   Exit Function
End If

If Trim(ComboComercios.Text) = "" Then
   datosok = False
   MsgE "Debe seleccionar un comercio"
   ComboComercios.SetFocus
   Exit Function
End If

If Trim(TxtObservaciones.Text) = "" Then
   TxtObservaciones.Text = vbNullString
End If

'si no permito registrar creditos diferidos
If Not VG_CREDITOSDIFERIDOS Then
   If CDate(DTPicker1.Value) <> CDate(Date) Then
      datosok = False
      MsgE "Verifique la fecha de refinanciacion...debe ser igual a la actual"
      Exit Function
   End If
Else
  'si permito diferidos verifico que no sean menores a 2000
   '***valido la fecha del credito
   If Year(DTPicker1.Value) < 2000 Then
      datosok = False
      MsgE "Verifique la fecha de refinanciacion...(el año debe ser superior a 2000)"
      DTPicker1.SetFocus
      Exit Function
   End If
   If CDate(DTPicker1.Value) > CDate(Date) Then
      datosok = False
      MsgE "Verifique la fecha de refinanciacion...(la fecha debe ser la actual)"
      DTPicker1.SetFocus
      Exit Function
   End If
End If

If CDate(DTPicker2.Value) < CDate(DTPicker1.Value) Then
   datosok = False
   MsgE "El primer vencimiento debe ser mayor o igual a la fecha de refinanciacion"
   DTPicker2.SetFocus
   Exit Function
End If

If CDate(DTPicker1.Value) < CDate(VG_FECHALIMITEINGRESO) Then
   datosok = False
   MsgE "La Fecha del credito es inferior a la fecha limite permitida"
   DTPicker1.SetFocus
   Exit Function
End If

'si el primer vto es en el mes actual de refinanciacion
If Month(CDate(DTPicker1.Value)) = Month(DTPicker2.Value) Then
   If Not MsgP("¿El primer vencimiento es en el mes de registracion del credito?") Then
      datosok = False
      Exit Function
   End If
End If

'si es un dia habil
If EsFeriado(DTPicker2.Value) Or EsSabado(DTPicker2.Value) Or EsDomingo(DTPicker2.Value) Then
   datosok = False
   MsgE "El 1º vencimiento no es un dia habil..."
   DTPicker2.SetFocus
   Exit Function
End If

If VG_APLICARSEGUNDOVENCIMIENTO Then
   If EsFeriado(DTPicker3.Value) Or EsSabado(DTPicker3.Value) Or EsDomingo(DTPicker3.Value) Then
      datosok = False
      MsgE "El 2º vencimiento no es un dia habil"
      DTPicker3.SetFocus
      Exit Function
   End If
   If CDate(DTPicker3.Value) < CDate(DTPicker2.Value) Then
      datosok = False
      MsgE "El 2º vencimiento debe ser mayor que el 1º vencimiento"
      DTPicker3.SetFocus
      Exit Function
   End If
End If

'verifico que haya cuotas en la lista
If lvcuotas.ListItems.Count = 0 Then
   datosok = False
   MsgE "No hay cuotas en la lista"
   Exit Function
End If

'reemplazo caracteres invalidos del teclado que pudieron
'cargar manualmente
Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-RefinanciarCreditos"
End Function
Private Sub CargarDeudasCliente(ByVal IdCliente As Long)
'carga todas las cuotas pendientes actualizadas si es necesario
Dim rec As rdoResultset
Dim sql As String
Dim Nitem As ListItem
Dim ImporteVencimiento2 As Currency
Dim ImporteMora As Currency
Dim Sumatotal As Currency
Dim ImporteParcial As Currency
Dim I As Long
Dim SaldoCuota As Currency
Dim SaldoCuota1erVenc As Currency
Dim Importe1erVenc As Currency
Dim IvaMora As Currency
Dim RecargoCuota As Currency
Dim ImporteMoraGral As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim IvaACobrarDevuelto  As Currency
On Error GoTo merror

'obtiene todas las cuotas pendientes
sql = "SELECT creditos.idcredito,creditos.codprestamo," & _
      "cuotas.numfactura,cuotas.numcuota,cuotas.cobrosparciales," & _
      "cuotas.fechavencimiento1,cuotas.fechavencimiento2," & _
      "cuotas.importecuota,cuotas.importegastos,cuotas.importeseguros," & _
      "cuotas.importeimpuestos,cuotas.importerecargovencimiento2," & _
      "cuotas.fechacobro,cuotas.importecobrado,cuotas.logic1 as exceptuada," & _
      "cuotas.importevencimiento1,cuotas.importevencimiento2 " & _
      "from creditos inner join cuotas on creditos.idcredito=cuotas.idcredito " & _
      "where creditos.fechafinalizacion is Null " & _
      "and creditos.fechabloqueo is Null " & _
      "and cuotas.cuotacomodin =0 " & _
      "and cuotas.fecharefinanciacion is Null " & _
      "and cuotas.fechacobro is null " & _
      "and creditos.idcliente=" & CLng(IdCliente) & " " & _
      "order by creditos.idcredito,cuotas.numcuota"

Set rec = cnSQL.OpenResultset(sql)
   
LvDeudas.ListItems.Clear

Sumatotal = 0

If Not rec.EOF Then
   I = 1
   Do While Not rec.EOF
      Set Nitem = LvDeudas.ListItems.Add(, , rec.rdoColumns("codprestamo"))
      Nitem.SubItems(1) = Format(rec.rdoColumns("idcredito"), "000000000") & vbNullString
      Nitem.SubItems(2) = Format(rec.rdoColumns("numfactura"), "000000000") & vbNullString
      Nitem.SubItems(3) = Format(rec.rdoColumns("numcuota"), "00") & vbNullString
      Nitem.SubItems(4) = rec.rdoColumns("fechavencimiento1") & vbNullString
      'este es el importe original al primer vto
      Nitem.SubItems(5) = Format(rec.rdoColumns("importevencimiento1"), "0.00") & vbNullString
      Nitem.SubItems(6) = rec.rdoColumns("fechavencimiento2") & vbNullString
      
      'el importe original al segundo vto
      ImporteVencimiento2 = CCur(rec.rdoColumns("importevencimiento1")) + CCur(rec.rdoColumns("importerecargovencimiento2"))
      Nitem.SubItems(7) = Format(rec.rdoColumns("ImporteVencimiento2"), "0.00") & vbNullString
                
      'el importe de cobros parciales
      ImporteParcial = ObtenerImporteParcialX(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
                
      SaldoCuota = ObtenerSaldoCuotaX(rec.rdoColumns("idcredito"), rec.rdoColumns("NumCuota"), DTPicker1.Value, SaldoCuota1erVenc)
      Importe1erVenc = ObtenerImporte1erVenc(rec.rdoColumns("idcredito"), rec.rdoColumns("NumCuota"))
      'verifico si hay mora
      If CDate(DTPicker1.Value) > CDate(rec.rdoColumns("fechavencimiento2")) Then
         'calculo la mora en forma habitual
         'puedo pasarle el campo [exceptuada]
         ImporteMora = CalculoMoraPendiente(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("exceptuada"), Importe1erVenc, rec.rdoColumns("fechavencimiento1"), CDate(DTPicker1.Value), IvaACobrarDevuelto)
         '''''''********ImporteMora = CalcularInteresMoraZZ(rec.rdoColumns("exceptuada"), Importe1erVenc, rec.rdoColumns("fechavencimiento1"), DTPicker1.Value)
         
         IvaMora = 0
         If VG_APLICARIMPUESTOS Then
            If VG_IMPUESTOSCREDIMACO Then
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
         
         'pongo en rojo sin cobrar en mora
         LvDeudas.ListItems.Item(I).ForeColor = vbRed
         LvDeudas.ListItems.Item(I).ListSubItems(1).ForeColor = vbRed
         LvDeudas.ListItems.Item(I).ListSubItems(2).ForeColor = vbRed
         LvDeudas.ListItems.Item(I).ListSubItems(3).ForeColor = vbRed
         LvDeudas.ListItems.Item(I).ListSubItems(4).ForeColor = vbRed
      Else
          'pongo en azul sin cobrar al dia
          LvDeudas.ListItems.Item(I).ForeColor = &HFF0000
          LvDeudas.ListItems.Item(I).ListSubItems(1).ForeColor = &HFF0000
          LvDeudas.ListItems.Item(I).ListSubItems(2).ForeColor = &HFF0000
          LvDeudas.ListItems.Item(I).ListSubItems(3).ForeColor = &HFF0000
          LvDeudas.ListItems.Item(I).ListSubItems(4).ForeColor = &HFF0000
      End If
                
      Nitem.SubItems(8) = Format(SaldoCuota, "0.00") & vbNullString
      Sumatotal = CCur(Sumatotal) + CCur(SaldoCuota)
      I = I + 1
      rec.MoveNext
   Loop
Else
   MsgI "El cliente no tiene deudas"
End If

Exit Sub
merror:
tratarerrores "Error cargando deudas de clientes-RefinanciarCreditos"
End Sub
Private Sub cmdgenerar_click()
'genero el nuevo credito
Dim sql As String
Dim IdCreditoAnterior As Long
Dim IdCredito As Long
Dim Nombregarante As String
Dim I As Long
Dim NumFactura As Long
Dim NumFactura2 As Long
Dim NumFactura3 As Long
Dim NumCuota As Long
Dim Vencimiento1 As Date
Dim Vencimiento2 As Date
Dim ImporteVencimiento1 As Currency
Dim ImporteVencimiento2 As Currency
Dim ImporteAmortizacion As Currency
Dim ImporteInteres As Currency
Dim ImporteCuota As Currency
Dim ImporteGastos As Currency
Dim ImporteSeguros As Currency
Dim ImporteImpuestos As Currency
Dim ImporteSellados As Currency
Dim ImporteRecargo As Currency
Dim ImporteRecargoVencimiento2 As Currency
Dim CodigoBarras As String
Dim Periodo As String
Dim Formula As String
Dim IdPlan As Long
Dim IvaInteres As Currency
Dim IvaSeguros As Currency
Dim IvaOtorgamientoGastos As Currency
Dim IdProvincia As Long
Dim DiasRecargo As Long
On Error GoTo merror
Call RefreshTimer

If Not datosok() Then Exit Sub

Formula = "AMORTIZACION SISTEMA FRANCES REFINANCIACION"

'valido el cliente
If Not ExisteCliente(IdCliente) Then
   MsgE "El cliente no existe"
   Exit Sub
End If

'lo vuelvo a generar porque no es mas manual
TxtNumCredito.Text = Format(UltimoId("idcredito", "creditos") + 1, "000000")
IdCredito = CLng(TxtNumCredito.Text)

'valido el credito nuevo
If ExisteCredito(IdCredito) Then
   MsgE "El numero de credito ya existe...(Debe ingresar uno nuevo)"
   TxtNumCredito.SetFocus
   Exit Sub
End If

'valido el numero de prestamo
If CHKRenov Then
    TxtCodPrestamo.Text = NuevoCodPrestamo(IdCliente, "W")
Else
    TxtCodPrestamo.Text = NuevoCodPrestamo(IdCliente, "M")
End If
If ExisteCodPrestamo(TxtCodPrestamo.Text) Then
   MsgE "El codigo de prestamo ya existe...(Debe ingresar uno nuevo)"
   TxtCodPrestamo.SetFocus
   Exit Sub
End If

'DESDE 2010 QUE NO USO MAS PLANES AL REFINANCIAR
'lo dejo en 1 para que despues muestre el credito en consultas
IdPlan = 1

'esto no funciona bien para cuotas distintas
If CLng(TxtCantidadCuotas.Text) > 1 Then
   ImporteCuota = CCur(lvcuotas.ListItems.Item(2).SubItems(10))
Else
   ImporteCuota = CCur(lvcuotas.ListItems.Item(1).SubItems(10))
End If

If Not MsgP("¿Confirma la Refinanciacion?") Then Exit Sub

'otras validaciones
'valido el cliente
If Not ExisteCliente(IdCliente) Then
   MsgE "El cliente no existe"
   Exit Sub
End If
'valido el credito
If ExisteCredito(IdCredito) Then
   MsgE "El numero de credito ya existe...(Debe ingresar uno nuevo)"
   TxtNumCredito.SetFocus
   Exit Sub
End If
'valido el numero de prestamo
If ExisteCodPrestamo(TxtCodPrestamo.Text) Then
   MsgE "El codigo de prestamo ya existe...(Debe ingresar uno nuevo)"
   TxtCodPrestamo.SetFocus
   Exit Sub
End If

IdProvincia = CLng(ComboProvincias.ItemData(ComboProvincias.ListIndex))

'inicio la transaccion
cnSQL.BeginTrans

'agrego un nuevo credito
'a partir de ahora a los refinanciados le pongo una marca en el campo [logic1]
'para diferenciarlos despues en consultas
sql = "insert into creditos (idcredito,idplan,idcliente," & _
      "tasa,importeafinanciar,importefinanciado,importegastos,importeseguros,importesellados," & _
      "ivainteres,ivaseguros,ivaotgastos,importerefinanciacion,numcuotas,fechacredito,observaciones," & _
      "formula,fechadesembolso,importetotal,importecuota,fechavencimiento1,codprestamo,idprovincia,logic1,motivobloqueo,cad1) " & _
      "values(" & CLng(IdCredito) & "," & CLng(IdPlan) & "," & CLng(IdCliente) & _
      "," & ConvertirDblSql(CDbl(TxtTasaFinanciacion.Text)) & _
      "," & ConvertirDblSql(TxtImporteAFinanciar.Text) & _
      "," & ConvertirDblSql(TxtImporteFinanciado.Text) & _
      "," & ConvertirDblSql(TxtTotalGastos.Text) & _
      "," & ConvertirDblSql(TxtTotalSeguros.Text) & _
      "," & ConvertirDblSql(TxtTotalSellados.Text) & "," & ConvertirDblSql(TxtTotalIvaInteres.Text) & _
      "," & ConvertirDblSql(TxtTotalIvaSeguros.Text) & "," & ConvertirDblSql(TxtTotalIvaGastos.Text) & "," & ConvertirDblSql(TxtImporteRecargo.Text) & _
      "," & CLng(TxtCantidadCuotas.Text) & _
      ",'" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & _
      "','" & CStr(TxtObservaciones.Text) & _
      "','" & CStr(Formula) & "'," & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & _
      "," & ConvertirDblSql(TxtImporteTotal.Text) & "," & ConvertirDblSql(ImporteCuota) & _
      ",'" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "','" & CStr(TxtCodPrestamo.Text) & "'," & CLng(IdProvincia) & ",1,'" & CStr(ComboComercios.Text) & "','" & CStr(VG_USUARIOLOGIN) & "')"
cnSQL.Execute sql

'recorro la lista grabando las nuevas cuotas ya calculadas
For I = 1 To CLng(lvcuotas.ListItems.Count())
    'genero una nueva factura
    NumFactura2 = ObtenerNuevoCupon()
    NumFactura = UltimoId("numfactura", "cuotas") + 1
    
    If NumFactura2 > NumFactura Then
        NumFactura = NumFactura2
    End If
       
    'saco una fila y la grabo
    NumCuota = CLng(lvcuotas.ListItems.Item(I))
    ImporteAmortizacion = CCur(lvcuotas.ListItems.Item(I).SubItems(1))
    ImporteInteres = CCur(lvcuotas.ListItems.Item(I).SubItems(2))
    'iva sobre interes
    IvaInteres = CCur(lvcuotas.ListItems.Item(I).SubItems(3))
    ImporteSeguros = CCur(lvcuotas.ListItems.Item(I).SubItems(5))
    'iva sobre seguros
    IvaSeguros = CCur(lvcuotas.ListItems.Item(I).SubItems(6))
    ImporteGastos = CCur(lvcuotas.ListItems.Item(I).SubItems(8))
    'iva sobre otorgamiento y gastos
    IvaOtorgamientoGastos = CCur(lvcuotas.ListItems.Item(I).SubItems(9))

    'este es el vto1
    ImporteVencimiento1 = CCur(lvcuotas.ListItems.Item(I).SubItems(10))
    Vencimiento1 = CDate(lvcuotas.ListItems.Item(I).SubItems(11))
    ImporteVencimiento2 = CCur(lvcuotas.ListItems.Item(I).SubItems(12))
    Vencimiento2 = CDate(lvcuotas.ListItems.Item(I).SubItems(13))
    'este es el importe cuota de antes original capital + interes
    ImporteCuota = CCur(ImporteAmortizacion) + CCur(ImporteInteres)
    'este es el total de impuestos
    ImporteImpuestos = CCur(IvaInteres) + CCur(IvaSeguros) + CCur(IvaOtorgamientoGastos)

    'revcargo por refinanciacion
    ImporteRecargo = CCur(lvcuotas.ListItems.Item(I).SubItems(14))
    
    'grabo el periodo para luego agrupar un reporte por meses
    Periodo = Format(CStr(Year(Vencimiento1)), "0000") & Format(CStr(Month(Vencimiento1)), "00")
    ImporteRecargoVencimiento2 = CCur(ImporteVencimiento2) - CCur(ImporteVencimiento1)
           
    'genero codigo de barras de rapipago
    DiasRecargo = CDate(Vencimiento2) - CDate(Vencimiento1)
    
    'ahora le paso los nuevos datos de rapipago
    CodigoBarras = GenerarCodigoBarras(VG_NUMEMPRESA, IdCliente, NumFactura, ImporteVencimiento1, Vencimiento1, ImporteRecargoVencimiento2, DiasRecargo)
    
    'grabo la cuota
    sql = "insert into cuotas (numfactura,idcredito,numcuota,importecuota," & _
          "fechavencimiento1,fechavencimiento2,importegastos,importeseguros," & _
          "importeimpuestos,importesellados,importerecargovencimiento2,codigobarras," & _
          "importeamortizacion,importeinteres,importerefinanciacion,periodo," & _
          "importevencimiento1,importevencimiento2,otorgamiento,ivainteres,ivaseguros,ivaotorgamientogastos) " & _
          "values(" & CLng(NumFactura) & "," & CLng(IdCredito) & _
          "," & CLng(NumCuota) & "," & ConvertirDblSql(ImporteCuota) & _
          ",'" & ConvertirFechaSql(CDate(Vencimiento1), "DD/MM/YYYY") & "','" & ConvertirFechaSql(CDate(Vencimiento2), "DD/MM/YYYY") & "'" & _
          "," & ConvertirDblSql(ImporteGastos) & "," & ConvertirDblSql(ImporteSeguros) & _
          "," & ConvertirDblSql(ImporteImpuestos) & "," & ConvertirDblSql(ImporteSellados) & _
          "," & ConvertirDblSql(ImporteRecargoVencimiento2) & ",'" & CStr(CodigoBarras) & _
          "'," & ConvertirDblSql(ImporteAmortizacion) & "," & ConvertirDblSql(ImporteInteres) & _
          "," & ConvertirDblSql(ImporteRecargo) & ",'" & CStr(Periodo) & _
          "'," & ConvertirDblSql(ImporteVencimiento1) & "," & ConvertirDblSql(ImporteVencimiento2) & ",0," & ConvertirDblSql(IvaInteres) & "," & ConvertirDblSql(IvaSeguros) & "," & ConvertirDblSql(IvaOtorgamientoGastos) & ")"
    cnSQL.Execute sql
    NumFactura3 = NumFactura
Next I

'marco las facturas refinanciadas deshabilitandolas
'tambien les asocio el plan actual al campo [NUM1] para luego poder anular planes
'y restablecer todo a lo anterior
For I = 1 To CLng(LvDeudas.ListItems.Count())
    'si esta marcada
    If LvDeudas.ListItems.Item(I).Checked Then
       NumFactura = CLng(LvDeudas.ListItems.Item(I).SubItems(2))
       sql = "update cuotas " & _
             "set fecharefinanciacion='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "',num1=" & CLng(IdCredito) & _
             " where cuotas.numfactura=" & CLng(NumFactura)
       cnSQL.Execute sql
    End If
Next I

'si no quedan mas cuotas pendientes del credito refinanciado finalizo
IdCreditoAnterior = 0
If VG_FINALIZARAUTOMATICAMENTE Then
    For I = 1 To CLng(LvDeudas.ListItems.Count())
        If LvDeudas.ListItems.Item(I).Checked Then
            If IdCreditoAnterior <> CLng(LvDeudas.ListItems.Item(I).SubItems(1)) Then
                    IdCreditoAnterior = CLng(LvDeudas.ListItems.Item(I).SubItems(1))
                    If CuotasImpagas(IdCreditoAnterior) = 0 Then
                        'la fecha de finalizacion del credito anterior es igual a la fecha de refinanciacion
                        sql = "update creditos set fechafinalizacion='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' " & _
                              "where creditos.idcredito=" & CLng(IdCreditoAnterior)
                        cnSQL.Execute sql
                    End If
            End If
        End If
    Next I
End If

'grabo la tasa de financiacion que se uso
sql = "update configuracionsistema set tasafinanciacion=" & ConvertirDblSql(CDbl(TxtTasaFinanciacion.Text)) & ", ultimocupon = " & CLng(NumFactura3)
cnSQL.Execute sql

'fin de la transaccion
cnSQL.CommitTrans

'refresco la lista de deudas del cliente
If IdCliente > 0 Then
   Call CargarDeudasCliente(IdCliente)
End If

'borro algunos campos
TxtCuotasARefinanciar.Text = 0
TxtImporteAFinanciar.Text = 0
DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date
TxtImporteFinanciado.Text = 0
TxtImporteRecargo.Text = 0
TxtTotalGastos.Text = 0
TxtTotalSeguros.Text = 0
TxtTotalIvaInteres.Text = 0
TxtTotalIvaSeguros.Text = 0
TxtTotalIvaGastos.Text = 0
TxtCantidadCuotas.Text = 0
TxtImporteTotal.Text = 0
TxtObservaciones.Text = ""
lvcuotas.ListItems.Clear

MsgI "La Refinanciacion fue registrada"

CmdGenerar.Enabled = False
ChTodas.Enabled = False

Exit Sub
merror:
tratarerrores "Error grabando la refinanciacion"
End Sub
Private Function DatosNumericosOk() As Boolean
'valida los campos numericos para calcular el importe a pagar en cuota y total financiado
On Error GoTo merror

DatosNumericosOk = True

'valido el importe a financiar
If Trim(TxtImporteAFinanciar.Text) = "" Then
   DatosNumericosOk = False
   Exit Function
End If
If Not IsNumeric(TxtImporteAFinanciar.Text) Then
   DatosNumericosOk = False
   Exit Function
End If
If CCur(TxtImporteAFinanciar.Text) <= 0 Then
   DatosNumericosOk = False
   Exit Function
End If

'valido la tasa de financiacion (se selecciona con el plan)
If Trim(TxtTasaFinanciacion.Text) = "" Then
   DatosNumericosOk = False
   Exit Function
End If
If Not IsNumeric(TxtTasaFinanciacion.Text) Then
   DatosNumericosOk = False
   Exit Function
End If
If CDbl(TxtTasaFinanciacion.Text) < 0 Then
   DatosNumericosOk = False
   Exit Function
End If

'valido la tasa2 (TEM)solo en credimaco (viene con el plan)
If Trim(TxtTasa2.Text) = "" Then
   DatosNumericosOk = False
   Exit Function
End If
If Not IsNumeric(TxtTasa2.Text) Then
   DatosNumericosOk = False
   Exit Function
End If
If CDbl(TxtTasa2.Text) < 0 Then
   DatosNumericosOk = False
   Exit Function
End If

'valido la tasa de REfinanciacion
If Trim(TxtTasaRefinanciacion.Text) = "" Then
   TxtTasaRefinanciacion.Text = 0
End If

'valido la cantidad de cuotas
If Trim(TxtCantidadCuotas.Text) = "" Then
   DatosNumericosOk = False
   Exit Function
End If
If Not IsNumeric(TxtCantidadCuotas.Text) Then
   DatosNumericosOk = False
   Exit Function
End If
If CLng(TxtCantidadCuotas.Text) <= 0 Then
   DatosNumericosOk = False
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosNumericosOk-RefinanciarCreditos"
End Function
Private Sub CalcularImportes()
'calcula los importes cada vez que se actualizan los campos
Dim CapitalCuota As Currency
Dim InteresCuota As Currency
Dim ImporteCuota As Currency
Dim ImporteGastos As Currency
Dim ImporteSeguros As Currency
Dim ImporteImpuestos As Currency
Dim ImporteVencimiento1 As Currency
Dim ImporteVencimiento2 As Currency
Dim Vencimiento1 As Date
Dim Vencimiento2 As Date
Dim DiaVencimiento As Long
Dim DiferenciaVencimientos As Long
Dim Nitem As ListItem
Dim I As Long
Dim FechaProxima As Date
Dim CapitalVivo As Currency
Dim TasaMensual As Double
Dim TasaFinanciacion As Double
Dim Frecuencia As Long
Dim InteresAnual As Currency
Dim ImporteRecargo As Currency
Dim AcumGastos As Currency
Dim AcumSeguros As Currency
Dim AcumImpuestos As Currency
Dim importevto1 As Currency
Dim ImporteVto2 As Currency
Dim TasaFin As Double
Dim ImporteTotal As Currency
Dim Importe1 As Currency
Dim Importe2 As Currency
Dim AcumSellados As Currency
Dim ParteSellados As Currency
Dim ImporteCuotaUnico As Currency
Dim TasaTem As Double
Dim ImporteIva1 As Double
Dim Resto As Currency
Dim Subtotal1 As Currency
Dim IvaInteres As Currency
Dim IvaSeguros As Currency
Dim IvaGastos As Currency
Dim IvaOtorgamiento As Currency
Dim OtorgamientoGastos As Currency
Dim ImporteOtorgamiento As Currency
Dim CalcularImpuestos As Boolean
Dim IvaRecargoRefin As Currency
Dim CapitalTotal As Currency
Dim InteresTotal As Currency
Dim CapitalMasInteresTotal As Currency
On Error GoTo merror

'cada vez que entra hay cambios
'entonces blanqueo todos los campos
lvcuotas.ListItems.Clear
TxtTotalInteres.Text = 0
TxtTotalPTF.Text = 0
TxtTotalGastos.Text = 0
TxtTotalSeguros.Text = 0
TxtTotalIvaInteres.Text = 0
TxtTotalIvaSeguros.Text = 0
TxtTotalIvaGastos.Text = 0
TxtTotalOtorgamiento.Text = 0
TxtImporteFinanciado.Text = 0
TxtImporteTotal.Text = 0
TxtImporteRecargo.Text = 0

ImporteCuota = 0
ImporteGastos = 0
ImporteSeguros = 0
ImporteImpuestos = 0
ImporteRecargo = 0
ImporteVencimiento1 = 0
ImporteVencimiento2 = 0
ImporteSellados = 0
importevto1 = 0
ImporteVto2 = 0
AcumGastos = 0
AcumSeguros = 0
AcumImpuestos = 0
Resto = 0

If Not DatosNumericosOk() Then Exit Sub

TasaFinanciacion = CDbl(TxtTasaFinanciacion.Text)

'para credimaco..esta tasa se calcula automatica al ingresar la tasa anterior
TasaTem = CDbl(TxtTasa2.Text)

'obtengo el primer vencimiento
Vencimiento1 = CDate(DTPicker2.Value)
'por defecto los dos vencimientos son iguales
Vencimiento2 = Vencimiento1

DiferenciaVencimientos = 0
'trata que venza siempre el 20 por ejemplo
DiaVencimiento = Day(DTPicker2.Value)

If VG_APLICARSEGUNDOVENCIMIENTO Then
   Vencimiento2 = CDate(DTPicker3.Value)
   DiferenciaVencimientos = DateDiff("d", Vencimiento1, Vencimiento2)
End If

'recargo por refinanciacion
'calculo la comision por refinanciacion de acuerdo al importe a refinanciar
TxtImporteRecargo.Text = Format(CalcularRecargoRefinanciacion(TxtImporteAFinanciar.Text), "0.00")

'importe de recargo por cuota
ImporteRecargo = CCur(CCur(TxtImporteRecargo.Text) / CLng(TxtCantidadCuotas.Text))
   
'aca toma de partida desde el subtotal que incluye desc y recargos
'antes partia del importeafinanciar
   
'esto devuelve el importe total de la cuota incluyendo capital e interes
ImporteCuota = Pmt(TasaTem / 100, TxtCantidadCuotas.Text, -(TxtSubtotal.Text))
TxtImporteFinanciado.Text = CDbl(ImporteCuota) * CLng(TxtCantidadCuotas.Text)
CapitalCuota = CCur(TxtSubtotal.Text) / CLng(TxtCantidadCuotas.Text)
InteresCuota = CCur(ImporteCuota) - CCur(CapitalCuota)

If ChkDDJJ.Value Then
    InteresCuota = InteresCuota * (1 + CDbl(VG_PORCENTAJEIVA) / 100)
End If

CapitalTotal = CCur(TxtSubtotal.Text)
InteresTotal = CCur(Format$(CCur(ImporteCuota) - CCur(CapitalCuota), "0.00")) * CLng(TxtCantidadCuotas.Text)
CapitalMasInteresTotal = CCur(TxtImporteFinanciado.Text)
  
'desde aca se parte..despues le va restando la cuota neta
'sirve para calcular el seguro de cada cuota
Resto = CCur(TxtImporteFinanciado.Text)
   
For I = 1 To CLng(TxtCantidadCuotas.Text)
    ImporteVencimiento1 = CCur(ImporteCuota)
    Set Nitem = lvcuotas.ListItems.Add(, , Format(I, "00"))

    'si aplico gastos
    ImporteGastos = 0
    If VG_APLICARGASTOS Then
       'si aplico a las refinanciaciones
       If Not VG_NOAPLICARGASTOSREFINANCIACION Then
          'si va de la 2 en adelante
          If VG_APLICARGASTOSCUOTA2 Then
             'solo de 2 en adelante
             If I > 1 Then
                If CCur(VG_IMPORTEGASTOSFIJOS) > 0 Then
                   ImporteGastos = CCur(VG_IMPORTEGASTOSFIJOS / 10)
                End If
                'si es un importe dividido por la cantidad de cuotas
                If CCur(VG_IMPORTEGASTOS) > 0 Then
                   ImporteGastos = CCur(VG_IMPORTEGASTOS / 10) / (CLng(TxtCantidadCuotas.Text) - 1)
                End If
                If VG_PORCCAPNOINT > 0 Then
                   ImporteGastos = CapitalTotal * (VG_PORCCAPNOINT / 100) / (CLng(TxtCantidadCuotas.Text) - 1)
                End If
                If VG_PORCFUNNOCAP > 0 Then
                   ImporteGastos = InteresTotal * (VG_PORCFUNNOCAP / 100) / (CLng(TxtCantidadCuotas.Text) - 1)
                End If
                If VG_PORCCAPINT > 0 Then
                   ImporteGastos = CapitalMasInteresTotal * (VG_PORCCAPINT / 100) / (CLng(TxtCantidadCuotas.Text) - 1)
                End If
             End If
          Else
             'le aplico a todas o solo a la primera
             If VG_APLICARGASTOSCUOTA1 Then
                If I = 1 Then
                   If CCur(VG_IMPORTEGASTOSFIJOS) > 0 Then
                      ImporteGastos = CCur(VG_IMPORTEGASTOSFIJOS / 10)
                   End If
                   'si es un importe dividido por la cantidad de cuotas
                   If CCur(VG_IMPORTEGASTOS) > 0 Then
                      ImporteGastos = CCur(VG_IMPORTEGASTOS / 10)
                   End If
                   If VG_PORCCAPNOINT > 0 Then
                      ImporteGastos = CapitalTotal * VG_PORCCAPNOINT / 100
                   End If
                   If VG_PORCFUNNOCAP > 0 Then
                      ImporteGastos = InteresTotal * VG_PORCFUNNOCAP / 100
                   End If
                   If VG_PORCCAPINT > 0 Then
                      ImporteGastos = CapitalMasInteresTotal * VG_PORCCAPINT / 100
                   End If
                End If
             Else
                'aplico a todas
                If CCur(VG_IMPORTEGASTOSFIJOS) > 0 Then
                   ImporteGastos = CCur(VG_IMPORTEGASTOSFIJOS / 10)
                End If
                If CCur(VG_IMPORTEGASTOS) > 0 Then
                   ImporteGastos = CCur(VG_IMPORTEGASTOS / 10) / CLng(TxtCantidadCuotas.Text)
                End If
                If VG_PORCCAPNOINT > 0 Then
                   ImporteGastos = CapitalTotal * (VG_PORCCAPNOINT / 100) / CLng(TxtCantidadCuotas.Text)
                End If
                If VG_PORCFUNNOCAP > 0 Then
                   ImporteGastos = InteresTotal * (VG_PORCFUNNOCAP / 100) / CLng(TxtCantidadCuotas.Text)
                End If
                If VG_PORCCAPINT > 0 Then
                   ImporteGastos = CapitalMasInteresTotal * (VG_PORCCAPINT / 100) / CLng(TxtCantidadCuotas.Text)
                End If
             End If 'si aplico a cuota1
          End If ' si aplico a cuota2
       End If ' si aplico a refinanciacion
    End If ' si aplico gastos
    
    Nitem.SubItems(8) = Format(ImporteGastos, "0.00") & vbNullString
        
    'si aplico seguros
    ImporteSeguros = 0
    If VG_APLICARSEGURO Then
       'si aplico seguros a las refinanciaciones
       If Not VG_NOAPLICARSEGUROSREFINANCIACION Then
          'si aplico a todas las cuotas
          If Not VG_APLICARSEGUROSCUOTA1 Then
             If CCur(VG_ALICUOTASEGUROS) > 0 Then
                'seguros
                ImporteSeguros = (VG_ALICUOTASEGUROS * CCur(Resto) / 100) / CDbl(1.21)
                'disminuyo el resto
                Resto = CCur(Resto) - CCur(ImporteCuota)
                'el resto sirve para seguir calculando el seguro a las demas cuotas
             End If
             If CCur(VG_SEGUROFIJO) > 0 Then
                ImporteSeguros = CCur(ImporteSeguros) + CCur(VG_SEGUROFIJO)
             End If
             If CCur(VG_IMPORTESEGURO) > 0 Then
                ImporteSeguros = CCur(ImporteSeguros) + (CCur(VG_IMPORTESEGURO) / CLng(TxtCantidadCuotas.Text))
             End If
          Else
             'solo a la cuota 1 el seguro que le corresponde
             If I = 1 Then
                If CCur(VG_ALICUOTASEGUROS) > 0 Then
                   ImporteSeguros = (VG_ALICUOTASEGUROS * CCur(Resto) / 100) / CDbl(1.21)
                End If
                If CCur(VG_SEGUROFIJO) > 0 Then
                   ImporteSeguros = CCur(ImporteSeguros) + CCur(VG_SEGUROFIJO)
                End If
                If CCur(VG_IMPORTESEGURO) > 0 Then
                   ImporteSeguros = CCur(ImporteSeguros) + (CCur(VG_IMPORTESEGURO) / CLng(TxtCantidadCuotas.Text))
                End If
             End If 'si i=1
          End If 'si aplico a cuota1
       End If 'si aplico a refin
    End If
        
    'capital
    Nitem.SubItems(1) = Format(CapitalCuota, "0.00") & vbNullString
    'interes
    Nitem.SubItems(2) = Format(InteresCuota, "0.00") & vbNullString
    'seguros
    Nitem.SubItems(5) = Format(ImporteSeguros, "0.00") & vbNullString
    
    'si aplico otorgamiento
    ImporteOtorgamiento = 0
    If VG_APLICAROTORGAMIENTO Then
       'si aplico otorgamiento a las refinanciaciones
       If Not VG_NOAPLICAROTREFIN Then
          'aplica los gastos de otorgamoento a la primer cuota
          If VG_APLICAROTORGAMIENTOCUOTA1 Then
             If I = 1 Then
                If CCur(VG_IMPORTEOTORGAMIENTO) > 0 Then
                    ImporteOtorgamiento = CCur(VG_IMPORTEOTORGAMIENTO)
                End If
                If VG_OTORCAPNOINT > 0 Then
                    ImporteOtorgamiento = CapitalTotal * VG_OTORCAPNOINT / 100
                End If
                If VG_OTORINTNOCAP > 0 Then
                    ImporteOtorgamiento = InteresTotal * VG_OTORINTNOCAP / 100
                End If
                If VG_OTORCAPMASINT > 0 Then
                    ImporteOtorgamiento = CapitalMasInteresTotal * VG_OTORCAPMASINT / 100
                End If
             End If
          Else
             'aplica a todas entonces lo divido en partes
             If CCur(VG_IMPORTEOTORGAMIENTO) > 0 Then
                ImporteOtorgamiento = CCur(VG_IMPORTEOTORGAMIENTO) / CLng(TxtCantidadCuotas.Text)
             End If
             If VG_OTORCAPNOINT > 0 Then
                ImporteOtorgamiento = (CapitalTotal * VG_OTORCAPNOINT / 100) / CLng(TxtCantidadCuotas.Text)
            End If
            If VG_OTORINTNOCAP > 0 Then
                ImporteOtorgamiento = (InteresTotal * VG_OTORINTNOCAP / 100) / CLng(TxtCantidadCuotas.Text)
            End If
            If VG_OTORCAPMASINT > 0 Then
                ImporteOtorgamiento = (CapitalMasInteresTotal * VG_OTORCAPMASINT / 100) / CLng(TxtCantidadCuotas.Text)
            End If
          End If
       End If
    End If
    Nitem.SubItems(7) = Format(ImporteOtorgamiento, "0.00") & vbNullString
  
    CalcularImpuestos = False
    'si aplico impuestos calculo el iva de otorgamiento
    If VG_APLICARIMPUESTOS Then
       'si aplico a las refinanciaciones
       If Not VG_NOAPLICARIMPUESTOSREFINANCIACION Then
          'si aplico a la cuota 1
          If VG_APLICARIMPUESTOSCUOTA1 Then
             'solo calculo impuestos a la cuota 1
             If I = 1 Then
                CalcularImpuestos = True
             End If 'firn cuota=1
          Else
             'aplico a todas o de la segunda en adelante
             If VG_APLICARIMPUESTOSCUOTA2 Then
                'solo calculo impuestos de cuota 2 en adelante
                If I > 1 Then
                   CalcularImpuestos = True
                End If 'fin si cuota>1
             Else
               'aplico a todas
               CalcularImpuestos = True
             End If 'fin cuota2
       
          End If 'fin VG_CUOTA1
       End If
    End If 'si aplico impuestos
    
    If CalcularImpuestos Then
       'si calculo segun credimaco
       If VG_IMPUESTOSCREDIMACO Then
          'reemplace el 21% por la variable global
          IvaOtorgamiento = CDbl(VG_PORCENTAJEIVA) * CCur(ImporteOtorgamiento / 100)
          IvaGastos = CDbl(VG_PORCENTAJEIVA) * CCur(ImporteGastos / 100)
          IvaSeguros = CCur(VG_PORCENTAJEIVA * ImporteSeguros / 100)
          If ChkDDJJ.Value Then
            IvaInteres = 0
          Else
            IvaInteres = CCur(VG_PORCENTAJEIVA * InteresCuota / 100)
          End If
       End If
       'si es un importe fijo
       If CCur(VG_IMPUESTOSFIJOS) > 0 Then
          'si es fijo solo asigno al iva interes
          If ChkDDJJ.Value Then
            IvaInteres = 0
          Else
            IvaInteres = CCur(VG_IMPUESTOSFIJOS)
          End If
          IvaOtorgamiento = 0
          IvaGastos = 0
          IvaSeguros = 0
       End If
       'si es un importe a dividir entre las cuotas
       If CCur(VG_IMPORTEIMPUESTOS) > 0 Then
          'si es fijo solo asigno al iva interes
          If ChkDDJJ.Value Then
            IvaInteres = 0
          Else
            IvaInteres = CCur(VG_IMPORTEIMPUESTOS) + CLng(TxtCantidadCuotas.Text)
          End If
          IvaOtorgamiento = 0
          IvaGastos = 0
          IvaSeguros = 0
       End If
    Else
       IvaOtorgamiento = 0
       IvaGastos = 0
       IvaSeguros = 0
       IvaInteres = 0
    End If
 
    'iva sobre el interes
    Nitem.SubItems(3) = Format(IvaInteres, "0.00") & vbNullString

    'iva seguros
    Nitem.SubItems(6) = Format(IvaSeguros, "0.00") & vbNullString
    
    'primer subtotal de los tres anteriores
    Subtotal1 = CCur(CapitalCuota) + CCur(InteresCuota) + CCur(IvaInteres)
    Nitem.SubItems(4) = Format(Subtotal1, "0.00") & vbNullString
    
    IvaRecargoRefin = 0
    If VG_APLICARIMPUESTOS Then
       'si aplico impuestos a la refin
       If Not VG_NOAPLICARIMPUESTOSREFINANCIACION Then
          If CCur(ImporteRecargo) > 0 Then
             IvaRecargoRefin = (VG_PORCENTAJEIVA * CCur(ImporteRecargo)) / 100
          End If
       End If
    End If
    'este es el iva recargo por refinanciacion
    Nitem.SubItems(15) = Format(IvaRecargoRefin, "0.00") & vbNullString
    
    'integro el nuevo iva al iva gastos
    IvaGastos = CCur(IvaGastos) + CCur(IvaRecargoRefin)
    Nitem.SubItems(9) = Format(IvaOtorgamiento + IvaGastos, "0.00") & vbNullString
        
    'aca tambien le sumo el recargo de refinanciacion
    ImporteVencimiento1 = CCur(Subtotal1) + CCur(ImporteOtorgamiento) + CCur(ImporteGastos) + CCur(ImporteSeguros) + CCur(IvaSeguros) + CCur(IvaGastos) + CCur(IvaOtorgamiento) + CCur(ImporteRecargo)
    'ahora el importe1 tiene gastos etc.
    Nitem.SubItems(10) = Format(ImporteVencimiento1, "0.00") & vbNullString
    
    'primer vencimiento
    Nitem.SubItems(11) = Vencimiento1
    
    'recalculo el importe del vencimiento2
    ImporteVencimiento2 = CalcularImporteVencimiento2(ImporteVencimiento1, Vencimiento1, Vencimiento2)
    Nitem.SubItems(12) = Format(ImporteVencimiento2, "0.00")
    Nitem.SubItems(13) = Vencimiento2
    'aca esta el recargo por refinanciacion
    Nitem.SubItems(14) = Format(ImporteRecargo, "0.00") & vbNullString
      
    'muevo los vencimientos
    If VG_DIASVENCIMIENTOREFINANCIACION = 30 Or VG_DIASVENCIMIENTOREFINANCIACION = 60 Then
       Frecuencia = CLng(VG_DIASVENCIMIENTOREFINANCIACION / 30)
       'incremento al primer vencimiento por meses exactos validos
       FechaProxima = ArmarFecha(DiaVencimiento, Vencimiento1, Frecuencia)
    Else
       'aplico vencimiento por dias
       FechaProxima = Vencimiento1 + VG_DIASVENCIMIENTOREFINANCIACION
    End If
    'muevo el vencimiento1
    Vencimiento1 = ObtenerFechaVencimiento(FechaProxima, VG_DIASVENCIMIENTOREFINANCIACION)
    Vencimiento2 = Vencimiento1
    
    If VG_APLICARSEGUNDOVENCIMIENTO Then
       'obtengo el segundo vencimiento respetando la diferencia inicial
       Vencimiento2 = ObtenerFechaVencimiento(Vencimiento1 + DiferenciaVencimientos, VG_DIASVENCIMIENTOREFINANCIACION)
    End If
    
Next I
   
'obtengo totales
For I = 1 To lvcuotas.ListItems.Count
    TxtTotalInteres.Text = CCur(TxtTotalInteres.Text) + CCur(lvcuotas.ListItems.Item(I).SubItems(2))
    TxtTotalGastos.Text = CCur(TxtTotalGastos.Text) + CCur(lvcuotas.ListItems.Item(I).SubItems(8))
    TxtTotalSeguros.Text = CCur(TxtTotalSeguros.Text) + CCur(lvcuotas.ListItems.Item(I).SubItems(5))
    TxtTotalIvaInteres.Text = CCur(TxtTotalIvaInteres.Text) + CCur(lvcuotas.ListItems.Item(I).SubItems(3))
    TxtTotalIvaSeguros.Text = CCur(TxtTotalIvaSeguros.Text) + CCur(lvcuotas.ListItems.Item(I).SubItems(6))
    TxtTotalIvaGastos.Text = CCur(TxtTotalIvaGastos.Text) + CCur(lvcuotas.ListItems.Item(I).SubItems(9))
    TxtTotalOtorgamiento.Text = CCur(TxtTotalOtorgamiento.Text) + CCur(lvcuotas.ListItems.Item(I).SubItems(7))
    TxtImporteTotal.Text = CCur(TxtImporteTotal.Text) + CCur(lvcuotas.ListItems.Item(I).SubItems(10))
Next I

'esto es capital + interes + seguros + otorgamiento + gastos
TxtTotalPTF.Text = CCur(TxtImporteAFinanciar.Text) + CCur(TxtTotalInteres.Text) + CCur(TxtTotalSeguros.Text) + CCur(TxtTotalOtorgamiento.Text) + CCur(TxtTotalGastos.Text)
TxtTotalPTF.Text = Format(TxtTotalPTF.Text, "0.00")

'formateo los totales
TxtImporteFinanciado.Text = Format(TxtImporteFinanciado.Text, "0.00")
TxtTotalInteres.Text = Format(TxtTotalInteres.Text, "0.00")
TxtTotalGastos.Text = Format(TxtTotalGastos.Text, "0.00")
TxtTotalSeguros.Text = Format(TxtTotalSeguros.Text, "0.00")
TxtTotalIvaInteres.Text = Format(TxtTotalIvaInteres.Text, "0.00")
TxtTotalIvaSeguros.Text = Format(TxtTotalIvaSeguros.Text, "0.00")
TxtTotalIvaGastos.Text = Format(TxtTotalIvaGastos.Text, "0.00")
TxtTotalOtorgamiento.Text = Format(TxtTotalOtorgamiento.Text, "0.00")
TxtImporteTotal.Text = Format(TxtImporteTotal.Text, "0.00")

'si redondeo lo aplico
If VG_REDONDEAR Then
   Call Redondear
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento CalcularImportes-RefinanciarCreditos"
End Sub
Private Sub LvDeudas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'se ejecuta al marcar/desmarcar una fila de la lista de deudas
On Error GoTo merror

'si la fila esta tildada
If LvDeudas.ListItems.Item(Item.Index).Checked Then
   'importe original de la fila
   TxtImporteAFinanciar.Text = CCur(TxtImporteAFinanciar.Text) + CCur(LvDeudas.ListItems.Item(Item.Index).SubItems(8))
   TxtCuotasARefinanciar.Text = CLng(TxtCuotasARefinanciar.Text) + 1
Else
   TxtImporteAFinanciar.Text = CCur(TxtImporteAFinanciar.Text) - CCur(LvDeudas.ListItems.Item(Item.Index).SubItems(8))
   TxtCuotasARefinanciar.Text = CLng(TxtCuotasARefinanciar.Text) - 1
End If

TxtImporteAFinanciar.Text = Format(TxtImporteAFinanciar.Text, "0.00")

If CheckDescuentos.Value = 1 Then
   If IsNumeric(TxtDescuento.Text) Then
      If IsNumeric(TxtImporteAFinanciar.Text) Then
         If CCur(TxtImporteAFinanciar.Text) > 0 Then
            If CCur(TxtDescuento.Text) < CCur(TxtImporteAFinanciar.Text) Then
               TxtSubtotal.Text = CCur(TxtImporteAFinanciar.Text) - CCur(TxtDescuento.Text)
            End If
         End If
      End If
   End If
End If

If CheckRecargos.Value = 1 Then
   If IsNumeric(TxtRecargo.Text) Then
      If IsNumeric(TxtImporteAFinanciar.Text) Then
         If CCur(TxtImporteAFinanciar.Text) > 0 Then
            TxtSubtotal.Text = CCur(TxtImporteAFinanciar.Text) + CCur(TxtRecargo.Text)
         End If
      End If
   End If
End If

'actualizo si hubo cambios en las cuotas seleccionadas
Call CalcularImportes
'ChTodas = False
Exit Sub
merror:
tratarerrores "Error seleccionando cuota a refinanciar"
End Sub
Private Sub TxtCliente_Change()

If Trim(TxtCliente.Text) <> "" Then
   Call CargarDeudasCliente(IdCliente)
   CmdGenerar.Enabled = True
   ChTodas.Enabled = True
End If

End Sub
Private Sub DTPicker2_Change()
'si cambio la fecha de vencimiento1
DTPicker3.Value = DTPicker2.Value
Call CalcularImportes
End Sub
Private Sub DTPicker3_Change()
'si cambia la fecha de vencimiento2
Call CalcularImportes
End Sub
Private Sub TxtCantidadCuotas_Change()
Call CalcularImportes
End Sub
Private Sub txtcantidadcuotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Call CalcularImportes
End If
End Sub
Private Sub TxtCantidadCuotas_LostFocus()
Call CalcularImportes
End Sub
Private Sub TxtCodPrestamo_LostFocus()
TxtCodPrestamo.Text = UCase(Trim(TxtCodPrestamo.Text))
End Sub
Private Sub TxtComercio_LostFocus()
TxtComercio.Text = UCase(Trim(TxtComercio.Text))
End Sub
Private Sub TxtImporteAFinanciar_Change()
Call CalcularSellados
TxtSubtotal.Text = TxtImporteAFinanciar.Text
End Sub
Private Sub TxtObservaciones_LostFocus()
TxtObservaciones.Text = UCase(Trim(TxtObservaciones.Text))
End Sub
Private Sub TxtTasaFinanciacion_Change()
Call CalcularTem
Call CalcularImportes
End Sub
Private Sub CalcularTem()
'calcula la tasa TEM y la aplica cuando cambia la tasa1 tna y cuando cambia
'el tipo de vencimiento mensual, etc
'(*)en las refinanciaciones el vencimiento siempre es mensual
'no se puede cambiar como al registrar creditos.
Dim Dias As Long
Dim Tem As Double
On Error GoTo merror

If Trim(TxtTasaFinanciacion.Text) = "" Then Exit Sub
If Not IsNumeric(TxtTasaFinanciacion.Text) Then Exit Sub
If CDbl(TxtTasaFinanciacion.Text) < 0 Then Exit Sub

TxtTasa2.Text = 0

'si es mensual
If VG_DIASVENCIMIENTOREFINANCIACION = 30 Then
   TxtTasa2.Text = CDbl(TxtTasaFinanciacion.Text) / 365 * 30
End If
'diario
If VG_DIASVENCIMIENTOREFINANCIACION = 1 Then
   TxtTasa2.Text = CDbl(TxtTasaFinanciacion.Text) / 365 * 1
End If
'semanal
If VG_DIASVENCIMIENTOREFINANCIACION = 7 Then
   TxtTasa2.Text = CDbl(TxtTasaFinanciacion.Text) / 365 * 7
End If
'quincenal
If VG_DIASVENCIMIENTOREFINANCIACION = 15 Then
   TxtTasa2.Text = CDbl(TxtTasaFinanciacion.Text) / 365 * 15
End If
'bimestral
If VG_DIASVENCIMIENTOREFINANCIACION = 60 Then
   TxtTasa2.Text = CDbl(TxtTasaFinanciacion.Text) / 365 * 60
End If

TxtTasa2.Text = Format(TxtTasa2.Text, "0.00")

Exit Sub
merror:
tratarerrores "Error calculando Tem"
End Sub
Private Sub Redondear()
'redondea los importes
Dim I As Long
Dim Capital As Currency
Dim Interes As Long
Dim Gastos As Currency
Dim Impuestos As Currency
Dim Seguros As Currency
Dim Sellados As Currency
Dim ImporteRefinanciacion As Currency
Dim Vencimiento1 As Date
Dim Vencimiento2 As Date
Dim ImporteCuota As Currency
Dim ImporteVencimiento1 As Currency
Dim ImporteVencimiento2 As Currency
Dim ImporteTotalGastos As Currency
Dim ImporteTotalSeguros As Currency
Dim ImporteTotalOtorgamiento As Currency
Dim ImporteTotalRefin As Currency
Dim ImporteTotal As Currency
Dim ImporteFinanciado As Currency
Dim IvaInteres As Currency
Dim Subtotal As Currency
Dim IvaSeguros As Currency
Dim Otorgamiento As Currency
Dim IvaGastos As Currency
Dim ImporteTotalIvaInteres As Currency
Dim ImporteTotalIvaSeguros As Currency
Dim ImporteTotalIvaGastos As Currency
On Error GoTo merror

ImporteFinanciado = 0
ImporteTotalGastos = 0
ImporteTotalSeguros = 0
ImporteTotalIvaInteres = 0
ImporteTotalIvaSeguros = 0
ImporteTotalIvaGastos = 0
ImporteTotalOtorgamiento = 0
ImporteTotal = 0

For I = 1 To lvcuotas.ListItems.Count
    'redondeo el capital
    Capital = CCur(lvcuotas.ListItems.Item(I).SubItems(1))
    Capital = Round(Capital)
    lvcuotas.ListItems.Item(I).SubItems(1) = Format(Capital, "0.00")
    
    'redondeo el interes
    Interes = CCur(lvcuotas.ListItems.Item(I).SubItems(2))
    Interes = Round(Interes)
    lvcuotas.ListItems.Item(I).SubItems(2) = Format(Interes, "0.00")
    
    'redondeo el ivainteres
    IvaInteres = CCur(lvcuotas.ListItems.Item(I).SubItems(3))
    IvaInteres = Round(IvaInteres)
    lvcuotas.ListItems.Item(I).SubItems(3) = Format(IvaInteres, "0.00")
        
    'subtotal credimaco..las partes ya estan redondeadas
    Subtotal = CCur(Capital) + CCur(Interes) + CCur(IvaInteres)
    lvcuotas.ListItems.Item(I).SubItems(4) = Format(Subtotal, "0.00")
    
    'importecuota ya esta redondeado por sus partes
    ImporteCuota = CCur(Capital) + CCur(Interes)
        
    'redondeo seguros
    Seguros = CCur(lvcuotas.ListItems.Item(I).SubItems(5))
    Seguros = Round(Seguros)
    lvcuotas.ListItems.Item(I).SubItems(5) = Format(Seguros, "0.00")
    
    'redondeo iva seguros
    IvaSeguros = CCur(lvcuotas.ListItems.Item(I).SubItems(6))
    IvaSeguros = Round(IvaSeguros)
    lvcuotas.ListItems.Item(I).SubItems(6) = Format(IvaSeguros, "0.00")
        
    'otorgamiento
    Otorgamiento = CCur(lvcuotas.ListItems.Item(I).SubItems(7))
    Otorgamiento = Round(Otorgamiento)
    lvcuotas.ListItems.Item(I).SubItems(7) = Format(Otorgamiento, "0.00")

    'gastos
    Gastos = CCur(lvcuotas.ListItems.Item(I).SubItems(8))
    Gastos = Round(Gastos)
    lvcuotas.ListItems.Item(I).SubItems(8) = Format(Gastos, "0.00")
    
    'ivagastos/otorgam
    IvaGastos = CCur(lvcuotas.ListItems.Item(I).SubItems(9))
    IvaGastos = Round(IvaGastos)
    lvcuotas.ListItems.Item(I).SubItems(9) = Format(IvaGastos, "0.00")
    
    'recargo refinanciacion
    ImporteRefinanciacion = CCur(lvcuotas.ListItems.Item(I).SubItems(14))
    ImporteRefinanciacion = Round(ImporteRefinanciacion)
    lvcuotas.ListItems.Item(I).SubItems(14) = Format(ImporteRefinanciacion, "0.00")
    
    'importe al primer vencimiento
    ImporteVencimiento1 = CCur(Subtotal) + CCur(Seguros) + CCur(IvaSeguros) + CCur(Otorgamiento) + CCur(Gastos) + CCur(IvaGastos) + CCur(ImporteRefinanciacion)
    lvcuotas.ListItems.Item(I).SubItems(10) = Format(ImporteVencimiento1, "0.00")
    
    'importe vencimiento2
    Vencimiento1 = CDate(lvcuotas.ListItems.Item(I).SubItems(11))
    Vencimiento2 = CDate(lvcuotas.ListItems.Item(I).SubItems(13))
    'lo calculo en base al importe 1 nuevo redondeado
    ImporteVencimiento2 = CalcularImporteVencimiento2(ImporteVencimiento1, Vencimiento1, Vencimiento2)
    ImporteVencimiento2 = Round(ImporteVencimiento2)
    lvcuotas.ListItems.Item(I).SubItems(12) = Format(ImporteVencimiento2, "0.00")
        
    'totales
    ImporteFinanciado = CCur(ImporteFinanciado) + CCur(ImporteCuota)
    ImporteTotalRefin = CCur(ImporteTotalRefin) + CCur(ImporteRefinanciacion)
    ImporteTotalGastos = CCur(ImporteTotalGastos) + CCur(Gastos)
    ImporteTotalSeguros = CCur(ImporteTotalSeguros) + CCur(Seguros)
    
    ImporteTotalIvaInteres = CCur(ImporteTotalIvaInteres) + CCur(IvaInteres)
    ImporteTotalIvaSeguros = CCur(ImporteTotalIvaSeguros) + CCur(IvaSeguros)
    ImporteTotalIvaGastos = CCur(ImporteTotalIvaGastos) + CCur(IvaGastos)
    
    ImporteTotalOtorgamiento = CCur(ImporteTotalOtorgamiento) + CCur(Otorgamiento)
    ImporteTotal = CCur(ImporteTotal) + CCur(ImporteVencimiento1)
Next I

TxtImporteFinanciado.Text = Format(ImporteFinanciado, "0.00")
TxtImporteRecargo.Text = Format(ImporteTotalRefin, "0.00")
TxtTotalGastos.Text = Format(ImporteTotalGastos, "0.00")
TxtTotalSeguros.Text = Format(ImporteTotalSeguros, "0.00")

TxtTotalIvaInteres.Text = Format(ImporteTotalIvaInteres, "0.00")
TxtTotalIvaSeguros.Text = Format(ImporteTotalIvaSeguros, "0.00")
TxtTotalIvaGastos.Text = Format(ImporteTotalIvaGastos, "0.00")

TxtImporteTotal.Text = Format(ImporteTotal, "0.00")

Exit Sub
merror:
tratarerrores "Error en funcion Redondear"
End Sub
Private Sub ComboProvincias_Click()
Call CalcularSellados
'si viene de comercios no haga nada
If BandComercio = False Then
    Call CargarComboWhere("comercios", ComboComercios, ComboProvincias)
End If
End Sub
Private Sub CalcularSellados()
'calcula el sellado de la provincia seleccionada
Dim IdProvincia As Long
Dim Porcentaje As Double
Dim Resultado As Currency
On Error GoTo merror

If ComboProvincias.Text = "" Then Exit Sub

If Trim(TxtImporteAFinanciar.Text) = "" Then Exit Sub
If Not IsNumeric(TxtImporteAFinanciar.Text) Then Exit Sub

IdProvincia = CLng(ComboProvincias.ItemData(ComboProvincias.ListIndex))

Porcentaje = ObtenerPorcentajeSellados(IdProvincia)
Resultado = ObtenerImporteSellados(Porcentaje, TxtImporteAFinanciar.Text)

TxtTotalSellados.Text = Format(Resultado, "0.00")

Exit Sub
merror:
tratarerrores "Error en funcion CalcularSellados"
End Sub
Private Function HayCreditosDistintos(ByVal lv As ListView) As Boolean
'verifico si en una lista se seleccionaron cuotas de distintos creditos
Dim I As Long
Dim IdCredito1 As Long
Dim IdCredito2 As Long
On Error GoTo merror

HayCreditosDistintos = False

'saco el primer credito seleccionado
For I = 1 To CLng(lv.ListItems.Count())
    If lv.ListItems.Item(I).Checked Then
       IdCredito1 = CLng(lv.ListItems.Item(I).SubItems(1))
       Exit For
    End If
Next I

'recorro la lista
For I = 1 To CLng(lv.ListItems.Count())
    If lv.ListItems.Item(I).Checked Then
       IdCredito2 = CLng(lv.ListItems.Item(I).SubItems(1))
       If IdCredito2 <> IdCredito1 Then
          HayCreditosDistintos = True
          Exit Function
       End If
    End If
Next I

Exit Function
merror:
tratarerrores "Error en funcion HayCreditosDistintos"
End Function
Private Sub CheckDescuentos_Click()
On Error GoTo merror

If CheckDescuentos.Value = 1 Then
   TxtDescuento.Enabled = True
   CheckRecargos.Value = 0
   TxtRecargo.Text = 0
   TxtRecargo.Enabled = False
Else
   TxtDescuento.Enabled = False
   TxtDescuento.Text = 0
End If

TxtDescuento.Text = Format(TxtDescuento.Text, "0.00")
  
Exit Sub
merror:
tratarerrores "Error aplicando descuentos"
End Sub
Private Sub CheckRecargos_Click()
On Error GoTo merror

If CheckRecargos.Value = 1 Then
   TxtRecargo.Enabled = True
   CheckDescuentos.Value = 0
   TxtDescuento.Text = 0
   TxtDescuento.Enabled = False
Else
   TxtRecargo.Enabled = False
   TxtRecargo.Text = 0
End If

TxtRecargo.Text = Format(TxtRecargo.Text, "0.00")

Exit Sub
merror:
tratarerrores "Error aplicando recargos"
End Sub
Private Sub TxtDescuento_Change()
On Error GoTo merror

If Trim(TxtImporteAFinanciar.Text) = "" Then Exit Sub
If CCur(TxtImporteAFinanciar.Text) = 0 Then Exit Sub
   
If Trim(TxtDescuento.Text) = "" Then Exit Sub
If Not IsNumeric(TxtDescuento.Text) Then Exit Sub
If CCur(TxtDescuento.Text) < 0 Then Exit Sub
If CCur(TxtDescuento.Text) > CCur(TxtImporteAFinanciar.Text) Then
   Exit Sub
End If

TxtSubtotal.Text = CCur(TxtImporteAFinanciar.Text) - CCur(TxtDescuento.Text)

Call CalcularImportes

Exit Sub
merror:
tratarerrores "Error cambiando el importe de descuento"
End Sub
Private Sub TxtRecargo_Change()
On Error GoTo merror

If Trim(TxtImporteAFinanciar.Text) = "" Then Exit Sub
If CCur(TxtImporteAFinanciar.Text) = 0 Then Exit Sub
   
If Trim(TxtRecargo.Text) = "" Then Exit Sub
If Not IsNumeric(TxtRecargo.Text) Then Exit Sub
If CCur(TxtRecargo.Text) < 0 Then Exit Sub

TxtSubtotal.Text = CCur(TxtImporteAFinanciar.Text) + CCur(TxtRecargo.Text)

Call CalcularImportes

Exit Sub
merror:
tratarerrores "Error cambiando el importe de recargo"
End Sub

