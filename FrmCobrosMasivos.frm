VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCobrosMasivos 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobrar cuotas multiples"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   HelpContextID   =   8
   Icon            =   "FrmCobrosMasivos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameCobrador 
      Caption         =   "Cobrador:"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4200
      TabIndex        =   28
      Top             =   4800
      Width           =   5175
      Begin VB.CheckBox CheckCobradores 
         Caption         =   "Seleccionar cobrador"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Permite seleccionar al cobrador de las cuotas"
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox ComboCobradores 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Lista de cobradores"
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame8 
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   120
      TabIndex        =   26
      Top             =   4320
      Width           =   4050
      Begin VB.TextBox TxtNumRecibo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox TxtMensaje 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   22
         ToolTipText     =   "Indica si se esta realizando un cobro total o parcial de cuotas."
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox TxtVuelto 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   21
         Tag             =   "N"
         ToolTipText     =   "Vuelto a entregar al cliente"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TxtImporteRecibido 
         Height          =   285
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   4
         Tag             =   "no"
         ToolTipText     =   "Importe entregado por el cliente"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox TxtImporteACobrar 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   17
         Tag             =   "N"
         ToolTipText     =   "Importe total seleccionado"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtSubtotal 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   20
         Tag             =   "N"
         ToolTipText     =   "Subtotal"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox TxtImporteRecargo 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   19
         Tag             =   "N"
         ToolTipText     =   "Importe de recargo"
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox CheckAplicarRecargos 
         Caption         =   "Recargos                  $:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Aplica un recargo a las cuotas seleccionadas (el % se configura en Opciones)"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox TxtImporteDescuento 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   18
         Tag             =   "N"
         ToolTipText     =   "Importe de descuento"
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox CheckAplicarDescuentos 
         Caption         =   "Descuentos              $:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Aplica un descuento a las cuotas seleccionadas (el % se configura en Opciones)"
         Top             =   600
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   600
         TabIndex        =   49
         ToolTipText     =   "Fecha de registracion del cobro de las cuotas seleccionadas"
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   57475073
         CurrentDate     =   39443
      End
      Begin VB.Label Label12 
         Caption         =   "Factura Nº:"
         Height          =   255
         Left            =   1920
         TabIndex        =   50
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de cobro:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Vuelto                             $:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Importe recibido              $:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Importe seleccionado     $:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Subtotal                          $:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1320
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      ToolTipText     =   "Blanquea la pantalla"
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "&Anular cobros"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      ToolTipText     =   "Anula el cobro de las cuotas seleccionadas"
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton CmdCobrar 
      Caption         =   "Co&brar cuotas"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Registra el cobro de las cuotas seleccionadas"
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Frame FrameFiltro 
      Caption         =   "Filtrar cuotas por:"
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   4200
      TabIndex        =   27
      Top             =   4200
      Width           =   5175
      Begin VB.CheckBox CheckTodos 
         Caption         =   "Vencidas masivas"
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         ToolTipText     =   "Muestra las cuotas vencidas de todos los clientes"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox CheckCobradas 
         Caption         =   "Incluir cobradas"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Tag             =   "no"
         ToolTipText     =   "Muestra las cuotas cobradas del credito seleccionado"
         Top             =   230
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox CheckCreditoActual 
         Caption         =   "Credito actual"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         ToolTipText     =   "Muestra cuotas de todos los creditos o solo del credito seleccionado"
         Top             =   230
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      ToolTipText     =   "Cierra la pantalla"
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccione el cliente:"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton CmdBuscarCliente 
         Height          =   375
         Left            =   8520
         Picture         =   "FrmCobrosMasivos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Permite seleccionar al cliente de una lista"
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox TxtCliente 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Cliente al cual le cobraremos cuotas"
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cuotas pendientes:"
      ForeColor       =   &H00FF0000&
      Height          =   3600
      Left            =   120
      TabIndex        =   25
      Top             =   600
      Width           =   9255
      Begin VB.CommandButton CmdRefrescar 
         Caption         =   "Refrescar"
         Height          =   255
         Left            =   5880
         TabIndex        =   23
         ToolTipText     =   "Refresca la lista de cuotas"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox TxtCuotasACobrar 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   15
         Tag             =   "N"
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox TxtImporteTotal 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7800
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   16
         Tag             =   "N"
         ToolTipText     =   "Saldo total de la lista"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   135
         Left            =   4800
         TabIndex        =   31
         ToolTipText     =   "Vencida en mora"
         Top             =   3240
         Width           =   135
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   135
         Left            =   3720
         TabIndex        =   30
         ToolTipText     =   "Pendiente al dia"
         Top             =   3240
         Width           =   135
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   135
         Left            =   2760
         TabIndex        =   29
         ToolTipText     =   "Cobrada o comodin(si tiene la letra ""C"")"
         Top             =   3240
         Width           =   135
      End
      Begin MSComctlLib.ListView lvcuotas 
         Height          =   2925
         Left            =   120
         TabIndex        =   1
         Tag             =   "BORRAR"
         ToolTipText     =   "Lista de cuotas del cliente seleccionado"
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5159
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
         NumItems        =   43
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cod.Prestamo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Credito"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cuota"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cupon Nº"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Capital"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Interes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Rec.Refin"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Rec.2ºVto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Gastos"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Otorgamiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Seguro"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Iva.Interes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Iva.Seguro"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Iva.Ot.Gastos"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "1º vto"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Imp.1º vto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "2º Vto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Imp.2º Vto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Dias.Mora"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Mora"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Iva.Mora"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Saldo.Total"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Fecha.Cobro"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "Total.Cobrado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Total.Parcial"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "Descuento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "Recargo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "Saldo.Capital"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Saldo.Interes"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "Saldo.Gastos"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "Saldo.Otorg"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "Saldo.Seg."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "Saldo.Iva.Interes"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "Saldo.Iva.Seguro"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "Saldo.Iva.Ot.Gastos"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "Saldo.Rec.2Vto"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "Saldo.Refin."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "CodeBar"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "PF"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "RP"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "Check"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Cuotas a cobrar:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Saldo $:"
         Height          =   255
         Left            =   7200
         TabIndex        =   37
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Parcial"
         Height          =   255
         Left            =   2160
         TabIndex        =   36
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label11 
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
         Left            =   1920
         TabIndex        =   35
         ToolTipText     =   "El asterisco indica quer la cuota tiene cobros parciales"
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "Vencida"
         Height          =   255
         Left            =   5040
         TabIndex        =   34
         ToolTipText     =   "Vencida en mora"
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Pendiente"
         Height          =   255
         Left            =   3960
         TabIndex        =   33
         ToolTipText     =   "Pendiente al dia"
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Cobrada"
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         ToolTipText     =   "Cobrada o comodin (si tiene la letra ""C"")"
         Top             =   3240
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Height          =   2100
      Left            =   4200
      TabIndex        =   45
      Top             =   5470
      Width           =   5175
      Begin VB.CommandButton CmdTodosExcedentes 
         Caption         =   "Imprim.Todos"
         Height          =   315
         Left            =   3840
         TabIndex        =   53
         ToolTipText     =   "Imprime los excedentes sin imputar de todos los clientes"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton CmdImprimirExcedentes 
         Caption         =   "Imprimir exced."
         Height          =   345
         Left            =   2520
         TabIndex        =   52
         ToolTipText     =   "Imprime la lista de excedentes"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CheckBox CheckExcedentes 
         Caption         =   "Cobrar con  excedentes RP"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         ToolTipText     =   "Permite usar los excedentes"
         Top             =   150
         Width           =   2295
      End
      Begin MSComctlLib.ListView Lv2 
         Height          =   1575
         Left            =   120
         TabIndex        =   46
         Tag             =   "BORRAR"
         ToolTipText     =   "Lista de excedentes del cliente seleccionado"
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2778
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
         Enabled         =   0   'False
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Obs"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IdExcedenteCliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "IdCliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Prestamo"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Credito"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cuota"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "FechaCobro"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Observaciones"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Archivo"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin VB.CommandButton CmdImportar 
      Caption         =   "Importar Cobros"
      Height          =   375
      Left            =   5760
      TabIndex        =   44
      ToolTipText     =   "Importa cobros desde una planilla Excel del disco rigido"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "FrmCobrosMasivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AQUI SE REGISTRAN COBROS DE CUOTAS EN FORMA INDIVIDUAL O GRUPAL
'TAMBIEN SE USAN LOS EXCEDENTES QUE QUEDAN DE COBROS DE PAGOFACIL O RAPIPAGO
'Y SE LOS USA PARA COBRAR CUOTAS MANUALMENTE EN ESTA PANTALLA

Public IdCliente As Long
Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call RefrescarOpcionesSistema

Call LimpiarCampos(Me)

If Not VG_APLICARSEGUNDOVENCIMIENTO Then
   lvcuotas.ColumnHeaders.Item(19).Width = 0
   lvcuotas.ColumnHeaders.Item(20).Width = 0
End If

'CARGA SOLO LOS COBRADORES ACTIVOS
Call CargarComboCobradores("cobradores", ComboCobradores, True, False)

DTPicker1.Value = Date

'cargo el proximo numero de factura de credimaco
TxtNumRecibo.Text = Format(UltimoId("ultimonumrecibo", "configuracionsistema"), "0000000")

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de cobro de cuotas"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
IdCliente = 0
Unload Me
End Sub
Private Sub cmdnuevo_Click()
Call RefreshTimer
Call LimpiarCampos(Me)
Call SetearEntorno
IdCliente = 0

'obtiene el proximo numero de factura de credimaco
TxtNumRecibo.Text = Format(UltimoId("ultimonumrecibo", "configuracionsistema"), "0000000")

CheckExcedentes.Value = 0

End Sub
Private Sub CmdRefrescar_Click()
Call RefreshTimer
Call ActualizarListas
End Sub
Private Sub CmdBuscarCliente_Click()
FrmClientesAbm.FormularioPadre = "COBROSMASIVOS"
Call CenterForm(FrmClientesAbm)
FrmClientesAbm.Show
End Sub
Private Sub CheckCreditoActual_Click()
'solo muestra las cuotas del credito seleccionado o muestra todas
Call ActualizarListas
End Sub
Private Sub CheckCobradores_Click()
If CheckCobradores.Value = 1 Then
   ComboCobradores.Enabled = True
Else
   ComboCobradores.ListIndex = -1
   ComboCobradores.Enabled = False
End If
End Sub
Private Sub CheckCobradas_Click()
Call ActualizarListas
End Sub
Private Sub CargarCuotasCreditos()
'carga las cuotas de todos los creditos vigentes del cliente seleccionado
Dim rec As rdoResultset
Dim Nitem As ListItem
Dim ImporteTotal As Currency
Dim ImporteMora As Currency
Dim IvaMora As Currency
Dim I As Long
Dim Cad1 As String
Dim SaldoCuota As Currency
Dim SaldoCuota1erVenc As Currency
Dim Importe1erVenc As Currency
Dim ImporteParcial As Currency
Dim ImporteCobrado As Currency
Dim DiasMora As Long
Dim Vencimiento2Cuota As Currency
Dim RecargoCuota As Currency
Dim CapitalRestante As Currency
Dim InteresRestante As Currency
Dim Vencimiento2Restante As Currency
Dim RefinRestante As Currency
Dim GastoRestante As Currency
Dim SeguroRestante As Currency
Dim OtorgamientoRestante As Currency
Dim IvaInteresRestante As Currency
Dim IvaSeguroRestante As Currency
Dim IvaOtorGastoRestante As Currency
Dim ImporteMoraGral As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
Dim IvaACobrarDevuelto  As Currency
On Error GoTo merror

TxtCuotasACobrar.Text = 0
TxtImporteACobrar.Text = 0

If CheckTodos.Value = 0 Then
   If Trim(TxtCliente.Text) = "" Then Exit Sub
End If

Set rec = CargarRecCuotasCreditos()

lvcuotas.ListItems.Clear

I = 1
ImporteTotal = 0
Do While Not rec.EOF
   ImporteMora = 0
   IvaMora = 0
   'pongo descripcion de cuotas
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
   Nitem.SubItems(2) = rec.rdoColumns("cliente") & vbNullString
   Nitem.SubItems(3) = Format(rec.rdoColumns("idcredito"), "000000") & vbNullString
   Nitem.SubItems(4) = Format(rec.rdoColumns("numcuota"), "00") & vbNullString
   Nitem.SubItems(5) = Format(rec.rdoColumns("numfactura"), "000000000") & vbNullString
   
   'importes originales de la cuota
   Nitem.SubItems(6) = Format(rec.rdoColumns("importeamortizacion"), "0.00") & vbNullString
   Nitem.SubItems(7) = Format(rec.rdoColumns("importeinteres"), "0.00") & vbNullString
   Nitem.SubItems(8) = Format(rec.rdoColumns("importerefinanciacion"), "0.00") & vbNullString
   Nitem.SubItems(9) = Format(rec.rdoColumns("importerecargovencimiento2"), "0.00") & vbNullString
   Nitem.SubItems(10) = Format(rec.rdoColumns("importegastos"), "0.00") & vbNullString
   Nitem.SubItems(11) = Format(rec.rdoColumns("otorgamiento"), "0.00") & vbNullString
   Nitem.SubItems(12) = Format(rec.rdoColumns("importeseguros"), "0.00") & vbNullString
   Nitem.SubItems(13) = Format(rec.rdoColumns("ivainteres"), "0.00") & vbNullString
   Nitem.SubItems(14) = Format(rec.rdoColumns("ivaseguros"), "0.00") & vbNullString
   Nitem.SubItems(15) = Format(rec.rdoColumns("ivaotorgamientogastos"), "0.00") & vbNullString
   
   Nitem.SubItems(16) = rec.rdoColumns("fechavencimiento1") & vbNullString
   Nitem.SubItems(17) = Format(rec.rdoColumns("importetotal"), "0.00") & vbNullString
   Nitem.SubItems(18) = rec.rdoColumns("fechavencimiento2") & vbNullString
   Nitem.SubItems(19) = Format(rec.rdoColumns("ImporteVencimiento2"), "0.00") & vbNullString
   
   'el importe parcial lo saca de ingresos sumando los items cobrados menos la mora e iva mora cobrada
   ImporteParcial = ObtenerImporteParcialX(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
      
   'trael el saldo de punto de partida exacto desde ingresos y cuota
   'trae el saldo incluyendo solo items fijos sin mora e iva mora cobrados
   SaldoCuota = ObtenerSaldoCuotaX(rec.rdoColumns("idcredito"), rec.rdoColumns("NumCuota"), DTPicker1.Value, SaldoCuota1erVenc)
   Importe1erVenc = ObtenerImporte1erVenc(rec.rdoColumns("idcredito"), rec.rdoColumns("NumCuota"))
   DiasMora = 0
   If CDate(DTPicker1.Value) > CDate(rec.rdoColumns("fechavencimiento1")) Then
       DiasMora = CDate(DTPicker1.Value) - CDate(rec.rdoColumns("fechavencimiento1"))
   End If
   
   IvaMora = 0
   'si no esta cobrada actualizo el importe si es necesario
   If IsNull(rec.rdoColumns("fechacobro")) And Not rec.rdoColumns("cuotacomodin") And IsNull(rec.rdoColumns("fecharefinanciacion")) Then
      'esto funciona para ambos vencimientos (si hay un solo vto ambos son iguales)
      If CDate(DTPicker1.Value) > CDate(rec.rdoColumns("fechavencimiento2")) Then
         'primero calculo la mora en forma habitual sobre lo restante
                           
         '***SI EL SALDO RESTANTE ES CADA VEZ MENOR, LA MORA SERA CADA VEZ MENOR
         '***SI EL SALDO RESTANTE SE MANTIENE ALTO PORQUE SE CUBRIO SOLO MORA
         '***AL PASAR LOS DIAS LA MORA SERA MAS ALTA
         ImporteMora = CalculoMoraPendiente(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"), rec.rdoColumns("exceptuada"), Importe1erVenc, rec.rdoColumns("fechavencimiento1"), CDate(DTPicker1.Value), IvaACobrarDevuelto)
         '''''''********ImporteMora = CalcularInteresMoraZZ(rec.rdoColumns("exceptuada"), SaldoCalculoMora, FechaCalculoMora, CDate(DTPicker1.Value))
         'tambien calculo el iva mora en forma habitual
         IvaMora = 0
         If VG_APLICARIMPUESTOS Then
            If VG_IMPUESTOSCREDIMACO Then
               'calculo el iva de la mora..uso la variable global
               '***CALCULO EL IVA DE LA MORA SIN TENER EN CUENTA AUN
               IvaMora = IvaACobrarDevuelto
            End If
         End If
         'trae solo la mora cobrada sin importar fechas (todo)
         '''''''********SoloMoraCobrada = ObtenerMoraCobrada(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
         
         '''''''********SoloIvaMoraCobrada = ObtenerIvaMoraCobrada(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
         '***
         '''''''********If CCur(ImporteMora) <= CCur(SoloMoraCobrada) Then
         '''''''********ImporteMora = 0
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
         lvcuotas.ListItems.Item(I).ForeColor = vbRed
         lvcuotas.ListItems.Item(I).ListSubItems(1).ForeColor = vbRed
         lvcuotas.ListItems.Item(I).ListSubItems(2).ForeColor = vbRed
         lvcuotas.ListItems.Item(I).ListSubItems(3).ForeColor = vbRed
         lvcuotas.ListItems.Item(I).ListSubItems(4).ForeColor = vbRed
         lvcuotas.ListItems.Item(I).ListSubItems(5).ForeColor = vbRed
      Else
         'pongo en azul sin cobrar al dia
         lvcuotas.ListItems.Item(I).ForeColor = &HFF0000
         lvcuotas.ListItems.Item(I).ListSubItems(1).ForeColor = &HFF0000
         lvcuotas.ListItems.Item(I).ListSubItems(2).ForeColor = &HFF0000
         lvcuotas.ListItems.Item(I).ListSubItems(3).ForeColor = &HFF0000
         lvcuotas.ListItems.Item(I).ListSubItems(4).ForeColor = &HFF0000
         lvcuotas.ListItems.Item(I).ListSubItems(5).ForeColor = &HFF0000
     End If
   Else
      DiasMora = 0
      ImporteMora = 0
      IvaMora = 0
      ImporteActualizado = 0
      SaldoCuota = 0
      
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
   'estas van aca porque se calculan mas arriba
   Nitem.SubItems(20) = Format(DiasMora, "000") & vbNullString
   Nitem.SubItems(21) = Format(ImporteMora, "0.00") & vbNullString
   Nitem.SubItems(22) = Format(IvaMora, "0.00") & vbNullString
   Nitem.SubItems(23) = Format(SaldoCuota, "0.00") & vbNullString
   Nitem.SubItems(24) = rec.rdoColumns("fechacobro") & vbNullString
    
   If Not IsNull(rec.rdoColumns("fechacobro")) Then
      ImporteCobrado = CCur(rec.rdoColumns("importecobrado"))
   Else
      ImporteCobrado = CCur(ImporteParcial)
   End If
   Nitem.SubItems(25) = Format(ImporteCobrado, "0.00") & vbNullString
   Nitem.SubItems(26) = Format(ImporteParcial, "0.00") & vbNullString
   Nitem.SubItems(27) = Format(rec.rdoColumns("importedescuentos"), "0.00") & vbNullString
   Nitem.SubItems(28) = Format(rec.rdoColumns("importerecargos"), "0.00") & vbNullString
   
   'desde aca MUESTRA LO QUE RESTA DE CADA ITEM
   CapitalRestante = CCur(rec.rdoColumns("importeamortizacion")) - ObtenerCapitalCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
   Nitem.SubItems(29) = Format(CapitalRestante, "0.00") & vbNullString
   
   InteresRestante = CCur(rec.rdoColumns("importeinteres")) - ObtenerInteresCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
   Nitem.SubItems(30) = Format(InteresRestante, "0.00") & vbNullString
   
   GastoRestante = CCur(rec.rdoColumns("importegastos")) - ObtenerGastosCobrados(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
   Nitem.SubItems(31) = Format(GastoRestante, "0.00") & vbNullString
   
   OtorgamientoRestante = CCur(rec.rdoColumns("otorgamiento")) - ObtenerOtorgamientoCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
   Nitem.SubItems(32) = Format(OtorgamientoRestante, "0.00") & vbNullString
   
   SeguroRestante = CCur(rec.rdoColumns("importeseguros")) - ObtenerSegurosCobrados(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
   Nitem.SubItems(33) = Format(SeguroRestante, "0.00") & vbNullString
   
   IvaInteresRestante = CCur(rec.rdoColumns("ivainteres")) - ObtenerIvaInteresCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
   Nitem.SubItems(34) = Format(IvaInteresRestante, "0.00") & vbNullString
   
   IvaSeguroRestante = CCur(rec.rdoColumns("ivaseguros")) - ObtenerIvaSegurosCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
   Nitem.SubItems(35) = Format(IvaSeguroRestante, "0.00") & vbNullString
   
   IvaOtorGastoRestante = CCur(rec.rdoColumns("ivaotorgamientogastos")) - ObtenerIvaOtorGastosCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
   Nitem.SubItems(36) = Format(IvaOtorGastoRestante, "0.00") & vbNullString
   
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
   Nitem.SubItems(37) = Format(Vencimiento2Restante, "0.00") & vbNullString
   RefinRestante = CCur(rec.rdoColumns("importerefinanciacion")) - ObtenerRefinCobrado(rec.rdoColumns("IdCredito"), rec.rdoColumns("NumCuota"))
   Nitem.SubItems(38) = Format(RefinRestante, "0.00") & vbNullString
   Nitem.SubItems(39) = rec.rdoColumns("codigobarras") & vbNullString
      
   If rec.rdoColumns("pagofacil") Then
      Nitem.SubItems(40) = "SI"
   Else
      Nitem.SubItems(40) = " "
   End If
   
   If rec.rdoColumns("rapipago") Then
      Nitem.SubItems(41) = "SI"
   Else
      Nitem.SubItems(41) = " "
   End If
   
   ImporteTotal = CCur(ImporteTotal) + CCur(SaldoCuota)
      
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

TxtImporteTotal.Text = Format(ImporteTotal, "0.00")

Exit Sub
merror:
tratarerrores "Error cargando cuotas de creditos"
End Sub
Private Function CargarRecCuotasCreditos() As rdoResultset
'carga las cuotas de un credito
Dim sql As String
Dim IdCredito As Long
Dim CondicionFiltro As String
On Error GoTo merror
  
If CheckTodos.Value = 1 Then
   'de todos los clientes vencidas e impagas
   CondicionFiltro = "and 1=1 and cuotas.fechacobro is null and cuotas.fechavencimiento2<='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "'"
Else
  'de un solo cliente
   CondicionFiltro = "and creditos.idcliente='" & CLng(IdCliente) & "'"
   
   'si imprimo cuotas cobradas
   If CheckCobradas.Value = 1 Then
      CondicionFiltro = CondicionFiltro
   Else
      CondicionFiltro = CondicionFiltro & " and cuotas.fechacobro is null"
   End If
End If

'si solo deseo ver el credito seleccionado (uno solo)
If CheckCreditoActual.Value = 1 Then
   If lvcuotas.ListItems.Count() > 0 Then
      IdCredito = CLng(lvcuotas.SelectedItem.SubItems(3))
      CondicionFiltro = CondicionFiltro + " and creditos.idcredito='" & CLng(IdCredito) & "'"
   End If
Else
   'ver todos los creditos del cliente seleccionado
   CondicionFiltro = CondicionFiltro
End If

sql = "select clientes.apellido + ', ' + clientes.nombre as cliente,creditos.numcuotas," & _
      "creditos.idcredito,creditos.codprestamo,cuotas.periodo,cuotas.numcuota,cuotas.numfactura,cuotas.cobrosparciales," & _
      "cuotas.fechavencimiento1,cuotas.fechavencimiento2,cuotas.importeamortizacion,cuotas.importeinteres," & _
      "cuotas.importecuota,cuotas.importerecargovencimiento2,cuotas.fecharefinanciacion,cuotas.pagofacil,cuotas.rapipago," & _
      "cuotas.importegastos,cuotas.importeseguros,cuotas.importeimpuestos,cuotas.ivainteres,cuotas.ivaseguros," & _
      "cuotas.ivaotorgamientogastos,cuotas.ivamora,cuotas.importeparcial," & _
      "cuotas.importedescuentos,cuotas.importerecargos,cuotas.importemora,cuotas.importerefinanciacion," & _
      "cuotas.fechacobro,cuotas.importecobrado,cuotas.codigobarras,cuotas.cuotacomodin,cuotas.otorgamiento,cuotas.logic1 as exceptuada," & _
      "(cuotas.importevencimiento1) as importetotal,cuotas.importevencimiento2,provincias.nombre as provincia " & _
      "from provincias inner join (clientes inner join (creditos inner join cuotas on " & _
      "creditos.idcredito=cuotas.idcredito) on " & _
      "clientes.idcliente=creditos.idcliente) on provincias.idprovincia=creditos.idprovincia " & _
      "where creditos.fechafinalizacion is null and creditos.fechabloqueo is null " & _
      "and cuotas.cuotacomodin = 0 and cuotas.fecharefinanciacion is Null " & CondicionFiltro & _
      " order by creditos.idcredito,cuotas.numcuota"

Set CargarRecCuotasCreditos = cnSQL.OpenResultset(sql)

Exit Function
merror:
tratarerrores "Error cargando el registro de cuotas"
End Function
Private Function ContarCuotasImpagas() As Long
'cuenta la cantidad de cuotas impagas que estan seleccionadas en la lista
'se usa para dividir el importre de descuento en partes iguales cuando
'se cobra totalmente lo seleccionado
'tambien puede servir para dividir el recargo en partes iguales
Dim I As Long
Dim Contador As Long
Dim IdCredito As Long
Dim NumCuota As Long
On Error GoTo merror

ContarCuotasImpagas = 0

Contador = 0

For I = 1 To lvcuotas.ListItems.Count
    'si la cuota esta marcada
    If lvcuotas.ListItems.Item(I).Checked Then
       IdCredito = CLng(lvcuotas.ListItems.Item(I).SubItems(3))
       NumCuota = CLng(lvcuotas.ListItems.Item(I).SubItems(4))
          
       'si no esta cobrada
       If Not CuotaCobrada(IdCredito, NumCuota) And Not CuotaRefinanciada(IdCredito, NumCuota) And Not CuotaEsComodin(IdCredito, NumCuota) Then
          Contador = Contador + 1
       End If

    End If
Next I

ContarCuotasImpagas = Contador

Exit Function
merror:
tratarerrores "Error en funcion ContarCuotasImpagas"
End Function
Private Sub CmdCobrar_Click()
Call RefreshTimer
CmdCobrar.Enabled = False
Call Cobrar
CmdCobrar.Enabled = True
End Sub
Private Sub Cobrar()
Dim sql As String
Dim rec As rdoResultset
Dim Observaciones As String
Dim I As Long
Dim NumCredito As Long
Dim NumCuota As Long
Dim CodPrestamo As String
Dim IdCobrador As Long
Dim IdIngreso As Long
Dim NumFactura As Long
Dim NumComprobante As Long
Dim ImporteCobrador As Currency
Dim SaldoCuota As Currency
Dim ImporteParcial As Currency
Dim ImporteZ As Currency
Dim ImporteCuota As Currency
Dim HuboCobros As Boolean
Dim CobroParcial As Boolean
Dim CantidadCuotasACobrar As Long
Dim DescuentoCuota As Currency
Dim RecargoCuota As Currency
Dim ImporteMora As Currency
Dim IvaMora As Currency
Dim CapitalCobrado As Currency
Dim InteresCobrado As Currency
Dim GastosCobrados As Currency
Dim OtorgamientoCobrado As Currency
Dim SegurosCobrados As Currency
Dim IvaInteresCobrado As Currency
Dim IvaSegurosCobrado As Currency
Dim IvaOtorGastosCobrado As Currency
Dim MoraCobrada As Currency
Dim IvaMoraCobrada As Currency
Dim RefinCobrado As Currency
Dim Vencimiento2Cobrado As Currency
Dim SelladosCobrados As Currency
Dim CapitalRestante As Currency
Dim InteresRestante As Currency
Dim Vencimiento2Restante As Currency
Dim RefinRestante As Currency
Dim GastoRestante As Currency
Dim SeguroRestante As Currency
Dim OtorgamientoRestante As Currency
Dim IvaInteresRestante As Currency
Dim IvaSeguroRestante As Currency
Dim IvaOtorGastoRestante As Currency
Dim ImporteTotalCobrado As Currency
Dim ImporteRealCobrado As Currency
Dim NumRecibo As String
Dim Cubre As Boolean
Dim CondicionCuotas As String
Dim FormaCobro As String
Dim ImporteXXX As Currency
Dim ImporteIngresos As Currency
Dim HuboIngresos As Boolean
Dim Vencimiento1 As Date
Dim Vencimiento2 As Date
Dim Diferencia As Currency
Dim ImporteVencimiento1 As Currency
Dim ImporteVencimiento2 As Currency
Dim SecuenciaIngreso As Long
Dim MontoCobradoSecuencia As Currency
On Error GoTo merror

FormaCobro = "EFECTIVO"

If Not HayFilasChequeadas(lvcuotas) Then
   MsgE "Debe seleccionar cuotas para cobrar"
   Exit Sub
End If

If CDate(DTPicker1.Value) < CDate(VG_FECHALIMITEINGRESO) Then
   datosok = False
   MsgE "La Fecha del cobro es inferior a la fecha limite permitida"
   DTPicker1.SetFocus
   Exit Sub
End If

'valido si permito cobros diferidos
If Not VG_COBROSDIFERIDOS Then
   If CDate(DTPicker1.Value) <> CDate(Date) Then
      MsgE "Verifique a fecha de cobro...debe ser igual a la actual"
      Exit Sub
   End If
Else
   'si permito diferidos
   If Year(DTPicker1.Value) < 2000 Then
      MsgE "Verifique la fecha de cobro...(el año debe ser superior a 2000)"
      DTPicker1.SetFocus
      Exit Sub
   End If
   If CDate(DTPicker1.Value) > CDate(Date) Then
      MsgE "Verifique la fecha de cobro...debe ser igual a la actual"
      DTPicker1.SetFocus
      Exit Sub
   End If
End If

'solo permito cobrar cuotas de un mismo credito/cliente(por el tema de excedentes)
If HayCreditosDistintos(lvcuotas) Then
   MsgE "Debe seleccionar cuotas de un mismo credito"
   Exit Sub
End If

'valido el importe seleccionado
If Trim(TxtImporteACobrar.Text) = "" Then
   MsgE "Debe seleccionar cuotas sin cobrar"
   Exit Sub
End If
If Not IsNumeric(TxtImporteACobrar.Text) Then
   MsgE "Debe seleccionar cuotas sin cobrar"
   Exit Sub
End If
If CCur(TxtImporteACobrar.Text) <= 0 Then
   MsgE "Debe seleccionar cuotas sin cobrar"
   Exit Sub
End If

'valido el numero de recibo=factura de credimaco
If Trim(TxtNumRecibo.Text) = "" Then
   MsgE "Debe ingresar el numero de recibo"
   TxtNumRecibo.SetFocus
   Exit Sub
End If
If Not IsNumeric(TxtNumRecibo.Text) Then
   MsgE "El numero de recibo debe ser numerico"
   TxtNumRecibo.SetFocus
   Exit Sub
End If
If CLng(TxtNumRecibo.Text) <= 0 Then
   MsgE "El numero de recibo debe ser mayor a cero"
   TxtNumRecibo.SetFocus
   Exit Sub
End If
If ExisteFacturaCredimaco(TxtNumRecibo.Text) Then
   If Not MsgP("El numero de recibo ya existe en la base de datos...¿Lo registra igual?") Then
      TxtNumRecibo.SetFocus
      Exit Sub
   Else
      'nada
   End If
End If

'valido el importe recibido
If Trim(TxtImporteRecibido.Text) = "" Then
   MsgE "Debe ingresar el importe entregado por el cliente"
   TxtImporteRecibido.SetFocus
   Exit Sub
End If
If Not IsNumeric(TxtImporteRecibido.Text) Then
   MsgE "El importe entregado por el cliente debe ser numerico"
   TxtImporteRecibido.SetFocus
   Exit Sub
End If
If CCur(TxtImporteRecibido.Text) <= 0 Then
   MsgE "El importe entregado por el cliente debe ser mayor a cero"
   TxtImporteRecibido.SetFocus
   Exit Sub
End If

'verifico si seleccionaron cobrador
IdCobrador = 0
If CheckCobradores.Value = 1 Then
   If ComboCobradores.Text <> "" Then
      IdCobrador = ComboCobradores.ItemData(ComboCobradores.ListIndex)
   End If
End If

'cuento la cantidad de cuotas seleccionadas que estan impagas
CantidadCuotasACobrar = ContarCuotasImpagas()
DescuentoCuota = 0
RecargoCuota = 0
'verifico que si es un cobro parcial no haya descuento
If CheckAplicarDescuentos.Value = 1 Then
   If Trim(TxtImporteDescuento.Text) = "" Then
      TxtImporteDescuento.Text = 0
   End If
   If Not IsNumeric(TxtImporteDescuento.Text) Then
      TxtImporteDescuento.Text = 0
   End If
   If CCur(TxtImporteDescuento.Text) > 0 Then
      If CCur(TxtImporteRecibido.Text) < CCur(TxtSubtotal.Text) Then
         MsgE "Solo se puede descontar si el cobro es total..(no con cobros parciales)"
         Exit Sub
      End If
      'verifico si el imorte de descuento es mayor al selecionado
      If CCur(TxtImporteDescuento.Text) > CCur(TxtImporteACobrar.Text) Then
         MsgE "El importe de descuento debe ser menor al importe seleccionado"
         Exit Sub
      End If
      'calculo la proporcion de descuentos y recargos solo si es cobro total
      DescuentoCuota = CCur(CCur(TxtImporteDescuento.Text) / CantidadCuotasACobrar)
   End If
End If

'verifico que si es un cobro total no haya recargos
If CheckAplicarRecargos.Value = 1 Then
   If Trim(TxtImporteRecargo.Text) = "" Then
      TxtImporteRecargo.Text = 0
   End If
   If Not IsNumeric(TxtImporteRecargo.Text) Then
      TxtImporteRecargo.Text = 0
   End If
   If CCur(TxtImporteRecargo.Text) > 0 Then
      If CCur(TxtImporteRecibido.Text) < CCur(TxtSubtotal.Text) Then
         MsgE "Solo se puede recargar si el cobro es total..(no con cobros parciales)"
         Exit Sub
      End If
      RecargoCuota = CCur(CCur(TxtImporteRecargo.Text) / CantidadCuotasACobrar)
   End If
End If

If Not MsgP("¿Confirma el cobro de las cuotas seleccionadas?") Then Exit Sub

HuboCobros = False
CobroParcial = False
ImporteRealCobrado = 0

'otras validaciones
'valido el cliente
If Not ExisteCliente(IdCliente) Then
   MsgE "El cliente no existe"
   Exit Sub
End If

'valido el cobrador
If IdCobrador > 0 Then
   If Not ExisteCobrador(IdCobrador) Then
      MsgE "El cobrador no existe"
      Exit Sub
   End If
End If

Cubre = False

If CCur(TxtImporteRecibido.Text) < CCur(TxtSubtotal.Text) Then
   '***COBRA UNA PARTE DE LA O LAS CUOTAS SELECCIONADAS***
   'luego abajo debe intentar cubrir por items hasta donde alcance
   Cubre = False
   MontoCobradoSecuencia = CCur(TxtImporteRecibido.Text)
Else
   'cubre todas las cuotas SELECCIONADAS si o si (nunca entra a parcial)
   Cubre = True
   MontoCobradoSecuencia = CCur(TxtSubtotal.Text)
End If

HuboIngresos = False

ImporteTotalCobrado = CCur(TxtImporteRecibido.Text)
SecuenciaIngreso = UltimoId("secuencia", "ingresos") + 1

'inicio transaccion
cnSQL.BeginTrans
   
For I = 1 To lvcuotas.ListItems.Count()
    IvaMora = 0
    'esto va solo si no cubre
    If Not Cubre Then
       ImporteRealCobrado = 0
    End If
       
    'si la cuota esta marcada
    If lvcuotas.ListItems.Item(I).Checked Then
       CodPrestamo = CStr(lvcuotas.ListItems.Item(I).SubItems(1))
       NumCredito = CLng(lvcuotas.ListItems.Item(I).SubItems(3))
       NumCuota = CLng(lvcuotas.ListItems.Item(I).SubItems(4))
       Vencimiento1 = CDate(lvcuotas.ListItems.Item(I).SubItems(16))
       Vencimiento2 = CDate(lvcuotas.ListItems.Item(I).SubItems(18))
       ImporteVencimiento1 = CCur(lvcuotas.ListItems.Item(I).SubItems(17))
       ImporteVencimiento2 = CCur(lvcuotas.ListItems.Item(I).SubItems(19))
          
       'voy validando el credito seleccionado
       If ExisteCredito(NumCredito) And Not CreditoFinalizado(NumCredito) And Not CreditoBloqueado1(NumCredito) Then
          NumFactura = CLng(lvcuotas.ListItems.Item(I).SubItems(5))
          
          'si no esta cobrada
          If Not CuotaCobrada(NumCredito, NumCuota) And Not CuotaRefinanciada(NumCredito, NumCuota) And Not CuotaEsComodin(NumCredito, NumCuota) Then
             'saldo original inlcuye mora y resta de parciales
             ImporteCuota = CCur(lvcuotas.ListItems.Item(I).SubItems(23))
             
             ImporteMora = CCur(lvcuotas.ListItems.Item(I).SubItems(21))
             IvaMora = CCur(lvcuotas.ListItems.Item(I).SubItems(22))
                         
             'los saldos restantes por items
             CapitalRestante = CCur(lvcuotas.ListItems.Item(I).SubItems(29))
             InteresRestante = CCur(lvcuotas.ListItems.Item(I).SubItems(30))
             GastoRestante = CCur(lvcuotas.ListItems.Item(I).SubItems(31))
             OtorgamientoRestante = CCur(lvcuotas.ListItems.Item(I).SubItems(32))
             SeguroRestante = CCur(lvcuotas.ListItems.Item(I).SubItems(33))
             IvaInteresRestante = CCur(lvcuotas.ListItems.Item(I).SubItems(34))
             IvaSeguroRestante = CCur(lvcuotas.ListItems.Item(I).SubItems(35))
             IvaOtorGastoRestante = CCur(lvcuotas.ListItems.Item(I).SubItems(36))
             Vencimiento2Restante = CCur(lvcuotas.ListItems.Item(I).SubItems(37))
             RefinRestante = CCur(lvcuotas.ListItems.Item(I).SubItems(38))

             SaldoCuota = CCur(ImporteCuota) - CCur(DescuentoCuota) + CCur(RecargoCuota)
          
             'verificar si tenia cobros parciales
             ImporteParcial = ObtenerImporteParcialX(NumCredito, NumCuota)
               
             'este es un truco para que siempre entre en la parte de arriba
             'cuando se cubra todas las cuotas
             If Cubre Then
                ImporteTotalCobrado = CCur(SaldoCuota) + 1
             End If
                
             'si el total actual permite saldar esta cuota
             If CCur(ImporteTotalCobrado) >= CCur(SaldoCuota) Then
                CobroParcial = False
                ImporteRealCobrado = CCur(SaldoCuota)
                ImporteIngresos = CCur(SaldoCuota)
                             
                ImporteParcial = CCur(ImporteParcial) + CCur(SaldoCuota)
                   
                sql = "update cuotas set fechacobro='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "'," & _
                      "importecobrado=" & ConvertirDblSql(CCur(ImporteParcial)) & "," & _
                      "formacobro='" & CStr(FormaCobro) & "', origen='MANUAL',idcobrador=" & IdCobrador & " " & _
                      "where idcredito=" & CLng(NumCredito) & " " & _
                      "and numcuota=" & CLng(NumCuota)
                cnSQL.Execute sql
                HuboCobros = True
                HuboIngresos = True
                
                'estas variables son para compatibilizar INGRESOS
                'como el cobro es total en ambos casos anteiores solo
                'graba como items cobrados al importe original
                CapitalCobrado = CCur(CapitalRestante)
                InteresCobrado = CCur(InteresRestante)
                GastosCobrados = CCur(GastoRestante)
                SegurosCobrados = CCur(SeguroRestante)
                OtorgamientoCobrado = CCur(OtorgamientoRestante)
                Vencimiento2Cobrado = CCur(Vencimiento2Restante)
                RefinCobrado = CCur(RefinRestante)
                IvaInteresCobrado = CCur(IvaInteresRestante)
                IvaSegurosCobrado = CCur(IvaSeguroRestante)
                IvaOtorGastosCobrado = CCur(IvaOtorGastoRestante)
                MoraCobrada = CCur(ImporteMora)
                IvaMoraCobrada = CCur(IvaMora)
                
                lvcuotas.ListItems.Item(I).SubItems(42) = "ok"
                
                If Not Cubre Then
                   'reduzco el importe total cobrado
                   ImporteTotalCobrado = CCur(ImporteTotalCobrado) - CCur(SaldoCuota)
                End If
             Else
                'es uncobro parcial
                'aca solo entraria si no cubre
                If Not Cubre Then
                   CobroParcial = True
                   
                   'el importe no permite saldar toda la cuota..es parcial si o si
                   If CCur(ImporteTotalCobrado) > 0 Then
                      'este se usa para reducirlo al cobrar items
                      ImporteRealCobrado = CCur(ImporteTotalCobrado)
                      ImporteIngresos = CCur(ImporteRealCobrado)
                      
                      'obtengo el total parcial de nuevo con el ultimo recien cargado
                      'parece que esta linea esta demas y quedo de antes
                      ImporteParcial = CCur(ImporteParcial) + CCur(ImporteTotalCobrado)
                      
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
                                       
                      'si queda aun resto
                      Vencimiento2Cobrado = 0
                      'si cobre despues del 1 vto
                      If CDate(DTPicker1.Value) > CDate(Vencimiento1) Then
                          If CCur(ImporteRealCobrado) > 0 Then
                             If CCur(Vencimiento2Restante) > 0 Then
                                'si cubro todo
                                If CCur(ImporteRealCobrado) >= CCur(Vencimiento2Restante) Then
                                   ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(Vencimiento2Restante)
                                   Vencimiento2Cobrado = CCur(Vencimiento2Restante)
                                Else
                                   'cubro solo una parte
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
                      
                      'grabo el cobro parcial
                      sql = "update cuotas set " & _
                            "cobrosparciales=1," & _
                            "formacobro='" & CStr(FormaCobro) & "'," & _
                            "origen='MANUAL'," & _
                            "idcobrador=" & IdCobrador & _
                            "where idcredito=" & CLng(NumCredito) & " " & _
                            "and numcuota=" & CLng(NumCuota)
                      cnSQL.Execute sql
                      
                      HuboIngresos = True
                      HuboCobros = True
                      
                      'reduzco el total cobrado
                      ImporteTotalCobrado = CCur(ImporteTotalCobrado) - CCur(SaldoCuota)
                      lvcuotas.ListItems.Item(I).SubItems(42) = "ok"
                                   
                   End If 'si importetotalcobrado>0
                   
                End If ' si no cubre
                
             End If 'si importetotalcobrado>=saldo
             
             If HuboIngresos Then
                'codigo nuevo
                IdIngreso = UltimoId("idingreso", "ingresos") + 1
                'obgtengo el recibo/factura asociada a esta cuota
                TxtNumRecibo.Text = Format(UltimoId("ultimonumrecibo", "configuracionsistema"), "0000000")

                'AHORA GRABO LOS CAMPOS CBR
                sql = "insert into ingresos (idingreso,idcredito,numcuota," & _
                      "fechacobro,importecobrado,numrecibo,codprestamo,numcomprobante,capitalcobrado,interescobrado,vencimiento2cobrado,refincobrado,gastoscobrados,seguroscobrados,otorgamientocobrado,ivainterescobrado,ivaseguroscobrado,ivaotorgastoscobrado,moracobrada,ivamoracobrada,descuentos,recargos,usuario,origen,fechaimputacion,idcobrador,secuencia,montocobradosecuencia) " & _
                      "values(" & CLng(IdIngreso) & "," & CLng(NumCredito) & "," & CLng(NumCuota) & _
                      ",'" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "'," & ConvertirDblSql(ImporteIngresos) & _
                      ",'" & CStr(TxtNumRecibo.Text) & "','" & CStr(CodPrestamo) & "'," & CLng(NumFactura) & "," & ConvertirDblSql(CCur(CapitalCobrado)) & "," & ConvertirDblSql(CCur(InteresCobrado)) & "," & ConvertirDblSql(CCur(Vencimiento2Cobrado)) & "," & ConvertirDblSql(RefinCobrado) & "," & ConvertirDblSql(GastosCobrados) & "," & ConvertirDblSql(SegurosCobrados) & "," & _
                      ConvertirDblSql(OtorgamientoCobrado) & "," & ConvertirDblSql(IvaInteresCobrado) & "," & ConvertirDblSql(IvaSegurosCobrado) & "," & ConvertirDblSql(IvaOtorGastosCobrado) & "," & ConvertirDblSql(MoraCobrada) & "," & ConvertirDblSql(IvaMoraCobrada) & "," & ConvertirDblSql(DescuentoCuota) & "," & ConvertirDblSql(RecargoCuota) & ",'" & CStr(VG_USUARIOLOGIN) & "','MANUAL',GetDate()," & IdCobrador & "," & SecuenciaIngreso & "," & ConvertirDblSql(MontoCobradoSecuencia) & ")"
                cnSQL.Execute sql
                
                'restablesco los ingresos a falso
                HuboIngresos = False
                
                'incremento el numero de recibo guardado
                sql = "update configuracionsistema set ultimonumrecibo=ultimonumrecibo + 1"
                cnSQL.Execute sql
             
                If VG_FINALIZARAUTOMATICAMENTE Then
                   'si es la ultima cuota finalizo el credito
                   Call FinalizarCredito(NumCredito, DTPicker1.Value)
                End If
                      
                'grabo datos de cobradores
                If IdCobrador > 0 Then
                   'obtengo el proximo id de cobradorespagos
                   IdCobradorPago = UltimoId("idcobradorpago", "cobradorespagos") + 1
                   ImporteCobrador = ObtenerComisionCobrador(IdCobrador, SaldoCuota)
                   CodPrestamo = lvcuotas.ListItems.Item(I).SubItems(1)
                                
                   sql = "insert into cobradorespagos (idcobradorpago,idcobrador,idcredito,numcuota,numfactura," & _
                         "importecobrador,fecha,codprestamo) " & _
                         "values('" & CLng(IdCobradorPago) & "','" & CLng(IdCobrador) & "','" & CLng(NumCredito) & "','" & CLng(NumCuota) & "','" & CLng(NumFactura) & _
                         "'," & ImporteCobrador & ",'" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "','" & CStr(CodPrestamo) & "')"
                   cnSQL.Execute sql
                End If
              
              End If 'si hubo cobros
          
          End If  'si no esta cobrada
         
        End If  'si existe el credito etc
      
     End If ' si esta marcada
    
  Next I
  
 'actualizo excedentes si se usaron
 If HuboCobros Then
    If CheckExcedentes.Value = 1 Then
       Call ActualizarExcedentes(IdCliente, NumCredito, CodPrestamo)
       Call CargarExcedentesClientes(IdCliente)
    End If
 End If
 
 'fin de transaccion
 cnSQL.CommitTrans
    
TxtCuotasACobrar.Text = 0
TxtImporteACobrar.Text = 0

If HuboCobros Then
   
   MsgI "El cobro fue realizado exitosamente"
   
   If MsgP("¿Desea imprimir la factura?") Then
      CondicionCuotas = ObtenerCondicion()
      Call ImprimirFacturaCredimaco(CondicionCuotas, CDate(DTPicker1.Value))
   End If
   Call ActualizarListas
Else
   MsgE "No hubo cobros...revise si las cuotas ya estan cobradas o no tienen valicez"
End If

'actualizo el numero de factura de credimaco
TxtNumRecibo.Text = Format(UltimoId("ultimonumrecibo", "configuracionsistema"), "0000000")

Exit Sub
merror:
tratarerrores "Error registrando cobro de cuotas simultaneas"
End Sub
Private Sub CmdAnular_Click()
Call RefreshTimer
CmdAnular.Enabled = False
Call Anular
CmdAnular.Enabled = True
End Sub
Private Sub Anular()
'anula el cobro de una o multiples cuotas
Dim sql As String
Dim NumCredito As Long
Dim NumCuota As Long
Dim I As Long
Dim HuboAnulacion As Boolean
Dim NumRecibo As String
On Error GoTo merror

'verifico si en la lista hay cuotas seleccionadas
If Not HayFilasChequeadas(lvcuotas) Then
   MsgE "Debe seleccionar cuotas para anular"
   Exit Sub
End If

'verificar si las cuotas seleccionadas estan cobradas
If Not HayCuotasCobradas() Then
   MsgE "Debe seleccionar cuotas cobradas"
   Exit Sub
End If

If Not MsgP("¿Confirma la anulacion del cobro de las cuotas seleccionadas?") Then Exit Sub

'chequeo complementario
If Not ExisteCliente(IdCliente) Then
   MsgE "El cliente no existe"
   Exit Sub
End If

HuboAnulacion = False

'registro la anulacion
'inicio transaccion
cnSQL.BeginTrans

For I = 1 To lvcuotas.ListItems.Count()
    'si la cuota esta seleccionada
    If lvcuotas.ListItems.Item(I).Checked Then
       'si la cuota no esta cobrada ni es comodin no esta refinanciada
       'valido si es vigente
       NumCuota = CLng(lvcuotas.ListItems.Item(I).SubItems(4))
       NumCredito = CLng(lvcuotas.ListItems.Item(I).SubItems(3))
       'valido cada credito
       If ExisteCredito(NumCredito) And Not CreditoFinalizado(NumCredito) And Not CreditoBloqueado1(NumCredito) Then
          If CuotaCobrada(NumCredito, NumCuota) And Not CuotaRefinanciada(NumCredito, NumCuota) _
             And Not CuotaEsComodin(NumCredito, NumCuota) Then
             NumRecibo = ""
             sql = "update cuotas set fechacobro = Null,importecobrado=0," & _
                   "importedescuentos=0,importerecargos=0,importeparcial=0," & _
                   "importemora=0,ivamora=0," & _
                   "idmoneda=0,pagofacil='False',rapipago='False',cobrosparciales='False' " & _
                   "where idcredito='" & CLng(NumCredito) & "' and numcuota='" & CLng(NumCuota) & "'"
             cnSQL.Execute (sql)
             
             'si el credito estaba finalizado lo desmarco como no finalizado
             If CreditoFinalizado(NumCredito) Then
                sql = "update creditos set fechafinalizacion = Null " & _
                      "where idcredito='" & CLng(NumCredito) & "'"
                cnSQL.Execute (sql)
             End If
             
             'borro las entradas de cobradorespagos
             sql = "delete from cobradorespagos " & _
                   "where idcredito='" & CLng(NumCredito) & "' and numcuota='" & CLng(NumCuota) & "'"
             cnSQL.Execute sql
             
             'borro sus entradas en INGRESOS
             sql = "delete from ingresos where idcredito='" & CLng(NumCredito) & "' and numcuota='" & CLng(NumCuota) & "'"
             cnSQL.Execute sql
             HuboAnulacion = True
         End If
       End If
    End If
    
Next I

'fin de transaccion
cnSQL.CommitTrans

TxtCuotasACobrar.Text = 0
TxtImporteACobrar.Text = 0

Call ActualizarListas

If HuboAnulacion Then
   MsgI "La anulacion fue realizada exitosamente"
Else
   MsgE "No hubo anulacion...(verifique si las cuotas estan cobradas)"
End If

Exit Sub
merror:
tratarerrores "Error registrando anulacion de cuotas simultaneas"
End Sub
Private Function HayCuotasCobradas() As Boolean
'verifica si en la lista hay seleccionadas cuotas cobradas..
Dim I As Long
On Error GoTo merror

HayCuotasCobradas = False

'busco si entre las seleccionadas hay alguna cobrada
For I = 1 To lvcuotas.ListItems.Count()
    'si esta marcada
    If lvcuotas.ListItems.Item(I).Checked Then
       'si hay una fecha de cobro
       If lvcuotas.ListItems.Item(I).SubItems(24) <> "" Then
          HayCuotasCobradas = True
          Exit Function
       End If
    End If
Next I

Exit Function
merror:
tratarerrores "Error en funcion HayCuotasCobradas"
End Function

Private Sub lvcuotas_DblClick()
'muestra los cobros parciales de la cuota seleccionada
Dim Credito As Long
Dim Cuota As Long
Dim Saldo As Currency
Dim Cliente As Long
Dim Factura As Long
On Error GoTo merror

If Not VerificarSeleccionLista(lvcuotas) Then Exit Sub

Credito = CLng(lvcuotas.SelectedItem.SubItems(3))
Cuota = CLng(lvcuotas.SelectedItem.SubItems(4))
Saldo = CCur(lvcuotas.SelectedItem.SubItems(23))
Factura = CLng(lvcuotas.SelectedItem.SubItems(5))

If TieneCobrosParciales(Credito, Cuota) Then
   'si la cuota tiene cobros parciales las muestra
   If ObtenerImporteParcialX(Credito, Cuota) > 0 Then
      FrmCobrosParciales.xnumcredito = Credito
      FrmCobrosParciales.xnumcuota = Cuota
      FrmCobrosParciales.ximporteactualizado = Saldo
      FrmCobrosParciales.xidcliente = IdCliente
      FrmCobrosParciales.xfactura = Factura
   
      Call CenterForm(FrmCobrosParciales)
      FrmCobrosParciales.Show vbModal
      'por si hubo cambios
      Call ActualizarListas
  End If
End If

Exit Sub
merror:
tratarerrores "Error seleccionando cobros parciales de cuotas"
End Sub
Private Sub LvCuotas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'marcar/desmarcar una fila de la lista de cuotas
Dim ImporteTotalCuota As Currency
Dim CapitalCuota As Currency
Dim InteresCuota As Currency
Dim ImporteDescuento As Currency
Dim Importe As Currency
On Error GoTo merror

'saldo
ImporteTotalCuota = CCur(lvcuotas.ListItems.Item(Item.Index).SubItems(23))
   
'si estoy tildando
If lvcuotas.ListItems.Item(Item.Index).Checked Then
     
   'si no esta cobrada incremento el total a cobrar
   If lvcuotas.ListItems.Item(Item.Index).SubItems(24) = "" Then
      'incremento el saldo
      TxtImporteACobrar.Text = CCur(TxtImporteACobrar.Text) + CCur(lvcuotas.ListItems.Item(Item.Index).SubItems(23))
      TxtCuotasACobrar.Text = CLng(TxtCuotasACobrar.Text) + 1
   End If
Else
   'si desmarco le saco el descuento y recargo
   'si no esta cobrada
   If lvcuotas.ListItems.Item(Item.Index).SubItems(24) = "" Then
      'decremento el saldo
      TxtImporteACobrar.Text = CCur(TxtImporteACobrar.Text) - CCur(lvcuotas.ListItems.Item(Item.Index).SubItems(23))
      TxtCuotasACobrar.Text = CLng(TxtCuotasACobrar.Text) - 1
   End If
   
End If

TxtImporteACobrar.Text = Format(TxtImporteACobrar.Text, "0.00")

TxtSubtotal.Text = ReconstruirImporte
TxtSubtotal.Text = Format(TxtSubtotal.Text, "0.00")

Exit Sub
merror:
tratarerrores "Error seleccionando cuota"
End Sub
Private Sub SetearEntorno()
On Error GoTo merror

TxtImporteACobrar.Text = 0
TxtImporteDescuento.Text = 0
TxtImporteRecargo.Text = 0
TxtSubtotal.Text = 0
TxtImporteRecibido.Text = ""
TxtVuelto.Text = 0
CheckAplicarDescuentos.Value = 0
CheckAplicarRecargos.Value = 0

If lvcuotas.ListItems.Count = 0 Then
   CmdCobrar.Enabled = False
   CmdAnular.Enabled = False
Else
   If VG_COBRA Then
      CmdCobrar.Enabled = True
   End If
   If VG_ANULA Then
      CmdAnular.Enabled = True
   End If
End If

Exit Sub
merror:
tratarerrores "Error seteando entorno-Cobros Masivos"
End Sub
Private Sub CheckAplicarDescuentos_Click()
On Error GoTo merror

If CheckAplicarDescuentos.Value = 1 Then
   CheckAplicarRecargos.Value = 0
   TxtImporteRecargo.Text = 0
   TxtImporteRecargo.Enabled = False
   TxtImporteDescuento.Enabled = True
   TxtImporteDescuento.BackColor = vbWhite
Else
   TxtImporteDescuento.Enabled = False
   TxtImporteDescuento.Text = 0
   TxtImporteDescuento.BackColor = &HFFFFC0
End If

If Trim(TxtImporteACobrar.Text) = "" Then Exit Sub
        
TxtSubtotal.Text = ReconstruirImporte

TxtSubtotal.Text = Format(TxtSubtotal.Text, "0.00")

Call ActualizarVuelto

Exit Sub
merror:
tratarerrores "Error aplicando descuentos"
End Sub
Private Sub CheckAplicarRecargos_Click()
On Error GoTo merror

If CheckAplicarRecargos.Value = 1 Then
   CheckAplicarDescuentos.Value = 0
   TxtImporteDescuento.Text = 0
   TxtImporteDescuento.Enabled = False
   TxtImporteRecargo.Enabled = True
   TxtImporteRecargo.BackColor = vbWhite
Else
   TxtImporteRecargo.Enabled = False
   TxtImporteRecargo.Text = 0
   TxtImporteRecargo.BackColor = &HFFFFC0
End If

If Trim(TxtImporteACobrar.Text) = "" Then Exit Sub

TxtSubtotal.Text = ReconstruirImporte

TxtSubtotal.Text = Format(TxtSubtotal.Text, "0.00")

Call ActualizarVuelto

Exit Sub
merror:
tratarerrores "Error aplicando recargos"
End Sub
Private Sub TxtImporteDescuento_Change()
On Error GoTo merror

If Trim(TxtImporteACobrar.Text) = "" Then Exit Sub
If CCur(TxtImporteACobrar.Text) = 0 Then Exit Sub
   
If Trim(TxtImporteDescuento.Text) = "" Then Exit Sub
If Not IsNumeric(TxtImporteDescuento.Text) Then Exit Sub
If CCur(TxtImporteDescuento.Text) < 0 Then Exit Sub
If CCur(TxtImporteDescuento.Text) > CCur(TxtImporteACobrar.Text) Then
   TxtSubtotal.Text = 0
   Exit Sub
End If

TxtSubtotal.Text = ReconstruirImporte()
TxtSubtotal.Text = Format(TxtSubtotal.Text, "0.00")

Call ActualizarVuelto

Exit Sub
merror:
tratarerrores "Error en procedimiento TxtImporteDescuento_Change"
End Sub
Private Sub TxtImporteRecargo_Change()
On Error GoTo merror

If Trim(TxtImporteACobrar.Text) = "" Then Exit Sub
If CCur(TxtImporteACobrar.Text) = 0 Then Exit Sub
   
If Trim(TxtImporteRecargo.Text) = "" Then Exit Sub
If Not IsNumeric(TxtImporteRecargo.Text) Then Exit Sub
If CCur(TxtImporteRecargo.Text) < 0 Then Exit Sub
   
TxtSubtotal.Text = ReconstruirImporte()
TxtSubtotal.Text = Format(TxtSubtotal.Text, "0.00")

Call ActualizarVuelto

Exit Sub
merror:
tratarerrores "Error en procedimiento TxtImporteRecargo_Change"
End Sub
Private Sub LvCuotas_KeyDown(KeyCode As Integer, Shift As Integer)
'si apretaron enter en una cuota
If KeyCode = vbKeyReturn Then
   Call lvcuotas_DblClick
End If
End Sub
Private Sub ActualizarListas()
Call CargarCuotasCreditos
Call SetearEntorno
End Sub
Private Sub TxtCliente_Change()
'si cambia el cliente
Call ActualizarListas
Call CargarExcedentesClientes(IdCliente)
End Sub
Private Sub CmdBuscar_Click()
'buscar por rango de fechas
Call ActualizarListas
End Sub
Private Function HayCreditosDistintos(ByVal lv As ListView) As Boolean
'verifico si en una lista se seleccionaron lineas de distintos creditos
Dim I As Long
Dim IdCredito1 As Long
Dim IdCredito2 As Long
On Error GoTo merror

HayCreditosDistintos = False

'saco el primer credito seleccionado
For I = 1 To CLng(lv.ListItems.Count())
    If lv.ListItems.Item(I).Checked Then
       IdCredito1 = CLng(lv.ListItems.Item(I).SubItems(3))
       Exit For
    End If
Next I

'recorro la lista
For I = 1 To CLng(lv.ListItems.Count())
    If lv.ListItems.Item(I).Checked Then
       IdCredito2 = CLng(lv.ListItems.Item(I).SubItems(3))
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
Private Sub TxtImporteACobrar_Change()
TxtSubtotal.Text = TxtImporteACobrar.Text
Call ActualizarVuelto
End Sub
Private Sub ActualizarVuelto()
Dim Importe As Currency
On Error GoTo merror

TxtVuelto.Text = 0

If Trim(TxtSubtotal.Text) = "" Then Exit Sub
   
If Not IsNumeric(TxtSubtotal.Text) Then Exit Sub
   
If CCur(TxtSubtotal.Text) <= 0 Then Exit Sub

If Trim(TxtImporteRecibido.Text) = "" Then Exit Sub
   
If Not IsNumeric(TxtImporteRecibido.Text) Then Exit Sub
   
If CCur(TxtImporteRecibido.Text) <= 0 Then Exit Sub
   
Importe = CCur(TxtImporteRecibido.Text)

If CCur(Importe) > CCur(TxtSubtotal.Text) Then
   TxtVuelto.Text = CCur(Importe) - CCur(TxtSubtotal.Text)
End If

TxtVuelto.Text = Format(CCur(TxtVuelto.Text), "0.00")

If CCur(TxtImporteRecibido.Text) >= CCur(TxtSubtotal.Text) Then
   TxtMensaje.Text = "COBRO TOTAL"
Else
   TxtMensaje.Text = "COBRO PARCIAL"
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento ActualizarVuelto"
End Sub
Private Sub TxtImporteRecibido_Change()
Call ActualizarVuelto
End Sub
Private Function DatosImpresionOk() As Boolean

DatosImpresionOk = True

'verificar si hay cuotas
If lvcuotas.ListItems.Count = 0 Then
   DatosImpresionOk = False
   Exit Function
End If

'verificar si hay filas chequeadas
If Not HayFilasChequeadas(lvcuotas) Then
   MsgE "No hay cuotas seleccionadas"
   DatosImpresionOk = False
   Exit Function
End If

End Function
Private Sub DTPicker1_Change()
Call ActualizarListas
End Sub
Private Function ReconstruirImporte() As Currency
Dim Interes As Currency
Dim Gastos As Currency
Dim Seguros As Currency
Dim Impuestos As Currency
Dim Descuentos As Currency
Dim Recargos As Currency
Dim Base As Currency
Dim ImporteReal As Currency
On Error GoTo merror

ReconstruirImporte = 0

Base = CCur(TxtImporteACobrar.Text)

Descuentos = 0
If CheckAplicarDescuentos.Value = 1 Then
   If Trim(TxtImporteDescuento.Text) = "" Then
      TxtImporteDescuento.Text = 0
   End If
   If Not IsNumeric(TxtImporteDescuento.Text) Then
      TxtImporteDescuento.Text = 0
   End If
   If CCur(TxtImporteDescuento.Text) <= 0 Then
      TxtImporteDescuento.Text = 0
   End If
   
   Descuentos = CCur(TxtImporteDescuento.Text)
End If

Recargos = 0
If CheckAplicarRecargos.Value = 1 Then
   If Trim(TxtImporteRecargo.Text) = "" Then
      TxtImporteRecargo.Text = 0
   End If
   If Not IsNumeric(TxtImporteRecargo.Text) Then
      TxtImporteRecargo.Text = 0
   End If
   If CCur(TxtImporteRecargo.Text) <= 0 Then
      TxtImporteRecargo.Text = 0
   End If
   
   Recargos = CCur(TxtImporteRecargo.Text)
End If
'**

ImporteReal = CCur(Base) - CCur(Descuentos) + CCur(Recargos)

ReconstruirImporte = CCur(ImporteReal)

Exit Function
merror:
tratarerrores "Error en funcion ReconstruirImporte"
End Function
Private Sub CheckTodos_Click()
'permite listar todas las cuotas vencidas
'si seleccione todos deshabilito la seleccion de cliente
lvcuotas.ListItems.Clear
Lv2.ListItems.Clear
CheckExcedentes.Value = 0

If CheckTodos.Value = 1 Then
   'deshabilito los demas checks
   CheckCobradas.Value = 0
   CheckCobradas.Enabled = False
   TxtCliente.Text = ""
   IdCliente = 0
   TxtCliente.Enabled = False
   CmdBuscarCliente.Enabled = False
Else
   'habilito los demas checks
   CheckCobradas.Enabled = True
   TxtCliente.Text = ""
   IdCliente = 0
   TxtCliente.Enabled = True
   CmdBuscarCliente.Enabled = True
End If

Call ActualizarListas

End Sub
Private Function ObtenerCondicion() As String
'recorre la lista de cuotas y obtiene las seleccionadas
'PRECAUCION DE NO SELECCIONAR MAS DE UNA CANTIDAD RAZONABLE 99 O 100 PORQUE
'SE CORRE EL RIEGO DE ERRORES AL EJECUTAR LAS SENTENCIAS DE CONSULTA SQL
Dim CondicionCuotas As String
Dim I As Long
Dim IdCredito As Long
Dim NumCuota As Long
Dim Vez As Long
On Error GoTo merror

ObtenerCondicion = ""

CondicionCuotas = ""

Vez = 1

For I = 1 To lvcuotas.ListItems.Count()
    If lvcuotas.ListItems.Item(I).Checked Then
       'si fue cobrada en ese momento
       If lvcuotas.ListItems.Item(I).SubItems(42) = "ok" Then
          IdCredito = CLng(lvcuotas.ListItems.Item(I).SubItems(3))
          NumCuota = CLng(lvcuotas.ListItems.Item(I).SubItems(4))
       
          If Vez = 1 Then
             CondicionCuotas = "(creditos.idcredito='" & CLng(IdCredito) & "' and cuotas.numcuota='" & CLng(NumCuota) & "')"
             Vez = Vez + 1
          Else
             CondicionCuotas = CondicionCuotas & " or (creditos.idcredito='" & CLng(IdCredito) & "' and cuotas.numcuota='" & CLng(NumCuota) & "')"
          End If
       End If
    End If
Next I

ObtenerCondicion = CondicionCuotas

Exit Function
merror:
tratarerrores "Error en funcion ObtenerCondicion"
End Function
'FUNCIONES DE EXCEDENTES
Private Sub CheckExcedentes_Click()
'se ejecuta al aplicar excedentes o no
Dim I As Long
On Error GoTo merror

TxtImporteRecibido.Text = 0

'si habilita el uso de excedentes
If CheckExcedentes.Value = 1 Then
   Lv2.Enabled = True
   TxtImporteRecibido.Locked = True
'   TxtImporteRecibido.BackColor = &H80000013
Else
   'si no usa excedentes deshabilita y desmarcar
   Lv2.Enabled = False
   TxtImporteRecibido.Locked = False
   TxtImporteRecibido.BackColor = vbWhite
   
   If Lv2.ListItems.Count() = 0 Then Exit Sub
   
   For I = 1 To Lv2.ListItems.Count
       If Lv2.ListItems.Item(I).Checked Then
          Lv2.ListItems.Item(I).Checked = False
       End If
   Next I
End If

Exit Sub
merror:
tratarerrores "Error habilitando/deshabilitando excedentes"
End Sub
Private Sub CargarExcedentesClientes(ByVal IdCliente As Long)
'carga la lista de excedentes del cliente actual
Dim sql As String
Dim rec As rdoResultset
Dim Cad1 As String
On Error GoTo merror

If Trim(TxtCliente.Text) = "" Then Exit Sub

Lv2.ListItems.Clear

'carga los excedentes aun no procesados
sql = "select * from excedentesclientes " & _
      "where idcliente='" & CLng(IdCliente) & "' " & _
      "and fechaproceso is Null"
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Do While Not rec.EOF
      Cad1 = ""
      Set Nitem = Lv2.ListItems.Add(, , Cad1)
      Nitem.SubItems(1) = Format(rec.rdoColumns("idexcedentecliente"), "000000") & vbNullString
      Nitem.SubItems(2) = Format(rec.rdoColumns("idcliente"), "000000") & vbNullString
      Nitem.SubItems(3) = rec.rdoColumns("codprestamo") & vbNullString
      Nitem.SubItems(4) = Format(rec.rdoColumns("idcredito"), "000000") & vbNullString
      Nitem.SubItems(5) = Format(rec.rdoColumns("numcuota"), "000") & vbNullString
      Nitem.SubItems(6) = rec.rdoColumns("fechacobro") & vbNullString
      Nitem.SubItems(7) = Format(rec.rdoColumns("importecobro"), "0.00") & vbNullString
      Nitem.SubItems(8) = rec.rdoColumns("observaciones") & vbNullString
      Nitem.SubItems(9) = rec.rdoColumns("archivorp") & vbNullString
      
      rec.MoveNext
   Loop
End If


Exit Sub
merror:
tratarerrores "Error cargando excedentes de clientes"
End Sub
Private Sub ActualizarExcedentes(ByVal IdCliente As Long, ByVal IdCredito As Long, ByVal CodPrestamo As String)
'esta se ejecuta solo al registrar un cobro con excedentes
'descuenta los excedentes usados
'verificar si uso todo o no(lo no usado lo recicla y lo agrega)
'por ejemplo si hubo vuelto no se uso todo
'si vuelto es cero uso todo lo seleccionado
Dim sql As String
Dim rec As rdoResultset
Dim I As Long
Dim IdExcedenteCliente As Long
Dim Observaciones As String
On Error GoTo merror

For I = 1 To Lv2.ListItems.Count
    If Lv2.ListItems.Item(I).Checked Then
       IdExcedenteCliente = Lv2.ListItems.Item(I).SubItems(1)
       'si uso todo lo doy por usado
       sql = "update excedentesclientes set fechaproceso='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "'" & _
             "where idexcedentecliente='" & CLng(IdExcedenteCliente) & "'"
       cnSQL.Execute sql
    End If
Next I

'si no uso todo
If CCur(TxtVuelto.Text) > 0 Then
   Observaciones = "Excedente agregado por el sist.(hubo vuelto)"
  'si no uso todo agrego el vuelto como excednte no usado
   IdExcedenteCliente = UltimoId("idexcedentecliente", "excedentesclientes") + 1
   sql = "insert into excedentesclientes(idexcedentecliente,idcliente,codprestamo,idcredito,fechacobro,importecobro,observaciones) " & _
         "values('" & CLng(IdExcedenteCliente) & "','" & CLng(IdCliente) & "','" & CStr(CodPrestamo) & "','" & CLng(IdCredito) & "','" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "','" & ConvertirDblSql(CCur(TxtVuelto.Text)) & "','" & CStr(Observaciones) & "')"
   cnSQL.Execute (sql)
End If

Exit Sub
merror:
tratarerrores "Error actualizando excedentes de clientes"
End Sub
Private Sub Lv2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'al marcar/desmarcar excedentes
Dim Importe As Currency
On Error GoTo merror

'si no hay items
If Lv2.ListItems.Count() = 0 Then Exit Sub

'saldo
Importe = CCur(Lv2.ListItems.Item(Item.Index).SubItems(7))

If Trim(TxtImporteRecibido.Text) = "" Then
   TxtImporteRecibido.Text = 0
End If
If Not IsNumeric(TxtImporteRecibido.Text) Then
   TxtImporteRecibido.Text = 0
End If

'si estoy tildando
If Lv2.ListItems.Item(Item.Index).Checked Then
   'incremento el saldo
   TxtImporteRecibido.Text = CCur(TxtImporteRecibido.Text) + CCur(Importe)
Else
   'si desmarco le saco el importe
   TxtImporteRecibido.Text = CCur(TxtImporteRecibido.Text) - CCur(Importe)
End If

Exit Sub
merror:
tratarerrores "Error tildando excedentes de clientes"
End Sub
Private Sub CmdImprimirExcedentes_Click()
'imprime la lista de excedentes de rapipago del cliente actual
Dim sql As String
Dim rec As rdoResultset
Dim Mreporte As New ARListadoExcedentes
Dim Titulo As String
On Error GoTo merror
Call RefreshTimer

'si imprimo todos los cobros parciales de la cuota
Titulo = "Listado de excedentes de RapiPago del cliente:" & TxtCliente.Text & " a la fecha:" & CStr(DTPicker1.Value)

If Trim(TxtCliente.Text) = "" Then Exit Sub
If CLng(IdCliente) <= 0 Then Exit Sub

If Lv2.ListItems.Count() = 0 Then Exit Sub

sql = "select clientes.apellido + ' ' + clientes.nombre as cliente," & _
      "excedentesclientes.codprestamo,excedentesclientes.archivorp," & _
      "excedentesclientes.fechacobro,excedentesclientes.importecobro," & _
      "excedentesclientes.observaciones,excedentesclientes.idcliente," & _
      "excedentesclientes.numcuota " & _
      "from clientes inner join excedentesclientes on clientes.idcliente=excedentesclientes.idcliente " & _
      "where excedentesclientes.idcliente='" & CLng(IdCliente) & "' " & _
      "and fechaproceso is Null"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Mreporte.RDODataControl1.Resultset = rec
   Mreporte.Caption = "Imprimir listado de excedentes de RapiPago de clientes"
   Mreporte.LabelTitulo = Titulo
   Mreporte.Show vbModal
Else
   MsgE "No hay excedentes para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo listado de excedentes"
End Sub
Private Sub CmdTodosExcedentes_Click()
'imprime la lista total de excedentes de rapipago aun no procesados o
'descontados
Dim sql As String
Dim rec As rdoResultset
Dim Mreporte As New ARListadoExcedentes
Dim Titulo As String
On Error GoTo merror
Call RefreshTimer

'si imprimo todos los cobros parciales de la cuota
Titulo = "Listado total de excedentes de RapiPago a la fecha:" & CStr(DTPicker1.Value)

sql = "select clientes.apellido + ' ' + clientes.nombre as cliente," & _
      "excedentesclientes.codprestamo,excedentesclientes.archivorp," & _
      "excedentesclientes.fechacobro,excedentesclientes.importecobro," & _
      "excedentesclientes.observaciones,excedentesclientes.idcliente," & _
      "excedentesclientes.numcuota " & _
      "from clientes inner join excedentesclientes on clientes.idcliente=excedentesclientes.idcliente " & _
      "where fechaproceso is Null"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Mreporte.RDODataControl1.Resultset = rec
   Mreporte.Caption = "Imprimir listado total de excedentes de RapiPago de clientes"
   Mreporte.LabelTitulo = Titulo
   Mreporte.Show vbModal
Else
   MsgE "No hay excedentes para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo lista total de excedentes"
End Sub

