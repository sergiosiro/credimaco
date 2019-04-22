VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRegistrarCreditos 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Nuevos Creditos"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   HelpContextID   =   16
   Icon            =   "FrmRegistrarCreditos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdImprimirResumen 
      Caption         =   "&Imprimir resumen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      ToolTipText     =   "Imprime el resumen del credito y cuotas"
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5280
      TabIndex        =   17
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7920
      Width           =   2025
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "&Registrar credito"
      Height          =   375
      Left            =   120
      Picture         =   "FrmRegistrarCreditos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Registra un nuevo credito"
      Top             =   7920
      Width           =   2025
   End
   Begin VB.Frame fmeDatos 
      ForeColor       =   &H00FF0000&
      Height          =   7815
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   9975
      Begin VB.CheckBox ChkDDJJ 
         Caption         =   "CERTIFICACION"
         Height          =   195
         Left            =   8160
         TabIndex        =   77
         Top             =   6960
         Width           =   1575
      End
      Begin VB.ComboBox ComboVendedores 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2400
         Width           =   2415
      End
      Begin VB.ComboBox ComboProvincias 
         Height          =   315
         ItemData        =   "FrmRegistrarCreditos.frx":0D1C
         Left            =   4440
         List            =   "FrmRegistrarCreditos.frx":0D1E
         Style           =   2  'Dropdown List
         TabIndex        =   71
         ToolTipText     =   "Provincia donde se registro el credito"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox TxtTotalSellados 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   68
         Tag             =   "N"
         ToolTipText     =   "Importe de sellados"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox TxtCodPrestamo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   8
         ToolTipText     =   "Codigo de prestamo"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         Caption         =   "Seleccione un cliente:"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   9735
         Begin VB.CommandButton CmdVerificarRequisitosCliente 
            Height          =   255
            Left            =   8640
            Picture         =   "FrmRegistrarCreditos.frx":0D20
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Verifica si el cliente seleccionado cumple los requisitos de credito"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton CmdCliente 
            Height          =   375
            Left            =   7800
            Picture         =   "FrmRegistrarCreditos.frx":1722
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Permite seleccionar al cliente de una lista"
            Top             =   180
            Width           =   615
         End
         Begin VB.TextBox TxtTitular 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   19
            ToolTipText     =   "Cliente al cual le otorgaremos un credito"
            Top             =   240
            Width           =   7575
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   9735
         Begin VB.TextBox TxtFinanciar 
            Height          =   285
            Left            =   1560
            MaxLength       =   9
            TabIndex        =   3
            ToolTipText     =   "Importe solicitado por el cliente"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox ComboPlanes 
            Height          =   315
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox TxtImporteAFinanciar 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            MaxLength       =   9
            TabIndex        =   4
            ToolTipText     =   "Importe solicitado por el cliente"
            Top             =   240
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   255
            Left            =   6720
            TabIndex        =   6
            ToolTipText     =   "Fecha de vencimiento de la primer cuota"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   16711681
            CurrentDate     =   39018
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   255
            Left            =   8400
            TabIndex        =   7
            ToolTipText     =   "Fecha del segundo vencimiento de la primer cuota"
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   16711681
            CurrentDate     =   39043
         End
         Begin VB.Label Label28 
            Caption         =   "Monto a Financiar:"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Plan:"
            Height          =   255
            Left            =   4320
            TabIndex        =   54
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "1º Vto:"
            Height          =   255
            Left            =   6240
            TabIndex        =   46
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Importe $:"
            Height          =   255
            Left            =   2640
            TabIndex        =   45
            Top             =   240
            Width           =   735
         End
         Begin VB.Label LabelVencimiento2 
            Caption         =   "2º Vto:"
            Height          =   255
            Left            =   7920
            TabIndex        =   47
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   9735
         Begin VB.ComboBox ComboComercios 
            Height          =   315
            ItemData        =   "FrmRegistrarCreditos.frx":1ADD
            Left            =   7320
            List            =   "FrmRegistrarCreditos.frx":1ADF
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox TxtComercio 
            Height          =   285
            Left            =   7440
            MaxLength       =   255
            TabIndex        =   72
            ToolTipText     =   "Comercio del credito"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtTotalPTF 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   5880
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   67
            Tag             =   "N"
            ToolTipText     =   "Importe total PTF"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtTasa2 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   21
            Tag             =   "N"
            ToolTipText     =   "Tasa TEM del plan seleccionado"
            Top             =   240
            Width           =   600
         End
         Begin VB.TextBox TxtCantidadCuotas 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   20
            Tag             =   "N"
            ToolTipText     =   "Nº de cuotas del nuevo credito"
            Top             =   240
            Width           =   480
         End
         Begin VB.TextBox TxtTasaFinanciacion 
            Height          =   285
            Left            =   3720
            MaxLength       =   6
            TabIndex        =   2
            Tag             =   "N"
            ToolTipText     =   "Tasa T.NA del plan seleccionado"
            Top             =   240
            Width           =   600
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Left            =   600
            TabIndex        =   1
            ToolTipText     =   "Fecha de otorgamiento del credito"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   16711681
            CurrentDate     =   39018
         End
         Begin VB.TextBox TxtNumCredito 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   43
            Tag             =   "N"
            ToolTipText     =   "Nº del nuevo credito"
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "Com:"
            Height          =   255
            Left            =   6960
            TabIndex        =   73
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label22 
            Caption         =   "PTF:"
            Height          =   255
            Left            =   5520
            TabIndex        =   66
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            Caption         =   "TEM:"
            Height          =   255
            Left            =   4440
            TabIndex        =   53
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Cuotas:"
            Height          =   255
            Left            =   2040
            TabIndex        =   52
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "TNA %:"
            Height          =   255
            Left            =   3120
            TabIndex        =   51
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Fecha:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label13 
            Caption         =   "Credito Nº:"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Lista de cuotas del nuevo credito:"
         ForeColor       =   &H00FF0000&
         Height          =   3735
         Left            =   120
         TabIndex        =   35
         Top             =   2760
         Width           =   9735
         Begin VB.TextBox TxtImporteTotal 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   7800
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   30
            Tag             =   "N"
            ToolTipText     =   "Importe total del credito"
            Top             =   3360
            Width           =   1695
         End
         Begin VB.TextBox TxtTotalOtorgamiento 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   6480
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   28
            Tag             =   "N"
            ToolTipText     =   "Gastos de otorgamiento"
            Top             =   3360
            Width           =   975
         End
         Begin VB.TextBox TxtTotalIvaGastos 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   5400
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   27
            Tag             =   "N"
            ToolTipText     =   "Importe total de impuestos"
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox TxtTotalIvaSeguros 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   4440
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   26
            Tag             =   "N"
            ToolTipText     =   "Importe total de impuestos"
            Top             =   3360
            Width           =   975
         End
         Begin VB.TextBox TxtTotalIvaInteres 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   3360
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   25
            Tag             =   "N"
            ToolTipText     =   "Importe total de impuestos"
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox TxtTotalSeguros 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   2280
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   24
            Tag             =   "N"
            ToolTipText     =   "Importe total de seguros"
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox TxtTotalGastos 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   23
            Tag             =   "N"
            ToolTipText     =   "Importe total de gastos administrativos"
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox TxtTotalInteres 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   22
            Tag             =   "N"
            ToolTipText     =   "Importe total de interes"
            Top             =   3360
            Width           =   1095
         End
         Begin MSComctlLib.ListView lvcuotas 
            Height          =   2895
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Lista de cuotas del nuevo credito"
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   5106
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
            NumItems        =   14
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
               Object.Width           =   2117
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
               Object.Width           =   2117
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
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "2º Vto"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.TextBox TxtImporteFinanciado 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   57
            Tag             =   "N"
            ToolTipText     =   "Importe solicitado financiado"
            Top             =   3360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TxtTotalImpuestos 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   7800
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   29
            Tag             =   "N"
            ToolTipText     =   "Importe total de impuestos"
            Top             =   3360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Credito total $:"
            Height          =   255
            Left            =   7800
            TabIndex        =   65
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Otorgam."
            Height          =   255
            Left            =   6480
            TabIndex        =   64
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "Iva Ot/Gast"
            Height          =   255
            Left            =   5400
            TabIndex        =   63
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Iva seguros"
            Height          =   255
            Left            =   4440
            TabIndex        =   62
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "Iva Interes"
            Height          =   255
            Left            =   3360
            TabIndex        =   61
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Seguros:"
            Height          =   255
            Left            =   2280
            TabIndex        =   60
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Gastos:"
            Height          =   255
            Left            =   1200
            TabIndex        =   59
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Total interes:"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.financiado"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   3240
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Forma de entrega del prestamo al cliente:"
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   120
         TabIndex        =   36
         Top             =   6480
         Width           =   9735
         Begin VB.CheckBox ChEntrega 
            Caption         =   "Producto con Entrega"
            Height          =   195
            Left            =   7680
            TabIndex        =   79
            Top             =   120
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   255
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Fecha de entrega del prestamo al cliente"
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Format          =   16711681
            CurrentDate     =   39098
         End
         Begin VB.TextBox TxtSonPesos 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   31
            ToolTipText     =   "Monto del prestamo en letras"
            Top             =   840
            Width           =   7095
         End
         Begin VB.TextBox TxtNumCheque 
            Height          =   285
            Left            =   4200
            MaxLength       =   25
            TabIndex        =   13
            ToolTipText     =   "Nº de cheque"
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox ComboBancos 
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Banco del cheque"
            Top             =   480
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.ComboBox ComboModalidad 
            Height          =   315
            ItemData        =   "FrmRegistrarCreditos.frx":1AE1
            Left            =   1680
            List            =   "FrmRegistrarCreditos.frx":1AEE
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Tag             =   "no"
            ToolTipText     =   "Forma de entrega del prestamo al cliente"
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label15 
            Caption         =   "Forma de entrega:"
            Height          =   255
            Left            =   1680
            TabIndex        =   49
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label14 
            Caption         =   "Fecha de entrega:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label24 
            Caption         =   "Son pesos:"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Width           =   855
         End
         Begin VB.Label LabelNumCheque 
            Caption         =   "Nº cheque:"
            Height          =   255
            Left            =   4320
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label LabelBanco 
            Caption         =   "Banco:"
            Height          =   255
            Left            =   6120
            TabIndex        =   39
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.TextBox TxtObservaciones 
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   9
         ToolTipText     =   "Observaciones del credito"
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label27 
         Caption         =   "Vendedor:"
         Height          =   255
         Left            =   7320
         TabIndex        =   75
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   4440
         TabIndex        =   70
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Sellado:"
         Height          =   255
         Left            =   6240
         TabIndex        =   69
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Cod.Prestamo:"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Obs:"
         Height          =   255
         Left            =   1560
         TabIndex        =   38
         Top             =   2160
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   7800
      TabIndex        =   18
      ToolTipText     =   "Cierra la pantalla"
      Top             =   7920
      Width           =   2025
   End
End
Attribute VB_Name = "FrmRegistrarCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***AQUI SE REGISTRAN LOS CREDITOS NUEVOS
Public IdCliente As Long
Public BandComercio As Boolean
Public VF_CANT_DIAS As Integer

Private Sub ChEntrega_Click()
Call TxtFinanciar_LostFocus
End Sub

Private Sub ChkDDJJ_Click()
'si cambio la condicion de certificacion
Call CalcularImportes
End Sub

Private Sub ComboComercios_Click()
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

BandComercio = True
sql = "select * from comercios WHERE idcomercio = '" & CLng(ComboComercios.ItemData(ComboComercios.ListIndex)) & "'"

Set rec = cnSQL.OpenResultset(sql)
If Not rec.EOF Then
For I = 0 To comboprovincias.ListCount - 1
      If comboprovincias.ItemData(I) = rec.rdoColumns("idprovincia") Then
       
       comboprovincias.ListIndex = I
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

BandComercio = False
Call RefrescarOpcionesSistema

IdCliente = 0
VF_CANT_DIAS = 0

Call LimpiarCampos(Me)

ComboModalidad.ListIndex = 0

Call CargarComboProvincias("provincias", comboprovincias)

Call CargarCombo2("bancos", ComboBancos)

Call CargarCombo2("comercios", ComboComercios)

Call CargarCombo2("vendedores", ComboVendedores)

'cargo los planes activos de creditos
Call CargarComboPlanes(ComboPlanes)
ComboPlanes.ListIndex = -1

TxtNumCredito.Text = Format(UltimoId("idcredito", "creditos") + 1, "000000")

'si hay segundo vencimiento muestro las columnas
If VG_APLICARSEGUNDOVENCIMIENTO Then
   DTPicker3.Visible = True
   LabelVencimiento2.Visible = True
Else
   lvcuotas.ColumnHeaders(13).Width = 0
   lvcuotas.ColumnHeaders(14).Width = 0
End If

Exit Sub
merror:
tratarerrores "Error cargando la pantalla RegistrarCreditos"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
IdCliente = 0
Unload Me
End Sub
Private Sub cmdnuevo_Click()
Call RefreshTimer
CmdGenerar.Enabled = True
Call LimpiarCampos(Me)
TxtNumCredito.Text = Format(UltimoId("idcredito", "creditos") + 1, "000000")
IdCliente = 0
CmdImprimirResumen.Enabled = False
End Sub
Private Sub CmdCliente_Click()
'seleccionar cliente
FrmClientesAbm.FormularioPadre = "REGISTRARCREDITOS1"
Call CenterForm(FrmClientesAbm)
FrmClientesAbm.Show
End Sub
Private Sub CmdVerificarRequisitoscliente_Click()
CmdVerificarRequisitosCliente.Enabled = False

If DatosClienteOk() Then
   MsgI "El cliente seleccionado cumple con los requisitos"
End If

CmdVerificarRequisitosCliente.Enabled = True
End Sub
Private Sub CmdImprimirResumen_Click()
'imprime el resumen del credito
Call RefreshTimer
CmdImprimirResumen.Enabled = False
If DatosImpresionResumenOk() Then
   Call ImprimirResumenCredito(TxtNumCredito.Text, Date)
End If
CmdImprimirResumen.Enabled = True
End Sub
Private Function DatosImpresionResumenOk() As Boolean
'valido solo los mas importantes..si el credito existe lo muestra
On Error GoTo merror

DatosImpresionResumenOk = True

'valido el cliente
If Trim(TxtTitular.Text) = "" Then
   DatosImpresionResumenOk = False
   Exit Function
End If

'valido el numero de credito
If Trim(TxtNumCredito.Text) = "" Then
   DatosImpresionResumenOk = False
   Exit Function
End If
If Not IsNumeric(TxtNumCredito.Text) Then
   DatosImpresionResumenOk = False
   Exit Function
End If
If CLng(TxtNumCredito.Text) <= 0 Then
   DatosImpresionResumenOk = False
   Exit Function
End If
If Not ExisteCredito(TxtNumCredito.Text) Then
   DatosImpresionResumenOk = False
   Exit Function
End If

'valido las cuotas
If lvcuotas.ListItems.Count() = 0 Then
   DatosImpresionResumenOk = False
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosImpresionResumenOk"
End Function
Private Function DatosClienteOk() As Boolean
On Error GoTo merror

DatosClienteOk = True

If Trim(TxtTitular.Text) = "" Then
   DatosClienteOk = False
   MsgE "Debe seleccionar el cliente titular del credito"
   Exit Function
End If

'verifico los requisitos generales de credito
If Not RequisitosGeneralesClienteOk(IdCliente) Then
   DatosClienteOk = False
   Exit Function
End If

'esto solo se valida para pasarle a la funcion de abajo

'valido el MONTO a financiar
If Trim(TxtFinanciar.Text) = "" Then
   DatosClienteOk = False
   MsgE "Debe ingresar el Monto a financiar"
   TxtFinanciar.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtFinanciar.Text) Then
   DatosClienteOk = False
   MsgE "El monto a financiar debe ser numerico"
   TxtFinanciar.SetFocus
   Exit Function
End If
If CCur(TxtFinanciar.Text) <= 0 Then
   DatosClienteOk = False
   MsgE "El monto a financiar debe ser mayor a cero"
   TxtFinanciar.SetFocus
   Exit Function
End If

'winikmora

Dias_Mora_Cliente = ObtenerDiasMoraMaximo(IdCliente, DTPicker1.Value)

If Dias_Mora_Cliente > VG_DIAS_MORA Then
    MsgE "Algún  crédito del cliente supera la máxima cantidad de días de mora admitida"
    DatosClienteOk = False
    Exit Function
End If


'valido el importe a financiar
If Trim(TxtImporteAFinanciar.Text) = "" Then
   DatosClienteOk = False
   MsgE "Debe ingresar el importe a financiar"
   TxtImporteAFinanciar.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtImporteAFinanciar.Text) Then
   DatosClienteOk = False
   MsgE "El importe a financiar debe ser numerico"
   TxtImporteAFinanciar.SetFocus
   Exit Function
End If
If CCur(TxtImporteAFinanciar.Text) <= 0 Then
   DatosClienteOk = False
   MsgE "El importe a financiar debe ser mayor a cero"
   TxtImporteAFinanciar.SetFocus
   Exit Function
End If

'valido el PTF
If Trim(TxtTotalPtf.Text) = "" Then
   DatosClienteOk = False
   MsgE "Falta el importe PTF"
   TxtImporteAFinanciar.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtTotalPtf.Text) Then
   DatosClienteOk = False
   MsgE "El importe PTF debe ser numerico"
   TxtImporteAFinanciar.SetFocus
   Exit Function
End If
If CCur(TxtTotalPtf.Text) <= 0 Then
   DatosClienteOk = False
   MsgE "El importe PTF debe ser mayor a cero"
   'TxtImporteAFinanciar.SetFocus
   TxtTotalPtf.SetFocus
   Exit Function
End If

'verifico requisitos particulares del cliente..si supera el creditomaximo
If Not RequisitosBasicosClienteOk(IdCliente, TxtTotalPtf.Text) Then
   DatosClienteOk = False
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosClienteOk-RegistrarCreditos"
End Function
Private Function DatosGaranteOk(ByVal IdCliente As Long) As Boolean
'chequea los datos del garante de un cliente
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

DatosGaranteOk = True

sql = "select idcliente,apellidogarante,nombregarante " & _
      "from clientes where idcliente=" & CLng(IdCliente)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   'si no hay datos de garante
   If Trim(rec.rdoColumns("apellidogarante")) = "" Then
      DatosGaranteOk = False
      MsgE "El cliente no tiene garante"
      Exit Function
   End If
   
   If Trim(rec.rdoColumns("nombregarante")) = "" Then
      DatosGaranteOk = False
      MsgE "Los datos del garante estan incompletos (falta el nombre)"
      Exit Function
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosGaranteOk-RegistrarCreditos"
End Function
Private Function datosok() As Boolean
'se ejecuta al grabar un credito
Dim CantDiasVenc As Integer
On Error GoTo merror

datosok = True

If Not DatosClienteOk() Then
   datosok = False
   Exit Function
End If

'si exijo garante lo valido
If VG_GARANTE Then
   If Not DatosGaranteOk(IdCliente) Then
      datosok = False
      Exit Function
   End If
End If

'valido el numero de credito (se genera solo)
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

'valido el plan
If ComboPlanes.Text = "" Then
   datosok = False
   MsgE "Debe seleccionar un plan"
   ComboPlanes.SetFocus
   Exit Function
End If

'valido el codigo de prestamo..solo valido que se haya ingresado
'porque puede ser alfanumerico
'If Trim(TxtCodPrestamo.Text) = "" Then
'   datosok = False
'   MsgE "Debe ingresar el codigo de prestamo"
'   TxtCodPrestamo.SetFocus
'   Exit Function
'End If

'valido la tasa de financiacion TNA (depende del plan)
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
'aca no comparo con el igual a cero porque debo poder reg cred con tasa cero
If CDbl(TxtTasaFinanciacion.Text) < 0 Then
   datosok = False
   MsgE "La tasa de financiacion debe ser mayor a cero"
   TxtTasaFinanciacion.SetFocus
   Exit Function
End If

'valido la cantidad de cuotas (depende del plan)
If Trim(TxtCantidadCuotas.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar la cantidad de cuotas del credito"
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

If ComboComercios.Text = "" Then
   datosok = False
   MsgE "Debe seleccionar un comercio para el credito"
   ComboComercios.SetFocus
   Exit Function
End If

'nuevo 2011 vendedor
If ComboVendedores.Text = "" Then
   datosok = False
   MsgE "Debe seleccionar un vendedor para el credito"
   ComboVendedores.SetFocus
   Exit Function
End If

'valido la provincia por el tema sellados
If comboprovincias.Text = "" Then
   datosok = False
   MsgE "Debe seleccionar una provincia"
   comboprovincias.SetFocus
   Exit Function
End If

'si no permito registrar creditos diferidos
If Not VG_CREDITOSDIFERIDOS Then
   If CDate(DTPicker1.Value) <> CDate(Date) Then
      datosok = False
      MsgE "Verifique la fecha del credito...debe ser igual a la actual"
      Exit Function
   End If
Else
   'si permito diferidos verifico que no sean menores al año 2000
   '***valido la fecha del credito
   If Year(DTPicker1.Value) < 2000 Then
      datosok = False
      MsgE "Verifique la fecha del credito...(el año debe ser superior a 2000)"
      DTPicker1.SetFocus
      Exit Function
   End If
   If CDate(DTPicker1.Value) > CDate(Date) Then
      datosok = False
      MsgE "Verifique la fecha del credito...(la fecha debe ser la actual)"
      DTPicker1.SetFocus
      Exit Function
   End If
End If

'si el primer vto es menor a la fecha del credito
If CDate(DTPicker2.Value) < CDate(DTPicker1.Value) Then
   datosok = False
   MsgE "El primer vencimiento debe ser mayor o igual a la fecha del credito"
   DTPicker2.SetFocus
   Exit Function
End If

If Month(CDate(DTPicker1.Value)) = Month(DTPicker2.Value) Then
   If Not MsgP("¿El primer vencimiento es en el mes de registracion del credito?") Then
      datosok = False
      Exit Function
   End If
End If

If CDate(DTPicker1.Value) < CDate(VG_FECHALIMITEINGRESO) Then
   datosok = False
   MsgE "La Fecha del credito es inferior a la fecha limite permitida"
   DTPicker1.SetFocus
   Exit Function
End If


'la fecha de 1er vencimiento no pueda superar la fecha actual + cant dias (ingersado en config sistema)

If VF_CANT_DIAS > 0 Then
    CantDiasVenc = VF_CANT_DIAS
Else
    CantDiasVenc = VG_CANT_DIAS
End If
If CDate(DTPicker3.Value) > (DTPicker1.Value + CantDiasVenc) Then
   datosok = False
   MsgE "La fecha del Segundo Vencimiento supera los limites permitidos"
   DTPicker3.SetFocus
   Exit Function
End If


'chequeo la primera fecha de vencimiento
If EsFeriado(DTPicker2.Value) Then
   datosok = False
   MsgE "El primer vencimiento es feriado"
   DTPicker2.SetFocus
   Exit Function
End If

'chequeo la primera fecha de vencimiento
If EsSabado(DTPicker2.Value) Then
   datosok = False
   MsgE "El primer vencimiento es sabado"
   DTPicker2.SetFocus
   Exit Function
End If

'chequeo la primera fecha de vencimiento
If EsDomingo(DTPicker2.Value) Then
   datosok = False
   MsgE "El primer vencimiento es domingo"
   DTPicker2.SetFocus
   Exit Function
End If

'si uso segundo vencimiento de cuotas
If VG_APLICARSEGUNDOVENCIMIENTO Then
   If CDate(DTPicker3.Value) <= CDate(DTPicker2.Value) Then
      datosok = False
      MsgE "El segundo vencimiento debe ser mayor que el primer vencimiento"
      DTPicker3.SetFocus
      Exit Function
   End If
   'chequeo la segunda fecha de vencimiento
   If EsFeriado(DTPicker3.Value) Then
      datosok = False
      MsgE "El segundo vencimiento es feriado"
      DTPicker3.SetFocus
      Exit Function
   End If
   If EsSabado(DTPicker3.Value) Then
      datosok = False
      MsgE "El segundo vencimiento es sabado"
      DTPicker3.SetFocus
      Exit Function
   End If
   If EsDomingo(DTPicker3.Value) Then
      datosok = False
      MsgE "El segundo vencimiento es domingo"
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

'chequeo la fecha de entrega del prestamo
'si la fecha de entrega es mayor al primer vto
If CDate(DTPicker4.Value) > CDate(DTPicker2.Value) Then
   datosok = False
   MsgE "La fecha de entrega del prestamo debe ser menor que el 1º vencimiento"
   Exit Function
End If
   
If ComboModalidad.Text = "Cheque" Then
   If Trim(TxtNumCheque.Text) = "" Then
      datosok = False
      MsgE "Debe ingresar el Nº de cheque entregado al cliente"
      TxtNumCheque.SetFocus
      Exit Function
   End If
   If Trim(ComboBancos.Text) = "" Then
      datosok = False
      MsgE "Debe seleccionar el banco del cheque entregado al cliente"
      ComboBancos.SetFocus
      Exit Function
   End If
End If

Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-RegistrarCreditos"
End Function
Private Sub cmdgenerar_click()
'grabo el nuevo credito ya calculado en la funcion calcularimportes()
Dim sql As String
Dim IdCredito As Long
Dim I As Long
Dim NumFactura As Long
Dim NumFactura2 As Long
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
Dim ImporteRecargoVencimiento2 As Currency
Dim Periodo As String
Dim CodigoBarras As String
Dim Banco As String
Dim NumCheque As String
Dim IdPlan As Long
Dim IvaInteres As Currency
Dim IvaSeguros As Currency
Dim IvaOtorgamientoGastos As Currency
Dim Otorgamiento As Currency
Dim Formula As String
Dim IdProvincia As Long
Dim CodPrestamo As String
Dim DiasRecargo As Long
Dim TasaAInsertar As Double
On Error GoTo merror
Call RefreshTimer
Call ComboPlanes_Click
'verifico requisitos de titular y garante,ademas verifico los campos
If Not datosok() Then Exit Sub

Formula = "AMORTIZACION SISTEMA FRANCES"

'valido el cliente
If Not ExisteCliente(IdCliente) Then
   MsgE "El cliente seleccionado no existe"
   Exit Sub
End If

'lo vuelvo a generar porque no es mas manual
TxtNumCredito.Text = Format(UltimoId("idcredito", "creditos") + 1, "000000")
IdCredito = CLng(TxtNumCredito.Text)

'verifico si el credito ya existe
If ExisteCredito(IdCredito) Then
   MsgE "El numero de credito ya existe...(reintente la operacion)"
   TxtNumCredito.SetFocus
   Exit Sub
End If

'valido el plan
IdPlan = ComboPlanes.ItemData(ComboPlanes.ListIndex)
If Not ExistePlan(IdPlan) Then
   MsgE "El plan seleccionado no existe"
   Exit Sub
End If

'verifico el codigo de prestamo
If ChkDDJJ.Value = 1 Then
    TxtCodPrestamo.Text = NuevoCodPrestamo(IdCliente, "C")
    If ExisteCodPrestamo(TxtCodPrestamo.Text) Then
       MsgE "El codigo de prestamo ya existe...ingrese otro diferente"
       TxtCodPrestamo.SetFocus
       Exit Sub
    End If
Else

    TxtCodPrestamo.Text = NuevoCodPrestamo(IdCliente, "R")
    If ExisteCodPrestamo(TxtCodPrestamo.Text) Then
       MsgE "El codigo de prestamo ya existe...ingrese otro diferente"
       TxtCodPrestamo.SetFocus
       Exit Sub
    End If
End If

Banco = ""
NumCheque = ""

If ComboModalidad.Text = "CHEQUE" Then
   NumCheque = UCase(Trim(TxtNumCheque.Text))
   Banco = ComboBancos.Text
End If

'esto no anda si las cuotas son distintas
'obtengo el importe de cada cuota del credito..saco de la 2 porque la uno
'tal vez tenga sellados
If CLng(TxtCantidadCuotas.Text) > 1 Then
   ImporteCuota = CCur(lvcuotas.ListItems.Item(2).SubItems(10))
Else
   ImporteCuota = CCur(lvcuotas.ListItems.Item(1).SubItems(10))
End If

'Verifico q la fecha de alta del credito sea mayor o igual a la cargada en conf del sistema



If Not MsgP("¿Confirma el nuevo credito?") Then Exit Sub

'otras validaciones
'valido el cliente
If Not ExisteCliente(IdCliente) Then
   MsgE "El cliente no existe"
   Exit Sub
End If

'valido el credito
If ExisteCredito(IdCredito) Then
   MsgE "El numero de credito ya existe...(reintente la operacion)"
   TxtNumCredito.SetFocus
   Exit Sub
End If

'valido el plan
If Not ExistePlan(IdPlan) Then
   MsgE "El plan no existe"
   Exit Sub
End If

'valido el codigo de prestamo
If ExisteCodPrestamo(TxtCodPrestamo.Text) Then
   MsgE "El codigo de prestamo ya existe..ingrese otro diferente"
   TxtCodPrestamo.SetFocus
   Exit Sub
End If

IdProvincia = CLng(comboprovincias.ItemData(comboprovincias.ListIndex))


TasaAInsertar = TxtTasaFinanciacion.Text
'If ChkDDJJ.Value Then
'      TasaAInsertar = Rate(TxtCantidadCuotas.Text, CDbl(lvcuotas.ListItems.Item(1).SubItems(1)) + CDbl(lvcuotas.ListItems.Item(1).SubItems(2)), -(TxtImporteAFinanciar.Text)) * 100 * 12
'End If

'If ChEntrega = True Then
'    STREntrega = "TRUE"
'Else
    
'End If

'inicio transaccion
cnSQL.BeginTrans

',MontoOriginal,ProductoEntrega,TasaUtEntrega
'agrego un nuevo credito
sql = "insert into creditos " & _
      "(idcredito,idplan,idcliente," & _
      "tasa,importeafinanciar,importeinteres,importefinanciado,ptf," & _
      "importegastos,importeseguros,importeotorgamiento,importesellados,ivainteres,ivaseguros,ivaotgastos," & _
      "numcuotas,fechacredito,observaciones," & _
      "sonpesos,modalidadentrega,numcheque," & _
      "banco,formula,fechadesembolso,importetotal,fechavencimiento1,importecuota,codprestamo,idprovincia,motivobloqueo,cad1,cad2,MontoOriginal,ProductoEntrega,TasaUtEntrega) " & _
      "values(" & CLng(IdCredito) & "," & CLng(IdPlan) & "," & CLng(IdCliente) & _
      "," & ConvertirDblSql(TasaAInsertar) & _
      "," & ConvertirDblSql(CCur(TxtImporteAFinanciar.Text)) & _
      "," & ConvertirDblSql(CCur(TxtTotalInteres.Text)) & "," & ConvertirDblSql(CCur(TxtImporteFinanciado.Text)) & _
      "," & ConvertirDblSql(CCur(TxtTotalPtf.Text)) & "," & ConvertirDblSql(CCur(TxtTotalGastos.Text)) & _
      "," & ConvertirDblSql(CCur(TxtTotalSeguros.Text)) & _
      "," & ConvertirDblSql(CCur(TxtTotalOtorgamiento.Text)) & "," & ConvertirDblSql(CCur(TxtTotalSellados.Text)) & "," & ConvertirDblSql(CCur(TxtTotalIvaInteres.Text)) & _
      "," & ConvertirDblSql(CCur(TxtTotalIvaSeguros.Text)) & _
      "," & ConvertirDblSql(CCur(TxtTotalIvaGastos.Text)) & _
      "," & CLng(TxtCantidadCuotas.Text) & _
      ",'" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & _
      "','" & CStr(txtObservaciones.Text) & _
      "','" & CStr(TxtSonPesos.Text) & _
      "','" & CStr(ComboModalidad.Text) & _
      "','" & CStr(NumCheque) & "','" & CStr(Banco) & _
      "','" & CStr(Formula) & _
      "','" & ConvertirFechaSql(CDate(DTPicker4.Value), "DD/MM/YYYY") & "'," & ConvertirDblSql(CCur(TxtImporteTotal.Text)) & _
      ",'" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "'," & ConvertirDblSql(CCur(ImporteCuota)) & ",'" & CStr(TxtCodPrestamo.Text) & "'," & CLng(IdProvincia) & ",'" & CStr(ComboComercios.Text) & "','" & CStr(VG_USUARIOLOGIN) & "','" & CStr(ComboVendedores.Text) & "'," & ConvertirDblSql(CCur(TxtFinanciar.Text)) & ",'" & CStr(ChEntrega) & "'," & CLng(ValorPorcentaje) & " )"
      
cnSQL.Execute sql

'recorro la lista grabando las cuotas ya calculadas
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
    'la columna 4 muestra el subtotal
    ImporteSeguros = CCur(lvcuotas.ListItems.Item(I).SubItems(5))
    'iva sobre seguros
    IvaSeguros = CCur(lvcuotas.ListItems.Item(I).SubItems(6))
    Otorgamiento = CCur(lvcuotas.ListItems.Item(I).SubItems(7))
    ImporteGastos = CCur(lvcuotas.ListItems.Item(I).SubItems(8))
    'iva sobre otorgamiento y gastos
    IvaOtorgamientoGastos = CCur(lvcuotas.ListItems.Item(I).SubItems(9))
    
    'este es el vto1 es el total de la cuota completa
    ImporteVencimiento1 = CCur(lvcuotas.ListItems.Item(I).SubItems(10))
    Vencimiento1 = CDate(lvcuotas.ListItems.Item(I).SubItems(11))
    
    ImporteVencimiento2 = CCur(lvcuotas.ListItems.Item(I).SubItems(12))
    Vencimiento2 = CDate(lvcuotas.ListItems.Item(I).SubItems(13))
    
    'este es el importe cuota de antes original capital + interes
    ImporteCuota = CCur(ImporteAmortizacion) + CCur(ImporteInteres)
    
    'este es el total de impuestos (reuno los impuestos)
    ImporteImpuestos = CCur(IvaInteres) + CCur(IvaSeguros) + CCur(IvaOtorgamientoGastos)
        
    'grabo el periodo para luego agrupar un reporte por meses
    Periodo = Format(CStr(Year(Vencimiento1)), "0000") & Format(CStr(Month(Vencimiento1)), "00")
    
    'ahora en credimaco no hay 2 vto
    ImporteRecargoVencimiento2 = CCur(ImporteVencimiento2) - CCur(ImporteVencimiento1)
           
    DiasRecargo = CDate(Vencimiento2) - CDate(Vencimiento1)
    
    'genero codigo de barras
    'ahora le paso los nuevos datos de rapipago
    CodigoBarras = GenerarCodigoBarras(VG_NUMEMPRESA, IdCliente, NumFactura, ImporteVencimiento1, Vencimiento1, ImporteRecargoVencimiento2, DiasRecargo)
    
    'grabo la cuota
    sql = "insert into cuotas (numfactura,idcredito,numcuota,importecuota," & _
          "fechavencimiento1,fechavencimiento2,importegastos,importeseguros," & _
          "importeimpuestos,importerecargovencimiento2,codigobarras," & _
          "importeamortizacion,importeinteres,periodo,importevencimiento1,importevencimiento2,otorgamiento,ivainteres,ivaseguros,ivaotorgamientogastos) " & _
          "values(" & CLng(NumFactura) & "," & CLng(IdCredito) & _
          "," & CLng(NumCuota) & "," & ConvertirDblSql(CCur(ImporteCuota)) & _
          ",'" & ConvertirFechaSql(CDate(Vencimiento1), "DD/MM/YYYY") & "','" & ConvertirFechaSql(CDate(Vencimiento2), "DD/MM/YYYY") & _
          "'," & ConvertirDblSql(CCur(ImporteGastos)) & "," & ConvertirDblSql(CCur(ImporteSeguros)) & _
          "," & ConvertirDblSql(CCur(ImporteImpuestos)) & "," & ConvertirDblSql(CCur(ImporteRecargoVencimiento2)) & ",'" & CStr(CodigoBarras) & _
          "'," & ConvertirDblSql(CCur(ImporteAmortizacion)) & "," & ConvertirDblSql(CCur(ImporteInteres)) & _
          ",'" & CStr(Periodo) & "'," & ConvertirDblSql(CCur(ImporteVencimiento1)) & _
          "," & ConvertirDblSql(CCur(ImporteVencimiento2)) & "," & ConvertirDblSql(CCur(Otorgamiento)) & "," & ConvertirDblSql(CCur(IvaInteres)) & _
          "," & ConvertirDblSql(CCur(IvaSeguros)) & "," & ConvertirDblSql(CCur(IvaOtorgamientoGastos)) & ")"
    
    cnSQL.Execute sql
Next I

'grabo la ultima tasa del sistema utilizada..
sql = "update configuracionsistema " & _
      "set tasafinanciacion=" & ConvertirDblSql(CDbl(TxtTasaFinanciacion.Text)) & ", ultimocupon = " & CLng(NumFactura)
cnSQL.Execute sql

'fin de la transaccion
cnSQL.CommitTrans

MsgI "El credito fue registrado"

'deshabilito la generacion hasta que pulsen el boton [Nuevo]
CmdGenerar.Enabled = False

'habilito la impresion del resumen del credito
CmdImprimirResumen.Enabled = True

Exit Sub
merror:
tratarerrores "Error grabando el nuevo credito"
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

'valido la tasa TNA
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

'valido la tasa (TEM)solo en credimaco
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
tratarerrores "Error en funcion DatosNumericosOk-RegistrarCreditos"
End Function
Private Sub CalcularImportes()
'calcula las cuotas cada vez que se actualizan los campos
Dim CapitalCuota As Currency
Dim InteresCuota As Currency
Dim ImporteCuota As Currency
Dim ImporteGastos As Currency
Dim ImporteSeguros As Currency
Dim ImporteImpuestos As Currency
Dim ImporteSellados As Currency
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
Dim TotalInteres As Currency
Dim NuevoPorcentaje As Double
Dim CalcularImpuestos As Boolean
Dim CapitalTotal As Currency
Dim InteresTotal As Currency
Dim CapitalMasInteresTotal As Currency
On Error GoTo merror

lvcuotas.ListItems.Clear

TxtTotalInteres.Text = 0
TxtTotalPtf.Text = 0
TxtTotalGastos.Text = 0
TxtTotalSeguros.Text = 0
TxtTotalIvaInteres.Text = 0
TxtTotalIvaSeguros.Text = 0
TxtTotalIvaGastos.Text = 0
TxtTotalOtorgamiento.Text = 0
TxtImporteFinanciado.Text = 0
TxtImporteTotal.Text = 0

ImporteCuota = 0
ImporteGastos = 0
ImporteSeguros = 0
ImporteImpuestos = 0
ImporteSellados = 0
ImporteVencimiento1 = 0
ImporteVencimiento2 = 0
importevto1 = 0
ImporteVto2 = 0
AcumGastos = 0
AcumSeguros = 0
AcumImpuestos = 0
Resto = 0

If Not DatosNumericosOk() Then Exit Sub

TasaFinanciacion = CDbl(TxtTasaFinanciacion.Text)

TasaTem = CDbl(TxtTasa2.Text)

'obtengo el primer vencimiento
Vencimiento1 = CDate(DTPicker2.Value)
Vencimiento2 = Vencimiento1

DiferenciaVencimientos = 0

'trata que venza siempre el 20 por ejemplo
DiaVencimiento = Day(DTPicker2.Value)

'en credimaco no hay 2 vto
If VG_APLICARSEGUNDOVENCIMIENTO Then
   Vencimiento2 = CDate(DTPicker3.Value)
   DiferenciaVencimientos = DateDiff("d", Vencimiento1, Vencimiento2)
End If

'esto devuelve el importe total de la cuota incluyendo capital e interes
ImporteCuota = Pmt(TasaTem / 100, TxtCantidadCuotas.Text, -(TxtImporteAFinanciar.Text))
TxtImporteFinanciado.Text = CDbl(ImporteCuota) * CLng(TxtCantidadCuotas.Text)
CapitalCuota = CCur(TxtImporteAFinanciar.Text) / CLng(TxtCantidadCuotas.Text)
InteresCuota = CCur(ImporteCuota) - CCur(CapitalCuota)

' If ChkDDJJ.Value Then
'    InteresCuota = InteresCuota * (1 + CDbl(VG_PORCENTAJEIVA) / 100)
'End If

CapitalTotal = CCur(TxtImporteAFinanciar.Text)
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
       'si va a 1
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
          'aplico a todas o de la 2 a la ultima
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
          End If
       End If
    End If
    
    Nitem.SubItems(8) = Format(ImporteGastos, "0.00") & vbNullString
             
    'si aplico seguros
    ImporteSeguros = 0
    If VG_APLICARSEGURO Then
       'si aplico solo a la cuota 1
       If VG_APLICARSEGUROSCUOTA1 Then
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
          End If
       Else
          'aplico a todas
          If CCur(VG_ALICUOTASEGUROS) > 0 Then
             'seguros
             ImporteSeguros = (VG_ALICUOTASEGUROS * CCur(Resto) / 100) / CDbl(1.21)
             'disminuyo el resto (parecido a capital vivo)
             Resto = CCur(Resto) - CCur(ImporteCuota)
             'el resto sirve para seguir calculando el seguro a las demas cuotas
          End If
          If CCur(VG_SEGUROFIJO) > 0 Then
             'esto es para seguros simultaneos
             ImporteSeguros = CCur(ImporteSeguros) + CCur(VG_SEGUROFIJO)
          End If
          If CCur(VG_IMPORTESEGURO) > 0 Then
             'esto es para seguros simultaneos
             ImporteSeguros = CCur(ImporteSeguros) + (CCur(VG_IMPORTESEGURO) / CLng(TxtCantidadCuotas.Text))
          End If
       End If
    End If
               
    Nitem.SubItems(1) = Format(CapitalCuota, "0.00") & vbNullString
    Nitem.SubItems(2) = Format(InteresCuota, "0.00") & vbNullString
    Nitem.SubItems(5) = Format(ImporteSeguros, "0.00") & vbNullString
    
    'si aplico otorgamiento
    ImporteOtorgamiento = 0
    If VG_APLICAROTORGAMIENTO Then
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
    Nitem.SubItems(7) = Format(ImporteOtorgamiento, "0.00") & vbNullString
      
    CalcularImpuestos = False
    'si aplico impuestos calculo el iva de otorgamiento
    If VG_APLICARIMPUESTOS Then
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
    End If 'si aplico impuestos
    
    If CalcularImpuestos Then
       'si calculo segun credimaco
       If VG_IMPUESTOSCREDIMACO Then
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
    
    'primer subtotal de los tres anteriores
    Subtotal1 = CCur(CapitalCuota) + CCur(InteresCuota) + CCur(IvaInteres)
    Nitem.SubItems(4) = Format(Subtotal1, "0.00") & vbNullString
    
    'iva seguros
    Nitem.SubItems(6) = Format(IvaSeguros, "0.00") & vbNullString
    'aca va el iva otorgamiento y el iva gastos en el mismo
    Nitem.SubItems(9) = Format(IvaOtorgamiento + IvaGastos, "0.00") & vbNullString
    
    ImporteVencimiento1 = CCur(Subtotal1) + CCur(ImporteSeguros) + CCur(IvaSeguros) + CCur(ImporteOtorgamiento) + CCur(IvaOtorgamiento) + CCur(ImporteGastos) + CCur(IvaGastos)
    Nitem.SubItems(10) = Format(ImporteVencimiento1, "0.00") & vbNullString
    
    'primer vencimiento
    Nitem.SubItems(11) = Vencimiento1
    
    'recalculo el importe del vencimiento2
    ImporteVencimiento2 = CalcularImporteVencimiento2(ImporteVencimiento1, Vencimiento1, Vencimiento2)
    
    Nitem.SubItems(12) = Format(ImporteVencimiento2, "0.00")
    
    'segundo vencimiento
    Nitem.SubItems(13) = Vencimiento2
          
    'hay que mover los vencimientos
    'si aplico vencimiento por meses
    If VG_DIASVENCIMIENTOFINANCIACION = 30 Or VG_DIASVENCIMIENTOFINANCIACION = 60 Then
       Frecuencia = CLng(VG_DIASVENCIMIENTOFINANCIACION / 30)
       'incremento al primer vencimiento por meses exactos validos
       FechaProxima = ArmarFecha(DiaVencimiento, Vencimiento1, Frecuencia)
    Else
       'aplico vencimiento por dias
       FechaProxima = Vencimiento1 + VG_DIASVENCIMIENTOFINANCIACION
    End If
    
    'muevo el vencimiento1
    Vencimiento1 = ObtenerFechaVencimiento(FechaProxima, VG_DIASVENCIMIENTOFINANCIACION)
    Vencimiento2 = Vencimiento1
       
    If VG_APLICARSEGUNDOVENCIMIENTO Then
       'obtengo el segundo vencimiento respetando la diferencia inicial
       Vencimiento2 = ObtenerFechaVencimiento(Vencimiento1 + DiferenciaVencimientos, VG_DIASVENCIMIENTOFINANCIACION)
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
 'TxtTotalPTF.Text = CCur(TxtImporteAFinanciar.Text) + CCur(TxtTotalInteres.Text) + CCur(TxtTotalSeguros.Text) + CCur(TxtTotalOtorgamiento.Text) + CCur(TxtTotalGastos.Text)
 'Pedido el 22/10 todavia no aprobado presupuesto MW
TxtTotalPtf.Text = CCur(TxtImporteAFinanciar.Text) + CCur(TxtTotalInteres.Text) + CCur(TxtTotalGastos.Text)

TxtTotalPtf.Text = Format(TxtTotalPtf.Text, "0.00")

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

'muestro el importe total del credito en letras
TxtSonPesos.Text = SonPesos(TxtImporteAFinanciar.Text)

Exit Sub
merror:
tratarerrores "Error en procedimiento CalcularImportes-RegistrarCreditos"
End Sub
Private Sub ComboModalidad_Click()
If ComboModalidad.Text = "CHEQUE" Then
   LabelBanco.Visible = True
   LabelNumCheque.Visible = True
   ComboBancos.Visible = True
   TxtNumCheque.Visible = True
Else
   LabelBanco.Visible = False
   LabelNumCheque.Visible = False
   ComboBancos.Visible = False
   TxtNumCheque.Visible = False
End If
End Sub

Private Sub TxtCodPrestamo_LostFocus()
TxtCodPrestamo.Text = UCase(Trim(TxtCodPrestamo.Text))
End Sub
Private Sub TxtComercio_LostFocus()
TxtComercio.Text = UCase(Trim(TxtComercio.Text))
End Sub

Private Sub TxtFinanciar_LostFocus()

Dim IdPlan As Long
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

If ChEntrega = False Then
    TxtImporteAFinanciar = TxtFinanciar
    ValorPorcentaje = 0
Else
    'al seleccionar un plan actualiza las tasas, etc
    If ComboPlanes.Text = "" Then Exit Sub
    
    IdPlan = ComboPlanes.ItemData(ComboPlanes.ListIndex)
    
    sql = "select * " & _
          "from planes " & _
          "where idplan=" & CLng(IdPlan)
    Set rec = cnSQL.OpenResultset(sql)
    
         If Not rec.EOF Then
           If Not IsNull(rec.rdoColumns("porcprodentre")) Then
            ValorPorcentaje = rec.rdoColumns("porcprodentre")
            TxtImporteAFinanciar = TxtFinanciar * (100 - Int(rec.rdoColumns("porcprodentre"))) / 100
            End If
         End If
   'TxtImporteAFinanciar =
End If
Exit Sub
merror:
tratarerrores "Error seleccionando Monto a Financiar"
End Sub

Private Sub TxtImporteAFinanciar_Change()

'si cambio el importe a financiar
Call CalcularImportes

'calculo los sellados
Call CalcularSellados

End Sub
Private Sub TxtCantidadCuotas_Change()
'si cambio la cantidad de cuotas
Call CalcularImportes
End Sub
Private Sub DTPicker2_Change()
'si cambio la fecha de vencimiento1
DTPicker3.Value = DTPicker2.Value
Call CalcularImportes
End Sub
Private Sub DTPicker3_Change()
Call CalcularImportes
End Sub
Private Sub TxtObservaciones_LostFocus()
txtObservaciones.Text = UCase(Trim(txtObservaciones.Text))
End Sub
Private Sub TxtTasaFinanciacion_Change()
'si cambia la tasa
Call CalcularTem
Call CalcularImportes
End Sub
Private Sub Redondear()
'redondea los importes
Dim I As Long
Dim Capital As Currency
Dim Interes As Long
Dim Gastos As Currency
Dim Seguros As Currency
Dim Impuestos As Currency
Dim Sellados As Currency
Dim Vencimiento1 As Date
Dim Vencimiento2 As Date
Dim ImporteCuota As Currency
Dim ImporteVencimiento1 As Currency
Dim ImporteVencimiento2 As Currency
Dim ImporteTotalGastos As Currency
Dim ImporteTotalSeguros As Currency
Dim ImporteTotalOtorgamiento As Currency
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
Dim ImporteTotalInteres As Currency
Dim ImportePTF As Currency
On Error GoTo merror

ImporteFinanciado = 0
ImporteTotalGastos = 0
ImporteTotalSeguros = 0
ImporteTotalImpuestos = 0
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
        
    'importecuota.las partes ya estan redondeadas
    ImporteCuota = CCur(Capital) + CCur(Interes)
       
    'seguros
    Seguros = CCur(lvcuotas.ListItems.Item(I).SubItems(5))
    Seguros = Round(Seguros)
    lvcuotas.ListItems.Item(I).SubItems(5) = Format(Seguros, "0.00")
    
    'ivaseguros
    IvaSeguros = CCur(lvcuotas.ListItems.Item(I).SubItems(6))
    IvaSeguros = Round(IvaSeguros)
    lvcuotas.ListItems.Item(I).SubItems(6) = Format(IvaSeguros, "0.00")
       
    'otortgamiento credimaco
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
        
    ImporteVencimiento1 = CCur(Subtotal) + CCur(Seguros) + CCur(IvaSeguros) + CCur(Otorgamiento) + CCur(Gastos) + CCur(IvaGastos)
    lvcuotas.ListItems.Item(I).SubItems(10) = Format(ImporteVencimiento1, "0.00")
       
    'importe vencimiento2
    Vencimiento1 = CDate(lvcuotas.ListItems.Item(I).SubItems(11))
    Vencimiento2 = CDate(lvcuotas.ListItems.Item(I).SubItems(13))
    
    'lo calculo en base al importe 1 nuevo redondeado
    ImporteVencimiento2 = CalcularImporteVencimiento2(ImporteVencimiento1, Vencimiento1, Vencimiento2)
    ImporteVencimiento2 = Round(ImporteVencimiento2)
    lvcuotas.ListItems.Item(I).SubItems(12) = Format(ImporteVencimiento2, "0.00")
    
    
    'ImportePTF = CCur(ImporteFinanciado) + CCur(ImporteTotalInteres) + CCur(ImporteTotalGastos)
    'ImportePTF = Round(ImportePTF)
         
    'totales
    ImporteFinanciado = CCur(ImporteFinanciado) + CCur(ImporteCuota)
    ImporteTotalGastos = CCur(ImporteTotalGastos) + CCur(Gastos)
    ImporteTotalSeguros = CCur(ImporteTotalSeguros) + CCur(Seguros)
    
    ImporteTotalIvaInteres = CCur(ImporteTotalIvaInteres) + CCur(IvaInteres)
    ImporteTotalIvaSeguros = CCur(ImporteTotalIvaSeguros) + CCur(IvaSeguros)
    ImporteTotalIvaGastos = CCur(ImporteTotalIvaGastos) + CCur(IvaGastos)
    
    ImporteTotalOtorgamiento = CCur(ImporteTotalOtorgamiento) + CCur(Otorgamiento)
    ImporteTotal = CCur(ImporteTotal) + CCur(ImporteVencimiento1)
    ImporteTotalInteres = CCur(ImporteTotalInteres) + CCur(Interes)
    
    ImportePTF = CCur(ImportePTF) + CCur(Capital)
    
Next I

ImportePTF = CCur(ImportePTF) + CCur(ImporteTotalInteres) + CCur(ImporteTotalGastos)

TxtImporteFinanciado.Text = Format(ImporteFinanciado, "0.00")
TxtTotalGastos.Text = Format(ImporteTotalGastos, "0.00")
TxtTotalSeguros.Text = Format(ImporteTotalSeguros, "0.00")
TxtTotalIvaInteres.Text = Format(ImporteTotalIvaInteres, "0.00")
TxtTotalIvaSeguros.Text = Format(ImporteTotalIvaSeguros, "0.00")
TxtTotalIvaGastos.Text = Format(ImporteTotalIvaGastos, "0.00")
TxtTotalSellados.Text = Format(ImporteTotalOtorgamiento, "0.00")
TxtImporteTotal.Text = Format(ImporteTotal, "0.00")
TxtTotalInteres.Text = Format(ImporteTotalInteres, "0.00")
TxtTotalOtorgamiento.Text = Format(ImporteTotalOtorgamiento, "0.00")
TxtTotalPtf.Text = Format(ImportePTF, "0.00")

Exit Sub
merror:
tratarerrores "Error en funcion Redondear"
End Sub
Private Sub ComboPlanes_Click()
'al seleccionar un plan actualiza las tasas, etc
Dim IdPlan As Long
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

'por defecto el vencimiento es mensual
VG_DIASVENCIMIENTOFINANCIACION = 30

If ComboPlanes.Text = "" Then Exit Sub

IdPlan = ComboPlanes.ItemData(ComboPlanes.ListIndex)

sql = "select * " & _
      "from planes " & _
      "where idplan=" & CLng(IdPlan)
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idplan")) Then
      TxtTasaFinanciacion.Text = CDbl(rec.rdoColumns("tasa1"))
      TxtTasa2.Text = CDbl(rec.rdoColumns("tasa2"))
      TxtCantidadCuotas.Text = CLng(rec.rdoColumns("cantcuotas"))
      VG_DIASVENCIMIENTOFINANCIACION = CLng(rec.rdoColumns("diasvencimiento"))
      
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
    
      'Gastos de otorgamiento
      VG_APLICAROTORGAMIENTO = rec.rdoColumns("aplicarotorgamiento")
      VG_APLICAROTORGAMIENTOCUOTA1 = rec.rdoColumns("aplicarotorgamientocuota1")
      VG_IMPORTEOTORGAMIENTO = CCur(rec.rdoColumns("importeotorgamiento"))
      VG_OTORCAPNOINT = rec.rdoColumns("OtorCapNoInt")
      VG_OTORINTNOCAP = rec.rdoColumns("OtorIntNoCap")
      VG_OTORCAPMASINT = rec.rdoColumns("OtorCapmasInt")
      VG_NOAPLICAROTREFIN = rec.rdoColumns("noaplicarotorrefin")
      
      VF_CANT_DIAS = 0
      If Not IsNull(rec.rdoColumns("cantdiasvenc")) Then
        If rec.rdoColumns("cantdiasvenc") <> 0 Then
            VF_CANT_DIAS = rec.rdoColumns("cantdiasvenc")
        End If
      End If
      
      'esto lo cambie para que solo actualice si esta todo ok
      Call CalcularImportes
      Call TxtFinanciar_LostFocus
   End If
Else
   MsgI "El plan no existe"
End If

Exit Sub
merror:
tratarerrores "Error seleccionando plan"
End Sub
Private Sub CalcularTem()
'calcula la tasa TEM y la aplica cuando cambia la tasa1 tna y cuando cambia
'el tipo de vencimiento mensual, etc
Dim Dias As Long
Dim Tem As Double
On Error GoTo merror

If Trim(TxtTasaFinanciacion.Text) = "" Then Exit Sub
If Not IsNumeric(TxtTasaFinanciacion.Text) Then Exit Sub
If CDbl(TxtTasaFinanciacion.Text) < 0 Then Exit Sub

TxtTasa2.Text = 0

'si es mensual
If VG_DIASVENCIMIENTOFINANCIACION = 30 Then
   TxtTasa2.Text = CDbl(TxtTasaFinanciacion.Text) / 365 * 30
End If
'diario
If VG_DIASVENCIMIENTOFINANCIACION = 1 Then
   TxtTasa2.Text = CDbl(TxtTasaFinanciacion.Text) / 365 * 1
End If
'semanal
If VG_DIASVENCIMIENTOFINANCIACION = 7 Then
   TxtTasa2.Text = CDbl(TxtTasaFinanciacion.Text) / 365 * 7
End If
'quincenal
If VG_DIASVENCIMIENTOFINANCIACION = 15 Then
   TxtTasa2.Text = CDbl(TxtTasaFinanciacion.Text) / 365 * 15
End If
'bimestral
If VG_DIASVENCIMIENTOFINANCIACION = 60 Then
   TxtTasa2.Text = CDbl(TxtTasaFinanciacion.Text) / 365 * 60
End If

TxtTasa2.Text = Format(TxtTasa2.Text, "0.00")

Exit Sub
merror:
tratarerrores "Error calculando Tem"
End Sub
Private Sub ComboProvincias_Click()
Call CalcularSellados
'si viene de comercios no haga nada
If BandComercio = False Then
    Call CargarComboWhere("comercios", ComboComercios, comboprovincias)
End If
End Sub
Private Sub CalcularSellados()
'calcula el sellado segun la provincia seleccionada
Dim IdProvincia As Long
Dim Porcentaje As Double
Dim Resultado As Currency
On Error GoTo merror

If comboprovincias.Text = "" Then Exit Sub

If Trim(TxtImporteAFinanciar.Text) = "" Then Exit Sub
If Not IsNumeric(TxtImporteAFinanciar.Text) Then Exit Sub

IdProvincia = CLng(comboprovincias.ItemData(comboprovincias.ListIndex))

Porcentaje = ObtenerPorcentajeSellados(IdProvincia)
Resultado = ObtenerImporteSellados(Porcentaje, TxtImporteAFinanciar.Text)

TxtTotalSellados.Text = Format(Resultado, "0.00")

Exit Sub
merror:
tratarerrores "Error en funcion CalcularSellados"
End Sub

