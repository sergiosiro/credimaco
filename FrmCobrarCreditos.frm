VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCobrarCreditos 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobrar Cuotas de Creditos"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   HelpContextID   =   18
   Icon            =   "FrmCobrarCreditos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtIvaInteresRestante 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   92
      Tag             =   "N"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Cierra la pantalla"
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   4200
      TabIndex        =   35
      ToolTipText     =   "Limpia la pantalla para ingresar nuevos datos"
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton CmdAnularCobro 
      Caption         =   "&Anular cobro"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   34
      ToolTipText     =   "Anula el cobro de la factura ingresada"
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Registrar Cobro"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Registra el cobro de la factura seleccionada"
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   630
      Width           =   7935
      Begin VB.TextBox TxtCodPrestamo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   63
         ToolTipText     =   "Codigo de prestamo de la cuota ingresada"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtSaldo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   54
         Tag             =   "N"
         ToolTipText     =   "Saldo del credito seleccionado"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtTotalCuotas 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   4
         Tag             =   "N"
         ToolTipText     =   "Nº de cuotas del credito de la factura ingresada"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TxtNumCredito 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "N"
         ToolTipText     =   "Nº de credito de la factura ingresada"
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TxtCliente 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Cliente titular de la factura"
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label23 
         Caption         =   "Prestamo Nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Saldo Credito $:"
         Height          =   255
         Left            =   5160
         TabIndex        =   53
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Cuotas:"
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Credito Nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del cobro:"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   20
      Top             =   6120
      Width           =   7935
      Begin VB.TextBox TxtCotizacion 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   61
         Tag             =   "no"
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TxtNumRecibo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   59
         Tag             =   "N"
         ToolTipText     =   "Numero de recibo que se asignara al cobro de la cuota"
         Top             =   1080
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1440
         TabIndex        =   57
         ToolTipText     =   "Fecha de cobro de la cuota"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54132737
         CurrentDate     =   39443
      End
      Begin VB.TextBox TxtMensaje 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   55
         ToolTipText     =   "Indica si se esta cobrando total o parcialmente la cuota"
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox ComboCobradores 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   44
         ToolTipText     =   "Lista de cobradores"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox CheckCobradores 
         Caption         =   "Cobrador"
         Height          =   195
         Left            =   3000
         TabIndex        =   43
         ToolTipText     =   "Permite seleccionar al cobrador de la cuota"
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox TxtVuelto 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   13
         Tag             =   "N"
         ToolTipText     =   "Vuelto a entregar al cliente"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TxtImporteRecibido 
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   1
         ToolTipText     =   "Importe recibido del cliente en la caja"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox ComboMonedas 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "no"
         ToolTipText     =   "Tipo de moneda del cobro (pesos, dolares, etc)"
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label LabelRecibo 
         Caption         =   "Factura Nº:"
         Height          =   255
         Left            =   3000
         TabIndex        =   60
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Fecha de cobro:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "Tipo de cobro:"
         Height          =   255
         Left            =   3000
         TabIndex        =   56
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Vuelto                $:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Importe recibido $:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Forma de cobro:"
         Height          =   255
         Left            =   6360
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del comprobante:"
      ForeColor       =   &H00FF0000&
      Height          =   4570
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   7935
      Begin VB.TextBox TxtExceptuada 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   100
         ToolTipText     =   "Indica si la cuota esta exceptuada de mora"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox TxtImporteACobrar 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   98
         Tag             =   "N"
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox TxtDiasMora 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   96
         Tag             =   "N"
         ToolTipText     =   "Dias de atraso de la factura al dia de hoy"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox TxtIvaOtorGastoRestante 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   93
         Tag             =   "N"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TxtIvaSeguroRestante 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   88
         Tag             =   "N"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TxtRefinRestante 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   87
         Tag             =   "N"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TxtVencimiento2Restante 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   86
         Tag             =   "N"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TxtOtorgamientoRestante 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   85
         Tag             =   "N"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TxtSeguroRestante 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   82
         Tag             =   "N"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TxtGastoRestante 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   80
         Tag             =   "N"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtInteresRestante 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   79
         Tag             =   "N"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TxtCapitalRestante 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   76
         Tag             =   "N"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtImporteRecargoVencimiento2 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   75
         Tag             =   "N"
         ToolTipText     =   "Recargo al segundo vencimiento"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtImporteVencimiento2 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   73
         Tag             =   "N"
         ToolTipText     =   "Importe al segundo vencimiento"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox TxtFechaVencimiento2 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   71
         ToolTipText     =   "2º vencimiento de la factura"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox TxtIvaOtorGastos 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   69
         Tag             =   "N"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtIvaSeguros 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   68
         Tag             =   "N"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtIvaInteres 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   67
         Tag             =   "N"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtImporteOtorgamiento 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   66
         Tag             =   "N"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtIvaMora 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   64
         Tag             =   "N"
         ToolTipText     =   "Importe de recargo por el atraso"
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox TxtImporteInteres 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   52
         Tag             =   "N"
         ToolTipText     =   "Interes de la cuota"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame FrameCobrada 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3360
         TabIndex        =   47
         Top             =   4080
         Visible         =   0   'False
         Width           =   4455
         Begin VB.TextBox TxtImporteCobrado 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   3240
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   51
            Tag             =   "N"
            ToolTipText     =   "Importe cobrado de la factura (si ya esta cobrada)"
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox TxtFechaCobro 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   49
            ToolTipText     =   "Fecha de cobro de la factura (si ya esta cobrada)"
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Imp.Cobrado $:"
            Height          =   255
            Left            =   2040
            TabIndex        =   50
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha cobro:"
            Height          =   255
            Left            =   0
            TabIndex        =   48
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.TextBox TxtCobrosParciales 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   45
         Tag             =   "N"
         ToolTipText     =   "Muestra el importe de cobros parciales si los tiene"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox TxtImporteRefinanciacion 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   42
         Tag             =   "N"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtImporteImpuestos 
         Height          =   285
         Left            =   6960
         MaxLength       =   9
         TabIndex        =   41
         Tag             =   "N"
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtImporteSeguros 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   40
         Tag             =   "N"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtImporteGastos 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   39
         Tag             =   "N"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtImporteCuota 
         Height          =   285
         Left            =   6960
         MaxLength       =   9
         TabIndex        =   38
         Tag             =   "N"
         Top             =   2640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtImporteAmortizacion 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   37
         Tag             =   "N"
         ToolTipText     =   "Capital de la cuota"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtImporteRecargo 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         MaxLength       =   9
         TabIndex        =   10
         Tag             =   "N"
         ToolTipText     =   "Importe de recargo"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox TxtImporteDescuento 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   9
         TabIndex        =   26
         Tag             =   "N"
         ToolTipText     =   "Importe de descuento"
         Top             =   3840
         Width           =   975
      End
      Begin VB.CheckBox CheckRecargo 
         Caption         =   "Recargos  $:"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         ToolTipText     =   "Permite aplicar recargos"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox CheckDescuento 
         Caption         =   "Descuentos         $:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Permite aplicar descuentos"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox TxtImporteRecargoMora 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   8
         Tag             =   "N"
         ToolTipText     =   "Importe de recargo por el atraso"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TxtNumCuota 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "N"
         ToolTipText     =   "Nº de cuota de la factura ingresada"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtImporteActualizado 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   11
         Tag             =   "N"
         ToolTipText     =   "Importe final a cobrar al cliente"
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox TxtImporteVencimiento1 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   7
         Tag             =   "N"
         ToolTipText     =   "Importe al primer vencimiento de la factura"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TxtFechaVencimiento1 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "1º vencimiento de la factura"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtNumFactura 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   32
         Tag             =   "N"
         ToolTipText     =   "Nº de comprobante"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Subtotal                      $:"
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Dias de mora:"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label52 
         Caption         =   "SaldoIvaOt.Gast"
         Height          =   255
         Left            =   5640
         TabIndex        =   95
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label51 
         Caption         =   "Saldo Iva.Int:"
         Height          =   255
         Left            =   5640
         TabIndex        =   94
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label50 
         Caption         =   "SaldoRec2Vto"
         Height          =   255
         Left            =   3360
         TabIndex        =   91
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label49 
         Caption         =   "SaldoIvaSeg:"
         Height          =   255
         Left            =   5640
         TabIndex        =   90
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label48 
         Caption         =   "Saldo Refin:"
         Height          =   255
         Left            =   3360
         TabIndex        =   89
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label47 
         Caption         =   "Saldo Otorg:"
         Height          =   255
         Left            =   3360
         TabIndex        =   84
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label46 
         Caption         =   "Saldo Seguros"
         Height          =   255
         Left            =   5640
         TabIndex        =   83
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label45 
         Caption         =   "Saldo Gastos"
         Height          =   255
         Left            =   5640
         TabIndex        =   81
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label44 
         Caption         =   "Saldo Interes"
         Height          =   255
         Left            =   3360
         TabIndex        =   78
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label43 
         Caption         =   "Saldo Capital"
         Height          =   255
         Left            =   3360
         TabIndex        =   77
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label41 
         Caption         =   "Cupon Nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label LabelImporte2 
         Caption         =   "Importe 2º Vto             $:"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label LabelVencimiento2 
         Caption         =   "Fecha 2º Vencimiento:"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label24 
         Caption         =   "Iva Mora                     $:"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Cobro Parcial:"
         Height          =   255
         Left            =   3360
         TabIndex        =   46
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Imp.Mora                     $:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Nº de cuota:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Importe total a cobrar $:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Importe 1º Vto             $:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha 1º vencimiento:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Nº de comprobante:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Codigo de barras/Nº de comprobante:"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   7935
      Begin VB.TextBox TxtCodigo 
         Height          =   285
         Left            =   120
         MaxLength       =   44
         TabIndex        =   0
         Tag             =   "NO"
         ToolTipText     =   "Puede leer el codigo de barras con el lector optico o ingresar el numero de factura manualmente"
         Top             =   240
         Width           =   6015
      End
   End
End
Attribute VB_Name = "FrmCobrarCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***AQUI SE COBRAN CUOTAS EN FORMA INDIVIDUAL

Dim VG_NUMFACTURA As Double

Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call RefrescarOpcionesSistema

Call LimpiarCampos(Me)

'cargo los cobradores
Call CargarComboCobradores("cobradores", ComboCobradores, True, False)

DTPicker1.Value = Date

VG_NUMFACTURA = 0

'cargo el proximo numero de factura de credimaco
TxtNumRecibo.Text = Format(UltimoId("ultimonumrecibo", "configuracionsistema"), "0000000")

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de cobro de creditos"
End Sub
Private Sub CmdLimpiar_Click()
'limpia la pantalla
Call RefreshTimer
Call LimpiarCampos(Me)
Call SetearEntorno
TxtCodigo.Text = ""
TxtCodigo.SetFocus
VG_NUMFACTURA = 0

'proximo numero de factura de credimaco
TxtNumRecibo.Text = Format(UltimoId("ultimonumrecibo", "configuracionsistema"), "0000000")

End Sub
Private Sub DTPicker1_Change()
Call BuscarFactura(2)
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
VG_NUMFACTURA = 0
Unload Me
End Sub
Private Sub CmdAnularCobro_Click()
'anula el cobro de una factura
Call RefreshTimer
Call AnularCobro
End Sub
Private Sub AnularCobro()
'anula el cobro de una cuota
Dim sql As String
Dim NumRecibo As String
On Error GoTo merror

If Trim(TxtNumCredito.Text) = "" Then Exit Sub
If Not IsNumeric(TxtNumCredito.Text) Then Exit Sub

If Trim(TxtNumCuota.Text) = "" Then Exit Sub
If Not IsNumeric(TxtNumCuota.Text) Then Exit Sub

If Not ExisteCuota(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "La cuota seleccionada no existe"
   Exit Sub
End If

If CuotaRefinanciada(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "La cuota seleccionada esta refinanciada (no esta vigente)"
   Exit Sub
End If

If CreditoBloqueado1(TxtNumCredito.Text) Then
   MsgE "La cuota pertenece a un credito bloqueado"
   Exit Sub
End If

If CuotaEsComodin(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "La cuota es comodin"
   Exit Sub
End If

If Not CuotaCobrada(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "La cuota no esta cobrada..."
   Exit Sub
End If

If Not MsgP("¿Confirma la anulacion del cobro?") Then Exit Sub

'otras validaciones
If Not ExisteCredito(TxtNumCredito.Text) Then
   MsgE "El credito no existe"
   Exit Sub
End If
If CreditoFinalizado(TxtNumCredito.Text) Then
   MsgE "El credito esta finalizado"
   Exit Sub
End If
If CreditoBloqueado1(TxtNumCredito.Text) Then
   MsgE "El credito esta bloqueado"
   Exit Sub
End If
If Not ExisteCuota(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "El comprobante ingresado no existe"
   Exit Sub
End If
If CuotaRefinanciada(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "La cuotao esta refinanciada (no esta vigente)"
   Exit Sub
End If
If CuotaEsComodin(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "La cuota es comodin"
   Exit Sub
End If
If Not CuotaCobrada(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "La cuota esta cobrada..."
   Exit Sub
End If

NumRecibo = ""

'inicio la transaccion
cnSQL.BeginTrans

'esta cobrada y esta en condiciones de anularse el cobro
sql = "update cuotas set fechacobro=null,importecobrado=ccur(0)," & _
      "importedescuentos=ccur(0),importerecargos=ccur(0)," & _
      "importemora=ccur(0),ivamora=0,importeparcial=ccur(0)," & _
      "cobrosparciales=false,idmoneda=clng(0),pagofacil=false,rapipago=false " & _
      "where idcredito='" & CLng(TxtNumCredito.Text) & "' and numcuota='" & CLng(TxtNumCuota.Text) & "'"

cnSQL.Execute sql

'si el credito estaba finalizado lo desmarco como no finalizado
If CreditoFinalizado(TxtNumCredito.Text) Then
   sql = "update creditos set fechafinalizacion=null " & _
         "where idcredito='" & CLng(TxtNumCredito.Text) & "'"
   cnSQL.Execute (sql)
End If

'Ahora borro las entradas de cobradores idem anterior
sql = "delete from cobradorespagos " & _
      "where idcredito=" & CLng(TxtNumCredito.Text) & " and numcuota=" & CLng(TxtNumCuota.Text)
cnSQL.Execute sql

'BORRO LAS ENTRADAS EN INGRESOS DE ESA CUOTA
sql = "delete from ingresos where idcredito=" & CLng(TxtNumCredito.Text) & " and numcuota=" & CLng(TxtNumCuota.Text)
cnSQL.Execute sql

'finalizo la transaccion
cnSQL.CommitTrans

'blanqueo el campo de codigo de barras
Call LimpiarCampos(Me)
Call SetearEntorno

TxtCodigo.Text = ""
TxtCodigo.SetFocus
VG_NUMFACTURA = 0

MsgI "Se anulo el cobro de la cuota"

Exit Sub
merror:
tratarerrores "Error anulando el cobro de cuotas"
End Sub
Private Sub BuscarFactura(ByVal Opcion As Long)
'busca la boleta del codigo de barras o del nº de factura
Dim rec As rdoResultset
Dim sql As String
Dim NumFactura As Long
Dim SaldoCuota As Currency
Dim SaldoCuota1erVenc As Currency
Dim ImporteParcial As Currency
Dim ImporteMora As Currency
Dim IvaMora As Currency
Dim RecargoCuota As Currency
Dim ImporteMoraGral As Currency
Dim SoloMoraCobrada As Currency
Dim SoloIvaMoraCobrada As Currency
On Error GoTo merror

TxtDiasMora.Text = 0
TxtImporteRecargoMora.Text = 0

If Opcion = 1 Then
   If Not DatosFacturaOk() Then Exit Sub
   NumFactura = CapturarNumFactura()
   VG_NUMFACTURA = NumFactura
   If Not ExisteFactura(NumFactura) Then
      MsgE "El comprobante no existe"
      Exit Sub
   End If
Else
   NumFactura = VG_NUMFACTURA
End If

'verifico que sea una boleta vigente
If FacturaRefinanciada(NumFactura) Then
   MsgE "El comprobante esta refinanciado (no tiene vigencia)"
   Exit Sub
End If

If CreditoBloqueado(NumFactura) Then
   MsgE "El comprobante pertenece a un credito bloqueado"
   Exit Sub
End If

If CreditoFinalizado2(NumFactura) Then
   MsgE "El comprobante pertenece a un credito finalizado"
   Exit Sub
End If

If FacturaEsComodin(NumFactura) Then
   MsgE "El comprobante pertenece a una cuota comodin"
   Exit Sub
End If

'busco los datos de la factura
sql = "select creditos.idcredito as numcredito,creditos.numcuotas,creditos.codprestamo," & _
      "clientes.idcliente,clientes.apellido + ' ' + clientes.nombre as titular," & _
      "cuotas.*,cuotas.logic1 as exceptuada," & _
      "(cuotas.importevencimiento1) as importetotal " & _
      "from clientes inner join (creditos inner join cuotas " & _
      "on creditos.idcredito=cuotas.idcredito) " & _
      "on clientes.idcliente=creditos.idcliente " & _
      "where cuotas.numfactura='" & CLng(NumFactura) & "'"

Set rec = cnSQL.OpenResultset(sql)

TxtImporteRecibido.Text = ""
TxtVuelto.Text = 0

If Not rec.EOF Then
   'saldo del credito incluye mora e iva mora si se aplica impuestocredimaco
   TxtSaldo.Text = Format(ObtenerSaldoCredito(rec.rdoColumns("idcredito"), DTPicker1.Value), "0.00")
   
   'obtengo los campos de cliente y factura
   txtnumfactura.Text = Format(rec.rdoColumns("numfactura"), "000000000")
   TxtCodPrestamo.Text = rec.rdoColumns("codprestamo")
   
   'campos que componen la cuota
   TxtImporteCuota.Text = CCur(rec.rdoColumns("importecuota"))
   TxtImporteAmortizacion.Text = Format(CCur(rec.rdoColumns("importeamortizacion")), "0.00")
   TxtImporteInteres.Text = Format(CCur(rec.rdoColumns("importeinteres")), "0.00")
   TxtImporteRefinanciacion.Text = CCur(rec.rdoColumns("importerefinanciacion"))
   TxtImporteGastos.Text = CCur(rec.rdoColumns("importegastos"))
   TxtImporteSeguros.Text = CCur(rec.rdoColumns("importeseguros"))
   TxtIvaInteres.Text = CCur(rec.rdoColumns("ivainteres"))
   TxtIvaSeguros.Text = CCur(rec.rdoColumns("ivaseguros"))
   TxtIvaOtorGastos.Text = CCur(rec.rdoColumns("ivaotorgamientogastos"))
   TxtImporteOtorgamiento.Text = CCur(rec.rdoColumns("otorgamiento"))
   'verificar estos..porque se calcula ahora
   TxtImporteRecargoMora.Text = 0
   TxtIvaMora.Text = 0
   TxtImporteRecargoVencimiento2.Text = CCur(rec.rdoColumns("importerecargovencimiento2"))
   TxtImporteRecargoVencimiento2.Text = Format(TxtImporteRecargoVencimiento2.Text, "0.00")
   If rec.rdoColumns("exceptuada") Then
      TxtExceptuada.Text = "(E)"
   Else
      TxtExceptuada.Text = " "
   End If
   
   'nuevos restantes
   TxtCapitalRestante.Text = Format(CCur(TxtImporteAmortizacion.Text) - ObtenerCapitalCobrado(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota")), "0.00")
   TxtInteresRestante.Text = Format(CCur(TxtImporteInteres.Text) - ObtenerInteresCobrado(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota")), "0.00")
   
   'si esta pendiente
   If IsNull(rec.rdoColumns("fechacobro")) Then
      'y se aplica el 2 vto
      If CDate(DTPicker1.Value) > CDate(rec.rdoColumns("fechavencimiento1")) Then
         TxtVencimiento2Restante.Text = CCur(TxtImporteRecargoVencimiento2.Text) - ObtenerVencimiento2Cobrado(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
      End If
   Else
      'si esta cobrada despues del vto 2
      If CDate(rec.rdoColumns("fechacobro")) > CDate(rec.rdoColumns("fechavencimiento1")) Then
         TxtVencimiento2Restante.Text = CCur(TxtImporteRecargoVencimiento2.Text) - ObtenerVencimiento2Cobrado(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota"))
      End If
   End If
   TxtVencimiento2Restante.Text = Format(TxtVencimiento2Restante.Text, "0.00")
   TxtRefinRestante.Text = Format(CCur(TxtImporteRefinanciacion.Text) - ObtenerRefinCobrado(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota")), "0.00")
   TxtGastoRestante.Text = Format(CCur(TxtImporteGastos.Text) - ObtenerGastosCobrados(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota")), "0.00")
   TxtSeguroRestante.Text = Format(CCur(TxtImporteSeguros.Text) - ObtenerSegurosCobrados(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota")), "0.00")
   TxtOtorgamientoRestante.Text = Format(CCur(TxtImporteOtorgamiento.Text) - ObtenerOtorgamientoCobrado(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota")), "0.00")
   TxtIvaInteresRestante.Text = Format(CCur(TxtIvaInteres.Text) - ObtenerIvaInteresCobrado(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota")), "0.00")
   TxtIvaSeguroRestante.Text = Format(CCur(TxtIvaSeguros.Text) - ObtenerIvaSegurosCobrado(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota")), "0.00")
   TxtIvaOtorGastoRestante.Text = Format(CCur(TxtIvaOtorGastos.Text) - ObtenerIvaOtorGastosCobrado(rec.rdoColumns("idcredito"), rec.rdoColumns("numcuota")), "0.00")
   TxtCliente.Text = rec.rdoColumns("titular") & vbNullString
   TxtNumCredito.Text = Format(rec.rdoColumns("numcredito"), "000000")
   TxtNumCuota.Text = Format(rec.rdoColumns("numcuota"), "00")
   TxtTotalCuotas.Text = Format(rec.rdoColumns("numcuotas"), "00")
   TxtFechaVencimiento1.Text = rec.rdoColumns("fechavencimiento1")
   TxtFechaVencimiento2.Text = rec.rdoColumns("fechavencimiento2")
   
   ImporteParcial = ObtenerImporteParcialX(TxtNumCredito.Text, TxtNumCuota.Text)
   
   SaldoCuota = ObtenerSaldoCuotaX(TxtNumCredito.Text, TxtNumCuota.Text, DTPicker1.Value, SaldoCuota1erVenc)
   
   'aca pongo el orginal
   TxtImporteVencimiento1.Text = Format(CCur(rec.rdoColumns("importetotal")), "0.00")
   
   TxtImporteVencimiento2.Text = Format(rec.rdoColumns("ImporteVencimiento2"), "0.00")
   
   'si esta sin cobrar si es necesario actualizo importes
   'por defecto el importe de pago es el primero
   TxtImporteActualizado.Text = Format(CCur(TxtImporteVencimiento1.Text), "0.00")
   
   'obtengo el importe de cobros parciales hasta la fecha
   TxtCobrosParciales.Text = Format(ImporteParcial, "0.00")
   
   'obtiene el proximo numero de recibo
   TxtNumRecibo.Text = Format(UltimoId("ultimonumrecibo", "configuracionsistema"), "0000000")

   If IsNull(rec.rdoColumns("fechacobro")) Then
      CheckDescuento.Value = 0
      CheckRecargo.Value = 0
      TxtImporteDescuento.Text = 0
      TxtImporteRecargo.Text = 0
      FrameCobrada.Visible = False
      txtFechacobro.Text = ""
      TxtImporteCobrado.Text = 0
      TxtImporteDescuento.Enabled = False
      TxtImporteRecargo.Enabled = False
         
      'si estoy antes del primer vencimiento
      If CDate(DTPicker1.Value) <= CDate(TxtFechaVencimiento1.Text) Then
         TxtImporteActualizado.Text = Format(CCur(SaldoCuota), "0.00")
         TxtImporteRecargoMora.Text = 0
         TxtIvaMora.Text = 0
      End If
   
      'si estoy despues del 1º vencimiento pero antes del segundo
      If CDate(DTPicker1.Value) > CDate(TxtFechaVencimiento1.Text) And CDate(DTPicker1.Value) <= CDate(TxtFechaVencimiento2.Text) Then
         'vale el segundo..no hay mora
         TxtImporteActualizado.Text = Format(SaldoCuota, "0.00")
         TxtImporteRecargoMora.Text = 0
         TxtIvaMora.Text = 0
      End If
            
      'si estoy despues del segundo vencimiento hay mora (vale para los dos vtos)
      If CDate(DTPicker1.Value) > CDate(TxtFechaVencimiento2.Text) Then
         TxtDiasMora.Text = DateDiff("d", CDate(TxtFechaVencimiento1.Text), CDate(DTPicker1.Value))
         
         'calculo la mora en forma habitual
         'puedo pasarle el campo [exceptuada]
         ImporteMora = CalcularInteresMoraZZ(rec.rdoColumns("exceptuada"), SaldoCuota, TxtFechaVencimiento2.Text, CDate(DTPicker1.Value), TxtFechaVencimiento1.Text)
         
         IvaMora = 0
         If VG_APLICARIMPUESTOS Then
            If VG_IMPUESTOSCREDIMACO Then
               IvaMora = CCur(VG_PORCENTAJEIVA * ImporteMora / 100)
            End If
         End If
         SoloMoraCobrada = ObtenerMoraCobrada(TxtNumCredito.Text, TxtNumCuota.Text)
         SoloIvaMoraCobrada = ObtenerIvaMoraCobrada(TxtNumCredito.Text, TxtNumCuota.Text)
         'hago el ajuste necesario
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
         'revizar si debo acumular la mora y el ivamora o manejarlas por separado
         TxtImporteActualizado.Text = Format(CCur(SaldoCuota) + CCur(ImporteMora) + CCur(IvaMora), "0.00")
         TxtImporteRecargoMora.Text = Format(ImporteMora, "0.00")
         'ahora separe la mora del iva mora para mayor claridad
         TxtIvaMora.Text = Format(CCur(IvaMora), "0.00")
      End If
        
   Else
      'si esta cobrada
      FrameCobrada.Visible = True
      txtFechacobro.Text = rec.rdoColumns("fechacobro") & vbNullString
      TxtImporteCobrado.Text = Format(CCur(rec.rdoColumns("importecobrado")), "0.00")
      TxtDiasMora.Text = 0
      TxtImporteRecargoMora.Text = 0
      TxtIvaMora.Text = 0
      TxtImporteRecargo.Text = 0
      TxtImporteDescuento.Text = 0
      TxtImporteActualizado.Text = 0
      TxtCobrosParciales.Text = 0
      CheckDescuento.Value = 0
      CheckDescuento.Enabled = False
      CheckRecargo.Value = 0
      CheckRecargo.Enabled = False
   End If
   
   TxtImporteACobrar.Text = TxtImporteActualizado.Text
   
End If

Exit Sub
merror:
tratarerrores "Error buscando el comprobante"
End Sub
Private Function CreditoFinalizado2(ByVal NumFactura As Long) As Boolean
'verifica si un credito esta finalizado segun su numero de factura
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

CreditoFinalizado2 = False

sql = "select creditos.fechafinalizacion " & _
      "from creditos inner join cuotas on creditos.idcredito=cuotas.idcredito " & _
      "where cuotas.numfactura='" & CLng(NumFactura) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("fechafinalizacion")) Then
      CreditoFinalizado2 = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CreditoFinalizado2"
End Function
Private Sub cmdaceptar_Click()
Call RefreshTimer
Call RegistrarCobro
End Sub
Private Sub RegistrarCobro()
'registro el cobro
Dim rec As rdoResultset
Dim sql As String
Dim IDmoneda As Long
Dim ImporteVencimiento2 As Currency
Dim IdCobrador As Long
Dim ImporteCobrador As Currency
Dim CobroParcial As Boolean
Dim UltimaParcial As Boolean
Dim IdCobroParcial As Long
Dim IdCobradorPago As Long
Dim Observaciones As String
Dim ImporteParcial As Currency
Dim ImporteRealCobrado As Currency
Dim ImporteRealCobrado2 As Currency
Dim IdIngreso As Long
Dim NumRecibo As String
Dim CodPrestamo As String
Dim IvaMora As Currency
Dim OtorgamientoCuota As Currency
Dim IvaOtorGastos As Currency
Dim ImporteXXX As Currency
Dim MoraCobrada As Currency
Dim IvaMoraCobrada As Currency
Dim Vencimiento2Cobrado As Currency
Dim RefinCobrado As Currency
Dim GastosCobrados As Currency
Dim OtorgamientoCobrado As Currency
Dim IvaOtorGastosCobrado As Currency
Dim SegurosCobrados As Currency
Dim IvaSegurosCobrado As Currency
Dim CapitalCobrado As Currency
Dim InteresCobrado As Currency
Dim IvaInteresCobrado As Currency
Dim Item As String
Dim IdItemParcial As Long
Dim ImporteItem As Currency
Dim Diferencia As Currency
Dim Vencimiento2Cuota As Currency
Dim DescuentoCuota As Currency
Dim RecargoCuota As Currency
Dim Diferencia2 As Currency
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
Dim Condicion As String
On Error GoTo merror

If Trim(TxtNumCredito.Text) = "" Then Exit Sub
If Not IsNumeric(TxtNumCredito.Text) Then Exit Sub

If Trim(TxtNumCuota.Text) = "" Then Exit Sub
If Not IsNumeric(TxtNumCuota.Text) Then Exit Sub

'valido si esta permitiendo cobros diferidos
'si no permito cobros diferidos
If Not VG_COBROSDIFERIDOS Then
   If CDate(DTPicker1.Value) <> CDate(Date) Then
      MsgE "Verifique la fecha de cobro...debe ser igual a la actual"
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

'solo valido si aplico recibos=factura en credimaco
If VG_APLICARRECIBOS Then
   'valido el numero de recibo si lo ingresaron a mano
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
End If

'verifico si existe
If Not ExisteCuota(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "El comprobante no existe"
   Exit Sub
End If

'valido si es vigente
If CuotaRefinanciada(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "El comprobante esta refinanciado (no esta vigente)"
   Exit Sub
End If

If CreditoBloqueado1(TxtNumCredito.Text) Then
   MsgE "El comprobante pertenece a un credito bloqueado"
   Exit Sub
End If

If CreditoFinalizado(TxtNumCredito.Text) Then
   MsgE "El comprobante pertenece a un credito finalizado"
   Exit Sub
End If

If CuotaCobrada(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "El comprobante ya esta cobrado!!!"
   Exit Sub
End If

'verifica si es comodin o ultima comodin
If CuotaEsComodin(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "El comprobante pertenece a una cuota comodin"
   Exit Sub
End If

'si no permito salteadas verifico
If Not VG_PAGARCUOTASDESORDENADAS Then
   If HayCuotasImpagasPrevias(TxtNumCredito.Text, TxtNumCuota.Text) Then
      MsgE "Hay cuotas impagas anteriores a esta"
      Exit Sub
   End If
End If

If Trim(TxtImporteActualizado.Text) = "" Then
   MsgE "Falta el importe total del comprobante"
   Exit Sub
End If

'valido el importe recibido
If Trim(TxtImporteRecibido.Text) = "" Then
   MsgE "Debe ingresar el importe a cobrar recibido del cliente"
   TxtImporteRecibido.SetFocus
   Exit Sub
End If
If Not IsNumeric(TxtImporteRecibido.Text) Then
   MsgE "El importe a cobrar debe ser numerico"
   TxtImporteRecibido.SetFocus
   Exit Sub
End If
If CCur(TxtImporteRecibido.Text) <= 0 Then
   MsgE "El importe a cobrar debe ser mayor a cero"
   TxtImporteRecibido.SetFocus
   Exit Sub
End If

If CheckDescuento.Value = 1 Then
   'verifico los campos descuentos y recargos
   If Trim(TxtImporteDescuento.Text) = "" Then
      TxtImporteDescuento.Text = 0
   End If
   If Not IsNumeric(TxtImporteDescuento.Text) Then
      TxtImporteDescuento.Text = 0
   End If
   If CCur(TxtImporteDescuento.Text) < 0 Then
      TxtImporteDescuento.Text = 0
   End If
   If CCur(TxtImporteDescuento.Text) > 0 Then
      If CCur(TxtImporteRecibido.Text) < CCur(TxtImporteActualizado.Text) Then
         MsgE "Solo se puede descontar si el cobro es total..(no con cobros parciales)"
         Exit Sub
      End If
      If CCur(TxtImporteDescuento.Text) > CCur(TxtImporteActualizado.Text) Then
         MsgE "El importe de descuento debe ser menor al importe a pagar"
         Exit Sub
      End If
   End If
Else
   TxtImporteDescuento.Text = 0
End If

If CheckRecargo.Value = 1 Then
   'verifico los campos descuentos y recargos
   If Trim(TxtImporteRecargo.Text) = "" Then
      TxtImporteRecargo.Text = 0
   End If
   If Not IsNumeric(TxtImporteRecargo.Text) Then
      TxtImporteRecargo.Text = 0
   End If
   If CCur(TxtImporteRecargo.Text) < 0 Then
      TxtImporteRecargo.Text = 0
   End If
   
   If CCur(TxtImporteRecargo.Text) > 0 Then
      If CCur(TxtImporteRecibido.Text) < CCur(TxtImporteActualizado.Text) Then
         MsgE "Solo se puede recargar si el cobro es total..(no con cobros parciales)"
         Exit Sub
      End If
   End If
Else
   TxtImporteRecargo.Text = 0
End If

IDmoneda = 1

IdCobrador = 0
ImporteCobrador = 0

'verifico si seleccionaron cobrador
If CheckCobradores.Value = 1 Then
   If ComboCobradores.Text <> "" Then
      IdCobrador = CLng(ComboCobradores.ItemData(ComboCobradores.ListIndex))
      ImporteCobrador = ObtenerComisionCobrador(IdCobrador, ImporteRealCobrado)
   End If
End If

CobroParcial = False
UltimaParcial = False

'si no permito cobros parciales exijo el importe total
If (CCur(TxtImporteRecibido.Text)) < CCur(TxtImporteActualizado.Text) Then
   If VG_APLICARCOBROSPARCIALES Then
      If Not MsgP("¿Confirma el cobro PARCIAL de la cuota?") Then Exit Sub
      CobroParcial = True
   Else
      MsgE "El importe a cobrar es incorrecto...debe ser igual o mayor al importe a pagar"
      TxtImporteRecibido.SetFocus
      Exit Sub
   End If
Else
   'si el importe es igual no significa que no es parcial,porque puedo estar grabando
   'la ultima parcial y es igual al importe a cobrar.
   If CCur(TxtCobrosParciales.Text) > 0 Then
      'estoy saldando la ultima cuota
      If Not MsgP("¿Confirma el cobro de la ULTIMA cuota parcial?") Then Exit Sub
      CobroParcial = True
      UltimaParcial = True
   Else
      If Not MsgP("¿Confirma el cobro de la cuota?") Then Exit Sub
   End If
End If

'valido el credito
If Not ExisteCredito(TxtNumCredito.Text) Then
   MsgE "El credito no existe"
   Exit Sub
End If

If CreditoFinalizado(TxtNumCredito.Text) Then
   MsgE "El credito esta finalizado"
   Exit Sub
End If

If CreditoBloqueado1(TxtNumCredito.Text) Then
   MsgE "El credito esta bloqueado"
   Exit Sub
End If

'verifico si existe
If Not ExisteCuota(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "El comprobante no existe"
   Exit Sub
End If

'valido si es vigente
If CuotaRefinanciada(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "El comprobante esta refinanciado (no esta vigente)"
   Exit Sub
End If

If CuotaCobrada(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "El comprobante ya esta cobrado!!!"
   Exit Sub
End If

'verifica si es comodin o ultima comodin
If CuotaEsComodin(TxtNumCredito.Text, TxtNumCuota.Text) Then
   MsgE "El comprobante pertenece a una cuota comodin"
   Exit Sub
End If

'valido el cobrador si esta seleccionado
If IdCobrador > 0 Then
   If Not ExisteCobrador(IdCobrador) Then
      MsgE "El cobrador no existe"
      Exit Sub
   End If
End If

ImporteRealCobrado = CCur(TxtImporteRecibido.Text) - CCur(TxtVuelto.Text)
ImporteRealCobrado2 = CCur(TxtImporteRecibido.Text) - CCur(TxtVuelto.Text)

'este es el recargo por vencimiento2
Vencimiento2Cuota = CCur(TxtImporteVencimiento2.Text) - CCur(TxtImporteVencimiento1.Text)

If CDate(DTPicker1.Value) <= CDate(TxtFechaVencimiento1.Text) Then
   Vencimiento2Cuota = 0
End If

'inicio transaccion
cnSQL.BeginTrans
   
'si registro el pago total de la cuota correspondiente
If Not CobroParcial Then
   'aca debo cubrir todos los items...salda toda la cuota
   sql = "update cuotas set fechacobro='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "'," & _
         "importecobrado='" & ConvertirDblSql(CCur(TxtImporteActualizado.Text)) & "'," & _
         "formacobro='" & CStr(FormaCobro) & "',idmoneda='" & CLng(IDmoneda) & "'" & _
         " where idcredito='" & CLng(TxtNumCredito.Text) & "' and numcuota='" & CLng(TxtNumCuota.Text) & "'"
   cnSQL.Execute sql
   
   'esto es para compatibilizar con ingresos
   CapitalCobrado = CCur(TxtImporteAmortizacion.Text)
   InteresCobrado = CCur(TxtImporteInteres.Text)
   GastosCobrados = CCur(TxtImporteGastos.Text)
   SegurosCobrados = CCur(TxtImporteSeguros.Text)
   OtorgamientoCobrado = CCur(TxtImporteOtorgamiento.Text)
   Vencimiento2Cobrado = CCur(Vencimiento2Cuota)
   RefinCobrado = CCur(TxtImporteRefinanciacion.Text)
   IvaInteresCobrado = CCur(TxtIvaInteres.Text)
   IvaSegurosCobrado = CCur(TxtIvaSeguros.Text)
   IvaOtorGastosCobrado = CCur(TxtIvaOtorGastos.Text)
   MoraCobrada = CCur(TxtImporteRecargoMora.Text)
   IvaMoraCobrada = CCur(TxtIvaMora.Text)
   DescuentoCuota = CCur(TxtImporteDescuento.Text)
   RecargoCuota = CCur(TxtImporteRecargo.Text)
  
Else
   'es un cobro parcial
   
   'si era la ultima parcial termino la cuota
   ImporteParcial = ObtenerImporteParcialX(TxtNumCredito.Text, TxtNumCuota.Text) + CCur(ImporteRealCobrado)
   
   'si es la ultima parcial que salda la cuota
   If UltimaParcial Then
      'grabo la ultima SIN IMPORTAR EL ORDEN DE LOS ITEMS
          
       sql = "update cuotas set fechacobro='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "'," & _
             "importecobrado='" & ConvertirDblSql(CCur(ImporteParcial)) & "'," & _
             "importedescuentos='" & ConvertirDblSql(CCur(TxtImporteDescuento.Text)) & "'," & _
             "importerecargos='" & ConvertirDblSql(CCur(TxtImporteRecargo.Text)) & "'," & _
             "formacobro='" & CStr(FormaCobro) & "' " & _
             "where idcredito='" & CLng(TxtNumCredito.Text) & "' and numcuota='" & CLng(TxtNumCuota.Text) & "'"
        cnSQL.Execute sql
        
        'esto es para ingresos
        CapitalCobrado = CCur(TxtCapitalRestante.Text)
        InteresCobrado = CCur(TxtInteresRestante.Text)
        Vencimiento2Cobrado = CCur(TxtVencimiento2Restante.Text)
        RefinCobrado = CCur(TxtRefinRestante.Text)
        GastosCobrados = CCur(TxtGastoRestante.Text)
        SegurosCobrados = CCur(TxtSeguroRestante.Text)
        OtorgamientoCobrado = CCur(TxtOtorgamientoRestante.Text)
        IvaInteresCobrado = CCur(TxtIvaInteresRestante.Text)
        IvaSegurosCobrado = CCur(TxtIvaSeguroRestante.Text)
        IvaOtorGastosCobrado = CCur(TxtIvaOtorGastoRestante.Text)
        MoraCobrada = CCur(TxtImporteRecargoMora.Text)
        IvaMoraCobrada = CCur(TxtIvaMora.Text)
   
   Else 'si ultimaparcial
      'ES UN COBRO PARCIAL COMUN...ACA SI IMPORTA EL ORDEN
                    
      'si queda resto
      IvaMoraCobrada = 0
      If CCur(ImporteRealCobrado) > 0 Then
         'si hay IVA mora intento cubrirla
         If CCur(TxtIvaMora.Text) > 0 Then
            If CCur(ImporteRealCobrado) >= CCur(TxtIvaMora.Text) Then
               ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtIvaMora.Text)
               IvaMoraCobrada = CCur(TxtIvaMora.Text)
            Else
               IvaMoraCobrada = CCur(ImporteRealCobrado)
               ImporteRealCobrado = 0
            End If
        End If
      End If
      
      'ya esta agregado el parcial arriba en cuotasparciales
      'primero la mora
      MoraCobrada = 0
      If CCur(ImporteRealCobrado) > 0 Then
         If CCur(TxtImporteRecargoMora.Text) > 0 Then
            If CCur(ImporteRealCobrado) >= CCur(TxtImporteRecargoMora.Text) Then
               ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtImporteRecargoMora.Text)
               MoraCobrada = CCur(TxtImporteRecargoMora.Text)
            Else
               MoraCobrada = CCur(ImporteRealCobrado)
               ImporteRealCobrado = 0
            End If
         End If
      End If
                      
      'a partir de aqui tienen techo
      'vencimiento2
      'si queda aun resto
      Vencimiento2Cobrado = 0
      'si cobre despues del 1 vto
      If CDate(DTPicker1.Value) > CDate(TxtFechaVencimiento1.Text) Then
         If CCur(ImporteRealCobrado) > 0 Then
            'si sigue intento cubrir
            If CCur(TxtVencimiento2Restante.Text) > 0 Then
               If CCur(ImporteRealCobrado) >= CCur(TxtVencimiento2Restante.Text) Then
                  ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtVencimiento2Restante.Text)
                  Vencimiento2Cobrado = CCur(TxtVencimiento2Restante.Text)
               Else
                  Vencimiento2Cobrado = CCur(ImporteRealCobrado)
                  ImporteRealCobrado = 0
               End If
            End If
          End If
       End If
                     
        'refinanciacion
        RefinCobrado = 0
        If CCur(ImporteRealCobrado) > 0 Then
          'si sigue intento cubrr el iva interes
          If CCur(TxtRefinRestante.Text) > 0 Then
             If CCur(ImporteRealCobrado) >= CCur(TxtRefinRestante.Text) Then
                ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtRefinRestante.Text)
                RefinCobrado = CCur(TxtRefinRestante.Text)
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
          If CCur(TxtIvaOtorGastoRestante.Text) > 0 Then
             If CCur(ImporteRealCobrado) >= CCur(TxtIvaOtorGastoRestante.Text) Then
                ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtIvaOtorGastoRestante.Text)
                IvaOtorGastosCobrado = CCur(TxtIvaOtorGastoRestante.Text)
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
           If CCur(TxtIvaSeguroRestante.Text) > 0 Then
              If CCur(ImporteRealCobrado) >= CCur(TxtIvaSeguroRestante.Text) Then
                 ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtIvaSeguroRestante)
                 IvaSegurosCobrado = CCur(TxtIvaSeguroRestante.Text)
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
           If CCur(TxtIvaInteresRestante.Text) > 0 Then
              If CCur(ImporteRealCobrado) >= CCur(TxtIvaInteresRestante.Text) Then
                 ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtIvaInteresRestante.Text)
                 IvaInteresCobrado = CCur(TxtIvaInteresRestante.Text)
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
          If CCur(TxtOtorgamientoRestante.Text) > 0 Then
             If CCur(ImporteRealCobrado) >= CCur(TxtOtorgamientoRestante.Text) Then
                ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtOtorgamientoRestante.Text)
                OtorgamientoCobrado = CCur(TxtOtorgamientoRestante.Text)
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
          If CCur(TxtGastoRestante.Text) > 0 Then
             If CCur(ImporteRealCobrado) >= CCur(TxtGastoRestante.Text) Then
                ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtGastoRestante.Text)
                GastosCobrados = CCur(TxtGastoRestante.Text)
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
           If CCur(TxtSeguroRestante.Text) > 0 Then
              If CCur(ImporteRealCobrado) >= CCur(TxtSeguroRestante.Text) Then
                 ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtSeguroRestante.Text)
                 SegurosCobrados = CCur(TxtSeguroRestante.Text)
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
           If CCur(TxtInteresRestante.Text) > 0 Then
              If CCur(ImporteRealCobrado) >= CCur(TxtInteresRestante.Text) Then
                 ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtInteresRestante.Text)
                 InteresCobrado = CCur(TxtInteresRestante.Text)
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
           If CCur(TxtCapitalRestante.Text) > 0 Then
              If CCur(ImporteRealCobrado) >= CCur(TxtCapitalRestante.Text) Then
                 ImporteRealCobrado = CCur(ImporteRealCobrado) - CCur(TxtCapitalRestante.Text)
                 CapitalCobrado = CCur(TxtCapitalRestante.Text)
              Else
                 CapitalCobrado = CCur(ImporteRealCobrado)
                 ImporteRealCobrado = 0
              End If
           End If
        End If
                      
        'grabo el cobro parcial
        sql = "update cuotas set " & _
              "cobrosparciales='True',formacobro='" & CStr(FormaCobro) & "' " & _
              "where idcredito='" & CLng(TxtNumCredito.Text) & "' and numcuota='" & CLng(TxtNumCuota.Text) & "'"
        cnSQL.Execute sql
                    
        HuboCobros = True
   End If
  
End If

txtFechacobro.Text = ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY")
TxtImporteCobrado.Text = CCur(TxtImporteActualizado.Text)

If VG_FINALIZARAUTOMATICAMENTE Then
   'si es la ultima cuota finalizo el credito
   Call FinalizarCredito(TxtNumCredito.Text, DTPicker1.Value)
End If

IdIngreso = UltimoId("idingreso", "ingresos") + 1

'aca grabo ahora los items cobrados
sql = "insert into ingresos (idingreso,idcredito,numcuota," & _
      "fechacobro,importecobrado,numrecibo,codprestamo,numcomprobante,capitalcobrado,interescobrado,vencimiento2cobrado,refincobrado,gastoscobrados,seguroscobrados,otorgamientocobrado,ivainterescobrado,ivaseguroscobrado,ivaotorgastoscobrado,moracobrada,ivamoracobrada,descuentos,recargos,usuario) " & _
      "values('" & CLng(IdIngreso) & "','" & CLng(TxtNumCredito.Text) & "','" & CLng(TxtNumCuota.Text) & _
      "','" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "','" & ConvertirDblSql(CCur(ImporteRealCobrado2)) & "','" & CStr(TxtNumRecibo.Text) & "','" & CStr(TxtCodPrestamo.Text) & _
      "','" & CLng(txtnumfactura.Text) & "','" & ConvertirDblSql(CCur(CapitalCobrado)) & "','" & ConvertirDblSql(CCur(InteresCobrado)) & "','" & ConvertirDblSql(CCur(Vencimiento2Cobrado)) & "','" & ConvertirDblSql(CCur(RefinCobrado)) & "','" & ConvertirDblSql(CCur(GastosCobrados)) & "','" & ConvertirDblSql(CCur(SegurosCobrados)) & "','" & ConvertirDblSql(CCur(OtorgamientoCobrado)) & "','" & ConvertirDblSql(CCur(IvaInteresCobrado)) & "','" & ConvertirDblSql(CCur(IvaSegurosCobrado)) & "','" & ConvertirDblSql(CCur(IvaOtorGastosCobrado)) & "','" & ConvertirDblSql(CCur(MoraCobrada)) & "','" & ConvertirDblSql(CCur(IvaMoraCobrada)) & "','" & ConvertirDblSql(CCur(TxtImporteDescuento.Text)) & "','" & ConvertirDblSql(CCur(TxtImporteRecargo.Text)) & "','" & CStr(VG_USUARIOLOGIN) & "')"
cnSQL.Execute sql

'grabo datos de cobradores
If ComboCobradores.Text <> "" Then
   'obtengo el proximo id de cobradorespagos
   IdCobradorPago = UltimoId("idcobradorpago", "cobradorespagos") + 1
   ImporteCobrador = ObtenerComisionCobrador(IdCobrador, ImporteRealCobrado)
  
   sql = "insert into cobradorespagos (idcobradorpago,idcobrador,idcredito,numcuota,numfactura," & _
         "importecobrador,fecha) " & _
         "values('" & CLng(IdCobradorPago) & "','" & CLng(IdCobrador) & "','" & CLng(TxtNumCredito.Text) & "','" & CLng(TxtNumCuota.Text) & "','" & CLng(txtnumfactura.Text) & _
         "','" & ConvertirDblSql(CCur(ImporteCobrador)) & "','" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "')"
   cnSQL.Execute sql
End If

'ahora actualizo el numero de recibo y lo pongo en el proximo a imprimir(o a asociar)
sql = "update configuracionsistema set ultimonumrecibo=ultimonumrecibo + 1"
cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

MsgI "Se registro el cobro de la cuota"

If MsgP("¿Desea imprimir la factura?") Then
   Condicion = "creditos.idcredito='" & CLng(TxtNumCredito.Text) & "' and cuotas.numcuota='" & CLng(TxtNumCuota.Text) & "'"
   Call ImprimirFacturaCredimaco(Condicion, DTPicker1.Value)
End If
'actualizo el nuevo numero de recibo
TxtNumRecibo.Text = Format(UltimoId("ultimonumrecibo", "configuracionsistema"), "0000000")

'blanqueo el campo de codigo de barras
Call LimpiarCampos(Me)
Call SetearEntorno

TxtCodigo.Text = ""
TxtCodigo.SetFocus

Exit Sub
merror:
tratarerrores "Error registrando el cobro de cuota"
End Sub
Private Function DatosFacturaOk() As Boolean
'valida si se cargo el codigo de barras o solo la factura
On Error GoTo merror

DatosFacturaOk = True

'si no hay factura no se cargo un codigo de barras o numero de factura a mano
If Trim(txtnumfactura.Text) = "" Then
   DatosFacturaOk = False
   MsgE "Debe leer un codigo de barras o ingresar un numero de comprobante"
   TxtCodigo.SetFocus
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosFacturaOk-CobrarCreditos"
End Function
Private Function CapturarNumFactura() As Long
'extrae el numero de factura del codigo de barras o de la factura
On Error GoTo merror

'el numero de factura va desde la posicion 12 ocupando 11 posiciones
'por ahora le saco solo 9 porque se cuelga por overflow
CapturarNumFactura = 0

'si el numero es muy largo lo trato como barras
If Len(Trim(TxtCodigo.Text)) >= 44 Then
   CapturarNumFactura = CDbl(Mid(Trim(TxtCodigo.Text), 14, 9))
   Exit Function
End If

'si es mas corto lo trato como factura
If Len(Trim(TxtCodigo.Text)) >= 9 Then
   'tiene menos carcteres entonces puede ser un numero de factura
   CapturarNumFactura = CDbl(Mid(Trim(TxtCodigo.Text), 1, 9))
   Exit Function
Else
   'si es menor que 10 de largo todo es numero de factura
   CapturarNumFactura = CDbl(Trim(TxtCodigo.Text))
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion CapturarNumFactura"
End Function
Private Sub ActualizarVuelto()
Dim Importe As Currency
On Error GoTo merror

TxtVuelto.Text = 0

If Trim(TxtImporteActualizado.Text) = "" Then Exit Sub
   
If Not IsNumeric(TxtImporteActualizado.Text) Then Exit Sub
   
If CCur(TxtImporteActualizado.Text) <= 0 Then Exit Sub

If Trim(TxtImporteRecibido.Text) = "" Then Exit Sub
   
If Not IsNumeric(TxtImporteRecibido.Text) Then Exit Sub
   
If CCur(TxtImporteRecibido.Text) <= 0 Then Exit Sub
   
Importe = CCur(TxtImporteRecibido.Text)

If CCur(Importe) > CCur(TxtImporteActualizado.Text) Then
   TxtVuelto.Text = CCur(Importe) - CCur(TxtImporteActualizado.Text)
End If

TxtVuelto.Text = Format(CCur(TxtVuelto.Text), "0.00")

If CCur(TxtImporteRecibido.Text) >= CCur(TxtImporteActualizado.Text) Then
   TxtMensaje.Text = "COBRO TOTAL"
Else
   TxtMensaje.Text = "COBRO PARCIAL"
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento ActualizarVuelto"
End Sub
Private Sub CheckDescuento_Click()
'solo se ejecuta si aplicamos descuentos personalizados
On Error GoTo merror

'si habilito el descuento
If CheckDescuento.Value = 1 Then
   CheckRecargo.Value = 0
   TxtImporteRecargo.Text = 0
   TxtImporteRecargo.Enabled = False
   TxtImporteDescuento.Enabled = True
   TxtImporteDescuento.BackColor = vbWhite
Else
   'si saco el descuento
   TxtImporteDescuento.Text = 0
   TxtImporteDescuento.Enabled = False
   TxtImporteDescuento.BackColor = &HFFFFC0
End If

TxtImporteActualizado.Text = Format(ReconstruirImporte(), "0.00")
Call ActualizarVuelto
  
Exit Sub
merror:
tratarerrores "Error aplicando descuentos"
End Sub
Private Sub CheckRecargo_Click()
'solo tiene validez si aplicamos recargos
On Error GoTo merror

If CheckRecargo.Value = 1 Then
   CheckDescuento.Value = 0
   TxtImporteDescuento.Text = 0
   TxtImporteDescuento.Enabled = False
   TxtImporteRecargo.Enabled = True
   TxtImporteRecargo.BackColor = vbWhite
Else
   TxtImporteRecargo.Text = 0
   TxtImporteRecargo.Enabled = False
   TxtImporteRecargo.BackColor = &HFFFFC0
End If

TxtImporteActualizado.Text = Format(ReconstruirImporte(), "0.00")
Call ActualizarVuelto

Exit Sub
merror:
tratarerrores "Error aplicando recargos"
End Sub

Private Sub TxtImporteDescuento_Change()
On Error GoTo merror

TxtImporteActualizado.Text = Format(ReconstruirImporte(), "0.00")
Call ActualizarVuelto

Exit Sub
merror:
tratarerrores "Error cambiando el importe de descuento"
End Sub
Private Sub TxtImporteRecargo_Change()
On Error GoTo merror

TxtImporteActualizado.Text = Format(ReconstruirImporte(), "0.00")
Call ActualizarVuelto

Exit Sub
merror:
tratarerrores "Error cambiando el importe de recargo"
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
If CheckDescuento.Value = 1 Then
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
If CheckRecargo.Value = 1 Then
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

ImporteReal = CCur(Base) - CCur(Descuentos) + CCur(Recargos)

ReconstruirImporte = CCur(ImporteReal)

Exit Function
merror:
tratarerrores "Error en funcion ReconstruirImporte"
End Function
Private Sub TxtImporteRecibido_Change()
Call ActualizarVuelto
End Sub
Private Sub TxtImporteRecibido_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Call ActualizarVuelto
End If
End Sub
Private Sub TxtImporteRecibido_LostFocus()
Call ActualizarVuelto
End Sub
Private Sub Txtcodigo_Change()
On Error GoTo merror

'si cambia el numero de factura?
If Trim(TxtCodigo.Text) = "" Then Exit Sub

If Not IsNumeric(TxtCodigo.Text) Then Exit Sub

'uso val en vez de clng por overflow
If Val(TxtCodigo.Text) <= 0 Then Exit Sub

'si tiene 44 es un codigo de barras entonces busco
If Len(Trim(TxtCodigo.Text)) = 44 Then
   'esto es nuevo para los cobros diferidos
   Call BuscarFactura(1)
   Call SetearEntorno
   TxtCodigo.Text = ""
End If
'si tiene mas de 1 caracter y menos de 44 no hace nada
'solo espera que alguien ingrese un valor y apriete enter
'entonces se ejecutara el keydown

Exit Sub
merror:
tratarerrores "Error ingresando el codigo de barras/Nº de comprobante"
End Sub
Private Sub Txtcodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(TxtCodigo.Text) = "" Then Exit Sub

If Not IsNumeric(TxtCodigo.Text) Then Exit Sub

'uso val porque clng hace overflow
If Val(TxtCodigo.Text) <= 0 Then Exit Sub

If KeyCode = vbKeyReturn Then
   'esto es nuevo para los cobros diferidos
   Call BuscarFactura(1)
   Call SetearEntorno
   TxtCodigo.Text = ""
End If

End Sub
Private Sub SetearEntorno()
'restablece el estado de los botones
On Error GoTo merror

cmdAceptar.Enabled = False
CmdAnularCobro.Enabled = False
CheckDescuento.Enabled = True
CheckRecargo.Enabled = True
If Trim(TxtNumCuota.Text) = "0" Then Exit Sub

'si la cuota esta cobrada
If Trim(txtFechacobro.Text) <> "" Then
   'si el tipo de usuario anula cobros
   If VG_ANULA Then
      CmdAnularCobro.Enabled = True
   End If
   CheckDescuento.Enabled = False
   CheckRecargo.Enabled = False
Else
   cmdAceptar.Enabled = True
   If VG_APLICARRECIBOS Then
      LabelRecibo.Visible = True
      TxtNumRecibo.Visible = True
   End If
End If

Exit Sub
merror:
tratarerrores "Error seteando el entorno-CobrarCreditos"
End Sub
Private Sub CheckCobradores_Click()
On Error GoTo merror

If CheckCobradores.Value = 1 Then
   ComboCobradores.Enabled = True
Else
   ComboCobradores.ListIndex = -1
   ComboCobradores.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error seleccionando cobradores"
End Sub
