VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmOpciones 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   HelpContextID   =   33
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame9 
      Height          =   495
      Left            =   1560
      TabIndex        =   58
      Top             =   6720
      Width           =   6495
      Begin VB.Label Label11 
         Caption         =   "(*) Cada vez que realiza cambios debe pulsar el boton [Aplicar]"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   59
         Top             =   195
         Width           =   4575
      End
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   6480
      TabIndex        =   48
      ToolTipText     =   "Cierra la pantalla"
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   120
      TabIndex        =   47
      ToolTipText     =   "Graba las ultimas modificaciones"
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Frame FrameMutuo 
      Caption         =   "Texto de mutuo acuerdo:"
      ForeColor       =   &H00FF0000&
      Height          =   5895
      Left            =   240
      TabIndex        =   88
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox TxtParrafo6 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   95
         ToolTipText     =   "Parrafos del acuerdo mutuo"
         Top             =   3600
         Width           =   8295
      End
      Begin VB.TextBox TxtParrafo3 
         Height          =   285
         Left            =   1440
         TabIndex        =   92
         ToolTipText     =   "Parrafos del acuerdo mutuo"
         Top             =   1320
         Width           =   5295
      End
      Begin VB.TextBox TxtParrafo1 
         Height          =   285
         Left            =   120
         TabIndex        =   90
         ToolTipText     =   "Parrafos del acuerdo mutuo"
         Top             =   360
         Width           =   6735
      End
      Begin VB.CommandButton CmdPredeterminadoAcuerdo 
         Caption         =   "Predeterminado"
         Height          =   255
         Left            =   7080
         TabIndex        =   89
         ToolTipText     =   "Restablece el texto predeterminado del sistema"
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox TxtParrafo2 
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   91
         ToolTipText     =   "Parrafos del acuerdo mutuo"
         Top             =   720
         Width           =   8295
      End
      Begin VB.TextBox TxtParrafo5 
         Height          =   1275
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   94
         ToolTipText     =   "Parrafos del acuerdo mutuo"
         Top             =   2040
         Width           =   8295
      End
      Begin VB.TextBox TxtParrafo4 
         Height          =   285
         Left            =   120
         TabIndex        =   93
         ToolTipText     =   "Parrafos del acuerdo mutuo"
         Top             =   1680
         Width           =   5295
      End
      Begin VB.Label Label20 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   101
         Top             =   4680
         Width           =   6855
      End
      Begin VB.Label Label19 
         Caption         =   "[DOMICILIO DEUDOR]"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   100
         ToolTipText     =   "Este campo lo completa el sistema"
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "[VENCIMIENTO]"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5520
         TabIndex        =   99
         ToolTipText     =   "Este campo lo completa el sistema"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "[IMPORTE CUOTA]"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6840
         TabIndex        =   98
         ToolTipText     =   "Este campo lo completa el sistema"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "[Nº DE CUOTAS]"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   97
         ToolTipText     =   "Este campo lo completa el sistema"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "[IMPORTE TOTAL]"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6960
         TabIndex        =   96
         ToolTipText     =   "Este campo lo completa el sistema"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame FrameLibreDeuda 
      Caption         =   "Texto del certificado de libre deuda:"
      ForeColor       =   &H00FF0000&
      Height          =   5895
      Left            =   240
      TabIndex        =   49
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CommandButton CmdRestablecerLibreDeuda 
         Caption         =   "Predeterminado"
         Height          =   255
         Left            =   6360
         TabIndex        =   29
         ToolTipText     =   "Restablece el texto predeterminado del sistema"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox TxtTextoLibreDeuda 
         Height          =   1455
         Left            =   480
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   28
         ToolTipText     =   "Texto del libre deuda"
         Top             =   480
         Width           =   7455
      End
      Begin VB.Label Label17 
         Caption         =   "Aclaracion: El cliente y la fecha lo pone el sistema al momento de imprimir el libre deuda correspondiente."
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   480
         TabIndex        =   50
         Top             =   3720
         Width           =   7695
      End
   End
   Begin VB.Frame FrameImpuestos 
      Caption         =   "Impuestos:"
      ForeColor       =   &H00FF0000&
      Height          =   5895
      Left            =   240
      TabIndex        =   42
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CheckBox CheckImpuestosCuota2 
         Caption         =   "Aplicar impuestos desde la cuota 2"
         Height          =   255
         Left            =   240
         TabIndex        =   156
         Top             =   4680
         Width           =   3015
      End
      Begin VB.TextBox TxtPorcentajeIva 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   125
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox CheckImpuestosCuota1 
         Caption         =   "Aplicar todo el impuesto en la  primer cuota"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   86
         ToolTipText     =   "Aplica todo el impuesto en la primer cuota del credito"
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox TxtImporteImpuestos 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   82
         Tag             =   "N"
         ToolTipText     =   "Porcentaje del importe de la cuota"
         Top             =   1560
         Width           =   720
      End
      Begin MSComCtl2.UpDown UpDown13 
         Height          =   285
         Left            =   6840
         TabIndex        =   65
         Top             =   1200
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtImporteFijoImpuestos"
         BuddyDispid     =   196639
         OrigLeft        =   7080
         OrigTop         =   720
         OrigRight       =   7335
         OrigBottom      =   975
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.CheckBox CheckNoAplicarImpuestosRefinanciacion 
         Caption         =   "No aplicar impuestos a las refinanciaciones"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   25
         ToolTipText     =   "Si marca la casilla no se aplicaran impuestos a los planes refinanciados"
         Top             =   5040
         Width           =   3375
      End
      Begin VB.OptionButton OptionImporteImpuestos 
         Caption         =   "Importe fijo a dividir entre las cuotas del credito                                  $"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   23
         ToolTipText     =   "Divide un importe fijo entre las cuotas del credito"
         Top             =   1560
         Value           =   -1  'True
         Width           =   5295
      End
      Begin VB.CheckBox CheckAplicarImpuestos 
         Caption         =   "Aplicar impuestos"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Aplica impuestos a las cuotas de los nuevos creditos"
         Top             =   360
         Width           =   1575
      End
      Begin MSComCtl2.UpDown UpDown22 
         Height          =   285
         Left            =   6840
         TabIndex        =   83
         Top             =   1560
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtImporteImpuestos"
         BuddyDispid     =   196635
         OrigLeft        =   7080
         OrigTop         =   1440
         OrigRight       =   7335
         OrigBottom      =   1695
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.TextBox TxtImporteFijoImpuestos 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   24
         Tag             =   "N"
         ToolTipText     =   "Importe fijo a dividir entre las cuotas"
         Top             =   1200
         Width           =   705
      End
      Begin VB.OptionButton OptionImpuestosCredimaco 
         Caption         =   "Credimaco (IVA sobre interes, seguros, otorgamiento, gastos y mora)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         TabIndex        =   121
         ToolTipText     =   "Esta opcion calcula los impuestos personalizados para Credimaco"
         Top             =   840
         Width           =   5175
      End
      Begin VB.OptionButton OptionImpuestosFijos 
         Caption         =   "Importe fijo por cuota                                                                          $"
         Height          =   255
         Left            =   840
         TabIndex        =   155
         Top             =   1200
         Width           =   5295
      End
      Begin VB.Label Label46 
         Caption         =   "Porcentaje de IVA   %:"
         Height          =   255
         Left            =   240
         TabIndex        =   124
         Top             =   3840
         Width           =   1695
      End
   End
   Begin VB.Frame FrameGastos 
      Caption         =   "Gastos administrativos"
      ForeColor       =   &H00FF0000&
      Height          =   5895
      Left            =   240
      TabIndex        =   54
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox TxtCapInt 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5400
         MaxLength       =   7
         TabIndex        =   176
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox TxtFuncIntnocap 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5400
         MaxLength       =   7
         TabIndex        =   174
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TxtcapNoint 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5400
         MaxLength       =   7
         TabIndex        =   172
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton PorccapNoint 
         Caption         =   "Porcentaje en función del capital y no del interés:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         TabIndex        =   171
         Top             =   1800
         Width           =   4335
      End
      Begin VB.OptionButton PorcIntNoCap 
         Caption         =   "Porcentaje en función del interés y no del capital:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         TabIndex        =   170
         Top             =   2160
         Width           =   4335
      End
      Begin VB.OptionButton PorcCapint 
         Caption         =   "Porcentaje en función del capital + intereses:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         TabIndex        =   169
         Top             =   2520
         Width           =   4335
      End
      Begin VB.TextBox TxtImporteGastosBis 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   164
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxtImporteGastosFijoBis 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   120
         ToolTipText     =   "Importe que se agrega al 1º vto"
         Top             =   1080
         Width           =   840
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   285
         Left            =   6240
         TabIndex        =   119
         Top             =   1080
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtImporteGastosFijo"
         BuddyDispid     =   196652
         OrigLeft        =   7320
         OrigTop         =   3120
         OrigRight       =   7575
         OrigBottom      =   3375
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.TextBox TxtImporteGastosFijo 
         BackColor       =   &H80000013&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         MaxLength       =   7
         TabIndex        =   118
         ToolTipText     =   "Importe que se agrega al 1º vto"
         Top             =   1080
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.OptionButton OptionGastosFijos 
         Caption         =   "Importe fijo por cuota  (Credimaco)                               $:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         TabIndex        =   117
         Top             =   1080
         Width           =   4335
      End
      Begin VB.CheckBox CheckGastosCuota2 
         Caption         =   "Aplicar gastos desde la cuota 2 en adelante"
         Height          =   255
         Left            =   120
         TabIndex        =   116
         ToolTipText     =   "Carga los gastos desde la cuota Nº 2 en adelante"
         Top             =   4560
         Width           =   3495
      End
      Begin VB.CheckBox CheckGastosCuota1 
         Caption         =   "Aplicar todo el gasto en la primer cuota"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   4200
         Width           =   3135
      End
      Begin MSComCtl2.UpDown UpDown7 
         Height          =   285
         Left            =   6240
         TabIndex        =   64
         Top             =   1440
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtImporteGastos"
         BuddyDispid     =   196657
         OrigLeft        =   7560
         OrigTop         =   960
         OrigRight       =   7815
         OrigBottom      =   1215
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.CheckBox CheckNoAplicarGastosRefinanciacion 
         Caption         =   "No aplicar gastos a las refinanciaciones"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Si marca la casilla no se aplicaran gastos a los planes refinanciados"
         Top             =   4920
         Width           =   3135
      End
      Begin VB.TextBox TxtImporteGastos 
         BackColor       =   &H80000013&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   17
         Tag             =   "N"
         ToolTipText     =   "Importe fijo a dividir entre las cuotas de cada credito"
         Top             =   1440
         Width           =   735
      End
      Begin VB.OptionButton OptionImporteGastos 
         Caption         =   "Importe fijo a dividir entre las cuotas del credito           $:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         TabIndex        =   16
         ToolTipText     =   "Permite un importe fijo de gastos que se divide entre las cuotas"
         Top             =   1440
         Width           =   4335
      End
      Begin VB.CheckBox CheckAplicarGastos 
         Caption         =   "Aplicar gastos"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Aplica gastos administrativos a los creditos nuevos"
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   6240
         TabIndex        =   173
         Top             =   1800
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtcapNoint"
         BuddyDispid     =   196646
         OrigLeft        =   7560
         OrigTop         =   960
         OrigRight       =   7815
         OrigBottom      =   1215
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown8 
         Height          =   285
         Left            =   6255
         TabIndex        =   175
         Top             =   2160
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtFuncIntnocap"
         BuddyDispid     =   196645
         OrigLeft        =   7560
         OrigTop         =   960
         OrigRight       =   7815
         OrigBottom      =   1215
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown9 
         Height          =   285
         Left            =   6255
         TabIndex        =   177
         Top             =   2520
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtCapInt"
         BuddyDispid     =   196644
         OrigLeft        =   7560
         OrigTop         =   960
         OrigRight       =   7815
         OrigBottom      =   1215
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
   End
   Begin VB.Frame FrameSeguroVida 
      Caption         =   "Seguros:"
      ForeColor       =   &H00FF0000&
      Height          =   5895
      Left            =   240
      TabIndex        =   52
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox TxtSeguroFijo 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   154
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TxtAlicuotaSeguros 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   123
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox CheckSegurosCuota1 
         Caption         =   "Aplicar seguro solo a la primer cuota"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   85
         ToolTipText     =   "Aplica todo el seguro en la primer cuota del credito"
         Top             =   3360
         Width           =   2895
      End
      Begin VB.CheckBox CheckNoAplicarSegurosRefinanciacion 
         Caption         =   "No aplicar seguro a las refinanciaciones"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Si marca la casilla no se aplicara seguro a los planes refinanciados"
         Top             =   3840
         Width           =   3735
      End
      Begin VB.CheckBox CheckAplicarSeguro 
         Caption         =   "Aplicar seguros"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Aplica seguro a las facturas de los nuevos creditos"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtImporteSeguro 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   20
         Tag             =   "N"
         ToolTipText     =   "Importe fijo a dividir entre las cuotas del credito"
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label Label25 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   120
         TabIndex        =   168
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label Label24 
         Caption         =   "Importe fijo a dividir entre las cuotas del credito"
         Height          =   255
         Left            =   2040
         TabIndex        =   167
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label12 
         Caption         =   "Importe fijo de seguros por cuota"
         Height          =   255
         Left            =   2040
         TabIndex        =   166
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label10 
         Caption         =   "Calcular segun Credimaco (con alicuota)"
         Height          =   255
         Left            =   2040
         TabIndex        =   165
         Top             =   840
         Width           =   3855
      End
   End
   Begin VB.Frame FrameTasas 
      Caption         =   "Tasas de interes:"
      ForeColor       =   &H00FF0000&
      Height          =   5895
      Left            =   240
      TabIndex        =   53
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.Frame Frame15 
         Caption         =   "Recargo por refinanciacion:"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   240
         TabIndex        =   107
         Top             =   3600
         Width           =   8055
         Begin VB.CheckBox CheckAplicarTasaRefinanciacion 
            Caption         =   "Aplicar comision por Refinanciacion de deudas %:"
            Height          =   255
            Left            =   120
            TabIndex        =   109
            ToolTipText     =   "Al refinanciar deudas agrega un recargo sobre el monto adeudado"
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox TxtTasaRefinanciacion 
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4080
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   108
            Tag             =   "N"
            ToolTipText     =   "Porcentaje de recargo sobre el monto a refinanciar"
            Top             =   360
            Width           =   570
         End
         Begin MSComCtl2.UpDown UpDown6 
            Height          =   285
            Left            =   4680
            TabIndex        =   110
            Top             =   360
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "TxtTasaRefinanciacion"
            BuddyDispid     =   196674
            OrigLeft        =   4920
            OrigTop         =   840
            OrigRight       =   5175
            OrigBottom      =   1095
            Max             =   99
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   0   'False
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tasa por mora:"
         ForeColor       =   &H00FF0000&
         Height          =   2175
         Left            =   240
         TabIndex        =   60
         Top             =   1200
         Width           =   8055
         Begin VB.TextBox TxtTasaMora2 
            Height          =   285
            Left            =   2040
            MaxLength       =   5
            TabIndex        =   129
            Tag             =   "N"
            ToolTipText     =   "Tasa mensual de interes por mora"
            Top             =   960
            Width           =   585
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   285
            Left            =   2640
            TabIndex        =   63
            Top             =   360
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "TxtTasaMora"
            BuddyDispid     =   196677
            OrigLeft        =   3000
            OrigTop         =   360
            OrigRight       =   3255
            OrigBottom      =   615
            Max             =   999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox TxtTasaMora 
            Height          =   285
            Left            =   2040
            MaxLength       =   5
            TabIndex        =   62
            Tag             =   "N"
            ToolTipText     =   "Tasa de interes anual que se aplica a todas las facturas en mora"
            Top             =   360
            Width           =   855
         End
         Begin MSComCtl2.UpDown UpDown18 
            Height          =   285
            Left            =   2640
            TabIndex        =   131
            Top             =   960
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "TxtTasaMora2"
            BuddyDispid     =   196676
            OrigLeft        =   3000
            OrigTop         =   360
            OrigRight       =   3255
            OrigBottom      =   615
            Max             =   999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label55 
            Caption         =   "Esta tasa se usara para moras inferiores a 60 dias"
            Height          =   255
            Left            =   3000
            TabIndex        =   130
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label Label52 
            Caption         =   "Esta tasa se usara para las moras iguales superiores a 60 dias de atraso."
            Height          =   375
            Left            =   3000
            TabIndex        =   128
            Top             =   960
            Width           =   4935
         End
         Begin VB.Label Label51 
            Caption         =   "Tasa Mensual por mora %:"
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label38 
            Caption         =   "(*)Para Credimaco ademas se calculara el IVA sobre la mora."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   1800
            Width           =   4335
         End
         Begin VB.Label Label26 
            Caption         =   "(*)La tasa diaria se calcula dividiendo la tasa anual por 365 o dividiendo la tasa mensual por 30."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   1560
            Width           =   7815
         End
         Begin VB.Label Label13 
            Caption         =   "Tasa anual por mora      %:"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tasa de financiacion:"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   240
         TabIndex        =   67
         Top             =   360
         Width           =   8055
         Begin VB.Label Label2 
            Caption         =   "La tasa de financiacion depende de cada plan (ver pantalla [PLANES])"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   360
            Width           =   5535
         End
      End
   End
   Begin VB.Frame FrameCartaReclamo 
      Caption         =   "Texto de la carta reclamo:"
      ForeColor       =   &H00FF0000&
      Height          =   5895
      Left            =   240
      TabIndex        =   43
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CommandButton CmdRestablecerCarta 
         Caption         =   "Predeterminado"
         Height          =   255
         Left            =   6600
         TabIndex        =   27
         ToolTipText     =   "Restablece el texto predeterminado del sistema"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox TxtTextoCartaReclamo2 
         Height          =   1245
         Left            =   1680
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   26
         ToolTipText     =   "Texto de la carta reclamo"
         Top             =   1080
         Width           =   6375
      End
      Begin VB.TextBox TxtTextoCartaReclamo1 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   525
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   480
         Width           =   6375
      End
      Begin VB.Label Label21 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   360
         TabIndex        =   51
         Top             =   4080
         Width           =   7815
      End
      Begin VB.Label Label47 
         Caption         =   "Texto personalizado:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame FrameEmpresa 
      Caption         =   "Datos de nuestra empresa:"
      ForeColor       =   &H00FF0000&
      Height          =   6015
      Left            =   240
      TabIndex        =   30
      Top             =   840
      Width           =   8535
      Begin VB.TextBox TxtWebsite 
         Height          =   285
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   8
         ToolTipText     =   "Pagina web de nuestra empresa"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox TxtLugaresPago 
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   10
         ToolTipText     =   "Lugares habilitados para el pago de facturas"
         Top             =   3600
         Width           =   5535
      End
      Begin VB.TextBox TxtCiudad 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Ciudad donde reside nuestra empresa"
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox TxtEmpresa 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Nombre de nuestra empresa"
         Top             =   360
         Width           =   5535
      End
      Begin VB.TextBox TxtCuit 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "C.U.I.T de la empresa"
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox TxtHorarioAtencion 
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   9
         ToolTipText     =   "Horarios de atencion al publico"
         Top             =   3240
         Width           =   5535
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   7
         ToolTipText     =   "Correo electronico de nuestra empresa"
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox TxtTelefono 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   6
         ToolTipText     =   "Telefonos-Fax"
         Top             =   2520
         Width           =   5535
      End
      Begin VB.TextBox TxtDomicilio 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "Domicilio de nuestra empresa"
         Top             =   2160
         Width           =   5535
      End
      Begin VB.TextBox TxtIngresosBrutos 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Nº de ingresos brutos de nuestra empresa"
         Top             =   1800
         Width           =   5535
      End
      Begin VB.TextBox TxtIva 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Tipo de iva de la empresa"
         Top             =   1440
         Width           =   5535
      End
      Begin VB.Label Label71 
         Caption         =   "Website:"
         Height          =   255
         Left            =   4680
         TabIndex        =   55
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label49 
         Caption         =   "Lugares de pago:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label28 
         Caption         =   "Ciudad:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "CUIT:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Horario de atencion:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Telefono/Fax:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "E-mail:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Direccion:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Ingresos Brutos:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de IVA:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de la empresa:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   2085
         Left            =   3720
         Top             =   3890
         Width           =   2025
      End
   End
   Begin VB.Frame FrameImpresion 
      Caption         =   "Opciones de impresion:"
      ForeColor       =   &H00FF0000&
      Height          =   5895
      Left            =   240
      TabIndex        =   56
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.Frame Frame7 
         Caption         =   "Codigos RapiPago: "
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   240
         TabIndex        =   157
         Top             =   4560
         Width           =   8175
         Begin VB.TextBox TxtNumEmpresa 
            Height          =   285
            Left            =   3600
            MaxLength       =   3
            TabIndex        =   163
            ToolTipText     =   "Es el numero de nuestra empresa (ej:753)"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox TxtCodigoRapipago 
            Height          =   285
            Left            =   960
            MaxLength       =   4
            TabIndex        =   158
            ToolTipText     =   "Este numero es la extension de los archivos de rapipago (ej: 1032)"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label30 
            Caption         =   "Extension:"
            Height          =   255
            Left            =   120
            TabIndex        =   162
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label29 
            Caption         =   "Nº Empresa:"
            Height          =   255
            Left            =   2640
            TabIndex        =   159
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Modelo de Credimaco:"
         ForeColor       =   &H00FF0000&
         Height          =   2415
         Left            =   0
         TabIndex        =   136
         Top             =   360
         Width           =   8415
         Begin VB.CheckBox CheckImprimirMoraIva 
            Caption         =   "Imprimir Mora e IVA Mora"
            Height          =   255
            Left            =   3480
            TabIndex        =   152
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox TxtNumRecibo 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   6720
            MaxLength       =   9
            TabIndex        =   151
            ToolTipText     =   "Numero de inicio de la numeracion de las facturas"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox TxtNumCopias 
            Height          =   285
            Left            =   4440
            MaxLength       =   1
            TabIndex        =   149
            ToolTipText     =   "Nº de copias de la impresion de facturas"
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox CheckMostrarRecuadros 
            Caption         =   "Mostrar recuadros"
            Height          =   255
            Left            =   3480
            TabIndex        =   147
            ToolTipText     =   "Indica si en la factura se imprimiran los recuadros"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox TxtBotom 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   145
            Top             =   1440
            Width           =   600
         End
         Begin VB.TextBox TxtLeft 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   142
            Top             =   1080
            Width           =   600
         End
         Begin VB.TextBox TxtTop 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   139
            Top             =   720
            Width           =   600
         End
         Begin VB.CheckBox CheckModeloFactura1 
            Caption         =   "Modelo de factura de Credimaco"
            Height          =   255
            Left            =   120
            TabIndex        =   137
            ToolTipText     =   "Selecciona el 1º modelo de recibo"
            Top             =   360
            Width           =   2655
         End
         Begin MSComCtl2.UpDown UpDown4 
            Height          =   285
            Left            =   3000
            TabIndex        =   140
            Top             =   720
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "TxtTop"
            BuddyDispid     =   196729
            OrigLeft        =   4920
            OrigTop         =   2160
            OrigRight       =   5175
            OrigBottom      =   2415
            Increment       =   90
            Max             =   720
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown5 
            Height          =   285
            Left            =   3000
            TabIndex        =   143
            Top             =   1080
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "TxtLeft"
            BuddyDispid     =   196728
            OrigLeft        =   4920
            OrigTop         =   2520
            OrigRight       =   5175
            OrigBottom      =   2775
            Increment       =   90
            Max             =   720
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown17 
            Height          =   285
            Left            =   3000
            TabIndex        =   146
            Top             =   1440
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "TxtBotom"
            BuddyDispid     =   196727
            OrigLeft        =   4800
            OrigTop         =   1680
            OrigRight       =   5055
            OrigBottom      =   1935
            Increment       =   90
            Max             =   720
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label42 
            Caption         =   "(*) Los margenes serviran para acomodar la hoja A4 a la factura de Credimaco."
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1080
            TabIndex        =   153
            Top             =   1800
            Width           =   5055
         End
         Begin VB.Label Label16 
            Caption         =   "Ultimo Nº de factura:"
            Height          =   255
            Left            =   5160
            TabIndex        =   150
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label48 
            Caption         =   "Nº de copias:"
            Height          =   255
            Left            =   3480
            TabIndex        =   148
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label43 
            Caption         =   "Margen inferior:"
            Height          =   255
            Left            =   1080
            TabIndex        =   144
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label41 
            Caption         =   "Margen izquierdo:"
            Height          =   255
            Left            =   1080
            TabIndex        =   141
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label40 
            Caption         =   "Margen superior:"
            Height          =   255
            Left            =   1080
            TabIndex        =   138
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Modelo de cupon:"
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Left            =   120
         TabIndex        =   132
         Top             =   2880
         Width           =   8295
         Begin VB.CheckBox CheckModeloFactura4 
            Caption         =   "Modelo resumido con tres talones"
            Height          =   255
            Left            =   240
            TabIndex        =   135
            Top             =   1080
            Width           =   2775
         End
         Begin VB.CheckBox CheckModeloFactura3 
            Caption         =   "Modelo resumido con dos talones"
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Top             =   720
            Width           =   2775
         End
         Begin VB.CheckBox CheckModeloFactura2 
            Caption         =   "Modelo con dos talones y detalles"
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   360
            Width           =   2775
         End
      End
   End
   Begin VB.Frame FrameComprobantes 
      Caption         =   "Configuracion de opciones de vencimientos:"
      ForeColor       =   &H00FF0000&
      Height          =   5895
      Left            =   240
      TabIndex        =   40
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.Frame Frame3 
         Caption         =   "Segundo vencimiento de cuotas:"
         Height          =   2415
         Left            =   120
         TabIndex        =   71
         Top             =   1320
         Width           =   8295
         Begin VB.OptionButton OptionVencimiento2Mora 
            Caption         =   "Aplicar dias de mora con formula de Credimaco"
            Enabled         =   0   'False
            Height          =   255
            Left            =   720
            TabIndex        =   160
            Top             =   1560
            Width           =   4815
         End
         Begin VB.TextBox TxtVencimiento2ImporteBis 
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5520
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   87
            ToolTipText     =   "Importe que se agrega al 1º vto"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TxtVencimiento2Porcentaje 
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5520
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   76
            Tag             =   "N"
            ToolTipText     =   "Porcentaje fijo que se suma al monto del 1º vto(el % se calcula sobre el importe del 1º vto)"
            Top             =   1080
            Width           =   585
         End
         Begin VB.TextBox TxtVencimiento2Importe 
            BackColor       =   &H80000013&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5520
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   75
            Tag             =   "N"
            ToolTipText     =   "Importe fijo que se recargara al monto del 1º vto"
            Top             =   720
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.OptionButton OptionVencimiento22 
            Caption         =   "Recargar al 1º Vto un porcentaje personalizado                   %:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   720
            TabIndex        =   74
            ToolTipText     =   "Agrega un % sobre el importe del 1º vto"
            Top             =   1080
            Width           =   4695
         End
         Begin VB.CheckBox CheckAplicarSegundoVencimiento 
            Caption         =   "Habilitar segundo vencimiento de cuotas"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            ToolTipText     =   "Habilita que se utilice el segundo vencimiento de facturas"
            Top             =   360
            Width           =   3255
         End
         Begin MSComCtl2.UpDown UpDown16 
            Height          =   285
            Left            =   6120
            TabIndex        =   77
            Top             =   720
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "TxtVencimiento2Importe"
            BuddyDispid     =   196746
            OrigLeft        =   6240
            OrigTop         =   720
            OrigRight       =   6495
            OrigBottom      =   1005
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   0   'False
         End
         Begin MSComCtl2.UpDown UpDown19 
            Height          =   285
            Left            =   6120
            TabIndex        =   78
            Top             =   1080
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "TxtVencimiento2Porcentaje"
            BuddyDispid     =   196745
            OrigLeft        =   6240
            OrigTop         =   1080
            OrigRight       =   6495
            OrigBottom      =   1365
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   0   'False
         End
         Begin VB.OptionButton OptionVencimiento21 
            Caption         =   "Recargar al 1º Vto con un importe fijo                                    $:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   720
            TabIndex        =   73
            ToolTipText     =   "Le suma un importe fijo al monto del 1º vencimiento"
            Top             =   720
            Value           =   -1  'True
            Width           =   4695
         End
         Begin VB.Label Label102 
            Caption         =   "(*)  Los recargos anteriores se agregan al importe del 1º vencimiento."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   720
            TabIndex        =   79
            Top             =   2040
            Width           =   5295
         End
      End
      Begin VB.Frame Frame10 
         Height          =   615
         Left            =   120
         TabIndex        =   80
         Top             =   480
         Width           =   8295
         Begin VB.CheckBox CheckAplicarVencimientoSabados 
            Caption         =   "Aplicar vencimiento los dias sabados"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            ToolTipText     =   "Establece si las cuotas pueden vencer los dias sabados"
            Top             =   240
            Width           =   3015
         End
      End
   End
   Begin VB.Frame FrameRequisitos 
      Caption         =   "Requisitos generales de credito: "
      ForeColor       =   &H00FF0000&
      Height          =   5775
      Left            =   240
      TabIndex        =   39
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox txtedad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   189
         Top             =   4560
         Width           =   375
      End
      Begin VB.CheckBox CheckCobrosDiferidos 
         Caption         =   "Permitir cobrar cuotas diferidas"
         Height          =   255
         Left            =   360
         TabIndex        =   106
         ToolTipText     =   "Si permite cobrar cuotas con fechas anteriores a la fecha"
         Top             =   3360
         Width           =   2655
      End
      Begin VB.CheckBox CheckCreditosDiferidos 
         Caption         =   "Permitir registrar creditos diferidos"
         Height          =   255
         Left            =   360
         TabIndex        =   105
         ToolTipText     =   "Si permite registrar creditos anteriores a la fecha"
         Top             =   3000
         Width           =   2775
      End
      Begin VB.CheckBox CheckAplicarRecibo 
         Caption         =   "Habilitar el uso de recibos"
         Height          =   255
         Left            =   360
         TabIndex        =   104
         ToolTipText     =   "Pregunta si desea imprimir el recibo o no al momento de cobrar cuotas"
         Top             =   3720
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox CheckRedondear 
         Caption         =   "Redondear el importe de las cuotas"
         Height          =   255
         Left            =   360
         TabIndex        =   103
         ToolTipText     =   "Si marca la casilla se redondearan las cuotas de los nuevos creditos"
         Top             =   2640
         Width           =   3255
      End
      Begin VB.CheckBox CheckAplicarCobrosParciales 
         Caption         =   "Permitir cobros parciales de cuotas"
         Height          =   255
         Left            =   360
         TabIndex        =   70
         ToolTipText     =   "Si marca la casilla se podra cobrar cuotas en forma parcial"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.CheckBox CheckFinalizarAutomaticamente 
         Caption         =   "Finalizar creditos automaticamente al cobrar la ultima cuota o al refinanciarlo"
         Height          =   255
         Left            =   360
         TabIndex        =   69
         ToolTipText     =   "Al cobrar la ultima cuota o al refinanciarlo finaliza el credito"
         Top             =   1560
         Width           =   5655
      End
      Begin VB.CheckBox CheckPagarCuotasDesordenadas 
         Caption         =   "Permitir cobrar cuotas salteadas"
         Height          =   255
         Left            =   360
         TabIndex        =   66
         ToolTipText     =   "Si permite que se cobren cuotas teniendo impagas anteriores"
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CheckBox CheckClienteJudicial 
         Caption         =   "Otorgar creditos a clientes que tienen otros creditos bloqueados a su nombre"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         ToolTipText     =   "Permite clientes que son titulares de otros creditos en tramite legal"
         Top             =   840
         Width           =   5775
      End
      Begin VB.CheckBox CheckClienteSimultaneo 
         Caption         =   "Permitir que los clientes tengan creditos simultaneos"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         ToolTipText     =   "Permite que un cliente tenga mas de un credito simultaneo vigente"
         Top             =   480
         Width           =   4695
      End
      Begin VB.CheckBox CheckGarante 
         Caption         =   "Exigir un garante al otorgar nuevos creditos"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         ToolTipText     =   "Si marca la casilla se exigira un garante al otorgar un nuevo credito"
         Top             =   1200
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker FecLimite 
         Height          =   255
         Left            =   1920
         TabIndex        =   192
         ToolTipText     =   "Seleccione la fecha de fin de consulta"
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   50266113
         CurrentDate     =   42528
      End
      Begin VB.Label Label31 
         Caption         =   "Fecha limite ingreso:"
         Height          =   255
         Left            =   360
         TabIndex        =   191
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label23 
         Caption         =   "Edad maxima permitida para otorgar un credito:"
         Height          =   255
         Left            =   360
         TabIndex        =   190
         Top             =   4560
         Width           =   4575
      End
   End
   Begin VB.Frame FrameOtorgamiento 
      Caption         =   "Cargos de otorgamiento:"
      ForeColor       =   &H00FF0000&
      Height          =   6015
      Left            =   240
      TabIndex        =   111
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.OptionButton OptImpOtor 
         Caption         =   "Importe de otorgamiento $:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   188
         ToolTipText     =   "Permite un importe fijo de gastos que se divide entre las cuotas"
         Top             =   840
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.TextBox TxtCapmasInt 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6480
         MaxLength       =   7
         TabIndex        =   185
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TxtPorIntNoCap 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6480
         MaxLength       =   7
         TabIndex        =   183
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TxtPorCapNoInt 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6480
         MaxLength       =   7
         TabIndex        =   181
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton OptCapInt 
         Caption         =   "Porcentaje en función del capital + intereses:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   178
         ToolTipText     =   "Permite un importe fijo de gastos que se divide entre las cuotas"
         Top             =   1920
         Width           =   4335
      End
      Begin VB.CheckBox CheckNoAplicarOtRefin 
         Caption         =   "No incluir en las refinanciaciones"
         Height          =   255
         Left            =   240
         TabIndex        =   161
         Top             =   3960
         Width           =   2655
      End
      Begin VB.CheckBox CheckAplicarOtorgamiento 
         Caption         =   "Aplicar Cargos de Otorgamiento"
         Height          =   255
         Left            =   120
         TabIndex        =   114
         ToolTipText     =   "Si aplica gastos de otorgamiento"
         Top             =   360
         Width           =   2655
      End
      Begin VB.CheckBox CheckOtorgamiento1 
         Caption         =   "Cargar otorgamiento solo a la primer cuota"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   113
         ToolTipText     =   "Si el cargo de otorgamiento se carga solo a la primer cuota o a todas en partes iguales"
         Top             =   2760
         Width           =   3375
      End
      Begin MSComCtl2.UpDown UpDown11 
         Height          =   285
         Left            =   7215
         TabIndex        =   182
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtPorCapNoInt"
         BuddyDispid     =   196771
         OrigLeft        =   7560
         OrigTop         =   960
         OrigRight       =   7815
         OrigBottom      =   1215
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown12 
         Height          =   285
         Left            =   7215
         TabIndex        =   184
         Top             =   1560
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtPorIntNoCap"
         BuddyDispid     =   196770
         OrigLeft        =   7560
         OrigTop         =   960
         OrigRight       =   7815
         OrigBottom      =   1215
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown14 
         Height          =   285
         Left            =   7215
         TabIndex        =   186
         Top             =   1920
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtCapmasInt"
         BuddyDispid     =   196769
         OrigLeft        =   7560
         OrigTop         =   960
         OrigRight       =   7815
         OrigBottom      =   1215
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown15 
         Height          =   285
         Left            =   6136
         TabIndex        =   187
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtImporteOtorgamiento"
         BuddyDispid     =   196777
         OrigLeft        =   7560
         OrigTop         =   960
         OrigRight       =   7815
         OrigBottom      =   1215
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.OptionButton OptIntNoCap 
         Caption         =   "Porcentaje en función del interés y no del capital:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   179
         ToolTipText     =   "Permite un importe fijo de gastos que se divide entre las cuotas"
         Top             =   1560
         Width           =   4335
      End
      Begin VB.TextBox TxtImporteOtorgamiento 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         MaxLength       =   7
         TabIndex        =   112
         ToolTipText     =   "Importe de gastos de otorgamiento"
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton OptCapNoint 
         Caption         =   "Porcentaje en función del capital y no del interés:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   180
         ToolTipText     =   "Permite un importe fijo de gastos que se divide entre las cuotas"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label27 
         Caption         =   "(*)Si no lo carga a la primer cuota, el importe se dividira entre las cuotas en partes iguales."
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   115
         Top             =   3120
         Width           =   6135
      End
      Begin VB.Label Label33 
         Caption         =   "(Por ejemplo:20,66)"
         Height          =   255
         Left            =   6960
         TabIndex        =   126
         Top             =   840
         Width           =   1455
      End
   End
   Begin MSComctlLib.TabStrip TabStripOpciones 
      Height          =   7215
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   12726
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   12
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Datos de la empresa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tasas de interes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisitos generales de Credito"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vencimientos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Texto de carta reclamo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Texto de libre deuda"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gastos administrativos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seguros"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Impuestos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Impresion"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Mutuo acuerdo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gtos.Otorgamiento"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label57 
      Caption         =   "(*) Cada vez que realiza cambios debe pulsar el boton Aplicar"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   57
      Top             =   6360
      Width           =   4575
   End
End
Attribute VB_Name = "FrmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ACA SE CONFIGURAN PARAMETROS GENERALES DEL SISTEMA A TRAVES
'DE VARIABLES GLOBALES que se usan en todo el sistema
'Esas variables se declaran en el modulo MAIN

Private Sub Form_Load()
On Error GoTo merror

Call CargarOpciones

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de Opciones"
End Sub
Private Sub CmdCerrar_Click()
Unload Me
End Sub
Private Sub CargarOpciones()
'cargo las opciones del sistema en las variables globales
On Error GoTo merror

Call RefrescarOpcionesSistema

'datos de la empresa
TxtEmpresa.Text = VG_EMPRESA
TxtCiudad.Text = VG_CIUDAD
TxtCuit.Text = VG_CUIT
TxtIngresosBrutos.Text = VG_INGRESOSBRUTOS
TxtIva.Text = VG_IVA
TxtTelefono.Text = VG_TELEFONO
TxtDomicilio.Text = VG_DOMICILIO
TxtEmail.Text = VG_EMAIL
TxtWebsite.Text = VG_WEBSITE
TxtHorarioAtencion.Text = VG_HORARIOATENCION
TxtLugaresPago.Text = VG_LUGARESPAGO

'libre deuda
TxtTextoLibreDeuda.Text = VG_TEXTOLIBREDEUDA1 + VG_TEXTOLIBREDEUDA2

'carta reclamo
TxtTextoCartaReclamo2.Text = VG_TEXTOCARTARECLAMO1 + VG_TEXTOCARTARECLAMO2

'acuerdo mutuo
TxtParrafo1.Text = VG_TEXTOACUERDOMUTUO1 & vbNullString
TxtParrafo2.Text = VG_TEXTOACUERDOMUTUO2 & vbNullString
TxtParrafo3.Text = VG_TEXTOACUERDOMUTUO3 & vbNullString
TxtParrafo4.Text = VG_TEXTOACUERDOMUTUO4 & vbNullString
TxtParrafo5.Text = VG_TEXTOACUERDOMUTUO5 & VG_TEXTOACUERDOMUTUO6 & VG_TEXTOACUERDOMUTUO7 & vbNullString
TxtParrafo6.Text = VG_TEXTOACUERDOMUTUO8 & VG_TEXTOACUERDOMUTUO9 & VG_TEXTOACUERDOMUTUO10 & vbNullString

'tasas
TxtTasaMora.Text = CDbl(VG_TASAMORA)

'recargo por refinanciacion
TxtTasaRefinanciacion.Text = CDbl(VG_TASAREFINANCIACION)
If VG_APLICARTASAREFINANCIACION Then
   CheckAplicarTasaRefinanciacion.Value = 1
Else
   CheckAplicarTasaRefinanciacion.Value = 0
End If

'gastos administrativos
TxtImporteGastos.Text = CCur(VG_IMPORTEGASTOS)
TxtImporteGastosFijo.Text = CCur(VG_IMPORTEGASTOSFIJOS)
TxtcapNoint.Text = VG_PORCCAPNOINT
TxtCapInt.Text = VG_PORCCAPINT
TxtFuncIntnocap.Text = VG_PORCFUNNOCAP




If VG_APLICARGASTOS Then
   'dentro actualiza los checks derivados
   CheckAplicarGastos.Value = 1
Else
   CheckAplicarGastos.Value = 0
End If

'si no aplico gastos a las refinanciaciones
If VG_NOAPLICARGASTOSREFINANCIACION Then
   CheckNoAplicarGastosRefinanciacion.Value = 1
Else
   CheckNoAplicarGastosRefinanciacion.Value = 0
End If
   
'si aplico todo el gasto en la primer cuota
If VG_APLICARGASTOSCUOTA1 Then
   CheckGastosCuota1.Value = 1
Else
   CheckGastosCuota1.Value = 0
End If

'si aplico todo el gasto en la primer cuota
If VG_APLICARGASTOSCUOTA2 Then
   CheckGastosCuota2.Value = 1
Else
   CheckGastosCuota2.Value = 0
End If

'seguros
TxtImporteSeguro.Text = CCur(VG_IMPORTESEGURO)
TxtAlicuotaSeguros.Text = CDbl(VG_ALICUOTASEGUROS)
TxtSeguroFijo.Text = CCur(VG_SEGUROFIJO)

If VG_APLICARSEGURO Then
   CheckAplicarSeguro.Value = 1
Else
   CheckAplicarSeguro.Value = 0
End If

'si no aplico seguros a las refinanciaciones
If VG_NOAPLICARSEGUROSREFINANCIACION Then
   CheckNoAplicarSegurosRefinanciacion.Value = 1
Else
   CheckNoAplicarSegurosRefinanciacion.Value = 0
End If

'si aplico todo el seguro en la primer cuota
If VG_APLICARSEGUROSCUOTA1 Then
   CheckSegurosCuota1.Value = 1
Else
   CheckSegurosCuota1.Value = 0
End If

'impuestos
TxtImporteImpuestos.Text = CCur(VG_IMPORTEIMPUESTOS)
TxtImporteFijoImpuestos.Text = CCur(VG_IMPUESTOSFIJOS)
TxtPorcentajeIva = VG_PORCENTAJEIVA

If VG_APLICARIMPUESTOS Then
   CheckAplicarImpuestos.Value = 1
Else
   CheckAplicarImpuestos.Value = 0
End If

'si no aplico impuestos a las refinanciaciones
If VG_NOAPLICARIMPUESTOSREFINANCIACION Then
   CheckNoAplicarImpuestosRefinanciacion.Value = 1
Else
   CheckNoAplicarImpuestosRefinanciacion.Value = 0
End If

'si aplico todo el impuesto en la primer cuota
If VG_APLICARIMPUESTOSCUOTA1 Then
   CheckImpuestosCuota1.Value = 1
Else
   CheckImpuestosCuota1.Value = 0
End If

'si aplico todo el impuesto en la primer cuota
If VG_APLICARIMPUESTOSCUOTA2 Then
   CheckImpuestosCuota2.Value = 1
Else
   CheckImpuestosCuota2.Value = 0
End If

'requisitos generales
If VG_CLIENTESIMULTANEO Then
   CheckClienteSimultaneo.Value = 1
Else
   CheckClienteSimultaneo.Value = 0
End If
   
If VG_GARANTE Then
   CheckGarante.Value = 1
Else
   CheckGarante.Value = 0
End If

If VG_CLIENTEJUDICIAL Then
   CheckClienteJudicial.Value = 1
Else
   CheckClienteJudicial.Value = 0
End If


txtedad.Text = VG_EDAD
FecLimite.Value = VG_FECHALIMITEINGRESO


'cuotas salteadas
'el chequeo de cuotas salteadas solo se aplica en la
'pantalla de cobros individuales (no esta en la de cobros multiples)
If VG_PAGARCUOTASDESORDENADAS Then
   CheckPagarCuotasDesordenadas.Value = 1
Else
   CheckPagarCuotasDesordenadas.Value = 0
End If
  
'finalizacion automatica
If VG_FINALIZARAUTOMATICAMENTE Then
   CheckFinalizarAutomaticamente.Value = 1
Else
   CheckFinalizarAutomaticamente.Value = 0
End If

'si acepto cobros parciales de cuotas
If VG_APLICARCOBROSPARCIALES Then
   CheckAplicarCobrosParciales.Value = 1
Else
   CheckAplicarCobrosParciales.Value = 0
End If

'redondeos
If VG_REDONDEAR Then
   CheckRedondear.Value = 1
Else
   CheckRedondear.Value = 0
End If

'si confirmo el recibo al momento del cobro
If VG_APLICARRECIBOS Then
   CheckAplicarRecibo.Value = 1
Else
   CheckAplicarRecibo.Value = 0
End If

'si premite registrar creditos diferidos
If VG_CREDITOSDIFERIDOS Then
   CheckCreditosDiferidos.Value = 1
Else
   CheckCreditosDiferidos.Value = 0
End If

'si permite cobrar cuotas diferidas
If VG_COBROSDIFERIDOS Then
   CheckCobrosDiferidos.Value = 1
Else
   CheckCobrosDiferidos.Value = 0
End If

'si puede haber vencimiento los sabados
If VG_APLICARVENCIMIENTOSABADOS Then
   CheckAplicarVencimientoSabados.Value = 1
Else
   CheckAplicarVencimientoSabados.Value = 0
End If

'segundo vencimiento de facturas
TxtVencimiento2Importe.Text = CCur(VG_VENCIMIENTO2IMPORTE)
TxtVencimiento2Porcentaje.Text = CDbl(VG_VENCIMIENTO2PORCENTAJE)

If VG_APLICARSEGUNDOVENCIMIENTO Then
   CheckAplicarSegundoVencimiento.Value = 1
Else
   CheckAplicarSegundoVencimiento.Value = 0
End If

If VG_MODELOFACTURA1 Then
   CheckModeloFactura1.Value = 1
End If
If VG_MODELOFACTURA2 Then
   CheckModeloFactura2.Value = 1
End If
If VG_MODELOFACTURA3 Then
   CheckModeloFactura3.Value = 1
End If
If VG_MODELOFACTURA4 Then
   CheckModeloFactura4.Value = 1
End If

TxtNumRecibo.Text = VG_ULTIMONUMRECIBO

TxtNumCopias.Text = VG_NUMCOPIAS

'Gastos de otorgamiento
TxtImporteOtorgamiento.Text = CCur(VG_IMPORTEOTORGAMIENTO)
TxtPorCapNoInt.Text = VG_OTORCAPNOINT
TxtPorIntNoCap.Text = VG_OTORINTNOCAP
TxtCapmasInt.Text = VG_OTORCAPMASINT

If VG_APLICAROTORGAMIENTO Then
   CheckAplicarOtorgamiento.Value = 1
Else
   CheckAplicarOtorgamiento.Value = 0
End If




If VG_APLICAROTORGAMIENTOCUOTA1 Then
   CheckOtorgamiento1.Value = 1
Else
   CheckOtorgamiento1.Value = 0
End If

'ubicacion del reporte
TxtTop.Text = VG_TOP
TxtLeft.Text = VG_LEFT
TxtBotom.Text = VG_BOTOM

If VG_MOSTRARRECUADROS Then
   CheckMostrarRecuadros.Value = 1
Else
   CheckMostrarRecuadros.Value = 0
End If

If VG_IMPRIMIRMORAIVA Then
   CheckImprimirMoraIva.Value = 1
Else
   CheckImprimirMoraIva.Value = 0
End If

TxtCodigoRapipago.Text = VG_CODIGOAUTOMATICO

'nueva tasa por mora de acuerdo a los dias de mora
TxtTasaMora2.Text = VG_TASAMORA2

If VG_NOAPLICAROTREFIN Then
   CheckNoAplicarOtRefin.Value = 1
Else
   CheckNoAplicarOtRefin.Value = 0
End If

'numero de empresa ante rapipago
TxtNumEmpresa.Text = VG_NUMEMPRESA

Exit Sub
merror:
tratarerrores "Error cargando opciones del sistema"
End Sub
Private Function DatosEmpresaOk() As Boolean
'verifico datos de empresa
On Error GoTo merror

DatosEmpresaOk = True

If Trim(TxtEmpresa.Text) = "" Then
   DatosEmpresaOk = False
   MsgE "Debe ingresar el nombre de la empresa"
   TxtEmpresa.SetFocus
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosEmpresaOk-Opciones"
End Function
Private Function DatosRequisitosOk() As Boolean
On Error GoTo merror

DatosRequisitosOk = True
  
If Trim(txtedad.Text) = "" Then
   DatosRequisitosOk = False
   MsgE "Ingrese edad maxima para otorgar un credito"
   txtedad.SetFocus
   Exit Function
End If
  
Exit Function
merror:
tratarerrores "Error en funcion DatosRequisitosOk"
End Function
Private Function DatosTasasOk() As Boolean
'verifico datos de las tasas de interes por mora
On Error GoTo merror

DatosTasasOk = True

'por ahora no hay chequeo porque tiene updown
If Trim(TxtTasaMora.Text) = "" Then
   DatosTasasOk = False
   MsgE "Debe ingresar la tasa por mora"
   TxtTasaMora.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtTasaMora.Text) Then
   DatosTasasOk = False
   MsgE "La tasa por mora debe ser numerica"
   TxtTasaMora.SetFocus
   Exit Function
End If
If CDbl(TxtTasaMora.Text) < 0 Then
   DatosTasasOk = False
   MsgE "La tasa por mora debe ser mayor a cero"
   TxtTasaMora.SetFocus
   Exit Function
End If

'valido la segunda tasa por mora
If Trim(TxtTasaMora2.Text) = "" Then
   DatosTasasOk = False
   MsgE "Debe ingresar la tasa por mora mensual"
   TxtTasaMora2.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtTasaMora2.Text) Then
   DatosTasasOk = False
   MsgE "La tasa por mora mensual debe ser numerica"
   TxtTasaMora2.SetFocus
   Exit Function
End If
If CDbl(TxtTasaMora2.Text) < 0 Then
   DatosTasasOk = False
   MsgE "La tasa por mora mensual debe ser mayor a cero"
   TxtTasaMora2.SetFocus
   Exit Function
End If

If CheckAplicarTasaRefinanciacion.Value = 0 Then
   TxtTasaRefinanciacion.Text = 0
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosTasasOk-Opciones"
End Function
Private Function DatosGastosOk() As Boolean
'verifico datos de los gastos
On Error GoTo merror

DatosGastosOk = True


If Not OptionImporteGastos.Value Then
   TxtImporteGastos.Text = 0
End If

If Not OptionGastosFijos.Value Then
   TxtImporteGastosFijo.Text = 0
End If

 If Not PorccapNoint.Value Then
    TxtcapNoint.Text = 0
 End If

If Not PorcIntNoCap.Value Then
    TxtFuncIntnocap.Text = 0
End If

If Not PorcCapint.Value Then
    TxtCapInt.Text = 0
End If

'valido el importe fijo
If CheckAplicarGastos.Value = 1 Then
   If OptionGastosFijos.Value Then
      If Trim(TxtImporteGastosFijo.Text) = "" Then
         DatosGastosOk = False
         MsgE "Debe ingresar un importe"
         Exit Function
      End If
      If Not IsNumeric(TxtImporteGastosFijo.Text) Then
         DatosGastosOk = False
         MsgE "El importe debe ser numerico"
         TxtImporteGastosFijo.Text = 0
         Exit Function
      End If
      If CCur(TxtImporteGastosFijo.Text) < 0 Then
         DatosGastosOk = False
         Exit Function
      End If
   End If
   If OptionImporteGastos.Value Then
      If Trim(TxtImporteGastosBis.Text) = "" Then
         DatosGastosOk = False
         MsgE "Debe ingresar un importe"
         Exit Function
      End If
      If Not IsNumeric(TxtImporteGastosBis.Text) Then
         DatosGastosOk = False
         MsgE "El importe debe ser numerico"
         TxtImporteGastosBis.Text = 0
         Exit Function
      End If
      If CCur(TxtImporteGastosBis.Text) < 0 Then
         DatosGastosOk = False
         Exit Function
      End If
    End If
    
     If PorccapNoint.Value Then
      If Trim(TxtcapNoint.Text) = "" Then
         DatosGastosOk = False
         MsgE "Debe ingresar un Porcentaje"
         Exit Function
      End If
      If Not IsNumeric(TxtcapNoint.Text) Then
         DatosGastosOk = False
         MsgE "El Porcentaje debe ser numerico"
         TxtcapNoint.Text = 0
         Exit Function
      End If
      If CCur(TxtcapNoint.Text) < 0 Then
         DatosGastosOk = False
         Exit Function
      End If
    End If
    
     If PorcIntNoCap.Value Then
      If Trim(TxtFuncIntnocap.Text) = "" Then
         DatosGastosOk = False
         MsgE "Debe ingresar un Porcentaje"
         Exit Function
      End If
      If Not IsNumeric(TxtFuncIntnocap.Text) Then
         DatosGastosOk = False
         MsgE "El Porcentaje debe ser numerico"
         TxtFuncIntnocap.Text = 0
         Exit Function
      End If
      If CCur(TxtFuncIntnocap.Text) < 0 Then
         DatosGastosOk = False
         Exit Function
      End If
    End If
    
    If PorcCapint.Value Then
      If Trim(TxtCapInt.Text) = "" Then
         DatosGastosOk = False
         MsgE "Debe ingresar un Porcentaje"
         Exit Function
      End If
      If Not IsNumeric(TxtCapInt.Text) Then
         DatosGastosOk = False
         MsgE "El Porcentaje debe ser numerico"
         TxtCapInt.Text = 0
         Exit Function
      End If
      If CCur(TxtCapInt.Text) < 0 Then
         DatosGastosOk = False
         Exit Function
      End If
    End If
    
Else
   TxtImporteGastosFijo.Text = 0
   TxtImporteGastos.Text = 0
   TxtcapNoint = 0
   TxtFuncIntnocap = 0
   TxtCapInt = 0
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosGastosOk-Opciones"
End Function
Private Function DatosSegurosOk() As Boolean
'verifico datos de los seguros
On Error GoTo merror

DatosSegurosOk = True

'si no aplico seguros
If CheckAplicarSeguro.Value = 0 Then
   TxtAlicuotaSeguros.Text = 0
   TxtSeguroFijo.Text = 0
   TxtImporteSeguro.Text = 0
Else
   If Trim(TxtAlicuotaSeguros.Text) = "" Then
      TxtAlicuotaSeguros.Text = 0
   End If
   If Not IsNumeric(TxtAlicuotaSeguros.Text) Then
      TxtAlicuotaSeguros.Text = 0
   End If
   If CDbl(TxtAlicuotaSeguros.Text) < 0 Then
      TxtAlicuotaSeguros.Text = 0
   End If

   If Trim(TxtImporteSeguro.Text) = "" Then
      TxtImporteSeguro.Text = 0
   End If
   If Not IsNumeric(TxtImporteSeguro.Text) Then
      TxtImporteSeguro.Text = 0
   End If
   If CCur(TxtImporteSeguro.Text) < 0 Then
      TxtImporteSeguro.Text = 0
   End If

   If Trim(TxtSeguroFijo.Text) = "" Then
      TxtSeguroFijo.Text = 0
   End If
   If Not IsNumeric(TxtSeguroFijo.Text) Then
      TxtSeguroFijo.Text = 0
   End If
   If CCur(TxtSeguroFijo.Text) < 0 Then
      TxtSeguroFijo.Text = 0
   End If
   
   If CDbl(TxtAlicuotaSeguros.Text) = 0 And CCur(TxtImporteSeguro.Text) = 0 And CCur(TxtSeguroFijo.Text) = 0 Then
      DatosSegurosOk = False
      MsgE "Debe ingresar el valor de alguno de los seguros"
      Exit Function
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosSegurosOk-Opciones"
End Function
Private Function DatosImpuestosOk() As Boolean
On Error GoTo merror

DatosImpuestosOk = True

If Not OptionImpuestosFijos.Value Then
   TxtImporteFijoImpuestos.Text = 0
End If

If Not OptionImporteImpuestos.Value Then
   TxtImporteImpuestos.Text = 0
End If

If CheckAplicarImpuestos.Value = 0 Then
   TxtPorcentajeIva.Text = 0
Else
   If Trim(TxtPorcentajeIva.Text) = "" Then
      DatosImpuestosOk = False
      MsgE "Debe ingresar el porcentaje de IVA"
      TxtPorcentajeIva.SetFocus
      Exit Function
   End If
   If Not IsNumeric(TxtPorcentajeIva.Text) Then
      DatosImpuestosOk = False
      MsgE "El porcentaje de IVA debe ser numerico"
      TxtPorcentajeIva.SetFocus
      Exit Function
   End If
   If CDbl(TxtPorcentajeIva.Text) < 0 Then
      DatosImpuestosOk = False
      MsgE "El porcentaje de IVA debe ser mayor a cero"
      TxtPorcentajeIva.SetFocus
      Exit Function
   End If
   
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosImpuestosOk-Opciones"
End Function
Private Function DatosCartaReclamoOk() As Boolean
On Error GoTo merror

DatosCartaReclamoOk = True

If Trim(TxtTextoCartaReclamo2.Text) = "" Then
   DatosCartaReclamoOk = False
   MsgE "Debe ingresar el texto de la carta reclamo"
   TxtTextoCartaReclamo2.SetFocus
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosCartaReclamoOk-Opciones"
End Function
Private Function DatosLibreDeudaOk() As Boolean
On Error GoTo merror

DatosLibreDeudaOk = True

If Trim(TxtTextoLibreDeuda.Text) = "" Then
   DatosLibreDeudaOk = False
   MsgE "Debe ingresar el texto del libre deuda"
   TxtTextoLibreDeuda.SetFocus
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosLibreDeudaOk-Opciones"
End Function
Private Function DatosImpresionOk() As Boolean
On Error GoTo merror

DatosImpresionOk = True

If CheckModeloFactura1.Value = 0 Then
   DatosImpresionOk = False
   MsgE "Debe seleccionar el modelo de factura de credimaco"
   Exit Function
End If

If CheckModeloFactura2.Value = 0 And CheckModeloFactura3.Value = 0 And CheckModeloFactura4.Value = 0 Then
   DatosImpresionOk = False
   MsgE "Debe seleccionar algun modelo de cupon"
   Exit Function
End If

If Trim(TxtNumCopias.Text) = "" Then
   DatosImpresionOk = False
   MsgE "Debe ingresar la cantidad de copias"
   TxtNumCopias.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtNumCopias.Text) Then
   DatosImpresionOk = False
   MsgE "La cantidad de copias debe ser numerica"
   TxtNumCopias.SetFocus
   Exit Function
End If
If CLng(TxtNumCopias.Text) <= 0 Then
   DatosImpresionOk = False
   MsgE "La cantidad de copias debe ser mayor a cero"
   TxtNumCopias.SetFocus
   Exit Function
End If

If Trim(TxtNumRecibo.Text) = "" Then
   DatosImpresionOk = False
   MsgE "Debe ingresar el numero de factura"
   TxtNumRecibo.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtNumRecibo.Text) Then
   DatosImpresionOk = False
   MsgE "El numero de factura debe ser numerico"
   TxtNumRecibo.SetFocus
   Exit Function
End If
If CLng(TxtNumRecibo.Text) <= 0 Then
   DatosImpresionOk = False
   MsgE "El numero de recibo debe ser mayor a cero"
   TxtNumRecibo.SetFocus
   Exit Function
End If

'codigo de rapipago
If Trim(TxtCodigoRapipago.Text) = "" Then
   DatosImpresionOk = False
   MsgE "Debe ingresar el numero de extension de RapiPago"
   TxtCodigoRapipago.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtCodigoRapipago.Text) Then
   DatosImpresionOk = False
   MsgE "El numero de extension de rapipago debe ser numerico"
   TxtCodigoRapipago.SetFocus
   Exit Function
End If
If CLng(TxtCodigoRapipago.Text) <= 0 Then
   DatosImpresionOk = False
   MsgE "El numero de extension de rapipago debe ser mayor a cero"
   TxtCodigoRapipago.SetFocus
   Exit Function
End If

'numero de empresa rapipago
If Trim(TxtNumEmpresa.Text) = "" Then
   DatosImpresionOk = False
   MsgE "Debe ingresar el numero de empresa RapiPago"
   TxtNumEmpresa.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtNumEmpresa.Text) Then
   DatosImpresionOk = False
   MsgE "El numero de empresa rapipago debe ser numerico"
   TxtNumEmpresa.SetFocus
   Exit Function
End If
If CLng(TxtNumEmpresa.Text) <= 0 Then
   DatosImpresionOk = False
   MsgE "El numero de empresa de rapipago debe ser mayor a cero"
   TxtNumEmpresa.SetFocus
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosImpresionOk-Opciones"
End Function
Private Function DatosComprobantesOk() As Boolean
'datos de vencimientos
On Error GoTo merror

DatosComprobantesOk = True

If Not OptionVencimiento21.Value Then
   TxtVencimiento2Importe.Text = 0
End If

If Not OptionVencimiento22.Value Then
   TxtVencimiento2Porcentaje.Text = 0
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosComprobantesOk-Opciones"
End Function
Private Function DatosAcuerdoMutuoOk() As Boolean
On Error GoTo merror

DatosAcuerdoMutuoOk = True

If Trim(TxtParrafo1.Text) = "" Then
   DatosAcuerdoMutuoOk = False
   MsgE "Debe completar el texto del acuerdo mutuo"
   Exit Function
End If
If Trim(TxtParrafo2.Text) = "" Then
   DatosAcuerdoMutuoOk = False
   MsgE "Debe completar el texto del acuerdo mutuo"
   Exit Function
End If
If Trim(TxtParrafo3.Text) = "" Then
   DatosAcuerdoMutuoOk = False
   MsgE "Debe completar el texto del acuerdo mutuo"
   Exit Function
End If
If Trim(TxtParrafo4.Text) = "" Then
   DatosAcuerdoMutuoOk = False
   MsgE "Debe completar el texto del acuerdo mutuo"
   Exit Function
End If
If Trim(TxtParrafo5.Text) = "" Then
   DatosAcuerdoMutuoOk = False
   MsgE "Debe completar el texto del acuerdo mutuo"
   Exit Function
End If
If Trim(TxtParrafo6.Text) = "" Then
   DatosAcuerdoMutuoOk = False
   MsgE "Debe completar el texto del acuerdo mutuo"
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosAcuerdoMutuoOk"
End Function
Private Function DatosOtorgamientoOk() As Boolean
On Error GoTo merror

DatosOtorgamientoOk = True

If CheckAplicarOtorgamiento.Value = 1 Then
   If Trim(TxtImporteOtorgamiento.Text) = "" Then
      DatosOtorgamientoOk = False
      MsgE "Debe ingresar el importe de otorgamiento"
      TxtImporteOtorgamiento.SetFocus
      Exit Function
   End If
   If Not IsNumeric(TxtImporteOtorgamiento.Text) Then
      DatosOtorgamientoOk = False
      MsgE "El Porcentaje debe ser numerico"
      TxtImporteOtorgamiento.SetFocus
      Exit Function
   End If
   
   
   If Not IsNumeric(TxtPorCapNoInt.Text) Then
      DatosOtorgamientoOk = False
      MsgE "El Porcentaje debe ser numerico"
      TxtPorCapNoInt.Text = 0
      TxtPorCapNoInt.SetFocus
      Exit Function
   End If
   
   If Not IsNumeric(TxtPorIntNoCap.Text) Then
      DatosOtorgamientoOk = False
      MsgE "El Porcentaje debe ser numerico"
      TxtPorIntNoCap.Text = 0
      TxtPorIntNoCap.SetFocus
      Exit Function
   End If
   
   If Not IsNumeric(TxtCapmasInt.Text) Then
      DatosOtorgamientoOk = False
      MsgE "El Porcentaje debe ser numerico"
      TxtCapmasInt.Text = 0
      TxtCapmasInt.SetFocus
      Exit Function
   End If
   
   'If CCur(TxtImporteOtorgamiento.Text) <= 0 Then
    '  DatosOtorgamientoOk = False
     ' MsgE "El importe de otorgamiento debe ser mayor a cero"
      'TxtImporteOtorgamiento.SetFocus
      'Exit Function
   'End If
   
    If Not OptImpOtor.Value Then
        TxtImporteOtorgamiento.Text = 0
    End If

    If Not OptCapNoint.Value Then
       TxtPorCapNoInt.Text = 0
    End If

    If Not OptIntNoCap.Value Then
        TxtPorIntNoCap.Text = 0
    End If

   If Not OptCapInt.Value Then
        TxtCapmasInt.Text = 0
    End If
   
Else
   TxtImporteOtorgamiento.Text = 0
   TxtPorCapNoInt.Text = 0
   TxtPorIntNoCap.Text = 0
   TxtCapmasInt.Text = 0
End If

Exit Function
merror:
tratarerrores "Error en funcion Datos OtorgamientoOk"
End Function
Private Sub cmdAplicar_Click()
'si acepta grabo los cambios
Dim sql As String
Dim MopcionVencimiento21 As Long
Dim MopcionVencimiento22 As Long
Dim Tam As Long
Dim Cadena1 As String
Dim Cadena2 As String
Dim Cadena3 As String
Dim Cadena4 As String
Dim Cadena5 As String
Dim Cadena6 As String
Dim Credimac As Long
Dim Vencimiento2Mora As Long
On Error GoTo merror

MopcionVencimiento21 = 0
MopcionVencimiento22 = 0

Call ReemplazarComillas(Me)

'si estoy en solapa de datos empresa
If TabStripOpciones.SelectedItem.Index = 1 Then
   If Not DatosEmpresaOk() Then Exit Sub
   
   If Not MsgP("¿Confirma los datos de la empresa?") Then Exit Sub
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "UPDATE configuracionsistema SET " & _
         "empresa='" & CStr(TxtEmpresa.Text) & _
         "',ciudad='" & CStr(TxtCiudad.Text) & "',cuit='" & CStr(TxtCuit.Text) & _
         "',ingresosbrutos='" & CStr(TxtIngresosBrutos.Text) & "',iva='" & CStr(TxtIva.Text) & _
         "',telefono='" & CStr(TxtTelefono.Text) & "',domicilio='" & CStr(TxtDomicilio.Text) & "',email='" & CStr(TxtEmail.Text) & _
         "',website='" & CStr(TxtWebsite.Text) & "',horarioatencion='" & CStr(TxtHorarioAtencion.Text) & _
         "',lugarespago='" & CStr(TxtLugaresPago.Text) & "'"
         
      
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans
   
   MsgI "Los datos de la empresa fueron actualizados"
End If

'si estoy en la solapa de tasas de interes
If TabStripOpciones.SelectedItem.Index = 2 Then
   If Not DatosTasasOk() Then Exit Sub
   
   If Not MsgP("¿Confirma las tasas de interes?") Then Exit Sub
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update configuracionsistema set " & _
         "tasamora= '" & ConvertirDblSql(CDbl(TxtTasaMora.Text)) & "'," & _
         "aplicartasarefinanciacion=" & CheckAplicarTasaRefinanciacion.Value & _
         ",tasarefinanciacion='" & ConvertirDblSql(CDbl(TxtTasaRefinanciacion.Text)) & "',tasamora2='" & ConvertirDblSql(CDbl(TxtTasaMora2.Text)) & "'"
 
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans
   
   MsgI "Las tasas de interes por mora y refinanciacion fueron actualizadas"
   
End If

'si estoy en solapa requisitos
If TabStripOpciones.SelectedItem.Index = 3 Then
   If Not DatosRequisitosOk() Then Exit Sub
   
   If Not MsgP("¿Confirma los requisitos generales de credito?") Then Exit Sub
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update configuracionsistema set garante=" & CheckGarante.Value & _
         ",clientesimultaneo=" & CheckClienteSimultaneo.Value & _
         ",clientejudicial=" & CheckClienteJudicial.Value & _
         ",finalizarautomaticamente=" & CheckFinalizarAutomaticamente.Value & _
         ",pagarcuotasdesordenadas=" & CheckPagarCuotasDesordenadas.Value & _
         ",aplicarcobrosparciales=" & CheckAplicarCobrosParciales.Value & _
         ",redondearcuotas=" & CheckRedondear.Value & _
         ",aplicarrecibo=" & CheckAplicarRecibo.Value & _
         ",creditosdiferidos=" & CheckCreditosDiferidos.Value & _
         ",EdadMaxCredito=" & txtedad.Text & _
         ",fechalimite= '" & ConvertirFechaSql(CDate(FecLimite), "DD/MM/YYYY") & _
         "',cobrosdiferidos=" & CheckCobrosDiferidos.Value

   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans
   
   MsgI "Los requisitos de credito fueron actualizados"
   
End If

'datos de vencimientos
If TabStripOpciones.SelectedItem.Index = 4 Then
   If Not DatosComprobantesOk() Then Exit Sub
   
   If Not MsgP("¿Confirma las opciones vencimientos?") Then Exit Sub
   
   If OptionVencimiento21.Value Then
      MopcionVencimiento21 = 1
   End If
   
   If OptionVencimiento22.Value Then
      MopcionVencimiento22 = 1
   End If
   
   If OptionVencimiento2Mora.Value Then
      Vencimiento2Mora = 1
   Else
      Vencimiento2Mora = 0
   End If
        
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update configuracionsistema set " & _
         "aplicarsegundovencimiento=" & CheckAplicarSegundoVencimiento.Value & _
         ",vencimiento2importe=" & ConvertirDblSql(CCur(TxtVencimiento2Importe.Text)) & _
         ",vencimiento2porcentaje=" & ConvertirDblSql(CCur(TxtVencimiento2Porcentaje.Text)) & _
         ",aplicarvencimiento2mora=" & CLng(Vencimiento2Mora) & _
         ",aplicarvencimientosabados=" & CheckAplicarVencimientoSabados.Value
         
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans
   
   MsgI "Las opciones de vencimientos fueron actualizadas"
   
End If

'cartareclamo
If TabStripOpciones.SelectedItem.Index = 5 Then
   
   If Not DatosCartaReclamoOk() Then Exit Sub
   
   If Not MsgP("¿Confirma el texto de la carta reclamo?") Then Exit Sub
   
   Tam = Len(TxtTextoCartaReclamo2.Text)
   
   If Tam <= 255 Then
      Cadena1 = Mid(TxtTextoCartaReclamo2.Text, 1, Tam)
      Cadena2 = " "
   Else
      Cadena1 = Mid(TxtTextoCartaReclamo2.Text, 1, 255)
      Cadena2 = Mid(TxtTextoCartaReclamo2.Text, 256, Tam - 255)
   End If
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update configuracionsistema " & _
         "set textocartareclamo1='" & CStr(Cadena1) & _
         "',textocartareclamo2='" & CStr(Cadena2) & "'"
   
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans
   
   MsgI "El texto de la carta reclamo fue actualizado"
End If

'libre deuda
If TabStripOpciones.SelectedItem.Index = 6 Then
   
   If Not DatosLibreDeudaOk() Then Exit Sub
   
   If Not MsgP("¿Confirma el texto de libre deuda?") Then Exit Sub
   
   Tam = Len(TxtTextoLibreDeuda.Text)
   
   If Tam <= 255 Then
      Cadena1 = Mid(TxtTextoLibreDeuda.Text, 1, Tam)
      Cadena2 = " "
   Else
      Cadena1 = Mid(TxtTextoLibreDeuda.Text, 1, 255)
      Cadena2 = Mid(TxtTextoLibreDeuda.Text, 256, Tam - 255)
   End If
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update configuracionsistema " & _
         "set textolibredeuda1='" & CStr(Cadena1) & _
         "',textolibredeuda2= '" & CStr(Cadena2) & "'"
   
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans
   
   MsgI "El texto de libre deuda fue actualizado"
End If


'si estoy en solapa gastos
If TabStripOpciones.SelectedItem.Index = 7 Then
   If Not DatosGastosOk() Then Exit Sub
   
   If Not MsgP("¿Confirma las opciones de gastos administrativos?") Then Exit Sub
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update configuracionsistema set aplicargastos=" & CheckAplicarGastos.Value & _
         ",noaplicargastosrefinanciacion=" & CheckNoAplicarGastosRefinanciacion.Value & _
         ",aplicargastoscuota1=" & CheckGastosCuota1.Value & _
         ",aplicargastoscuota2=" & CheckGastosCuota2.Value & _
         ",importegastosfijos=" & ConvertirDblSql(CCur(TxtImporteGastosFijo.Text)) & _
         ",importegastos=" & ConvertirDblSql(CCur(TxtImporteGastos.Text)) & _
         ",PorcentajeCapitalyNoInt = " & ConvertirDblSql(CDbl(TxtcapNoint.Text)) & "" & _
         ",PorcentajefuncNoCapital = " & ConvertirDblSql(CDbl(TxtFuncIntnocap.Text)) & "" & _
         ",PorcentajeCapitalInteres = " & ConvertirDblSql(CDbl(TxtCapInt.Text)) & ""
         
        
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans
   
   MsgI "Las opciones de gastos administrativos fueron actualizadas"
   
End If

'si estoy en solapa seguros
If TabStripOpciones.SelectedItem.Index = 8 Then
   If Not DatosSegurosOk() Then Exit Sub
   
   If Not MsgP("¿Confirma las opciones de seguros?") Then Exit Sub
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update configuracionsistema set aplicarseguro=" & CheckAplicarSeguro.Value & _
         ",noaplicarsegurosrefinanciacion=" & CheckNoAplicarSegurosRefinanciacion.Value & _
         ",aplicarseguroscuota1=" & CheckSegurosCuota1.Value & _
         ",importeseguro=" & ConvertirDblSql(CCur(TxtImporteSeguro.Text)) & _
         ",alicuotaseguros=" & ConvertirDblSql(CDbl(TxtAlicuotaSeguros.Text)) & _
         ",importesegurosfijos=" & ConvertirDblSql(CCur(TxtSeguroFijo.Text))
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans
   
   MsgI "Las opciones de seguros fueron actualizadas"
   
End If

'si estoy en solapa impuestos
If TabStripOpciones.SelectedItem.Index = 9 Then
   
   If Not DatosImpuestosOk() Then Exit Sub
   
   If Not MsgP("¿Confirma las opciones de impuestos?") Then Exit Sub
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   If OptionImpuestosCredimaco.Value = True Then
      Credimac = 1
   Else
      Credimac = 0
   End If
   sql = "update configuracionsistema set aplicarimpuestos=" & CheckAplicarImpuestos.Value & _
         ",importeimpuestos=" & ConvertirDblSql(CCur(TxtImporteImpuestos.Text)) & _
         ",importeimpuestosfijos=" & ConvertirDblSql(CCur(TxtImporteFijoImpuestos.Text)) & _
         ",porcentajeiva=" & ConvertirDblSql(CDbl(TxtPorcentajeIva.Text)) & _
         ",noaplicarimpuestosrefinanciacion=" & CheckNoAplicarImpuestosRefinanciacion.Value & _
         ",aplicarimpuestoscuota1=" & CheckImpuestosCuota1.Value & "," & _
         "aplicarimpuestoscuota2=" & CheckImpuestosCuota2.Value & _
         ",impuestoscredimaco=" & CLng(Credimac)
         
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans
   
   MsgI "Las opciones de impuestos fueron actualizadas"
End If


'impresion
If TabStripOpciones.SelectedItem.Index = 10 Then
   If Not DatosImpresionOk() Then Exit Sub
   
   If Not MsgP("¿Confirma las opciones de impresion?") Then Exit Sub
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update configuracionsistema set " & _
         "modelofactura1=" & CheckModeloFactura1.Value & _
         ",modelofactura2=" & CheckModeloFactura2.Value & _
         ",modelofactura3=" & CheckModeloFactura3.Value & _
         ",modelofactura4=" & CheckModeloFactura4.Value & _
         ",margentop=" & CLng(TxtTop.Text) & ",margenleft=" & CLng(TxtLeft.Text) & ",margenbotom=" & CLng(TxtBotom.Text) & ",mostrarrecuadros=" & CheckMostrarRecuadros.Value & ",numcopias=" & CLng(TxtNumCopias.Text) & _
         ",codigoautomatico=" & CLng(TxtCodigoRapipago.Text) & _
         ",num1=" & CLng(TxtNumEmpresa.Text) & ",ultimonumrecibo=" & CLng(TxtNumRecibo.Text) & _
         ",imprimirmoraiva=" & CheckImprimirMoraIva.Value
 
         
   cnSQL.Execute sql
   
  
   'fin de transaccion
   cnSQL.CommitTrans
   
   MsgI "Las opciones de impresion fueron actualizadas"
End If

'pagare
If TabStripOpciones.SelectedItem.Index = 11 Then
   
   If Not DatosAcuerdoMutuoOk() Then Exit Sub
   
   If Not MsgP("¿Confirma el texto del mutuo acuerdo?") Then Exit Sub
   
   'saco las dos cadenas del 5
   Tam = Len(TxtParrafo5.Text)
   If Tam <= 255 Then
      Cadena1 = Mid(TxtParrafo5.Text, 1, Tam)
      Cadena2 = " "
      Cadena3 = " "
      
   Else
      Cadena1 = Mid(TxtParrafo5.Text, 1, 255)
      Cadena2 = Mid(TxtParrafo5.Text, 256, 255)
      If Tam > 510 Then
         Cadena3 = Mid(TxtParrafo5.Text, 511, Tam - 510)
      End If
   End If
   
   'saco las dos cadenas del 6
   Tam = Len(TxtParrafo6.Text)
   If Tam <= 255 Then
      Cadena4 = Mid(TxtParrafo6.Text, 1, Tam)
      Cadena5 = " "
      Cadena6 = " "
   Else
      Cadena4 = Mid(TxtParrafo6.Text, 1, 255)
      Cadena5 = Mid(TxtParrafo6.Text, 256, 255)
      If Tam > 510 Then
         Cadena6 = Mid(TxtParrafo6.Text, 511, Tam - 510)
      End If
   End If
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update configuracionsistema " & _
         "set textoacuerdomutuo1='" & CStr(TxtParrafo1.Text) & _
         "',textoacuerdomutuo2='" & CStr(TxtParrafo2.Text) & _
         "',textoacuerdomutuo3='" & CStr(TxtParrafo3.Text) & _
         "',textoacuerdomutuo4='" & CStr(TxtParrafo4.Text) & _
         "',textoacuerdomutuo5='" & CStr(Cadena1) & _
         "',textoacuerdomutuo6='" & CStr(Cadena2) & _
         "',textoacuerdomutuo7='" & CStr(Cadena3) & _
         "',textoacuerdomutuo8='" & CStr(Cadena4) & _
         "',textoacuerdomutuo9='" & CStr(Cadena5) & _
         "',textoacuerdomutuo10='" & CStr(Cadena6) & "'"
   
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans
   
   MsgI "El texto del acuerdo mutuo fue actualizado"
End If

'otorgamiento
If TabStripOpciones.SelectedItem.Index = 12 Then
   
   If Not DatosOtorgamientoOk() Then Exit Sub
   
   If Not MsgP("¿Confirma el importe de otorgamiento?") Then Exit Sub
   
   'inicio transaccion
   cnSQL.BeginTrans
   
   sql = "update configuracionsistema " & _
   "set aplicarotorgamiento=" & CheckAplicarOtorgamiento.Value & _
   ",importeotorgamiento=" & ConvertirDblSql(CCur(TxtImporteOtorgamiento.Text)) & _
   ",noaplicarotorrefin=" & CheckNoAplicarOtRefin.Value & _
   ",aplicarotorgamientocuota1=" & CheckOtorgamiento1.Value & _
   ",OtorCapNoInt=" & ConvertirDblSql(CDbl(TxtPorCapNoInt)) & _
   ",OtorIntNoCap=" & ConvertirDblSql(CDbl(TxtPorIntNoCap)) & _
   ",OtorCapmasInt=" & ConvertirDblSql(CDbl(TxtCapmasInt))

   
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans
   MsgI "El gasto de otorgamiento fue actualizado"
End If

'actualizo los cambios en la pantalla
Call CargarOpciones

Exit Sub
merror:
tratarerrores "Error actualizando las opciones del sistema"
End Sub
Private Sub CheckImpuestosCuota1_Click()
If CheckImpuestosCuota1.Value = 1 Then
   CheckImpuestosCuota2.Value = 0
   CheckImpuestosCuota2.Enabled = False
Else
   CheckImpuestosCuota2.Enabled = True
End If
End Sub
Private Sub CheckImpuestosCuota2_Click()
If CheckImpuestosCuota2.Value = 1 Then
   CheckImpuestosCuota1.Value = 0
   CheckImpuestosCuota1.Enabled = False
Else
   CheckImpuestosCuota1.Enabled = True
End If
End Sub
'todos los botones
Private Sub CmdRestablecerLibreDeuda_Click()
On Error GoTo merror

TxtTextoLibreDeuda.Text = "Certificamos que el cliente arriba mencionado " & _
                          "no tiene deudas al dia de la fecha en concepto " & _
                          "de cuotas vencidas. El presente certificado " & _
                          "reemplaza a todas las boletas y recibos emitidos " & _
                          "hasta la fecha.No asi a las boletas aun no abonadas."
Exit Sub
merror:
tratarerrores "Error restableciendo texto de libre deuda"
End Sub
Private Sub CmdRestablecerCarta_Click()
On Error GoTo merror

TxtTextoCartaReclamo2.Text = "en concepto de documentos impagos." & _
                             "Los mismos deberan ser cancelados en un plazo " & _
                             "maximo de 3 dias habiles en nuestras oficinas." & _
                             "Caso contrario iniciaremos los tramites legales " & _
                             "correspondientes.Sin otro particular saludamos a " & _
                             "UD. atentamente."
Exit Sub
merror:
tratarerrores "Error restableciendo texto de carta reclamo"
End Sub
Private Sub CmdPredeterminadoAcuerdo_Click()
On Error GoTo merror

TxtParrafo1.Text = "Pagaremos solidariamente en su sede de ciudad Capital la cantidad de pesos"
TxtParrafo2.Text = "de dinero por el que se firman dos pagares por el total." & _
"Los que seran reintegrados al firmante titular dentro de los 45 dias de cancelada la operacion." & _
"Dicha cantidad la abonaremos en"
TxtParrafo3.Text = "cuotas con vencimientos mensuales consecutivos"
TxtParrafo4.Text = "cuyo primer vencimiento operara el dia "
TxtParrafo5.Text = "La falta de pago de cualesquiera de las cuotas en las fechas convenidas producira la caducidad de los plazos otorgados para el pago y dara derecho a exigir por via ejecutiva, el pago del saldo que adeudare por esta obligacion, mas el interes compensatorio del 45 % anual sobre saldos y el interes punitorio del 6 % mensual." & _
"La mora se producira por el mero transcurso del plazo estipulado en el presente, sin que sea necesaria ninguna clase de intimacion judicial o extrajudicial alguna." & _
"A todos los efectos el/los firmantes se somete/n a la competencia de los tribunales de Capital Federal renunciando a cualquier otro fuero o jurisdiccion." & _
"El deudor constituye domicilio legal en"
TxtParrafo6.Text = "donde se realizara la cobranza en una unica oportunidad por mes y sin cargo por la gestion de cobranza, de las cuotas pactadas, teniendo las visitas adicionales un costo de 5 pesos cada una." & _
"En el mismo domicilio denunciado se entregaran las notificaciones que se practicaren tanto judiciales como extrajudiciales." & _
"En caso de promoverse juicio, renunciamos al derecho de recusar sin causa, y para el caso de subasta de los bienes que se embarguen aceptamos se designe el martillero que proponga la acreedora."

Exit Sub
merror:
tratarerrores "Error restableciendo texto del mutuo acuerdo"
End Sub
Private Sub CheckAplicarTasaRefinanciacion_Click()
On Error GoTo merror

If CheckAplicarTasaRefinanciacion.Value = 1 Then
   UpDown6.Enabled = True
   TxtTasaRefinanciacion.BackColor = vbWhite
   TxtTasaRefinanciacion.Enabled = True
   If FrameTasas.Visible Then
      TxtTasaRefinanciacion.SetFocus
   End If
Else
   UpDown6.Enabled = False
   TxtTasaRefinanciacion.BackColor = &HFFFFC0
   TxtTasaRefinanciacion.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error aplicando tasa de refinanciacion"
End Sub
Private Sub CheckAplicarSegundoVencimiento_Click()
On Error GoTo merror

If CheckAplicarSegundoVencimiento.Value = 1 Then
   OptionVencimiento21.Enabled = True
   OptionVencimiento22.Enabled = True
   OptionVencimiento2Mora.Enabled = True
   
   OptionVencimiento21.Value = False
   OptionVencimiento21.Value = True
   
   If CCur(VG_VENCIMIENTO2IMPORTE) > 0 Then
      OptionVencimiento21.Value = True
   End If
   If CCur(VG_VENCIMIENTO2PORCENTAJE) > 0 Then
      OptionVencimiento22.Value = True
   End If
   'si aplico mora entre el primer y segundo vto
   If VG_APLICARVENCIMIENTO2MORA Then
      OptionVencimiento2Mora.Value = True
   End If
Else
   UpDown16.Enabled = False
   UpDown19.Enabled = False
   OptionVencimiento21.Value = False
   OptionVencimiento22.Value = False
   
   OptionVencimiento2Mora.Enabled = False
   OptionVencimiento21.Enabled = False
   OptionVencimiento22.Enabled = False
   
   TxtVencimiento2ImporteBis.BackColor = &HFFFFC0
   TxtVencimiento2Porcentaje.BackColor = &HFFFFC0
   TxtVencimiento2ImporteBis.Enabled = False
   TxtVencimiento2Porcentaje.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error aplicando segundo vencimiento"
End Sub
Private Sub CheckAplicarImpuestos_Click()
On Error GoTo merror

If CheckAplicarImpuestos.Value = 1 Then
   OptionImpuestosCredimaco.Enabled = True
   OptionImpuestosFijos.Enabled = True
   OptionImporteImpuestos.Enabled = True
   CheckNoAplicarImpuestosRefinanciacion.Enabled = True
   CheckImpuestosCuota1.Enabled = True
   CheckImpuestosCuota2.Enabled = True
   TxtPorcentajeIva.Enabled = True
   TxtPorcentajeIva.BackColor = vbWhite
   
   'por defecto pone habil la primera
   'este es un truco para que ejecute el option
   OptionImpuestosCredimaco.Value = False
   OptionImpuestosCredimaco.Value = True
      
   If VG_IMPUESTOSCREDIMACO Then
      OptionImpuestosCredimaco.Value = True
   End If
   If CCur(TxtImporteFijoImpuestos.Text) > 0 Then
      OptionImpuestosFijos.Value = True
   End If
   
   If CCur(TxtImporteImpuestos.Text) > 0 Then
      OptionImporteImpuestos.Value = True
   End If
  
Else
   OptionImpuestosCredimaco.Enabled = False
   OptionImpuestosFijos.Enabled = False
   OptionImporteImpuestos.Enabled = False
   CheckNoAplicarImpuestosRefinanciacion.Enabled = False
   CheckImpuestosCuota1.Enabled = False
   CheckImpuestosCuota2.Enabled = False
   TxtPorcentajeIva.Enabled = False
   TxtPorcentajeIva.BackColor = &H80000013
   UpDown13.Enabled = False
   UpDown22.Enabled = False
   TxtImporteFijoImpuestos.Enabled = False
   TxtImporteImpuestos.Enabled = False
   TxtImporteFijoImpuestos.BackColor = &HFFFFC0
   TxtImporteImpuestos.BackColor = &HFFFFC0
End If

Exit Sub
merror:
tratarerrores "Error aplicando impuestos"
End Sub
Private Sub CheckaplicarGastos_Click()
On Error GoTo merror

If CheckAplicarGastos.Value = 1 Then
   OptionGastosFijos.Enabled = True
   OptionImporteGastos.Enabled = True
   PorccapNoint.Enabled = True
   PorcIntNoCap.Enabled = True
   PorcCapint.Enabled = True
   CheckGastosCuota1.Enabled = True
   CheckGastosCuota2.Enabled = True
   PorcCapint.Enabled = True
   PorcIntNoCap.Enabled = True
   PorccapNoint.Enabled = True
   CheckNoAplicarGastosRefinanciacion.Enabled = True
   
   OptionGastosFijos.Value = False
   OptionGastosFijos.Value = True
   
   If CCur(TxtImporteGastosFijoBis.Text) > 0 Then
      OptionGastosFijos.Value = True
   End If
   
   If CCur(TxtImporteGastos.Text) > 0 Then
      OptionImporteGastos.Value = True
   End If
   
   If TxtcapNoint.Text > 0 Then
        PorccapNoint.Value = True
   End If
   
   If TxtFuncIntnocap.Text > 0 Then
        PorcIntNoCap.Value = True
   End If
   
   If TxtCapInt.Text > 0 Then
        PorcCapint.Value = True
   End If
   
 
Else
   OptionGastosFijos.Enabled = False
   OptionImporteGastos.Enabled = False
   PorccapNoint.Enabled = False
   PorcIntNoCap.Enabled = False
   PorcCapint.Enabled = False
   CheckGastosCuota1.Enabled = False
   CheckGastosCuota2.Enabled = False
   CheckNoAplicarGastosRefinanciacion.Enabled = False
   TxtImporteGastosFijoBis.BackColor = &HFFFFC0
   TxtImporteGastosFijoBis.Enabled = False
   TxtImporteGastosBis.BackColor = &HFFFFC0
   TxtImporteGastosBis.Enabled = False
   UpDown3.Enabled = False
   UpDown7.Enabled = False
   UpDown1.Enabled = False
   UpDown8.Enabled = False
   UpDown9.Enabled = False
   TxtcapNoint.BackColor = &HFFFFC0
   TxtcapNoint.Enabled = False
   TxtFuncIntnocap.BackColor = &HFFFFC0
   TxtFuncIntnocap.Enabled = False
   TxtCapInt.BackColor = &HFFFFC0
   TxtCapInt.Enabled = False
   
End If

Exit Sub
merror:
tratarerrores "Error aplicando gastos administrativos"
End Sub
Private Sub CheckaplicarSeguro_Click()
On Error GoTo merror

If CheckAplicarSeguro.Value = 1 Then
   TxtAlicuotaSeguros.Enabled = True
   TxtAlicuotaSeguros.BackColor = vbWhite
   TxtSeguroFijo.Enabled = True
   TxtSeguroFijo.BackColor = vbWhite
   TxtImporteSeguro.Enabled = True
   TxtImporteSeguro.BackColor = vbWhite
   
   CheckNoAplicarSegurosRefinanciacion.Enabled = True
   CheckSegurosCuota1.Enabled = True
Else
   TxtAlicuotaSeguros.Enabled = False
   TxtAlicuotaSeguros.BackColor = &HFFFFC0
   TxtSeguroFijo.Enabled = False
   TxtSeguroFijo.BackColor = &HFFFFC0
   TxtImporteSeguro.Enabled = False
   TxtImporteSeguro.BackColor = &HFFFFC0
   
   CheckSegurosCuota1.Value = 0
   CheckSegurosCuota1.Enabled = False
   
   CheckNoAplicarSegurosRefinanciacion.Value = 0
   CheckNoAplicarSegurosRefinanciacion.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error aplicando seguros"
End Sub
Private Sub CheckModeloFactura2_Click()
On Error GoTo merror

If CheckModeloFactura2.Value = 1 Then
   CheckModeloFactura3.Value = 0
   CheckModeloFactura4.Value = 0
End If

Exit Sub
merror:
tratarerrores "Error estableciendo el modelo de factura2"
End Sub
Private Sub CheckModeloFactura3_Click()
On Error GoTo merror

If CheckModeloFactura3.Value = 1 Then
   CheckModeloFactura2.Value = 0
   CheckModeloFactura4.Value = 0
End If

Exit Sub
merror:
tratarerrores "Error estableciendo el modelo de factura3"
End Sub
Private Sub CheckModeloFactura4_Click()
On Error GoTo merror

If CheckModeloFactura4.Value = 1 Then
   CheckModeloFactura2.Value = 0
   CheckModeloFactura3.Value = 0
End If

Exit Sub
merror:
tratarerrores "Error estableciendo el modelo de factura4"
End Sub

Private Sub OptCapInt_Click()
On Error GoTo merror

If OptCapInt.Value Then
   
   UpDown12.Enabled = False
   UpDown15.Enabled = False
   UpDown11.Enabled = False
   UpDown14.Enabled = True
   
   TxtCapmasInt.Enabled = True
   TxtCapmasInt.BackColor = vbWhite
   If FrameOtorgamiento.Visible Then
      TxtCapmasInt.SetFocus
   End If
   
   'deshabilito los otros
   TxtImporteOtorgamiento.BackColor = &HFFFFC0
   TxtImporteOtorgamiento.Enabled = False
   
   TxtPorCapNoInt.BackColor = &HFFFFC0
   TxtPorCapNoInt.Enabled = False
   
   TxtPorIntNoCap.BackColor = &HFFFFC0
   TxtPorIntNoCap.Enabled = False
   
 
End If

Exit Sub
merror:
tratarerrores "Error estableciendo Porcentaje en función del capital + intereses"
End Sub

Private Sub OptCapNoint_Click()
On Error GoTo merror

If OptCapNoint.Value Then
   
   UpDown14.Enabled = False
   UpDown15.Enabled = False
   UpDown12.Enabled = False
   UpDown11.Enabled = True
   
   TxtPorCapNoInt.Enabled = True
   TxtPorCapNoInt.BackColor = vbWhite
   If FrameOtorgamiento.Visible Then
      TxtPorCapNoInt.SetFocus
   End If
   
   'deshabilito los otros
   TxtImporteOtorgamiento.BackColor = &HFFFFC0
   TxtImporteOtorgamiento.Enabled = False
   
   TxtPorIntNoCap.BackColor = &HFFFFC0
   TxtPorIntNoCap.Enabled = False
   
   TxtCapmasInt.BackColor = &HFFFFC0
   TxtCapmasInt.Enabled = False
   
 
End If

Exit Sub
merror:
tratarerrores "Error estableciendo Por. en fun. del capital y no del interés"
End Sub

Private Sub OptImpOtor_Click()
On Error GoTo merror

If OptImpOtor.Value Then
   
   
   UpDown14.Enabled = False
   UpDown11.Enabled = False
   UpDown12.Enabled = False
   UpDown15.Enabled = True
   
   TxtImporteOtorgamiento.Enabled = True
   TxtImporteOtorgamiento.BackColor = vbWhite
   If FrameOtorgamiento.Visible Then
      TxtImporteOtorgamiento.SetFocus
   End If
   
   'deshabilito los otros
   TxtPorCapNoInt.BackColor = &HFFFFC0
   TxtPorCapNoInt.Enabled = False
   
   TxtPorIntNoCap.BackColor = &HFFFFC0
   TxtPorIntNoCap.Enabled = False
   
   TxtCapmasInt.BackColor = &HFFFFC0
   TxtCapmasInt.Enabled = False
   
 
End If

Exit Sub
merror:
tratarerrores "Error estableciendo importe de Otorgamiento"
End Sub

Private Sub OptIntNoCap_Click()
On Error GoTo merror

If OptIntNoCap.Value Then
   
   UpDown14.Enabled = False
   UpDown15.Enabled = False
   UpDown11.Enabled = False
   UpDown12.Enabled = True
   
   TxtPorIntNoCap.Enabled = True
   TxtPorIntNoCap.BackColor = vbWhite
   If FrameOtorgamiento.Visible Then
      TxtPorIntNoCap.SetFocus
   End If
   
   'deshabilito los otros
   TxtImporteOtorgamiento.BackColor = &HFFFFC0
   TxtImporteOtorgamiento.Enabled = False
   
   TxtPorCapNoInt.BackColor = &HFFFFC0
   TxtPorCapNoInt.Enabled = False
   
   TxtCapmasInt.BackColor = &HFFFFC0
   TxtCapmasInt.Enabled = False
   
 
End If

Exit Sub
merror:
tratarerrores "Error estableciendo Porc. en fun. del interés y no del capital"
End Sub

Private Sub OptionGastosFijos_Click()
On Error GoTo merror

If OptionGastosFijos.Value Then
   
   UpDown7.Enabled = False
   UpDown1.Enabled = False
   UpDown8.Enabled = False
   UpDown9.Enabled = False
   UpDown3.Enabled = True
   
   TxtImporteGastosFijoBis.Enabled = True
   TxtImporteGastosFijoBis.BackColor = vbWhite
   If FrameGastos.Visible Then
      TxtImporteGastosFijoBis.SetFocus
   End If
   
   'deshabilito los otros
   
   TxtImporteGastosBis.BackColor = &HFFFFC0
   TxtImporteGastosBis.Enabled = False
      
   TxtcapNoint.BackColor = &HFFFFC0
   TxtcapNoint.Enabled = False
   
   TxtFuncIntnocap.BackColor = &HFFFFC0
   TxtFuncIntnocap.Enabled = False
   
   TxtCapInt.BackColor = &HFFFFC0
   TxtCapInt.Enabled = False
   
   
   
End If

Exit Sub
merror:
tratarerrores "Error estableciendo importe fijo"
End Sub
'todos los options
Private Sub OptionImportegastos_Click()
On Error GoTo merror

If OptionImporteGastos.Value Then
   UpDown7.Enabled = True
   UpDown3.Enabled = False
   UpDown1.Enabled = False
   UpDown8.Enabled = False
   UpDown9.Enabled = False
   
   TxtImporteGastosBis.Enabled = True
   TxtImporteGastosBis.BackColor = vbWhite
   If FrameGastos.Visible Then
      TxtImporteGastosBis.SetFocus
   End If
   'deshabilito los otros
   
   TxtImporteGastosFijoBis.Enabled = False
   TxtImporteGastosFijoBis.BackColor = &HFFFFC0
   

   TxtcapNoint.BackColor = &HFFFFC0
   TxtcapNoint.Enabled = False

   TxtFuncIntnocap.BackColor = &HFFFFC0
   TxtFuncIntnocap.Enabled = False
   
   TxtCapInt.BackColor = &HFFFFC0
   TxtCapInt.Enabled = False
   
End If

Exit Sub
merror:
tratarerrores "Error estableciendo Importe fijo a dividir entre las cuotas del credito"
End Sub
Private Sub OptionImpuestosCredimaco_Click()
If OptionImpuestosCredimaco.Value Then
   TxtImporteFijoImpuestos.Enabled = False
   TxtImporteImpuestos.Enabled = False
   TxtImporteFijoImpuestos.BackColor = &HFFFFC0
   TxtImporteImpuestos.BackColor = &HFFFFC0
   UpDown13.Enabled = False
   UpDown22.Enabled = False
End If
End Sub
Private Sub OptionImpuestosFijos_Click()
If OptionImpuestosFijos.Value Then
   TxtImporteFijoImpuestos.Enabled = True
   UpDown13.Enabled = True
   TxtImporteFijoImpuestos.BackColor = vbWhite
   If FrameImpuestos.Visible Then
      TxtImporteFijoImpuestos.SetFocus
   End If
   'deshabilito los demas
   TxtImporteImpuestos.Enabled = False
   TxtImporteImpuestos.BackColor = &HFFFFC0
   UpDown22.Enabled = False
End If
End Sub
Private Sub OptionImporteImpuestos_Click()
On Error GoTo merror

If OptionImporteImpuestos.Value Then
   TxtImporteImpuestos.Enabled = True
   UpDown22.Enabled = True
   TxtImporteImpuestos.BackColor = vbWhite
   If FrameImpuestos.Visible Then
      TxtImporteImpuestos.SetFocus
   End If
   'deshabilito los demas
   TxtImporteFijoImpuestos.Enabled = False
   TxtImporteFijoImpuestos.BackColor = &HFFFFC0
   UpDown13.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error estableciendo importe de impuestos"
End Sub
Private Sub OptionVencimiento21_Click()
On Error GoTo merror

'importe fijo por todo el periodo
If OptionVencimiento21.Value Then
   UpDown16.Enabled = True
   UpDown19.Enabled = False
   
   TxtVencimiento2ImporteBis.Enabled = True
   TxtVencimiento2ImporteBis.BackColor = vbWhite
   
   If FrameComprobantes.Visible Then
      TxtVencimiento2ImporteBis.SetFocus
   End If
   
   'deshabilito los demas
   TxtVencimiento2Porcentaje.BackColor = &HFFFFC0
   TxtVencimiento2Porcentaje.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error estableciendo importe de vencimiento21"
End Sub
Private Sub OptionVencimiento22_Click()
On Error GoTo merror

'importe diario fijo
If OptionVencimiento22.Value Then
   UpDown16.Enabled = False
   UpDown19.Enabled = True
   
   TxtVencimiento2Porcentaje.Enabled = True
   TxtVencimiento2Porcentaje.BackColor = vbWhite
   
   If FrameComprobantes.Visible Then
      TxtVencimiento2Porcentaje.SetFocus
   End If
   
   'deshabilito los demas
   TxtVencimiento2ImporteBis.BackColor = &HFFFFC0
   TxtVencimiento2ImporteBis.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error estableciendo importe de vencimiento22"
End Sub
Private Sub OptionVencimiento2Mora_Click()
If OptionVencimiento2Mora.Value Then
   'deshabilito todos los demas
   UpDown16.Enabled = False
   UpDown19.Enabled = False
   TxtVencimiento2ImporteBis.Enabled = False
   TxtVencimiento2ImporteBis.BackColor = &HFFFFC0
   TxtVencimiento2Porcentaje.BackColor = &HFFFFC0
   TxtVencimiento2Porcentaje.Enabled = False
End If
End Sub

Private Sub PorcCapint_Click()
On Error GoTo merror

If PorcCapint.Value Then
   
   UpDown1.Enabled = False
   UpDown3.Enabled = False
   UpDown7.Enabled = False
   UpDown8.Enabled = False
   UpDown9.Enabled = True
   
   TxtCapInt.Enabled = True
   TxtCapInt.BackColor = vbWhite
   If FrameGastos.Visible Then
      TxtCapInt.SetFocus
   End If
   
   'deshabilito los otros
   TxtImporteGastosBis.BackColor = &HFFFFC0
   TxtImporteGastosBis.Enabled = False
      
   TxtImporteGastosFijoBis.Enabled = False
   TxtImporteGastosFijoBis.BackColor = &HFFFFC0
      
   TxtcapNoint.BackColor = &HFFFFC0
   TxtcapNoint.Enabled = False
      
   TxtFuncIntnocap.BackColor = &HFFFFC0
   TxtFuncIntnocap.Enabled = False
    
End If

Exit Sub
merror:
tratarerrores "Error estableciendo Porc. en func. del capital + intereses:"
End Sub

Private Sub PorccapNoint_Click()
On Error GoTo merror

If PorccapNoint.Value Then
   
   UpDown7.Enabled = False
   UpDown8.Enabled = False
   UpDown9.Enabled = False
   UpDown3.Enabled = False
   UpDown1.Enabled = True
   
   TxtcapNoint.Enabled = True
   TxtcapNoint.BackColor = vbWhite
   If FrameGastos.Visible Then
      TxtcapNoint.SetFocus
   End If
   
   'deshabilito el otro
   
   TxtImporteGastosBis.BackColor = &HFFFFC0
   TxtImporteGastosBis.Enabled = False
   
   TxtImporteGastosFijoBis.Enabled = False
   TxtImporteGastosFijoBis.BackColor = &HFFFFC0
     
   TxtFuncIntnocap.BackColor = &HFFFFC0
   TxtFuncIntnocap.Enabled = False
   
   TxtCapInt.BackColor = &HFFFFC0
   TxtCapInt.Enabled = False
   
End If

Exit Sub
merror:
tratarerrores "Error estableciendo Capital  No interes"
End Sub

Private Sub PorcIntNoCap_Click()
On Error GoTo merror

If PorcIntNoCap.Value Then
   
   UpDown1.Enabled = False
   UpDown3.Enabled = False
   UpDown7.Enabled = False
   UpDown9.Enabled = False
   UpDown8.Enabled = True
   
   TxtFuncIntnocap.Enabled = True
   TxtFuncIntnocap.BackColor = vbWhite
   If FrameGastos.Visible Then
      TxtFuncIntnocap.SetFocus
   End If
   
   'deshabilito el otro
   TxtImporteGastosBis.BackColor = &HFFFFC0
   TxtImporteGastosBis.Enabled = False
      
   TxtImporteGastosFijoBis.Enabled = False
   TxtImporteGastosFijoBis.BackColor = &HFFFFC0
   
   
   TxtcapNoint.BackColor = &HFFFFC0
   TxtcapNoint.Enabled = False
      
   TxtCapInt.BackColor = &HFFFFC0
   TxtCapInt.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error estableciendo Porc. en fun. del interés y no del capital"
End Sub

Private Sub TabStripOpciones_Click()
On Error GoTo merror

FrameEmpresa.Visible = False
FrameTasas.Visible = False
FrameRequisitos.Visible = False
FrameComprobantes.Visible = False
FrameImpuestos.Visible = False
FrameCartaReclamo.Visible = False
FrameLibreDeuda.Visible = False
FrameSeguroVida.Visible = False
FrameGastos.Visible = False
FrameImpresion.Visible = False
FrameMutuo.Visible = False
FrameOtorgamiento.Visible = False

'datos empresa
If TabStripOpciones.SelectedItem.Index = 1 Then
   FrameEmpresa.Visible = True
End If

'tasas de interes
If TabStripOpciones.SelectedItem.Index = 2 Then
   FrameTasas.Visible = True
End If

'requisitos
If TabStripOpciones.SelectedItem.Index = 3 Then
   FrameRequisitos.Visible = True
End If

'comprobantes
If TabStripOpciones.SelectedItem.Index = 4 Then
   FrameComprobantes.Visible = True
End If

'carta reclamo
If TabStripOpciones.SelectedItem.Index = 5 Then
   FrameCartaReclamo.Visible = True
End If

'libre deuda
If TabStripOpciones.SelectedItem.Index = 6 Then
   FrameLibreDeuda.Visible = True
End If

'GASTOS
If TabStripOpciones.SelectedItem.Index = 7 Then
   FrameGastos.Visible = True
End If

'SEGUROS
If TabStripOpciones.SelectedItem.Index = 8 Then
   FrameSeguroVida.Visible = True
End If

'IMPUESTOS
If TabStripOpciones.SelectedItem.Index = 9 Then
   FrameImpuestos.Visible = True
End If

'impresion
If TabStripOpciones.SelectedItem.Index = 10 Then
   FrameImpresion.Visible = True
End If

'mutuo
If TabStripOpciones.SelectedItem.Index = 11 Then
   FrameMutuo.Visible = True
End If

'otorgamiento
If TabStripOpciones.SelectedItem.Index = 12 Then
   FrameOtorgamiento.Visible = True
End If

Exit Sub
merror:
tratarerrores "Error seleccionando Opciones"
End Sub
'keycodes de campos
Private Sub TxtEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtCiudad.SetFocus
End If
End Sub
Private Sub TxtCiudad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtCuit.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtEmpresa.SetFocus
End If
End Sub
Private Sub TxtCuit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtIva.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtCiudad.SetFocus
End If
End Sub
Private Sub TxtImporteGastos_Change()
TxtImporteGastosBis.Text = Format(CCur(TxtImporteGastos.Text) / 10, "0.00")
End Sub

Private Sub TxtImporteGastosFijo_Change()
TxtImporteGastosFijoBis.Text = Format(CCur(TxtImporteGastosFijo.Text) / 10, "0.00")
End Sub

Private Sub TxtIva_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtIngresosBrutos.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtCuit.SetFocus
End If
End Sub
Private Sub TxtIngresosBrutos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtDomicilio.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtIva.SetFocus
End If
End Sub
Private Sub TxtDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtTelefono.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtIngresosBrutos.SetFocus
End If
End Sub
Private Sub TxtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtEmail.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtDomicilio.SetFocus
End If
End Sub
Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtWebsite.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtTelefono.SetFocus
End If
End Sub
Private Sub TxtWebsite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtHorarioAtencion.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtEmail.SetFocus
End If
End Sub
Private Sub TxtHorarioAtencion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtLugaresPago.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtWebsite.SetFocus
End If
End Sub
Private Sub TxtLugaresPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   TxtHorarioAtencion.SetFocus
End If
End Sub
Private Sub TxtVencimiento2Importe_Change()
'Hay algunos campos ocultos que permiten que se muestren importes o porcentajes
'con decimales
TxtVencimiento2ImporteBis.Text = Format(CCur(TxtVencimiento2Importe.Text) / 10, "0.00")
End Sub
'mayusculas y trim de campos
Private Sub TxtTextoCartaReclamo2_LostFocus()
TxtTextoCartaReclamo2.Text = Trim(TxtTextoCartaReclamo2.Text)
End Sub
Private Sub TxtTextoLibreDeuda_LostFocus()
TxtTextoLibreDeuda.Text = Trim(TxtTextoLibreDeuda.Text)
End Sub
Private Sub CheckAplicarRecibo_Click()
If CheckAplicarRecibo.Value = 1 Then
   TxtNumRecibo.Enabled = True
   If FrameRequisitos.Visible Then
      TxtNumRecibo.SetFocus
   End If
Else
   TxtNumRecibo.Enabled = False
End If
End Sub
Private Sub CheckAplicarotorgamiento_Click()
On Error GoTo merror

If CheckAplicarOtorgamiento.Value = 1 Then
   TxtImporteOtorgamiento.BackColor = vbWhite
   TxtImporteOtorgamiento.Enabled = True
   UpDown15.Enabled = True
   OptImpOtor.Enabled = True
   OptCapInt.Enabled = True
   OptCapNoint.Enabled = True
   OptIntNoCap.Enabled = True
   If FrameOtorgamiento.Visible Then
      TxtImporteOtorgamiento.SetFocus
   End If
   
   If Val(TxtImporteOtorgamiento.Text) > 0 Then
      OptImpOtor.Value = True
   End If
  
   If Val(TxtPorCapNoInt.Text) > 0 Then
      OptCapNoint.Value = True
   End If
   
   If Val(TxtPorIntNoCap.Text) > 0 Then
        OptIntNoCap.Value = True
   End If
   
   If Val(TxtCapmasInt.Text) > 0 Then
        OptCapInt.Value = True
   End If
   CheckOtorgamiento1.Enabled = True
   
Else
   TxtImporteOtorgamiento.BackColor = &HFFFFC0
   TxtImporteOtorgamiento.Enabled = False
   TxtPorCapNoInt.BackColor = &HFFFFC0
   TxtPorCapNoInt.Enabled = False
   TxtPorIntNoCap.BackColor = &HFFFFC0
   TxtPorIntNoCap.Enabled = False
   TxtCapmasInt.BackColor = &HFFFFC0
   TxtCapmasInt.Enabled = False
   CheckOtorgamiento1.Value = 0
   CheckOtorgamiento1.Enabled = False
   OptImpOtor.Enabled = False
   OptCapInt.Enabled = False
   OptCapNoint.Enabled = False
   OptIntNoCap.Enabled = False
   UpDown11.Enabled = False
   UpDown12.Enabled = False
   UpDown14.Enabled = False
   UpDown15.Enabled = False
   
End If

Exit Sub
merror:
tratarerrores "Error aplicando gastos de otorgamiento"
End Sub
Private Sub CheckGastosCuota1_Click()
If CheckGastosCuota1.Value = 1 Then
   CheckGastosCuota2.Value = 0
   CheckGastosCuota2.Enabled = False
Else
   CheckGastosCuota2.Enabled = True
End If
End Sub
Private Sub CheckGastosCuota2_Click()
If CheckGastosCuota2.Value = 1 Then
   CheckGastosCuota1.Value = 0
   CheckGastosCuota1.Enabled = False
Else
   CheckGastosCuota1.Enabled = True
End If
End Sub

