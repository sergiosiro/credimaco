VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPlanesAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar planes de creditos"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   HelpContextID   =   35
   Icon            =   "FrmPlanesAbm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameSeguroVida 
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   240
      TabIndex        =   68
      Top             =   3840
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox TxtImporteSeguro 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         MaxLength       =   7
         TabIndex        =   74
         Tag             =   "N"
         ToolTipText     =   "Importe fijo a dividir entre las cuotas del credito"
         Top             =   1440
         Width           =   960
      End
      Begin VB.CheckBox CheckAplicarSeguro 
         Caption         =   "Aplicar seguros"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         ToolTipText     =   "Aplica seguro a las facturas de los nuevos creditos"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox CheckNoAplicarSegurosRefinanciacion 
         Caption         =   "No aplicar seguro a las refinanciaciones"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   72
         ToolTipText     =   "Si marca la casilla no se aplicara seguro a los planes refinanciados"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.CheckBox CheckSegurosCuota1 
         Caption         =   "Aplicar seguro solo a la primer cuota"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   71
         ToolTipText     =   "Aplica todo el seguro en la primer cuota del credito"
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox TxtAlicuotaSeguros 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         MaxLength       =   7
         TabIndex        =   70
         Tag             =   "N"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtSeguroFijo 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         MaxLength       =   7
         TabIndex        =   69
         Tag             =   "N"
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Calcular segun Credimaco (con alicuota)"
         Height          =   255
         Left            =   720
         TabIndex        =   78
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label12 
         Caption         =   "Importe fijo de seguros por cuota"
         Height          =   255
         Left            =   720
         TabIndex        =   77
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label24 
         Caption         =   "Importe fijo a dividir entre las cuotas del credito"
         Height          =   255
         Left            =   720
         TabIndex        =   76
         Top             =   1440
         Width           =   3855
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
         TabIndex        =   75
         Top             =   2160
         Width           =   8175
      End
   End
   Begin VB.Frame FrameGastos 
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   240
      TabIndex        =   28
      Top             =   3840
      Width           =   6975
      Begin VB.TextBox TxtCapInt 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   45
         Tag             =   "N"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox TxtFuncIntnocap 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   44
         Tag             =   "N"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox TxtcapNoint 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   43
         Tag             =   "N"
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton PorccapNoint 
         Caption         =   "Porcentaje en función del capital y no del interés:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   4335
      End
      Begin VB.OptionButton PorcIntNoCap 
         Caption         =   "Porcentaje en función del interés y no del capital:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   4335
      End
      Begin VB.OptionButton PorcCapint 
         Caption         =   "Porcentaje en función del capital + intereses:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2040
         Width           =   4335
      End
      Begin VB.TextBox TxtImporteGastosBis 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   39
         Tag             =   "NA"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtImporteGastosFijoBis 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   38
         Tag             =   "NA"
         ToolTipText     =   "Importe que se agrega al 1º vto"
         Top             =   600
         Width           =   840
      End
      Begin VB.TextBox TxtImporteGastosFijo 
         BackColor       =   &H80000013&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   36
         ToolTipText     =   "Importe que se agrega al 1º vto"
         Top             =   600
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.OptionButton OptionGastosFijos 
         Caption         =   "Importe fijo por cuota  (Credimaco)                               $:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   4575
      End
      Begin VB.CheckBox CheckGastosCuota2 
         Caption         =   "Aplicar gastos desde la cuota 2 en adelante"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Carga los gastos desde la cuota Nº 2 en adelante"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.CheckBox CheckGastosCuota1 
         Caption         =   "Aplicar todo el gasto en la primer cuota"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2400
         Width           =   3135
      End
      Begin VB.CheckBox CheckNoAplicarGastosRefinanciacion 
         Caption         =   "No aplicar gastos a las refinanciaciones"
         Height          =   255
         Left            =   3360
         TabIndex        =   32
         ToolTipText     =   "Si marca la casilla no se aplicaran gastos a los planes refinanciados"
         Top             =   2400
         Width           =   3135
      End
      Begin VB.TextBox TxtImporteGastos 
         BackColor       =   &H80000013&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   31
         Tag             =   "N"
         ToolTipText     =   "Importe fijo a dividir entre las cuotas de cada credito"
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton OptionImporteGastos 
         Caption         =   "Importe fijo a dividir entre las cuotas del credito           $:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Permite un importe fijo de gastos que se divide entre las cuotas"
         Top             =   960
         Width           =   4335
      End
      Begin VB.CheckBox CheckAplicarGastos 
         Caption         =   "Aplicar gastos"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "Aplica gastos administrativos a los creditos nuevos"
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.UpDown UpDown4 
         Height          =   285
         Left            =   5400
         TabIndex        =   37
         Top             =   1320
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtcapNoint"
         BuddyDispid     =   196626
         OrigLeft        =   7320
         OrigTop         =   3120
         OrigRight       =   7575
         OrigBottom      =   3375
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown7 
         Height          =   285
         Left            =   5400
         TabIndex        =   46
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtImporteGastosFijo"
         BuddyDispid     =   196632
         OrigLeft        =   7560
         OrigTop         =   960
         OrigRight       =   7815
         OrigBottom      =   1215
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown5 
         Height          =   285
         Left            =   5415
         TabIndex        =   47
         Top             =   960
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtImporteGastos"
         BuddyDispid     =   196637
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
         Left            =   5415
         TabIndex        =   48
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtFuncIntnocap"
         BuddyDispid     =   196625
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
         Left            =   5415
         TabIndex        =   49
         Top             =   2040
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtCapInt"
         BuddyDispid     =   196624
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
   Begin VB.Frame FrameOtorgamiento 
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   240
      TabIndex        =   50
      Top             =   3840
      Visible         =   0   'False
      Width           =   7095
      Begin MSComCtl2.UpDown UpDown15 
         Height          =   285
         Left            =   3736
         TabIndex        =   64
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtImporteOtorgamiento"
         BuddyDispid     =   196611
         OrigLeft        =   7560
         OrigTop         =   960
         OrigRight       =   7815
         OrigBottom      =   1215
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.OptionButton OptCapNoint 
         Caption         =   "Porcentaje en función del capital y no del interés:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   61
         ToolTipText     =   "Permite un importe fijo de gastos que se divide entre las cuotas"
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox TxtImporteOtorgamiento 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         MaxLength       =   7
         TabIndex        =   60
         Tag             =   "N"
         ToolTipText     =   "Importe de gastos de otorgamiento"
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton OptIntNoCap 
         Caption         =   "Porcentaje en función del interés y no del capital:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   59
         ToolTipText     =   "Permite un importe fijo de gastos que se divide entre las cuotas"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CheckBox CheckOtorgamiento1 
         Caption         =   "Cargar otorgamiento solo a la primer cuota"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   58
         ToolTipText     =   "Si el cargo de otorgamiento se carga solo a la primer cuota o a todas en partes iguales"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CheckBox CheckAplicarOtorgamiento 
         Caption         =   "Aplicar Cargos de Otorgamiento"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         ToolTipText     =   "Si aplica gastos de otorgamiento"
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox CheckNoAplicarOtRefin 
         Caption         =   "No incluir en las refinanciaciones"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   2640
         Width           =   2655
      End
      Begin VB.OptionButton OptCapInt 
         Caption         =   "Porcentaje en función del capital + intereses:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   55
         ToolTipText     =   "Permite un importe fijo de gastos que se divide entre las cuotas"
         Top             =   1560
         Width           =   4335
      End
      Begin VB.TextBox TxtPorCapNoInt 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   54
         Tag             =   "N"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TxtPorIntNoCap 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   53
         Tag             =   "N"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TxtCapmasInt 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   52
         Tag             =   "N"
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton OptImpOtor 
         Caption         =   "Importe de otorgamiento $:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   51
         ToolTipText     =   "Permite un importe fijo de gastos que se divide entre las cuotas"
         Top             =   480
         Value           =   -1  'True
         Width           =   2415
      End
      Begin MSComCtl2.UpDown UpDown11 
         Height          =   285
         Left            =   5655
         TabIndex        =   65
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtPorCapNoInt"
         BuddyDispid     =   196617
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
         Left            =   5655
         TabIndex        =   66
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtPorIntNoCap"
         BuddyDispid     =   196618
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
         Left            =   5655
         TabIndex        =   67
         Top             =   1560
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtCapmasInt"
         BuddyDispid     =   196619
         OrigLeft        =   7560
         OrigTop         =   960
         OrigRight       =   7815
         OrigBottom      =   1215
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.Label Label33 
         Caption         =   "(Por ejemplo:20,66)"
         Height          =   255
         Left            =   4680
         TabIndex        =   63
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "(*)Si no lo carga a la primer cuota, el importe se dividira entre las cuotas en partes iguales."
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   62
         Top             =   2040
         Width           =   6135
      End
   End
   Begin VB.CommandButton CmdRefrescar 
      Caption         =   "Refrescar"
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      ToolTipText     =   "Refresca los datos de la pantalla por si hubo cambios en otra PC en red"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de planes"
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5295
      Begin MSComctlLib.ListView lv 
         Height          =   2895
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Lista de planes disponibles"
         Top             =   240
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   5106
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Orden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   8820
         EndProperty
      End
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      ToolTipText     =   "Cierra la pantalla o cancela una operacion de agregado o modificacion"
      Top             =   6600
      Width           =   1425
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      ToolTipText     =   "Graba los datos de un plan"
      Top             =   6120
      Width           =   1425
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      ToolTipText     =   "Permite borrar el plan seleccionado"
      Top             =   5520
      Width           =   1425
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      ToolTipText     =   "Permite modificar los datos del plan seleccionado"
      Top             =   4920
      Width           =   1425
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      ToolTipText     =   "Permite agregar los datos de un nuevo plan"
      Top             =   4320
      Width           =   1425
   End
   Begin VB.Frame fmeDatos 
      ForeColor       =   &H00FF0000&
      Height          =   3225
      Left            =   5520
      TabIndex        =   16
      Top             =   120
      Width           =   5295
      Begin VB.TextBox TxtMensaje 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2520
         Width           =   3495
      End
      Begin VB.ComboBox ComboTem 
         Height          =   315
         ItemData        =   "FrmPlanesAbm.frx":0A02
         Left            =   1680
         List            =   "FrmPlanesAbm.frx":0A0C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Muestra dos formas de calcular el TEM"
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox TxtTasa1 
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "N"
         ToolTipText     =   "Tasa anual de financiacion"
         Top             =   1080
         Width           =   705
      End
      Begin VB.ComboBox ComboVencimientos 
         Height          =   315
         ItemData        =   "FrmPlanesAbm.frx":0A25
         Left            =   1680
         List            =   "FrmPlanesAbm.frx":0A38
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Frecuencia de vencimientos de cuotas (MENSUAL,DIARIO, etc)"
         Top             =   1440
         Width           =   2535
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   2400
         TabIndex        =   21
         Top             =   1800
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtTasa2"
         BuddyDispid     =   196654
         OrigLeft        =   1800
         OrigTop         =   720
         OrigRight       =   2055
         OrigBottom      =   975
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   720
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "TxtCuotas"
         BuddyDispid     =   196653
         OrigLeft        =   1920
         OrigTop         =   1080
         OrigRight       =   2175
         OrigBottom      =   1335
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox TxtNombre 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         ToolTipText     =   "Descripcion del plan"
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox TxtCuotas 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "N"
         ToolTipText     =   "Cantidad de cuotas del plan"
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox TxtTasa2 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "N"
         ToolTipText     =   "Tasa anual de financiacion"
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox CheckPredeterminada 
         Alignment       =   1  'Right Justify
         Caption         =   "Plan Inactivo:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Indica si el plan aparecera o no en la lista de planes en las demas pantallas"
         Top             =   2880
         Width           =   1695
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Top             =   1080
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtTasa1"
         BuddyDispid     =   196650
         OrigLeft        =   1800
         OrigTop         =   720
         OrigRight       =   2055
         OrigBottom      =   975
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Formula TEM:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "%"
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "Tasa T.N.A:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Vencimiento:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Cuotas:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Tasa T.E.M:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Nombre del plan:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1200
      End
   End
   Begin MSComctlLib.TabStrip TabStripopciones 
      Height          =   3495
      Left            =   120
      TabIndex        =   27
      Top             =   3480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6165
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gastos Administrativos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gastos Otorgamientos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seguros"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPlanesAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AQUI SE AGREGAN LOS PLANES DE CREDITOS QUE LUEGO SE SELECCIONAN EN LA
'PANTALLA DE REGISTRAR CREDITOS

Private Sub CheckAplicarGastos_Click()
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
   'UpDown3.Enabled = False
   UpDown7.Enabled = False
   'UpDown1.Enabled = False
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
tratarerrores "Error aplicando Planes"
End Sub

Private Sub CheckAplicarOtorgamiento_Click()
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

Private Sub CheckAplicarSeguro_Click()
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

Private Sub Form_Load()
On Error GoTo merror

Call RefrescarOpcionesSistema
Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de planes"
End Sub
Private Sub CmdCerrar_Click()
If TipoEdicion = "C" Then
   Unload Me
Else
   TipoEdicion = "C"
   Call SetearEntorno
   Call CargarDatos
   lv.SetFocus
End If
End Sub
Private Sub CmdRefrescar_Click()
'refresca la pantalla
Call RefrescarOpcionesSistema
Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno
End Sub
Private Sub ComboTem_Click()
Call CalcularTem
End Sub
Private Sub CargarLista()
'carga el listview con los planes
Dim rec As rdoResultset
Dim sql As String
Dim Nitem As ListItem
On Error GoTo merror
    
 sql = "SELECT IDplan,Nombre AS 'plan' " & _
       "FROM planes " & _
       "ORDER BY planes.idplan"

Set rec = cnSQL.OpenResultset(sql)

lv.ListItems.Clear
    
If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lv.ListItems.Add(, , Format(rec.rdoColumns("Idplan"), "000"))
      Nitem.SubItems(1) = rec.rdoColumns("plan") & vbNullString
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error cargando la lista de planes"
End Sub
Private Function PuedoBorrarPlan(ByVal IdPlan As Long) As Boolean
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

PuedoBorrarPlan = True

'verifico en tabla creditos
sql = "select creditos.idplan " & _
      "from creditos " & _
      "where idplan=" & CLng(IdPlan)
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idplan")) Then
      PuedoBorrarPlan = False
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion PuedoBorrarPlan"
End Function
Private Sub cmdborrar_Click()
Dim sql As String
On Error GoTo merror

If Not VerificarSeleccionLista(lv) Then Exit Sub

'primer chequeo
If Not ExistePlan(lv.SelectedItem) Then
   MsgE "El plan no existe"
   Exit Sub
End If

If Not PuedoBorrarPlan(lv.SelectedItem) Then
   MsgE "No se puede borrar el plan (tiene creditos relacionados)"
   Exit Sub
End If

If Not MsgP("¿Confirma el borrado del plan seleccionado?") Then Exit Sub

'otras validaciones
If Not ExistePlan(lv.SelectedItem) Then
   MsgE "El plan no existe"
   Exit Sub
End If

If Not PuedoBorrarPlan(lv.SelectedItem) Then
   MsgE "No se puede borrar el plan (tiene creditos relacionados)"
   Exit Sub
End If

'inicio de transaccion
cnSQL.BeginTrans

sql = "delete from planes WHERE idplan=" & CLng(lv.SelectedItem)

cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

MsgI "El plan fue borrado"
lv.SetFocus
   
Exit Sub
merror:
tratarerrores "Error borrando planes"
End Sub
Private Sub cmdGrabar_Click()
Dim sql As String
Dim Mensaje As String
Dim IdPlan As Long
Dim DiasVencimiento As Long
Dim CalculoTem As Long
On Error GoTo merror

If Not datosok() Then Exit Sub

DiasVencimiento = 0

If ComboVencimientos.Text = "MENSUAL" Then
   DiasVencimiento = 30
End If
If ComboVencimientos.Text = "BIMESTRAL" Then
   DiasVencimiento = 60
End If
If ComboVencimientos.Text = "DIARIO" Then
   DiasVencimiento = 1
End If
If ComboVencimientos.Text = "SEMANAL" Then
   DiasVencimiento = 7
End If
If ComboVencimientos.Text = "QUINCENAL" Then
   DiasVencimiento = 15
End If

If ComboTem.ListIndex = 0 Then
   CalculoTem = 1
Else
   CalculoTem = 2
End If
    
If TipoEdicion = "N" Then
   
   If Not DatosGastosOk() Then Exit Sub
   
   If Not MsgP("¿Confirma el nuevo plan?") Then Exit Sub
   
   IdPlan = UltimoId("idplan", "planes") + 1
   
   'otras validaciones
   If ExistePlan(IdPlan) Then
      MsgE "El plan ya existe"
      Exit Sub
   End If

   'inicio de transaccion
   cnSQL.BeginTrans
   sql = "INSERT INTO planes (Idplan,nombre,tasa1,tasa2,cantcuotas,diasvencimiento,predeterminada,calculotem,aplicargastos,Noaplicargastosrefinanciacion,aplicargastoscuota1,aplicargastoscuota2,importegastosfijos,importegastos,PorcentajeCapitalyNoInt,PorcentajefuncNoCapital,PorcentajeCapitalInteres,aplicarotorgamiento,importeotorgamiento,noaplicarotorrefin,aplicarotorgamientocuota1,OtorCapNoInt,OtorIntNoCap,OtorCapmasInt,noaplicarsegurosrefinanciacion,aplicarseguroscuota1,aplicarseguro,importeseguro,alicuotaseguros,importesegurosfijos) " & _
         "VALUES (" & CLng(IdPlan) & ",'" & CStr(TxtNombre.Text) & "'," & ConvertirDblSql(CDbl(TxtTasa1.Text)) & "," & ConvertirDblSql(CDbl(TxtTasa2.Text)) & "," & CLng(TxtCuotas.Text) & "," & CLng(DiasVencimiento) & "," & CheckPredeterminada.Value & "," & CLng(CalculoTem) & "," & CheckAplicarGastos.Value & "," & CheckNoAplicarGastosRefinanciacion.Value & "," & CheckGastosCuota1.Value & "," & CheckGastosCuota2.Value & ",'" & TxtImporteGastosFijo.Text & "','" & TxtImporteGastos.Text & "'," & ConvertirDblSql(CDbl(TxtcapNoint.Text)) & "," & ConvertirDblSql(CDbl(TxtFuncIntnocap.Text)) & "," & ConvertirDblSql(CDbl(TxtCapInt.Text)) & "" & _
         "," & CheckAplicarOtorgamiento.Value & "," & ConvertirDblSql(CCur(TxtImporteOtorgamiento.Text)) & "," & CheckNoAplicarOtRefin.Value & "," & CheckOtorgamiento1.Value & "," & ConvertirDblSql(CDbl(TxtPorCapNoInt)) & "," & ConvertirDblSql(CDbl(TxtPorIntNoCap)) & "," & ConvertirDblSql(CDbl(TxtCapmasInt)) & "," & CheckNoAplicarSegurosRefinanciacion.Value & "," & CheckSegurosCuota1.Value & "," & CheckAplicarSeguro.Value & "," & ConvertirDblSql(CCur(TxtImporteSeguro.Text)) & "," & ConvertirDblSql(CDbl(TxtAlicuotaSeguros.Text)) & "," & ConvertirDblSql(CCur(TxtSeguroFijo.Text)) & ") "
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El plan fue agregado"
   
   Call CargarLista
   Call CargarDatos

Else
   'primer chequeo
   If Not ExistePlan(lv.SelectedItem) Then
      MsgE "El plan no existe"
      Exit Sub
   End If
    
   If Not DatosGastosOk() Then Exit Sub
   If Not DatosOtorgamientoOk() Then Exit Sub
   If Not DatosSegurosOk() Then Exit Sub
   
   If Not MsgP("¿Confirma la modificacion del plan seleccionado?") Then Exit Sub
   
   'segundo chequeo
   If Not ExistePlan(lv.SelectedItem) Then
      MsgE "El plan no existe"
      Exit Sub
   End If

   'inicio de transaccion
   cnSQL.BeginTrans
   
   
   sql = "UPDATE planes SET nombre='" & CStr(TxtNombre.Text) & "',tasa1=" & ConvertirDblSql(CDbl(TxtTasa1.Text)) & _
         ",tasa2=" & ConvertirDblSql(CDbl(TxtTasa2.Text)) & ",cantcuotas=" & CLng(TxtCuotas.Text) & _
         ",diasvencimiento=" & CLng(DiasVencimiento) & ",predeterminada=" & CheckPredeterminada.Value & _
         ",calculotem=" & CLng(CalculoTem) & ",aplicargastos=" & CheckAplicarGastos.Value & _
         ",Noaplicargastosrefinanciacion=" & CheckNoAplicarGastosRefinanciacion.Value & _
         ",aplicargastoscuota1=" & CheckGastosCuota1.Value & ",aplicargastoscuota2=" & CheckGastosCuota2.Value & _
         ",importegastosfijos=" & ConvertirDblSql(CCur(TxtImporteGastosFijo.Text)) & _
         ",importegastos=" & ConvertirDblSql(CCur(TxtImporteGastos.Text)) & _
         ",PorcentajeCapitalyNoInt = " & ConvertirDblSql(CDbl(TxtcapNoint.Text)) & "" & _
         ",PorcentajefuncNoCapital = " & ConvertirDblSql(CDbl(TxtFuncIntnocap.Text)) & "" & _
         ",PorcentajeCapitalInteres = " & ConvertirDblSql(CDbl(TxtCapInt.Text)) & "" & _
         ",aplicarotorgamiento=" & CheckAplicarOtorgamiento.Value & _
         ",importeotorgamiento=" & ConvertirDblSql(CCur(TxtImporteOtorgamiento.Text)) & _
         ",noaplicarotorrefin=" & CheckNoAplicarOtRefin.Value & ",aplicarotorgamientocuota1=" & CheckOtorgamiento1.Value & _
         ",OtorCapNoInt=" & ConvertirDblSql(CDbl(TxtPorCapNoInt)) & ",OtorIntNoCap=" & ConvertirDblSql(CDbl(TxtPorIntNoCap)) & _
         ",OtorCapmasInt=" & ConvertirDblSql(CDbl(TxtCapmasInt)) & _
         ",aplicarseguro=" & CheckAplicarSeguro.Value & _
         ",aplicarseguroscuota1=" & CheckSegurosCuota1.Value & _
         ",noaplicarsegurosrefinanciacion=" & CheckNoAplicarSegurosRefinanciacion.Value & _
         ",importeseguro=" & ConvertirDblSql(CCur(TxtImporteSeguro.Text)) & _
         ",alicuotaseguros=" & ConvertirDblSql(CDbl(TxtAlicuotaSeguros.Text)) & _
         ",importesegurosfijos=" & ConvertirDblSql(CCur(TxtSeguroFijo.Text)) & _
         " WHERE Idplan=" & CLng(lv.SelectedItem)
         
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El plan fue modificado"
   
   lv.SelectedItem.ListSubItems(1).Text = TxtNombre.Text & vbNullString

End If




TipoEdicion = "C"
Call SetearEntorno

MsgI Mensaje

lv.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando planes"
End Sub
Private Sub cmdModificar_Click()
'predispone a modificar solo si hay datos en el listview y hay seleccion
   
If Not VerificarSeleccionLista(lv) Then Exit Sub

TipoEdicion = "M"
Call SetearEntorno

End Sub
Private Sub cmdnuevo_Click()
TipoEdicion = "N"
Call SetearEntorno
End Sub
Private Sub CargarDatos()
'Pone los datos del item seleccionado del listview en los campos de abajo
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror
    
    
If Not VerificarSeleccionLista(lv) Then Exit Sub
        
sql = "SELECT planes.IDplan,planes.nombre AS 'plan'," & _
      "planes.tasa1,planes.tasa2,planes.cantcuotas,planes.diasvencimiento,planes.tipovencimiento,planes.predeterminada,planes.calculotem, " & _
      "planes.Aplicargastos,planes.Noaplicargastosrefinanciacion,planes.aplicargastoscuota1,planes.aplicargastoscuota2,planes.importegastosfijos,planes.importegastos,planes.PorcentajeCapitalyNoInt,planes.PorcentajefuncNoCapital,planes.PorcentajeCapitalInteres," & _
      "planes.aplicarotorgamiento,planes.importeotorgamiento,planes.noaplicarotorrefin,planes.aplicarotorgamientocuota1,planes.OtorCapNoInt,planes.OtorIntNoCap,planes.OtorCapmasInt," & _
      "planes.aplicarseguro,planes.aplicarseguroscuota1,planes.noaplicarsegurosrefinanciacion,planes.importeseguro,planes.alicuotaseguros,planes.importesegurosfijos" & _
      " FROM planes " & _
      "WHERE Idplan=" & CLng(lv.SelectedItem)

Set rec = cnSQL.OpenResultset(sql)
        
If Not rec.EOF Then
   TxtNombre.Text = rec.rdoColumns("plan") & vbNullString
   TxtTasa1.Text = rec.rdoColumns("tasa1") & vbNullString
   TxtTasa2.Text = rec.rdoColumns("tasa2") & vbNullString
   
   TxtCuotas.Text = rec.rdoColumns("cantcuotas") & vbNullString
   
   If rec.rdoColumns("diasvencimiento") = 30 Then
      ComboVencimientos.Text = "MENSUAL"
   End If
   If rec.rdoColumns("diasvencimiento") = 60 Then
      ComboVencimientos.Text = "BIMESTRAL"
   End If
   If rec.rdoColumns("diasvencimiento") = 1 Then
      ComboVencimientos.Text = "DIARIO"
   End If
   If rec.rdoColumns("diasvencimiento") = 7 Then
      ComboVencimientos.Text = "SEMANAL"
   End If
   If rec.rdoColumns("diasvencimiento") = 15 Then
      ComboVencimientos.Text = "QUINCENAL"
   End If
     
   If rec.rdoColumns("predeterminada") Then
      CheckPredeterminada.Value = 1
   Else
      CheckPredeterminada.Value = 0
   End If
   
   If rec.rdoColumns("calculotem") = 1 Then
      ComboTem.ListIndex = 0
   Else
      ComboTem.ListIndex = 1
   End If
   
   'Solapa Gastos Administrativos
   TxtImporteGastosFijo = CCur(rec.rdoColumns("importegastosfijos"))
   If Val(TxtImporteGastosFijo) > 0 Then
    OptionGastosFijos.Value = 1
   End If
   
   TxtImporteGastos = CCur(rec.rdoColumns("importegastos"))
   If Val(TxtImporteGastos) > 0 Then
    OptionImporteGastos.Value = 1
   End If
   TxtcapNoint = rec.rdoColumns("PorcentajeCapitalyNoInt")
   If Val(TxtcapNoint) > 0 Then
    PorccapNoint.Value = 1
   End If
   
   TxtFuncIntnocap = rec.rdoColumns("PorcentajefuncNoCapital")
   If Val(TxtFuncIntnocap) > 0 Then
    PorcIntNoCap.Value = 1
   End If
   
   TxtCapInt = rec.rdoColumns("PorcentajeCapitalInteres")
   If Val(TxtCapInt) > 0 Then
    PorcCapint.Value = 1
   End If
   
   If rec.rdoColumns("aplicargastoscuota1") Then
    CheckGastosCuota1 = 1
   Else
    CheckGastosCuota1 = 0
   End If
   
   If rec.rdoColumns("aplicargastoscuota2") Then
    CheckGastosCuota2 = 1
   Else
    CheckGastosCuota2 = 0
   End If
   If rec.rdoColumns("noaplicargastosrefinanciacion") Then
    CheckNoAplicarGastosRefinanciacion = 1
   Else
    CheckNoAplicarGastosRefinanciacion = 0
   End If
    If rec.rdoColumns("aplicargastos") Then
        CheckAplicarGastos = 1
    Else
        CheckAplicarGastos = 0
    End If
    
   'solapa otorgamientos
   TxtImporteOtorgamiento = CCur(rec.rdoColumns("importeotorgamiento"))
   If Val(TxtImporteOtorgamiento) > 0 Then
    OptImpOtor.Value = 1
   End If
   
   TxtPorCapNoInt = CCur(rec.rdoColumns("OtorCapNoInt"))
   If Val(TxtPorCapNoInt) > 0 Then
    OptCapNoint.Value = 1
   End If
   
   TxtPorIntNoCap = CCur(rec.rdoColumns("OtorIntNoCap"))
   If Val(TxtPorIntNoCap) > 0 Then
    OptIntNoCap.Value = 1
   End If
   
   TxtCapmasInt = CCur(rec.rdoColumns("OtorCapmasInt"))
   If Val(TxtCapmasInt) > 0 Then
    OptCapInt.Value = 1
   End If
         
   If rec.rdoColumns("aplicarotorgamientocuota1") Then
    CheckOtorgamiento1 = 1
   Else
    CheckOtorgamiento1 = 0
   End If
   
  If rec.rdoColumns("noaplicarotorrefin") Then
    CheckNoAplicarOtRefin = 1
   Else
    CheckNoAplicarOtRefin = 0
   End If
         
   If rec.rdoColumns("aplicarotorgamiento") Then
        CheckAplicarOtorgamiento = 1
    Else
        CheckAplicarOtorgamiento = 0
    End If
   
   'Solapa Seguros
    TxtAlicuotaSeguros = CCur(rec.rdoColumns("alicuotaseguros"))
    TxtImporteSeguro = CCur(rec.rdoColumns("ImporteSeguro"))
    TxtSeguroFijo = CCur(rec.rdoColumns("importesegurosfijos"))
   If rec.rdoColumns("aplicarseguro") Then
        CheckAplicarSeguro = 1
    Else
        CheckAplicarSeguro = 0
    End If
    If rec.rdoColumns("aplicarseguroscuota1") Then
        CheckSegurosCuota1 = 1
    Else
        CheckSegurosCuota1 = 0
    End If
      If rec.rdoColumns("noaplicarsegurosrefinanciacion") Then
        CheckNoAplicarSegurosRefinanciacion = 1
    Else
        CheckNoAplicarSegurosRefinanciacion = 0
    End If
    
End If
        
Exit Sub
merror:
tratarerrores "Error cargando datos de planes"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(TxtNombre.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el nombre/descripcion del plan"
   TxtNombre.SetFocus
   Exit Function
End If

'valido las cuotas
If Trim(TxtCuotas.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar la cantidad de cuotas del plan"
   TxtCuotas.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtCuotas.Text) Then
   datosok = False
   MsgE "La cantidad de cuotas debe ser numerica"
   TxtCuotas.SetFocus
   Exit Function
End If
If CDbl(TxtCuotas.Text) < 0 Then
   datosok = False
   MsgE "La cantidad de cuotas debe ser mayor a cero"
   TxtCuotas.SetFocus
   Exit Function
End If

'valido la tasa
If Trim(TxtTasa1.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar la tasa de financiacion"
   TxtTasa1.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtTasa1.Text) Then
   datosok = False
   MsgE "La tasa debe ser numerica"
   TxtTasa1.SetFocus
   Exit Function
End If
If CDbl(TxtTasa1.Text) < 0 Then
   datosok = False
   MsgE "La tasa de financiacion debe ser mayor o igual a cero"
   TxtTasa1.SetFocus
   Exit Function
End If

If ComboVencimientos.Text = "" Then
   datosok = False
   MsgE "Debe seleccionar los dias de vencimiento"
   ComboVencimientos.SetFocus
   Exit Function
End If

Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-Planes"
End Function
Private Sub SetearEntorno()
On Error GoTo merror

Select Case TipoEdicion
       Case "C"
            fmeDatos.Enabled = False
            FrameSeguroVida.Enabled = False
            FrameGastos.Enabled = False
            FrameOtorgamiento.Enabled = False
            cmdGrabar.Enabled = False
            cmdNuevo.Enabled = True
            CmdRefrescar.Enabled = True
            If lv.ListItems.Count > 0 Then
               cmdModificar.Enabled = True
               cmdBorrar.Enabled = True
            Else
               cmdModificar.Enabled = False
               cmdBorrar.Enabled = False
               Call LimpiarCampos(Me)
            End If
            cmdcerrar.Caption = "&Cerrar"
            lv.Enabled = True
            Call ColorCyan(Me)
        Case "M"
        
            'solapas
            FrameSeguroVida.Enabled = True
            FrameGastos.Enabled = True
            FrameOtorgamiento.Enabled = True
            If OptionGastosFijos.Value = True Then
                Call OptionGastosFijos_Click
            End If
            
            If OptionImporteGastos.Value = True Then
               Call OptionImporteGastos_Click
            End If
            
            If PorccapNoint.Value = True Then
                 Call PorccapNoint_Click
            End If
            
            If PorcIntNoCap.Value = True Then
                 Call PorcIntNoCap_Click
            End If
            
            If PorcCapint.Value = True Then
                 Call PorcCapint_Click
            End If
        
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            cmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            cmdBorrar.Enabled = False
            CmdRefrescar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lv.Enabled = False
            TxtNombre.SetFocus
            Call ColorBlanco(Me)
        Case "N"
            Call LimpiarCampos(Me)
            fmeDatos.Enabled = True
            FrameSeguroVida.Enabled = True
            FrameGastos.Enabled = True
            FrameOtorgamiento.Enabled = True
            cmdGrabar.Enabled = True
            cmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            cmdBorrar.Enabled = False
            CmdRefrescar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lv.Enabled = False
            TxtNombre.SetFocus
            Call ColorBlanco(Me)
    End Select

Exit Sub
merror:
tratarerrores "Error seteando el entorno-PlanesAbm"
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'ordena el listview pero solo si tiene datos
Dim Orden As Integer
    
If lv.ListItems.Count > 1 Then
   lv.SortKey = ColumnHeader.Index - 1
   Orden = lv.SortKey
   lv.SortOrder = Abs(Not lv.SortOrder = 1)
   lv.Sorted = True
End If

End Sub
Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
'dentro de la funcion chequea que haya datos en el listview
Call CargarDatos
End Sub

Private Sub OptCapInt_Click()
On Error GoTo merror

If OptCapInt.Value And FrameOtorgamiento.Enabled = True Then
   
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

If OptCapNoint.Value And FrameOtorgamiento.Enabled = True Then
   
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

If OptImpOtor.Value And FrameOtorgamiento.Enabled = True Then
   
   
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

If OptIntNoCap.Value And FrameOtorgamiento.Enabled = True Then
   
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

If OptionGastosFijos.Value And FrameGastos.Enabled = True Then
   
   UpDown7.Enabled = True
   UpDown5.Enabled = False
   UpDown8.Enabled = False
   UpDown9.Enabled = False
   UpDown4.Enabled = False
   
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

Private Sub OptionImporteGastos_Click()
On Error GoTo merror

If OptionImporteGastos.Value And FrameGastos.Enabled = True Then
   UpDown4.Enabled = False
   UpDown5.Enabled = True
   UpDown7.Enabled = False
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

Private Sub PorcCapint_Click()
On Error GoTo merror

If PorcCapint.Value And FrameGastos.Enabled = True Then
   
   UpDown5.Enabled = False
   UpDown4.Enabled = False
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

If PorccapNoint.Value And FrameGastos.Enabled = True Then
   
   UpDown7.Enabled = False
   UpDown8.Enabled = False
   UpDown9.Enabled = False
   UpDown5.Enabled = False
   UpDown4.Enabled = True
   
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

If PorcIntNoCap.Value And FrameGastos.Enabled = True Then
   
   UpDown5.Enabled = False
   UpDown4.Enabled = False
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


FrameGastos.Visible = False
FrameOtorgamiento.Visible = False
FrameSeguroVida.Visible = False

'Gastos Administrayivos
If TabStripopciones.SelectedItem.Index = 1 Then
   FrameGastos.Visible = True
End If

'Gastos Otorgamiento
If TabStripopciones.SelectedItem.Index = 2 Then
   FrameOtorgamiento.Visible = True
End If

'Seguro de vida
If TabStripopciones.SelectedItem.Index = 3 Then
   FrameSeguroVida.Visible = True
End If



Exit Sub
merror:
tratarerrores "Error seleccionando Opciones"
End Sub

Private Sub TxtImporteGastos_Change()
TxtImporteGastosBis.Text = Format(CCur(TxtImporteGastos.Text) / 10, "0.00")
End Sub

Private Sub TxtImporteGastosFijo_Change()
If TxtImporteGastosFijo.Text <> "" Then
  TxtImporteGastosFijoBis.Text = Format(CCur(TxtImporteGastosFijo.Text) / 10, "0.00")
End If
End Sub

Private Sub TxtNombre_LostFocus()
TxtNombre.Text = UCase(Trim(TxtNombre.Text))
End Sub
Private Sub TxtTasa1_Change()
'calculo la tasa TEM (MENSUAL, ETC)
Call CalcularTem
End Sub
Private Sub ComboVencimientos_Change()
Call CalcularTem
End Sub
Private Sub ComboVencimientos_Click()
Call CalcularTem
End Sub
Private Sub CalcularTem()
'calcula la tasa TEM y la aplica cuando cambia la tasa1 tna y cuando cambia
'el tipo de vencimiento mensual, etc
Dim Dias As Long
Dim Tem As Double
On Error GoTo merror

If Trim(TxtTasa1.Text) = "" Then Exit Sub
If Not IsNumeric(TxtTasa1.Text) Then Exit Sub
If CDbl(TxtTasa1.Text) < 0 Then Exit Sub
If ComboVencimientos.Text = "" Then Exit Sub

TxtTasa2.Text = 0

Dias = ComboVencimientos.ListIndex

'si calculo aplicando los 365
If ComboTem.ListIndex = 0 Then
   'si es mensual
   If Dias = 0 Then
      TxtTasa2.Text = CDbl(TxtTasa1.Text) / 365 * 30
      TxtMensaje.Text = CStr(TxtTasa1.Text) & " / 365 x 30"
   End If
   'diario
   If Dias = 1 Then
      TxtTasa2.Text = CDbl(TxtTasa1.Text) / 365 * 1
      TxtMensaje.Text = CStr(TxtTasa1.Text) & " / 365 x 1"
   End If
   'semanal
   If Dias = 2 Then
      TxtTasa2.Text = CDbl(TxtTasa1.Text) / 365 * 7
      TxtMensaje.Text = CStr(TxtTasa1.Text) & " / 365 x 7"
   End If
   'quincenal
   If Dias = 3 Then
      TxtTasa2.Text = CDbl(TxtTasa1.Text) / 365 * 15
      TxtMensaje.Text = CStr(TxtTasa1.Text) & " / 365 x 15"
   End If
   'bimestral
   If Dias = 4 Then
      TxtTasa2.Text = CDbl(TxtTasa1.Text) / 365 * 60
      TxtMensaje.Text = CStr(TxtTasa1.Text) & " / 365 x 60"
   End If
Else
   'si es mensual
   If Dias = 0 Then
      TxtTasa2.Text = CDbl(TxtTasa1.Text) / 12
      TxtMensaje.Text = CStr(TxtTasa1.Text) & " / 12"
   End If
   'diario
   If Dias = 1 Then
      TxtTasa2.Text = CDbl(TxtTasa1.Text) / 12 / 30
      TxtMensaje.Text = CStr(TxtTasa1.Text) & " / 12 / 30"
   End If
   'semanal
   If Dias = 2 Then
      TxtTasa2.Text = CDbl(TxtTasa1.Text) / 12 / 30 * 7
      TxtMensaje.Text = CStr(TxtTasa1.Text) & " / 12 / 30 x 7"
   End If
   'quincenal
   If Dias = 3 Then
      TxtTasa2.Text = CDbl(TxtTasa1.Text) / 12 / 30 * 15
      TxtMensaje.Text = CStr(TxtTasa1.Text) & " / 12 / 30 x 15"
   End If
   'bimestral
   If Dias = 4 Then
      TxtTasa2.Text = CDbl(TxtTasa1.Text) / 12 * 2
      TxtMensaje.Text = CStr(TxtTasa1.Text) & " / 12 x 2"
   End If
End If

TxtTasa2.Text = Format(TxtTasa2.Text, "0.00")

Exit Sub
merror:
tratarerrores "Error calculando Tem"
End Sub
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

