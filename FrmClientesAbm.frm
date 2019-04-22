VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmClientesAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Clientes"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   HelpContextID   =   7
   Icon            =   "FrmClientesAbm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameGarante 
      Height          =   1695
      Left            =   240
      TabIndex        =   95
      Top             =   6120
      Width           =   6735
      Begin VB.ComboBox Cmbdiferenciando 
         Height          =   315
         ItemData        =   "FrmClientesAbm.frx":030A
         Left            =   5160
         List            =   "FrmClientesAbm.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   171
         ToolTipText     =   "Tipo Garante"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TxtProfesionGarante 
         Height          =   285
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   36
         ToolTipText     =   "Profesion del garante"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox TxtNacionalidadGarante 
         Height          =   285
         Left            =   720
         MaxLength       =   50
         TabIndex        =   35
         ToolTipText     =   "Nacionalidad del garante"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TxtCuitGarante 
         Height          =   285
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   33
         ToolTipText     =   "Nº de Cuit del garante"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TxtSueldoGarante 
         Height          =   285
         Left            =   720
         MaxLength       =   9
         TabIndex        =   32
         Tag             =   "N"
         ToolTipText     =   "Sueldo del garante"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtDomicilioGarante 
         Height          =   285
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   31
         ToolTipText     =   "Domicilio del garante"
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox TxtDocumentoGarante 
         Height          =   285
         Left            =   720
         MaxLength       =   50
         TabIndex        =   30
         ToolTipText     =   "Nº de documento del garante"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtNombreGarante 
         Height          =   285
         Left            =   720
         MaxLength       =   50
         TabIndex        =   28
         ToolTipText     =   "Nombre del garante"
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtTelefonoGarante 
         Height          =   285
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   34
         ToolTipText     =   "Telefono del garante"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtApellidoGarante 
         Height          =   285
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   29
         ToolTipText     =   "Apellido del garante"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label65 
         Caption         =   "tipo:"
         Height          =   255
         Left            =   4560
         TabIndex        =   172
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label44 
         Caption         =   "Profesion:"
         Height          =   255
         Left            =   2160
         TabIndex        =   126
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label43 
         Caption         =   "Pais:"
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label27 
         Caption         =   "Tel:"
         Height          =   255
         Left            =   4560
         TabIndex        =   102
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label26 
         Caption         =   "Cuit-Cuil:"
         Height          =   255
         Left            =   2160
         TabIndex        =   101
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "Sueldo:"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Domicilio:"
         Height          =   255
         Left            =   2160
         TabIndex        =   99
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "Docum.:"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Apellido:"
         Height          =   255
         Left            =   3000
         TabIndex        =   97
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdRefrescar 
      Caption         =   "Refrescar"
      Height          =   375
      Left            =   7440
      TabIndex        =   56
      ToolTipText     =   "Refresca los datos de la pantalla por si hubo cambios desde otra PC en red"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton CmdExportarClientes 
      Caption         =   "Exportar clientes"
      Height          =   375
      Left            =   7440
      TabIndex        =   61
      ToolTipText     =   "Exporta la lista de clientes a una planilla en la carpeta C:\EXPORTACIONEXCEL\CLIENTES.XLS"
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton CmdResumen 
      Caption         =   "Resumen clientes"
      Height          =   375
      Left            =   7440
      TabIndex        =   63
      ToolTipText     =   "Resumen de clientes"
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtrar por:"
      ForeColor       =   &H00FF0000&
      Height          =   2535
      Left            =   7440
      TabIndex        =   74
      Top             =   0
      Width           =   1815
      Begin VB.OptionButton Option5 
         Caption         =   "Por cupon"
         Height          =   255
         Left            =   120
         TabIndex        =   130
         ToolTipText     =   "Filtra por Nº de cupon (o comprobante)"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton CmdTodos 
         Caption         =   "Mostrar &todos"
         Height          =   255
         Left            =   240
         TabIndex        =   71
         ToolTipText     =   "Muestra todos los clientes"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton CmdFiltrar 
         Caption         =   "Filtrar"
         Height          =   250
         Left            =   240
         TabIndex        =   70
         ToolTipText     =   "Filtra la lista de clientes"
         Top             =   1850
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Por Nº cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         ToolTipText     =   "Filtra por Nº de cliente"
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Por cuit"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         ToolTipText     =   "Filtra clientes por Nº de cuil"
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por documento"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         ToolTipText     =   "Filtra clientes por Nº de documento"
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por apellido"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         ToolTipText     =   "Filtra clientes por apellido"
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox TxtCampo 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   69
         Tag             =   "no"
         ToolTipText     =   "Criterio de busqueda (Apellido, documento, etc)"
         Top             =   1500
         Width           =   1575
      End
   End
   Begin VB.Frame FrameClientes 
      Caption         =   "Lista de clientes:"
      ForeColor       =   &H00FF0000&
      Height          =   1870
      Left            =   120
      TabIndex        =   72
      Top             =   0
      Width           =   7215
      Begin MSComctlLib.ListView lv 
         Height          =   1575
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Muestra la lista de clientes"
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2778
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nº cliente"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   8820
         EndProperty
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir cliente"
      Height          =   375
      Left            =   7440
      TabIndex        =   62
      ToolTipText     =   "Imprime la lista de clientes"
      Top             =   6480
      Width           =   1785
   End
   Begin VB.CommandButton CmdSeleccionar 
      Caption         =   "&Seleccionar cliente"
      Height          =   375
      Left            =   7440
      TabIndex        =   55
      ToolTipText     =   "Selecciona el cliente marcado en la lista y lo carga en otra pantalla"
      Top             =   2640
      Width           =   1785
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   7440
      TabIndex        =   57
      ToolTipText     =   "Permite agregar los datos de un nuevo cliente"
      Top             =   3840
      Width           =   1785
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   7440
      TabIndex        =   58
      ToolTipText     =   "Modifica los datos del cliente seleccionado"
      Top             =   4320
      Width           =   1785
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "B&orrar"
      Height          =   375
      Left            =   7440
      TabIndex        =   59
      ToolTipText     =   "Borra el cliente seleccionado"
      Top             =   4800
      Width           =   1785
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7440
      TabIndex        =   60
      ToolTipText     =   "Graba los datos del cliente"
      Top             =   5280
      Width           =   1785
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   7440
      TabIndex        =   64
      ToolTipText     =   "Cierra la pantalla"
      Top             =   7560
      Width           =   1785
   End
   Begin VB.Frame fmeDatos 
      Caption         =   "Datos  del cliente"
      ForeColor       =   &H00FF0000&
      Height          =   3975
      Left            =   120
      TabIndex        =   73
      Top             =   1920
      Width           =   7215
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   16
         ToolTipText     =   "E-Mail del cliente"
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox TxtCelular 
         Height          =   285
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   15
         ToolTipText     =   "Celular del cliente"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox ComboSexo 
         Height          =   315
         ItemData        =   "FrmClientesAbm.frx":032B
         Left            =   3240
         List            =   "FrmClientesAbm.frx":0335
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtCP 
         Height          =   285
         Left            =   5160
         TabIndex        =   8
         ToolTipText     =   "Codigo postal diferenciado"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtTipoIva 
         Height          =   285
         Left            =   5640
         MaxLength       =   50
         TabIndex        =   22
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox TxtCreditoMaximo 
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   23
         Tag             =   "N"
         Top             =   3120
         Width           =   1080
      End
      Begin VB.TextBox TxtNacionalidad 
         Height          =   285
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   13
         ToolTipText     =   "Nacionalidad del cliente"
         Top             =   1680
         Width           =   2175
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   285
         Left            =   1440
         TabIndex        =   94
         Top             =   1680
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtResidencia"
         BuddyDispid     =   196664
         OrigLeft        =   5520
         OrigTop         =   1680
         OrigRight       =   5775
         OrigBottom      =   1935
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   6600
         TabIndex        =   93
         Top             =   2400
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtSueldo"
         BuddyDispid     =   196673
         OrigLeft        =   4200
         OrigTop         =   2400
         OrigRight       =   4455
         OrigBottom      =   2655
         Increment       =   50
         Max             =   1000000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   4440
         TabIndex        =   92
         Top             =   2400
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtAntiguedad"
         BuddyDispid     =   196674
         OrigLeft        =   2040
         OrigTop         =   2400
         OrigRight       =   2295
         OrigBottom      =   2655
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox TxtActividad 
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   17
         ToolTipText     =   "Profesion del cliente"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox ComboEstadoCivil 
         Height          =   315
         ItemData        =   "FrmClientesAbm.frx":034E
         Left            =   5160
         List            =   "FrmClientesAbm.frx":0364
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Estado civil del cliente"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtCodigoDescuento 
         Height          =   285
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   12
         ToolTipText     =   "Barrio de residencia del cliente"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox TxtNumCBU 
         Height          =   285
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   21
         ToolTipText     =   "Codigo bancario del cliente"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CheckBox CheckFacturaServicio 
         Caption         =   "Facturas de servicios"
         Height          =   195
         Left            =   5040
         TabIndex        =   25
         ToolTipText     =   "Si presento facturas de servicios"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.ComboBox ComboLocalidades 
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Ciudad del cliente"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox TxtResidencia 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   11
         Tag             =   "N"
         ToolTipText     =   "Años de residencia del cliente en la ciudad /pais"
         Top             =   1680
         Width           =   360
      End
      Begin VB.TextBox TxtDomicilio 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Domicilio del cliente (calle y Nº)"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CheckBox CheckReciboSueldo 
         Caption         =   "Recibo de sueldo"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         ToolTipText     =   "Si tiene recibo de sueldo"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox TxtCuil 
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   20
         ToolTipText     =   "Nº de CUIL del cliente (codigo laboral)"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox TxtNumLegajo 
         Height          =   285
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   3
         ToolTipText     =   "Legajo del cliente"
         Top             =   600
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         ToolTipText     =   "Fecha de nacimiento del cliente"
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54591489
         CurrentDate     =   39018
      End
      Begin VB.TextBox TxtObservaciones 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   26
         ToolTipText     =   "Comentarios varios"
         Top             =   3480
         Width           =   4095
      End
      Begin VB.TextBox TxtTelefono 
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   14
         ToolTipText     =   "Nº de telefono del cliente"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox TxtApellido 
         Height          =   285
         Left            =   4560
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Apellido del cliente"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox TxtNombre 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Nombre del cliente"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox TxtSueldo 
         Height          =   285
         Left            =   5640
         MaxLength       =   9
         TabIndex        =   19
         Tag             =   "N"
         ToolTipText     =   "Sueldo del cliente"
         Top             =   2400
         Width           =   945
      End
      Begin VB.TextBox TxtAntiguedad 
         Height          =   285
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   18
         Tag             =   "N"
         ToolTipText     =   "Nº de años de antiguedad laboral del cliente"
         Top             =   2400
         Width           =   480
      End
      Begin VB.TextBox TxtNumDocumento 
         Height          =   285
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Nº de documento del cliente"
         Top             =   600
         Width           =   1455
      End
      Begin MSComCtl2.UpDown UpDown7 
         Height          =   285
         Left            =   2520
         TabIndex        =   128
         Top             =   3120
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtCreditoMaximo"
         BuddyDispid     =   196656
         OrigLeft        =   2400
         OrigTop         =   600
         OrigRight       =   2655
         OrigBottom      =   885
         Increment       =   50
         Max             =   10000000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label49 
         Caption         =   "Mail:"
         Height          =   255
         Left            =   4320
         TabIndex        =   134
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label48 
         Caption         =   "CEL:"
         Height          =   255
         Left            =   2400
         TabIndex        =   133
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label47 
         Caption         =   "Sexo (*):"
         Height          =   255
         Left            =   2520
         TabIndex        =   132
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label46 
         Caption         =   "CP:"
         Height          =   255
         Left            =   4800
         TabIndex        =   131
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Tipo IVA:"
         Height          =   255
         Left            =   4920
         TabIndex        =   129
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label45 
         Caption         =   "Credito maximo $:"
         Height          =   255
         Left            =   120
         TabIndex        =   127
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label42 
         Caption         =   "Pais:"
         Height          =   255
         Left            =   4320
         TabIndex        =   124
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "Profesion:"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Est.civil:"
         Height          =   255
         Left            =   4560
         TabIndex        =   90
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Barrio:"
         Height          =   255
         Left            =   1920
         TabIndex        =   89
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "CBU:"
         Height          =   255
         Left            =   2880
         TabIndex        =   88
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Ciudad (*):"
         Height          =   255
         Left            =   3960
         TabIndex        =   87
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Residencia:"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "CUIL:"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Nº cliente (*):"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Fech.Nacim:"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Sueldo (*):"
         Height          =   255
         Left            =   4800
         TabIndex        =   81
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Antig.laboral:"
         Height          =   255
         Left            =   2880
         TabIndex        =   80
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "DNI (*):"
         Height          =   255
         Left            =   2400
         TabIndex        =   79
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Domicilio (*):"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "TEL:"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido (*):"
         Height          =   255
         Left            =   3720
         TabIndex        =   76
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Nombre    (*):"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Frame FrameObservaciones 
      Height          =   1695
      Left            =   240
      TabIndex        =   135
      Top             =   6120
      Width           =   6735
      Begin VB.TextBox TxtObservaciones1 
         Height          =   1365
         Left            =   240
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   136
         ToolTipText     =   "Nombre del garante"
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame FramePropiedad 
      Height          =   1695
      Left            =   240
      TabIndex        =   103
      Top             =   6120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox TxtMetrosPropiedad 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   39
         ToolTipText     =   "Metros cuadrados de  la propiedad"
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox ComboTipoPropiedad 
         Height          =   315
         ItemData        =   "FrmClientesAbm.frx":03AF
         Left            =   1080
         List            =   "FrmClientesAbm.frx":03C8
         Style           =   2  'Dropdown List
         TabIndex        =   40
         ToolTipText     =   "Tipo de propiedad"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox TxtValuacionPropiedad 
         Height          =   285
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   38
         Tag             =   "N"
         ToolTipText     =   "Valuacion estimada en pesos de la propiedad"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox TxtCatastroPropiedad 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   37
         ToolTipText     =   "Referencia catastral de la propiedad"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label30 
         Caption         =   "Superficie  :"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Tipo            :"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "Valuacion $:"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Nº catastro:"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FrameEmpleador 
      Height          =   1695
      Left            =   240
      TabIndex        =   114
      Top             =   6120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CheckBox CheckJubPen 
         Caption         =   "Jubilado/Pensionado"
         Height          =   255
         Left            =   3720
         TabIndex        =   141
         ToolTipText     =   "Si es Monotributista"
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox CmbMonotributista 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmClientesAbm.frx":0407
         Left            =   5400
         List            =   "FrmClientesAbm.frx":0432
         Style           =   2  'Dropdown List
         TabIndex        =   140
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox CHMonotributista 
         Caption         =   "Monotributista"
         Height          =   255
         Left            =   3720
         TabIndex        =   139
         ToolTipText     =   "Si es Monotributista"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtDomicilioEmpleador 
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   48
         ToolTipText     =   "Domicilio del empleador"
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox TxtTelefonoEmpleador 
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   49
         ToolTipText     =   "Telefono del empleador"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox TxtCuitEmpleador 
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   47
         ToolTipText     =   "Nº de Cuit del empleador"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtEmpresa 
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   46
         ToolTipText     =   "Empresa en la que trabaja el cliente"
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Txtactividad1 
         Height          =   285
         Left            =   4200
         MaxLength       =   25
         TabIndex        =   137
         ToolTipText     =   "Empresa en la que trabaja el cliente"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label50 
         Caption         =   "Actividad:"
         Height          =   255
         Left            =   3480
         TabIndex        =   138
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label37 
         Caption         =   "Domicilio:"
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label36 
         Caption         =   "Telefono:"
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "C.U.I.T:"
         Height          =   255
         Left            =   120
         TabIndex        =   116
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Empleador:"
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FrameVehiculo 
      Height          =   1695
      Left            =   240
      TabIndex        =   108
      Top             =   6120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.ComboBox ComboTipoVehiculo 
         Height          =   315
         ItemData        =   "FrmClientesAbm.frx":045D
         Left            =   960
         List            =   "FrmClientesAbm.frx":0479
         Style           =   2  'Dropdown List
         TabIndex        =   45
         ToolTipText     =   "Tipo de vehiculo"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtPatenteVehiculo 
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   43
         ToolTipText     =   "Nº de patente del vehiculo"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtValuacionVehiculo 
         Height          =   285
         Left            =   3600
         MaxLength       =   9
         TabIndex        =   44
         Tag             =   "N"
         ToolTipText     =   "Valuacion estimada en pesos del vehiculo"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtModeloVehiculo 
         Height          =   285
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   42
         ToolTipText     =   "Modelo del vehiculo"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtMarcaVehiculo 
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   41
         ToolTipText     =   "Marca del vehiculo"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label35 
         Caption         =   "Tipo           :"
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label34 
         Caption         =   "Nº patente :"
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label33 
         Caption         =   "Valuacion $:"
         Height          =   255
         Left            =   2640
         TabIndex        =   111
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "Modelo       :"
         Height          =   255
         Left            =   2640
         TabIndex        =   110
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Marca        :"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FrameCalificacion 
      Height          =   1695
      Left            =   240
      TabIndex        =   119
      Top             =   6120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox TxtAfip 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   54
         ToolTipText     =   "Informe de la Afip"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox TxtJudicial 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   53
         ToolTipText     =   "Informe del Poder Judicial"
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox TxtBancoCentral 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   50
         ToolTipText     =   "Calificacion crediticia brindada por el  banco central"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtAnses 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   52
         ToolTipText     =   "Informe del Anses"
         Top             =   600
         Width           =   2895
      End
      Begin VB.CheckBox Checkveraz 
         Caption         =   "Figura en el veraz"
         Height          =   255
         Left            =   4680
         TabIndex        =   51
         ToolTipText     =   "Marcar si el cliente figura en el veraz"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label41 
         Caption         =   "Informe AFIP (Aportes y contribuciones)        :"
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label40 
         Caption         =   "Informe poder judicial (Causas y ejecuciones):"
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label39 
         Caption         =   "Informe de aportes y contribuciones (Anses) :"
         Height          =   255
         Left            =   120
         TabIndex        =   121
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label38 
         Caption         =   "Calificacion crediticia (Banco central)           :"
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame FrameFacturacion 
      Height          =   2655
      Left            =   240
      TabIndex        =   142
      Top             =   6240
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox TxtFac7 
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   155
         ToolTipText     =   "Nombre del garante"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox TxtMonto7 
         Height          =   285
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   156
         ToolTipText     =   "Apellido del garante"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox TxtFac6 
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   153
         ToolTipText     =   "Nombre del garante"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox TxtMonto6 
         Height          =   285
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   154
         ToolTipText     =   "Apellido del garante"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox TxtFac5 
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   151
         ToolTipText     =   "Nombre del garante"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox TxtMonto5 
         Height          =   285
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   152
         ToolTipText     =   "Apellido del garante"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox TxtMonto4 
         Height          =   285
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   150
         ToolTipText     =   "Apellido del garante"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtFac4 
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   149
         ToolTipText     =   "Nombre del garante"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox TxtMonto3 
         Height          =   285
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   148
         ToolTipText     =   "Apellido del garante"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtFac3 
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   147
         ToolTipText     =   "Nombre del garante"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TxtMonto2 
         Height          =   285
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   146
         ToolTipText     =   "Apellido del garante"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtFac2 
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   145
         ToolTipText     =   "Nombre del garante"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TxtFac1 
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   143
         ToolTipText     =   "Nombre del garante"
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox TxtMonto1 
         Height          =   285
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   144
         ToolTipText     =   "Apellido del garante"
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label64 
         Caption         =   "Nro factura:"
         Height          =   255
         Left            =   120
         TabIndex        =   170
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label63 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   2640
         TabIndex        =   169
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label62 
         Caption         =   "Nro factura:"
         Height          =   255
         Left            =   120
         TabIndex        =   168
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label61 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   2640
         TabIndex        =   167
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label60 
         Caption         =   "Nro factura:"
         Height          =   255
         Left            =   120
         TabIndex        =   166
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label57 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   2640
         TabIndex        =   165
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label56 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   2640
         TabIndex        =   164
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label55 
         Caption         =   "Nro factura:"
         Height          =   255
         Left            =   120
         TabIndex        =   163
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label54 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   2640
         TabIndex        =   162
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label53 
         Caption         =   "Nro factura:"
         Height          =   255
         Left            =   120
         TabIndex        =   161
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label52 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   2640
         TabIndex        =   160
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label51 
         Caption         =   "Nro factura:"
         Height          =   255
         Left            =   120
         TabIndex        =   159
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label59 
         Caption         =   "Nro factura:"
         Height          =   255
         Left            =   120
         TabIndex        =   158
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label58 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   2640
         TabIndex        =   157
         Top             =   120
         Width           =   615
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2175
      Left            =   120
      TabIndex        =   27
      Top             =   5880
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3836
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Garante"
            Object.ToolTipText     =   "Datos del garante o persona de contacto"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Propiedad"
            Object.ToolTipText     =   "Propiedades del cliente en garantia"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vehiculo"
            Object.ToolTipText     =   "Vehiculos del cliente en garantia"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Empleador"
            Object.ToolTipText     =   "Datos del empleador del cliente"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Calificacion"
            Object.ToolTipText     =   "Calificacion crediticia del cliente"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Observaciones"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Datos Facturación"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmClientesAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AQUI SE AGREGAN LOS DATOS DE LOS CLIENTES A LOS CUALES LES OTORGAREMOS CREDITOS

Public FormularioPadre As String
Private Todos As Boolean

Private Function IntegrarCUIT() As String

    Dim cDocumento      As String
    Dim cSexo           As String
    
    cDocumento = Format$(Trim$(TxtNumDocumento), "00000000")
    
    Select Case UCase$(Trim$(ComboSexo.Text))
    Case "MASCULINO"
        cSexo = "M"
    Case "FEMENINO"
        cSexo = "F"
    Case Else
        cSexo = ""
    End Select
    
    IntegrarCUIT = ObtenerCUIT(cDocumento, cSexo)
    
End Function
Private Function NombreValido(cNombre As String) As Boolean
    Dim J       As Integer
    Dim cLetra  As String
    
    NombreValido = True
    For J = 1 To Len(cNombre)
        cLetra = Mid$(cNombre, J, 1)
        If Not ((cLetra >= "A" And cLetra <= "Z") Or (cLetra >= "a" And cLetra <= "z") Or (cLetra = " ") Or (cLetra = "'") Or (cLetra = "ñ") Or (cLetra = "Ñ") Or (cLetra >= "á" And cLetra <= "ú") Or (cLetra >= "Á" And cLetra <= "Ú")) Then
            NombreValido = False
        End If
    Next

End Function

Private Sub CHMonotributista_Click()
    If CHMonotributista.Value = 1 Then
        CmbMonotributista.Enabled = True
    Else
        CmbMonotributista.ListIndex = -1
        CmbMonotributista.Enabled = False
    End If
End Sub

Private Sub ComboSexo_LostFocus()
    TxtCuil.Text = IntegrarCUIT()
End Sub

Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call RefrescarOpcionesSistema

'si no la llamo ninguna otra pantalla deshabilito la seleccion
If Trim(FormularioPadre) = "" Then
   CmdSeleccionar.Enabled = False
End If

Call CargarCombo2("localidades", ComboLocalidades)
ComboLocalidades.ListIndex = -1

ComboEstadoCivil.ListIndex = 0
ComboTipoPropiedad.ListIndex = 0
ComboTipoVehiculo.ListIndex = 0
ComboSexo.ListIndex = 0
Call CargarLista
Call CargarDatos

TipoEdicion = "C"
Call SetearEntorno
Todos = True


Exit Sub
merror:
tratarerrores "Error cargando la pantalla de clientes"
End Sub
Private Sub Form_Unload(Cancel As Integer)
FormularioPadre = ""
Call RefreshTimer

End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
If TipoEdicion = "C" Then
   Unload Me
Else
   TipoEdicion = "C"
   Call SetearEntorno
   Call CargarDatos
   lv.SetFocus
End If
End Sub
Private Sub CmdFiltrar_Click()
Call RefreshTimer
If Trim(TxtCampo.Text) <> "" Then
   Call BuscarClientes
End If
End Sub
Private Sub CmdRefrescar_Click()
Call RefreshTimer
'refresca la pantalla
Call RefrescarOpcionesSistema
Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno
End Sub
Private Sub cmdModificar_Click()
Call RefreshTimer
        
If Not VerificarSeleccionLista(lv) Then Exit Sub
TipoEdicion = "M"
Call SetearEntorno
 
End Sub
Private Sub cmdnuevo_Click()
Call RefreshTimer
TipoEdicion = "N"
Call SetearEntorno
TxtNumLegajo.Text = ObtenerUltimoLegajo() + 1
End Sub
Private Sub CmdTodos_Click()
TxtCampo.Text = ""
Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno
lv.SetFocus
Todos = True
End Sub
Private Sub cmdimprimir_Click()
Call RefreshTimer

CmdImprimir.Enabled = False
If lv.ListItems.Count > 0 Then
   Call ImprimirClientes
End If
CmdImprimir.Enabled = True
End Sub
Private Sub ImprimirClientes()
Dim rec As rdoResultset
Dim Mreporte As New ARClientesNuevo
Dim Condicion As String
Dim Archivo As String
On Error GoTo merror

If Todos Then
   'Condicion = "1=1"
   'a pedido ahora se imprime en el q estas posicionado
   Condicion = "clientes.idcliente=" & CLng(lv.SelectedItem)
Else
   'por apellido
   If Option1.Value Then
      Condicion = "clientes.apellido='" & CStr(TxtCampo.Text) & "'"
   End If
   'por documento
   If Option2.Value Then
      Condicion = "clientes.numdocumento='" & CStr(TxtCampo.Text) & "'"
   End If
   'por cuit
   If Option3.Value Then
      Condicion = "clientes.cuil='" & CStr(TxtCampo.Text) & "'"
   End If
   'por legajo
   If Option4.Value Then
      Condicion = "clientes.numlegajo='" & CStr(TxtCampo.Text) & "'"
   End If
End If

Set rec = CargarRecClientes(Condicion)

If Not rec.EOF Then
   'si imprimo los datos de empresa
   Mreporte.LabelEmpresa = VG_EMPRESA & vbNullString
   Mreporte.Caption = "Imprimir la lista de clientes"
   Mreporte.LabelTitulo.Caption = "Solicitud del Cliente    " & CStr(Date)
   Mreporte.RDODataControl1.Resultset = rec
   Mreporte.Show vbModal
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo la lista de clientes"
End Sub
Private Sub CmdResumen_Click()
Call RefreshTimer

CmdImprimir.Enabled = False
If lv.ListItems.Count > 0 Then
   Call ImprimirResumenClientes
End If
CmdImprimir.Enabled = True
End Sub
Private Sub ImprimirResumenClientes()
Dim rec As rdoResultset
Dim Mreporte As New ARClientesLista
Dim Condicion As String
Dim Archivo As String
On Error GoTo merror

If Todos Then
   Condicion = "1=1"
Else
   'por apellido
   If Option1.Value Then
      Condicion = "clientes.apellido='" & CStr(TxtCampo.Text) & "'"
   End If
   'por documento
   If Option2.Value Then
      Condicion = "clientes.numdocumento'=" & CStr(TxtCampo.Text) & "'"
   End If
   'por cuit
   If Option3.Value Then
      Condicion = "clientes.cuil='" & CStr(TxtCampo.Text) & "'"
   End If
   'por legajo
   If Option4.Value Then
      Condicion = "clientes.numlegajo='" & CStr(TxtCampo.Text) & "'"
   End If
End If

Set rec = CargarRecClientes(Condicion)

If Not rec.EOF Then
   'si imprimo los datos de empresa
    Mreporte.LabelEmpresa = VG_EMPRESA & vbNullString
    Mreporte.Caption = "Imprimir resumen de clientes"
    Mreporte.LabelTitulo.Caption = "Resumen de clientes a la fecha " & CStr(Date)
    Mreporte.RDODataControl1.Resultset = rec
    Mreporte.Show vbModal
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo el resumen de clientes"
End Sub
Private Sub cmdborrar_Click()
Dim sql As String
On Error GoTo merror
Call RefreshTimer
         
If Not VerificarSeleccionLista(lv) Then Exit Sub

If Not PuedoBorrarCliente(lv.SelectedItem) Then
   MsgE "No se puede borrar el cliente porque tiene registros asociados"
   Exit Sub
End If

If Not MsgP("¿Confirma el borrado del cliente seleccionado?") Then Exit Sub

'otras validaciones
If Not ExisteCliente(lv.SelectedItem) Then
   MsgE "El cliente no existe"
   Exit Sub
End If

'inicio de transaccion
cnSQL.BeginTrans

sql = "delete from clientes WHERE idcliente=" & CLng(lv.SelectedItem)
       
cnSQL.Execute sql

'borro excedentes
sql = "delete from excedentesclientes WHERE idcliente=" & CLng(lv.SelectedItem)
cnSQL.Execute sql


'fin de transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

MsgI "El cliente fue borrado"
lv.SetFocus

Exit Sub
merror:
tratarerrores "Error borrando clientes"
End Sub
Private Function ExisteLegajoCliente(ByVal Numlegajo As String) As Boolean
'chequea el numero de legajo para no repetirlo
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteLegajoCliente = False

sql = "select idcliente from clientes " & _
      "where numlegajo='" & CStr(Numlegajo) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcliente")) Then
      ExisteLegajoCliente = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteLegajoCliente"
End Function
Private Function ExisteDocumentoCliente(ByVal NumDocumento As String) As Boolean
'chequea el numero de legajo para no repetirlo
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteDocumentoCliente = False

sql = "select idcliente from clientes " & _
      "where numdocumento='" & CStr(NumDocumento) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcliente")) Then
      ExisteDocumentoCliente = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteDocumentoCliente"
End Function
Private Function ExisteLegajoIgual(ByVal IdCliente As Long, ByVal Numlegajo As String) As Boolean
'verifica si el legajo esta asigado a un cliente distinto al parametro
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteLegajoIgual = False

sql = "select idcliente from clientes " & _
      "where numlegajo='" & CStr(Numlegajo) & _
      "' and idcliente<>" & CLng(IdCliente)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcliente")) Then
      ExisteLegajoIgual = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteLegajoIgual"
End Function
Private Function ExisteDocumentoIgual(ByVal IdCliente As Long, ByVal NumDocumento As String) As Boolean
'verifica si el legajo esta asigado a un cliente distinto al parametro
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteDocumentoIgual = False

sql = "select idcliente from clientes " & _
      "where numdocumento='" & CStr(NumDocumento) & _
      "' and idcliente<>" & CLng(IdCliente)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcliente")) Then
      ExisteDocumentoIgual = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteDocumentoIgual"
End Function
Private Sub cmdGrabar_Click()
'graba los registros nuevos y las modificaciones
Dim sql As String
Dim IdCliente As Long
Dim IdLocalidad As Long
Dim Mensaje As String
Dim Sexo As String
Dim a As Integer
On Error GoTo merror
Call RefreshTimer

If Not datosok() Then Exit Sub

Sexo = ""
If ComboSexo.Text = "MASCULINO" Then
   Sexo = "M"
End If
If ComboSexo.Text = "FEMENINO" Then
   Sexo = "F"
End If

IdLocalidad = CLng(ComboLocalidades.ItemData(ComboLocalidades.ListIndex))

If TipoEdicion = "N" Then
   'Verifico que no exista el mismo nº de legajo
   If ExisteLegajoCliente(TxtNumLegajo.Text) Then
      MsgE "El Nº de cliente ya existe"
      TxtNumLegajo.SetFocus
      Exit Sub
   End If
   
   'Verifico que no exista el mismo nº de documento
   If ExisteDocumentoCliente(TxtNumDocumento.Text) Then
      MsgE "El Nº de documento ya existe"
      TxtNumDocumento.SetFocus
      Exit Sub
   End If
   
   If TxtCuil.Text <> IntegrarCUIT() Then
        If Not MsgP("El CUIT ingresado no es el correcto de acuerdo al algoritmo, ¿desea continuar?") Then Exit Sub
   End If

   If Not MsgP("¿Confirma el nuevo cliente?") Then Exit Sub
   
   IdCliente = UltimoId("idcliente", "clientes") + 1
   
   'otras validaciones
   If ExisteCliente(IdCliente) Then
      MsgE "El cliente ya existe"
      Exit Sub
   End If
   
   'Verifico que no exista el mismo nº de legajo
   If ExisteLegajoCliente(TxtNumLegajo.Text) Then
      MsgE "El Nº de cliente ya existe"
      TxtNumLegajo.SetFocus
      Exit Sub
   End If
   
   'Verifico que no exista el mismo nº de documento
   If ExisteDocumentoCliente(TxtNumDocumento.Text) Then
      MsgE "El Nº de documento ya existe"
      TxtNumDocumento.SetFocus
      Exit Sub
   End If
   
   
   If Not ExisteLocalidad(IdLocalidad) Then
      MsgE "La localidad seleccionada no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans
   
   'agrego al cliente
   sql = "INSERT INTO clientes (Idcliente,nombre,apellido," & _
         "telefono,domicilio,numdocumento,nacionalidad,empresa,sueldo,antiguedad," & _
         "observaciones,fechanacimiento,numlegajo,recibosueldo,veraz," & _
         "residencia,cuil,idlocalidad,facturaservicio,codigodescuento," & _
         "cuitempleador,numcbu,estadocivil,profesion,nombregarante," & _
         "apellidogarante,documentogarante,cuitgarante,telefonogarante,domiciliogarante,sueldogarante,nacionalidadgarante," & _
         "profesiongarante,tipopropiedad,catastropropiedad,metrospropiedad,valuacionpropiedad,tipovehiculo," & _
         "patentevehiculo,marcavehiculo,modelovehiculo,valuacionvehiculo,bancocentral,anses,judicial,afip,domicilioempleador,telefonoempleador,creditomaximo,tipoiva,codigopostal,cad1,cad2,cad3,observaciones1,actividad,monotributista,jubilado,CatMonotributo,TipoGarante) " & _
         "VALUES (" & CLng(IdCliente) & _
         ",'" & CStr(TxtNombre.Text) & "','" & CStr(TxtApellido.Text) & _
         "','" & CStr(TxtTelefono.Text) & "','" & CStr(TxtDomicilio.Text) & _
         "','" & CStr(TxtNumDocumento.Text) & "','" & CStr(TxtNacionalidad.Text) & "','" & CStr(TxtEmpresa.Text) & _
         "'," & ConvertirDblSql(CCur(TxtSueldo.Text)) & "," & CLng(TxtAntiguedad.Text) & _
         ",'" & CStr(TxtCelular.Text) & "','" & ConvertirFechaSql(DTPicker1.Value, "DD/MM/YYYY") & _
         "','" & CStr(TxtNumLegajo.Text) & "'," & CheckReciboSueldo.Value & _
         "," & Checkveraz.Value & "," & CLng(TxtResidencia.Text) & _
         ",'" & CStr(TxtCuil.Text) & "'," & CLng(IdLocalidad) & _
         "," & CheckFacturaServicio.Value & ",'" & CStr(TxtCodigoDescuento.Text) & _
         "','" & CStr(TxtCuitEmpleador.Text) & "','" & CStr(TxtNumCBU.Text) & _
         "','" & CStr(ComboEstadoCivil.Text) & "','" & CStr(TxtActividad.Text) & "','" & CStr(TxtNombreGarante.Text) & _
         "','" & CStr(TxtApellidoGarante.Text) & "','" & CStr(TxtDocumentoGarante.Text) & "','" & CStr(TxtCuitGarante.Text) & _
         "','" & CStr(TxtTelefonoGarante.Text) & "','" & CStr(TxtDomicilioGarante.Text) & "'," & ConvertirDblSql(CCur(TxtSueldoGarante.Text)) & _
         ",'" & CStr(TxtNacionalidadGarante.Text) & "','" & CStr(TxtProfesionGarante.Text) & "','" & CStr(ComboTipoPropiedad.Text) & "','" & CStr(TxtCatastroPropiedad.Text) & _
         "','" & CStr(TxtMetrosPropiedad.Text) & "'," & ConvertirDblSql(CCur(TxtValuacionPropiedad.Text)) & ",'" & CStr(ComboTipoVehiculo.Text) & _
         "','" & CStr(TxtPatenteVehiculo.Text) & "','" & CStr(TxtMarcaVehiculo.Text) & "','" & CStr(TxtModeloVehiculo.Text) & "'," & ConvertirDblSql(CCur(TxtValuacionVehiculo.Text)) & ",'" & CStr(TxtBancoCentral.Text) & "','" & CStr(TxtAnses.Text) & "','" & CStr(TxtJudicial.Text) & "','" & CStr(TxtAfip.Text) & "','" & CStr(TxtDomicilioEmpleador.Text) & "','" & CStr(TxtTelefonoEmpleador.Text) & "'," & ConvertirDblSql(CCur(TxtCreditoMaximo.Text)) & ",'" & CStr(TxtTipoIva.Text) & "','" & CStr(TxtCP.Text) & "','" & CStr(Sexo) & "','" & CStr(TxtObservaciones.Text) & "','" & CStr(TxtEmail.Text) & "','" & CStr(TxtObservaciones1.Text) & "','" & CStr(Txtactividad1.Text) & "','" & CHMonotributista.Value & "'," & CheckJubPen.Value & ",'" & CmbMonotributista.Text & "','" & Cmbdiferenciando.Text & "')"
  
   cnSQL.Execute sql
   
   'agrego Datos de factura del cliente
     
      If Trim(TxtFac1) <> "" Then
        sql = "CargarDatosFactura " & CLng(IdCliente) & "," & 1 & ",'" & TxtFac1.Text & "'," & ConvertirDblSql(CCur(TxtMonto1.Text))
         cnSQL.Execute sql
      End If
      If Trim(TxtFac2) <> "" Then
        sql = "CargarDatosFactura " & CLng(IdCliente) & "," & 2 & ",'" & TxtFac2.Text & "'," & ConvertirDblSql(CCur(TxtMonto2.Text))
         cnSQL.Execute sql
      End If
      If Trim(TxtFac3) <> "" Then
        sql = "CargarDatosFactura " & CLng(IdCliente) & "," & 3 & ",'" & TxtFac3.Text & "'," & ConvertirDblSql(CCur(TxtMonto3.Text))
         cnSQL.Execute sql
      End If
      If Trim(TxtFac4) <> "" Then
        sql = "CargarDatosFactura " & CLng(IdCliente) & "," & 4 & ",'" & TxtFac4.Text & "'," & ConvertirDblSql(CCur(TxtMonto4.Text))
         cnSQL.Execute sql
      End If
      If Trim(TxtFac5) <> "" Then
        sql = "CargarDatosFactura " & CLng(IdCliente) & "," & 5 & ",'" & TxtFac5.Text & "'," & ConvertirDblSql(CCur(TxtMonto5.Text))
         cnSQL.Execute sql
      End If
      If Trim(TxtFac6) <> "" Then
        sql = "CargarDatosFactura " & CLng(IdCliente) & "," & 6 & ",'" & TxtFac6.Text & "'," & ConvertirDblSql(CCur(TxtMonto6.Text))
         cnSQL.Execute sql
      End If
      If Trim(TxtFac7) <> "" Then
        sql = "CargarDatosFactura " & CLng(IdCliente) & "," & 7 & ",'" & TxtFac7.Text & "'," & ConvertirDblSql(CCur(TxtMonto7.Text))
        cnSQL.Execute sql
      End If
      
  
   
   'fin de transaccion
   cnSQL.CommitTrans
    
    
   Mensaje = "El nuevo cliente fue agregado"
   
   Call CargarLista
   Call CargarDatos
Else
   
   'verifica que el legajo no exista en un cliente diferente al actual
   If ExisteLegajoIgual(lv.SelectedItem, TxtNumLegajo.Text) Then
      MsgE "El Nº de cliente ya esta asignado a otro cliente"
      TxtNumLegajo.SetFocus
      Exit Sub
   End If
   
   'verifica que el documento no exista en un cliente diferente al actual
   If ExisteDocumentoIgual(lv.SelectedItem, TxtNumDocumento.Text) Then
      MsgE "El Nº de documento ya esta asignado a otro cliente"
      TxtNumDocumento.SetFocus
      Exit Sub
   End If
      
   If TxtCuil.Text <> IntegrarCUIT() Then
        If Not MsgP("El CUIT ingresado no es el correcto de acuerdo al algoritmo, ¿desea continuar?") Then Exit Sub
   End If
     
   If Not MsgP("¿Confirma la modificacion del cliente seleccionado?") Then Exit Sub
   
   'otras validaciones
   If Not ExisteCliente(lv.SelectedItem) Then
      MsgE "El cliente no existe"
      Exit Sub
   End If
   
   'verifica que el legajo no exista en un cliente diferente al actual
   If ExisteLegajoIgual(lv.SelectedItem, TxtNumLegajo.Text) Then
      MsgE "El Nº de cliente ya esta asignado a otro cliente"
      TxtNumLegajo.SetFocus
      Exit Sub
   End If
   
   'verifica que el documento no exista en un cliente diferente al actual
   If ExisteDocumentoIgual(lv.SelectedItem, TxtNumDocumento.Text) Then
      MsgE "El Nº de documento ya esta asignado a otro cliente"
      TxtNumDocumento.SetFocus
      Exit Sub
   End If
   
   
   If Not ExisteLocalidad(IdLocalidad) Then
      MsgE "La localidad seleccionada no existe"
      Exit Sub
   End If

   'inicio de transaccion
   cnSQL.BeginTrans

   'estoy grabando una modificacion
   sql = "UPDATE clientes SET " & _
         "numlegajo='" & CStr(TxtNumLegajo.Text) & _
         "',nombre='" & CStr(TxtNombre.Text) & "',apellido='" & CStr(TxtApellido.Text) & _
         "',telefono='" & CStr(TxtTelefono.Text) & "',domicilio='" & CStr(TxtDomicilio.Text) & _
         "',numdocumento='" & CStr(TxtNumDocumento.Text) & "',nacionalidad='" & CStr(TxtNacionalidad.Text) & "',estadocivil='" & CStr(ComboEstadoCivil.Text & _
         "',empresa='" & CStr(TxtEmpresa.Text) & "',sueldo=" & ConvertirDblSql(CCur(TxtSueldo.Text)) & _
         ",antiguedad=" & CLng(TxtAntiguedad.Text) & ",profesion='" & CStr(TxtActividad.Text) & _
         "',observaciones='" & CStr(TxtCelular.Text) & "',fechanacimiento='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY")) & _
         "',residencia=" & CLng(TxtResidencia.Text) & ",cuil='" & CStr(TxtCuil.Text) & _
         "',idlocalidad=" & CLng(IdLocalidad) & ",numcbu='" & CStr(TxtNumCBU.Text) & _
         "',codigodescuento='" & CStr(TxtCodigoDescuento.Text) & "',cuitempleador='" & CStr(TxtCuitEmpleador.Text) & _
         "',telefonoempleador='" & CStr(TxtTelefonoEmpleador.Text) & "',domicilioempleador='" & CStr(TxtDomicilioEmpleador.Text) & _
         "',facturaservicio=" & CheckFacturaServicio.Value & ",recibosueldo=" & CheckReciboSueldo.Value & _
         ",veraz=" & Checkveraz.Value & ",nombregarante='" & CStr(TxtNombreGarante.Text) & "',apellidogarante='" & CStr(TxtApellidoGarante.Text) & _
         "',documentogarante='" & CStr(TxtDocumentoGarante.Text) & "',cuitgarante='" & CStr(TxtCuitGarante.Text) & _
         "',telefonogarante='" & CStr(TxtTelefonoGarante.Text) & "',domiciliogarante='" & CStr(TxtDomicilioGarante.Text) & "',profesiongarante='" & CStr(TxtProfesionGarante.Text) & _
         "',sueldogarante=" & ConvertirDblSql(CCur(TxtSueldoGarante.Text)) & ",nacionalidadgarante='" & CStr(TxtNacionalidadGarante.Text) & _
         "',tipopropiedad='" & CStr(ComboTipoPropiedad.Text) & "',catastropropiedad='" & CStr(TxtCatastroPropiedad.Text) & _
         "',metrospropiedad='" & CStr(TxtMetrosPropiedad.Text) & "',observaciones1='" & CStr(TxtObservaciones1.Text) & "',valuacionpropiedad=" & ConvertirDblSql(CCur(TxtValuacionPropiedad.Text)) & _
         ",tipovehiculo='" & CStr(ComboTipoVehiculo.Text) & "',patentevehiculo='" & CStr(TxtPatenteVehiculo.Text) & "',marcavehiculo='" & CStr(TxtMarcaVehiculo.Text) & "',modelovehiculo='" & CStr(TxtModeloVehiculo.Text) & "',valuacionvehiculo=" & ConvertirDblSql(CCur(TxtValuacionVehiculo.Text)) & _
         ",bancocentral='" & CStr(TxtBancoCentral.Text) & "',anses='" & CStr(TxtAnses.Text) & "',judicial='" & CStr(TxtJudicial.Text) & "',afip='" & CStr(TxtAfip.Text) & "',creditomaximo=" & ConvertirDblSql(CCur(TxtCreditoMaximo.Text)) & ",tipoiva='" & CStr(TxtTipoIva.Text) & "',codigopostal='" & CStr(TxtCP.Text) & "',cad1='" & CStr(Sexo) & "',cad2='" & CStr(TxtObservaciones.Text) & "',cad3='" & CStr(TxtEmail.Text) & _
         "',Actividad='" & CStr(Txtactividad1.Text) & "',monotributista='" & CHMonotributista.Value & "',jubilado = " & CheckJubPen.Value & ",CatMonotributo = '" & CmbMonotributista.Text & _
         "',TipoGarante = '" & CStr(Cmbdiferenciando.Text) & _
         "' WHERE Idcliente=" & CLng(lv.SelectedItem)
   
   cnSQL.Execute sql
   
   'agrego Datos de factura del cliente
     
      
        sql = "CargarDatosFactura " & CLng(lv.SelectedItem) & "," & 1 & ",'" & TxtFac1.Text & "'," & ConvertirDblSql(CCur(TxtMonto1.Text))
         cnSQL.Execute sql
       
        sql = "CargarDatosFactura " & CLng(lv.SelectedItem) & "," & 2 & ",'" & TxtFac2.Text & "'," & ConvertirDblSql(CCur(TxtMonto2.Text))
         cnSQL.Execute sql
       
        sql = "CargarDatosFactura " & CLng(lv.SelectedItem) & "," & 3 & ",'" & TxtFac3.Text & "'," & ConvertirDblSql(CCur(TxtMonto3.Text))
         cnSQL.Execute sql
       
        sql = "CargarDatosFactura " & CLng(lv.SelectedItem) & "," & 4 & ",'" & TxtFac4.Text & "'," & ConvertirDblSql(CCur(TxtMonto4.Text))
         cnSQL.Execute sql
      
        sql = "CargarDatosFactura " & CLng(lv.SelectedItem) & "," & 5 & ",'" & TxtFac5.Text & "'," & ConvertirDblSql(CCur(TxtMonto5.Text))
         cnSQL.Execute sql
      
        sql = "CargarDatosFactura " & CLng(lv.SelectedItem) & "," & 6 & ",'" & TxtFac6.Text & "'," & ConvertirDblSql(CCur(TxtMonto6.Text))
         cnSQL.Execute sql
      
        sql = "CargarDatosFactura " & CLng(lv.SelectedItem) & "," & 7 & ",'" & TxtFac7.Text & "'," & ConvertirDblSql(CCur(TxtMonto7.Text))
        cnSQL.Execute sql
      
     
   'fin de transaccion datos de factura del cliente
   cnSQL.CommitTrans
   

   Mensaje = "El cliente fue modificado"
   
   'actualizo cambios en la lista
   lv.SelectedItem.ListSubItems(1).Text = TxtNumLegajo.Text & vbNullString
   lv.SelectedItem.ListSubItems(2).Text = TxtApellido.Text & " " & TxtNombre.Text & vbNullString
End If

TipoEdicion = "C"
Call SetearEntorno

MsgI Mensaje

lv.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando clientes"
End Sub
Private Sub BuscarClientes()
'busco los clientes que cumplan una condicion determinada
Dim sql As String
Dim rec As rdoResultset
Dim Condicion As String
Dim Opcion As Long
On Error GoTo merror

If Trim(TxtCampo.Text) = "" Then Exit Sub

Opcion = 1
'por apellido
If Option1.Value Then
   Condicion = "charindex('" & CStr(TxtCampo.Text) & "',clientes.apellido + ' ' +clientes.nombre ) > 0"
End If

'por documento
If Option2.Value Then
   Condicion = "clientes.numdocumento='" & CStr(TxtCampo.Text) & "'"
End If

'por cuit
If Option3.Value Then
   Condicion = "clientes.cuil='" & CStr(TxtCampo.Text) & "'"
End If

'por legajo
If Option4.Value Then
   Condicion = "clientes.numlegajo='" & CStr(TxtCampo.Text) & "'"
End If

'por cupon
If Option5.Value Then
   If Not IsNumeric(TxtCampo.Text) Then
      TxtCampo = 0
   End If
   
   Opcion = 2
   Condicion = "cuotas.numfactura=" & CLng(TxtCampo.Text)
End If


Set rec = CargarRecClientes2(Condicion, Opcion)

If Not rec.EOF Then
   'cargo la lista
   If Not rec.EOF Then
      lv.ListItems.Clear
      Do While Not rec.EOF
         Set Nitem = lv.ListItems.Add(, , Format(rec.rdoColumns("Idcliente"), "000000"))
         Nitem.SubItems(1) = rec.rdoColumns("numlegajo") & vbNullString
         Nitem.SubItems(2) = rec.rdoColumns("apellido") & " " & rec.rdoColumns("Nombre") & vbNullString
         rec.MoveNext
      Loop
   End If
   Call CargarDatos
   TipoEdicion = "C"
   Call SetearEntorno
   lv.SetFocus
   Todos = False
Else
   lv.ListItems.Clear
   Call LimpiarCampos(Me)
   cmdModificar.Enabled = False
   CmdBorrar.Enabled = False
   CmdImprimir.Enabled = False
   CmdExportarClientes.Enabled = False
   CmdResumen.Enabled = False
   MsgE "No hay coincidencias en la base de datos"
   Todos = True
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento BuscarClientes-ClientesAbm"
End Sub
Private Function CargarRecClientes2(ByVal Condicion As String, ByVal Opcion As Long) As rdoResultset
Dim sql As String
On Error GoTo merror

If Opcion = 1 Then
   sql = "SELECT clientes.*," & _
         "clientes.apellido + ' ' + clientes.nombre as cliente," & _
         "localidades.nombre as localidad," & _
         "clientes.apellidogarante + ' ' + clientes.nombregarante as garante " & _
         "FROM localidades inner join clientes on localidades.idlocalidad=clientes.idlocalidad " & _
         "where " & Condicion & _
         " ORDER BY clientes.apellido,clientes.nombre"
Else
   'SI FILTRO POR CUPON debe traer todos los clientes que tienen ese nro de cupon
   sql = "SELECT clientes.*," & _
         "clientes.apellido + ' ' + clientes.nombre as cliente," & _
         "localidades.nombre as localidad," & _
         "clientes.apellidogarante + ' ' + clientes.nombregarante as garante " & _
         "FROM localidades inner join (clientes inner join (creditos inner join cuotas on creditos.idcredito=cuotas.idcredito) on clientes.idcliente=creditos.idcliente) on localidades.idlocalidad=clientes.idlocalidad " & _
         "where " & Condicion & _
         " ORDER BY clientes.apellido,clientes.nombre"
End If

Set CargarRecClientes2 = cnSQL.OpenResultset(sql)

Exit Function
merror:
tratarerrores "Error en funcion CargarRecClientes2"
End Function
Private Function CargarRecClientes(ByVal Condicion As String) As rdoResultset
Dim sql As String
On Error GoTo merror

    sql = "SELECT clientes.*," & _
      "clientes.apellido + ', ' + clientes.nombre as cliente," & _
      "localidades.nombre as localidad,provincias.nombre as provincia,clientes.codigopostal as codigopostal2," & _
      "clientes.apellidogarante + ' ' + clientes.nombregarante as garante,lower(tipogarante) + ':' as tipogarantemin " & _
      "FROM provincias inner join (localidades inner join clientes on localidades.idlocalidad=clientes.idlocalidad) on provincias.idprovincia=localidades.idprovincia " & _
      "where " & Condicion & _
      "ORDER BY clientes.apellido,clientes.nombre  "


Set CargarRecClientes = cnSQL.OpenResultset(sql)

Exit Function
merror:
tratarerrores "Error en funcion CargarRecClientes"
End Function

Private Function CargarDatosFactura(ByVal Condicion As String) As rdoResultset
Dim sql As String
On Error GoTo merror

sql = "SELECT * from datosfactura where " & Condicion & "" _


Set CargarDatosFactura = cnSQL.OpenResultset(sql)

Exit Function
merror:
tratarerrores "Error en funcion CargarDatosFactura"
End Function
Private Sub CargarLista()
'carga la lista con todos los clientes
Dim rec As rdoResultset
Dim Nitem As ListItem
Dim Condicion As String
On Error GoTo merror
    
lv.ListItems.Clear

Condicion = "1=1"

Set rec = CargarRecClientes(Condicion)
    
If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lv.ListItems.Add(, , Format(rec.rdoColumns("Idcliente"), "000000"))
      Nitem.SubItems(1) = rec.rdoColumns("numlegajo") & vbNullString
      Nitem.SubItems(2) = rec.rdoColumns("apellido") & " " & rec.rdoColumns("Nombre") & vbNullString
       
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error cargando la lista de clientes"
End Sub
Private Sub CargarDatos()
'carga el elemento seleccionado del listview en los cuadros de texto
Dim rec As rdoResultset
Dim rec1 As rdoResultset
Dim Condicion As String
On Error GoTo merror
    
If Not VerificarSeleccionLista(lv) Then Exit Sub
        
Condicion = "idcliente=" & CLng(lv.SelectedItem)

Set rec = CargarRecClientes(Condicion)

If Not rec.EOF Then
   TxtNombre.Text = rec.rdoColumns("nombre") & vbNullString
   TxtApellido.Text = rec.rdoColumns("apellido") & vbNullString
   TxtTelefono.Text = rec.rdoColumns("telefono") & vbNullString
   TxtNumDocumento.Text = rec.rdoColumns("numdocumento") & vbNullString
   TxtSueldo.Text = rec.rdoColumns("sueldo") & vbNullString
   TxtAntiguedad.Text = rec.rdoColumns("antiguedad") & vbNullString
   TxtNumLegajo.Text = rec.rdoColumns("numlegajo") & vbNullString
   TxtDomicilio.Text = rec.rdoColumns("domicilio") & vbNullString
   TxtResidencia.Text = rec.rdoColumns("residencia") & vbNullString
   DTPicker1.Value = rec.rdoColumns("fechanacimiento")
   TxtCuil.Text = rec.rdoColumns("cuil") & vbNullString
   TxtNumCBU.Text = rec.rdoColumns("numcbu") & vbNullString
   TxtCodigoDescuento.Text = rec.rdoColumns("codigodescuento") & vbNullString
   TxtActividad.Text = rec.rdoColumns("profesion") & vbNullString
   TxtCreditoMaximo.Text = rec.rdoColumns("creditomaximo") & vbNullString
   TxtTipoIva.Text = rec.rdoColumns("tipoiva") & vbNullString
   TxtCP.Text = rec.rdoColumns("codigopostal2") & vbNullString
   TxtNombreGarante.Text = rec.rdoColumns("nombregarante") & vbNullString
   TxtApellidoGarante.Text = rec.rdoColumns("apellidogarante") & vbNullString
   TxtDocumentoGarante.Text = rec.rdoColumns("documentogarante") & vbNullString
   TxtCuitGarante.Text = rec.rdoColumns("cuitgarante") & vbNullString
   TxtTelefonoGarante.Text = rec.rdoColumns("telefonogarante") & vbNullString
   TxtDomicilioGarante.Text = rec.rdoColumns("domiciliogarante") & vbNullString
   TxtSueldoGarante.Text = rec.rdoColumns("sueldogarante") & vbNullString
   TxtNacionalidad.Text = rec.rdoColumns("nacionalidad") & vbNullString
   TxtNacionalidadGarante.Text = rec.rdoColumns("nacionalidadgarante") & vbNullString
   TxtProfesionGarante.Text = rec.rdoColumns("profesiongarante") & vbNullString
   TxtObservaciones1.Text = Trim(rec.rdoColumns("observaciones1")) & vbNullString
   
   If rec.rdoColumns("tipogarante") <> "" Then
      Cmbdiferenciando.Text = rec.rdoColumns("tipogarante")
   Else
      Cmbdiferenciando.ListIndex = -1
   End If
   
   Checkveraz.Value = 0
   If rec.rdoColumns("veraz") Then
      Checkveraz.Value = 1
   End If
    
   CheckReciboSueldo.Value = 0
   If rec.rdoColumns("recibosueldo") Then
      CheckReciboSueldo.Value = 1
   End If
   
   CheckFacturaServicio.Value = 0
   If rec.rdoColumns("facturaservicio") Then
      CheckFacturaServicio.Value = 1
   End If
   
   'corregir usando la funcion de la libreria
   If ComboLocalidades.ListCount > 0 Then
      ComboLocalidades.Text = rec.rdoColumns("localidad")
   End If
   
   'si hay estado civil lo muestro
   If rec.rdoColumns("estadocivil") <> "" Then
      ComboEstadoCivil.Text = rec.rdoColumns("estadocivil")
   Else
      ComboEstadoCivil.ListIndex = -1
   End If
   
   'datos de la propiedad
   TxtCatastroPropiedad.Text = rec.rdoColumns("catastropropiedad") & vbNullString
   TxtMetrosPropiedad.Text = rec.rdoColumns("metrospropiedad") & vbNullString
   TxtValuacionPropiedad.Text = rec.rdoColumns("valuacionpropiedad") & vbNullString
   
   'si hay un tipo de propiedad actualizo el combo
   If rec.rdoColumns("tipopropiedad") <> "" Then
      ComboTipoPropiedad.Text = rec.rdoColumns("tipopropiedad")
   Else
      ComboTipoPropiedad.ListIndex = -1
   End If
   
   'datos del vehiculo
   TxtPatenteVehiculo.Text = rec.rdoColumns("patentevehiculo") & vbNullString
   TxtMarcaVehiculo.Text = rec.rdoColumns("marcavehiculo") & vbNullString
   TxtModeloVehiculo.Text = rec.rdoColumns("modelovehiculo") & vbNullString
   TxtValuacionVehiculo.Text = rec.rdoColumns("valuacionvehiculo") & vbNullString
   
   'si hay un tipo de vehiculo actualizo el combo
   If rec.rdoColumns("tipovehiculo") <> "" Then
      ComboTipoVehiculo.Text = rec.rdoColumns("tipovehiculo")
   Else
      ComboTipoVehiculo.ListIndex = -1
   End If
   
   ComboSexo.ListIndex = -1
   If rec.rdoColumns("cad1") = "M" Then
      ComboSexo = "MASCULINO"
   End If
   
   If rec.rdoColumns("cad1") = "F" Then
      ComboSexo = "FEMENINO"
   End If
      
   'datos del empleador
   TxtEmpresa.Text = rec.rdoColumns("empresa") & vbNullString
   TxtCuitEmpleador.Text = rec.rdoColumns("cuitempleador") & vbNullString
   TxtTelefonoEmpleador.Text = rec.rdoColumns("telefonoempleador") & vbNullString
   TxtDomicilioEmpleador.Text = rec.rdoColumns("domicilioempleador") & vbNullString
   Txtactividad1.Text = rec.rdoColumns("actividad") & vbNullString
   If Not IsNull(rec.rdoColumns("monotributista")) Then
    If rec.rdoColumns("monotributista") Then
       CHMonotributista.Value = 1
    Else
        CHMonotributista.Value = 0
    End If
   Else
    CHMonotributista.Value = 0
   End If
   
   If Not IsNull(rec.rdoColumns("jubilado")) Then
    If rec.rdoColumns("jubilado") Then
       CheckJubPen.Value = 1
    Else
        CheckJubPen.Value = 0
    End If
   Else
    CheckJubPen.Value = 0
   End If
    If rec.rdoColumns("CatMonotributo") <> "" And rec.rdoColumns("CatMonotributo") <> " " Then
      CmbMonotributista.Text = rec.rdoColumns("CatMonotributo")
   Else
      ComboTipoPropiedad.ListIndex = -1
   End If
   
   'datos de la calificacion
   TxtBancoCentral.Text = rec.rdoColumns("bancocentral") & vbNullString
   TxtAnses.Text = rec.rdoColumns("anses") & vbNullString
   TxtJudicial.Text = rec.rdoColumns("judicial") & vbNullString
   TxtAfip.Text = rec.rdoColumns("afip") & vbNullString
   
   'en celular cargo las observaciones de ellos para evitar traspasos
   'porque usaban obs para celular
   TxtCelular.Text = rec.rdoColumns("observaciones") & vbNullString
   'este lo cargo del nuevo cad2
   TxtObservaciones.Text = rec.rdoColumns("cad2") & vbNullString
   
   TxtEmail.Text = rec.rdoColumns("cad3") & vbNullString
   
   'Cargo los datos de la factura
   Condicion = "idcliente=" & CLng(lv.SelectedItem)

   TxtFac1.Text = ""
   TxtMonto1.Text = ""
   TxtFac2.Text = ""
   TxtMonto2.Text = ""
   TxtFac3.Text = ""
   TxtMonto3.Text = ""
   TxtFac4.Text = ""
   TxtMonto4.Text = ""
   TxtFac5.Text = ""
   TxtMonto5.Text = ""
   TxtFac6.Text = ""
   TxtMonto6.Text = ""
   TxtFac7.Text = ""
   TxtMonto7.Text = ""
   Set rec1 = CargarDatosFactura(Condicion)
   If Not rec1.EOF Then
        Do While Not rec1.EOF
            Select Case rec1.rdoColumns("Secuencia")
            Case 1
                TxtFac1.Text = Trim(rec1.rdoColumns("NroFactura"))
                TxtMonto1.Text = rec1.rdoColumns("Monto")
            Case 2
                TxtFac2.Text = Trim(rec1.rdoColumns("NroFactura"))
                TxtMonto2.Text = rec1.rdoColumns("Monto")
            Case 3
                TxtFac3.Text = Trim(rec1.rdoColumns("NroFactura"))
                TxtMonto3.Text = rec1.rdoColumns("Monto")
            Case 4
                TxtFac4.Text = Trim(rec1.rdoColumns("NroFactura"))
                TxtMonto4.Text = rec1.rdoColumns("Monto")
            Case 5
                TxtFac5.Text = Trim(rec1.rdoColumns("NroFactura"))
                TxtMonto5.Text = rec1.rdoColumns("Monto")
            Case 6
                TxtFac6.Text = Trim(rec1.rdoColumns("NroFactura"))
                TxtMonto6.Text = rec1.rdoColumns("Monto")
            Case 7
                TxtFac7.Text = Trim(rec1.rdoColumns("NroFactura"))
                TxtMonto7.Text = rec1.rdoColumns("Monto")
            End Select
            rec1.MoveNext
        Loop
   End If
   
   
End If
        
Exit Sub
merror:
tratarerrores "Error cargando datos de clientes"
End Sub
Private Function PuedoBorrarCliente(ByVal IdCliente As Long) As Boolean
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

PuedoBorrarCliente = True

'verifico en creditos...
sql = "select idcliente from creditos " & _
      "where idcliente=" & CLng(IdCliente)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcliente")) Then
      PuedoBorrarCliente = False
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion PuedoBorrarCliente"
End Function
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(TxtNombre.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el nombre del cliente"
   TxtNombre.SetFocus
   Exit Function
Else
   If IsNumeric(TxtNombre.Text) _
           Or InStr(1, TxtNombre.Text, "+") > 0 _
           Or InStr(1, TxtNombre.Text, "-") > 0 _
           Or InStr(1, TxtNombre.Text, ".") > 0 _
           Or InStr(1, TxtNombre.Text, "*") > 0 _
           Or InStr(1, TxtNombre.Text, "/") > 0 _
           Or InStr(1, TxtNombre.Text, ";") > 0 _
           Or InStr(1, TxtNombre.Text, ",") > 0 Then
     datosok = False
     MsgE "Nombre invalido"
     TxtNombre.SetFocus
     Exit Function
    End If
End If
    

    
If Trim(TxtApellido.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el apellido del cliente"
   TxtApellido.SetFocus
   Exit Function
End If


If Not NombreValido(Trim(TxtNombre.Text)) Then
   datosok = False
   MsgE "El nombre del cliente tiene caracteres inválidos"
   TxtNombre.SetFocus
   Exit Function
End If

If Not NombreValido(Trim(TxtApellido.Text)) Then
   datosok = False
   MsgE "El apellido del cliente tiene caracteres inválidos"
   TxtApellido.SetFocus
   Exit Function
End If


'valido el numero de legajo
If Trim(TxtNumLegajo.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el numero de cliente"
   TxtNumLegajo.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtNumLegajo.Text) Then
   datosok = False
   MsgE "El numero de cliente debe ser numerico"
   TxtNumLegajo.SetFocus
   Exit Function
End If
If CCur(TxtNumLegajo.Text) <= 0 Then
   datosok = False
   MsgE "El numero de cliente debe ser mayor a cero"
   TxtNumLegajo.SetFocus
   Exit Function
End If

If Trim(TxtNumDocumento.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar numero de documento"
   TxtNumDocumento.SetFocus
   Exit Function
End If

If Not IsNumeric(TxtNumDocumento.Text) Or InStr(1, TxtNumDocumento.Text, ".") > 0 Or _
   InStr(1, TxtNumDocumento.Text, ",") > 0 Or _
   InStr(1, TxtNumDocumento.Text, "d") > 0 Or _
   InStr(1, TxtNumDocumento.Text, "e") > 0 Or _
   InStr(1, TxtNumDocumento.Text, "D") > 0 Or _
   InStr(1, TxtNumDocumento.Text, "E") > 0 Or _
   InStr(1, TxtNumDocumento.Text, " ") > 0 Then
   datosok = False
   MsgE "El número de documento debe ser numérico"
   TxtNumDocumento.SetFocus
   Exit Function
End If

If Len(TxtNumDocumento.Text) > 8 Then
   datosok = False
   MsgE "El número de documento debe contener como máximo 8 dígitos"
   TxtNumDocumento.SetFocus
   Exit Function
End If

'chequeo de fechanacimiento
If Year(DTPicker1.Value) < 1900 Then
   datosok = False
   MsgE "Verifique el año de nacimiento (debe ser superior al año 1900)"
   DTPicker1.SetFocus
   Exit Function
End If
If CDate(DTPicker1.Value) >= CDate(Date) Then
   datosok = False
   MsgE "Verifique la fecha de nacimiento...(debe ser menor a la fecha actual)"
   DTPicker1.SetFocus
   Exit Function
End If

If (Year(Date) - Year(DTPicker1.Value)) < 18 Then
   datosok = False
   MsgE "Verifique la fecha de nacimiento...el cliente debe ser mayor de edad"
   DTPicker1.SetFocus
   Exit Function
End If

If DateDiff("d", DTPicker1.Value, Date) >= (VG_EDAD * 365) Then
   datosok = False
   MsgE "El Cliente es mayor a " & VG_EDAD & " años"
   DTPicker1.SetFocus
   Exit Function
End If

If ComboSexo.ListIndex = -1 Then
   datosok = False
   MsgBox "Debe ingresar el sexo del cliente"
   ComboSexo.SetFocus
   Exit Function
End If


If Trim(TxtDomicilio.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el domicilio"
   TxtDomicilio.SetFocus
   Exit Function
End If

'si hay localidad
If ComboLocalidades.Text = "" Then
   datosok = False
   MsgE "Debe seleccionar una localidad"
   ComboLocalidades.SetFocus
   Exit Function
End If

'valido el sueldo
If Trim(TxtSueldo.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el sueldo del cliente"
   TxtSueldo.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtSueldo.Text) Then
   datosok = False
   MsgE "El sueldo debe ser numerico"
   TxtSueldo.SetFocus
   Exit Function
End If
If CCur(TxtSueldo.Text) <= 0 Then
   datosok = False
   MsgE "El sueldo debe ser mayor a cero"
   TxtSueldo.SetFocus
   Exit Function
End If
   
'si ingresaron la antiguedad la valido
If Trim(TxtAntiguedad.Text) = "" Then
   TxtAntiguedad.Text = 0
End If

If Not IsNumeric(TxtAntiguedad.Text) Then
   TxtAntiguedad.Text = 0
End If
   
If CLng(TxtAntiguedad.Text) < 0 Then
   TxtAntiguedad.Text = 0
End If
   
If Trim(TxtResidencia.Text) = "" Then
   TxtResidencia.Text = 0
End If

If Not IsNumeric(TxtResidencia.Text) Then
   TxtResidencia.Text = 0
End If

If CLng(TxtResidencia.Text) < 0 Then
   TxtResidencia.Text = 0
End If
   
If CDate(DTPicker1.Value) > Date Then
   DTPicker1.Value = Date
End If

If Trim(TxtNacionalidad.Text) = "" Then
   TxtNacionalidad.Text = vbNullString
End If

TxtTelefono.Text = Trim$(TxtTelefono.Text)
TxtCelular.Text = Trim$(TxtCelular.Text)
TxtTelefonoGarante.Text = Trim$(TxtTelefonoGarante.Text)
TxtTelefonoEmpleador.Text = Trim$(TxtTelefonoEmpleador.Text)

  If Len(TxtTelefono.Text) < 10 Then
     datosok = False
     MsgE "El Telefono debe ser de 10 digitos"
     TxtTelefono.SetFocus
     Exit Function
  Else
    If Not IsNumeric(TxtTelefono.Text) _
           Or InStr(1, TxtTelefono.Text, "e") > 0 _
           Or InStr(1, TxtTelefono.Text, "E") > 0 _
           Or InStr(1, TxtTelefono.Text, "+") > 0 _
           Or InStr(1, TxtTelefono.Text, "-") > 0 _
           Or InStr(1, TxtTelefono.Text, ".") > 0 _
           Or InStr(1, TxtTelefono.Text, ",") > 0 Then
     datosok = False
     MsgE "El Telefono debe ser numerico"
     TxtTelefono.SetFocus
     Exit Function
    End If
  End If

   If Len(TxtCelular.Text) < 10 Then
     datosok = False
     MsgE "El Celular debe ser de 10 digitos"
     TxtCelular.SetFocus
     Exit Function
  Else
    If Not IsNumeric(TxtCelular.Text) _
           Or InStr(1, TxtCelular.Text, "e") > 0 _
           Or InStr(1, TxtCelular.Text, "E") > 0 _
           Or InStr(1, TxtCelular.Text, "+") > 0 _
           Or InStr(1, TxtCelular.Text, "-") > 0 _
           Or InStr(1, TxtCelular.Text, ".") > 0 _
           Or InStr(1, TxtCelular.Text, ",") > 0 Then
     datosok = False
     MsgE "El Celular debe ser numerico"
     TxtCelular.SetFocus
     Exit Function
    End If
  End If

If Trim(TxtObservaciones.Text) = "" Then
   TxtObservaciones.Text = vbNullString
End If

If Trim(TxtObservaciones1.Text) = "" Then
   TxtObservaciones1.Text = vbNullString
End If

If Trim(TxtCuil.Text) = "" Then
   TxtCuil.Text = vbNullString
End If

If Trim(TxtNumCBU.Text) = "" Then
   TxtNumCBU.Text = vbNullString
End If

'empleador
If Trim(TxtCuitEmpleador.Text) = "" Then
   TxtCuitEmpleador.Text = vbNullString
End If

If Trim(TxtCodigoDescuento.Text) = "" Then
   TxtCodigoDescuento.Text = vbNullString
End If

If Trim(TxtTelefonoEmpleador.Text) = "" Then
   TxtTelefonoEmpleador.Text = vbNullString
Else
   If Len(TxtTelefonoEmpleador.Text) < 10 Then
     datosok = False
     MsgE "El Telefono del empleador debe ser de 10 digitos"
     Exit Function
  Else
    If Not IsNumeric(TxtTelefonoEmpleador.Text) _
           Or InStr(1, TxtTelefonoEmpleador.Text, "e") > 0 _
           Or InStr(1, TxtTelefonoEmpleador.Text, "E") > 0 _
           Or InStr(1, TxtTelefonoEmpleador.Text, "+") > 0 _
           Or InStr(1, TxtTelefonoEmpleador.Text, "-") > 0 _
           Or InStr(1, TxtTelefonoEmpleador.Text, ".") > 0 _
           Or InStr(1, TxtTelefonoEmpleador.Text, ",") > 0 Then
     datosok = False
     MsgE "El Telefono del Empleador debe ser numerico"
     Exit Function
    End If
  End If
End If

If Trim(TxtActividad.Text) = "" Then
   TxtActividad.Text = vbNullString
End If

If CHMonotributista.Value = 1 Then
    If CmbMonotributista.Text = "" Then
        datosok = False
        MsgE "Debe ingresar la categoría de monotributo"
        CmbMonotributista.SetFocus
        Exit Function
    End If
    
End If

'datos del garante
If Trim(TxtNombreGarante.Text) = "" Then
   TxtNombreGarante.Text = vbNullString
Else
    If Cmbdiferenciando.Text = "" Then
        datosok = False
        MsgE "Debe seleccionar un tipo de Garante"
        Cmbdiferenciando.SetFocus
        Exit Function
    End If
End If
If Trim(TxtApellidoGarante.Text) = "" Then
   TxtApellidoGarante.Text = vbNullString
End If
If Trim(TxtDocumentoGarante.Text) = "" Then
   TxtDocumentoGarante.Text = vbNullString
End If
If Trim(TxtTelefonoGarante.Text) = "" Then
   TxtTelefonoGarante.Text = vbNullString
Else
  If Len(TxtTelefonoGarante.Text) < 10 Then
     datosok = False
     MsgE "El Telefono del Garante debe ser de 10 digitos"
     Exit Function
  Else
    If Not IsNumeric(TxtTelefonoGarante.Text) _
           Or InStr(1, TxtTelefonoGarante.Text, "e") > 0 _
           Or InStr(1, TxtTelefonoGarante.Text, "E") > 0 _
           Or InStr(1, TxtTelefonoGarante.Text, "+") > 0 _
           Or InStr(1, TxtTelefonoGarante.Text, "-") > 0 _
           Or InStr(1, TxtTelefonoGarante.Text, ".") > 0 _
           Or InStr(1, TxtTelefonoGarante.Text, ",") > 0 Then
     datosok = False
     MsgE "El telefono del Garante debe ser numerico"
     Exit Function
    End If
  End If
End If
If Trim(TxtDomicilioGarante.Text) = "" Then
   TxtDomicilioGarante.Text = vbNullString
End If
If Trim(TxtCuitGarante.Text) = "" Then
   TxtCuitGarante.Text = vbNullString
End If

If Trim(TxtNacionalidadGarante.Text) = "" Then
   TxtNacionalidadGarante.Text = vbNullString
End If
If Trim(TxtProfesionGarante.Text) = "" Then
   TxtProfesionGarante.Text = vbNullString
End If

'valido el sueldo del garante
If Trim(TxtSueldoGarante.Text) = "" Then
   TxtSueldoGarante.Text = 0
End If

If Not IsNumeric(TxtSueldoGarante.Text) Then
   TxtSueldoGarante.Text = 0
End If

If CCur(TxtSueldoGarante.Text) < 0 Then
   TxtSueldoGarante.Text = 0
End If

'si hay localidad

'propiedades
If Trim(TxtCatastroPropiedad.Text) = "" Then
   TxtCatastroPropiedad.Text = vbNullString
End If

If Trim(TxtMetrosPropiedad.Text) = "" Then
   TxtMetrosPropiedad.Text = vbNullString
End If

If Trim(TxtValuacionPropiedad.Text) = "" Then
   TxtValuacionPropiedad.Text = 0
End If

If Not IsNumeric(TxtValuacionPropiedad.Text) Then
   TxtValuacionPropiedad.Text = 0
End If

If Val(TxtValuacionPropiedad.Text) < 0 Then
   TxtValuacionPropiedad.Text = 0
End If

'vehiculo
If Trim(TxtPatenteVehiculo.Text) = "" Then
   TxtPatenteVehiculo.Text = vbNullString
End If

If Trim(TxtMarcaVehiculo.Text) = "" Then
   TxtMarcaVehiculo.Text = vbNullString
End If

If Trim(TxtModeloVehiculo.Text) = "" Then
   TxtModeloVehiculo.Text = vbNullString
End If

If Trim(TxtValuacionVehiculo.Text) = "" Then
   TxtValuacionVehiculo.Text = 0
End If

If Not IsNumeric(TxtValuacionVehiculo.Text) Then
   TxtValuacionVehiculo.Text = 0
End If

If Val(TxtValuacionVehiculo.Text) < 0 Then
   TxtValuacionVehiculo.Text = 0
End If

If Trim(TxtCreditoMaximo.Text) = "" Then
   TxtCreditoMaximo.Text = 0
End If

If Not IsNumeric(TxtCreditoMaximo.Text) Then
   TxtCreditoMaximo.Text = 0
End If
If CCur(TxtCreditoMaximo.Text) = 0 Then
   datosok = False
   MsgE "Debe ingresar el limite maximo de credito"
   TxtCreditoMaximo.SetFocus
   Exit Function
End If

'Datos Factura
If Trim(TxtFac1) = "" Then
   TxtFac1.Text = vbNullString
End If

If Trim(TxtMonto1) = "" Then
   TxtMonto1.Text = "0"
End If

If Trim(TxtFac2) = "" Then
   TxtFac2.Text = vbNullString
End If

If Trim(TxtMonto2) = "" Then
   TxtMonto2.Text = "0"
End If

If Trim(TxtFac3) = "" Then
   TxtFac3.Text = vbNullString
End If

If Trim(TxtMonto3) = "" Then
   TxtMonto3.Text = "0"
End If

If Trim(TxtFac4) = "" Then
   TxtFac4.Text = vbNullString
End If

If Trim(TxtMonto4) = "" Then
   TxtMonto4.Text = "0"
End If

If Trim(TxtFac5) = "" Then
   TxtFac5.Text = vbNullString
End If

If Trim(TxtMonto5) = "" Then
   TxtMonto5.Text = "0"
End If

If Trim(TxtFac6) = "" Then
   TxtFac6.Text = vbNullString
End If

If Trim(TxtMonto6) = "" Then
   TxtMonto6.Text = "0"
End If

If Trim(TxtFac7) = "" Then
   TxtFac7.Text = vbNullString
End If

If Trim(TxtMonto7) = "" Then
   TxtMonto7.Text = "0"
End If

If Not IsNumeric(TxtMonto1.Text) Then
   datosok = False
   MsgE "El monto debe ser numerico"
   TxtMonto1.SetFocus
   Exit Function
End If

If Not IsNumeric(TxtMonto2.Text) Then
   datosok = False
   MsgE "El monto debe ser numerico"
   TxtMonto2.SetFocus
   Exit Function
End If

If Not IsNumeric(TxtMonto3.Text) Then
   datosok = False
   MsgE "El monto debe ser numerico"
   TxtMonto3.SetFocus
   Exit Function
End If

If Not IsNumeric(TxtMonto4.Text) Then
   datosok = False
   MsgE "El monto debe ser numerico"
   TxtMonto4.SetFocus
   Exit Function
End If

If Not IsNumeric(TxtMonto5.Text) Then
   datosok = False
   MsgE "El monto debe ser numerico"
   TxtMonto5.SetFocus
   Exit Function
End If

If Not IsNumeric(TxtMonto6.Text) Then
   datosok = False
   MsgE "El monto debe ser numerico"
   TxtMonto6.SetFocus
   Exit Function
End If

If Not IsNumeric(TxtMonto7.Text) Then
   datosok = False
   MsgE "El monto debe ser numerico"
   TxtMonto7.SetFocus
   Exit Function
End If


'reemplazo comillas no permitidas de los
'campos de pantalla por ej: los apostrofes, y simbolos raros del teclado
Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-AbmClientes"
End Function
Private Sub SetearEntorno()
'habilita o deshabilita los botones correspondientes
On Error GoTo merror
    
Select Case TipoEdicion
       Case "C"
            fmeDatos.Enabled = False
            FrameGarante.Enabled = False
            FramePropiedad.Enabled = False
            FrameVehiculo.Enabled = False
            FrameEmpleador.Enabled = False
            FrameCalificacion.Enabled = False
            FrameObservaciones.Enabled = False
            FrameFacturacion.Enabled = False
            
            Frame1.Enabled = True
            cmdGrabar.Enabled = False
            CmdNuevo.Enabled = True
            CmdTodos.Enabled = True
            CmdRefrescar.Enabled = True
               
            If lv.ListItems.Count > 0 Then
               CmdBorrar.Enabled = True
               cmdModificar.Enabled = True
               CmdImprimir.Enabled = True
               CmdResumen.Enabled = True
               If Trim(FormularioPadre) <> "" Then
                  CmdSeleccionar.Enabled = True
               End If
               If VG_EXPORTA Then
                  CmdExportarClientes.Enabled = True
                  CmdResumen.Enabled = True
                  
               Else
                  CmdExportarClientes.Enabled = False
                  CmdResumen.Enabled = False
               End If
            Else
               cmdModificar.Enabled = False
               CmdBorrar.Enabled = False
               CmdSeleccionar.Enabled = False
               CmdImprimir.Enabled = False
               CmdResumen.Enabled = False
               CmdExportarClientes.Enabled = False
               Call LimpiarCampos(Me)
            End If
            cmdcerrar.Caption = "&Cerrar"
            lv.Enabled = True
            Call ColorCyan(Me)
        Case "M"
            fmeDatos.Enabled = True
            FrameGarante.Enabled = True
            FramePropiedad.Enabled = True
            FrameVehiculo.Enabled = True
            FrameEmpleador.Enabled = True
            FrameCalificacion.Enabled = True
            FrameObservaciones.Enabled = True
            FrameFacturacion.Enabled = True
            Frame1.Enabled = False
            cmdGrabar.Enabled = True
            CmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            CmdBorrar.Enabled = False
            CmdImprimir.Enabled = False
            CmdResumen.Enabled = False
            CmdExportarClientes.Enabled = False
            CmdRefrescar.Enabled = False
            CmdSeleccionar.Enabled = False
            CmdTodos.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lv.Enabled = False
            TxtNombre.SetFocus
            Call ColorBlanco(Me)
        Case "N"
            Call LimpiarCampos(Me)
            fmeDatos.Enabled = True
            FrameGarante.Enabled = True
            FramePropiedad.Enabled = True
            FrameVehiculo.Enabled = True
            FrameEmpleador.Enabled = True
            FrameCalificacion.Enabled = True
            FrameObservaciones.Enabled = True
            FrameFacturacion.Enabled = True
            Frame1.Enabled = False
            cmdGrabar.Enabled = True
            CmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            CmdBorrar.Enabled = False
            CmdImprimir.Enabled = False
            CmdResumen.Enabled = False
            CmdExportarClientes.Enabled = False
            CmdRefrescar.Enabled = False
            CmdSeleccionar.Enabled = False
            CmdTodos.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lv.Enabled = False
            TxtNombre.SetFocus
            Call ColorBlanco(Me)
    End Select
    'si el usuario no administra clientes deshabilito los botones
    If Not VG_ADMCLIENTES Then
       CmdNuevo.Enabled = False
       cmdGrabar.Enabled = False
       cmdModificar.Enabled = False
       CmdBorrar.Enabled = False
    End If
  
Exit Sub
merror:
tratarerrores "Error seteando entorno-ClientesAbm"
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'al clickear sobre el encabezado del listview los datos se ordenan en asc y en desc
Dim Orden As Integer
  
If lv.ListItems.Count > 1 Then
   lv.SortKey = ColumnHeader.Index - 1
   Orden = lv.SortKey
   lv.SortOrder = Abs(Not lv.SortOrder = 1)
   lv.Sorted = True
End If

End Sub
Private Sub lv_DblClick()
Call CmdSeleccionar_Click
End Sub
Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
'al clickear sobre algun elemento del listview los datos se cargan en los dos text
Call CargarDatos
End Sub
Private Sub CmdSeleccionar_Click()
'selecciona un cliente y lo carga en otra pantalla
On Error GoTo merror
Call RefreshTimer

If FormularioPadre = "" Then Exit Sub

If Not VerificarSeleccionLista(lv) Then Exit Sub

If FormularioPadre = "REGISTRARCREDITOS1" Then
   FrmRegistrarCreditos.IdCliente = CLng(lv.SelectedItem)
   FrmRegistrarCreditos.TxtTitular.Text = lv.SelectedItem.SubItems(2)
End If

If FormularioPadre = "CONSULTARDEUDORES" Then
   FrmCreditosConsultarDeudores.IdCliente2 = CLng(lv.SelectedItem)
   FrmCreditosConsultarDeudores.Numlegajocliente2 = CLng(lv.SelectedItem.SubItems(1))
   FrmCreditosConsultarDeudores.TxtCliente.Text = lv.SelectedItem.SubItems(2)
   FrmCreditosConsultarDeudores.NumDni = TxtNumDocumento
End If

If FormularioPadre = "REFINANCIARCREDITOS" Then
   FrmReFinanciarCreditos.IdCliente = CLng(lv.SelectedItem)
   FrmReFinanciarCreditos.TxtCliente.Text = lv.SelectedItem.SubItems(2)
End If

If FormularioPadre = "CONSULTARCREDITOS" Then
   FrmConsultarCreditos.IdCliente = CLng(lv.SelectedItem)
   FrmConsultarCreditos.TxtCliente.Text = lv.SelectedItem.SubItems(2)
End If

If FormularioPadre = "COBROSMASIVOS" Then
   FrmCobrosMasivos.IdCliente = CLng(lv.SelectedItem)
   FrmCobrosMasivos.TxtCliente.Text = lv.SelectedItem.SubItems(2)
End If

Unload Me

Exit Sub
merror:
tratarerrores "Error seleccionando clientes-ClientesAbm"
End Sub
Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
If lv.ListItems.Count() = 0 Then Exit Sub

If KeyCode = vbKeyReturn Then
   Call CmdSeleccionar_Click
End If
End Sub
Private Sub Option1_Click()
TxtCampo.Text = ""
If Option1.Value Then
   TxtCampo.MaxLength = 50
End If
TxtCampo.SetFocus
End Sub
Private Sub Option2_Click()
TxtCampo.Text = ""
If Option2.Value Then
   TxtCampo.MaxLength = 50
End If
TxtCampo.SetFocus
End Sub
Private Sub Option3_Click()
TxtCampo.Text = ""
If Option3.Value Then
   TxtCampo.MaxLength = 50
End If
TxtCampo.SetFocus
End Sub
Private Sub Option4_Click()
TxtCampo.Text = ""
TxtCampo.SetFocus
If Option4.Value Then
   TxtCampo.MaxLength = 50
End If
End Sub
Private Sub Option5_Click()
TxtCampo.Text = ""
If Option5.Value Then
   TxtCampo.MaxLength = 9
End If
TxtCampo.SetFocus
End Sub
'keydowns de campos
Private Sub ComboDomicilios_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   TxtNumDocumento.SetFocus
End If
End Sub
Private Sub TxtActividad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtAntiguedad.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtEmail.SetFocus
End If
End Sub

Private Sub TxtCampo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   If Trim(TxtCampo.Text) <> "" Then
      Call BuscarClientes
   End If
End If
End Sub
Private Sub TxtCelular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtEmail.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtTelefono.SetFocus
End If

End Sub
Private Sub TxtCelular_LostFocus()
TxtCelular.Text = UCase(Trim(TxtCelular.Text))
End Sub
Private Sub TxtCodigoDescuento_LostFocus()
TxtCodigoDescuento.Text = UCase(Trim(TxtCodigoDescuento.Text))
End Sub
Private Sub TxtCP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtDomicilio.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtNumDocumento.SetFocus
End If
End Sub
Private Sub TxtCreditoMaximo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtObservaciones.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtTipoIva.SetFocus
End If

End Sub
Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtActividad.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtCelular.SetFocus
End If

End Sub
Private Sub TxtEmail_LostFocus()
TxtEmail.Text = UCase(Trim(TxtEmail.Text))
End Sub

Private Sub TxtNacionalidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtTelefono.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtCodigoDescuento.SetFocus
End If
End Sub
Private Sub TxtNacionalidadGarante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtProfesionGarante.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtTelefonoGarante.SetFocus
End If
End Sub

Private Sub TxtNumDocumento_LostFocus()
    TxtCuil.Text = IntegrarCUIT()
End Sub

Private Sub TxtNumLegajo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtNumDocumento.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtApellido.SetFocus
End If
End Sub
Private Sub TxtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtApellido.SetFocus
End If
End Sub
Private Sub TxtApellido_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtNumLegajo.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtNombre.SetFocus
End If
End Sub
Private Sub TxtNumDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtCP.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtNumLegajo.SetFocus
End If
End Sub
Private Sub TxtProfesionGarante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   TxtNacionalidadGarante.SetFocus
End If
End Sub
Private Sub TxtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtCelular.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtNacionalidad.SetFocus
End If
End Sub
Private Sub TxtDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtResidencia.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtCP.SetFocus
End If
End Sub
Private Sub TxtResidencia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtCodigoDescuento.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtDomicilio.SetFocus
End If
End Sub
Private Sub TxtCuil_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtNumCBU.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtSueldo.SetFocus
End If
End Sub
Private Sub TxtAntiguedad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtSueldo.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtActividad.SetFocus
End If
End Sub
Private Sub TxtSueldo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtCuil.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtAntiguedad.SetFocus
End If
End Sub
Private Sub TxtNumCBU_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtTipoIva.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtCuil.SetFocus
End If
End Sub
Private Sub TxtCodigoDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtNacionalidad.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtResidencia.SetFocus
End If
End Sub
Private Sub TxtObservaciones_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   If FrameGarante.Visible Then
      TxtNombreGarante.SetFocus
   End If
   
   If FramePropiedad.Visible Then
      TxtCatastroPropiedad.SetFocus
   End If
   
   If FrameVehiculo.Visible Then
      TxtMarcaVehiculo.SetFocus
   End If
   
   If FrameEmpleador.Visible Then
      TxtEmpresa.SetFocus
   End If
   
   If FrameCalificacion.Visible Then
      TxtBancoCentral.SetFocus
   End If
   
   If FrameObservaciones.Visible Then
      TxtObservaciones1.SetFocus
   End If
   
End If

If KeyCode = vbKeyUp Then
   TxtCreditoMaximo.SetFocus
End If
End Sub
'garante
Private Sub TxtApellidoGarante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtDocumentoGarante.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtNombreGarante.SetFocus
End If
End Sub
Private Sub TxtCuitGarante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtTelefonoGarante.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtSueldoGarante.SetFocus
End If
End Sub
Private Sub TxtDocumentoGarante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtDomicilioGarante.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtApellidoGarante.SetFocus
End If
End Sub
Private Sub TxtDomicilioGarante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtSueldoGarante.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtDocumentoGarante.SetFocus
End If
End Sub
Private Sub TxtNombreGarante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtApellidoGarante.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtObservaciones.SetFocus
End If
End Sub
Private Sub TxtSueldoGarante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtCuitGarante.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtDomicilioGarante.SetFocus
End If
End Sub
'empleador
Private Sub TxtEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtCuitEmpleador.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtObservaciones.SetFocus
End If
End Sub
Private Sub TxtCuitEmpleador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtDomicilioEmpleador.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtEmpresa.SetFocus
End If
End Sub
Private Sub TxtDomicilioEmpleador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtTelefonoEmpleador.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtCuitEmpleador.SetFocus
End If
End Sub
Private Sub TxtTelefono_LostFocus()
TxtTelefono.Text = UCase(Trim(TxtTelefono.Text))
End Sub
Private Sub TxtTelefonoEmpleador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   TxtDomicilioEmpleador.SetFocus
End If
End Sub
'propiedades
Private Sub TxtCatastroPropiedad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtValuacionPropiedad.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtObservaciones.SetFocus
End If
End Sub
Private Sub TxtMetrosPropiedad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   TxtValuacionPropiedad.SetFocus
End If
End Sub
Private Sub TxtTipoIva_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtCreditoMaximo.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtNumCBU.SetFocus
End If
End Sub
Private Sub TxtTipoIva_LostFocus()
TxtTipoIva.Text = UCase(Trim(TxtTipoIva.Text))
End Sub
Private Sub TxtValuacionPropiedad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtMetrosPropiedad.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtCatastroPropiedad.SetFocus
End If
End Sub
'vehiculo
Private Sub TxtPatenteVehiculo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtValuacionVehiculo.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtModeloVehiculo.SetFocus
End If

End Sub
Private Sub TxtMarcaVehiculo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtModeloVehiculo.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtObservaciones.SetFocus
End If
End Sub
Private Sub TxtValuacionVehiculo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   TxtPatenteVehiculo.SetFocus
End If
End Sub
Private Sub TxtModeloVehiculo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtPatenteVehiculo.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtMarcaVehiculo.SetFocus
End If
End Sub
'calificacion
Private Sub TxtBancoCentral_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtAnses.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtObservaciones.SetFocus
End If
End Sub
Private Sub TxtAnses_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtJudicial.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtBancoCentral.SetFocus
End If
End Sub
Private Sub TxtJudicial_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtAfip.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtAnses.SetFocus
End If
End Sub
Private Sub TxtAfip_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   TxtJudicial.SetFocus
End If
End Sub
Private Sub TxtTelefonoGarante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtNacionalidadGarante.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtCuitGarante.SetFocus
End If
End Sub
'mayusculas de campos
Private Sub TxtNombre_LostFocus()

TxtNombre.Text = UCase(Trim(TxtNombre.Text))
End Sub
Private Sub TxtApellido_LostFocus()
TxtApellido.Text = UCase(Trim(TxtApellido.Text))
End Sub
Private Sub TxtDomicilio_LostFocus()
TxtDomicilio.Text = UCase(Trim(TxtDomicilio.Text))
End Sub
Private Sub TxtObservaciones_LostFocus()
TxtObservaciones.Text = UCase(Trim(TxtObservaciones.Text))
End Sub
Private Sub TxtNumLegajo_LostFocus()
TxtNumLegajo.Text = UCase(Trim(TxtNumLegajo.Text))
End Sub
Private Sub TxtActividad_LostFocus()
TxtActividad.Text = UCase(Trim(TxtActividad.Text))
End Sub
Private Sub TxtNacionalidad_LostFocus()
TxtNacionalidad.Text = UCase(Trim(TxtNacionalidad.Text))
End Sub
Private Sub TxtNombreGarante_LostFocus()
TxtNombreGarante.Text = UCase(Trim(TxtNombreGarante.Text))
End Sub
Private Sub TxtApellidoGarante_LostFocus()
TxtApellidoGarante.Text = UCase(Trim(TxtApellidoGarante.Text))
End Sub
Private Sub TxtDomicilioGarante_LostFocus()
TxtDomicilioGarante.Text = UCase(Trim(TxtDomicilioGarante.Text))
End Sub
Private Sub TxtTelefonoGarante_LostFocus()
TxtTelefonoGarante.Text = UCase(Trim(TxtTelefonoGarante.Text))
End Sub
Private Sub TxtNacionalidadGarante_LostFocus()
TxtNacionalidadGarante.Text = UCase(Trim(TxtNacionalidadGarante.Text))
End Sub
Private Sub TxtProfesionGarante_LostFocus()
TxtProfesionGarante.Text = UCase(Trim(TxtProfesionGarante.Text))
End Sub
Private Sub TxtCuitGarante_LostFocus()
TxtCuitGarante.Text = UCase(Trim(TxtCuitGarante.Text))
End Sub
Private Sub TxtDocumentoGarante_LostFocus()
TxtDocumentoGarante.Text = UCase(Trim(TxtDocumentoGarante.Text))
End Sub
Private Sub TxtPatenteVehiculo_LostFocus()
TxtPatenteVehiculo.Text = UCase(Trim(TxtPatenteVehiculo.Text))
End Sub
Private Sub TxtMarcaVehiculo_LostFocus()
TxtMarcaVehiculo.Text = UCase(Trim(TxtMarcaVehiculo.Text))
End Sub
Private Sub TxtModeloVehiculo_LostFocus()
TxtModeloVehiculo.Text = UCase(Trim(TxtModeloVehiculo.Text))
End Sub
Private Sub TxtBancoCentral_LostFocus()
TxtBancoCentral.Text = UCase(Trim(TxtBancoCentral.Text))
End Sub
Private Sub TxtAnses_LostFocus()
TxtAnses.Text = UCase(Trim(TxtAnses.Text))
End Sub
Private Sub TxtJudicial_LostFocus()
TxtJudicial.Text = UCase(Trim(TxtJudicial.Text))
End Sub
Private Sub TxtAfip_LostFocus()
TxtAfip.Text = UCase(Trim(TxtAfip.Text))
End Sub
Private Sub TxtCampo_LostFocus()
TxtCampo.Text = UCase(Trim(TxtCampo.Text))
End Sub
Private Sub TxtEmpresa_LostFocus()
TxtEmpresa.Text = UCase(Trim(TxtEmpresa.Text))
End Sub
Private Sub TxtCuitEmpleador_LostFocus()
TxtCuitEmpleador.Text = UCase(Trim(TxtCuitEmpleador.Text))
End Sub
Private Sub TxtDomicilioEmpleador_LostFocus()
TxtDomicilioEmpleador.Text = UCase(Trim(TxtDomicilioEmpleador.Text))
End Sub
Private Sub TxtTelefonoEmpleador_LostFocus()
TxtTelefonoEmpleador.Text = UCase(Trim(TxtTelefonoEmpleador.Text))
End Sub
Private Sub TxtCatastroPropiedad_LostFocus()
TxtCatastroPropiedad.Text = UCase(Trim(TxtCatastroPropiedad.Text))
End Sub
Private Sub TabStrip1_Click()
On Error GoTo merror

FrameGarante.Visible = False
FramePropiedad.Visible = False
FrameVehiculo.Visible = False
FrameEmpleador.Visible = False
FrameCalificacion.Visible = False
FrameObservaciones.Visible = False
FrameFacturacion.Visible = False

'garante
If TabStrip1.SelectedItem.Index = 1 Then
   FrameGarante.Visible = True
   TabStrip1.Height = 2175
   FrmClientesAbm.Height = 8505
End If

'propiedad
If TabStrip1.SelectedItem.Index = 2 Then
   FramePropiedad.Visible = True
   TabStrip1.Height = 2175
   FrmClientesAbm.Height = 8505
End If

'vehiculo
If TabStrip1.SelectedItem.Index = 3 Then
   FrameVehiculo.Visible = True
   TabStrip1.Height = 2175
   FrmClientesAbm.Height = 8505
End If

'empleador
If TabStrip1.SelectedItem.Index = 4 Then
   FrameEmpleador.Visible = True
   TabStrip1.Height = 2175
   FrmClientesAbm.Height = 8505
End If

'calificacion
If TabStrip1.SelectedItem.Index = 5 Then
   FrameCalificacion.Visible = True
   TabStrip1.Height = 2175
   FrmClientesAbm.Height = 8505
End If

'Observaciones
If TabStrip1.SelectedItem.Index = 6 Then
   FrameObservaciones.Visible = True
   TabStrip1.Height = 2175
   FrmClientesAbm.Height = 8505
End If

'Datos Facturacion
If TabStrip1.SelectedItem.Index = 7 Then
   FrameFacturacion.Visible = True
   TabStrip1.Height = 3000
   FrmClientesAbm.Height = 9345
End If


Exit Sub
merror:
tratarerrores "Error seleccionando solapas-AbmClientes"
End Sub
Private Sub CmdExportarClientes_Click()
Call RefreshTimer

CmdExportarClientes.Enabled = False
Me.MousePointer = vbHourglass
Call ExportarClientesTXT
'Call ExportarClientes
CmdExportarClientes.Enabled = True
Me.MousePointer = vbDefault
End Sub

Private Sub ExportarClientesTXT()
Dim Lin As String, n As Integer
Dim rec As rdoResultset
Dim rec1 As rdoResultset
Dim SexoSTR As String
Dim StrActividad, StrJubilado As String
Dim StrCategoria As String
Dim I As Integer
Dim ObservacioneSolapa As String


On Error GoTo merror


If lv.ListItems.Count() = 0 Then Exit Sub

Dia = Format(CStr(Day(Date)), "00")
Mes = Format(CStr(Month(Date)), "00")
Ano = Format(CStr(Year(Date)), "0000")

Archi = "Clientes"

Archi = Archi + "-" + Dia + "-" + Mes + "-" + Ano + ".xls"

If Not MsgP("¿Confirma la exportacion del resumen de Clientes hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & " ?") Then Exit Sub

'inicio transaccion
cnSQL.BeginTrans

'si no existe me voy pero eso es imposible porque si no existe la crea dentro
If Not ExisteCarpeta() Then Exit Sub

'verifico si ya existe en la ubicacion
Archivo = Dir("c:\exportacionexcel\" & Archi)

If Trim(Archivo) <> "" Then
  'si ya existe lo borro para que despues no haya errores con la pantallita
  'que despliega el excel
   Kill ("c:\exportacionexcel\" & Archi)
End If


Open "c:\exportacionexcel\" & Archi For Output As #1

'titulo principal
Print #1, "Resumen de clientes a la fecha:" & CStr(Date)
Print #1, "Cod.Cliente"; Chr$(9); "Nro.Cliente"; Chr$(9); "Nombre y Apellido"; Chr$(9); "Sexo"; Chr$(9); "Documento"; Chr$(9); "Telefono"; Chr$(9); "Celular"; Chr$(9); "EMail"; Chr$(9); "Domicilio"; Chr$(9); "Ciudad"; Chr$(9); "Provincia"; Chr$(9); "Nacionalidad"; Chr$(9); "CP"; Chr$(9); "Residencia"; Chr$(9); "Profesion"; Chr$(9); "CUIL"; Chr$(9); "AntLaboral"; Chr$(9); "Sueldo"; Chr$(9); "LimiteCredito"; Chr$(9); "SaldoCredito"; Chr$(9); "Observaciones"; Chr$(9); "Empleador"; Chr$(9); "DomEmpleador"; Chr$(9); "TelEmpleador"; Chr$(9); "Actividad"; Chr$(9); "Es Monotributista"; Chr$(9); "Categoria"; Chr$(9); "Es Jubilado/Pensionado"; Chr$(9); "Garante"; Chr$(9); "DocGarante"; Chr$(9); "DomGarante"; Chr$(9); "TelGarante"; Chr$(9); "m2 propiedad"; Chr$(9); "Nro catastro"; Chr$(9); "Valuacion"; Chr$(9); "MarcaVehiculo"; Chr$(9); "Modelo"; Chr$(9); "Patente"; Chr$(9); "Valuacion"; Chr$(9); "Banco Central"; Chr$(9); "Anses"; Chr$(9); "Afip"; Chr$(9); "Judicial"; Chr$(9); "Observaciones pestaña"; _
Chr$(9); "Nro Factura"; Chr$(9); "Monto"; Chr$(9); "Nro Factura"; Chr$(9); "Monto"; Chr$(9); "Nro Factura"; Chr$(9); "Monto"; Chr$(9); "Nro Factura"; Chr$(9); "Monto"; Chr$(9); "Nro Factura"; Chr$(9); "Monto"; Chr$(9); "Nro Factura"; Chr$(9); "Monto"; Chr$(9); "Nro Factura"; Chr$(9); "Monto"; Chr$(9); "Tipo Garante"

'cargo el registro de creditos segun las condiciones de la pantalla

Condicion = "1=1"
Set rec = CargarRecClientes(Condicion)
Call RefreshTimer

If Not rec.EOF Then
   Do While Not rec.EOF
      'grabo cada usuario en la planilla
       If rec.rdoColumns("cad1") = "M" Then
        SexoSTR = "Masculino"
      ElseIf rec.rdoColumns("cad1") = "F" Then
        SexoSTR = "Femenino"
      End If
      
      If Not IsNull(rec.rdoColumns("actividad")) Then
        StrActividad = rec.rdoColumns("actividad") & vbNullString
      End If
      
      If Not IsNull(rec.rdoColumns("monotributista")) Then
        If rec.rdoColumns("monotributista") Then
          StrMonotibutista = "SI"
        Else
          StrMonotibutista = "NO"
        End If
      End If
      
      If Not IsNull(rec.rdoColumns("catmonotributo")) Then
        StrCategoria = rec.rdoColumns("catmonotributo") & vbNullString
      End If
      
      If Not IsNull(rec.rdoColumns("jubilado")) Then
        If rec.rdoColumns("jubilado") Then
          StrJubilado = "SI"
        Else
          StrJubilado = "NO"
        End If
      End If
      
      If Not IsNull(rec.rdoColumns("observaciones1")) Then
        ObservacioneSolapa = ReemplazarEnter(rec.rdoColumns("observaciones1"))
      End If
      
      'Obtengo datos de facturacion
      sql = "SELECT * FROM DatosFactura WHERE IdCliente = " & rec.rdoColumns("idcliente") & ""
      Set rec1 = cnSQL.OpenResultset(sql)
      I = 1
      Do While Not rec1.EOF
        
        Select Case I
            Case 1
                If Not IsNull(rec1.rdoColumns("nrofactura")) Then
                  StrNumFactura1 = Trim(rec1.rdoColumns("nrofactura")) & vbNullString
                End If
                If Not IsNull(rec1.rdoColumns("monto")) Then
                  StrMontoFact1 = Trim(rec1.rdoColumns("monto")) & vbNullString
                End If
            Case 2
                If Not IsNull(rec1.rdoColumns("nrofactura")) Then
                   StrNumFactura2 = Trim(rec1.rdoColumns("nrofactura")) & vbNullString
                End If
                If Not IsNull(rec1.rdoColumns("monto")) Then
                   StrMontoFact2 = Trim(rec1.rdoColumns("monto")) & vbNullString
                End If
            Case 3
                If Not IsNull(rec1.rdoColumns("nrofactura")) Then
                   StrNumFactura3 = Trim(rec1.rdoColumns("nrofactura")) & vbNullString
                End If
                If Not IsNull(rec1.rdoColumns("monto")) Then
                   StrMontoFact3 = Trim(rec1.rdoColumns("monto")) & vbNullString
                End If
             Case 4
                If Not IsNull(rec1.rdoColumns("nrofactura")) Then
                   StrNumFactura4 = Trim(rec1.rdoColumns("nrofactura")) & vbNullString
                End If
                If Not IsNull(rec1.rdoColumns("monto")) Then
                   StrMontoFact4 = Trim(rec1.rdoColumns("monto")) & vbNullString
                End If
             Case 5
                If Not IsNull(rec1.rdoColumns("nrofactura")) Then
                   StrNumFactura5 = Trim(rec1.rdoColumns("nrofactura")) & vbNullString
                End If
                If Not IsNull(rec1.rdoColumns("monto")) Then
                   StrMontoFact5 = Trim(rec1.rdoColumns("monto")) & vbNullString
                End If
             Case 6
                If Not IsNull(rec1.rdoColumns("nrofactura")) Then
                   StrNumFactura6 = Trim(rec1.rdoColumns("nrofactura")) & vbNullString
                End If
                If Not IsNull(rec1.rdoColumns("monto")) Then
                   StrMontoFact6 = Trim(rec1.rdoColumns("monto")) & vbNullString
                End If
             Case 7
                If Not IsNull(rec1.rdoColumns("nrofactura")) Then
                   StrNumFactura7 = Trim(rec1.rdoColumns("nrofactura")) & vbNullString
                End If
                If Not IsNull(rec1.rdoColumns("monto")) Then
                   StrMontoFact7 = Trim(rec1.rdoColumns("monto")) & vbNullString
                End If
        End Select
        
        I = I + 1
        rec1.MoveNext
     Loop
     rec1.Close
     'Fin obtener Datos Facturacion
            
      Print #1, rec.rdoColumns("idcliente") & vbNullString; Chr$(9); rec.rdoColumns("numlegajo") & vbNullString; Chr$(9); rec.rdoColumns("cliente") & vbNullString; Chr$(9); SexoSTR; Chr$(9); rec.rdoColumns("numdocumento") & vbNullString; Chr$(9); rec.rdoColumns("telefono") & vbNullString; Chr$(9); rec.rdoColumns("observaciones"); Chr$(9); rec.rdoColumns("cad3") & vbNullString; Chr$(9); rec.rdoColumns("domicilio") & vbNullString; Chr$(9); rec.rdoColumns("localidad"); Chr$(9); rec.rdoColumns("provincia"); Chr$(9); rec.rdoColumns("nacionalidad"); Chr$(9); rec.rdoColumns("codigopostal"); Chr$(9); rec.rdoColumns("residencia"); Chr$(9); rec.rdoColumns("profesion"); Chr$(9); rec.rdoColumns("cuil") & vbNullString; Chr$(9); rec.rdoColumns("antiguedad") & vbNullString _
      ; Chr$(9); rec.rdoColumns("sueldo"); Chr$(9); rec.rdoColumns("creditomaximo"); Chr$(9); ObtenerSaldoCliente(rec.rdoColumns("idcliente")); Chr$(9); rec.rdoColumns("cad2") & vbNullString; Chr$(9); rec.rdoColumns("empresa"); Chr$(9); rec.rdoColumns("domicilioempleador"); Chr$(9); rec.rdoColumns("telefonoempleador") & vbNullString; Chr$(9); StrActividad; Chr$(9); StrMonotibutista; Chr$(9); StrCategoria _
      ; Chr$(9); StrJubilado; Chr$(9); rec.rdoColumns("garante"); Chr$(9); rec.rdoColumns("documentogarante"); Chr$(9); rec.rdoColumns("domiciliogarante"); Chr$(9); rec.rdoColumns("telefonogarante"); Chr$(9); rec.rdoColumns("metrospropiedad"); Chr$(9); rec.rdoColumns("catastropropiedad"); Chr$(9); rec.rdoColumns("valuacionpropiedad"); Chr$(9); rec.rdoColumns("marcavehiculo"); Chr$(9); rec.rdoColumns("modelovehiculo"); Chr$(9); rec.rdoColumns("patentevehiculo"); Chr$(9); rec.rdoColumns("valuacionvehiculo"); Chr$(9); rec.rdoColumns("BancoCentral"); Chr$(9); rec.rdoColumns("anses"); Chr$(9); rec.rdoColumns("Afip"); Chr$(9); rec.rdoColumns("judicial"); Chr$(9); ObservacioneSolapa & vbNullString; Chr$(9); StrNumFactura1; Chr$(9); StrMontoFact1; Chr$(9); StrNumFactura2; Chr$(9); StrMontoFact2; Chr$(9); StrNumFactura3; Chr$(9); StrMontoFact3; Chr$(9); StrNumFactura4; Chr$(9); StrMontoFact4; Chr$(9); StrNumFactura5; Chr$(9); StrMontoFact5; Chr$(9); StrNumFactura6; Chr$(9); StrMontoFact6 _
      ; Chr$(9); StrNumFactura7; Chr$(9); StrMontoFact7; Chr$(9); rec.rdoColumns("TipoGarante") & vbNullString
             
      rec.MoveNext
      Call RefreshTimer

   StrNumFactura1 = ""
   StrMontoFact1 = ""
   StrNumFactura2 = ""
   StrMontoFact2 = ""
   StrNumFactura3 = ""
   StrMontoFact3 = ""
   StrNumFactura4 = ""
   StrMontoFact4 = ""
   StrNumFactura5 = ""
   StrMontoFact5 = ""
   StrNumFactura6 = ""
   StrMontoFact6 = ""
   StrNumFactura7 = ""
   StrMontoFact7 = ""
   ObservacioneSolapa = ""
      
   Loop
   rec.Close
   Close #1
   
   Mensaje = "Se exporto la lista de clientes...a la planilla C:\ExportacionExcel\" & Archi
Else
   Mensaje = "No hay datos para exportar"
End If
   
'fin de transaccion
cnSQL.CommitTrans

MsgI Mensaje

Exit Sub
merror:
tratarerrores "Error Exportando clientes...verifique que los archivos de Excel esten cerrados"
End Sub

Private Sub ExportarClientes()
'exporta todos los clientes
Dim MiExcel As Excel.APPLICATION ' Aplicacion Excel
Dim MiLibro As Excel.Workbook ' Un libro de Excel
Dim MiHoja As Excel.Worksheet ' Una hoja de Excel
Dim Filas As Long
Dim FilaTitulos As Long
Dim Archivo As String
Dim Archi As String
Dim rec As rdoResultset
Dim rec1 As rdoResultset
Dim sql As String
Dim SaldoCredito As Currency
Dim SaldoCapital As Currency
Dim CuotasCobradas As Long
Dim CuotasVencidas As Long
Dim CuotasPendientes As Long
Dim Fecha As Date
Dim Dia As String
Dim Mes As String
Dim Ano As String
Dim Mensaje As String
Dim columna As Integer
On Error GoTo merror

If lv.ListItems.Count() = 0 Then Exit Sub

Dia = Format(CStr(Day(Date)), "00")
Mes = Format(CStr(Month(Date)), "00")
Ano = Format(CStr(Year(Date)), "0000")

Archi = "Clientes"

Archi = Archi + "-" + Dia + "-" + Mes + "-" + Ano + ".xls"

If Not MsgP("¿Confirma la exportacion del resumen de Clientes hacia la carpeta C:\EXPORTACIONEXCEL\" & Archi & " ?") Then Exit Sub

'inicio transaccion
cnSQL.BeginTrans

'si no existe me voy pero eso es imposible porque si no existe la crea dentro
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
MiHoja.Cells(1, 1).Value = "Resumen de clientes a la fecha:" & CStr(Date)

'establezco el ancho de todas las columnas
MiHoja.Columns.ColumnWidth = 17


FilaTitulos = 2

MiHoja.Cells(FilaTitulos, 1).Value = "Cod.Cliente"
MiHoja.Cells(FilaTitulos, 2).Value = "Nro.Cliente"
MiHoja.Cells(FilaTitulos, 3).Value = "Nombre y Apellido"
MiHoja.Cells(FilaTitulos, 4).Value = "Sexo"
MiHoja.Cells(FilaTitulos, 5).Value = "Documento"
MiHoja.Cells(FilaTitulos, 6).Value = "Telefono"
MiHoja.Cells(FilaTitulos, 7).Value = "Celular"
MiHoja.Cells(FilaTitulos, 8).Value = "EMail"
MiHoja.Cells(FilaTitulos, 9).Value = "Domicilio"
MiHoja.Cells(FilaTitulos, 10).Value = "Ciudad"
MiHoja.Cells(FilaTitulos, 11).Value = "Provincia"
MiHoja.Cells(FilaTitulos, 12).Value = "Nacionalidad"
MiHoja.Cells(FilaTitulos, 13).Value = "CP"
MiHoja.Cells(FilaTitulos, 14).Value = "Residencia"
MiHoja.Cells(FilaTitulos, 15).Value = "Profesion"
MiHoja.Cells(FilaTitulos, 16).Value = "CUIL"
MiHoja.Cells(FilaTitulos, 17).Value = "AntLaboral"
MiHoja.Cells(FilaTitulos, 18).Value = "Sueldo"
MiHoja.Cells(FilaTitulos, 19).Value = "LimiteCredito"
MiHoja.Cells(FilaTitulos, 20).Value = "SaldoCredito"
MiHoja.Cells(FilaTitulos, 21).Value = "Observaciones"

'empleador
MiHoja.Cells(FilaTitulos, 22).Value = "Empleador"
MiHoja.Cells(FilaTitulos, 23).Value = "DomEmpleador"
MiHoja.Cells(FilaTitulos, 24).Value = "TelEmpleador"
MiHoja.Cells(FilaTitulos, 25).Value = "Actividad"
MiHoja.Cells(FilaTitulos, 26).Value = "Es Monotributista"
MiHoja.Cells(FilaTitulos, 27).Value = "Categoria"
MiHoja.Cells(FilaTitulos, 28).Value = "Es Jubilado/Pensionado"

'garante
MiHoja.Cells(FilaTitulos, 29).Value = "Garante"
MiHoja.Cells(FilaTitulos, 30).Value = "DocGarante"
MiHoja.Cells(FilaTitulos, 31).Value = "DomGarante"
MiHoja.Cells(FilaTitulos, 32).Value = "TelGarante"
MiHoja.Cells(FilaTitulos, 33).Value = "m2 propiedad"
MiHoja.Cells(FilaTitulos, 34).Value = "Nro catastro"
MiHoja.Cells(FilaTitulos, 35).Value = "Valuacion"
MiHoja.Cells(FilaTitulos, 36).Value = "MarcaVehiculo"
MiHoja.Cells(FilaTitulos, 37).Value = "Modelo"
MiHoja.Cells(FilaTitulos, 38).Value = "Patente"
MiHoja.Cells(FilaTitulos, 39).Value = "Valuacion"

'calificaciones
MiHoja.Cells(FilaTitulos, 40).Value = "Banco Central"
MiHoja.Cells(FilaTitulos, 41).Value = "Anses"
MiHoja.Cells(FilaTitulos, 42).Value = "Afip"
MiHoja.Cells(FilaTitulos, 43).Value = "Judicial"

MiHoja.Cells(FilaTitulos, 44).Value = "Observaciones pestaña"

'Datos Facturación
MiHoja.Cells(FilaTitulos, 45).Value = "Nro Factura"
MiHoja.Cells(FilaTitulos, 46).Value = "Monto"
MiHoja.Cells(FilaTitulos, 47).Value = "Nro Factura"
MiHoja.Cells(FilaTitulos, 48).Value = "Monto"
MiHoja.Cells(FilaTitulos, 49).Value = "Nro Factura"
MiHoja.Cells(FilaTitulos, 50).Value = "Monto"
MiHoja.Cells(FilaTitulos, 51).Value = "Nro Factura"
MiHoja.Cells(FilaTitulos, 52).Value = "Monto"
MiHoja.Cells(FilaTitulos, 53).Value = "Nro Factura"
MiHoja.Cells(FilaTitulos, 54).Value = "Monto"
MiHoja.Cells(FilaTitulos, 55).Value = "Nro Factura"
MiHoja.Cells(FilaTitulos, 56).Value = "Monto"
MiHoja.Cells(FilaTitulos, 57).Value = "Nro Factura"
MiHoja.Cells(FilaTitulos, 58).Value = "Monto"
MiHoja.Cells(FilaTitulos, 59).Value = "Tipo Garante"




'cargo el registro de creditos segun las condiciones de la pantalla

Condicion = "1=1"
Set rec = CargarRecClientes(Condicion)

'comienzo en la fila 3 porque las anteriores son de titulos
Filas = 3
If Not rec.EOF Then
   Do While Not rec.EOF
      'grabo cada usuario en la planilla
      MiHoja.Cells(Filas, 1).Value = rec.rdoColumns("idcliente") & vbNullString
      MiHoja.Cells(Filas, 2).Value = rec.rdoColumns("numlegajo") & vbNullString
      MiHoja.Cells(Filas, 3).Value = rec.rdoColumns("cliente") & vbNullString
      If rec.rdoColumns("cad1") = "M" Then
        MiHoja.Cells(Filas, 4).Value = "Masculino"
      ElseIf rec.rdoColumns("cad1") = "F" Then
        MiHoja.Cells(Filas, 4).Value = "Femenino"
      End If
      
      
      MiHoja.Cells(Filas, 5).Value = rec.rdoColumns("numdocumento") & vbNullString
      
      MiHoja.Cells(Filas, 6).Value = rec.rdoColumns("telefono") & vbNullString
      MiHoja.Cells(Filas, 7).Value = rec.rdoColumns("observaciones") & vbNullString
      MiHoja.Cells(Filas, 8).Value = rec.rdoColumns("cad3") & vbNullString
      
      MiHoja.Cells(Filas, 9).Value = rec.rdoColumns("domicilio") & vbNullString
      MiHoja.Cells(Filas, 10).Value = rec.rdoColumns("localidad") & vbNullString
      
      MiHoja.Cells(Filas, 11).Value = rec.rdoColumns("provincia") & vbNullString
      
      MiHoja.Cells(Filas, 12).Value = rec.rdoColumns("nacionalidad") & vbNullString
      MiHoja.Cells(Filas, 13).Value = rec.rdoColumns("codigopostal") & vbNullString
      MiHoja.Cells(Filas, 14).Value = rec.rdoColumns("residencia") & vbNullString
      MiHoja.Cells(Filas, 15).Value = rec.rdoColumns("profesion") & vbNullString
      MiHoja.Cells(Filas, 16).Value = rec.rdoColumns("cuil") & vbNullString
      MiHoja.Cells(Filas, 17).Value = rec.rdoColumns("antiguedad") & vbNullString
      MiHoja.Cells(Filas, 18).Value = rec.rdoColumns("sueldo") & vbNullString
      MiHoja.Cells(Filas, 19).Value = rec.rdoColumns("creditomaximo") & vbNullString
      Saldo = ObtenerSaldoCliente(rec.rdoColumns("idcliente"))
      MiHoja.Cells(Filas, 20).Value = Format(Saldo, "0.00") & vbNullString
      MiHoja.Cells(Filas, 21).Value = rec.rdoColumns("cad2") & vbNullString
      
      'empleador
      MiHoja.Cells(Filas, 22).Value = rec.rdoColumns("empresa") & vbNullString
      MiHoja.Cells(Filas, 23).Value = rec.rdoColumns("domicilioempleador") & vbNullString
      MiHoja.Cells(Filas, 24).Value = rec.rdoColumns("telefonoempleador") & vbNullString
      
      If Not IsNull(rec.rdoColumns("actividad")) Then
        MiHoja.Cells(Filas, 25).Value = rec.rdoColumns("actividad") & vbNullString
      End If
      
      If Not IsNull(rec.rdoColumns("monotributista")) Then
        If rec.rdoColumns("monotributista") Then
          MiHoja.Cells(Filas, 26).Value = "SI"
        Else
          MiHoja.Cells(Filas, 26).Value = "NO"
        End If
      End If
      If Not IsNull(rec.rdoColumns("catmonotributo")) Then
        MiHoja.Cells(Filas, 27).Value = rec.rdoColumns("catmonotributo") & vbNullString
      End If
      
      If Not IsNull(rec.rdoColumns("jubilado")) Then
        If rec.rdoColumns("jubilado") Then
          MiHoja.Cells(Filas, 28).Value = "SI"
        Else
          MiHoja.Cells(Filas, 28).Value = "NO"
        End If
      End If
      'garante
      MiHoja.Cells(Filas, 29).Value = rec.rdoColumns("garante") & vbNullString
      MiHoja.Cells(Filas, 30).Value = rec.rdoColumns("documentogarante") & vbNullString
      MiHoja.Cells(Filas, 31).Value = rec.rdoColumns("domiciliogarante") & vbNullString
      MiHoja.Cells(Filas, 32).Value = rec.rdoColumns("telefonogarante") & vbNullString
      
      MiHoja.Cells(Filas, 33).Value = rec.rdoColumns("metrospropiedad") & vbNullString
      MiHoja.Cells(Filas, 34).Value = rec.rdoColumns("catastropropiedad") & vbNullString
      MiHoja.Cells(Filas, 35).Value = rec.rdoColumns("valuacionpropiedad") & vbNullString
      
      MiHoja.Cells(Filas, 36).Value = rec.rdoColumns("marcavehiculo") & vbNullString
      MiHoja.Cells(Filas, 37).Value = rec.rdoColumns("modelovehiculo") & vbNullString
      MiHoja.Cells(Filas, 38).Value = rec.rdoColumns("patentevehiculo") & vbNullString
      MiHoja.Cells(Filas, 39).Value = rec.rdoColumns("valuacionvehiculo") & vbNullString
      
      'calificacion
      MiHoja.Cells(Filas, 40).Value = rec.rdoColumns("BancoCentral") & vbNullString
      MiHoja.Cells(Filas, 41).Value = rec.rdoColumns("Anses") & vbNullString
      MiHoja.Cells(Filas, 42).Value = rec.rdoColumns("Afip") & vbNullString
      MiHoja.Cells(Filas, 43).Value = rec.rdoColumns("Judicial") & vbNullString
      MiHoja.Cells(Filas, 44).Value = rec.rdoColumns("observaciones1") & vbNullString
      
      MiHoja.Cells(Filas, 59).Value = rec.rdoColumns("TipoGarante") & vbNullString
      
      
      'Obtengo datos de facturacion
     columna = 45
     sql = "SELECT * FROM DatosFactura WHERE IdCliente = " & rec.rdoColumns("idcliente") & ""
     Set rec1 = cnSQL.OpenResultset(sql)
     Do While Not rec1.EOF
          If Not IsNull(rec1.rdoColumns("nrofactura")) Then
            MiHoja.Cells(Filas, columna).Value = Trim(rec1.rdoColumns("nrofactura")) & vbNullString
          End If
          If Not IsNull(rec1.rdoColumns("monto")) Then
            MiHoja.Cells(Filas, columna + 1).Value = Trim(rec1.rdoColumns("monto")) & vbNullString
          End If
        columna = columna + 2
        rec1.MoveNext
     Loop
     rec1.Close
   'Fin obtener Datos Facturacion
      
      
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
   Mensaje = "Se exporto la lista de clientes...a la planilla C:\ExportacionExcel\" & Archi
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
tratarerrores "Error Exportando clientes...verifique que los archivos de Excel esten cerrados"
End Sub


