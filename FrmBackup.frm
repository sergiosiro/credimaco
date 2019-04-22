VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBackup 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de la base de datos (Backup)"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   HelpContextID   =   3
   Icon            =   "FrmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCopiar 
      Caption         =   "Co&piar la base de datos"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Copia la base de datos del sistema al lugar elegido por el usuario"
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton CmdRestaurar 
      Caption         =   "&Restituir la base de datos"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      ToolTipText     =   "Restituye una copia de seguridad efectuada anteriormente"
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      ToolTipText     =   "Cierra la pantalla"
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Frame FrameCopiar 
      Height          =   6255
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   8175
      Begin VB.Frame Frame3 
         Height          =   2055
         Left            =   4800
         TabIndex        =   23
         Top             =   1920
         Width           =   3255
         Begin VB.Label Label1 
            Caption         =   "(*)Se recomienda hacer periodicamente un backup hacia un pendrive desde esta misma pantalla."
            ForeColor       =   &H00FF0000&
            Height          =   1695
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Opciones:"
         Height          =   855
         Left            =   4800
         TabIndex        =   22
         Top             =   960
         Width           =   3255
         Begin VB.CheckBox CheckBorrarAnterior 
            Caption         =   "Borrar copias de seguridad anteriores"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Al efectuar la copia de seguridad borrara las copias anteriores"
            Top             =   480
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   720
         Picture         =   "FrmBackup.frx":030A
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   16
         Top             =   5160
         Width           =   615
      End
      Begin VB.Frame Frame2 
         Caption         =   "Carpeta donde se guardara la copia de seguridad:"
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   7935
         Begin VB.TextBox TxtArchivoDestino 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   4
            ToolTipText     =   "Destino y nombre de la copia de seguridad"
            Top             =   240
            Width           =   7695
         End
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   5280
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame Frame5 
         Caption         =   "Seleccionar la carpeta destino de la copia:"
         Height          =   3015
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   4575
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            ToolTipText     =   "Selecciona la unidad donde se ubicara la copia de seguridad"
            Top             =   360
            Width           =   4335
         End
         Begin VB.DirListBox Dir1 
            Height          =   2115
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Selecciona la carpeta donde se ubicara la copia de seguridad"
            Top             =   720
            Width           =   4335
         End
      End
      Begin VB.Label Label2 
         Caption         =   "El Backup copia la base de datos a un lugar seguro como por ejemplo un PenDrive o carpeta del disco rigido."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Frame FrameReponer 
      Height          =   6255
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   8175
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   480
         Picture         =   "FrmBackup.frx":5F1C
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   20
         Top             =   5280
         Width           =   615
      End
      Begin VB.Frame Frame6 
         Caption         =   "Restaurar base desde:"
         Height          =   3615
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   7935
         Begin VB.FileListBox File1 
            Height          =   2430
            Left            =   4080
            TabIndex        =   7
            ToolTipText     =   "Selecciona cual copia restituir"
            Top             =   240
            Width           =   3735
         End
         Begin VB.Frame Frame8 
            Caption         =   "Nombre y ubicacion de la copia de seguridad a restituir:"
            Height          =   615
            Left            =   120
            TabIndex        =   19
            Top             =   2880
            Width           =   7695
            Begin VB.TextBox txtnombreArchivo 
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   8
               ToolTipText     =   "Muestra el nombre de la copia elegida para restituir"
               Top             =   240
               Width           =   7455
            End
         End
         Begin VB.DriveListBox Drive2 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Selecciona la unidad donde se ubica la copia que deseamos restituir"
            Top             =   240
            Width           =   3615
         End
         Begin VB.DirListBox Dir2 
            Height          =   2115
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Selecciona la carpeta donde se ubica la copia que deseamos restituir"
            Top             =   600
            Width           =   3615
         End
      End
      Begin MSComctlLib.ProgressBar PB2 
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   5400
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label3 
         Caption         =   $"FrmBackup.frx":6226
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   7815
      End
   End
   Begin MSComctlLib.TabStrip TabStripOpciones 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   11880
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Realizar copia de seguridad"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Restituir copia de seguridad"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AQUI SE REALIZAN LAS COPIAS DE SEGURIDAD DE LA BASE DE DATOS
'TAMBIEN LA RESTAURACION DE BASES ANTERIORES EN CASO DE FALLAS.

Private Sub Form_Load()
On Error GoTo merror

Call RefrescarOpcionesSistema

'seteos de la solapa backup
Dir1.Path = Drive1

'seteos de la pestaña de restauracion
Dir2.Path = Drive2
File1.Pattern = "Creditos-Backup-*.mdb"
File1.Path = Dir2.Path

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de Backup"
End Sub
Private Sub CmdCerrar_Click()
Unload Me
End Sub
Private Sub cmdCopiar_Click()
If TabStripOpciones.SelectedItem.Index = 1 Then
   CmdCopiar.Enabled = False
   If DatosCopiaOk() Then
      If MsgP("¿Confirma el backup?...verifique que la base de datos no este en uso") Then
         Me.MousePointer = vbHourglass
         Call Copiar
         Me.MousePointer = vbDefault
      End If
   End If
   CmdCopiar.Enabled = True
   PB1.Value = 0
End If
End Sub
Private Sub Copiar()
Dim NombreArchivo As String
Dim Mpath As String
On Error GoTo merror

PB1.Value = 10

Mpath = App.Path

'si no estoy en raiz le agrego la barra
If Mid(Mpath, Len(Mpath)) <> "\" Then
   Mpath = Mpath + "\"
End If

'cierro todo para poder hacer el backup
Call CLOSE_MODULE

PB1.Value = 30
If CheckBorrarAnterior.Value = 1 Then
   Call EliminarBackupAnterior
   PB1.Value = 50
End If

PB1.Value = 70

'hago la copia
'nombre del archivo a backupear
NombreArchivo = Mpath & "database\creditos.mdb"

'si ya existe una copia del dia actual la borro para que no haya errores
If Len(Dir$(TxtArchivoDestino.Text)) Then Kill TxtArchivoDestino.Text

FileCopy NombreArchivo, TxtArchivoDestino.Text

Call AbrirBase

PB1.Value = 100

MsgI "La copia de seguridad se realizo exitosamente"

Exit Sub
merror:
tratarerrores "Error efectuando la copia de seguridad"
End Sub
Private Function DatosCopiaOk() As Boolean
On Error GoTo merror

DatosCopiaOk = True

If Trim(TxtArchivoDestino.Text) = "" Then
   DatosCopiaOk = False
   MsgE "Debe seleccionar la carpeta destino de la copia"
   Exit Function
End If

'verifico que no estoy backupeando a la misma carpeta de la base
If UCase(Dir1.Path) = UCase(App.Path) & "\DATABASE" Then
   DatosCopiaOk = False
   MsgE "No puede hacer el backup a la misma carpeta del sistema de Creditos"
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosCopiaOk"
End Function
Private Sub cmdrestaurar_Click()
If TabStripOpciones.SelectedItem.Index = 2 Then
   CmdRestaurar.Enabled = False
   If DatosRestaurarOk() Then
      If MsgP("¿Confirma la restauracion de la base?...verifique que la base de datos no este en uso") Then
         Me.MousePointer = vbHourglass
         Call Restaurar
         Me.MousePointer = vbDefault
      End If
   End If
   CmdRestaurar.Enabled = True
   PB2.Value = 0
End If
End Sub
Private Sub Restaurar()
Dim Archivorestaurar As String
Dim NombreBaseCopia As String
Dim Mpath As String
On Error GoTo merror

PB2.Value = 30

Mpath = App.Path

'si no estoy en raiz le agrego la barra
If Mid(Mpath, Len(Mpath)) <> "\" Then
   Mpath = Mpath + "\"
End If

'cierro la base para poder restaurar
Call CLOSE_MODULE

'establezco el nombre de la copia de seguridad mia anterior
NombreBaseCopia = Mpath & "database\creditos-anterior.mdb"

'si existe borro la copia de seguridad anterior
If Len(Dir$(NombreBaseCopia)) Then Kill NombreBaseCopia

PB2.Value = 50

'renombro la base actual y la resguardo
Name Mpath & "database\creditos.mdb" As Mpath & "database\creditos-anterior.mdb"

'debe restaurar a la carpeta de la base

PB2.Value = 70

'hago la copia
FileCopy txtnombreArchivo.Text, Mpath & "database\creditos.mdb"
 
Call AbrirBase

PB2.Value = 100

MsgI "La base de datos se restituyo exitosamente"

Exit Sub
merror:
tratarerrores "Error restituyendo la base de datos"
End Sub
Private Function DatosRestaurarOk() As Boolean
On Error GoTo merror

DatosRestaurarOk = True

If Trim(txtnombreArchivo.Text) = "" Then
   DatosRestaurarOk = False
   MsgE "Debe seleccionar la base de datos a restaurar"
   Exit Function
End If

'reviso si ingresaron un nombre de archivo y no una carpeta sola
If InStr(Trim(UCase(txtnombreArchivo.Text)), ".MDB") = 0 Then
   DatosRestaurarOk = False
   MsgE "Debe seleccionar una base de datos valida para restaurar"
   Exit Function
End If

'reviso si la base que esta restaurando es de credimaco
If InStr(Trim(txtnombreArchivo.Text), "Creditos-Backup") = 0 Then
   DatosRestaurarOk = False
   MsgE "La base a restaurar no es una copia de seguridad del software de credimaco"
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosRestaurarOk"
End Function
Private Sub Drive2_Change()
'si cambio el drive actualizo las carpetas,la lista de archivos y el nombre de archivo
On Error GoTo merror

Dir2.Path = Drive2

Exit Sub
merror:
tratarerrores "Error cambiando de unidad"
txtnombreArchivo.Text = ""
Drive2.Drive = "c:"
Dir2.Path = Drive2
Call Dir2_Change
End Sub
Private Sub Dir2_Change()
'si cambio la carpeta actualizo la lista de archivos y el nombre de archivo
On Error GoTo merror

File1.Path = Dir2.Path

If Mid(Dir2.Path, Len(Dir2.Path)) <> "\" Then
   txtnombreArchivo.Text = UCase(Dir2.Path) & "\" & File1
Else
   txtnombreArchivo.Text = UCase(Dir2.Path) & File1
End If

Exit Sub
merror:
tratarerrores "Error cambiando de carpeta"
End Sub
Private Sub File1_Click()
'si cambio el archivo actualizo el nombre de archivo
On Error GoTo merror

If Mid(Dir2.Path, Len(Dir2.Path)) <> "\" Then
   txtnombreArchivo.Text = UCase(Dir2.Path) & "\" & File1
Else
   txtnombreArchivo.Text = UCase(Dir2.Path) & File1
End If

Exit Sub
merror:
tratarerrores "Error seleccionando archivo"
End Sub
'esto es del backup
Private Function ObtenerNombreBackup() As String
'Armo el nombre del archivo nuevo backupeado respetando un formato
'de 2 digitos para el dia y el mes y 4 para el año
Dim Dia As String
Dim Mes As String
Dim Año As String
On Error GoTo merror

Dia = Format(CStr(Day(Date)), "00")
Mes = Format(CStr(Month(Date)), "00")
Año = Format(CStr(Year(Date)), "0000")

ObtenerNombreBackup = "Creditos-Backup-" & CStr(Dia) & "-" & CStr(Mes) & "-" & CStr(Año) & ".mdb"

Exit Function
merror:
tratarerrores "Error en funcion ObtenerNombreBackup"
End Function
Private Sub Drive1_Change()
On Error GoTo merror

Dir1.Path = Drive1

Exit Sub
merror:
tratarerrores "Error cambiando de unidad"
TxtArchivoDestino.Text = ""
Drive1.Drive = "c:"
Dir1.Path = Drive1
Call Dir1_Change
End Sub
Private Sub Dir1_Change()
Dim NombreBackup As String
On Error GoTo merror

NombreBackup = ObtenerNombreBackup()

If Mid(Dir1.Path, Len(Dir1.Path)) <> "\" Then
   TxtArchivoDestino.Text = UCase(Dir1.Path) & "\" & NombreBackup
Else
   TxtArchivoDestino.Text = UCase(Dir1.Path) & NombreBackup
End If

Exit Sub
merror:
tratarerrores "Error cambiando de carpeta"
End Sub
Private Sub EliminarBackupAnterior()
'borra todos los backups anteriores sin importar la fecha
Dim BackupAnterior As String
On Error GoTo merror

'si no es raiz
If Mid(Dir1.Path, Len(Dir1.Path)) <> "\" Then
   BackupAnterior = Dir1.Path + "\Creditos-Backup-*.mdb"
Else
   BackupAnterior = Dir1.Path + "Creditos-Backup-*.mdb"
End If

'si existe el backup anterior lo borro
If Len(Dir$(BackupAnterior)) Then Kill BackupAnterior

Exit Sub
merror:
tratarerrores "Error borrando Backup anterior"
End Sub
Private Sub TabStripOpciones_Click()
'datos empresa
If TabStripOpciones.SelectedItem.Index = 1 Then
   FrameCopiar.Visible = True
   FrameReponer.Visible = False
   CmdRestaurar.Enabled = False
   CmdCopiar.Enabled = True
Else
   FrameCopiar.Visible = False
   FrameReponer.Visible = True
   CmdRestaurar.Enabled = True
   CmdCopiar.Enabled = False
End If
End Sub

