VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFeriadosAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Feriados (dias no habiles para vencimientos)"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   HelpContextID   =   15
   Icon            =   "FrmFeriadosAbm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdRefrescar 
      Caption         =   "Refrescar"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      ToolTipText     =   "Refresca los datos de la pantalla por si hubo cambios desde otra PC en red"
      Top             =   120
      Width           =   1305
   End
   Begin VB.Frame FrameFeriados 
      Caption         =   "Lista de feriados registrados:"
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   5655
      Begin MSComctlLib.ListView lv 
         Height          =   2655
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Muestra la lista de feriados"
         Top             =   240
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   4683
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
            Text            =   "Fecha"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame fmeDatos 
      Caption         =   "Datos"
      ForeColor       =   &H00FF0000&
      Height          =   1065
      Left            =   15
      TabIndex        =   10
      Top             =   3240
      Width           =   5640
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "Fecha del feriado"
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54132737
         CurrentDate     =   38814
      End
      Begin VB.TextBox txtDatos 
         Height          =   285
         Index           =   1
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Descripcion del feriado"
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "Permite agregar un nuevo feriado"
      Top             =   960
      Width           =   1305
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      ToolTipText     =   "Permite modificar los datos del feriado seleccionado"
      Top             =   1560
      Width           =   1305
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      ToolTipText     =   "Borra el feriado seleccionado"
      Top             =   2160
      Width           =   1305
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      ToolTipText     =   "Graba los datos del feriado"
      Top             =   2760
      Width           =   1305
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      ToolTipText     =   "Cierra la pantalla"
      Top             =   3960
      Width           =   1305
   End
End
Attribute VB_Name = "FrmFeriadosAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***AQUI SE REGISTRAN LOS FERIADOS DEL SISTEMA PARA QUE LUEGO NO SE GENEREN
'FECHAS DE VENCIMIENTO DE CUOTAS EN ESOS DIAS

Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call RefrescarOpcionesSistema

Call CargarLista
Call CargarDatos

TipoEdicion = "C"
Call SetearEntorno

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de Feriados"
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
Private Sub CmdRefrescar_Click()
'refresca la pantalla
Call RefreshTimer
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
txtDatos(1).SetFocus
    
End Sub
Private Sub cmdnuevo_Click()
Call RefreshTimer
TipoEdicion = "N"
Call SetearEntorno
End Sub
Private Sub cmdborrar_Click()
Call RefreshTimer
Call Borrar
End Sub
Private Sub Borrar()
'solicita la confirmacion del borrado
Dim sql As String
On Error GoTo merror
   
If Not VerificarSeleccionLista(lv) Then Exit Sub
    
If Not MsgP("¿Confirma el borrado del feriado seleccionado?") Then Exit Sub

'inicio de transaccion
cnSQL.BeginTrans

sql = "delete from feriados WHERE fecha='" & ConvertirFechaSql(lv.SelectedItem, "DD/MM/YYYY") & "'"

cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

MsgI "El Feriado fue borrado"

lv.SetFocus

Exit Sub
merror:
tratarerrores "Error borrando feriados"
End Sub
Private Function ExisteFeriado(ByVal Fecha As Date) As Boolean
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ExisteFeriado = False

sql = "select fecha from feriados where fecha='" & ConvertirFechaSql(Fecha, "DD/MM/YYYY") & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("fecha")) Then
      ExisteFeriado = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ExisteFeriado"
End Function
Private Sub cmdGrabar_Click()
Call RefreshTimer
Call Grabar
End Sub
Private Sub Grabar()
'graba los registros nuevos y las modificaciones
Dim sql As String
Dim Mensaje As String
On Error GoTo merror

If Not datosok() Then Exit Sub
  
If TipoEdicion = "N" Then
   If ExisteFeriado(DTPicker1.Value) Then
      MsgE "El feriado ya esta registrado con anterioridad"
      Exit Sub
   End If

   If Not MsgP("¿Confirma el nuevo feriado?") Then Exit Sub
   
   'otras validaciones
   If ExisteFeriado(DTPicker1.Value) Then
      MsgE "El feriado ya esta registrado con anterioridad"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans

   sql = "INSERT INTO feriados(fecha,descripcion) " & _
         "VALUES ('" & ConvertirFechaSql(DTPicker1.Value, "DD/MM/YYYY") & "','" & CStr(txtDatos(1).Text) & "')"
        
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El feriado fue agregado"
   
   Call CargarLista
   Call CargarDatos

Else
   If Not MsgP("¿Confirma la modificacion del feriado seleccionado?") Then Exit Sub
   
   'otras validaciones
   If Not ExisteFeriado(DTPicker1.Value) Then
      MsgE "El feriado no existe"
      Exit Sub
   End If

   'inicio de transaccion
   cnSQL.BeginTrans

   'estoy grabando una modificacion
   sql = "UPDATE feriados SET descripcion='" & CStr(txtDatos(1).Text) & "' " & _
         "WHERE fecha='" & ConvertirFechaSql(lv.SelectedItem, "DD/MM/YYYY") & "'"
        
   cnSQL.Execute sql

   'fin de transaccion
   cnSQL.CommitTrans
        
   Mensaje = "El feriado fue modificado"
   
   lv.SelectedItem.ListSubItems(1).Text = txtDatos(1).Text & vbNullString

End If

TipoEdicion = "C"
Call SetearEntorno

MsgI Mensaje

lv.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando feriados"
End Sub
Private Sub CargarLista()
'carga el listview con los feriados del presente año
Dim sql As String
Dim rec As rdoResultset
Dim Nitem As ListItem
On Error GoTo merror
    
sql = "SELECT * FROM feriados " & _
      "where year(fecha)=year('" & ConvertirFechaSql(Mid$(Now, 1, 10), "DD/MM/YYYY") & "') " & _
      "ORDER BY fecha"

Set rec = cnSQL.OpenResultset(sql)
    
lv.ListItems.Clear

If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lv.ListItems.Add(, , rec.rdoColumns("fecha"))
      Nitem.SubItems(1) = rec.rdoColumns("descripcion") & vbNullString
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error cargando la lista de feriados"
End Sub
Private Sub CargarDatos()
'carga el elemento seleccionado del listview en los cuadros de texto
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror
    
If Not VerificarSeleccionLista(lv) Then Exit Sub
        
sql = "SELECT fecha,descripcion FROM feriados " & _
      "WHERE fecha='" & ConvertirFechaSql(lv.SelectedItem, "DD/MM/YYYY") & "'"

Set rec = cnSQL.OpenResultset(sql)
        
If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("fecha")) Then
      DTPicker1.Value = CDate(rec.rdoColumns("fecha"))
      txtDatos(1).Text = rec.rdoColumns("descripcion") & vbNullString
   End If
End If

Exit Sub
merror:
tratarerrores "Error cargando datos de feriados"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If CDate(DTPicker1.Value) < CDate(Date) Then
   datosok = False
   MsgE "Verifique la fecha del feriado...debe ser mayor a la actual"
   DTPicker1.SetFocus
   Exit Function
End If

If Trim(txtDatos(1).Text) = "" Then
   datosok = False
   MsgE "Debe ingresar la descripcion del feriado"
   txtDatos(1).SetFocus
   Exit Function
End If

Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-Feriados"
End Function
Private Sub SetearEntorno()
'habilita o desahbilita los botones correspondientes
On Error GoTo merror
    
Select Case TipoEdicion
       Case "C"
            fmeDatos.Enabled = False
            cmdGrabar.Enabled = False
            CmdRefrescar.Enabled = True
            If lv.ListItems.Count > 0 Then
               cmdModificar.Enabled = True
               CmdBorrar.Enabled = True
            Else
               cmdModificar.Enabled = False
               CmdBorrar.Enabled = False
               Call LimpiarCampos(Me)
            End If
            CmdNuevo.Enabled = True
            cmdcerrar.Caption = "&Cerrar"
            lv.Enabled = True
            Call ColorCyan(Me)
       Case "M"
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            CmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            CmdBorrar.Enabled = False
            CmdRefrescar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lv.Enabled = False
            Call ColorBlanco(Me)
        Case "N"
            Call LimpiarCampos(Me)
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            CmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            CmdBorrar.Enabled = False
            CmdRefrescar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lv.Enabled = False
            DTPicker1.Value = Date
            DTPicker1.SetFocus
            Call ColorBlanco(Me)
End Select

Exit Sub
merror:
tratarerrores "Error seteando entorno-Feriados"
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
Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
'al clickear sobre algun elemento del listview los datos se cargan en los dos text
Call CargarDatos
End Sub
Private Sub TxtDatos_LostFocus(Index As Integer)
'al perder el enfoque los text transforman el contenido a mayusculas
txtDatos(Index).Text = UCase(Trim(txtDatos(Index).Text))
End Sub

