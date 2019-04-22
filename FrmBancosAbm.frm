VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBancosAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Bancos"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   HelpContextID   =   12
   Icon            =   "FrmBancosAbm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdRefrescar 
      Caption         =   "Refrescar"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      ToolTipText     =   "Refresca los datos de la pantalla por si hubo cambios desde otra PC en red"
      Top             =   120
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de bancos:"
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5535
      Begin MSComctlLib.ListView lvbancos 
         Height          =   2295
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Lista de bancos"
         Top             =   240
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   4048
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   8820
         EndProperty
      End
   End
   Begin VB.Frame fmeDatos 
      Caption         =   "Datos"
      ForeColor       =   &H00FF0000&
      Height          =   1050
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   5535
      Begin VB.CheckBox CheckPredeterminado 
         Alignment       =   1  'Right Justify
         Caption         =   "Predeterminado"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   720
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Nombre del banco"
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "Permite agregar un nuevo banco"
      Top             =   960
      Width           =   1185
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      ToolTipText     =   "Permite modificar el banco seleccionado"
      Top             =   1440
      Width           =   1185
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      ToolTipText     =   "Permite borrar el banco seleccionado"
      Top             =   1920
      Width           =   1185
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      ToolTipText     =   "Graba los datos de un banco"
      Top             =   2400
      Width           =   1185
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      ToolTipText     =   "Cierra la pantalla"
      Top             =   3480
      Width           =   1185
   End
End
Attribute VB_Name = "FrmBancosAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'LISTA DE BANCOS QUE SE SELECCIONARAN EN LA PANTALLA DE REGISTRAR CREDITOS
'ASOCIADOS A UN PRESTAMO ENTREGADO EN CHEQUE (NO EN EFECTIVO)

Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de bancos"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
If TipoEdicion = "C" Then
   Unload Me
Else
   TipoEdicion = "C"
   Call SetearEntorno
   Call CargarDatos
   lvbancos.SetFocus
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
Private Sub cmdModificar_Click()
Call RefreshTimer
    
If Not VerificarSeleccionLista(lvbancos) Then Exit Sub
TipoEdicion = "M"
Call SetearEntorno
    
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
Dim sql As String
On Error GoTo merror

If Not VerificarSeleccionLista(lvbancos) Then Exit Sub

If Not MsgP("¿Desea borrar el banco seleccionado?") Then Exit Sub

'otras validaciones
If Not ExisteBanco(lvbancos.SelectedItem) Then
   MsgE "El banco no existe"
   Exit Sub
End If

'inicio de transaccion
cnSQL.BeginTrans

sql = "delete from bancos WHERE idbanco=" & CLng(lvbancos.SelectedItem)
cnSQL.Execute sql

'fin de la transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos

TipoEdicion = "C"
Call SetearEntorno

MsgI "El banco fue borrado"

lvbancos.SetFocus

Exit Sub
merror:
tratarerrores "Error borrando banco"
End Sub
Private Sub cmdGrabar_Click()
Call RefreshTimer
If datosok() Then
   Call Grabar
End If
End Sub
Private Sub Grabar()
'graba los registros nuevos y las modificaciones
Dim sql As String
Dim rec As rdoResultset
Dim Mensaje As String
Dim IdBanco As Long
On Error GoTo merror

If TipoEdicion = "N" Then
   If Not MsgP("¿Confirma el nuevo banco?") Then Exit Sub
   
   IdBanco = UltimoId("idbanco", "bancos") + 1
   
   'otras validaciones
   If ExisteBanco(IdBanco) Then
      MsgE "El banco ya existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans
   
   If CheckPredeterminado.Value = 1 Then
      'le saco el predeterminado a las demas
      sql = "update bancos set predeterminada=0"
      cnSQL.Execute sql
   End If
   
   sql = "INSERT INTO bancos (idbanco,nombre,predeterminada) " & _
         "VALUES (" & CLng(IdBanco) & ",'" & CStr(TxtDescripcion.Text) & "'," & CheckPredeterminado.Value & ")"
               
   cnSQL.Execute sql
   
   'fin de la transaccion
   cnSQL.CommitTrans
  
   Mensaje = "El banco fue agregado"
   
   'solo actualizo la pantalla si agrego
   Call CargarLista
   Call CargarDatos
Else
   If Not MsgP("¿Confirma la modificacion del banco seleccionado?") Then Exit Sub
   
   'otras validaciones
   If Not ExisteBanco(lvbancos.SelectedItem) Then
      MsgE "El banco no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans
   
   If CheckPredeterminado.Value = 1 Then
      'le saco el predeterminado a las demas
      sql = "update bancos set predeterminada=0"
      cnSQL.Execute sql
   End If
   
   'estoy grabando una modificacion
   sql = "UPDATE bancos SET nombre='" & CStr(TxtDescripcion.Text) & _
         "',predeterminada=" & CheckPredeterminado.Value & _
         " WHERE Idbanco=" & CLng(lvbancos.SelectedItem)
        
   cnSQL.Execute sql

   'fin de transaccion
   cnSQL.CommitTrans
   
   'actualizo solo la fila de la lista
   lvbancos.SelectedItem.ListSubItems(1).Text = TxtDescripcion.Text & vbNullString
   Mensaje = "El banco fue modificado"
End If

TipoEdicion = "C"
Call SetearEntorno

MsgI (Mensaje)

lvbancos.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando bancos"
End Sub
Private Sub CargarLista()
'carga el listview
Dim rec As rdoResultset
Dim sql As String
Dim Nitem As ListItem
On Error GoTo merror

sql = "SELECT bancos.idbanco,bancos.nombre " & _
      "FROM bancos " & _
      "ORDER BY bancos.nombre"

Set rec = cnSQL.OpenResultset(sql)

lvbancos.ListItems.Clear

If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lvbancos.ListItems.Add(, , Format(rec.rdoColumns("Idbanco"), "000"))
      Nitem.SubItems(1) = rec.rdoColumns("Nombre") & vbNullString
      rec.MoveNext
   Loop
End If
          
Exit Sub
merror:
tratarerrores "Error cargando la lista de bancos"
End Sub
Private Sub CargarDatos()
'carga el elemento seleccionado del listview en los cuadros de texto
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror
    
If Not VerificarSeleccionLista(lvbancos) Then Exit Sub
        
sql = "SELECT * FROM bancos " & _
      "WHERE Idbanco=" & CLng(lvbancos.SelectedItem)
        
Set rec = cnSQL.OpenResultset(sql)
        
If Not rec.EOF Then
   TxtDescripcion.Text = rec.rdoColumns("nombre") & vbNullString
   If rec.rdoColumns("predeterminada") Then
      CheckPredeterminado.Value = 1
   Else
      CheckPredeterminado.Value = 0
   End If
End If
   
Exit Sub
merror:
tratarerrores "Error cargando datos de bancos"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(TxtDescripcion.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el nombre del banco"
   TxtDescripcion.SetFocus
   Exit Function
End If

'reemplazo las comillas no permitidas en todos los campos de texto
Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk_BancosAbm"
End Function
Private Sub SetearEntorno()
'habilita o deshabilita los botones correspondientes
On Error GoTo merror
    
Select Case TipoEdicion
       Case "C"
            fmeDatos.Enabled = False
            cmdGrabar.Enabled = False
            CmdNuevo.Enabled = True
            cmdcerrar.Caption = "&Cerrar"
            lvbancos.Enabled = True
            CmdRefrescar.Enabled = True
            If lvbancos.ListItems.Count() > 0 Then
               cmdModificar.Enabled = True
               CmdBorrar.Enabled = True
            Else
               cmdModificar.Enabled = False
               CmdBorrar.Enabled = False
               Call LimpiarCampos(Me)
            End If
            Call ColorCyan(Me)
       Case "M"
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            CmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            CmdBorrar.Enabled = False
            CmdRefrescar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lvbancos.Enabled = False
            TxtDescripcion.SetFocus
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
            lvbancos.Enabled = False
            TxtDescripcion.SetFocus
            Call ColorBlanco(Me)
End Select

Exit Sub
merror:
tratarerrores "Error seteando entorno-BancosAbm"
End Sub
Private Sub lvbancos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'al clickear sobre el encabezado del listview los datos se ordenan en asc y en desc
Dim Orden As Integer
  
If lvbancos.ListItems.Count > 1 Then
   lvbancos.SortKey = ColumnHeader.Index - 1
   Orden = lvbancos.SortKey
   lvbancos.SortOrder = Abs(Not lvbancos.SortOrder = 1)
   lvbancos.Sorted = True
End If

End Sub
Private Sub lvbancos_ItemClick(ByVal Item As MSComctlLib.ListItem)
'al clickear sobre algun elemento del listview los datos se cargan en los dos text
Call CargarDatos
End Sub
Private Sub TxtDescripcion_LostFocus()
  TxtDescripcion.Text = UCase(Trim(TxtDescripcion.Text))
End Sub
