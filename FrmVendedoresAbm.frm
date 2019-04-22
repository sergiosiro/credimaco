VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVendedoresAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Vendedores"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "FrmVendedoresAbm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdRefrescar 
      Caption         =   "Refrescar"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      ToolTipText     =   "Refresca los datos de la pantalla por si hubo cambios desde otra PC en red"
      Top             =   120
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de vendedores:"
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5415
      Begin MSComctlLib.ListView lvvendedores 
         Height          =   2415
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Lista de vendedores"
         Top             =   240
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   4260
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fmeDatos 
      Caption         =   "Datos"
      ForeColor       =   &H00FF0000&
      Height          =   1305
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   5415
      Begin VB.CheckBox CheckPredeterminada 
         Caption         =   "Predeterminado"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Si marca la casilla, el vendedor seleccionado sera la predeterminada en las demas pantallas del sistema"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Nombre del vendedor"
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Nombre del vendedor:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      ToolTipText     =   "Permite agregar los datos de un nuevo vendedor"
      Top             =   960
      Width           =   1185
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      ToolTipText     =   "Permite modificar los datos del vendedor seleccionado en la lista"
      Top             =   1560
      Width           =   1185
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      ToolTipText     =   "Permite borrar el vendedor seleccionado en la lista"
      Top             =   2160
      Width           =   1185
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      ToolTipText     =   "Graba los datos de un vendedor"
      Top             =   2760
      Width           =   1185
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      ToolTipText     =   "Cierra la pantalla o cancela una operacion de agregado o modificacion"
      Top             =   3960
      Width           =   1185
   End
End
Attribute VB_Name = "FrmVendedoresAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***AQUI SE REGISTRAN LOS VENDEDORES DE CREDITOS
'ESTO SE AGREGO EN 2011 Y SOLO SIRVE PARA
'SELECCIONAR UNA LISTA DE VENDEDORES EN LA PANTALLA
'DE REGISTRAR CREDITOS. POR ELLO LA TABLA VENDEDORES NO APARECE EN LAS CONSULTAS SQL

Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de Vendedores"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
If TipoEdicion = "C" Then
   Unload Me
Else
   TipoEdicion = "C"
   Call SetearEntorno
   Call CargarDatos
   lvvendedores.SetFocus
End If
End Sub
Private Sub CmdRefrescar_Click()
'refresca la pantalla por si hubo cambios desde otra maquina
Call RefreshTimer
Call RefrescarOpcionesSistema
Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno
End Sub
Private Sub cmdModificar_Click()
Call RefreshTimer
    
If Not VerificarSeleccionLista(lvvendedores) Then Exit Sub
TipoEdicion = "M"
Call SetearEntorno
TxtDescripcion.SetFocus
    
End Sub
Private Sub cmdnuevo_Click()
Call RefreshTimer
TipoEdicion = "N"
Call SetearEntorno
End Sub
Private Sub cmdborrar_Click()
Dim sql As String
On Error GoTo merror
Call RefreshTimer

If Not VerificarSeleccionLista(lvvendedores) Then Exit Sub
    
'***no se chequea la integridad referencial porque los
'vendedores solo se usan como campo de texto asociado a creditos

If Not MsgP("¿Confirma el borrado del vendedor seleccionado?") Then Exit Sub

'otras validaciones
If Not ExisteVendedor(lvvendedores.SelectedItem) Then
   MsgE "El vendedor no existe"
   Exit Sub
End If

'inicio de transaccion
cnSQL.BeginTrans

sql = "delete from vendedores WHERE idvendedor=" & CLng(lvvendedores.SelectedItem)

cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

MsgI "El vendedor fue borrado"
lvvendedores.SetFocus

Exit Sub
merror:
tratarerrores "Error borrando vendedores"
End Sub
Private Sub cmdGrabar_Click()
'graba los registros nuevos y las modificaciones
Dim sql As String
Dim IdVendedor As Long
Dim Mensaje As String
On Error GoTo merror
Call RefreshTimer

If Not datosok() Then Exit Sub
    
If TipoEdicion = "N" Then
   If Not MsgP("¿Confirma el nuevo vendedor?") Then Exit Sub
   
   IdVendedor = UltimoId("idvendedor", "vendedores") + 1
   
   'otras validaciones
   If ExisteVendedor(IdVendedor) Then
      MsgE "El vendedor ya existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans

   'si es predeterminada saco el predeterminada al resto
   If CheckPredeterminada.Value = 1 Then
      sql = "update vendedores set predeterminada=0"
      cnSQL.Execute sql
   End If
           
   sql = "INSERT INTO vendedores (Idvendedor, nombre,predeterminada) " & _
         "VALUES (" & CLng(IdVendedor) & ",'" & CStr(TxtDescripcion.Text) & "'," & CheckPredeterminada.Value & ")"
        
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El nuevo vendedor fue agregado"
   
   Call CargarLista
   Call CargarDatos
Else
   If Not MsgP("¿Confirma la modificacion del vendedor seleccionado?") Then Exit Sub
   
   'otras validaciones
   If Not ExisteVendedor(lvvendedores.SelectedItem) Then
      MsgE "El vendedor no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans

   If CheckPredeterminada.Value = 1 Then
      sql = "update vendedores set predeterminada=0"
      cnSQL.Execute sql
   End If
   
   'estoy grabando una modificacion
   sql = "UPDATE vendedores SET nombre='" & CStr(TxtDescripcion.Text) & _
         "',predeterminada=" & CheckPredeterminada.Value & _
         " WHERE Idvendedor=" & CLng(lvvendedores.SelectedItem)
        
   cnSQL.Execute sql

   'fin de transaccion
   cnSQL.CommitTrans
        
   Mensaje = "El vendedor fue modificado"
   
   lvvendedores.SelectedItem.ListSubItems(1).Text = TxtDescripcion.Text & vbNullString
End If

TipoEdicion = "C"
Call SetearEntorno

MsgI Mensaje

lvvendedores.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando vendedores"
End Sub
Private Sub CargarLista()
'carga el listview con los vendedores
Dim sql As String
Dim rec As rdoResultset
Dim Nitem As ListItem
On Error GoTo merror
    
lvvendedores.ListItems.Clear

sql = "SELECT * FROM vendedores " & _
      "ORDER BY vendedores.nombre"

Set rec = cnSQL.OpenResultset(sql)
    
If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lvvendedores.ListItems.Add(, , Format(rec.rdoColumns("Idvendedor"), "000"))
      Nitem.SubItems(1) = rec.rdoColumns("Nombre") & vbNullString
      
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error cargando la lista de vendedores"
End Sub
Private Sub CargarDatos()
'carga el elemento seleccionado del listview en los cuadros de texto
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror
    
If Not VerificarSeleccionLista(lvvendedores) Then Exit Sub
        
sql = "SELECT * FROM vendedores WHERE Idvendedor=" & CLng(lvvendedores.SelectedItem)
Set rec = cnSQL.OpenResultset(sql)
        
If Not rec.EOF Then
   TxtDescripcion.Text = rec.rdoColumns("nombre") & vbNullString
   
   If rec.rdoColumns("predeterminada") Then
      CheckPredeterminada.Value = 1
   Else
      CheckPredeterminada.Value = 0
   End If
End If

Exit Sub
merror:
tratarerrores "Error cargando datos de vendedores"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(TxtDescripcion.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el nombre del vendedor"
   TxtDescripcion.SetFocus
   Exit Function
End If

Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-VendedoresAbm"
End Function
Private Sub SetearEntorno()
'habilita o desahbilita los botones correspondientes
On Error GoTo merror

Select Case TipoEdicion
       Case "C"
            fmeDatos.Enabled = False
            cmdGrabar.Enabled = False
            CmdNuevo.Enabled = True
            CmdRefrescar.Enabled = True
            If lvvendedores.ListItems.Count > 0 Then
               CmdBorrar.Enabled = True
               cmdModificar.Enabled = True
            Else
               cmdModificar.Enabled = False
               CmdBorrar.Enabled = False
               Call LimpiarCampos(Me)
            End If
            cmdcerrar.Caption = "&Cerrar"
            lvvendedores.Enabled = True
            Call ColorCyan(Me)
        Case "M"
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            CmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            CmdBorrar.Enabled = False
            CmdRefrescar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lvvendedores.Enabled = False
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
            lvvendedores.Enabled = False
            TxtDescripcion.SetFocus
            Call ColorBlanco(Me)
    End Select

Exit Sub
merror:
tratarerrores "Error seteando entorno-VendedoresAbm"
End Sub
Private Sub lvvendedores_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'al clickear sobre el encabezado del listview los datos se ordenan
Dim Orden As Integer
  
If lvvendedores.ListItems.Count > 1 Then
   lvvendedores.SortKey = ColumnHeader.Index - 1
   Orden = lvvendedores.SortKey
   lvvendedores.SortOrder = Abs(Not lvvendedores.SortOrder = 1)
   lvvendedores.Sorted = True
End If

End Sub
Private Sub lvvendedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
'al clickear sobre algun elemento del listview los datos se cargan en los text
Call CargarDatos
End Sub
Private Sub TxtDescripcion_LostFocus()
TxtDescripcion.Text = UCase(Trim(TxtDescripcion.Text))
End Sub
