VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProvinciasAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Provincias"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   HelpContextID   =   14
   Icon            =   "FrmProvinciasAbm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdRefrescar 
      Caption         =   "Refrescar"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      ToolTipText     =   "Refresca los datos de la pantalla por si hubo cambios desde otra PC en red"
      Top             =   120
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de provincias:"
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5415
      Begin MSComctlLib.ListView lvprovincias 
         Height          =   2415
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Lista de provincias"
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
            Text            =   "Porcentaje.Sellado"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fmeDatos 
      Caption         =   "Datos"
      ForeColor       =   &H00FF0000&
      Height          =   1665
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   5415
      Begin VB.TextBox TxtPorcentajeSellados 
         Height          =   285
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox CheckPredeterminada 
         Caption         =   "Predeterminada"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Si marca la casilla, la provincia seleccionada sera la predeterminada en las demas pantallas del sistema"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Nombre de la provincia"
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje sellados:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Provincia/estado:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      ToolTipText     =   "Permite agregar los datos de una nueva provincia"
      Top             =   960
      Width           =   1185
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      ToolTipText     =   "Permite modificar los datos de la provincia seleccionada"
      Top             =   1560
      Width           =   1185
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      ToolTipText     =   "Permite borrar la provincia seleccionada"
      Top             =   2160
      Width           =   1185
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      ToolTipText     =   "Graba los datos de una provincia"
      Top             =   2760
      Width           =   1185
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      ToolTipText     =   "Cierra la pantalla o cancela una operacion de agregado o modificacion"
      Top             =   4320
      Width           =   1185
   End
End
Attribute VB_Name = "FrmProvinciasAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***AQUI SE AGREGAN LAS PROVINCIAS DEL SISTEMA.LUEGO SE LE ASOCIARAN CIUDADES
'EN LA PANTALLA DE CIUDADES
Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de Provincias"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
If TipoEdicion = "C" Then
   Unload Me
Else
   TipoEdicion = "C"
   Call SetearEntorno
   Call CargarDatos
   lvprovincias.SetFocus
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
If Not VerificarSeleccionLista(lvprovincias) Then Exit Sub
TipoEdicion = "M"
Call SetearEntorno
TxtDescripcion.SetFocus
    
End Sub
Private Sub cmdnuevo_Click()
Call RefreshTimer
TipoEdicion = "N"
Call SetearEntorno
End Sub
Private Function PuedoBorrarProvincia(ByVal IdProvincia As Long) As Boolean
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

PuedoBorrarProvincia = True

'verifico en tabla localidades
sql = "select idprovincia " & _
      "from localidades " & _
      "where idprovincia=" & CLng(IdProvincia)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idprovincia")) Then
      PuedoBorrarProvincia = False
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion PuedoBorrarProvincia"
End Function
Private Sub cmdborrar_Click()
Dim sql As String
On Error GoTo merror
Call RefreshTimer

If Not VerificarSeleccionLista(lvprovincias) Then Exit Sub
    
If Not PuedoBorrarProvincia(lvprovincias.SelectedItem) Then
   MsgE "No se puede borrar la provincia,...tiene localidades asociadas"
   Exit Sub
End If

If Not MsgP("¿Confirma el borrado de la provincia seleccionada?") Then Exit Sub

'otras validaciones
'esta validacion  va por la pausa que se produce en la pregunta
'anterior. En ese lapso pueden haber borrado la provincia
'desde otra maquina de la red.
If Not ExisteProvincia(lvprovincias.SelectedItem) Then
   MsgE "La provincia no existe"
   Exit Sub
End If

'inicio de transaccion
cnSQL.BeginTrans

sql = "delete from provincias WHERE idprovincia=" & CLng(lvprovincias.SelectedItem)

cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

MsgI "La provincia fue borrada"
lvprovincias.SetFocus

Exit Sub
merror:
tratarerrores "Error borrando provincias"
End Sub
Private Sub cmdGrabar_Click()
'graba los registros nuevos y las modificaciones
Dim sql As String
Dim IdProvincia As Long
Dim Mensaje As String
On Error GoTo merror
Call RefreshTimer

If Not datosok() Then Exit Sub
    
If TipoEdicion = "N" Then
   If Not MsgP("¿Confirma la nueva provincia?") Then Exit Sub
   
   IdProvincia = UltimoId("idprovincia", "provincias") + 1
   
   'otras validaciones
   If ExisteProvincia(IdProvincia) Then
      MsgE "La provincia ya existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans

   'si es predeterminada saco el predeterminada al resto
   If CheckPredeterminada.Value = 1 Then
      sql = "update provincias set predeterminada=0"
      cnSQL.Execute sql
   End If
           
   sql = "INSERT INTO provincias (Idprovincia, nombre,predeterminada,porcentajesellados) " & _
         "VALUES (" & CLng(IdProvincia) & ",'" & CStr(TxtDescripcion.Text) & "'," & CheckPredeterminada.Value & "," & ConvertirDblSql(CDbl(TxtPorcentajeSellados.Text)) & ")"
        
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "La nueva provincia fue agregada"
   
   Call CargarLista
   Call CargarDatos
Else
   If Not MsgP("¿Confirma la modificacion de la provincia seleccionada?") Then Exit Sub
   
   'otras validaciones
   If Not ExisteProvincia(lvprovincias.SelectedItem) Then
      MsgE "La provincia no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans

   If CheckPredeterminada.Value = 1 Then
      sql = "update provincias set predeterminada=0"
      cnSQL.Execute sql
   End If
   
   'estoy grabando una modificacion
   sql = "UPDATE provincias SET nombre='" & CStr(TxtDescripcion.Text) & _
         "',predeterminada=" & CheckPredeterminada.Value & _
         ",porcentajesellados=" & ConvertirDblSql(CDbl(TxtPorcentajeSellados.Text)) & _
         "WHERE Idprovincia=" & CLng(lvprovincias.SelectedItem)
        
   cnSQL.Execute sql

   'fin de transaccion
   cnSQL.CommitTrans
        
   Mensaje = "La provincia fue modificada"
   
   lvprovincias.SelectedItem.ListSubItems(1).Text = TxtDescripcion.Text & vbNullString
   lvprovincias.SelectedItem.ListSubItems(2).Text = Format(TxtPorcentajeSellados.Text, "0.00") & vbNullString

End If

TipoEdicion = "C"
Call SetearEntorno

MsgI Mensaje

lvprovincias.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando provincias"
End Sub
Private Sub CargarLista()
'carga el listview con las provincias
Dim sql As String
Dim rec As rdoResultset
Dim Nitem As ListItem
On Error GoTo merror
    
lvprovincias.ListItems.Clear

sql = "SELECT * FROM provincias " & _
      "ORDER BY provincias.nombre"

Set rec = cnSQL.OpenResultset(sql)
    
If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lvprovincias.ListItems.Add(, , Format(rec.rdoColumns("Idprovincia"), "000"))
      Nitem.SubItems(1) = rec.rdoColumns("Nombre") & vbNullString
      Nitem.SubItems(2) = Format(rec.rdoColumns("porcentajesellados"), "0.00") & vbNullString
      
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error cargando la lista de provincias"
End Sub
Private Sub CargarDatos()
'carga el elemento seleccionado del listview en los cuadros de texto
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror
    
If Not VerificarSeleccionLista(lvprovincias) Then Exit Sub
        
sql = "SELECT * FROM provincias WHERE Idprovincia=" & CLng(lvprovincias.SelectedItem)
Set rec = cnSQL.OpenResultset(sql)
        
If Not rec.EOF Then
   TxtDescripcion.Text = rec.rdoColumns("nombre") & vbNullString
   TxtPorcentajeSellados.Text = Format(rec.rdoColumns("porcentajesellados"), "0.00") & vbNullString
   
   If rec.rdoColumns("predeterminada") Then
      CheckPredeterminada.Value = 1
   Else
      CheckPredeterminada.Value = 0
   End If
End If

Exit Sub
merror:
tratarerrores "Error cargando datos de provincias"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(TxtDescripcion.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el nombre de la provincia"
   TxtDescripcion.SetFocus
   Exit Function
End If

'valido el porcentaje
If Trim(TxtPorcentajeSellados.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el porcentaje de sellados"
   TxtPorcentajeSellados.SetFocus
   Exit Function
End If
If Not IsNumeric(TxtPorcentajeSellados.Text) Then
   datosok = False
   MsgE "El porcentaje de sellados debe ser numerico"
   TxtPorcentajeSellados.SetFocus
   Exit Function
End If
If CCur(TxtPorcentajeSellados.Text) < 0 Then
   datosok = False
   MsgE "El porcentaje de sellados debe ser mayor o igual a cero"
   TxtPorcentajeSellados.SetFocus
   Exit Function
End If


'reemplazo los caracteres invalidos del teclado como el simbolo apostrofe
Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-ProvinciasAbm"
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
            If lvprovincias.ListItems.Count > 0 Then
               CmdBorrar.Enabled = True
               cmdModificar.Enabled = True
            Else
               cmdModificar.Enabled = False
               CmdBorrar.Enabled = False
               Call LimpiarCampos(Me)
            End If
            cmdcerrar.Caption = "&Cerrar"
            lvprovincias.Enabled = True
            Call ColorCyan(Me)
        Case "M"
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            CmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            CmdBorrar.Enabled = False
            CmdRefrescar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lvprovincias.Enabled = False
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
            lvprovincias.Enabled = False
            TxtDescripcion.SetFocus
            Call ColorBlanco(Me)
    End Select

Exit Sub
merror:
tratarerrores "Error seteando entorno-ProvinciasAbm"
End Sub
Private Sub lvprovincias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'al clickear sobre el encabezado del listview los datos se ordenan en asc y en desc
Dim Orden As Integer
  
If lvprovincias.ListItems.Count > 1 Then
   lvprovincias.SortKey = ColumnHeader.Index - 1
   Orden = lvprovincias.SortKey
   lvprovincias.SortOrder = Abs(Not lvprovincias.SortOrder = 1)
   lvprovincias.Sorted = True
End If

End Sub
Private Sub lvprovincias_ItemClick(ByVal Item As MSComctlLib.ListItem)
'al clickear sobre algun elemento del listview los datos se cargan en los text
Call CargarDatos
End Sub
Private Sub TxtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   TxtPorcentajeSellados.SetFocus
End If
End Sub
Private Sub TxtDescripcion_LostFocus()
TxtDescripcion.Text = UCase(Trim(TxtDescripcion.Text))
End Sub
