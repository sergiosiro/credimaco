VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLocalidadesAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Localidades"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   HelpContextID   =   13
   Icon            =   "FrmLocalidadesAbm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdRefrescar 
      Caption         =   "Refrescar"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      ToolTipText     =   "Refresca los datos de la pantalla por si hubo cambios desde otra PC en red"
      Top             =   120
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de localidades:"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5175
      Begin MSComctlLib.ListView lvlocalidad 
         Height          =   1935
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Lista de localidades"
         Top             =   240
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   3413
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
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      ToolTipText     =   "Cierra la pantalla o cancela una operacion"
      Top             =   3720
      Width           =   1185
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      ToolTipText     =   "Graba los datos de una localidad"
      Top             =   2760
      Width           =   1185
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      ToolTipText     =   "Permite borrar la localidad seleccionada"
      Top             =   2160
      Width           =   1185
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      ToolTipText     =   "Permite modificar los datos de la localidad seleccionada"
      Top             =   1560
      Width           =   1185
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Permite agregar los datos de una nueva localidad"
      Top             =   960
      Width           =   1185
   End
   Begin VB.Frame fmeDatos 
      Caption         =   "Datos"
      ForeColor       =   &H00FF0000&
      Height          =   1545
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   5175
      Begin VB.CheckBox CheckPredeterminada 
         Alignment       =   1  'Right Justify
         Caption         =   "Predeterminada:"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         ToolTipText     =   "Si marca la casilla, la localidad sera la predeterminada en las demas pantallas del sistema"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtDatos 
         Height          =   285
         Index           =   2
         Left            =   960
         MaxLength       =   12
         TabIndex        =   2
         ToolTipText     =   "Codigo postal de la localidad"
         Top             =   720
         Width           =   1590
      End
      Begin VB.ComboBox comboprovincias 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Provincia a la cual pertenece la localidad"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txtDatos 
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Nombre de la localidad"
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.Postal:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Provincia:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmLocalidadesAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AQUI SE AGREGAN LAS CIUDADES DEL SISTEMA QUE LUEGO SE ASOCIARAN A LOS
'CLIENTES DEL SISTEMA

Private Sub Form_Load()
Call RefreshTimer
On Error GoTo merror

Call CargarComboProvincias("provincias", ComboProvincias)

Call CargarLista
Call CargarDatos

TipoEdicion = "C"
Call SetearEntorno

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de localidades"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
If TipoEdicion = "C" Then
   Unload Me
Else
   TipoEdicion = "C"
   Call SetearEntorno
   Call CargarDatos
   lvlocalidad.SetFocus
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
Private Function PuedoBorrarLocalidad(ByVal IdLocalidad As Long) As Boolean
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

PuedoBorrarLocalidad = True

'verifico en tabla clientes
sql = "select idlocalidad " & _
      "from clientes " & _
      "where idlocalidad=" & CLng(IdLocalidad)
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idlocalidad")) Then
      PuedoBorrarLocalidad = False
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion PuedoBorrarLocalidad"
End Function
Private Sub cmdborrar_Click()
Dim sql As String
On Error GoTo merror
Call RefreshTimer

If Not VerificarSeleccionLista(lvlocalidad) Then Exit Sub

If Not PuedoBorrarLocalidad(lvlocalidad.SelectedItem) Then
   MsgE "No se puede borrar la localidad (tiene clientes relacionados)"
   Exit Sub
End If

If Not MsgP("¿Confirma el borrado de la localidad seleccionada?") Then Exit Sub

'otras validaciones
If Not ExisteLocalidad(lvlocalidad.SelectedItem) Then
   MsgE "La localidad no existe"
   Exit Sub
End If

'inicio de transaccion
cnSQL.BeginTrans

sql = "delete from localidades WHERE idlocalidad=" & CLng(lvlocalidad.SelectedItem)

cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

MsgI "La localidad fue borrada"
lvlocalidad.SetFocus
   
Exit Sub
merror:
tratarerrores "Error borrando localidades"
End Sub
Private Sub cmdGrabar_Click()
Dim sql As String
Dim Mensaje As String
Dim IdLocalidad As Long
Dim IdProvincia As Long
On Error GoTo merror
Call RefreshTimer

If Not datosok() Then Exit Sub
    
IdProvincia = CLng(ComboProvincias.ItemData(ComboProvincias.ListIndex))

If TipoEdicion = "N" Then
   If Not MsgP("¿Confirma la nueva localidad?") Then Exit Sub
   
   IdLocalidad = UltimoId("idlocalidad", "localidades") + 1
   
   'otras validaciones
   If ExisteLocalidad(IdLocalidad) Then
      MsgE "La localidad ya existe"
      Exit Sub
   End If
   
   If Not ExisteProvincia(IdProvincia) Then
      MsgE "La provincia no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans

   'si es predeterminada saco el predeterminada al resto
   If CheckPredeterminada.Value = 1 Then
      sql = "update localidades set predeterminada=0"
      cnSQL.Execute sql
   End If
   
   sql = "INSERT INTO Localidades (Idlocalidad, nombre, idprovincia, codigopostal,predeterminada) " & _
         "VALUES (" & CLng(IdLocalidad) & ",'" & CStr(txtDatos(1).Text) & "'," & CLng(IdProvincia) & ",' " & CStr(txtDatos(2).Text) & "'," & CheckPredeterminada.Value & ")"
      
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "La localidad fue agregada"
   
   Call CargarLista
   Call CargarDatos

Else
   If Not MsgP("¿Confirma la modificacion de la localidad seleccionada?") Then Exit Sub
   
   'otras validaciones
   If Not ExisteLocalidad(lvlocalidad.SelectedItem) Then
      MsgE "La localidad no existe"
      Exit Sub
   End If
   
   If Not ExisteProvincia(IdProvincia) Then
      MsgE "La provincia no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans
   
   'si es predeterminada saco el predeterminada al resto
   If CheckPredeterminada.Value = 1 Then
      sql = "update localidades set predeterminada=0"
      cnSQL.Execute sql
   End If
   
   sql = "UPDATE localidades SET " & _
         "nombre='" & CStr(txtDatos(1).Text) & _
         "',codigopostal='" & CStr(txtDatos(2).Text) & _
         "',idprovincia=" & CLng(IdProvincia) & _
         ",predeterminada=" & CheckPredeterminada.Value & _
         " WHERE Idlocalidad=" & CLng(lvlocalidad.SelectedItem)
   
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "La localidad fue modificada"
   
   lvlocalidad.SelectedItem.ListSubItems(1).Text = txtDatos(1).Text & vbNullString

End If

TipoEdicion = "C"
Call SetearEntorno

MsgI Mensaje

lvlocalidad.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando localidades"
End Sub
Private Sub cmdModificar_Click()
'predispone a modificar solo si hay datos en el listview y hay seleccion
Call RefreshTimer
   
If Not VerificarSeleccionLista(lvlocalidad) Then Exit Sub

TipoEdicion = "M"
Call SetearEntorno

End Sub
Private Sub cmdnuevo_Click()
Call RefreshTimer
TipoEdicion = "N"
Call SetearEntorno
End Sub
Private Sub CargarLista()
'carga el listview con las localidades y su correspondiente provincia
Dim rec As rdoResultset
Dim sql As String
Dim Nitem As ListItem
On Error GoTo merror
    
 sql = "SELECT IDlocalidad,Nombre AS Localidad " & _
       "FROM localidades " & _
       "ORDER BY localidades.nombre"

Set rec = cnSQL.OpenResultset(sql)

lvlocalidad.ListItems.Clear
    
If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lvlocalidad.ListItems.Add(, , Format(rec.rdoColumns("Idlocalidad"), "0000"))
      Nitem.SubItems(1) = rec.rdoColumns("Localidad") & vbNullString
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error cargando la lista de localidades"
End Sub
Private Sub CargarDatos()
'Pone los datos del item seleccionado del listview en los campos de abajo
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror
    
    
If Not VerificarSeleccionLista(lvlocalidad) Then Exit Sub
        
sql = "SELECT localidades.IDlocalidad,localidades.nombre AS Localidad," & _
      "localidades.Codigopostal,provincias.nombre AS Provincia," & _
      "localidades.idprovincia,localidades.predeterminada " & _
      "FROM provincias " & _
      "INNER JOIN LOCALIDADES ON PROVINCIAS.IDprovincia=LOCALIDADES.idprovincia " & _
      "WHERE Idlocalidad=" & CLng(lvlocalidad.SelectedItem)

Set rec = cnSQL.OpenResultset(sql)
        
If Not rec.EOF Then
   txtDatos(1).Text = rec.rdoColumns("localidad") & vbNullString
   txtDatos(2).Text = rec.rdoColumns("Codigopostal") & vbNullString
   
   If ComboProvincias.ListCount() > 0 Then
      ComboProvincias.Text = rec.rdoColumns("provincia")
   End If
   
   If rec.rdoColumns("predeterminada") Then
      CheckPredeterminada.Value = 1
   Else
      CheckPredeterminada.Value = 0
   End If
End If
        
Exit Sub
merror:
tratarerrores "Error cargando datos de localidades"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(txtDatos(1).Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el nombre de la localidad"
   txtDatos(1).SetFocus
   Exit Function
End If

If Trim(txtDatos(2).Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el codigo postal"
   txtDatos(2).SetFocus
   Exit Function
End If

If ComboProvincias.Text = "" Then
   datosok = False
   MsgE "Debe seleccionar la provincia"
   ComboProvincias.SetFocus
   Exit Function
End If

Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-Localidades"
End Function
Private Sub SetearEntorno()
On Error GoTo merror

Select Case TipoEdicion
       Case "C"
            fmeDatos.Enabled = False
            cmdGrabar.Enabled = False
            CmdNuevo.Enabled = True
            CmdRefrescar.Enabled = True
            If lvlocalidad.ListItems.Count > 0 Then
               cmdModificar.Enabled = True
               CmdBorrar.Enabled = True
            Else
               cmdModificar.Enabled = False
               CmdBorrar.Enabled = False
               Call LimpiarCampos(Me)
            End If
            cmdcerrar.Caption = "&Cerrar"
            lvlocalidad.Enabled = True
            Call ColorCyan(Me)
        Case "M"
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            CmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            CmdBorrar.Enabled = False
            CmdRefrescar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lvlocalidad.Enabled = False
            txtDatos(1).SetFocus
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
            lvlocalidad.Enabled = False
            txtDatos(1).SetFocus
            Call ColorBlanco(Me)
    End Select

Exit Sub
merror:
tratarerrores "Error seteando el entorno-LocalidadesAbm"
End Sub
Private Sub lvlocalidad_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'ordena el listview pero solo si tiene datos
Dim Orden As Integer
    
If lvlocalidad.ListItems.Count > 1 Then
   lvlocalidad.SortKey = ColumnHeader.Index - 1
   Orden = lvlocalidad.SortKey
   lvlocalidad.SortOrder = Abs(Not lvlocalidad.SortOrder = 1)
   lvlocalidad.Sorted = True
End If

End Sub
Private Sub lvlocalidad_ItemClick(ByVal Item As MSComctlLib.ListItem)
'dentro de la funcion chequea que haya datos en el listview
Call CargarDatos
End Sub
Private Sub TxtDatos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   If Index < 2 Then
      txtDatos(Index + 1).SetFocus
   Else
      ComboProvincias.SetFocus
   End If
End If
If KeyCode = vbKeyUp Then
   If Index > 1 Then
      txtDatos(Index - 1).SetFocus
   End If
End If
End Sub
Private Sub TxtDatos_LostFocus(Index As Integer)
txtDatos(Index).Text = UCase(Trim(txtDatos(Index).Text))
End Sub

