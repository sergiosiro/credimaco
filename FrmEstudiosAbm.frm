VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEstudiosAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Estudios"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   Icon            =   "FrmEstudiosAbm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fmeDatos 
      BorderStyle     =   0  'None
      Caption         =   "Registrar Localidades"
      ForeColor       =   &H00FF0000&
      Height          =   1545
      Left            =   0
      TabIndex        =   11
      Top             =   2400
      Width           =   5175
      Begin VB.TextBox txtDatos 
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   50
         TabIndex        =   7
         ToolTipText     =   "Nombre de la localidad"
         Top             =   360
         Width           =   4095
      End
      Begin VB.ComboBox comboprovincias 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Provincia a la cual pertenece la localidad"
         Top             =   720
         Width           =   4095
      End
      Begin VB.CheckBox CheckPredeterminado 
         Alignment       =   1  'Right Justify
         Caption         =   "Predeterminado:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Si marca la casilla, la localidad sera la predeterminada en las demas pantallas del sistema"
         Top             =   1200
         Width           =   1455
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
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Provincia:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdNuevo 
      BackColor       =   &H00C0C000&
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Permite agregar los datos de una nueva localidad"
      Top             =   840
      Width           =   1185
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00C0C000&
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      ToolTipText     =   "Permite modificar los datos de la localidad seleccionada"
      Top             =   1440
      Width           =   1185
   End
   Begin VB.CommandButton cmdBorrar 
      BackColor       =   &H00C0C000&
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      ToolTipText     =   "Permite borrar la localidad seleccionada"
      Top             =   2040
      Width           =   1185
   End
   Begin VB.CommandButton cmdGrabar 
      BackColor       =   &H00C0C000&
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      ToolTipText     =   "Graba los datos de una localidad"
      Top             =   2640
      Width           =   1185
   End
   Begin VB.CommandButton cmdcerrar 
      BackColor       =   &H00C0C000&
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      ToolTipText     =   "Cierra la pantalla o cancela una operacion"
      Top             =   3600
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de localidades:"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5175
      Begin MSComctlLib.ListView lvestudio 
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
   Begin VB.CommandButton CmdRefrescar 
      BackColor       =   &H00C0C000&
      Caption         =   "Refrescar"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "Refresca los datos de la pantalla por si hubo cambios desde otra PC en red"
      Top             =   0
      Width           =   1185
   End
End
Attribute VB_Name = "FrmEstudiosAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call RefreshTimer
On Error GoTo merror

Call CargarComboProvincias("provincias", comboprovincias)

Call CargarLista
Call CargarDatos

TipoEdicion = "C"
Call SetearEntorno

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de estudios"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
If TipoEdicion = "C" Then
   Unload Me
Else
   TipoEdicion = "C"
   Call SetearEntorno
   Call CargarDatos
   lvestudio.SetFocus
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
Private Function PuedoBorrarEstudio(ByVal IdEstudio As Long) As Boolean
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

PuedoBorrarEstudio = True

'verifico en tabla clientes
sql = "select idestudio " & _
      "from creditosbloqueados " & _
      "where idestudio=" & CLng(IdEstudio)
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idestudio")) Then
      PuedoBorrarEstudio = False
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion PuedoBorrarestudio"
End Function
Private Sub cmdborrar_Click()
Dim sql As String
On Error GoTo merror
Call RefreshTimer

If Not VerificarSeleccionLista(lvestudio) Then Exit Sub

If Not PuedoBorrarEstudio(lvestudio.SelectedItem) Then
   MsgE "No se puede borrar el estudio (tiene créditos bloqueados relacionados)"
   Exit Sub
End If

If Not MsgP("¿Confirma el borrado del estudio seleccionada?") Then Exit Sub

'otras validaciones
If Not ExisteEstudio(lvestudio.SelectedItem) Then
   MsgE "El estudio no existe"
   Exit Sub
End If

'inicio de transaccion
cnSQL.BeginTrans

sql = "delete from estudios WHERE idestudio=" & CLng(lvestudio.SelectedItem)

cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

MsgI "El estudio fue borrado"
lvestudio.SetFocus
   
Exit Sub
merror:
tratarerrores "Error borrando estudios"
End Sub
Private Sub CmdGrabar_Click()
Dim sql As String
Dim Mensaje As String
Dim IdEstudio As Long
Dim IdProvincia As Long
On Error GoTo merror
Call RefreshTimer

If Not datosok() Then Exit Sub
    
IdProvincia = CLng(comboprovincias.ItemData(comboprovincias.ListIndex))

If TipoEdicion = "N" Then
   If Not MsgP("¿Confirma el nuevo estudio?") Then Exit Sub
   
   IdEstudio = UltimoId("idestudio", "estudios") + 1
   
   'otras validaciones
   If ExisteEstudio(IdEstudio) Then
      MsgE "El estudio ya existe"
      Exit Sub
   End If
   
   If Not ExisteProvincia(IdProvincia) Then
      MsgE "La provincia no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans

   'si es predeterminado saco el predeterminado al resto
   If CheckPredeterminado.Value = 1 Then
      sql = "update estudios set predeterminado=0"
      cnSQL.Execute sql
   End If
   
   sql = "INSERT INTO estudios (Idestudio, nombre, idprovincia, predeterminado) " & _
         "VALUES (" & CLng(IdEstudio) & ",'" & CStr(txtDatos(1).Text) & "'," & CLng(IdProvincia) & "," & CheckPredeterminado.Value & ")"
      
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El estudio fue agregado"
   
   Call CargarLista
   Call CargarDatos

Else
   If Not MsgP("¿Confirma la modificacion del estudio seleccionado?") Then Exit Sub
   
   'otras validaciones
   If Not ExisteEstudio(lvestudio.SelectedItem) Then
      MsgE "El estudio no existe"
      Exit Sub
   End If
   
   If Not ExisteProvincia(IdProvincia) Then
      MsgE "La provincia no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans
   
   'si es predeterminado saco el predeterminado al resto
   If CheckPredeterminado.Value = 1 Then
      sql = "update estudios set predeterminado=0"
      cnSQL.Execute sql
   End If
   
   sql = "UPDATE estudios SET " & _
         "nombre='" & CStr(txtDatos(1).Text) & _
         "',idprovincia=" & CLng(IdProvincia) & _
         ",predeterminado=" & CheckPredeterminado.Value & _
         " WHERE Idestudio=" & CLng(lvestudio.SelectedItem)
   
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El estudio fue modificado"
   
   lvestudio.SelectedItem.ListSubItems(1).Text = txtDatos(1).Text & vbNullString

End If

TipoEdicion = "C"
Call SetearEntorno

MsgI Mensaje

lvestudio.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando estudios"
End Sub
Private Sub cmdModificar_Click()
'predispone a modificar solo si hay datos en el listview y hay seleccion
Call RefreshTimer
   
If Not VerificarSeleccionLista(lvestudio) Then Exit Sub

TipoEdicion = "M"
Call SetearEntorno

End Sub
Private Sub cmdnuevo_Click()
Call RefreshTimer
TipoEdicion = "N"
Call SetearEntorno
End Sub
Private Sub CargarLista()
'carga el listview con los estudios y su correspondiente provincia
Dim rec As rdoResultset
Dim sql As String
Dim Nitem As ListItem
On Error GoTo merror
    
 sql = "SELECT IDestudio,Nombre AS estudio " & _
       "FROM estudios " & _
       "ORDER BY estudios.nombre"

Set rec = cnSQL.OpenResultset(sql)

lvestudio.ListItems.Clear
    
If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lvestudio.ListItems.Add(, , Format(rec.rdoColumns("Idestudio"), "0000"))
      Nitem.SubItems(1) = rec.rdoColumns("estudio") & vbNullString
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error cargando la lista de estudios"
End Sub
Private Sub CargarDatos()
'Pone los datos del item seleccionado del listview en los campos de abajo
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror
    
    
If Not VerificarSeleccionLista(lvestudio) Then Exit Sub
        
sql = "SELECT estudios.IDestudio,estudios.nombre AS estudio," & _
      "provincias.nombre AS Provincia," & _
      "estudios.idprovincia,estudios.predeterminado " & _
      "FROM provincias " & _
      "INNER JOIN estudios ON PROVINCIAS.IDprovincia=estudios.idprovincia " & _
      "WHERE Idestudio=" & CLng(lvestudio.SelectedItem)

Set rec = cnSQL.OpenResultset(sql)
        
If Not rec.EOF Then
   txtDatos(1).Text = rec.rdoColumns("estudio") & vbNullString
   
   If comboprovincias.ListCount() > 0 Then
      comboprovincias.Text = rec.rdoColumns("provincia")
   End If
   
   If rec.rdoColumns("predeterminado") Then
      CheckPredeterminado.Value = 1
   Else
      CheckPredeterminado.Value = 0
   End If
End If
        
Exit Sub
merror:
tratarerrores "Error cargando datos de estudios"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(txtDatos(1).Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el nombre del estudio"
   txtDatos(1).SetFocus
   Exit Function
End If

If comboprovincias.Text = "" Then
   datosok = False
   MsgE "Debe seleccionar la provincia"
   comboprovincias.SetFocus
   Exit Function
End If

Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-estudios"
End Function
Private Sub SetearEntorno()
On Error GoTo merror

Select Case TipoEdicion
       Case "C"
            fmeDatos.Enabled = False
            cmdGrabar.Enabled = False
            cmdNuevo.Enabled = True
            CmdRefrescar.Enabled = True
            If lvestudio.ListItems.Count > 0 Then
               cmdModificar.Enabled = True
               cmdBorrar.Enabled = True
            Else
               cmdModificar.Enabled = False
               cmdBorrar.Enabled = False
               Call LimpiarCampos(Me)
            End If
            cmdcerrar.Caption = "&Cerrar"
            lvestudio.Enabled = True
            Call ColorCyan(Me)
        Case "M"
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            cmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            cmdBorrar.Enabled = False
            CmdRefrescar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lvestudio.Enabled = False
            txtDatos(1).SetFocus
            Call ColorBlanco(Me)
        Case "N"
            Call LimpiarCampos(Me)
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            cmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            cmdBorrar.Enabled = False
            CmdRefrescar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lvestudio.Enabled = False
            txtDatos(1).SetFocus
            Call ColorBlanco(Me)
    End Select

Exit Sub
merror:
tratarerrores "Error seteando el entorno-estudiosAbm"
End Sub
Private Sub lvestudio_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'ordena el listview pero solo si tiene datos
Dim Orden As Integer
    
If lvestudio.ListItems.Count > 1 Then
   lvestudio.SortKey = ColumnHeader.Index - 1
   Orden = lvestudio.SortKey
   lvestudio.SortOrder = Abs(Not lvestudio.SortOrder = 1)
   lvestudio.Sorted = True
End If

End Sub
Private Sub lvestudio_ItemClick(ByVal Item As MSComctlLib.ListItem)
'dentro de la funcion chequea que haya datos en el listview
Call CargarDatos
End Sub
Private Sub TxtDatos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   If Index < 2 Then
      txtDatos(Index + 1).SetFocus
   Else
      comboprovincias.SetFocus
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



