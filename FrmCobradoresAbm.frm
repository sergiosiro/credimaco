VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCobradoresAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Cobradores"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   HelpContextID   =   9
   Icon            =   "FrmCobradoresAbm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdRefrescar 
      Caption         =   "Refrescar"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      ToolTipText     =   "Refresca los datos de la pantalla por si hubo cambios desde otra PC en red"
      Top             =   120
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de cobradores:"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   5175
      Begin MSComctlLib.ListView lv 
         Height          =   1935
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Lista de cobradores"
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
            Text            =   "Cobrador"
            Object.Width           =   8820
         EndProperty
      End
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      ToolTipText     =   "Cierra la pantalla o cancela una operacion"
      Top             =   5880
      Width           =   1305
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      ToolTipText     =   "Graba los datos de un cobrador "
      Top             =   4320
      Width           =   1305
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   5400
      TabIndex        =   14
      ToolTipText     =   "Permite borrar el cobrador seleccionado"
      Top             =   3720
      Width           =   1305
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      ToolTipText     =   "Permite modificar los datos del cobrador seleccionado"
      Top             =   3120
      Width           =   1305
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      ToolTipText     =   "Permite agregar los datos de un nuevo cobrador"
      Top             =   2520
      Width           =   1305
   End
   Begin VB.Frame fmeDatos 
      Caption         =   "Datos"
      ForeColor       =   &H00FF0000&
      Height          =   3825
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   5175
      Begin VB.TextBox TxtPorcentajeComision 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3720
         MaxLength       =   9
         TabIndex        =   25
         Top             =   3000
         Width           =   840
      End
      Begin VB.TextBox TxtImporteComision 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3720
         MaxLength       =   9
         TabIndex        =   24
         Top             =   2640
         Width           =   840
      End
      Begin VB.CheckBox CheckAplicarComision 
         Caption         =   "Aplicar comision por cobro de cuotas"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   3135
      End
      Begin VB.OptionButton OptionPorcentaje 
         Caption         =   "Aplicar porcentaje de comision   %:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   3000
         Width           =   2775
      End
      Begin VB.OptionButton OptionImporte 
         Caption         =   "Aplicar importe de comision        $:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CheckBox CheckActivo 
         Caption         =   "Cobrador activo"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         ToolTipText     =   "Establece que el cobrador pueda efectuar cobros"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtDatos 
         Height          =   285
         Index           =   4
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "Telefono del cobrador"
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox txtDatos 
         Height          =   285
         Index           =   3
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Domicilio del cobrador"
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox txtDatos 
         Height          =   285
         Index           =   0
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Nombre del cobrador"
         Top             =   360
         Width           =   3975
      End
      Begin VB.CheckBox CheckPredeterminada 
         Caption         =   "Predeterminado"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Si marca la casilla, el cobrador aparecera primero en las demas pantallas"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtDatos 
         Height          =   285
         Index           =   2
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Nº de documento del cobrador"
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtDatos 
         Height          =   285
         Index           =   1
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Apellido del cobrador"
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "Domicilio:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Documento:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Telefono:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmCobradoresAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AQUI SE REGISTRAN LOS COBRADORES DE CUOTAS DEL SISTEMA Y A LOS CUALES
'SE LES ASIGNA UNA COMISION POR COBRO REGISTRADO.LUEGO ESAS COMISIONES SE LIQUIDAN
'EN LA PANTALLA DE ADMINISTRACION DE COBRADORES

Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call CargarLista
Call CargarDatos

TipoEdicion = "C"
Call SetearEntorno
Exit Sub
merror:
tratarerrores "Error cargando la pantalla de cobradores"
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
Private Function PuedoBorrarCobrador(ByVal IdCobrador As Long) As Boolean
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

PuedoBorrarCobrador = True

'verifico en tabla cobradorespagos
sql = "select idcobrador " & _
      "from cobradorespagos " & _
      "where idcobrador=" & CLng(IdCobrador)
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcobrador")) Then
      PuedoBorrarCobrador = False
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion PuedoBorrarCobrador"
End Function
Private Sub cmdborrar_Click()
Dim sql As String
On Error GoTo merror
Call RefreshTimer

If Not VerificarSeleccionLista(lv) Then Exit Sub

If Not PuedoBorrarCobrador(lv.SelectedItem) Then
   MsgE "No se puede borrar el cobrador (tiene registros relacionados)"
   Exit Sub
End If

If Not MsgP("¿Confirma el borrado del cobrador seleccionado?") Then Exit Sub

'otras validaciones
If Not ExisteCobrador(lv.SelectedItem) Then
   MsgE "El cobrador no existe"
   Exit Sub
End If

'inicio de transaccion
cnSQL.BeginTrans

sql = "delete from cobradores WHERE idcobrador=" & CLng(lv.SelectedItem)

cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

MsgI "El cobrador fue borrado"
lv.SetFocus
   
Exit Sub
merror:
tratarerrores "Error borrando cobradores"
End Sub
Private Sub cmdGrabar_Click()
Dim sql As String
Dim Mensaje As String
Dim IdCobrador As Long
On Error GoTo merror
Call RefreshTimer

If Not datosok() Then Exit Sub
    
If TipoEdicion = "N" Then
   If Not MsgP("¿Confirma el nuevo cobrador?") Then Exit Sub
   
   IdCobrador = UltimoId("idcobrador", "cobradores") + 1
   
   'otras validaciones
   If ExisteCobrador(IdCobrador) Then
      MsgE "El cobrador ya existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans

   'si es predeterminada saco el predeterminada al resto
   If CheckPredeterminada.Value = 1 Then
      sql = "update cobradores set predeterminada=0"
      cnSQL.Execute sql
   End If
   
   sql = "INSERT INTO cobradores (Idcobrador,nombre,apellido,documento,domicilio,telefono,predeterminada,activo,importecomision,porcentajecomision,aplicarcomision) " & _
         "VALUES (" & CLng(IdCobrador) & ",'" & CStr(txtDatos(0).Text) & "','" & CStr(txtDatos(1).Text) & "','" & CStr(txtDatos(2).Text) & "','" & CStr(txtDatos(3).Text) & "','" & CStr(txtDatos(4).Text) & "'," & CheckPredeterminada.Value & "," & CheckActivo.Value & "," & ConvertirDblSql(CCur(TxtImporteComision.Text)) & "," & ConvertirDblSql(CDbl(TxtPorcentajeComision.Text)) & "," & CheckAplicarComision.Value & ")"
      
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El nuevo cobrador fue agregado"
   
   Call CargarLista
   Call CargarDatos

Else
   If Not MsgP("¿Confirma la modificacion del cobrador seleccionado?") Then Exit Sub
   
   'otras validaciones
   If Not ExisteCobrador(lv.SelectedItem) Then
      MsgE "El cobrador no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans
   
   'si es predeterminada saco el predeterminada al resto
   If CheckPredeterminada.Value = 1 Then
      sql = "update cobradores set predeterminada=0"
      cnSQL.Execute sql
   End If
   
   sql = "UPDATE cobradores SET " & _
         "nombre='" & CStr(txtDatos(0).Text) & _
         "',apellido='" & CStr(txtDatos(1).Text) & _
         "',documento='" & CStr(txtDatos(2).Text) & _
         "',domicilio='" & CStr(txtDatos(3).Text) & _
         "',telefono='" & CStr(txtDatos(4).Text) & _
         "',importecomision=" & ConvertirDblSql(CCur(TxtImporteComision.Text)) & _
         ",porcentajecomision=" & ConvertirDblSql(CDbl(TxtPorcentajeComision.Text)) & _
         ",predeterminada=" & CheckPredeterminada.Value & _
         ",activo=" & CheckActivo.Value & _
         ",aplicarcomision=" & CheckAplicarComision.Value & _
         " WHERE Idcobrador=" & CLng(lv.SelectedItem)
   
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El cobrador fue modificado"
   
   lv.SelectedItem.ListSubItems(1).Text = txtDatos(1).Text & " " & txtDatos(0).Text & vbNullString

End If

TipoEdicion = "C"
Call SetearEntorno

MsgI Mensaje

lv.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando cobradores"
End Sub
Private Sub cmdModificar_Click()
'predispone a modificar solo si hay datos en el listview y hay seleccion
Call RefreshTimer
   
If Not VerificarSeleccionLista(lv) Then Exit Sub

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
    
 sql = "SELECT idcobrador,apellido + ' ' + nombre as cobrador " & _
       "FROM cobradores " & _
       "ORDER BY apellido,nombre"

Set rec = cnSQL.OpenResultset(sql)

lv.ListItems.Clear
    
If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lv.ListItems.Add(, , Format(rec.rdoColumns("Idcobrador"), "000"))
      Nitem.SubItems(1) = rec.rdoColumns("cobrador") & vbNullString
      
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error cargando la lista de cobradores"
End Sub
Private Sub CargarDatos()
'Pone los datos del item seleccionado del listview en los campos de abajo
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror
    
If Not VerificarSeleccionLista(lv) Then Exit Sub
        
sql = "SELECT * " & _
      "FROM cobradores " & _
      "WHERE Idcobrador=" & CLng(lv.SelectedItem)

Set rec = cnSQL.OpenResultset(sql)
        
If Not rec.EOF Then
   txtDatos(0).Text = rec.rdoColumns("nombre") & vbNullString
   txtDatos(1).Text = rec.rdoColumns("apellido") & vbNullString
   txtDatos(2).Text = rec.rdoColumns("documento") & vbNullString
   txtDatos(3).Text = rec.rdoColumns("domicilio") & vbNullString
   txtDatos(4).Text = rec.rdoColumns("telefono") & vbNullString
   TxtImporteComision.Text = rec.rdoColumns("importecomision") & vbNullString
   TxtPorcentajeComision.Text = rec.rdoColumns("porcentajecomision") & vbNullString
   
   If rec.rdoColumns("aplicarcomision") Then
      CheckAplicarComision.Value = 1
   Else
      CheckAplicarComision.Value = 0
   End If
   
   If CCur(TxtImporteComision.Text) > 0 Then
      OptionImporte.Value = True
   End If
   
   If CDbl(TxtPorcentajeComision.Text) > 0 Then
      OptionPorcentaje.Value = True
   End If
   
   If rec.rdoColumns("predeterminada") Then
      CheckPredeterminada.Value = 1
   Else
      CheckPredeterminada.Value = 0
   End If
   
   If rec.rdoColumns("activo") Then
      CheckActivo.Value = 1
   Else
      CheckActivo.Value = 0
   End If
   
End If
        
Exit Sub
merror:
tratarerrores "Error cargando datos de cobradores"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(txtDatos(0).Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el nombre del cobrador"
   txtDatos(0).SetFocus
   Exit Function
End If

If Trim(txtDatos(1).Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el apellido del cobrador"
   txtDatos(1).SetFocus
   Exit Function
End If

If Trim(txtDatos(2).Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el documento del cobrador"
   txtDatos(2).SetFocus
   Exit Function
End If

If Trim(txtDatos(3).Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el domicilio del cobrador"
   txtDatos(3).SetFocus
   Exit Function
End If

If Trim(txtDatos(4).Text) = "" Then
   txtDatos(4).Text = vbNullString
End If

If CheckAplicarComision.Value = 0 Then
   datosok = False
   MsgE "Debe seleccionar la comision"
   Exit Function
End If

If Not OptionImporte.Value Then
   TxtImporteComision.Text = 0
Else
   If Trim(TxtImporteComision.Text) = "" Then
      datosok = False
      MsgE "Debe ingresar el importe de comision"
      TxtImporteComision.SetFocus
      Exit Function
   End If

   If Not IsNumeric(TxtImporteComision.Text) Then
      datosok = False
      MsgE "El importe de comision debe ser numerico"
      TxtImporteComision.SetFocus
      Exit Function
   End If

   If CCur(TxtImporteComision.Text) = 0 Then
      datosok = False
      MsgE "El importe de comision debe ser mayor a cero"
      TxtImporteComision.SetFocus
      Exit Function
   End If
End If

If Not OptionPorcentaje.Value Then
   TxtPorcentajeComision.Text = 0
Else
   If Trim(TxtPorcentajeComision.Text) = "" Then
      datosok = False
      MsgE "Debe ingresar el porcentaje de comision"
      TxtPorcentajeComision.SetFocus
      Exit Function
   End If
   If Not IsNumeric(TxtPorcentajeComision.Text) Then
      datosok = False
      MsgE "El porcentaje de comision debe ser numerico"
      TxtPorcentajeComision.SetFocus
      Exit Function
   End If

   If CDbl(TxtPorcentajeComision.Text) = 0 Then
      datosok = False
      MsgE "El porcentaje de comision debe ser mayor a cero"
      TxtPorcentajeComision.SetFocus
      Exit Function
   End If
End If

Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-Cobradores"
End Function
Private Sub SetearEntorno()
On Error GoTo merror

Select Case TipoEdicion
       Case "C"
            fmeDatos.Enabled = False
            cmdGrabar.Enabled = False
            CmdNuevo.Enabled = True
            CmdRefrescar.Enabled = True
            
            If lv.ListItems.Count > 0 Then
               cmdModificar.Enabled = True
               CmdBorrar.Enabled = True
            Else
               cmdModificar.Enabled = False
               CmdBorrar.Enabled = False
               Call LimpiarCampos(Me)
            End If
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
            txtDatos(0).SetFocus
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
            txtDatos(0).SetFocus
            Call ColorBlanco(Me)
    End Select

Exit Sub
merror:
tratarerrores "Error seteando el entorno-CobradoresAbm"
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
Private Sub CheckAplicarComision_Click()
On Error GoTo merror

If CheckAplicarComision.Value = 1 Then
   OptionImporte.Enabled = True
   OptionPorcentaje.Enabled = True
   OptionImporte.Value = False
   OptionImporte.Value = True
   If CCur(TxtImporteComision.Text) > 0 Then
      OptionImporte.Value = True
   End If
   If CDbl(TxtPorcentajeComision.Text) > 0 Then
      OptionPorcentaje.Value = True
   End If
Else
   OptionImporte.Enabled = False
   OptionPorcentaje.Enabled = False
   TxtImporteComision.Text = 0
   TxtPorcentajeComision.Text = 0
   TxtImporteComision.BackColor = &HFFFFC0
   TxtPorcentajeComision.BackColor = &HFFFFC0
   TxtImporteComision.Enabled = False
   TxtPorcentajeComision.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error aplicando comision de cobradores"
End Sub
Private Sub OptionImporte_Click()
On Error GoTo merror

If OptionImporte.Value Then
   TxtImporteComision.Enabled = True
   TxtImporteComision.BackColor = vbWhite
   TxtPorcentajeComision.BackColor = &HFFFFC0
   TxtPorcentajeComision.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error estableciendo importe de comision"
End Sub
Private Sub OptionPorcentaje_Click()
On Error GoTo merror

If OptionPorcentaje.Value Then
   TxtPorcentajeComision.Enabled = True
   TxtPorcentajeComision.BackColor = vbWhite
   TxtImporteComision.BackColor = &HFFFFC0
   TxtImporteComision.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error estableciendo porcentaje de comision"
End Sub
Private Sub TxtDatos_LostFocus(Index As Integer)
txtDatos(Index).Text = UCase(Trim(txtDatos(Index).Text))
End Sub
Private Sub TxtPorcentajeComision_Change()
'TxtPorcentajeComisionBis.Text = Format(CCur(TxtPorcentajeComision.Text) / 10, "0.00")
End Sub
