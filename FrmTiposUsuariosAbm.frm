VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTiposUsuariosAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Tipos de Usuarios"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   HelpContextID   =   32
   Icon            =   "FrmTiposUsuariosAbm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameUsuarios 
      Caption         =   "Lista de tipos de usuarios:"
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   5775
      Begin MSComctlLib.ListView lvUsuarios 
         Height          =   2055
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Lista de tipos de usuarios"
         Top             =   240
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3625
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo de usuario"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "(*)El tipo de usuario ADMINISTRADOR no se puede modificar ni borrar."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   5535
      End
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   6000
      TabIndex        =   22
      ToolTipText     =   "Cierra la pantalla o cancela una operacion de agregado o modificacion"
      Top             =   6000
      Width           =   1305
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6000
      TabIndex        =   21
      ToolTipText     =   "Graba los datos de un tipo de usuario"
      Top             =   2400
      Width           =   1305
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   6000
      TabIndex        =   20
      ToolTipText     =   "Permite borrar al tipo de usuario seleccionado"
      Top             =   1680
      Width           =   1305
   End
   Begin VB.CommandButton CmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   6000
      TabIndex        =   19
      ToolTipText     =   "Permite modificar los datos del tipo de usuario seleccionado"
      Top             =   960
      Width           =   1305
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      ToolTipText     =   "Permite agregar los datos de un nuevo tipo de usuario"
      Top             =   240
      Width           =   1305
   End
   Begin VB.Frame fmeDatos 
      Caption         =   "Datos del tipo de usuario:"
      ForeColor       =   &H00FF0000&
      Height          =   3450
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   5775
      Begin VB.CheckBox CheckImporta 
         Caption         =   "Importa cobros Rapipago, Pagofacil"
         Height          =   255
         Left            =   2520
         TabIndex        =   30
         Top             =   2760
         Width           =   2895
      End
      Begin VB.CheckBox Checkacercade 
         Caption         =   "Acerca de"
         Height          =   255
         Left            =   2520
         TabIndex        =   29
         Top             =   3000
         Width           =   2895
      End
      Begin VB.CheckBox CheckIngresos 
         Caption         =   "Consulta cobros por periodo"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CheckBox CheckExporta 
         Caption         =   "Exporta a Excel"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Puede exportar a planillas excel"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CheckBox CheckAdmPlanes 
         Caption         =   "Administra planes"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         ToolTipText     =   "Si el tipo de usuario puede registrar planes de creditos"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox CheckImprimeCuotas 
         Caption         =   "Imprime cuotas"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         ToolTipText     =   "Establece si podra imprimir cuotas"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox CheckEmiteCartaReclamo 
         Caption         =   "Emite carta reclamo"
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         ToolTipText     =   "Establece si podra imprimir la cartar reclamo"
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CheckBox CheckEfectuaBackup 
         Caption         =   "Efectua backups"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Establece si podra efectuar copias de seguridad de la base de datos"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox CheckActualizaOpciones 
         Caption         =   "Actualiza Opciones"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Establece si podra modificar opciones del sistema"
         Top             =   840
         Width           =   2175
      End
      Begin VB.CheckBox CheckAdmClientes 
         Caption         =   "Registra clientes"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Establece si podra registrar clientes"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox CheckConsultaDeudores 
         Caption         =   "Consulta deudores"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         ToolTipText     =   "Establece si podra consultar deudores"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CheckBox CheckConsulta 
         Caption         =   "Consulta creditos"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         ToolTipText     =   "Establece si podra consultar creditos"
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox CheckEmiteLibreDeuda 
         Caption         =   "Emite libre deuda"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         ToolTipText     =   "Establece si podra emitir el libre deuda"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox CheckAdmCobradores 
         Caption         =   "Administra cobradores"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         ToolTipText     =   "Establece si podra liquidar comisiones a cobradores"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CheckBox CheckAdmCreditos 
         Caption         =   "Borra/Bloquea/Finaliza creditos"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         ToolTipText     =   "Establece si podra borrar/bloquear creditos"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox CheckAnula 
         Caption         =   "Anula cobros"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Establece si podra anular cobro de cuotas"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CheckBox CheckCobra 
         Caption         =   "Cobra cuotas"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Establece si podra cobrar cuotas"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox CheckRefinancia 
         Caption         =   "Refinancia creditos"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Establece si podra refinanciar creditos"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox CheckRegistra 
         Caption         =   "Registra creditos"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Establece si podra registrar nuevos creditos"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox CheckActualiza 
         Caption         =   "Efectua actualizaciones"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Establece si podra registrar monedas, bancos, localidades, provincias, etc"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox TxtNombre 
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Descripcion del tipo de usuario"
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Detalle del tipo de usuario:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1875
      End
   End
End
Attribute VB_Name = "FrmTiposUsuariosAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***AQUI SE CARGAN LOS GRUPOS DE USUARIOS DEL SISTEMA. CADA GRUPO TIENE PERMISOS
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
tratarerrores "Error cargando la pantalla de Tipos de Usuarios"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
If TipoEdicion = "C" Then
   Unload Me
Else
   TipoEdicion = "C"
   Call SetearEntorno
   Call CargarDatos
   lvUsuarios.SetFocus
End If
End Sub
Private Sub cmdnuevo_Click()
Call RefreshTimer
TipoEdicion = "N"
Call SetearEntorno
End Sub
Private Function VerificarCasillas() As Boolean
'chequea que se han seleccionado al menos una casilla
On Error GoTo merror

VerificarCasillas = True

If CheckActualizaOpciones.Value = 0 And CheckEfectuaBackup.Value = 0 And CheckActualiza.Value = 0 _
   And CheckAdmClientes.Value = 0 And CheckRegistra.Value = 0 And CheckRefinancia.Value = 0 And CheckCobra.Value = 0 _
   And CheckAnula.Value = 0 And CheckConsulta.Value = 0 And CheckAdmCreditos.Value = 0 _
   And CheckImprimeCuotas.Value = 0 And CheckEmiteLibreDeuda.Value = 0 And CheckConsultaDeudores.Value = 0 _
   And CheckEmiteCartaReclamo.Value = 0 And CheckAdmCobradores.Value = 0 And CheckAdmPlanes.Value = 0 And CheckExporta.Value = 0 And CheckImporta.Value = 0 And CheckIngresos.Value = 0 And Checkacercade.Value = 0 Then
   VerificarCasillas = False
End If
  
Exit Function
merror:
tratarerrores "Error en funcion VerificarCasillas"
End Function
Private Sub cmdModificar_Click()
Call RefreshTimer
If Not VerificarSeleccionLista(lvUsuarios) Then Exit Sub

If UCase(Trim(lvUsuarios.SelectedItem.SubItems(1))) = "ADMINISTRADOR" Then
   MsgE "No se puede modificar al tipo de usuario ADMINISTRADOR"
   Exit Sub
End If

TipoEdicion = "M"
Call SetearEntorno

If fmeDatos.Enabled Then
   TxtNombre.SetFocus
End If
End Sub
Private Sub CargarLista()
'carga la lista de tipos de usuarios
Dim rec As rdoResultset
Dim sql As String
Dim Nitem As ListItem
On Error GoTo merror

lvUsuarios.ListItems.Clear

sql = "select * from tipousuario ORDER BY nombre"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lvUsuarios.ListItems.Add(, , Format(rec.rdoColumns("Idtipousuario"), "00"))
      Nitem.SubItems(1) = rec.rdoColumns("nombre") & vbNullString
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error cargando la lista de tipos de usuarios"
End Sub
Private Sub CargarDatos()
'carga los campos de abajo con un usuario
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror
    
If Not VerificarSeleccionLista(lvUsuarios) Then Exit Sub
       
sql = "SELECT nombre,consultaingresos,actualiza,registra,refinancia,cobra,anula,consulta,admcobradores,emitelibredeuda,consultadeudores,admcreditos,actualizaopciones,admclientes,efectuabackup,emitecartareclamo,imprimecuotas,admplanes,acercade,adminversores as exporta,admcomodines as importa " & _
      "FROM tipousuario " & _
      "WHERE Idtipousuario=" & CLng(lvUsuarios.SelectedItem)
      
Set rec = cnSQL.OpenResultset(sql)
        
If Not rec.EOF Then
   TxtNombre.Text = rec.rdoColumns("nombre") & vbNullString
   If rec.rdoColumns("actualiza") Then
      CheckActualiza.Value = 1
   Else
      CheckActualiza.Value = 0
   End If
   
   If rec.rdoColumns("acercade") Then
      Checkacercade.Value = 1
   Else
      Checkacercade.Value = 0
   End If
   
   If rec.rdoColumns("registra") Then
      CheckRegistra.Value = 1
   Else
      CheckRegistra.Value = 0
   End If
   If rec.rdoColumns("refinancia") Then
      CheckRefinancia.Value = 1
   Else
      CheckRefinancia.Value = 0
   End If
   
   If rec.rdoColumns("cobra") Then
      CheckCobra.Value = 1
   Else
      CheckCobra.Value = 0
   End If
   
   If rec.rdoColumns("anula") Then
      CheckAnula.Value = 1
   Else
      CheckAnula.Value = 0
   End If
   
   If rec.rdoColumns("consulta") Then
      CheckConsulta.Value = 1
   Else
      CheckConsulta.Value = 0
   End If
   
   If rec.rdoColumns("admcreditos") Then
      CheckAdmCreditos.Value = 1
   Else
      CheckAdmCreditos.Value = 0
   End If
   
   If rec.rdoColumns("admcobradores") Then
      CheckAdmCobradores.Value = 1
   Else
      CheckAdmCobradores.Value = 0
   End If
   
   If rec.rdoColumns("emitelibredeuda") Then
      CheckEmiteLibreDeuda.Value = 1
   Else
      CheckEmiteLibreDeuda.Value = 0
   End If
   
   If rec.rdoColumns("consultadeudores") Then
      CheckConsultaDeudores.Value = 1
   Else
      CheckConsultaDeudores.Value = 0
   End If
   
   If rec.rdoColumns("admclientes") Then
      CheckAdmClientes.Value = 1
   Else
      CheckAdmClientes.Value = 0
   End If
   
   If rec.rdoColumns("actualizaopciones") Then
      CheckActualizaOpciones.Value = 1
   Else
      CheckActualizaOpciones.Value = 0
   End If
   
   If rec.rdoColumns("efectuabackup") Then
      CheckEfectuaBackup.Value = 1
   Else
      CheckEfectuaBackup.Value = 0
   End If
   
   If rec.rdoColumns("emitecartareclamo") Then
      CheckEmiteCartaReclamo.Value = 1
   Else
      CheckEmiteCartaReclamo.Value = 0
   End If
   
   If rec.rdoColumns("imprimecuotas") Then
      CheckImprimeCuotas.Value = 1
   Else
      CheckImprimeCuotas.Value = 0
   End If
   
   If rec.rdoColumns("admplanes") Then
      CheckAdmPlanes.Value = 1
   Else
      CheckAdmPlanes.Value = 0
   End If
   
   If rec.rdoColumns("exporta") Then
      CheckExporta.Value = 1
   Else
      CheckExporta.Value = 0
   End If
   
   If rec.rdoColumns("importa") Then
      CheckImporta.Value = 1
   Else
      CheckImporta.Value = 0
   End If
   
   If rec.rdoColumns("consultaingresos") Then
      CheckIngresos.Value = 1
   Else
      CheckIngresos.Value = 0
   End If
End If

Exit Sub
merror:
tratarerrores "Error cargando datos de usuarios"
End Sub
Private Function PuedoBorrarTipoUsuario(ByVal Idtipousuario As Long) As Boolean
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

PuedoBorrarTipoUsuario = True

'verifico en usuarios
sql = "select idusuario from usuarios " & _
"where idtipousuario=" & CLng(Idtipousuario)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull("idusuario") Then
      PuedoBorrarTipoUsuario = False
      Exit Function
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion PuedoBorrarTipoUsuario"
End Function
Private Sub cmdborrar_Click()
Dim sql As String
Dim nombre As String
On Error GoTo merror
Call RefreshTimer

If Not VerificarSeleccionLista(lvUsuarios) Then Exit Sub

nombre = UCase(Trim(lvUsuarios.SelectedItem.SubItems(1)))

'si deseo borrar un administrador
If nombre = "ADMINISTRADOR" Then
   MsgE "No puede borrar el tipo de usuario ADMINISTRADOR"
   Exit Sub
End If

If Not PuedoBorrarTipoUsuario(lvUsuarios.SelectedItem) Then
   MsgE "No se puede borrar el tipo de usuario...tiene usuarios asociados"
   Exit Sub
End If

If Not MsgP("¿Confirma el borrado del tipo de usuario seleccionado?") Then Exit Sub

'otras validaciones
If Not ExisteTipoUsuario(lvUsuarios.SelectedItem) Then
   MsgE "El tipo de usuario no existe"
   Exit Sub
End If

'inicio transaccion
cnSQL.BeginTrans

sql = "delete from tipousuario " & _
      "where idtipousuario=" & CLng(lvUsuarios.SelectedItem)

cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

MsgI "El tipo de usuario fue borrado"

lvUsuarios.SetFocus

Exit Sub
merror:
tratarerrores "Error borrando tipos de usuarios"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(TxtNombre.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el nombre del tipo de usuario"
   TxtNombre.SetFocus
   Exit Function
End If

If TipoEdicion = "N" Then
   If UCase(Trim(TxtNombre.Text)) = "ADMINISTRADOR" Then
      MsgE "No se puede agregar otro tipo de usuario ADMINISTRADOR"
      datosok = False
      Exit Function
   End If
End If

If TipoEdicion = "M" Then
   If UCase(lvUsuarios.SelectedItem.SubItems(1)) = "ADMINISTRADOR" Then
      MsgE "No se puede modificar al administrador"
      datosok = False
      Exit Function
   End If
End If

If Not VerificarCasillas() Then
   datosok = False
   MsgE "Debe marcar casillas de opciones"
   Exit Function
End If

'reemplazo los caracteres invalidos
Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-TiposUsuariosAbm"
End Function
Private Sub cmdGrabar_Click()
Dim sql As String
Dim Idtipousuario As Long
Dim Mensaje As String
On Error GoTo merror
Call RefreshTimer

If Not datosok() Then Exit Sub

If TipoEdicion = "N" Then
   
   If Not MsgP("¿Confirma el nuevo tipo de usuario?") Then Exit Sub
   
   Idtipousuario = UltimoId("idtipousuario", "tipousuario") + 1
   
   'otras validaciones
   If ExisteTipoUsuario(Idtipousuario) Then
      MsgE "El tipo de usuario ya existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans
        
   sql = "INSERT INTO tipousuario (Idtipousuario,nombre,actualiza,registra," & _
         "refinancia,cobra,anula,consulta,admcreditos,imprimecuotas," & _
         "admcobradores,emitelibredeuda,consultadeudores," & _
         "admclientes,actualizaopciones,efectuabackup,emitecartareclamo,admplanes,adminversores,admcomodines,consultaingresos,acercade) " & _
         "VALUES (" & CLng(Idtipousuario) & ",'" & CStr(TxtNombre.Text) & _
         "'," & CheckActualiza.Value & "," & CheckRegistra.Value & _
         "," & CheckRefinancia.Value & "," & CheckCobra.Value & _
         "," & CheckAnula.Value & "," & CheckConsulta.Value & _
         "," & CheckAdmCreditos.Value & "," & CheckImprimeCuotas.Value & _
         "," & CheckAdmCobradores.Value & _
         "," & CheckEmiteLibreDeuda.Value & "," & CheckConsultaDeudores.Value & _
         "," & CheckAdmClientes.Value & "," & CheckActualizaOpciones.Value & _
         "," & CheckEfectuaBackup.Value & "," & CheckEmiteCartaReclamo.Value & "," & CheckAdmPlanes.Value & "," & CheckExporta.Value & "," & CheckImporta.Value & "," & CheckIngresos.Value & "," & Checkacercade.Value & ")"
        
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El tipo de usuario fue agregado"
   
   Call CargarLista
   Call CargarDatos
Else
   
   If Not MsgP("¿Confirma la modificacion del tipo de usuario seleccionado?") Then Exit Sub
  
   'otras validaciones
   If Not ExisteTipoUsuario(lvUsuarios.SelectedItem) Then
      MsgE "El tipo de usuario no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans

   sql = "UPDATE tipousuario SET nombre='" & CStr(TxtNombre.Text) & _
         "',actualiza=" & CheckActualiza.Value & _
         ",registra=" & CheckRegistra.Value & _
         ",refinancia=" & CheckRefinancia.Value & _
         ",cobra=" & CheckCobra.Value & ",anula=" & CheckAnula.Value & _
         ",consulta=" & CheckConsulta.Value & _
         ",admcreditos=" & CheckAdmCreditos.Value & _
         ",imprimecuotas=" & CheckImprimeCuotas.Value & _
         ",admcobradores=" & CheckAdmCobradores.Value & _
         ",emitelibredeuda=" & CheckEmiteLibreDeuda.Value & _
         ",consultadeudores=" & CheckConsultaDeudores.Value & _
         ",admclientes=" & CheckAdmClientes.Value & _
         ",actualizaopciones=" & CheckActualizaOpciones.Value & _
         ",efectuabackup=" & CheckEfectuaBackup.Value & _
         ",acercade=" & Checkacercade.Value & _
         ",emitecartareclamo=" & CheckEmiteCartaReclamo.Value & _
         ",consultaingresos=" & CheckIngresos.Value & _
         ",admplanes=" & CheckAdmPlanes.Value & ",adminversores=" & CheckExporta.Value & ",admcomodines=" & CheckImporta.Value & _
         " WHERE Idtipousuario=" & CLng(lvUsuarios.SelectedItem)
        
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El tipo de usuario fue modificado"
   
   lvUsuarios.SelectedItem.ListSubItems(1).Text = TxtNombre.Text & vbNullString
    
End If

TipoEdicion = "C"
Call SetearEntorno

MsgI Mensaje

lvUsuarios.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando tipos de usuarios"
End Sub
Private Sub SetearEntorno()
On Error GoTo merror

Select Case TipoEdicion
       Case "C"
            fmeDatos.Enabled = False
            cmdGrabar.Enabled = False
            CmdNuevo.Enabled = True
            If lvUsuarios.ListItems.Count > 0 Then
               cmdModificar.Enabled = True
               CmdBorrar.Enabled = True
            Else
               cmdModificar.Enabled = False
               CmdBorrar.Enabled = False
               Call LimpiarCampos(Me)
            End If
            cmdcerrar.Caption = "&Cerrar"
            lvUsuarios.Enabled = True
            Call ColorCyan(Me)
       Case "M"
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            CmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            CmdBorrar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lvUsuarios.Enabled = False
            Call ColorBlanco(Me)
        Case "N"
            Call LimpiarCampos(Me)
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            CmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            CmdBorrar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lvUsuarios.Enabled = False
            If fmeDatos.Enabled Then
               TxtNombre.SetFocus
            End If
            Call ColorBlanco(Me)
        End Select

Exit Sub
merror:
tratarerrores "Error seteando entorno-TiposUsuariosAbm"
End Sub
Private Sub lvUsuarios_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Orden As Integer
    
If lvUsuarios.ListItems.Count > 1 Then
   lvUsuarios.SortKey = ColumnHeader.Index - 1
   Orden = lvUsuarios.SortKey
   lvUsuarios.SortOrder = Abs(Not lvUsuarios.SortOrder = 1)
   lvUsuarios.Sorted = True
End If

End Sub
Private Sub lvUsuarios_ItemClick(ByVal Item As MSComctlLib.ListItem)
 Call CargarDatos
End Sub
Private Sub TxtNombre_LostFocus()
TxtNombre.Text = UCase(Trim(TxtNombre.Text))
End Sub
