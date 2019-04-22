VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUsuariosAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Usuarios del sistema"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   HelpContextID   =   5
   Icon            =   "FrmUsuariosAbm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameUsuarios 
      Caption         =   "Lista de usuarios:"
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5055
      Begin MSComctlLib.ListView lvUsuarios 
         Height          =   2055
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Lista de usuarios"
         Top             =   240
         Width           =   4875
         _ExtentX        =   8599
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
            Text            =   "Usuario"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      ToolTipText     =   "Cierra la pantalla o cancela una operacion de agregado o modificacion"
      Top             =   4320
      Width           =   1305
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      ToolTipText     =   "Graba los datos de un usuario"
      Top             =   2280
      Width           =   1305
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      ToolTipText     =   "Permite borrar al usuario seleccionado"
      Top             =   1560
      Width           =   1305
   End
   Begin VB.CommandButton CmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      ToolTipText     =   "Permite modificar los datos del usuario seleccionado"
      Top             =   840
      Width           =   1305
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      ToolTipText     =   "Permite agregar los datos de un nuevo usuario"
      Top             =   120
      Width           =   1305
   End
   Begin VB.Frame fmeDatos 
      Caption         =   "Datos del usuario:"
      ForeColor       =   &H00FF0000&
      Height          =   2010
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   5055
      Begin VB.TextBox TxtRepeticion 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Contraseña del usuario (debe ser mayor a 6 digitos)(se recomienda que tenga letras y numeros)"
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox CheckPredeterminada 
         Caption         =   "Predeterminado"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         ToolTipText     =   "Establece que el usuario aparezca primero en las demas pantallas del sistema"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox TxtContraseña 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Contraseña del usuario (debe ser mayor a 6 digitos)(se recomienda que tenga letras y numeros)"
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox TxtUsuario 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Nombre del usuario"
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox ComboTipoUsuario 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Categoria del usuario"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Repeticion contraseña:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Contraseña:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmUsuariosAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***AQUI SE CARGAN LOS USUARIOS DEL SISTEMA
Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call RefrescarOpcionesSistema

Call CargarCombo2("tipousuario", ComboTipoUsuario)

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de Usuarios"
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
Private Sub cmdModificar_Click()
Call RefreshTimer

If Not VerificarSeleccionLista(lvUsuarios) Then Exit Sub
TipoEdicion = "M"
Call SetearEntorno
TxtUsuario.SetFocus

End Sub
Private Sub CargarLista()
'carga la lista de usuarios
Dim rec As rdoResultset
Dim sql As String
Dim Nitem As ListItem
On Error GoTo merror

lvUsuarios.ListItems.Clear

sql = "select * from usuarios ORDER BY Usuario"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lvUsuarios.ListItems.Add(, , Format(rec.rdoColumns("Idusuario"), "000"))
      Nitem.SubItems(1) = rec.rdoColumns("Usuario") & vbNullString
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error cargando la lista de usuarios"
End Sub
Private Sub CargarDatos()
'carga los campos de abajo con un usuario
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror
    
If Not VerificarSeleccionLista(lvUsuarios) Then Exit Sub
       
sql = "SELECT usuarios.usuario,usuarios.contraseña,usuarios.predeterminada,tipousuario.nombre " & _
      "FROM tipousuario inner join Usuarios " & _
      "on tipousuario.idtipousuario=usuarios.idtipousuario " & _
      "WHERE Idusuario=" & CLng(lvUsuarios.SelectedItem)
      
Set rec = cnSQL.OpenResultset(sql)
        
If Not rec.EOF Then
   TxtUsuario.Text = rec.rdoColumns("usuario") & vbNullString
   'pongo la clave desencriptada por si la modifican
   'para grabarla despues encriptada.
   TxtContraseña.Text = Desencriptar(rec.rdoColumns("contraseña")) & vbNullString
   TxtRepeticion.Text = TxtContraseña.Text & vbNullString
   ComboTipoUsuario.Text = rec.rdoColumns("nombre")
   If rec.rdoColumns("predeterminada") Then
      CheckPredeterminada.Value = 1
   Else
      CheckPredeterminada.Value = 0
   End If
End If

Exit Sub
merror:
tratarerrores "Error cargando datos de usuarios"
End Sub
Private Function CantidadAdministradores() As Long
'cuenta la cantidad de administradores del sistema
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

CantidadAdministradores = 0

sql = "select count(idtipousuario) as cantidad " & _
      "from usuarios " & _
      "where idtipousuario=1"
      
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("cantidad")) Then
      CantidadAdministradores = CLng(rec.rdoColumns("cantidad"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CantidadAdministradores"
End Function
Private Function ObtenerTipoUsuario(ByVal idusuario As Long) As Long
'obtiene el tipo de un usuario
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

ObtenerTipoUsuario = 0

sql = "select idusuario,idtipousuario " & _
      "from usuarios " & _
      "where idusuario=" & CLng(idusuario)

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idusuario")) Then
      ObtenerTipoUsuario = CLng(rec.rdoColumns("idtipousuario"))
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion ObtenerTipoUsuario"
End Function
Private Sub cmdborrar_Click()
Dim sql As String
Dim Idtipousuario As Long
On Error GoTo merror
Call RefreshTimer

If Not VerificarSeleccionLista(lvUsuarios) Then Exit Sub

Idtipousuario = ObtenerTipoUsuario(lvUsuarios.SelectedItem)

'si deseo borrar un administrador
If Idtipousuario = 1 Then
   If CantidadAdministradores() = 1 Then
      MsgE "No puede borrar al ultimo administrador (debe haber al menos un administrador)"
      Exit Sub
   End If
End If

If Not MsgP("¿Confirma el borrado del usuario seleccionado?") Then Exit Sub

'otras validaciones
If Not ExisteUsuario(lvUsuarios.SelectedItem) Then
   MsgE "El usuario no existe"
   Exit Sub
End If

'inicio transaccion
cnSQL.BeginTrans

sql = "delete from usuarios " & _
      "where idusuario=" & CLng(lvUsuarios.SelectedItem)

cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

MsgI "El usuario fue borrado"

lvUsuarios.SetFocus

Exit Sub
merror:
tratarerrores "Error borrando usuarios"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(TxtUsuario.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el nombre de usuario"
   TxtUsuario.SetFocus
   Exit Function
End If

If Trim(TxtContraseña.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar la contraseña del usuario"
   TxtContraseña.SetFocus
   Exit Function
End If

If Len(Trim(TxtContraseña.Text)) < 3 Then
   datosok = False
   MsgE "La contraseña debe tener 3 caracteres como minimo"
   TxtContraseña.SetFocus
   Exit Function
End If

If Trim(TxtRepeticion.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar la repeticion de la contraseña"
   TxtRepeticion.SetFocus
   Exit Function
End If

If Trim(TxtContraseña.Text) <> Trim(TxtRepeticion.Text) Then
   datosok = False
   MsgE "La contraseña y la repeticion deben ser iguales"
   TxtContraseña.SetFocus
   Exit Function
End If

If ComboTipoUsuario.Text = "" Then
   datosok = False
   MsgE "Debe seleccionar el tipo de usuario"
   ComboTipoUsuario.SetFocus
   Exit Function
End If

'reemplazo los caracteres de teclado invalidos
Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-UsuariosAbm"
End Function
Private Sub cmdGrabar_Click()
Dim sql As String
Dim Encriptada As String
Dim idusuario As Long
Dim Idtipousuario As Long
Dim IdTipoUsuarioActual As Long
Dim Mensaje As String
On Error GoTo merror
Call RefreshTimer

If Not datosok() Then Exit Sub

'el tipo de usuario seleccionado
Idtipousuario = CLng(ComboTipoUsuario.ItemData(ComboTipoUsuario.ListIndex))
Encriptada = Encriptar(TxtContraseña.Text)

If TipoEdicion = "N" Then
   
   If Not MsgP("¿Confirma el nuevo usuario?") Then Exit Sub
   
   idusuario = UltimoId("idusuario", "usuarios") + 1
   
   'otras validaciones
   If ExisteUsuario(idusuario) Then
      MsgE "El usuario ya existe"
      Exit Sub
   End If
   
   If Not ExisteTipoUsuario(Idtipousuario) Then
      MsgE "El tipo de usuario no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans
   
   'si es predeterminada saco el predeterminada al resto
   If CheckPredeterminada.Value = 1 Then
      sql = "update usuarios set predeterminada=0"
      cnSQL.Execute sql
   End If
           
   sql = "INSERT INTO usuarios (Idusuario, usuario, contraseña,idtipousuario,predeterminada) " & _
         "VALUES (" & CLng(idusuario) & ",'" & TxtUsuario.Text & "','" & CStr(Encriptada) & "'," & CLng(Idtipousuario) & "," & CheckPredeterminada.Value & ")"
        
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El usuario fue agregado"
   
   Call CargarLista
   Call CargarDatos

Else
   If Not MsgP("¿Confirma la modificacion del usuario seleccionado?") Then Exit Sub
   
   'este busca que tipo de usuario es el actual(no al que cambio)
   'para evitar que se elimine el ultimo administrador
   IdTipoUsuarioActual = ObtenerTipoUsuario(lvUsuarios.SelectedItem)
   
   'otras validaciones
   If Not ExisteUsuario(lvUsuarios.SelectedItem) Then
      MsgE "El usuario no existe"
      Exit Sub
   End If
   
   If Not ExisteTipoUsuario(Idtipousuario) Then
      MsgE "El tipo de usuario no existe"
      Exit Sub
   End If
   
   'si el tipo de usuario es administrador solo permito cambiar
   'el nombre de usuario la clave y predeterminada
   If IdTipoUsuarioActual = 1 And CantidadAdministradores() = 1 Then
      'inicio de transaccion
      cnSQL.BeginTrans
      
      'si es predeterminada saco el predeterminada al resto
      If CheckPredeterminada.Value = 1 Then
         sql = "update usuarios set predeterminada=0"
         cnSQL.Execute sql
      End If
   
      sql = "UPDATE usuarios SET usuario='" & CStr(TxtUsuario.Text) & _
            "',contraseña='" & CStr(Encriptada) & "',predeterminada=" & CheckPredeterminada.Value & _
            " WHERE Idusuario=" & CLng(lvUsuarios.SelectedItem)
      cnSQL.Execute sql
   
      'fin de transaccion
      cnSQL.CommitTrans
   Else
      'si es usuario comun o es administrador pero hay mas de uno
              
      'inicio de transaccion
      cnSQL.BeginTrans
      
      'si es predeterminada saco el predeterminada al resto
      If CheckPredeterminada.Value = 1 Then
         sql = "update usuarios set predeterminada=0"
         cnSQL.Execute sql
      End If
   
      sql = "UPDATE usuarios SET usuario='" & CStr(TxtUsuario.Text) & _
            "',contraseña='" & CStr(Encriptada) & "',idtipousuario=" & CLng(Idtipousuario) & ",predeterminada=" & CheckPredeterminada.Value & _
            " WHERE Idusuario=" & CLng(lvUsuarios.SelectedItem)
      cnSQL.Execute sql
   
      'fin de transaccion
      cnSQL.CommitTrans
   End If
    
   Mensaje = "El usuario fue modificado"
   lvUsuarios.SelectedItem.ListSubItems(1).Text = TxtUsuario.Text & vbNullString
End If

TipoEdicion = "C"
Call SetearEntorno

MsgI Mensaje

lvUsuarios.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando usuarios"
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
            TxtUsuario.SetFocus
            Call ColorBlanco(Me)
        End Select

Exit Sub
merror:
tratarerrores "Error seteando entorno-UsuariosAbm"
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
Private Sub TxtRepeticion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   TxtContraseña.SetFocus
End If
End Sub
Private Sub TxtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtContraseña.SetFocus
End If
End Sub
Private Sub TxtContraseña_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
   TxtRepeticion.SetFocus
End If
If KeyCode = vbKeyUp Then
   TxtUsuario.SetFocus
End If
End Sub
