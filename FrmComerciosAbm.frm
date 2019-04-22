VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmComerciosAbm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar comercios"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "FrmComerciosAbm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
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
      Caption         =   "Lista de comercios:"
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5415
      Begin MSComctlLib.ListView lvcomercios 
         Height          =   2415
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Lista de comercios"
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
      Height          =   1665
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   5415
      Begin VB.ComboBox ComboProvincias 
         Height          =   315
         ItemData        =   "FrmComerciosAbm.frx":57E2
         Left            =   1800
         List            =   "FrmComerciosAbm.frx":57E4
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Provincia donde se registro el credito"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox CheckPredeterminada 
         Caption         =   "Predeterminado"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Si marca la casilla, el comercio seleccionado sera la predeterminada en las demas pantallas del sistema"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Nombre del comercio"
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Provincia:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   705
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Nombre del comercio:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      ToolTipText     =   "Permite agregar los datos de un nuevo comercio"
      Top             =   960
      Width           =   1185
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      ToolTipText     =   "Permite modificar los datos del comercio seleccionado en la lista"
      Top             =   1560
      Width           =   1185
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      ToolTipText     =   "Permite borrar el comercio seleccionado en la lista"
      Top             =   2160
      Width           =   1185
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      ToolTipText     =   "Graba los datos de un comercio"
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
Attribute VB_Name = "FrmComerciosAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AQUI SE AGREGAN LOS COMERCIOS QUE LUEGO SE SELECCIONAN EN LA PANTALLA
'DE REGISTRAR CREDITOS Y REFINANCIAR. ES SOLO INFORMATIVO Y NO SE AGREGAN
'A LAS CONSULTAS SQL.(DE USO SIMILAR A VENDEDORES)

Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call CargarLista
Call CargarComboProvincias("provincias", ComboProvincias)
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno



Exit Sub
merror:
tratarerrores "Error cargando la pantalla de Comercios"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
If TipoEdicion = "C" Then
   Unload Me
Else
   TipoEdicion = "C"
   Call SetearEntorno
   Call CargarDatos
   lvcomercios.SetFocus
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
    
If Not VerificarSeleccionLista(lvcomercios) Then Exit Sub
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

If Not VerificarSeleccionLista(lvcomercios) Then Exit Sub
    

'***No se chequea la integridad referencial porque los comercios
'solo se usan como una lista en la pantalla de registrar creditos
If Not MsgP("¿Confirma el borrado del comercio seleccionado?") Then Exit Sub

'otras validaciones
If Not ExisteComercio(lvcomercios.SelectedItem) Then
   MsgE "El comercio no existe"
   Exit Sub
End If

'inicio de transaccion
cnSQL.BeginTrans

sql = "delete from comercios WHERE idcomercio=" & CLng(lvcomercios.SelectedItem)

cnSQL.Execute sql

'fin de transaccion
cnSQL.CommitTrans

Call CargarLista
Call CargarDatos
TipoEdicion = "C"
Call SetearEntorno

MsgI "El comercio fue borrado"
lvcomercios.SetFocus

Exit Sub
merror:
tratarerrores "Error borrando comercios"
End Sub
Private Sub cmdGrabar_Click()
'graba los registros nuevos y las modificaciones
Dim sql As String
Dim IdComercio As Long
Dim Mensaje As String
Dim IdProvincia As Long
On Error GoTo merror
Call RefreshTimer

If Not datosok() Then Exit Sub
    
If TipoEdicion = "N" Then
   If Not MsgP("¿Confirma el nuevo comercio?") Then Exit Sub
   
   IdComercio = UltimoId("idcomercio", "comercios") + 1
   
   'otras validaciones
   If ExisteComercio(IdComercio) Then
      MsgE "El comercio ya existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans

   'si es predeterminada saco el predeterminada al resto
   If CheckPredeterminada.Value = 1 Then
      sql = "update comercios set predeterminada=0"
      cnSQL.Execute sql
   End If
           
   IdProvincia = CLng(ComboProvincias.ItemData(ComboProvincias.ListIndex))
   sql = "INSERT INTO comercios (Idcomercio, nombre,predeterminada,Idprovincia) " & _
         "VALUES (" & CLng(IdComercio) & ",'" & CStr(TxtDescripcion.Text) & "'," & CheckPredeterminada.Value & "," & CLng(IdProvincia) & ")"
        
   cnSQL.Execute sql
   
   'fin de transaccion
   cnSQL.CommitTrans

   Mensaje = "El nuevo comercio fue agregado"
   
   Call CargarLista
   Call CargarDatos
Else
   If Not MsgP("¿Confirma la modificacion del comercio seleccionado?") Then Exit Sub
   
   'otras validaciones
   If Not ExisteComercio(lvcomercios.SelectedItem) Then
      MsgE "El comercio no existe"
      Exit Sub
   End If
   
   'inicio de transaccion
   cnSQL.BeginTrans

   If CheckPredeterminada.Value = 1 Then
      sql = "update comercios set predeterminada=0"
      cnSQL.Execute sql
   End If
   
   IdProvincia = CLng(ComboProvincias.ItemData(ComboProvincias.ListIndex))
   'estoy grabando una modificacion
   sql = "UPDATE comercios SET nombre='" & CStr(TxtDescripcion.Text) & _
         "',predeterminada=" & CheckPredeterminada.Value & _
         " ,Idprovincia = " & CLng(IdProvincia) & _
         " WHERE Idcomercio=" & CLng(lvcomercios.SelectedItem)
        
   cnSQL.Execute sql

   'fin de transaccion
   cnSQL.CommitTrans
        
   Mensaje = "El comercio fue modificado"
   
   lvcomercios.SelectedItem.ListSubItems(1).Text = TxtDescripcion.Text & vbNullString

End If

TipoEdicion = "C"
Call SetearEntorno

MsgI Mensaje

lvcomercios.SetFocus

Exit Sub
merror:
tratarerrores "Error actualizando comercios"
End Sub
Private Sub CargarLista()
'carga el listview con los comercios
Dim sql As String
Dim rec As rdoResultset
Dim Nitem As ListItem
On Error GoTo merror
    
lvcomercios.ListItems.Clear

sql = "SELECT * FROM comercios " & _
      "ORDER BY comercios.nombre"

Set rec = cnSQL.OpenResultset(sql)
    
If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lvcomercios.ListItems.Add(, , Format(rec.rdoColumns("Idcomercio"), "000"))
      Nitem.SubItems(1) = rec.rdoColumns("Nombre") & vbNullString
      rec.MoveNext
   Loop
End If

Exit Sub
merror:
tratarerrores "Error cargando la lista de comercios"
End Sub
Private Sub CargarDatos()
'carga el elemento seleccionado del listview en los cuadros de texto
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror
    
If Not VerificarSeleccionLista(lvcomercios) Then Exit Sub
        
sql = "SELECT * FROM comercios WHERE Idcomercio=" & CLng(lvcomercios.SelectedItem)
Set rec = cnSQL.OpenResultset(sql)
        
If Not rec.EOF Then
   TxtDescripcion.Text = rec.rdoColumns("nombre") & vbNullString
   For I = 0 To ComboProvincias.ListCount - 1
      If ComboProvincias.ItemData(I) = rec.rdoColumns("idprovincia") Then
       ComboProvincias.ListIndex = I
       Exit For
      End If
   Next I
   
   If rec.rdoColumns("predeterminada") Then
      CheckPredeterminada.Value = 1
   Else
      CheckPredeterminada.Value = 0
   End If
End If

Exit Sub
merror:
tratarerrores "Error cargando datos de comercios"
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(TxtDescripcion.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar el nombre del comercio"
   TxtDescripcion.SetFocus
   Exit Function
End If

If Trim(ComboProvincias.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar La provincia"
   ComboProvincias.SetFocus
   Exit Function
End If

Call ReemplazarComillas(Me)

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-ComerciosAbm"
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
            If lvcomercios.ListItems.Count > 0 Then
               CmdBorrar.Enabled = True
               cmdModificar.Enabled = True
            Else
               cmdModificar.Enabled = False
               CmdBorrar.Enabled = False
               Call LimpiarCampos(Me)
            End If
            cmdcerrar.Caption = "&Cerrar"
            lvcomercios.Enabled = True
            Call ColorCyan(Me)
        Case "M"
            fmeDatos.Enabled = True
            cmdGrabar.Enabled = True
            CmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            CmdBorrar.Enabled = False
            CmdRefrescar.Enabled = False
            cmdcerrar.Caption = "&Cancelar"
            lvcomercios.Enabled = False
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
            lvcomercios.Enabled = False
            TxtDescripcion.SetFocus
            Call ColorBlanco(Me)
    End Select

Exit Sub
merror:
tratarerrores "Error seteando entorno-ComerciosAbm"
End Sub
Private Sub lvcomercios_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'al clickear sobre el encabezado del listview los datos se ordenan
Dim Orden As Integer
  
If lvcomercios.ListItems.Count > 1 Then
   lvcomercios.SortKey = ColumnHeader.Index - 1
   Orden = lvcomercios.SortKey
   lvcomercios.SortOrder = Abs(Not lvcomercios.SortOrder = 1)
   lvcomercios.Sorted = True
End If

End Sub
Private Sub lvcomercios_ItemClick(ByVal Item As MSComctlLib.ListItem)
'al clickear sobre algun elemento del listview los datos se cargan en los text
Call CargarDatos
End Sub
Private Sub TxtDescripcion_LostFocus()
TxtDescripcion.Text = UCase(Trim(TxtDescripcion.Text))
End Sub
