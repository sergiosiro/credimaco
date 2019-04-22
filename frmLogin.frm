VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresar al sistema de creditos"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   HelpContextID   =   1
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos del usuario:"
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtPassword 
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "Ingrese su contraseña respetando mayusculas y minusculas"
         Top             =   960
         Width           =   2895
      End
      Begin VB.ComboBox cboUsuarios 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Seleccione su nombre de usuario"
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Verifique siempre si el teclado esta en mayusculas o minusculas. La contraseña debe ser exacta."
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Contraseña:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Cancela el ingreso al sistema volviendo a Windows"
      Top             =   2040
      Width           =   1785
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Permite ingresar al sistema Credit-Click"
      Top             =   2040
      Width           =   1785
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PANTALLA DE LOGIN DEL SISTEMA

Private Sub Form_Load()
VG_IDTIPOUSUARIOLOGIN = 0
VG_IDUSUARIOLOGIN = 0
VG_USUARIOLOGIN = ""
VG_CLAVELOGIN = ""
VG_LOGIN = False
End Sub
Private Sub cboUsuarios_Click()
txtPassword.SetFocus
End Sub
Private Sub CmdCerrar_Click()
Unload Me
End Sub
Private Sub Form_Activate()
On Error GoTo merror

Me.Caption = App.Title
Call CargarComboUsuarios("usuarios", cboUsuarios)

If cboUsuarios.ListCount > 0 Then
   cboUsuarios.ListIndex = 0
End If

Exit Sub
merror:
tratarerrores "Error activando la pantalla de Login"
End Sub
Private Sub cmdOK_Click()
Dim sql As String
'Dim rec As rdoResultset
Dim rec As rdoResultset
Dim idusuario As Long
On Error GoTo merror
    
'Valido los datos de ingreso.
If Not datosok() Then Exit Sub
        

        
idusuario = CLng(cboUsuarios.ItemData(cboUsuarios.ListIndex))

'valido usuario y contraseña.
sql = "SELECT * from USUARIOS where IdUsuario=" & CLng(idusuario)
               
Set rec = cnSQL.OpenResultset(sql)
      
        
If rec.EOF Then
   MsgE "El usuario no existe"
   cboUsuarios.SetFocus
   Exit Sub
End If
        
'si existe, tomo los datos para variables globales y obtengo el id del usuario
VG_IDUSUARIOLOGIN = CLng(rec.rdoColumns("IdUsuario"))
VG_USUARIOLOGIN = rec.rdoColumns("Usuario") & vbNullString
VG_CLAVELOGIN = rec.rdoColumns("Contraseña") & vbNullString
VG_IDTIPOUSUARIOLOGIN = CLng(rec.rdoColumns("IdTipoUsuario"))

'verifico la contraseña
If Desencriptar(VG_CLAVELOGIN) <> Trim(txtPassword.Text) Then
   MsgE "La contraseña es incorrecta"
   txtPassword.Text = ""
   txtPassword.SetFocus
   Exit Sub
End If
    
VG_LOGIN = True

Unload Me

Exit Sub
merror:
tratarerrores "Error ingresando al sistema-Pantalla Login"
End Sub
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then Call cmdOK_Click
End Sub
Private Function datosok() As Boolean
On Error GoTo merror

datosok = True

If Trim(cboUsuarios.Text) = "" Then
   datosok = False
   MsgE "Debe seleccionar el nombre de usuario"
   cboUsuarios.SetFocus
   Exit Function
End If

If Trim(txtPassword.Text) = "" Then
   datosok = False
   MsgE "Debe ingresar la contraseña"
   txtPassword.SetFocus
   Exit Function
End If

Exit Function
merror:
tratarerrores "Error en funcion DatosOk-Login"
End Function
