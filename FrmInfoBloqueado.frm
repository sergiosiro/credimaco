VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfoBloqueado 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bloquear Crédito"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   Icon            =   "FrmInfoBloqueado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmTitulo 
      Caption         =   "Crédito a Bloquear:"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9255
      Begin VB.Label lblIdCredito 
         Caption         =   "IdCredito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lblCapital 
         Caption         =   "Capital"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   21
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblCuotas 
         Caption         =   "Cuotas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblCliente 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label lblCodPrestamo 
         Caption         =   "Cód. Préstamo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Crédito:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Cuotas:"
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Capital:"
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Cód. Préstamo:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Bloquear Crédito"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Importa los cobros de los archivos seleccionados"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del Bloqueo: "
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   9255
      Begin VB.ComboBox cboSituacion 
         Height          =   315
         ItemData        =   "FrmInfoBloqueado.frx":0442
         Left            =   4200
         List            =   "FrmInfoBloqueado.frx":047F
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Seleccione el tipo de consulta"
         Top             =   840
         Width           =   4455
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   975
         Left            =   1440
         TabIndex        =   3
         Top             =   1320
         Width           =   7695
      End
      Begin VB.ComboBox cboEstudio 
         Height          =   315
         ItemData        =   "FrmInfoBloqueado.frx":05F0
         Left            =   4080
         List            =   "FrmInfoBloqueado.frx":05F2
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Seleccione el tipo de consulta"
         Top             =   360
         Width           =   4575
      End
      Begin VB.ComboBox cboEstado 
         Height          =   315
         ItemData        =   "FrmInfoBloqueado.frx":05F4
         Left            =   840
         List            =   "FrmInfoBloqueado.frx":0601
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione el tipo de consulta"
         Top             =   360
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpFechaEnvio 
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         ToolTipText     =   "Fecha de nacimiento del cliente"
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   16711681
         CurrentDate     =   39018
      End
      Begin VB.Label Label10 
         Caption         =   "Situación:"
         Height          =   255
         Left            =   3360
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Estudio:"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Estado:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha de Envío:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInfoBloqueado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboEstado_Click()
    If cboEstado.ListIndex = 0 Then
        If cboEstudio.ListIndex = -1 Then
            cboEstudio.ListIndex = 0
        End If
        cboEstudio.Enabled = True
        dtpFechaEnvio.Enabled = True
    Else
        cboEstudio.Enabled = False
        dtpFechaEnvio.Enabled = False
    End If
End Sub

Private Sub CmdCancelar_Click()
    If VG_ADMCREDITOS Then
        If Me.Tag = "Bloqueo" Then
            If MsgP("¿Desea salir sin realizar el bloqueo?") Then
                Unload Me
            End If
        Else
            If MsgP("¿Desea salir sin actualizar la información del bloqueo?") Then
                Unload Me
            End If
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub CmdGrabar_Click()
Dim sql As String
Dim IdCredito As Long
Dim nSecuencia As Long
Dim cEstado As String
Dim IdEstudio As Long
Dim cSituacion As String
On Error GoTo merror


IdCredito = CLng(Trim$(lblIdCredito.Caption))

If cboEstado.ListIndex = -1 Then
   MsgE "Debe seleccionar un estado"
   cboEstado.SetFocus
   Exit Sub
End If

If cboEstudio.Enabled Then
    IdEstudio = CLng(cboEstudio.ItemData(cboEstudio.ListIndex))
End If
cSituacion = Trim$(cboSituacion.Text)
If Me.Tag = "Bloqueo" Then
    If Not MsgP("¿Confirma el bloqueo del credito seleccionado?") Then Exit Sub
Else
    If Not MsgP("¿Confirma la actualización de los datos del crédito bloqueado?") Then Exit Sub
End If
  
'otras validaciones
If Not ExisteCredito(IdCredito) Then
   MsgE "El credito no existe"
   Exit Sub
End If
   
'inicio de la transaccion
cnSQL.BeginTrans
   
If Me.Tag = "Bloqueo" Then
    sql = "UPDATE creditos SET creditos.fechabloqueo= '" & ConvertirFechaSql(Mid$(Now, 1, 10), "DD/MM/YYYY") & "'" & _
          " WHERE (creditos.idcredito)=" & IdCredito
    cnSQL.Execute sql
End If

nSecuencia = 1
sql = "SELECT MAX(Secuencia) as secuencia FROM CREDITOSBLOQUEADOS WHERE IdCredito = " & IdCredito
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
    If Not IsNull(rec.rdoColumns("secuencia")) Then
        nSecuencia = rec.rdoColumns("secuencia") + 1
    End If
End If

Select Case cboEstado.Text
Case "Enviar a Estudio"
    cEstado = "EE"
Case "No Enviar"
    cEstado = "NE"
Case "Pendiente de Envio"
    cEstado = "PE"
End Select

sql = "INSERT INTO CREDITOSBLOQUEADOS (IdCredito,Secuencia,FechaEstado,Estado,Observaciones, Situacion) " & _
      "VALUES (" & IdCredito & "," & nSecuencia & ",CURRENT_TIMESTAMP,'" & cEstado & "','" & Trim$(txtObservaciones.Text) & "','" & cSituacion & "')"
cnSQL.Execute sql

If cboEstudio.Enabled Then
    sql = "UPDATE CREDITOSBLOQUEADOS SET IdEstudio = " & IdEstudio & ", FechaEnvio = '" & ConvertirFechaSql(dtpFechaEnvio, "DD/MM/YYYY") & "' " & _
          "WHERE IdCredito = " & IdCredito & " AND Secuencia = " & nSecuencia
    cnSQL.Execute sql
End If

'fin de la transaccion
cnSQL.CommitTrans
   
If Me.Tag = "Bloqueo" Then
    MsgI "El credito fue bloqueado exitosamente"
Else
    MsgI "Los datos del crédito bloqueado se han actualizado exitosamente"
End If
Unload Me

Exit Sub
merror:
tratarerrores "Error en el proceso de bloqueo o actualización de datos: " & Err.Number & "(" & Err.Description & ")"

End Sub

Private Sub Form_Load()
    Call CargarCombo2("estudios", cboEstudio)
    If cboEstudio.ListCount = 0 Then
        MsgE "Error! Debe cargar estudios!"
        End
    End If
    cboEstado.ListIndex = 0
    cboEstudio.ListIndex = 0
    cboSituacion.ListIndex = 0
    If VG_ADMCREDITOS Then
        CmdGrabar.Enabled = True
    Else
        CmdGrabar.Enabled = False
    End If
End Sub
