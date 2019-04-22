VERSION 5.00
Begin VB.Form FrmAcercaDe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de"
   ClientHeight    =   4530
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7110
   ClipControls    =   0   'False
   Icon            =   "FrmAcercaDe.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3126.686
   ScaleMode       =   0  'User
   ScaleWidth      =   6676.658
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdBorrarCreditos 
      BackColor       =   &H00C0E0FF&
      Caption         =   "1) Borrar creditos masivos(finalizados hasta 2009)"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Borra todos los creditos finalizados viejos hasta el 2009 inclusive"
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CommandButton CmdBorrarClientes 
      BackColor       =   &H00C0C0FF&
      Caption         =   "2) Borrar clientes masivos(sin creditos)"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Borra clientes que no tengan ningun tipo de creditos"
      Top             =   3360
      Width           =   3615
   End
   Begin VB.CommandButton CmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Default         =   -1  'True
      Height          =   345
      Left            =   5160
      TabIndex        =   0
      Top             =   4080
      Width           =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "Aqui estaba el logotipo de Software Bariloche"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label3 
      Caption         =   "Borra clientes que no tienen ningun credito en el sistema (ni vigentes, ni bloqueados, ni finalizados)"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Borra creditos finalizados hasta 2009 inclusive sin cuotas pendientes, ni refinanciadas."
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label LabelNumRegistracion 
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label LabelContacto 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6855
   End
End
Attribute VB_Name = "FrmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'nuevo
Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Me.Caption = "Acerca de " & App.Title & " (Version Registrada)"
LabelContacto.Caption = App.CompanyName
LabelNumRegistracion.Caption = "Nº de licencia:S3-1016"

Exit Sub
merror:
tratarerrores "Error cargando la pantalla AcercaDe"
End Sub
Private Sub CmdCerrar_Click()
Unload Me
End Sub
Private Sub CmdBorrarClientes_Click()
CmdBorrarClientes.Enabled = False
Me.MousePointer = vbHourglass
Call BorrarClientesViejos
Me.MousePointer = vbDefault
CmdBorrarClientes.Enabled = True
End Sub
Private Sub CmdBorrarCreditos_Click()
CmdBorrarCreditos.Enabled = False
Me.MousePointer = vbHourglass
Call BorrarCreditosViejos
Me.MousePointer = vbDefault
CmdBorrarCreditos.Enabled = True
End Sub
Private Function PuedoBorrarCliente2(ByVal IdCliente As Long) As Boolean
Dim rec As rdoResultset
Dim sql As String
On Error GoTo merror

PuedoBorrarCliente2 = True

'verifico en creditos...
sql = "select idcliente from creditos " & _
      "where idcliente='" & CLng(IdCliente) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcliente")) Then
      PuedoBorrarCliente2 = False
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion PuedoBorrarClientesMasivos"
End Function
Private Function EsCreditoViejo(ByVal IdCredito As Long) As Boolean
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

EsCreditoViejo = False

sql = "select fechafinalizacion " & _
      "from creditos " & _
      "where idcredito='" & CLng(IdCredito) & "'"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Year(rec.rdoColumns("fechafinalizacion")) <= 2009 Then
      EsCreditoViejo = True
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion EsCreditoViejo"
End Function
Private Sub BorrarCreditosViejos()
'no permito borrar creditos vigentes ni bloqueados..
'deben estar finalizados sin cuotas pendientes ni refinanciadas
Dim sql As String
Dim rec As rdoResultset
Dim IdCredito As Long
Dim HuboBorrados As Boolean
Dim cont As Long
On Error GoTo merror

If Not MsgP("¿Confirma el borrado masivo de creditos finalizados viejos?") Then Exit Sub

'para mayor seguridad saco solo los creditos finalizados
'(que tienen fecha de finalizacion)
'para achicar el rango trae solo creados hasta 2009 inclusive
'descarta los finalizados creados en 2010 y 2011
sql = "select idcredito " & _
      "from creditos " & _
      "where fechafinalizacion is not Null and year(fechacredito)<=2009 " & _
      "order by idcredito"
Set rec = cnSQL.OpenResultset(sql)

HuboBorrados = False
cont = 0

If Not rec.EOF Then
   'inicio de la transaccion
   cnSQL.BeginTrans
   
   Do While Not rec.EOF
      IdCredito = rec.rdoColumns("idcredito")
      If EsCreditoViejo(IdCredito) Then
         If CreditoFinalizado(IdCredito) Then
            'si no tiene cuotas impagas puras(que no tienen fecha de cobro)
            'y que no estan refin ni son comodin
            If CuotasImpagas(IdCredito) = 0 Then
               If Not CreditoTieneRefinanciadas(IdCredito) Then
                  'borro el credito
                  sql = "delete from creditos WHERE (creditos.idcredito)='" & CLng(IdCredito) & "'"
                  cnSQL.Execute sql
                  
                  'borro sus cuotas
                  sql = "delete from cuotas WHERE (cuotas.idcredito)='" & CLng(IdCredito) & "'"
                  cnSQL.Execute sql
            
                  'borro las liquidaciones a cobradores de ese credito
                  sql = "delete from cobradorespagos WHERE (cobradorespagos.idcredito)='" & CLng(IdCredito) & "'"
                  cnSQL.Execute sql

                  'borro sus excedentes
                  sql = "delete from excedentesclientes WHERE (excedentesclientes.idcredito)='" & CLng(IdCredito) & "'"
                  cnSQL.Execute sql
            
                  'borro SUS INGRESOS
                  sql = "delete from ingresos WHERE (ingresos.idcredito)='" & CLng(IdCredito) & "'"
                  cnSQL.Execute sql

                  'pagofacil
                  sql = "delete from pagofacil WHERE (pagofacil.idcredito)='" & CLng(IdCredito) & "'"
                  cnSQL.Execute sql
            
                  'pagofacil2
                  sql = "delete from pagofacil2 WHERE (pagofacil2.idcredito)='" & CLng(IdCredito) & "'"
                  cnSQL.Execute sql
                  
                  'pagofaciltemp
                  sql = "delete from pagofaciltemp WHERE (pagofaciltemp.idcredito)='" & CLng(IdCredito) & "'"
                  cnSQL.Execute sql
                  
                  'rapipago
                  sql = "delete from rapipago WHERE (rapipago.idcredito)='" & CLng(IdCredito) & "'"
                  cnSQL.Execute sql
                        
                  'templistadeudores
                  sql = "delete from templistadeudores WHERE (templistadeudores.idcredito)='" & CLng(IdCredito) & "'"
                  cnSQL.Execute sql
                  
                  'temporaldeudores
                  sql = "delete from temporaldeudores WHERE (temporaldeudores.idcredito)='" & CLng(IdCredito) & "'"
                  cnSQL.Execute sql
                              
                  HuboBorrados = True
                  cont = cont + 1
               End If
            End If

         End If
      End If
      rec.MoveNext
   Loop
   'fin de la transaccion
   cnSQL.CommitTrans
End If
   

If HuboBorrados Then
   MsgI "fueron borrados " & CStr(cont) & " creditos finalizados"
Else
   MsgE "No hubo borrado de creditos"
End If

Exit Sub
merror:
tratarerrores "Error borrando creditos finalizados masivos"
End Sub
Private Sub BorrarClientesViejos()
Dim sql As String
Dim rec As rdoResultset
Dim IdCliente As Long
Dim HuboBorrados As Boolean
Dim cont As Long
On Error GoTo merror
         
If Not MsgP("¿Confirma el borrado masivo de clientes?") Then Exit Sub

'obtengo todos los clientes
sql = "select idcliente from clientes order by idcliente"
Set rec = cnSQL.OpenResultset(sql)

cont = 0
HuboBorrados = False

If Not rec.EOF Then
   'inicio de transaccion
   cnSQL.BeginTrans

   Do While Not rec.EOF
      IdCliente = rec.rdoColumns("idcliente")
      'si no tiene creditos asociados
      If PuedoBorrarCliente2(IdCliente) Then
         'borro a ese cliente
         sql = "delete from clientes " & _
               "WHERE idcliente='" & CLng(IdCliente) & "'"
         cnSQL.Execute sql
         
         'borro por las dudas en creditos tambien (aunque si tiene no borraria
         'al cliente)para evitar que queden descolgados posibles registros
         sql = "delete from creditos " & _
               "WHERE creditos.idcliente='" & CLng(IdCliente) & "'"
         cnSQL.Execute sql
         
         'borro excedentes
         sql = "delete from excedentesclientes " & _
               "WHERE idcliente='" & CLng(IdCliente) & "'"
         cnSQL.Execute sql
         
         'borro en otras tablas
         'pagofacil
         sql = "delete from pagofacil " & _
               "WHERE idcliente='" & CLng(IdCliente) & "'"
         cnSQL.Execute sql
         
         'pagofaciltemp
         sql = "delete from pagofaciltemp " & _
               "WHERE numcliente='" & CLng(IdCliente) & "'"
         cnSQL.Execute sql
         
         'rapipago
         sql = "delete from rapipago " & _
               "WHERE idcliente='" & CLng(IdCliente) & "'"
         cnSQL.Execute sql
         
         'rapipagotemp
         sql = "delete from rapipagotemp " & _
               "WHERE numcliente='" & CLng(IdCliente) & "'"
         cnSQL.Execute sql
         
         'templistadeudores
         sql = "delete from templistadeudores " & _
               "WHERE idcliente='" & CLng(IdCliente) & "'"
         cnSQL.Execute sql
         
         'temporaldeudores
         sql = "delete from temporaldeudores " & _
               "WHERE idcliente='" & CLng(IdCliente) & "'"
         cnSQL.Execute sql
         
         HuboBorrados = True
         cont = cont + 1
      End If
      rec.MoveNext
   Loop
   'fin de transaccion
   cnSQL.CommitTrans
End If

If HuboBorrados Then
   MsgI "Fueron borrados " & CStr(cont) & " clientes en forma exitosa"
Else
   MsgE "No hubo borrados de clientes"
End If

Exit Sub
merror:
tratarerrores "Error borrando clientes en forma masiva"
End Sub
