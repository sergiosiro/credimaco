VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCobrosParciales 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobros parciales de la cuota seleccionada"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   HelpContextID   =   17
   Icon            =   "FrmCobrosParciales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdImprimirTodos 
      Caption         =   "Imprimir cobros parciales"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Imprime la lista de cobros parciales"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de cobros parciales:"
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6615
      Begin VB.TextBox TxtImporteParcial 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   8
         Tag             =   "N"
         ToolTipText     =   "Importe total cobrado parcialmente hasta la fecha"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox TxtCantParcial 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Tag             =   "N"
         ToolTipText     =   "Cantidad de cobros parciales de la cuota"
         Top             =   2520
         Width           =   855
      End
      Begin MSComctlLib.ListView lv 
         Height          =   2295
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Lista de cobros parciales"
         Top             =   240
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   4048
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº de cobro"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Credito Nº"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cod.Prestamo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cuota Nº"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Importe parcial"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Int. Historico"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "(*) Solo se pueden borrar cobros parciales en orden desde el mas reciente (ultimo) hacia atras."
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   6375
      End
      Begin VB.Label Label2 
         Caption         =   "Total Parcial:"
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cant.cobros:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdBorrarParcial 
      Caption         =   "&Borrar cobro parcial"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      ToolTipText     =   "Borra el cobro parcial seleccionado"
      Top             =   3600
      Width           =   2025
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      ToolTipText     =   "Cierra la pantalla"
      Top             =   3600
      Width           =   1905
   End
End
Attribute VB_Name = "FrmCobrosParciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AQUI SE CONSULTAN LOS COBROS PARCIALES DE CUOTAS
'VARIABLES GLOBALES QUE RECIBE ESTA PANTALLA DESDE OTRA
Public xnumcredito As Long
Public xcodprestamo As String
Public xnumcuota As Long
Public xidcliente As Long
Public xfactura As Long
Public ximporteactualizado As Currency
Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call CargarLista
Call SetearEntorno

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de cobros parciales"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
xnumcredito = 0
xcodprestamo = ""
xnumcuota = 0
xidcliente = 0
xfactura = 0
ximporteactualizado = 0
Unload Me
End Sub
Private Sub cmdborrarParcial_Click()
Call RefreshTimer
CmdBorrarParcial.Enabled = False
Call BorrarCobroParcial
CmdBorrarParcial.Enabled = True
End Sub
Private Sub BorrarCobroParcial()
'borra el cobro parcial seleccionado en la lista
Dim sql As String
Dim rec As rdoResultset
Dim ImporteParcial As Currency
Dim CobroParcial As Long
Dim FechaParcial As Date
Dim Item As String
Dim ImporteItem As Currency
Dim IdIngreso As Long
Dim Pos As Long
Dim Indicador As Long
On Error GoTo merror

If Not VerificarSeleccionLista(lv) Then Exit Sub

'verifico si estoy borrando un cobro anterior (no el ultimo)
'este es la cantidad de filas (o sea la ultima posicion)
Pos = lv.ListItems.Count

'si el que quiero borrar tiene una pos menor que el ultimo
If lv.SelectedItem.Index < Pos Then
   MsgE "Solo se puede borrar cobros parciales desde el ultimo hacia atras"
   Exit Sub
End If

If CuotaRefinanciada(xnumcredito, xnumcuota) Then
   MsgE "El comprobante esta refinanciado...no se puede anular el cobro parcial"
   Exit Sub
End If

If CuotaCobrada(xnumcredito, xnumcuota) Then
   MsgE "No se pueden borrar cobros parciales de cuotas ya cobradas"
   Exit Sub
End If

If Not MsgP("¿Desea borrar el cobro parcial seleccionado?") Then Exit Sub

IdIngreso = CLng(lv.SelectedItem)
FechaParcial = CDate(lv.SelectedItem.SubItems(4))
ImporteParcial = CCur(lv.SelectedItem.SubItems(5))

'otras validaciones
If Not ExisteCredito(xnumcredito) Then
   MsgE "El credito no existe"
   Exit Sub
End If

'inicio de transaccion
cnSQL.BeginTrans

'borro el cobro parcial de la tabla cobrosparciales
sql = "delete from ingresos " & _
      "Where idingreso='" & CLng(IdIngreso) & "'"
cnSQL.Execute sql

CobroParcial = 1
If ObtenerImporteParcialX(xnumcredito, xnumcuota) <= 0 Then
   CobroParcial = 0
End If

sql = "update cuotas set cobrosparciales='" & CLng(CobroParcial) & "',importeparcial=0 " & _
      "where idcredito='" & CLng(xnumcredito) & "' and numcuota='" & CLng(xnumcuota) & "'"
cnSQL.Execute sql

'si no hay mas cobros parciales de esa cuota le blanqueo los campos de importes
If CobroParcial = 0 Then
   Cad = ""
   sql = "update cuotas set pagofacil='" & CLng(Indicador) & "'," & _
         "rapipago='" & CLng(Indicador) & "',importedescuentos=0," & _
         "importerecargos=0,importemora=0,ivamora=0," & _
         "importeparcial=0 " & _
         "where idcredito='" & CLng(xnumcredito) & "' and numcuota='" & CLng(xnumcuota) & "'"
   cnSQL.Execute sql
End If

'fin de la transaccion
cnSQL.CommitTrans

'actualizo la lista de cobros parciales
Call CargarLista
Call SetearEntorno

If lv.ListItems.Count > 0 Then
   lv.SetFocus
End If

MsgI "El cobro parcial fue borrado"

Exit Sub
merror:
tratarerrores "Error borrando el cobro parcial seleccionado"
End Sub
Private Sub CargarLista()
'carga la lista de cobros parciales
Dim rec As rdoResultset
Dim sql As String
Dim Nitem As ListItem
Dim ImporteTotal As Currency
On Error GoTo merror

If xnumcredito = 0 Then Exit Sub

If xnumcuota = 0 Then Exit Sub

sql = "SELECT * " & _
      "FROM ingresos " & _
      "where idcredito='" & CLng(xnumcredito) & "' and numcuota='" & CLng(xnumcuota) & "' " & _
      "ORDER BY idingreso,idcredito,numcuota"

Set rec = cnSQL.OpenResultset(sql)

lv.ListItems.Clear

If Not rec.EOF Then
   Do While Not rec.EOF
      Set Nitem = lv.ListItems.Add(, , Format(rec.rdoColumns("Idingreso"), "000"))
      Nitem.SubItems(1) = Format(rec.rdoColumns("idcredito"), "000000") & vbNullString
      Nitem.SubItems(2) = rec.rdoColumns("codprestamo") & vbNullString
      Nitem.SubItems(3) = Format(rec.rdoColumns("numcuota"), "00") & vbNullString
      Nitem.SubItems(4) = rec.rdoColumns("fechacobro") & vbNullString
      Nitem.SubItems(5) = Format(rec.rdoColumns("importecobrado"), "0.00") & vbNullString
      
      ImporteTotal = CCur(ImporteTotal) + CCur(rec.rdoColumns("importecobrado"))
      rec.MoveNext
   Loop
End If
          
TxtCantParcial.Text = lv.ListItems.Count
TxtImporteParcial.Text = Format(ImporteTotal, "0.00")

Exit Sub
merror:
tratarerrores "Error cargando la lista de cobros parciales"
End Sub
Private Sub SetearEntorno()
'habilita o deshabilita los botones correspondientes
On Error GoTo merror
    
If lv.ListItems.Count > 0 Then
    If VG_ANULA Then
        CmdBorrarParcial.Enabled = True
    Else
        CmdBorrarParcial.Enabled = False
    End If
    CmdImprimirTodos.Enabled = True
Else
    CmdBorrarParcial.Enabled = False
    CmdImprimirTodos.Enabled = False
End If

Exit Sub
merror:
tratarerrores "Error seteando entorno-CobrosParciales"
End Sub
Private Sub CmdImprimirTodos_Click()
Call RefreshTimer
Call ImprimirCobrosParciales
End Sub
Private Sub ImprimirCobrosParciales()
'imprime la lista de cobros parciales
Dim sql As String
Dim rec As rdoResultset
Dim Archivo As String
Dim Mreporte As New ARCobrosParciales
Dim Titulo As String
On Error GoTo merror

'si imprimo todos los cobros parciales de la cuota
Titulo = "Lista de cobros parciales de la cuota Nº:" & Format(CStr(xnumcuota), "00") & " del prestamo Nº: " & CStr(xcodprestamo)

sql = "SELECT creditos.numcuotas,ingresos.* " & _
      "FROM creditos inner join ingresos on creditos.idcredito=ingresos.idcredito " & _
      "where ingresos.idcredito='" & CLng(xnumcredito) & "' and ingresos.numcuota='" & CLng(xnumcuota) & "' " & _
      "ORDER BY ingresos.idingreso,ingresos.idcredito,ingresos.numcuota"

Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   Mreporte.RDODataControl1.Resultset = rec
   Mreporte.Caption = "Imprimir cobros parciales de cuota"
   Mreporte.LabelTitulo = Titulo
   Mreporte.Show vbModal
Else
   MsgE "No hay cobros parciales para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo cobros parciales"
End Sub

