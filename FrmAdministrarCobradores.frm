VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAdministrarCobradores 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrar Cobradores (liquidar comisiones)"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   HelpContextID   =   22
   Icon            =   "FrmAdministrarCobradores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdImprimirMovimientos 
      Caption         =   "I&mprimir liquidaciones"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      ToolTipText     =   "Imprime las cuotas de liquidacion"
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      ToolTipText     =   "Cierra la pantalla"
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton CmdPagar 
      Caption         =   "&Liquidar al cobrador"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Registra el pago al cobrador de las cuotas seleccionadas en la lista"
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar liquidacion"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "Borra las liquidaciones seleccionadas en la lista"
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones de busqueda:"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton CmdBuscar 
         Height          =   375
         Left            =   8520
         Picture         =   "FrmAdministrarCobradores.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Busca los cobros efectuados por los cobradores"
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox CheckLiquidadas 
         Caption         =   "Incluir cuotas ya liquidadas"
         Height          =   195
         Left            =   3120
         TabIndex        =   2
         ToolTipText     =   "Muestra las cuotas que ya fueron liquidadas al cobrador"
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox CheckCobrador 
         Caption         =   "Filtrar por cobrador"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Si marca la casilla podra seleccionar un cobrador"
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox ComboCobradores 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Lista de cobradores"
         Top             =   480
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   6960
         TabIndex        =   4
         ToolTipText     =   "Fecha final de la consulta"
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   55115777
         CurrentDate     =   39140
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   5520
         TabIndex        =   3
         ToolTipText     =   "Fecha inicial de la consulta"
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   55115777
         CurrentDate     =   39140
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   6960
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   5520
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cuotas cobradas al cliente para liquidar a los cobradores:"
      ForeColor       =   &H00FF0000&
      Height          =   4335
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   9255
      Begin MSComctlLib.ListView LvCuotas 
         Height          =   3675
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Lista de cuotas cobradas al cliente que se deben liquidar a los cobradores"
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6482
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod.Prestamo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Credito Nº"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cuota Nº"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Comprobante Nº"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cobrador"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fecha cobro"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Importe liquidacion"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Fecha Liquidacion"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.TextBox TxtCuotasALiquidar 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   12
         Tag             =   "N"
         ToolTipText     =   "Nº de cuotas seleccionadas"
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox TxtImporteALiquidar 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   13
         Tag             =   "N"
         ToolTipText     =   "Importe total seleccionado"
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox TxtContador2 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   11
         Tag             =   "N"
         ToolTipText     =   "Nº de cuotas cobradas"
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cuotas seleccionadas:"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Importe seleccionado:"
         Height          =   255
         Left            =   6600
         TabIndex        =   19
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Nº de cuotas:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3960
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmAdministrarCobradores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AQUI SE LIQUIDAN(PAGAN) LAS COMISIONES A COBRADORES DE CUOTAS
'PARA QUE UN COBRADOR TENGA COMISIONES HAY QUE SELECCIONARLO AL MOMENTO
'DE COBRAR CUOTAS

Private Sub Form_Load()
On Error GoTo merror
Call RefreshTimer

Call RefrescarOpcionesSistema

Call LimpiarCampos(Me)

'debe cargar todos los cobradores (no solo los activos)
Call CargarComboCobradores("cobradores", ComboCobradores, False, True)
ComboCobradores.ListIndex = -1

Exit Sub
merror:
tratarerrores "Error cargando la pantalla de administracion de cobradores"
End Sub
Private Sub CmdCerrar_Click()
Call RefreshTimer
Unload Me
End Sub
Private Sub cmdborrar_Click()
'borra las liquidaciones seleccionadas
Call RefreshTimer
CmdBorrar.Enabled = False
Call BorrarLiquidacion
CmdBorrar.Enabled = True
End Sub
Private Sub CmdPagar_Click()
'liquida las cuotas seleccionadas
Call RefreshTimer
CmdPagar.Enabled = False
Call LiquidarCuotas
TxtImporteALiquidar.Text = 0
CmdPagar.Enabled = True
End Sub
Private Sub CmdBuscar_Click()
Call ActualizarListas
End Sub
Private Sub CargarMovimientos()
'carga los cobros efectuados por los cobradores
Dim rec As rdoResultset
Dim Nitem As ListItem
Dim I As Long
On Error GoTo merror

lvcuotas.ListItems.Clear

TxtContador2.Text = 0
TxtCuotasALiquidar.Text = 0
TxtImporteALiquidar.Text = 0

Set rec = CargarRecMovimientos()

I = 1
Do While Not rec.EOF
   Set Nitem = lvcuotas.ListItems.Add(, , rec.rdoColumns("codprestamo"))
   Nitem.SubItems(1) = Format(rec.rdoColumns("idcredito"), "0000000") & vbNullString
   Nitem.SubItems(2) = Format(rec.rdoColumns("numcuota"), "000") & vbNullString
   
   Nitem.SubItems(3) = Format(rec.rdoColumns("numfactura"), "0000000") & vbNullString
   Nitem.SubItems(4) = rec.rdoColumns("cobrador") & vbNullString
   Nitem.SubItems(5) = rec.rdoColumns("fecha") & vbNullString
   Nitem.SubItems(6) = Format(rec.rdoColumns("importecobrador"), "0.00") & vbNullString
   Nitem.SubItems(7) = rec.rdoColumns("fechapago") & vbNullString
   
   'si no esta liquidada la pongo en azul
   If IsNull(rec.rdoColumns("fechapago")) Then
      lvcuotas.ListItems.Item(I).ForeColor = vbBlue
      lvcuotas.ListItems.Item(I).ListSubItems(1).ForeColor = vbBlue
      lvcuotas.ListItems.Item(I).ListSubItems(2).ForeColor = vbBlue
      lvcuotas.ListItems.Item(I).ListSubItems(3).ForeColor = vbBlue
      lvcuotas.ListItems.Item(I).ListSubItems(4).ForeColor = vbBlue
   Else
      'si esta pagada la pongo en verde
      lvcuotas.ListItems.Item(I).ForeColor = &H8000&
      lvcuotas.ListItems.Item(I).ListSubItems(1).ForeColor = &H8000&
      lvcuotas.ListItems.Item(I).ListSubItems(2).ForeColor = &H8000&
      lvcuotas.ListItems.Item(I).ListSubItems(3).ForeColor = &H8000&
      lvcuotas.ListItems.Item(I).ListSubItems(4).ForeColor = &H8000&
   End If
   'en bold
   lvcuotas.ListItems.Item(I).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(1).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(2).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(3).Bold = True
   lvcuotas.ListItems.Item(I).ListSubItems(4).Bold = True
   
   I = I + 1
   rec.MoveNext
Loop

TxtContador2.Text = lvcuotas.ListItems.Count

Exit Sub
merror:
tratarerrores "Error cargando la lista de liquidacion de cobrador"
End Sub
Private Function CargarRecMovimientos() As rdoResultset
'carga las cuotas de liquidacion a cobradores
Dim sql As String
Dim Condicion1 As String
Dim Condicion2 As String
Dim IdCobrador As Long
On Error GoTo merror
  
Condicion1 = " and 1=1"
If CheckCobrador.Value = 1 And ComboCobradores.Text <> "" Then
   'carga para un cobrador
   IdCobrador = CLng(ComboCobradores.ItemData(ComboCobradores.ListIndex))
   Condicion1 = " and cobradores.idcobrador='" & CLng(IdCobrador) & "'"
End If

'si solo las no liquidadas
Condicion2 = " and 1=1"
If CheckLiquidadas.Value = 0 Then
   Condicion2 = " and cobradorespagos.fechapago is Null"
End If

sql = "select cobradores.idcobrador,cobradores.apellido + ' '+ cobradores.nombre as cobrador," & _
      "cobradorespagos.idcredito,cobradorespagos.numcuota,cobradorespagos.codprestamo,cobradorespagos.numfactura,cobradorespagos.importecobrador,cobradorespagos.fecha," & _
      "cobradorespagos.fechapago " & _
      "from cobradores inner join cobradorespagos on cobradores.idcobrador=cobradorespagos.idcobrador " & _
      "where cobradorespagos.fecha>='" & ConvertirFechaSql(CDate(DTPicker1.Value), "DD/MM/YYYY") & "' and cobradorespagos.fecha<='" & ConvertirFechaSql(CDate(DTPicker2.Value), "DD/MM/YYYY") & "'" & _
      Condicion1 & Condicion2 & _
      " order by cobradores.apellido,cobradores.nombre,cobradorespagos.idcredito,cobradorespagos.numfactura"

Set CargarRecMovimientos = cnSQL.OpenResultset(sql)

Exit Function
merror:
tratarerrores "Error cargando el registro de liquidaciones de cobrador"
End Function
Private Function CuotaLiquidada(ByVal IdCredito As Long, ByVal NumFactura As Long) As Boolean
'verifica si la cuota seleccionada ya esta liquidada al cobrador
Dim sql As String
Dim rec As rdoResultset
On Error GoTo merror

CuotaLiquidada = False

sql = "select idcobrador,fechapago " & _
      "from cobradorespagos " & _
      "where idcredito='" & CLng(IdCredito) & "' and numfactura='" & CLng(NumFactura) & "'"
            
Set rec = cnSQL.OpenResultset(sql)

If Not rec.EOF Then
   If Not IsNull(rec.rdoColumns("idcobrador")) Then
      If Not IsNull(rec.rdoColumns("fechapago")) Then
         CuotaLiquidada = True
      End If
   End If
End If

Exit Function
merror:
tratarerrores "Error en funcion CuotaLiquidada-AdmCobradores"
End Function
Private Sub LiquidarCuotas()
'registra la liquidacion de las cuotas seleccionadas al cobrador correspondiente
Dim sql As String
Dim IdCredito As Long
Dim NumCuota As Long
Dim NumFactura As Long
Dim HuboLiquidacion As Boolean
Dim I As Long
On Error GoTo merror

If (lvcuotas.ListItems.Count = 0) Then
   MsgE "No tiene cuotas para liquidar al cobrador"
   Exit Sub
End If

If Not HayFilasChequeadas(lvcuotas) Then
   MsgE "Debe seleccionar las cuotas a liquidar al cobrador"
   lvcuotas.SetFocus
   Exit Sub
End If

If Not MsgP("¿Confirma la liquidacion al cobrador de las cuotas seleccionadas?") Then Exit Sub

HuboLiquidacion = False

'inicio de la transaccion
cnSQL.BeginTrans

'voy liquidando las cuotas seleccionadas si no estan ya liquidadas
For I = 1 To lvcuotas.ListItems.Count()
   If lvcuotas.ListItems.Item(I).Checked Then
      IdCredito = CLng(lvcuotas.ListItems.Item(I).SubItems(1))
      NumCuota = CLng(lvcuotas.ListItems.Item(I).SubItems(2))
      NumFactura = CLng(lvcuotas.ListItems.Item(I).SubItems(3))
      If Not CuotaLiquidada(IdCredito, NumFactura) Then
         sql = "update cobradorespagos set fechapago= '" & ConvertirFechaSql(Mid$(Now, 1, 10), "DD/MM/YYYY") & "'" & _
               "where idcredito='" & CLng(IdCredito) & "' and numfactura='" & CLng(NumFactura) & "'"
         cnSQL.Execute sql
         HuboLiquidacion = True
      End If
   End If
Next I

'fin de la transaccion
cnSQL.CommitTrans

Call ActualizarListas

If HuboLiquidacion Then
   MsgI "Se registro la liquidacion exitosamente"
Else
   MsgI "No hubo liquidacion...(no se pueden liquidar cuotas liquidadas anteriormente)"
End If

Exit Sub
merror:
tratarerrores "Error en procedimiento LiquidarCuotas de cobradores"
End Sub
Private Sub BorrarLiquidacion()
'borra liquidaciones seleccionadas si estan liquidadas al cobrador
Dim sql As String
Dim IdCredito As Long
Dim NumCuota As Long
Dim NumFactura As Long
Dim HuboBorrado As Boolean
Dim I As Long
On Error GoTo merror
   
If lvcuotas.ListItems.Count = 0 Then
   MsgE "No hay liquidaciones para borrar"
   Exit Sub
End If

If Not HayFilasChequeadas(lvcuotas) Then
   MsgE "Debe seleccionar cuotas para borrar (ya liquidadas a los cobradores)"
   lvcuotas.SetFocus
   Exit Sub
End If

If Not MsgP("¿Confirma el borrado de las liquidaciones seleccionadas?") Then Exit Sub

HuboBorrado = False

'inicio de la transaccion
cnSQL.BeginTrans
   
'voy borrando las liquidaciones seleccionadas
For I = 1 To lvcuotas.ListItems.Count()
    'si esta seleccionada
    If lvcuotas.ListItems.Item(I).Checked Then
       IdCredito = CLng(lvcuotas.ListItems.Item(I).SubItems(1))
       NumCuota = CLng(lvcuotas.ListItems.Item(I).SubItems(2))
       NumFactura = CLng(lvcuotas.ListItems.Item(I).SubItems(3))
       If CuotaLiquidada(IdCredito, NumFactura) Then
          'borro la operacion
          sql = "delete from cobradorespagos " & _
                "WHERE idcredito=clng('" & IdCredito & "') and numfactura=clng('" & NumFactura & "')"
          cnSQL.Execute sql
          HuboBorrado = True
       End If
    End If
Next I

'fin de la transaccion
cnSQL.CommitTrans

Call ActualizarListas

If HuboBorrado Then
   MsgI ("Las liquidaciones seleccionadas fueron borradas!")
Else
   MsgE "No hubo borrado (no se pueden borrar cuotas aun no liquidadas a los cobradores)"
End If

Exit Sub
merror:
tratarerrores "Error borrando liquidaciones de cobradores"
End Sub
Private Sub cmdimprimirMovimientos_Click()
Call RefreshTimer
CmdImprimirMovimientos.Enabled = False

Call ImprimirMovimientos

CmdImprimirMovimientos.Enabled = True
End Sub
Private Sub ImprimirMovimientos()
'imprime los movimientos
Dim rec As rdoResultset
Dim Archivo As String
Dim Texto As String
Dim Mreporte As New ARMovimientosCobradores
On Error GoTo merror

If lvcuotas.ListItems.Count() = 0 Then Exit Sub

Set rec = CargarRecMovimientos()

If Not rec.EOF Then
   Mreporte.RDODataControl1.Resultset = rec
   Mreporte.Caption = "Imprimir liquidaciones de cobradores"
   
   Texto = "Lista de liquidacion de cobradores en el periodo:" & CStr(DTPicker1.Value) & " al " & CStr(DTPicker2.Value)
   
   If CheckCobrador.Value = 1 Then
      If ComboCobradores.Text <> "" Then
         Texto = "Lista de liquidacion del cobrador " & ComboCobradores.Text & " en el periodo:" & CStr(DTPicker1.Value) & " al " & CStr(DTPicker2.Value)
      End If
   End If
   
   Mreporte.LabelTitulo.Caption = Texto
   
   'si imprimo los datos de empresa
   Mreporte.LabelEmpresa = VG_EMPRESA & vbNullString
   Mreporte.Show vbModal
Else
   MsgE "No hay liquidaciones de cobrador para imprimir"
End If

Exit Sub
merror:
tratarerrores "Error imprimiendo liquidaciones de cobradores"
End Sub
Private Sub lvcuotas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Orden As Long
  
If lvcuotas.ListItems.Count > 1 Then
   lvcuotas.SortKey = ColumnHeader.Index - 1
   Orden = lvcuotas.SortKey
   lvcuotas.SortOrder = Abs(Not lvcuotas.SortOrder = 1)
   lvcuotas.Sorted = True
End If

End Sub
Private Sub ActualizarListas()
Call CargarMovimientos
Call SetearEntorno
End Sub
Private Sub SetearEntorno()
If lvcuotas.ListItems.Count = 0 Then
   CmdBorrar.Enabled = False
   CmdPagar.Enabled = False
   CmdImprimirMovimientos.Enabled = False
Else
   CmdBorrar.Enabled = True
   CmdPagar.Enabled = True
   CmdImprimirMovimientos.Enabled = True
End If
End Sub
Private Sub Combocobradores_Click()
'si cambio el cobrador actualizo la lista
Call ActualizarListas
End Sub
Private Sub DTPicker1_Change()
'si cambia la fecha inicial del rango
Call ActualizarListas
End Sub
Private Sub DTPicker2_Change()
'si cambia la fecha final del rango
Call ActualizarListas
End Sub
Private Sub Checkcobrador_Click()
'si selecciono un cobrador
If CheckCobrador.Value = 1 Then
   ComboCobradores.Enabled = True
   ComboCobradores.BackColor = vbWhite
Else
   ComboCobradores.Enabled = False
   ComboCobradores.ListIndex = -1
   ComboCobradores.BackColor = &HFFFFC0
End If
End Sub
Private Sub CheckLiquidadas_Click()
Call ActualizarListas
End Sub
Private Sub LvCuotas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'se ejecuta al marcar/desmarcar una fila de la lista de cuotas
On Error GoTo merror

'si la fila esta tildada
If lvcuotas.ListItems.Item(Item.Index).Checked Then
   'importe original de la fila
   TxtImporteALiquidar.Text = CCur(TxtImporteALiquidar.Text) + CCur(lvcuotas.ListItems.Item(Item.Index).SubItems(6))
   TxtCuotasALiquidar.Text = CLng(TxtCuotasALiquidar.Text) + 1
Else
   'importe original de la fila
   TxtImporteALiquidar.Text = CCur(TxtImporteALiquidar.Text) - CCur(lvcuotas.ListItems.Item(Item.Index).SubItems(6))
   TxtCuotasALiquidar.Text = CLng(TxtCuotasALiquidar.Text) - 1
End If

TxtImporteALiquidar.Text = Format(TxtImporteALiquidar.Text, "0.00")

Exit Sub
merror:
tratarerrores "Error seleccionando cuota a liquidar-AdmCobradores"
End Sub
