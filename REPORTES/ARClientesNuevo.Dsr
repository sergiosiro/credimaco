VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARClientesNuevo 
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "ARClientesNuevo.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARClientesNuevo.dsx":0442
End
Attribute VB_Name = "ARClientesNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ActiveReport_ReportEnd()

    Unload subDatosFactura.object
    Set subDatosFactura.object = Nothing
    
    Unload subPrestamoVigente.object
    Set subPrestamoVigente.object = Nothing
    
End Sub

Private Sub ActiveReport_ReportStart()

    Set subDatosFactura.object = New ARDatosFactura
    Set subPrestamoVigente.object = New ARCreditoVigente
    
End Sub

Private Sub Detail_Format()
Dim sql As String
Dim rec As rdoResultset
Dim rec1 As rdoResultset

On Error GoTo merror

With RDODataControl1.Resultset
     If .rdoColumns("veraz") Then
        FieldVeraz.Text = "SI"
     Else
        FieldVeraz.Text = "NO"
     End If
     If .rdoColumns("facturaservicio") Then
        FieldFacturas.Text = "SI"
     Else
        FieldFacturas.Text = "NO"
     End If
     If .rdoColumns("recibosueldo") Then
        FieldRecibo.Text = "SI"
     Else
        FieldRecibo.Text = "NO"
     End If
     
     If IsNull(.rdoColumns("monotributista")) Then
        FieldMonotibuto.Text = ""
     Else
        If .rdoColumns("monotributista") Then
            FieldMonotibuto.Text = "SI"
        Else
            FieldMonotibuto.Text = "NO"
        End If
     End If
     
     If IsNull(.rdoColumns("Jubilado")) Then
        FieldJubilado.Text = ""
     Else
        If .rdoColumns("Jubilado") Then
            FieldJubilado.Text = "SI"
        Else
            FieldJubilado.Text = "NO"
        End If
     End If
     
     If IsNull(.rdoColumns("cad1")) Then
        Fieldsexo.Text = ""
     Else
        If .rdoColumns("cad1") = "M" Then
            Fieldsexo.Text = "MASCULINO"
        ElseIf .rdoColumns("cad1") = "F" Then
            Fieldsexo.Text = "FEMENINO"
        End If
     End If
     
     FieldEdad.Text = CStr(Year(Date) - Year(.rdoColumns("fechanacimiento")))
     FieldSaldo.Text = ObtenerSaldoCliente(.rdoColumns("idcliente"))
     
     sql = "SELECT * FROM DatosFactura WHERE IdCliente = " & .rdoColumns("IdCliente")
     Set rec = cnSQL.OpenResultset(sql)
     subDatosFactura.object.RDODataControl1.Resultset = rec
     
     sql = "SELECT * FROM Creditos WHERE IdCliente = " & .rdoColumns("IdCliente") & " and fechabloqueo is null and fechafinalizacion is  null order by codprestamo"
     Set rec1 = cnSQL.OpenResultset(sql)
     subPrestamoVigente.object.RDODataControl1.Resultset = rec1
     
End With



Exit Sub
merror:
tratarerrores "Error en reporte ARClientesLista2 " & Err.Number & " " & Err.Description
End Sub

