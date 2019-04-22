VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExportarVeraz 
   BackColor       =   &H00C0C000&
   Caption         =   "Exportar Archivos a Veraz"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " Datos de la exportacion: "
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6375
      Begin VB.CheckBox chkVuelcoInicial 
         Caption         =   "Vuelco Inicial"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   255
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Fecha de nacimiento del cliente"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54722561
         CurrentDate     =   39018
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         ToolTipText     =   "Fecha de nacimiento del cliente"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54722561
         CurrentDate     =   39018
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Hasta:"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha Desde:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "Exportar archivos"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Importa los cobros de los archivos seleccionados"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "frmExportarVeraz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function DetalleAltas(rec As rdoResultset, dFechaFin As Date) As String

    Dim sql             As String
    Dim cMatriz         As String
    Dim cSucursal       As String
    Dim cSector         As String
    Dim cTipo           As String
    Dim cNroOperacion   As String
    Dim cVinculacion    As String
    Dim cNroCliente     As String
    Dim cSinUso         As String
    Dim cApellidoNombre As String
    Dim cFechaNac       As String
    Dim cNroDoc1        As String
    Dim cNroDoc2        As String
    Dim cPciaCedula     As String
    Dim cSinUso2        As String
    Dim cTipoSocPers    As String
    Dim cEstadoCivil    As String
    Dim cSinUso3        As String
    Dim cMarcaDireccion As String
    Dim cDireccion      As String
    Dim cLocalidad      As String
    Dim cProvincia      As String
    Dim cCodPostal      As String
    Dim cFechaServicio  As String
    Dim cCargo          As String
    Dim cSinUso4        As String
    Dim cRetorno        As String
    Dim cTelefono       As String
    Dim cNacionalidad   As String
    Dim cSinUso5        As String
    Dim cPrimerParteDet As String
    Dim cSinUso6        As String
    Dim cTelAux         As String
    Dim nPos            As Integer
    Dim nSaldo          As Currency
    Dim recCliente      As rdoResultset
    Dim recLocalidad    As rdoResultset
    Dim recPrimerCredit As rdoResultset

    On Error GoTo merror
    
    If chkVuelcoInicial.Value Then
        nSaldo = ObtenerSaldoCredito(rec.rdoColumns("IdCredito"), dFechaFin)
        If nSaldo < 0 Then
            nSaldo = 0
        End If
        If Not IsNull(rec.rdoColumns("FechaFinalizacion")) Then
            DetalleAltas = "NoGrabar"
            Exit Function
        End If
        If nSaldo < 1 Then
            DetalleAltas = "NoGrabar"
            Exit Function
        End If
    End If
    
    sql = "DatosCliente " & rec.rdoColumns("IdCliente")
    Set recCliente = cnSQL.OpenResultset(sql)
    If recCliente.EOF Then
        tratarerrores "Error grave recuperando cliente al generar detalle del archivo de Altas de Veraz: " & rec.rdoColumns("IdCliente")
        End
    End If
    
    sql = "DatosLocalidad " & recCliente.rdoColumns("IdLocalidad")
    Set recLocalidad = cnSQL.OpenResultset(sql)
    If recLocalidad.EOF Then
        tratarerrores "Error grave recuperando la localidad al generar detalle del archivo de Altas de Veraz: " & recCliente.rdoColumns("IdLocalidad")
        End
    End If
    
    sql = "PrimerCredito " & rec.rdoColumns("IdCliente")
    Set recPrimerCredit = cnSQL.OpenResultset(sql)
    If recPrimerCredit.EOF Then
        tratarerrores "Error grave recuperando el primer credito al generar detalle del archivo de Altas de Veraz: " & rec.rdoColumns("IdCliente")
        End
    End If
    
    cMatriz = "C13712"
    cSucursal = "0000"
    cSector = "PP"
    cTipo = "1P"
    cNroOperacion = Format(rec.rdoColumns("IdCredito"), "0000000000") & Space(10)
    cVinculacion = "T"
    cNroCliente = Format(rec.rdoColumns("IdCliente"), "00000000000000000000")
    cSinUso = Space(1)
    cApellidoNombre = Trim$(recCliente.rdoColumns("Apellido")) & "," & Trim$(recCliente.rdoColumns("Nombre")) & Space(72 - Len(Trim$(recCliente.rdoColumns("Apellido")) & "," & Trim$(recCliente.rdoColumns("Nombre"))))
    cSinUso2 = Space(25)
    cFechaNac = Format(recCliente.rdoColumns("FechaNacimiento"), "YYYYMMDD")
    cNroDoc1 = Format(recCliente.rdoColumns("NumDocumento"), "00000000000")
    cNroDoc2 = "00000000000"
    cPciaCedula = Space(1)
    cSinUso3 = Space(3)

    If recCliente.rdoColumns("cad1") = "M" Or recCliente.rdoColumns("cad1") = "F" Then
        cTipoSocPers = recCliente.rdoColumns("cad1")
    Else
        cTipoSocPers = "I"
    End If
    
    Select Case recCliente.rdoColumns("EstadoCivil")
    Case "CASADO/A"
        cEstadoCivil = "C"
    Case "SOLTERO/A"
        cEstadoCivil = "S"
    Case "SOLTERO/A CON HIJOS"
        cEstadoCivil = "S"
    Case "DIVORCIADO/A"
        cEstadoCivil = "D"
    Case "VIUDO/A"
        cEstadoCivil = "V"
    Case Else
        cEstadoCivil = Space(1)
    End Select

    cSinUso4 = Space(8)
    cMarcaDireccion = "J"
    
    If Len(Trim$(recCliente.rdoColumns("Domicilio"))) < 39 Then
        cDireccion = Trim$(recCliente.rdoColumns("Domicilio")) & Space(39 - Len(Trim$(recCliente.rdoColumns("Domicilio"))))
    Else
        cDireccion = Mid$(Trim$(recCliente.rdoColumns("Domicilio")), 1, 39)
    End If
    
    nPos = InStr(1, recLocalidad.rdoColumns("Nombre"), "-")
    If nPos = 0 Or nPos > 20 Then
        cLocalidad = Mid$(Trim$(recLocalidad.rdoColumns("Nombre")), 1, 20) & Space(20 - Len(Mid$(Trim$(recLocalidad.rdoColumns("Nombre")), 1, 20)))
    Else
        cLocalidad = Mid$(Trim$(recLocalidad.rdoColumns("Nombre")), 1, nPos - 1) & Space(20 - Len(Mid$(Trim$(recLocalidad.rdoColumns("Nombre")), 1, nPos - 1)))
    End If
    
    Select Case rec.rdoColumns("IdProvincia")
    Case 2
        cProvincia = "B"
    Case 4
        cProvincia = "N"
    Case 5
        cProvincia = "C"
        cLocalidad = Space(20)
    Case 6
        cProvincia = "H"
    Case 7
        cProvincia = "W"
    Case 8
        cProvincia = "P"
    Case 9
        cProvincia = "E"
    Case 10
        cProvincia = "S"
    Case Else
        cProvincia = " "
    End Select
    
    cCodPostal = Mid$(recCliente.rdoColumns("CodigoPostal"), 1, 8) & Space(8 - Len(Mid$(recCliente.rdoColumns("CodigoPostal"), 1, 8)))
    cFechaServicio = Format(recPrimerCredit.rdoColumns("FechaCredito"), "YYYYMMDD")
    cCargo = Space(2)
    cSinUso5 = Space(6)
    cRetorno = Space(2)
    
    cTelefono = ""
    cTelAux = recCliente.rdoColumns("Telefono")
    For nPos = 1 To 50
        If Mid$(cTelAux, nPos, 1) >= "0" And Mid$(cTelAux, nPos, 1) <= "9" Then
            cTelefono = Trim$(cTelefono) & Mid$(cTelAux, nPos, 1)
        End If
    Next
    cTelefono = Mid$(Trim$(cTelefono), 1, 14) & Space(14 - Len(Mid$(Trim$(cTelefono), 1, 14)))

    Select Case Mid$(recCliente.rdoColumns("Nacionalidad"), 1, 1)
    Case " "
        cNacionalidad = "A"
    Case "A"
        cNacionalidad = "A"
    Case Else
        cNacionalidad = "E"
    End Select
    
    cSinUso6 = Space(1)

    recCliente.Close
    recLocalidad.Close
    recPrimerCredit.Close
    
    cPrimerParteDet = cMatriz & _
                      cSucursal & _
                      cSector & _
                      cTipo & _
                      cNroOperacion & _
                      cVinculacion & _
                      cNroCliente & _
                      cSinUso & _
                      cApellidoNombre & _
                      cSinUso2 & _
                      cFechaNac & _
                      cNroDoc1 & _
                      cNroDoc2 & _
                      cPciaCedula & _
                      cSinUso3 & _
                      cTipoSocPers & _
                      cEstadoCivil & _
                      cSinUso4
                   
        DetalleAltas = cPrimerParteDet & _
                       cMarcaDireccion & _
                       cDireccion & _
                       cLocalidad & _
                       cProvincia & _
                       cCodPostal & _
                       cFechaServicio & _
                       cCargo & _
                       cSinUso5 & _
                       cRetorno & _
                       cTelefono & _
                       cNacionalidad & _
                       cSinUso6

Exit Function
merror:
tratarerrores "Error generando detalle del archivo de Altas de Veraz: " & Err.Number & " " & Err.Description
End Function

Private Function DetalleModif(rec As rdoResultset, dFechaInicio As Date, dFechaFin As Date, ByRef nCampo4 As Currency) As String

    Dim sql             As String
    Dim cMatriz         As String
    Dim cSucursal       As String
    Dim cSector         As String
    Dim cTipo           As String
    Dim cNroOperacion   As String
    Dim cNroCliente     As String
    Dim cUltCompra      As String
    Dim cStatus         As String
    Dim cCampo1         As String
    Dim cCampo2         As String
    Dim cCampo3         As String
    Dim cCampo4         As String
    Dim cCampo5         As String
    Dim cCampo6         As String
    Dim cCampo7         As String
    Dim cRetorno        As String
    Dim cActualizado    As String
    Dim cFechaInform    As String
    Dim cCampo9         As String
    Dim nSaldo          As Currency
    Dim nCampo2         As Currency
    Dim nDiasDeuda      As Integer
    Dim nCuotasPend     As Integer
    Dim bSinCuoVigente  As Boolean
    Dim bSinCuoImpaga   As Boolean
    Dim bFinalizado     As Boolean
    Dim bSinSaldoVenc   As Boolean
    Dim bSinInfo        As Boolean
    Dim recCuota        As rdoResultset
    Dim recCuotasImpaga As rdoResultset
    Dim recPrimerCuota  As rdoResultset
    Dim recPrimerImpaga As rdoResultset
    
    On Error GoTo merror
    
    bFinalizado = False
    bSinSaldoVenc = False
    bSinInfo = False
    
    nSaldo = ObtenerSaldoCredito(rec.rdoColumns("IdCredito"), dFechaFin)
    If nSaldo < 0 Then
        nSaldo = 0
    End If
    
    If chkVuelcoInicial.Value Then
        sql = "CuotaDelMes " & rec.rdoColumns("IdCredito") & ",'" & ConvertirFechaSql("01/" & Format(dFechaFin, "MM/YYYY"), "DD/MM/YYYY") & "','" & ConvertirFechaSql(Format(dFechaFin, "DD/MM/YYYY"), "DD/MM/YYYY") & "'"
        If Not IsNull(rec.rdoColumns("FechaFinalizacion")) Then
            DetalleModif = "NoGrabar"
            Exit Function
        End If
        If nSaldo < 1 Then
            DetalleModif = "NoGrabar"
            Exit Function
        End If
    Else
        sql = "CuotaDelMes " & rec.rdoColumns("IdCredito") & ",'" & ConvertirFechaSql(Format(dFechaInicio, "DD/MM/YYYY"), "DD/MM/YYYY") & "','" & ConvertirFechaSql(Format(dFechaFin, "DD/MM/YYYY"), "DD/MM/YYYY") & "'"
        If nSaldo < 1 Then
            bFinalizado = True
        End If
    End If
    Set recCuota = cnSQL.OpenResultset(sql)
    bSinCuoVigente = False
    If recCuota.EOF Then
        bSinCuoVigente = True
    End If
    
    sql = "PrimerCuota " & rec.rdoColumns("IdCredito")
    Set recPrimerCuota = cnSQL.OpenResultset(sql)
    If recPrimerCuota.EOF Then
        tratarerrores "Error grave recuperando la primer cuota del credito: " & rec.rdoColumns("IdCredito")
        End
    End If
    
    sql = "PrimerCuotaImpaga " & rec.rdoColumns("IdCredito")
    Set recPrimerImpaga = cnSQL.OpenResultset(sql)
    bSinCuoImpaga = False
    If recPrimerImpaga.EOF Then
        bSinCuoImpaga = True
    End If
         
    sql = "CantCuotasImpagas " & rec.rdoColumns("IdCredito")
    Set recCuotasImpaga = cnSQL.OpenResultset(sql)
    If recCuotasImpaga.EOF Then
        tratarerrores "Error grave recuperando cantidad de cuotas impagas al generar detalle del archivo de Actualizaciones de Veraz: " & rec.rdoColumns("IdCredito")
        End
    End If
   
    cMatriz = "C13712"
    cSucursal = "0000"
    cSector = "PP"
    cTipo = "1P"
    cNroOperacion = Format(rec.rdoColumns("IdCredito"), "0000000000") & Space(10)
    cNroCliente = Format(rec.rdoColumns("IdCliente"), "00000000000000000000")
    cUltCompra = Space(6)
       
    If rec.rdoColumns("logic1") Then
        cCampo1 = "PR/SG" & " ARS"
    Else
        cCampo1 = "PP/SG" & " ARS"
    End If
    
    If recPrimerCuota.rdoColumns("FechaVencimiento1") > dFechaFin Then
        bSinInfo = True
    End If
    
    nCampo2 = 0
    If bSinCuoVigente Then
        If bSinInfo Then
            nCampo2 = 0
        Else
            nCampo2 = nSaldo
        End If
        cCampo6 = Format(recPrimerCuota.rdoColumns("ImporteVencimiento1"), "000000000")
    Else
        nCampo2 = recCuota.rdoColumns("ImporteVencimiento1")
        cCampo6 = Format(recCuota.rdoColumns("ImporteVencimiento1"), "000000000")
    End If
    
    cCampo3 = Format(rec.rdoColumns("ImporteAFinanciar"), "000000000")
    cCampo4 = Format(nSaldo, "000000000")
    nCampo4 = CCur(cCampo4)
    
    If bSinInfo Then
        nCuotasPend = rec.rdoColumns("NumCuotas")
    Else
        If recCuotasImpaga.rdoColumns("CANTIDAD") = rec.rdoColumns("NumCuotas") Then
            If rec.rdoColumns("NumCuotas") <> 1 Then
                nCuotasPend = recCuotasImpaga.rdoColumns("CANTIDAD") - 1
            Else
                nCuotasPend = 1
            End If
        Else
            nCuotasPend = recCuotasImpaga.rdoColumns("CANTIDAD")
        End If
    End If
    
    cCampo5 = Space(3 - Len(Trim$(CStr(rec.rdoColumns("NumCuotas"))))) & Trim$(CStr(rec.rdoColumns("NumCuotas"))) & "M /" & Space(3 - Len(Trim$(CStr(nCuotasPend)))) & Trim$(CStr(nCuotasPend))
    
    If chkVuelcoInicial.Value Then
        If bSinCuoVigente Then
            If bSinInfo Then
                nSaldo = 0
            Else
                nSaldo = ObtenerSaldoCredito(rec.rdoColumns("IdCredito"), dFechaFin)
            End If
        Else
            nSaldo = ObtenerSaldoVencido(rec.rdoColumns("IdCredito"), "01/" & Format(dFechaFin, "MM/YYYY"))
            nCampo2 = nCampo2 + nSaldo
        End If
    Else
        If bSinCuoVigente Then
            If bSinInfo Then
                nSaldo = 0
            Else
                nSaldo = ObtenerSaldoCredito(rec.rdoColumns("IdCredito"), dFechaFin)
            End If
        Else
            nSaldo = ObtenerSaldoVencido(rec.rdoColumns("IdCredito"), dFechaInicio)
            nCampo2 = nCampo2 + nSaldo
        End If
    End If
    
    If nCampo2 > nCampo4 Then
        nCampo2 = nCampo4
    End If
    
    cCampo2 = Format(nCampo2, "000000000")
    
    If nSaldo < 1 Then
        nSaldo = 0
        bSinSaldoVenc = True
    End If
    cCampo7 = Format(nSaldo, "000000000")
    
    cRetorno = Space(2)
    cActualizado = Space(1)
    cFechaInform = Format(dFechaFin, "YYYYMMDD")
    cCampo9 = Space(5)
        
    nDiasDeuda = 0
    cStatus = ""
    If Not bFinalizado Then
        If bSinInfo Then
            cStatus = "0"
        Else
            If bSinSaldoVenc Then
                cStatus = "1"
            Else
                nDiasDeuda = DateDiff("d", recPrimerImpaga.rdoColumns("FechaVencimiento1"), dFechaFin)
            End If
        End If
    Else
        cCampo2 = "000000000"
        cCampo4 = "000000000"
        cCampo5 = Mid(cCampo5, 1, 8) & "0"
        cCampo7 = "000000000"
        cStatus = "C"
    End If
    
    If Not IsNull(rec.rdoColumns("FechaBloqueo")) Then
        cStatus = "G"
    End If
        
    If nDiasDeuda < 0 Then
        nDiasDeuda = nDiasDeuda * (-1)
    End If
    
    If cStatus <> "C" And cStatus <> "G" Then
        If nDiasDeuda > 0 Then
            If nDiasDeuda > 0 And nDiasDeuda <= 30 Then
                cStatus = "1"
            ElseIf nDiasDeuda >= 31 And nDiasDeuda <= 60 Then
                cStatus = "2"
            ElseIf nDiasDeuda >= 61 And nDiasDeuda <= 90 Then
                cStatus = "3"
            ElseIf nDiasDeuda >= 91 And nDiasDeuda <= 120 Then
                cStatus = "4"
            ElseIf nDiasDeuda >= 121 And nDiasDeuda <= 150 Then
                cStatus = "5"
            ElseIf nDiasDeuda >= 151 And nDiasDeuda <= 180 Then
                cStatus = "6"
            ElseIf nDiasDeuda >= 181 Then
                cStatus = "9"
            End If
        Else
            If cStatus = "" Then
                cStatus = "1"
            End If
        End If
    End If
    
    recCuota.Close
    recCuotasImpaga.Close
    
    DetalleModif = cMatriz & _
                   cSucursal & _
                   cSector & _
                   cTipo & _
                   cNroOperacion & _
                   cNroCliente & _
                   cUltCompra & _
                   cStatus & _
                   cCampo1 & _
                   cCampo2 & _
                   cCampo3 & _
                   cCampo4 & _
                   cCampo5 & _
                   cCampo6 & _
                   cCampo7 & _
                   cRetorno & _
                   cActualizado & _
                   cFechaInform & _
                   cCampo9

Exit Function
merror:
tratarerrores "Error generando detalle del archivo de Altas de Veraz: " & Err.Number & " " & Err.Description
End Function

Private Function GeneraVerazAltas() As Boolean

    Dim rec                 As rdoResultset
    Dim dFechaDesde         As Date
    Dim dFechaHasta         As Date
    Dim nMes                As Integer
    Dim nAnio               As Integer
    Dim nCantRegDetalle     As Integer
    Dim cNombreArchivo      As String
    Dim cArchivoCompleto    As String
    Dim cNombreArchivo2     As String
    Dim cArchivoCompleto2   As String
    Dim cHeader             As String
    Dim cDetalle            As String
    Dim cTrailer            As String
    Dim cPeriodo            As String
    Dim sql                 As String

    On Error GoTo merror
    Call RefreshTimer

    GeneraVerazAltas = False

    dFechaDesde = CDate(dtpFechaDesde.Value)
    dFechaHasta = CDate(dtpFechaHasta.Value)

    nMes = Month(dFechaHasta)
    nAnio = Year(dFechaHasta)
    cPeriodo = Format(nAnio, "0000") & Format(nMes, "00")
    
    cNombreArchivo = "tempVerazAltas_C13712.txt"

    If Not ExisteCarpeta() Then Exit Function

    cArchivoCompleto = "c:\ExportacionExcel\" & cNombreArchivo

    If Trim(Dir(cArchivoCompleto)) <> "" Then
        Kill (cArchivoCompleto)
    End If
       
    Open cArchivoCompleto For Output As #1
    
    cHeader = HeaderAltas(cPeriodo)
    Print #1, cHeader
    
    sql = "SELECT * From Creditos WHERE FechaCredito >= '" & ConvertirFechaSql(dFechaDesde, "DD/MM/YYYY") & "' AND FechaCredito <= '" & ConvertirFechaSql(dFechaHasta, "DD/MM/YYYY") & "'"
    Set rec = cnSQL.OpenResultset(sql)
    Call RefreshTimer

    nCantRegDetalle = 0
    If Not rec.EOF Then
        Do While Not rec.EOF
            cDetalle = DetalleAltas(rec, dFechaHasta)
            If cDetalle <> "NoGrabar" Then
                Print #1, cDetalle
                Call RefreshTimer
                nCantRegDetalle = nCantRegDetalle + 1
            End If
            rec.MoveNext
       Loop
    End If
    
    cTrailer = TrailerAltas(nCantRegDetalle)
    Print #1, cTrailer
   
    Close #1
    
    rec.Close
    
    cNombreArchivo2 = "C13712_altas1P" & cPeriodo & "_" & Format(nCantRegDetalle + 2, "0000")
    
    cArchivoCompleto2 = "c:\ExportacionExcel\" & cNombreArchivo2

    If Trim(Dir(cArchivoCompleto2)) <> "" Then
        Kill (cArchivoCompleto2)
    End If
       
    Name cArchivoCompleto As cArchivoCompleto2
    
    GeneraVerazAltas = True

Exit Function
merror:
tratarerrores "Error exportando el archivo de Altas de Veraz: " & Err.Number & " " & Err.Description
End Function

Private Function GeneraVerazModif() As Boolean

    Dim rec                 As rdoResultset
    Dim dFechaDesde         As Date
    Dim dFechaHasta         As Date
    Dim nMes                As Integer
    Dim nAnio               As Integer
    Dim nCantRegDetalle     As Integer
    Dim nSumCampo4          As Currency
    Dim nCampo4             As Currency
    Dim cDesde              As String
    Dim cHasta              As String
    Dim cNombreArchivo      As String
    Dim cArchivoCompleto    As String
    Dim cNombreArchivo2     As String
    Dim cArchivoCompleto2   As String
    Dim cHeader             As String
    Dim cDetalle            As String
    Dim cTrailer            As String
    Dim cPeriodo            As String
    Dim sql                 As String

    On Error GoTo merror
    Call RefreshTimer

    GeneraVerazModif = False

    dFechaDesde = CDate(dtpFechaDesde.Value)
    dFechaHasta = CDate(dtpFechaHasta.Value)

    nMes = Month(dFechaHasta)
    nAnio = Year(dFechaHasta)
    cPeriodo = Format(nAnio, "0000") & Format(nMes, "00")
    
    cNombreArchivo = "tempVerazModif_C13712.txt"

    If Not ExisteCarpeta() Then Exit Function

    cArchivoCompleto = "c:\ExportacionExcel\" & cNombreArchivo

    If Trim(Dir(cArchivoCompleto)) <> "" Then
        Kill (cArchivoCompleto)
    End If
       
    Open cArchivoCompleto For Output As #1
    
    cHeader = HeaderModif(cPeriodo)
    Print #1, cHeader
    
    cDesde = ConvertirFechaSql(dFechaDesde, "DD/MM/YYYY")
    cHasta = ConvertirFechaSql(dFechaHasta, "DD/MM/YYYY")
    sql = "SELECT * From Creditos " & _
          "WHERE FechaCredito <= '" & cHasta & "' AND " & _
          "((FechaFinalizacion >= '" & cDesde & "' AND FechaFinalizacion <= '" & cHasta & "') OR " & _
          "(FechaBloqueo >= '" & cDesde & "' AND FechaBloqueo <= '" & cHasta & "') OR " & _
          "(FechaFinalizacion IS NULL AND FechaBloqueo IS NULL))"
    
    Set rec = cnSQL.OpenResultset(sql)
    Call RefreshTimer

    nCantRegDetalle = 0
    nSumCampo4 = 0
    If Not rec.EOF Then
        Do While Not rec.EOF
            cDetalle = DetalleModif(rec, dFechaDesde, dFechaHasta, nCampo4)
            If cDetalle <> "NoGrabar" Then
                Print #1, cDetalle
                nCantRegDetalle = nCantRegDetalle + 1
                nSumCampo4 = nSumCampo4 + nCampo4
            End If
            rec.MoveNext
            Call RefreshTimer
       Loop
    End If
    
    cTrailer = TrailerModif(nCantRegDetalle, nSumCampo4)
    Print #1, cTrailer
   
    Close #1
    
    rec.Close
    
    cNombreArchivo2 = "C13712_datos1P" & cPeriodo & "_" & Format(nCantRegDetalle + 2, "0000")
    
    cArchivoCompleto2 = "c:\ExportacionExcel\" & cNombreArchivo2

    If Trim(Dir(cArchivoCompleto2)) <> "" Then
        Kill (cArchivoCompleto2)
    End If
       
    Name cArchivoCompleto As cArchivoCompleto2
    
    GeneraVerazModif = True

Exit Function
merror:
tratarerrores "Error exportando el archivo de Altas de Veraz: " & Err.Number & " " & Err.Description
End Function


Private Function HeaderAltas(ByVal cPeriodoPar As String) As String

    Dim cMatriz         As String
    Dim cTipoReg        As String
    Dim cTipo           As String
    Dim cArchivo        As String
    Dim cFecha          As String
    Dim cHora           As String
    Dim cMedio          As String
    Dim cBloqueo        As String
    Dim cPeriodo        As String
    Dim cCodifOrigen    As String
    Dim cMarcaTarjeta   As String
    Dim cFiller         As String
    Dim cY2K            As String
    
    cMatriz = "C13712"
    cTipoReg = "HHHHHH"
    cTipo = "1P"
    cArchivo = "ALTAS"
    cFecha = Format(Date, "YYYYMMDD")
    cHora = Format(Time, "HHMMSS")
    cMedio = "FT"
    cBloqueo = "000000"
    cPeriodo = cPeriodoPar
    cCodifOrigen = "A"
    cMarcaTarjeta = Space(2)
    cFiller = Space(247)
    cY2K = "Y2K"
    
    HeaderAltas = cMatriz & _
                  cTipoReg & _
                  cTipo & _
                  cArchivo & _
                  cFecha & _
                  cHora & _
                  cMedio & _
                  cBloqueo & _
                  cPeriodo & _
                  cCodifOrigen & _
                  cMarcaTarjeta & _
                  cFiller & _
                  cY2K
    
End Function

Private Function HeaderModif(ByVal cPeriodoPar As String) As String

    Dim cMatriz         As String
    Dim cTipoReg        As String
    Dim cTipo           As String
    Dim cArchivo        As String
    Dim cFecha          As String
    Dim cHora           As String
    Dim cMedio          As String
    Dim cBloqueo        As String
    Dim cPeriodo        As String
    Dim cCodifOrigen    As String
    Dim cMarcaTarjeta   As String
    Dim cFiller         As String
    Dim cY2K            As String
    
    cMatriz = "C13712"
    cTipoReg = "HHHHHH"
    cTipo = "1P"
    cArchivo = "DATOS"
    cFecha = Format(Date, "YYYYMMDD")
    cHora = Format(Time, "HHMMSS")
    cMedio = "FT"
    cBloqueo = "000000"
    cPeriodo = cPeriodoPar
    cCodifOrigen = "A"
    cMarcaTarjeta = Space(2)
    cFiller = Space(87)
    cY2K = "Y2K"
    
    HeaderModif = cMatriz & _
                  cTipoReg & _
                  cTipo & _
                  cArchivo & _
                  cFecha & _
                  cHora & _
                  cMedio & _
                  cBloqueo & _
                  cPeriodo & _
                  cCodifOrigen & _
                  cMarcaTarjeta & _
                  cFiller & _
                  cY2K
    
End Function

Private Function TrailerAltas(ByVal nCantReg As Long) As String

    Dim cMatriz         As String
    Dim cTipoReg        As String
    Dim cTipo           As String
    Dim cCantReg        As String
    Dim cFiller         As String
    
    cMatriz = "C13712"
    cTipoReg = "TTTTTT"
    cTipo = "1P"
    cCantReg = Format(nCantReg, "00000000")
    cFiller = Space(278)
    
    TrailerAltas = cMatriz & _
                   cTipoReg & _
                   cTipo & _
                   cCantReg & _
                   cFiller
    
End Function

Private Function TrailerModif(ByVal nCantReg As Long, nSumCampo4 As Currency) As String

    Dim cMatriz         As String
    Dim cTipoReg        As String
    Dim cTipo           As String
    Dim cCantReg        As String
    Dim cSumCampo4      As String
    Dim cFiller         As String
    
    cMatriz = "C13712"
    cTipoReg = "TTTTTT"
    cTipo = "1P"
    cCantReg = Format(nCantReg, "00000000")
    cSumCampo4 = Format(nSumCampo4, "000000000000000000")
    cFiller = Space(100)
    
    TrailerModif = cMatriz & _
                   cTipoReg & _
                   cTipo & _
                   cCantReg & _
                   cSumCampo4 & _
                   cFiller
    
End Function

Private Sub CmdCerrar_Click()
Call RefreshTimer
    Unload Me
End Sub

Private Sub CmdExportar_Click()
    
    Dim bResultadoProceso       As Boolean
    Call RefreshTimer
    
    If Day(dtpFechaDesde) <> 1 Then
       MsgE "La fecha de inicio debe empezar el dìa 1 del mes"
       Exit Sub
    End If
    
    If chkVuelcoInicial.Value Then
        If dtpFechaDesde > dtpFechaHasta Then
            MsgE "La fecha final debe ser superior a la inicial"
            Exit Sub
        End If
    Else
        If Month(dtpFechaDesde) <> Month(dtpFechaHasta) Or _
           Year(dtpFechaDesde) <> Year(dtpFechaHasta) Then
            MsgE "El mes y año de ambas fechas debe ser el mismo"
            Exit Sub
        End If
    End If
    
    If Not MsgP("¿Confirma la exportacion de los archivos para Veraz del " & dtpFechaDesde & " al " & dtpFechaHasta & "?") Then Exit Sub
    
    CmdExportar.Enabled = False
    Me.MousePointer = vbHourglass
    
    bResultadoProceso = GeneraVerazAltas()
    If bResultadoProceso Then
        bResultadoProceso = GeneraVerazModif()
    End If
    
    If bResultadoProceso Then
        MsgI "Se generaron con èxito los archivos para Veraz en la carpeta C:\ExportacionExcel\"
    End If
    
    Me.MousePointer = vbDefault
    CmdExportar.Enabled = True
    
End Sub

Private Sub Form_Load()
    
    Dim dFechaActual        As Date
    Dim nMesActual          As Integer
    Dim nAnioActual         As Integer
    Dim cFechaDesde         As String
    Dim nAnioDesde          As Integer
    Dim nMesDesde           As Integer
    Dim cFechaHasta         As String
    Dim nDiaHasta           As Integer
    
    Call RefreshTimer
    
    dFechaActual = Date
    nMesActual = Month(dFechaActual)
    nAnioActual = Year(dFechaActual)
    If nMesActual = 1 Then
        nMesDesde = 12
        nAnioDesde = nAnioActual - 1
    Else
        nMesDesde = nMesActual - 1
        nAnioDesde = nAnioActual
    End If
    
    Select Case nMesDesde
    Case 1, 3, 5, 7, 8, 10, 12
        nDiaHasta = 31
    Case 4, 6, 9, 11
        nDiaHasta = 30
    Case 2
        If nAnioDesde Mod 4 = 0 Then
            nDiaHasta = 29
        Else
            nDiaHasta = 28
        End If
    End Select
    
    cFechaDesde = "01/" & Format$(nMesDesde, "00") & "/" & Format$(nAnioDesde, "0000")
    cFechaHasta = Format$(nDiaHasta, "00") & "/" & Format$(nMesDesde, "00") & "/" & Format$(nAnioDesde, "0000")
    
    dtpFechaDesde = cFechaDesde
    dtpFechaHasta = cFechaHasta
    
End Sub
