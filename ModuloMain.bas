Attribute VB_Name = "ModuloMain"
'***MODULO INICIAL DEL SISTEMA (PUNTO DE PARTIDA)
Option Explicit


'Conexion con la base de datos SQL Server
Public enSQL As rdoEnvironment
Public cnSQL As rdoConnection
Public cConexion   As String

Public cAmbiente   As String

'variables globales usadas en todas las pantallas del sistema
'nombre de base
'Public VG_NOMBREBASE As String
'variable global para referenciar la base de datos
Public Db As Database
'Espacio de trabajo de la base de datos abierta
Public wrkODBC As Workspace
'Variable donde se encuentra la CONEXION de la
'base de datos ej. "ODBC;DSN=BDCoop;UID=Cesar..."
'Public CONEXION As String
'Motor que se utilizara para abrir la Base
'Public MOTOR As Variant

'para la base
'Public VG_USUARIOBASE As String
'Public VG_CLAVEBASE As String

'para el login
Public VG_LOGIN As Boolean
Public VG_IDUSUARIOLOGIN As Long
Public VG_IDTIPOUSUARIOLOGIN As Long
Public VG_USUARIOLOGIN As String
Public VG_CLAVELOGIN As String

'variables de tipo de usuario para habilitar pantallas
Public VG_ACTUALIZA As Boolean
Public VG_REGISTRA As Boolean
Public VG_REFINANCIA As Boolean
Public VG_CONSULTA As Boolean
Public VG_COBRA As Boolean
Public VG_ANULA As Boolean
Public VG_ADMCLIENTES As Boolean
Public VG_ADMCOBRADORES As Boolean
Public VG_ADMCREDITOS As Boolean
Public VG_EFECTUABACKUP As Boolean
Public VG_EMITELIBREDEUDA As Boolean
Public VG_EMITECARTARECLAMO As Boolean
Public VG_CONSULTADEUDORES As Boolean
Public VG_ACTUALIZAOPCIONES As Boolean
Public VG_ADMCOMODINES As Boolean
Public VG_IMPRIMECUOTAS As Boolean
Public VG_ADMPLANES As Boolean
Public VG_EXPORTA As Boolean
Public VG_IMPORTA As Boolean
Public VG_CONSULTAINGRESOS As Boolean
Public VG_ACERCADE As Boolean

'datos empresa
Public VG_EMPRESA As String
Public VG_CIUDAD As String
Public VG_CUIT As String
Public VG_IVA As String
Public VG_INGRESOSBRUTOS As String
Public VG_DOMICILIO As String
Public VG_TELEFONO As String
Public VG_EMAIL As String
Public VG_WEBSITE As String
Public VG_HORARIOATENCION As String
Public VG_LUGARESPAGO As String

'requisitos generales
'si exijo garante
Public VG_GARANTE As Boolean
'si permito creditos simultaneos
Public VG_CLIENTESIMULTANEO As Boolean
'si permito clientes que tienen creditos bloqueados
Public VG_CLIENTEJUDICIAL As Boolean
'pagar salteadas
Public VG_PAGARCUOTASDESORDENADAS As Boolean
'finalizar automaticamente
Public VG_FINALIZARAUTOMATICAMENTE As Boolean
'si aplico cobros parciales
Public VG_APLICARCOBROSPARCIALES As Boolean
'si redondeo
Public VG_REDONDEAR As Boolean
'este es para ver si pregunta o no por la impresion del recibo al cobrar
Public VG_APLICARRECIBOS As Boolean
'si permite creditos diferidos
Public VG_CREDITOSDIFERIDOS As Boolean
'si permite cuotas diferidas
Public VG_COBROSDIFERIDOS As Boolean
'Si hay vencimiento los sabados
Public VG_APLICARVENCIMIENTOSABADOS As Boolean
'Edad maxima para otorgar un credito
Public VG_EDAD As Long
'Cantidad de dias q no puede superar el Credito la primer fecha de vencimiento
Public VG_CANT_DIAS As Integer
'Fecha limite de ingreso de un credito
Public VG_FECHALIMITEINGRESO As Date
'Cantidad de dias de mora permitida
Public VG_DIAS_MORA As Integer
'Antiguedad de la la mora permitida
Public VG_ANTIGUEDAD_MORA As Integer
'Monto máximo de la mora permitida
Public VG_MONTO_MORA As Currency
'Tiempo para dar logout automático
Public VG_TIEMPO_LOGOUT As Long

'tasas
Public VG_TASAMORA As Double
Public VG_TASAFINANCIACION As Double
Public VG_TASAREFINANCIACION As Double
Public VG_APLICARTASAREFINANCIACION As Boolean

'seguros
Public VG_APLICARSEGURO As Boolean
Public VG_NOAPLICARSEGUROSREFINANCIACION As Boolean
Public VG_APLICARSEGUROSCUOTA1 As Boolean
Public VG_IMPORTESEGURO As Currency
Public VG_SEGUROFIJO As Currency
Public VG_ALICUOTASEGUROS As Double

'gastos administrativos
Public VG_APLICARGASTOS As Boolean
Public VG_IMPORTEGASTOS As Currency
Public VG_IMPORTEGASTOSFIJOS As Currency
Public VG_PORCCAPNOINT As Double
Public VG_PORCCAPINT As Double
Public VG_PORCFUNNOCAP As Double
Public VG_NOAPLICARGASTOSREFINANCIACION As Boolean
Public VG_APLICARGASTOSCUOTA1 As Boolean
Public VG_APLICARGASTOSCUOTA2 As Boolean

'impuestos
Public VG_APLICARIMPUESTOS As Boolean
Public VG_IMPORTEIMPUESTOS As Currency
Public VG_IMPUESTOSFIJOS As Currency
Public VG_NOAPLICARIMPUESTOSREFINANCIACION As Boolean
Public VG_APLICARIMPUESTOSCUOTA1 As Boolean
Public VG_APLICARIMPUESTOSCUOTA2 As Boolean
Public VG_PORCENTAJEIVA As Double
Public VG_IMPUESTOSCREDIMACO As Boolean

'comprobantes
Public VG_APLICARSEGUNDOVENCIMIENTO As Boolean
Public VG_VENCIMIENTO2IMPORTE As Currency
Public VG_VENCIMIENTO2PORCENTAJE As Double
Public VG_APLICARVENCIMIENTO2MORA As Boolean
Public VG_DIASVENCIMIENTOFINANCIACION As Long
Public VG_DIASVENCIMIENTOREFINANCIACION As Long

'libre deuda
Public VG_TEXTOLIBREDEUDA1 As String
Public VG_TEXTOLIBREDEUDA2 As String

'carta reclamo
Public VG_TEXTOCARTARECLAMO1 As String
Public VG_TEXTOCARTARECLAMO2 As String

'pagare
Public VG_TEXTOACUERDOMUTUO1 As String
Public VG_TEXTOACUERDOMUTUO2 As String
Public VG_TEXTOACUERDOMUTUO3 As String
Public VG_TEXTOACUERDOMUTUO4 As String
Public VG_TEXTOACUERDOMUTUO5 As String
Public VG_TEXTOACUERDOMUTUO6 As String
Public VG_TEXTOACUERDOMUTUO7 As String
Public VG_TEXTOACUERDOMUTUO8 As String
Public VG_TEXTOACUERDOMUTUO9 As String
Public VG_TEXTOACUERDOMUTUO10 As String

'impresion
Public VG_MODELOFACTURA1 As Boolean
Public VG_MODELOFACTURA2 As Boolean
Public VG_MODELOFACTURA3 As Boolean
Public VG_MODELOFACTURA4 As Boolean
Public VG_IMPRIMIRMORAIVA As Boolean
Public VG_TOP As Long
Public VG_LEFT As Long
Public VG_BOTOM As Long
Public VG_NUMCOPIAS As Long
Public VG_MOSTRARRECUADROS As Boolean

'proximo numero de reciboa emitir
Public VG_ULTIMONUMRECIBO As Long

Public VG_APLICAROTORGAMIENTO As Boolean
Public VG_APLICAROTORGAMIENTOCUOTA1 As Boolean
Public VG_IMPORTEOTORGAMIENTO As Currency

Public VG_OTORCAPNOINT As Double
Public VG_OTORINTNOCAP As Double
Public VG_OTORCAPMASINT As Double

'Esta tasa se usa para moras mayores a 60 dias sino se usa la tasa anual
Public VG_TASAMORA2 As Double

Public VG_NOAPLICAROTREFIN As Boolean

'usada para pagofacil
Public VG_CODIGOAUTOMATICO As Long

'inicio de posicion de lectura en archivos rapipago y pagofacil
Public VG_INICIORP As Long
Public VG_INICIOPF As Long

'numero de empresa credimaco ante rapipago (753)
'que va dentro del codigo de barras(no es la extension)
Public VG_NUMEMPRESA As Long

Public TipoEdicion As String

Public ValorPorcentaje As Integer

Public minutosLogout As Integer


' Para leer del .ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    
    
Sub Main()
Dim EstaIni     As String
Dim NombreIni   As String
Dim sql         As String
Dim bOpenOK     As Boolean
Dim nSecuencia  As Integer
On Error GoTo merror

'obtengo el nombre del ini
NombreIni = App.Path + "\configuracion.ini"

EstaIni = Dir(NombreIni)
If Trim(EstaIni) = "" Then
      MsgE "Falta el archivo CONFIGURACION.INI (debe ubicarse dentro de la carpeta donde tenga instalado el sistema)"
   'se va del sistema
   End
End If
bOpenOK = False
nSecuencia = 16
Do While nSecuencia <= 50 And Not bOpenOK
    bOpenOK = OPEN_MODULE(NombreIni, nSecuencia)
    nSecuencia = nSecuencia + 1
Loop

If bOpenOK Then
   Call RefrescarOpcionesSistema
   'El nombre del manual de ayuda originalmente era [Manual-Credit-Click.CHM]
   'ahora lo hemos cambiado a: Manual-Credimaco.CHM
   App.HelpFile = App.Path & "\AYUDA\Manual-Credimaco.chm"
   
   'llamo a la pantalla de login
   Call CenterForm(FrmLogin)
   FrmLogin.Show vbModal
   
    'si el login fue exitoso
   If VG_LOGIN Then
      '***TABLAS NUEVAS CREADAS POR CODIGO***
      'tabla historica nueva
      'si no existe creo la tabla por codigo
      If Not ExisteTabla("pagofacilhistorico2") Then
         'inicio transaccion
         cnSQL.BeginTrans
            
         sql = "CREATE TABLE pagofacilhistorico2 " & _
               "(NombreArchivo CHAR(50),FechaArchivo DATETIME," & _
               "FechaProceso DATETIME,Usuario CHAR(50))"
         cnSQL.Execute sql
        
        'fin de transaccion
         cnSQL.CommitTrans
      End If
      
      'tabla pagofacil2 nueva
      If Not ExisteTabla("pagofacil2") Then
         'inicio transaccion
         cnSQL.BeginTrans
            
         sql = "CREATE TABLE pagofacil2 " & _
               "(IdPagoFacil INTEGER NOT NULL," & _
               "NombreArchivo CHAR(59),CodPrestamo CHAR(50),IdCredito INTEGER,FechaCobro DATETIME," & _
               "ImporteCobro MONEY,Cliente CHAR(50),Negocio CHAR(50)," & _
               "Recibo CHAR(50),FechaImportacion DATETIME,Excedentes CHAR(50))"
         cnSQL.Execute sql
         
         sql = "ALTER TABLE [dbo].[pagofacil2] WITH NOCHECK ADD CONSTRAINT [aaaaapagofacil2_PK] PRIMARY KEY  NONCLUSTERED  ([IdPagoFacil])  ON [PRIMARY]"
         cnSQL.Execute sql

        'fin de transaccion
         cnSQL.CommitTrans
      End If
      
      'tabla pagofaciltemp2
      If Not ExisteTabla("pagofaciltemp2") Then
         'inicio transaccion
         cnSQL.BeginTrans
            
         sql = "CREATE TABLE pagofaciltemp2 " & _
               "(NombreArchivo CHAR(50),Negocio CHAR(50),Cliente CHAR(50)," & _
               "CodPrestamo CHAR(50),FechaCobro DATETIME," & _
               "ImporteCobro MONEY,Recibo CHAR(50))"
         cnSQL.Execute sql
           
        'fin de transaccion
         cnSQL.CommitTrans
      End If
      
      'si no existe creo la tabla por codigo
      If Not ExisteTabla("templistadeudores") Then
         'inicio transaccion
         cnSQL.BeginTrans
            
         sql = "CREATE TABLE templistadeudores " & _
               "(idcliente INTEGER,numlegajo CHAR(50),cliente CHAR(255)," & _
               "sexo CHAR(1),numdocumento CHAR(50),telefono CHAR(50)," & _
               "idcredito INTEGER,codprestamo CHAR(50),fechacredito DATETIME,provcredito CHAR(255)," & _
               "domicilio CHAR(50),localidad CHAR(50),provincia CHAR(50),codigopostal CHAR(50)," & _
               "nacionalidad CHAR(50),fechanacimiento CHAR(50)," & _
               "capital MONEY,numcuotas INTEGER,cuotasenmora INTEGER," & _
               "deudareal MONEY,diasmora INTEGER,maxdiasmora INTEGER,importemora MONEY,ivamora MONEY," & _
               "saldoenmora MONEY,saldocredito MONEY," & _
               "observaciones CHAR(255),Vencimiento DATETIME," & _
               "comercio CHAR(50)," & _
               "num1 INTEGER,num2 INTEGER,num3 INTEGER," & _
               "mon1 MONEY,mon2 MONEY,mon3 MONEY," & _
               "cad1 CHAR(50),cad2 CHAR(50),cad3 CHAR(50)," & _
               "logic1 BIT,logic2 BIT,logic3 BIT," & _
               "fecha1 DATETIME,fecha2 DATETIME,fecha3 DATETIME)"
               
         cnSQL.Execute sql
         
                         
        'fin de transaccion
         cnSQL.CommitTrans
      End If
      
      'si no existe creo la tabla por codigo
      If Not ExisteTabla("comercios") Then
         'inicio transaccion
         cnSQL.BeginTrans
            
         sql = "CREATE TABLE comercios " & _
               "(idcomercio INTEGER CONSTRAINT MyFieldConstraint PRIMARY KEY,nombre CHAR(50),predeterminada BIT," & _
               "cad1 CHAR(255),cad2 CHAR(255),cad3 CHAR(255)," & _
               "num1 INTEGER,num2 INTEGER,num3 INTEGER," & _
               "mon1 MONEY,mon2 MONEY,mon3 MONEY," & _
               "logic1 BIT,logic2 BIT,logic3 BIT," & _
               "fecha1 DATETIME,fecha2 DATETIME,fecha3 DATETIME)"
         cnSQL.Execute sql
           
        'fin de transaccion
         cnSQL.CommitTrans
      End If
      
      'nuevo agregado en mayo 2011
      'si no existe creo la tabla por codigo
      If Not ExisteTabla("vendedores") Then
         'inicio transaccion
         cnSQL.BeginTrans
            
         sql = "CREATE TABLE vendedores " & _
               "(idvendedor INTEGER NOT NULL,nombre CHAR(50),predeterminada BIT," & _
               "documento CHAR(50),direccion CHAR(50),telefono CHAR(50)," & _
               "ciudad CHAR(50),num1 INTEGER,num2 INTEGER,num3 INTEGER," & _
               "mon1 MONEY,mon2 MONEY,mon3 MONEY," & _
               "logic1 BIT,logic2 BIT,logic3 BIT," & _
               "fecha1 DATETIME,fecha2 DATETIME,fecha3 DATETIME);"
         cnSQL.Execute sql
         
         sql = "ALTER TABLE [dbo].[vendedores] WITH NOCHECK ADD CONSTRAINT [DF__vendedore__prede__6F7F8B4B] DEFAULT (0) FOR [predeterminada],    CONSTRAINT [DF__vendedore__logic__7073AF84] DEFAULT (0) FOR [logic1],    CONSTRAINT [DF__vendedore__logic__7167D3BD] DEFAULT (0) FOR [logic2],    CONSTRAINT [DF__vendedore__logic__725BF7F6] DEFAULT (0) FOR [logic3],    CONSTRAINT [aaaaavendedores_PK] PRIMARY KEY  NONCLUSTERED    (        [IdVendedor]    )  ON [PRIMARY]"
         cnSQL.Execute sql
         
        'fin de transaccion
         cnSQL.CommitTrans
      End If
      
      
      MDIPrincipal.Show
   Else
      Call CLOSE_MODULE
      End
   End If
Else
   'si hubo problemas con el openmodule se va del sistema
   End
End If
        
Exit Sub
merror:
tratarerrores "Error ingresando al sistema (Main)"
End Sub

Function OPEN_MODULE(archivo_ini As String, nSecuencia As Integer) As Boolean

Dim cDriver     As String
Dim cServer     As String
Dim cBase       As String
Dim cUsuario    As String
Dim cPassword   As String
Dim rsSQL       As rdoResultset

       
'la base de datos antes se llamaba "credit-click.mdb" pero ahora debe cambiar
'y pasar a llamarse por ej "CREDITOS.mdb"

'VG_NOMBREBASE = App.Path + "\database\creditos.mdb"
   
'VG_USUARIOBASE = "Admin"
'VG_CLAVEBASE = ""
'MOTOR = dbUseJet

On Error GoTo merror

Set enSQL = rdoEngine.rdoEnvironments(0)
cAmbiente = Leer_Ini(archivo_ini, "AMBIENTE", "XXX")
cServer = Leer_Ini(archivo_ini, "SERVER", "XXX")

If cAmbiente = "DESA" Then
'    cServer = "XPSP2\SQLEXPRESS"
    cBase = "creditosSQL"
    cUsuario = "UserCreditosSQL"
    cPassword = "River20" & nSecuencia
ElseIf cAmbiente = "DESAW" Then
    'cServer = "1025S90\SQLEXPRESS" 'Mariano
    cBase = "creditosSQL"
ElseIf cAmbiente = "DESAN" Then
'    cServer = "PERSONALREG" 'Notebook
    cBase = "CreditosSQL"
    cUsuario = "UserCreditosSQL"
    cPassword = "River20" & nSecuencia
ElseIf cAmbiente = "TEST" Then
'    cServer = "CREDIMACO\SQLEXPRESS" 'Ambiente de Test de Credimaco
    cBase = "CreditosSQL_TEST"
    cUsuario = "UserCreditosSQL"
    cPassword = "River20" & nSecuencia
ElseIf cAmbiente = "VIRT" Then
'    cServer = "VM-1560740" 'Servidor Virtual IPLAN
    cBase = "CreditosSQL"
    cUsuario = "UserCreditosSQL"
    cPassword = "River2000" & nSecuencia
ElseIf cAmbiente = "TELE" Then
'    cServer = "CREDIMACO\SQLEXPRESS"
    cBase = "creditosSQL"
    cUsuario = "sa"
    cPassword = "River2000"
Else
'    cServer = "CREDIMACO01\SQLEXPRESS"
    cBase = "creditosSQL"
    cUsuario = "sa"
    cPassword = "River2000"
End If

cDriver = "{SQL Server}"
cConexion = "driver=" & cDriver & ";server=" & cServer & ";database=" & cBase & ";uid=" & cUsuario & ";pwd=" & cPassword
Set cnSQL = enSQL.OpenConnection("", rdDriverNoPrompt, False, cConexion)

'creo el espacio de trabajo jet (esta linea va en ambos casos jet o net)
'Set wrkODBC = CreateWorkspace("MiWorkspace", VG_USUARIOBASE, VG_CLAVEBASE, MOTOR)

'conecto a jet con red
'con el segundo parametro logico en true abre en modo exclusivo
'con el segundo parametro logico en false es para (modo compartido)
'Set Db = cnsql.OpenDatabase(VG_NOMBREBASE, False, False, "Ms Access; pwd=bche2008k")

'CONEXION = Db.Connect

'Db.QueryTimeout = 360
OPEN_MODULE = True
        
Exit Function
merror:
OPEN_MODULE = False
End Function
Public Sub CLOSE_MODULE()
On Error Resume Next

If Not (Db Is Nothing) Then Db.Close
If Not (wrkODBC Is Nothing) Then cnSQL.Close
        
Set Db = Nothing
Set wrkODBC = Nothing

On Error GoTo 0
End Sub
Public Sub AbrirBase()
On Error GoTo merror
 
'abre la base de datos y el espacio de trabajo
'Set wrkODBC = CreateWorkspace("ODBCWorkspace", VG_USUARIOBASE, VG_CLAVEBASE, MOTOR)
 
'conecto a jet
'Set Db = cnsql.OpenDatabase(VG_NOMBREBASE, False, False, "Ms Access; pwd=bche2008k")

'CONEXION = Db.Connect
'Db.QueryTimeout = 360

Exit Sub
merror:
tratarerrores "Error en procedimiento AbrirBase"
End Sub

Private Function Leer_Ini(Path_INI As String, Key As String, Default As Variant) As String
Const APPLICATION As String = "Datos_Programa"
Dim bufer As String * 256
Dim Len_Value As Long
  
    Len_Value = GetPrivateProfileString(APPLICATION, _
                                         Key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         Path_INI)
          
    Leer_Ini = Left$(bufer, Len_Value)
  
End Function
