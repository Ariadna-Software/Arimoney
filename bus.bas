Attribute VB_Name = "bus"
  Option Explicit



Public vUsu As Usuario  'Datos usuario
Public vConfig As Configuracion
Public vParam As Cparametros
Public vEmpresa As Cempresa

'Formato de fecha
Public FormatoFecha As String
Public FormatoImporte As String

Public CadenaDesdeOtroForm As String
'Public DB As Database
Public Conn As ADODB.Connection


'Global para nº de registro eliminado
Public NumRegElim  As Long

'Para algunos campos de texto suletos controlarlos
Public miTag As CTag

'Variable para saber si se ha actualizado algun asiento
Public AlgunAsientoActualizado As Boolean
Public TieneIntegracionesPendientes As Boolean

Public miRsAux As ADODB.Recordset

Public AnchoLogin As String  'Para fijar los anchos de columna






'/////////////////////////////////////////////////////////////////
'// Se trata de identificar el PC en la BD. Asi conseguiremos tener
'// los nombres de los PC para poder asignarles un codigo
'// UNa vez asignado el codigo  se lo sumaremos (x 1000) al codusu
'// con lo cual el usuario sera distinto( aunque sea con el mismo codigo de entrada)
'// dependiendo desde k PC trabaje

Public Sub GestionaPC()
CadenaDesdeOtroForm = ComputerName
If CadenaDesdeOtroForm <> "" Then
    FormatoFecha = DevuelveDesdeBD("codpc", "Usuarios.pcs", "nompc", CadenaDesdeOtroForm, "T")
    If FormatoFecha = "" Then
        NumRegElim = 0
        FormatoFecha = "Select max(codpc) from Usuarios.pcs"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open FormatoFecha, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            NumRegElim = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        NumRegElim = NumRegElim + 1
        If NumRegElim > 32000 Then
            MsgBox "Error en numero de PC's activos. Demasiados PC en BD. Llame a soporte técnico.", vbCritical
            End
        End If
        FormatoFecha = "INSERT INTO Usuarios.pcs (codpc, nompc) VALUES (" & NumRegElim & ", '" & CadenaDesdeOtroForm & "')"
        Conn.Execute FormatoFecha
    End If
End If
End Sub


Public Sub OtrasAcciones()
On Error Resume Next

    FormatoFecha = "yyyy-mm-dd"
    FormatoImporte = "#,###,###,##0.00"


    'Borramos uno de los archivos temporales
    If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
    
    
    'Borramos tmp bloqueos
    'Borramos temporal
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    NumRegElim = Len(CadenaDesdeOtroForm)
    If NumRegElim = 0 Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = " WHERE codusu = " & vUsu.Codigo
    End If
    Conn.Execute "Delete from zBloqueos " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
    
    'Buscamos si hay integraciones pendientes. Lo almacenamios en la variable global: TieneIntegracionesPendientes
    TieneIntegracionesPendientes = BuscarIntegraciones(False, "??")
    
End Sub


'Usuario As String, Pass As String --> Directamente el usuario
Public Function AbrirConexion() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
    
    
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vUsuarios;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER

    
    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password
    '---- Laura: 29/09/2006
    Cad = Cad & ";PORT=3306;OPTION=3;STMT=;"
    
    Cad = Cad & "Persist Security Info=true"
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Conn.ConnectionString = Cad
    Conn.Open
    Conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión.", Err.Description
End Function




Public Function AbrirConexionUsuarios() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionUsuarios = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient
    Conn.CursorLocation = adUseServer
    
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=USUARIOS;DATA SOURCE= vUsuarios;SERVER=" & vConfig.SERVER
    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password
    Cad = Cad & ";PORT=3306;OPTION=3;STMT=;"
    
    
    Conn.ConnectionString = Cad
    Conn.Open
    AbrirConexionUsuarios = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión usuarios.", Err.Description
End Function


'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    Conn.Execute "commit"
    Conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    Conn.Execute "commit"
    Conn.Execute "set autocommit=1"
End Sub


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosComas(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ".")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "," & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaPuntosComas = CADENA
End Function


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ",")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "." & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaComasPuntos = CADENA
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosHoras(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ".")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & ":" & Mid(CADENA, I + 1)
        End If
    Loop Until I = 0
    TransformaPuntosHoras = CADENA
End Function


Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"
                    DBLet = ""
                Case "N"
                    DBLet = 0
                Case "F"
                    DBLet = "0:00:00"
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function



Public Function DBSet1(ByRef T As TextBox, Tipo As String) As String
    If T.Text <> "" Then
        
        
            Select Case Tipo
                
                Case "N"
                    DBSet1 = Val(T.Text)
                Case "F"
                    DBSet1 = "'" & Format(T.Text, FormatoFecha) & "'"
                Case "D"  'DECIMAL , moneda
                    If InStr(1, T.Text, ",") > 0 Then
                        'Importe formateado
                        DBSet1 = TransformaComasPuntos(CStr(ImporteFormateado(T.Text)))
                    Else
                        'Sin formatear
                        DBSet1 = T.Text
                    End If
                Case Else
                    'Case "T"
                    DBSet1 = "'" & DevNombreSQL(T.Text) & "'"
            End Select
    Else
        DBSet1 = "NULL"
    End If
End Function


'MODIFICADO. Conta nueva. Ambito fechas
'  12 MAYO 2008
Public Function FechaCorrecta2(vFecha As Date, MostrarMensaje As Boolean) As Byte
Dim Mens As String
'--------------------------------------------------------
'   Dada una fecha dira si pertenece o no
'   al intervalo de fechas que maneja la apliacion
'   Resultados:
'       0 .- Año actual
'       1 .- Siguiente
'       2 .- Ambito fecha. Fecha menor a la del ambito !!!!! NUEVO !!!!
'       3 .- Anterior al inicio
'       4 .- Posterior al fin
'--------------------------------------------------------
    
    If vFecha >= vParam.fechaini Then
        'Mayor que fecha inicio
        If vFecha >= vParam.fechaAmbito Then
            If vFecha <= vParam.fechafin Then
                FechaCorrecta2 = 0
            Else
                'Compruebo si el año siguiente
                If vFecha <= DateAdd("yyyy", 1, vParam.fechafin) Then
                    FechaCorrecta2 = 1
                Else
                    FechaCorrecta2 = 4   'Fuera ejercicios
                    Mens = "mayor que fin ejercicios"
                End If
            End If
        Else
            Mens = "menor que fecha activa"
            FechaCorrecta2 = 2   'Menor que fecha actvia
        End If
    Else            '< fecha ini
        FechaCorrecta2 = 3
        Mens = "anterior al inicio de ejercicios"
    End If
    
    If FechaCorrecta2 > 1 Then
        If MostrarMensaje Then
            Mens = "Fecha " & Mens & ". Fecha: " & vFecha
            MsgBox Mens, vbExclamation
        End If
    End If
End Function






Public Sub MuestraError(numero As Long, Optional CADENA As String, Optional Desc As String)
    Dim Cad As String
    Dim Aux As String
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    Cad = "Se ha producido un error: " & vbCrLf
    If CADENA <> "" Then
        Cad = Cad & vbCrLf & CADENA & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If Conn.Errors.Count > 0 Then
        ControlamosError Aux
        Conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then Cad = Cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then
        If numero <> 513 Then Cad = Cad & "Número: " & numero & vbCrLf & "Descripción: " & Error(numero)
    End If
    MsgBox Cad, vbExclamation
End Sub

Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function


Public Function RellenaCodigoCuenta(vCodigo As String) As String
    Dim I As Integer
    Dim J As Integer
    Dim CONT As Integer
    Dim Cad As String

    RellenaCodigoCuenta = vCodigo
    If Len(vCodigo) > vEmpresa.DigitosUltimoNivel Then Exit Function
    I = 0: CONT = 0
    Do
        I = I + 1
        I = InStr(I, vCodigo, ".")
        If I > 0 Then
            If CONT > 0 Then CONT = 1000
            CONT = CONT + I
        End If
    Loop Until I = 0

    'Habia mas de un punto
    If CONT > 1000 Or CONT = 0 Then Exit Function

    'Cambiamos el punto por 0's  .-Utilizo la variable maximocaracteres, para no tener k definir mas
    I = Len(vCodigo) - 1 'el punto lo quito
    J = vEmpresa.DigitosUltimoNivel - I
    Cad = ""
    For I = 1 To J
        Cad = Cad & "0"
    Next I

    Cad = Mid(vCodigo, 1, CONT - 1) & Cad
    Cad = Cad & Mid(vCodigo, CONT + 1)
    RellenaCodigoCuenta = Cad
End Function


Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef OtroCampo As String) As String
    Dim RS As Recordset
    Dim Cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    Cad = "Select " & kCampo
    If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
    Cad = Cad & " FROM " & Ktabla
    Cad = Cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        Cad = Cad & ValorCodigo
    Case "T", "F"
        Cad = Cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        DevuelveDesdeBD = DBLet(RS.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(RS.Fields(1))
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function

'Obvio
Public Function EsCuentaUltimoNivel(Cuenta As String) As Boolean
    EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresa.DigitosUltimoNivel)
End Function


Public Function CuentaCorrectaUltimoNivelTXT(TCta As TextBox, TDesc As TextBox) As Boolean
Dim C1 As String
Dim C2 As String

    C1 = TCta.Text
    C2 = ""
    CuentaCorrectaUltimoNivelTXT = CuentaCorrectaUltimoNivel(C1, C2)
    TCta.Text = C1
    TDesc.Text = C2
End Function

Public Function CuentaCorrectaUltimoNivel(ByRef Cuenta As String, ByRef Devuelve As String) As Boolean
'Comprueba si es numerica
Dim SQL As String

CuentaCorrectaUltimoNivel = False
If Cuenta = "" Then
    Devuelve = "Cuenta vacia"
    Exit Function
End If
If Not IsNumeric(Cuenta) Then
    Devuelve = "La cuenta debe de ser numérica: " & Cuenta
    Exit Function
End If

'Rellenamos si procede
Cuenta = RellenaCodigoCuenta(Cuenta)

If Not EsCuentaUltimoNivel(Cuenta) Then
    Devuelve = "No es cuenta de último nivel: " & Cuenta
    Exit Function
End If


SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
If SQL = "" Then
    Devuelve = "No existe la cuenta : " & Cuenta
    Exit Function
End If

'Llegados aqui, si que existe la cuenta
CuentaCorrectaUltimoNivel = True
Devuelve = SQL
End Function


'-------------------------------------------------------------------------
''
''   Es la misma solo k no si no existe cuenta no da error
Public Function CuentaCorrectaUltimoNivelSIN(ByRef Cuenta As String, ByRef Devuelve As String) As Byte
'Comprueba si es numerica
Dim SQL As String

CuentaCorrectaUltimoNivelSIN = 0
If Cuenta = "" Then
    Devuelve = "Cuenta vacia"
    Exit Function
End If
If Not IsNumeric(Cuenta) Then
    Devuelve = "La cuenta debe de ser numérica: " & Cuenta
    Exit Function
End If

'Rellenamos si procede
Cuenta = RellenaCodigoCuenta(Cuenta)

CuentaCorrectaUltimoNivelSIN = 1
If Not EsCuentaUltimoNivel(Cuenta) Then
    SQL = "No es cuenta de último nivel"
Else
    SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
    If SQL = "" Then
        SQL = "No existe la cuenta  "
    Else
        CuentaCorrectaUltimoNivelSIN = 2
    End If
End If

'Llegados aqui, si que existe la cuenta
Devuelve = SQL
End Function


Public Function CuentaBloqeada(Cuenta As String, Fecha As Date, MostrarMensaje As Boolean) As Boolean
Dim RS As ADODB.Recordset

    On Error GoTo ECtaB
    CuentaBloqeada = False
    Set RS = New ADODB.Recordset
    RS.Open "Select fecbloq from cuentas where codmacta = '" & Cuenta & "'", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS!FecBloq) Then
            If RS!FecBloq <= Fecha Then
                CuentaBloqeada = True
                If MostrarMensaje Then _
                    MsgBox "Cuenta bloqueada: " & Cuenta & " -  Fecha: " & Format(RS!FecBloq, "dd/mm/yyyy"), vbExclamation
            End If
        End If
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
ECtaB:
    MuestraError Err.Number
End Function




'Devuelve, para un nivel determinado, cuantos digitos tienen las cuentas
' a ese nivel
Public Function DigitosNivel(numnivel As Integer) As Integer
    Select Case numnivel
    Case 1
        DigitosNivel = vEmpresa.numdigi1

    Case 2
        DigitosNivel = vEmpresa.numdigi2

    Case 3
        DigitosNivel = vEmpresa.numdigi3

    Case 4
        DigitosNivel = vEmpresa.numdigi4

    Case 5
        DigitosNivel = vEmpresa.numdigi5

    Case 6
        DigitosNivel = vEmpresa.numdigi6

    Case 7
        DigitosNivel = vEmpresa.numdigi7

    Case 8
        DigitosNivel = vEmpresa.numdigi8

    Case 9
        DigitosNivel = vEmpresa.numdigi9

    Case 10
        DigitosNivel = vEmpresa.numdigi10

    Case Else
        DigitosNivel = -1
    End Select
End Function

Public Function NivelCuenta(CodigoCuenta As String) As Integer
Dim lon As Integer
Dim niv As Integer
Dim I As Integer
    NivelCuenta = -1
    lon = Len(CodigoCuenta)
    I = 0
    Do
       I = I + 1
       niv = DigitosNivel(I)
       If niv > 0 Then
            If niv = lon Then
                NivelCuenta = I
                I = 11 'para salir del bucle
            End If
        Else
            I = 11 'salimos pq ya no hay nveles para las cuentas de longitud lon
        End If
    Loop Until I > 10
End Function



'Public Function ExistenSubcuentas(ByRef Cuenta As String, Nivel As Integer) As Boolean
'Dim i As Integer
'Dim B As Boolean
'Dim cad As String
'
'    i = DigitosNivel(Nivel)
'    cad = Mid(Cuenta, 1, i)
'    cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", cad, "T")
'    If cad = "" Then
'        'NO existe la subcuenta de nivel N
'        'salimos
'        ExistenSubcuentas = False
'        Exit Function
'    End If
'    If Nivel > 1 Then
'        ExistenSubcuentas = ExistenSubcuentas(Cuenta, Nivel - 1)
'    Else
'        ExistenSubcuentas = True
'    End If
'End Function
'
'
'Public Function CreaSubcuentas(ByRef Cuenta, HastaNivel As Integer, TEXTO As String) As Boolean
'Dim i As Integer
'Dim J As Integer
'Dim cad As String
'Dim Cta As String
'
'On Error GoTo ECreaSubcuentas
'CreaSubcuentas = False
'For i = 1 To HastaNivel
'    J = DigitosNivel(i)
'    Cta = Mid(Cuenta, 1, J)
'    cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cta, "T")
'    If cad = "" Then
'        'CreaCuenta
'        cad = "INSERT INTO cuentas (codmacta, nommacta, apudirec, model347, razosoci, "
'        cad = cad & " dirdatos, codposta, despobla, desprovi, nifdatos, maidatos, webdatos,"
'        cad = cad & " obsdatos) VALUES ("
'        cad = cad & " '" & Cta
'        cad = cad & " ', '" & TEXTO
'        cad = cad & " ', "
'        cad = cad & " 'N', 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
'        Conn.Execute cad
'    End If
'Next i
'CreaSubcuentas = True
'Exit Function
'ECreaSubcuentas:
'    MuestraError Err.Number, "Creando subcuentas", Err.Description
'End Function
'



Public Function CambiarBarrasPATH(ParaGuardarBD As Boolean, CADENA) As String
Dim I As Integer
Dim Ch As String
Dim Ch2 As String

If ParaGuardarBD Then
    Ch = "\"
    Ch2 = "/"
Else
    Ch = "/"
    Ch2 = "\"
End If
I = 0
Do
    I = I + 1
    I = InStr(1, CADENA, Ch)
    If I > 0 Then CADENA = Mid(CADENA, 1, I - 1) & Ch2 & Mid(CADENA, I + 1)
Loop Until I = 0
CambiarBarrasPATH = CADENA
End Function


Public Function ImporteSinFormato(CADENA As String) As String
Dim I As Integer
'Quitamos puntos
Do
    I = InStr(1, CADENA, ".")
    If I > 0 Then CADENA = Mid(CADENA, 1, I - 1) & Mid(CADENA, I + 1)
Loop Until I = 0
ImporteSinFormato = TransformaPuntosComas(CADENA)
End Function



''Periodo vendran las fechas Ini y fin con pipe final
'Public Sub SaldoHistorico(Cuenta As String, Periodo As String, DescCuenta As String)
'Dim RS As Recordset
'Dim SQL As String
'Dim RC2 As String
'    Screen.MousePointer = vbHourglass
'    SQL = "Select Sum(timporteD),sum(timporteH) from hlinapu"
'    SQL = SQL & " WHERE codmacta='" & Cuenta & "'"
'    SQL = SQL & " AND fechaent>='" & Format(vParam.fechaini, FormatoFecha) & "' AND punteada "
'    Set RS = New ADODB.Recordset
'    RC2 = Cuenta & "|"
'    'PUNTEADO
'    RS.Open SQL & "='S';", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not RS.EOF Then
'       RC2 = RC2 & Format(RS.Fields(0), FormatoImporte) & "|"
'       RC2 = RC2 & Format(RS.Fields(1), FormatoImporte) & "|"
'    Else
'        RC2 = RC2 & "||"
'    End If
'    RS.Close
'    'SIN puntear
'    RS.Open SQL & "<>'S';", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not RS.EOF Then
'       RC2 = RC2 & Format(RS.Fields(0), FormatoImporte) & "|"
'       RC2 = RC2 & Format(RS.Fields(1), FormatoImporte) & "|"
'    Else
'        RC2 = RC2 & "||"
'    End If
'    RS.Close
'
'    'En el periodo. Para cuando viene de puntear
'    If Periodo <> "" Then
'        SQL = "Select Sum(timporteD) - sum(timporteH) from hlinapu"
'        SQL = SQL & " WHERE codmacta='" & Cuenta & "' AND "
'        SQL = SQL & Periodo
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        If Not RS.EOF Then
'            RC2 = RC2 & Format(RS.Fields(0), FormatoImporte) & "|"
'        Else
'            RC2 = RC2 & "|"
'        End If
'    Else
'        RC2 = RC2 & "|"
'    End If
'    RC2 = RC2 & DescCuenta & "|"
'    Set RS = Nothing
'    'Mostramos la ventanita de mesaje
'    frmMensajes.Opcion = 1
'    frmMensajes.Parametros = RC2
'    frmMensajes.Show vbModal
'
'End Sub

'Lo que hace es comprobar que si la resolucion es mayor
'que 800x600 lo pone en el 400
Public Sub AjustarPantalla(ByRef formulario As Form)
    If Screen.Width > 13000 Then
        formulario.Top = 400
        formulario.Left = 400
    Else
        formulario.Top = 0
        formulario.Left = 0
    End If
    formulario.Width = 12000
    formulario.Height = 9000
End Sub


'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256.98
'   Tiene que venir numérico
Public Function ImporteFormateado(Importe As String) As Currency
Dim I As Integer

If Importe = "" Then
    ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateado = Importe
End If
End Function





Public Function DiasMes(Mes As Byte, Anyo As Integer) As Integer
    Select Case Mes
    Case 2
        If (Anyo Mod 4) = 0 Then
            DiasMes = 29
        Else
            DiasMes = 28
        End If
    Case 1, 3, 5, 7, 8, 10, 12
        DiasMes = 31
    Case Else
        DiasMes = 30
    End Select
End Function





Public Function ComprobarEmpresaBloqueada(codusu As Integer, ByRef Empresa As String) As Boolean
Dim Cad As String

ComprobarEmpresaBloqueada = False

'Antes de nada, borramos las entradas de usuario, por si hubiera kedado algo
Conn.Execute "Delete from Usuarios.vBloqBD where codusu=" & codusu

'Ahora comprobamos k nadie bloquea la BD
Cad = DevuelveDesdeBD("codusu", "Usuarios.vBloqBD", "conta", Empresa, "T")
If Cad <> "" Then
    'En teoria esta bloqueada. Puedo comprobar k no se haya kedado el bloqueo a medias
    
    Set miRsAux = New ADODB.Recordset
    Cad = "show processlist"
    miRsAux.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        If miRsAux.Fields(3) = Empresa Then
            Cad = miRsAux.Fields(2)
            miRsAux.MoveLast
        End If
    
        'Siguiente
        miRsAux.MoveNext
    Wend
    
    If Cad = "" Then
        'Nadie esta utilizando la aplicacion, luego se puede borrar la tabla
        Conn.Execute "Delete from Usuarios.vBloqBD where conta ='" & Empresa & "'"
        
    Else
        MsgBox "BD bloqueada.", vbCritical
        ComprobarEmpresaBloqueada = True
    End If
End If

Conn.Execute "commit"
End Function


Public Function Bloquear_DesbloquearBD(Bloquear As Boolean) As Boolean

On Error GoTo EBLo
    Bloquear_DesbloquearBD = False
    If Bloquear Then
        CadenaDesdeOtroForm = "INSERT INTO Usuarios.vBloqBD (codusu, conta) VALUES (" & vUsu.Codigo & ",'" & vUsu.CadenaConexion & "')"
    Else
        CadenaDesdeOtroForm = "DELETE FROM  Usuarios.vBloqBD WHERE codusu =" & vUsu.Codigo & " AND conta = '" & vUsu.CadenaConexion & "'"
    End If
    Conn.Execute CadenaDesdeOtroForm
    Bloquear_DesbloquearBD = True
    Exit Function
EBLo:
    'MuestraError Err.Number, "Bloq. BD"
    Err.Clear
End Function




Public Function OtrosPCsContraContabiliad() As String
Dim MiRS As Recordset
Dim Cad As String
Dim Equipo As String
Dim EquipoConBD As Boolean


    Set MiRS = New ADODB.Recordset
    EquipoConBD = (vUsu.PC = vConfig.SERVER Or LCase(vConfig.SERVER) = "localhost")
    Cad = "show processlist"
    MiRS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not MiRS.EOF
        If UCase(MiRS.Fields(3)) = UCase(vUsu.CadenaConexion) Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vUsu.PC Then
                    
                    NumRegElim = 0
                    If Equipo <> "LOCALHOST" Then
                        'Si no es localhost
                        NumRegElim = 1
                    Else
                        'HAy un proceso de loclahost. Luego, si el equipo no tiene la BD
                        If Not EquipoConBD Then NumRegElim = 1
                    End If
                    
                    
                    
                    'Si hay que insertar
                    If NumRegElim = 1 Then
                        If InStr(1, Cad, Equipo & "|") = 0 Then Cad = Cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraContabiliad = Cad
End Function


Public Function EsNumerico(TEXTO As String) As Boolean
Dim I As Integer
Dim C As Integer
Dim L As Integer
Dim Cad As String
    
    EsNumerico = False
    Cad = ""
    If Not IsNumeric(TEXTO) Then
        Cad = "El campo debe ser numérico"
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            I = InStr(L, TEXTO, ".")
            If I > 0 Then
                L = I + 1
                C = C + 1
            End If
        Loop Until I = 0
        If C > 1 Then Cad = "Numero de puntos incorrecto"
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                I = InStr(L, TEXTO, ",")
                If I > 0 Then
                    L = I + 1
                    C = C + 1
                End If
            Loop Until I = 0
            If C > 1 Then Cad = "Numero incorrecto"
        End If
        
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
    Else
        EsNumerico = True
    End If
End Function



Public Function EsFechaOK(ByRef T As TextBox) As Boolean
Dim Cad As String
    
    Cad = T.Text
    If InStr(1, Cad, "/") = 0 Then
        If Len(T.Text) = 8 Then
            Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        Else
            If Len(T.Text) = 6 Then Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/20" & Mid(Cad, 5)
        End If
    End If
    
    If IsDate(Cad) Then
        EsFechaOK = True
        T.Text = Format(Cad, "dd/mm/yyyy")
    Else
        EsFechaOK = False
    End If
End Function



Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim Cad As String
    
    Cad = T
    If InStr(1, Cad, "/") = 0 Then
        If Len(T) = 8 Then
            Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        Else
            If Len(T) = 6 Then Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/20" & Mid(Cad, 5)
        End If
    End If
    If IsDate(Cad) Then
        EsFechaOKString = True
        T = Format(Cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function

'Devuelve si hay archivos
'                                                        Llevara la forma: 01, 02  para la empresa 1 o 2..
Public Function BuscarIntegraciones(Errores As Boolean, Empresa As String) As Boolean
Dim Cad As String
On Error GoTo Ebuscarintegraciones
    
    'Exit Function
    BuscarIntegraciones = False
    If vConfig.Integraciones = "" Then Exit Function
    
    Cad = vConfig.Integraciones
    If Right(Cad, 1) <> "\" Then Cad = Cad & "\"
    If Dir(Cad, vbDirectory) = "" Then
        MsgBox "Carpeta de errores no encontrada: " & vConfig.Integraciones, vbExclamation
        Exit Function
    End If
    
    If Errores Then
        Cad = Cad & "ERRORES"
    Else
        Cad = Cad & "INTEGRA"
    End If
    
    
    'Facturas clientes
'    If Dir(cad & "\FRACLI\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If
'
'    'Facturas Proveedores
'    If Dir(cad & "\FRAPRO\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If
'
'    'Asientos al diario
'    If Dir(cad & "\ASIDIA\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If
'
'    'Asientos al historico
'    If Dir(cad & "\ASIHCO\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If



    If Dir(Cad & "\SCOBRO\*.?" & Empresa) <> "" Then
        BuscarIntegraciones = True
        Exit Function
    End If
    
    'Facturas Proveedores
    If Dir(Cad & "\SPAGO\*.?" & Empresa) <> "" Then
        BuscarIntegraciones = True
        Exit Function
    End If
    
    
   

    Exit Function
Ebuscarintegraciones:
    MuestraError Err.Number, Err.Description, "Buscar archivos integraciones" & vbCrLf
End Function


Public Sub ComprobarFuncionamientoEspia()
Dim I As Integer
Dim Cad As String
Dim Ruta As String
Dim F As Date
Dim TieneAntiguos As Boolean

On Error GoTo Ebuscarintegraciones
    
    
    
    TieneAntiguos = False
    For I = 1 To 2
    
        Cad = vConfig.Integraciones
        If Right(Cad, 1) <> "\" Then Cad = Cad & "\"
        Cad = Cad & "AUTOMA"

        If Dir(Cad, vbDirectory) <> "" Then
            Select Case I
            Case 1
                Ruta = Cad & "\SCOBRO"
            Case 2
                Ruta = Cad & "\SPAGO"
            
            End Select
    
            If Dir(Ruta, vbDirectory) <> "" Then
                
                Cad = Dir(Ruta & "\*.*", vbArchive)
                Do While Cad <> ""
                    F = FileDateTime(Ruta & "\" & Cad)
                    NumRegElim = Abs(DateDiff("n", Now, F))
                    'MAYOR QUE 10 MINUTOS
                    If NumRegElim > 10 Then
                        'ERROR
                        TieneAntiguos = True
                        Cad = "" 'Para que se salga sin hacer mas
                        Exit For
                    Else
                        Cad = Dir
                    End If
    
                Loop
                
            End If
            
        End If
    
    Next I

    
    If TieneAntiguos Then
        Ruta = ""
        For I = 1 To 50
            Ruta = Ruta & "*"
        Next I
        Cad = Ruta & vbCrLf & vbCrLf & vbCrLf
        Cad = Cad & "        Existen archivos en la carpeta AUTOMA pendientes de " & vbCrLf
        Cad = Cad & "        integrar (tesoreria) desde hace más de 10 minutos." & vbCrLf & vbCrLf & vbCrLf & Ruta
        MsgBox Cad, vbCritical
        
    End If
    NumRegElim = 0

    Exit Sub
Ebuscarintegraciones:
    MuestraError Err.Number, Err.Description, "Buscar archivos AUTOMA" & vbCrLf


End Sub


'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef CADENA As String)
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
End Sub

Public Function DevNombreSQL(CADENA As String) As String
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSQL = CADENA
End Function



'Para los balnces
Public Function FechaInicioIGUALinicioEjerecicio(FecIni As Date, EjerciciosCerrados1 As Boolean) As Byte
Dim Fecha As Date
Dim Salir As Boolean
Dim I As Integer
On Error GoTo EfechaInicioIGUALinicioEjerecicio

    FechaInicioIGUALinicioEjerecicio = 1
    If EjerciciosCerrados1 Then
        I = -1 'En ejercicios cerrados empèzamos mirando un año por debajo fecini
    Else
        I = 1
    End If
    Fecha = DateAdd("yyyy", I, vParam.fechaini)
    Salir = False
    While Not Salir
        If FecIni = Fecha Then
            'Fecha inicio del listado contiene es fecha incio ejercicio
            FechaInicioIGUALinicioEjerecicio = 0
            Salir = True
        Else
            If FecIni < Fecha Then
                Fecha = DateAdd("yyyy", -1, Fecha)
            Else
                Salir = True
            End If
        End If
    Wend
    
    Exit Function
EfechaInicioIGUALinicioEjerecicio:
    Err.Clear  'No tiene importancia
End Function



'Public Function DevuelveDigitosNivelAnterior() As Integer
'Dim J As Integer
'    DevuelveDigitosNivelAnterior = 3
'    If vEmpresa Is Nothing Then Exit Function
'    If vEmpresa.numnivel < 2 Then Exit Function
'    J = vEmpresa.numnivel - 1
'    J = DigitosNivel(J)
'    If J < 3 Then J = 3
'    DevuelveDigitosNivelAnterior = J
'End Function



'--------------------------------------------------------------------------
' Los numeros vendran formateados o sin formatear, pero siempre viene texto
'
Public Function CadenaCurrency(TEXTO As String, ByRef Importe As Currency) As Boolean
Dim I As Integer

    On Error GoTo ECadenaCurrency
    Importe = 0
    CadenaCurrency = False
    If Not IsNumeric(TEXTO) Then Exit Function
    I = InStr(1, TEXTO, ",")
    If I = 0 Then
        'Significa k el numero no esta  formateado y como mucho tiene punto
        Importe = CCur(TransformaPuntosComas(TEXTO))
    Else
        Importe = ImporteFormateado(TEXTO)
    End If
    CadenaCurrency = True
    
    Exit Function
ECadenaCurrency:
    Err.Clear
End Function

Public Sub FormatTextImporte(ByRef T As TextBox)
Dim J As Integer
Dim Im As Currency
    If T.Text = "" Then Exit Sub
    If InStr(1, T.Text, ",") > 0 Then
        'Tiene comas.
        Do
            J = InStr(1, T.Text, ".")
            If J > 0 Then T.Text = Mid(T.Text, 1, J - 1) & Mid(T.Text, J + 1)
        Loop Until J = 0
        If Not IsNumeric(T.Text) Then
            MsgBox "Campo no numerico:" & T.Text, vbExclamation
            T.Text = ""
            Exit Sub
        End If
        Im = CCur(T.Text)
        
    
    Else
        If Not IsNumeric(T.Text) Then
            MsgBox "Campo no numerico " & T.Text, vbExclamation
            T.Text = ""
            Exit Sub
        End If
        Im = CCur(TransformaPuntosComas(T.Text))
        
    End If
    T.Text = Format(Im, FormatoImporte)
End Sub



Public Function UsuariosConectados() As Boolean
Dim I As Integer
Dim Cad As String
Dim metag As String
Dim SQL As String
Cad = OtrosPCsContraContabiliad
UsuariosConectados = False
If Cad <> "" Then
    UsuariosConectados = True
    I = 1
    'metag = "Los siguientes PC's están conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
    metag = "Los siguientes PC's están conectados :" & vbCrLf & vbCrLf
    Do
        SQL = RecuperaValor(Cad, I)
        If SQL <> "" Then
            metag = metag & "    - " & SQL & vbCrLf
            I = I + 1
        End If
    Loop Until SQL = ""
    MsgBox metag, vbExclamation
End If
End Function



Public Function BloqueoManual(Bloquear As Boolean, Tabla As String, Clave As String) As Boolean
Dim SQL As String
    If Bloquear Then
        SQL = "INSERT INTO zbloqueos (codusu, tabla, clave) VALUES (" & vUsu.Codigo
        SQL = SQL & ",'" & UCase(Tabla) & "','" & UCase(Clave) & "')"
    Else
        SQL = "DELETE FROM zbloqueos where codusu = " & vUsu.Codigo & " AND tabla ='"
        SQL = SQL & Tabla & "'"
    End If
    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        Err.Clear
        BloqueoManual = False
    Else
        BloqueoManual = True
    End If
End Function




Public Function TextoAimporte(Importe As String) As Currency
Dim I As Integer
    If Importe = "" Then
        TextoAimporte = 0
    Else
        If InStr(1, Importe, ",") > 0 Then
            'Primero quitamos los puntos
            Do
                I = InStr(1, Importe, ".")
                If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
            Loop Until I = 0
            TextoAimporte = Importe
        
        
        Else
            'No tiene comas. El punto es el decimal
            TextoAimporte = TransformaPuntosComas(Importe)
        End If
    End If

End Function



Public Sub TirarAtrasTransaccion()
    On Error Resume Next
    Conn.RollbackTrans
    If Err.Number <> 0 Then
        If Conn.Errors(0).NativeError = 1196 Then
            'NO PASA NADA. YA sabemos que las tblas tmp no se van a hacer rollbacktrans
            Conn.Cancel
            Conn.RollbackTrans
        Else
            MsgBox "Deshaciendo transacciones:" & Err.Description, vbExclamation
        End If
        Err.Clear
        Conn.Errors.Clear
        
    End If
    
End Sub








'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'
'   Imprimir listado caja.
'
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Public Function ImpirmirListadoCaja(ByRef vSQL As String, SaldoArrastrado As Boolean) As Boolean
Dim miSql As String
Dim L As Long
Dim Cad As String
Dim Caja As String
Dim CtaCaja As String
Dim Tipo As Integer
Dim RT As ADODB.Recordset

    ImpirmirListadoCaja = False
    Conn.Execute "DELETE from Usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo
    
    Set miRsAux = New ADODB.Recordset
    miSql = "Select slicaja.*,nommacta from slicaja,cuentas,susucaja where slicaja.codmacta=cuentas.codmacta " & vSQL
    miSql = miSql & " ORDER BY slicaja.codusu,feccaja,numlinea"
    miRsAux.Open miSql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    L = 1
    'INSERT INTO ztesoreriacomun (
    'codusu, fecha1,codigo, texto1, texto2,opcion, texto3, texto4, texto5, texto6,
    'importe1, importe2,   fecha3,
    'observa1, observa2) VALUES (
    
    vSQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, fecha1,codigo, texto1, texto2,texto4,opcion, texto3, observa1, "
    vSQL = vSQL & "texto5,importe1 ,importe2,texto6 ) VALUES (" & vUsu.Codigo & ",'"
    CtaCaja = ""
    While Not miRsAux.EOF
        If miRsAux!codusu <> CtaCaja Then
            CtaCaja = miRsAux!codusu
            
            Caja = DevuelveDesdeBD("nomusu", "usuarios.usuarios", "codusu", miRsAux!codusu, "N")
            Caja = DevNombreSQL(Caja)
            Caja = ",'" & CtaCaja & "','" & Caja & "'"
            
            'Si lleva saldo arrastrado entonces lo obtengo del datos de usucaja
            If SaldoArrastrado Then
                Cad = "Select saldo from susucaja where codusu =" & CtaCaja
                Set RT = New ADODB.Recordset
                RT.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                If Not RT.EOF Then
                    'Inserto una primera linea con fecha 1900 con el saldo de la caja
                    Cad = "1900-01-01'," & L & Caja
                    For Tipo = 1 To 4
                        Cad = Cad & ",NULL"
                    Next Tipo
                      Cad = Cad & ",'Saldo caja :',"
                    If RT!saldo >= 0 Then
                        Cad = Cad & TransformaComasPuntos(CStr(RT!saldo)) & ",0"
                    Else
                        Cad = Cad & "0," & TransformaComasPuntos(CStr(Abs(RT!saldo)))
                    End If
                    Cad = vSQL & Cad & ",NULL)"
                    Conn.Execute Cad
                    'Sumo L
                    L = L + 1
                End If
                RT.Close
                Set RT = Nothing
            End If
        End If
        
        Cad = Format(miRsAux!feccaja, FormatoFecha) & "'," & L & Caja
        If miRsAux!tipomovi = 1 Then
            Tipo = 1
            'FACTURAS PROVEEDORES
            Cad = Cad & ",'FRAPRO',1,'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
            'Numero de factura
            Cad = Cad & DevNombreSQL(DBLet(miRsAux!numfacpr))
            If Not IsNull(miRsAux!numvenci) Then Cad = Cad & " - Vto: " & miRsAux!numvenci
            Cad = Cad & "',"
        Else
            If miRsAux!tipomovi >= 2 Then
                'TRASPASO o PAGO
                Tipo = Val(miRsAux!tipomovi)
                Cad = Cad & ",'"
                If Tipo = 2 Then
                    Cad = Cad & "PAGO"
                Else
                    Cad = Cad & "TRASPASO"
                End If
                Cad = Cad & "'," & Tipo & ",'"
                Cad = Cad & "','" & DevNombreSQL(miRsAux!Ampliaci) & "',NULL,"
                ''" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
            Else
                'FACTURA CLIENTE
                Tipo = 0
                Cad = Cad & ",'FRACLI',0,'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
                'Numero factura
                If Not IsNull(miRsAux!NUmSerie) Then Cad = Cad & miRsAux!NUmSerie
                If Not IsNull(miRsAux!numfaccl) Then Cad = Cad & Format(miRsAux!numfaccl, "0000000000")
                If Not IsNull(miRsAux!numvenci) Then Cad = Cad & " - Vto: " & miRsAux!numvenci
                Cad = Cad & "',"
            End If
        End If
        'El importe
        Cad = Cad & TransformaComasPuntos(CStr(DBLet(miRsAux!ImporteD, "N")))
        Cad = Cad & "," & TransformaComasPuntos(CStr(DBLet(miRsAux!ImporteH, "N")))
        
        
        'Texto 6: numero de linea
        Cad = Cad & "," & Format(miRsAux!NumLinea, "00000")
        
        Cad = vSQL & Cad & ")"
        Conn.Execute Cad
        
        miRsAux.MoveNext
        L = L + 1
    Wend
    miRsAux.Close
    '
    'INSERT INTO ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4, texto5, texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion) VALUES (
    ImpirmirListadoCaja = True
End Function






Public Function ListadoCtaBanco() As Boolean
Dim SQL As String
Dim C As String
        
    On Error GoTo eztesoreriacomun
    ListadoCtaBanco = False
    C = "Delete from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    Ejecuta C
    NumRegElim = 1
    
    C = "Select ctabancaria.*,cuentas.nommacta from ctabancaria,cuentas where ctabancaria.codmacta=cuentas.codmacta"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        C = "'" & DBLet(miRsAux!codccost, "T") & "',"
        C = C & NumRegElim & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
        
        SQL = DBLet(miRsAux!ctaingreso, "T")
        If SQL = "" Then SQL = "   ---   "
        C = C & SQL & " / "
        SQL = DBLet(miRsAux!ctagastos, "T")
        If SQL = "" Then SQL = "   ---   "
        C = C & SQL & "','"
        
        'Id cedante - sufijo OEM
        C = C & Right("---" & DBLet(miRsAux!sufijoem, "T"), 3) & "  /  "
        
        SQL = DBLet(miRsAux!idcedente, "T")
        If SQL = "" Then
            C = C & "    ----- "
        Else
            C = C & miRsAux!idcedente
        End If
        C = C & "','"
        
        If DBLet(miRsAux!Entidad, "N") = 0 Then
            'VACIO
            C = C & " Sin asignar"
        Else
            C = C & Format(miRsAux!Entidad, "0000") & "  " & Format(DBLet(miRsAux!oficina, "N"), "0000") & " "
            If IsNull(miRsAux!Control) Then
                C = C & "**"
            Else
                C = C & Right(miRsAux!Control & "00", 2)
            End If
            
            C = C & " " & Format(miRsAux!CtaBanco, "0000000000")
        End If
        
        
        SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, observa1, codigo, texto1, texto2, "
        SQL = SQL & " texto3, texto4,texto5 ) VALUES (" & vUsu.Codigo & ","
        C = SQL & C & "')"
        Conn.Execute C
        
        NumRegElim = NumRegElim + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If NumRegElim = 1 Then
        MsgBox "Ningun dato para mostrar", vbExclamation
        Set miRsAux = Nothing
        Exit Function
    End If
    
    
    
    'Si k hay.....
    'Updateare
    '
    SQL = "UPDATE Usuarios.ztesoreriacomun Set "
    
    
    
    'BORRAR EN CUANTO SE PUEDA
'    'Cuenta banco gastos
'
'    C = "Select texto3 from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo & " GROUP BY texto3"
'    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not miRsAux.EOF
'        If miRsAux!texto3 <> "" Then
'            C = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", miRsAux!texto3, "T")
'            C = "texto4 = '" & DevNombreSQL(C) & "'"
'            C = SQL & C & " WHERE codusu =" & vUsu.Codigo
'        Else
'            C = ""
'        End If
'        miRsAux.MoveNext
'        If C <> "" Then Ejecuta C
'    Wend
'    miRsAux.Close
'
'    C = "Select texto5 from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo & " GROUP BY texto3"
'    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not miRsAux.EOF
'        If miRsAux!texto5 <> "" Then
'            C = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", miRsAux!texto5, "T")
'            C = "texto4 = '" & DevNombreSQL(C) & "'"
'            C = SQL & C & " WHERE codusu =" & vUsu.Codigo
'        Else
'            C = ""
'        End If
'        miRsAux.MoveNext
'        If C <> "" Then Ejecuta C
'    Wend
'    miRsAux.Close
    
    'Centro coste
    C = "Select observa1 from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo & " GROUP BY observa1"
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If miRsAux!observa1 <> "" Then
            C = DevuelveDesdeBD("nomccost", "cabccost", "codccost", miRsAux!observa1, "T")
            C = miRsAux!observa1 & " " & C
            C = "observa1 = '" & DevNombreSQL(C) & "'"
            C = SQL & C & " WHERE codusu =" & vUsu.Codigo
        Else
            C = ""
        End If
        miRsAux.MoveNext
        If C <> "" Then Ejecuta C
    Wend
    miRsAux.Close
    
    
    
    ListadoCtaBanco = True
    
    
    Exit Function
eztesoreriacomun:
    MuestraError Err.Number, , C
    Set miRsAux = Nothing
End Function

Public Function ListadoFormaPago(ByRef SQL As String) As Boolean

    On Error GoTo EListadoFormaPago
    ListadoFormaPago = False
    
    Conn.Execute "DELETE from Usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo
        
    'MONTO EL SQL AL REVES. Empezando por el where

    SQL = " WHERE sforpa.tipforpa = stipoformapago.tipoformapago " & SQL
    SQL = " FROM sforpa ,stipoformapago" & SQL
    SQL = " sforpa.codforpa,sforpa.nomforpa,stipoformapago.descformapago " & SQL
    SQL = "INSERT INTO Usuarios.ztesoreriacomun(codusu,codigo,texto1,texto2) Select " & vUsu.Codigo & "," & SQL
    'INSERT INTO Usuarios.ztesoreriacomun (codusu, observa1, codigo,
    'texto1, texto2,  texto3, texto4 ,texto5) VALUES (
    
    
    
    Conn.Execute SQL
    
    Set miRsAux = New ADODB.Recordset
    SQL = "select count(*) from Usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then SQL = ""
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If SQL <> "" Then
        MsgBox "Ningun dato se ha generado", vbExclamation
    Else
        ListadoFormaPago = True
    End If
    Exit Function
EListadoFormaPago:
    MuestraError Err.Number, "ListadoFormaPago "
End Function

Public Function ListadoEfectosDevueltos(ByRef vSQL As String) As Boolean
Dim SQL As String

    ListadoEfectosDevueltos = False
    Conn.Execute "DELETE from Usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo
    
    
    SQL = "SELECT sefecdev.*,scobro.codmacta as cta,scobro.impvenci , cuentas.nommacta"
    SQL = SQL & " FROM (sefecdev LEFT JOIN scobro ON (sefecdev.numorden = scobro.numorden) AND "
    SQL = SQL & "(sefecdev.fecfaccl = scobro.fecfaccl) AND (sefecdev.codfaccl = scobro.codfaccl) AND "
    SQL = SQL & "(sefecdev.numserie = scobro.numserie)) LEFT JOIN cuentas ON scobro.codmacta = "
    SQL = SQL & "cuentas.codmacta"
    
    
    
    If vSQL <> "" Then SQL = SQL & " WHERE sefecdev.numorden>=0 " & vSQL
    SQL = SQL & " ORDER BY fechadev"
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 1
    vSQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4,texto5,"
    vSQL = vSQL & "fecha1,fecha2,importe1 ,importe2 ) VALUES (" & vUsu.Codigo & ","
    
    While Not miRsAux.EOF
        SQL = NumRegElim & ",'"
        'Si se hubiera producido errores y la cuenta estuviera mal
        If IsNull(miRsAux!Cta) Then
            SQL = "ERROR','RECIBO INCORRECTO"
        Else
            SQL = miRsAux!Cta & "','"
            If IsNull(miRsAux!Nommacta) Then
                SQL = SQL & "CTA NO EXISTE"
            Else
                SQL = SQL & DevNombreSQL(miRsAux!Nommacta)
            End If
        End If
        SQL = NumRegElim & ",'" & SQL & "','"
        SQL = SQL & miRsAux!NUmSerie & "','" & Format(miRsAux!codfaccl, "0000000000") & "','" & miRsAux!numorden
        SQL = SQL & "','" & Format(miRsAux!fecfaccl, FormatoFecha) & "','" & Format(miRsAux!fechadev, FormatoFecha) & "',"
        SQL = SQL & TransformaComasPuntos(CStr(DBLet(miRsAux!impvenci, "N"))) & ","
        SQL = SQL & TransformaComasPuntos(CStr(miRsAux!gastodev)) & ")"
        Conn.Execute vSQL & SQL
    
        NumRegElim = NumRegElim + 1
        miRsAux.MoveNext
    Wend
    
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    ListadoEfectosDevueltos = True
End Function


Public Sub CargaIconoListview(ByRef QueListview As ListView)
On Error Resume Next
    If Dir(App.Path & "\listview.dat", vbArchive) <> "" Then
        QueListview.Picture = LoadPicture(App.Path & "\listview.dat")
        QueListview.PictureAlignment = lvwTopLeft
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


'QUITAR###
'Este procedimiento, probablemente, no servira
'Public Sub EliminarRecibosRemesa2()
'Dim C As String
'Dim J As Integer
'Dim F As Date
'Dim Dias As Integer
'
'
'
'    C = DevuelveDesdeBD("diaselimrem", "paramtesor", "codigo", "1", "N")
'    Dias = Val(C)
'    If Dias = 0 Then
'        MsgBox "Parametro dias eliminacion remesa: 0", vbInformation
'        Exit Sub
'    End If
'
'    Set miRsAux = New ADODB.Recordset
'    C = "select count(*),fecremesa from remesas,scobro where remesas.codigo =scobro.codrem and remesas.anyo=scobro.anyorem"
'    'C = C & " and remesas.situacion ='Q' and fecremesa >='" & Format(DateAdd("d", -J, Now), FormatoFecha) & "'"
'    C = C & " and remesas.situacion ='Q' group by fecremesa"
'    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    J = 0
'    While Not miRsAux.EOF
'        If Not IsNull(miRsAux!fecremesa) Then
'            F = DateAdd("d", Dias, miRsAux!fecremesa)
'            If F < Now Then J = J + miRsAux.Fields(0)
'        End If
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'    If J = 0 Then
'        MsgBox "No existe ningun efecto a eliminar transcurridos " & Dias & " dias", vbInformation
'        Exit Sub
'    Else
'        C = "Existen " & J & " efecto(s) para borrar transcurridos los " & Dias & " dias"
'        C = C & vbCrLf & vbCrLf & "¿Desea eliminar " & J & " efectos?"
'        If MsgBox(C, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
'    End If
'
'
'    'Llegado aqui, borrara los
'    C = "select * from remesas where remesas.situacion ='Q'"
'    miRsAux.Open C, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    While Not miRsAux.EOF
'        F = DateAdd("d", Dias, miRsAux!fecremesa)
'        If F < Now Then
'
'            C = "Delete from scobro where codrem =" & miRsAux!Codigo & " and anyorem =" & miRsAux!Anyo
'            Conn.Execute C
'
'            C = "UPDATE remesas Set situacion=""Z"" where codigo =" & miRsAux!Codigo & " and anyo =" & miRsAux!Anyo
'            Conn.Execute DevNombreSQL(C)
'        End If
'        miRsAux.MoveNext
'        'FALTA##
'        'Faltaria ver si los efectos son imagados
'    Wend
'    miRsAux.Close
'End Sub



Public Function EjecutarSQL(CadenaSQL As String) As Boolean
    On Error Resume Next
    Conn.Execute CadenaSQL
    If Err.Number <> 0 Then
         
         MuestraError Err.Number, "Error ejecutando SQL: " & vbCrLf & CadenaSQL, Err.Description
         EjecutarSQL = False
    Else
         EjecutarSQL = True
    End If
    
End Function

Public Function Memo_Leer(ByRef C As ADODB.Field) As String
    On Error Resume Next
    Memo_Leer = C.Value
    If Err.Number <> 0 Then
        Err.Clear
        Memo_Leer = ""
    End If
End Function


Public Sub DeseleccionaGrid(ByRef DataGrid1 As DataGrid)
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub



Public Function DevuelveNombreUsuario(Codigo As Integer) As String
Dim RS As ADODB.Recordset
    On Error GoTo ED
    Set RS = New ADODB.Recordset
    DevuelveNombreUsuario = ""
    RS.Open "Select nomusu from usuarios.usuarios where codusu = " & Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then DevuelveNombreUsuario = DBLet(RS.Fields(0), "T")
    RS.Close
    Set RS = Nothing
    
    Exit Function
ED:
    MuestraError Err.Number
End Function



Public Function DevuelveNombreInformeSCRYST(NumInforme As Integer, Titulo As String) As String
Dim Cad As String

        DevuelveNombreInformeSCRYST = ""
        Cad = DevuelveDesdeBD("informe", "scryst", "codigo", CStr(NumInforme))

        If Cad = "" Then
            MsgBox "No existe el informe para: " & Titulo & " (" & NumInforme & ")", vbExclamation
            Exit Function
        End If
        
        
        If Dir(App.Path & "\InformesT\" & Cad, vbArchive) = "" Then
            MsgBox "No se encuentra el archivo: " & Cad & vbCrLf & "Opcion: " & Titulo, vbExclamation
            Exit Function
        End If
        DevuelveNombreInformeSCRYST = Cad
            
End Function





'********************************************************************************
'********************************************************************************
'   Carga iconos de un formulario
'   -----------------------------
'       Opciones:   Colection  El col de imagenes
'                   Tipo    1.- Lupa
'                           2.- Fecha
'                           3.- Ayuda
Public Sub CargaImagenesAyudas(ByRef Colec, Tipo As Byte, Optional ToolTipText_ As String)
Dim I As Image

    

    For Each I In Colec
            I.Picture = frmPpal.imgIcoForms.ListImages(Tipo).Picture
            If I.ToolTipText = "" Then
                If ToolTipText_ <> "" Then
                    I.ToolTipText = ToolTipText_
                Else
                    If Tipo = 3 Then
                        I.ToolTipText = "Ayuda"
                    ElseIf Tipo = 2 Then
                        I.ToolTipText = "Seleccionar fecha"
                    Else
                        I.ToolTipText = "Buscar"
                    End If
                End If
            End If
    Next
End Sub


Public Sub Carga1ImagenAyuda(ByRef I As Image, Tipo As Byte)
        I.Picture = frmPpal.imgIcoForms.ListImages(Tipo).Picture
        If I.ToolTipText = "" Then
            If Tipo = 3 Then
                I.ToolTipText = "Ayuda"
            ElseIf Tipo = 2 Then
                I.ToolTipText = "Seleccionar fecha"
            Else
                I.ToolTipText = "Buscar"
            End If
        End If
End Sub

'********************************************************************************
'********************************************************************************



Public Function RemesaSeleccionTipoRemesa(chkEfec As Boolean, chkPaga As Boolean, chkTalon As Boolean) As String
Dim C As String
    C = ""
    
    If chkEfec And chkPaga And chkTalon Then
        'LOS QUIERE TODOS, NO hacemos nada
        
    Else
    
        If Not chkEfec And Not chkPaga And Not chkTalon Then
            'NO QUIERE NINGUNO. Tampoco hago nada
            
        Else
            
            If chkEfec Then
                If chkPaga Then
                    C = " <> 3 "
                Else
                    If chkTalon Then
                        C = " <> 2 "
                    Else
                        C = " = 1" 'Solo efectos
                    End If
                End If
            Else
                If chkPaga Then
                    If chkTalon Then
                        C = " <> 1"
                    Else
                        C = " = 2 "
                    End If
                Else
                    C = " =3 "
                End If
            End If
        End If
    End If
    If C <> "" Then C = " tiporem  " & C
    RemesaSeleccionTipoRemesa = C
End Function




'*******************************************************************
'*******************************************************************
'*******************************************************************
'
'  Letra serie 3 Digitos
'  Con lo cual para algunas campos (numdocum de linapu) son un maximo de
'   10 posiciones. Como antes era un digito letra ser, formateabamos con 9
'       numerofactura debe ser NUMERICO
Public Function SerieNumeroFactura(Posiciones As Integer, Serie As String, Numerofactura As String)
Dim I As Integer
Dim Cad As String
    
    I = Posiciones - Len(Numerofactura) - Len(Serie)
    If I <= 0 Then
        'Hay menos posiciones de las que podemos meter
        Cad = Right(Numerofactura, Posiciones - Len(Numerofactura))
    Else
        Cad = String(I, "0") & Numerofactura
    End If
    SerieNumeroFactura = Serie & Cad
    
    
End Function





'-------------------------------------------------------------------------
'CCargar LISTVIEW con las mempresas de tesoreria
Private Function DevuelveProhibidas() As String
Dim I As Integer


    On Error GoTo EDevuelveProhibidas
    DevuelveProhibidas = ""

    I = vUsu.Codigo Mod 100
    miRsAux.Open "Select * from usuarios.usuarioempresaT WHERE codusu =" & I, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    DevuelveProhibidas = ""
    While Not miRsAux.EOF
        DevuelveProhibidas = DevuelveProhibidas & miRsAux.Fields(1) & "|"
        miRsAux.MoveNext
    Wend
    If DevuelveProhibidas <> "" Then DevuelveProhibidas = "|" & DevuelveProhibidas
    miRsAux.Close
    Exit Function
EDevuelveProhibidas:
    MuestraError Err.Number, "Cargando empresas prohibidas"
    Err.Clear
End Function



Public Sub cargaEmpresasTesor(ByRef Lis As ListView)
Dim Prohibidas As String
Dim IT
Dim Aux As String

    Set miRsAux = New ADODB.Recordset

    Prohibidas = DevuelveProhibidas
    
    Lis.ListItems.Clear
    Aux = "Select * from Usuarios.empresas where tesor=1"
    
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
    
        Aux = "|" & miRsAux!codempre & "|"
        If InStr(1, Prohibidas, Aux) = 0 Then
            Set IT = Lis.ListItems.Add
            IT.Key = "C" & miRsAux!codempre
            If vEmpresa.codempre = miRsAux!codempre Then IT.Checked = True
            IT.Text = miRsAux!nomempre
            IT.Tag = miRsAux!codempre
        End If
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

End Sub

