Attribute VB_Name = "BaseDato"
Option Explicit

Private SQL As String
'
Dim ImpD As Currency
Dim ImpH As Currency
Dim RT As ADODB.Recordset
'
'
'Dim d As String
'Dim H As String
''Para los balances
'Dim M1 As Integer   ' años y kmeses para el balance
'Dim M2 As Integer
'Dim M3 As Integer
'Dim A1 As Integer
'Dim A2 As Integer
'Dim A3 As Integer
'Dim vCta As String
'Dim ImAcD As Currency  'importes
'Dim ImAcH As Currency
'Dim ImPerD As Currency  'importes
'Dim ImPerH As Currency
'Dim ImCierrD As Currency  'importes
'Dim ImCierrH As Currency
'Dim Contabilidad As Integer
'Dim Aux As String
'Dim vFecha1 As Date
'Dim vFecha2 As Date
'Dim VFecha3 As Date
'Dim Codigo As String
'Dim EjerciciosCerrados As Boolean
'Dim NumAsiento As Integer
'Dim Nulo1 As Boolean
'Dim Nulo2 As Boolean
'
'Dim VarConsolidado(2) As String
'
'Dim EsBalancePerdidas_y_ganancias As Boolean

'--------------------------------------------------------------------
'--------------------------------------------------------------------
'Private Function ImporteASQL(ByRef Importe As Currency) As String
'ImporteASQL = ","
'If Importe = 0 Then
'    ImporteASQL = ImporteASQL & "NULL"
'Else
'    ImporteASQL = ImporteASQL & TransformaComasPuntos(CStr(Importe))
'End If
'End Function
'
'
'
''--------------------------------------------------------------------
''--------------------------------------------------------------------
'' El dos sera para k pinte el 0. Ya en el informe lo trataremos.
'' Con esta opcion se simplifica bastante la opcion de totales
'Private Function ImporteASQL2(ByRef Importe As Currency) As String
'    ImporteASQL2 = "," & TransformaComasPuntos(CStr(Importe))
'End Function



'--------------------------------------------------------------------
'--------------------------------------------------------------------




'Public Function FacturaCorrecta(NumF As Long, AnoF As Integer, ByRef Serie As String) As String
'
'
'On Error GoTo EFacturaCorrecta
'FacturaCorrecta = ""
'
'Set RT = New ADODB.Recordset
'
''Calculamos el total de los importes
'SQL = "Select * from cabfact where numserie = '" & Serie & "' AND codfaccl = " & NumF
'SQL = SQL & " AND anofaccl = " & AnoF
'RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'If RT.EOF Then
'    FacturaCorrecta = "No existe la factura(Serire/Numero/año): " & Serie & " / " & NumF & " / " & AnoF
'Else
'    'Si que existe la factura
'    'Sumamos bases imponibles
'    ImpD = RT!ba1faccl
'    If Not IsNull(RT!ba2faccl) Then ImpD = ImpD + RT!ba2faccl
'    If Not IsNull(RT!ba3faccl) Then ImpD = ImpD + RT!ba3faccl
'
'    'IVAS
'    ImpD = ImpD + RT!ti1faccl
'    If Not IsNull(RT!ti2faccl) Then ImpD = ImpD + RT!ti2faccl
'    If Not IsNull(RT!ti3faccl) Then ImpD = ImpD + RT!ti3faccl
'
'    'Retenciones
'    If Not IsNull(RT!tr1faccl) Then ImpD = ImpD + RT!tr1faccl
'    If Not IsNull(RT!tr2faccl) Then ImpD = ImpD + RT!tr2faccl
'    If Not IsNull(RT!tr3faccl) Then ImpD = ImpD + RT!tr3faccl
'
'    'Importe retencion  (SE LER RESTA)
'    If Not IsNull(RT!trefaccl) Then ImpD = ImpD - RT!trefaccl
'
'    'Comprobamos que el importe que pone es el que corresponde
'    If ImpD <> RT!totfaccl Then
'        FacturaCorrecta = "La suma de bases, ivas, retenciones no coincide con el total factura: " & ImpD & " /  " & RT!totfaccl
'    Else
'        'Si coincide la suma de las facturas
'        FacturaCorrecta = ""
'End If
'RT.Close
'If FacturaCorrecta = "" Then
'    'Ahora comprobamos que la suma de lineas coincide con el totalfac->impd
'    SQL = "SELECT sum(impbascl) FROM linfact  WHERE linfact.numserie= '" & Serie & "'"
'    SQL = SQL & " AND linfact.codfaccl= " & NumF
'    SQL = SQL & " AND linfact.anofaccl=" & AnoF & ";"
'    RT.Open SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
'    ImpH = 0
'    If Not RT.EOF Then
'        If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
'    End If
'    RT.Close
'    If ImpD <> ImpH Then _
'        FacturaCorrecta = "El importe indicado en la cabecera no coincide con el de la suma de lineas: " & ImpD & " / " & ImpH
'
'End If
'
'EFacturaCorrecta:
'    If Err.Number <> 0 Then _
'        FacturaCorrecta = Err.Number & " - " & Err.Description
'        Err.Clear
'    End If
'    Set RT = Nothing
'End Function
'
'
'
'
'
'Public Function SeparaCampoBusqueda(Tipo As String, Campo As String, CADENA As String, ByRef DevSQL As String) As Byte
'Dim Cad As String
'Dim Aux As String
'Dim Ch As String
'Dim Fin As Boolean
'Dim I, J As String
'
'On Error GoTo ErrSepara
'SeparaCampoBusqueda = 1
'DevSQL = ""
'Cad = ""
'Select Case Tipo
'Case "N"
'    '----------------  NUMERICO  ---------------------
'    I = CararacteresCorrectos(CADENA, "N")
'    If I > 0 Then Exit Function  'Ha habido un error y salimos
'    'Comprobamos si hay intervalo ':'
'    I = InStr(1, CADENA, ":")
'    If I > 0 Then
'        'Intervalo numerico
'        Cad = Mid(CADENA, 1, I - 1)
'        Aux = Mid(CADENA, I + 1)
'        If Not IsNumeric(Cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
'        'Intervalo correcto
'        'Construimos la cadena
'        DevSQL = Campo & " >= " & Cad & " AND " & Campo & " <= " & Aux
'        '----
'        'ELSE
'        Else
'            'Prueba
'            'Comprobamos que no es el mayor
'            If CADENA = ">>" Or CADENA = "<<" Then
'                DevSQL = "1=1"
'             Else
'                    Fin = False
'                    I = 1
'                    Cad = ""
'                    Aux = "NO ES NUMERO"
'                    While Not Fin
'                        Ch = Mid(CADENA, I, 1)
'                        If Ch = ">" Or Ch = "<" Or Ch = "=" Then
'                            Cad = Cad & Ch
'                            Else
'                                Aux = Mid(CADENA, I)
'                                Fin = True
'                        End If
'                        I = I + 1
'                        If I > Len(CADENA) Then Fin = True
'                    Wend
'                    'En aux debemos tener el numero
'                    If Not IsNumeric(Aux) Then Exit Function
'                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
'                    If Cad = "" Then Cad = " = "
'                    DevSQL = Campo & " " & Cad & " " & Aux
'            End If
'        End If
'Case "F"
'     '---------------- FECHAS ------------------
'    I = CararacteresCorrectos(CADENA, "F")
'    If I = 1 Then Exit Function
'    'Comprobamos si hay intervalo ':'
'    I = InStr(1, CADENA, ":")
'    If I > 0 Then
'        'Intervalo de fechas
'        Cad = Mid(CADENA, 1, I - 1)
'        Aux = Mid(CADENA, I + 1)
'        If Not EsFechaOKString(Cad) Or Not EsFechaOKString(Aux) Then Exit Function  'Fechas incorrectas
'        'Intervalo correcto
'        'Construimos la cadena
'        Cad = Format(Cad, FormatoFecha)
'        Aux = Format(Aux, FormatoFecha)
'        'En my sql es la ' no el #
'        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
'        DevSQL = Campo & " >='" & Cad & "' AND " & Campo & " <= '" & Aux & "'"
'        '----
'        'ELSE
'        Else
'            'Comprobamos que no es el mayor
'            If CADENA = ">>" Or CADENA = "<<" Then
'                  DevSQL = "1=1"
'            Else
'                Fin = False
'                I = 1
'                Cad = ""
'                Aux = "NO ES FECHA"
'                While Not Fin
'                    Ch = Mid(CADENA, I, 1)
'                    If Ch = ">" Or Ch = "<" Or Ch = "=" Then
'                        Cad = Cad & Ch
'                        Else
'                            Aux = Mid(CADENA, I)
'                            Fin = True
'                    End If
'                    I = I + 1
'                    If I > Len(CADENA) Then Fin = True
'                Wend
'                'En aux debemos tener el numero
'                If Not EsFechaOKString(Aux) Then Exit Function
'                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
'                Aux = "'" & Format(Aux, FormatoFecha) & "'"
'                If Cad = "" Then Cad = " = "
'                DevSQL = Campo & " " & Cad & " " & Aux
'            End If
'        End If
'
'
'
'
'Case "T"
'    '---------------- TEXTO ------------------
'    I = CararacteresCorrectos(CADENA, "T")
'    If I = 1 Then Exit Function
'
'    'Comprobamos que no es el mayor
'     If CADENA = ">>" Or CADENA = "<<" Then
'        DevSQL = "1=1"
'        Exit Function
'    End If
'    'Cambiamos el * por % puesto que en ADO es el caraacter para like
'    I = 1
'    Aux = CADENA
'    While I <> 0
'        I = InStr(1, Aux, "*")
'        If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "%" & Mid(Aux, I + 1)
'    Wend
'    'Cambiamos el ? por la _ pue es su omonimo
'    I = 1
'    While I <> 0
'        I = InStr(1, Aux, "?")
'        If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "_" & Mid(Aux, I + 1)
'    Wend
'    Cad = Mid(CADENA, 1, 2)
'    If Cad = "<>" Then
'        Aux = Mid(CADENA, 3)
'        DevSQL = Campo & " LIKE '!" & Aux & "'"
'        Else
'        DevSQL = Campo & " LIKE '" & Aux & "'"
'    End If
'
'
'
'
'Case "B"
'    'Como vienen de check box o del option box
'    'los escribimos nosotros luego siempre sera correcta la
'    'sintaxis
'    'Los booleanos. Valores buenos son
'    'Verdadero , Falso, True, False, = , <>
'    'Igual o distinto
'    I = InStr(1, CADENA, "<>")
'    If I = 0 Then
'        'IGUAL A valor
'        Cad = " = "
'        Else
'            'Distinto a valor
'        Cad = " <> "
'    End If
'    'Verdadero o falso
'    I = InStr(1, CADENA, "V")
'    If I > 0 Then
'            Aux = "True"
'            Else
'            Aux = "False"
'    End If
'    'Ponemos la cadena
'    DevSQL = Campo & " " & Cad & " " & Aux
'
'Case Else
'    'No hacemos nada
'        Exit Function
'End Select
'SeparaCampoBusqueda = 0
'ErrSepara:
'    If Err.Number <> 0 Then MuestraError Err.Number
'End Function
'
'
'Private Function CararacteresCorrectos(vCad As String, Tipo As String) As Byte
'Dim I As Integer
'Dim Ch As String
'Dim Error As Boolean
'
'CararacteresCorrectos = 1
'Error = False
'Select Case Tipo
'Case "N"
'    'Numero. Aceptamos numeros, >,< = :
'    For I = 1 To Len(vCad)
'        Ch = Mid(vCad, I, 1)
'        Select Case Ch
'            Case "0" To "9"
'            Case "<", ">", ":", "=", ".", " ", "-"
'            Case Else
'                Error = True
'                Exit For
'        End Select
'    Next I
'Case "T"
'    'Texto aceptamos numeros, letras y el interrogante y el asterisco
'    For I = 1 To Len(vCad)
'        Ch = Mid(vCad, I, 1)
'        Select Case Ch
'            Case "a" To "z"
'            Case "A" To "Z"
'            Case "0" To "9"
'            Case "*", "%", "?", "_", "\", "/", ":", ".", " ", "-" ' estos son para un caracter sol no esta demostrado , "%", "&"
'            'Esta es opcional
'            Case "<", ">"
'            Case "Ñ", "ñ"
'            Case Else
'                Error = True
'                Exit For
'        End Select
'    Next I
'Case "F"
'    'Numeros , "/" ,":"
'    For I = 1 To Len(vCad)
'        Ch = Mid(vCad, I, 1)
'        Select Case Ch
'            Case "0" To "9"
'            Case "<", ">", ":", "/", "="
'            Case Else
'                Error = True
'                Exit For
'        End Select
'    Next I
'Case "B"
'    'Numeros , "/" ,":"
'    For I = 1 To Len(vCad)
'        Ch = Mid(vCad, I, 1)
'        Select Case Ch
'            Case "0" To "9"
'            Case "<", ">", ":", "/", "=", " "
'            Case Else
'                Error = True
'                Exit For
'        End Select
'    Next I
'End Select
''Si no ha habido error cambiamos el retorno
'If Not Error Then CararacteresCorrectos = 0
'End Function
'
'
'
'
'
'
''Este modulo estaba antes del ADOBUS
'Public Function BloquearAsiento(NA As String, ND As String, NF As String) As Boolean
'Dim RB As Recordset
'
'    'Pensar en la coicidencia en el tiempo de dos transacciones es improbable.Teimpo de acceso en milisegudos
'    BloquearAsiento = False
'    SQL = "SELECT * from cabapu "
'    SQL = SQL & " WHERE numdiari =" & ND
'    SQL = SQL & " AND fechaent='" & NF
'    SQL = SQL & "' AND numasien=" & NA
'    Set RB = New Recordset
'    RB.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'    If RB.EOF Then
'        MsgBox "Asiento inexistente o ha sido borrado", vbExclamation
'        Exit Function
'    End If
'    If RB!bloqactu = 0 Then
'        'Asiento no bloqueado
'        'Tratar de modificarlo
'        SQL = "UPDATE cabapu set bloqactu=1 "
'        SQL = SQL & " WHERE numdiari =" & ND
'        SQL = SQL & " AND fechaent='" & NF
'        SQL = SQL & " ' AND numasien=" & NA
'        RB.Close
'        On Error Resume Next
'        Conn.Execute SQL
'        If Err.Number <> 0 Then
'            Err.Clear
'        Else
'            BloquearAsiento = True
'        End If
'        On Error GoTo 0  'quitamos los errores
'    Else
'        RB.Close
'    End If
'    Set RB = Nothing
'End Function
'
'
'
'
'
'
'
'
'
'
'
'
'Public Function DesbloquearAsiento(NA As String, ND As String, NF As String) As Boolean
'On Error Resume Next
'    SQL = "UPDATE cabapu "
'    SQL = SQL & " SET bloqactu=0 "
'    SQL = SQL & " WHERE numdiari =" & ND
'    SQL = SQL & " AND fechaent='" & NF
'    SQL = SQL & "' AND numasien=" & NA
'    Conn.Execute SQL
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, "Desbloqueo asiento: " & NA & " /  " & ND & " /     " & NF
'        DesbloquearAsiento = False
'    Else
'        DesbloquearAsiento = True
'    End If
'End Function
'
'

'-------------------------------------------------------------------

Public Function CargaDatosConExt(ByRef Cuenta As String, fec1 As Date, fec2 As Date, ByRef vSQL As String, ByRef DescCuenta As String) As Byte
Dim ACUM As Double  'Acumulado anterior

On Error GoTo ECargaDatosConExt
CargaDatosConExt = 1


'DELETES
SQL = "DELETE FROM tmpconextcab where codusu = " & vUsu.Codigo
Conn.Execute SQL
SQL = "DELETE FROM tmpconext where codusu = " & vUsu.Codigo
Conn.Execute SQL


'Insertamos en los campos de cabecera de cuentas
NombreSQL DescCuenta
SQL = Cuenta & "    -    " & DescCuenta
SQL = "INSERT INTO tmpconextcab (codusu,cta,fechini,fechfin,cuenta) VALUES (" & vUsu.Codigo & ", '" & Cuenta & "','" & Format(fec1, "dd/mm/yyyy") & "','" & Format(fec2, "dd/mm/yyyy") & "','" & SQL & "')"
Conn.Execute SQL


'los totatales
'Dim T1, cad
'cad = "Cuenta: " & DescCuenta & vbCrLf
'T1 = Timer


If Not CargaAcumuladosTotales(Cuenta) Then Exit Function
'cad = cad & "Acum Total:" & Format(Timer - T1, "0.000") & vbCrLf

'Los caumulados anteriores
If Not CargaAcumuladosAnteriores(Cuenta, fec1, ACUM) Then Exit Function
'cad = cad & "Anterior:   " & Format(Timer - T1, "0.000") & vbCrLf

'GENERAMOS LA TBLA TEMPORAL
If Not CargaTablaTemporalConExt(Cuenta, vSQL, ACUM) Then Exit Function


'cad = cad & "Tabla:    " & Format(Timer - T1, "0.000") & vbCrLf
'MsgBox cad


CargaDatosConExt = 0
Exit Function
ECargaDatosConExt:
    CargaDatosConExt = 2
    MuestraError Err.Number, "Gargando datos temporales. Cta: " & Cuenta, Err.Description
End Function



Private Function CargaAcumuladosTotales(ByRef Cta As String) As Boolean
    CargaAcumuladosTotales = False
    SQL = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH"
    SQL = SQL & " from hlinapu where codmacta='" & Cta & "'"
    SQL = SQL & " AND fechaent >=  '" & Format(vParam.fechaini, FormatoFecha) & "'"
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImpD = 0
        Else
        ImpD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImpH = 0
    Else
        ImpH = RT.Fields(1)
    End If
    RT.Close
    Set RT = Nothing
    SQL = "UPDATE tmpconextcab SET acumtotD= " & TransformaComasPuntos(CStr(ImpD)) 'Format(ImpD, "#,###,##0.00")
    SQL = SQL & ", acumtotH= " & TransformaComasPuntos(CStr(ImpH)) 'Format(ImpH, "#,###,##0.00")
    ImpD = ImpD - ImpH
    SQL = SQL & ", acumtotT= " & TransformaComasPuntos(CStr(ImpD)) 'Format(ImpD, "#,###,##0.00")
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute SQL
    CargaAcumuladosTotales = True
End Function


Private Function CargaAcumuladosAnteriores(ByRef Cta As String, ByRef FI As Date, ByRef ACUM As Double) As Boolean
Dim F1 As Date

    CargaAcumuladosAnteriores = False
    SQL = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH"
    SQL = SQL & " from hlinapu where codmacta='" & Cta & "'"
    F1 = vParam.fechaini

    Do
        If FI < F1 Then F1 = DateAdd("yyyy", -1, F1)
    Loop Until F1 <= FI
    'SQL = SQL & " AND fechaent >=  '" & Format(vParam.fechaini, FormatoFecha) & "'"
    SQL = SQL & " AND fechaent >=  '" & Format(F1, FormatoFecha) & "'"
    SQL = SQL & " AND fechaent <  '" & Format(FI, FormatoFecha) & "'"
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImpD = 0
    Else
        ImpD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImpH = 0
    Else
        ImpH = RT.Fields(1)
    End If
    RT.Close
    ACUM = ImpD - ImpH
    SQL = "UPDATE tmpconextcab SET acumantD= " & TransformaComasPuntos(CStr(ImpD))
    SQL = SQL & ", acumantH= " & TransformaComasPuntos(CStr(ImpH))
    SQL = SQL & ", acumantT= " & TransformaComasPuntos(CStr(ACUM))
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute SQL
    Set RT = Nothing
    CargaAcumuladosAnteriores = True
End Function



Private Function CargaTablaTemporalConExt(Cta As String, vSele As String, ByRef ACUM As Double) As Boolean
Dim Aux As Currency
Dim ImporteD As String
Dim ImporteH As String
Dim Contador As Integer
Dim RC As String
On Error GoTo Etmpconext


'TIEMPOS
'Dim T1, Cadenita
'T1 = Timer
'Cadenita = "Cuenta: " & Cta & vbCrLf

CargaTablaTemporalConExt = False

'Conn.Execute "Delete from tmpconext where codusu =" & vUsu.Codigo
Set RT = New ADODB.Recordset
SQL = "Select * from hlinapu where codmacta='" & Cta & "'"
SQL = SQL & " AND " & vSele & " ORDER BY fechaent,numasien"
RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

'Cadenita = Cadenita & "Select: " & Format(Timer - T1, "0.0000") & vbCrLf
'T1 = Timer

SQL = "INSERT INTO tmpconext (codusu, POS,numdiari, fechaent, numasien, linliapu, timporteD, timporteH, saldo, Punteada,nomdocum,ampconce,cta,contra,ccost) VALUES ("
'ImpD = 0 ASI LLEVAMOS EL ACUMULADO
'ImpH = 0
Contador = 0
While Not RT.EOF
    Contador = Contador + 1
    If Not IsNull(RT!timported) Then
        Aux = RT!timported
        ImpD = ImpD + Aux
        ImporteD = TransformaComasPuntos(RT!timported)
        ImporteH = "Null"
    Else
        Aux = RT!timporteH
        ImporteD = "Null"
        ImporteH = TransformaComasPuntos(RT!timporteH)
        ImpH = ImpH + Aux
        Aux = -1 * Aux
    End If
    ACUM = ACUM + Aux
    
    'Insertar
    RC = vUsu.Codigo & "," & Contador & "," & RT!numdiari & ",'" & Format(RT!fechaent, FormatoFecha) & "'," & RT!Numasien & "," & RT!Linliapu & ","
    RC = RC & ImporteD & "," & ImporteH
    If RT!punteada <> 0 Then
        ImporteD = "SI"
        Else
        ImporteD = ""
    End If
    RC = RC & "," & TransformaComasPuntos(CStr(ACUM)) & ",'" & ImporteD & "','"
    RC = RC & RT!Numdocum & "','" & DevNombreSQL(RT!ampconce) & "','" & Cta & "',"
    If IsNull(RT!ctacontr) Then
        RC = RC & "NULL"
    Else
        RC = RC & "'" & RT!ctacontr & "'"
    End If
    RC = RC & ","
    If IsNull(RT!codccost) Then
        RC = RC & "NULL"
    Else
        RC = RC & "'" & RT!codccost & "'"
    End If
    RC = RC & ")"
    
    'IMPORTANTE###
    '------------------------
    'NO EJECUTO
    'Conn.Execute SQL & RC
    'Sig
    RT.MoveNext
Wend
RT.Close


'Cadenita = Cadenita & "Recorrer: " & Format(Timer - T1, "0.0000") & vbCrLf
'T1 = Timer

    SQL = "UPDATE tmpconextcab SET acumperD= " & TransformaComasPuntos(CStr(ImpD))
    SQL = SQL & ", acumperH= " & TransformaComasPuntos(CStr(ImpH))
    ImpD = ImpD - ImpH
    SQL = SQL & ", acumperT= " & TransformaComasPuntos(CStr(ImpD))
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute SQL

    CargaTablaTemporalConExt = True
    
'Cadenita = Cadenita & "Actualizar: " & Format(Timer - T1, "0.0000") & vbCrLf
'MsgBox Cadenita
Exit Function
Etmpconext:
    MuestraError Err.Number, "Generando datos saldos"
    Set RT = Nothing
End Function








Public Function DevuelveLaCtaBanco(ByRef Cta As String) As String
Dim RS As ADODB.Recordset
    
    DevuelveLaCtaBanco = "|||||"
    Set RS = New ADODB.Recordset
    RS.Open "Select entidad,oficina,cc,cuentaba,iban from cuentas where codmacta ='" & Cta & "'", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If DBLet(RS!Entidad, "N") > 0 Then DevuelveLaCtaBanco = Format(DBLet(RS!Entidad, "T"), "0000") & "|" & Format(DBLet(RS!oficina, "T"), "0000") & "|" & DBLet(RS!CC, "T") & "|" & DBLet(RS!cuentaba, "T") & "|" & UCase(DBLet(RS!IBAN, "T")) & "|"
    End If
        
    RS.Close
    Set RS = Nothing
End Function
