Attribute VB_Name = "Norma34"
Option Explicit
        
        
'******************************************************************************************************************
'******************************************************************************************************************
'
'       Normas 34 y 68
'
'******************************************************************************************************************
'******************************************************************************************************************
    
        
Dim AuxD As String
Private NumeroTransferencia As Integer
'----------------------------------------------------------------------
'  Copia fichero generado bajo
Public Sub CopiarFicheroNorma43(Es34 As Boolean, Destino As String)

    
    'If Not CopiarEnDisquette(True, 3) Then
        AuxD = Destino
        CopiarEnDisquette False, 0, Es34 'A disco
    
        
End Sub

Private Function CopiarEnDisquette(A_disquetera As Boolean, Intentos As Byte, Es34 As Boolean) As Boolean
Dim I As Integer
Dim Cad As String

On Error Resume Next

    CopiarEnDisquette = False
    
    If A_disquetera Then
        For I = 1 To Intentos
            Cad = "Introduzca un disco vacio. (" & I & ")"
            MsgBox Cad, vbInformation
            FileCopy App.Path & "\norma34.txt", "a:\norma34.txt"
            If Err.Number <> 0 Then
                MuestraError Err.Number, "Copiar En Disquette"
            Else
                CopiarEnDisquette = True
                Exit For
            End If
        Next I
    Else
        If AuxD = "" Then
            Cad = Format(Now, "ddmmyyhhnn")
            Cad = App.Path & "\" & Cad & ".txt"
        Else
            Cad = AuxD
        End If
        If Es34 Then
            FileCopy App.Path & "\norma34.txt", Cad
        Else
            FileCopy App.Path & "\norma68.txt", Cad
        End If
        If Err.Number <> 0 Then
            MsgBox "Error creando copia fichero. Consulte soporte técnico." & vbCrLf & Err.Description, vbCritical
        Else
            MsgBox "El fichero esta guardado como: " & Cad, vbInformation
        End If
            
    End If
End Function



'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma34(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, Pagos As Boolean) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Cad As String


    On Error GoTo EGen
    GeneraFicheroNorma34 = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
    'Cargamos la cuenta
    Cad = "Select * from ctabancaria where codmacta='" & CuentaPropia & "'"
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 10)
    If RS.EOF Then
        Cad = ""
    Else
        If IsNull(RS!entidad) Then
            Cad = ""
        Else
            Cad = Format(RS!entidad, "0000") & "|" & Format(DBLet(RS!oficina, "T"), "0000") & "|" & DBLet(RS!Control, "T") & "|" & Format(DBLet(RS!CtaBanco, "T"), "0000000000") & "|"
            CuentaPropia = Cad
        End If
        
        
        'Identificador norma bancaria
        If Not IsNull(RS!idnorma34) Then Aux = RS!idnorma34
    End If
    RS.Close
    Set RS = Nothing
    If Cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NFich
    
    
    
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 10)  'CIF EMPRESA
    
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, Cad
    Cabecera2 NFich, CodigoOrdenante, Cad
    Cabecera3 NFich, CodigoOrdenante, Cad
    Cabecera4 NFich, CodigoOrdenante, Cad
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set RS = New ADODB.Recordset
    If Pagos Then
        Aux = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,despobla from spagop,cuentas"
        Aux = Aux & " where codmacta=ctaprove and transfer =" & NumeroTransferencia
    Else
        'ABONOS
         '
        Aux = "Select scobro.codbanco as entidad,scobro.codsucur as oficina,scobro.cuentaba,scobro.digcontr as CC"
        Aux = Aux & ",nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta from scobro,cuentas"
        Aux = Aux & " where cuentas.codmacta=scobro.codmacta and transfer =" & NumeroTransferencia
    End If
    RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If RS.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not RS.EOF
            If Pagos Then
                Im = DBLet(RS!imppagad, "N")
                Im = RS!ImpEfect - Im
                Aux = RellenaAceros(RS!ctaprove, False, 12)
            
            Else
                Im = Abs(RS!impvenci)
                Aux = RellenaAceros(RS!codmacta, False, 12)
            End If
            
            'Cad = "06"
            'Cad = Cad & "56"
            'Cad = Cad & " "
            'Aux = "06" & "56" & " " & CodigoOrdenante & Aux  'Ordenante y socio juntos
        
            Aux = "06" & "56" & CodigoOrdenante & Aux   'Ordenante y socio juntos
        
            Linea1 NFich, Aux, RS, Im, Cad, ConceptoTransferencia
            Linea2 NFich, Aux, RS, Cad
            Linea3 NFich, Aux, RS, Cad
            Linea4 NFich, Aux, RS, Cad
            Linea5 NFich, Aux, RS, Cad
            Linea6 NFich, Aux, RS, Cad, ConceptoTr, Pagos
            If Pagos Then Linea7 NFich, Aux, RS, Cad
        
        
        
        
            Importe = Importe + Im
            Regs = Regs + 1
            RS.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, Cad, Pagos
    End If
    RS.Close
    Set RS = Nothing
    Close (NFich)
    If Regs > 0 Then GeneraFicheroNorma34 = True
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function


Private Function RellenaABlancos(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim Cad As String
    
    Cad = Space(Longitud)
    If PorLaDerecha Then
        Cad = CADENA & Cad
        RellenaABlancos = Left(Cad, Longitud)
    Else
        Cad = Cad & CADENA
        RellenaABlancos = Right(Cad, Longitud)
    End If
    
End Function



Private Function RellenaAceros(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim Cad As String
    
    Cad = Mid("00000000000000000000", 1, Longitud)
    If PorLaDerecha Then
        Cad = CADENA & Cad
        RellenaAceros = Left(Cad, Longitud)
    Else
        Cad = Cad & CADENA
        RellenaAceros = Right(Cad, Longitud)
    End If
    
End Function



'Private Sub Cabecera1(NF As Integer,ByRef CodOrde As String)
'Dim Cad As String
'
'End Sub

Private Sub Cabecera1(NF As Integer, ByRef CodOrde As String, Fecha As Date, Cta As String, ByRef Cad As String)

    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "001"
    Cad = Cad & Format(Now, "ddmmyy")
    Cad = Cad & Format(Fecha, "ddmmyy")
    'Cuenta bancaria
    Cad = Cad & RecuperaValor(Cta, 1)
    Cad = Cad & RecuperaValor(Cta, 2)
    Cad = Cad & RecuperaValor(Cta, 4)
    Cad = Cad & "0"  'Sin relacion
    Cad = Cad & "   " & RecuperaValor(Cta, 3)  'Digito de control bancario
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



Private Sub Cabecera2(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)
    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "002"
    
    Cad = Cad & RellenaABlancos(vEmpresa.nomempre, True, 30)   'Nombre empresa
  
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Cabecera3(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)
    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "003"
    
    
    AuxD = DevuelveDesdeBD("direccion", "empresa2", "codigo", 1, "N")
    Cad = Cad & RellenaABlancos(AuxD, True, 30)   'Nombre empresa
    Cad = Cad & RellenaABlancos("", True, 30)   'Nombre empresa
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



Private Sub Cabecera4(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)

    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "004"
    
    AuxD = DevuelveDesdeBD("codpos", "empresa2", "codigo", 1, "N")
    Cad = Cad & RellenaABlancos(AuxD, False, 5)
    Cad = Cad & " "
    AuxD = DevuelveDesdeBD("provincia", "empresa2", "codigo", 1, "N")
    Cad = Cad & RellenaABlancos(AuxD, True, 30)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



'ConceptoTransferencia
'1.- Abono nomina
'9.- Transferencia ordinaria
Private Sub Linea1(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Importe1 As Currency, ByRef Cad As String, vConceptoTransferencia As String)


   
    '
    Cad = CodOrde   'llevara tb la ID del socio
    Cad = Cad & "010"
    Cad = Cad & RellenaAceros(CStr(Round(Importe1, 2) * 100), False, 12)
    
    Cad = Cad & RellenaAceros(CStr(RS1!entidad), False, 4)     'Entidad
    Cad = Cad & RellenaAceros(CStr(RS1!oficina), False, 4)   'Sucur
    Cad = Cad & RellenaAceros(CStr(RS1!cuentaba), False, 10)  'Cta
    Cad = Cad & "1" & vConceptoTransferencia
    Cad = Cad & "  "
    Cad = Cad & RellenaAceros(CStr(RS1!CC), False, 2)  'CC
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea2(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "011"
    Cad = Cad & RellenaABlancos(RS1!Nommacta, False, 36)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea3(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "012"
    Cad = Cad & RellenaABlancos(DBLet(RS1!dirdatos, "T"), False, 36)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea4(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "013"
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea5(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "014"
    Cad = Cad & RellenaABlancos(DBLet(RS1!codposta, "T"), False, 5) & " "
    Cad = Cad & RellenaABlancos(DBLet(RS1!despobla, "T"), False, 30)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea6(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String, ByRef ConceptoT As String, Pagos As Boolean)
Dim Aux As String

    Aux = ConceptoT
    If Pagos Then
        'Tiene dos campos para las descripcion. Si no tiene nada pondre la descripcion de la transferencia
        Aux = Trim(DBLet(RS1!text1csb, "T"))
        If Aux = "" Then Aux = ConceptoT
    End If

    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "016"
    Cad = Cad & RellenaABlancos(Aux, False, 35)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea7(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)


    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "017"
    Cad = Cad & RellenaABlancos(DBLet(RS1!text2csb, "T"), False, 35)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub




Private Sub Totales(NF As Integer, ByRef CodOrde As String, Total As Currency, Registros As Integer, ByRef Cad As String, Pagos As Boolean)
    Cad = "08" & "56"
    Cad = Cad & CodOrde    'llevara tb la ID del socio
    Cad = Cad & Space(15)
    Cad = Cad & RellenaAceros(CStr(Int(Round(Total * 100, 2))), False, 12)
    Cad = Cad & RellenaAceros(CStr(Registros), False, 8)
    If Pagos Then
        Cad = Cad & RellenaAceros(CStr((Registros * 7) + 4 + 1), False, 10)
    Else
        Cad = Cad & RellenaAceros(CStr((Registros * 6) + 4 + 1), False, 10)
    End If
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub











'******************************************************************************************************************
'******************************************************************************************************************
'
'       Genera fichero NORMA 68
'
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma68(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Cad As String


    On Error GoTo EGen
    GeneraFicheroNorma68 = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
    'Cargamos la cuenta
    Cad = "Select * from ctabancaria where codmacta='" & CuentaPropia & "'"
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 9)
    Aux = Mid(CIF & Space(10), 1, 9)
    If RS.EOF Then
        Cad = ""
    Else
        If IsNull(RS!entidad) Then
            Cad = ""
        Else
            
            CodigoOrdenante = Format(RS!entidad, "0000") & Format(DBLet(RS!oficina, "N"), "0000") & Format(DBLet(RS!Control, "N"), "00") & Format(DBLet(RS!CtaBanco, "T"), "0000000000")
            
            If Not DevuelveIBAN2("ES", CodigoOrdenante, Cad) Then Cad = ""
            CuentaPropia = "ES" & Cad & CodigoOrdenante
            
        End If
        
        
    End If
    RS.Close
    Set RS = Nothing
    If Cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\norma68.txt" For Output As #NFich
    
    
    
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 9)  'CIF EMPRESA
    CodigoOrdenante = CodigoOrdenante & "000" 'el sufijo
    
    'CABECERA
    Cabecera1_68 NFich, CodigoOrdenante, Fecha, CuentaPropia, Cad
   
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set RS = New ADODB.Recordset
    Aux = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,despobla,nifdatos,razosoci,desprovi,pais from spagop,cuentas"
    Aux = Aux & " where codmacta=ctaprove and transfer =" & NumeroTransferencia
    RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If RS.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not RS.EOF
            
                Im = DBLet(RS!imppagad, "N")
                Im = RS!ImpEfect - Im
                Aux = RellenaABlancos(RS!nifdatos, True, 12)
            

            
            
            Aux = "06" & "59" & CodigoOrdenante & Aux   'Ordenante y nifprove
        
            Linea1_68 NFich, Aux, RS, Cad
            Linea2_68 NFich, Aux, RS, Cad
            Linea3_68 NFich, Aux, RS, Cad
            Linea4_68 NFich, Aux, RS, Cad
            Linea5_68 NFich, Aux, RS, Cad, Fecha, Im
            Linea6_68 NFich, Aux, RS, Im, Cad, ConceptoTr
            'If Pagos Then Linea7 NFich, Aux, RS, Cad
        
        
        
        
            Importe = Importe + Im
            Regs = Regs + 1
            RS.MoveNext
        Wend
        'Imprimimos totales
        Totales68 NFich, CodigoOrdenante, Importe, Regs, Cad
    End If
    RS.Close
    Set RS = Nothing
    Close (NFich)
    If Regs > 0 Then
        GeneraFicheroNorma68 = True
    Else
        MsgBox "No se han leido registros en la tabala de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function





Private Sub Cabecera1_68(NF As Integer, ByRef CodOrde As String, Fecha As Date, IBAN As String, ByRef Cad As String)

    Cad = "03"
    Cad = Cad & "59"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "001"
    
    Cad = Cad & Format(Fecha, "ddmmyy")
    
    'Cuenta bancaria
    Cad = Cad & Space(9)
    Cad = Cad & IBAN
    Cad = RellenaABlancos(Cad, True, 100)
    Cad = Mid(Cad, 1, 100)
    Print #NF, Cad
End Sub







Private Sub Linea1_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "010"
    If IsNull(RS1!razosoci) Then
        Cad = Cad & RellenaABlancos(RS1!Nommacta, True, 40)
    Else
        Cad = Cad & RellenaABlancos(RS1!razosoci, True, 40)
    End If
    Cad = RellenaABlancos(Cad, True, 100)
    Cad = Mid(Cad, 1, 100)
    Print #NF, Cad
End Sub


Private Sub Linea2_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "011"
    Cad = Cad & RellenaABlancos(DBLet(RS1!dirdatos, "T"), True, 45)
    Cad = RellenaABlancos(Cad, True, 100)
    Cad = Mid(Cad, 1, 100)
    Print #NF, Cad
End Sub





Private Sub Linea3_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "012"
    Cad = Cad & RellenaABlancos(DBLet(RS1!codposta, "T"), True, 5) & " "
    Cad = Cad & RellenaABlancos(DBLet(RS1!despobla, "T"), True, 40)
    Cad = RellenaABlancos(Cad, True, 100)
    Cad = Mid(Cad, 1, 100)
    Print #NF, Cad
End Sub

Private Sub Linea4_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "013"
    'De mommento pongo balancos, ya que es para extranjero
    'Cad = Cad & RellenaABlancos(DBLet(RS1!codposta, "T"), False, 5) & " "
    Cad = Cad & "     "
    Cad = Cad & RellenaABlancos(DBLet(RS1!desprovi, "T"), True, 30)   'desprovi,pais
    Cad = Cad & RellenaABlancos(DBLet(RS1!PAIS, "T"), True, 20)   'desprovi,pais
    Cad = RellenaABlancos(Cad, True, 100)
    Cad = Mid(Cad, 1, 100)
    Print #NF, Cad
End Sub

Private Sub Linea5_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String, ByRef Fechadoc As Date, ByRef Importe1 As Currency)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "014"

    Cad = Cad & "00000000" 'Numero de pago domiciliado
    Cad = Cad & Format(Fechadoc, "ddmmyyyy") 'fecha pago
   
    Cad = Cad & RellenaAceros(CStr(Round(Importe1, 2) * 100), False, 12)
    Cad = Cad & "0" 'presentacion
    Cad = Cad & "ES1" 'presentacion
    Cad = RellenaABlancos(Cad, True, 100)
    Cad = Mid(Cad, 1, 99) & "1"
    Print #NF, Cad
End Sub


Private Sub Linea6_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Importe1 As Currency, ByRef Cad As String, vConceptoTransferencia As String)


   
    '
    Cad = CodOrde   'llevara tb la ID del socio
    Cad = Cad & "015"
    Cad = Cad & "00000000" 'Numero de pago domiciliado
    Cad = Cad & RellenaABlancos(RS1!numfactu, False, 12)
    Cad = Cad & Format(RS1!fecfactu, "ddmmyyyy") 'fecha fac

    Cad = Cad & RellenaAceros(CStr(Round(Importe1, 2) * 100), False, 12)
    
    Cad = Cad & "H"
    'Cad = Cad & RellenaABlancos(vConceptoTransferencia, False, 26)
    Cad = Cad & "ADJUNTAMOS PAGO FACTURA     "
    Cad = RellenaABlancos(Cad, True, 100)
    Cad = Mid(Cad, 1, 100)
    Print #NF, Cad
End Sub



Private Sub Totales68(NF As Integer, ByRef CodOrde As String, Total As Currency, Registros As Integer, ByRef Cad As String)
    Cad = "08" & "59"
    Cad = Cad & CodOrde    'llevara tb la ID del socio
    Cad = Cad & Space(15)
    Cad = Cad & RellenaAceros(CStr(Int(Round(Total * 100, 2))), False, 12)
    'Cad = Cad & RellenaAceros(CStr(Registros), False, 8)
    Cad = Cad & RellenaAceros(CStr((Registros * 6) + 1 + 1), False, 10)
    Cad = RellenaABlancos(Cad, True, 100)
    Print #NF, Cad
End Sub

