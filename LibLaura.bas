Attribute VB_Name = "LibLaura"
Option Explicit

Public Sub DesplazamientoVisible(ByRef toolb As Toolbar, iniBoton As Byte, Bol As Boolean, nreg As Byte)
'Oculta o Muestra las botones de desplazamiento de la toolbar
Dim I As Byte
 

    Select Case nreg
        Case 0, 1
            For I = iniBoton To iniBoton + 3
                toolb.Buttons(I).Visible = False
            Next I
        Case Else
            For I = iniBoton To iniBoton + 3
                toolb.Buttons(I).Visible = Bol
            Next I
    End Select
End Sub

 

 

Public Sub BloquearText1(ByRef formulario As Form, Modo As Byte)

'Bloquea controles q se llamen TEXT1 si no estamos en Modo: 3.-Insertar, 4.-Modificar

'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización

'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)

Dim I As Byte
Dim B As Boolean
On Error Resume Next

 

    With formulario
        B = (Modo = 3 Or Modo = 4)
        For I = 0 To .Text1.Count - 1
            .Text1(I).Locked = ((Not B) And (Modo <> 1))

            .Text1(I).BackColor = vbWhite 'Blanco

            If Modo = 3 Then .Text1(I).Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)

        Next I

    End With

    If Err.Number <> 0 Then Err.Clear

End Sub

 

 

Public Sub ConseguirFoco(ByRef Text As TextBox, Modo As Byte)

'Acciones que se realizan en el evento:GotFocus de los TextBox:Text1

'en los formularios de Mantenimiento

On Error Resume Next

 

    If (Modo <> 0 And Modo <> 2) Then

        If Modo = 1 Then 'Modo 1: Busqueda

            Text.BackColor = vbYellow

        End If

        Text.SelStart = 0

        Text.SelLength = Len(Text.Text)

    End If

    If Err.Number <> 0 Then Err.Clear

End Sub

 

Public Function PerderFocoGnral(ByRef Text As TextBox, Modo As Byte) As Boolean

Dim Comprobar As Boolean

On Error Resume Next

    With Text

    

        'Quitamos blancos por los lados

        .Text = Trim(.Text)

        

        If .BackColor = vbYellow Then

    '        Text1(Index).BackColor = &H80000018

            .BackColor = vbWhite

        End If

        

        'Si no estamos en modo: 3=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion

        If (Modo <> 3 And Modo <> 4 And Modo <> 1) Then

            PerderFocoGnral = False

            Exit Function

        End If

        

        If Modo = 1 Then

            'Si estamos en modo busqueda y contiene un caracter especial no realizar

            'las comprobaciones

            Comprobar = ContieneCaracterBusqueda(.Text)

            If Comprobar Then

                PerderFocoGnral = False

                Exit Function

            End If

        End If

        PerderFocoGnral = True

    End With

    If Err.Number <> 0 Then Err.Clear

End Function

 




Public Sub Ponerfoco(ByRef T As TextBox)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerfocoObj(ByRef Obj As Object)
    On Error Resume Next
    Obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub KEYdown(KeyCode As Integer)

'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.

On Error Resume Next

    Select Case KeyCode

        Case 38 'Desplazamieto Fecha Hacia Arriba

            SendKeys "+{tab}"

        Case 40 'Desplazamiento Flecha Hacia Abajo

            SendKeys "{tab}"

    End Select

    If Err.Number <> 0 Then Err.Clear

End Sub





Public Function ContieneCaracterBusqueda(CADENA As String) As Boolean

'Comprueba si la cadena contiene algun caracter especial de busqueda

' >,>,>=,: , ....

'si encuentra algun caracter de busqueda devuelve TRUE y sale

Dim B As Boolean
Dim I As Integer
Dim Ch As String
 

    'For i = 1 To Len(cadena)

    I = 1
    B = False
    Do
        Ch = Mid(CADENA, I, 1)
        Select Case Ch
            Case "<", ">", ":", "="
                B = True
            Case "*", "%", "?", "_", "\", ":" ', "."
                B = True
            Case Else
                B = False
        End Select

    'Next i

        I = I + 1

    Loop Until (B = True) Or (I > Len(CADENA))

    ContieneCaracterBusqueda = B

End Function



Public Sub PonerIndicador(ByRef lblIndicador As Label, Modo As Byte)
'Pone el titulo del label lblIndicador
    Select Case Modo
        Case 0    'Modo Inicial
            lblIndicador.Caption = ""
        Case 1 'Modo Buscar
            lblIndicador.Caption = "BUSQUEDA"
        Case 2    'Preparamos para que pueda Modificar

        Case 3 'Modo Insertar
            lblIndicador.Caption = "INSERTAR"
        Case 4 'MODIFICAR
            lblIndicador.Caption = "MODIFICAR"
        Case Else
            lblIndicador.Caption = ""
    End Select
End Sub




Public Sub DesplazamientoData(ByRef vData As Adodc, Index As Integer)
'Para desplazarse por los registros de control Data
    If vData.Recordset.EOF Then Exit Sub
    Select Case Index
        Case 0 'Primer Registro
            If Not vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 1 'Anterior
            vData.Recordset.MovePrevious
            If vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 2 'Siguiente
            vData.Recordset.MoveNext
            If vData.Recordset.EOF Then vData.Recordset.MoveLast
        Case 3 'Ultimo
            vData.Recordset.MoveLast
    End Select

End Sub




'Esto es para que cuando pincha en siguiente le sugerimos

'Se puede comentar todo y asi no hace nada ni da error

'El SQL es propio de cada tabla

Public Function SugerirCodigoSiguienteStr(NomTabla As String, NomCodigo As String, Optional CondLineas As String) As String
Dim SQL As String
Dim Rs As ADODB.Recordset
On Error GoTo ESugerirCodigo

    SQL = "Select Max(" & NomCodigo & ") from " & NomTabla
    If CondLineas <> "" Then
        SQL = SQL & " WHERE " & CondLineas

    End If

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, , , adCmdText
    SQL = "1"
    If Not Rs.EOF Then

        If Not IsNull(Rs.Fields(0)) Then

            If IsNumeric(Rs.Fields(0)) Then

                SQL = CStr(Rs.Fields(0) + 1)

            Else

                If Asc(Left(Rs.Fields(0), 1)) <> 122 Then 'Z

                SQL = Left(Rs.Fields(0), 1) & CStr(Asc(Right(Rs.Fields(0), 1)) + 1)

                End If

            End If

        End If

    End If

    Rs.Close

    Set Rs = Nothing

    SugerirCodigoSiguienteStr = SQL

ESugerirCodigo:

    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation

End Function

 

 

Public Sub BloquearTxt(ByRef Text As TextBox, B As Boolean, Optional EsContador As Boolean)

'Bloquea un control de tipo TextBox

'Si lo bloquea lo poner de color amarillo claro sino lo pone en color blanco (sino es contador)

'pero si es contador lo pone color azul claro

On Error Resume Next
    Text.Locked = B
    If Not B And Text.Enabled = False Then Text.Enabled = True
    If B Then
        If EsContador Then
            'Si Es un campo que se obtiene de un contador poner color azul
            Text.BackColor = &H80000013 'Azul Claro
        Else
            Text.BackColor = &H80000018 'Amarillo Claro
        End If
    Else
        Text.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

 

 

Public Function SituarDataTrasEliminar(ByRef vData As Adodc, NumReg) As Boolean
On Error GoTo ESituarDataElim

        vData.Refresh
        If Not vData.Recordset.EOF Then    'Solo habia un registro
            If NumReg > vData.Recordset.RecordCount Then
                vData.Recordset.MoveLast
            Else
                vData.Recordset.MoveFirst
                vData.Recordset.Move NumReg - 1
            End If

            SituarDataTrasEliminar = True
        Else
            SituarDataTrasEliminar = False
        End If

ESituarDataElim:

    If Err.Number <> 0 Then

        Err.Clear

        SituarDataTrasEliminar = False

    End If

End Function

Public Function SituarData(ByRef vData1 As Adodc, vWhere As String, Indicador As String) As Boolean
    On Error GoTo ES
    SituarData = False
    vData1.Recordset.Find vWhere
    If vData1.Recordset.EOF Then Exit Function
    SituarData = True
    Exit Function
ES:
    Err.Clear
End Function


Public Sub KeyPressGral(ByRef KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub
