Attribute VB_Name = "modLostFocus"
Option Explicit




Public Function ComprobarCampoENlazado(ByRef T As TextBox, TDesc As TextBox, Tipo As String) As Byte

    T.Text = Trim(T.Text)
    If T.Text = "" Then
        ComprobarCampoENlazado = 0 'NO HA PUESTO NADA
        TDesc.Text = ""
        Exit Function
    End If
    
    Select Case Tipo
    Case "N"
        If Not IsNumeric(T.Text) Then
            MsgBox "El campo debe ser numérico: " & T.Text, vbExclamation
            TDesc.Text = ""
            T.Text = ""
            ComprobarCampoENlazado = 1
        Else
            ComprobarCampoENlazado = 2
        End If
    End Select
        
End Function
