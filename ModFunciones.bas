Attribute VB_Name = "ModFunciones"
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
'   En este modulo estan las funciones que recorren el form
'   usando el each for
'   Estas son
'
'   CamposSiguiente -> Nos devuelve el el text siguiente en
'           el orden del tabindex
'
'   CompForm -> Compara los valores con su tag
'
'   InsertarDesdeForm - > Crea el sql de insert e inserta
'
'   Limpiar -> Pone a "" todos los objetos text de un form
'
'   ObtenerBusqueda -> A partir de los text crea el sql a
'       partir del WHERE ( sin el).
'
'   ModifcarDesdeFormulario -> Opcion modificar. Genera el SQL
'
'   PonerDatosForma -> Pone los datos del RECORDSET en el form
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
Option Explicit

Public Const ValorNulo = "Null"

Public Function CompForm(ByRef formulario As Form) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Carga As Boolean
    Dim Correcto As Boolean
       
    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            Carga = mTag.Cargar(Control)
            If Carga = True Then
                Correcto = mTag.Comprobar(Control)
                If Not Correcto Then Exit Function
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                    
                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                            Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function


Public Sub Limpiar(ByRef formulario As Form)
    Dim Control As Object
    
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        Else
            If TypeOf Control Is ComboBox Then
                Control.ListIndex = -1
            End If
        End If
    Next Control
End Sub


Public Function CampoSiguiente(ByRef formulario As Form, Valor As Integer) As Control
Dim Fin As Boolean
Dim Control As Object

On Error GoTo ECampoSiguiente

    'Debug.Print "Llamada:  " & Valor
    'Vemos cual es el siguiente
    Do
        Valor = Valor + 1
        For Each Control In formulario.Controls
            'Debug.Print "-> " & Control.Name & " - " & Control.TabIndex
            'Si es texto monta esta parte de sql
            If Control.TabIndex = Valor Then
                    Set CampoSiguiente = Control
                    Fin = True
                    Exit For
            End If
        Next Control
        If Not Fin Then
            Valor = -1
        End If
    Loop Until Fin
    Exit Function
ECampoSiguiente:
    Set CampoSiguiente = Nothing
    Err.Clear
End Function




Private Function ValorParaSQL(Valor, ByRef vTag As CTag) As String
Dim Dev As String
Dim D As Single
Dim I As Integer
Dim v
    Dev = ""
    If Valor <> "" Then
        Select Case vTag.TipoDato
        Case "N"
            v = Valor
            If InStr(1, Valor, ",") Then
                If InStr(1, Valor, ".") Then
                    'ABRIL 2004
                
                    'Ademas de la coma lleva puntos
                    v = ImporteFormateado(CStr(Valor))
                    Valor = v
                Else
                
                    v = CSng(Valor)
                    Valor = v
                End If
            Else
         
            End If
            Dev = TransformaComasPuntos(CStr(Valor))
            
        Case "F"
            Dev = "'" & Format(Valor, FormatoFecha) & "'"
        Case "T"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & Valor & "'"
        End Select
        
    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vTag.Vacio = "S" Then Dev = ValorNulo
    End If
    ValorParaSQL = Dev
End Function

Public Function InsertarDesdeForm(ByRef formulario As Form) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Izda As String
    Dim Der As String
    Dim cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Debug.Print Control.Name
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            'Debug.Print Control.Tag
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.Columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & ""
                    
                        'Parte VALUES
                        cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.Columna & ""
                If Control.Value = 1 Then
                    cad = "1"
                    Else
                    cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                Der = Der & cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.Columna & ""
                    If Control.ListIndex = -1 Then
                        cad = ValorNulo
                        Else
                        cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.Tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
    
   
    Conn.Execute cad, , adCmdText
    
    
    InsertarDesdeForm = True
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function





Public Function PonerCamposForma(ByRef formulario As Form, ByRef vData As Adodc) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim cad As String
    Dim Valor As Variant
    Dim Campo As String  'Campo en la base de datos
    Dim I As Integer


    On Error GoTo EPonerCamposForma

    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In formulario.Controls
        'TEXTO
        Debug.Print Control.Tag
        If TypeOf Control Is TextBox Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                
                    'If mTag.Columna = "entidad" Then Stop
                    'Columna en la BD
                    If mTag.Columna <> "" Then
                        Campo = mTag.Columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(Campo))
                        Else
                            Valor = vData.Recordset.Fields(Campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                cad = Format(Valor, mTag.Formato)
                                Control.Text = cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = Valor
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    Campo = mTag.Columna
                    If mTag.Vacio = "S" Then
                        Valor = DBLet(vData.Recordset.Fields(Campo), mTag.TipoDato)
                    Else
                        Valor = vData.Recordset.Fields(Campo)
                    End If
                    Else
                        Valor = 0
                End If
                Control.Value = Valor
            End If
            
         'COMBOBOX
         ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Campo = mTag.Columna
                    If mTag.Vacio = "S" Then
                        If IsNull(vData.Recordset.Fields(Campo)) Then
                            Valor = -1
                        Else
                            Valor = vData.Recordset.Fields(Campo)
                        End If
                    Else
                        Valor = vData.Recordset.Fields(Campo)
                    End If
                    I = 0
                    For I = 0 To Control.ListCount - 1
                        If Control.ItemData(I) = Val(Valor) Then
                            Control.ListIndex = I
                            Exit For
                        End If
                    Next I
                    If I = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control
    
    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. " & mTag.Columna
End Function

Private Function ObtenerMaximoMinimo(ByRef vSQL As String) As String
Dim Rs As Recordset
ObtenerMaximoMinimo = ""
Set Rs = New ADODB.Recordset
Rs.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not Rs.EOF Then
    If Not IsNull(Rs.EOF) Then
        ObtenerMaximoMinimo = CStr(Rs.Fields(0))
    End If
End If
Rs.Close
Set Rs = Nothing
End Function


Public Function ObtenerBusqueda(ByRef formulario As Form, Optional CHECK As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim cad As String
    Dim SQL As String
    Dim Tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        cad = " MAX(" & mTag.Columna & ")"
                    Else
                        cad = " MIN(" & mTag.Columna & ")"
                    End If
                    SQL = "Select " & cad & " from " & mTag.Tabla
                    SQL = ObtenerMaximoMinimo(SQL)
                    Select Case mTag.TipoDato
                    Case "N"
                        SQL = mTag.Tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
                    Case "F"
                        SQL = mTag.Tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                    Case Else
                        SQL = mTag.Tabla & "." & mTag.Columna & " = '" & SQL & "'"
                    End Select
                    SQL = "(" & SQL & ")"
                End If
            End If
        End If
    Next

    
    
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    SQL = mTag.Tabla & "." & mTag.Columna & " is NULL"
                    SQL = "(" & SQL & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            'Cargamos el tag
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.Tabla <> "" Then
                            Tabla = mTag.Tabla & "."
                            Else
                            Tabla = ""
                        End If
                        RC = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.Columna, Aux, cad)
                        If RC = 0 Then
                            If SQL <> "" Then SQL = SQL & " AND "
                            SQL = SQL & "(" & cad & ")"
                        End If
                    End If
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
            
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            mTag.Cargar Control
            If mTag.Cargado Then
                If Control.ListIndex > -1 Then
                    cad = Control.ItemData(Control.ListIndex)
                    cad = mTag.Tabla & "." & mTag.Columna & " = " & cad
                    If SQL <> "" Then SQL = SQL & " AND "
                    SQL = SQL & "(" & cad & ")"
                End If
            End If
        
        
        'CHECK
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    
                    Aux = ""
                    If CHECK <> "" Then
                        Tabla = DBLet(Control.Index, "T")
                        If Tabla <> "" Then Tabla = "(" & Tabla & ")"
                        Tabla = Control.Name & Tabla & "|"
                        If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
                        cad = mTag.Tabla & "." & mTag.Columna & " = " & Aux
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & cad & ")"
                    End If
                End If
            End If
        End If
        
    Next Control
    ObtenerBusqueda = SQL
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. "
End Function




Public Function ModificaDesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.Columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        'Debug.Print mTag.Columna
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.Columna & " = " & Aux & ")"
                             
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    Conn.Execute Aux, , adCmdText






ModificaDesdeFormulario = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Function ParaGrid(ByRef Control As Control, AnchoPorcentaje As Integer, Optional Desc As String) As String
Dim mTag As CTag
Dim cad As String

'Montamos al final: "Cod Diag.|idDiag|N|10·"

ParaGrid = ""
cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Desc <> "" Then
                cad = Desc
            Else
                cad = mTag.Nombre
            End If
            cad = cad & "|"
            cad = cad & mTag.Columna & "|"
            cad = cad & mTag.TipoDato & "|"
            cad = cad & AnchoPorcentaje & "·"
            
                
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            
        ElseIf TypeOf Control Is ComboBox Then
        
        
        End If 'De los elseif
    End If
Set mTag = Nothing
ParaGrid = cad
End If



End Function

'////////////////////////////////////////////////////
' Monta a partir de una cadena devuelta por el formulario
'de busqueda el sql para situar despues el datasource
Public Function ValorDevueltoFormGrid(ByRef Control As Control, ByRef CadenaDevuelta As String, Orden As Integer) As String
Dim mTag As CTag
Dim cad As String
Dim Aux As String
'Montamos al final: " columnatabla = valordevuelto "

ValorDevueltoFormGrid = ""
cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            Aux = RecuperaValor(CadenaDevuelta, Orden)
            If Aux <> "" Then cad = mTag.Columna & " = " & ValorParaSQL(Aux, mTag)
                
            
            
                
        'CheckBOX
       ' ElseIf TypeOf Control Is CheckBox Then
       '
       ' ElseIf TypeOf Control Is ComboBox Then
       '
       '
        End If 'De los elseif
    End If
End If
Set mTag = Nothing
ValorDevueltoFormGrid = cad
End Function


Public Sub FormateaCampo(vTex As TextBox, ByRef mTag As CTag)
    
    Dim cad As String
    On Error GoTo EFormateaCampo
    
    
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                cad = TransformaPuntosComas(vTex.Text)
                cad = Format(cad, mTag.Formato)
                vTex.Text = cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
 
End Sub


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef CADENA As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim CONT As Integer
Dim cad As String

I = 0
CONT = 1
cad = ""
Do
    J = I + 1
    I = InStr(J, CADENA, "|")
    If I > 0 Then
        If CONT = Orden Then
            cad = Mid(CADENA, J, I - J)
            I = Len(CADENA) 'Para salir del bucle
            Else
                CONT = CONT + 1
        End If
    End If
Loop Until I = 0
RecuperaValor = cad
End Function


'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim I As Integer
Dim J As Integer


On Error GoTo EPonerOpcionesMenuGeneral


'Añadir, modificar y borrar deshabilitados si no nivel
With formulario

    'LA TOOLBAR  .--> Requisito, k se llame toolbar1
    For I = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(I).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(I).Tag)
            If J < vUsu.Nivel Then
                .Toolbar1.Buttons(I).Enabled = False
            End If
        End If
    Next I
    
    'Esto es un poco salvaje. Por si acaso , no existe en este trozo pondremos los errores on resume next
    
    On Error Resume Next
    
    'Los MENUS
    'K sean mnAlgo
    J = Val(.mnNuevo.HelpContextID)
    If J < vUsu.Nivel Then .mnNuevo.Enabled = False
    
    J = Val(.mnModificar.HelpContextID)
    If J < vUsu.Nivel Then .mnModificar.Enabled = False
    
    J = Val(.mnEliminar.HelpContextID)
    If J < vUsu.Nivel Then .mnEliminar.Enabled = False
    On Error GoTo 0
End With




Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



'Este modifica las claves prinipales y todo
'la sentenca del WHERE cod=1 and .. viene en claves
Public Function ModificaDesdeFormularioClaves(ByRef formulario As Form, Claves As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String
Dim I As Integer

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    cadWHERE = Claves
    'Construimos el SQL
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    Conn.Execute Aux, , adCmdText






ModificaDesdeFormularioClaves = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function







Public Function BLOQUEADesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                         cadWHERE = cadWHERE & "(" & mTag.Columna & " = " & Aux & ")"
                    End If
                End If
            End If
        End If
    Next Control
    
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        
    Else
        Aux = "select * FROM " & mTag.Tabla
        Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"
        
        'Intenteamos bloquear
        PreparaBloquear
        Conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function




Public Function BloqueaRegistroForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    Next Control
    
    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        
    Else
        Aux = "Insert into zBloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.Tabla
        Aux = Aux & "',""" & AuxDef & """)"
        Conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If Conn.Errors.Count > 0 Then
            If Conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
        SQL = "DELETE from zBloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.Tabla & "'"
        Conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function





Public Function SeparaCampoBusqueda(Tipo As String, Campo As String, CADENA As String, ByRef DevSQL As String) As Byte
Dim cad As String
Dim Aux As String
Dim Ch As String
Dim Fin As Boolean
Dim I, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda = 1
DevSQL = ""
cad = ""
Select Case Tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    I = CararacteresCorrectos(CADENA, "N")
    If I > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo numerico
        cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not IsNumeric(cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = Campo & " >= " & cad & " AND " & Campo & " <= " & Aux
        '----
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                DevSQL = "1=1"
             Else
                    Fin = False
                    I = 1
                    cad = ""
                    Aux = "NO ES NUMERO"
                    While Not Fin
                        Ch = Mid(CADENA, I, 1)
                        If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                            cad = cad & Ch
                            Else
                                Aux = Mid(CADENA, I)
                                Fin = True
                        End If
                        I = I + 1
                        If I > Len(CADENA) Then Fin = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If cad = "" Then cad = " = "
                    DevSQL = Campo & " " & cad & " " & Aux
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(CADENA, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo de fechas
        cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not EsFechaOKString(cad) Or Not EsFechaOKString(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        cad = Format(cad, FormatoFecha)
        Aux = Format(Aux, FormatoFecha)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = Campo & " >='" & cad & "' AND " & Campo & " <= '" & Aux & "'"
        '----
        'ELSE
        Else
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                I = 1
                cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    Ch = Mid(CADENA, I, 1)
                    If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                        cad = cad & Ch
                        Else
                            Aux = Mid(CADENA, I)
                            Fin = True
                    End If
                    I = I + 1
                    If I > Len(CADENA) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOKString(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                Aux = "'" & Format(Aux, FormatoFecha) & "'"
                If cad = "" Then cad = " = "
                DevSQL = Campo & " " & cad & " " & Aux
            End If
        End If
    
    
    
    
Case "T"
    '---------------- TEXTO ------------------
    I = CararacteresCorrectos(CADENA, "T")
    If I = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
     If CADENA = ">>" Or CADENA = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    'Cambiamos el * por % puesto que en ADO es el caraacter para like
    I = 1
    Aux = CADENA
    While I <> 0
        I = InStr(1, Aux, "*")
        If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "%" & Mid(Aux, I + 1)
    Wend
    'Cambiamos el ? por la _ pue es su omonimo
    I = 1
    While I <> 0
        I = InStr(1, Aux, "?")
        If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "_" & Mid(Aux, I + 1)
    Wend
    cad = Mid(CADENA, 1, 2)
    If cad = "<>" Then
        Aux = Mid(CADENA, 3)
        DevSQL = Campo & " LIKE '!" & Aux & "'"
        Else
        DevSQL = Campo & " LIKE '" & Aux & "'"
    End If
    


    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    I = InStr(1, CADENA, "<>")
    If I = 0 Then
        'IGUAL A valor
        cad = " = "
        Else
            'Distinto a valor
        cad = " <> "
    End If
    'Verdadero o falso
    I = InStr(1, CADENA, "V")
    If I > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = Campo & " " & cad & " " & Aux
    
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Private Function CararacteresCorrectos(vCad As String, Tipo As String) As Byte
Dim I As Integer
Dim Ch As String
Dim Error As Boolean

CararacteresCorrectos = 1
Error = False
Select Case Tipo
Case "N"
    'Numero. Aceptamos numeros, >,< = :
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "=", ".", " ", "-"
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "T"
    'Texto aceptamos numeros, letras y el interrogante y el asterisco
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "a" To "z"
            Case "A" To "Z"
            Case "0" To "9"
            Case "*", "%", "?", "_", "\", "/", ":", ".", " " ' estos son para un caracter sol no esta demostrado , "%", "&"
            'Esta es opcional
            Case "<", ">"
            Case "Ñ", "ñ"
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "F"
    'Numeros , "/" ,":"
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "B"
    'Numeros , "/" ,":"
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "=", " "
            Case Else
                Error = True
                Exit For
        End Select
    Next I
End Select
'Si no ha habido error cambiamos el retorno
If Not Error Then CararacteresCorrectos = 0
End Function





Public Function DameClavesADODCForm(ByRef formulario As Form, ByRef ado As Adodc) As String
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte
Dim Valor2 As String

On Error GoTo EBLOQ
    DameClavesADODCForm = ""
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Valor2 = ado.Recordset(mTag.Columna).Value
                        Aux = ValorParaSQL(Valor2, mTag)
                        AuxDef = AuxDef & " AND " & mTag.Columna & " = " & Aux
                    End If
                End If
            End If
        End If
    Next Control
    
    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        
    Else
        DameClavesADODCForm = Mid(AuxDef, 5)
    End If
EBLOQ:
    If Err.Number <> 0 Then

            MuestraError Err.Number, "Obteniendo claves"

           

    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Sub PonerFocoGral(ByRef OB As Object)
    On Error Resume Next
    OB.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub ObtenerFocoGral(ByRef T1 As TextBox)
    On Error Resume Next
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
    If Err.Number <> 0 Then Err.Clear
End Sub
