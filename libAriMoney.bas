Attribute VB_Name = "libAriMoney"
Option Explicit

Public Sub CommitConexion()
    On Error Resume Next
    Conn.Execute "Commit"
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub Main()
    
       Load frmIdentifica
       CadenaDesdeOtroForm = ""
       
       'Necesitaremos el archivo arifon.dat
       frmIdentifica.Show vbModal
        
       If CadenaDesdeOtroForm = "" Then
            'NO se ha identificado
            Set Conn = Nothing
            End
       End If
       
       CadenaDesdeOtroForm = ""
       frmLogin.Show vbModal
       If CadenaDesdeOtroForm = "" Then
            'No ha seleccionado nonguna empresa
            Set Conn = Nothing
            End
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass

        'Cerramos la conexion
        Conn.Close

        
        If AbrirConexion() = False Then
            MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
            End
        End If
        Screen.MousePointer = vbHourglass
        
        
        'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
        GestionaPC
        
        LeerEmpresaParametros
        
        
        'Fijamos el primer dia de la semana para que lo tulice el calendar
        FijarPrimerDiaSemana
        
        'Otras acciones
        OtrasAcciones
        Screen.MousePointer = vbHourglass
        Load frmPpal
        Screen.MousePointer = vbHourglass
        frmPpal.Show
End Sub



Public Function LeerEmpresaParametros()
        'Abrimos la empresa
        Set vEmpresa = New Cempresa
        If vEmpresa.Leer = 1 Then
            MsgBox "No se han podido cargar datos empresa. Debe configurar la aplicación.", vbExclamation
            Set vEmpresa = Nothing
        End If
        Set vParam = New Cparametros
        If vParam.Leer = 1 Then
            MsgBox "No se han podido leer los parametros de la empresa", vbExclamation
            Set vParam = Nothing
        End If
    End Function




Public Function Ejecuta(ByRef SQL As String) As Boolean

    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cadena: " & SQL & vbCrLf
        Ejecuta = False
    Else
        Ejecuta = True
    End If
End Function




