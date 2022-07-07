Attribute VB_Name = "Functions"
Option Private Module
Public contador As Integer
'---------------------------------------------------------------------------------------
' Module    : Functions
' Author    : MVP, Sergio Alejandro Campos
' Date      : 21/sep/2015
' Purpose   : Funciones para permitir sólo texto, número y números con decimales
'---------------------------------------------------------------------------------------
'
Sub MostrarFormulario()
'
    frmValidaciones.Show
    '
End Sub
'
Function SoloTexto(Texto As Variant)
'
    Dim Caracter As Variant
    Dim Largo As String
    On Error Resume Next
    Largo = Len(Texto)
    '
    For i = 1 To Largo
        Caracter = CInt(Mid(Texto, i, 1))
        '
        If Caracter <> "" Then
            If Not Application.WorksheetFunction.IsText(Caracter) Then
                Texto = Replace(Texto, Caracter, "")
                SoloTexto = Texto
            Else
            End If
        End If
        '
    Next i
    '
    SoloTexto = Texto
    On Error GoTo 0
    '
End Function
'
Function SoloNumero(Texto As Variant)
'
    Dim Caracter As Variant
    Dim Largo As Integer
    On Error Resume Next
    Largo = Len(Texto)
    '
    For i = 1 To Largo
        Caracter = Mid(CStr(Texto), i, 1)
        '
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Texto = Replace(Texto, Caracter, "")
                SoloNumero = Texto
            Else
            End If

        End If
        '
    Next i
    '
    SoloNumero = Texto
    On Error GoTo 0
    '
End Function
'
Function SoloNumeroDecimal(Texto As Variant, Optional Separador As String)
'---------------------------------------------------------------------------------------
' Module    : Functions
' Author    : MVP, Sergio Alejandro Campos
' Date      : 21/sep/2015
' Purpose   : Funciones para permitir sólo texto, número y números con decimales
'---------------------------------------------------------------------------------------
    Dim Caracter As Variant
    Dim Largo As Integer
    On Error Resume Next
    Punto = 0
' ----------------------------------------------------------------------------------------------------------- '
' --- Author    : Milagros Huerta Gómez de Merodio                                                        --- '
' --- Modifica  : 01/jul/2022                                                                             --- '
' --- Introduzco en la Función de Sergio la variable SEPARADOR, para que se pueda escoger si PUNTO o COMA --- '
' ----------------------------------------------------------------------------------------------------------- '
' --- Si el separador puesto es incorrecto, o no se pone, utiliza por defecto el PUNTO                    --- '
    
    If Separador <> "," And Separador <> "." Then Separador = "."

' ----------------------------------------------------------------------------------------------------------- '
' --- Además, quito los CEROS de delante del SEPARADOR, por si se pusieran ---------------------------------- '
' ----------------------------------------------------------------------------------------------------------- '
    If Texto = Separador Or Texto = 0 Then      ' Si empieza por CERO o por el separador, pone "0," o "0."
        Texto = 0 & Separador
    ElseIf Left(Texto, 1) = 0 And Mid(Texto, 2, 1) <> Separador Then
        Texto = Mid(Texto, 2, Len(Texto))
    ElseIf Left(Texto, 1) = Separador Then
        Texto = 0 & Separador & Mid(Texto, 2, Len(Texto))
    End If
' ----------------------------------------------------------------------------------------------------------- '
    Largo = Len(Texto)
    '
    For i = 1 To Largo
        Caracter = Mid(CStr(Texto), i, 1)
        If Caracter <> "" Then
            '
            If Caracter = Separador Then
                Punto = Punto + 1
                If Punto > 1 Then
                    Texto = WorksheetFunction.Replace(Texto, i, 1, "")
                    SoloNumeroDecimal = Texto
                    Punto = 0
                End If
            Else
                If Caracter < Chr(48) Or Caracter > Chr(57) Then
                    Texto = Replace(Texto, Caracter, "")
                    SoloNumeroDecimal = Texto
                Else
                End If
                '
            End If
            '
        End If
    Next i
    '
    SoloNumeroDecimal = Texto
    On Error GoTo 0
    '
End Function
'---------------------------------------------------------------------------------------
' Module    : Functions
' Author    : Milagros Huerta Gómez de Merodio
' Date      : 29/jun/2022
' Purpose   : Funcion para permitir números enteros positivos y negativos
'---------------------------------------------------------------------------------------
Function NumeroEntero(Texto_E As Variant, Optional ValorMax As Variant, Optional ValorMin As Variant)
Dim Texto_N As String
    On Error Resume Next
    If Left(Texto_E, 1) = "-" Then
        If Len(Texto_E) > 1 Then
            Texto_N = Mid(Texto_E, 2, Len(Texto_E))
            ' ---------------------------------------------------------------------- '
            ' --- Quito los CEROS de detrás del signo MENOS , por si se pusieran --- '
            ' ---------------------------------------------------------------------- '
            If Left(Texto_N, 1) = 0 Then Texto_N = Mid(Texto_N, 2, Len(Texto_N))
            Texto_E = "-" & SoloNumero(Texto_N)
        End If
    Else
        If Left(Texto_E, 1) = 0 And Len(Texto_E) > 1 Then Texto_E = Mid(Texto_E, 2, Len(Texto_E))
        Texto_E = SoloNumero(Texto_E)
    End If
    If Len(Texto_E) > 15 Or Texto_E = "-" Or Texto_E = "" Then
        NumeroEntero = Texto_E
    Else
        If IsMissing(ValorMax) And IsMissing(ValorMin) Then
            NumeroEntero = Texto_E
        ElseIf IsMissing(ValorMin) Then
                NumeroEntero = Application.Min(Texto_E, ValorMax)
        ElseIf IsMissing(ValorMax) Then
            NumeroEntero = Application.Max(Texto_E, ValorMin)
        Else
            NumeroEntero = Application.Max(Application.Min(Texto_E, ValorMax), ValorMin)
        End If
    End If

    On Error GoTo 0
End Function
'---------------------------------------------------------------------------------------
' Module    : Functions
' Author    : Milagros Huerta Gómez de Merodio
' Date      : 29/jun/2022
' Purpose   : Funcion para permitir números enteros positivos y negativos
'---------------------------------------------------------------------------------------
Function NumeroDecimal(Texto_E As Variant, Separador As String, Optional N_Decimales As Integer, Optional ValorMax As Variant, Optional ValorMin As Variant)
Dim Texto_N As String
'Dim N_Decimales As Integer
    On Error Resume Next
    If Left(Texto_E, 1) = "-" Then
        If Len(Texto_E) > 1 Then
            Texto_N = Mid(Texto_E, 2, Len(Texto_E))
            ' ---------------------------------------------------------------------- '
            ' --- Quito los CEROS de detrás del signo MENOS , por si se pusieran --- '
            ' ---------------------------------------------------------------------- '
            If Left(Texto_N, 1) = 0 Then Texto_N = Mid(Texto_N, 2, Len(Texto_N))
            Texto_E = "-" & SoloNumeroDecimal(Texto_E, Separador)
        End If
    Else
        If Left(Texto_E, 1) = 0 And Len(Texto_E) > 1 Then Texto_E = Mid(Texto_E, 2, Len(Texto_E))
        Texto_E = SoloNumeroDecimal(Texto_E, Separador)
    End If
' Saca la cadena de números antes del separador de decimales, para que no sea un número de más de 15 dígitos
    If Len(InStr(Texto_E, Separador)) > 15 Or InStr(Texto_E, Separador) > 15 Or Texto_E = "-" Or Texto_E = "" Or Right(Texto_E, 1) = "0" Then
        NumeroDecimal = Texto_E
    Else
        If (IsMissing(ValorMax) And IsMissing(ValorMin)) Or Right(Texto_E, 1) = Separador Then
            NumeroDecimal = Texto_E
        ElseIf IsMissing(ValorMin) Then
                ValorMax = CLng(ValorMax)
                If Separador = "," Then
                    NumeroDecimal = Application.Min(Replace(Texto_E, ",", "."), ValorMax)
                Else
                    NumeroDecimal = Replace(Application.Min(Texto_E, ValorMax), ",", ".")
                End If
        ElseIf IsMissing(ValorMax) Then
                ValorMin = CLng(ValorMin)
                If Separador = "," Then
                    NumeroDecimal = Application.Max(Replace(Texto_E, ",", "."), ValorMin)
                Else
                    NumeroDecimal = Replace(Application.Max(Texto_E, ValorMin), ",", ".")
                End If
        Else
                ValorMax = CLng(ValorMax)
                ValorMin = CLng(ValorMin)
                If Separador = "," Then
                    NumeroDecimal = Application.Max(Application.Min(Replace(Texto_E, ",", "."), ValorMax), ValorMin)
                Else
                    NumeroDecimal = Replace(Application.Max(Application.Min(Texto_E, ValorMax), ValorMin), ",", ".")
                End If
        End If
    End If
    If N_Decimales <> 0 And InStr(Texto_E, Separador) <> 0 Then
            NumeroDecimal = Left(NumeroDecimal, InStr(Texto_E, Separador) + N_Decimales)
    End If
    
    On Error GoTo 0
End Function
