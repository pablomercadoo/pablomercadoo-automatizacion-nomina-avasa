Attribute VB_Name = "modAdmin"
Option Explicit

Public Sub BuscarMatrizPeriodoAdmin()

    Dim codigo As String
    Dim ws As Worksheet
    Dim hojaEsperada As Worksheet
    Dim lista() As String
    Dim n As Long, i As Long
    Dim msg As String
    Dim resp As Variant
    Dim nombreEsperado As String
    
    '------------------------------------
    ' Pedir periodo: 2025_12_Q1, 2026_03_S2, etc.
    '------------------------------------
    codigo = InputBox( _
        Prompt:="Pega el periodo con el formato:" & vbCrLf & _
                "AAAA_MM_Q#  o  AAAA_MM_S#" & vbCrLf & _
                "Ejemplo:  2025_12_Q1", _
        Title:="Buscar matriz de incidencias")
    
    codigo = Trim(UCase$(codigo))
    codigo = Replace(codigo, " ", "")
    
    If codigo = "" Then Exit Sub
    
    ' Validación muy ligera: 2 guiones bajos
    If UBound(Split(codigo, "_")) <> 2 Then
        MsgBox "El texto no parece un periodo válido (ej. 2025_12_Q1).", vbExclamation
        Exit Sub
    End If
    
    '------------------------------------
    ' 1) Intento directo con el patrón "M_LOC_periodo"
    '    (si el archivo está bien hecho, cada archivo solo tiene 1 LOC)
    '------------------------------------
    On Error Resume Next
    If gLoc <> "" Then
        nombreEsperado = "M_" & gLoc & "_" & codigo
        Set hojaEsperada = ThisWorkbook.Worksheets(nombreEsperado)
    End If
    On Error GoTo 0
    
    If Not hojaEsperada Is Nothing Then
        hojaEsperada.Visible = xlSheetVisible
        hojaEsperada.Activate
        Exit Sub
    End If
    
    '------------------------------------
    ' 2) Plan B: buscar cualquier hoja que termine en "_periodo"
    '    Ej: * _2025_12_Q1
    '------------------------------------
    n = 0
    For Each ws In ThisWorkbook.Worksheets
        If UCase$(ws.Name) <> "MENU" Then
            If ws.Name Like "*_" & codigo Then
                n = n + 1
                ReDim Preserve lista(1 To n)
                lista(n) = ws.Name
            End If
        End If
    Next ws
    
    Select Case n
        Case 0
            MsgBox "No se encontró ninguna hoja cuyo nombre termine en '" & codigo & "'.", _
                   vbInformation
            Exit Sub
        
        Case 1
            Set hojaEsperada = ThisWorkbook.Worksheets(lista(1))
            hojaEsperada.Visible = xlSheetVisible
            hojaEsperada.Activate
        
        Case Else
            ' Hay varias coincidencias, que elijas una
            msg = "Se encontraron varias hojas con el periodo " & codigo & ":" & vbCrLf & vbCrLf
            For i = 1 To n
                msg = msg & i & ") " & lista(i) & vbCrLf
            Next i
            msg = msg & vbCrLf & "Escribe el número de la hoja a la que quieres ir:"
            
            resp = InputBox(msg, "Seleccionar hoja", "1")
            If resp = "" Then Exit Sub
            If Not IsNumeric(resp) Then
                MsgBox "Valor no válido.", vbExclamation
                Exit Sub
            End If
            i = CLng(resp)
            If i < 1 Or i > n Then
                MsgBox "El número no está en la lista.", vbExclamation
                Exit Sub
            End If
            
            Set hojaEsperada = ThisWorkbook.Worksheets(lista(i))
            hojaEsperada.Visible = xlSheetVisible
            hojaEsperada.Activate
    End Select

End Sub


