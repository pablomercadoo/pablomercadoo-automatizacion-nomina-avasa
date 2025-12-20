Attribute VB_Name = "modCalendario"
Option Explicit

Private dictFestivos As Object          ' key: CLng(Date)  value: TRUE
Private dictFestivoNombre As Object     ' key: CLng(Date)  value: Nombre

'==============================
'  API PÚBLICA
'==============================
Public Sub AplicarCalendarioAMatriz(ByVal wsMatriz As Worksheet)
    Dim rngFechas As Range
    Set rngFechas = GetHeaderFechasRange(wsMatriz)

    If Not rngFechas Is Nothing Then
        PintarDiasEspeciales wsMatriz, rngFechas
    End If
End Sub

'==============================
'  CORE
'==============================
Public Function GetDiaTipo(ByVal dt As Date) As String
    If EsFestivo(dt) Then
        GetDiaTipo = "DF"
    ElseIf Weekday(dt, vbMonday) = 7 Then
        GetDiaTipo = "PD"
    Else
        GetDiaTipo = "NORMAL"
    End If
End Function

Public Function EsFestivo(ByVal dt As Date) As Boolean
    EsFestivo = False
    If dictFestivos Is Nothing Then Exit Function
    EsFestivo = dictFestivos.Exists(CLng(DateValue(dt)))
End Function

Public Function GetFestivoNombre(ByVal dt As Date) As String
    GetFestivoNombre = ""
    If dictFestivoNombre Is Nothing Then Exit Function

    Dim k As Long
    k = CLng(DateValue(dt))

    If dictFestivoNombre.Exists(k) Then
        GetFestivoNombre = CStr(dictFestivoNombre(k))
    End If
End Function

'==============================
'  MATRIZ
'==============================
Private Function GetHeaderFechasRange(ByVal ws As Worksheet) As Range
    Dim startCol As Long, c As Long, lastCol As Long

    startCol = 9 ' Columna I
    lastCol = startCol

    ' Avanzar mientras haya números (días)
    For c = startCol To ws.Columns.Count
        If IsNumeric(ws.Cells(2, c).Value) Then
            lastCol = c
        Else
            Exit For
        End If
    Next c

    If lastCol < startCol Then Exit Function

    Set GetHeaderFechasRange = ws.Range(ws.Cells(2, startCol), ws.Cells(2, lastCol))
End Function

Private Sub PintarDiasEspeciales(ByVal ws As Worksheet, ByVal headerRange As Range)
    Dim c As Range
    Dim tipo As String
    Dim fechaReal As Date
    Dim dia As Long

    'IMPORTANTE: si hay Formato Condicional, puede estar sobre-escribiendo el color que pongas por VBA
    On Error Resume Next
    headerRange.FormatConditions.Delete
    On Error GoTo 0

    For Each c In headerRange.Cells

        If IsNumeric(c.Value) Then
            dia = CLng(c.Value)

            'Color base SIEMPRE (amarillo)
            c.Interior.Color = RGB(240, 200, 80)
            c.Font.Bold = True
            c.Font.Italic = False

            'Fecha real del periodo
            fechaReal = DateSerial(gAnio, gMes, dia)
            tipo = GetDiaTipo(fechaReal)

            Select Case tipo
                Case "DF"
                    'Festivo
                    c.Interior.Color = RGB(255, 180, 180) 'rojo suave

                Case "PD"
                    'Domingo
                    c.Interior.Color = RGB(220, 220, 220) 'gris suave
            End Select
        End If

    Next c
End Sub

'==============================
'  CARGA FESTIVOS A MEMORIA
'==============================
Public Sub CargarFestivosEnMemoria()
    Dim lo As ListObject
    Dim lr As ListRow
    Dim vFecha As Variant, vNombre As Variant, vActivo As Variant
    Dim f As Date, key As Long

    Set dictFestivos = CreateObject("Scripting.Dictionary")
    Set dictFestivoNombre = CreateObject("Scripting.Dictionary")

    Set lo = ThisWorkbook.Worksheets("Config").ListObjects("tblFestivos")

    For Each lr In lo.ListRows
        vFecha = lr.Range(1, 1).Value  ' Fecha
        vNombre = lr.Range(1, 2).Value ' Nombre
        vActivo = lr.Range(1, 5).Value ' Activo

        If IsDate(vFecha) Then
            f = DateValue(CDate(vFecha))
            key = CLng(f)

            If EsVerdadero(vActivo) Then
                dictFestivos(key) = True
                dictFestivoNombre(key) = CStr(vNombre)
            End If
        End If
    Next lr
End Sub

Private Function EsVerdadero(ByVal v As Variant) As Boolean
    On Error GoTo Falso
    If IsError(v) Or IsEmpty(v) Then GoTo Falso

    If VarType(v) = vbBoolean Then
        EsVerdadero = (v = True)
        Exit Function
    End If

    If IsNumeric(v) Then
        EsVerdadero = (CLng(v) <> 0)
        Exit Function
    End If

    Select Case UCase$(Trim$(CStr(v)))
        Case "TRUE", "VERDADERO", "SI", "SÍ", "1", "X"
            EsVerdadero = True
        Case Else
            EsVerdadero = False
    End Select
    Exit Function

Falso:
    EsVerdadero = False
End Function

