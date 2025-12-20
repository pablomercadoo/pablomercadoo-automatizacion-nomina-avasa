Attribute VB_Name = "modMantenimientoMatrices"
Option Explicit

'====================================================
' modMantenimientoMatrices
' Limpia hojas de matriz M_... por locación,
' conservando SOLO:
'   - periodo actual (según fecha hoy)
'   - periodo anterior inmediato
' Para SEMANAL y QUINCENAL.
'====================================================

Public Sub LimpiarMatricesRelevantes(Optional ByVal locCode As String = "", Optional ByVal pwdLibro As String = "AVASA")

    Dim loc As String
    loc = Trim$(locCode)
    If loc = "" Then loc = Trim$(gLoc)
    If loc = "" Then Exit Sub

    '1) Calcular "targets" (actual y anterior) para semanal y quincenal
    Dim y As Long, m As Long
    Dim tipo As String, per As Long

    Dim keepDict As Object
    Set keepDict = CreateObject("Scripting.Dictionary") 'key: sheetname => True

    '--- QUINCENAL ---
    y = Year(Date): m = Month(Date)
    tipo = "QUINCENAL"
    per = PeriodoActualPorFecha(tipo, Date)

    AddKeep keepDict, BuildNombreMatriz(loc, y, m, tipo, per)
    AddKeep keepDict, NombreMatrizAnterior(loc, y, m, tipo, per)

    '--- SEMANAL ---
    y = Year(Date): m = Month(Date)
    tipo = "SEMANAL"
    per = PeriodoActualPorFecha(tipo, Date)

    AddKeep keepDict, BuildNombreMatriz(loc, y, m, tipo, per)
    AddKeep keepDict, NombreMatrizAnterior(loc, y, m, tipo, per)

    '2) Borrar cualquier M_ de esta loc que NO esté en keep
    Dim ws As Worksheet
    Dim ok As Boolean
    Dim pLoc As String, pAnio As Long, pMes As Long, pTipo As String, pNumPer As Long

    Dim estabaProtegido As Boolean
    estabaProtegido = ThisWorkbook.ProtectStructure

    If estabaProtegido Then
        On Error Resume Next
        ThisWorkbook.Unprotect pwdLibro
        On Error GoTo 0
    End If

    Application.DisplayAlerts = False

    Dim toDelete As Collection
    Set toDelete = New Collection

    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 2) = "M_" Then
            ok = TryParseNombreMatriz(ws.Name, pLoc, pAnio, pMes, pTipo, pNumPer)
            If ok Then
                If UCase$(pLoc) = UCase$(loc) Then
                    If Not keepDict.Exists(ws.Name) Then
                        toDelete.Add ws.Name
                    End If
                End If
            End If
        End If
    Next ws

    Dim nm As Variant
    For Each nm In toDelete
        On Error Resume Next
        ThisWorkbook.Worksheets(CStr(nm)).Delete
        On Error GoTo 0
    Next nm

    Application.DisplayAlerts = True

    If estabaProtegido Then
        On Error Resume Next
        ThisWorkbook.Protect Password:=pwdLibro, Structure:=True
        On Error GoTo 0
    End If

End Sub

'========================
' Helpers (KEEP)
'========================
Private Sub AddKeep(ByVal dict As Object, ByVal sheetName As String)
    If Len(sheetName) = 0 Then Exit Sub
    If Not dict.Exists(sheetName) Then dict.Add sheetName, True
End Sub

Private Function BuildNombreMatriz(ByVal loc As String, ByVal anio As Long, ByVal mes As Long, ByVal tipo As String, ByVal numPeriodo As Long) As String
    Dim suf As String
    If UCase$(tipo) = "SEMANAL" Then
        suf = "S" & CStr(numPeriodo)
    Else
        suf = "Q" & CStr(numPeriodo)
    End If

    BuildNombreMatriz = "M_" & loc & "_" & CStr(anio) & "_" & Format$(mes, "00") & "_" & suf

    'Si tu NombreHojaMatriz truncó a 31, aquí igualamos para coincidir
    If Len(BuildNombreMatriz) > 31 Then BuildNombreMatriz = Left$(BuildNombreMatriz, 31)
End Function

Private Function NombreMatrizAnterior(ByVal loc As String, ByVal anio As Long, ByVal mes As Long, ByVal tipo As String, ByVal numPeriodo As Long) As String

    Dim y2 As Long, m2 As Long, p2 As Long
    y2 = anio: m2 = mes: p2 = numPeriodo - 1

    If UCase$(tipo) = "QUINCENAL" Then
        If p2 < 1 Then
            'ir al mes anterior y quedarse con Q2
            BackOneMonth y2, m2
            p2 = 2
        End If

    ElseIf UCase$(tipo) = "SEMANAL" Then
        If p2 < 1 Then
            'ir al mes anterior y quedarse con S4
            BackOneMonth y2, m2
            p2 = 4
        End If
    End If

    NombreMatrizAnterior = BuildNombreMatriz(loc, y2, m2, tipo, p2)
End Function

Private Sub BackOneMonth(ByRef anio As Long, ByRef mes As Long)
    mes = mes - 1
    If mes < 1 Then
        mes = 12
        anio = anio - 1
    End If
End Sub

'========================
' Periodo actual por fecha
'========================
Private Function PeriodoActualPorFecha(ByVal tipo As String, ByVal dt As Date) As Long
    Dim d As Long, uDia As Long
    d = Day(dt)
    uDia = Day(DateSerial(Year(dt), Month(dt) + 1, 0))

    Select Case UCase$(tipo)
        Case "QUINCENAL"
            If d <= 15 Then
                PeriodoActualPorFecha = 1
            Else
                PeriodoActualPorFecha = 2
            End If

        Case "SEMANAL"
            If d <= 7 Then
                PeriodoActualPorFecha = 1
            ElseIf d <= 14 Then
                PeriodoActualPorFecha = 2
            ElseIf d <= 21 Then
                PeriodoActualPorFecha = 3
            Else
                PeriodoActualPorFecha = 4
            End If

        Case Else
            PeriodoActualPorFecha = 0
    End Select
End Function

'========================
' Parseo robusto del nombre
' Soporta loc con underscores.
' Formato esperado:
'   M_<LOC>_<YYYY>_<MM>_<Q|S><N>
' Ej: M_CUU_2025_12_Q1
'========================
Public Function TryParseNombreMatriz( _
        ByVal sheetName As String, _
        ByRef loc As String, _
        ByRef anio As Long, _
        ByRef mes As Long, _
        ByRef tipo As String, _
        ByRef numPeriodo As Long) As Boolean

    On Error GoTo Fail

    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")

    re.Pattern = "^M_(.+)_(\d{4})_(\d{2})_([QS])(\d+)$"
    re.IgnoreCase = True
    re.Global = False

    If Not re.Test(sheetName) Then GoTo Fail

    Set m = re.Execute(sheetName)(0)

    loc = CStr(m.SubMatches(0))
    anio = CLng(m.SubMatches(1))
    mes = CLng(m.SubMatches(2))

    Dim t As String
    t = UCase$(CStr(m.SubMatches(3)))
    If t = "Q" Then
        tipo = "QUINCENAL"
    Else
        tipo = "SEMANAL"
    End If

    numPeriodo = CLng(m.SubMatches(4))

    TryParseNombreMatriz = True
    Exit Function

Fail:
    TryParseNombreMatriz = False
End Function




