Attribute VB_Name = "modPeriodos"
Option Explicit

Private Const WS_CFG As String = "Config"
Private Const LO_PERIODOS As String = "tblPeriodos"

'========================
' Helpers
'========================
Private Function GetLO() As ListObject
    Set GetLO = ThisWorkbook.Worksheets(WS_CFG).ListObjects(LO_PERIODOS)
End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then NzStr = "" Else NzStr = Trim$(CStr(v))
End Function

Private Function NzLng(ByVal v As Variant, Optional ByVal def As Long = 0) As Long
    If IsError(v) Or IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then NzLng = def Else NzLng = CLng(v)
End Function

Private Function FindRow(ByVal loc As String, ByVal anio As Long, ByVal mes As Long, _
                         ByVal tipo As String, ByVal periodo As Long) As ListRow
    Dim lo As ListObject, lr As ListRow
    Set lo = GetLO()

    loc = UCase$(Trim$(loc))
    tipo = UCase$(Trim$(tipo))

    For Each lr In lo.ListRows
        If UCase$(NzStr(lr.Range(1, lo.ListColumns("LocCode").Index).Value)) = loc _
           And NzLng(lr.Range(1, lo.ListColumns("Anio").Index).Value) = anio _
           And NzLng(lr.Range(1, lo.ListColumns("Mes").Index).Value) = mes _
           And UCase$(NzStr(lr.Range(1, lo.ListColumns("TipoPeriodo").Index).Value)) = tipo _
           And NzLng(lr.Range(1, lo.ListColumns("Periodo").Index).Value) = periodo Then
            Set FindRow = lr
            Exit Function
        End If
    Next lr
End Function

'========================
' Public API
'========================
Public Sub EnsurePeriodoRow(ByVal loc As String, ByVal anio As Long, ByVal mes As Long, _
                            ByVal tipo As String, ByVal periodo As Long)
    On Error GoTo Fail

    Dim lo As ListObject, lr As ListRow
    Set lo = GetLO()

    Set lr = FindRow(loc, anio, mes, tipo, periodo)
    If Not lr Is Nothing Then Exit Sub

    Set lr = lo.ListRows.Add
    lr.Range(1, lo.ListColumns("LocCode").Index).Value = UCase$(Trim$(loc))
    lr.Range(1, lo.ListColumns("Anio").Index).Value = anio
    lr.Range(1, lo.ListColumns("Mes").Index).Value = mes
    lr.Range(1, lo.ListColumns("TipoPeriodo").Index).Value = UCase$(Trim$(tipo))
    lr.Range(1, lo.ListColumns("Periodo").Index).Value = periodo
    lr.Range(1, lo.ListColumns("Status").Index).Value = "CAPTURA"

    On Error Resume Next
    lr.Range(1, lo.ListColumns("UpdatedAt").Index).Value = Now
    On Error GoTo 0

Fail:
    'Si falla, no rompemos nada: el sistema seguirá con la lógica actual por fecha
End Sub

Public Function GetPeriodoStatus(ByVal loc As String, ByVal anio As Long, ByVal mes As Long, _
                                ByVal tipo As String, ByVal periodo As Long) As String
    On Error GoTo Fail

    Dim lo As ListObject, lr As ListRow
    Set lo = GetLO()
    Set lr = FindRow(loc, anio, mes, tipo, periodo)

    If lr Is Nothing Then
        GetPeriodoStatus = ""
    Else
        GetPeriodoStatus = UCase$(NzStr(lr.Range(1, lo.ListColumns("Status").Index).Value))
    End If
    Exit Function

Fail:
    GetPeriodoStatus = ""
End Function

Public Function GetLockOverrideHours(ByVal loc As String, ByVal anio As Long, ByVal mes As Long, _
                                    ByVal tipo As String, ByVal periodo As Long) As Variant
    'Devuelve Empty si no hay override / no existe fila / error
    On Error GoTo Fail

    Dim lo As ListObject, lr As ListRow
    Set lo = GetLO()
    Set lr = FindRow(loc, anio, mes, tipo, periodo)

    If lr Is Nothing Then
        GetLockOverrideHours = Empty
    Else
        Dim v As Variant
        v = lr.Range(1, lo.ListColumns("LockWindowHoursOverride").Index).Value
        If Len(Trim$(CStr(v))) = 0 Then
            GetLockOverrideHours = Empty
        Else
            GetLockOverrideHours = CDbl(v)
        End If
    End If
    Exit Function

Fail:
    GetLockOverrideHours = Empty
End Function

Public Sub SetPeriodoStatus(ByVal loc As String, ByVal anio As Long, ByVal mes As Long, _
                            ByVal tipo As String, ByVal periodo As Long, ByVal newStatus As String)
    On Error GoTo Fail

    Dim lo As ListObject, lr As ListRow
    newStatus = UCase$(Trim$(newStatus))
    If newStatus <> "CAPTURA" And newStatus <> "ENVIADO" And newStatus <> "CERRADO" Then Exit Sub

    Set lo = GetLO()
    Set lr = FindRow(loc, anio, mes, tipo, periodo)

    If lr Is Nothing Then
        EnsurePeriodoRow loc, anio, mes, tipo, periodo
        Set lr = FindRow(loc, anio, mes, tipo, periodo)
        If lr Is Nothing Then Exit Sub
    End If

    lr.Range(1, lo.ListColumns("Status").Index).Value = newStatus
    On Error Resume Next
    lr.Range(1, lo.ListColumns("UpdatedAt").Index).Value = Now
    On Error GoTo 0

Fail:
End Sub


