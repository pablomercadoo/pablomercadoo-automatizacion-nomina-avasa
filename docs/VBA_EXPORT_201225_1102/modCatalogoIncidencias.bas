Attribute VB_Name = "modCatalogoIncidencias"
Option Explicit

Private Const WS_CFG As String = "Config"
Private Const LO_CAT As String = "tblCatalogoIncidencias"

'=========================================
' Normalización
'=========================================
' Normaliza para comparar: mayúsculas, sin espacios, sin "/"
Public Function NormalizaCodigo(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace$(s, " ", "")
    s = Replace$(s, "/", "")
    NormalizaCodigo = s
End Function

' Canoniza alias históricos y limpia basura
' - "T/D" -> "TD"
' - "FI"  -> "F"
' - "0"   -> "" (lo eliminamos)
Public Function CanonizarCodigo(ByVal codigo As String) As String
    Dim norm As String
    norm = NormalizaCodigo(codigo)

    Select Case norm
        Case "", "0"
            CanonizarCodigo = ""
        Case "TD", "T/D"
            CanonizarCodigo = "TD"
        Case "FI"
            CanonizarCodigo = "F"
        Case Else
            CanonizarCodigo = norm
    End Select
End Function

'=========================================
' Catálogo
'=========================================
' Recalcula la columna [Normalizado] = NormalizaCodigo([Codigo])
Public Sub Catalogo_RecalcularNormalizados()
    Dim lo As ListObject, lr As ListRow
    Dim idxCod As Long, idxNorm As Long

    Set lo = ThisWorkbook.Worksheets(WS_CFG).ListObjects(LO_CAT)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    idxCod = lo.ListColumns("Codigo").Index
    idxNorm = lo.ListColumns("Normalizado").Index

    For Each lr In lo.ListRows
        lr.Range(1, idxNorm).Value = NormalizaCodigo(CStr(lr.Range(1, idxCod).Value))
    Next lr
End Sub

' Devuelve array 1D con códigos ACTIVOS (para dropdown)
Public Function GetCodigosActivos() As Variant
    Dim lo As ListObject
    Dim arr, out(), i As Long, n As Long
    Dim idxCod As Long, idxActivo As Long

    Set lo = ThisWorkbook.Worksheets(WS_CFG).ListObjects(LO_CAT)
    If lo.DataBodyRange Is Nothing Then Exit Function

    arr = lo.DataBodyRange.Value
    idxCod = lo.ListColumns("Codigo").Index
    idxActivo = lo.ListColumns("Activo").Index

    ReDim out(1 To UBound(arr, 1)) 'máximo
    n = 0

    For i = 1 To UBound(arr, 1)
        If VBA.CBool(arr(i, idxActivo)) = True Then
            n = n + 1
            out(n) = CStr(arr(i, idxCod))
        End If
    Next i

    If n = 0 Then Exit Function
    ReDim Preserve out(1 To n)

    GetCodigosActivos = out
End Function

' Valida si un código (o alias) existe en catálogo y está ACTIVO
' OJO: permite "" (vacío) como válido
Public Function EsCodigoValido(ByVal codigo As String) As Boolean
    Dim lo As ListObject
    Dim rngNorm As Range
    Dim norm As String, r As Variant
    Dim idxActivo As Long

    norm = CanonizarCodigo(codigo)
    If Len(norm) = 0 Then
        EsCodigoValido = True
        Exit Function
    End If

    Set lo = ThisWorkbook.Worksheets(WS_CFG).ListObjects(LO_CAT)
    Set rngNorm = lo.ListColumns("Normalizado").DataBodyRange
    idxActivo = lo.ListColumns("Activo").Index

    r = Application.Match(norm, rngNorm, 0)
    If IsError(r) Then
        EsCodigoValido = False
    Else
        EsCodigoValido = VBA.CBool(lo.DataBodyRange.Cells(CLng(r), idxActivo).Value)
    End If
End Function

'=========================================
' BD (opcional): canoniza toda la columna O
'=========================================
Public Sub BD_CanonizarCodigos()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim v As String, canon As String

    Set ws = ThisWorkbook.Worksheets("BDIncidencias_Local")
    lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row ' O = CodigoInc

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error Resume Next
    ws.Unprotect "AVASA"
    ws.Unprotect "IncidenciasAVASA"
    On Error GoTo 0

    For i = 2 To lastRow
        v = CStr(ws.Cells(i, "O").Value)
        If Len(Trim$(v)) > 0 Then
            canon = CanonizarCodigo(v)
            If canon <> v Then ws.Cells(i, "O").Value = canon
        End If
    Next i

    On Error Resume Next
    ws.Protect Password:="AVASA"
    On Error GoTo 0

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Listo. Canonizados códigos en BD: columna O.", vbInformation
End Sub

