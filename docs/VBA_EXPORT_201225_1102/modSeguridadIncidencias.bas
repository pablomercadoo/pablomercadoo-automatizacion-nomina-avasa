Attribute VB_Name = "modSeguridadIncidencias"


Option Explicit

'==========================
'  SEGURIDAD / PROTECCIÓN
'==========================
Public Const PASS As String = "AVASA"
Public Const SECURITY_ON As Boolean = False   'DEV: False / RELEASE: True

'Deja la hoja "editable para macros" (UserInterfaceOnly)
'pero protegida contra edición manual (si SECURITY_ON=True)
Public Sub EnsureSheetEditable(ByVal ws As Worksheet)
    If Not SECURITY_ON Then Exit Sub

    On Error Resume Next
    ws.Unprotect Password:=PASS
    ws.Protect Password:=PASS, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub

'Protege una hoja de matriz (bloquea todo)
Public Sub ProtegerHojaMatriz(ByVal ws As Worksheet)
    If Not SECURITY_ON Then Exit Sub

    On Error Resume Next
    ws.Unprotect Password:=PASS
    ws.Cells.Locked = True
    ws.Cells.FormulaHidden = False
    ws.Protect Password:=PASS, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub

'Protege todas las hojas M_*
Public Sub InicializarProtecciones()
    If Not SECURITY_ON Then Exit Sub

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 2) = "M_" Then ProtegerHojaMatriz ws
    Next ws
End Sub

Public Sub UnprotectSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String)
    If Not SECURITY_ON Then Exit Sub
    On Error Resume Next
    wb.Worksheets(sheetName).Unprotect Password:=PASS
    On Error GoTo 0
End Sub

Public Sub ProtectSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String)
    If Not SECURITY_ON Then Exit Sub
    On Error Resume Next
    wb.Worksheets(sheetName).Protect Password:=PASS, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

'==========================
'  CIERRE AUTOMÁTICO
'==========================
'Se considera "cerrado" cuando:
'   Now >= FechaFinPeriodo + (LockWindowHours / 24)
'LockWindowHours viene de Config (ej 48)
'
'NEW:
' - Si tblPeriodos.Status = CERRADO -> bloquea siempre.
' - Si tblPeriodos.LockWindowHoursOverride tiene valor -> usa ese valor en vez del default.

Public Function GetLockWindowHours_Default() As Double
    'Default: 48 horas
    Dim v As Variant
    v = GetConfig("LockWindowHours", 48)

    If IsNumeric(v) Then
        GetLockWindowHours_Default = CDbl(v)
    Else
        GetLockWindowHours_Default = 48#
    End If
End Function

Public Function GetLockWindowHours_Efectivo() As Double
    'Devuelve override si existe; si no, default de Config
    On Error GoTo Fail

    Dim vOvr As Variant
    vOvr = modPeriodos.GetLockOverrideHours(gLoc, gAnio, gMes, gTipoPeriodo, gPeriodo)

    If IsEmpty(vOvr) Then
        GetLockWindowHours_Efectivo = GetLockWindowHours_Default()
    ElseIf IsNumeric(vOvr) Then
        GetLockWindowHours_Efectivo = CDbl(vOvr)
    Else
        GetLockWindowHours_Efectivo = GetLockWindowHours_Default()
    End If
    Exit Function

Fail:
    GetLockWindowHours_Efectivo = GetLockWindowHours_Default()
End Function

Public Function GetFechaFinPeriodoActual() As Date
    'Calcula fecha fin del periodo actual usando globals
    'Requiere: gAnio, gMes, gTipoPeriodo, gPeriodo (modGlobal)
    Dim diaIni As Long, diaFin As Long
    diaIni = 0: diaFin = 0

    Call ObtenerRangoPeriodo_Local(gAnio, gMes, gTipoPeriodo, gPeriodo, diaIni, diaFin)
    GetFechaFinPeriodoActual = DateSerial(gAnio, gMes, diaFin)
End Function

Public Function GetFechaCierrePeriodoActual() As Date
    GetFechaCierrePeriodoActual = GetFechaFinPeriodoActual() + (GetLockWindowHours_Efectivo() / 24#)
End Function

Public Function PeriodoActualEstaCerrado(Optional ByRef cierre As Date) As Boolean
    cierre = GetFechaCierrePeriodoActual()
    PeriodoActualEstaCerrado = (Now >= cierre)
End Function

Public Function PermiteEdicionPeriodo(Optional ByRef msgBloqueo As String) As Boolean
    On Error GoTo Fail

    '-----------------------------------------
    ' 1) OVERRIDE MANUAL: Status=CERRADO
    '-----------------------------------------
    Dim st As String
    st = modPeriodos.GetPeriodoStatus(gLoc, gAnio, gMes, gTipoPeriodo, gPeriodo)

    If st = "CERRADO" Then
        msgBloqueo = "Periodo CERRADO (manual)." & vbCrLf & _
                     "Si necesitas reabrir, solicítalo a Nómina."
        PermiteEdicionPeriodo = False
        Exit Function
    End If

'-----------------------------------------
' 2) CIERRE AUTOMÁTICO: fin + horas (override o default)
'-----------------------------------------
Dim cierre As Date
If PeriodoActualEstaCerrado(cierre) Then

    'NEW: dejar huella del cierre automático
    On Error Resume Next
    modPeriodos.SetPeriodoStatus gLoc, gAnio, gMes, gTipoPeriodo, gPeriodo, "CERRADO"
    On Error GoTo 0

    msgBloqueo = "Periodo CERRADO." & vbCrLf & _
                 "Cierre automático: " & Format$(cierre, "dd/mm/yyyy hh:mm") & vbCrLf & _
                 "Si necesitas reabrir, solicítalo a Nómina."
    PermiteEdicionPeriodo = False
Else
    PermiteEdicionPeriodo = True
End If
Exit Function


Fail:
    'Si algo falla, no bloqueamos por error: usamos la lógica actual (por fecha + default)
    Dim cierre2 As Date
    cierre2 = GetFechaFinPeriodoActual() + (GetLockWindowHours_Default() / 24#)
    If Now >= cierre2 Then
        msgBloqueo = "Periodo CERRADO." & vbCrLf & _
                     "Cierre automático: " & Format$(cierre2, "dd/mm/yyyy hh:mm") & vbCrLf & _
                     "Si necesitas reabrir, solicítalo a Nómina."
        PermiteEdicionPeriodo = False
    Else
        PermiteEdicionPeriodo = True
    End If
End Function

Public Sub ValidarPeriodoAbiertoOrExit()
    Dim m As String
    If Not PermiteEdicionPeriodo(m) Then
        MsgBox m, vbExclamation, "Bloqueo por cierre"
        Err.Raise vbObjectError + 513, "modSeguridadIncidencias", "Periodo cerrado"
    End If
End Sub

'Aplica el cierre a la hoja matriz:
'- Si está abierto: quita protección (por si quedó protegida)
'- Si está cerrado: protege (solo consulta)
Public Sub AplicarCierreEnMatriz(ByVal wsMat As Worksheet)
    Dim m As String

    If wsMat Is Nothing Then Exit Sub

    '--- PERIODO ABIERTO: quitar protección ---
    If PermiteEdicionPeriodo(m) Then
        On Error Resume Next
        wsMat.Unprotect PASS
        wsMat.EnableSelection = xlNoRestrictions
        On Error GoTo 0
        Exit Sub
    End If

    '--- PERIODO CERRADO: proteger hoja ---
    On Error Resume Next
    wsMat.Unprotect PASS
    On Error GoTo 0

    On Error Resume Next
    wsMat.Cells.Locked = True
    wsMat.Protect Password:=PASS, _
                  UserInterfaceOnly:=True, _
                  AllowFiltering:=True, _
                  AllowSorting:=True
    wsMat.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub

'-----------------------------------------
' Helper local: rango periodo
'-----------------------------------------
Private Sub ObtenerRangoPeriodo_Local( _
        ByVal anio As Long, _
        ByVal mes As Long, _
        ByVal tipoPeriodo As String, _
        ByVal numPeriodo As Long, _
        ByRef diaIni As Long, _
        ByRef diaFin As Long)

    Dim uDia As Long
    uDia = Day(DateSerial(anio, mes + 1, 0))

    Select Case UCase$(tipoPeriodo)
        Case "SEMANAL"
            Select Case numPeriodo
                Case 1: diaIni = 1:  diaFin = 7
                Case 2: diaIni = 8:  diaFin = 14
                Case 3: diaIni = 15: diaFin = 21
                Case 4: diaIni = 22: diaFin = uDia
            End Select

        Case "QUINCENAL"
            Select Case numPeriodo
                Case 1: diaIni = 1:  diaFin = 15
                Case 2: diaIni = 16: diaFin = uDia
            End Select
    End Select
End Sub


