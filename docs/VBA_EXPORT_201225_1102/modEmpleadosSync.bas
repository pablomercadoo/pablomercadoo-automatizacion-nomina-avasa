Attribute VB_Name = "modEmpleadosSync"
Option Explicit

'========================================================
' BuildPeriodID: "2025-12-Q1" o "2025-12-S3"
'========================================================
Public Function BuildPeriodID(ByVal anio As Long, ByVal mes As Long, ByVal tipo As String, ByVal periodo As Long) As String
    Dim suf As String
    If UCase$(Trim$(tipo)) = "SEMANAL" Then
        suf = "S" & CStr(periodo)
    Else
        suf = "Q" & CStr(periodo)
    End If
    BuildPeriodID = CStr(anio) & "-" & Format$(mes, "00") & "-" & suf
End Function

'========================================================
' SyncEmpleados_PeriodoActual:
' - Inserta en hoja "Empleados" SOLO una vez por PeriodID y Locación
' - No borra meses pasados (esto es lo que quieres)
'========================================================
Public Sub SyncEmpleados_PeriodoActual(ByVal periodID As String, ByVal force As Boolean)

    Dim dbPath As String, dbTable As String
    dbPath = GetConfig("EmployeeDBPath", "")
    dbTable = GetConfig("EmployeeDBTable", "")

    If Len(Trim$(dbPath)) = 0 Or Len(Trim$(dbTable)) = 0 Then
        MsgBox "Falta configurar EmployeeDBPath y/o EmployeeDBTable en Config.", vbCritical
        Exit Sub
    End If

    'Si no forzamos y ya se sincronizó ese PeriodID, no hacemos nada
    If Not force Then
        If UCase$(GetConfig("EmployeeLastSyncPeriodID", "")) = UCase$(periodID) Then Exit Sub
    End If

    Dim wbDB As Workbook, wsEmp As Worksheet
    Dim loDB As ListObject
    Dim loTarget As ListObject

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'Abrir BD (solo lectura)
    Set wbDB = Workbooks.Open(fileName:=dbPath, ReadOnly:=True)

    'Buscar la tabla por nombre en cualquier hoja
    Set loDB = FindTableInWorkbook(wbDB, dbTable)
    If loDB Is Nothing Then
        wbDB.Close SaveChanges:=False
        GoTo SafeExitWithMsg
    End If

    'Hoja destino en archivo de locación
    Set wsEmp = ThisWorkbook.Worksheets("Empleados")

    'Si ya tienes una tabla en Empleados, úsala; si no, creamos/armamos un rango tabla
    Set loTarget = Nothing
    On Error Resume Next
    If wsEmp.ListObjects.Count > 0 Then Set loTarget = wsEmp.ListObjects(1)
    On Error GoTo 0

    'Generar/actualizar tabla destino con empleados activos SOLO de esta locación (gLoc)
    WriteEmpleadosToLocal loDB, wsEmp, loTarget, periodID, gLoc

    'Guardar marca de sync
    SetConfig "EmployeeLastSyncPeriodID", periodID
    SetConfig "LastEmployeeSync", Format$(Now, "yyyy-mm-dd hh:nn:ss")

    wbDB.Close SaveChanges:=False

SafeExit:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

SafeExitWithMsg:
    MsgBox "No encontré la tabla '" & dbTable & "' dentro de: " & vbCrLf & dbPath, vbCritical
    Resume SafeExit
End Sub


'========================================================
' Helpers
'========================================================
Private Function FindTableInWorkbook(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If UCase$(lo.Name) = UCase$(tableName) Then
                Set FindTableInWorkbook = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set FindTableInWorkbook = Nothing
End Function

Private Sub WriteEmpleadosToLocal(ByVal loDB As ListObject, ByVal wsEmp As Worksheet, ByVal loTarget As ListObject, _
                                 ByVal periodID As String, ByVal locCode As String)

    'Columnas esperadas en la BD:
    'PeriodID, GRUPO, CIUDAD, NumeroEmpleado, UsuarioCARs+, DriverCARs+, PUESTO, ACTIVIDAD, NOMBRE, FechaIngreso, FechaBaja

    Dim idxPID As Long, idxGrupo As Long, idxCiudad As Long, idxNum As Long
    Dim idxUsr As Long, idxDrv As Long, idxPuesto As Long, idxAct As Long, idxNom As Long
    Dim idxIng As Long, idxBaja As Long

    idxPID = loDB.ListColumns("PeriodID").Index
    idxGrupo = loDB.ListColumns("GRUPO").Index
    idxCiudad = loDB.ListColumns("CIUDAD").Index
    idxNum = loDB.ListColumns("NumeroEmpleado").Index
    idxUsr = loDB.ListColumns("UsuarioCARs+").Index
    idxDrv = loDB.ListColumns("DriverCARs+").Index
    idxPuesto = loDB.ListColumns("PUESTO").Index
    idxAct = loDB.ListColumns("ACTIVIDAD").Index
    idxNom = loDB.ListColumns("NOMBRE").Index
    idxIng = loDB.ListColumns("FechaIngreso").Index
    idxBaja = loDB.ListColumns("FechaBaja").Index

Dim outRow As Long

On Error Resume Next
wsEmp.Unprotect Password:="AVASA"
On Error GoTo 0

wsEmp.Cells.Clear



    'Headers destino (lo que ya usa tu frmIncidencias: C=Numero, D=Usuario, E=Driver, F=Puesto, G=Actividad, H=Nombre)
    wsEmp.Range("A1").Value = "GRUPO"
    wsEmp.Range("B1").Value = "CIUDAD"
    wsEmp.Range("C1").Value = "NumeroEmpleado"
    wsEmp.Range("D1").Value = "UsuarioCARs+"
    wsEmp.Range("E1").Value = "DriverCARs+"
    wsEmp.Range("F1").Value = "PUESTO"
    wsEmp.Range("G1").Value = "ACTIVIDAD"
    wsEmp.Range("H1").Value = "NOMBRE"
    wsEmp.Range("I1").Value = "FechaIngreso"
    wsEmp.Range("J1").Value = "FechaBaja"

    outRow = 2

    Dim r As ListRow, fb As Variant, pid As String, grp As String
    For Each r In loDB.ListRows
grp = Trim$(CStr(r.Range.Cells(1, idxGrupo).Value))
fb = r.Range.Cells(1, idxBaja).Value

'Filtro: misma locación + activo (FechaBaja vacía)
If UCase$(grp) = UCase$(locCode) Then
    If Len(Trim$(CStr(fb))) = 0 Then

        wsEmp.Cells(outRow, 1).Value = r.Range.Cells(1, idxGrupo).Value
        wsEmp.Cells(outRow, 2).Value = r.Range.Cells(1, idxCiudad).Value
        wsEmp.Cells(outRow, 3).Value = r.Range.Cells(1, idxNum).Value
        wsEmp.Cells(outRow, 4).Value = r.Range.Cells(1, idxUsr).Value
        wsEmp.Cells(outRow, 5).Value = r.Range.Cells(1, idxDrv).Value
        wsEmp.Cells(outRow, 6).Value = r.Range.Cells(1, idxPuesto).Value
        wsEmp.Cells(outRow, 7).Value = r.Range.Cells(1, idxAct).Value
        wsEmp.Cells(outRow, 8).Value = r.Range.Cells(1, idxNom).Value
        wsEmp.Cells(outRow, 9).Value = r.Range.Cells(1, idxIng).Value
        wsEmp.Cells(outRow, 10).Value = r.Range.Cells(1, idxBaja).Value

        outRow = outRow + 1
    End If
End If
    Next r

    'Convertir a tabla local (para que quede ordenado y fácil)
    Dim lastRow As Long
    lastRow = wsEmp.Cells(wsEmp.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim rng As Range
    Set rng = wsEmp.Range("A1").Resize(lastRow, 10)

    On Error Resume Next
    wsEmp.ListObjects(1).Unlist
    On Error GoTo 0

    Dim loNew As ListObject
    Set loNew = wsEmp.ListObjects.Add(xlSrcRange, rng, , xlYes)
    loNew.Name = "tblEmpleados_Local"

    wsEmp.Columns("A:J").AutoFit
    
    '--- Volver a proteger Empleados ---
If modSeguridadIncidencias.SECURITY_ON Then
    On Error Resume Next
    wsEmp.Protect Password:="AVASA", UserInterfaceOnly:=True
    On Error GoTo 0
End If

End Sub


