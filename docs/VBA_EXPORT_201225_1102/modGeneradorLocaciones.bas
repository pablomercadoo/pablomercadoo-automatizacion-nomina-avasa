Attribute VB_Name = "modGeneradorLocaciones"
Option Explicit

'==========================
' GENERADOR 62 ARCHIVOS
'==========================
Public Sub GenerarArchivosPorLocacion()

    Dim lo As ListObject, r As ListRow
    Dim idxActive As Long, idxCode As Long, idxName As Long, idxCC As Long

    Dim basePath As String
    Dim locCode As String, locName As String, cc As String

    Dim fileName As String, fullPath As String
    Dim tempCopyPath As String
    Dim wbNew As Workbook

    On Error GoTo SafeExit

    ' 1) BasePath (SALIDA) -> Config!MasterDBPath
    basePath = GetConfig("MasterDBPath", "")
    If Len(Trim$(basePath)) = 0 Then
        MsgBox "Config.MasterDBPath está vacío. Ej: C:\AVASA_TMP\OUT\", vbCritical
        Exit Sub
    End If
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    EnsureFolderExists basePath

    ' 2) Tabla locaciones
    Set lo = Nothing
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets("Locaciones").ListObjects("tblLocaciones")
    On Error GoTo 0

    If lo Is Nothing Then
        MsgBox "No encontré la tabla tblLocaciones en la hoja 'Locaciones'.", vbCritical
        Exit Sub
    End If

    If Not TableHasColumn(lo, "Active") _
       Or Not TableHasColumn(lo, "LocationCode") _
       Or Not TableHasColumn(lo, "LocationName") _
       Or Not TableHasColumn(lo, "CC") Then
        MsgBox "tblLocaciones debe tener: LocationCode, LocationName, CC, Active", vbCritical
        Exit Sub
    End If

    idxActive = lo.ListColumns("Active").Index
    idxCode = lo.ListColumns("LocationCode").Index
    idxName = lo.ListColumns("LocationName").Index
    idxCC = lo.ListColumns("CC").Index

    ' 3) Requisito: template guardado en disco
    If Len(Trim$(ThisWorkbook.Path)) = 0 Then
        MsgBox "Guarda primero este template (.xlsm) en una carpeta antes de generar copias.", vbCritical
        Exit Sub
    End If

    ' 4) Performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ' 5) Loop locaciones activas
    For Each r In lo.ListRows

        If Val(r.Range.Cells(1, idxActive).Value) = 1 Then

            locCode = Trim$(CStr(r.Range.Cells(1, idxCode).Value))
            locName = Trim$(CStr(r.Range.Cells(1, idxName).Value))
            cc = Trim$(CStr(r.Range.Cells(1, idxCC).Value))

            If locCode <> "" Then

                fileName = "Incidencias_" & locCode & ".xlsm"
                fullPath = basePath & fileName

                ' Temporal SIEMPRE local
                tempCopyPath = Environ$("TEMP") & "\__TMP__" & fileName
                If FileExists(tempCopyPath) Then Kill tempCopyPath

                ' Copiar template -> temp
                ThisWorkbook.SaveCopyAs tempCopyPath

                ' Abrir temp
                Set wbNew = Workbooks.Open(fileName:=tempCopyPath, ReadOnly:=False)

                ' *** CLAVE: desproteger Config ANTES de escribir ***
              If modSeguridadIncidencias.SECURITY_ON Then
    UnprotectSheetIfExists wbNew, "Config", "AVASA"
End If

                ' Setear Config
                SetConfig_InWorkbook wbNew, "LocationCode", locCode
                SetConfig_InWorkbook wbNew, "LocationName", locName
                SetConfig_InWorkbook wbNew, "LocationDisplay", locCode & " - " & locName
                SetConfig_InWorkbook wbNew, "CC", cc
                SetConfig_InWorkbook wbNew, "TemplateVersion", GetConfig("TemplateVersion", "1.0.0")
                SetConfig_InWorkbook wbNew, "IsTestFile", "0"
                SetConfig_InWorkbook wbNew, "IsTemplate", "0"

                ' Volver a proteger Config
If modSeguridadIncidencias.SECURITY_ON Then
    ProtectSheetIfExists wbNew, "Config", "AVASA"
End If

                ' (Opcional) Ocultar Locaciones
                On Error Resume Next
                wbNew.Worksheets("Locaciones").Visible = xlSheetHidden
                On Error GoTo 0

                ' Reemplazar si ya existe
                If FileExists(fullPath) Then Kill fullPath

                wbNew.SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
                wbNew.Close SaveChanges:=True

                ' Limpieza temp
                If FileExists(tempCopyPath) Then Kill tempCopyPath

            End If
        End If

    Next r

    MsgBox "Listo. Archivos generados en: " & basePath, vbInformation

SafeExit:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

'==========================
' Helpers Config
'==========================
Private Sub SetConfig_InWorkbook(ByVal wb As Workbook, ByVal cfgKey As String, ByVal cfgValue As String)
    Dim lo As ListObject, rr As ListRow

    Set lo = wb.Worksheets("Config").ListObjects("tblConfig")

    For Each rr In lo.ListRows
        If UCase$(Trim$(CStr(rr.Range.Cells(1, 1).Value))) = UCase$(Trim$(cfgKey)) Then
            rr.Range.Cells(1, 2).Value = cfgValue
            Exit Sub
        End If
    Next rr

    Set rr = lo.ListRows.Add
    rr.Range.Cells(1, 1).Value = cfgKey
    rr.Range.Cells(1, 2).Value = cfgValue
End Sub

'==========================
' Helpers protección
'==========================
Private Sub UnprotectSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String, ByVal pwd As String)
    On Error Resume Next
    wb.Worksheets(sheetName).Unprotect Password:=pwd
    On Error GoTo 0
End Sub

Private Sub ProtectSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String, ByVal pwd As String)
    On Error Resume Next
    wb.Worksheets(sheetName).Protect Password:=pwd, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

'==========================
' Helpers varios
'==========================
Private Function TableHasColumn(ByVal lo As ListObject, ByVal colName As String) As Boolean
    On Error GoTo Fail
    Dim tmp As Long
    tmp = lo.ListColumns(colName).Index
    TableHasColumn = True
    Exit Function
Fail:
    TableHasColumn = False
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(fullPath)) > 0)
    On Error GoTo 0
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
End Sub





