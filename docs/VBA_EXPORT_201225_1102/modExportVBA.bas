Attribute VB_Name = "modExportVBA"
Option Explicit

'====================================================
' EXPORTACIÓN DE TODO EL PROYECTO VBA (AVASA)
'
' ? Exporta TODO a 1 TXT con timestamp:
'    VBA_COMPLETO_ddmmyy_hhnn.txt
'
' ? Exporta por componente a carpeta con timestamp:
'    VBA_EXPORT_ddmmyy_hhnn\  (dentro del GetExportPath)
'    - .bas / .cls / .frm (+ .frx si aplica)
' ? Genera INDEX.txt (manifiesto) dentro de la carpeta
' ? (Opcional) Excluye componentes con nombre que empiece con "_"
'
' Nota:
' Para exportar componentes, Excel requiere:
' Trust access to the VBA project object model (en Trust Center)
'====================================================

'----------------------------------------------------
' Ruta base donde se guardan exports
'----------------------------------------------------
Private Function GetExportPath() As String
    GetExportPath = "C:\Users\OPERADOR1\OneDrive - AVASA\Documentos\Comisiones e incidencias AVASA\Proyecto automatización\CodigoBVA\"
End Function

'----------------------------------------------------
' Util: asegura que la ruta termine con "\"
'----------------------------------------------------
Private Function NormalizePath(ByVal p As String) As String
    p = Trim$(p)
    If Len(p) = 0 Then
        NormalizePath = ""
    ElseIf Right$(p, 1) = "\" Then
        NormalizePath = p
    Else
        NormalizePath = p & "\"
    End If
End Function

'----------------------------------------------------
' Util: timestamp ddmmyy_hhnn (24h)
'----------------------------------------------------
Private Function GetStamp() As String
    GetStamp = Format(Now, "ddmmyy_hhnn")
End Function

'----------------------------------------------------
' Util: obtener extensión por tipo de componente
' vbext_ComponentType:
' 1=StdModule, 2=ClassModule, 3=MSForm, 100=Document
'----------------------------------------------------
Private Function ExtByType(ByVal compType As Long) As String
    Select Case compType
        Case 1:   ExtByType = ".bas"
        Case 2:   ExtByType = ".cls"
        Case 3:   ExtByType = ".frm"
        Case 100: ExtByType = ".cls"
        Case Else: ExtByType = ".txt"
    End Select
End Function

'----------------------------------------------------
' Util: crear carpeta si no existe (simple)
'----------------------------------------------------
Private Sub EnsureFolder(ByVal folderPath As String)
    If Len(folderPath) = 0 Then Exit Sub
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
End Sub

'----------------------------------------------------
' Exporta TODO el proyecto VBA a un solo TXT con timestamp
'----------------------------------------------------
Public Sub ExportarVBA_A_UnSoloTXT()

    Dim exportBase As String
    Dim exportFile As String
    Dim stamp As String
    Dim vbComp As Object
    Dim ff As Integer
    Dim i As Long

    exportBase = NormalizePath(GetExportPath())
    If exportBase = "" Then
        MsgBox "GetExportPath() está vacío.", vbCritical
        Exit Sub
    End If

    If Dir(exportBase, vbDirectory) = "" Then
        MsgBox "La carpeta no existe:" & vbCrLf & exportBase, vbCritical
        Exit Sub
    End If

    stamp = GetStamp()
    exportFile = exportBase & "VBA_COMPLETO_" & stamp & ".txt"

    ff = FreeFile
    Open exportFile For Output As #ff

    ' Encabezado
    Print #ff, "===================================================="
    Print #ff, "VBA_COMPLETO — Export automático"
    Print #ff, "Fecha: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #ff, "Workbook: " & ThisWorkbook.Name
    Print #ff, "===================================================="
    Print #ff, vbCrLf

    ' Exportación por componente (pegado de código)
    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        'Opcional: saltar componentes con "_" al inicio
        If Left$(vbComp.Name, 1) = "_" Then GoTo NextCompTXT

        Print #ff, "========================================"
        Print #ff, "COMPONENTE: " & vbComp.Name
        Print #ff, "TIPO: " & vbComp.Type
        Print #ff, "========================================"

        If vbComp.CodeModule.CountOfLines > 0 Then
            For i = 1 To vbComp.CodeModule.CountOfLines
                Print #ff, vbComp.CodeModule.Lines(i, 1)
            Next i
        End If

        Print #ff, vbCrLf

NextCompTXT:
    Next vbComp

    Close #ff

    MsgBox "Archivo creado correctamente:" & vbCrLf & exportFile, vbInformation

End Sub

'----------------------------------------------------
' Exporta cada componente a archivo individual en subcarpeta timestamp
' + genera INDEX.txt
'----------------------------------------------------
Public Sub ExportarVBA_AVASA()

    Dim exportBase As String
    Dim exportPath As String
    Dim stamp As String

    Dim vbComp As Object
    Dim ext As String
    Dim outFile As String

    Dim ff As Integer
    Dim exportOK As Boolean
    Dim errMsg As String

    exportBase = NormalizePath(GetExportPath())
    If exportBase = "" Then
        MsgBox "GetExportPath() está vacío.", vbCritical
        Exit Sub
    End If

    If Dir(exportBase, vbDirectory) = "" Then
        MsgBox "La carpeta no existe:" & vbCrLf & exportBase, vbCritical
        Exit Sub
    End If

    stamp = GetStamp()
    exportPath = exportBase & "VBA_EXPORT_" & stamp & "\"
    EnsureFolder exportPath

    'INDEX.txt
    ff = FreeFile
    Open exportPath & "INDEX.txt" For Output As #ff
    Print #ff, "Export VBA - " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #ff, "Workbook: " & ThisWorkbook.Name
    Print #ff, "Folder: " & exportPath
    Print #ff, String(70, "-")
    Print #ff, "Name | Type | File | Result"
    Print #ff, String(70, "-")

    exportOK = True

    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        'Opcional: saltar componentes con "_" al inicio
        If Left$(vbComp.Name, 1) = "_" Then GoTo NextComp

        ext = ExtByType(vbComp.Type)
        outFile = exportPath & vbComp.Name & ext

        Err.Clear
        On Error Resume Next
        vbComp.Export outFile
        If Err.Number <> 0 Then
            exportOK = False
            errMsg = "ERR " & Err.Number & ": " & Err.Description
            Print #ff, vbComp.Name & " | " & vbComp.Type & " | " & (vbComp.Name & ext) & " | " & errMsg
        Else
            Print #ff, vbComp.Name & " | " & vbComp.Type & " | " & (vbComp.Name & ext) & " | OK"
        End If
        On Error GoTo 0

NextComp:
    Next vbComp

    Close #ff

    If exportOK Then
        MsgBox "Exportación por componentes completada en:" & vbCrLf & exportPath, vbInformation
    Else
        MsgBox "Exportación terminó con algunos errores." & vbCrLf & _
               "Revisa INDEX.txt en:" & vbCrLf & exportPath, vbExclamation
    End If

End Sub

'----------------------------------------------------
' Exportación COMPLETA:
' 1) TXT completo
' 2) Carpeta con archivos individuales + INDEX
'----------------------------------------------------
Public Sub ExportarVBA_COMPLETO()

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo Fail

    Call ExportarVBA_A_UnSoloTXT
    Call ExportarVBA_AVASA

    MsgBox "Exportación COMPLETA finalizada correctamente.", vbInformation

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error durante la exportación COMPLETA:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
End Sub

