Attribute VB_Name = "modConfig"

Option Explicit

'==========================
' CONFIG: Get / Set por Key
'==========================

Public Function GetConfig(ByVal cfgKey As String, Optional ByVal defaultValue As String = "") As String
    Dim lo As ListObject, r As ListRow
    Set lo = ThisWorkbook.Worksheets("Config").ListObjects("tblConfig")

    For Each r In lo.ListRows
        If UCase$(Trim$(CStr(r.Range.Cells(1, 1).Value))) = UCase$(Trim$(cfgKey)) Then
            GetConfig = CStr(r.Range.Cells(1, 2).Value)
            Exit Function
        End If
    Next r

    GetConfig = defaultValue
End Function

Public Sub SetConfig(ByVal cfgKey As String, ByVal cfgValue As String)
    Dim lo As ListObject, r As ListRow
    Set lo = ThisWorkbook.Worksheets("Config").ListObjects("tblConfig")

    For Each r In lo.ListRows
        If UCase$(Trim$(CStr(r.Range.Cells(1, 1).Value))) = UCase$(Trim$(cfgKey)) Then
            r.Range.Cells(1, 2).Value = cfgValue
            Exit Sub
        End If
    Next r

    Set r = lo.ListRows.Add
    r.Range.Cells(1, 1).Value = cfgKey
    r.Range.Cells(1, 2).Value = cfgValue
End Sub

' Para archivos de locación: bloquea Config (úsalo ya hasta "release")
Public Sub LockConfigSheet(Optional ByVal pwd As String = "AVASA")
    With ThisWorkbook.Worksheets("Config")
        .Protect Password:=pwd, UserInterfaceOnly:=True
        .Visible = xlSheetVeryHidden
    End With
End Sub

