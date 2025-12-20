VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private Sub Workbook_Open()

    Application.EnableEvents = True

    '1) Leer config
    Dim isTemplate As String
    isTemplate = GetConfig("IsTemplate", "0")

    '2) Setear globals SIEMPRE
    gIsTemplate = (Trim$(isTemplate) = "1")

    If gIsTemplate Then
        gLoc = ""
        gLocDisplay = "TEMPLATE"
    Else
        gLoc = Trim$(GetConfig("LocationCode", ""))
        gLocDisplay = Trim$(GetConfig("LocationDisplay", ""))

        If gLoc = "" Or gLocDisplay = "" Then
            MsgBox "Este archivo no tiene locación configurada (Config).", vbCritical
            Exit Sub
        End If
    End If

    '3) Mostrar Menu
    On Error Resume Next
    Me.Worksheets("Menu").Visible = xlSheetVisible
    Me.Worksheets("Menu").Activate
    On Error GoTo 0

    '4) Protecciones (UIOnly se pierde al abrir, por eso aquí)
    InicializarProtecciones

    '5) Mostrar menú
    frmMenuPrincipal.Show

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim ws As Worksheet

    On Error Resume Next
    modMantenimientoMatrices.LimpiarMatricesRelevantes
    On Error GoTo 0

    Application.ScreenUpdating = False

    For Each ws In Me.Worksheets
        If ws.Name = "Menu" Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetHidden
        End If
    Next ws

    On Error Resume Next
    Me.Worksheets("Menu").Activate
    On Error GoTo 0

    Application.ScreenUpdating = True
End Sub

