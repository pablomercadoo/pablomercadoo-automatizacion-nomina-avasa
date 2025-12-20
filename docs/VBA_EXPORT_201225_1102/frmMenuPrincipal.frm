VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenuPrincipal 
   Caption         =   "UserForm1"
   ClientHeight    =   5530
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8270.001
   OleObjectBlob   =   "frmMenuPrincipal.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private minAnio As Long          ' Primer año con registros
Private mesesNombres As Variant  ' Array con nombres de meses



'====================================================
'  INICIALIZAR MENÚ PRINCIPAL
'====================================================

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long
    Dim v As Variant
    Dim anioActual As Long

    '---------------------------------
    ' 0) Locación fija (ya viene de Workbook_Open)
    '---------------------------------
    Me.lblLocacion.caption = gLocDisplay

    ' Si NO es template y no hay locación, no puede seguir
    If Not gIsTemplate Then
        If Len(Trim$(gLoc)) = 0 Or Len(Trim$(gLocDisplay)) = 0 Then
            MsgBox "Este archivo no tiene locación configurada (Config).", vbCritical
            Unload Me
            Exit Sub
        End If
    End If

    '---------------------------------
    ' 1) Determinar primer año en BDIncidencias_Local (col I = Año)
    '---------------------------------
    anioActual = Year(Date)
    minAnio = anioActual

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("BDIncidencias_Local")
    On Error GoTo 0

    If Not ws Is Nothing Then
        ultFila = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        If ultFila >= 2 Then
            For i = 2 To ultFila
                v = ws.Cells(i, "I").Value
                If IsNumeric(v) Then
                    If CLng(v) < minAnio Then minAnio = CLng(v)
                End If
            Next i
        End If
    End If

    '---------------------------------
    ' 2) Llenar Años
    '---------------------------------
    cboAnio.Clear
    For i = minAnio To anioActual
        cboAnio.AddItem i
    Next i

    '---------------------------------
    ' 3) Meses
    '---------------------------------
    mesesNombres = Array("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", _
                         "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")

    '---------------------------------
    ' 4) Tipo de periodo
    '---------------------------------
    cboTipoPeriodo.Clear
    cboTipoPeriodo.AddItem "Semanal"
    cboTipoPeriodo.AddItem "Quincenal"

    cboMes.Clear
    cboPeriodo.Clear

End Sub


'====================================================
'  CAMBIO DE AÑO -> ACTUALIZAR MESES
'====================================================
Private Sub cboAnio_Change()

    Dim anioSel As Long
    Dim anioActual As Long
    Dim mesLimite As Long
    Dim i As Long

    If cboAnio.Value = "" Then Exit Sub
    If Not IsNumeric(cboAnio.Value) Then Exit Sub

    anioSel = CLng(cboAnio.Value)
    anioActual = Year(Date)

    If anioSel < anioActual Then
        mesLimite = 12
    ElseIf anioSel = anioActual Then
        mesLimite = Month(Date)
    Else
        Exit Sub
    End If

    cboMes.Clear
    For i = 0 To mesLimite - 1
        cboMes.AddItem UCase$(Format$(DateSerial(anioSel, i + 1, 1), "mmmm"))
    Next i

    cboTipoPeriodo.Value = ""
    cboPeriodo.Clear

End Sub

'====================================================
'  CAMBIO DE MES / TIPO PERIODO -> ACTUALIZAR PERIODOS
'====================================================
Private Sub cboMes_Change()
    ActualizarPeriodos
End Sub

Private Sub cboTipoPeriodo_Change()
    ActualizarPeriodos
End Sub

Private Sub ActualizarPeriodos()
    Dim anioSel As Long, mesSel As Long
    Dim anioAct As Long, mesAct As Long, diaHoy As Long
    Dim maxPeriodo As Long
    Dim i As Long

    cboPeriodo.Clear

    If cboAnio.Value = "" Or cboMes.Value = "" Or cboTipoPeriodo.Value = "" Then Exit Sub

    anioSel = CLng(cboAnio.Value)
    mesSel = cboMes.ListIndex + 1

    anioAct = Year(Date)
    mesAct = Month(Date)
    diaHoy = Day(Date)

    Select Case UCase$(cboTipoPeriodo.Value)

        Case "QUINCENAL"
            If (anioSel < anioAct) Or (anioSel = anioAct And mesSel < mesAct) Then
                maxPeriodo = 2
            ElseIf anioSel = anioAct And mesSel = mesAct Then
                If diaHoy <= 15 Then
                    maxPeriodo = 1
                Else
                    maxPeriodo = 2
                End If
            Else
                Exit Sub
            End If

        Case "SEMANAL"
            If (anioSel < anioAct) Or (anioSel = anioAct And mesSel < mesAct) Then
                maxPeriodo = 4
            ElseIf anioSel = anioAct And mesSel = mesAct Then
                Select Case diaHoy
                    Case 1 To 7:   maxPeriodo = 1
                    Case 8 To 14:  maxPeriodo = 2
                    Case 15 To 21: maxPeriodo = 3
                    Case Else:     maxPeriodo = 4
                End Select
            Else
                Exit Sub
            End If

    End Select

    For i = 1 To maxPeriodo
        cboPeriodo.AddItem i
    Next i
End Sub

'====================================================
'  ACEPTAR -> FIJAR GLOBALES, GENERAR MATRIZ Y MOSTRARLA
'====================================================
Private Sub cmdAceptar_Click()

    If gLoc = "" Or _
       cboAnio.Value = "" Or cboMes.Value = "" Or _
       cboTipoPeriodo.Value = "" Or cboPeriodo.Value = "" Then
        MsgBox "Selecciona año, mes, tipo y periodo.", vbExclamation
        Exit Sub
    End If

    gAnio = CLng(cboAnio.Value)
    gMes = cboMes.ListIndex + 1
    gTipoPeriodo = cboTipoPeriodo.Value
    gPeriodo = CLng(cboPeriodo.Value)

    '=========================================
    ' NEW: registrar periodo en tblPeriodos (CAPTURA por default)
    '=========================================
    modPeriodos.EnsurePeriodoRow gLoc, gAnio, gMes, gTipoPeriodo, gPeriodo

    '=============================
    ' 1) Sync empleados (solo 1 vez por periodo)
    '=============================
    Dim pid As String
    pid = modEmpleadosSync.BuildPeriodID(gAnio, gMes, gTipoPeriodo, gPeriodo)

    'False = NO fuerza; si ya se sincronizó ese periodo, no vuelve a tocarlo
    modEmpleadosSync.SyncEmpleados_PeriodoActual pid, False

    '=============================
    ' 2) Generar matriz del periodo
    '=============================
    modReporteIncidencias.GenerarMatrizPeriodoActual
    Unload Me

End Sub

'====================================================
'  CERRAR MENÚ
'====================================================
Private Sub cmdCerrar_Click()
    Unload Me
End Sub


