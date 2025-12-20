Attribute VB_Name = "modReporteIncidencias"
Option Explicit

' gLoc, gAnio, gMes, gTipoPeriodo, gPeriodo
' vienen de modGlobal


'====================================================
' 1. Nombre de la hoja de matriz del periodo actual
'====================================================
Private Function NombreHojaMatriz() As String
    Dim sufTipo As String
    Dim nombre As String
    
    If UCase$(gTipoPeriodo) = "SEMANAL" Then
        sufTipo = "S" & gPeriodo
    Else
        sufTipo = "Q" & gPeriodo
    End If
    
    'Ejemplo base: M_CUU_2025_12_Q1
    nombre = "M_" & gLoc & "_" & gAnio & "_" & _
             Format$(gMes, "00") & "_" & sufTipo
             
    'Por si algún día gLoc trae algo largo (ej. "CANCUN_AEROPUERTO")
    If Len(nombre) > 31 Then
        nombre = Left$(nombre, 31)
    End If
    
    NombreHojaMatriz = nombre
End Function


'====================================================
' 2. Obtener (o crear) la hoja de matriz del periodo
'====================================================
Private Function GetHojaMatrizPeriodo() As Worksheet
    Dim ws As Worksheet
    Dim nombre As String
    
    nombre = NombreHojaMatriz()
    
    '¿Ya existe la hoja?
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    
    If ws Is Nothing Then
        
        'Si la estructura del libro está protegida, no se pueden agregar hojas
        If ThisWorkbook.ProtectStructure Then
            MsgBox "No se puede crear la hoja de matriz '" & nombre & "' porque " & _
                   "la estructura del libro está protegida." & vbCrLf & vbCrLf & _
                   "Desprotege el libro (Revisar > Proteger libro) y vuelve a intentar.", _
                   vbCritical
            Set GetHojaMatrizPeriodo = Nothing
            Exit Function
        End If
        
        'Crear hoja al final
        On Error GoTo ErrCrear
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombre
        On Error GoTo 0
    End If
    
    '?? IMPORTANTE: asegurarnos de que la hoja esté visible
    ws.Visible = xlSheetVisible
    
    Set GetHojaMatrizPeriodo = ws
    Exit Function

ErrCrear:
    MsgBox "Error al crear la hoja de matriz '" & nombre & "'." & vbCrLf & _
           "Detalle: " & Err.Description, vbCritical
    Set GetHojaMatrizPeriodo = Nothing
End Function


'====================================================
' 3. Rango (primer y último día) del periodo
'====================================================
Private Sub GetRangoPeriodo( _
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


'====================================================
' 4. Buscar CódigoInc en BDIncidencias_Local
'    para un empleado / día / periodo actual
'====================================================
Private Function BuscarCodigoInc( _
        ByVal numEmp As Long, _
        ByVal dia As Long) As String

    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("BDIncidencias_Local")
    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultFila
        If ws.Cells(i, "A").Value = gLoc And _
           ws.Cells(i, "C").Value = numEmp And _
           ws.Cells(i, "I").Value = gAnio And _
           ws.Cells(i, "J").Value = gMes And _
           ws.Cells(i, "K").Value = gTipoPeriodo And _
           ws.Cells(i, "L").Value = gPeriodo And _
           ws.Cells(i, "M").Value = dia Then
            
            BuscarCodigoInc = CStr(ws.Cells(i, "O").Value) 'CodigoInc col O
            Exit Function
        End If
    Next i

    BuscarCodigoInc = ""
End Function


'====================================================
' 5. Obtener Adicional, Observaciones y Bono comedor
'    en el periodo actual (primer valor no vacío)
'====================================================
Private Sub ObtenerAdicionalYObs( _
        ByVal numEmp As Long, _
        ByRef adicional As String, _
        ByRef obs As String, _
        Optional ByRef bonoComedor As Variant)

    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long

    adicional = ""
    obs = ""
    bonoComedor = Empty

    Set ws = ThisWorkbook.Worksheets("BDIncidencias_Local")
    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultFila
        If ws.Cells(i, "A").Value = gLoc And _
           ws.Cells(i, "C").Value = numEmp And _
           ws.Cells(i, "I").Value = gAnio And _
           ws.Cells(i, "J").Value = gMes And _
           ws.Cells(i, "K").Value = gTipoPeriodo And _
           ws.Cells(i, "L").Value = gPeriodo Then

            If adicional = "" Then adicional = CStr(ws.Cells(i, "P").Value) 'Adicional
            If obs = "" Then obs = CStr(ws.Cells(i, "Q").Value)            'Observacion

            ' Bono comedor en columna 22 (V), primer valor no vacío
            If IsEmpty(bonoComedor) Or bonoComedor = "" Then
                bonoComedor = ws.Cells(i, 22).Value
            End If
            
            If adicional <> "" And obs <> "" And _
               Not (IsEmpty(bonoComedor) Or bonoComedor = "") Then Exit For
        End If
    Next i
End Sub


'====================================================
' 6. ¿Ya existen incidencias para este empleado / periodo?
'====================================================
Public Function ExisteIncidenciasEmpleadoPeriodo(ByVal numEmp As Long) As Boolean

    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("BDIncidencias_Local")
    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultFila
        If ws.Cells(i, 1).Value = gLoc _
           And ws.Cells(i, 3).Value = numEmp _
           And ws.Cells(i, 9).Value = gAnio _
           And ws.Cells(i, 10).Value = gMes _
           And ws.Cells(i, 11).Value = gTipoPeriodo _
           And ws.Cells(i, 12).Value = gPeriodo Then

            ExisteIncidenciasEmpleadoPeriodo = True
            Exit Function
        End If
    Next i

End Function


'====================================================
' 7. Borrar TODAS las incidencias del empleado / periodo
'====================================================
Public Sub BorrarIncidenciasEmpleadoPeriodo(ByVal numEmp As Long)

    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("BDIncidencias_Local")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Asegurar que la hoja NO esté protegida
    On Error Resume Next
    ws.Unprotect "AVASA"
    ws.Unprotect "IncidenciasAVASA"   'por si quedó con la vieja
    On Error GoTo 0

    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    'Borrar de abajo hacia arriba
    For i = ultFila To 2 Step -1
        If ws.Cells(i, 1).Value = gLoc _
           And ws.Cells(i, 3).Value = numEmp _
           And ws.Cells(i, 9).Value = gAnio _
           And ws.Cells(i, 10).Value = gMes _
           And ws.Cells(i, 11).Value = gTipoPeriodo _
           And ws.Cells(i, 12).Value = gPeriodo Then

            ws.Rows(i).Delete
        End If
    Next i

    ' Si quieres que la BD quede protegida:
    On Error Resume Next
    ws.Protect Password:="AVASA"
    On Error GoTo 0

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

'====================================================
' 8. Crear / actualizar un botón (shape) en la hoja
'====================================================
Private Sub CrearOBoton( _
        ByVal ws As Worksheet, _
        ByVal nombre As String, _
        ByVal texto As String, _
        ByVal macro As String, _
        ByVal col As Long)

    Dim sh As Shape
    Dim leftPos As Double, topPos As Double
    Dim ancho As Double, alto As Double

    'Buscar si ya existe el botón
    On Error Resume Next
    Set sh = ws.Shapes(nombre)
    On Error GoTo 0

    'Altura fija de la primera fila
    ws.Rows(1).RowHeight = 37.5

    'Posición y tamaño dentro de la columna indicada
    leftPos = ws.Columns(col).Left + 3
    topPos = ws.Rows(1).Top + 2
    ancho = ws.Columns(col).Width - 6
    alto = ws.Rows(1).Height - 12

    If ancho < 40 Then ancho = 40
    If alto < 18 Then alto = 18

    'Crear o actualizar el shape
    If sh Is Nothing Then
        Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, ancho, alto)
        sh.Name = nombre
    Else
        sh.Left = leftPos
        sh.Top = topPos
        sh.Width = ancho
        sh.Height = alto
    End If

    'Formato del botón
    With sh
        'Texto centrado usando TextFrame (no TextFrame2)
        .TextFrame.Characters.Text = texto
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter

        'Macro asignada
        .OnAction = macro

        'Estilo visual
        .Fill.ForeColor.RGB = RGB(0, 102, 153)
        .Line.ForeColor.RGB = RGB(0, 51, 102)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Size = 10
    End With

End Sub
'====================================================
' 9. Generar / actualizar matriz del periodo actual
'    SIEMPRE desde Empleados (todos los de la locación)
'    y overlay de códigos desde BDIncidencias_Local
'====================================================
Public Sub GenerarMatrizPeriodoActual()

    Dim wsMat As Worksheet
    Dim wsBD As Worksheet
    Dim wsEmp As Worksheet

    Dim diaIni As Long, diaFin As Long, numDias As Long
    Dim baseColDias As Long, lastDiaCol As Long
    Dim colAdic As Long, colObs As Long, colBono As Long
    Dim lastHeaderCol As Long

    Dim fIni As Date, fFin As Date, mesNombre As String, titulo As String

    Dim dictCod As Object, dictAdic As Object, dictObs As Object, dictBono As Object

    Dim lastEmpRow As Long, r As Long, filaOut As Long
    Dim numEmp As Long, d As Long, k As String, cod As String
    Dim adicional As String, obs As String, bono As Variant
    Dim empGrupo As String

    '----------------------------------------
    ' Validación de globals
    '----------------------------------------
    If gLoc = "" Or gAnio = 0 Or gMes = 0 Or gTipoPeriodo = "" Or gPeriodo = 0 Then
        MsgBox "Falta definir locación o periodo. Entra por el menú.", vbCritical
        Exit Sub
    End If

    Set wsBD = ThisWorkbook.Worksheets("BDIncidencias_Local")
    Set wsEmp = ThisWorkbook.Worksheets("Empleados")
    Set wsMat = GetHojaMatrizPeriodo()
    If wsMat Is Nothing Then Exit Sub

    '----------------------------------------
    ' Botones fila 1 (si quieres)
    '----------------------------------------
    CrearOBoton wsMat, "btnAgregarIncidencia", "Agregar", "'" & ThisWorkbook.Name & "'!modReporteIncidencias.BotonAgregarIncidencia", 1
    CrearOBoton wsMat, "btnEditarIncidencia", "Editar", "'" & ThisWorkbook.Name & "'!modReporteIncidencias.BotonEditarIncidencia", 2
    CrearOBoton wsMat, "btnEliminarIncidencia", "Eliminar", "'" & ThisWorkbook.Name & "'!modReporteIncidencias.BotonEliminarIncidencia", 3
    CrearOBoton wsMat, "btnMenuIncidencias", "Menú", "'" & ThisWorkbook.Name & "'!modReporteIncidencias.BotonMenuPrincipal", 4

    '----------------------------------------
    ' Rango de días y título
    '----------------------------------------
    GetRangoPeriodo gAnio, gMes, gTipoPeriodo, gPeriodo, diaIni, diaFin
    numDias = diaFin - diaIni + 1

    baseColDias = 9 ' I
    lastDiaCol = baseColDias + numDias - 1

    fIni = DateSerial(gAnio, gMes, diaIni)
    fFin = DateSerial(gAnio, gMes, diaFin)
    mesNombre = UCase$(Format$(fIni, "mmmm"))

    titulo = "Incidencias AVASA " & LCase$(gTipoPeriodo) & " " & _
             Format$(fIni, "dd") & "-" & Format$(fFin, "dd") & " " & _
             mesNombre & " " & gAnio

    '----------------------------------------
    ' Limpiar matriz (dejando fila 1)
    '----------------------------------------
    With wsMat
        .Rows("2:2000").ClearContents
        .Rows("2:2000").Interior.ColorIndex = xlNone
        .Rows("2:2000").Borders.LineStyle = xlNone

        .Rows(1).RowHeight = 30
        .Rows(1).UnMerge

        If lastDiaCol < 6 Then lastDiaCol = 6

        .Range(.Cells(1, 6), .Cells(1, lastDiaCol)).Merge
        With .Range("F1")
            .Value = titulo
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 14
        End With

        'Encabezados
        .Range("A2").Value = "GRUPO"
        .Range("B2").Value = "CIUDAD"
        .Range("C2").Value = "NumeroEmpleado"
        .Range("D2").Value = "UsuarioCARs+"
        .Range("E2").Value = "DriverCARs+"
        .Range("F2").Value = "Puesto"
        .Range("G2").Value = "Actividad"
        .Range("H2").Value = "Nombre"

        For d = 0 To numDias - 1
            .Cells(2, baseColDias + d).Value = diaIni + d
        Next d

        colAdic = baseColDias + numDias
        colObs = colAdic + 1

        .Cells(2, colAdic).Value = "Adicional"
        .Cells(2, colObs).Value = "Observaciones"

        If UCase$(gLoc) = "CAP" Then
            colBono = colObs + 1
            .Cells(2, colBono).Value = "Bono comedor"
        Else
            colBono = 0
        End If

        lastHeaderCol = IIf(colBono > 0, colBono, colObs)

        With .Range(.Cells(2, 1), .Cells(2, lastHeaderCol))
            .Interior.Color = RGB(240, 200, 80)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With

        .Columns("A:H").ColumnWidth = 12
        For d = baseColDias To lastDiaCol
            .Columns(d).ColumnWidth = 3
        Next d
        .Columns(colAdic).ColumnWidth = 12
        .Columns(colObs).ColumnWidth = 18
        If colBono > 0 Then .Columns(colBono).ColumnWidth = 14
    End With

    '----------------------------------------
    ' Diccionarios desde BD (solo periodo actual)
    '----------------------------------------
    Set dictCod = CreateObject("Scripting.Dictionary")   ' key: emp|dia  -> codigo
    Set dictAdic = CreateObject("Scripting.Dictionary")  ' key: emp      -> adicional
    Set dictObs = CreateObject("Scripting.Dictionary")   ' key: emp      -> obs
    Set dictBono = CreateObject("Scripting.Dictionary")  ' key: emp      -> bono

    Dim ultFilaBD As Long, i As Long
    Dim diaBD As Long, empBD As Long

    ultFilaBD = wsBD.Cells(wsBD.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultFilaBD
        If wsBD.Cells(i, "A").Value = gLoc _
           And wsBD.Cells(i, "I").Value = gAnio _
           And wsBD.Cells(i, "J").Value = gMes _
           And UCase$(CStr(wsBD.Cells(i, "K").Value)) = UCase$(CStr(gTipoPeriodo)) _
           And wsBD.Cells(i, "L").Value = gPeriodo Then

            empBD = CLng(wsBD.Cells(i, "C").Value)
            diaBD = DiaAsLong(wsBD.Cells(i, "M").Value)

            If diaBD >= diaIni And diaBD <= diaFin Then
                k = CStr(empBD) & "|" & CStr(diaBD)
                cod = modCatalogoIncidencias.CanonizarCodigo(CStr(wsBD.Cells(i, "O").Value))
                dictCod(k) = cod
            End If

            If Not dictAdic.Exists(CStr(empBD)) Then
                If Len(Trim$(CStr(wsBD.Cells(i, "P").Value))) > 0 Then dictAdic(CStr(empBD)) = CStr(wsBD.Cells(i, "P").Value)
            End If

            If Not dictObs.Exists(CStr(empBD)) Then
                If Len(Trim$(CStr(wsBD.Cells(i, "Q").Value))) > 0 Then dictObs(CStr(empBD)) = CStr(wsBD.Cells(i, "Q").Value)
            End If

            If colBono > 0 Then
                If Not dictBono.Exists(CStr(empBD)) Then
                    If Len(Trim$(CStr(wsBD.Cells(i, 22).Value))) > 0 Then dictBono(CStr(empBD)) = wsBD.Cells(i, 22).Value
                End If
            End If
        End If
    Next i

    '----------------------------------------
    ' Pintar matriz SIEMPRE desde Empleados (todos locación)
    '----------------------------------------
    lastEmpRow = wsEmp.Cells(wsEmp.Rows.Count, "A").End(xlUp).Row
    filaOut = 3

    For r = 2 To lastEmpRow
        empGrupo = Trim$(CStr(wsEmp.Cells(r, "A").Value))
        If UCase$(empGrupo) = UCase$(gLoc) Then

            If IsNumeric(wsEmp.Cells(r, "C").Value) Then
                numEmp = CLng(wsEmp.Cells(r, "C").Value)
            Else
                GoTo NextEmp
            End If

            wsMat.Cells(filaOut, 1).Value = wsEmp.Cells(r, "A").Value
            wsMat.Cells(filaOut, 2).Value = wsEmp.Cells(r, "B").Value
            wsMat.Cells(filaOut, 3).Value = numEmp
            wsMat.Cells(filaOut, 4).Value = wsEmp.Cells(r, "D").Value
            wsMat.Cells(filaOut, 5).Value = wsEmp.Cells(r, "E").Value
            wsMat.Cells(filaOut, 6).Value = wsEmp.Cells(r, "F").Value
            wsMat.Cells(filaOut, 7).Value = wsEmp.Cells(r, "G").Value
            wsMat.Cells(filaOut, 8).Value = wsEmp.Cells(r, "H").Value

            'Overlay de códigos por día desde BD
            For d = diaIni To diaFin
                k = CStr(numEmp) & "|" & CStr(d)
                If dictCod.Exists(k) Then
                    cod = CStr(dictCod(k))
                    If Len(cod) > 0 Then wsMat.Cells(filaOut, baseColDias + (d - diaIni)).Value = cod
                End If
            Next d

            'Adic/Obs/Bono
            adicional = ""
            obs = ""
            bono = Empty

            If dictAdic.Exists(CStr(numEmp)) Then adicional = dictAdic(CStr(numEmp))
            If dictObs.Exists(CStr(numEmp)) Then obs = dictObs(CStr(numEmp))
            If colBono > 0 Then
                If dictBono.Exists(CStr(numEmp)) Then bono = dictBono(CStr(numEmp))
            End If

            wsMat.Cells(filaOut, colAdic).Value = adicional
            wsMat.Cells(filaOut, colObs).Value = obs
            If colBono > 0 Then wsMat.Cells(filaOut, colBono).Value = bono

            filaOut = filaOut + 1
        End If

NextEmp:
    Next r

    wsMat.Activate
    wsMat.Range("A1").Select

    If filaOut = 3 Then
        MsgBox "No encontré empleados en la hoja Empleados para la locación " & gLoc & ".", vbInformation
    End If

End Sub




'====================================================
' 10. Botón "Agregar incidencia" en la matriz
'====================================================
Public Sub BotonAgregarIncidencia()
    On Error GoTo Salir
    If Not SetGlobalsDesdeHojaMatriz(ActiveSheet) Then
    MsgBox "Esta hoja no parece ser una matriz válida (M_LOC_AAAA_MM_Q#/S#).", vbExclamation
    Exit Sub
End If
    modSeguridadIncidencias.ValidarPeriodoAbiertoOrExit

    Load frmIncidencias
    frmIncidencias.EnEdicion = False
    frmIncidencias.Show

Salir:
    If Err.Number <> 0 Then Exit Sub
End Sub

'====================================================
' 11. Botón "Editar incidencia" en la matriz
'====================================================
Public Sub BotonEditarIncidencia()

On Error GoTo Salir
If Not SetGlobalsDesdeHojaMatriz(ActiveSheet) Then
    MsgBox "Esta hoja no parece ser una matriz válida (M_LOC_AAAA_MM_Q#/S#).", vbExclamation
    Exit Sub
End If
modSeguridadIncidencias.ValidarPeriodoAbiertoOrExit
    Dim ws As Worksheet
    Dim fila As Long
    Dim numEmp As Long
    
    Set ws = ActiveSheet
    fila = ActiveCell.Row
    
    If fila < 3 Then
        MsgBox "Selecciona primero la fila del empleado que quieres editar.", vbExclamation
        Exit Sub
    End If
    
    If ws.Cells(fila, 3).Value = "" Then
        MsgBox "La fila seleccionada no contiene número de empleado.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(ws.Cells(fila, 3).Value) Then
        MsgBox "El número de empleado de la fila seleccionada no es válido.", vbExclamation
        Exit Sub
    End If
    
    numEmp = CLng(ws.Cells(fila, 3).Value)
    
    Load frmIncidencias
    frmIncidencias.CargarDesdeBD numEmp
    frmIncidencias.Show
    
Salir:
If Err.Number <> 0 Then Exit Sub



End Sub


'====================================================
' 12. Botón "Eliminar incidencia" en la matriz
'====================================================
Public Sub BotonEliminarIncidencia()
On Error GoTo Salir
If Not SetGlobalsDesdeHojaMatriz(ActiveSheet) Then
    MsgBox "Esta hoja no parece ser una matriz válida (M_LOC_AAAA_MM_Q#/S#).", vbExclamation
    Exit Sub
End If
modSeguridadIncidencias.ValidarPeriodoAbiertoOrExit
    Dim ws As Worksheet
    Dim fila As Long
    Dim numEmp As Long
    Dim resp As VbMsgBoxResult
    
    Set ws = ActiveSheet
    fila = ActiveCell.Row
    
    If fila < 3 Then
        MsgBox "Selecciona primero la fila del empleado que quieres eliminar.", vbExclamation
        Exit Sub
    End If
    
    If ws.Cells(fila, 3).Value = "" Then
        MsgBox "La fila seleccionada no contiene número de empleado.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(ws.Cells(fila, 3).Value) Then
        MsgBox "El número de empleado de la fila seleccionada no es válido.", vbExclamation
        Exit Sub
    End If
    
    numEmp = CLng(ws.Cells(fila, 3).Value)
    
    If Not ExisteIncidenciasEmpleadoPeriodo(numEmp) Then
        MsgBox "No hay incidencias registradas para este empleado en el periodo actual.", vbInformation
        Exit Sub
    End If
    
    resp = MsgBox( _
        "Se eliminarán TODAS las incidencias de este empleado " & vbCrLf & _
        "para la locación y periodo actuales." & vbCrLf & vbCrLf & _
        "¿Deseas continuar?", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar incidencias")
    
    If resp = vbNo Then Exit Sub
    
    BorrarIncidenciasEmpleadoPeriodo numEmp
    GenerarMatrizPeriodoActual
Salir:
If Err.Number <> 0 Then Exit Sub
End Sub


'====================================================
' 13. Botón "Menú" en la matriz
'====================================================
Public Sub BotonMenuPrincipal()

    'Opcional: activar hoja "Menu" si existe
    On Error Resume Next
    ThisWorkbook.Worksheets("Menu").Activate
    On Error GoTo 0

    'Mostrar el menú principal
    frmMenuPrincipal.Show

End Sub

Private Function DiaAsLong(ByVal v As Variant) As Long
    On Error GoTo Fail

    If IsError(v) Or IsEmpty(v) Then GoTo Fail

    If IsDate(v) Then
        DiaAsLong = Day(CDate(v))
        Exit Function
    End If

    If IsNumeric(v) Then
        DiaAsLong = CLng(v)
        Exit Function
    End If

Fail:
    DiaAsLong = 0
End Function

Private Function SetGlobalsDesdeHojaMatriz(ByVal ws As Worksheet) As Boolean
    On Error GoTo Fail

    Dim p() As String, tipoTag As String

    If ws Is Nothing Then GoTo Fail
    If Left$(ws.Name, 2) <> "M_" Then GoTo Fail

    p = Split(ws.Name, "_")
    ' Esperado: M, LOC, YYYY, MM, Q1/S4
    If UBound(p) < 4 Then GoTo Fail

    gLoc = p(1)
    gAnio = CLng(p(2))
    gMes = CLng(p(3))

    tipoTag = UCase$(p(4)) 'Q1 o S4
    If Left$(tipoTag, 1) = "Q" Then
        gTipoPeriodo = "QUINCENAL"
        gPeriodo = CLng(Mid$(tipoTag, 2))
    ElseIf Left$(tipoTag, 1) = "S" Then
        gTipoPeriodo = "SEMANAL"
        gPeriodo = CLng(Mid$(tipoTag, 2))
    Else
        GoTo Fail
    End If

    SetGlobalsDesdeHojaMatriz = True
    Exit Function

Fail:
    SetGlobalsDesdeHojaMatriz = False
End Function




