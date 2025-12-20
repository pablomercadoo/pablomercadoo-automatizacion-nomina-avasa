VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIncidencias 
   Caption         =   "UserForm2"
   ClientHeight    =   6260
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11070
   OleObjectBlob   =   "frmIncidencias.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmIncidencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--- bandera para saber si estamos editando o capturando nuevo ---
Public EnEdicion As Boolean

'--- rango de días del periodo seleccionado ---
Private gDiaIni As Long
Private gDiaFin As Long

'--- datos fijos del empleado cargados desde la hoja Empleados ---
Private mGrupo As String
Private mCiudad As String

'--- glosario de incidencias disponibles ---
Private tiposInc() As String


'===================================================
'  INICIALIZAR FORMULARIO
'===================================================
Private Sub UserForm_Initialize()

    Dim i As Long, k As Long
    Dim codigos As Variant

    EnEdicion = False   'por default

    '-----------------------------------------
    ' 1) TIPOS DE INCIDENCIA (desde catálogo activo)
    '-----------------------------------------
    codigos = modCatalogoIncidencias.GetCodigosActivos()

    ' Preparar combos/labels de día
    For i = 1 To 16

        With Me.Controls("cboDia" & i)
            .Clear
            .AddItem "" 'permitimos vacío

            If IsArray(codigos) Then
                For k = LBound(codigos) To UBound(codigos)
                    .AddItem codigos(k)
                Next k
            End If

            .Visible = False

            ' Solo lista (sin captura manual)
            On Error Resume Next
            .Style = fmStyleDropDownList
            On Error GoTo 0
        End With

        With Me.Controls("lblDia" & i)
            .caption = ""
            .Visible = False
        End With

    Next i

    '-----------------------------------------
    ' 2) Mostrar locación desde Config
    '-----------------------------------------
    lblLocacion.caption = GetConfig("LocationDisplay", gLoc)

    '-----------------------------------------
    ' 2b) Bono comedor: solo para CAP
    '-----------------------------------------
    If UCase$(gLoc) = "CAP" Then
        txtBonoComedor.Enabled = True
        txtBonoComedor.Visible = True
        On Error Resume Next
        lblBonoComedor.Visible = True
        On Error GoTo 0
    Else
        txtBonoComedor.Value = ""
        txtBonoComedor.Enabled = False
        txtBonoComedor.Visible = False
        On Error Resume Next
        lblBonoComedor.Visible = False
        On Error GoTo 0
    End If

    '-----------------------------------------
    ' 3) Configurar periodo (labels de días y lblPeriodo)
    '-----------------------------------------
    ConfigurarPeriodo

End Sub

'===================================================
'  ÚLTIMO DÍA DEL MES
'===================================================
Private Function UltimoDiaMes(anio As Long, mes As Long) As Long
    UltimoDiaMes = Day(DateSerial(anio, mes + 1, 0))
End Function


'===================================================
'  OBTENER RANGO DE DÍAS DEL PERIODO
'===================================================
Private Sub ObtenerRangoPeriodo( _
        ByVal anio As Long, _
        ByVal mes As Long, _
        ByVal tipoPeriodo As String, _
        ByVal numPeriodo As Long, _
        ByRef diaIni As Long, _
        ByRef diaFin As Long)

    Dim uDia As Long
    uDia = UltimoDiaMes(anio, mes)

    Select Case UCase(tipoPeriodo)

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


'===================================================
'  CONFIGURAR LABELS/COMBOS DE DÍA Y TEXTO DEL PERIODO
'===================================================
Private Sub ConfigurarPeriodo()
    Dim numDias As Long
    Dim i As Long
    Dim dia As Long
    Dim fIni As Date, fFin As Date
    Dim nombreMes As String

    If gAnio = 0 Or gMes = 0 Or gTipoPeriodo = "" Or gPeriodo = 0 Then Exit Sub

    ObtenerRangoPeriodo gAnio, gMes, gTipoPeriodo, gPeriodo, gDiaIni, gDiaFin
    numDias = gDiaFin - gDiaIni + 1

    fIni = DateSerial(gAnio, gMes, gDiaIni)
    fFin = DateSerial(gAnio, gMes, gDiaFin)
    nombreMes = UCase(Format(fIni, "mmmm"))

    lblPeriodo.caption = Format(fIni, "dd") & "-" & _
                         Format(fFin, "dd") & " " & _
                         nombreMes & " " & gAnio

    For i = 1 To 16
        With Me.Controls("lblDia" & i)
            If i <= numDias Then
                dia = gDiaIni + i - 1
                .caption = CStr(dia)
                .Visible = True
            Else
                .caption = ""
                .Visible = False
            End If
        End With

        With Me.Controls("cboDia" & i)
            If i <= numDias Then
                .Visible = True
            Else
                .Value = ""
                .Visible = False
            End If
        End With
    Next i
End Sub


'===================================================
'  CARGAR DATOS DEL EMPLEADO DESDE HOJA "Empleados"
'===================================================
Private Function CargarEmpleado(ByVal numEmp As Long) As Boolean
    Dim ws As Worksheet
    Dim cel As Range

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets("Empleados")

    Set cel = ws.Columns("C").Find(What:=numEmp, LookAt:=xlWhole) 'C = NumeroEmpleado
    If cel Is Nothing Then
        CargarEmpleado = False
        Exit Function
    End If

    With cel.EntireRow
        mGrupo = .Cells(1, "A").Value
        mCiudad = .Cells(1, "B").Value

        txtNumEmpleado.Value = .Cells(1, "C").Value
        txtUsuarioCars.Value = .Cells(1, "D").Value
        txtDriverCars.Value = .Cells(1, "E").Value
        txtPuesto.Value = .Cells(1, "F").Value
        txtActividad.Value = .Cells(1, "G").Value
        txtNombre.Value = .Cells(1, "H").Value
    End With

    CargarEmpleado = True
    Exit Function

ErrHandler:
    CargarEmpleado = False
End Function


Private Sub txtNumEmpleado_AfterUpdate()
    Dim n As Long

    If Trim(txtNumEmpleado.Value) = "" Then Exit Sub

    If Not IsNumeric(txtNumEmpleado.Value) Then
        MsgBox "El número de empleado debe ser numérico.", vbExclamation
        txtNumEmpleado.SetFocus
        Exit Sub
    End If

    n = CLng(txtNumEmpleado.Value)

    If Not CargarEmpleado(n) Then
        MsgBox "Empleado no encontrado en la hoja 'Empleados'.", vbExclamation
        txtNumEmpleado.SetFocus
    End If
End Sub


'===================================================
'  Validar que un código de incidencia esté en glosario
'===================================================
Private Function EsCodigoValido(ByVal cod As String) As Boolean
    Dim k As Long
    cod = UCase$(Trim$(cod))

    'Vacío siempre es válido (sin incidencia)
    If cod = "" Then
        EsCodigoValido = True
        Exit Function
    End If

    For k = LBound(tiposInc) To UBound(tiposInc)
        If cod = tiposInc(k) Then
            EsCodigoValido = True
            Exit Function
        End If
    Next k

    EsCodigoValido = False
End Function


'===================================================
'  CARGAR INCIDENCIAS DESDE LA BD (para botón EDITAR)
'===================================================
Public Sub CargarDesdeBD(ByVal numEmp As Long)

    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long
    Dim dia As Long
    Dim idx As Long
    Dim adicional As String
    Dim obs As String
    Dim bono As Variant
    Dim cod As String

    EnEdicion = True

    ' Limpiar antes de cargar
    LimpiarIncidencias
    txtAdicional.Value = ""
    txtObservaciones.Value = ""
    On Error Resume Next
    txtBonoComedor.Value = ""
    On Error GoTo 0

    ' Datos del empleado
    If Not CargarEmpleado(numEmp) Then
        MsgBox "Empleado no encontrado en la hoja 'Empleados'.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("BDIncidencias_Local")
    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    adicional = ""
    obs = ""
    bono = ""

    For i = 2 To ultFila

        If ws.Cells(i, "A").Value = gLoc _
           And ws.Cells(i, "C").Value = numEmp _
           And ws.Cells(i, "I").Value = gAnio _
           And ws.Cells(i, "J").Value = gMes _
           And ws.Cells(i, "K").Value = gTipoPeriodo _
           And ws.Cells(i, "L").Value = gPeriodo Then

            dia = CLng(ws.Cells(i, "M").Value)
            idx = dia - gDiaIni + 1    '1..16

            If idx >= 1 And idx <= 16 Then

                cod = CStr(ws.Cells(i, "O").Value)
                cod = modCatalogoIncidencias.CanonizarCodigo(cod)

                ' Si el código no existe en el dropdown (catálogo), NO lo metas.
                ' Mejor avisamos: así detectas cat incompleto.
                If cod <> "" Then
                    If Not modCatalogoIncidencias.EsCodigoValido(cod) Then
                        MsgBox "Código en BD no existe en el catálogo: '" & cod & "'" & vbCrLf & _
                               "Empleado: " & numEmp & " | Día: " & dia & vbCrLf & _
                               "Solución: agrega el código al catálogo o ajusta alias en CanonizarCodigo.", _
                               vbExclamation, "Catálogo incompleto"
                    Else
                        Me.Controls("cboDia" & idx).Value = cod
                    End If
                End If

            End If

            If adicional = "" Then adicional = CStr(ws.Cells(i, "P").Value)
            If obs = "" Then obs = CStr(ws.Cells(i, "Q").Value)

            ' Bono comedor en col 22
            If bono = "" Then
                bono = ws.Cells(i, 22).Value
            End If

        End If
    Next i

    txtAdicional.Value = adicional
    txtObservaciones.Value = obs

    If UCase$(gLoc) = "CAP" Then
        txtBonoComedor.Value = bono
    End If

End Sub

'===================================================
'  GUARDAR INCIDENCIAS EN BDIncidencias_Local (UPSERT por UID)
'===================================================
Private Sub cmdGuardar_Click()
    On Error GoTo Salir

    modSeguridadIncidencias.ValidarPeriodoAbiertoOrExit

    Dim ws As Worksheet
    Dim i As Long
    Dim diaReal As Long
    Dim cod As String
    Dim f As Date
    Dim numEmp As Long
    Dim resp As VbMsgBoxResult
    Dim tieneNueva As Boolean

    Dim uid As String
    Dim rowUID As Long
    Dim ultimaFila As Long
    Dim nextID As Long

    '------------------------------------
    ' Validaciones básicas
    '------------------------------------
    If Trim$(txtNumEmpleado.Value) = "" Then
        MsgBox "Captura el número de empleado.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtNumEmpleado.Value) Then
        MsgBox "El número de empleado debe ser numérico.", vbExclamation
        Exit Sub
    End If

    If gAnio = 0 Or gMes = 0 Or gTipoPeriodo = "" Or gPeriodo = 0 Or gLoc = "" Then
        MsgBox "No hay periodo o locación seleccionados. Cierra y vuelve a entrar desde el menú.", vbCritical
        Exit Sub
    End If

    numEmp = CLng(txtNumEmpleado.Value)

    '------------------------------------
    ' Validar códigos contra catálogo (canonizando)
    '------------------------------------
    For i = 1 To 16
        If Me.Controls("cboDia" & i).Visible Then

            cod = modCatalogoIncidencias.CanonizarCodigo(Me.Controls("cboDia" & i).Value)

            If Not modCatalogoIncidencias.EsCodigoValido(cod) Then
                diaReal = CLng(Me.Controls("lblDia" & i).caption)
                MsgBox "El código '" & cod & "' del día " & diaReal & " no es válido.", vbExclamation
                Exit Sub
            End If
        End If
    Next i

    '------------------------------------
    ' ¿Hay algo capturado?
    '------------------------------------
    tieneNueva = False
    For i = 1 To 16
        If Me.Controls("cboDia" & i).Visible Then
            cod = modCatalogoIncidencias.CanonizarCodigo(Me.Controls("cboDia" & i).Value)
            If cod <> "" Then
                tieneNueva = True
                Exit For
            End If
        End If
    Next i

    If Not tieneNueva And Not EnEdicion Then
        MsgBox "No capturaste ninguna incidencia.", vbInformation
        Exit Sub
    End If

    '------------------------------------
    ' Hoja BD
    '------------------------------------
    Set ws = ThisWorkbook.Worksheets("BDIncidencias_Local")

    On Error Resume Next
    ws.Unprotect "AVASA"
    ws.Unprotect "IncidenciasAVASA"
    On Error GoTo 0

    ws.Columns(13).NumberFormat = "0"                   'M Dia
    ws.Columns(14).NumberFormat = "dd/mm/yyyy"          'N Fecha
    ws.Columns(19).NumberFormat = "yyyy-mm-dd hh:mm:ss" 'S FechaHora

    '------------------------------------
    ' Confirmación si NO es edición y ya existe data del empleado en ese periodo
    '------------------------------------
    If Not EnEdicion Then
        If ExisteIncidenciasEmpleadoPeriodo(numEmp) Then
            resp = MsgBox( _
                "Este empleado ya tiene incidencias en este periodo." & vbCrLf & _
                "Se actualizarán/mezclarán por día (UID)." & vbCrLf & vbCrLf & _
                "¿Continuar?", _
                vbQuestion + vbYesNo + vbDefaultButton2, "Actualizar incidencias")
            If resp = vbNo Then Exit Sub
        End If
    End If

    '------------------------------------
    ' NEXT ID (solo si insertamos nuevas filas)
    '------------------------------------
    nextID = GetNextID(ws)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '------------------------------------
    ' Recorrer días visibles (1..16)
    '------------------------------------
    For i = 1 To 16

        If Me.Controls("cboDia" & i).Visible Then

            diaReal = CLng(Me.Controls("lblDia" & i).caption)
            cod = modCatalogoIncidencias.CanonizarCodigo(Me.Controls("cboDia" & i).Value)
            f = DateSerial(gAnio, gMes, diaReal)

            uid = modUID.BuildUID_Incidencia(gLoc, numEmp, gAnio, gMes, gTipoPeriodo, gPeriodo, diaReal)

            rowUID = FindRowByUID(ws, uid) 'busca en columna W

            '--- Si está vacío y estamos EDITANDO: borrar el registro existente para ese día ---
            If cod = "" Then
                If EnEdicion Then
                    If rowUID > 0 Then
                        ws.Rows(rowUID).Delete
                    End If
                End If
                GoTo NextI
            End If

            '--- Insertar o actualizar ---
            If rowUID = 0 Then
                ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
                rowUID = ultimaFila

                ws.Cells(rowUID, 21).Value = nextID 'IDRegistro col U
                nextID = nextID + 1

                ws.Cells(rowUID, "W").Value = uid 'UID col W
            End If

            With ws
                .Cells(rowUID, 1).Value = gLoc                        'A GRUPO/loc
                .Cells(rowUID, 2).Value = lblLocacion.caption         'B CIUDAD
                .Cells(rowUID, 3).Value = numEmp                      'C NumeroEmpleado
                .Cells(rowUID, 4).Value = txtUsuarioCars.Value        'D UsuarioCARs+
                .Cells(rowUID, 5).Value = txtDriverCars.Value         'E DriverCARs+
                .Cells(rowUID, 6).Value = txtPuesto.Value             'F Puesto
                .Cells(rowUID, 7).Value = txtActividad.Value          'G Actividad
                .Cells(rowUID, 8).Value = txtNombre.Value             'H Nombre

                .Cells(rowUID, 9).Value = gAnio                       'I Año
                .Cells(rowUID, 10).Value = gMes                       'J Mes
                .Cells(rowUID, 11).Value = gTipoPeriodo               'K TipoPeriodo
                .Cells(rowUID, 12).Value = gPeriodo                   'L Periodo
                .Cells(rowUID, 13).Value = CLng(diaReal)              'M Dia
                .Cells(rowUID, 14).Value = f                          'N Fecha
                .Cells(rowUID, 14).NumberFormat = "dd/mm/yyyy"

                .Cells(rowUID, 15).Value = cod                        'O CodigoInc (YA CANONIZADO)
                .Cells(rowUID, 16).Value = txtAdicional.Value         'P Adicional
                .Cells(rowUID, 17).Value = UCase$(txtObservaciones.Value) 'Q Observacion

                .Cells(rowUID, 18).Value = Environ$("username")       'R CapturadoPor
                .Cells(rowUID, 19).Value = Now                        'S FechaHora
                .Cells(rowUID, 20).Value = "BORRADOR"                 'T Estatus

                If UCase$(gLoc) = "CAP" Then
                    .Cells(rowUID, 22).Value = txtBonoComedor.Value   'V Bono comedor
                Else
                    .Cells(rowUID, 22).Value = ""
                End If
            End With

        End If

NextI:
    Next i

    ' Proteger BD
    On Error Resume Next
    ws.Protect Password:="AVASA"
    On Error GoTo 0

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Incidencias guardadas/actualizadas.", vbInformation

    ' Limpieza UI
    LimpiarIncidencias
    txtAdicional.Value = ""
    txtObservaciones.Value = ""
    On Error Resume Next
    txtBonoComedor.Value = ""
    On Error GoTo 0

    txtNumEmpleado.Value = ""
    txtUsuarioCars.Value = ""
    txtDriverCars.Value = ""
    txtPuesto.Value = ""
    txtActividad.Value = ""
    txtNombre.Value = ""

    EnEdicion = False

    modReporteIncidencias.GenerarMatrizPeriodoActual

Salir:
    If Err.Number <> 0 Then
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "Error al guardar: " & Err.Description, vbCritical
    End If
End Sub

Private Sub LimpiarIncidencias()
    Dim i As Long
    For i = 1 To 16
        If Me.Controls("cboDia" & i).Visible Then
            Me.Controls("cboDia" & i).Value = ""
        End If
    Next i

    On Error Resume Next
    txtBonoComedor.Value = ""
    On Error GoTo 0
End Sub

Private Function FindRowByUID(ByVal ws As Worksheet, ByVal uid As String) As Long
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, "W").End(xlUp).Row
    For i = 2 To lastRow
        If CStr(ws.Cells(i, "W").Value) = uid Then
            FindRowByUID = i
            Exit Function
        End If
    Next i
    FindRowByUID = 0
End Function

Private Function GetNextID(ByVal ws As Worksheet) As Long
    'IDRegistro = columna U (21)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row

    If lastRow < 2 Or ws.Cells(lastRow, "U").Value = "" Then
        GetNextID = 1
    Else
        GetNextID = CLng(ws.Cells(lastRow, "U").Value) + 1
    End If
End Function

'===================================================
'  CERRAR
'===================================================
Private Sub cmdCerrar_Click()
    Unload Me
End Sub


