Attribute VB_Name = "modUID"


Option Explicit

'==========================
' UID OFICIAL (INCIDENCIAS POR DÍA)
' Formato: LOC|NUMEMP|AÑO|MM|TIPO|PERIODO|DIA
'==========================
Public Function BuildUID_Incidencia( _
    ByVal loc As String, ByVal numEmp As Long, ByVal anio As Long, _
    ByVal mes As Long, ByVal tipoPeriodo As String, ByVal periodo As Long, _
    ByVal dia As Long) As String

    BuildUID_Incidencia = UCase$(Trim$(loc)) & "|" & _
                          CStr(numEmp) & "|" & _
                          CStr(anio) & "|" & Format$(mes, "00") & "|" & _
                          UCase$(Trim$(tipoPeriodo)) & "|" & _
                          CStr(periodo) & "|" & CStr(dia)
End Function

'==========================
' UID ALTERNATIVO (POR FECHA REAL)
' Formato: LOC|AÑO|MM|TIPO|PERIODO|NUMEMP|YYYYMMDD
' Útil si luego ocupas llave por fecha completa.
'==========================
Public Function BuildUID_Fecha( _
    ByVal locCode As String, ByVal anio As Long, ByVal mes As Long, _
    ByVal tipoPeriodo As String, ByVal periodo As Long, _
    ByVal numEmp As Long, ByVal dt As Date) As String

    BuildUID_Fecha = UCase$(Trim$(locCode)) & "|" & _
                     CStr(anio) & "|" & Format$(mes, "00") & "|" & _
                     UCase$(Trim$(tipoPeriodo)) & "|" & CStr(periodo) & "|" & _
                     CStr(numEmp) & "|" & Format$(dt, "yyyymmdd")
End Function

'Convierte "día del periodo" a fecha real del periodo actual
Public Function FechaDeDiaPeriodo(ByVal anio As Long, ByVal mes As Long, ByVal dia As Long) As Date
    FechaDeDiaPeriodo = DateSerial(anio, mes, dia)
End Function


