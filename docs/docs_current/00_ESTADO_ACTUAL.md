# 00 — ESTADO ACTUAL DEL SISTEMA

## Información general
- Proyecto: Automatización Incidencias de Nómina AVASA
- Responsable: Pablo Mercado
- Fecha de última actualización: YYYY-MM-DD

## Versión
- Versión lógica: vX.Y.Z
- Estatus: 🟢 Estable / 🟡 Estable con riesgos / 🔴 Inestable
- Base de esta versión: [commit / tag / descripción]

## Producción
- Producción = archivos Excel generados por locación (≈62)
- Ruta estándar:
  REPORTES GERENTES\<LOC>\REPORTE DE INCIDENCIAS DE NOMINA\
- Template base en producción: [nombre exacto]

## Funcionalidad activa en producción
- Menú de selección de periodo: Sí
- Sync de empleados:
  - Modo local (tabla inyectada): Sí
  - Modo externo (archivo RH): No / Sí (especificar)
- Matriz de incidencias: Sí
- Precarga checador:
  - Locaciones: [lista]
- Cierre de periodo: Sí

## Observabilidad
- Sistema de LOG activo: Sí
- Ruta: C:\AVASA_TMP\_LOGS
- Eventos logueados actualmente:
  - Workbook_Open
  - cmdAceptar_Click
  - SyncEmpleados_PeriodoActual
  - PrecargarBDDesdeChecador
- Pendiente de log:
  - GenerarMatrizPeriodoActual
  - IrAMatrizPeriodoActual
  - CerrarPeriodo

## Componentes críticos
Si alguno falla, el sistema no puede usarse:
- Workbook_Open
- frmMenuPrincipal.cmdAceptar_Click
- modEmpleadosSync
- modReporteIncidencias
- BDIncidencias_Local
- modSeguridadIncidencias

## Riesgos actuales
- R1: Dependencia alta de macros → impacto alto
- R2: Cambios históricos no documentados → riesgo de regresión
- R3: Diferencias entre locaciones → errores específicos

## Pendientes estratégicos
1. Depuración de módulos redundantes
2. Completar mapeo técnico del sistema
3. Estabilizar versión “óptima” para producción
