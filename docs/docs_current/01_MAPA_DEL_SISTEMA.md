# 01 — Mapa del sistema (para handoff)

Este documento explica **qué hace cada componente**, **quién lo llama** y **qué toca**.

## Idea central
- **BDIncidencias_Local = verdad** (fuente única).
- **Matriz = vista** (se puede regenerar).
- **Forms = UI** (no deben contener reglas duplicadas).
- **Módulos = lógica** (dominios claros: config, empleados, periodos, incidencias, checador, catálogos).

## Flujo principal (ruta diaria)
1) `ThisWorkbook.Workbook_Open`
2) `frmMenuPrincipal` selecciona periodo
3) `modEmpleadosSync.BuildPeriodID`
4) `modEmpleadosSync.SyncEmpleados_PeriodoActual`
5) `modReporteIncidencias` genera/abre matriz
6) `frmOpciones` menú operativo
7) `frmIncidencias` captura manual (alta/edición/borrado)
8) `modChecadorPrecarga` precarga checador (si aplica)
9) `modPeriodos` + `modSeguridadIncidencias` cierre y enforcement

## Componentes

### Entry points
- **ThisWorkbook**: arranque (leer config, setear globals, ocultar hojas técnicas, abrir menú).
- **frmMenuPrincipal**: entrada de usuario (periodo) y arranque del flujo.
- **frmOpciones**: acciones del periodo (captura, completar, limpiar, eliminar empleado, cerrar periodo).

### Dominios

#### Config / Globals
- `modConfig`: `GetConfig`, `SetConfig`, bloqueo de hoja Config.
- `modGlobal`: variables globales (estado de locación y periodo).

#### Empleados
- `modEmpleadosSync`: construir `periodID`, sincronizar empleados.
- `modEmpleadosEliminados`: reglas por periodo para excluir empleados.
- `modRefreshEmpleados` / `modPushEmpleados`: utilidades masivas (admin).

#### Incidencias / Matriz
- `modReporteIncidencias`: motor de matriz + helpers de BD.
- `modMantenimientoMatrices`: limpieza / parsing de matrices.
- `modUID`: UIDs por incidencia/fecha.

#### Checador
- `modChecadorLectura`: lectura/validación.
- `modChecadorPrecarga`: merge a BD (idempotente: no duplicar manual).

#### Periodos / Seguridad
- `modPeriodos`: estado abierto/cerrado + override.
- `modSeguridadIncidencias`: proteger/desproteger + validar periodo.

#### Catálogos
- `modCatalogoIncidencias`: validación/canonización de códigos.
- `modCatalogos`: mapeos locación/CC/checador.
- `modCatalogosPuestoActividad` + `modCachePuestoActividad`: combos y cache.
- `modCalendario`: festivos.

## Red flags a vigilar
- **Shadowing de nombres** (misma función en dos módulos/forms) → causa bugs fantasma.
- **MsgBox de OK** → sustituir por log/status.
- Utilerías duplicadas → centralizar en un solo módulo.

