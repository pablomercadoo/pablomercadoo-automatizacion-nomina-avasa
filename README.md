# pablomercadoo-automatizacion-nomina-avasa
Sistema en Excel/VBA para gestión de incidencias de nómina AVASA
AUTOMATIZACIÓN NOMINA AVASA — INCIDENCIAS (README v2)
====================================================

Contexto del proyecto
---------------------
Sistema en Excel/VBA para automatizar la gestión de incidencias de empleados en AVASA.
Objetivo: capturar/editar incidencias por periodo (semanal o quincenal), generar matrices de control por locación,
y preparar información para nómina. Especial foco en la excepción operativa CAP (Cancún Aeropuerto).

Principios clave (regla mental)
-------------------------------
- BDIncidencias_Local = verdad (fuente única)
- Matriz = vista temporal (se reconstruye, no se edita a mano)
- Forms = UI (captura/edición)
- Globals = estado (locación + periodo seleccionado)
- Config = reglas del negocio (cero hardcode si es regla de negocio)

Arquitectura por componentes (mapa rápido)
------------------------------------------
1) ThisWorkbook
   - Arranque/cierre:
     - Lee Config (locación, template).
     - Setea globals (gLoc, gLocDisplay, gIsTemplate).
     - Aplica protecciones (UIOnly) al abrir.
     - Muestra frmMenuPrincipal.
   - Al cerrar:
     - Limpia matrices antiguas/relevantes.
     - Oculta hojas y deja "Menu" visible.

2) modGlobal
   - Variables globales del estado actual:
     gAnio, gMes, gTipoPeriodo, gPeriodo, gLoc, gLocDisplay, gIsTemplate

3) frmMenuPrincipal (puerta de entrada del usuario)
   - Selección: año, mes, tipo (Semanal/Quincenal), periodo (#).
   - Valida no futuro (según fecha del sistema).
   - Sync empleados (una vez por periodo) via modEmpleadosSync.
   - Genera matriz del periodo via modReporteIncidencias.

4) modEmpleadosSync (sincronización empleados)
   - Lee DB externa (EmployeeDBPath + EmployeeDBTable en Config).
   - Filtra por locación (gLoc) y empleados activos (FechaBaja vacía).
   - Escribe hoja Empleados y crea tabla local tblEmpleados_Local.
   - Marca último PeriodID sincronizado en Config.

5) modReporteIncidencias (motor matriz)
   - Crea o recupera hoja matriz: M_LOC_AAAA_MM_Q#/S#
   - Reconstruye matriz completa:
     - Encabezados + días del periodo
     - Empleados (desde hoja Empleados)
     - Overlay incidencias (desde BDIncidencias_Local)
     - Botones (Agregar/Editar/Eliminar/Menú)
   - Los botones disparan apertura de frmIncidencias y operaciones de BD.

6) frmIncidencias (captura/edición)
   - Carga empleado desde Empleados (por NumeroEmpleado).
   - Muestra rango de días del periodo (hasta 16).
   - Valida códigos contra catálogo (tblCatalogoIncidencias).
   - Guarda en BDIncidencias_Local con UPSERT por UID.
   - En edición:
     - Si día queda vacío => borra registro del día.

7) modCatalogoIncidencias (catálogo de códigos)
   - Tabla: Config!tblCatalogoIncidencias
   - Normaliza / canoniza alias:
     - "T/D" -> "TD"
     - "FI" -> "F"
     - "0"  -> "" (vacío)
   - GetCodigosActivos() alimenta dropdowns del form.
   - EsCodigoValido() valida contra catálogo (activo).

8) modSeguridadIncidencias (cierre/protecciones)
   - LockWindowHours desde Config (default 48).
   - Periodo se considera cerrado si:
     Now >= FechaFinPeriodo + LockWindowHours/24
   - ValidarPeriodoAbiertoOrExit bloquea cambios si ya cerró.
   - SECURITY_ON controla si se aplican protecciones (DEV: False).

9) modMantenimientoMatrices
   - Limpia hojas M_ de una locación.
   - Conserva SOLO:
     - Periodo actual por fecha
     - Periodo anterior inmediato
   - Soporta semanal y quincenal.

10) modCalendario
   - Carga festivos de Config!tblFestivos a memoria.
   - Pinta encabezados de días:
     - DF (festivo) rojo suave
     - PD (domingo) gris suave
   - Sólo apoyo visual/validación (no “pisa” incidencias manuales).

11) modAdmin
   - Búsqueda/navegación de matrices históricas por código: AAAA_MM_Q#/S#

Modelo de datos
---------------
A) BDIncidencias_Local (tabla/hoja)
Cada fila = 1 incidencia de 1 empleado en 1 día de 1 periodo y 1 locación.
Campos clave (según implementación actual):
- Locación (A), Ciudad (B), NumeroEmpleado (C)
- UsuarioCARs+, DriverCARs+, Puesto, Actividad, Nombre
- Año (I), Mes (J), TipoPeriodo (K), Periodo (L), Día (M), Fecha (N)
- CodigoInc (O), Adicional (P), Observación (Q)
- CapturadoPor (R), FechaHora (S), Estatus (T)
- IDRegistro (U), BonoComedor (V), UID (W)

B) UID (clave lógica para UPSERT)
Formato (vigente en frmIncidencias):
  LOC|EMP|AÑO|MM|TIPO|PERIODO|DIA

Importante:
- UID evita duplicados y permite “mezclar” capturas por día.
- Si la lógica de UID cambia, hay que migrar/compatibilizar (ver TODOs).

Flujo end-to-end (operación normal)
-----------------------------------
1) Abrir archivo .xlsm (por locación o template)
2) Workbook_Open:
   - carga Config y globals
   - aplica protecciones
   - abre frmMenuPrincipal
3) En menú:
   - seleccionar año/mes/tipo/periodo
   - Sync empleados del periodo (una sola vez)
   - Generar matriz del periodo
4) En matriz:
   - Agregar / Editar / Eliminar incidencias (botones)
5) frmIncidencias:
   - validar catálogo
   - guardar (UPSERT) en BDIncidencias_Local
6) Regenerar matriz (vista actualizada)
7) Cierre automático:
   - después de la fecha fin del periodo + ventana

Excepción CAP (Cancún Aeropuerto)
--------------------------------
- Checador:
  - Entrada + salida = asistencia.
  - Traducción sugerida: domingo => PD, festivo => DF, cualquier otro => X.
  - Importante: checador NO debe pisar incidencias manuales.
- Bonos:
  - Bono fijo mensual (según puesto) calculado sobre 14 días.
  - Bono comedor: se paga por días asistidos.
- Reglas deben vivir en tablas/configuración (no en código).

Convenciones y nombres
----------------------
- Hojas matriz:
  M_<LOC>_<AAAA>_<MM>_<Q#|S#>
  Ej: M_CAP_2025_12_Q1
- Hoja de UI: "Menu"
- Hoja/tabla config: "Config" / tblConfig
- Tablas recomendadas:
  - tblConfig
  - tblCatalogoIncidencias
  - tblFestivos
  - (Locaciones) tblLocaciones
  - (Empleados local) tblEmpleados_Local

Repo / Git (sugerencia de estructura)
-------------------------------------
/docs
  README.md (este documento)
  MAPA.md
  BITACORA.md
  DECISION_LOG.md
/export
  VBA_TODO.txt (dump completo)
  components/ (export por .bas/.cls/.frm si aplica)
/excel
  template/ (archivo maestro)
  locaciones/ (salidas generadas)

Bitácora y disciplina de cambios (para no marearte)
---------------------------------------------------
Regla de oro: "Si cambia una regla del negocio, se documenta en DECISION_LOG antes de codificar."

Formato recomendado de entrada (DECISION_LOG.md):
- Fecha (YYYY-MM-DD)
- Contexto (qué problema)
- Decisión (qué se cambió)
- Impacto (qué módulos toca)
- Riesgos/pendientes
- Checklist de pruebas

Roadmap corto (próximos pasos típicos)
--------------------------------------
- Consolidar reglas CAP (checador + bonos) a tablas.
- Unificar implementación de UID (evitar duplicados de BuildUID en distintos módulos).
- Agregar “cierre por periodo seleccionado” (no solo por fecha del sistema, si aplica).
- Exportación a nómina + validaciones de consistencia.
- Pruebas: set de casos (semana/quincena, fin de mes, periodos cerrados, CAP vs no CAP).

Cómo usar este chat (reglas del juego)
--------------------------------------
- Este chat es el proyecto único.
- Cada vez que cambiemos algo, lo registramos como:
  1) DECISIÓN (qué regla cambió)
  2) CAMBIO (qué módulos/código)
  3) PRUEBA (cómo validamos)
- Si vas a pegar una guía nueva o un bloque grande:
  - pégalo con título y fecha (ej. "README v1 – 2025-12-18")

Add initial project README (v2)
