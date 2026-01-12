# pablomercadoo-automatizacion-nomina-avasa

Sistema en **Excel / VBA** para la gestión de incidencias de nómina AVASA.  
Diseñado para operar por **locación y periodo**, con control de cierre, precarga de checador y trazabilidad para Nómina.

**README v3 — 2025-12-20**

---

## 🎯 Contexto del proyecto

Este sistema automatiza la captura, edición y control de incidencias de empleados en AVASA.

**Objetivo principal**:
- Capturar incidencias por empleado y por día
- Operar por periodo **semanal o quincenal**
- Generar matrices visuales por locación
- Consolidar información en una base única para Nómina
- Controlar cierres automáticos y manuales de periodo

Existe una **excepción operativa importante** para **CAP (Cancún Aeropuerto)** que se maneja por reglas de negocio específicas.

---

## 🧠 Principios clave (regla mental del sistema)

- **BDIncidencias_Local = verdad**
  - Fuente única de información
- **Matriz = vista temporal**
  - Se reconstruye siempre, no se edita a mano
- **Forms = UI**
  - Captura y edición controlada
- **Globals = estado**
  - Locación + periodo seleccionado
- **Config = reglas del negocio**
  - Cero hardcode de reglas operativas

---

## 🧱 Arquitectura por componentes (mapa rápido)

### 1) ThisWorkbook
- **Al abrir**:
  - Lee `Config`
  - Setea globals (`gLoc`, `gLocDisplay`, `gIsTemplate`)
  - Aplica protecciones (`UIOnly`)
  - Muestra `frmMenuPrincipal`
- **Al cerrar**:
  - Limpia matrices antiguas
  - Oculta hojas técnicas
  - Deja visible `Menu`




---

### 2) modGlobal
Variables de estado:
gAnio, gMes, gTipoPeriodo, gPeriodo
gLoc, gLocDisplay, gIsTemplate


---

### 3) frmMenuPrincipal (entrada del usuario)
- Selección de:
  - Año / Mes
  - Tipo de periodo (Semanal / Quincenal)
  - Número de periodo
- Validaciones:
  - No permite periodos futuros
- Acciones:
  - Sincroniza empleados (`modEmpleadosSync`)
  - Genera matriz (`modReporteIncidencias`)

## 🧭 frmOpciones (botones reales)

- **AGREGAR INCIDENCIA**  
  Solo cuando el empleado **no aparece** en el registro. Se captura manualmente su número y datos base para que ya pueda trabajarse en el periodo.

- **EDITAR INCIDENCIA**  
  Aquí se capturan y corrigen las incidencias por día.  
  ⚠️ Para “agregar incidencias”, SIEMPRE se usa este botón (aunque el empleado haya sido agregado manualmente).

- **COMPLETAR INCIDENCIAS**  
  Autocompleta asistencias faltantes (X/PD/DF) para cerrar limpio el periodo.  
  Regla: aquí se generan las **no-incidencias** (asistencias), no el gerente.

- **LIMPIAR INCIDENCIAS**  
  Borra incidencias del periodo actual (acción delicada).

- **ELIMINAR EMPLEADO**  
  Quita al empleado del periodo actual (si fue agregado por error o no corresponde).

- **CAMBIAR PERIODO**  
  Salir y seleccionar otro periodo.

- **CERRAR PERIODO**  
  Bloquea el periodo y deja el archivo en solo lectura para ese periodo.

- **CERRAR MENÚ**  
  Cierra la ventana del menú.


---

### 4) modEmpleadosSync
- Lee DB externa de empleados (ruta y tabla desde `Config`)
- Filtra:
  - Por locación (`gLoc`)
  - Solo empleados activos
- Escribe hoja `Empleados`
- Genera `tblEmpleados_Local`
- Marca último periodo sincronizado en `Config`

---

### 5) modReporteIncidencias (motor de matrices)
- Crea / recupera hojas:
M_<LOC><AAAA><MM>_<Q#|S#>

- Reconstruye completamente la matriz:
- Encabezados y días del periodo
- Empleados
- Overlay de incidencias desde BD
- Botones de acción

---

### 6) frmIncidencias (captura / edición)
- Carga datos del empleado desde `Empleados`
- Muestra hasta **16 días** por periodo
- Valida códigos contra catálogo
- Guarda en `BDIncidencias_Local` usando **UPSERT por UID**
- En edición:
- Día vacío ⇒ borra registro del día

---

### 7) modCatalogoIncidencias
- Tabla: `Config!tblCatalogoIncidencias`
- Funciones:
- Canonización de alias  
  (`"T/D" → "TD"`, `"FI" → "F"`, `"0" → ""`)
- `GetCodigosActivos()` para dropdowns
- `EsCodigoValido()` para validación

---

### 8) modSeguridadIncidencias
- Ventana de cierre configurable (`LockWindowHours`, default 48)
- Periodo cerrado si:
Now >= FechaFinPeriodo + LockWindowHours

- `ValidarPeriodoAbiertoOrExit` bloquea cambios
- `SECURITY_ON` permite modo DEV / PROD

---

### 9) modMantenimientoMatrices
- Limpia hojas `M_`
- Conserva solo:
- Periodo actual
- Periodo inmediato anterior
- Soporta semanal y quincenal

---

### 10) modCalendario
- Carga festivos desde `Config!tblFestivos`
- Pintado visual:
- Domingo → gris (PD)
- Festivo → rojo suave (DF)
- **No pisa incidencias**

---

### 11) modAdmin
- Navegación de matrices históricas
- Acceso por código de periodo (`AAAA_MM_Q#/S#`)

---

## 🗄️ Modelo de datos

### A) BDIncidencias_Local
Cada fila representa **1 incidencia de 1 empleado en 1 día**.

Campos principales:
- Locación, Ciudad, NumeroEmpleado
- UsuarioCARs+, DriverCARs+, Puesto, Actividad, Nombre
- Año, Mes, TipoPeriodo, Periodo, Día, Fecha
- CodigoInc, Adicional, Observación
- CapturadoPor, FechaHora, Estatus
- IDRegistro, BonoComedor, UID

---

### B) UID (clave lógica)
Formato vigente:
LOC|EMP|AÑO|MM|TIPO|PERIODO|DIA


- Evita duplicados
- Permite mezcla de capturas
- Cualquier cambio al UID requiere migración

---

## 🔁 Flujo end-to-end (operación real)

1. Abrir el archivo de la locación `Incidencias_<LOC>.xlsm`
2. Habilitar macros
3. Seleccionar periodo (año/mes/SEM o QUIN/número de periodo)
4. El sistema sincroniza empleados y genera la matriz del periodo
5. Capturar incidencias **solo por excepción** usando el menú:
   - Agregar incidencia (solo si el empleado no existe)
   - Editar incidencia (captura real de incidencias por día)
6. Ejecutar **Completar incidencias** para llenar automáticamente asistencias faltantes:
   - Día normal → `X`
   - Domingo → `PD`
   - Festivo → `DF`
7. Revisar y finalmente **Cerrar periodo** (bloquea cambios)

---

## 📦 Distribución V1 (62 archivos por locación)

La V1 se opera con **1 archivo por locación** (ej. `Incidencias_CUN.xlsm`).

### Estructura de carpetas (estándar)
Cada archivo se guarda en:

`<RAIZ>\ <LOC>\ REPORTE DE INCIDENCIAS DE NOMINA\ Incidencias_<LOC>.xlsm`

Donde:
- **RAIZ** = carpeta raíz de gerentes (OneDrive) definida en el template.
- **LOC** = código de locación (ej. CUN, MID, MTY, etc.)

### Cómo se generan los 62 archivos
Desde el **template** (archivo base), se ejecuta el generador:
- Lee `tblLocaciones` y toma las locaciones con `Active = 1`
- Crea carpetas faltantes
- Genera una copia `.xlsm` por locación
- En cada archivo nuevo “setea” su configuración (LocationCode, LocationName, CC, etc.)
- Marca `IsTemplate = 0` en cada archivo generado

📌 Resultado: cada gerente recibe **solo el archivo de su locación**, y trabaja por periodo.


## ✈️ Excepción CAP (Cancún Aeropuerto)

- Checador:
  - Entrada + salida = asistencia
  - Domingo → PD
  - Festivo → DF
  - Otro → X
- Checador **no debe pisar** incidencias manuales
- Bono comedor:
  - Pago por días asistidos
- Reglas deben vivir en **tablas**, no en código

---

## 📂 Convenciones

- Matrices:
- M_<LOC><AAAA><MM>_<Q#|S#>

- - UI: `Menu`
- Configuración: `Config / tblConfig`
- Tablas clave:
- `tblCatalogoIncidencias`
- `tblFestivos`
- `tblEmpleados_Local`
- `tblLocaciones`

---

## 🧾 Disciplina de cambios

**Regla de oro**:  
> Si cambia una regla de negocio, se documenta antes de codificar.

Formato `DECISION_LOG.md`:
- Fecha
- Contexto
- Decisión
- Impacto
- Riesgos
- Checklist de pruebas

---

## 🛣️ Roadmap corto

- Consolidar reglas CAP a tablas
- Unificar UID en todos los módulos
- Cierre por periodo seleccionado
- Exportación a Nómina
- Set de pruebas formales

---

README v3 cerrado — 2025-12-20


