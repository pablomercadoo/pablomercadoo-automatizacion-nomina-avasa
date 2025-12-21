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

## 🔁 Flujo end-to-end

1. Abrir archivo `.xlsm`
2. `Workbook_Open`
3. Selección de periodo en menú
4. Sync empleados
5. Generar matriz
6. Captura / edición en formulario
7. Guardado en BD
8. Regenerar matriz
9. Cierre automático del periodo

---

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


