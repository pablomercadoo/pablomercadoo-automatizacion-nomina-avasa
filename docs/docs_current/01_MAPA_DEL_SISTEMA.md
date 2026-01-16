# 01 — MAPA DEL SISTEMA

## Objetivo
Este documento permite que cualquier persona técnica (o tú mismo en 6 meses):
- Entienda el sistema completo
- Sepa qué módulo hace qué
- Evite duplicar lógica
- Sepa dónde tocar y dónde NO tocar

---

## Visión general del sistema

Flujo principal (macro-nivel):

Workbook_Open  
→ frmMenuPrincipal  
→ Sync empleados  
→ Generar matriz  
→ Captura incidencias  
→ (Opcional) Precarga checador  
→ Cierre de periodo  

---

## Variables globales críticas
(definidas principalmente en `ModGlobal`)

- gLoc
- gAnio
- gMes
- gTipoPeriodo
- gPeriodo
- periodID
- gIsTemplate

Estas variables definen **el contexto completo del sistema**.

---

## Catálogo REAL de módulos (según export VBA)

### 🔴 CORE DEL SISTEMA (alto riesgo)

#### ModGlobal
**Responsabilidad:** Variables globales y helpers base  
**Riesgo:** Alto  
**Notas:** Cambios aquí afectan todo el sistema.

---

#### modConfig
**Responsabilidad:** Lectura y escritura de configuración (`tblConfig`)  
**Funciones clave:** GetConfig, SetConfig  
**Riesgo:** Alto  
**Regla:** Nunca duplicar lógica de configuración fuera de este módulo.

---

#### modPeriodos
**Responsabilidad:** Lógica de periodos (validación, rangos, tipos)  
**Riesgo:** Alto  
**Usado por:** menú, checador, seguridad.

---

#### modReporteIncidencias
**Responsabilidad:** Generar y navegar matrices de incidencias  
**Funciones clave:**
- GenerarMatrizPeriodoActual
- IrAMatrizPeriodoActual  
**Riesgo:** Alto

---

#### modSeguridadIncidencias
**Responsabilidad:** Seguridad y control de edición  
**Funciones clave:**
- PermiteEdicionPeriodo
- ProtegerHojaMatriz
- SafeProtectSheets  
**Riesgo:** Alto

---

### 🟠 EMPLEADOS

#### modEmpleadosSync
**Responsabilidad:** Sincronización de empleados  
**Modos:**
- Local (tabla inyectada)
- Externo (archivo RH, si aplica)  
**Funciones clave:**
- SyncEmpleados_PeriodoActual
- BuildPeriodID  
**Riesgo:** Alto

---

#### modRefreshEmpleados
**Responsabilidad:** Refrescar empleados ya existentes  
**Riesgo:** Medio

---

#### modPushEmpleados
**Responsabilidad:** Empujar empleados a estructuras locales  
**Riesgo:** Medio

---

#### modEmpleadosEliminados
**Responsabilidad:** Manejo de empleados dados de baja  
**Riesgo:** Medio

---

### 🟡 INCIDENCIAS / CATÁLOGOS

#### modCatalogoIncidencias
**Responsabilidad:** Catálogo de tipos de incidencias  
**Riesgo:** Medio

---

#### modCatalogosPuestoActividad
**Responsabilidad:** Catálogo puesto / actividad  
**Riesgo:** Medio

---

#### modCachePuestoActividad
**Responsabilidad:** Cacheo de valores únicos (combos)  
**Riesgo:** Medio

---

#### modUID
**Responsabilidad:** Generación de identificadores únicos  
**Función clave:** BuildUID_Incidencia  
**Riesgo:** Medio

---

### 🟢 CHECADOR (locaciones específicas)

#### modChecadorLectura
**Responsabilidad:** Lectura del archivo de checador  
**Riesgo:** Medio

---

#### modChecadorPrecarga
**Responsabilidad:** Precarga del checador a BDIncidencias_Local  
**Riesgo:** Alto  
**Regla:** El checador SOLO pisa registros marcados como CHECADOR.

---

### 🔵 CALENDARIO / FECHAS

#### modCalendario
**Responsabilidad:** Manejo de fechas, días, rangos  
**Riesgo:** Medio

---

### 🟣 GENERACIÓN / MANTENIMIENTO

#### modGeneradorLocaciones
**Responsabilidad:** Generar archivos por locación  
**Riesgo:** Alto

---

#### modUpdaterLocaciones
**Responsabilidad:** Actualizar archivos de locación existentes  
**Riesgo:** Alto

---

#### modMantenimientoMatrices
**Responsabilidad:** Limpieza y mantenimiento de matrices  
**Riesgo:** Medio

---

### ⚪ ADMIN / UTILIDADES

#### modAdmin
**Responsabilidad:** Funciones administrativas internas  
**Riesgo:** Bajo

---

#### modAutoFixPaths
**Responsabilidad:** Corregir rutas automáticamente  
**Riesgo:** Medio

---

#### modExportVBA
**Responsabilidad:** Exportación del código VBA  
**Riesgo:** Bajo

---

#### modGeneradorTests
**Responsabilidad:** Generación de datos / pruebas internas  
**Riesgo:** Bajo

---

### 🖥️ FORMULARIOS (UI)

- frmMenuPrincipal — selección de periodo
- frmIncidencias — captura principal
- frmAgregarIncidencias — alta directa
- frmOpciones — opciones/configuración

**Regla UI:**  
Los forms NO deben contener reglas de negocio complejas.

---

## Reglas de arquitectura (no negociables)

- Un módulo = una responsabilidad
- No duplicar lógica entre módulos
- Forms = UI, no negocio
- Configuración solo vía modConfig
