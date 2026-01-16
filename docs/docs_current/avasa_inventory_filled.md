# INVENTARIO COMPLETADO - Sistema AVASA Nómina

> Análisis realizado el 2026-01-16 basado en export completo del código VBA

---

## ✅ MÓDULOS CORE (100% confirmado que se usan)

| Módulo | ¿Se usa? | ¿Para qué? | Notas |
|--------|----------|------------|-------|
| **ThisWorkbook** | ✅ SÍ | Entry point del sistema | Llama AutoFix, config, menú |
| **ModGlobal** | ✅ SÍ | Variables globales críticas | gLoc, gAnio, gMes, gTipoPeriodo, gPeriodo |
| **modConfig** | ✅ SÍ | Único acceso a config | GetConfig/SetConfig usado por TODOS |
| **modPeriodos** | ✅ SÍ | Lógica de rangos/períodos | Usado por matriz, sync, seguridad |
| **modReporteIncidencias** | ✅ SÍ | **GIGANTE (1000+ líneas)** | ⚠️ Generar matriz, CRUD, TODO mezclado |
| **modSeguridadIncidencias** | ✅ SÍ | Protección y permisos | Usado en menú, matriz, CRUD |
| **modEmpleadosSync** | ✅ SÍ | Sync de empleados | SyncEmpleados_PeriodoActual + BuildPeriodID |
| **modUID** | ✅ SÍ | Generación de UIDs | BuildUID_Incidencia usado en CRUD |
| **frmMenuPrincipal** | ✅ SÍ | Selección de período | Entry UI principal |
| **frmIncidencias** | ✅ SÍ | Captura/edición | Form principal de incidencias |
| **frmAgregarIncidencias** | ✅ SÍ | Selector manual/checador | Usado por matriz |
| **frmOpciones** | ✅ SÍ | Menú de opciones | Usado por matriz |

---

## ⚠️ MÓDULOS DE SOPORTE (Se usan, pero menos críticos)

| Módulo | ¿Se usa? | ¿Para qué? | Decisión |
|--------|----------|------------|----------|
| **modCalendario** | ✅ SÍ | Festivos, domingos | Usado por matriz, checador | ✅ KEEP |
| **modCatalogoIncidencias** | ✅ SÍ | Validación códigos | Usado por forms, matriz | ✅ KEEP |
| **modCatalogos** | ✅ SÍ | Terminal→CC→Loc | Usado por checador | ✅ KEEP |
| **modCatalogosPuestoActividad** | ✅ SÍ | Puestos/Actividades | Usado por forms, matriz | ✅ KEEP |
| **modChecadorPrecarga** | ✅ SÍ | Precarga checador | Llamado por frmAgregar | ✅ KEEP |
| **modChecadorLectura** | ✅ SÍ | Leer archivo checador | Llamado por Precarga | ✅ KEEP |
| **modAutoFixPaths** | ✅ SÍ | Corregir rutas | Llamado en Workbook_Open | ✅ KEEP |
| **modLog** | ✅ SÍ | Sistema de logs | Usado en varios módulos | ✅ KEEP |
| **modAdmin** | ✅ SÍ | Funciones admin | BuscarMatrizPeriodoAdmin | ✅ KEEP |
| **modEmpleadosEliminados** | ✅ SÍ | Control eliminados | Usado en matriz, completar | ✅ KEEP |

---

## 🔶 MÓDULOS TOOLING (Solo para mantenimiento)

| Módulo | ¿Se usa? | ¿Cuándo? | Decisión |
|--------|----------|----------|----------|
| **modGeneradorLocaciones** | ⚠️ OCASIONAL | Solo al crear archivos nuevos | ✅ KEEP (admin) |
| **modExportVBA** | ⚠️ OCASIONAL | Solo para export de código | ✅ KEEP (admin) |
| **modMantenimientoMatrices** | ⚠️ OCASIONAL | Limpieza de matrices viejas | 🟡 REVIEW |
| **modUpdaterLocaciones** | ⚠️ OCASIONAL | Actualizar archivos masivos | ✅ KEEP (admin) |

---

## 🔴 MÓDULOS SOSPECHOSOS (Posibles candidatos a eliminar)

| Módulo | ¿Alguien lo llama? | Análisis | Decisión |
|--------|-------------------|----------|----------|
| **modRefreshEmpleados** | ❌ NO (en código activo) | Solo define `Refresh_Empleados_Todos()` pero NADIE la llama | 🔴 **DELETE** |
| **modPushEmpleados** | ❌ NO (en código activo) | Solo define `PushEmpleados_A_TodasLasLocaciones()` pero NADIE la llama | 🔴 **DELETE** |
| **modCachePuestoActividad** | ⚠️ DUDOSO | Define `CacheUnicosPuestoActividad_Global()` que SÍ es llamado por `modEmpleadosSync` | ✅ KEEP |
| **modGeneradorTests** | ❌ NO (en código activo) | Solo para desarrollo/testing | 🟡 **MOVE** a archivo DEV separado |

---

## 🔍 FUNCIONES DUPLICADAS ENCONTRADAS

### ❌ DUPLICADO CONFIRMADO: `NormalizarTipoPeriodo`

**Ubicaciones encontradas:**
1. ✅ **modReporteIncidencias** (línea ~1450) ← **CANÓNICA**
2. ❌ **modChecadorPrecarga** (línea ~220) ← **DUPLICADO - ELIMINAR**
3. ❌ **modEmpleadosSync** (?) - revisar si existe

**Acción:** Centralizar en `modReporteIncidencias.NormalizarTipoPeriodo()` y eliminar duplicados

---

### ❌ DUPLICADO CONFIRMADO: `CleanPath`

**Ubicaciones encontradas:**
1. ✅ **modGeneradorLocaciones** (línea ~400) ← **CANÓNICA**
2. ❌ **modPushEmpleados** (línea ~250)
3. ❌ **modRefreshEmpleados** (?) - revisar

**Acción:** Mover a módulo de utilidades o dejar en `modGeneradorLocaciones` y que los demás la llamen

---

### ❌ DUPLICADO CONFIRMADO: `GetConfigValue`

**Ubicaciones encontradas:**
1. ✅ **modGeneradorLocaciones** (línea ~350)
2. ❌ **modPushEmpleados** (línea ~15)
3. ❌ **modGeneradorTests** (línea ~50)

**Acción:** TODOS deberían usar `modConfig.GetConfig()` directamente - ELIMINAR estas versiones

---

### ❌ DUPLICADO CONFIRMADO: `FindTableByName`

**Ubicaciones encontradas:**
1. **modEmpleadosSync** (línea ~120)
2. **modRefreshEmpleados** (línea ~180)
3. **modPushEmpleados** (línea ~90)

**Acción:** Centralizar en UN módulo de utilidades

---

### ❌ DUPLICADO CONFIRMADO: `GetFilteredTableData_Fast` / `GetFilteredTableData`

**Ubicaciones encontradas:**
1. **modGeneradorLocaciones** (línea ~450)
2. **modRefreshEmpleados** (línea ~200)
3. **modPushEmpleados** (línea ~110)

**Acción:** Centralizar en UN solo módulo

---

### ✅ NO DUPLICADO: `BuildPeriodID`

**Ubicación:** SOLO en `modEmpleadosSync` ← **CORRECTO**

---

### ✅ NO DUPLICADO: `GetPeriodRange` / `ObtenerRangoPeriodo`

**Ubicaciones:**
- `modReporteIncidencias.GetRangoPeriodo()` ← Privada, OK
- `modSeguridadIncidencias.ObtenerRangoPeriodo_Local()` ← Privada, OK
- `modSeguridadIncidencias.ObtenerRangoPeriodo_Public()` ← Wrapper público

**Acción:** Considerar centralizar en `modPeriodos` pero NO urgente

---

### ❌ DUPLICADO CONFIRMADO: `IsWebPath`

**Ubicaciones:**
1. **ThisWorkbook** (línea ~100) ← Privada
2. **modEmpleadosSync** (línea ~200) ← Privada

**Acción:** Mover a `modAutoFixPaths` como función pública

---

## 📊 RESUMEN EJECUTIVO

### Conteo Total

**Total de módulos:** 31 componentes VBA
- **Módulos estándar (.bas):** 23
- **Forms (.frm):** 4 
- **Hojas (Document):** 8

**Módulos que SÍ uso (CORE + SOPORTE):** 22
**Módulos que POSIBLEMENTE no uso:** 3
- modRefreshEmpleados (DELETE)
- modPushEmpleados (DELETE)
- modGeneradorTests (MOVE a DEV)

**Funciones duplicadas encontradas:** 7 casos
- NormalizarTipoPeriodo (2-3 ubicaciones)
- CleanPath (2-3 ubicaciones)
- GetConfigValue (3 ubicaciones)
- FindTableByName (3 ubicaciones)
- GetFilteredTableData (3 ubicaciones)
- IsWebPath (2 ubicaciones)
- ReplaceLocalTableData (2 ubicaciones)

---

## 🎯 PLAN DE ACCIÓN INMEDIATO

### FASE 1: Eliminar lo obvio (1 hora)

1. ❌ **Eliminar modRefreshEmpleados**
   - Nadie lo llama
   - Era para actualizar empleados desde master externo
   - Ya no se usa porque tienes `modEmpleadosSync`

2. ❌ **Eliminar modPushEmpleados**
   - Nadie lo llama
   - Era para "push masivo" de empleados
   - Ya no se usa

3. 📦 **Mover modGeneradorTests** a archivo separado
   - Solo para desarrollo
   - No debe estar en producción

---

### FASE 2: Consolidar duplicados (2-3 horas)

4. ✅ **Centralizar utilidades de tabla**
   ```
   Crear: modUtilsTablas (nuevo)
   └─ FindTableByName()
   └─ GetFilteredTableData()
   └─ ReplaceLocalTableData()
   ```

5. ✅ **Centralizar utilidades de path**
   ```
   modAutoFixPaths (ya existe)
   └─ CleanPath() ← mover aquí
   └─ IsWebPath() ← ya existe, hacer pública
   └─ PathExists()
   └─ FileExists()
   ```

6. ✅ **Eliminar GetConfigValue duplicado**
   - Todos deben usar `modConfig.GetConfig()` directamente

7. ✅ **Consolidar NormalizarTipoPeriodo**
   - Dejar solo en `modReporteIncidencias`
   - Hacer pública
   - Actualizar llamadas en `modChecadorPrecarga`

---

### FASE 3: Refactorizar modReporteIncidencias (1 semana)

8. 🔨 **Dividir el monstruo de 1000+ líneas**
   ```
   modReporteIncidencias (1000+ líneas) DIVIDIR EN:
   
   modMatrizUI (~200 líneas)
   └─ GenerarMatrizPeriodoActual()
   └─ PintarEncabezadoDias()
   └─ CrearOBoton()
   
   modIncidenciasBusiness (~150 líneas)
   └─ Validaciones
   └─ Reglas de negocio
   └─ GetOverridesEmpleadoPeriodo()
   
   modIncidenciasData (~200 líneas)
   └─ CRUD a BDIncidencias_Local
   └─ ExisteIncidenciasEmpleadoPeriodo()
   └─ BorrarIncidenciasEmpleadoPeriodo()
   
   modEmpleadosHelpers (~100 líneas)
   └─ UpsertEmpleadoTemporal()
   └─ EmpleadoExisteOficial()
   └─ Headers robustos (GetEmpCol_*)
   ```

---

## ✅ RESULTADO ESPERADO

### Antes de limpieza:
- 23 módulos estándar
- ~7 funciones duplicadas
- 1 módulo gigante (1000+ líneas)
- 3 módulos sin usar

### Después de limpieza:
- 22 módulos estándar (eliminados 3, creados 2 nuevos)
- 0 funciones duplicadas
- Módulo más grande: ~200 líneas
- Todo código en uso activo

### Beneficios:
- ✅ Más fácil de mantener
- ✅ Más fácil de entender
- ✅ Menos bugs por duplicación
- ✅ Mejor organización

---

## 🚦 SIGUIENTE PASO

**¿Qué prefieres hacer primero?**

1. 🔴 **Rápido (1 hora):** Eliminar los 3 módulos que NO se usan
2. 🟡 **Medio (3 horas):** Consolidar duplicados + eliminar módulos
3. 🟢 **Completo (1 semana):** Todo lo anterior + refactorizar modReporteIncidencias

**Mi recomendación:** Empezar con opción 2 (Medio) para ver mejora rápida sin riesgo alto.
