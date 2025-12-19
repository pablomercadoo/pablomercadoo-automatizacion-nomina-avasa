# ROADMAP — Automatización Incidencias AVASA

Este documento define el orden de trabajo del proyecto.
Si algo no está aquí, NO se programa.

Última actualización: 2025-12-18

---

## VISIÓN GENERAL

- Todas las locaciones generan **el mismo entregable final**:
  - `BDIncidenciasLocal`
  - `tblPeriodos`
- El concentrado (Master) **no distingue locaciones**, solo lee datos estándar.
- Existe un **proceso CORE general** y **extensiones por locación** (ej. Cancún).
- Cancún es la locación piloto por ser la más grande.

---

## FASE 1 — CORE ESTABLE (CERRADA)

### Objetivo
Tener una base sólida y confiable para captura y edición de incidencias.

### Incluye
- UID único por incidencia (por día)
- Captura por excepción (no asistencias)
- BDIncidenciasLocal como fuente única
- Catálogo de incidencias normalizado
- Export automático del código VBA
- Documentación mínima viva

### Decisiones clave
- UID oficial incidencias:
  `LOC|NUMEMP|AÑO|MM|TIPO|PERIODO|DIA`
- UPSERT por día (editar no duplica)

### Estado
✅ Cerrada (2025-12-18)

---

## FASE 2 — DISEÑO DE PROCESO (GENERAL + CANCÚN)

### Objetivo
Homologar procesos sin romper a las locaciones actuales.

### Proceso CORE (todas las locaciones)
1. Nómina mantiene una BD central que se actualiza al cierre del periodo.
2. El gerente puede capturar incidencias desde el día 1 del periodo,
   aunque la BD central aún no esté actualizada.
3. La matriz se genera desde empleados activos disponibles.
4. El gerente captura **solo incidencias NO asistencias**.
5. Empleados faltantes se agregan como **temporales**.
6. Al actualizarse la BD Nómina:
   - se validan altas y bajas
   - se rellena automáticamente lo no capturado como:
     - asistencias
     - PD (domingo)
     - DF (festivo)
   - excepto días previos a alta y posteriores a baja.
7. Se validan incidencias y, si aplica, bonos.
8. El periodo queda listo para el pull al concentrado.

### Extensión Cancún (NO sistema paralelo)
Cancún agrega una **fuente adicional**:
- BD Reloj Checador (entradas y salidas)

Reglas:
- Si un empleado aparece en el reloj:
  - se considera activo operativamente
  - puede crearse como temporal si aún no existe en BD central.
- BDIncidenciasLocal puede prellenarse desde el reloj.
- El proceso humano y el entregable final **son los mismos**.

### Estado
🟡 Diseño validado (en curso)

---

## FASE 3 — PERIODOS Y CONTROL DE FLUJO

### Objetivo
Controlar cuándo se puede capturar, validar, calcular y enviar.

### tblPeriodos (por locación)
Campos base:
- LocCode
- Anio, Mes, TipoPeriodo, Periodo
- FechaIni, FechaFin
- CloseTS
- StatusPeriodo
- NominaDBReady
- LastPulledAt
- UpdatedAt, UpdatedBy

### StatusPeriodo (configurable por locación)
CORE (todas):
- CAPTURA
- ENVIADO
- CERRADO (automático por tiempo)

EXTENDIDO (Cancún):
- CAPTURA
- LISTO_PARA_CALCULO
- VALIDADO
- ENVIADO
- CERRADO

Reglas:
- CERRADO bloquea siempre.
- ENVIADO se marca cuando el concentrado hace el pull.
- Cancún puede tener pasos intermedios antes de ENVIADO.

### Estado
🟡 Pendiente de implementación

---

## FASE 4 — CHECADOR (EXTENSIÓN)

### Objetivo
Integrar fuentes automáticas de información cuando existan.

### Incluye
- Importación de BD reloj checador
- Generación de incidencias base (asistencias)
- Detección de empleados no presentes en BD central

### Alcance
- Inicialmente solo Cancún
- Otras locaciones pueden migrar si existe fuente equivalente

### Estado
🔴 Bloqueada (requiere Fase 3)

---

## FASE 5 — BONOS Y VALIDACIÓN

### Objetivo
Calcular y validar bonos solo cuando la información esté completa.

### Reglas
- Bonos solo se calculan cuando:
  - BD Nómina esté actualizada
  - incidencias estén completas
- El gerente valida resultados antes de ENVIADO.

### Pendiente
- Validar reglas específicas con RH (Juanita):
  - descansos mínimos/máximos
  - reglas duras de nómina
  - condiciones de baja

### Estado
🔴 Bloqueada (decisiones RH)

---

## FASE 6 — MASTER / CONCENTRADO

### Objetivo
Centralizar información sin intervención de locaciones.

### Flujo
1. El master recorre carpetas de locaciones.
2. Lee `BDIncidenciasLocal` y `tblPeriodos`.
3. Marca el periodo como ENVIADO.
4. Consolida información para nómina.

### Estado
🔴 Bloqueada (requiere Fase 3)

---

## FASE 7 — v2.0 (EVOLUCIÓN)

### Objetivo
Separar APP y DATA.

- Locaciones = solo data files
- Template único = app
- Un solo punto de actualización de macros

### Estado
🧊 Futuro (cuando v1 esté estable)

---

## REGLA DE ORO
- Un solo entregable.
- Un solo concentrado.
- Cancún es extensión, no excepción.
- Si algo no está en este roadmap, NO se programa.
