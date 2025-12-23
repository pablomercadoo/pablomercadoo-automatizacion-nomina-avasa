## 2025-12-18 — Cierre de jornada

### Qué se logró
- Se cerró y documentó el CORE del sistema de incidencias.
- Se definió el proceso general y el proceso Cancún.
- Se identificó que Cancún es una EXTENSIÓN del proceso, no un sistema aparte.
- Se actualizó el ROADMAP con una visión unificada y escalable.

### Decisiones clave
- Un solo entregable final: BDIncidenciasLocal.
- Proceso CORE común para todas las locaciones.
- Cancún agrega una fuente adicional (reloj checador).
- El sistema será configurable por locación, no duplicado.
- RH deberá validar reglas duras de incidencias (pendiente).

### Pendientes
- Validar reglas de incidencias con RH (Juanita).
- Bajar Fase 3 a diseño técnico (tblPeriodos y estados).
- Detallar integración del reloj checador.

### Próximo paso
- Diseñar e implementar control de periodos (tblPeriodos).

# 📒 Bitácora de trabajo — Automatización Incidencias AVASA

---

## 🗓️ 22 de diciembre de 2025

### Contexto
Sesión enfocada en **estabilizar v1 del sistema de incidencias AVASA**, cerrar bugs críticos y validar el flujo real de operación en la locación **CAP**, trabajando ya con datos reales y precarga desde checador.

---

### ✅ Avances logrados

#### 1. Matriz funcional end-to-end
- La **matriz del periodo**:
  - Se genera correctamente desde `Empleados`.
  - Se rellena con incidencias existentes desde `BDIncidencias_Local`.
  - Respeta overlay de datos (no borra incidencias manuales).
- Colores de domingos y festivos **funcionando correctamente**.
- Freeze panes correcto (filas 1–2 y columnas A–H).

---

#### 2. Botones de matriz (Agregar / Editar / Eliminar)
- Se corrigió error crítico que impedía abrir el formulario:
  - El **form ahora abre correctamente** en:
    - Agregar incidencia
    - Editar incidencia (precarga correcta desde BD).
- Eliminar incidencias:
  - Elimina registros en `BDIncidencias_Local` correctamente.
  - La matriz **no elimina al empleado** (comportamiento correcto por diseño).
  - Al regenerar la matriz, el empleado aparece sin incidencias.

---

#### 3. Precarga desde checador (robusta)
- Se validó y dejó operativa la macro de precarga:
  - Permite **múltiples cargas dentro del mismo periodo**.
  - Soporta ambos escenarios:
    1. Archivos parciales (ej. días 16–18, luego 19–20).
    2. Archivos acumulados (ej. días 16–21).
- Regla aplicada:
  - El checador **solo pisa registros capturados por checador**.
  - Nunca pisa incidencias manuales.
- Uso de **UID único por día** evita duplicados y permite upsert seguro.

---

#### 4. Flujo real validado
- Se probó el flujo completo:
  1. Precargar checador.
  2. Editar incidencias manuales.
  3. Volver a precargar.
  4. Regenerar matriz.
- Resultado:
  - **Sin duplicados**.
  - **Sin pérdida de información**.
  - Comportamiento consistente y estable.

---

### 🧠 Decisiones de diseño tomadas

1. La matriz **siempre se genera desde `Empleados`**, no desde incidencias.
2. Eliminar incidencias **no elimina empleados** (correcto por diseño).
3. La precarga desde checador:
   - Se ejecuta desde el **menú principal**, no desde la matriz.
4. No todas las locaciones tendrán checador:
   - Se agrega bandera `TieneChecador` en `tblLocaciones`.

---

### 📝 Pendientes definidos (no implementados)

#### A. Modos de carga por locación
Para cada locación se debe definir:
- `TieneChecador = TRUE / FALSE`.

Flujos a implementar:
1. Precarga desde checador (si aplica).
2. Captura manual por formulario (siempre disponible).
3. Alta temporal de empleado.

---

#### B. Alta temporal de empleado (pendiente)
- Usar **el mismo formulario** de incidencias.
- Flujo propuesto:
  - Si el número de empleado no existe en `Empleados`:
    - Preguntar si desea agregarlo temporalmente.
    - Habilitar campos superiores (nombre, puesto, etc.).
    - Guardar en una tabla temporal por periodo (`Empleados_Temp`).
- La matriz deberá:
  - Incluir empleados oficiales + empleados temporales del periodo.

---

#### C. Histórico y performance (pendiente)
- Definir estrategia para:
  - Manejo de históricos de incidencias cerradas.
  - Evitar crecimiento excesivo del archivo en el tiempo.
- Posible solución futura:
  - Migrar periodos cerrados a una BD histórica.
  - Limpiar BD activa.
- Este punto se considera **ajuste final** y no bloquea la v1.

---

### 📌 Estado actual del proyecto

- **Versión:** v1 funcional (operativa en CAP).
- **Riesgos críticos:** mitigados.
- **Siguiente sesión:**  
  Implementar **modos de carga por locación** y **alta temporal de empleado**.

---
# Bitácora — 23 de diciembre de 2025 (5:00 pm)

## Contexto
Sesión enfocada en **estabilizar y cerrar la v1 funcional** del sistema de incidencias.
El objetivo no fue agregar features nuevos, sino **blindar reglas de negocio, UX y consistencia de datos**.

---

## ✅ Trabajo completado

### 1. Core de incidencias (cerrado)
- `BDIncidencias_Local` definida como **fuente única de verdad**.
- UID único por **empleado + periodo + día** funcionando correctamente.
- Flujo **Agregar / Editar / Eliminar** validado:
  - Editar sobrescribe por UID.
  - Eliminar borra de BD y limpia la matriz al regenerar.
- La matriz **siempre se regenera desde BD**, no se edita manualmente.
- Al guardar en modo edición:
  - Se refresca la matriz.
  - El formulario se cierra automáticamente.

---

### 2. Catálogo de incidencias y normalización
- Catálogo activo y normalizado validado.
- Aliases resueltos (ej. `T/D → TD`, `FI → F`, etc.).
- La lógica ya **no depende del texto capturado**, sino del código canonizado.
- Incidencia **B (Baja)**:
  - Siempre aparece al final de la lista.
  - Requiere confirmación explícita al guardar.

---

### 3. Reglas por tipo de día (blindaje completo)
Se implementó un sistema **a prueba de errores humanos**:

#### Domingos (PD) y días feriados (DF)
- ❌ Se elimina la opción **X (Asistencia)**.
- ✅ Solo se permiten:
  - PD / DF (según aplique)
  - B (Baja)
  - Otras incidencias válidas (vacaciones, incapacidades, descansos, etc.).

#### Días normales
- ❌ No se permite seleccionar PD ni DF.
- Si vienen cargados desde BD:
  - Se corrigen automáticamente (PD/DF → X o vacío).
- Todas las demás incidencias son válidas.

#### Blindaje doble
- Las reglas se aplican:
  - Al cargar el formulario.
  - Antes de guardar (blindaje final).
- Aunque el usuario intente forzar un valor, **el sistema lo corrige**.

---

### 4. Formulario `frmIncidencias`
- Inicialización estable.
- Precarga desde BD funcionando correctamente.
- Reglas de combos por día se reaplican siempre.
- No guarda estados inválidos.
- UX consistente y predecible.

---

### 5. Precarga desde checador (CAP)
- Soporta cargas:
  - Parciales.
  - Acumuladas.
  - Múltiples veces por periodo.
- Regla crítica cumplida:
  - **Checador solo pisa registros de checador**.
  - Manual nunca se sobreescribe.
- Sin duplicados ni pérdida de información.
- Matriz se regenera correctamente tras cada carga.

👉 **CAP puede operar en producción controlada.**

---

## ⚠️ Pendientes identificados (no implementados)

### 1. Modos de carga por locación
- Falta agregar campo `TieneChecador` en `tblLocaciones`.
- El botón **Agregar** aún no pregunta:
  - Manual
  - Precarga desde checador
  - Alta temporal

---

### 2. Alta temporal de empleados
- No existe aún `Empleados_Temp`.
- El formulario requiere que el empleado exista en `Empleados`.
- Falta el flujo:
  - Empleado no existe → alta temporal por periodo.
- La matriz aún no hace UNION con empleados temporales.

---

### 3. Estados del periodo (decisión consciente)
No se trabajó en:
- Estados BORRADOR / ENVIADO / CERRADO.
- Bloqueo real del libro.
- Archivado histórico.

(Se decidió conscientemente **no tocar esto en esta sesión**).

---

## Estado final de la sesión
- ✅ v1 del sistema **estable, consistente y usable**.
- 🧠 Reglas de negocio críticas correctamente implementadas.
- 🟡 Features estructurales grandes (modos de carga y alta temporal) **diferidas** para evitar sobrecarga.

---

## Próximo paso sugerido
Cuando se retome el proyecto:
1. Definir `tblLocaciones.TieneChecador`.
2. Selector de modos en botón **Agregar**.
3. Implementar `Empleados_Temp`.
4. UNION de empleados base + temporales en matriz.

---

**Sesión cerrada a las 17:00 hrs.**
