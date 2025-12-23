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
