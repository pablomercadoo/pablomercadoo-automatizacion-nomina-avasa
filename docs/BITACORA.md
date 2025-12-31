# 📒 Bitácora de trabajo — Automatización Incidencias AVASA

---

## 🗓️ 18 de diciembre de 2025 — Cierre de jornada

### Contexto
Sesión enfocada en **cerrar y documentar el CORE del sistema de incidencias**, con énfasis en arquitectura general y definición de alcance real del proyecto.

### ✅ Qué se logró
- Se cerró y documentó el **CORE del sistema de incidencias**.
- Se definió el **proceso general** y el **proceso específico de Cancún**.
- Se identificó que **Cancún es una extensión del proceso**, no un sistema independiente.
- Se actualizó el **ROADMAP** con una visión unificada y escalable.

### 🧠 Decisiones clave
- Un solo entregable final: `BDIncidencias_Local`.
- Proceso CORE común para todas las locaciones.
- Cancún agrega una fuente adicional (reloj checador).
- El sistema será **configurable por locación**, no duplicado.
- RH deberá validar reglas duras de incidencias (pendiente).

### 📌 Pendientes
- Validar reglas de incidencias con RH (Juanita).
- Bajar Fase 3 a diseño técnico (`tblPeriodos` y estados).
- Detallar integración del reloj checador.

### ▶️ Próximo paso
- Diseñar e implementar control de periodos (`tblPeriodos`).

---

## 🗓️ 22 de diciembre de 2025

### Contexto
Sesión enfocada en **estabilizar la v1 del sistema de incidencias**, cerrar bugs críticos y validar el flujo real de operación en la locación **CAP**, trabajando con datos reales y precarga desde checador.

### ✅ Avances logrados

#### 1. Matriz funcional end-to-end
- La matriz del periodo:
  - Se genera correctamente desde `Empleados`.
  - Se rellena con incidencias existentes desde `BDIncidencias_Local`.
  - Respeta overlay de datos (no borra incidencias manuales).
- Colores de domingos y festivos funcionando correctamente.
- Freeze panes correcto (filas 1–2 y columnas A–H).

#### 2. Botones de matriz (Agregar / Editar / Eliminar)
- Se corrigió error crítico que impedía abrir el formulario.
- El formulario abre correctamente en:
  - Agregar incidencia.
  - Editar incidencia (precarga correcta desde BD).
- Eliminar incidencias:
  - Borra registros en `BDIncidencias_Local`.
  - No elimina al empleado de la matriz (correcto por diseño).

#### 3. Precarga desde checador (robusta)
- Soporta:
  - Cargas parciales.
  - Cargas acumuladas.
  - Múltiples cargas por periodo.
- Regla crítica:
  - El checador solo pisa registros de checador.
  - Nunca pisa incidencias manuales.
- Uso de UID único por día evita duplicados.

#### 4. Flujo real validado
- Flujo probado:
  1. Precargar checador.
  2. Editar incidencias manuales.
  3. Volver a precargar.
  4. Regenerar matriz.
- Resultado:
  - Sin duplicados.
  - Sin pérdida de información.
  - Comportamiento consistente.

### 🧠 Decisiones de diseño
- La matriz siempre se genera desde `Empleados`.
- Eliminar incidencias no elimina empleados.
- La precarga desde checador se ejecuta desde el menú principal.
- No todas las locaciones tendrán checador:
  - Se define bandera `TieneChecador` en `tblLocaciones`.

### 📌 Pendientes
- Modos de carga por locación.
- Alta temporal de empleados.
- Definir estrategia de histórico y performance.

---

## 🗓️ 23 de diciembre de 2025 — Cierre v1 funcional (17:00 hrs)

### Contexto
Sesión enfocada en **blindar reglas de negocio, UX y consistencia de datos**, evitando agregar features nuevos.

### ✅ Trabajo completado

#### 1. Core de incidencias (cerrado)
- `BDIncidencias_Local` definida como **fuente única de verdad**.
- UID único por empleado + periodo + día funcionando.
- Flujo Agregar / Editar / Eliminar validado.
- La matriz siempre se regenera desde BD.
- En modo edición:
  - Se refresca la matriz.
  - El formulario se cierra automáticamente.

#### 2. Catálogo y normalización
- Catálogo canonizado y aliases resueltos.
- La lógica depende del código normalizado, no del texto.
- Incidencia **B (Baja)**:
  - Siempre al final.
  - Requiere confirmación explícita.

#### 3. Blindaje por tipo de día
- Domingos y festivos:
  - No se permite X.
  - Solo PD / DF / B u otras válidas.
- Días normales:
  - No se permite PD / DF.
- Blindaje aplicado:
  - Al cargar el formulario.
  - Antes de guardar.

#### 4. Formulario `frmIncidencias`
- Inicialización estable.
- Precarga desde BD funcionando.
- UX consistente y predecible.

#### 5. Precarga desde checador (CAP)
- Cargas parciales y acumuladas.
- Manual nunca se sobreescribe.
- Sin duplicados.
- Matriz se regenera correctamente.

👉 **CAP operable en producción controlada.**

### ⚠️ Pendientes (decisión consciente)
- Modos de carga por locación.
- Alta temporal de empleados.
- Estados del periodo (BORRADOR / ENVIADO / CERRADO).

---

## 🗓️ 30 de diciembre de 2025 — Sesión de rescate y reestructura (22:21 hrs)

### Contexto
Sesión intensiva para **rescatar, ordenar y unificar** el sistema previo al arranque operativo.
Se trabajó bajo presión real de fecha, priorizando **arquitectura, estabilidad y trazabilidad**.

### ✅ Avances logrados

#### 1. ETL / Base de empleados
- ETL corregido y estabilizado:
  - `OUT_EmpleadosMaster` se genera correctamente.
  - Normalización de **PuestoCanon** y **EsOperativo**.
  - Integración correcta de **UsuarioCARs / DriverCARs** desde `TI_RAW`.
- Decisión de **centralizar catálogos en el ETL**.
- Exportación a `.xlsx` funcional.

#### 2. Matriz de incidencias
- La matriz vuelve a mostrar incidencias correctamente.
- Número de empleado recuperado y bloqueado en edición.
- UsuarioCARs+ y DriverCARs+:
  - Correctos en la matriz (oficiales y temporales).
- Alta temporal funcional y homogénea en formato.

#### 3. UI / Reestructura
- Eliminación del enfoque de múltiples botones en hoja.
- Nuevo enfoque:
  - **Un solo botón “OPCIONES” en la matriz**.
- Creación de `frmOpciones` con:
  - Agregar
  - Editar
  - Limpiar incidencias
  - Eliminar empleado de la matriz
  - Cambiar periodo
  - Cerrar periodo (preparado, no activo)
- Labels de contexto:
  - `lblEmpleado`
  - `lblPeriodo`

#### 4. Seguridad de periodo
- Implementado `modSeguridadIncidencias`:
  - `SECURITY_ON`.
  - Cierre automático por fecha.
  - Override por status `CERRADO`.
- Infraestructura lista para cierre formal de periodos.

### ❌ Problema NO resuelto (documentado)
- **Editar incidencia abre el formulario pero no carga datos**:
  - Contexto correcto (locación, periodo, empleado).
  - Campos vacíos en el form.
- Confirmado:
  - No es problema de la matriz.
  - No es problema de selección.
  - Es un bug localizado en `CargarDesdeBD`.
- Se decidió **detener trabajo** para evitar deuda técnica por cansancio.

### 📌 Pendientes inmediatos
1. Auditar `CargarDesdeBD` con trazas controladas.
2. Corregir lectura desde `BDIncidencias_Local`.
3. Optimizar cierre / limpieza lenta del formulario (secundario).

### 📊 Estado al cierre
- ETL empleados: 🟢 OK
- Matriz periodo: 🟢 OK
- Alta temporal: 🟢 OK
- Guardado incidencias: 🟢 OK
- UI Opciones: 🟢 OK
- Editar incidencias: 🔴 Abre sin cargar

---
# 🧾 Bitácora técnica — 31/12

⏰ Hora de cierre: 13:00  
🎯 Estado general: **V1 funcional, estable y usable**  
🧠 Modo de trabajo: ejecución, sin refactors grandes

---

## ✅ Logros del día

### 🔧 Estabilidad general
- El proyecto **compila en verde** sin errores.
- Flujo completo operativo desde **frmOpciones**:
  - Agregar
  - Editar
  - Limpiar incidencias
  - Eliminar empleado de la matriz
  - Completar periodo (AUTO)
  - Cerrar periodo

### 📊 Matriz de incidencias
- Se consolidó el modelo:
  - **La matriz se genera SIEMPRE desde `Empleados`**
  - Las incidencias se leen exclusivamente desde `BDIncidencias_Local`
- Se corrigió definitivamente:
  - Puesto / Actividad (ya no aparecen como `1`)
  - UsuarioCARs / DriverCARs
- El botón único **OPCIONES** reemplaza todos los botones de hoja.

### 🧑‍💼 Empleados
- Empleados oficiales + temporales funcionan correctamente.
- Se implementó **eliminación por periodo**:
  - El empleado eliminado:
    - desaparece de la matriz
    - NO se borra de BD (queda respaldo)
    - NO se completa en AUTO
- El flujo ya distingue correctamente:
  - Oficial
  - Temporal
  - Eliminado por periodo

### 🧠 Completar periodo (AUTO)
- La macro **CompletarPeriodoActual**:
  - Inserta incidencias en **BD**, no solo en la matriz.
  - Recorre **solo empleados visibles** en la matriz.
  - Respeta:
    - manual vs AUTO
    - domingos (PD)
    - festivos (DF)
    - normales (X)
- Se integra con seguridad de periodo abierto/cerrado.

### 🔐 Seguridad
- Periodo cerrado:
  - Bloquea agregar / editar / limpiar / eliminar
  - Deja el sistema en **solo lectura**
- `modSeguridadIncidencias` ya gobierna toda la UI.

### 📦 Catálogos (decisión importante)
- Se **eliminan catálogos locales** de Puesto / Actividad.
- Los dropdowns se alimentan de:
  - **valores únicos globales** desde la BD del ETL
- Esto permite:
  - crear puestos nuevos en cualquier locación
  - sin romper reglas futuras

---

## 🧭 Decisiones importantes del día

- ✔️ La **fuente de verdad** son las BD, no las matrices.
- ✔️ El consolidado futuro se hará **desde BDIncidencias**, no desde hojas.
- ✔️ Eliminar empleado ≠ borrar BD (se marca por periodo).
- ✔️ V1 prioriza **operación real** sobre perfección visual.

---

## 📌 Estado al cierre

- Sistema **usable para gerentes**
- Flujo completo de captura y cierre
- Pendientes ya claramente acotados (ver `PENDIENTES.md`)

⛔ Se cierra sesión sin abrir nuevos frentes.


