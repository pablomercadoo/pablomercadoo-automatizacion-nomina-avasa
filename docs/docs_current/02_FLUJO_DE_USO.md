## 3. Captura de incidencias (COMPLETO)

### Objetivo
Registrar incidencias del periodo activo y reflejarlas correctamente en:
- BDIncidencias_Local (fuente de verdad)
- Matriz del periodo (vista operativa)

---

### Formularios involucrados
- frmIncidencias (principal)
- frmAgregarIncidencias (alta específica)
- frmOpciones (configuración auxiliar)

---

### Flujo operativo

#### 3.1 Alta de incidencia
**Usuario:**
- Abre formulario
- Selecciona empleado, día y tipo de incidencia
- Guarda

**Sistema:**
- Valida periodo abierto
- Genera UID si no existe
- Asigna IDRegistro
- Inserta en BDIncidencias_Local
- Actualiza matriz

---

#### 3.2 Edición de incidencia
**Usuario:**
- Selecciona incidencia existente
- Modifica datos
- Guarda

**Sistema:**
- Actualiza registro existente
- NO duplica filas
- Refresca matriz

---

#### 3.3 Eliminación de incidencia
**Usuario:**
- Elimina incidencia

**Sistema:**
- Elimina o marca como eliminado (según implementación)
- Mantiene integridad de BD
- Refresca matriz

---

### Reglas de seguridad
- Si el periodo está cerrado:
  - No se permite alta / edición / eliminación
- No se permite capturar fuera del periodo seleccionado

---

### Convivencia con checador
- Registros manuales NO pueden ser pisados por checador
- Registros CHECADOR solo pueden ser modificados por checador

---

### Criterios de éxito
- La incidencia aparece correctamente en la matriz
- No se duplican registros
- El cierre de periodo bloquea correctamente la captura
