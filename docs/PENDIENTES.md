# 📌 Pendientes — Post V1 (estado real)

Fecha: 31 de diciembre  
Estado: **pendientes conscientes, no bloqueantes**

---

## 🔴 Pendientes funcionales (importantes)

### 1) Completar (AUTO) con empleados que ya tienen incidencias
- Caso a revisar:
  - Empleado con **algunas incidencias manuales**
  - AUTO no siempre completa correctamente los días faltantes
- Objetivo:
  - Por cada (empleado, día):
    - si NO existe en BD → insertar AUTO
    - si YA existe → no tocar
- Impacto: medio (no bloquea operación diaria, pero sí cierre limpio)

---

### 2) Definición final del UID
- Falta decidir:
  - ¿Se guarda UID en `BDIncidencias_Local`?
  - ¿O se calcula solo en el consolidado master?
- UID propuesto:
LOC|AÑO|MES|TIPO|PERIODO|NUMEMP|FECHA

yaml
Copy code
- Impacto: alto para consolidación, bajo para operación local

---

### 3) Consolidación de los 62 reportes (Master)
- Pendiente crear archivo **MASTER** que:
- Recorra carpeta de locaciones
- Abra cada `Incidencias_XXX.xlsm` en read-only
- Lea `BDIncidencias_Local`
- Unifique todo en una sola tabla
- Este paso **no bloquea V1**, pero es clave para V1.5

---

## 🟡 Pendientes de UX / Formato

### 4) Ajuste visual de la matriz
- Ocultar columnas **Locación** y **Ciudad** en la matriz
- Mantenerlas en BD (necesarias para consolidado)
- Reajustar:
- anchos de columnas
- freeze panes
- posición visual (menos desplazamiento horizontal)

---

### 5) Automatización de Bono Comedor (CAP)
- Definir regla exacta:
- basada en asistencias X
- o en días trabajados
- Integrar al flujo de **CompletarPeriodoActual**
- Actualmente:
- la columna existe
- no está automatizada

---

## 🟢 Pendientes menores / futuros (V2)

- Diferenciar empleados **semanales vs quincenales** desde RH
- Automatizar cierre por fecha (CloseTS)
- Separar App / Data (template único)
- Auditoría visual de AUTO vs manual
- Mejora estética general

---

## ✅ Lo que NO es pendiente
- Menú único (frmOpciones)
- Eliminación de empleado
- Dropdowns de puesto/actividad
- Seguridad por periodo
- Regeneración de matriz
- Flujo completo de captura

---

🧠 Regla:  
Todo lo anterior **ya no bloquea operación**.  
Lo pendiente se ataca con cabeza fría en enero.


- Ejecutar pruebas funcionales SEM / QUIN en matriz
- Validar conteo correcto de empleados por tipo de nómina
- Confirmar comportamiento con empleados TEMP
- Integrar futuras bases de Usuarios / Drivers (pendiente Mauricio)
