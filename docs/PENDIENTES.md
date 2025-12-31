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
