## 2025-12-18 — UID único para incidencias + normalización de catálogo

**Problema**
- Existían 2 implementaciones de UID (riesgo de duplicados / comportamiento inconsistente).
- Error “código R no es válido” pese a estar activo (catálogo sin columna Normalizado poblada).

**Decisiones**
1) **UID oficial de incidencias (por día):**  
   `LOC|NUMEMP|AÑO|MM|TIPO|PERIODO|DIA`

2) **Normalización de catálogo:**  
   Todo código activo debe tener un valor en **Normalizado** (aliases apuntan al código canon).

**Motivo**
- Permite UPSERT por día (editar no duplica).
- Evita dependencias de timestamp/fecha completa.
- La validación se vuelve determinista (catálogo canonizado).

**Cambios**
- `frmIncidencias` ya no define `BuildUID` (se eliminó la versión duplicada).
- `frmIncidencias` ahora usa `modUID.BuildUID_Incidencia`.
- Se llenó la columna **Normalizado** en `tblCatalogoIncidencias`:
  - códigos canónicos → se normalizan a sí mismos (R→R, X→X, etc.)
  - aliases → apuntan al canon (T/D→TD, FI→F, etc.)

**Pruebas**
- Editar la misma incidencia en el mismo día **no duplica** (UPSERT OK).
- Códigos `R`, `TD` y alias `T/D` validan correctamente.
