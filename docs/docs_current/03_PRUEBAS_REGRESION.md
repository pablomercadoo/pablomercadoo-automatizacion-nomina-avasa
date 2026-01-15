# 03 — Pruebas de regresión (obligatorias)

> Se corren **SIEMPRE** antes de publicar.

## Setup
- Usar copia **DEV** del archivo (nunca PROD).
- Usar un periodo de prueba (no mezclar con operación real).

## Checklist (8 pruebas mínimas)

### 1) Abrir archivo
- [ ] Abre sin error.
- [ ] No borra `MasterDBPath/EmployeeDBPath`.
- [ ] Abre `frmMenuPrincipal`.

### 2) Selección de periodo
- [ ] Permite seleccionar periodo válido.
- [ ] Construye `periodID` correcto (ej. `2026-01-Q2`).

### 3) Sync empleados
- [ ] Corre sync sin duplicar.
- [ ] Si falta ruta/config, falla con mensaje claro **y deja log**.

### 4) Matriz
- [ ] Genera/abre matriz del periodo.
- [ ] Lista empleados correctos.

### 5) Alta manual
- [ ] En `frmIncidencias`, guardar una incidencia.
- [ ] Se refleja en `BDIncidencias_Local`.

### 6) Editar/Borrar
- [ ] Editar una incidencia existente.
- [ ] Borrar una incidencia y validar que desaparece.

### 7) Checador (idempotencia)
- [ ] Precargar checador 1 vez.
- [ ] Precargar checador 2 veces → NO duplica registros.
- [ ] No pisa cambios manuales.

### 8) Cerrar periodo
- [ ] Cerrar periodo bloquea edición.
- [ ] Validar que UI/acciones respetan el bloqueo.

## Evidencia
- Adjuntar (o copiar) últimas 30 líneas del `_LOG`.

