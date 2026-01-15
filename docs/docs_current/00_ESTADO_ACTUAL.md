# 00 — Estado actual (fuente única)

Este documento es el **punto de verdad** para saber *qué versión corre*, *dónde vive* y *qué se publicó*.

## Versión actual
- **Nombre de versión:** `PROD_2026-01-15_0905`
- **Fecha/hora:** 2026-01-15 09:05 (America/Cancun)
- **Fuente:** `VBA_EXPORT_20260115_0905.zip` + archivo(s) Excel de plantilla/locación

## Artefactos
- **Export VBA:** `VBA_EXPORT_20260115_0905.zip`
- **VBA consolidado:** `VBA_TODO.txt` (si aplica)
- **Release notes:** ver `docs_current/04_BITACORA_CAMBIOS.md`

## Reglas de operación
- **PROD no se edita.**
- Todo cambio se hace en una copia **DEV** y solo se publica cuando pasa `03_PRUEBAS_REGRESION.md`.

## Objetivo inmediato (depuración)
- Depurar parches acumulados sin romper operación.
- Hacer el sistema **diagnosticable** (logs + healthcheck).

