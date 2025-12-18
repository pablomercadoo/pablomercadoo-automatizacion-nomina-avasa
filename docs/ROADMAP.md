# ROADMAP — Automatización Incidencias AVASA

Este documento define el orden de trabajo del proyecto.
Si algo no está aquí, NO se toca.

---

## FASE 1 — CORE ESTABLE (cerrada)

### Objetivo
Tener un sistema base confiable para captura y edición de incidencias.

### Pasos
1. UID único por incidencia (por día)
2. UPSERT sin duplicados
3. Catálogo de incidencias normalizado
4. Export automático del código VBA
5. Documentación mínima (README, DECISION_LOG)

### Estado
✅ Cerrada (2025-12-18)

### No tocar en esta fase
- CAP
- Checador
- Bonos
- Festivos automáticos

---

## FASE 2 — DISEÑO CAP (sin código)

### Objetivo
Definir reglas claras para CAP antes de programar.

### Pasos
1. Definir reglas de asistencia (entrada + salida)
2. Definir traducción de días (PD / DF / X)
3. Definir prioridad: manual vs checador
4. Definir reglas de bonos (fijo 14 días, comedor)
5. Diseñar tablas necesarias (sin VBA)

### Estado
🟡 Pendiente

### No tocar en esta fase
- Programar checador
- Automatizar bonos
- Modificar matrices

---

## FASE 3 — IMPLEMENTACIÓN CAP

### Objetivo
Implementar lo diseñado en Fase 2.

### Pasos
1. Crear tablas CAP en Config
2. Importar datos de checador
3. Merge checador + incidencias
4. Cálculo de bonos
5. Validaciones y pruebas

### Estado
🔴 Bloqueada (requiere Fase 2)

---

## FASE 4 — CIERRES Y MANTENIMIENTO

### Objetivo
Asegurar integridad histórica del sistema.

### Pasos
1. Cierre automático por ventana (48h)
2. Limpieza segura de matrices
3. Navegación histórica
4. Exportación a nómina

### Estado
🔴 Bloqueada
