# 🚑 RESCATE EXPRESS — V1
**Sistema de Incidencias AVASA**

📅 Fecha: 30 de diciembre  
🎯 Objetivo: Sistema **operativo y usable** el 1° de enero  
⏱️ Enfoque: **Funcionalidad > Elegancia > Refactor**

---

## 🧭 REGLAS DEL RESCATE (NO NEGOCIABLES)

- ❌ No refactors grandes
- ❌ No mejoras “nice to have”
- ❌ No cambios sin checklist
- ❌ No romper flujo existente que ya funciona
- ✅ Cambios pequeños, verificables
- ✅ Commit por bloque
- ✅ Todo lo que no esté aquí → **se ignora**

---

## 🧩 ESTADO ACTUAL DEL SISTEMA

### ✅ FUNCIONA
- Captura manual de incidencias
- Edición de incidencias existentes
- Validación de códigos por día (PD / DF / X)
- UID por día / empleado / periodo
- Guardado en `BDIncidencias_Local`
- Regeneración de matriz (aunque con errores de datos)

### ⚠️ FUNCIONA CON ERRORES
- Puesto y Actividad se escriben como `1`
- UsuarioCARs / DriverCARs no se cargan correctamente
- Matriz a veces no refleja cambios tras eliminar
- Mezcla de personal semanal y quincenal

### ❌ NO IMPLEMENTADO
- Cierre automático de periodo
- Bloqueo por periodo cerrado
- Menú único (UserForm) para acciones
- Diferenciación de tipo de nómina (semanal / quincenal)
- Flag `TieneChecador` por locación

---

## 🧪 BLOQUE 0 — ESTABILIZACIÓN (OBLIGATORIO)
⏱️ 20 min

- [ ] Confirmar que el sistema **abre sin errores**
- [ ] Confirmar que se puede:
  - Agregar incidencia
  - Editar incidencia
  - Guardar sin error
- [ ] **NO TOCAR LÓGICA**, solo asegurar punto de partida
- [ ] Commit: `chore: baseline stable before rescue`

---

## 🔧 BLOQUE 1 — CORRECCIÓN CRÍTICA (PUESTO / ACTIVIDAD)
⏱️ 40 min

### Objetivo
Eliminar definitivamente el error donde **Puesto / Actividad aparecen como `1`**.

### Checklist
- [ ] Localizar **exactamente** dónde se escriben en la matriz
- [ ] Verificar:
  - Tipo de dato (Value vs Value2)
  - Uso incorrecto de índices / booleanos
- [ ] Corregir escritura para que:
  - Sea texto
  - Respete catálogo canon
- [ ] Validar con 2 empleados reales
- [ ] Commit: `fix: correct puesto and actividad values in matrix`

---

## 🔧 BLOQUE 2 — CARGA CORRECTA DE BD EMPLEADOS
⏱️ 30 min

### Objetivo
Que **UsuarioCARs, DriverCARs, Puesto y Actividad** se carguen correctamente desde la BD.

### Checklist
- [ ] Revisar flujo ETL → `Base de datos empleados.xlsx`
- [ ] Confirmar que:
  - Las columnas existen
  - No hay corrimiento de índices
- [ ] Ajustar lectura **sin crear nuevas funciones**
- [ ] Validar:
  - Empleado con datos
  - Empleado sin CARs (campos vacíos)
- [ ] Commit: `fix: load cars, puesto, actividad from empleados DB`

---

## 🔁 BLOQUE 3 — ELIMINAR / LIMPIAR + REGENERAR MATRIZ
⏱️ 25 min

### Objetivo
Que **cualquier cambio** refleje la matriz **sin intervención manual**.

### Checklist
- [ ] Revisar:
  - `EliminarIncidenciasEmpleadoPeriodo`
- [ ] Al final:
  - Llamar SIEMPRE a `GenerarMatrizPeriodoActual`
- [ ] Probar:
  - Eliminar → matriz se actualiza
  - Limpiar → matriz se actualiza
- [ ] Commit: `fix: matrix always regenerates after delete/clean`

---

## ⚙️ BLOQUE 4 — DIFERENCIAR SEMANAL / QUINCENAL
⏱️ 30 min

### Objetivo
Que el sistema **NO mezcle empleados** con distinta forma de pago.

### Checklist
- [ ] Identificar columna tipo nómina en RH
- [ ] Al generar matriz:
  - Si periodo = semanal → solo empleados semanales
  - Si periodo = quincenal → solo empleados quincenales
- [ ] Sin excepciones
- [ ] Commit: `feat: filter employees by payroll type`

---

## 🧠 BLOQUE 5 — CIERRE AUTOMÁTICO DE PERIODO
⏱️ 35 min

### Objetivo
Permitir cerrar el periodo **sin capturar todo manualmente**.

### Checklist
- [ ] Crear macro `CompletarPeriodoActual`
- [ ] Para cada empleado visible:
  - Si no existe incidencia:
    - DF si festivo
    - PD si domingo
    - X si normal
- [ ] No pisar capturas manuales
- [ ] Marcar como `AUTO`
- [ ] Regenerar matriz
- [ ] Botón único: “Completar / Cerrar”
- [ ] Commit: `feat: auto-complete and close period`

---

## 🔒 BLOQUE 6 — BLOQUEO POR PERIODO CERRADO
⏱️ 15 min

### Objetivo
Periodo cerrado = **solo lectura**.

### Checklist
- [ ] Bloquear:
  - Agregar
  - Editar
  - Eliminar
  - Precarga
- [ ] Permitir:
  - Ver
  - Generar matriz
- [ ] Commit: `feat: lock system when period is closed`

---

## 🧭 BLOQUE 7 — MENÚ ÚNICO (USERFORM)
⏱️ 20 min

### Objetivo
Eliminar botones sueltos en hojas.

### Checklist
- [ ] Crear UserForm menú:
  - Agregar
  - Editar
  - Limpiar
  - Eliminar empleado
  - Completar periodo
- [ ] Conectar a macros existentes (NO duplicar lógica)
- [ ] Commit: `feat: unified menu userform`

---

## ✅ DEFINICIÓN DE “V1 TERMINADA”

- [ ] Sistema usable por un gerente sin soporte
- [ ] Matriz siempre consistente
- [ ] No aparecen valores `1`
- [ ] Empleados correctos por tipo de nómina
- [ ] Periodo se puede cerrar
- [ ] Repo limpio, con commits claros

---

## 🏁 NOTA FINAL

Cualquier idea nueva → **V2**  
Cualquier refactor → **V2**  
Cualquier mejora estética → **V2**

**V1 se cierra hoy.**
