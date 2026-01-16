# 02 — FLUJO DE USO DEL SISTEMA

## Usuarios
- Gerentes
- Auxiliares / asistentes
- Administrador (operador del sistema)

---

## Flujo estándar

### 1. Abrir archivo
**Sistema:**
- Ejecuta Workbook_Open
- Lee configuración
- Muestra menú
**Log esperado:** Workbook_Open START/END

---

### 2. Selección de periodo
**Usuario:**
- Selecciona locación, año, mes, tipo, periodo
- Presiona Aceptar
**Sistema:**
- Construye periodID
- Sincroniza empleados
- Genera matriz
- Navega a matriz
**Log esperado:** cmdAceptar_Click

---

### 3. Captura de incidencias
**Usuario:**
- Agrega / edita / elimina incidencias
**Sistema:**
- Actualiza BDIncidencias_Local
- Refleja cambios en matriz

---

### 4. Precarga checador (si aplica)
**Usuario:**
- Selecciona archivo checador
**Sistema:**
- Lee archivo
- Filtra por locación
- Inserta/actualiza solo registros checador
**Log esperado:** PrecargarChecador

---

### 5. Cierre de periodo
**Usuario:**
- Ejecuta cierre
**Sistema:**
- Bloquea edición
- Marca periodo cerrado
**Log esperado:** CerrarPeriodo

---

## Reglas operativas
- No se permite editar periodo cerrado
- El checador no pisa capturas manuales
- No deben mostrarse errores técnicos al usuario final
