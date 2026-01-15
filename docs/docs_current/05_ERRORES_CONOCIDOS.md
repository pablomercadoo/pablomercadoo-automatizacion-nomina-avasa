# 05 — Errores conocidos

## E001 — (ejemplo) Config se borra al abrir
- **Síntoma:** al abrir, rutas quedan vacías, sync falla.
- **Causa típica:** shadowing de `AutoFix_ConfigPaths`.
- **Solución:** asegurar que solo exista una implementación y no borre keys.

## E002 — (ejemplo) Duplicados en precarga
- **Síntoma:** al precargar dos veces, duplica.
- **Solución:** usar UID como clave y hacer UPSERT.

