## UID único para incidencias

**Decisión**
El UID oficial del sistema para incidencias será:

LOC|NUMEMP|AÑO|MM|TIPO_PERIODO|PERIODO|DIA

**Motivo**
- Permite upsert por día
- Evita duplicados
- No depende de timestamp

**Implementación**
- BuildUID vive en modUID
- frmIncidencias NO define su propia versión

Decisión: UID oficial incidencias = LOC|NUMEMP|AÑO|MM|TIPO|PERIODO|DIA

Cambio: frmIncidencias ya no define BuildUID; ahora usa modUID.BuildUID_Incidencia

Riesgo evitado: 2 BuildUID con formatos distintos

Prueba: “editar misma incidencia mismo día no duplica”
