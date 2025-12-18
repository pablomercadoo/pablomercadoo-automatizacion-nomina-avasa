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

