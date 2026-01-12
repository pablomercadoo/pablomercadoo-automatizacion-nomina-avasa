# Arquitectura (V1)

## Componentes
1) **Template / Generador**: contiene VBA, formularios, Config, catÃ¡logo locaciones.
2) **Archivos por locaciÃ³n (62)**: Incidencias_<LOC>.xlsm
3) **MasterData externo**: Base de datos empleados.xlsx con 	blEmpleados

## Config (tblConfig) â€“ claves tÃ­picas
- MasterDBPath (raÃ­z REPORTES GERENTES)
- EmployeeDBPath (ruta al xlsx)
- EmployeeDBTable (tblEmpleados)
- LocationCode / CC
- IsTemplate
- LockWindowHours (48)
- LockPassword (AVASA)