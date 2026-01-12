## 📍 MAPA DEL PROYECTO – Automatización Incidencias AVASA

### 🔹 ThisWorkbook

* **Qué hace**:
  Arranque y cierre del sistema.
* **Responsabilidades**:

  * Lee configuración inicial (locación, template).
  * Inicializa variables globales.
  * Muestra el menú principal al abrir.
  * Limpia matrices relevantes al cerrar.
* **Punto crítico**: si aquí falla algo, **el sistema ni siquiera arranca**.

---

### 🔹 modGlobal

* **Qué hace**:
  Guarda el **estado global** del sistema.
* **Variables clave**:

  * Año, mes, tipo de periodo, número de periodo
  * Locación (código y display)
  * Bandera de template
* **Regla mental**:
  Todo el sistema asume que estas variables están bien seteadas.

---

### 🔹 frmMenuPrincipal

* **Qué hace**:
  Es la **puerta de entrada del usuario**.
* **Responsabilidades**:

  * Elegir año, mes, tipo y periodo.
  * Validar que el periodo sea lógico (no futuro).
  * Sincronizar empleados del periodo.
  * Disparar la generación de la matriz.
* **Regla clave**:
  Aquí se define **qué periodo estás tocando**.

---

### 🔹 modReporteIncidencias

* **Qué hace**:
  Es el **motor principal del sistema**.
* **Responsabilidades**:

  * Crear o recuperar la hoja matriz del periodo (`M_LOC_AAAA_MM_Q#/S#`)
  * Pintar encabezados, días, columnas extra
  * Cargar empleados desde la hoja `Empleados`
  * Hacer overlay de incidencias desde `BDIncidencias_Local`
  * Crear botones (Agregar / Editar / Eliminar / Menú)
* **Regla clave**:

  * La matriz **siempre se reconstruye**, nunca se edita “a mano”.

---

### 🔹 frmIncidencias

* **Qué hace**:
  UI para **capturar o editar incidencias** de un empleado.
* **Responsabilidades**:

  * Cargar datos del empleado.
  * Mostrar días válidos del periodo.
  * Validar códigos de incidencia contra catálogo.
  * Guardar incidencias en `BDIncidencias_Local` (UPSERT por UID).
* **Detalle importante**:

  * Usa UID (`LOC|EMP|AÑO|MES|TIPO|PERIODO|DIA`) para evitar duplicados.
  * Maneja modo **nuevo vs edición**.

---

### 🔹 modSeguridadIncidencias

* **Qué hace**:
  Controla **cuándo un periodo se puede editar**.
* **Responsabilidades**:

  * Definir cierre automático del periodo.
  * Bloquear edición si el periodo ya cerró.
  * Proteger hojas de matriz.
* **Concepto clave**:

  * El cierre depende de la fecha fin del periodo + ventana de horas (Config).

---

### 🔹 modConfig

* **Qué hace**:
  Acceso centralizado a la hoja `Config`.
* **Responsabilidades**:

  * Leer valores por clave (`GetConfig`)
  * Evitar valores hardcodeados en el sistema.
* **Regla**:

  * Cualquier parámetro “de negocio” debería vivir aquí.

---

### 🔹 modAdmin

* **Qué hace**:
  Herramientas **administrativas / soporte**.
* **Responsabilidades**:

  * Buscar matrices de periodos pasados.
  * Navegar entre hojas históricas.
* **Uso típico**:

  * Auditoría
  * Soporte
  * Consultas históricas

---

### 🔹 Hojas del libro (Document / HojaX)

* **Qué hacen**:

  * La mayoría no tiene lógica directa.
  * Algunas solo existen como contenedores visibles/ocultos.
* **Regla**:

  * No meter lógica aquí salvo que sea estrictamente UI.

---

## 🧠 Regla mental final (importantísima)

* **BDIncidencias_Local = verdad**
* **Matriz = vista temporal**
* **Forms = UI**
* **Globals = estado**
* **Config = reglas del negocio**



---

### Flujo de empleados (actual)

Nómina (SEM / QUIN)
↓
ETL Empleados
↓
Master: Base de datos empleados.xlsx
  - Incluye: TipoNomina
↓
Archivo de locación
  - Hoja: Empleados
↓
Matriz de incidencias
  - Filtra por:
    - Locación
    - TipoNomina = TipoPeriodo
   


---

### 🔹 Generador de archivos por locación (Distribución V1)

* **Qué hace**:
  Genera automáticamente los **62 archivos** (1 por locación) desde el template.

* **Responsabilidades**:
  * Lee `tblLocaciones` (solo `Active = 1`)
  * Crea carpetas estándar por locación
  * Genera `Incidencias_<LOC>.xlsm` dentro de:
    `<RAIZ>\<LOC>\REPORTE DE INCIDENCIAS DE NOMINA\`
  * Setea valores en `Config` de cada archivo nuevo:
    `LocationCode`, `LocationName`, `LocationDisplay`, `CC`, `IsTemplate=0`, etc.

* **Regla clave**:
  El generador **solo se corre desde el template** (archivo maestro).

