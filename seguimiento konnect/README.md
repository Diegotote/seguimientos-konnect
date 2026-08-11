# KONNECT Seguimiento V20

Construida sobre la estructura visual y de despliegue de la V18/V19.

## Cambios V20

- Nancy se excluye únicamente cuando aparece en **Director Comercial**.
- Los prospectos referenciados por Nancy permanecen visibles en la sección **Referenciadas**.
- Pedro, Erika y Ericka continúan fuera del alcance cuando aparecen como Director Comercial.
- José Ernesto García y Carlos Alejandro Álvarez ya no aparecen como cierres prioritarios del mes.
- Nueva diapositiva visible: **Cierres de meses transcurridos**, alimentada con `Membresias_KONNECT_Ejecutivo.xlsx`.
- Se conservan el diseño, navegación, tablas, selector y presentación operativa.

## GitHub / Vercel

Sube el contenido de esta carpeta respetando exactamente su estructura. En Vercel usa:

- Framework Preset: Other
- Root Directory: `./`
- Output Directory: `dist`

El archivo `vercel.json` ya publica la carpeta `dist` sin instalar ni compilar dependencias.


## Ajuste V20.1

- Nancy continúa excluida únicamente cuando aparece como Directora Comercial.
- Los prospectos referenciados por Nancy sí se incluyen en Referenciados.
- Los prospectos referenciados por Erika o Ericka quedan excluidos.


## V21

- Se agregó una nueva diapositiva en la presentación operativa para mostrar actividad de comentarios de financieras en plataforma.
- Lado izquierdo: gráfica de pastel con las financieras que sí comentaron y su desglose.
- Lado derecho superior: tabla de cobertura por financiera con operaciones totales, operaciones comentadas, cobertura y comentarios.
- Lado derecho inferior: tabla de financieras sin comentarios y sus operaciones activas.


## V22
- Se reacomodó la diapositiva de actividad en plataforma: dona arriba y listado de ops. comentadas abajo.
- La tabla de cobertura ahora muestra únicamente: financiera, ops. totales, ops. comentadas y cobertura.
- Se eliminó la diapositiva “Integración · Motivos de detención”.
- La diapositiva de actividad en plataforma se colocó antes de la lámina final de agradecimiento.


## V23
- Actualizada la diapositiva comercial de cierres históricos con nueva lectura visual (gráfica vertical + cierres de agosto y julio).
- Integrada la nueva alta de agosto (Melany Mariel Torres Gudiño) en el histórico de membresías.


## V24
- La diapositiva comercial de cierres por mes ahora se alimenta automáticamente de la hoja `CIERRES 2026` del archivo operativo.
- Al actualizar Seguimiento Operativo, se actualizan la gráfica mensual, el listado del mes actual y el listado del mes anterior.
- El detalle muestra mes, director, financiera, broker, cliente y monto.


## V25
- Viabilidad y las demás etapas del pipeline se filtran por el mes calendario actual.
- En agosto ya no arrastra operaciones fechadas en julio.
- El comparativo toma exactamente el mes calendario anterior.
- El título y contador superior de dispersiones se actualizan dinámicamente.
- Se cambió la clave local de almacenamiento para evitar que sobrevivan datos viejos de julio.


## V26
- La diapositiva comercial “Cierres de membresías por mes” ahora lee exclusivamente la hoja CIERRES 2026 del archivo comercial.
- Reconoce bloques mensuales de marzo a agosto con encabezados tipo Titular, Oficina, Correo, Teléfono, Región y Membresía.
- Ya no reutiliza los cierres del archivo operativo ni los totales enero-marzo.
- Se renovó la clave de almacenamiento local para evitar datos heredados.


## V31
- Corrección de metas mensuales en inicio operativo.
- Se agregó/forzó bloque de metas posteriores con gap.
- La estructura por estatus considera fecha de comentario para capturar actividad del periodo actual.
- Corrección visual en gráfica de comentarios de financieras.


## V31.1
- Corrección de carga de JavaScript en Vercel: `dist/index.html` ahora usa XLSX por CDN y `/assets/updater-v20.js`, por lo que vuelve a funcionar el botón “Actualizar presentaciones”.
