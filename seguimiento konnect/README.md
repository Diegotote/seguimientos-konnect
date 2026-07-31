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
