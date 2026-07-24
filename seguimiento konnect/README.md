# KONNECT Seguimiento V19

Versión construida sobre el repositorio V18 que ya desplegaba correctamente en Vercel.

## Cambios comerciales V19

- Excluye por completo cualquier registro que contenga Nancy, Pedro, Erika o Ericka en cualquier campo.
- Mantiene únicamente el alcance comercial válido.
- Separa operaciones de Diego y Jorge.
- Distingue cartera 100% propia y cartera referenciada.
- Muestra el origen de las referencias y la distribución por estatus.
- Conserva los cierres del mes actual y del mes siguiente.
- Mantiene intacta la presentación operativa y el diseño visual V18.

## Subida a GitHub

Reemplaza todo el contenido del repositorio con el contenido de esta carpeta.

Vercel debe usar:

- Framework Preset: Other
- Root Directory: ./

No agregues comandos manuales. `vercel.json` publica directamente la carpeta `dist`, igual que la V18.
