# Tablero de Tableros PowerBI

Página simple para que el directorio acceda rápido a todos los reportes de PowerBI,
sin navegar por el portal. Agrupa por unidad de negocio, con logos y buscador.

## La fuente de la verdad: `LINKS POWER BI.xlsx`

La página **lee este mismo Excel** que ya venías usando. Si agregás, borrás o cambiás
una fila, el tablero se actualiza solo (no hay que tocar código).

Columnas que usa (hoja 1):

| Columna      | Para qué                                                              |
|--------------|----------------------------------------------------------------------|
| **Agrupacion** | Unidad de negocio que agrupa las tarjetas: `HERGO - DISTRIBUCION`, `MENOR COSTE`, `FINANZAS`. |
| **Nombre**     | Título del reporte.                                                 |
| **Link**       | URL del tablero de PowerBI.                                         |

### Informes mensuales (automático)
Si el nombre termina en ` MM-AAAA` (ej. `COMPARATIVO 06-2026`, `HECTOLITROS 05-2026`),
la página los **agrupa solos** en una sola tarjeta con un **menú desplegable de meses**,
mostrando por defecto el más reciente. Para publicar un mes nuevo solo agregás la fila
`... 07-2026` y aparece arriba automáticamente.

> Si agregás una agrupación nueva (ej. `BLACK`), aparece igual con un estilo neutro.
> Para darle color y logo propios, editá el objeto `GRUPOS` en `index.html` (está comentado).

## Archivos de esta carpeta

| Archivo                 | Para qué                                              |
|-------------------------|-------------------------------------------------------|
| `index.html`            | La página. No hace falta editarla para uso normal.   |
| `LINKS POWER BI.xlsx`   | **La lista de tableros.** Acá cargás todo.            |
| `assets/logos/`         | Logos de las unidades de negocio (van a Vercel).     |

## Probar localmente

Abrir el `.html` con doble clic NO funciona (el navegador bloquea leer el Excel desde
`file://`). Desde esta carpeta, en PowerShell:

```powershell
python -m http.server 8000
```

Y entrá a http://localhost:8000

## Publicar en Vercel

### Opción A — Rápida (arrastrar la carpeta)
1. https://vercel.com → **Add New… → Project → Deploy**.
2. Arrastrá toda esta carpeta (con el Excel y `assets/`).
3. Vercel te da una URL pública.
4. Para actualizar: volvés a subir la carpeta con el Excel nuevo.

### Opción B — Automática (recomendada, con Git)
Cada `push` redepliega solo cuando cambiás el Excel.
1. Subí esta carpeta a un repo (GitHub/GitLab).
2. Vercel → **Add New… → Project → Import** ese repo. Framework: **Other**.
3. Cuando actualices reportes:
   ```powershell
   git add "LINKS POWER BI.xlsx"
   git commit -m "Actualizo tableros"
   git push
   ```

## Notas
- Usa la librería **SheetJS** desde CDN para leer el Excel (requiere internet, igual que PowerBI).
- Tiene buscador por nombre o agrupación y es responsive (PC, tablet, celular).
- Logos tomados de `...\Uso Corporativo\` (Hergo, Alta=Distribución, Menor Coste, Black).
