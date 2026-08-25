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

---

## Automatización de los links de Power BI (`sincronizar.bat`)

Los links de Power BI **ya no se pegan a mano**. Un script los trae de la API y los
escribe en el Excel. Los links cargados a mano (Vercel, portales, apps) no se tocan.

### El Excel ahora tiene dos hojas

| Hoja | Para qué |
|---|---|
| **`Tableros`** | Lo que lee la página. Igual que antes, más una columna **`Origen`**: `MANUAL` (lo cargás vos) o `API` (lo genera el script). |
| **`Catalogo`** | **Todos** los reportes que existen en las áreas de trabajo de Power BI, con una columna **`Publicar`**. Solo los `SI` bajan a `Tableros`. |

Por qué el `Catalogo`: entre las 5 áreas de trabajo hay **219 reportes** y vos publicás
unos 45. Volcarlos todos taparía el tablero. Así los ves listados, tildás lo que
querés, y nunca más buscás un `reportId`.

### Cómo se usa

1. Doble clic en **`sincronizar.bat`**. Escribe `LINKS POWER BI _test.xlsx`.
2. El script informa qué apareció nuevo desde la última corrida.
3. Abrís la hoja `Catalogo` (lo nuevo viene resaltado en amarillo) y en `Publicar`
   ponés `SI` a lo que quieras mostrar, o `NO` a lo que no quieras que vuelva a
   proponerse. Vacío = no se publica.
4. Volvés a correr `sincronizar.bat`. Lo marcado baja a `Tableros`.
5. Mirás cómo quedó abriendo `index_test.html` (lee el `_test`, no toca producción).
6. Cuando está bien, pasás a producción:
   ```powershell
   python sincronizar_powerbi.py --entrada "LINKS POWER BI _test.xlsx" --salida "LINKS POWER BI.xlsx"
   ```
   y después `actualizar.bat` para publicar en Vercel.

Para ver qué haría sin escribir nada: `python sincronizar_powerbi.py --dry-run`

### Actualizar los links: qué hacés en cada caso

| Lo que pasó | Qué tenés que hacer |
|---|---|
| **Publicaste un informe nuevo** en cualquier área de trabajo | Nada. Doble clic en `sincronizar.bat` y aparece. |
| **Subiste el mes nuevo** de un informe mensual | Nada. El `.bat` lo detecta y reemplaza al mes más viejo. |
| **Borraste un informe** de Power BI Service | Nada. El `.bat` lo saca del tablero y te lo avisa. |
| **Renombraste un informe** en Power BI | Nada. El nombre del tablero se guarda por `reportId`, no cambia. |
| **Movés un informe** a otra área de trabajo | Nada. Pasa a la `Area` nueva. |
| **No querés que un informe se vea** | En la hoja `Catalogo`, poné `NO` en `Publicar`. |
| **Querés cambiar el nombre, la agrupación o la descripción** | Editalo en la hoja `Catalogo`. Queda para siempre. |
| **Sumás un link a mano** (Vercel, portal, app) | Fila nueva en la hoja `Tableros` con `Origen = MANUAL`. |
| **Cambia la URL de una app replicada** | Editás **una sola** fila; el script actualiza las 5 copias. |
| **Sumás un área de trabajo de Power BI** | Una línea en `areas` dentro de `powerbi_areas.json`. |

El ciclo completo es siempre el mismo:

```powershell
sincronizar.bat        # trae Power BI al Excel de prueba
                       # (mirás index_test.html si querés revisar)
actualizar.bat         # commit + push -> Vercel republica
```

### Si borro un informe de Power BI, ¿se borra del tablero?

**Sí, solo.** Sale de la hoja `Tableros` y también del `Catalogo`, y el script te lo
lista en el resumen para que sepas qué se fue. Está probado: se metió un informe
inexistente en el Excel a propósito y el script lo eliminó de las dos hojas.

Los links **`MANUAL`** no pasan por ese control — son tuyos y no viven en Power BI,
así que los de Vercel nunca se borran solos.

Si un informe borrado te ensucia el resumen todas las corridas, poné su `reportId`
en `descartar_reportid` dentro de `powerbi_areas.json` y deja de avisarte.

### Lo que el script respeta y no pisa

- **Links a mano** (`Origen = MANUAL`): intactos, siempre.
- **`Nombre`, `Agrupacion` y `Descripcion`**: son tuyos. Se guardan por `reportId`,
  así que si en Power BI el reporte se renombra, tu nombre queda. Hace falta porque
  los nombres no coinciden: Power BI dice `Consumos por cliente` y el tablero dice
  `CONSUMOS DE CLIENTES`, Power BI dice `MARGENES` y el tablero `MARGENES MAYORISTA`.
- **La `Area` que elegiste**: gana sobre el área de trabajo de origen. `COMPARATIVO
  08-2026` vive en *Grupo Hergo - Jefes* pero lo publicás en `DIRECTORIO`, y así queda.
- **Links que desaparecieron de Power BI**: no se borran. Pasan a `MANUAL` y el
  script te los lista para que decidas.

### Informes mensuales: el renombre es automático

Power BI los llama `COMPARATIVO JUNIO 2026` o `Informe Hctos - Junio 26`. La página
necesita el formato ` MM-AAAA` para colapsarlos en una tarjeta con desplegable de
meses. El script traduce solo:

| Power BI | Publicado |
|---|---|
| `COMPARATIVO JUNIO 2026` | `COMPARATIVO 06-2026` |
| `Informe Hctos - Junio 26` | `HECTOLITROS 06-2026` |

Y **un mes nuevo de una serie que ya publicás se marca `SI` solo**: cuando subas
`COMPARATIVO SEPTIEMBRE 2026`, corrés el `.bat` y ya está arriba. Cero trabajo manual.

Al tablero bajan solo los últimos **`meses_max`** meses (hoy 2: actual y anterior).
Los meses viejos quedan en `SI` en el `Catalogo` — no se pierde el tilde — pero no
se publican. Si algún día querés ver un año entero, subís `meses_max` a `12` y
aparecen todos, sin tocar nada más.

**Si el mes no está en el nombre, se busca en la subcarpeta.** Power BI organiza los
workspaces en carpetas (`Agosto - 26`, `Mayo 2026`); el script las lee por la API de
Fabric. Así que si algún mes subís el informe sin ponerle el mes al nombre pero lo
guardás en la carpeta correcta, igual lo detecta. Hoy los 106 informes que están en
carpetas mensuales ya tienen el mes en el nombre, así que esto es solo una red.

**Un informe que está en dos áreas de trabajo se publica una vez.** `COMPARATIVO
AGOSTO 2026` existe en *Directorio* y en *Grupo Hergo - Jefes*; si los dos caen en la
misma `Area`, gana el del área de trabajo que le corresponde y el script te avisa
cuál descartó.

### `powerbi_areas.json`

| Clave | Para qué |
|---|---|
| `areas` | Área de trabajo de Power BI → columna `Area`. Para sumar un área nueva, agregá una línea acá. |
| `agrupacion_default` | Qué `Agrupacion` se le pone a un reporte recién descubierto. Después la corregís en el Excel y queda. |
| `alias_series` | Nombre base de una serie mensual → cómo se llama en el tablero (`informe hctos` → `HECTOLITROS`). |
| `meses_max` | Cuántos meses de cada serie bajan al tablero. Está en **`2`** (mes actual y anterior) para que el desplegable no se sature. `0` = todos. Se aplica al publicar, no al tildar: podés cambiarlo cuando quieras sin tener que re-tildar nada en el `Catalogo`. |
| `ignorar` | Reportes que nunca entran al `Catalogo` (por ahora los `Usage Metrics Report`). |
| `publicar_nuevos` | `SI` = el Excel replica las áreas de trabajo: todo lo que existe se publica y el `Catalogo` sirve para excluir con `NO`. Vacío = al revés, hay que tildar cada uno. |
| `agrupacion_por_nombre` | Si el nombre contiene ese texto, va a esa agrupación (`MARGENES MC` → MENOR COSTE). Se evalúa antes que `agrupacion_default`. |
| `apps_en_todas_las_areas` | Links a mano que se replican en **todas** las áreas de trabajo. Los cargás una vez y el script hace las copias. |
| `agrupacion_apps_replicadas` | En qué agrupación caen esas copias (`HERRAMIENTAS`), para que queden juntas y al final de cada área. |
| `descartar_reportid` | Informes borrados de Power BI que ya no querés ver avisados en el resumen. |

### Autenticación

Reutiliza la sesión de Power BI del **orquestador**
(`Scripts\orquestador\.pbi_token.json`, device-code con refresh token). No hay
secretos en esta carpeta ni en el repo. Si el script dice que no hay sesión válida,
abrí el orquestador y volvé a loguear con el botón de Power BI.

Ve solo las áreas de trabajo donde tu usuario es miembro — que son justamente estas 5.
