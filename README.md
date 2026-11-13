# Tesis de Arquitectura CUJAE

Sitio web del registro de tesis de la Facultad de Arquitectura de la CUJAE: buscador con filtros, estadísticas y una página que explica cómo se hizo el registro y qué errores puede tener.

Es un sitio estático: no necesita servidor, base de datos ni instalar nada. **Todos los datos salen del Excel `datos/registro.xlsx`**; el sitio lo lee directamente en el navegador cada vez que alguien entra.

## Publicarlo en GitHub (una sola vez)

1. Crea una cuenta en github.com si no la tienes y pulsa **New repository**. Ponle un nombre (por ejemplo `tesis-arquitectura-cujae`), déjalo **Public** y créalo.
2. En el repositorio vacío, pulsa **uploading an existing file** y arrastra **todo el contenido de esta carpeta** (el archivo `index.html` y las carpetas `css`, `js`, `fuentes` y `datos`). Pulsa **Commit changes**.
3. Ve a **Settings → Pages**. En *Build and deployment*, elige **Deploy from a branch**, rama **main**, carpeta **/ (root)**, y pulsa **Save**.
4. Espera uno o dos minutos. La dirección del sitio aparece en esa misma página: `https://TU-USUARIO.github.io/tesis-arquitectura-cujae/`.

## Actualizar los datos cuando cambie el Excel

1. Guarda el Excel revisado con el nombre **`registro.xlsx`**.
2. En GitHub, entra en la carpeta **`datos`**, pulsa **Add file → Upload files**, arrastra el archivo y pulsa **Commit changes**. Como se llama igual, sustituye al anterior.
3. En uno o dos minutos el sitio muestra los datos nuevos. El buscador, los filtros, las estadísticas y las cifras de la página «Sobre el registro» se recalculan solos.

No hay que tocar nada más.

### Qué puede cambiar en el Excel y qué no

- **Puedes** corregir cualquier celda, añadir o borrar filas, cambiar el orden de las columnas, añadir columnas nuevas y borrar las que no quieras publicar.
- **Conserva** el nombre de la hoja (`Registro`) y los encabezados de la primera fila. El sitio reconoce cada columna por cómo empieza su encabezado:

| Encabezado que empieza por… | Uso en el sitio | ¿Obligatoria? |
|---|---|---|
| `Año` | año de la tesis | sí |
| `Autor` | autores (uno por renglón dentro de la celda) | sí |
| `Título` | título | sí |
| `Tutor` | tutores y cotutores (uno por renglón; el rol va entre paréntesis: `Nombre (cotutor)`) | no |
| `Asesor` | asesores y consultantes | no |
| `Resumen` | resumen | no |
| `N.º` | número de la tesis en el registro | no |
| `Archivo`, `Pág` | archivo PDF de origen y página del resumen | no |
| `Observaciones` | notas que se muestran en la ficha | no |
| `Tipo de aviso` | sirve para excluir filas (ver abajo) | no |
| `Verificación` | cómo se comprobó el resumen | no |
| `Enlace` | dirección web del PDF; si existe, la ficha muestra «Abrir el PDF» | no |

- Las columnas cuyo encabezado contiene «tal como» (los nombres de tutores sin estandarizar) no se usan.
- Las filas cuyo «Tipo de aviso» contiene `No es tesis` no se muestran ni se cuentan.
- Las estadísticas por tutor cuentan los nombres tal como estén escritos en la columna de tutores: dos grafías distintas cuentan como dos personas.

## Cambiar textos y opciones

Abre `js/config.js` (en GitHub: entra al archivo y pulsa el lápiz). Ahí están el título, el subtítulo, el correo o repositorio para avisar de errores, qué filas se excluyen y si se muestran las observaciones.

El texto de la página «Sobre el registro» está en `js/app.js`, en la función `pintarAcerca`. Las cifras de esa página se calculan a partir del Excel.

## Verlo en tu computadora antes de publicar

Si abres `index.html` con doble clic, el navegador no deja leer el Excel de la carpeta y la página te pide elegirlo a mano: pulsa el botón, elige tu `registro.xlsx` y verás el sitio completo con esos datos. Sirve también para probar un Excel revisado antes de subirlo.

## Qué hay en cada carpeta

- `index.html` — la página.
- `css/estilo.css` — el diseño.
- `js/app.js` — el buscador, los filtros y las estadísticas.
- `js/config.js` — textos y opciones.
- `js/vendor/xlsx.mini.min.js` — biblioteca SheetJS (licencia Apache 2.0), que lee el Excel.
- `fuentes/` — las tipografías (licencia OFL).
- `datos/registro.xlsx` — los datos.
