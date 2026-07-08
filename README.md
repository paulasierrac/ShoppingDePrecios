# Shopping de Precios — Colsubsidio

Robot RPA (migración de Automation Anywhere a Python/Playwright) que extrae precios de farmacias colombianas en línea.

**Desarrollado por:** KPMG Advisory, Tax & Legal SAS  
**Cliente:** Colsubsidio

---

## Tabla de contenido

- [Descripción general](#descripción-general)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Farmacias implementadas](#farmacias-implementadas)
- [Requisitos e instalación](#requisitos-e-instalación)
- [Configuración](#configuración)
- [Ejecución](#ejecución)
- [Modo debug](#modo-debug)
- [Base de datos](#base-de-datos)
- [Convenciones de precios](#convenciones-de-precios)
- [Estados de registros](#estados-de-registros)
- [Notas técnicas por farmacia](#notas-técnicas-por-farmacia)
- [Errores frecuentes y soluciones](#errores-frecuentes-y-soluciones)

---

## Descripción general

El robot consulta una lista de productos (identificados por EAN) en los sitios web de farmacias colombianas, extrae los precios y genera reportes Excel. El flujo por farmacia se divide en tres historias de usuario (HU):

| HU | Nombre | Descripción |
|----|--------|-------------|
| HU00 | DespliegueAmbiental | Carga parámetros desde SQL Server, limpia carpetas y tablas según fecha |
| HU01 | ValidacionYCargaInsumo | Lee el archivo Excel de insumo y lo carga en la tabla `TicketInsumo` de la BD |
| HU02 | ConsultaYReporte | Scraping web por EAN, guarda resultados en BD y genera Excel de reporte |

---

## Estructura del proyecto

```
ShoppingDePrecios/
│
├── ShoppingDePrecios<Farmacia>/    # Una carpeta por farmacia
│   ├── main.py                     # Orquestador: llama HU00 → HU01 → HU02
│   └── HU/
│       ├── HU00_DespliegueAmbiental.py
│       ├── HU01_ValidacionYCargaInsumo.py
│       └── HU02_ConsultaYReporte.py
│
├── Funciones/
│   └── utils.py                    # write_log, conectar_bd, enviar_correo, etc.
│
├── Config/
│   └── Configuracion.py            # CargarVault() — Azure Key Vault
│
├── Insumo/
│   └── InsumoPricing.xlsx          # Archivo de entrada (PLU / EAN / DESCRIPCION / PROVEEDOR / CATEGORIA)
│
├── Resultado/                      # Reportes Excel generados (producción)
├── pruebas.db                      # SQLite local para modo debug
└── requirements.txt
```

---

## Farmacias implementadas

| Farmacia | HU00 | HU01 | HU02 | Motor scraping | Estado |
|----------|------|------|------|----------------|--------|
| Locatel | ✓ | ✓ | ✓ | Playwright | Completo |
| Exito | ✓ | ✓ | ✓ | Playwright | Completo |
| Cafam | ✓ | ✓ | ✓ | Playwright | Completo |
| Farmatodo | ✓ | ✓ | ✓ | Playwright | Completo |
| CruzVerde | ✓ | ✓ | ✓ | Playwright | Completo |
| LaRebaja | ✓ | ✓ | stub | — | Pendiente ZIP |
| Medipiel | ✓ | ✓ | stub | — | Pendiente ZIP |
| Olimpica | ✓ | ✓ | stub | — | Pendiente ZIP |
| Ortopedicos | ✓ | ✓ | stub | — | Pendiente ZIP |
| Pasteur | ✓ | ✓ | stub | — | Pendiente ZIP |
| Alemana | ✓ | ✓ | stub | — | Pendiente ZIP |
| Comfandi | ✓ | ✓ | stub | — | Pendiente ZIP |

---

## Requisitos e instalación

### 1. Dependencias Python

```bash
pip install -r requirements.txt
```

| Librería | Versión mínima | Para qué se usa |
|----------|----------------|-----------------|
| `pandas` | 2.0 | Lectura del Excel de insumo, generación de reportes Excel y manipulación de datos tabulares en memoria |
| `selenium` | 4.10 | Dependencia legada — todas las farmacias activas migraron a Playwright. Ya no se usa en HU02 de ninguna farmacia |
| `playwright` | 1.40 | Motor de scraping web para las cinco farmacias completas (Locatel, Éxito, Cafam, Farmatodo, Cruz Verde). Requiere ejecutar `playwright install chromium` por cada usuario de Windows |
| `pyodbc` | 5.0 | Conexión a SQL Server mediante ODBC Driver 17/18. Usado en todas las HU para leer parámetros, cargar insumo y guardar resultados |
| `openpyxl` | 3.1 | Motor de escritura de archivos `.xlsx` usado por pandas (`pd.ExcelWriter(..., engine="openpyxl")`) y para incrustar imágenes en el reporte de Cafam |
| `Pillow` | 10.0 | Incrustación de thumbnails (screenshots) dentro de los reportes Excel de Cafam |
| `azure-identity` | 1.12 | Autenticación con Azure (credenciales de cuenta de servicio o identidad administrada) para acceder al Key Vault |
| `azure-keyvault-secrets` | 4.6 | Lectura de secretos desde Azure Key Vault (usuario, contraseña y servidor de BD) |
| `python-dotenv` | 1.0 | Carga de variables de entorno desde archivo `.env` en desarrollo local |
| `pydantic` | 2.0 | Validación de modelos de configuración internos |
| `pydantic-settings` | 2.0 | Carga de settings desde variables de entorno usando modelos Pydantic |
| `python-dateutil` | 2.8 | Parsing de fechas en distintos formatos (usado internamente por pandas) |

### 2. Navegador Playwright

Debe ejecutarse **una vez por usuario de Windows** en la máquina donde corre el bot:

```bash
playwright install chromium
```

> **Nota:** Si el bot corre bajo un usuario de servicio (ej. `TEMP.COLSUBSIDIO.XXX`), este comando debe ejecutarse con ese mismo usuario. El bot intenta instalarlo automáticamente si detecta que el binario no existe.

### 3. Credenciales

Las credenciales de base de datos se obtienen exclusivamente desde **Azure Key Vault** mediante:

```python
CargarVault(filtro_tags={"shared": "true", "environment": "dev"}, strip_prefix="Dev")
```

No se usa ningún archivo `config.json`. Toda la configuración operativa proviene de la tabla `[ShoppingDePrecios].[Parametros]` en SQL Server.

---

## Configuración

Todos los parámetros se leen de la tabla `[ShoppingDePrecios].[Parametros]` en SQL Server. **No existe ningún archivo de configuración local** — las rutas llegan completas desde la BD.

### Rutas de trabajo

| Parámetro | Descripción | Ejemplo |
|-----------|-------------|---------|
| `RutaInsumos` | Ruta completa donde se deposita el archivo `InsumoPricing.xlsx` antes de ejecutar el bot. Tras cargarlo a BD, el archivo se mueve a `CarpetaProcesados`. | `\\servidor\ShoppingPrecios\Insumos\` |
| `CarpetaProcesados` | Subcarpeta (relativa a `RutaInsumos`) donde se archiva el insumo ya procesado con sello de fecha. | `Procesados\` |
| `RutaReporte` | Ruta completa donde se guardan los Excel de resultado generados por HU02. | `\\servidor\ShoppingPrecios\Resultados\` |
| `RutaScreenshots` | Ruta base para capturas de pantalla del scraping. Se crean subcarpetas automáticas por farmacia y fecha (`Farmacia\AAAA\MM\DD\`). | `\\servidor\ShoppingPrecios\Screenshots\` |
| `RutaTemp` | Ruta para el archivo CSV temporal (`Insumo.csv`) que se genera durante HU01. | `C:\Temp\ShoppingPrecios\` |
| `RutaRed` | Ruta de red general del proyecto (usada por HU00 para limpieza). | `\\servidor\ShoppingPrecios\` |
| `PathLog` | Ruta donde se escriben los archivos de log `.txt` de cada ejecución. Si no es accesible, el log se imprime solo en consola. | `\\servidor\ShoppingPrecios\Logs\` |

### Nombres de archivo y hoja

| Parámetro | Descripción |
|-----------|-------------|
| `ArchivoInsumo` | Nombre del archivo Excel de insumo. Valor esperado: `InsumoPricing.xlsx` |
| `SheetTicketInsumo` | Nombre de la hoja dentro del Excel de insumo. Valor esperado: `TicketInsumo` |
| `NombreResultado` | Prefijo del archivo Excel de resultado. Ej: `ReportePricingExito_` |
| `NombreHojaResultado` | Nombre de la hoja dentro del Excel de resultado. Ej: `ReportePricingExito` |

### URLs de scraping

Cada farmacia tiene su propia URL. El bot reemplaza el token `REEMPLAZAR` por el EAN antes de navegar.

| Parámetro | Farmacia | Ejemplo de URL |
|-----------|----------|----------------|
| `UrlExito` | Éxito | `https://www.exito.com/s?q=REEMPLAZAR&sort=score_desc&page=0` |
| `UrlLocatel` | Locatel | `https://www.locatelcolombia.com/REEMPLAZAR` |
| `UrlCafam` | Cafam | `https://www.cafam.com.co/...?q=REEMPLAZAR` |
| `UrlFarmatodo` | Farmatodo | `https://www.farmatodo.com.co/buscar?product=REEMPLAZAR&` |
| `UrlCruzVerde` | Cruz Verde | `https://www.cruzverde.com.co/...?q=REEMPLAZAR` |

### Control de scraping por farmacia

Cada farmacia tiene su propio par de parámetros para controlar la velocidad del scraping:

| Parámetro | Farmacia | Descripción |
|-----------|----------|-------------|
| `CantExito` | Éxito | Número de EANs por lote antes de reiniciar el contexto del navegador |
| `SegExito` | Éxito | Segundos de espera entre lotes (evita bloqueos por rate limiting) |
| `LoteCafam` | Cafam | EANs por lote |
| `DelayCafam` | Cafam | Segundos de espera entre lotes |
| `LoteCruzVerde` | Cruz Verde | EANs por lote |
| `DelayCruzVerde` | Cruz Verde | Segundos de espera entre lotes |
| `CantFarmatodo` | Farmatodo | EANs por lote |
| `SegFarmatodo` | Farmatodo | Segundos de espera entre lotes |

> **Recomendación:** Valores de lote entre 20 y 50, y delays de 3 a 10 segundos. Valores muy bajos pueden causar bloqueos por parte del sitio web.

### Navegador

| Parámetro | Descripción | Valores |
|-----------|-------------|---------|
| `HeadlessChrome` | Controla si el navegador corre en modo invisible (sin interfaz gráfica). En producción debe ser `"true"`. | `"true"` / `"false"` |

### Mantenimiento y debug

| Parámetro | Descripción |
|-----------|-------------|
| `LimpiezaDB` | Fecha de la última limpieza de registros históricos en BD (formato `YYYY-MM-DD`). HU00 compara esta fecha con la del día actual para evitar limpiar más de una vez por día. Se actualiza automáticamente tras cada limpieza. |
| `LoteDebug` | Número máximo de EANs a consultar cuando el bot corre en **modo debug**. Permite hacer pruebas rápidas sin procesar todo el insumo. Valor recomendado: `3` a `10`. |

### Nombres de farmacias (para correos)

| Parámetro | Descripción |
|-----------|-------------|
| `DrogueriaExito` | Nombre legible de la farmacia, usado en el asunto y cuerpo del correo de resultado. Ej: `Éxito` |
| `DrogueriaLocatel` | Ídem para Locatel |
| `DrogueriaCafam` | Ídem para Cafam |
| _(etc.)_ | |

### Carpetas de screenshots por farmacia

| Parámetro | Descripción |
|-----------|-------------|
| `CarpetaExito` | Subcarpeta (relativa a `RutaScreenshots`) para las capturas de Éxito. Ej: `Exito\` |
| `CarpetaLocatel` | Ídem para Locatel |
| _(etc.)_ | |

---

## Ejecución

Cada farmacia tiene su propio `main.py`:

```bash
# Producción
python ShoppingDePreciosExito/main.py
python ShoppingDePreciosLocatel/main.py
# etc.
```

El orquestador ejecuta HU00 → HU01 → HU02 en secuencia. Si alguna HU retorna error, el proceso se detiene y envía correo de notificación.

---

## Modo debug

Activa el modo debug con la variable de entorno `RPA_DEBUG`:

```powershell
# PowerShell
$env:RPA_DEBUG = "true"
python ShoppingDePreciosExito/main.py
```

```cmd
:: CMD
set RPA_DEBUG=true
python ShoppingDePreciosExito/main.py
```

### Comportamiento en modo debug

| HU | Comportamiento |
|----|----------------|
| HU00 | Lee parámetros de SQL Server (solo lectura — sin cambios) |
| HU01 | Lee `Insumo/InsumoPricing.xlsx` local → CSV en `debug/temp/` → INSERT en `pruebas.db` (no mueve el archivo) |
| HU02 | Lee de `pruebas.db` → Chrome **visible** → Escribe resultados en `pruebas.db` → Excel en `debug/` |
| Correos | Omitidos |
| SQL Server (escrituras) | Ninguna escritura en producción; solo lecturas en HU00 |

El archivo de insumo local de pruebas se encuentra en `Insumo/InsumoPricing.xlsx` y **no se mueve ni elimina** en modo debug, lo que permite re-ejecutar sin regenerar el archivo.

---

## Base de datos

### Esquema principal: `[ShoppingDePrecios]`

| Tabla | Descripción |
|-------|-------------|
| `TicketInsumo` | Insumo cargado desde Excel (EAN, PLU, Descripcion, Proveedor, Categoria) |
| `Locatel` | Resultados de scraping Locatel |
| `Exito` | Resultados de scraping Éxito |
| `Cafam` | Resultados de scraping Cafam |
| `Farmatodo` | Resultados de scraping Farmatodo |
| `CruzVerde` | Resultados de scraping Cruz Verde |
| `Parametros` | Parámetros de configuración del robot |
| `Selectores` | Selectores CSS de scraping para Farmatodo (**opcional** — el código usa fallbacks hardcodeados si la tabla no existe o está vacía) |
| `EnvioCorreos` | Plantillas de correos de notificación |

### Grupos de tablas de resultados

**Grupo A** (Locatel, Cafam, CruzVerde) — incluye columnas `Observaciones`, `Reintentos` y `Categoria`.

**Grupo B** (Éxito, Farmatodo) — sin `Observaciones` ni `Reintentos`.

### Formato de precios colombiano

- Separador de miles: `.` (punto) → se elimina con `REPLACE('.', '')`
- Cafam usa `,` como decimal → se convierte con `REPLACE(',', '.')`

---

## Convenciones de precios

Todos los HU02 siguen la misma convención para asignar los campos de precio:

| Campo | Contenido |
|-------|-----------|
| `PrecioSinDescuento` | Precio regular (precio de lista, sin descuento aplicado) |
| `PrecioConDescuento` | Precio final tras aplicar el descuento. **Vacío si no hay descuento activo.** |
| `Porc.Descuento` | Porcentaje de descuento, calculado como `(SinDesc - ConDesc) * 100 / SinDesc` |
| `PrecioFidelizacion` | Precio especial para clientes del programa de fidelización (Éxito Club, etc.) |
| `PrecioUnitario` | Precio por unidad de medida (PUM). Ej: `(Ml a $ 12,72)` o `Mililitros a $ 14.08` |

> **Regla clave:** Cuando el sitio muestra un solo precio sin tachado (sin descuento activo), el valor se guarda en `PrecioSinDescuento` y `PrecioConDescuento` queda vacío. Esto aplica a Éxito, Farmatodo y Cafam.

---

## Estados de registros

Aplican a todas las tablas de resultados (`Locatel`, `Exito`, `Cafam`, etc.):

| Estado | Significado |
|--------|-------------|
| `1` | Pendiente de consultar |
| `2` | Producto encontrado |
| `3` | Sin coincidencia (nombre no corresponde al EAN buscado) |
| `99` | Sin información (producto no aparece en la búsqueda) |
| `100` | Consultado y reportado (fue Estado=2) |
| `199` | Consultado y reportado (fue Estado=99) |

---

## Notas técnicas por farmacia

### Estrategia de extracción JavaScript

Todas las farmacias usan `page.evaluate()` de Playwright para extraer datos del DOM, pero con enfoques distintos:

| Farmacia | Estrategia | Detalle |
|----------|-----------|---------|
| Cafam | **Bloque JS unificado** | Un solo `page.evaluate()` recorre las tarjetas Doofinder y devuelve un objeto con todos los campos |
| Cruz Verde | **Bloque JS unificado** | Un solo `page.evaluate()` extrae nombre, precio, PUM e INVIMA de la página de detalle Angular |
| Éxito | **JS → parse Python** | JS obtiene el `innerHTML` completo de las tarjetas; Python lo parsea con `_entre()` buscando clases CSS por nombre |
| Locatel | **JS por campo** | Un `page.evaluate()` independiente por cada dato (título, precio, disponibilidad, marca) |
| Farmatodo | **JS por campo** | Helper `_js_selector(selector)` ejecuta `querySelector(sel)?.textContent` por cada campo |

> **Nota sobre Éxito:** el enfoque de parsear HTML con `_entre()` es frágil ante cambios de clases CSS (los nombres como `styles_name__qQJiK` son generados por el bundler y pueden cambiar entre deploys). Las farmacias con bloque JS unificado son más robustas porque trabajan con la estructura semántica del DOM, no con nombres de clases.

---

### Éxito
- El scraping se realiza sobre la **página de resultados de búsqueda** (`/s?q=EAN`). La URL del producto almacenada puede ser la URL de búsqueda (si el sitio no redirige) o la URL directa del producto (si Éxito la incluye en el HTML de resultados).
- El `PrecioUnitario` se extrae de la clase `.product-unit_price-unit__text__qeheS`. No todos los productos lo muestran en la página de resultados.

### Cafam
- Cafam usa **Doofinder** con Phoenix LiveView (WebSocket). Para evitar que el sitio detecte múltiples conexiones como tráfico de bot, **se mantiene una sola página por lote** y se reutiliza el campo de búsqueda Doofinder (`.dfd-searchbox-input`) para los EANs subsiguientes.
- El `PrecioUnitario` se transforma del formato `PUM: $ 14.00 ML` al formato `(ML a $ 14.00)` antes de guardar.
- Los screenshots se **incrustan como thumbnails (160×120 px) directamente en el Excel** de resultado, además de guardarse en disco. Requiere `Pillow>=10.0`.

### Farmatodo
- Los selectores CSS se cargan desde la tabla `[ShoppingDePrecios].[Selectores]` con `Competencia='FARMATODO'`. Si la tabla no existe o está vacía, el código usa los siguientes **selectores por defecto** basados en el HTML real del sitio (Angular SPA):

  | Campo | Selector CSS |
  |-------|-------------|
  | Nombre | `.text-title` |
  | Marca | `.text-brand` |
  | Precio con descuento | `.price__text-price` |
  | Precio sin descuento | `.price__text-offer-price` |
  | PUM | `.price__text-price-unit` |
  | Banner | `.offer-description__text` |
  | URL producto | `a.content-product[href]` |

### Locatel
- Locatel redirige directamente a la página de detalle del producto cuando hay un único resultado por EAN (URL con `/p?skuId=`). El scraper detecta ambos casos (página de detalle vs. página de resultados).
- El `Porc.Descuento` se calcula en BD durante la generación del reporte, antes de marcar los registros como Estado=100.

### Cruz Verde
- Cruz Verde es una SPA Angular con componentes `ml-card-product`. Los datos se extraen via JavaScript sobre el DOM de la página de detalle del producto.
- Los campos `RegistroInvima` y `PrecioUnitario` se incluyen cuando están disponibles en la página.

---

## Errores frecuentes y soluciones

### `Executable doesn't exist at ...ms-playwright\chromium...`

Playwright está instalado como librería pero los binarios del navegador no se descargaron para el usuario actual. El bot lo resuelve automáticamente en el siguiente intento. Si persiste, ejecutar manualmente:

```bash
playwright install chromium
```

### `Invalid object name 'ShoppingDePrecios.Selectores'`

La tabla de selectores CSS de Farmatodo no existe en la base de datos. **No bloquea la ejecución** — el código registra el warning y continúa usando los selectores hardcodeados. Para activar selectores configurables desde BD, crear la tabla con las columnas `[Clave]`, `[Selector]` y `[Competencia]` e insertar las filas con `Competencia='FARMATODO'`.

### `Incorrect syntax near 'LIMIT'`

SQL Server no soporta `LIMIT`. Usar `TOP(N)` en su lugar. Este error ya fue corregido en el código.

### `Cannot insert explicit value for identity column`

La tabla tiene una columna `Id` con IDENTITY. Requiere `SET IDENTITY_INSERT ON` antes del INSERT. Todos los HU02 ya tienen esta corrección.

### `ERR_CONNECTION_RESET` o `SSL handshake failed`

El proxy corporativo hace inspección SSL (Deep Packet Inspection). Playwright está configurado con `ignore_https_errors=True` en el contexto del navegador para manejar esto. Si el error persiste verificar que el proxy del sistema esté bien configurado en Windows.

### `net::ERR_ABORTED` en Cafam (segundo EAN en adelante)

Ocurre si se crea una nueva página (`context.new_page()`) por cada EAN en Cafam. El sitio detecta múltiples conexiones HTTP como bot y las aborta. La solución implementada es usar **una sola página por lote** y cambiar el EAN usando el campo de búsqueda de Doofinder.

### Registros duplicados en tablas de resultados

Causado por ejecutar HU01 múltiples veces con `DELETE FROM` en lugar de `TRUNCATE TABLE`. `TRUNCATE` resetea el contador IDENTITY, evitando que se generen IDs duplicados. Todos los HU01 ya usan `TRUNCATE TABLE`.

### `Porc.Descuento` vacío en Locatel

Ocurre si el cálculo del porcentaje de descuento se ejecuta después de que el estado ya cambió a `100`. El UPDATE de `Porc.Descuento` debe correr mientras los registros aún están en Estado=`2`. Esta corrección ya está aplicada en el código.

### Imágenes no incrustadas en el Excel de Cafam

Requiere `Pillow>=10.0`. Verificar con `pip install Pillow>=10.0`. Si `Pillow` no está disponible, `_incrustar_imagenes` falla silenciosamente y el Excel se genera sin thumbnails (las rutas de imagen siguen presentes en la columna `RutaImagen`).
