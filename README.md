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
- [Estados de registros](#estados-de-registros)
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
| `playwright` | 1.40 | Motor de scraping web para Locatel y Éxito. Más estable que Selenium en entornos con proxy SSL corporativo. Requiere ejecutar `playwright install chromium` por cada usuario de Windows |
| `pyodbc` | 5.0 | Conexión a SQL Server mediante ODBC Driver 17/18. Usado en todas las HU para leer parámetros, cargar insumo y guardar resultados |
| `openpyxl` | 3.1 | Motor de escritura de archivos `.xlsx` usado por pandas (`pd.ExcelWriter(..., engine="openpyxl")`) |
| `azure-identity` | 1.12 | Autenticación con Azure (credenciales de cuenta de servicio o identidad administrada) para acceder al Key Vault |
| `azure-keyvault-secrets` | 4.6 | Lectura de secretos desde Azure Key Vault (usuario, contraseña y servidor de BD) |
| `python-dotenv` | 1.0 | Carga de variables de entorno desde archivo `.env` en desarrollo local |
| `pydantic` | 2.0 | Validación de modelos de configuración internos |
| `pydantic-settings` | 2.0 | Carga de settings desde variables de entorno usando modelos Pydantic |
| `python-dateutil` | 2.8 | Parsing de fechas en distintos formatos (usado internamente por pandas) |

### 2. Navegador Playwright (Locatel y Éxito)

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
| `UrlFarmatodo` | Farmatodo | `https://www.farmatodo.com.co/...?q=REEMPLAZAR` |
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
| `LoteDebug` | Número máximo de EANs a consultar cuando el bot corre en **modo debug**. Permite hacer pruebas rápidas sin procesar todo el insumo. Valor recomendado: `3` a `10`. Para procesar todos los registros en debug, igualar este valor al total de registros en el insumo. |

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
| HU01 | Lee `Insumo/InsumoPricing.xlsx` local → CSV en `debug/temp/` → INSERT en BD dev (no mueve el archivo) |
| HU02 | Lee de BD dev → Chrome **visible** → Escribe resultados en BD dev → Excel en `debug/` |
| Correos | Omitidos |
| SQL Server (escrituras) | Solo en BD dev, ninguna escritura en producción |

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
| `Selectores` | Selectores CSS de scraping (usado por Farmatodo) |
| `EnvioCorreos` | Plantillas de correos de notificación |

### Grupos de tablas de resultados

**Grupo A** (Locatel, Cafam, CruzVerde) — 25 columnas, incluye `Observaciones` y `Reintentos`.

**Grupo B** (Éxito, Farmatodo) — 20 columnas, sin `Observaciones`.

### Formato de precios colombiano

- Separador de miles: `.` (punto) → se elimina con `REPLACE('.', '')`
- Cafam usa `,` como decimal → se convierte con `REPLACE(',', '.')`

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

## Errores frecuentes y soluciones

### `Executable doesn't exist at ...ms-playwright\chromium...`

Playwright está instalado como librería pero los binarios del navegador no se descargaron para el usuario actual. El bot lo resuelve automáticamente en el siguiente intento. Si persiste, ejecutar manualmente:

```bash
playwright install chromium
```

### `Incorrect syntax near 'LIMIT'`

SQL Server no soporta `LIMIT`. Usar `TOP(N)` en su lugar. Este error ya fue corregido en el código.

### `Cannot insert explicit value for identity column`

La tabla tiene una columna `Id` con IDENTITY. Requiere `SET IDENTITY_INSERT ON` antes del INSERT. Todos los HU02 ya tienen esta corrección.

### `ERR_CONNECTION_RESET` o `SSL handshake failed`

El proxy corporativo hace inspección SSL (Deep Packet Inspection). Playwright está configurado con `ignore_https_errors=True` en el contexto del navegador para manejar esto. Si el error persiste verificar que el proxy del sistema esté bien configurado en Windows.

### Registros duplicados en tablas de resultados

Causado por ejecutar HU01 múltiples veces con `DELETE FROM` en lugar de `TRUNCATE TABLE`. `TRUNCATE` resetea el contador IDENTITY, evitando que se generen Ids duplicados. Todos los HU01 ya usan `TRUNCATE TABLE`.

### Correos no enviados — columnas no encontradas

El archivo `EnvioCorreos.xlsx` usa encabezados en español. La función `cargar_tabla_envio_correos` en `utils.py` hace el mapeo automático a los nombres que espera la BD.
