# Shopping de Precios — Contexto del Proyecto

## Descripción
Robot RPA (migración de Automation Anywhere a Python/Selenium) que extrae precios de farmacias colombianas en línea. Desarrollado para Colsubsidio por KPMG Advisory, Tax & Legal SAS.

## Estructura
Cada farmacia tiene su propia carpeta `ShoppingDePrecios<Farmacia>/` con:
```
ShoppingDePrecios<Farmacia>/
├── main.py                          # Orquestador principal
└── HU/
    ├── HU00_DespliegueAmbiental.py  # Carga parámetros desde BD, limpieza de carpetas
    ├── HU01_ValidacionYCargaInsumo.py  # Lee Excel → carga TicketInsumo en BD
    └── HU02_ConsultaYReporte.py     # Scraping web + reporte Excel
```

Carpetas compartidas:
```
Funciones/utils.py     # write_log, conectar_bd, conectar_bd_debug, enviar_correo, etc.
Config/Configuracion.py  # CargarVault() — Azure Key Vault
Insumo/                # Archivo Excel de insumo (InsumoPricing.xlsx)
Resultado/             # Reportes generados
pruebas.db             # SQLite para modo debug (tablas: TicketInsumo, Locatel, Exito, Cafam, Farmatodo, CruzVerde, Parametros)
requirements.txt
```

## Farmacias implementadas
| Farmacia    | HU00 | HU01 | HU02 | Estado     |
|-------------|------|------|------|------------|
| Locatel     | ✓    | ✓    | ✓    | Completo   |
| Exito       | ✓    | ✓    | ✓    | Completo   |
| Cafam       | ✓    | ✓    | ✓    | Completo   |
| Farmatodo   | ✓    | ✓    | ✓    | Completo   |
| CruzVerde   | ✓    | ✓    | ✓    | Completo   |
| LaRebaja    | ✓    | ✓    | stub | Sin ZIP    |
| Medipiel    | ✓    | ✓    | stub | Sin ZIP    |
| Olimpica    | ✓    | ✓    | stub | Sin ZIP    |
| Ortopedicos | ✓    | ✓    | stub | Sin ZIP    |
| Pasteur     | ✓    | ✓    | stub | Sin ZIP    |
| Alemana     | ✓    | ✓    | stub | Sin ZIP    |
| Comfandi    | ✓    | ✓    | stub | Sin ZIP    |

## Reglas de seguridad (NUNCA violar)
1. **Sin config.json** — toda configuración viene de `[ShoppingDePrecios].[Parametros]` en SQL Server
2. **Credenciales de BD** exclusivamente desde Azure Key Vault via `CargarVault(filtro_tags={"shared":"true","environment":"dev"}, strip_prefix="Dev")`
3. **Sin escrituras en BD en modo debug** — usar `pruebas.db` (SQLite local)

## Modo Debug
Activar: `set RPA_DEBUG=true` (o `$env:RPA_DEBUG="true"` en PowerShell)

Comportamiento en debug:
- HU00: lee parámetros de SQL Server normalmente (solo lectura)
- HU01: lee `Insumo/InsumoPricing.xlsx` **local** → CSV en `debug/temp/` → INSERT en `pruebas.db.TicketInsumo` (no mueve el archivo)
- HU02: lee de `pruebas.db.TicketInsumo` → Chrome visible → escribe en `pruebas.db.<Farmacia>` → Excel en `debug/`
- Correos: enviados igual que en producción (Excel adjunto desde `debug/YYYY/MM/DD/`)
- SQL Server: solo lecturas en HU00, ninguna escritura

## Parámetros clave en `[ShoppingDePrecios].[Parametros]`
Todos los parámetros se cargan directamente desde esta tabla. Las rutas vienen del DB:
- `RutaInsumos`, `PathLog`, `RutaScreenshots`, `RutaReporte`, `RutaTemp`, `RutaRed`
- `UrlExito`, `UrlCafam`, `UrlFarmatodo`, `UrlCruzVerde`, etc.
- `LoteCafam`, `LoteCruzVerde`, `CantExito`, `CantFarmatodo` — tamaño de lote scraping
- `DelayCafam`, `DelayCruzVerde`, `SegExito`, `SegFarmatodo` — delay entre lotes (segundos)
- `LimpiezaDB` — fecha última limpieza de BD (evita limpiar más de una vez por día)
- `HeadlessChrome` — "true"/"false"
- `LoteDebug` — número de EANs a procesar en modo debug (default 3)

**IMPORTANTE:** La tabla NO tiene parámetro `fserver`. Las rutas llegan completas desde la BD.

## Funciones utilitarias (Funciones/utils.py)
- `obtener_config()` — carga Key Vault + retorna `{"_db": {...}, "_correo": {...}, "Scheme": "..."}`
- `conectar_bd(config)` — pyodbc a SQL Server (fallback chain: driver 17 → 18 → 13 → "SQL Server")
- `conectar_bd_debug()` — sqlite3 a `pruebas.db` (para modo debug)
- `write_log(state, msg, task, config)` — siempre imprime a consola; escribe a archivo si PathLog accesible
- `enviar_correo(...)` — lee tabla `EnvioCorreos` de BD, envía por SMTP
- `excel_a_csv(...)` / `csv_a_excel(...)`

## Patrones de scraping
- **Exito**: JavaScript sobre clases CSS del DOM de resultados
- **Cafam**: JS → tarjetas `dfd-card-link`, navega a detalle, CSS selectors
- **Farmatodo**: selectores CSS cargados desde `[ShoppingDePrecios].[Selectores]` WHERE Competencia='FARMATODO'
- **CruzVerde**: Angular `ml-card-product`, extrae innerHTML con JS, parsea con `_entre()`
- **Locatel**: similar a Exito

## Formato de precios colombiano
- Separador de miles: `.` (punto) → eliminar con `REPLACE('.', '')`
- Cafam usa `,` como decimal → `REPLACE(',', '.')`

## Dependencias
Ver `requirements.txt`. Instalación: `pip install -r requirements.txt`
ChromeDriver: gestionado automáticamente por Selenium 4.x.

## Estados en tablas de resultados
- `1` = Pendiente de consultar
- `2` = Información encontrada (scraping exitoso)
- `99` = Sin información / error al extraer

Ambos estados (`2` y `99`) se eliminan de la tabla tras generar el reporte (quedan capturados en el Excel). La tabla queda vacía al finalizar cada ejecución.

> Estado=3, Estado=100 y Estado=199 fueron eliminados. Los casos que antes usaban Estado=3 ahora se registran como Estado=99 con observación descriptiva.
