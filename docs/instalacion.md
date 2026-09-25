# Instalación en otra PC (p. ej. la laptop de la empresa)

Para: quien instala y valida el proyecto en un equipo nuevo. Probado originalmente en Windows 11 + Python 3.13 +
Excel 2016 (FRANCISCO-PC).

## 1. Qué necesita el equipo

| Componente | Para qué | ¿Requiere administrador? |
|---|---|---|
| Git | Clonar el repo | No (instalador "solo para mí" o Git portable) |
| Python ≥ 3.13 (python.org) | Todo el proyecto | No: instalar "for me only" y marcar *Add python.exe to PATH* |
| Excel de escritorio (Office) | Macros por COM (`run-macros`) | No |
| ODBC Driver 18 for SQL Server | Conexión a Azure SQL | **Sí** (instalador `.msi` de Microsoft): pedir a TI si no está |
| Acceso a `trina-etl.database.windows.net:1433` | Guardar corridas en Azure | Firewall de Azure (tú) y de la red de la empresa (TI) |

Verificar el driver ODBC (PowerShell):

```powershell
Get-OdbcDriver -Name "ODBC Driver 18 for SQL Server"
```

Verificar la red hacia Azure (debe decir `TcpTestSucceeded : True`; si no, la red de la empresa bloquea el
puerto 1433 y hay que pedirlo a TI):

```powershell
Test-NetConnection trina-etl.database.windows.net -Port 1433
```

## 2. Instalar

```powershell
git clone https://github.com/franciscobarrientoscaceres/ETL_Arena.git
cd ETL_Arena
python -m venv .venv
.venv\Scripts\python -m pip install -e ".[dev]"
```

- **No clonar dentro de OneDrive** (ni carpetas sincronizadas): la sincronización bloquea los `.xlsm` mientras
  Excel los usa. Usar p. ej. `C:\dev\ETL_Arena`.
- Si la red de la empresa usa proxy y `pip` no descarga: `pip install --proxy http://<proxy>:<puerto> …`
  (pedir el proxy a TI).
- Los libros de `data/` vienen en el repo; `data/work/`, `data/processed/` y `.env` **no** (son locales).

## 3. Configurar `.env`

Copiar `.env.example` a `.env` y dejar:

```ini
ETL_ARENA_DB_URL=mssql+pyodbc://@trina-etl.database.windows.net:1433/trina_etl?driver=ODBC+Driver+18+for+SQL+Server&Encrypt=yes&TrustServerCertificate=no
ETL_ARENA_DB_AUTH=entra
# opcional, para el entorno de prueba (runbook §6b):
ETL_ARENA_DB_URL_PRUEBA=mssql+pyodbc://@trina-etl.database.windows.net:1433/trina_etl_prueba?driver=ODBC+Driver+18+for+SQL+Server&Encrypt=yes&TrustServerCertificate=no
```

Sin contraseñas: la primera conexión abre el navegador para iniciar sesión con la cuenta @trinasolar.com.

**Firewall de Azure:** portal → servidor `trina-etl` → Redes → *Agregar la dirección IPv4 del cliente* (una vez por
red desde la que te conectes: oficina, casa, VPN).

## 4. Validar que todo funciona igual

En orden; cada paso debe terminar sin errores.

| # | Comando | Resultado esperado |
|---|---|---|
| 1 | `.venv\Scripts\python -m pytest -q` | ~370 `passed`; se omiten los tests de Excel (y los de SQL si no hay `ETL_ARENA_TEST_DB_URL`) |
| 2 | `.venv\Scripts\python -m ruff check src tests scripts` | `All checks passed!` |
| 3 | `$env:ETL_ARENA_EXCEL = "1"; .venv\Scripts\python -m pytest tests\com -q` | `4 passed` (~2 min; abre Excel, no usarlo mientras) |
| 4 | Ensayo sin base (abajo) | `C14 = 104134.2920000001`, reconciliación `pass` |
| 5 | Ensayo en la base de prueba (runbook §6b) | Corrida `success` visible en `v_kpi_vigente` de `trina_etl_prueba` |

Ensayo sin base (paso 4), con el libro de septiembre:

```powershell
.venv\Scripts\python scripts\run_lunes.py --stage all --omitir-acquire --sin-bd --corte ensayo `
  --work $env:TEMP\etl_ensayo --libro-preparado "data\AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm"
```

Valores de referencia (septiembre 1–21, libro de septiembre): C12 = 1975, C14 = 104134.2920000001,
C16 = 0.9819924099052362, 334 eventos.

**Diferencias esperables según la versión de Excel:**
- Excel 2016 (build < 10000): el log muestra dos avisos "sin SortFields.Add2" (F-40). Es normal.
- Excel 2019/365: no aparecen esos avisos y `ListOfFaults!N:Q` queda ordenado. El KPI no cambia.
- Si Excel bloquea las macros por política de TI: agregar la carpeta `data\work\` como ubicación de confianza.

## 5. Una sola PC oficial

`data/work/libro_base.json` (el libro base) y `data/work/cola_cierres.json` (cierres pendientes) viven solo en la
PC donde se corrió. Las corridas oficiales de cada lunes deben salir **siempre del mismo equipo**. Para validar en
otra PC, usar `--sin-bd` o `--entorno prueba`. Si se cambia de PC oficial, copiar la carpeta `data/work/` completa
(incluidos los cortes a los que apunta `libro_base.json`). `libro_base.json` guarda una ruta absoluta: si la
carpeta del repo cambia, corregir el campo `ruta` (el sha256 se verifica igual).
