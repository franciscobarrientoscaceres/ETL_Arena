# Modelo conceptual de datos

Para: cualquiera que quiera entender **qué se guarda en la base de datos y por qué**, sin leer SQL.
Las palabras técnicas están en el [glosario](./glosario.md). El detalle exacto de cada tabla está en `sql/02_corrida.sql`.

> **La idea en una frase:** la base guarda **una sola versión vigente de cada mes**. Cuando se vuelve a cargar un
> mes (por ejemplo, el lunes siguiente con más días), sus datos **se reemplazan**: no se acumulan copias. Aparte, un
> **registro de ejecuciones** anota cada carga (cuándo, qué archivo, si cuadró con el Excel y qué datos cambiaron).
> Decisión: [ADR-12](./adr/ADR-12-estado-vigente.md).

Los diagramas se ven como dibujos en GitHub y en VS Code (con la extensión *Markdown Preview Mermaid Support*).

Este modelo existe **igual en las dos bases** del servidor `trina-etl.database.windows.net`: **PROD** (`trina_etl`, lo oficial) y **TEST/QA** (`trina_etl_prueba`, para practicar). Ver el [README](../README.md#las-dos-bases-de-datos-prod-y-testqa).

---

## 1. El viaje de un dato

Desde que un PCS informa cuántas baterías tiene, hasta que aparece en el reporte de Power BI:

```mermaid
%%{init: {'theme':'base','themeVariables':{'fontSize':'15px'}}}%%
flowchart LR
    subgraph P["🏭 Planta Arena"]
        PCS["61 PCS<br/>× 4 baterías<br/>× 12 racks"]
    end
    subgraph S["📡 SCADA"]
        EXP["Export semanal<br/>RawData-PCS"]
    end
    subgraph A["👤 Alex (fin de mes)"]
        EM["Exclusion_Matrix<br/>0 / 1 / 2 por PCS"]
    end
    subgraph X["📗 Libro Excel (copia de trabajo)"]
        HOJAS["RawData-PCS<br/>Exclusion_Matrix<br/>PlantActivity"]
        MAC["Macros<br/>(referencia)"]
    end
    subgraph PY["🐍 Python"]
        CALC["Cálculo<br/>disponibilidad + eventos"]
        REC["Reconciliación<br/>Python vs Excel"]
    end
    subgraph DB["🗄️ Azure SQL"]
        TAB["Tablas de estado<br/>(una versión por mes)"]
        REG["Registro de ejecuciones<br/>(historia de cada carga)"]
    end
    PBI["📊 Power BI<br/>(Misael)"]

    PCS -- "cada 15 min" --> EXP
    EXP -- "TeamViewer" --> HOJAS
    EM --> HOJAS
    HOJAS --> MAC
    HOJAS --> CALC
    MAC --> REC
    CALC --> REC
    REC -- "si cuadra: reemplaza el mes" --> TAB
    REC -- "siempre" --> REG
    TAB --> PBI

    classDef planta fill:#FFF4D6,stroke:#C9A227,color:#3D3000
    classDef fuente fill:#E3F2FD,stroke:#1E88E5,color:#0D2B45
    classDef excel fill:#E8F5E9,stroke:#43A047,color:#1B3D1F
    classDef py fill:#F3E5F5,stroke:#8E24AA,color:#3A1245
    classDef bd fill:#FFEBEE,stroke:#E53935,color:#4A1212
    classDef bi fill:#FFF3E0,stroke:#FB8C00,color:#4A2A00
    class PCS planta
    class EXP,EM fuente
    class HOJAS,MAC excel
    class CALC,REC py
    class TAB,REG bd
    class PBI bi
```

---

## 2. El mapa general: tres cajones

Piensa en la base de datos como **tres cajones**:

1. **Maestros**: lo que casi nunca cambia (la planta y el catálogo de fallas).
2. **Estado vigente**: lo que Power BI muestra. **Una sola versión por mes**; se reemplaza al recargar ese mes.
3. **Registro de ejecuciones**: la historia. Una fila por cada carga y lo que pasó en ella. Nunca se borra.

```mermaid
%%{init: {'theme':'base','themeVariables':{'fontSize':'15px'}}}%%
flowchart TB
    subgraph M["🗂️ 1. Maestros — casi nunca cambian"]
        direction LR
        PR["Proyecto<br/><small>proyecto</small>"]
        CF["Catálogo de fallas<br/><small>tipo_detencion</small>"]
    end
    subgraph E["📊 2. Estado vigente — una versión por mes (lo que ve Power BI)"]
        direction TB
        E1["Bloques de 15 min<br/><small>muestra_pcs · muestra_planta</small>"]
        E2["Detenciones<br/><small>detencion</small>"]
        E3["Día y mes<br/><small>disponibilidad_diaria · disponibilidad_mensual</small>"]
        E4["Problemas de calidad<br/><small>calidad_dato</small>"]
    end
    subgraph R["🧾 3. Registro de ejecuciones — historia, nunca se borra"]
        direction TB
        R1["Ejecución<br/><small>etl_run</small>"]
        R2["Comparación con Excel<br/><small>excel_reference_run · reconciliation_result</small>"]
        R3["Cambios y entregas<br/><small>correccion_dato · exclusion_matrix_carga</small>"]
        R4["Revisiones humanas<br/><small>detencion_revision</small>"]
    end

    PR -- "1 proyecto : muchos meses" --> E
    CF -. "clasifica" .-> E2
    R1 -- "reemplaza el mes<br/>(solo si cuadra)" --> E
    R4 -. "se muestra junto a" .-> E2

    classDef m fill:#E3F2FD,stroke:#1E88E5,color:#0D2B45
    classDef e fill:#E8F5E9,stroke:#43A047,color:#1B3D1F,stroke-width:2px
    classDef r fill:#FFF4D6,stroke:#C9A227,color:#3D3000
    class PR,CF m
    class E1,E2,E3,E4 e
    class R1,R2,R3,R4 r
```

### Cómo funciona una recarga

Ejemplo: el lunes 21 se cargó septiembre del 1 al 20. El lunes 28 se vuelve a correr septiembre del 1 al 27.

```mermaid
%%{init: {'theme':'base','themeVariables':{'fontSize':'14px'}}}%%
sequenceDiagram
    participant L as run_lunes.py
    participant B as Base de datos
    L->>B: 1. Registra la ejecución N° 15 (etl_run, "running")
    L->>L: 2. Calcula septiembre 1–27 y lo compara con el Excel
    alt Python y Excel cuadran
        L->>B: 3. En UNA transacción: anota los datos que cambiaron (correccion_dato)
        L->>B: borra septiembre y guarda la versión nueva (muestras, días, calidad)
        L->>B: detenciones: actualiza las que siguen, agrega las nuevas, quita las que desaparecieron
        L->>B: reemplaza la fila de septiembre en disponibilidad_mensual
        L->>B: 4. etl_run N° 15 = success, Publicada = sí, "2026-09"
    else No cuadran (parity_failed)
        L->>B: 4. etl_run N° 15 = parity_failed, Publicada = no
        Note over B: Septiembre 1–20 sigue tal cual:<br/>Power BI no ve nada a medias
    end
```

Si algo falla a mitad de la transacción, **nada** cambia: el mes anterior queda intacto.

---

## 3. El detalle de cada cajón (modelo entidad-relación)

Cómo leer las "patitas" de las líneas:

| Símbolo | Se lee |
|---|---|
| `\|\|` | exactamente uno |
| `o\|` | cero o uno |
| `o{` | cero o muchos |
| `\|{` | uno o muchos |
| línea punteada | relación de negocio (no es una llave directa en la base) |

### 3.1 El estado vigente (lo que lee Power BI)

Cada fila se identifica por **qué es** (proyecto, fecha, PCS), no por qué carga la escribió. Por eso al recargar se
reemplaza en vez de duplicarse. La columna `NumCorrida` solo dice qué ejecución la escribió por última vez.

```mermaid
%%{init: {'theme':'neutral'}}%%
erDiagram
    PROYECTO ||--o{ MUESTRA_PCS : "tiene"
    PROYECTO ||--o{ MUESTRA_PLANTA : "tiene"
    PROYECTO ||--o{ DETENCION : "tiene"
    PROYECTO ||--o{ DIA : "tiene"
    PROYECTO ||--o{ MES : "tiene"
    PROYECTO ||--o{ CALIDAD : "tiene"
    CATALOGO_FALLAS ||--o{ DETENCION : "clasifica"
    MES ||..o{ MUESTRA_PCS : "suma C14 de"
    MES ||..o{ DETENCION : "agrupa"

    PROYECTO {
        int IdProyecto PK
        string Nombre "Arena BESS"
        int NumPCS "61"
        int NumBateriasPorPCS "4"
        int NumRacksPorBAC "12"
        int MinutosMuestreo "15"
    }
    MUESTRA_PCS {
        int IdProyecto PK
        float SerialFecha PK "fecha exacta del Excel"
        int Ocurrencia PK "2 si la hora se repite"
        int NumeroPCS PK "1 a 61"
        int Anio
        int Mes
        int NumeroFilaOrigen "fila del Excel"
        float ModulosRaw "NUMBER_OF_MODULES"
        bool ModulosDisponiblesNulo "venia vacio"
        string FallaRaw "CURRENT FAULT"
        int ValorExclusion "0, 1 o 2"
        float BateriasIndisponibles "4 menos M"
        float ImpactoRackPonderado "aporte al C14"
        int NumCorrida "ultima carga"
    }
    MUESTRA_PLANTA {
        int IdProyecto PK
        float SerialFecha PK
        int Ocurrencia PK
        float FactorOperacionalRaw "solo para C21"
        float PorcentajeSOC
        string ComentarioExclusion "de Alex"
    }
    DETENCION {
        bigint IdDetencion PK "no cambia al recargar"
        int Anio
        int Mes
        int NumeroPCS
        datetime FechaInicio
        datetime FechaTermino
        int DuracionSegundos
        float DuracionHoras
        string CodigoFalla "F13, F55..."
        float HorasRackIndisponibles
        bool EsExcusable
        int NumCorrida
    }
    DIA {
        int IdProyecto PK
        datetime Dia PK
        float Disponibilidad "acumulada al dia"
        float Variacion
    }
    MES {
        int IdProyecto PK
        int Anio PK
        int Mes PK
        datetime UltimoDato "hasta donde llega"
        bool MesCompleto
        int BloquesMuestreo "C12"
        float BloquesRacksIndisponibles "C14"
        float DisponibilidadMensual "C16, el KPI"
        string EstadoExclusiones "sin o con"
        string Origen "corrida o excel_manual"
        int NumCorrida
    }
    CALIDAD {
        bigint IdCalidad PK
        int Anio
        int Mes
        string Tipo "hueco, cambio de hora..."
        string Severidad
        string Detalle
    }
    CATALOGO_FALLAS {
        int IdTipoDetencion PK
        string CodigoFalla
        string Significado
    }
```

### 3.2 El registro de ejecuciones (la historia)

Aquí nada se borra ni se reemplaza. Sirve para responder *¿quién cargó qué, cuándo, y qué cambió?*

```mermaid
%%{init: {'theme':'neutral'}}%%
erDiagram
    PROYECTO ||--o{ EJECUCION : "tiene"
    EJECUCION }o--o| REFERENCIA_EXCEL : "se compara con"
    EJECUCION ||--o{ RECONCILIACION : "resultado"
    EJECUCION ||--o{ CORRECCION : "detecta"
    PROYECTO ||--o{ CARGA_MATRIZ : "recibe cada mes"
    CARGA_MATRIZ |o..o| EJECUCION : "la usó el cierre"
    PROYECTO ||--o{ REVISION : "tiene"

    PROYECTO {
        int IdProyecto PK
    }
    EJECUCION {
        uuid IdCorrida PK "llave interna"
        int NumCorrida UK "1, 2, 3... para personas"
        string TipoCorrida "semanal, cierre_mensual..."
        datetime InicioPeriodo
        datetime FinPeriodo
        string ArchivoOrigen "con su sha256"
        string Estado "success, parity_failed, failed"
        bool Publicada "reemplazo el mes"
        string MesesPublicados "2026-09"
    }
    REFERENCIA_EXCEL {
        uuid IdReferencia PK
        string Corte
        int C12
        float C14
        float C16
    }
    RECONCILIACION {
        int Nivel "1 a 5"
        string Metrica "resumen o diferencia"
        float Delta
        bool Aprobado
    }
    CORRECCION {
        string TipoCorreccion "recarga, exclusion_matrix..."
        string Hoja
        int NumeroFilaOrigen
        int NumeroPCS
        string Campo "MODULES, PCS01..."
        string ValorAnterior
        string ValorNuevo
    }
    CARGA_MATRIZ {
        bigint IdCarga PK
        int Anio
        int Mes
        string ArchivoOrigen
        uuid IdCorridaCierre
    }
    REVISION {
        int NumeroPCS
        datetime FechaInicio
        string EstadoRevision
        string Observacion
        string RevisadoPor
    }
```

## 4. Cada concepto, en una línea

### 🗂️ 1. Maestros — lo que casi nunca cambia

| Concepto | Tabla | Qué guarda |
|---|---|---|
| Proyecto | `proyecto` | Cada planta con sus números: Arena = 61 PCS, 4 baterías, 12 racks, 15 min. (Hay 3 plantas más registradas, aún sin datos.) |
| Catálogo de fallas | `tipo_detencion` | Los códigos de falla de `PCS-Fault` (F0…F257) con su significado. |

### 📊 2. Estado vigente — una versión por mes

| Concepto | Tabla | Qué guarda |
|---|---|---|
| Bloque de un PCS | `muestra_pcs` | Una fila por **bloque de 15 minutos y PCS**: el dato crudo de SCADA (`NUMBER_OF_MODULES`, falla, estado), la exclusión aplicada y **cuánto aportó al C14**. |
| Bloque de la planta | `muestra_planta` | Una fila por bloque: lo de `PlantActivity` (solo influye si C21 = "Yes") y el comentario de la `Exclusion_Matrix`. |
| Detención | `detencion` | Cada evento de falla (`ListOfFaults`): PCS, inicio, fin, duración en **segundos** y en **horas**, código, horas-rack perdidas. Su `IdDetencion` **no cambia** al recargar. |
| Día | `disponibilidad_diaria` | La curva diaria (`Daily`): disponibilidad acumulada hasta cada día y su variación. |
| Mes | `disponibilidad_mensual` | **Una fila por mes** con los números clave: C12, C14, **C16 (el KPI)**, C19, L10, hasta qué dato llega, si está completo y si es "Sin" o "Con Exclusiones". Julio y agosto 2026 vienen del Excel (`excel_manual`). |
| Problema de calidad | `calidad_dato` | Cosas raras en los datos del mes: huecos, duplicados, cambio de hora, módulos vacíos, códigos desconocidos… |

### 🧾 3. Registro de ejecuciones — la historia

| Concepto | Tabla | Qué guarda |
|---|---|---|
| Ejecución | `etl_run` | **Una fila por cada carga**: período, tipo, archivo leído (con su huella sha256), cómo terminó, si **publicó** y qué meses. |
| Referencia Excel | `excel_reference_run` | Los KPI que calculó el Excel con sus macros (C12, C14, C16, C19). El detalle completo queda en `referencia_excel.json` del corte. |
| Reconciliación | `reconciliation_result` | La comparación Python vs Excel: un resumen por nivel y solo las diferencias que no cuadraron. |
| Corrección | `correccion_dato` | Cada dato que **cambió** respecto de la versión anterior del mes: al recargar (`recarga`) o al cargar la matriz (`exclusion_matrix`). Valor antes y después. |
| Carga de la matriz | `exclusion_matrix_carga` | Cada entrega mensual de Alex: qué archivo, cuándo y qué cierre la usó. |
| Revisión | `detencion_revision` | Lo que una persona anota al revisar una detención. Se conserva aunque el mes se recargue. |

Cada ejecución tiene **dos identificadores**:

| Campo | Ejemplo | Para qué |
|---|---|---|
| `NumCorrida` | `12` | Para **personas**: "revisa la ejecución 12". Correlativo 1, 2, 3… que asigna la base |
| `IdCorrida` | `4d763e71-a9d5-…` | Para el **sistema**: se crea en Python antes de tocar la base (logs, carpeta del corte) |

**Misael no necesita ninguno de los dos**: las tablas de estado ya tienen solo la versión vigente de cada mes.

---

## 5. Las reglas que el modelo siempre cumple

1. **Una sola versión por mes.** Recargar un mes lo reemplaza entero; nunca quedan filas repetidas.
2. **Solo se publica lo que cuadra.** Si Python y el Excel no coinciden (`parity_failed`), el mes vigente no se toca.
3. **Todo o nada.** El reemplazo de un mes ocurre en una sola transacción: Power BI nunca ve un mes a medias.
4. **Protecciones.** No se borra por accidente: una carga que llega a una fecha **anterior** a lo vigente, un mes
   importado del Excel (`excel_manual`) o un mes "Con Exclusiones" solo se reemplazan con una opción explícita
   (`--permitir-recorte`, `--reemplazar-manual`, `--forzar-sin-exclusiones`).
5. **La historia no se pierde.** El registro de ejecuciones nunca se borra, y cada dato que cambia entre cargas queda
   en `correccion_dato` con su valor anterior.
6. **Las revisiones sobreviven.** Una detención conserva su `IdDetencion` y su revisión aunque el mes se recargue.
7. **Los datos raros no se esconden**: los `NUMBER_OF_MODULES` vacíos se cuentan como disponibles (igual que el
   Excel), pero quedan marcados y listados.
8. **Fechas en `DATETIME`** y además el número de serie original del Excel, para no perder precisión.

## 6. Lo que ve Power BI

Power BI lee **directo las tablas de estado** (ya tienen solo lo vigente) y unas pocas vistas que suman o juntan:

```mermaid
%%{init: {'theme':'base'}}%%
flowchart LR
    M["disponibilidad_mensual"] --> V1{{"v_disponibilidad_anual<br/>acumulado del año"}}
    DT["detencion"] --> V2{{"v_resumen_codigo_mensual<br/>Pareto por código"}}
    DT --> V3{{"v_detencion<br/>+ última revisión"}}
    RV["detencion_revision"] --> V3
    D["disponibilidad_diaria"] --> PBI["📊 Power BI"]
    M --> PBI
    V1 & V2 & V3 --> PBI
    classDef v fill:#FFF3E0,stroke:#FB8C00,color:#4A2A00
    classDef t fill:#E8F5E9,stroke:#43A047,color:#1B3D1F
    class V1,V2,V3 v
    class M,DT,D,RV t
```

Vistas de control: `v_modulos_nulos` (bloques sin dato de módulos), `v_correccion_dato` (qué datos cambiaron y en
qué carga) y `v_ejecuciones` (el registro de cargas). Detalle para Power BI en [`pbi-handoff.md`](./pbi-handoff.md).

## 7. Ejemplo: seguir un número hacia atrás

*"¿Por qué septiembre dio C14 = 104.134,292?"*

1. `disponibilidad_mensual` (2026, 9) → **C14 = 104.134,292**, escrito por la ejecución `NumCorrida` 15.
2. `muestra_pcs` de septiembre → la suma de `ImpactoRackPonderado` de todos sus bloques y PCS da exactamente ese
   número. Se puede ver qué PCS y qué días aportaron más.
3. En esas mismas filas está el `NUMBER_OF_MODULES` que informó SCADA (`ModulosRaw`) y la fila del Excel
   (`NumeroFilaOrigen`).
4. `v_ejecuciones` con `NumCorrida = 15` dice qué archivo se leyó (con su sha256) y si cuadró con el Excel; y
   `v_correccion_dato` muestra qué datos habían cambiado respecto de la carga anterior.

Las consultas listas para hacer este recorrido están en `sql/07_audit_queries.sql`.
