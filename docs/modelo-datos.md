# Modelo conceptual de datos

Para: cualquiera que quiera entender **qué se guarda en la base de datos y por qué**, sin leer SQL.
Las palabras técnicas están en el [glosario](./glosario.md). El detalle exacto de cada tabla está en `sql/02_corrida.sql`.

> **La idea en una frase:** cada vez que se calcula la disponibilidad se crea una **corrida**. Todo lo que entra
> (datos de SCADA, exclusiones), todo lo que sale (KPI, eventos, días, meses) y todos los controles (comparación con
> el Excel, problemas de calidad) quedan **colgando de esa corrida**, y nunca se borran.

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
        TAB["Tablas<br/>(todo, por corrida)"]
        VIS["Vistas v_*_vigente<br/>(solo lo oficial)"]
    end
    PBI["📊 Power BI<br/>(Misael)"]

    PCS -- "cada 15 min" --> EXP
    EXP -- "TeamViewer" --> HOJAS
    EM --> HOJAS
    HOJAS --> MAC
    HOJAS --> CALC
    MAC --> REC
    CALC --> REC
    CALC --> TAB
    REC --> TAB
    TAB --> VIS
    VIS --> PBI

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
    class TAB,VIS bd
    class PBI bi
```

---

## 2. El mapa general: cinco cajones

Piensa en la base de datos como **cinco cajones**. En el centro está la **corrida**: todo lo demás se cuelga de ella.
Cada caja agrupa **conceptos** ("cosas" que guardamos); en letra chica, las tablas donde viven.

```mermaid
%%{init: {'theme':'base','themeVariables':{'fontSize':'15px'}}}%%
flowchart TB
    subgraph M["🗂️ 1. Maestros — casi nunca cambian"]
        direction LR
        PR["Proyecto<br/><small>proyecto</small>"]
        CF["Catálogo de fallas<br/><small>tipo_detencion</small>"]
    end
    CO(["🧾 2. CORRIDA<br/><small>etl_run</small><br/>una por cada cálculo"])
    subgraph E["📥 3. Lo que entra"]
        direction TB
        E1["Muestras de SCADA<br/><small>raw_pcs_sample · raw_pcs_column_map</small>"]
        E2["Actividad y exclusiones<br/><small>plant_activity_sample · exclusion_matrix_sample</small>"]
    end
    subgraph S["📤 4. Lo que sale"]
        direction TB
        S1["Resultado por bloque y KPI<br/><small>availability_sample_result · availability_run_result</small>"]
        S2["Eventos y detenciones<br/><small>fault_event · fault_code_summary · detencion · detencion_revision</small>"]
        S3["Día, mes y año<br/><small>daily_availability · monthly_official_kpi · annual_availability</small>"]
    end
    subgraph K["✅ 5. Los controles"]
        direction TB
        K1["Comparación con Excel<br/><small>excel_reference_* · reconciliation_result</small>"]
        K2["Calidad y cambios<br/><small>data_quality_issue · correccion_dato · exclusion_matrix_carga</small>"]
    end

    PR -- "1 proyecto : muchas corridas" --> CO
    CF -. "clasifica las detenciones" .-> S2
    CO -- "lee" --> E
    CO -- "calcula" --> S
    CO -- "se controla con" --> K

    classDef m fill:#E3F2FD,stroke:#1E88E5,color:#0D2B45
    classDef c fill:#FFF4D6,stroke:#C9A227,color:#3D3000,stroke-width:3px
    classDef e fill:#E8F5E9,stroke:#43A047,color:#1B3D1F
    classDef s fill:#F3E5F5,stroke:#8E24AA,color:#3A1245
    classDef k fill:#FFEBEE,stroke:#E53935,color:#4A1212
    class PR,CF m
    class CO c
    class E1,E2 e
    class S1,S2,S3 s
    class K1,K2 k
```

## 3. El detalle de cada cajón (modelo entidad-relación)

Estos diagramas muestran los datos principales de cada concepto. Cómo leer las "patitas" de las líneas:

| Símbolo | Se lee |
|---|---|
| `\|\|` | exactamente uno |
| `o\|` | cero o uno |
| `o{` | cero o muchos |
| `\|{` | uno o muchos |
| línea punteada | relación de negocio (no es una llave directa en la base) |

### 3.1 La corrida y lo que entra

Una corrida lee el libro Excel y guarda **lo que vio**: cada muestra de SCADA, la actividad de planta y las marcas de
exclusión, siempre con la fila del Excel de donde salió (`NumeroFilaOrigen`).

```mermaid
%%{init: {'theme':'neutral'}}%%
erDiagram
    PROYECTO ||--o{ CORRIDA : "tiene"
    PROYECTO ||--o{ CARGA_MATRIZ : "recibe cada mes"
    CORRIDA ||--o{ MUESTRA_SCADA : "lee"
    CORRIDA ||--o{ MAPA_COLUMNAS : "reconoce"
    CORRIDA ||--o{ ACTIVIDAD_PLANTA : "lee"
    CORRIDA ||--o{ MARCA_EXCLUSION : "lee"
    CARGA_MATRIZ ||..o{ MARCA_EXCLUSION : "trae"

    PROYECTO {
        int IdProyecto PK
        string Nombre "Arena BESS"
        int NumPCS "61"
        int NumBateriasPorPCS "4"
        int NumRacksPorBAC "12"
        int MinutosMuestreo "15"
    }
    CORRIDA {
        uuid IdCorrida PK "llave interna"
        int NumCorrida UK "1, 2, 3... para personas"
        string TipoCorrida "semanal o cierre_mensual"
        bool EsOficial
        string EstadoExclusiones "sin o con"
        datetime InicioPeriodo "C5"
        datetime FinPeriodo "C7"
        string ArchivoOrigen "con su sha256"
        string Estado "success o parity_failed"
    }
    MUESTRA_SCADA {
        int NumeroFilaOrigen "fila del Excel"
        int NumeroPCS "1 a 61"
        float SerialFechaExcelOrigen "fecha exacta"
        float ModulosDisponibles "NUMBER_OF_MODULES"
        bool ModulosDisponiblesNulo "venia vacio"
        string FallaRaw "CURRENT FAULT"
    }
    MAPA_COLUMNAS {
        int NumeroPCS
        string Campo "FAULT, MODULES..."
        int NumeroColumna
    }
    ACTIVIDAD_PLANTA {
        int NumeroFilaOrigen
        float FactorOperacionalRaw "solo para C21"
        float PorcentajeSOC
    }
    MARCA_EXCLUSION {
        int NumeroFilaOrigen
        int NumeroPCS
        float ValorExclusion "1 o 2"
        float BateriasPrevias "para el 2"
        string Comentario "causa, de Alex"
    }
    CARGA_MATRIZ {
        bigint IdCarga PK
        int Anio
        int Mes
        string ArchivoOrigen
        string Sha256Archivo
        uuid IdCorridaCierre "cierre que la usó"
    }
```

### 3.2 Lo que sale

Con esos datos la corrida calcula el resultado de cada bloque, el KPI, los eventos y las tablas diaria, mensual y
anual.

```mermaid
%%{init: {'theme':'neutral'}}%%
erDiagram
    CORRIDA ||--o{ RESULTADO_MUESTRA : "calcula"
    CORRIDA ||--|| KPI_CORRIDA : "resume en"
    CORRIDA ||--o{ EVENTO_FALLA : "detecta"
    CORRIDA ||--o{ RESUMEN_CODIGO : "agrupa"
    CORRIDA ||--o{ DIA : "calcula"
    CORRIDA |o--o| KPI_MENSUAL : "registra si es oficial"
    CORRIDA ||--o{ ACUMULADO_ANUAL : "calcula"
    CORRIDA ||--o{ DETENCION : "publica"
    CATALOGO_FALLAS ||--o{ DETENCION : "clasifica"
    DETENCION ||..o{ REVISION : "se revisa en"

    CORRIDA {
        uuid IdCorrida PK
    }
    RESULTADO_MUESTRA {
        int NumeroPCS
        datetime MarcaTiempoMuestra
        float BateriasIndisponibles "4 menos M"
        float ValorExclusion "0, 1 o 2"
        float ImpactoRackPonderado "aporte al C14"
    }
    KPI_CORRIDA {
        int BloquesMuestreo "C12"
        float BloquesRacksIndisponibles "C14"
        float DisponibilidadPeriodo "C16, el KPI"
        float DisponibilidadAnualAcumulada "C19"
        float HorasRackEventos "L10"
    }
    EVENTO_FALLA {
        int NumeroPCS
        datetime MarcaTiempoInicio
        datetime MarcaTiempoFin
        float DuracionHoras
        string CodigoFalla "F13, F55..."
        float HorasRackIndisponibles
        bool TieneExclusion
    }
    RESUMEN_CODIGO {
        int Ranking
        string CodigoFalla
        float HorasRackIndisponibles
        float Porcentaje
    }
    DIA {
        int DiaN "1 a 31"
        datetime Dia
        float Disponibilidad "acumulada al dia"
        float Variacion
    }
    KPI_MENSUAL {
        int Anio
        int Mes
        float BloquesRacksIndisponibles
        string Origen "corrida o excel_manual"
    }
    ACUMULADO_ANUAL {
        int Mes
        float DisponibilidadMensual
        float DisponibilidadAcumulada "KPI del año"
        float DisponibilidadContractual "0,98"
    }
    CATALOGO_FALLAS {
        int IdTipoDetencion PK
        string CodigoFalla
        string Significado
    }
    DETENCION {
        bigint IdDetencion PK
        int NumeroPCS
        datetime FechaInicio
        datetime FechaTermino
        int DuracionSegundos
        float DuracionHoras "misma duración en horas"
        bool EsExcusable
    }
    REVISION {
        int NumeroPCS
        datetime FechaInicio
        string EstadoRevision
        string Observacion
        string RevisadoPor
    }
```

### 3.3 Los controles

Cada corrida se compara con el Excel y deja anotado todo lo raro y todo lo que se corrigió.

```mermaid
%%{init: {'theme':'neutral'}}%%
erDiagram
    CORRIDA }o--o| REFERENCIA_EXCEL : "se compara con"
    CORRIDA ||--o{ RECONCILIACION : "resultado"
    REFERENCIA_EXCEL ||--o{ RECONCILIACION : "aporta el valor Excel"
    CORRIDA ||--o{ PROBLEMA_CALIDAD : "anota"
    CORRIDA ||--o{ CORRECCION : "aplica"
    CARGA_MATRIZ ||..o{ CORRECCION : "genera"

    CORRIDA {
        uuid IdCorrida PK
    }
    REFERENCIA_EXCEL {
        uuid IdReferencia PK
        string Corte
        int C12
        float C14
        float C16
        string HashLibro "huella del libro"
    }
    RECONCILIACION {
        int Nivel "1 a 5"
        string Metrica "C14, evento..."
        string ValorPython
        string ValorExcel
        float Delta "diferencia"
        bool Aprobado
    }
    PROBLEMA_CALIDAD {
        string Tipo "hueco, pa_vacio..."
        string Severidad "info, advertencia, error"
        int NumeroPCS
        string Detalle
    }
    CORRECCION {
        string TipoCorreccion "exclusion_matrix..."
        int NumeroFilaOrigen
        int NumeroPCS
        string Campo "PCS01, Comments..."
        string ValorAnterior
        string ValorNuevo
    }
    CARGA_MATRIZ {
        bigint IdCarga PK
        int Mes
    }
```

## 4. Cada concepto, en una línea

### 🗂️ 1. Maestros — lo que casi nunca cambia

| Concepto | Tabla | Qué guarda |
|---|---|---|
| Proyecto | `proyecto` | Cada planta con sus números: Arena = 61 PCS, 4 baterías, 12 racks, 15 min. (Hay 3 plantas más registradas, aún sin datos.) |
| Catálogo de fallas | `tipo_detencion` | Los códigos de falla de `PCS-Fault` (F0…F257) con su significado. |

### 🧾 2. La corrida — el centro de todo

| Concepto | Tabla | Qué guarda |
|---|---|---|
| Corrida | `etl_run` | **Un registro por cada cálculo**: qué período, qué tipo (semanal / cierre mensual), si es oficial, si es "Sin" o "Con Exclusiones", los interruptores usados (C21, C31, L14), qué archivo se leyó (con su huella sha256), si terminó bien y qué versión del algoritmo se usó. |

Todo lo demás se "cuelga" de la corrida con su `IdCorrida`. Así se puede responder siempre: *¿de dónde salió este
número?*

Cada corrida tiene **dos identificadores**:

| Campo | Ejemplo | Para qué |
|---|---|---|
| `NumCorrida` | `12` | Para **personas**: "revisa la corrida 12". Correlativo 1, 2, 3… que asigna la base al guardar |
| `IdCorrida` | `4d763e71-a9d5-…` | Para el **sistema**: la llave que une todas las tablas |

¿Por qué no basta con el número? `IdCorrida` se crea en Python **antes** de tocar la base (lo usan los logs, la
carpeta del corte y los ensayos sin base de datos), y nunca se repite aunque haya varios computadores y dos bases
(PROD y TEST/QA). Un número, en cambio, solo existe después de guardar y se repite entre bases: la corrida 12 de
TEST/QA no es la corrida 12 de PROD.

### 📥 3. Lo que entra

| Concepto | Tabla | Qué guarda |
|---|---|---|
| Muestra de SCADA | `raw_pcs_sample` | Una fila por **bloque de 15 minutos y PCS**: el `NUMBER_OF_MODULES` tal como vino, el código de falla, el estado, y de qué fila del Excel salió. |
| Mapa de columnas | `raw_pcs_column_map` | En qué columna del Excel estaba cada dato de cada PCS (para auditar). |
| Actividad de planta | `plant_activity_sample` | Lo de `PlantActivity`. Solo influye si C21 = "Yes". |
| Marca de exclusión | `exclusion_matrix_sample` | Solo las celdas marcadas con **1 o 2** en la `Exclusion_Matrix`, con el comentario de Alex. |
| Carga de la matriz | `exclusion_matrix_carga` | Cada vez que Alex entrega la matriz de un mes: qué archivo, cuándo y qué cierre mensual la usó (`IdCorridaCierre`). |

### 📤 4. Lo que sale

| Concepto | Tabla | Qué guarda |
|---|---|---|
| Resultado por muestra | `availability_sample_result` | Para cada bloque y PCS: cuántas baterías faltaban, si había exclusión y **cuánto aportó al C14**. |
| KPI de la corrida | `availability_run_result` | **Una fila por corrida** con los números clave: C12, C14, **C16 (el KPI)**, C19 y L10. |
| Evento de falla | `fault_event` | La lista de eventos (`ListOfFaults`): PCS, inicio, fin, duración, código, horas-rack perdidas. |
| Resumen por código | `fault_code_summary` | Horas-rack perdidas por cada código de falla, ordenadas (el "Pareto"). |
| Día | `daily_availability` | La curva diaria (`Daily`): disponibilidad acumulada hasta cada día y su variación. |
| KPI mensual | `monthly_official_kpi` | El número oficial de cada mes. Julio y agosto 2026 vienen del Excel (`excel_manual`). |
| Acumulado anual | `annual_availability` | La tabla `Annual_AVA`: cada mes y la **disponibilidad acumulada del año** contra el 98 % contractual. |
| Detención | `detencion` | Los eventos listos para operación, con su tipo del catálogo y la duración en **segundos** (`DuracionSegundos`) y en **horas** (`DuracionHoras`). |
| Revisión | `detencion_revision` | Lo que una persona anota al revisar una detención (estado, observación, quién y cuándo). Se conserva aunque el mes se recalcule. |

### ✅ 5. Los controles

| Concepto | Tabla | Qué guarda |
|---|---|---|
| Referencia Excel | `excel_reference_run` (+ `_sample`, `_fault_event`, `_daily`) | Lo que calculó el Excel con sus macros para el mismo período. |
| Reconciliación | `reconciliation_result` | La comparación Python vs Excel en 5 niveles, con la diferencia (`Delta`) de cada número. |
| Problema de calidad | `data_quality_issue` | Cosas raras en los datos: huecos, duplicados, cambio de hora, módulos vacíos, códigos desconocidos… |
| Corrección | `correccion_dato` | Cada celda que cambió por una carga de la matriz o una corrección: valor antes y después. |

---

## 5. Las reglas que el modelo siempre cumple

1. **Nunca se borra nada** (*append-only*). Si algo se recalcula, se crea una corrida nueva y la anterior queda como historia.
2. **Todo tiene dueño**: cada dato calculado apunta a su corrida (`IdCorrida`), y cada corrida a su archivo de origen (con su sha256).
3. **Se puede auditar hacia atrás**: KPI → aporte de cada bloque → dato crudo de SCADA → fila y columna del Excel.
4. **Una sola verdad por mes para Power BI**: las vistas eligen la corrida vigente (oficial, terminada bien, y
   "Con Exclusiones" si existe).
5. **Los datos raros no se esconden**: los `NUMBER_OF_MODULES` vacíos se cuentan como disponibles (igual que el
   Excel), pero quedan marcados y listados.
6. **Fechas en `DATETIME`** y además el número de serie original del Excel, para no perder precisión.

## 6. Lo que ve Power BI: las vistas

Power BI **no lee las tablas**: lee **vistas**, que ya eligen lo oficial. Así el reporte nunca se equivoca de corrida.

```mermaid
%%{init: {'theme':'base'}}%%
flowchart LR
    R["etl_run<br/>(todas las corridas)"] --> V0{{"v_corrida_oficial_vigente<br/>1 corrida por mes"}}
    V0 --> V1["v_kpi_vigente"]
    V0 --> V2["v_daily_vigente"]
    V0 --> V3["v_fault_event_vigente"]
    V0 --> V4["v_fault_code_vigente"]
    V0 --> V5["v_detencion_vigente"]
    V0 --> V6["v_annual_vigente"]
    MK["monthly_official_kpi"] --> V7["v_monthly_kpi_vigente"]
    V1 & V2 & V3 & V4 & V5 & V6 & V7 --> PBI["📊 Power BI"]
    classDef v fill:#FFF3E0,stroke:#FB8C00,color:#4A2A00
    classDef t fill:#FFEBEE,stroke:#E53935,color:#4A1212
    class V0,V1,V2,V3,V4,V5,V6,V7 v
    class R,MK t
```

Además hay vistas de control: `v_calidad_corrida` (salud de cada corrida), `v_modulos_nulos_historico`
(bloques sin dato de módulos) y `v_correccion_dato` (qué celdas cambiaron). Detalle para Power BI en
[`pbi-handoff.md`](./pbi-handoff.md).

## 7. Ejemplo: seguir un número hacia atrás

*"¿Por qué septiembre 1–21 dio C14 = 104.134,292?"*

1. `availability_run_result` de la corrida vigente de septiembre → **C14 = 104.134,292**.
2. `availability_sample_result` de esa corrida → la suma de `ImpactoRackPonderado` de todos sus bloques y PCS da
   exactamente ese número. Se puede ver qué PCS y qué días aportaron más.
3. `raw_pcs_sample` → para cualquiera de esos bloques, el `NUMBER_OF_MODULES` que informó SCADA.
4. `NumeroFilaOrigen` y `raw_pcs_column_map` → la fila y la columna exactas del libro Excel, y `etl_run` dice qué
   archivo era (con su sha256).

Las consultas listas para hacer este recorrido están en `sql/07_audit_queries.sql`.
