# MODEL DE DADES — TORNS CONSORCI v3.0

## 1. Objectiu

Definir l’estructura de dades definitiva de l’aplicació de gestió de torns.

Principi central:

```text
PLANIFICACIÓ BASE
        +
INCIDÈNCIES
        =
PLANIFICACIÓ FINAL
```

La informació es gestionarà per dates, encara que després es mostri en vistes diàries, setmanals, mensuals o anuals. Excel actuarà com a capa de persistència mitjançant taules estructurades (`ListObject`).

## 2. Principis generals

- Cap formulari accedeix directament als fulls.
- Només `clsRepository` coneix els fulls i les taules Excel.
- Les referències a camps es fan pel nom de la columna, mai per número.
- Les vistes del calendari no són la font de veritat.
- Els colors són presentació; no formen part de les dades.
- Els canvis puntuals es registren com a incidències.
- La planificació base no es destrueix quan s’aplica una incidència.
- Tot canvi rellevant deixa traça a l’historial.
- Les dates es guarden com a dates reals d’Excel.
- Els identificadors són estables i no depenen de la posició de les files.

## 3. Taules del sistema

### 3.1 `tblOperaris`

Full recomanat: `OPERARIS`

| Camp | Tipus | Obligatori | Exemple | Descripció |
|---|---|:---:|---|---|
| `OperariId` | Text | Sí | `OP001` | Identificador únic i estable |
| `Nom` | Text | Sí | `Josep Galera` | Nom complet |
| `Actiu` | Boolean | Sí | `TRUE` | Participa en la planificació |
| `DataAlta` | Date | No | `01/01/2020` | Alta al sistema |
| `DataBaixa` | Date | No |  | Baixa definitiva |
| `CandidatIntensiu` | Boolean | Sí | `TRUE` | Pot fer intensius |
| `CandidatGuardia` | Boolean | Sí | `TRUE` | Pot fer guàrdies |
| `IntensiveCount` | Long | Sí | `6` | Intensius realitzats |
| `LastIntensiveDate` | Date | No | `07/01/2027` | Últim intensiu realitzat |
| `RotationState` | Text | Sí | `MATI` | Estat actual del cicle |
| `Observacions` | Text | No |  | Informació complementària |

Regles:

- `OperariId` no es pot repetir.
- Els operaris inactius es conserven per mantenir l’historial.
- `DataBaixa` no equival a una baixa mèdica temporal.

### 3.2 `tblPlanificacioBase`

Full recomanat: `PLANIFICACIO_BASE`

La granularitat és diària.

| Camp | Tipus | Obligatori | Exemple | Descripció |
|---|---|:---:|---|---|
| `AssignacioId` | Text | Sí | `A-2027-000001` | Identificador únic |
| `VersioId` | Text | Sí | `V2027-DRAFT-01` | Versió de planificació |
| `OperariId` | Text | Sí | `OP001` | Operari assignat |
| `Data` | Date | Sí | `07/01/2027` | Dia de l’assignació |
| `TornBase` | Text | Sí | `I` | Torn generat pel motor |
| `EsGuardiaPrincipal` | Boolean | Sí | `TRUE` | Guàrdia principal |
| `EsGuardiaReforc` | Boolean | Sí | `FALSE` | Guàrdia de reforç |
| `Origen` | Text | Sí | `MOTOR` | Origen de l’assignació |
| `DataCreacio` | DateTime | Sí |  | Data de generació |
| `UsuariCreacio` | Text | Sí | `Marina` | Usuari que genera |

Regles:

- Clau funcional: `VersioId + OperariId + Data`.
- Només una assignació base per operari i dia dins d’una versió.
- El motor treballa sempre per dates.

### 3.3 `tblIncidencies`

Full recomanat: `INCIDENCIES`

| Camp | Tipus | Obligatori | Exemple | Descripció |
|---|---|:---:|---|---|
| `IncidenciaId` | Text | Sí | `INC-2027-00001` | Identificador únic |
| `OperariId` | Text | Sí | `OP004` | Operari afectat |
| `DataInici` | Date | Sí | `15/09/2027` | Inici |
| `DataFi` | Date | Sí | `18/09/2027` | Final |
| `Tipus` | Text | Sí | `BAIXA` | Tipus d’incidència |
| `TornResultant` | Text | No | `B` | Torn visible final |
| `OperariSubstitutId` | Text | No | `OP008` | Operari substitut |
| `Motiu` | Text | Sí | `Baixa mèdica` | Justificació |
| `EsManual` | Boolean | Sí | `TRUE` | Canvi manual |
| `Estat` | Text | Sí | `ACTIVA` | Activa, anul·lada o tancada |
| `DataCreacio` | DateTime | Sí |  | Alta de la incidència |
| `UsuariCreacio` | Text | Sí | `Marina` | Usuari responsable |
| `DataModificacio` | DateTime | No |  | Última modificació |
| `UsuariModificacio` | Text | No |  | Usuari modificador |

Tipus inicials:

- `BAIXA`
- `VACANCES`
- `PERMIS`
- `CANVI_TORN`
- `SUBSTITUCIO`
- `DESCANS`
- `ALTRES`

Regles:

- `DataFi` no pot ser anterior a `DataInici`.
- Una incidència no elimina la planificació base.
- Les incidències anul·lades es conserven.
- Una substitució pot afectar un sol dia o un interval.

### 3.4 `tblGuardies`

Full recomanat: `GUARDIES`

| Camp | Tipus | Obligatori | Exemple | Descripció |
|---|---|:---:|---|---|
| `GuardiaId` | Text | Sí | `G-2027-0012` | Identificador únic |
| `VersioId` | Text | Sí | `V2027-PUB-01` | Versió associada |
| `OperariId` | Text | Sí | `OP007` | Operari de guàrdia |
| `DataInici` | Date | Sí | `07/01/2027` | Inici |
| `DataFi` | Date | Sí | `10/01/2027` | Final |
| `Tipus` | Text | Sí | `PRINCIPAL` | Principal o reforç |
| `Origen` | Text | Sí | `MOTOR` | Motor o manual |
| `Motiu` | Text | No |  | Observacions |
| `Activa` | Boolean | Sí | `TRUE` | Vigència |
| `DataCreacio` | DateTime | Sí |  | Traçabilitat |
| `UsuariCreacio` | Text | Sí |  | Traçabilitat |

### 3.5 `tblVersionsPlanificacio`

Full recomanat: `VERSIONS`

| Camp | Tipus | Obligatori | Exemple | Descripció |
|---|---|:---:|---|---|
| `VersioId` | Text | Sí | `V2027-PUB-01` | Identificador únic |
| `Any` | Long | Sí | `2027` | Any |
| `NumeroVersio` | Long | Sí | `1` | Número seqüencial |
| `Estat` | Text | Sí | `PUBLICADA` | Esborrany, publicada o històrica |
| `DataCreacio` | DateTime | Sí |  | Creació |
| `UsuariCreacio` | Text | Sí |  | Creador |
| `DataPublicacio` | DateTime | No |  | Publicació |
| `UsuariPublicacio` | Text | No |  | Publicador |
| `Observacions` | Text | No |  | Notes |

Regles:

- Només una versió publicada per any.
- Una versió publicada no es modifica directament.
- Les versions històriques són immutables.

### 3.6 `tblConfiguracio`

Full recomanat: `CONFIGURACIO`

| Camp | Tipus | Obligatori | Exemple | Descripció |
|---|---|:---:|---|---|
| `Clau` | Text | Sí | `MorningWeeks` | Nom del paràmetre |
| `Valor` | Text | Sí | `2` | Valor serialitzat |
| `TipusDada` | Text | Sí | `LONG` | Tipus esperat |
| `Descripcio` | Text | No |  | Ajuda funcional |
| `Actiu` | Boolean | Sí | `TRUE` | Vigència |

Paràmetres inicials:

- `MorningWeeks`
- `AfternoonWeeks`
- `IntensiveStartWeekday`
- `IntensiveEndWeekday`
- `DutyBackupCount`
- `MorningStartTime`
- `MorningEndTime`
- `AfternoonStartTime`
- `AfternoonEndTime`
- `PlanningYear`
- `ApplicationEnvironment`

### 3.7 `tblHistorial`

Full recomanat: `HISTORIAL`

| Camp | Tipus | Obligatori | Exemple | Descripció |
|---|---|:---:|---|---|
| `HistorialId` | Text | Sí | `H-2027-000012` | Identificador únic |
| `DataHora` | DateTime | Sí |  | Moment de l’acció |
| `Usuari` | Text | Sí | `Marina` | Usuari |
| `Entitat` | Text | Sí | `INCIDENCIA` | Tipus d’objecte |
| `EntitatId` | Text | Sí | `INC-2027-00001` | Identificador afectat |
| `Accio` | Text | Sí | `CREAR` | Acció realitzada |
| `ValorAnterior` | Text | No | `M` | Estat anterior |
| `ValorNou` | Text | No | `B` | Estat nou |
| `Motiu` | Text | No | `Baixa mèdica` | Justificació |
| `Detall` | Text | No |  | Informació addicional |

### 3.8 `tblParametresTorns`

Full recomanat: `PARAMETRES`

| Camp | Tipus | Obligatori | Exemple |
|---|---|:---:|---|
| `Codi` | Text | Sí | `M` |
| `Descripcio` | Text | Sí | `Matí` |
| `Prioritat` | Long | Sí | `10` |
| `Actiu` | Boolean | Sí | `TRUE` |

Codis inicials:

| Codi | Significat |
|---|---|
| `M` | Matí |
| `T` | Tarda |
| `I` | Intensiu |
| `D` | Descans |
| `B` | Baixa |
| `V` | Vacances |
| `P` | Permís |

Les guàrdies es gestionen a `tblGuardies`, no com a substitut del torn.

## 4. Relacions

```text
tblOperaris
   │
   ├──< tblPlanificacioBase
   ├──< tblIncidencies
   └──< tblGuardies

tblVersionsPlanificacio
   │
   ├──< tblPlanificacioBase
   └──< tblGuardies

tblIncidencies
   └── OperariSubstitutId → tblOperaris.OperariId

Totes les operacions rellevants
   └──< tblHistorial
```

## 5. Resolució del torn final

Per a un operari i una data:

1. Cercar la planificació base publicada.
2. Obtenir el `TornBase`.
3. Cercar incidències actives que cobreixin la data.
4. Aplicar la incidència de més prioritat.
5. Consultar les guàrdies per separat.
6. Retornar el torn final i les metadades.

Exemple:

```text
Base: M
Incidència: BAIXA
Resultat final: B
```

Exemple de substitució:

```text
Operari original
Base: I
Incidència: BAIXA
Resultat: B

Operari substitut
Base: T
Incidència: SUBSTITUCIO
Resultat: D de dilluns a dimecres
Resultat: I de dijous a diumenge
Calendari posterior: sense canvis
```

## 6. Vistes derivades

Les vistes no guarden dades pròpies:

- vista diària;
- vista setmanal;
- vista mensual;
- vista anual;
- fitxa d’operari.

## 7. Validacions mínimes

- Operari existent.
- Data dins del període de planificació.
- Codi de torn vàlid.
- Duplicats no permesos a la planificació base.
- Incidències amb interval coherent.
- Solapaments detectats.
- Substitut actiu.
- Versió publicada existent.
- Historial creat després de cada canvi.
- Cap modificació directa d’una versió publicada.

## 8. Criteris d’acceptació

El model queda aprovat quan:

- permet canvis d’un sol dia;
- permet canvis per interval;
- manté la planificació original;
- permet substitucions sense recalcular les setmanes posteriors;
- genera vistes diàries, setmanals, mensuals i anuals;
- manté l’historial;
- admet esborranys i publicacions;
- no depèn dels colors ni de la posició de les columnes;
- pot ser llegit exclusivament per `clsRepository`;
- pot ser consumit per `clsCalendarService`.
