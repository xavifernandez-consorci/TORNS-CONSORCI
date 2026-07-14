# TORNS CONSORCI

## 1. Objectiu

Aplicació Excel VBA per planificar i gestionar els torns del Consorci del Bages per a la Gestió de Residus.

L'aplicació ha de:

- generar una proposta de planificació;
- mantenir separades la lògica de negoci, la persistència i la interfície;
- permetre substitucions manuals sense recalcular automàticament la resta del calendari;
- conservar la traçabilitat dels canvis;
- facilitar l'evolució futura de les regles de servei.

El motor no modifica automàticament el calendari persistent. Els canvis s'han de confirmar explícitament.

---

## 2. Arquitectura

### Workbook

- `ThisWorkbook.cls`

Responsabilitat:

- gestionar l'obertura i el tancament del llibre;
- delegar la inicialització i el tancament al bootstrap.

### Application

- `modApplicationBootstrap.bas`
- `modConfiguration.bas`

Responsabilitats:

- inicialitzar l'aplicació;
- carregar i exposar la configuració activa;
- coordinar el tancament;
- centralitzar els errors d'inici i finalització.

### Domain

- `clsConfiguracio.cls`
- `clsOperari.cls`
- `clsScheduleContext.cls`
- `clsShiftAssignment.cls`

Responsabilitats:

- representar la configuració;
- representar els operaris;
- mantenir el context de planificació;
- mantenir i indexar les assignacions;
- registrar substitucions manuals.

### Services

- `modMotorRotacions.bas`
- `modGuardies.bas`
- `modSubstitucions.bas`

Responsabilitats:

- generar la planificació;
- gestionar guàrdies i reforços;
- aplicar canvis manuals sense recalcular la resta del calendari.

### Data

- `modDataAccess.bas`

Responsabilitat:

- aïllar l'accés als fulls Excel;
- carregar i desar configuració, operaris i assignacions.

### Tests

- `modTests.bas`

Responsabilitat:

- executar proves internes del domini i del motor;
- detectar regressions abans de cada commit.

### Presentation

- `frmInici.frm`
- `frmConfiguration.frm`
- `frmPlanificacio.frm`

Responsabilitat:

- oferir la interfície d'usuari;
- delegar les operacions als serveis;
- no contenir lògica de negoci.

---

## 3. Principis de disseny

- Cap lògica de negoci dins dels formularis.
- Cap accés directe a fulls des del motor.
- Cap persistència automàtica.
- Cap recalculació global per una incidència puntual.
- Els valors configurables no s'han de codificar directament al motor.
- Els canvis manuals han de quedar identificats.
- Cada component ha de tenir una única responsabilitat.
- Cada canvi funcional ha de compilar i superar els tests abans del commit.

---

## 4. Regles de negoci confirmades

### BR-001 — Dotació habitual

Actualment hi ha 14 operaris.

Aquest valor és configurable i no s'ha de codificar com una constant del motor.

### BR-002 — Distribució ordinària

Excloent l'operari assignat a intensiu:

- aproximadament un 60% dels operaris treballen de matí;
- aproximadament un 40% treballen de tarda.

Amb 14 operaris:

- 1 operari queda afectat pel cicle d'intensiu;
- els 13 restants es distribueixen habitualment en 8 de matí i 5 de tarda.

### BR-003 — Cicle base

El cicle ordinari és:

- 2 setmanes de matí;
- 2 setmanes de tarda.

### BR-004 — Setmana d'intensiu planificat

Per a l'operari seleccionat:

- dilluns: descans;
- dimarts: descans;
- dimecres: descans;
- dijous: intensiu;
- divendres: intensiu;
- dissabte: intensiu;
- diumenge: intensiu.

La resta d'operaris mantenen el torn base que ja tenien planificat.

### BR-005 — Seqüència posterior a l'intensiu planificat

Després del diumenge d'intensiu:

- 1 setmana de tarda;
- 2 setmanes de matí;
- 2 setmanes de tarda;
- continuació del cicle fins al següent intensiu.

### BR-006 — Selecció equilibrada

El següent operari d'intensiu es determina així:

1. només operaris actius;
2. només candidats a intensiu;
3. menor nombre d'intensius acumulats;
4. en cas d'empat, qui fa més temps que no en fa;
5. si l'empat continua, ordre estable de la col·lecció.

### BR-007 — Guàrdia principal

L'operari d'intensiu és també l'operari de guàrdia principal.

### BR-008 — Reforços

El nombre de reforços és configurable.

La seva selecció correspon a `modGuardies`.

### BR-009 — Substitució manual per baixa

Quan l'operari planificat per a l'intensiu està de baixa des del dilluns:

#### Operari original

- dilluns a diumenge: `BAIXA`.

#### Operari substitut

- dilluns a dimecres: `DESCANS`;
- dijous a diumenge: `INTENSIU`.

#### Calendari posterior

- a partir del dilluns següent, el substitut recupera el calendari que ja tenia;
- no se li aplica la seqüència posterior ordinària;
- no es recalcula cap setmana posterior;
- la resta del calendari queda intacta.

### BR-010 — Historial realitzat

Quan hi ha una substitució:

- l'operari original no incrementa el recompte d'intensius;
- l'operari substitut incrementa `IntensiveCount`;
- l'operari substitut actualitza `LastIntensiveDate`;
- aquest canvi d'historial no altera la planificació futura.

### BR-011 — Canvis manuals

Cada assignació modificada manualment ha de conservar:

- `IsManualOverride`;
- `OriginalEmployeeId`;
- `OverrideReason`.

---

## 5. Codis de torn actuals

| Codi | Significat |
|---|---|
| `M` | Matí |
| `T` | Tarda |
| `I` | Intensiu |
| `D` | Descans |
| `B` | Baixa |

Aquests codis s'han de traslladar progressivament a configuració.

---

## 6. Estat actual

### Implementat

- estructura Git i repositori remot;
- cicle de vida del workbook;
- bootstrap de l'aplicació;
- gestió central de configuració;
- contractes de persistència;
- model de configuració;
- model d'operari;
- context de planificació indexat;
- model d'assignació;
- metadades de substitució manual;
- motor base de matí i tarda;
- detecció dels dijous;
- selecció equilibrada d'intensius;
- creació del bloc intensiu;
- actualització de l'historial d'intensius;
- servei de substitució manual per baixa;
- infraestructura de proves;
- projecte VBA compilant.

### Tests

Últim resultat confirmat:

- 7 tests executats;
- 7 tests correctes;
- 0 tests fallits.

---

## 7. Estat dels mòduls

### `modMotorRotacions.bas`

Implementat:

- validació del context;
- preparació de generació;
- cicle base;
- detecció de dates candidates;
- selecció del següent intensiu;
- creació d'assignacions d'intensiu;
- actualització de l'historial.

Pendent:

- aplicar completament el descans previ de dilluns a dimecres en l'intensiu planificat;
- aplicar la seqüència posterior;
- validar la distribució 60/40;
- validació final de la planificació.

### `modSubstitucions.bas`

Implementat:

- baixa de l'operari original de dilluns a diumenge;
- descans del substitut de dilluns a dimecres;
- intensiu del substitut de dijous a diumenge;
- manteniment intacte del calendari posterior;
- registre de canvi manual;
- actualització de l'historial realitzat.

Pendent:

- test específic;
- anul·lació d'una substitució;
- substitució d'una única assignació;
- altres motius: vacances, permisos i incidències.

### `modGuardies.bas`

Pendent:

- operador principal;
- reforços;
- validacions;
- substitucions de guàrdia.

### `modDataAccess.bas`

Pendent:

- estructura física dels fulls;
- càrrega real de configuració;
- càrrega d'operaris;
- desament d'assignacions;
- historial de canvis.

### Formularis

Pendent:

- activar `frmInici`;
- configuració;
- planificació;
- substitucions;
- guàrdies;
- historial.

---

## 8. Proper ordre de treball

1. Afegir tests de `modSubstitucions`.
2. Implementar el registre d'auditoria.
3. Crear o completar `frmInici`.
4. Crear el formulari de substitucions.
5. Adaptar el motor a la setmana real:
   - descans dilluns-dimecres;
   - intensiu dijous-diumenge;
   - seqüència posterior.
6. Implementar guàrdies i reforços.
7. Implementar persistència real.
8. Generar la visualització del calendari.

---

## 9. Flux de desenvolupament

Per a cada canvi:

1. modificar un únic objectiu funcional;
2. compilar `VBAProject`;
3. executar `RunAllTests`;
4. confirmar que no hi ha regressions;
5. fer commit;
6. fer push a `origin/main`.

Exemple:

```powershell
git add .
git commit -m "Descripció concreta del canvi"
git push origin main
```

---

## 10. Regla de seguretat operativa

El motor genera una proposta.

Una incidència puntual:

- no ha de regenerar el calendari anual;
- no ha de moure automàticament setmanes posteriors;
- només ha de modificar el bloc que l'usuari confirma;
- ha de mantenir la traçabilitat del canvi.
