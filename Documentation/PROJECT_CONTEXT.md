# TORNS CONSORCI

## Estat del projecte

Versió: 0.1
Data inici: Juliol 2026

Projecte desenvolupat amb:

- Excel VBA
- GitHub
- Visual Studio Code
- GitHub Copilot
- ChatGPT (Arquitectura)

---

# Objectiu

Desenvolupar una aplicació professional de planificació de torns per al Consorci del Bages.

No és una plantilla Excel.

És una aplicació completa.

---

# Regles de negoci

## Cicle normal

2 setmanes Matí

↓

2 setmanes Tarda

↓

espera fins a Intensiu

## Intensiu

Només un operari per setmana.

L'intensiu comença DIJOUS.

Finalitza DIMECRES.

Cada operari fa un intensiu cada X setmanes (segons nombre d'operaris).

Ha de ser el més equitatiu possible.

## Després de l'intensiu

1 setmana Tarda

↓

2 setmanes Matí

↓

2 setmanes Tarda

↓

espera fins al següent intensiu.

## Guardies

L'operari d'Intensiu és també l'operari de Guàrdia.

Existeix un segon reforç.

Preparat per un tercer reforç.

---

# Principis del projecte

- No hi ha valors codificats.
- Tot configurable.
- SOLID.
- Clean Code.
- Un fitxer per tasca.
- Un commit per funcionalitat.
- Arquitectura modular.

---

# Arquitectura

## Mòduls

modInici

modConfiguracio

modDataAccess

modMotorRotacions

modGuardies

modUtilitats

## Classes

clsOperari

clsConfiguracio

## Formularis

frmInici

frmOperaris

frmConfiguracio

frmPlanificacio

frmGuardies

---

# Fulls Excel

CONFIGURACIO

OPERARIS

CALENDARI

GUARDIES

FESTIUS

INCIDENCIES

DASHBOARD

---

# Flux de desenvolupament

Sprint 1

Configuració

Sprint 2

Operaris

Sprint 3

Motor de rotacions

Sprint 4

Guardies

Sprint 5

Planificador

Sprint 6

Dashboard

Sprint 7

Informes

Sprint 8

Optimització

---

# Regles de desenvolupament

Mai generar diversos fitxers en un sol prompt.

Sempre un únic fitxer.

Sempre revisar abans d'acceptar.

Mai generar codi provisional.

Només codi de producció.

---

# Objectiu final

Disposar d'una aplicació estable, mantenible i escalable que permeti gestionar la planificació anual dels torns dels operaris del Consorci del Bages.