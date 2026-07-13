# TORNS CONSORCI — Motor de rotacions

## 1. Objectiu

Generar una proposta de planificació de torns mantenint l’equilibri entre els operaris.

El motor no modifica ni desa automàticament el calendari. Només genera assignacions en memòria.

## 2. Estats d’un operari

Durant la planificació, un operari pot trobar-se en un dels estats següents:

- `MATI`
- `TARDA`
- `INTENSIU`
- `POST_INTENSIU_TARDA`
- `POST_INTENSIU_MATI`
- `POST_INTENSIU_TARDA_2`
- `ESPERA_INTENSIU`

## 3. Cicle normal

Quan l’operari no està afectat per un període intensiu, segueix aquest cicle:

```text
2 setmanes MATI
2 setmanes TARDA