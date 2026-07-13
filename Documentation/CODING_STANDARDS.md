# CODING STANDARDS

## Llengua

Tot el projecte es desenvolupa en català.

Variables, comentaris, formularis i procediments.

---

## Option Explicit

Tots els mòduls han d'incloure:

Option Explicit

---

## Noms

Mòduls

modConfiguracio

modMotorRotacions

modGuardies

modUtilitats

Classes

clsOperari

clsConfiguracio

Formularis

frmInici

frmOperaris

frmConfiguracio

frmPlanificacio

frmGuardies

---

## Procediments

Els procediments han de tenir una única responsabilitat.

No més de 60 línies sempre que sigui possible.

---

## Funcions

Una funció només ha de retornar un resultat.

No modificar dades externes.

---

## Errors

No utilitzar:

On Error Resume Next

Excepte quan sigui absolutament necessari.

Sempre registrar els errors.

---

## Configuració

Cap valor fix dins del codi.

Tot ha de provenir del full CONFIGURACIO.

---

## Dates

Treballar sempre amb Date.

Mai amb String.

---

## Colors

Mai utilitzar valors RGB escrits al codi.

Els colors es defineixen a CONFIGURACIO.

---

## Fulls Excel

Mai utilitzar:

Sheets("CALENDARI")

Utilitzar constants o funcions centralitzades.

---

## Comentaris

Cada procediment ha de començar amb:

'=========================================================
' Nom:
' Descripció:
' Autor:
' Data:
'=========================================================

---

## Git

Un commit = una funcionalitat.

Mai diversos canvis grans en un únic commit.

---

## Copilot

Copilot no decideix l'arquitectura.

Només escriu codi.

Les decisions tècniques les pren l'arquitecte del projecte.