---
description: Enregistrement ecran/audio avec transcription live (whisper.cpp), fusion, et gestion des chunks
agent: build
---

Execute la commande PowerShell ci-dessous. Retourne le texte produit par la commande dans un bloc de code markdown (``` ... ```), sans modification du contenu.

Le script `rec.ps1` gere l'enregistrement ecran/audio avec ou sans transcription live automatique.

Commande normale :

```powershell
$script = Join-Path (Get-Location) '.opencode\scripts\rec\rec.ps1'
$rawArgs = '$ARGUMENTS'
& powershell -NoProfile -ExecutionPolicy Bypass -Command "& '$script' -RawArgs '$rawArgs'"
```
