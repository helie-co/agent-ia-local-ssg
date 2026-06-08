# /ppt-v1

Commande OpenCode projet pour preparer puis generer une presentation PowerPoint a partir de texte, Markdown, image ou PPTX source, en respectant un template `.pptx`.

## Installation automatique depuis OpenCode

Depuis OpenCode Desktop, dans le projet ou la commande doit etre disponible, demander :

```text
Installer la commande open code https://github.com/helie-co/agent-ia-local-ssg/ppt-v1
```

L'agent doit lire ce README puis executer l'installation projet-only suivante :

```powershell
& ([scriptblock]::Create((Invoke-RestMethod "https://raw.githubusercontent.com/helie-co/agent-ia-local-ssg/main/ppt-v1/install.ps1"))) -ProjectPath (Get-Location).Path
```

L'installation copie uniquement des fichiers dans le projet courant :

- `./.opencode/commands/ppt-v1.md`
- `./.opencode/commands/ppt-v1.README.md`
- `./.opencode/scripts/ppt-v1/`

Elle ne modifie pas la configuration globale OpenCode dans `~/.config/opencode/`.

Apres installation, redemarrer OpenCode Desktop depuis ce projet pour charger la commande.

## Installation manuelle

Depuis PowerShell, a la racine du projet cible :

```powershell
& ([scriptblock]::Create((Invoke-RestMethod "https://raw.githubusercontent.com/helie-co/agent-ia-local-ssg/main/ppt-v1/install.ps1"))) -ProjectPath (Get-Location).Path
```

Avec un chemin projet explicite :

```powershell
& ([scriptblock]::Create((Invoke-RestMethod "https://raw.githubusercontent.com/helie-co/agent-ia-local-ssg/main/ppt-v1/install.ps1"))) -ProjectPath "C:\chemin\vers\mon-projet"
```

## Utilisation

Apres redemarrage d'OpenCode Desktop :

```text
/ppt-v1 --help
/ppt-v1 Genere 5 slides de presentation sur l'IA
/ppt-v1 @source.md
/ppt-v1 Reprends les slides de @source.pptx avec le template @template.pptx
```

Pour installer les dependances OCR optionnelles :

```text
/ppt-v1 --install
```
