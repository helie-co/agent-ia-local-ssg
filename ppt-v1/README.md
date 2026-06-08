# /ppt-v1

Commande OpenCode projet pour preparer puis generer une presentation PowerPoint a partir de texte, Markdown, image ou PPTX source, en respectant un template `.pptx`.

## Installation automatique depuis OpenCode

Cette page sert uniquement a installer la commande projet `/ppt-v1`.

Prerequis obligatoire : OpenCode Desktop est deja installe et le consultant est deja dans OpenCode Desktop.

Ne pas installer OpenCode Desktop.

Ne pas installer opencode, npm, choco, scoop ou git.

Git n'est pas requis.

Ne pas utiliser git.

Ne pas faire `git clone`.

Ne pas verifier le depot avec git.

L'installation se fait uniquement par telechargement HTTP des fichiers necessaires depuis `raw.githubusercontent.com`.

Ne pas modifier la configuration globale `~/.config/opencode/`.

Depuis OpenCode Desktop, dans le projet ou la commande doit etre disponible, demander :

```text
Installer uniquement la commande projet /ppt-v1 depuis https://github.com/helie-co/agent-ia-local-ssg/ppt-v1. OpenCode Desktop est deja installe, ne pas l'installer. Ne pas utiliser git.
```

Si l'utilisateur demande `Installer la commande open code https://github.com/helie-co/agent-ia-local-ssg/ppt-v1`, l'agent doit comprendre qu'il s'agit d'installer uniquement la commande `/ppt-v1`, pas OpenCode Desktop.

L'agent doit lire ce README puis executer directement et uniquement l'installation projet-only suivante, sans utiliser git :

```powershell
& ([scriptblock]::Create((Invoke-RestMethod "https://raw.githubusercontent.com/helie-co/agent-ia-local-ssg/main/ppt-v1/install.ps1"))) -ProjectPath (Get-Location).Path
```

L'installation copie uniquement des fichiers dans le projet courant :

- `./.opencode/commands/ppt-v1.md`
- `./.opencode/commands/ppt-v1.README.md`
- `./.opencode/scripts/ppt-v1/`

Elle ne modifie pas la configuration globale OpenCode dans `~/.config/opencode/`.

Apres installation, repondre :

```text
La commande /ppt-v1 a ete installee dans ce projet. Redemarrez OpenCode Desktop depuis ce projet pour charger la commande.
```

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
