# agent-ia-local-ssg

Commandes OpenCode installables localement dans un projet.

## Commandes disponibles

- [`/ppt-v1`](./ppt-v1/README.md) : prepare puis genere une presentation PowerPoint a partir de texte, Markdown, image ou PPTX source.

## Installation d'une commande depuis OpenCode Desktop

Prerequis : OpenCode Desktop est deja installe et le consultant est deja dans OpenCode Desktop.

Ne pas installer OpenCode Desktop, opencode, npm, choco, scoop ou git. Cette procedure installe uniquement une commande projet.

Git n'est pas requis. Ne pas utiliser `git clone` et ne pas verifier le depot avec git. L'installation se fait uniquement par telechargement HTTP des fichiers necessaires.

Depuis OpenCode Desktop, dans le projet cible, demander :

```text
Installer uniquement la commande projet /ppt-v1 depuis https://github.com/helie-co/agent-ia-local-ssg/tree/main/ppt-v1. OpenCode Desktop est deja installe, ne pas l'installer. Ne pas utiliser git.
```

Si l'utilisateur donne l'ancienne URL courte `https://github.com/helie-co/agent-ia-local-ssg/ppt-v1`, l'interpreter comme le dossier `ppt-v1` de la branche `main`. Cette URL courte peut renvoyer `404` dans GitHub et ne doit pas declencher de `git clone`.

L'installation est locale au projet courant : elle cree ou met a jour `./.opencode/commands/` et `./.opencode/scripts/`, sans modifier la configuration globale OpenCode.

Fallback sans git : si le telechargement brut fichier par fichier echoue, telecharger l'archive HTTP de la branche `main` et extraire uniquement `ppt-v1/.opencode` vers le projet cible. Ne pas copier le reste du depot.

Apres installation, redemarrer OpenCode Desktop depuis ce projet pour charger la commande.

## Desinstallation d'une commande projet par l'agent

Si l'utilisateur demande de desinstaller `/ppt-v1` :

Ne pas explorer la structure du projet.
Ne pas chercher de fichiers.
Ne pas utiliser git.

Executer directement ces 3 commandes PowerShell depuis la racine du projet cible :

```powershell
Remove-Item -LiteralPath ".\.opencode\commands\ppt-v1.md" -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath ".\.opencode\commands\ppt-v1.README.md" -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath ".\.opencode\scripts\ppt-v1" -Recurse -Force -ErrorAction SilentlyContinue
```

Apres execution, repondre :

```text
La commande /ppt-v1 a ete desinstallee du projet. Redemarrez OpenCode Desktop depuis ce projet pour terminer la desinstallation.
```

Cette desinstallation ne supprime pas OpenCode Desktop, ne modifie pas `~/.config/opencode/`, ne supprime pas les fichiers de travail `ppt-v1/`, ne supprime pas les presentations `.pptx` generees et ne desinstalle pas les dependances OCR optionnelles.
