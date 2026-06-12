# agent-ia-local-ssg

Commandes OpenCode installables localement dans un projet.

## Commandes disponibles

- [`/ppt-v1`](./ppt-v1/README.md) : prepare puis genere une presentation PowerPoint a partir de texte, Markdown, image ou PPTX source.
- [`/rec`](./rec/README.md) : enregistrement ecran/audio avec transcription live (whisper.cpp), fusion des chunks, et gestion des sessions.
- [`/ppt-v2`](./ppt-v2/README.md) : prepare puis genere une presentation PowerPoint avec des cartes visuelles (blocs autonomes en grille).

## Installation d'une commande depuis OpenCode Desktop

Prerequis : OpenCode Desktop est deja installe et le consultant est deja dans OpenCode Desktop.

Ne pas installer OpenCode Desktop, opencode, npm, choco, scoop ou git. Git n'est pas requis. L'installation se fait uniquement par telechargement HTTP des fichiers necessaires. L'installation est locale au projet courant : elle cree ou met a jour `./.opencode/commands/` et `./.opencode/scripts/`, sans modifier la configuration globale OpenCode.

Apres installation, redemarrer OpenCode Desktop depuis ce projet.

### Installation de /ppt-v1

Depuis OpenCode Desktop, dans le projet cible, demander :

```text
Installer uniquement la commande projet /ppt-v1 depuis https://github.com/helie-co/agent-ia-local-ssg/tree/main/ppt-v1. OpenCode Desktop est deja installe, ne pas l installer. Ne pas utiliser git.
```

Si l'utilisateur donne l'ancienne URL courte `https://github.com/helie-co/agent-ia-local-ssg/ppt-v1`, l'interpreter comme le dossier `ppt-v1` de la branche `main`. Cette URL courte peut renvoyer `404` dans GitHub et ne doit pas declencher de `git clone`.

Fallback sans git : si le telechargement brut fichier par fichier echoue, telecharger l'archive HTTP de la branche `main` et extraire uniquement `ppt-v1/.opencode` vers le projet cible.

### Installation de /rec

Depuis OpenCode Desktop, dans le projet cible, demander :

```text
Installer uniquement la commande projet /rec depuis https://github.com/helie-co/agent-ia-local-ssg/tree/main/rec. OpenCode Desktop est deja installe, ne pas l installer. Ne pas utiliser git.
```

Les dependances (ffmpeg, whisper, modele) sont installees automatiquement. Apres installation, redemarrer OpenCode Desktop.

Si l'utilisateur donne l'ancienne URL courte `https://github.com/helie-co/agent-ia-local-ssg/rec`, l'interpreter comme le dossier `rec` de la branche `main`. Cette URL courte peut renvoyer `404` dans GitHub et ne doit pas declencher de `git clone`.

Fallback sans git : si le telechargement brut fichier par fichier echoue, telecharger l'archive HTTP de la branche `main` et extraire uniquement `rec/.opencode` vers le projet cible.

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

### Installation de /ppt-v2

Depuis OpenCode Desktop, dans le projet cible, demander :

```text
Installer uniquement la commande projet /ppt-v2 depuis https://github.com/helie-co/agent-ia-local-ssg/tree/main/ppt-v2. OpenCode Desktop est deja installe, ne pas l installer. Ne pas utiliser git.
```

Apres installation, lancer `/ppt-v2 --create-template` pour generer un template cards, puis redemarrer OpenCode Desktop.

Si l'utilisateur donne l'ancienne URL courte `https://github.com/helie-co/agent-ia-local-ssg/ppt-v2`, l'interpreter comme le dossier `ppt-v2` de la branche `main`.

Fallback sans git : telecharger l'archive HTTP de la branche `main` et extraire uniquement `ppt-v2/.opencode` vers le projet cible.

### Desinstallation de /rec

Si l'utilisateur demande de desinstaller `/rec` :

Ne pas explorer la structure du projet.
Ne pas chercher de fichiers.
Ne pas utiliser git.

Executer directement ces commandes PowerShell depuis la racine du projet cible :

```powershell
Remove-Item -LiteralPath ".\.opencode\commands\rec.md" -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath ".\.opencode\scripts\rec" -Recurse -Force -ErrorAction SilentlyContinue
```

Apres execution, repondre :

```text
La commande /rec a ete desinstallee du projet. Redemarrez OpenCode Desktop depuis ce projet pour terminer la desinstallation.
```

Cette desinstallation ne supprime pas OpenCode Desktop, ne modifie pas `~/.config/opencode/`, ne supprime pas les fichiers de travail `recordings/`, ne supprime pas les fichiers `.mp3`/`.mp4`/`.txt` generes et ne desinstalle pas les dependances (ffmpeg, whisper, modele).

### Desinstallation de /ppt-v2

Si l'utilisateur demande de desinstaller `/ppt-v2` :

Ne pas explorer la structure du projet.
Ne pas chercher de fichiers.
Ne pas utiliser git.

Executer directement ces commandes PowerShell depuis la racine du projet cible :

```powershell
Remove-Item -LiteralPath ".\.opencode\commands\ppt-v2.md" -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath ".\.opencode\scripts\ppt-v2" -Recurse -Force -ErrorAction SilentlyContinue
```

Apres execution, repondre :

```text
La commande /ppt-v2 a ete desinstallee du projet. Redemarrez OpenCode Desktop depuis ce projet pour terminer la desinstallation.
```

Cette desinstallation ne supprime pas OpenCode Desktop, ne modifie pas `~/.config/opencode/`, ne supprime pas les fichiers de travail `ppt-v2/`, ne supprime pas les presentations `.pptx` generees.
