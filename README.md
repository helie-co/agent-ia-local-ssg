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
Installer uniquement la commande projet /ppt-v1 depuis https://github.com/helie-co/agent-ia-local-ssg/ppt-v1. OpenCode Desktop est deja installe, ne pas l'installer. Ne pas utiliser git.
```

L'installation est locale au projet courant : elle cree ou met a jour `./.opencode/commands/` et `./.opencode/scripts/`, sans modifier la configuration globale OpenCode.

Apres installation, redemarrer OpenCode Desktop depuis ce projet pour charger la commande.
