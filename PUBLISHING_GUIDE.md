# 📦 Guide de Publication sur GitHub

Ce guide vous explique comment publier **MCP-OUTLOOK** sur GitHub pour le partager avec la communauté.

---

## Préparation Complétée

Le code a été nettoyé de toutes les références spécifiques :
- Références "Disney" supprimées
- Références "Vincent PAPUCHON" supprimées  
- Exemples génériques (company.com, Acme Corp)
- DISNEY_COMPLIANCE.md supprimé
- pyproject.toml avec auteur générique
- README.md nettoyé

---

## 📋 Étapes de Publication

### 1. Créer un Compte GitHub (si nécessaire)

Si vous n'avez pas encore de compte GitHub :
1. Allez sur https://github.com
2. Cliquez sur "Sign up"
3. Suivez les instructions

### 2. Créer un Nouveau Repository

1. Connectez-vous à GitHub
2. Cliquez sur le bouton "+" en haut à droite
3. Sélectionnez "New repository"
4. Configurez le repository :
   - **Repository name** : `mcp-outlook`
   - **Description** : `Model Context Protocol server for Microsoft Outlook - Email, Calendar & Contacts integration`
   - **Visibility** : Public - **Initialize** : Ne cochez RIEN (pas de README, pas de .gitignore, pas de licence)
5. Cliquez sur "Create repository"

### 3. Initialiser Git Localement

Ouvrez PowerShell dans le dossier du projet et exécutez :

```powershell
cd "C:\Users\vpapuchon\source\repos\MCP-OUTLOOK"

# Initialiser le repo git (si pas déjà fait)
git init

# Ajouter tous les fichiers
git add .

# Créer le premier commit
git commit -m "Initial commit: MCP-OUTLOOK v1.0.0"
```

### 4. Lier au Repository GitHub

Remplacez `YOUR_USERNAME` par votre nom d'utilisateur GitHub :

```powershell
# Ajouter le remote
git remote add origin https://github.com/YOUR_USERNAME/mcp-outlook.git

# Vérifier
git remote -v
```

### 5. Pousser le Code sur GitHub

```powershell
# Renommer la branche en main (si nécessaire)
git branch -M main

# Pousser le code
git push -u origin main
```

### 6. Ajouter une Licence

1. Sur GitHub, allez dans votre repository
2. Cliquez sur "Add file" > "Create new file"
3. Nommez le fichier `LICENSE`
4. Cliquez sur "Choose a license template"
5. Sélectionnez **MIT License** (recommandé pour l'open source)
6. Remplissez votre nom
7. Cliquez sur "Review and submit"
8. Commitez le fichier

### 7. Créer une Release

1. Sur GitHub, allez dans l'onglet "Releases"
2. Cliquez sur "Create a new release"
3. Configurez la release :
   - **Tag version** : `v1.0.0`
   - **Release title** : `MCP-OUTLOOK v1.0.0 - Initial Release`
   - **Description** :
     ```markdown
     # First Public Release
     
     ## Features
     - Email management (read, send, search, draft)
     - 📅 Calendar management (events, meetings)
     - 👥 Contact management
     - 📁 Custom folder support
     - Performance optimizations for large mailboxes
     
     ## Requirements
     - Windows OS
     - Microsoft Outlook installed
     - Python 3.10+
     
     See README.md for installation and usage instructions.
     ```
4. Cliquez sur "Publish release"

---

## Après la Publication

### Ajouter des Topics

Sur la page principale de votre repo GitHub :
1. Cliquez sur l'icône ⚙️ à côté de "About"
2. Ajoutez ces topics :
   - `mcp`
   - `model-context-protocol`
   - `outlook`
   - `microsoft-outlook`
   - `email`
   - `calendar`
   - `windows`
   - `python`
   - `fastmcp`
   - `ai-assistant`

### Créer un README Badge

Ajoutez ces badges en haut de votre README.md :

```markdown
# MCP Outlook

[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io)
```

### Partager le Projet

Vous pouvez maintenant partager votre projet :
- Sur les réseaux sociaux (Twitter/X, LinkedIn)
- Dans les communautés MCP
- Sur les forums Python
- Dans les discussions Cursor/Claude

---

## 🔄 Mises à Jour Futures

Pour publier une mise à jour :

```powershell
# Faire vos modifications
git add .
git commit -m "Description des changements"
git push

# Créer une nouvelle release sur GitHub
# Incrémenter la version (v1.0.1, v1.1.0, v2.0.0)
```

---

## Configuration Utilisateur Personnelle

Pour vos besoins personnels (dossier "Vincent PAPUCHON (PERSO)"), ajoutez dans vos **User Rules** de Cursor :

```
Pour la gestion des emails Outlook :
Mes nouveaux emails arrivent dans le dossier "Vincent PAPUCHON (PERSO)/My Mails" et ses sous-dossiers via une règle automatique. L'Inbox est toujours vide. Quand je demande "mes emails", "emails reçus", "nouveaux emails" ou "emails non lus", utilise search_emails_in_custom_folder() avec folder_path="Vincent PAPUCHON (PERSO)/My Mails" au lieu de get_inbox_emails().
```

---

## Félicitations !

Votre projet MCP-OUTLOOK est maintenant public et disponible pour la communauté ! **URL du projet** : `https://github.com/YOUR_USERNAME/mcp-outlook`

N'oubliez pas de :
- ⭐ Mettre une étoile sur votre propre projet
- 📢 Partager le lien
- Répondre aux issues
- Accepter les pull requests

---

**Besoin d'aide ?**
- Documentation Git : https://git-scm.com/doc
- Documentation GitHub : https://docs.github.com
- MCP Documentation : https://modelcontextprotocol.io

