# Guide de Contribution - MCP Outlook

Merci de votre intérêt pour contribuer à MCP Outlook ! Ce guide vous aidera à participer au projet.

## Table des Matières

- [Code de Conduite](#code-de-conduite)
- [Comment Contribuer](#comment-contribuer)
- [Setup Développement](#setup-développement)
- [Standards de Code](#standards-de-code)
- [Process de Pull Request](#process-de-pull-request)
- [Roadmap](#roadmap)

---

## Code de Conduite

### Notre Engagement

Nous nous engageons à faire de la participation à ce projet une expérience sans harcèlement pour tous, indépendamment de :
- L'âge
- La taille corporelle
- Le handicap
- L'ethnicité
- L'identité et l'expression de genre
- Le niveau d'expérience
- La nationalité
- L'apparence personnelle
- La race
- La religion
- L'identité et l'orientation sexuelles

### Comportements Attendus

- Utiliser un langage accueillant et inclusif
- Respecter les points de vue et expériences différents
- Accepter gracieusement les critiques constructives
- Se concentrer sur ce qui est le mieux pour la communauté
- Faire preuve d'empathie envers les autres membres

### Comportements Inacceptables

- L'utilisation de langage ou d'images sexualisés
- Les commentaires trolls, insultants ou dérogatoires
- Le harcèlement public ou privé
- La publication d'informations privées sans permission explicite
- Toute autre conduite inappropriée dans un cadre professionnel

---

## Comment Contribuer

### Signaler des Bugs

Avant de créer une issue :

1. **Vérifiez les issues existantes** pour éviter les doublons
2. **Testez avec la dernière version**
3. **Reproduisez le bug** de manière fiable

#### Template d'Issue pour Bug

```markdown
**Description du Bug**
Description claire et concise du bug.

**Steps to Reproduce**
1. Aller à '...'
2. Cliquer sur '...'
3. Voir l'erreur

**Comportement Attendu**
Ce qui devrait se passer.

**Comportement Actuel**
Ce qui se passe réellement.

**Environment**
- OS: Windows 10/11
- Outlook Version: 2016/2019/365
- Python Version: 3.10/3.11/3.12
- MCP Outlook Version: 1.2.0

**Logs/Screenshots**
Output de `python tests/test_connection.py` et messages d'erreur.

**Contexte Additionnel**
Toute autre information pertinente.
```

### Proposer des Nouvelles Fonctionnalités

#### Template d'Issue pour Feature

```markdown
**Problème/Besoin**
Décrivez le problème ou le besoin que cette feature résoudrait.

**Solution Proposée**
Décrivez la solution que vous envisagez.

**Alternatives Considérées**
Autres approches que vous avez envisagées.

**Impact**
- Sur les utilisateurs existants
- Sur les performances
- Sur la compatibilité

**Implémentation**
Sketch de l'implémentation si vous avez des idées.
```

### Améliorer la Documentation

La documentation est aussi importante que le code !

Contributions bienvenues :
- Corriger les fautes de frappe/grammaire
- Ajouter des exemples
- Clarifier les explications
- Traduire (si multilingue à l'avenir)

---

## Setup Développement

### Prérequis

- **Windows 10/11**
- **Microsoft Outlook** installé et configuré
- **Python 3.10+**
- **Git** pour le contrôle de version

### Installation

```bash
# 1. Fork le projet sur GitHub

# 2. Cloner votre fork
git clone https://github.com/YOUR_USERNAME/mcp-outlook.git
cd mcp-outlook

# 3. Ajouter le remote upstream
git remote add upstream https://github.com/ORIGINAL_OWNER/mcp-outlook.git

# 4. Créer un environnement virtuel
python -m venv venv
venv\Scripts\activate  # Sur Windows

# 5. Installer les dépendances
pip install -r requirements.txt

# 6. Installer les dépendances de développement
pip install pytest black ruff

# 7. Vérifier l'installation
python tests/test_connection.py
```

### Structure du Projet

```
mcp-outlook/
├── src/
│   ├── __init__.py
│   └── outlook_mcp.py       # Serveur MCP principal
├── tests/
│   ├── __init__.py
│   ├── test_connection.py   # Test de connexion Outlook
│   ├── test_outlook_mcp.py  # Tests unitaires
│   ├── test_advanced.py     # Tests avancés
│   └── test_tools.py        # Tests des outils
├── docs/                    # Documentation (si nécessaire)
├── .gitignore
├── pyproject.toml
├── requirements.txt
├── README.md
├── DOCUMENTATION.md
├── CONTRIBUTING.md          # Ce fichier
├── CHANGELOG.md
└── LICENSE
```

---

## Standards de Code

### Style Python

Ce projet suit **PEP 8** avec quelques ajustements :

```python
# Longueur de ligne : 100 caractères (pas 79)
# Utiliser Black pour le formatage automatique

# Bon
def my_function(param1: str, param2: int = 10) -> str:
    """
    Description de la fonction.
    
    Args:
        param1: Description du paramètre 1
        param2: Description du paramètre 2 (default: 10)
    
    Returns:
        Description du retour
    """
    return f"{param1}: {param2}"

# Mauvais
def myFunction(p1,p2=10):
    return f"{p1}: {p2}"
```

### Outils de Qualité

#### Black (Formatage Automatique)

```bash
# Formater tout le code
black src/ tests/

# Vérifier sans modifier
black --check src/ tests/
```

#### Ruff (Linter)

```bash
# Linter
ruff check src/ tests/

# Auto-fix
ruff check --fix src/ tests/
```

### Conventions de Nommage

```python
# Variables et fonctions : snake_case
user_name = "John"
def get_inbox_emails(): ...

# Classes : PascalCase
class EmailManager: ...

# Constantes : UPPER_SNAKE_CASE
MAX_EMAIL_LIMIT = 50
OUTLOOK_FOLDER_INBOX = 6

# Privé : préfixe _
_FOLDER_CACHE = {}
def _get_folder_by_path(): ...
```

### Docstrings

Utilisez le style **Google** :

```python
def send_email(
    to: str,
    subject: str,
    body: str,
    cc: Optional[str] = None
) -> str:
    """
    Send an email via Outlook.
    
    Creates and sends a new email through the user's Outlook account.
    The email is sent immediately and a copy is saved in the Sent Items folder.
    
    Args:
        to: Recipient email address(es), semicolon-separated for multiple.
            Example: "user1@example.com" or "user1@example.com; user2@example.com"
        subject: Email subject line
        body: Email body content (plain text format)
        cc: CC recipients (optional), semicolon-separated
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "message": str
        }
    
    Examples:
        >>> send_email("colleague@company.com", "Meeting", "See you at 2pm")
        {"success": true, "message": "Email sent to colleague@company.com"}
    
    Raises:
        ValueError: If Outlook is not accessible
    
    Notes:
        - Recipient addresses are logged but email content is not
        - BCC recipients are never logged for privacy
    """
    # Implementation
```

### Type Hints

Utilisez les type hints partout :

```python
from typing import Optional, Dict, Any, List

def format_email(mail_item) -> Dict[str, Any]:
    """Format an email item."""
    ...

def get_inbox_emails(
    limit: int = 10,
    unread_only: bool = False
) -> str:
    """Get inbox emails."""
    ...
```

### Gestion des Erreurs

```python
# Retour JSON cohérent
try:
    # Code
    return json.dumps({
        "success": True,
        "data": "..."
    }, indent=2)
except Exception as e:
    logger.error("Failed to ...", exc_info=True, extra={
        "param1": value1,
        "param2": value2
    })
    return json.dumps({
        "success": False,
        "error": str(e)
    })
```

### Tests

Chaque nouvelle feature doit avoir des tests :

```python
# tests/test_new_feature.py
import pytest
from src.outlook_mcp import new_function

def test_new_function_success():
    """Test new_function with valid input."""
    result = new_function("valid_input")
    assert result["success"] is True

def test_new_function_error():
    """Test new_function with invalid input."""
    result = new_function("invalid_input")
    assert result["success"] is False
    assert "error" in result
```

---

## Process de Pull Request

### Workflow Git

```bash
# 1. Synchroniser avec upstream
git fetch upstream
git checkout main
git merge upstream/main

# 2. Créer une branche pour votre feature
git checkout -b feature/my-awesome-feature
# OU
git checkout -b fix/bug-description

# 3. Faire vos modifications
# Éditez les fichiers...

# 4. Tester
python tests/test_connection.py
pytest tests/

# 5. Formater et linter
black src/ tests/
ruff check --fix src/ tests/

# 6. Commit
git add .
git commit -m "feat: add awesome feature"
# OU
git commit -m "fix: resolve bug with email attachments"

# 7. Push
git push origin feature/my-awesome-feature

# 8. Créer une Pull Request sur GitHub
```

### Convention de Commit

Utilisez **Conventional Commits** :

```
<type>(<scope>): <description>

[optional body]

[optional footer]
```

#### Types

- `feat`: Nouvelle fonctionnalité
- `fix`: Correction de bug
- `docs`: Documentation seulement
- `style`: Formatage, indentation (pas de changement de code)
- `refactor`: Refactoring (pas de nouvelle feature ni fix)
- `perf`: Amélioration de performance
- `test`: Ajout ou correction de tests
- `chore`: Maintenance (dépendances, config, etc.)

#### Exemples

```bash
feat(email): add support for HTML email attachments
fix(calendar): resolve timezone issue in event creation
docs(readme): update installation instructions
refactor(contacts): simplify search logic
perf(folders): optimize folder cache lookup
test(email): add tests for send_email with attachments
chore(deps): update pywin32 to v306
```

### Checklist Pull Request

Avant de soumettre une PR, vérifiez :

- [ ] Le code suit les standards de style (Black + Ruff)
- [ ] Les tests passent (`pytest tests/`)
- [ ] Les nouveaux tests sont ajoutés pour les nouvelles features
- [ ] La documentation est mise à jour (README, DOCUMENTATION, CHANGELOG)
- [ ] Les docstrings sont complètes
- [ ] Pas de code commenté ou de debug prints
- [ ] Les commits suivent Conventional Commits
- [ ] La PR a une description claire

### Template de Pull Request

```markdown
## Description

Brève description des changements.

## Type de Changement

- [ ] Bug fix (changement non-breaking qui corrige une issue)
- [ ] New feature (changement non-breaking qui ajoute une fonctionnalité)
- [ ] Breaking change (fix ou feature qui casserait des fonctionnalités existantes)
- [ ] Documentation update

## Comment Tester

1. Step 1
2. Step 2
3. Résultat attendu

## Checklist

- [ ] Mon code suit les standards du projet
- [ ] J'ai effectué une auto-review de mon code
- [ ] J'ai commenté le code dans les parties difficiles
- [ ] J'ai mis à jour la documentation
- [ ] Mes changements ne génèrent pas de nouveaux warnings
- [ ] J'ai ajouté des tests qui prouvent que mon fix/feature fonctionne
- [ ] Les tests unitaires passent localement
- [ ] J'ai mis à jour le CHANGELOG.md

## Screenshots (si applicable)

![Screenshot](url)

## Issues Liées

Fixes #123
Relates to #456
```

---

## Roadmap

Consultez le [CHANGELOG.md](CHANGELOG.md) pour la roadmap complète.

### Priorités Actuelles

#### High Priority
- [ ] Gestion des tâches (tasks)
- [ ] Gestion des dossiers (create, move, delete)
- [ ] Filtres avancés (flags, categories)

#### Medium Priority
- [ ] Gestion des règles emails (create, modify, delete)
- [ ] Prévisualisation de pièces jointes
- [ ] Opérations en batch

#### Low Priority
- [ ] Support cross-platform (explore alternatives MAPI)
- [ ] Interface web (optionnel)

### Features Completed

- [x] Email management (v1.0.0)
- [x] Calendar management (v1.0.0)
- [x] Contact management (v1.0.0)
- [x] Folder management (v1.0.0)
- [x] HTML email support (v1.1.0)
- [x] Outlook signature integration (v1.1.0)
- [x] Attachment management (v1.2.0)
- [x] Meeting response handling (v1.2.0)
- [x] Out-of-Office settings (v1.2.0)

---

## Questions Fréquentes

### Q: Mon PR a été rejeté, que faire ?

**R**: Ne vous découragez pas ! Lisez les commentaires des reviewers, effectuez les modifications demandées, et re-soumettez. C'est un processus d'apprentissage.

### Q: Je ne sais pas par où commencer ?

**R**: Regardez les issues étiquetées `good first issue` ou `help wanted`. Ce sont de bons points de départ pour les nouveaux contributeurs.

### Q: Je peux contribuer sans savoir coder ?

**R**: Absolument ! Vous pouvez :
- Améliorer la documentation
- Traduire (si multilingue)
- Signaler des bugs
- Suggérer des améliorations
- Aider d'autres utilisateurs dans les issues

### Q: Comment tester mes changements ?

**R**: 
1. Exécutez `python tests/test_connection.py` pour les tests basiques
2. Exécutez `pytest tests/` pour tous les tests
3. Testez manuellement avec un vrai Outlook
4. Vérifiez que les fonctionnalités existantes marchent toujours

### Q: Mes tests échouent, que faire ?

**R**: 
1. Lisez les messages d'erreur attentivement
2. Vérifiez qu'Outlook est en cours d'exécution
3. Vérifiez votre environnement Python
4. Demandez de l'aide dans l'issue ou la PR

---

## Ressources

### Documentation Externe

- **Python**: https://docs.python.org/3/
- **FastMCP**: https://github.com/jlowin/fastmcp
- **pywin32**: https://github.com/mhammond/pywin32
- **Outlook COM API**: https://docs.microsoft.com/en-us/office/vba/api/overview/outlook
- **Model Context Protocol**: https://modelcontextprotocol.io

### Documentation du Projet

- [README.md](README.md) - Vue d'ensemble
- [DOCUMENTATION.md](DOCUMENTATION.md) - Documentation technique complète
- [QUICK_START.md](QUICK_START.md) - Guide de démarrage rapide
- [EXAMPLES.md](EXAMPLES.md) - Exemples d'utilisation
- [CHANGELOG.md](CHANGELOG.md) - Historique des versions

---

## Remerciements

Merci à tous les contributeurs qui ont fait de MCP Outlook ce qu'il est aujourd'hui !

### Comment Être Listé

Si vous contribuez de manière significative :
- Bug fixes importants
- Nouvelles features
- Améliorations de documentation
- Aide à la communauté

Votre nom sera ajouté à la section remerciements dans le README !

---

## Contact

- **Issues GitHub**: [Créer une issue](https://github.com/YOUR_USERNAME/mcp-outlook/issues)
- **Discussions**: [GitHub Discussions](https://github.com/YOUR_USERNAME/mcp-outlook/discussions)
- **Email**: Pour les questions sensibles uniquement

---

**Merci de contribuer à MCP Outlook !** 

Chaque contribution, petite ou grande, fait une différence.

**Version**: 1.2.0  
**Dernière mise à jour**: 17 décembre 2025

