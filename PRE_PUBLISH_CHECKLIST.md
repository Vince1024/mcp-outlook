# ✅ Checklist Pré-Publication

Utilisez cette checklist pour vérifier que tout est prêt avant de publier sur GitHub.

---

## 📋 Vérifications Automatiques

### 1. Vérifier l'absence de références spécifiques

```powershell
# Rechercher "Disney" dans le code
git grep -i "disney" -- "*.py" "*.toml" "*.md" | grep -v "PREPARATION_SUMMARY\|PUBLISHING_GUIDE\|PRE_PUBLISH_CHECKLIST"

# Rechercher "Vincent PAPUCHON" dans le code
git grep -i "vincent papuchon" -- "*.py" "*.toml" "*.md" | grep -v "PREPARATION_SUMMARY\|PUBLISHING_GUIDE\|PRE_PUBLISH_CHECKLIST"
```

**Résultat attendu** : Aucune correspondance (sauf dans les guides de publication)

### 1.1 Vérifier que .vscode/ sera ignoré

```powershell
# Vérifier que .gitignore contient .vscode/
findstr /C:".vscode" .gitignore
```

**Résultat attendu** : `.vscode/` doit être présent dans .gitignore

**Note** : Le dossier `.vscode/` contient des configurations personnelles (chemins absolus, préférences d'éditeur) qui ne doivent PAS être publiées. Le `.gitignore` est déjà configuré pour l'ignorer automatiquement.

### 2. Vérifier la structure du projet

```powershell
tree /F /A
```

**Fichiers attendus** :
- ✅ README.md (avec badges)
- ✅ LICENSE (MIT)
- ✅ CHANGELOG.md
- ✅ PUBLISHING_GUIDE.md
- ✅ PREPARATION_SUMMARY.md
- ✅ PRE_PUBLISH_CHECKLIST.md (ce fichier)
- ✅ requirements.txt
- ✅ pyproject.toml
- ✅ .gitignore
- ✅ src/outlook_mcp.py
- ✅ tests/

**Fichiers à NE PAS avoir** :
- ❌ DISNEY_COMPLIANCE.md (supprimé)

### 3. Tester le serveur localement

```powershell
# Installer les dépendances
pip install -r requirements.txt

# Tester la connexion Outlook
python -c "from src.outlook_mcp import get_outlook_application; print('OK' if get_outlook_application() else 'FAIL')"

# Lancer le serveur (Ctrl+C pour arrêter)
python src/outlook_mcp.py
```

---

## 📝 Checklist Manuelle

### Code et Configuration

- [ ] **Code nettoyé** : Aucune référence à "Disney" ou "Vincent PAPUCHON" dans le code source
- [ ] **pyproject.toml** : Auteur générique "MCP-Outlook Contributors"
- [ ] **Version** : 1.0.0 dans pyproject.toml
- [ ] **EXCLUDED_STORES** : Liste vide ou avec commentaire exemple uniquement

### Documentation

- [ ] **README.md** : 
  - [ ] Badges ajoutés en haut
  - [ ] Pas de références internes
  - [ ] Licence MIT mentionnée
  - [ ] Exemples génériques (company.com, Acme Corp)
  
- [ ] **LICENSE** : Fichier MIT License présent

- [ ] **CHANGELOG.md** : Version 1.0.0 documentée

- [ ] **PUBLISHING_GUIDE.md** : Guide complet de publication

### Git et GitHub

- [ ] **Git installé** : `git --version` fonctionne

- [ ] **Git configuré** :
  ```powershell
  git config --global user.name "Votre Nom"
  git config --global user.email "votre@email.com"
  ```

- [ ] **Repository local initialisé** : `.git` existe

- [ ] **Fichiers ajoutés** : `git add .` exécuté

- [ ] **Commit initial créé** : 
  ```powershell
  git commit -m "Initial commit: MCP-OUTLOOK v1.0.0 - Ready for public release"
  ```

- [ ] **Repository GitHub créé** :
  - Nom : `mcp-outlook`
  - Visibilité : Public
  - Pas de README/LICENSE/.gitignore initialisé

- [ ] **Remote configuré** :
  ```powershell
  git remote add origin https://github.com/YOUR_USERNAME/mcp-outlook.git
  ```

### Publication

- [ ] **Code poussé** :
  ```powershell
  git branch -M main
  git push -u origin main
  ```

- [ ] **Release créée** :
  - Tag : `v1.0.0`
  - Title : `MCP-OUTLOOK v1.0.0 - Initial Release`
  - Description complète

- [ ] **Topics ajoutés** :
  - mcp
  - model-context-protocol
  - outlook
  - microsoft-outlook
  - email
  - calendar
  - windows
  - python
  - fastmcp
  - ai-assistant

### Post-Publication

- [ ] **Repository vérifié** : URL accessible publiquement

- [ ] **README s'affiche correctement** : Badges visibles

- [ ] **Release visible** : v1.0.0 dans l'onglet Releases

- [ ] **Clone test** :
  ```powershell
  cd %TEMP%
  git clone https://github.com/YOUR_USERNAME/mcp-outlook.git
  cd mcp-outlook
  pip install -r requirements.txt
  python src/outlook_mcp.py
  ```

---

## 🚀 Commandes Rapides

### Publication Automatique

```powershell
# Utiliser le script de publication
.\publish.bat
```

### Publication Manuelle

```powershell
# Initialiser et commiter
git init
git add .
git commit -m "Initial commit: MCP-OUTLOOK v1.0.0 - Ready for public release"

# Ajouter le remote (remplacer YOUR_USERNAME)
git remote add origin https://github.com/YOUR_USERNAME/mcp-outlook.git

# Pousser sur GitHub
git branch -M main
git push -u origin main
```

---

## ⚠️ Points d'Attention

### Avant de Publier

1. **Assurez-vous qu'Outlook fonctionne** sur votre machine
2. **Testez le serveur localement** avant de publier
3. **Vérifiez que Git est configuré** avec vos identifiants
4. **Créez le repository sur GitHub** avant de pousser

### Après la Publication

1. **Ne commitez jamais de credentials** ou données sensibles
2. **Répondez aux issues** rapidement
3. **Acceptez les pull requests** de qualité
4. **Maintenez le CHANGELOG** à jour

---

## 📞 Aide

### Problèmes Courants

**"Git n'est pas reconnu"**
- Installez Git : https://git-scm.com/download/win
- Redémarrez votre terminal

**"Permission denied (publickey)"**
- Configurez SSH : https://docs.github.com/en/authentication/connecting-to-github-with-ssh
- Ou utilisez HTTPS avec token

**"Repository not found"**
- Vérifiez que le repository existe sur GitHub
- Vérifiez l'URL du remote : `git remote -v`

**"Push rejected"**
- Le repository a peut-être été initialisé avec des fichiers
- Utilisez `git pull origin main --allow-unrelated-histories` puis `git push`

---

## ✅ Validation Finale

Une fois tous les points cochés :

```powershell
echo "🎉 MCP-OUTLOOK est prêt pour la publication !"
echo "URL: https://github.com/YOUR_USERNAME/mcp-outlook"
```

**Félicitations ! Votre projet est maintenant open source ! 🚀**

---

**Consultez PUBLISHING_GUIDE.md pour des instructions détaillées.**

