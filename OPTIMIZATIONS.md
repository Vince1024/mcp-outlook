# 🚀 Optimisations Outlook MCP

## Vue d'ensemble

Ce MCP Outlook a été optimisé pour **minimiser le gel d'Outlook** pendant les requêtes COM et **améliorer les performances** sur les boîtes mail volumineuses.

---

## ✅ Optimisations Implémentées

### 1. **Système de Cache pour les Dossiers**
- Cache global `_FOLDER_CACHE` qui mémorise les chemins de dossiers résolus
- Fonction `_get_folder_by_path()` avec support du cache
- **Gain** : Évite la traversée coûteuse de tous les stores Outlook à chaque requête
- **Résultat** : Première recherche ~45s, recherches suivantes ~1s (45x plus rapide)

### 2. **Suppression des Appels `items.Count`**
- Remplacé par indexation directe `items[i+1]` dans toutes les fonctions
- `items.Count` peut prendre **plusieurs minutes** sur de grandes boîtes mail
- **Fonctions optimisées** :
  - `get_inbox_emails()`
  - `get_sent_emails()`
  - `search_emails()`
  - `search_emails_in_custom_folder()`
  - `get_contacts()`

### 3. **Filtre par Date pour Réduire le Scope**
- Paramètre `days_back` dans `search_emails_in_custom_folder()`
- Par défaut : **2 derniers jours** seulement (configurable)
- Utilise `Restrict()` côté serveur **avant** l'itération
- **Gain** : Réduit drastiquement le nombre d'emails à parcourir
- **Résultat** : Moins de gel d'Outlook (quelques secondes au lieu de minutes)

### 4. **Réduction des Limites par Défaut**
```python
DEFAULT_EMAIL_LIMIT = 5        # Réduit de 10 → 5
MAX_EMAIL_LIMIT = 50           # Réduit de 100 → 50
DEFAULT_DAYS_BACK = 2          # Seulement 2 derniers jours par défaut
```
**Raison** : Moins d'emails = moins de gel d'Outlook

### 5. **`list_outlook_folders()` Ultra-Rapide**
- Paramètre `include_counts=False` par défaut
- Ne calcule **pas** les `item_count` et `unread_count` (très coûteux)
- **Gain** : Passe de plusieurs minutes à quelques secondes

### 6. **Indexation Directe au Lieu de GetFirst()/GetNext()**
- `items[i+1]` au lieu de `GetFirst()` / `GetNext()`
- Plus rapide sur les collections filtrées
- Gestion des exceptions pour la fin de collection

### 7. **Exclusion des Boîtes d'Équipe et Partagées**
- Liste `EXCLUDED_STORES` pour exclure les boîtes mail d'équipe/partagées
- Par défaut : `"DLP IS SPID"` (boîte d'équipe Disney)
- **Gain** : Évite de scanner des milliers d'emails d'équipe inutilement
- **Résultat** : Recherches plus rapides et résultats plus pertinents
- **Configuration** : Ajoutez simplement le nom du store dans la liste

```python
EXCLUDED_STORES = [
    "DLP IS SPID",                 # Team mailbox
    "Autre Boite Partagée",        # Autre exemple
]
```

---

## 📊 Performances

### Avant Optimisations
| Opération | Durée |
|-----------|-------|
| Recherche dans "My Mails" (sans cache) | ~45s |
| Recherche dans "My Mails" (répétée) | ~45s |
| `list_outlook_folders()` avec counts | Plusieurs minutes |
| Gel d'Outlook pendant les requêtes | Très long (minutes) |

### Après Optimisations
| Opération | Durée |
|-----------|-------|
| Recherche dans "My Mails" (1ère fois) | ~45s (recherche dossier) |
| Recherche dans "My Mails" (avec cache) | ~1s (lookup) |
| Recherche d'emails (2 derniers jours) | Variable selon volume* |
| `list_outlook_folders()` sans counts | Quelques secondes |
| Gel d'Outlook | Réduit (secondes au lieu de minutes) |

_*Note : Sur des dossiers avec énormément d'emails même récents, le gel peut persister. C'est une limitation structurelle d'Outlook COM._

---

## 🔧 Configuration

### Variables de Configuration (src/outlook_mcp.py)

```python
DEFAULT_EMAIL_LIMIT = 5            # Limite par défaut pour les emails
MAX_EMAIL_LIMIT = 50               # Limite maximum
DEFAULT_DAYS_BACK = 2              # Jours en arrière pour la recherche
```

### Utilisation

**Recherche standard (2 derniers jours) :**
```python
search_emails_in_custom_folder("Vincent PAPUCHON (PERSO)/My Mails")
```

**Recherche étendue (30 derniers jours) :**
```python
search_emails_in_custom_folder("Vincent PAPUCHON (PERSO)/My Mails", days_back=30)
```

**Recherche TOUS les emails (lent, peut geler Outlook) :**
```python
search_emails_in_custom_folder("Vincent PAPUCHON (PERSO)/My Mails", days_back=0)
```

---

## ⚠️ Limitations Connues

### Gel d'Outlook
Malgré toutes les optimisations, **Outlook COM est single-threaded** :
- Pendant une requête MCP, Outlook ne peut pas répondre à vos clics
- C'est une limitation architecturale de l'API COM Outlook
- Le gel est **réduit** mais **pas éliminé complètement**

### Solutions :
1. ✅ Fermer Outlook pendant l'utilisation du MCP
2. ✅ Utiliser des dossiers plus spécifiques (moins d'emails)
3. ✅ Réduire `days_back` au minimum nécessaire
4. ✅ Réduire les `limit` de résultats

---

## 📝 UserRule Recommandée (Cursor)

Pour une utilisation optimale avec Cursor, ajoutez cette UserRule :

```
Pour mes emails Outlook : mes nouveaux emails arrivent dans "Vincent PAPUCHON (PERSO)/My Mails" via une règle automatique. L'Inbox est toujours vide. Quand je demande "mes emails", "emails reçus", "nouveaux emails" ou "emails non lus", utilise TOUJOURS search_emails_in_custom_folder() avec folder_path="Vincent PAPUCHON (PERSO)/My Mails" au lieu de get_inbox_emails(). Par défaut, cherche sur les 2 derniers jours (days_back=2).
```

---

## 🎯 Recommandations

### Pour un Usage Optimal

1. **Spécifiez toujours un sous-dossier spécifique** si possible :
   ```python
   "Vincent PAPUCHON (PERSO)/My Mails/Incidents"
   "Vincent PAPUCHON (PERSO)/My Mails/Jira"
   ```

2. **Utilisez des plages de dates courtes** :
   - `days_back=1` pour aujourd'hui
   - `days_back=2` pour avant-hier et aujourd'hui (défaut)
   - `days_back=7` pour la semaine

3. **Réduisez les limites** si vous n'avez pas besoin de beaucoup d'emails :
   ```python
   search_emails_in_custom_folder(..., limit=10)
   ```

4. **Fermez Outlook** si vous faites beaucoup de requêtes MCP d'affilée

---

## 🔬 Tests

Des tests sont disponibles dans le dossier `tests/` :
- `test_connection.py` : Test de connexion Outlook
- `test_outlook_mcp.py` : Tests unitaires des fonctions MCP
- `test_advanced.py` : Tests avancés
- `test_tools.py` : Tests des outils

---

## 📚 Documentation

- `README.md` : Documentation principale
- `QUICK_START.md` : Guide de démarrage rapide
- `EXAMPLES.md` : Exemples d'utilisation
- `CHANGELOG.md` : Historique des changements

---

**Date** : 16 décembre 2025  
**Version** : 1.1.0 (optimisée)  
**Auteur** : Disney DLP-SPID Team

