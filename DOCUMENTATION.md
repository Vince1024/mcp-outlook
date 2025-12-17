# Documentation Complète - MCP Outlook

Documentation technique détaillée pour MCP Outlook v1.2.0

## Table des Matières

- [Architecture](#architecture)
- [Installation](#installation)
- [Configuration](#configuration)
- [Outils Email](#outils-email)
- [Outils Calendrier](#outils-calendrier)
- [Outils Contacts](#outils-contacts)
- [Outils Dossiers](#outils-dossiers)
- [Outils Out-of-Office](#outils-out-of-office)
- [Gestion des Erreurs](#gestion-des-erreurs)
- [Performances](#performances)
- [Sécurité](#sécurité)
- [Limitations](#limitations)

---

## Architecture

### Vue d'ensemble

MCP Outlook est un serveur MCP (Model Context Protocol) qui permet aux assistants IA d'interagir avec Microsoft Outlook via l'API COM Windows.

```
┌─────────────────┐
│  AI Assistant   │
│ (Cursor/Claude) │
└────────┬────────┘
         │
         │ MCP Protocol
         │
┌────────▼────────┐
│   MCP Outlook   │
│     Server      │
└────────┬────────┘
         │
         │ COM Automation
         │
┌────────▼────────┐
│    Microsoft    │
│     Outlook     │
└─────────────────┘
```

### Technologies

- **Python 3.10+** - Langage de base
- **FastMCP** - Framework MCP
- **pywin32** - COM automation
- **dateutil** - Parsing de dates flexible

### Structure du Projet

```
mcp-outlook/
├── src/
│   ├── __init__.py
│   └── outlook_mcp.py      # Serveur MCP principal (1840 lignes)
├── tests/
│   ├── __init__.py
│   ├── test_connection.py
│   ├── test_outlook_mcp.py
│   ├── test_advanced.py
│   └── test_tools.py
├── pyproject.toml          # Configuration du projet
├── requirements.txt        # Dépendances
├── README.md              # Documentation utilisateur
├── DOCUMENTATION.md       # Ce fichier
├── CONTRIBUTING.md        # Guide de contribution
├── CHANGELOG.md           # Historique des versions
├── LICENSE                # Licence MIT
├── QUICK_START.md         # Guide de démarrage rapide
└── EXAMPLES.md            # Exemples d'utilisation
```

---

## Installation

### Prérequis

#### Système d'Exploitation
- **Windows 10/11** (requis pour COM automation)
- Impossible sur Linux/macOS (API COM Windows uniquement)

#### Logiciels
- **Microsoft Outlook** (version 2010+)
  - Installé et configuré
  - Au moins un compte email configuré
  - Outlook doit être en cours d'exécution

#### Python
- **Python 3.10 ou supérieur**
- Vérifier : `python --version`
- Télécharger : https://www.python.org/downloads/

### Installation des Dépendances

```bash
# Cloner/télécharger le projet
git clone https://github.com/YOUR_USERNAME/mcp-outlook.git
cd mcp-outlook

# Installer les dépendances
pip install -r requirements.txt

# OU installer en mode développement
pip install -e .
```

### Dépendances Python

Le fichier `requirements.txt` contient :

```txt
fastmcp>=0.1.0
pywin32>=306
python-dateutil>=2.8.2
```

#### Détails des dépendances

- **fastmcp** - Framework MCP pour créer des serveurs
- **pywin32** - Accès aux API Windows COM
- **python-dateutil** - Parsing flexible des dates

### Post-Installation pywin32

Si vous rencontrez des erreurs avec `win32com`, exécutez :

```bash
python Scripts/pywin32_postinstall.py -install
```

### Vérification de l'Installation

```bash
python tests/test_connection.py
```

Résultat attendu :
```
✓ PASS: Imports
✓ PASS: Outlook Connection
✓ PASS: Server File
```

---

## Configuration

### Configuration MCP

#### Pour Cursor

Fichier : `~/.cursor/mcp.json` ou workspace settings

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": [
        "C:/Users/YOUR_USERNAME/path/to/mcp-outlook/src/outlook_mcp.py"
      ],
      "env": {}
    }
  }
}
```

#### Pour Claude Desktop

Fichier : `%APPDATA%/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": [
        "C:/Users/YOUR_USERNAME/path/to/mcp-outlook/src/outlook_mcp.py"
      ]
    }
  }
}
```

### Variables de Configuration

Dans `src/outlook_mcp.py` :

```python
# Limites par défaut
DEFAULT_EMAIL_LIMIT = 5        # Limite par défaut d'emails retournés
MAX_EMAIL_LIMIT = 50           # Limite maximale autorisée
DEFAULT_CONTACT_LIMIT = 50     # Limite par défaut de contacts
MAX_CONTACT_LIMIT = 200        # Limite maximale de contacts
EMAIL_BODY_PREVIEW_LENGTH = 500  # Longueur du preview du body
DEFAULT_DAYS_BACK = 2          # Recherche des 2 derniers jours par défaut

# Stores exclus (boîtes mail d'équipe/partagées)
EXCLUDED_STORES = [
    # Exemple: "Team Mailbox Name",
]
```

### Logging

Le logging est configuré en mode silencieux par défaut :

```python
# Niveau CRITICAL uniquement (pas de spam)
logging.basicConfig(
    level=logging.CRITICAL,
    format='%(message)s',
    handlers=[logging.NullHandler()]
)
```

Pour activer le debugging, modifiez dans `outlook_mcp.py` :

```python
logger.setLevel(logging.DEBUG)  # Au lieu de CRITICAL
```

---

## Outils Email

### `get_inbox_emails`

Récupère les emails de la boîte de réception.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `limit` | int | 5 | Nombre max d'emails à retourner |
| `unread_only` | bool | False | Ne retourner que les emails non lus |

#### Retour

```json
{
  "success": true,
  "count": 3,
  "emails": [
    {
      "subject": "Q1 Budget Review",
      "sender": "Jane Smith",
      "sender_email": "jane.smith@company.com",
      "recipients": "team@company.com",
      "cc": "",
      "bcc": "",
      "received_time": "2025-12-17 09:30:00+00:00",
      "sent_on": "2025-12-17 09:29:45+00:00",
      "body": "Hi team, please review...",
      "body_length": 1245,
      "has_attachments": true,
      "attachment_count": 2,
      "attachments": [
        {
          "filename": "report.pdf",
          "size": 245678,
          "type": 1
        }
      ],
      "importance": 1,
      "unread": true,
      "categories": "",
      "entry_id": "00000000..."
    }
  ]
}
```

#### Exemple

```python
# 10 derniers emails non lus
get_inbox_emails(limit=10, unread_only=True)

# 5 derniers emails (lus et non lus)
get_inbox_emails(limit=5)
```

#### Notes

- Limité à `MAX_EMAIL_LIMIT` (50) pour éviter les gels d'Outlook
- Le body est tronqué à `EMAIL_BODY_PREVIEW_LENGTH` (500 caractères)
- Les emails sont triés par date de réception décroissante

---

### `get_sent_emails`

Récupère les emails envoyés.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `limit` | int | 5 | Nombre max d'emails à retourner |

#### Retour

Format identique à `get_inbox_emails`.

#### Exemple

```python
get_sent_emails(limit=10)
```

---

### `search_emails`

Recherche des emails dans les dossiers standards.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `query` | str | *requis* | Terme de recherche |
| `folder` | str | "inbox" | Dossier ("inbox", "sent", "drafts", "deleted", "all") |
| `limit` | int | 20 | Nombre max de résultats |

#### Retour

```json
{
  "success": true,
  "query": "project alpha",
  "count": 5,
  "emails": [...]
}
```

#### Exemple

```python
# Rechercher dans la boîte de réception
search_emails(query="meeting", folder="inbox", limit=10)

# Rechercher partout
search_emails(query="budget", folder="all", limit=50)
```

#### Notes

- Recherche dans le sujet, le body et l'expéditeur
- Utilise la syntaxe DASL d'Outlook pour l'efficacité
- folder="all" cherche dans inbox, sent et drafts

---

### `send_email`

Envoie un email via Outlook.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `to` | str | *requis* | Destinataire(s), séparés par ";" |
| `subject` | str | *requis* | Sujet de l'email |
| `body` | str | *requis* | Contenu (texte brut) |
| `cc` | str | None | Destinataires en copie |
| `bcc` | str | None | Destinataires en copie cachée |
| `importance` | str | "normal" | "low", "normal" ou "high" |
| `html_body` | str | None | Contenu HTML (prioritaire sur body) |
| `signature_name` | str | None | Nom de la signature Outlook |

#### Retour

```json
{
  "success": true,
  "message": "Email sent to colleague@company.com"
}
```

#### Exemples

```python
# Email simple
send_email(
    to="colleague@company.com",
    subject="Meeting Follow-up",
    body="Thanks for the meeting today.",
    importance="normal"
)

# Email avec copie et importance haute
send_email(
    to="team@company.com",
    subject="Urgent: Server Down",
    body="The production server is down.",
    cc="manager@company.com",
    importance="high"
)

# Email HTML avec signature
send_email(
    to="client@example.com",
    subject="Project Update",
    html_body="<h1>Update</h1><p>The project is on track.</p>",
    signature_name="VP DXT"
)
```

#### Notes sur les Signatures

- Si `signature_name` est fourni, Outlook ajoute automatiquement la signature
- La signature est insérée via `Display(False)` pour préserver les images
- Outlook ajoute ~2 lignes blanches avant la signature (comportement natif)
- Les signatures sont cherchées dans `%APPDATA%\Microsoft\Signatures`

---

### `create_draft_email`

Crée un brouillon d'email sans l'envoyer.

#### Paramètres

Identiques à `send_email` (sauf pas de `importance`).

#### Retour

```json
{
  "success": true,
  "message": "Draft email created"
}
```

#### Exemple

```python
create_draft_email(
    to="team@company.com",
    subject="Weekly Report",
    body="Please review before sending.",
    cc="manager@company.com",
    signature_name="VP DXT"
)
```

#### Notes

- Le brouillon est sauvegardé dans le dossier Drafts
- Peut être modifié et envoyé manuellement depuis Outlook

---

### `get_email_attachments`

Liste les pièces jointes d'un email.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `entry_id` | str | *requis* | EntryID de l'email |

#### Retour

```json
{
  "success": true,
  "count": 2,
  "attachments": [
    {
      "filename": "report.pdf",
      "size": 245678,
      "type": 1,
      "index": 1
    },
    {
      "filename": "data.xlsx",
      "size": 98234,
      "type": 1,
      "index": 2
    }
  ]
}
```

#### Types de pièces jointes

- `type: 1` - Fichier standard
- `type: 5` - Item Outlook embarqué
- `type: 6` - Objet OLE

#### Exemple

```python
# Obtenir l'entry_id depuis get_inbox_emails
emails = get_inbox_emails(limit=1)
entry_id = emails["emails"][0]["entry_id"]

# Lister les pièces jointes
attachments = get_email_attachments(entry_id)
```

---

### `download_email_attachment`

Télécharge une pièce jointe sur le disque.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `entry_id` | str | *requis* | EntryID de l'email |
| `attachment_index` | int | *requis* | Index de la PJ (1-based) |
| `save_path` | str | *requis* | Chemin complet de sauvegarde |

#### Retour

```json
{
  "success": true,
  "message": "Attachment 'report.pdf' downloaded successfully",
  "saved_path": "C:/Downloads/report.pdf",
  "filename": "report.pdf",
  "size": 245678
}
```

#### Exemple

```python
download_email_attachment(
    entry_id="00000000...",
    attachment_index=1,
    save_path="C:/Users/user/Downloads/report.pdf"
)
```

#### Notes

- Les dossiers parents sont créés automatiquement
- Les fichiers existants sont écrasés sans confirmation
- L'index commence à 1 (pas 0)

---

### `send_email_with_attachments`

Envoie un email avec des pièces jointes.

#### Paramètres

Paramètres de `send_email` + :

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `attachments` | str | *requis* | Chemin(s) de fichier, séparés par ";" |

#### Retour

```json
{
  "success": true,
  "message": "Email sent to colleague@company.com",
  "attachments_added": 2
}
```

#### Exemple

```python
send_email_with_attachments(
    to="colleague@company.com",
    subject="Monthly Report",
    body="Please find attached the report and summary.",
    attachments="C:/Users/user/report.pdf; C:/Users/user/summary.xlsx",
    signature_name="VP DXT"
)
```

#### Notes

- Tous les fichiers doivent exister
- Chemins absolus recommandés
- Les gros fichiers peuvent ralentir l'envoi

---

## Outils Calendrier

### `get_calendar_events`

Récupère les événements du calendrier.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `days_ahead` | int | 7 | Nombre de jours à l'avance |
| `include_past` | bool | False | Inclure les événements passés d'aujourd'hui |

#### Retour

```json
{
  "success": true,
  "count": 3,
  "events": [
    {
      "subject": "Team Standup",
      "start": "2025-12-17 09:00:00",
      "end": "2025-12-17 09:30:00",
      "location": "Conference Room A",
      "organizer": "manager@company.com",
      "required_attendees": "team@company.com",
      "optional_attendees": "",
      "body": "Daily standup meeting",
      "is_all_day_event": false,
      "reminder_set": true,
      "reminder_minutes": 15,
      "categories": "",
      "busy_status": 2
    }
  ]
}
```

#### Busy Status

- `0` - Free
- `1` - Tentative
- `2` - Busy
- `3` - Out of Office

#### Exemple

```python
# Événements des 7 prochains jours
get_calendar_events(days_ahead=7)

# Événements d'aujourd'hui (y compris passés)
get_calendar_events(days_ahead=0, include_past=True)
```

---

### `create_calendar_event`

Crée un nouvel événement dans le calendrier.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `subject` | str | *requis* | Titre de l'événement |
| `start_time` | str | *requis* | Date/heure de début |
| `end_time` | str | *requis* | Date/heure de fin |
| `location` | str | None | Lieu |
| `body` | str | None | Description |
| `required_attendees` | str | None | Participants requis, séparés par ";" |
| `optional_attendees` | str | None | Participants optionnels |
| `reminder_minutes` | int | 15 | Minutes avant le rappel |
| `is_all_day` | bool | False | Événement toute la journée |

#### Formats de Date Supportés

- ISO: `"2025-12-20 14:00"`
- Natural language: `"tomorrow 2pm"`, `"next Monday at 9am"`

#### Retour

```json
{
  "success": true,
  "message": "Calendar event 'Team Meeting' created for 2025-12-20 14:00"
}
```

#### Exemple

```python
create_calendar_event(
    subject="Sprint Planning",
    start_time="2025-12-20 14:00",
    end_time="2025-12-20 15:30",
    location="Conference Room B",
    body="Planning for next sprint",
    required_attendees="team@company.com",
    reminder_minutes=30
)
```

#### Notes

- Si des participants sont spécifiés, une invitation est envoyée automatiquement
- Le parsing de dates utilise `python-dateutil` pour la flexibilité

---

### `search_calendar_events`

Recherche des événements par mot-clé.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `query` | str | *requis* | Terme de recherche |
| `days_range` | int | 30 | Jours à chercher (passés et futurs) |

#### Exemple

```python
# Chercher "standup" dans les 30 derniers et prochains jours
search_calendar_events(query="standup", days_range=30)

# Chercher "Conference Room A" dans la semaine
search_calendar_events(query="Conference Room A", days_range=7)
```

#### Notes

- Recherche dans le sujet ET le lieu
- Recherche insensible à la casse

---

### `get_meeting_requests`

Récupère les invitations de réunion en attente de réponse.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `days_range` | int | 30 | Jours à l'avance |

#### Retour

```json
{
  "success": true,
  "count": 2,
  "meeting_requests": [
    {
      "subject": "Team Meeting",
      "organizer": "manager@company.com",
      "start": "2025-12-20 14:00:00",
      "end": "2025-12-20 15:00:00",
      "location": "Conference Room A",
      "body": "Quarterly review...",
      "required_attendees": "team@company.com",
      "optional_attendees": "",
      "entry_id": "00000000...",
      "response_status": "Not Responded"
    }
  ]
}
```

#### Response Status

- `"Not Responded"` - Pas encore répondu
- `"Tentative"` - Accepté provisoirement

#### Exemple

```python
get_meeting_requests(days_range=7)
```

---

### `respond_to_meeting`

Répond à une invitation de réunion.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `entry_id` | str | *requis* | EntryID de la réunion |
| `response` | str | *requis* | "accept", "decline" ou "tentative" |
| `send_response` | bool | True | Envoyer la réponse à l'organisateur |
| `comment` | str | None | Commentaire optionnel |

#### Retour

```json
{
  "success": true,
  "message": "Meeting accepted and response sent"
}
```

#### Exemples

```python
# Accepter
respond_to_meeting(
    entry_id="00000000...",
    response="accept",
    send_response=True
)

# Décliner avec commentaire
respond_to_meeting(
    entry_id="00000000...",
    response="decline",
    send_response=True,
    comment="Sorry, I have a conflict."
)

# Accepter provisoirement sans notifier
respond_to_meeting(
    entry_id="00000000...",
    response="tentative",
    send_response=False
)
```

#### Notes

- `accept` : Ajoute la réunion au calendrier
- `decline` : Supprime la réunion du calendrier
- `tentative` : Marque comme provisoire
- `send_response=False` : Mise à jour silencieuse

---

## Outils Contacts

### `get_contacts`

Récupère les contacts.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `limit` | int | 50 | Nombre max de contacts |
| `search_name` | str | None | Filtre par nom |

#### Retour

```json
{
  "success": true,
  "count": 3,
  "contacts": [
    {
      "full_name": "Jane Smith",
      "email1": "jane.smith@company.com",
      "email2": "",
      "email3": "",
      "company": "Acme Corp",
      "job_title": "Product Manager",
      "business_phone": "+1-555-1234",
      "mobile_phone": "+1-555-5678",
      "home_phone": "",
      "business_address": "123 Main St",
      "categories": ""
    }
  ]
}
```

#### Exemple

```python
# Tous les contacts
get_contacts(limit=50)

# Chercher "Smith"
get_contacts(limit=20, search_name="Smith")
```

---

### `create_contact`

Crée un nouveau contact.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `full_name` | str | *requis* | Nom complet |
| `email` | str | *requis* | Email principal |
| `company` | str | None | Entreprise |
| `job_title` | str | None | Titre du poste |
| `business_phone` | str | None | Téléphone professionnel |
| `mobile_phone` | str | None | Mobile |
| `home_phone` | str | None | Téléphone personnel |

#### Exemple

```python
create_contact(
    full_name="John Doe",
    email="john.doe@techcorp.com",
    company="TechCorp",
    job_title="Senior Engineer",
    business_phone="+1-555-1234"
)
```

---

### `search_contacts`

Recherche des contacts.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `query` | str | *requis* | Terme de recherche |

#### Exemple

```python
# Chercher par nom
search_contacts(query="John Smith")

# Chercher par email
search_contacts(query="@acmecorp.com")

# Chercher par entreprise
search_contacts(query="Acme Corp")
```

#### Notes

- Recherche dans nom, email ET entreprise
- Insensible à la casse
- Pas de limite (tous les contacts correspondants)

---

## Outils Dossiers

### `list_outlook_folders`

Liste tous les dossiers Outlook.

#### Retour

```json
{
  "success": true,
  "count": 25,
  "folders": [
    {
      "name": "Inbox",
      "path": "Inbox"
    },
    {
      "name": "Archive",
      "path": "Inbox/Archive"
    },
    {
      "name": "Personal",
      "path": "Personal"
    },
    {
      "name": "My Mails",
      "path": "Personal/My Mails"
    }
  ]
}
```

#### Exemple

```python
list_outlook_folders()
```

#### Notes

- **Ultra-rapide** : N'inclut PAS les compteurs d'items (évite les gels)
- Inclut tous les dossiers récursivement
- Saute les dossiers système inaccessibles

---

### `search_emails_in_custom_folder`

Recherche des emails dans un dossier personnalisé.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `folder_path` | str | *requis* | Chemin du dossier |
| `query` | str | None | Terme de recherche (optionnel) |
| `limit` | int | 20 | Nombre max de résultats |
| `days_back` | int | 2 | Jours en arrière à chercher |

#### Retour

```json
{
  "success": true,
  "folder": "Personal/My Mails",
  "query": "invoice",
  "count": 5,
  "days_back": 2,
  "info": "Searched emails from last 2 days only",
  "emails": [...]
}
```

#### Exemples

```python
# Chercher dans "Personal/My Mails" (2 derniers jours)
search_emails_in_custom_folder(
    folder_path="Personal/My Mails"
)

# Chercher avec query et plus de jours
search_emails_in_custom_folder(
    folder_path="Personal/My Mails",
    query="invoice",
    days_back=30,
    limit=50
)

# ATTENTION : Chercher TOUS les emails (peut geler Outlook!)
search_emails_in_custom_folder(
    folder_path="Personal/My Mails",
    days_back=0  # 0 = tous les emails
)
```

#### Notes

- **IMPORTANT** : `days_back=2` par défaut pour éviter les gels
- `days_back=0` ou négatif cherche TOUS les emails (très lent)
- Utilisez `list_outlook_folders()` pour trouver les chemins
- Le chemin est sensible à la casse

---

## Outils Out-of-Office

### `get_out_of_office_settings`

Récupère les paramètres de réponse automatique.

#### Retour

```json
{
  "success": true,
  "enabled": true,
  "scheduled": true,
  "start_time": "2025-12-20 00:00:00",
  "end_time": "2025-12-27 00:00:00",
  "internal_reply": "I'm out of office until next week.",
  "external_reply": "I'm currently unavailable.",
  "external_audience": "Known"
}
```

#### External Audience

- `"None"` - Pas de réponses externes
- `"Known"` - Seulement contacts/organisation
- `"All"` - Tous les expéditeurs

#### Exemple

```python
get_out_of_office_settings()
```

---

### `set_out_of_office`

Configure les réponses automatiques.

#### Paramètres

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `enabled` | bool | *requis* | Activer les réponses auto |
| `internal_reply` | str | *requis* | Message pour internes |
| `external_reply` | str | None | Message pour externes (défaut: internal_reply) |
| `external_audience` | str | "Known" | "None", "Known" ou "All" |
| `scheduled` | bool | False | Planifier dans le temps |
| `start_time` | str | None | Date/heure de début (requis si scheduled) |
| `end_time` | str | None | Date/heure de fin (requis si scheduled) |

#### Exemples

```python
# Activer immédiatement
set_out_of_office(
    enabled=True,
    internal_reply="I'm out of office until next week.",
    external_reply="I'm currently unavailable.",
    external_audience="Known"
)

# Planifier pour des dates spécifiques
set_out_of_office(
    enabled=True,
    internal_reply="On vacation",
    external_reply="I'm on vacation until Dec 27th.",
    scheduled=True,
    start_time="2025-12-20 00:00",
    end_time="2025-12-27 00:00",
    external_audience="All"
)
```

#### Notes

- Format de date : ISO `"YYYY-MM-DD HH:MM"`
- Les réponses planifiées se désactivent automatiquement après `end_time`
- Requiert **Outlook 2010+ avec Exchange Server**

---

### `disable_out_of_office`

Désactive les réponses automatiques.

#### Retour

```json
{
  "success": true,
  "message": "Out-of-Office disabled"
}
```

#### Exemple

```python
disable_out_of_office()
```

#### Notes

- Les messages sont préservés (pas supprimés)
- Peut être réactivé plus tard avec `set_out_of_office`

---

## Gestion des Erreurs

### Format Standard

Tous les outils retournent un format d'erreur cohérent :

```json
{
  "success": false,
  "error": "Description de l'erreur"
}
```

### Erreurs Communes

#### Outlook non accessible

```json
{
  "success": false,
  "error": "Unable to connect to Outlook. Make sure Outlook is installed and properly configured."
}
```

**Solutions** :
- Vérifier qu'Outlook est en cours d'exécution
- Vérifier qu'un compte est configuré
- Redémarrer Outlook
- Exécuter en tant qu'administrateur

#### Dossier non trouvé

```json
{
  "success": false,
  "error": "Folder 'Custom/Path' not found. Use list_outlook_folders() to see available folders."
}
```

**Solutions** :
- Utiliser `list_outlook_folders()` pour voir les chemins
- Vérifier la casse du chemin
- Vérifier que le dossier existe

#### EntryID invalide

```json
{
  "success": false,
  "error": "Failed to get item from EntryID"
}
```

**Solutions** :
- Vérifier que l'EntryID est valide
- L'email a peut-être été supprimé

#### Fichier non trouvé

```json
{
  "success": false,
  "error": "Attachment file not found or couldn't be attached: C:/file.pdf"
}
```

**Solutions** :
- Vérifier le chemin du fichier
- Utiliser un chemin absolu
- Vérifier les permissions

#### Feature non disponible

```json
{
  "success": false,
  "error": "Out-of-Office settings not accessible via COM automation on this Outlook version."
}
```

**Solutions** :
- Feature requiert Outlook 2010+ avec Exchange
- Utiliser l'interface Outlook directement
- Vérifier la version d'Outlook

---

## Performances

### Optimisations Implémentées

#### 1. Cache de Dossiers

```python
_FOLDER_CACHE: Dict[str, Any] = {}
```

- Cache les objets dossiers résolus
- **Gain** : 45x plus rapide sur recherches répétées
- Première recherche : ~45s
- Recherches suivantes : ~1s

#### 2. Filtre par Date

```python
DEFAULT_DAYS_BACK = 2  # Seulement 2 derniers jours par défaut
```

- Réduit drastiquement le nombre d'emails à parcourir
- Utilise `Restrict()` côté serveur AVANT itération
- **Gain** : Secondes au lieu de minutes

#### 3. Indexation Directe

```python
mail = items[i + 1]  # Au lieu de GetFirst()/GetNext()
```

- Plus rapide sur collections filtrées
- Évite l'appel coûteux à `items.Count`

#### 4. Limites Réduites

```python
DEFAULT_EMAIL_LIMIT = 5    # Au lieu de 10
MAX_EMAIL_LIMIT = 50       # Au lieu de 100
```

- Moins d'emails = moins de gel d'Outlook

#### 5. Pas de Compteurs dans list_outlook_folders

```python
_get_all_folders(folder, include_counts=False)
```

- `folder.Items.Count` peut prendre **plusieurs minutes**
- **Gain** : Quelques secondes au lieu de minutes

### Benchmarks

| Opération | Avant | Après |
|-----------|-------|-------|
| Recherche dossier (1ère fois) | ~45s | ~45s |
| Recherche dossier (cache) | ~45s | ~1s (45x) |
| list_outlook_folders() | Minutes | Secondes |
| Recherche emails (2j) | Variable | Rapide |

### Limitations de Performance

**Outlook COM est single-threaded** :
- Pendant une requête MCP, Outlook ne répond pas aux clics
- C'est une limitation architecturale de l'API COM
- Le gel est **réduit** mais **pas éliminé**

### Recommandations

1. **Utiliser des dossiers spécifiques** (moins d'emails)
2. **Réduire `days_back`** au minimum nécessaire
3. **Réduire `limit`** si peu de résultats suffisent
4. **Fermer Outlook** pendant l'utilisation intensive du MCP

---

## Sécurité

### Accès aux Données

- Utilise Windows COM automation (pas de credentials stockés)
- Toutes les opérations utilisent les permissions du profil Outlook actuel
- Outlook doit être en cours d'exécution

### Body Truncation

```python
EMAIL_BODY_PREVIEW_LENGTH = 500  # Limite à 500 caractères
```

- Empêche la fuite excessive de données
- Réduit l'utilisation de tokens pour l'IA

### Logging

- Logging en mode CRITICAL (minimal)
- Pas de contenu sensible dans les logs
- Les emails/BCC ne sont PAS loggés

### Bonnes Pratiques

1. **Ne pas stocker d'EntryID** dans des fichiers partagés
2. **Vérifier les permissions** avant de télécharger des PJ
3. **Utiliser BCC** pour les emails de masse
4. **Respecter la politique de l'organisation**

---

## Limitations

### Plateforme

- **Windows uniquement** (API COM)
- Impossible sur Linux/macOS

### Outlook

- **Outlook doit être installé** et en cours d'exécution
- Un **compte configuré** est requis
- Fonctionne avec le **profil par défaut** uniquement

### Performance

- **Single-threaded** : Outlook gèle pendant les requêtes
- **Boîtes mail volumineuses** : Recherches potentiellement lentes
- **Limite de 50 emails** par requête

### Fonctionnalités

#### Out-of-Office
- Requiert **Outlook 2010+ avec Exchange**
- Ne fonctionne **pas avec POP3/IMAP**
- Peut ne pas être accessible via COM sur certaines configurations

#### Pièces Jointes
- **Types supportés** : Fichiers standards uniquement
- Les objets OLE et items embarqués ne peuvent pas être téléchargés
- **Taille** : Les gros fichiers peuvent ralentir

#### Recherche
- **Syntaxe DASL** : Limitée par Outlook
- **Cas sensible** : Chemins de dossiers
- **Cache** : Invalidé au redémarrage d'Outlook

---

## Support et Contribution

### Obtenir de l'Aide

- **Issues GitHub** : [Créer une issue](https://github.com/YOUR_USERNAME/mcp-outlook/issues)
- **Documentation** : [README.md](README.md), [EXAMPLES.md](EXAMPLES.md)
- **Tests** : Exécuter `python tests/test_connection.py`

### Contribuer

Voir [CONTRIBUTING.md](CONTRIBUTING.md) pour :
- Comment ouvrir une pull request
- Conventions de code
- Guide de test
- Roadmap

### Informations Utiles pour les Issues

Quand vous créez une issue, incluez :

- Version de Windows
- Version d'Outlook
- Version de Python
- Output de `python tests/test_connection.py`
- Message d'erreur complet
- Steps to reproduce

---

## Historique des Versions

Voir [CHANGELOG.md](CHANGELOG.md) pour l'historique détaillé.

### Version Actuelle : 1.2.0

**Nouveautés** :
- Gestion des pièces jointes (3 outils)
- Réponse aux invitations de réunion (2 outils)
- Paramètres Out-of-Office (3 outils)
- Métadonnées enrichies dans les emails

**Documentation Complète** : Voir [CHANGELOG.md](CHANGELOG.md)

---

## Annexes

### Constantes Outlook

```python
# Folders
OUTLOOK_FOLDER_INBOX = 6
OUTLOOK_FOLDER_SENT = 5
OUTLOOK_FOLDER_DRAFTS = 16
OUTLOOK_FOLDER_DELETED = 3
OUTLOOK_FOLDER_CALENDAR = 9
OUTLOOK_FOLDER_CONTACTS = 10

# Item types
OUTLOOK_ITEM_MAIL = 0
OUTLOOK_ITEM_APPOINTMENT = 1
OUTLOOK_ITEM_CONTACT = 2

# Importance
IMPORTANCE_LOW = 0
IMPORTANCE_NORMAL = 1
IMPORTANCE_HIGH = 2
```

### Structure Email Complète

```json
{
  "subject": "string",
  "sender": "string",
  "sender_email": "string",
  "recipients": "string",
  "cc": "string",
  "bcc": "string",
  "received_time": "datetime",
  "sent_on": "datetime",
  "body": "string (truncated)",
  "body_length": "int",
  "has_attachments": "bool",
  "attachment_count": "int",
  "attachments": [
    {
      "filename": "string",
      "size": "int (bytes)",
      "type": "int"
    }
  ],
  "importance": "int (0-2)",
  "unread": "bool",
  "categories": "string",
  "entry_id": "string"
}
```

---

**Version** : 1.2.0  
**Date** : 17 décembre 2025  
**Auteur** : MCP Outlook Contributors  
**Licence** : MIT

