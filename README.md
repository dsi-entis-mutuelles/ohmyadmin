# Userflow — Administration Système Entis Mutuelles

**Version :** 2.1  
**Date :** Mars 2026

---

## Présentation

Userflow est un outil d'administration système **unique** qui automatise la gestion du cycle de vie des utilisateurs.

**Un seul fichier à utiliser : `userflow.py`**

---

## Installation

### 1. Récupérer le projet

```bash
git clone https://github.com/dsi-entis-mutuelles/ohmyadmin.git
cd ohmyadmin
```

### 2. Créer settings.json

Créez un fichier `settings.json` avec vos paramètres :

```json
{
    "github": {
        "repo_owner": "dsi-entis-mutuelles",
        "repo_name": "ohmyadmin"
    },
    "sharepoint": {
        "glpi_site_url": "https://entis.sharepoint.com/sites/GLPI",
        "arrivants_site_url": "https://entis-my.sharepoint.com/...",
        "app_client_id": "VOTRE_CLIENT_ID",
        "tenant_id": "VOTRE_TENANT_ID",
        "config_library": "/sites/GLPI/Data",
        "keepass_path": "/sites/GLPI/Data/python-vault.kdbx"
    },
    "active_directory": {
        "server_fqdn": "srvad04.votredomaine.pri",
        "domain": "votredomaine.pri"
    },
    "modules": {
        "script_bases": ["mod_user_create", "mod_license_assign", ...]
    },
    "referentiels": ["Adresse_Agences.json", "GD_Agence.json", ...]
}
```

### 3. Configurer le coffre-fort KeePass

Ajoutez dans votre fichier `python-vault.kdbx` :

| Entry | Champ | Valeur |
|-------|-------|--------|
| GitHub | Password | Token GitHub |
| Azure App | Username | Client ID Azure |
| Azure App | Password | Client Secret Azure |

---

## Utilisation

```bash
python userflow.py
```

**Le script fait TOUT automatiquement :**

1. Crée le venv et installe les dépendances
2. Charge `settings.json` localement
3. Se connecte à SharePoint (MFA navigateur)
4. Télécharge la configuration complète (referentiels)
5. Demande le mot de passe KeePass
6. Se reconnecte via Azure App (silent)
7. Télécharge les modules depuis GitHub
8. Demande les identifiants AD
9. Affiche le menu principal

---

## Configuration Azure

Pour SharePoint silent, créez une App Registration Azure avec :

**Permissions applicatives :**
- Sites.Read.All
- Sites.ReadWrite.All
- User.Read.All
- Directory.Read.All

Accordez le **Admin Consent**.

---

## Fichiers

```
ohmyadmin/
├── userflow.py       # Script unique (A UTILISER)
├── settings.json     # Configuration minimale (A CREER)
├── requirements.txt # Dépendances
└── README.md
```

---

## Comment ça marche

```
python userflow.py
    │
    ├─> 1. Auto venv + dependances
    │
    ├─> 2. Lit settings.json local
    │
    ├─> 3. Auth SharePoint (navigateur)
    │
    ├─> 4. Telecharge config depuis SharePoint
    │       - Adresse_Agences.json
    │       - GD_Agence.json
    │       - Groupe_service.json
    │       - Letsign-it.json
    │
    ├─> 5. Mot de passe KeePass (1 saisie)
    │       - GitHub token
    │       - Azure credentials
    │
    ├─> 6. SharePoint silent via Azure App
    │
    ├─> 7. Telecharge modules depuis GitHub
    │
    ├─> 8. Identifiants AD (1 saisie)
    │
    └─> 9. Menu principal
```

---

## Dépendances

Installées automatiquement :
- requests, packaging, pykeepass
- office365, ldap3, pandas
- reportlab, azure-identity

---

*Document mis à jour Mars 2026*
