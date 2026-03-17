# Userflow — Administration Système Entis Mutuelles

**Version :** 2.0  
**Date :** Mars 2026

---

## Présentation

Userflow est un outil d'administration système qui automatise la gestion du cycle de vie des utilisateurs dans le SI du Groupe Entis Mutuelles.

**Fonctionnalités :**
- Création de comptes utilisateurs AD
- Attribution de licences Microsoft 365
- Remise de matériel informatique
- Gestion du parc GLPI
- Processus de départ utilisateur (offboarding)
- Synchronisation RH

---

## Installation

### 1. Récupérer les fichiers

```bash
git clone https://github.com/dsi-entis-mutuelles/ohmyadmin.git
cd ohmyadmin
```

### 2. Créer le dossier de configuration

Créez un dossier `config/` contenant :

```
config/
├── settings.json
├── Adresse_Agences.json
├── GD_Agence.json
├── Groupe_service.json
└── Letsign-it.json
```

### 3. Configurer le coffre-fort KeePass

Le fichier `python-vault.kdbx` doit contenir :

| Entry | Champ | Valeur |
|-------|-------|--------|
| GitHub | Password | Token d'accès GitHub |
| Azure App | Username | Client ID |
| Azure App | Password | Client Secret |

---

## Utilisation

### Script recommandé : `userflow.py`

C'est le script unique qui fait tout automatiquement.

```bash
python userflow.py
```

**Étapes :**
1. Le script crée automatiquement le venv et installe les dépendances
2. Il charge la configuration depuis le dossier `config/`
3. Il demande le mot de passe KeePass (1 seule saisie)
4. Il se connecte à SharePoint silencieusement via Azure App
5. Il téléchargera les modules depuis GitHub
6. Il demande les identifiants AD (1 saisie)
7. Le menu principal s'affiche

---

## Configuration Azure

Pour SharePoint silent, créez une App Registration Azure avec ces permissions :

- Sites.Read.All
- Sites.ReadWrite.All
- User.Read.All
- Directory.Read.All

Accordez le Admin Consent.

---

## Structure des fichiers

```
ohmyadmin/
├── userflow.py           # Script principal (recommandé)
├── hub_central.py       # Orchestrateur alternatif
├── requirements.txt     # Dépendances Python
├── config/              # Fichiers de configuration
│   ├── settings.json
│   ├── Adresse_Agences.json
│   ├── GD_Agence.json
│   ├── Groupe_service.json
│   └── Letsign-it.json
└── mod_*.py            # Modules (téléchargés depuis GitHub)
```

---

## Dépendances

Les bibliothèques nécessaires sont automatiquement installées :

- `requests` - Appels HTTP
- `packaging` - Gestion de versions
- `pykeepass` - Coffre KeePass
- `office365` - SharePoint API
- `ldap3` - Active Directory
- `pandas` - Fichiers Excel
- `reportlab` - Génération PDF

---

## Sécurité

- Les identifiants ne sont jamais stockés en clair
- Le KeePass contient uniquement les secrets (GitHub, Azure)
- Les identifiants AD sont saisis manuellement à chaque session
- SharePoint utilise l'authentification Azure App (silent)

---

## Contribution

Pour ajouter un nouveau module :

1. Créer un fichier `mod_nom_module.py`
2. Implémenter une fonction `async def run(...)`
3. Le fichier sera détecté et chargé automatiquement

---

*Document mis à jour Mars 2026*
