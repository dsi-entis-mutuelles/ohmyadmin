# Userflow — HUB d'Administration Entis Mutuelles

## Document de Spécification Produit (PRD)

**Version du document :** 1.3  
**Date de création :** Mars 2026  
**Dernière mise à jour :** Mars 2026  
**Auteur :** JEL pour Groupe Entis  
**Statut :** Actif — En évolution continue

---

## 1. Vue d'Ensemble du Projet

### 1.1 Description fonctionnelle

Userflow est une plateforme d'orchestration centralisée destinée à automatiser l'ensemble du cycle de vie des utilisateurs au sein du système d'information du Groupe Entis Mutuelles. Cette suite d'outils permet aux administrateurs système de gérer de manière intégrés les processus d'arrivée, de mobilité et de départ des collaborateurs, en s'appuyant sur l'écosystème Microsoft 365 et les outils de gestion du parc informatique existants.

Le projet répond à un besoin critique d'harmonisation et d'automatisation des tâches administratives récurrentes liées à la gestion des identité numériques. Chaque année, des centaines de nouveaux collaborateurs rejoignent le groupe tandis que d'autres quittent l'entreprise. Ces mouvements génèrent une charge de travail significative pour les équipes du service informatique, qui doivent créer des comptes, attribuer des licences, configurer des postes de travail, gérer les droits d'accès et enfin désactiver les accès lors des départs. Userflow permet de réduire considérablement ces manipulations manuelles tout en minimisant les risques d'erreurs.

La plateforme s'intègre nativement avec plusieurs composants majeurs de l'infrastructure : l'Active Directory local sur Windows Server, Azure Active Directory (désormais Microsoft Entra ID), Microsoft Intune pour la gestion des appareils mobiles, Exchange Online pour la messagerie, SharePoint pour le stockage documentaire, et GLPI pour la gestion du parc informatique. Cette intégration profonde permet de garantir la cohérence des données across l'ensemble du système d'information et d'offrir une expérience utilisateur fluide aux administrateurs.

### 1.2 Objectifs stratégiques

L'objectif principal de ce projet est d'atteindre une FULL automatisation des workflows d'entrée et sortie utilisateur tout en maintenant un niveau élevé de traçabilité et de sécurité. Les objectifs secondaires incluent la réduction du temps de traitement des demandes, l'amélioration de la qualité des données dans les annuaires, et la simplification des procédures pour les équipes du support informatique.

### 1.3 Périmètre fonctionnel

Le périmètre de Userflow couvre l'ensemble des opérations suivantes : la création de comptes utilisateurs dans l'Active Directory local avec génération automatique d'identifiants uniques, l'attribution de licences Microsoft 365 et la configuration des boites aux lettres Exchange Online, la gestion des (postes Windows et Mac) via Intune et l'Active Directory, la remise de matériel aux nouveaux arrivants avec génération de documents contractuels, la gestion centralisée du parc GLPI incluant l'inventaire, les attributions et les restitutions, et enfin le processus complet de départ utilisateur avec désactivation, archivage et suppression des ressources.

---

## 2. Architecture Technique

### 2.1 Vue d'ensemble de l'architecture

L'architecture de Userflow repose sur un modèle hub-and-spoke où le script principal `hub_central.py` agit comme orchestrateur central et les modules spécialisés traitent des cas d'usage spécifiques. Cette architecture présente l'avantage de permettre une maintenance facilitée des composants individuels tout en garantissant une cohérence globale grâce à des mécanismes de partage de session et de configuration centralisés.

Le hub central est responsable de l'amorçage de l'environnement, de l'authentification auprès des différents services, du chargement dynamique des modules et de la gestion du flux utilisateur principal. Chaque module est conçu pour fonctionner de manière autonome une fois chargé, mais bénéficie des éléments de contexte partagés établis par le hub.

### 2.2 Composants principaux

Le système se compose de sept éléments scripts Python distincts, chacun remplissant une fonction métier précise selon la politique de nommage quality friendly. Le premier composant, `hub_central.py` (anciennement `main_v1.py`), constitue le point d'entrée unique de l'application et l'orchestrateur central. Il gère l'authentification SSO via SharePoint, le déverrouillage du coffre KeePass contenant les secrets, la synchronisation des modules et de la configuration depuis GitHub, l'établissement de la connexion LDAP vers l'Active Directory local, le chargement dynamique des modules métier, et enfin l'affichage du menu principal avec gestion des choix utilisateurs.

Le deuxième composant, `mod_user_create.py` (anciennement `creer-utilisateur_v1.py`), est dédié à la création des comptes utilisateurs dans l'Active Directory local. Il lit les demandes de nouveaux arrivants depuis SharePoint, calcule dynamiquement les identifiants uniques en évitant les homonymes, crée les objets utilisateur dans l'OU appropriée, attribue les groupes de sécurité selon le profil, et met à jour le fichier de suivi SharePoint.

Le troisième composant, `mod_license_assign.py` (anciennement `attribuer-licence_v1.py`), gère l'attribution des licences Microsoft 365. Il permet d'afficher l'état des stocks de licences, de sélectionner les packs de licences à attribuer, de configurer les boites aux lettres Exchange Online avec le fuseau horaire approprié, de gérer l'authentification multifacteur, et de configurer les signatures email Letsign-it via les attributs personnalisés de l'Active Directory.

Le quatrième composant, `mod_device_provision.py` (anciennement `remise-pc_v1.py`), traite la remise de matériel informatique aux nouveaux arrivants. Il peut fonctionner en mode automatique depuis le registre SharePoint ou en mode manuel, attribue les appareils dans Intune en tant que Primary User, déplace les ordinateurs dans les OU Active Directory correspondantes, génère la fiche d'attribution PDF avec les clauses contractuelles, et archive le document sur SharePoint.

Le cinquième composant, `mod_asset_glpi.py` (anciennement `gestion-parc-glpi_v1.py`), offre une interface de gestion du parc GLPI. Il permet de lister le matériel affecté à un utilisateur, d'attribuer du nouveau matériel depuis le stock, de gérer les restitutions de manière groupée, et de générer les bons de restitution PDF.

Le sixième composant, `mod_user_offboard.py` (anciennement `sortie-utilisateur_v1.py`), automatise le processus de départ. Il effectue une recherche multicritère dans l'Active Directory et Azure, désactive les comptes AD et les déplace vers l'OU d'archivage, retire les licences Microsoft 365 via les groupes de sécurité, supprime les appareils Intune et Azure AD associés, et génère un rapport d'audit des opérations effectuées.

Le septième composant, `mod_hr_sync.py` (anciennement `verification-rh_v1.py`), réalise la synchronisation et la vérification des données RH. Il compare les informations du fichier Excel RH avec l'Active Directory, vérifie et met à jour les attributs (manager, fonction, département, téléphone mobile), contrôle les numéros de téléphone dans Intune, et génère un rapport PDF des modifications effectuées.

### 2.3 Flux de données et dépendances

Le flux de données global du système peut se résumer ainsi : l'utilisateur lance le hub central, celui-ci s'authentifie sur SharePoint et récupère la configuration centralisée, puis il déverrouille le coffre KeePass pour accéder aux credentials, établit la connexion LDAP avec l'Active Directory, synchronise les modules métier et les fichiers de configuration, présente le menu principal, et enfin exécute les modules sélectionnés en leur transmettant les objets de connexion partagés.

Chaque module communique avec les services externos selon ses besoins spécifiques. Les modules de création et de gestion des utilisateurs dialoguent avec l'Active Directory local via le protocole LDAP. Les modules d'attribution de licences, de remise de matériel et de sortie utilisateur interagissent avec Microsoft Graph API pour Azure, Intune et Exchange Online. Les modules de gestion du parc utilisent l'API REST de GLPI. Enfin, plusieurs modules lisent et écrivent des fichiers Excel ou JSON sur SharePoint pour le suivi des demandes et le stockage des configurations.

---

## 2.4 Politique de Nommage des Fichiers

### 2.4.1 Principes fondamentaux

La politique de nommage des fichiers Python dans le projet Userflow suit une convention « Quality Friendly » visant à garantir lisibilité, traçabilité et maintenabilité. Cette politique définit des règles strictes pour identifier rapidement la nature, le rôle et le statut de chaque composant sans avoir à ouvrir le fichier.

Tous les noms de fichiers utilisent exclusively le format snake_case (lettres minuscules avec underscores comme séparateurs). Cette convention est cohérente avec les standards Python PEP 8 et facilite le tri alphabétique naturel dans les gestionnaires de fichiers.

### 2.4.2 Structure de nommage

Le format recommandé pour les noms de fichiers est le suivant :

```
[préfixe]_[nom_descriptif][_qualité].py
```

Chaque composante a une signification précise. Le préfixe indique la catégorie du fichier, le nom descriptif identifie la fonctionnalité, et le suffixe optionnel de qualité indique le statut de maturité du code.

### 2.4.3 Préfixes standardisés

Les préfixes suivants doivent être utilisés selon la catégorie du fichier :

| Préfixe | Catégorie | Description |
|---------|-----------|-------------|
| `hub_` | Hub central | Point d'entrée principal de l'application |
| `mod_` | Module métier | Modules fonctionnels thématiques |
| `util_` | Utilitaires | Fonctions et classes helper |
| `lib_` | Librairies | Composants réutilisables |
| `test_` | Tests | Scripts de test automatisés |
| `setup_` | Installation | Scripts de configuration et déploiement |

### 2.4.4 Indicateurs de qualité (suffixes)

Les suffixes optionnels permettent d'indiquer le statut de maturité du composant. Ces indicateurs sont particulièrement utiles lors des phases de développement et de transition entre versions.

| Suffixe | Signification | Usage |
|---------|---------------|-------|
| `_draft` | Brouillon en cours de développement | Nouvelles fonctionnalités en cours de test |
| `_beta` | Fonctionnalité en phase de validation | Tests utilisateurs en cours |
| `_stable` | Fonctionnalité éprouvée | Prête pour la production |
| `_legacy` | Ancienne version conservée pour compatibilité | À migrer vers la nouvelle version |
| `_deprecated` | Obsolète | À ne plus utiliser, sera supprimé |

### 2.4.5 Convention de versioning

Le numéro de version ne doit plus apparaître dans le nom du fichier. Le versioning est géré séparément via le dépôt Git et les tags de release. Le format de version utilise le schéma sémantique MAJOR.MINOR (exemple : v1.0, v2.1).

Cette approche présente plusieurs avantages : les noms de fichiers restent stables entre les versions mineures, le système de détection automatique des versions via GitHub (dans le hub central) gère déjà la résolution de la dernière version, et la suppression de la version dans le nom simplifie les références dans le code et la documentation.

### 2.4.6 Tableau de correspondance des noms

Le tableau suivant présente la correspondance entre les noms actuels et les nouveaux noms selon la politique de nommage quality friendly :

| Ancien nom | Nouveau nom | Catégorie | Statut |
|------------|-------------|-----------|--------|
| `main_v1.py` | `hub_central.py` | Hub | _stable |
| `creer-utilisateur_v1.py` | `mod_user_create.py` | Module | _stable |
| `attribuer-licence_v1.py` | `mod_license_assign.py` | Module | _stable |
| `remise-pc_v1.py` | `mod_device_provision.py` | Module | _stable |
| `gestion-parc-glpi_v1.py` | `mod_asset_glpi.py` | Module | _stable |
| `sortie-utilisateur_v1.py` | `mod_user_offboard.py` | Module | _stable |
| `verification-rh_v1.py` | `mod_hr_sync.py` | Module | _stable |

### 2.4.7 Règles de migration

Lors de la migration vers la nouvelle politique de nommage, les règles suivantes s'appliquent. Premièrement, la rétrocompatibilité temporaire est assurée en conservant les anciens fichiers avec un suffixe `_legacy` pendant une période de transition de deux releases. Deuxièmement, les nouveaux modules suivent impérativement la nouvelle convention dès leur création. Troisièmement, le fichier `settings.json` doit être mis à jour avec les nouveaux noms de modules dans la section `script_bases`. Quatrièmement, après validation complète, les fichiers marqués `_legacy` peuvent être supprimés du dépôt.

### 2.4.8 Exemples de nouveaux noms

Voici quelques exemples illustrant l'application de la politique pour de futures évolutions :

- `hub_central.py` (hub central, version stable)
- `mod_user_create.py` (module création utilisateur)
- `mod_user_create_draft.py` (nouvelle version en développement)
- `mod_user_create_v2.py` → renommé en `mod_user_create.py` une fois stable
- `util_ldap_helpers.py` (librairie d'aide LDAP)
- `setup_environment.py` (script d'installation)

---

## 3. Spécifications des Modules

### 3.1 Module de Création d'Utilisateurs (mod_user_create.py)

Ce module constitue le point de départ du processus d'intégration d'un nouveau collaborateur. Il permet de transformer une demande d'arrivée, généralement soumise via un formulaire SharePoint, en un compte Active Directory pleinement opérationnel. Le processus suit plusieurs étapes clés qui garantissent la qualité et la cohérence des données créées.

La première étape consiste à récupérer les demandes depuis le fichier Excel hébergé sur SharePoint. Le module lit le formulaire qui contient les informations essentielles : nom, prénom, lieu de travail, service, fonction, type de contrat, date de fin de contrat éventuelle, et équipements demandés. L'administrateur peut sélectionner les demandes à traiter parmi les plus récentes.

La deuxième étape est le calcul de l'identifiant unique. Le module génère automatiquement un login (sAMAccountName) et une adresse email en évitant les conflits avec les comptes existants. Pour les noms composés, il utilise des stratégies de variation : ajout d'initiales, troncature du prénom, ou numérotation en dernier recours. Cette logique garantit que chaque collaborateur dispose d'identifiants uniques et cohérents avec la convention de nommage de l'entreprise.

La troisième étape détermine l'organisation dans l'Active Directory. En fonction du lieu de travail indiqué dans le formulaire, le module identifie l'OU (Organizational Unit) de destination et les groupes de sécurité à attribuer. La configuration des agences est stockée dans des fichiers JSON qui définissent les correspondances entre les villes, les OU et les groupes.

La quatrième étape crée effectivement le compte utilisateur dans l'Active Directory avec tous les attributs nécessaires : nom complet, prénom, nom, email, description, bureau, adresse, téléphone, fonction, département, entreprise, et mot de passe temporaire. Les groupes de sécurité sont ajoutés post-création.

Enfin, le module met à jour le fichier de suivi SharePoint en ajoutant les nouveaux utilisateurs créés avec le statut « À traiter » pour les étapes suivantes (attribution de licences, remise de matériel).

### 3.2 Module d'Attribution de Licences (mod_license_assign.py)

Ce module intervient après la création du compte utilisateur pour configurer l'environnement Microsoft 365. Il permet d'attribuer les licences appropriées selon le profil du collaborateur et de paramétrer les différents services.

Le module commence par afficher l'état des stocks de licences disponibles. Il interroge Microsoft Graph pour récupérer les informations sur les licences souscrites et calcule le nombre de licences restantes pour chaque type. Cette visibilité permet à l'administrateur de faire des choix éclairés.

L'administrateur peut ensuite sélectionner le pack de licences souhaité parmi plusieurs options : M365 Business Premium, Exchange Online Plan 2, M365 E3 sans Teams, Microsoft Teams EEA, ou des combinaisons personnalisées. L'attribution se fait via l'appartenance à des groupes de sécurité spécifiques dans Azure AD, qui déclenchent automatiquement le provisionnement des licences.

Le module configure également la boite aux lettres Exchange Online. Il récupère les informations de fuseau horaire et de langue depuis la configuration de l'agence et applique ces paramètres à la BAL de l'utilisateur. Une logique de retry est implémentée pour gérer le délai de provisionnement des boites aux lettres.

La configuration de l'authentification multifacteur (MFA) est également prise en charge. Le module vérifie si l'utilisateur a déjà des méthodes d'authentification enregistrées et peut, sur confirmation, l'ajouter au groupe de sécurité MFA qui强制 l'authentification forte.

Enfin, le module gère la configuration des signatures email Letsign-it. Il utilise les attributs personnalisés extensionAttribute2 et extensionAttribute3 de l'Active Directory pour stocker les identifiants de signature. L'administrateur peut sélectionner le modèle de signature approprié depuis une liste définie dans les fichiers de configuration.

### 3.3 Module de Remise de Matériel (mod_device_provision.py)

Ce module gère l'attribution physique des équipements informatiques aux nouveaux collaborateurs. Il peut fonctionner en mode automatique, en lisant les demandes depuis le registre SharePoint, ou en mode manuel, en recherchant un utilisateur spécifique dans Azure AD.

Pour chaque utilisateur à traiter, le module demande le nom du poste (format CW pour Windows, CM pour Mac), recherche l'ordinateur dans Intune, et l'attribue à l'utilisateur comme Primary User via Microsoft Graph. Pour les postes Windows, il déplace également l'objet ordinateur dans l'OU appropriée de l'Active Directory en fonction du site de l'utilisateur, et l'ajoute au groupe de conformité Intune.

Le module génère automatiquement une fiche d'attribution PDF conforme aux exigences juridiques de l'entreprise. Cette fiche inclut les coordonnées de l'entreprise, les caractéristiques techniques du matériel, les accessoires fournis, et les clauses contractuelles relatives à l'utilisation du matériel. Le document est généré localement et automatiquement ouvert dans le navigateur pour impression, puis archivé sur SharePoint dans le dossier approprié.

Le registre SharePoint des demandes est mis à jour pour marquer les équipements comme attribués, permettant ainsi le suivi des demandes en attente.

### 3.4 Module de Gestion du Parc GLPI (mod_asset_glpi.py)

Ce module offre une interface intégrée pour la gestion du parc informatique via l'outil GLPI. Il permet aux administrateurs de visualiser et gérer les équipements attribués aux utilisateurs sans quitter l'application centrale.

La fonctionnalité d'inventaire permet de lister l'ensemble du matériel affecté à un collaborateur, en recherchant dans GLPI tous les équipements (ordinateurs, écrans, docking, casques) liés au compte utilisateur. Les résultats sont affichés dans un tableau structuré avec tri par numéro d'inventaire.

La fonctionnalité d'attribution permet de sélectionner un type d'équipement, de consulter le stock disponible en statut « En Stock », et de choisir l'élément à attribuer. Le module met à jour le statut dans GLPI et associe l'équipement à l'utilisateur.

La fonctionnalité de restitution groupée permet de sélectionner plusieurs équipements à restituer simultanément. Le module met à jour les statuts dans GLPI, génère un bon de restitution PDF unique, et archive le document sur SharePoint dans le dossier de restitution.

### 3.5 Module de Sortie Utilisateur (mod_user_offboard.py)

Ce module automatise le processus complet de départ d'un collaborateur, aussi appelé « leaver process ». Il garantit une désactivation contrôlée et traçable de tous les accès numériques.

Le module commence par une recherche multicritère permettant de trouver l'utilisateur dans l'Active Directory local et dans Azure AD. Cette recherche croisée permet de traiter aussi bien les comptes synchronisés que les comptes cloud-only.

Une fois l'utilisateur identifié, le module récupère les informations sur les équipements associés : ordinateurs gérés dans l'AD, appareils Intune (PC et mobiles), et objets Azure AD. Un récapitulatif est présenté à l'administrateur pour confirmation.

L'étape de traitement Active Directory désactive le compte en modifiant le flag userAccountControl, met à jour la description avec la date et l'administrateur responsable, et déplace l'objet vers l'OU d'archivage dédiée.

L'étape de traitement Entra ID retire l'utilisateur des groupes de licences, ce qui provoque automatiquement la suppression des licences Microsoft 365.

L'étape de purge du matériel supprime les appareils Intune associés à l'utilisateur, supprime les objets correspondants dans Azure AD, et pour les postes Windows, supprime également l'objet ordinateur de l'Active Directory local.

Un rapport récapitulatif est affiché à l'administrateur listant toutes les actions effectuées.

### 3.6 Module de Vérification RH (mod_hr_sync.py)

Ce module réalise une synchronisation bidirectionnelle entre les données RH issues du fichier Excel et l'Active Directory. Il permet de maintenir à jour les informations des collaborateurs et de détecter les anomalies.

Le module lit le fichier Excel RH contenant les informations des salariés et les numéros de téléphone. Pour chaque enregistrement, il recherche l'utilisateur correspondant dans l'Active Directory et compare les attributs.

Les vérifications et mises à jour concernent plusieurs éléments : le manager hiérarchique (avec gestion interactive des homonymes), la fonction (title), la société (company), le département (department), le numéro de téléphone mobile, et la date de fin de contrat (accountExpires).

Le module peut également vérifier les numéros de téléphone dans Intune pour les appareils iOS. Si un numéro Rh ne correspond pas aux appareils enrollés, une alerte est générée dans le rapport.

Le module fonctionne en mode simulation (par défaut) ou production. En mode simulation, il affiche les modifications prévues sans les appliquer. En mode production, il effectue réellement les mises à jour dans l'Active Directory.

Un rapport PDF est généré listant toutes les modifications et alertes, puis automatiquement ouvert dans le navigateur et archivé sur SharePoint.

---

## 4. Configuration et Dépendances

### 4.1 Architecture de configuration centralisée

Userflow utilise une architecture de configuration centralisée. Le fichier principal `settings.json` définit les paramètres essentiels au fonctionnement de la plateforme. Cette approche permet de modifier la configuration sans impacter le code source.

### 4.2 Fichiers de configuration

Les fichiers de configuration JSON sont hébergés sur SharePoint dans `/sites/GLPI/Data/`. Ils sont téléchargés automatiquement par le hub central au démarrage.

| Fichier | Rôle | Impact sur les scripts |
|---------|------|----------------------|
| `settings.json` | Configuration principale (GitHub, SharePoint, AD, modules) | Définit quels modules charger, URLs, credentials Azure |
| `Adresse_Agences.json` | Informations des agences (adresse, téléphone, OU AD, fuseau horaire) | Utilisé par `mod_user_create` pour l'adresse et `mod_license_assign` pour le fuseau horaire Exchange |
| `GD_Agence.json` | Groupes de sécurité par agence | Utilisé par `mod_user_create` pour attribuer les groupes AD selon le lieu de travail |
| `Groupe_service.json` | Mapping services → groupes AD | Utilisé par `mod_user_create` pour attribuer les groupes selon le service |
| `Letsign-it.json` | Modèles de signatures email | Utilisé par `mod_license_assign` pour configurer les signatures |

### 4.3 Fichiers SharePoint (données)

| Fichier | Emplacement SharePoint | Rôle | Scripts concernés |
|---------|----------------------|------|-----------------|
| `python-vault.kdbx` | `/sites/GLPI/Data/` | Coffre KeePass contenant les secrets | Tous (via hub_central) |
| `Formulaire nouvel arrivant.xlsx` | `/personal/cs_automate_entis/Documents/` | Demandes nouveaux arrivants | `mod_user_create` |
| `new_users.xlsx` | `/sites/GLPI/Data/` | Suivi utilisateurs créés | `mod_user_create`, `mod_license_assign`, `mod_device_provision` |

### 4.4 Dépendances Python

L'exécution de Userflow nécessite plusieurs bibliothèques Python tierces. Les bibliothèques critiques installées au démarrage par le mécanisme d'amorçage sont `requests` pour les appels HTTP, `packaging` pour la gestion des versions, et `pykeepass` pour la manipulation des coffres KeePass.

Les bibliothèques métierchargées dynamiquement incluent `office365` pour l'interaction avec SharePoint et OneDrive, `ldap3` pour la communication avec l'Active Directory local, `msgraph` et `azure-identity` pour Microsoft Graph API, `pandas` pour la manipulation des fichiers Excel, `reportlab` pour la génération des documents PDF, `msal` pour l'authentification Microsoft, et `unidecode` pour la normalisation des chaînes de caractères.

---

## 5. Sécurité et Conformité

### 5.1 Modèle de sécurité

La sécurité de Userflow repose sur plusieurs principes fondamentaux. L'authentification multi-facteurs est requise pour l'accès à SharePoint via le mécanisme d'authentification interactive Azure. Le coffre-fort KeePass est déverrouillé manuellement par l'administrateur à chaque session, ce qui évite le stockage de mots de passe en clair.

Les connexions vers l'Active Directory local utilisent le protocole LDAPS (LDAP over SSL) pour garantir la confidentialité des échanges. Les appels vers Microsoft Graph utilisent des tokens d'accès obtained via l'authentification applicative (client credentials).

### 5.2 Traçabilité des opérations

Chaque opération effectuée par Userflow génère des logs détaillés. Les actions sensibles (création de compte, désactivation, suppression) sont accompagnées d'informations de traçabilité incluant la date, l'heure, et l'identité de l'administrateur effectuant l'action. Ces informations sont stockées dans les attributs de l'Active Directory (description, notes) et dans les rapports PDF générés.

### 5.3 Gestion des erreurs

Le code intègre de nombreux mécanismes de gestion des erreurs pour garantir la robustesse du système. Les exceptions sont capturées et affichées de manière claire à l'administrateur. Les opérations critiques demandent confirmation avant exécution. Un mécanisme de rollback est prévu dans la mesure du possible (par exemple, si la création d'un utilisateur échoue en cours de route, les opérations déjà effectuées sont annulées).

---

## 6. Maintenance et Évolution

### 6.1 Gestion des versions

Chaque module suit un système de versioning sémantique. Le numéro de version est géré via les tags Git et non plus dans le nom du fichier. Le hub central détecte automatiquement la dernière version disponible de chaque module sur GitHub en analysant les tags de release. Cette approche permet de déployer des mises à jour sans impacter le hub central ni les noms de fichiers.

Les règles de versioning sont les suivantes : les versions de production utilisent le format `vMAJOR.MINOR` (ex: v1.0, v2.1), les changements majeurs (incompatibilité API) incrémentent le numéro MAJOR, les ajouts de fonctionnalités rétrocompatibles incrémentent le numéro MINOR, et les correctifs de bugs utilisent des hotfixes avec le format `vMAJOR.MINOR.patch`.

### 6.2 Procédure de mise à jour

Pour mettre à jour un module, il suffit de créer une nouvelle version dans le dépôt GitHub avec un tag de release correspondant (ex: v1.2). Lors de la prochaine exécution, le hub détectera automatiquement la nouvelle version et la téléchargera. Il est recommandé de tester les nouvelles versions en environnement de validation avant déploiement en production. Les notes de release doivent documenter les changements effectués.

### 6.3 Ajout de nouveaux modules

L'architecture modulaire permet d'ajouter facilement de nouvelles fonctionnalités. Un nouveau module doit respecter les conventions suivantes : utiliser le préfixe `mod_` suivi d'un nom descriptif en snake_case (ex: `mod_user_create.py`), ne pas inclure la version dans le nom du fichier (le versioning est géré par Git), suivre les indicateurs de qualité selon le stade de développement (ajouter `_draft`, `_beta` ou `_stable`), contenir une fonction asynchrone `run(ctx, kp, tenant_id, ad_conn, config_dir, settings, **kwargs)` comme point d'entrée, être autonome dans ses imports, et suivre les mêmes conventions d'interface utilisateur (couleurs, formatage).

Après développement, le module doit être tagué avec une version Git (ex: v1.0) avant d'être utilisé en production.

---

## 7. Fonctionnalités Futures

Plusieurs évolutions sont envisageables pour étendre les capacités de Userflow. On peut citer l'intégration avec un système de ticketing (GLPI ou autre) pour automatiser la création de tickets lors des opérations, l'ajout de workflows d'approbation pour les opérations sensibles, l'extension à d'autres systèmes de provisionning (Google Workspace, SaaS divers), l'ajout de rapports analytics sur les opérations effectuées, l'intégration avec Azure B2B pour la gestion des utilisateurs externes, et l'automatisation des relances pour les tâches en attente.

---

## 8. Annexes

### 8.1 Glossaire

Les termes techniques utilisés dans ce document sont définis comme suit. L'Active Directory (AD) est le service d'annuaire de Microsoft sur Windows Server. Azure AD (désormais Microsoft Entra ID) est le service d'identité cloud de Microsoft. Microsoft Graph est l'API unifiée pour accéder aux services Microsoft 365. Intune est la solution de gestion des appareils mobiles (MDM) de Microsoft. GLPI est l'outil de gestion de parc informatique open source. SharePoint est la plateforme de collaboration et de stockage de Microsoft 365. KeePass est un gestionnaire de mots de passe open source. Letsign-it est une solution de gestion des signatures email. Leaver Process désigne le processus de départ d'un collaborateur.

### 8.2 Références techniques

Pour toute question technique ou demande d'évolution, contacter le service informatique du Groupe Entis. La documentation technique détaillée se trouve sur le SharePoint du service DSI.

---

*Ce document est vivante et doit être mis à jour à chaque évolution significative du projet.*
