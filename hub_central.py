# -*- coding: utf-8 -*-

"""
Auteur : JEL pour Groupe Entis
Rôle   : Orchestrateur centralisé pour l'administration système.
          Gère l'amorçage de l'environnement, la synchronisation sécurisée via GitHub (token en KDBX)
          et l'exécution des modules métiers avec partage de session KeePass.

Version : 1.9.0
Date    : 14/03/2026

Historique :
-----------
v1.9.0 (14/03/2026) : Création automatique du venv et gestion autonome des dépendances
v1.8.3 (27/02/2026) :  Externalisation de la configuration dans settings.json (SharePoint) et chargement dynamique au démarrage
v1.8.2 (26/12/2025) : Sécurisation du Token GitHub via KeePass et partage de l'objet de connexion aux modules.
v1.8.1 (26/12/2025) : Amélioration de l'explicativité des erreurs et normalisation "Active Directory".
"""

import os
import sys
import subprocess
import tempfile
import shutil
import getpass
import re
import importlib.util
import platform
import time
import json

VENV_DIR = "venv"
REQUIREMENTS_FILE = "requirements.txt"


# ==============================================================================
# INTERFACE UTILISATEUR ET STYLISATION (pour gestion venv)
# ==============================================================================


class Color:
    HEADER = "\033[95m"
    BLUE = "\033[94m"
    CYAN = "\033[96m"
    GREEN = "\033[92m"
    WARNING = "\033[93m"
    YELLOW = "\033[33m"
    FAIL = "\033[91m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    UNDERLINE = "\033[4m"


def print_header(message):
    width = 80
    print(f"\n{Color.HEADER}{Color.BOLD}{'=' * width}{Color.ENDC}")
    print(f"{Color.HEADER}{Color.BOLD} {message.center(width - 2)} {Color.ENDC}")
    print(f"{Color.HEADER}{Color.BOLD}{'=' * width}{Color.ENDC}")


def print_success(message):
    print(f"{Color.GREEN}[OK] {message}{Color.ENDC}")


def print_warning(message):
    print(f"{Color.WARNING}[!] {message}{Color.ENDC}")


def print_error(message):
    print(f"{Color.FAIL}[X] {message}{Color.ENDC}")


def print_info(message):
    print(f"{Color.CYAN}[i] {message}{Color.ENDC}")


def print_browser_alert(message):
    print(f"{Color.YELLOW}[!] {message}{Color.ENDC}")


# ==============================================================================
# GESTION AUTOMATIQUE DU VENV
# ==============================================================================


def is_in_venv():
    """Vérifie si le script s'exécute dans un environnement virtuel."""
    return hasattr(sys, "real_prefix") or (
        hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix
    )


def get_venv_python():
    """Retourne le chemin Python du venv."""
    if platform.system() == "Windows":
        return os.path.join(VENV_DIR, "Scripts", "python.exe")
    return os.path.join(VENV_DIR, "bin", "python")


def get_venv_pip():
    """Retourne le chemin pip du venv."""
    if platform.system() == "Windows":
        return os.path.join(VENV_DIR, "Scripts", "pip.exe")
    return os.path.join(VENV_DIR, "bin", "pip")


def venv_exists():
    """Vérifie si le répertoire venv existe et contient Python."""
    venv_python = get_venv_python()
    return os.path.exists(venv_python)


def create_venv():
    """Crée un environnement virtuel Python."""
    print("\n" + "=" * 60)
    print("CRÉATION DE L'ENVIRONNEMENT VIRTUEL")
    print("=" * 60)
    print(f"Création du venv dans le répertoire : {VENV_DIR}")

    try:
        subprocess.check_call([sys.executable, "-m", "venv", VENV_DIR])
        print_success(f"Environnement virtuel créé avec succès.")
        return True
    except subprocess.CalledProcessError as e:
        print_error(f"Échec de création du venv : {e}")
        return False


def install_requirements_in_venv():
    """Installe les dépendances dans le venv."""
    print("\n" + "=" * 60)
    print("INSTALLATION DES DÉPENDANCES")
    print("=" * 60)

    if not os.path.exists(REQUIREMENTS_FILE):
        print_error(f"Fichier {REQUIREMENTS_FILE} introuvable.")
        return False

    venv_python = get_venv_python()

    print(f"Installation des dépendances depuis {REQUIREMENTS_FILE}...")
    try:
        subprocess.check_call(
            [venv_python, "-m", "pip", "install", "--upgrade", "pip", "-q"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        subprocess.check_call(
            [venv_python, "-m", "pip", "install", "-q", "-r", REQUIREMENTS_FILE]
        )
        print_success("Dépendances installées avec succès.")
        return True
    except subprocess.CalledProcessError as e:
        print_error(f"Échec de l'installation des dépendances : {e}")
        return False


def run_in_venv():
    """Relance le script courant dans le venv."""
    print("\n" + "=" * 60)
    print("RELANCEMENT DANS L'ENVIRONNEMENT VIRTUEL")
    print("=" * 60)

    venv_python = get_venv_python()
    script_path = os.path.abspath(__file__)

    print(f"Relancement avec : {venv_python} {script_path}")

    os.execv(venv_python, [venv_python, script_path] + sys.argv[1:])


def check_dependencies_installed():
    """Vérifie si les dépendances critiques sont installées dans le venv."""
    critical_deps = ["office365", "ldap3", "pandas", "msgraph"]
    venv_python = get_venv_python()
    for dep in critical_deps:
        result = subprocess.run(
            [venv_python, "-c", f"import {dep.split('.')[0]}"],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            return False
    return True


def setup_venv_and_continue():
    """
    Configure le venv et retourne True si on doit continuer l'exécution,
    ou False si le script va être relancé dans le venv.
    """
    if is_in_venv():
        print_info("Exécution dans un environnement virtuel detectée.")
        if not check_dependencies_installed():
            print_warning("Dépendances incomplètes. Installation...")
            if not install_requirements_in_venv():
                return False
        return True

    if venv_exists():
        print_info("Environnement virtuel existant détecté.")
        if not check_dependencies_installed():
            print_warning("Dépendances incomplètes. Installation...")
            if not install_requirements_in_venv():
                return False
        run_in_venv()
        return False

    print_warning("Aucun environnement virtuel détecté.")
    print_info("Création automatique du venv...")

    if not create_venv():
        return False

    if not install_requirements_in_venv():
        return False

    run_in_venv()
    return False


# ==============================================================================
# AMORÇAGE ET DEPENDANCES CRITIQUES
# ==============================================================================


def bootstrap_dependencies():
    critical_deps = ["requests", "packaging", "pykeepass"]
    missing_deps = []

    for dep in critical_deps:
        try:
            importlib.import_module(dep)
        except ImportError:
            missing_deps.append(dep)

    if missing_deps:
        print(
            f"Dépendances d'amorçage manquantes ({', '.join(missing_deps)}). Installation immédiate..."
        )
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "-q"] + missing_deps
            )
            print("Installation réussie. Redémarrage du processus interne.")
        except subprocess.CalledProcessError as e:
            print(f"ERREUR FATALE : Impossible d'amorcer l'environnement : {e}")
            sys.exit(1)


# ==============================================================================
# POINT D'ENTRÉE PRINCIPAL
# ==============================================================================


if __name__ == "__main__":
    if not setup_venv_and_continue():
        sys.exit(0)


bootstrap_dependencies()

import requests
from packaging import version
from pykeepass import PyKeePass

# Chargement différé pour les libs métiers
try:
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.files.file import File
    from ldap3 import Server, Connection, ALL, SIMPLE
except ImportError:
    pass

# ==============================================================================
# INTERFACE UTILISATEUR ET STYLISATION
# ==============================================================================


def show_progress_bar(iteration, total, prefix="", length=40):
    percent = ("{0:.1f}").format(100 * (iteration / float(total)))
    filled_length = int(length * iteration // total)
    bar = "█" * filled_length + "-" * (length - filled_length)
    sys.stdout.write(f"\r{Color.CYAN}{prefix.ljust(25)} |{bar}| {percent}%{Color.ENDC}")
    sys.stdout.flush()
    if iteration == total:
        print()


# ==============================================================================
# SERVICES RÉSEAU ET INTEGRATION GITHUB
# ==============================================================================


def check_ad_connectivity(ad_server):
    param = "-n" if platform.system().lower() == "windows" else "-c"
    command = ["ping", param, "1", ad_server]
    try:
        return (
            subprocess.run(
                command, capture_output=True, text=True, timeout=5
            ).returncode
            == 0
        )
    except:
        return False


def download_file_from_github(token, repo_owner, repo_name, file_path, save_path):
    url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/contents/{file_path}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3.raw",
    }
    try:
        response = requests.get(url, headers=headers, stream=True)
        response.raise_for_status()
        with open(save_path, "wb") as f:
            shutil.copyfileobj(response.raw, f)
        return True
    except:
        return False


def get_latest_script_version(token, repo_owner, repo_name, base_name):
    url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/contents"
    headers = {"Authorization": f"token {token}"}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        files = response.json()
        script_regex = re.compile(rf"{base_name}[-_]v([\d\.]+)\.py")
        versions = []
        for file_info in files:
            if "name" in file_info:
                match = script_regex.match(file_info["name"])
                if match:
                    versions.append((version.parse(match.group(1)), file_info["name"]))
        if versions:
            versions.sort(key=lambda x: x[0], reverse=True)
            return versions[0][1]
        return None
    except:
        return None


def handle_dependencies(token, repo_owner, repo_name, temp_dir):
    print_header("VÉRIFICATION DES COMPOSANTS")
    req_file = "requirements.txt"
    local_req_path = os.path.join(temp_dir, req_file)
    if not download_file_from_github(
        token, repo_owner, repo_name, req_file, local_req_path
    ):
        return False

    print_info("Mise à jour des bibliothèques système...")
    try:
        subprocess.check_call(
            [
                sys.executable,
                "-m",
                "pip",
                "install",
                "-q",
                "--upgrade",
                "-r",
                local_req_path,
            ]
        )
        print_success("Composants système à jour.")
        global ClientContext, File, Server, Connection, ALL, SIMPLE
        from office365.sharepoint.client_context import ClientContext
        from office365.sharepoint.files.file import File
        from ldap3 import Server, Connection, ALL, SIMPLE

        return True
    except:
        return False


# ==============================================================================
# GESTION DES FICHIERS SHAREPOINT ET MODULES DYNAMIQUES
# ==============================================================================


def download_sp_file_to_temp(ctx, sp_file_url, temp_dir):
    try:
        file_content = File.open_binary(ctx, sp_file_url).content
        local_path = os.path.join(temp_dir, os.path.basename(sp_file_url))
        with open(local_path, "wb") as f:
            f.write(file_content)
        return local_path
    except:
        return None


def find_and_import_module(base_name, script_dir):
    """
    Charge un module soit depuis le répertoire local (prioritaire),
    soit depuis le répertoire temporaire (téléchargé depuis GitHub).
    """
    try:
        # Mode local : chercher d'abord dans le répertoire courant
        local_path = os.path.join(os.getcwd(), f"{base_name}.py")
        if os.path.exists(local_path):
            print_info(f"  -> Chargement depuis fichier local : {base_name}.py")
            module_name = f"{base_name}_local"
            spec = importlib.util.spec_from_file_location(module_name, local_path)
            if spec is None or spec.loader is None:
                print_error(
                    f"Impossible de charger le module {base_name} : spec invalide"
                )
                return None
            module = importlib.util.module_from_spec(spec)
            sys.modules[module_name] = module
            spec.loader.exec_module(module)
            return module

        # Mode GitHub (fallback) : chercher dans le répertoire temporaire
        script_pattern = re.compile(rf"{base_name}[-_]v([\d\.]+)\.py")
        files = [f for f in os.listdir(script_dir) if script_pattern.match(f)]
        if not files:
            print_error(
                f"Fichier source introuvable pour le module '{base_name}' (local ou GitHub)."
            )
            return None

        latest_script = max(
            files, key=lambda f: version.parse(script_pattern.match(f).group(1))
        )
        module_name = f"{base_name}_latest"
        full_path = os.path.join(script_dir, latest_script)

        spec = importlib.util.spec_from_file_location(module_name, full_path)
        if spec is None or spec.loader is None:
            print_error(f"Impossible de charger {latest_script} : spec invalide")
            return None
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        return module
    except Exception as e:
        print_error(f"Erreur critique d'importation pour '{base_name}' : {e}")
        import traceback

        traceback.print_exc()
        return None
        return None


# ==============================================================================
# LOGIQUE PRINCIPALE ET MENU D'EXÉCUTION
# ==============================================================================


def main():
    if sys.platform == "win32":
        os.system("cls")

    banner = r"""
 _____ _   _ _____ ___ ____     ____  ____ ___
| ____| \ | |_   _|_ _/ ___|   |  _ \/ ___|_ _|
|  _| |  \| | | |   | |\___ \   | | | \___ \| |
| |___| |\  | | |   | | ___) | | |_| |___) | |
|_____|_| \_| |_|  |___|____/  |____/|____/___|
"""
    print(f"{Color.CYAN}{banner}{Color.ENDC}")
    print_header("HUB D'ADMINISTRATION DYNAMIQUE - ENTIS MUTUELLES")

    SP_GLPI_SITE_URL = "https://entis.sharepoint.com/sites/GLPI"
    SP_APP_CLIENT_ID = "162abf7d-37f3-4e58-87f2-be66877ac142"
    SP_TENANT_ID = "d3f5ec4a-5030-4f94-acf4-aa1997d8865a"

    temp_dir = tempfile.mkdtemp(prefix="entis_hub_")
    ad_conn = None
    kp_session = None
    GITHUB_TOKEN = None
    settings = None

    try:
        # Etape 1 : Authentification SharePoint (Nécessaire pour le KDBX et settings)
        print_header("AUTHENTIFICATION SHAREPOINT SSO")
        print_browser_alert("Utilisation du navigateur pour la connexion MFA.")
        ctx_glpi = None
        while not ctx_glpi:
            try:
                ctx_glpi = ClientContext(SP_GLPI_SITE_URL).with_interactive(
                    SP_TENANT_ID, SP_APP_CLIENT_ID
                )
                ctx_glpi.load(ctx_glpi.web)
                ctx_glpi.execute_query()
                print_success("Session SharePoint GLPI active.")
            except Exception as e:
                print_error(f"Erreur d'authentification SharePoint : {e}")
                if input("Réessayer (o/n) ? ").lower() != "o":
                    sys.exit(1)

        # Chargement de la configuration externe
        print_info("Chargement de la configuration externe (settings.json)...")
        # Essayer d'abord le fichier local
        local_settings_path = os.path.join(os.getcwd(), "settings.json")
        if os.path.exists(local_settings_path):
            print_info("  -> Chargement depuis fichier local : settings.json")
            with open(local_settings_path, "r", encoding="utf-8-sig") as f:
                settings = json.load(f)
            print_success("Configuration locale chargée avec succès.")
        else:
            # Fallback : télécharger depuis SharePoint
            print_info("  → Fichier local introuvé, tentative SharePoint...")
            settings_path = download_sp_file_to_temp(
                ctx_glpi, "/sites/GLPI/Data/settings.json", temp_dir
            )
            if settings_path:
                with open(settings_path, "r", encoding="utf-8-sig") as f:
                    settings = json.load(f)
                print_success("Configuration SharePoint chargée avec succès.")
            else:
                print_error(
                    "Impossible de charger settings.json (local ou SharePoint)."
                )
                sys.exit(1)

        # Extraction des variables de configuration
        REPO_OWNER = settings["github"]["repo_owner"]
        REPO_NAME = settings["github"]["repo_name"]
        SP_ARRIVANTS_SITE_URL = settings["sharepoint"]["arrivants_site_url"]
        AD_SERVER_FQDN = settings["active_directory"]["server_fqdn"]
        AD_DOMAIN = settings["active_directory"]["domain"]
        SCRIPT_BASES = settings["modules"]["script_bases"]
        KDBX_PATH_SP = settings["sharepoint"]["keepass_path"]

        # Construction de la liste des fichiers de config
        SP_CONFIG_FILES = [KDBX_PATH_SP] + [
            f"/sites/GLPI/Data/{f}" for f in settings["referentiels"]
        ]

        # Etape 2 : Récupération du KeePass pour extraire les secrets
        print_info("Récupération du coffre-fort des secrets...")
        local_kdbx = download_sp_file_to_temp(ctx_glpi, KDBX_PATH_SP, temp_dir)

        if not local_kdbx:
            print_error("Impossible de télécharger le fichier KeePass.")
            sys.exit(1)

        print_header("DÉVERROUILLAGE DU COFFRE-FORT KEEPASS")
        while not kp_session:
            try:
                pwd_kp = getpass.getpass(
                    f"{Color.BLUE}Mot de passe Maître KeePass : {Color.ENDC}"
                )
                kp_session = PyKeePass(local_kdbx, password=pwd_kp)
                print_success("Coffre-fort déverrouillé.")
            except Exception:
                print_error("Mot de passe KeePass incorrect.")
                if input("Réessayer (o/n) ? ").lower() != "o":
                    sys.exit(1)

        # Extraction du Token GitHub
        entry_github = kp_session.find_entries(title="GitHub", first=True)
        if not entry_github:
            print_error("Entrée 'GitHub' introuvable dans le KeePass !")
            sys.exit(1)
        GITHUB_TOKEN = entry_github.password

        # Etape 3 : Amorçage de l'environnement via GitHub
        if not handle_dependencies(GITHUB_TOKEN, REPO_OWNER, REPO_NAME, temp_dir):
            sys.exit(1)

        # Etape 4 : Synchronisation GitHub des modules métiers
        print_header("SYNCHRONISATION DES OUTILS MÉTIERS")
        for i, base in enumerate(SCRIPT_BASES):
            show_progress_bar(i, len(SCRIPT_BASES), prefix="Scripts GitHub")
            latest = get_latest_script_version(
                GITHUB_TOKEN, REPO_OWNER, REPO_NAME, base
            )
            if latest:
                download_file_from_github(
                    GITHUB_TOKEN,
                    REPO_OWNER,
                    REPO_NAME,
                    latest,
                    os.path.join(temp_dir, latest),
                )
        show_progress_bar(len(SCRIPT_BASES), len(SCRIPT_BASES), prefix="Scripts GitHub")

        # Etape 5 : Configurations locales (autres fichiers JSON)
        print_info("Récupération des référentiels JSON...")
        for i, f_url in enumerate(SP_CONFIG_FILES):
            if f_url == KDBX_PATH_SP:
                continue  # Déjà fait
            show_progress_bar(i, len(SP_CONFIG_FILES), prefix="Paramétrage SharePoint")
            download_sp_file_to_temp(ctx_glpi, f_url, temp_dir)
        show_progress_bar(
            len(SP_CONFIG_FILES), len(SP_CONFIG_FILES), prefix="Paramétrage SharePoint"
        )
        print_success("Référentiels chargés.")

        # Etape 6 : Authentification Active Directory
        print_header("CONNEXION ACTIVE DIRECTORY (COMPTE ADMINISTRATEUR)")
        if not check_ad_connectivity(AD_SERVER_FQDN):
            print_error(f"Serveur Active Directory ({AD_SERVER_FQDN}) injoignable.")
            if input("Tenter quand même l'authentification (o/n) ? ").lower() != "o":
                sys.exit(1)
        else:
            print_success(f"Serveur Active Directory joignable.")

        server = Server(AD_SERVER_FQDN, get_info=ALL, use_ssl=True)
        while not ad_conn:
            try:
                print_info(
                    "Veuillez saisir vos identifiants administrateur Active Directory."
                )
                user = input(f"{Color.BLUE}Identifiant (admin.xxx) : {Color.ENDC}")
                pwd = getpass.getpass(f"{Color.BLUE}Mot de passe : {Color.ENDC}")
                user_dn = f"{user}@{AD_DOMAIN}" if "@" not in user else user
                ad_conn = Connection(
                    server,
                    user=user_dn,
                    password=pwd,
                    authentication=SIMPLE,
                    auto_bind=True,
                )
                print_success("Connexion LDAPS Active Directory établie avec succès.")
            except Exception as e:
                print_error(f"Échec de liaison Active Directory : {e}")
                if input("Réessayer (o/n) ? ").lower() != "o":
                    sys.exit(1)

        # Etape 7 : Préparation des modules
        print_header("CHARGEMENT DES LOGIQUES MÉTIERS")
        modules = {}
        loading_failed = False
        for base in SCRIPT_BASES:
            print_info(f"Chargement du module : {base}...")
            mod = find_and_import_module(base, temp_dir)
            if mod:
                modules[base] = mod
                print_success(f"Module '{base}' prêt.")
            else:
                loading_failed = True

        if loading_failed:
            print_error("Un ou plusieurs modules n'ont pas pu être chargés.")
            if input("Continuer avec les modules chargés (o/n) ? ").lower() != "o":
                sys.exit(1)

        # Etape 8 : Menu principal et orchestration
        operations = {
            "1": ("Créer un ou plusieurs utilisateurs", modules.get("mod_user_create")),
            "2": (
                "Attribution de licences & gestion des signatures",
                modules.get("mod_license_assign"),
            ),
            "3": (
                "Remise de matériel (nouvels arrivants)",
                modules.get("mod_device_provision"),
            ),
            "4": (
                "Sortir un ou plusieurs utilisateur",
                modules.get("mod_user_offboard"),
            ),
            "5": (
                "Gestion du parc GLPI (Inventaire, Prêt, Restitution)",
                modules.get("mod_asset_glpi"),
            ),
            "6": (
                "Vérification et synchronisation RH",
                modules.get("mod_hr_sync"),
            ),
            "X": ("Quitter", None),
        }

        while True:
            print_header("MENU PRINCIPAL")
            for key in sorted(operations.keys()):
                label, mod = operations[key]
                color = Color.CYAN if mod or key == "X" else Color.FAIL
                suffix = "" if mod or key == "X" else " (NON CHARGÉ)"
                print(f"   {color}{key}{Color.ENDC}. {label}{suffix}")

            try:
                choice = input(
                    f"\n{Color.BLUE}Saisissez votre choix : {Color.ENDC}"
                ).upper()
            except (EOFError, KeyboardInterrupt):
                print_warning("\nSaisie interrompue.")
                break

            if not choice:
                print_warning("Veuillez faire un choix.")
                continue

            if choice in operations:
                desc, mod = operations[choice]
                if choice == "X":
                    print_info("Au revoir!")
                    break
                if not mod:
                    print_error(f"Le module '{desc}' n'est pas disponible.")
                    continue

                func = mod.run
                print(f"\nExécution : {desc}\n{'-' * 80}")
                try:
                    import asyncio

                    if choice == "1":
                        print_browser_alert("Ouverture des demandes SharePoint...")
                        ctx_arr = ClientContext(SP_ARRIVANTS_SITE_URL).with_interactive(
                            SP_TENANT_ID, SP_APP_CLIENT_ID
                        )
                        func(
                            ctx_glpi=ctx_glpi,
                            ctx_arrivants=ctx_arr,
                            conn=ad_conn,
                            config_dir=temp_dir,
                            kp=kp_session,
                            settings=settings,
                        )

                    elif choice in ["2", "3", "4", "5"]:
                        asyncio.run(
                            func(
                                ctx=ctx_glpi,
                                kp=kp_session,
                                tenant_id=SP_TENANT_ID,
                                ad_conn=ad_conn,
                                config_dir=temp_dir,
                                settings=settings,
                            )
                        )

                except Exception as e:
                    print_error(f"Erreur durant l'opération : {e}")

                if (
                    input(f"\n{Color.BLUE}Autre opération (o/n) ? {Color.ENDC}").lower()
                    != "o"
                ):
                    break
            else:
                print_warning("Choix invalide.")

    except (KeyboardInterrupt, SystemExit):
        print_warning("\nInterruption de session.")
    finally:
        if ad_conn:
            ad_conn.unbind()
            print_info("Session Active Directory fermée.")
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
        print_info("Nettoyage et fermeture du HUB.")


if __name__ == "__main__":
    main()
