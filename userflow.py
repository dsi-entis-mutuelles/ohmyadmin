# -*- coding: utf-8 -*-

"""
USERFLOW - Script unique d'administration système
=================================================

Ce script autonome gère l'ensemble du workflow d'administration :
- Création/gestion des utilisateurs
- Attribution de licences
- Remise de matériel
- Gestion du parc GLPI
- Sortie utilisateur
- Sync RH

USAGE :
    python userflow.py

Le script gère automatiquement :
    - L'environnement virtuel (venv)
    - Les dépendances Python
    - La configuration (JSON)
    - Les secrets (KeePass)

Auteur : JEL pour Groupe Entis
Version : 2.0.0
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
import json

VENV_DIR = "venv"
CONFIG_DIR = "config"
REQUIREMENTS_FILE = "requirements.txt"


# ==============================================================================
# COULEURS POUR L'INTERFACE
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


def print_success(msg):
    print(f"{Color.GREEN}[OK] {msg}{Color.ENDC}")


def print_warning(msg):
    print(f"{Color.YELLOW}[!] {msg}{Color.ENDC}")


def print_error(msg):
    print(f"{Color.FAIL}[X] {msg}{Color.ENDC}")


def print_info(msg):
    print(f"{Color.CYAN}[i] {msg}{Color.ENDC}")


def print_browser_alert(msg):
    print(f"{Color.BLUE}{msg}{Color.ENDC}")


def print_header(msg):
    print(f"\n{Color.BOLD}{Color.HEADER}{'=' * 60}\n{msg}\n{'=' * 60}{Color.ENDC}\n")


# ==============================================================================
# GESTION AUTOMATIQUE DU VENV
# ==============================================================================


def is_in_venv():
    return hasattr(sys, "real_prefix") or (
        hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix
    )


def get_venv_python():
    if platform.system() == "Windows":
        return os.path.join(VENV_DIR, "Scripts", "python.exe")
    return os.path.join(VENV_DIR, "bin", "python")


def venv_exists():
    return os.path.exists(get_venv_python())


def create_venv():
    print_info(f"Creation du venv dans : {VENV_DIR}")
    subprocess.check_call([sys.executable, "-m", "venv", VENV_DIR])
    print_success("Environnement virtuel cree.")


def install_requirements():
    venv_python = get_venv_python()
    print_info("Installation des dependances...")
    subprocess.check_call(
        [venv_python, "-m", "pip", "install", "--upgrade", "pip", "-q"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    subprocess.check_call(
        [venv_python, "-m", "pip", "install", "-q", "-r", REQUIREMENTS_FILE],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    print_success("Dependances installees.")


def check_dependencies():
    """Verifie si les dependances sont presentees."""
    venv_python = get_venv_python()
    deps = ["office365", "ldap3", "pandas"]
    for dep in deps:
        result = subprocess.run(
            [venv_python, "-c", f"import {dep.split('.')[0]}"], capture_output=True
        )
        if result.returncode != 0:
            return False
    return True


def setup_environment():
    """Configure l'environnement automatiquement."""
    if is_in_venv():
        print_info("Execution dans venv detectee.")
        if not check_dependencies():
            print_warning("Dependances incompletes. Installation...")
            install_requirements()
        return True

    if venv_exists():
        print_info("Venv existant detecte.")
        if not check_dependencies():
            install_requirements()
        print_info("Relancement dans le venv...")
        venv_python = get_venv_python()
        os.execv(venv_python, [venv_python, __file__] + sys.argv[1:])

    print_warning("Aucun venv detecte. Creation automatique...")
    create_venv()
    install_requirements()
    print_info("Relancement dans le venv...")
    venv_python = get_venv_python()
    os.execv(venv_python, [venv_python, __file__] + sys.argv[1:])


# ==============================================================================
# CHARGEMENT DE LA CONFIGURATION
# ==============================================================================


def find_config_file(filename):
    """Cherche un fichier de config dans plusieurs emplacements."""
    search_paths = [
        os.getcwd(),
        os.path.join(os.getcwd(), CONFIG_DIR),
        os.path.dirname(__file__),
        os.path.join(os.path.dirname(__file__), CONFIG_DIR),
    ]
    for path in search_paths:
        filepath = os.path.join(path, filename)
        if os.path.exists(filepath):
            return filepath
    return None


def load_settings():
    """Charge le fichier settings.json."""
    print_info("Chargement de la configuration...")

    settings_path = find_config_file("settings.json")
    if not settings_path:
        print_error("settings.json introuvable!")
        print_info("Creez un dossier 'config/' avec settings.json")
        sys.exit(1)

    with open(settings_path, "r", encoding="utf-8-sig") as f:
        settings = json.load(f)

    print_success(f"Configuration chargee depuis : {settings_path}")
    return settings


def load_referentiels(settings):
    """Charge les fichiers de referentiels JSON."""
    print_info("Chargement des referentiels...")
    referentiels = {}
    for ref_file in settings.get("referentiels", []):
        path = find_config_file(ref_file)
        if path:
            with open(path, "r", encoding="utf-8-sig") as f:
                referentiels[ref_file] = json.load(f)
            print_info(f"  - {ref_file}")
        else:
            print_warning(f"  - {ref_file} (non trouve)")
    return referentiels


# ==============================================================================
# AUTHENTIFICATION ET KEEPASS
# ==============================================================================


def download_keepass(ctx, keepass_path, temp_dir):
    """Telecharge le fichier KeePass depuis SharePoint."""
    from office365.sharepoint.files.file import File

    local_path = os.path.join(temp_dir, "vault.kdbx")
    with open(local_path, "wb") as f:
        File.download(ctx, keepass_path).download(f).execute_query()
    return local_path


def unlock_keepass(ctx, settings, temp_dir):
    """Demande le mot de passe KeePass et deverrouille le coffre."""
    from pykeepass import PyKeePass

    print_header("DVERROUILLAGE DU COFFRE-FORT")

    keepass_path = settings["sharepoint"]["keepass_path"]
    local_kdbx = download_keepass(ctx, keepass_path, temp_dir)

    while True:
        pwd = getpass.getpass(f"{Color.BLUE}Mot de passe KeePass : {Color.ENDC}")
        try:
            kp = PyKeePass(local_kdbx, password=pwd)
            print_success("Coffre deverrouille.")
            return kp
        except Exception:
            print_error("Mot de passe incorrect. Reessayer ? (o/n)")
            if input("> ").lower() != "o":
                sys.exit(1)


def get_secrets(kp):
    """Extrait les secrets depuis KeePass."""
    secrets = {}

    # GitHub
    entry = kp.find_entries(title="GitHub", first=True)
    secrets["github_token"] = entry.password if entry else None

    # Azure
    entry = kp.find_entries(title="Azure", first=True) or kp.find_entries(
        title="Azure App Credentials", first=True
    )
    if entry:
        secrets["azure_client_id"] = entry.username
        secrets["azure_client_secret"] = entry.password

    return secrets


# ==============================================================================
# SHAREPOINT & GRAPH API CONNECTION (SILENT)
# ==============================================================================


def get_azure_token(settings, secrets):
    """Obtient un token Azure AD via Client Credentials Flow."""
    import requests

    tenant_id = settings["sharepoint"]["tenant_id"]
    client_id = secrets.get("azure_client_id")
    client_secret = secrets.get("azure_client_secret")

    if not client_id or not client_secret:
        print_error("Azure App credentials manquants dans KeePass!")
        print_info(
            "Ajoutez 'Azure App' avec Username=ClientId et Password=ClientSecret"
        )
        sys.exit(1)

    # Scope pour SharePoint Online
    sp_url = settings["sharepoint"]["glpi_site_url"]
    tenant_domain = sp_url.split("/sites/")[0].replace("https://", "")
    scope = f"https://{tenant_domain}/.default"

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": scope,
        "grant_type": "client_credentials",
    }

    response = requests.post(token_url, data=data)
    if response.status_code != 200:
        print_error(f"Erreur Azure auth: {response.status_code}")
        print_error(response.text)
        sys.exit(1)

    return response.json()["access_token"]


def connect_sharepoint(settings, secrets):
    """Etablit la connexion SharePoint via Azure App (silent)."""
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.azure_token_provider import AzureTokenProvider

    print_header("CONNEXION SHAREPOINT (SILENT)")

    print_info("Obtention du token Azure...")
    access_token = get_azure_token(settings, secrets)

    print_info("Connexion a SharePoint...")

    # Creer le contexte SharePoint
    ctx = ClientContext(settings["sharepoint"]["glpi_site_url"])

    # Utiliser le token provider
    token_provider = AzureTokenProvider(access_token)
    ctx._auth = token_provider

    ctx.load(ctx.web)
    ctx.execute_query()

    print_success("Session SharePoint active (silent).")
    return ctx


# ==============================================================================
# CHARGEMENT DES MODULES
# ==============================================================================


def load_modules(settings, temp_dir):
    """Charge les modules depuis GitHub."""
    print_header("SYNCHRONISATION DES MODULES")

    import requests
    from packaging import version

    github_token = settings.get("_secrets", {}).get("github_token")
    if not github_token:
        print_error("Token GitHub manquant dans KeePass!")
        sys.exit(1)

    repo_owner = settings["github"]["repo_owner"]
    repo_name = settings["github"]["repo_name"]
    modules = {}

    headers = {"Authorization": f"token {github_token}"}

    for base in settings["modules"]["script_bases"]:
        print_info(f"Chargement de {base}...")

        # Chercher fichier local
        local_path = os.path.join(os.getcwd(), f"{base}.py")
        if not local_path:
            local_path = os.path.join(os.path.dirname(__file__), f"{base}.py")

        if os.path.exists(local_path):
            print_info(f"  -> depuis fichier local")
            spec = importlib.util.spec_from_file_location(base, local_path)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            modules[base] = mod
            continue

        # Telecharger depuis GitHub
        url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/contents"
        resp = requests.get(url, headers=headers)
        files = resp.json()

        pattern = re.compile(rf"{base}[-_]v([\d\.]+)\.py")
        versions = []
        for f in files:
            match = pattern.match(f.get("name", ""))
            if match:
                versions.append((version.parse(match.group(1)), f["name"]))

        if versions:
            versions.sort(key=lambda x: x[0], reverse=True)
            latest = versions[0][1]

            download_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/contents/{latest}"
            resp = requests.get(download_url, headers=headers)
            content = resp.json()
            import base64

            file_content = base64.b64decode(content["content"])

            save_path = os.path.join(temp_dir, latest)
            with open(save_path, "wb") as f:
                f.write(file_content)

            spec = importlib.util.spec_from_file_location(base, save_path)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            modules[base] = mod
            print_success(f"  -> {latest}")
        else:
            print_warning(f"  -> Non trouve sur GitHub")

    return modules


# ==============================================================================
# ACTIVE DIRECTORY
# ==============================================================================


def connect_ad(settings):
    """Etablit la connexion Active Directory."""
    from ldap3 import Server, Connection, ALL, SIMPLE

    print_header("CONNEXION ACTIVE DIRECTORY")
    print_info("Identifiants administrateur requis.")

    ad_config = settings["active_directory"]
    server = Server(ad_config["server_fqdn"], get_info=ALL, use_ssl=True)

    while True:
        user = input(f"{Color.BLUE}Utilisateur (admin.xxx) : {Color.ENDC}")
        pwd = getpass.getpass(f"{Color.BLUE}Mot de passe : {Color.ENDC}")

        user_dn = f"{user}@{ad_config['domain']}"
        try:
            conn = Connection(
                server,
                user=user_dn,
                password=pwd,
                authentication=SIMPLE,
                auto_bind=True,
            )
            print_success("Connecte a l'Active Directory.")
            return conn
        except Exception as e:
            print_error(f"Echec : {e}")
            if input("Reessayer ? (o/n) ").lower() != "o":
                sys.exit(1)


# ==============================================================================
# MENU PRINCIPAL
# ==============================================================================


def show_menu(modules, settings):
    """Affiche le menu principal."""
    print_header("MENU PRINCIPAL - USERFLOW")

    operations = {
        "1": ("Creer un utilisateur", modules.get("mod_user_create")),
        "2": ("Attribuer licences", modules.get("mod_license_assign")),
        "3": ("Remise materiel", modules.get("mod_device_provision")),
        "4": ("Sortie utilisateur", modules.get("mod_user_offboard")),
        "5": ("Gestion parc GLPI", modules.get("mod_asset_glpi")),
        "6": ("Sync RH", modules.get("mod_hr_sync")),
        "X": ("Quitter", None),
    }

    for key in sorted(operations.keys()):
        label, mod = operations[key]
        color = Color.CYAN if mod else Color.FAIL
        suffix = "" if mod else " (NON DISPONIBLE)"
        print(f"   {color}{key}{Color.ENDC}. {label}{suffix}")

    return operations


def run_module(choice, modules, ctx, conn, kp, settings, referentiels):
    """Execute le module choisi."""
    operations = {
        "1": ("Creer un utilisateur", modules.get("mod_user_create"), True),
        "2": ("Attribuer licences", modules.get("mod_license_assign"), False),
        "3": ("Remise materiel", modules.get("mod_device_provision"), False),
        "4": ("Sortie utilisateur", modules.get("mod_user_offboard"), False),
        "5": ("Gestion parc GLPI", modules.get("mod_asset_glpi"), False),
        "6": ("Sync RH", modules.get("mod_hr_sync"), False),
    }

    if choice not in operations:
        print_error("Choix invalide.")
        return

    label, mod, need_arrivants = operations[choice]
    if not mod:
        print_error(f"Module {label} non charge.")
        return

    print(f"\n{Color.BOLD}=== {label} ==={Color.ENDC}\n")

    try:
        import asyncio

        if choice == "1":
            from office365.sharepoint.client_context import ClientContext

            sp = settings["sharepoint"]
            ctx_arr = ClientContext(sp["arrivants_site_url"]).with_interactive(
                sp["tenant_id"], sp["app_client_id"]
            )

            asyncio.run(
                mod.run(
                    ctx_glpi=ctx,
                    ctx_arrivants=ctx_arr,
                    conn=conn,
                    config_dir=os.path.join(os.getcwd(), CONFIG_DIR),
                    kp=kp,
                    settings=settings,
                )
            )
        else:
            asyncio.run(
                mod.run(
                    ctx=ctx,
                    kp=kp,
                    tenant_id=settings["sharepoint"]["tenant_id"],
                    ad_conn=conn,
                    config_dir=os.path.join(os.getcwd(), CONFIG_DIR),
                    settings=settings,
                )
            )

        print_success("Operation terminee.")
    except Exception as e:
        print_error(f"Erreur : {e}")
        import traceback

        traceback.print_exc()


# ==============================================================================
# POINT D'ENTREE PRINCIPAL
# ==============================================================================


def main():
    print(f"""
{Color.CYAN}
  _   _ _____ ___ ____     ____  ____ ___
 | \\ | |_   _|_ _/ ___|   |  _ \\/ ___|_ _|
 |  \\| | | |   | |\\___ \\   | | | \\___ \\| |
 | |\\  | | |   | | ___) |  | |_| |___) | |
 |_| \\_| |_|  |___|____/  |____/|____/___|
 
       USERFLOW - Administration Systeme
       Version 2.0 - Script Unique
{Color.ENDC}
""")

    # 1. Setup automatique de l'environnement
    setup_environment()

    # 2. Import des dependances
    import requests
    from packaging import version
    from pykeepass import PyKeePass
    from office365.sharepoint.client_context import ClientContext
    from ldap3 import Server, Connection, ALL, SIMPLE

    # 3. Chargement configuration
    settings = load_settings()
    referentiels = load_referentiels(settings)

    # 4. Deverrouillage KeePass (AVANT SharePoint pour avoir les secrets Azure)
    print_info("Deverrouillage du coffre-fort...")
    temp_dir = tempfile.mkdtemp(prefix="userflow_")

    # Telechargement KeePass via auth interactive
    from office365.sharepoint.client_context import ClientContext

    sp_config = settings["sharepoint"]

    print_info("Telechargement du coffre-fort KeePass...")
    ctx_temp = ClientContext(sp_config["glpi_site_url"]).with_interactive(
        sp_config["tenant_id"], sp_config["app_client_id"]
    )

    try:
        kp = unlock_keepass(ctx_temp, settings, temp_dir)
        secrets = get_secrets(kp)
        settings["_secrets"] = secrets
        print_success("Coffre-fort charge.")
    except Exception as e:
        print_error(f"ErreurKeePass: {e}")
        sys.exit(1)

    # 5. Connexion SharePoint via Azure App (SILENT)
    ctx = connect_sharepoint(settings, secrets)

    # 6. Chargement des modules
    modules = load_modules(settings, temp_dir)

    # 7. Connexion Active Directory
    conn = connect_ad(settings)

    # 8. Boucle principale
    while True:
        operations = show_menu(modules, settings)

        try:
            choice = input(f"\n{Color.BLUE}Choix : {Color.ENDC}").upper()
        except (EOFError, KeyboardInterrupt):
            print("\nAu revoir!")
            break

        if choice == "X":
            print("Aurevoir!")
            break

        run_module(choice, modules, ctx, conn, kp, settings, referentiels)

        input(f"\n{Color.BLUE}Entree pour continuer...{Color.ENDC}")

    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir, ignore_errors=True)
    print_info("Nettoyage termine.")


if __name__ == "__main__":
    main()
