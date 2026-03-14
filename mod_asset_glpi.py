# -*- coding: utf-8 -*-

"""
Auteur : JEL pour Groupe Entis
Rôle   : Module de gestion du parc GLPI. Permet l'inventaire, l'attribution
         et la restitution groupée de matériels avec génération d'un PDF unique.

Version : 1.8.0
Date    : 26/12/2024

Historique :
-----------
v1.8.0 (26/12/2024) : Harmonisation totale (List/Assign/Return) avec tableau strict (bordures),
                     tri par inventaire (ASC) et code couleur Orange. Support de la touche 'R'.
v1.7.0 (26/12/2024) : Standardisation UX et ajout touche 'R' pour retour menu.
v2.5.0 (26/12/2024) : Implémentation de la restitution groupée.
"""

import os
import sys
import json
import requests
import tempfile
import webbrowser
import asyncio
import re
from datetime import datetime
from requests.packages.urllib3.exceptions import InsecureRequestWarning

# Dépendances Microsoft & LDAP
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from ldap3 import Server, Connection, ALL, SIMPLE

# Dépendances PDF (ReportLab)
try:
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader
except ImportError:
    import subprocess

    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "reportlab"])
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader

# Désactivation des avertissements SSL
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

# ==============================================================================
# CONFIGURATION ET CONSTANTES
# ==============================================================================

GLPI_URL = "https://glpi.mutuelles-entis.fr/apirest.php"
APP_TOKEN = "R3IgdxlMhADjM5rfuFjlhEvRXyPmGRWRLu3czzAT"
USER_TOKEN = "yAomUk3W4D6kK5xaXx1gXitpXmMLf1dcd7jRvzyk"

SP_PDF_RESTITUTION_URL = "/sites/GLPI/Documents partages/Ordinateurs/Restitution"
SP_LOGO_URL = "/sites/GLPI/Data/groupe_entis.png"
AD_SEARCH_BASE = "dc=cetremut,dc=pri"

# Mapping IDs GLPI
CONTACT_FIELD_ID = "7"  # SamAccountName (Usager texte)
STATUS_FIELD_ID = "31"  # Statut
INVENTORY_FIELD_ID = "6"  # Numéro d'inventaire
MODEL_FIELD_ID = "40"  # Modèle
NAME_FIELD_ID = "1"  # Nom
SERIAL_FIELD_ID = "5"  # S/N

# Paramètres Métiers
STATUS_VALUE_STOCK = "En Stock"
STATUS_ID_SERVICE = 1
STATUS_ID_STOCK = 2
FOOTER_TEXT = (
    "Groupe Entis | 39, rue du Jourdil 74960 Cran-Gevrier | Tel : 09 69 39 96 96"
)

ITEM_TYPES = {
    "1": {"type": "Computer", "display": "Ordinateur", "subtype": None},
    "2": {"type": "Monitor", "display": "Écran", "subtype": None},
    "3": {"type": "Peripheral", "display": "Docking", "subtype": "Docking"},
    "4": {
        "type": "Peripheral",
        "display": "Casque sans fil",
        "subtype": "Casque audio",
    },
}

# ==============================================================================
# INTERFACE UTILISATEUR ET STYLISATION
# ==============================================================================


class Color:
    HEADER = "\033[95m"
    BLUE = "\033[94m"
    CYAN = "\033[96m"
    GREEN = "\033[92m"
    WARNING = "\033[93m"
    YELLOW = "\033[33m"
    ORANGE = "\033[38;5;208m"
    FAIL = "\033[91m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"


def print_header(message):
    width = 90
    print(f"\n{Color.HEADER}{Color.BOLD}{'=' * width}{Color.ENDC}")
    print(f"{Color.HEADER}{Color.BOLD} {message.center(width - 2)} {Color.ENDC}")
    print(f"{Color.HEADER}{Color.BOLD}{'=' * width}{Color.ENDC}")


def print_success(message):
    print(f"{Color.GREEN}[✓] {message}{Color.ENDC}")


def print_warning(message):
    print(f"{Color.WARNING}[!] {message}{Color.ENDC}")


def print_error(message):
    print(f"{Color.FAIL}[X] {message}{Color.ENDC}")


def print_info(message):
    print(f"{Color.CYAN}[i] {message}{Color.ENDC}")


# ==============================================================================
# LOGIQUE DE TRI ET D'AFFICHAGE (TABLEAU STRICT)
# ==============================================================================


def get_sorted_assets(assets_list):
    """Sépare et trie les matériels : Inventaire ASC d'abord, puis le reste."""
    with_inv = []
    without_inv = []

    for a in assets_list:
        inv_val = str(a.get(INVENTORY_FIELD_ID, "")).strip()
        if inv_val and inv_val.lower() not in ["none", "", "null", "nan"]:
            with_inv.append(a)
        else:
            without_inv.append(a)

    # Tri alphanumérique intelligent
    def sort_key(x):
        val = str(x.get(INVENTORY_FIELD_ID, ""))
        match = re.search(r"\d+", val)
        if match:
            return (0, int(match.group()))
        return (1, val.lower())

    with_inv.sort(key=sort_key)
    return with_inv, without_inv


def display_assets_table(assets_list, is_selection=False):
    """Affiche un tableau avec une structure de grille stricte pour éviter les décalages."""
    with_inv, without_inv = get_sorted_assets(assets_list)

    # Définition des largeurs de colonnes (Strictes)
    W_ID = 4
    W_TYPE = 12
    W_MODEL = 40
    W_INV = 15

    # Lignes de séparation
    sep_inner = f"+{'-' * (W_ID + 2)}+" if is_selection else "+"
    sep_inner += f"{'-' * (W_TYPE + 2)}+{'-' * (W_MODEL + 2)}+{'-' * (W_INV + 2)}+"

    # En-têtes
    head_inner = f"| {'ID':<{W_ID}} |" if is_selection else "|"
    head_inner += (
        f" {'Type':<{W_TYPE}} | {'Modèle / Nom':<{W_MODEL}} | {'Inventaire':<{W_INV}} |"
    )

    print(f"\n{sep_inner}")
    print(f"{Color.BOLD}{head_inner}{Color.ENDC}")
    print(sep_inner)

    final_ordered_list = with_inv + without_inv

    # 1. Matériel AVEC Inventaire
    for i, a in enumerate(with_inv):
        id_val = f"{i + 1:<{W_ID}}"
        t_val = f"{a.get('_display_type', a.get('_itype', 'N/A'))[:W_TYPE]:<{W_TYPE}}"
        m_val = f"{str(a.get(MODEL_FIELD_ID) or a.get(NAME_FIELD_ID, 'N/A'))[:W_MODEL]:<{W_MODEL}}"
        i_val = f"{str(a.get(INVENTORY_FIELD_ID))[:W_INV]:<{W_INV}}"

        row = f"| {id_val} |" if is_selection else "|"
        row += f" {t_val} | {m_val} | {i_val} |"
        print(row)

    # 2. Séparateur si mixité
    if with_inv and without_inv:
        print(sep_inner)

    # 3. Matériel SANS Inventaire (Orange)
    offset = len(with_inv)
    for i, a in enumerate(without_inv):
        id_val = f"{i + 1 + offset:<{W_ID}}"
        t_val = f"{a.get('_display_type', a.get('_itype', 'N/A'))[:W_TYPE]:<{W_TYPE}}"
        m_val = f"{str(a.get(MODEL_FIELD_ID) or a.get(NAME_FIELD_ID, 'N/A'))[:W_MODEL]:<{W_MODEL}}"
        i_val = f"{'SANS INV':<{W_INV}}"

        row_content = f"| {id_val} |" if is_selection else "|"
        row_content += f" {t_val} | {m_val} | {i_val} |"
        # On applique la couleur tout en gardant les pipes alignés
        print(f"{Color.ORANGE}{row_content}{Color.ENDC}")

    print(sep_inner)
    return final_ordered_list


# ==============================================================================
# LOGIQUE MÉTIER GLPI
# ==============================================================================


class GlpiModule:
    def __init__(self):
        self.api_url = GLPI_URL
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"user_token {USER_TOKEN}",
            "App-Token": APP_TOKEN,
        }
        self.session_token = None

    def connect(self):
        try:
            res = requests.get(
                f"{self.api_url}/initSession",
                headers=self.headers,
                verify=False,
                timeout=10,
            )
            self.session_token = res.json()["session_token"]
            self.headers["Session-Token"] = self.session_token
            return True
        except:
            return False

    def search(self, item_type, criteria):
        params = {"range": "0-100"}
        params.update(
            {
                f"forcedisplay[{i}]": fid
                for i, fid in enumerate(
                    [
                        NAME_FIELD_ID,
                        "2",
                        MODEL_FIELD_ID,
                        INVENTORY_FIELD_ID,
                        CONTACT_FIELD_ID,
                        STATUS_FIELD_ID,
                        SERIAL_FIELD_ID,
                    ]
                )
            }
        )
        params.update(
            {
                f"criteria[{i}][{k}]": v
                for i, c in enumerate(criteria)
                for k, v in c.items()
            }
        )
        try:
            res = requests.get(
                f"{self.api_url}/search/{item_type}",
                headers=self.headers,
                params=params,
                verify=False,
            )
            return res.json().get("data", [])
        except:
            return []

    def update(self, item_type, item_id, data):
        url = f"{self.api_url}/{item_type}/{item_id}"
        try:
            res = requests.put(
                url,
                headers=self.headers,
                data=json.dumps({"input": data}),
                verify=False,
            )
            return res.status_code in [200, 201]
        except:
            return False


# ==============================================================================
# LOGIQUE PDF CONSOLIDÉ
# ==============================================================================


def create_bulk_restitution_pdf(user_name, assets_list, local_logo_path=None):
    temp_dir = tempfile.gettempdir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    pdf_filename = f"Restitution_GROUPEE_{user_name.replace(' ', '_')}_{timestamp}.pdf"
    file_path = os.path.join(temp_dir, pdf_filename)

    def add_header_footer(canvas, doc):
        canvas.saveState()
        if local_logo_path and os.path.exists(local_logo_path):
            logo = ImageReader(local_logo_path)
            canvas.drawImage(
                logo,
                doc.width + doc.leftMargin - 2.5 * cm,
                doc.height + doc.topMargin - 1.5 * cm,
                width=2.5 * cm,
                preserveAspectRatio=True,
                mask="auto",
            )
        canvas.setFont("Helvetica", 8)
        canvas.line(
            doc.leftMargin,
            doc.bottomMargin,
            doc.width + doc.leftMargin,
            doc.bottomMargin,
        )
        canvas.drawString(doc.leftMargin, doc.bottomMargin - 0.5 * cm, FOOTER_TEXT)
        canvas.restoreState()

    doc = SimpleDocTemplate(
        file_path,
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=3 * cm,
        bottomMargin=2.5 * cm,
    )
    styles = getSampleStyleSheet()
    story = [
        Paragraph("BON DE RESTITUTION DE MATÉRIEL", styles["Heading1"]),
        Spacer(1, 1 * cm),
    ]
    intro = f"Le présent document atteste que <b>{user_name}</b> a restitué les équipements suivants au service informatique ce jour :"
    story.append(Paragraph(intro, styles["Normal"]))
    story.append(Spacer(1, 0.5 * cm))

    table_data = [["TYPE", "MODÈLE / NOM", "N° INVENTAIRE", "S/N"]]
    for asset in assets_list:
        table_data.append(
            [
                str(asset.get("_display_type", "N/A")),
                str(asset.get(MODEL_FIELD_ID) or asset.get(NAME_FIELD_ID, "N/A"))[:35],
                str(asset.get(INVENTORY_FIELD_ID, "N/A")),
                str(asset.get(SERIAL_FIELD_ID, "N/A")),
            ]
        )

    t = Table(table_data, colWidths=[3 * cm, 7 * cm, 3.5 * cm, 3.5 * cm])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEEEEE")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(t)
    story.append(Spacer(1, 2 * cm))
    story.append(
        Paragraph(
            f"Fait à Cran-Gevrier, le {datetime.now().strftime('%d/%m/%Y')}",
            styles["Normal"],
        )
    )
    sig_data = [["Signature Collaborateur", "Visa Service Informatique"]]
    sig_t = Table(sig_data, colWidths=[8 * cm, 8 * cm])
    sig_t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]
        )
    )
    story.append(sig_t)
    doc.build(story, onFirstPage=add_header_footer, onLaterPages=add_header_footer)
    return file_path, pdf_filename


# ==============================================================================
# ACTIONS WORKFLOW
# ==============================================================================


async def handle_list(api, sam, name):
    print_header(f"INVENTAIRE : {name}")
    assets = []
    prefix = sam.split("@")[0]
    for tid, info in ITEM_TYPES.items():
        results = api.search(
            info["type"],
            [{"field": CONTACT_FIELD_ID, "searchtype": "contains", "value": prefix}],
        )
        for r in results:
            r["_itype"] = info["type"]
            r["_display_type"] = info["display"]
            assets.append(r)

    if not assets:
        print_warning("Aucun matériel trouvé.")
        return

    display_assets_table(assets, is_selection=False)


async def handle_assign(api, sam, name):
    print_header(f"ATTRIBUTION : {name}")
    for k, v in ITEM_TYPES.items():
        print(f"  {Color.CYAN}{k}{Color.ENDC}. {v['display']}")

    c = (
        input(f"\n{Color.BLUE}Choix type (ou R pour retour) : {Color.ENDC}")
        .strip()
        .upper()
    )
    if c in ["R", "0", ""]:
        return
    if c not in ITEM_TYPES:
        return
    info = ITEM_TYPES[c]

    print_info("Recherche du stock disponible...")
    stock = api.search(
        info["type"],
        [
            {
                "field": STATUS_FIELD_ID,
                "searchtype": "contains",
                "value": STATUS_VALUE_STOCK,
            }
        ],
    )
    if not stock:
        print_warning("Stock vide.")
        return

    for item in stock:
        item["_display_type"] = info["display"]
        item["_itype"] = info["type"]

    ordered_stock = display_assets_table(stock, is_selection=True)

    ans = (
        input(f"\n{Color.BLUE}Numéro à attribuer (ou R pour retour) : {Color.ENDC}")
        .strip()
        .upper()
    )
    if ans in ["R", "0", ""]:
        return

    if ans.isdigit() and 0 < int(ans) <= len(ordered_stock):
        asset = ordered_stock[int(ans) - 1]
        inv_label = (
            asset.get(INVENTORY_FIELD_ID)
            if str(asset.get(INVENTORY_FIELD_ID)).lower() != "none"
            else "SANS INV"
        )
        if api.update(
            info["type"], asset["2"], {"states_id": STATUS_ID_SERVICE, "contact": sam}
        ):
            print_success(
                f"Matériel {inv_label} ({str(asset.get(MODEL_FIELD_ID) or asset.get(NAME_FIELD_ID))}) attribué."
            )
        else:
            print_error("Erreur de mise à jour GLPI.")


async def handle_return_bulk(api, sam, name, ctx, local_logo):
    print_header(f"RESTITUTION GROUPÉE : {name}")
    raw_assets = []
    prefix = sam.split("@")[0]

    print_info("Scan du matériel affecté...")
    for tid, info in ITEM_TYPES.items():
        results = api.search(
            info["type"],
            [{"field": CONTACT_FIELD_ID, "searchtype": "contains", "value": prefix}],
        )
        for r in results:
            r["_itype"] = info["type"]
            r["_display_type"] = info["display"]
            raw_assets.append(r)

    if not raw_assets:
        print_warning("Aucun matériel à restituer.")
        return

    # Affichage avec tableau trié et coloré
    ordered_assets = display_assets_table(raw_assets, is_selection=True)

    print(
        f"\n{Color.BLUE}Entrez les IDs séparés par des virgules (ex: 1,3), 'TOUT' ou 'R' pour retour : {Color.ENDC}"
    )
    selection_raw = input("Choix : ").strip().upper()

    if selection_raw in ["R", "0", ""]:
        return

    selected_assets = []
    if selection_raw == "TOUT":
        selected_assets = ordered_assets
    else:
        try:
            indices = [
                int(x.strip()) - 1
                for x in selection_raw.split(",")
                if x.strip().isdigit()
            ]
            selected_assets = [
                ordered_assets[i] for i in indices if 0 <= i < len(ordered_assets)
            ]
        except:
            print_error("Sélection invalide.")
            return

    if not selected_assets:
        print_warning("Aucune sélection valide.")
        return

    print_warning(
        f"Confirmation : Remise en stock de {len(selected_assets)} élément(s)."
    )
    if input(f"{Color.BOLD}Confirmer (o/n) ? {Color.ENDC}").lower() != "o":
        return

    success_list = []
    for asset in selected_assets:
        update_data = {
            "states_id": STATUS_ID_STOCK,
            "users_id": 0,
            "groups_id": 0,
            "contact": "",
        }
        if api.update(asset["_itype"], asset["2"], update_data):
            success_list.append(asset)
            inv_label = (
                asset.get(INVENTORY_FIELD_ID)
                if str(asset.get(INVENTORY_FIELD_ID)).lower() != "none"
                else "SANS INV"
            )
            print_success(f"Restitué : {inv_label}")
        else:
            print_error(f"Échec GLPI : {asset.get(INVENTORY_FIELD_ID)}")

    if success_list:
        path, filename = create_bulk_restitution_pdf(name, success_list, local_logo)
        try:
            with open(path, "rb") as f:
                ctx.web.get_folder_by_server_relative_url(
                    SP_PDF_RESTITUTION_URL
                ).upload_file(filename, f.read()).execute_query()
            print_success("Fiche groupée archivée sur SharePoint.")
            webbrowser.open(f"file://{os.path.realpath(path)}")
        except Exception as e:
            print_error(f"Erreur SharePoint : {e}")


# ==============================================================================
# POINT D'ENTRÉE
# ==============================================================================


async def run(ctx: ClientContext, ad_conn: Connection, **kwargs):
    print_header("GESTION DU PARC GLPI")
    api = GlpiModule()
    if not api.connect():
        print_error("Échec de connexion API GLPI.")
        return

    temp_dir = tempfile.gettempdir()
    local_logo = os.path.join(temp_dir, "groupe_entis.png")
    if not os.path.exists(local_logo):
        try:
            content = File.open_binary(ctx, SP_LOGO_URL).content
            with open(local_logo, "wb") as f:
                f.write(content)
        except:
            local_logo = None

    while True:
        print_header("SÉLECTION DU COLLABORATEUR")
        term = input(
            f"{Color.BLUE}Login ou Nom de l'usager (X pour quitter) : {Color.ENDC}"
        ).strip()
        if term.upper() == "X" or not term:
            break

        ad_conn.search(
            AD_SEARCH_BASE,
            f"(|(sAMAccountName=*{term}*)(displayName=*{term}*))",
            attributes=["sAMAccountName", "displayName"],
        )
        if not ad_conn.entries:
            print_warning("Utilisateur introuvable.")
            continue

        target = ad_conn.entries[0]
        if len(ad_conn.entries) > 1:
            for i, e in enumerate(ad_conn.entries):
                print(f"  {i + 1}. {e.displayName.value} [{e.sAMAccountName.value}]")
            c = input(f"\nSélectionnez le numéro (ou R pour retour) : ").strip().upper()
            if c in ["R", "0"]:
                continue
            target = (
                ad_conn.entries[int(c) - 1]
                if (c.isdigit() and int(c) <= len(ad_conn.entries))
                else None
            )

        if not target:
            continue
        curr_sam, curr_name = (
            str(target.sAMAccountName.value),
            str(target.displayName.value),
        )

        while True:
            print_header(f"GESTION : {curr_name.upper()}")
            print(f"  {Color.CYAN}1{Color.ENDC}. Lister le matériel affecté")
            print(f"  {Color.CYAN}2{Color.ENDC}. Attribuer un nouveau matériel")
            print(f"  {Color.CYAN}3{Color.ENDC}. Restituer du matériel (GROUPÉ)")
            print(f"  {Color.YELLOW}R{Color.ENDC}. Retour à la recherche utilisateur")

            choice = input(f"\nChoix : ").strip().upper()
            if choice == "1":
                await handle_list(api, curr_sam, curr_name)
            elif choice == "2":
                await handle_assign(api, curr_sam, curr_name)
            elif choice == "3":
                await handle_return_bulk(api, curr_sam, curr_name, ctx, local_logo)
            elif choice == "R":
                break

            input(f"\n{Color.BLUE}Appuyez sur Entrée pour continuer...{Color.ENDC}")


if __name__ == "__main__":
    print_error("Ce module doit être piloté par 'hub_central.py'.")
