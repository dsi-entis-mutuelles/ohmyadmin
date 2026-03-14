# -*- coding: utf-8 -*-

"""
Auteur : JEL pour Groupe Entis
Rôle   : Orchestration de la remise de matériel informatique (PC/MAC). Gère le cycle de vie
         dans l'AD local (Windows uniquement), l'assignation Intune et la génération de fiches de prêt PDF.

Version : 1.9.9
Date    : 26/12/2024

Historique :
-----------
v1.9.9 (26/12/2024) : Mise en conformité stricte des libellés PDF et sécurisation de la nomenclature SharePoint.
v1.9.8 (26/12/2024) : Ajout du logo en bas à gauche de la fiche PDF pour conformité totale.
v1.9.7 (26/12/2024) : Refonte totale du PDF pour conformité stricte avec le modèle officiel (clauses juridiques).
"""

import os
import sys
import getpass
import asyncio
import tempfile
import webbrowser
import pandas as pd
import io
import json
import requests
from datetime import datetime
import re

# Bibliothèques tierces
from pykeepass import PyKeePass
from pykeepass.exceptions import CredentialsError
import msal
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.users_request_builder import UsersRequestBuilder
from msgraph.generated.device_management.managed_devices.managed_devices_request_builder import (
    ManagedDevicesRequestBuilder,
)
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from ldap3 import Server, Connection, ALL, SIMPLE, MODIFY_REPLACE, MODIFY_ADD
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from ldap3.utils.dn import escape_rdn

# ==============================================================================
# CONFIGURATION ET CONSTANTES
# ==============================================================================

SHAREPOINT_SUIVI_FILE_URL = "/sites/GLPI/Data/new_users.xlsx"
SHAREPOINT_PDF_FOLDER_URL = "/sites/GLPI/Documents partages/Ordinateurs"
SHAREPOINT_LOGO_URL = "/sites/GLPI/Data/groupe_entis.png"
BASE_DN = "DC=cetremut,DC=pri"
GROUP_SEARCH_BASE_OU = f"OU=ENTIS,{BASE_DN}"
COMPLIANCE_GROUP_NAME = "GC_MDM_Intune_Compliance_Pilot"
AUTHORITY_URL = "https://login.microsoftonline.com/"
GRAPH_API_SCOPE = ["https://graph.microsoft.com/.default"]

# ==============================================================================
# INTERFACE UTILISATEUR ET STYLISATION
# ==============================================================================


class Color:
    HEADER, BLUE, CYAN, GREEN, WARNING, FAIL, ENDC, BOLD, UNDERLINE = (
        "\033[95m",
        "\033[94m",
        "\033[96m",
        "\033[92m",
        "\033[93m",
        "\033[91m",
        "\033[0m",
        "\033[1m",
        "\033[4m",
    )


def print_header(message):
    print(
        f"\n{Color.HEADER}{Color.BOLD}{'=' * 80}{Color.ENDC}\n{Color.HEADER}{Color.BOLD} {message.center(78)}{Color.ENDC}\n{Color.HEADER}{Color.BOLD}{'=' * 80}{Color.ENDC}"
    )


def print_success(message):
    print(f"{Color.GREEN}[✓] {message}{Color.ENDC}")


def print_warning(message):
    print(f"{Color.WARNING}[!] {message}{Color.ENDC}")


def print_error(message):
    print(f"{Color.FAIL}[X] {message}{Color.ENDC}")


def print_info(message):
    print(f"{Color.CYAN}[i] {message}{Color.ENDC}")


# ==============================================================================
# SERVICES RÉSEAU ET MICROSOFT GRAPH
# ==============================================================================


def download_file_from_sp(ctx, file_url):
    try:
        response = File.open_binary(ctx, file_url)
        return response.content
    except Exception as e:
        print_error(f"Erreur téléchargement {file_url}: {e}")
        raise e


async def assign_user_to_device(access_token, device_id, user_id):
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        request_body = {
            "@odata.id": f"https://graph.microsoft.com/v1.0/users/{user_id}"
        }
        request_url = f"https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{device_id}/users/$ref"

        loop = asyncio.get_running_loop()
        response = await loop.run_in_executor(
            None, lambda: requests.post(request_url, headers=headers, json=request_body)
        )
        response.raise_for_status()
        print_success("Assignation 'Primary User' Intune effectuée.")
    except Exception as e:
        print_error(f"Échec assignation Intune : {e}")
        raise


# ==============================================================================
# GESTION DES DOCUMENTS ET GÉNÉRATION PDF
# ==============================================================================


def get_accessories():
    print_header("SAISIE DES ACCESSOIRES")
    items = [
        "Chargeur",
        "Adaptateur",
        "Casque Sans fils",
        "Souris",
        "Tapis de souris",
        "Clavier",
        "Sacoche",
        "Docking",
        "Écran",
    ]
    selected = []
    for item in items:
        while True:
            val = input(f"{Color.BLUE}Quantité '{item}' (0 par défaut) : {Color.ENDC}")
            if val.isdigit():
                if int(val) > 0:
                    selected.append(f"{item} (x{val})")
                break
            print_warning("Nombre requis.")
    return ", ".join(selected) if selected else "Aucun"


def create_pdf(file_path, data, local_logo_path=None):
    def add_header_footer(canvas, doc):
        canvas.saveState()
        if local_logo_path and os.path.exists(local_logo_path):
            logo = ImageReader(local_logo_path)
            # Logo en haut à droite
            canvas.drawImage(
                logo,
                doc.width + doc.leftMargin - 3 * cm,
                doc.height + doc.topMargin - 1.5 * cm,
                width=3 * cm,
                preserveAspectRatio=True,
                mask="auto",
            )
            # Logo en bas à gauche
            canvas.drawImage(
                logo,
                doc.leftMargin,
                doc.bottomMargin - 0.8 * cm,
                width=2.5 * cm,
                preserveAspectRatio=True,
                mask="auto",
            )

        # Pied de page officiel
        canvas.setFont("Helvetica", 8)
        canvas.line(
            doc.leftMargin,
            doc.bottomMargin - 1 * cm,
            doc.width + doc.leftMargin,
            doc.bottomMargin - 1 * cm,
        )
        footer_text = "Groupe Entis | 39, rue du Jourdil 74960 Cran-Gevrier | Tel : 09 69 39 96 96"
        canvas.drawCentredString(A4[0] / 2, doc.bottomMargin - 1.5 * cm, footer_text)
        canvas.restoreState()

    doc = SimpleDocTemplate(
        file_path,
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=3 * cm,
        bottomMargin=3 * cm,
    )
    styles = getSampleStyleSheet()

    style_title = ParagraphStyle(
        "OfficialTitle",
        parent=styles["Heading1"],
        fontSize=16,
        alignment=1,
        spaceAfter=20,
    )
    style_body = ParagraphStyle(
        "OfficialBody", parent=styles["Normal"], fontSize=11, leading=16, spaceAfter=12
    )
    style_bold = ParagraphStyle(
        "OfficialBold", parent=style_body, fontName="Helvetica-Bold"
    )

    story = [
        Paragraph("Fiche d'attribution de matériel", style_title),
        Spacer(1, 1 * cm),
    ]

    # Texte officiel conforme à l'exemple
    intro = f"Entis confie à <b>{data['user_display_name']}</b> pendant toute la durée de son contrat de travail, le matériel suivant :"
    story.append(Paragraph(intro, style_body))
    story.append(Spacer(1, 0.5 * cm))

    # Détails techniques avec libellés officiels
    story.append(Paragraph(f"<b>Type de matériel :</b> {data['os_type']}", style_body))
    story.append(
        Paragraph(f"<b>Marque et modèle :</b> {data['device_model']}", style_body)
    )
    story.append(
        Paragraph(
            f"<b>Numéro de série :</b> {data['device_serial_number']}", style_body
        )
    )
    story.append(
        Paragraph(f"<b>Inventaire interne :</b> {data['pc_name']}", style_body)
    )
    story.append(Paragraph(f"<b>Accessoires :</b> {data['accessories']}", style_body))

    story.append(Spacer(1, 0.5 * cm))

    # Clauses contractuelles
    clauses = [
        "Ce matériel demeure la propriété de l'entreprise.",
        f"Il est convenu que <b>{data['user_display_name']}</b> ne pourra pas utiliser ce matériel à d'autres fins que professionnelles sans autorisation expresse et préalable de l'entreprise.",
        f"<b>{data['user_display_name']}</b> s'engage à restituer l'intégralité du matériel confié au moment de la rupture de ses relations contractuelles avec Entis et ce, quelques soient les motifs de cette rupture.",
        "Pour faire valoir ce que de droit.",
    ]

    for clause in clauses:
        story.append(Paragraph(clause, style_body))

    story.append(Spacer(1, 1 * cm))
    story.append(
        Paragraph(
            f"Fait à Cran-Gevrier, le {datetime.now().strftime('%d/%m/%Y')}", style_body
        )
    )
    story.append(Spacer(1, 1 * cm))

    story.append(
        Paragraph("Signature (précédée de la mention << lu et approuvé >>)", style_bold)
    )

    doc.build(story, onFirstPage=add_header_footer, onLaterPages=add_header_footer)


def get_unique_local_filepath(directory, filename):
    base, ext = os.path.splitext(filename)
    unique, counter = os.path.join(directory, filename), 1
    while os.path.exists(unique):
        unique = os.path.join(directory, f"{base} ({counter}){ext}")
        counter += 1
    return unique


# ==============================================================================
# OPÉRATIONS ACTIVE DIRECTORY
# ==============================================================================


def get_site_from_user_dn(user_dn: str) -> str:
    m = re.search(r"OU=([^,]+),OU=ENTIS,", user_dn, re.IGNORECASE)
    if not m:
        raise Exception("Site introuvable dans le DN utilisateur.")
    return m.group(1)


def find_computer_dn(conn, pc_name: str) -> str | None:
    conn.search(
        BASE_DN,
        f"(&(objectClass=computer)(cn={pc_name}))",
        attributes=["distinguishedName"],
    )
    return conn.entries[0].distinguishedName.value if conn.entries else None


def ensure_computer_in_ou(
    conn, computer_dn: str, target_ou_dn: str, pc_name: str
) -> str:
    if computer_dn.lower().endswith("," + target_ou_dn.lower()):
        print_info("L'ordinateur est déjà dans la bonne OU.")
        return computer_dn
    rdn = f"CN={escape_rdn(pc_name)}"
    print_info(f"Déplacement AD vers : {target_ou_dn}")
    conn.modify_dn(computer_dn, rdn, new_superior=target_ou_dn)
    return f"{rdn},{target_ou_dn}"


# ==============================================================================
# LOGIQUE PRINCIPALE ET ORCHESTRATION
# ==============================================================================


async def run(
    ctx: ClientContext,
    kp,
    tenant_id: str,
    ad_conn: Connection,
    config_dir: str,
    **kwargs,
):
    """
    Point d'entrée Hub v2.0.0.
    """
    print_header("SSO : AUTHENTIFICATION MICROSOFT GRAPH")
    try:
        entry = kp.find_entries(title="Azure App Credentials", first=True)
        if not entry:
            raise ValueError("Entrée KeePass 'Azure App Credentials' absente.")
        client_id, client_secret = entry.username, entry.password

        msal_app = msal.ConfidentialClientApplication(
            client_id,
            authority=f"{AUTHORITY_URL}{tenant_id}",
            client_credential=client_secret,
        )
        token_res = msal_app.acquire_token_for_client(scopes=GRAPH_API_SCOPE)
        access_token = token_res["access_token"]

        credential = ClientSecretCredential(tenant_id, client_id, client_secret)
        graph_client = GraphServiceClient(
            credentials=credential, scopes=GRAPH_API_SCOPE
        )
        print_success("Session Graph initialisée via SSO.")
    except Exception as e:
        print_error(f"Erreur d'authentification SSO : {e}")
        return

    # --- PRÉPARATION ---
    temp_dir = tempfile.gettempdir()
    logo_content = await asyncio.to_thread(
        download_file_from_sp, ctx, SHAREPOINT_LOGO_URL
    )
    local_logo = os.path.join(temp_dir, "groupe_entis.png")
    with open(local_logo, "wb") as f:
        f.write(logo_content)

    mode = input(
        f"\n{Color.BLUE}Mode : (1) Suivi SharePoint, (2) Manuel : {Color.ENDC}"
    )
    selected_users = []

    if mode == "1":
        print_info("Lecture du registre de suivi...")
        res = File.open_binary(ctx, SHAREPOINT_SUIVI_FILE_URL)
        df = pd.read_excel(io.BytesIO(res.content)).fillna("")
        pc_requests = df[df["materiel"].str.upper() == "PC"].copy()
        if pc_requests.empty:
            print_warning("Aucune demande PC trouvée.")
            return

        pc_requests["id"] = range(1, len(pc_requests) + 1)
        print(
            pc_requests[["id", "prenom", "nom", "email"]].set_index("id").to_markdown()
        )

        ids = input(f"\n{Color.BLUE}IDs à traiter (ex: 1,2 ou 'tous') : {Color.ENDC}")
        selected_users = (
            pc_requests.to_dict("records")
            if ids.lower() == "tous"
            else pc_requests[
                pc_requests["id"].isin([int(i.strip()) for i in ids.split(",")])
            ].to_dict("records")
        )

    elif mode == "2":
        email = input(f"{Color.BLUE}Email utilisateur : {Color.ENDC}").strip()
        q_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            filter=f"userPrincipalName eq '{email}'"
        )
        req_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=q_params
        )
        u_res = await graph_client.users.get(request_configuration=req_config)

        if u_res and u_res.value:
            selected_users = [
                {
                    "email": email,
                    "prenom": u_res.value[0].display_name,
                    "nom": "",
                    "manual_entry": True,
                }
            ]
        else:
            print_error(f"Utilisateur {email} introuvable dans Azure.")

    for user in selected_users:
        email, name = user["email"], f"{user.get('prenom')} {user.get('nom')}".strip()
        print_header(f"REMISE : {name}")
        try:
            pc_name = input(
                f"{Color.BLUE}Nom du poste (CW pour Windows, CM pour Mac) : {Color.ENDC}"
            ).upper()
            is_mac = pc_name.startswith("CM")

            ad_conn.search(
                BASE_DN,
                f"(&(objectClass=user)(|(mail={email})(userPrincipalName={email})))",
                attributes=["distinguishedName"],
            )
            if not ad_conn.entries:
                raise Exception("Utilisateur AD introuvable.")
            user_dn = ad_conn.entries[0].distinguishedName.value

            if not is_mac:
                print_info("Mise à jour Active Directory...")
                comp_dn = find_computer_dn(ad_conn, pc_name)
                if not comp_dn:
                    raise Exception("Ordinateur AD introuvable.")
                site = get_site_from_user_dn(user_dn)
                target_ou = (
                    f"OU=Laptop,OU=Windows,OU=Computer,OU={site},OU=ENTIS,{BASE_DN}"
                )
                new_comp_dn = ensure_computer_in_ou(
                    ad_conn, comp_dn, target_ou, pc_name
                )
                ad_conn.search(
                    GROUP_SEARCH_BASE_OU,
                    f"(&(objectClass=group)(cn={COMPLIANCE_GROUP_NAME}))",
                    attributes=["distinguishedName"],
                )
                if ad_conn.entries:
                    ad_conn.modify(
                        ad_conn.entries[0].distinguishedName.value,
                        {"member": [(MODIFY_ADD, [new_comp_dn])]},
                    )

            dev_res = await graph_client.device_management.managed_devices.get(
                request_configuration=ManagedDevicesRequestBuilder.ManagedDevicesRequestBuilderGetRequestConfiguration(
                    query_parameters=ManagedDevicesRequestBuilder.ManagedDevicesRequestBuilderGetQueryParameters(
                        filter=f"deviceName eq '{pc_name}'"
                    )
                )
            )
            if not dev_res or not dev_res.value:
                raise Exception(f"Poste {pc_name} absent d'Intune.")
            device = dev_res.value[0]

            user_az = await graph_client.users.get(
                request_configuration=UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
                    query_parameters=UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
                        filter=f"userPrincipalName eq '{email}'"
                    )
                )
            )
            await assign_user_to_device(access_token, device.id, user_az.value[0].id)

            # Nomenclature du fichier
            os_label = "macOS Laptop" if is_mac else "Windows Laptop"
            p_data = {
                "user_display_name": name,
                "pc_name": pc_name,
                "device_model": device.model,
                "device_serial_number": device.serial_number,
                "accessories": get_accessories(),
                "os_type": os_label,
            }

            # --- NOMENCLATURE DE NOMMAGE ---
            # Format: YYYYMMDD_NomPoste_Utilisateur.pdf
            base_filename = f"{datetime.now().strftime('%Y%m%d')}_{pc_name}_{name.replace(' ', '_')}.pdf"
            local_pdf_path = get_unique_local_filepath(temp_dir, filename=base_filename)
            final_filename = os.path.basename(
                local_pdf_path
            )  # Récupère le nom final (si doublon local)

            create_pdf(local_pdf_path, p_data, local_logo)
            webbrowser.open(f"file://{os.path.realpath(local_pdf_path)}")

            # --- REMONTÉE SHAREPOINT ---
            with open(local_pdf_path, "rb") as f:
                ctx.web.get_folder_by_server_relative_url(
                    SHAREPOINT_PDF_FOLDER_URL
                ).upload_file(final_filename, f.read()).execute_query()

            if not user.get("manual_entry"):
                reg = File.open_binary(ctx, SHAREPOINT_SUIVI_FILE_URL)
                rdf = pd.read_excel(io.BytesIO(reg.content))
                rdf.loc[rdf["email"] == email, "materiel"] = "Attribué"
                out = io.BytesIO()
                rdf.to_excel(out, index=False, engine="openpyxl")
                ctx.web.get_folder_by_server_relative_url(
                    os.path.dirname(SHAREPOINT_SUIVI_FILE_URL)
                ).upload_file(
                    os.path.basename(SHAREPOINT_SUIVI_FILE_URL), out.getvalue()
                ).execute_query()

            print_success(
                f"Dossier finalisé et archivé sur SharePoint : {final_filename}"
            )
        except Exception as e:
            print_error(f"Erreur {name} : {e}")


if __name__ == "__main__":
    print_error("Ce module doit être piloté par 'hub_central.py'.")
