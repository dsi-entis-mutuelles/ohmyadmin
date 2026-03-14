# -*- coding: utf-8 -*-
"""
Synopsis:
    Script de synchronisation des utilisateurs Active Directory et de vérification
    des numéros de téléphone dans Intune à partir d'un fichier Excel RH.

Description:
    Version finale restaurée : conserve l'interactivité originale (choix managers),
    gère ses propres dépendances critiques (inputimeout) et utilise le SSO du Hub.

Version : 1.3.4
Date    : 26/12/2024
"""

import os
import sys
import subprocess
import re
import logging
import tempfile
import io
import asyncio
import pandas as pd
import unicodedata
import requests
import webbrowser
from datetime import datetime, timedelta

# --- Gestion des dépendances spécifiques à l'interactivité ---
try:
    from inputimeout import inputimeout, TimeoutOccurred
except ImportError:
    print("[i] Installation de la dépendance interactive 'inputimeout'...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "inputimeout"])
    from inputimeout import inputimeout, TimeoutOccurred

from msal import ConfidentialClientApplication
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from ldap3 import Server, Connection, ALL, SIMPLE, MODIFY_REPLACE, MODIFY_DELETE
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    KeepTogether,
    PageBreak,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors

# ==============================================================================
# CONFIGURATION ET CONSTANTES
# ==============================================================================

AD_SEARCH_BASE = "DC=cetremut,DC=pri"
SP_DSI_SITE_URL = "https://entis.sharepoint.com/sites/PRI-DSI"
SP_DOC_LIBRARY_URL = "/sites/PRI-DSI/Documents Exploit/Active Directory/RH-SYNC"
SP_APP_CLIENT_ID = "162abf7d-37f3-4e58-87f2-be66877ac142"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]

# Mapping Excel RH
SHEET_SALARIE, SHEET_TELEPHONE = "Salariés et Utilisateurs", "Téléphones"
COL_NOM, COL_PRENOM, COL_EMAIL, COL_MANAGER_EMAIL = (
    "Nom",
    "Prénom",
    "Email du salarié",
    "Email Manager",
)
COL_FONCTION, COL_STRUCTURE, COL_DEPARTEMENT, COL_TELEPHONE = (
    "Fonction",
    "Structure",
    "Département",
    "Numéro",
)
COL_DATE_FIN_CONTRAT = "Date fin de contrat"

# Styles PDF
ENTIS_BLUE = colors.HexColor("#003333")
ENTIS_GREEN = colors.HexColor("#003333")
ENTIS_LIGHT_BLUE = colors.HexColor("#faf4de")


class Color:
    HEADER, BLUE, CYAN, GREEN, WARNING, FAIL, ENDC, BOLD = (
        "\033[95m",
        "\033[94m",
        "\033[96m",
        "\033[92m",
        "\033[93m",
        "\033[91m",
        "\033[0m",
        "\033[1m",
    )


# ==============================================================================
# FONCTIONS UTILITAIRES
# ==============================================================================


def print_header(m):
    print(
        f"\n{Color.HEADER}{Color.BOLD}{'=' * 80}\n {m.center(78)}\n{'=' * 80}{Color.ENDC}"
    )


def print_success(m):
    logging.info(f"{Color.GREEN}[✓] {m}{Color.ENDC}")


def print_warning(m):
    logging.warning(f"{Color.WARNING}[!] {m}{Color.ENDC}")


def print_error(m):
    logging.error(f"{Color.FAIL}[X] {m}{Color.ENDC}")


def print_info(m):
    logging.info(f"{Color.CYAN}[i] {m}{Color.ENDC}")


def print_dynamic_info(m):
    sys.stdout.write(f"\r{Color.CYAN}[i] {m.ljust(100)}{Color.ENDC}")
    sys.stdout.flush()


def clean_normalize(s):
    return (
        unicodedata.normalize("NFC", str(s).replace("_x000D_", " ").strip())
        if s and pd.notna(s)
        else ""
    )


def format_phone_for_display(phone_number):
    if phone_number and len(phone_number) == 10 and phone_number.isdigit():
        return " ".join(phone_number[i : i + 2] for i in range(0, 10, 2))
    return phone_number if phone_number else "Non défini"


def validate_format_mobile(p):
    if not p or pd.isna(p):
        return None
    normalized = re.sub(r"[\s./-]", "", str(p))
    match = re.match(r"^(?:0|\+33|0033|33)([67]\d{8})$", normalized)
    if match:
        return "0" + match.group(1)
    return None


def filetime_to_datetime(ft):
    if not ft:
        return None
    try:
        val = int(ft)
        if val == 0 or val == 9223372036854775807:
            return None
        return datetime(1601, 1, 1) + timedelta(microseconds=val / 10)
    except:
        return None


def datetime_to_filetime(dt):
    if not dt:
        return 0
    return int((dt - datetime(1601, 1, 1)).total_seconds() * 10000000)


def setup_logging(mode, file_name):
    log_dir = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(
        log_dir,
        f"rh_sync_{mode}_{os.path.splitext(file_name)[0]}_{datetime.now():%Y%m%d_%H%M%S}.log",
    )
    root_logger = logging.getLogger()
    root_logger.handlers.clear()
    root_logger.setLevel(logging.INFO)
    file_handler = logging.FileHandler(log_file, "w", "utf-8")
    file_handler.setFormatter(logging.Formatter("%(asctime)s - %(message)s"))
    root_logger.addHandler(file_handler)
    return log_file


# ==============================================================================
# LOGIQUE MICROSOFT GRAPH (SSO Hub)
# ==============================================================================


async def get_graph_token_from_hub(kp, tenant_id):
    """Utilise la session KeePass du Hub pour obtenir le token Graph."""
    try:
        entry = kp.find_entries(title="Azure App Credentials", first=True)
        if not entry:
            raise ValueError("Entrée 'Azure App Credentials' absente du KeePass.")

        app = ConfidentialClientApplication(
            client_id=entry.username,
            client_credential=entry.password,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
        )
        result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
        return result.get("access_token")
    except Exception as e:
        print_error(f"Échec authentification Graph via Hub : {e}")
        return None


def check_intune_phone_numbers(graph_token, upn, excel_phone_normalized):
    headers = {"Authorization": f"Bearer {graph_token}", "Accept": "application/json"}
    url = f"https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$filter=userPrincipalName eq '{upn}' and operatingSystem eq 'iOS'&$select=deviceName,phoneNumber"
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        devices = response.json().get("value", [])
        if not excel_phone_normalized or not devices:
            return []
        match_found, details = False, []
        for device in devices:
            intune_norm = validate_format_mobile(device.get("phoneNumber"))
            if intune_norm == excel_phone_normalized:
                match_found = True
                break
            details.append(
                f"'{device.get('deviceName')}' ({device.get('phoneNumber') or 'vide'})"
            )
        if not match_found:
            return [
                f"ALERTE INTUNE: Numéro RH ({format_phone_for_display(excel_phone_normalized)}) non trouvé sur les terminaux Intune : {', '.join(details)}."
            ]
        return []
    except:
        return [
            f"ALERTE INTUNE: Erreur lors de la vérification des terminaux pour {upn}."
        ]


# ==============================================================================
# GÉNÉRATION DE RAPPORT PDF
# ==============================================================================


def generate_pdf_report(file_path, modifications, mode, excel_file, logo_path, stats):
    def header_footer(canvas, doc):
        canvas.saveState()
        if logo_path and os.path.exists(logo_path):
            logo = ImageReader(logo_path)
            canvas.drawImage(
                logo,
                doc.width + doc.leftMargin - 3 * cm,
                doc.height + doc.topMargin + 1 * cm,
                width=2.5 * cm,
                preserveAspectRatio=True,
                mask="auto",
            )
        canvas.setStrokeColor(ENTIS_BLUE)
        canvas.line(
            doc.leftMargin,
            doc.bottomMargin,
            doc.width + doc.leftMargin,
            doc.bottomMargin,
        )
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(ENTIS_BLUE)
        canvas.drawString(
            doc.leftMargin,
            doc.bottomMargin - 0.5 * cm,
            "Groupe Entis - Rapport de Synchronisation Automatisé",
        )
        canvas.drawRightString(
            doc.width + doc.leftMargin,
            doc.bottomMargin - 0.5 * cm,
            f"Généré le {datetime.now():%d/%m/%Y %H:%M}",
        )
        canvas.restoreState()

    doc = SimpleDocTemplate(
        file_path,
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=3 * cm,
        bottomMargin=2.5 * cm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        name="Title", parent=styles["h1"], alignment=TA_CENTER, textColor=ENTIS_BLUE
    )

    story = [
        Paragraph("Rapport de Synchronisation RH & Intune", title_style),
        Spacer(1, 1 * cm),
    ]
    story.append(
        Paragraph(
            f"<b>Fichier source :</b> {excel_file}<br/><b>Mode :</b> {mode.upper()}",
            styles["Normal"],
        )
    )
    story.append(Spacer(1, 1 * cm))

    for user_data in modifications:
        table_data = [
            [Paragraph(f"<b>{user_data['name']}</b>", styles["Normal"]), None],
            ["Type", "Changement"],
        ]
        for change in user_data["changes"]:
            c_type = "Alerte" if "ALERTE" in change else "MàJ"
            table_data.append([c_type, Paragraph(change, styles["BodyText"])])

        t = Table(table_data, colWidths=[3.5 * cm, 14.5 * cm])
        t.setStyle(
            TableStyle(
                [
                    ("SPAN", (0, 0), (1, 0)),
                    ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#006666")),
                    ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
                    ("GRID", (0, 1), (-1, -1), 0.5, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ]
            )
        )
        story.append(KeepTogether(t))
        story.append(Spacer(1, 0.5 * cm))

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)


# ==============================================================================
# LOGIQUE MÉTIER ET INTERACTIVITÉ
# ==============================================================================


def find_ad_obj(conn, filt, attrs=None):
    conn.search(AD_SEARCH_BASE, filt, attributes=attrs or ["distinguishedName"])
    return conn.entries[0] if conn.entries else None


def get_ad_attr(entry, attr_name):
    for attr in entry.entry_attributes:
        if attr.lower() == attr_name.lower():
            return entry[attr]
    return None


# ==============================================================================
# POINT D'ENTRÉE DU MODULE (RUN)
# ==============================================================================


async def run(ctx, kp, tenant_id, ad_conn, config_dir, **kwargs):
    print_header("VÉRIFICATION RH & SYNCHRO MOBILE INTUNE")

    # 1. Sélection du mode
    print_info("Choisissez le mode d'exécution :")
    print(f"  {Color.CYAN}1{Color.ENDC}. Simulation (vérifications sans modifications)")
    print(f"  {Color.CYAN}2{Color.ENDC}. Production (applique les changements AD)")
    mode_choice = input(f"\n{Color.BLUE}Votre choix : {Color.ENDC}")
    exec_mode = "production" if mode_choice == "2" else "simulation"

    # 2. Token Graph via SSO Hub
    graph_token = await get_graph_token_from_hub(kp, tenant_id)

    # 3. Connexion SharePoint PRI-DSI pour le fichier RH
    print_info("Connexion au site SharePoint PRI-DSI...")
    try:
        ctx_dsi = ClientContext(SP_DSI_SITE_URL).with_interactive(
            tenant_id, SP_APP_CLIENT_ID
        )
        folder = ctx_dsi.web.get_folder_by_server_relative_url(SP_DOC_LIBRARY_URL)
        files = folder.files
        ctx_dsi.load(files)
        ctx_dsi.execute_query()
        hr_files = sorted(
            [
                f
                for f in files
                if re.match(r"Entrées_-_sorties_informatique_.*\.xlsx?$", f.name, re.I)
            ],
            key=lambda f: f.time_last_modified,
            reverse=True,
        )

        if not hr_files:
            print_error("Aucun fichier RH trouvé.")
            return
        target_file = hr_files[0]
        local_excel = os.path.join(config_dir, target_file.name)
        with open(local_excel, "wb") as f:
            target_file.download(f).execute_query()

        setup_logging(exec_mode, target_file.name)

        # 4. Traitement Excel
        df_salarie = pd.read_excel(local_excel, sheet_name=SHEET_SALARIE)
        df_telephone = pd.read_excel(local_excel, sheet_name=SHEET_TELEPHONE)
        phone_lookup = {
            (
                clean_normalize(r.get(COL_NOM)),
                clean_normalize(r.get(COL_PRENOM)),
            ): r.get(COL_TELEPHONE)
            for _, r in df_telephone.iterrows()
        }

        modifications_list = []
        stats = {"s": 0, "m": 0, "err": 0}
        p_prefix = "[SIMULATION] " if exec_mode == "simulation" else ""

        for idx, row in df_salarie.iterrows():
            email = clean_normalize(row.get(COL_EMAIL))
            if not email or row.get(COL_NOM) == "Nom":
                continue

            print_dynamic_info(f"Analyse {email} ({idx + 1}/{len(df_salarie)})...")
            user_ad = find_ad_obj(
                ad_conn,
                f"(&(objectClass=user)(mail={email}))",
                [
                    "distinguishedName",
                    "displayName",
                    "manager",
                    "title",
                    "company",
                    "department",
                    "mobile",
                    "accountExpires",
                    "userAccountControl",
                ],
            )

            if not user_ad:
                sys.stdout.write("\r" + " " * 120 + "\r")
                print_warning(f"Utilisateur {email} non trouvé dans l'AD.")
                stats["err"] += 1
                continue

            dn, chg, logs = user_ad.distinguishedName.value, {}, []
            user_fullname = (
                user_ad.displayName.value if "displayName" in user_ad else email
            )

            # --- LOGIQUE MANAGER INTERACTIVE (RESTAURÉE) ---
            manager_raw = row.get(COL_MANAGER_EMAIL)
            target_manager_dn = None
            current_manager_dn = user_ad.manager.value if "manager" in user_ad else None
            if pd.notna(manager_raw) and str(manager_raw).strip():
                emails = [
                    e.strip()
                    for e in re.split(r"[\s,;]+", clean_normalize(manager_raw))
                    if e.strip()
                ]
                if len(emails) > 1:
                    sys.stdout.write("\r" + " " * 120 + "\r")
                    print_warning(f"Plusieurs managers pour {user_fullname}:")
                    for i, m_mail in enumerate(emails):
                        print(f"  {Color.CYAN}{i + 1}{Color.ENDC}. {m_mail}")
                    try:
                        manager_email = emails[
                            int(
                                inputimeout(
                                    f"\n{Color.BLUE}Choisissez (défaut 2 en 30s): {Color.ENDC}",
                                    30,
                                )
                            )
                            - 1
                        ]
                    except (TimeoutOccurred, ValueError, IndexError):
                        manager_email = emails[1]
                        print_warning("Choix par défaut (2) utilisé.")
                else:
                    manager_email = emails[0] if emails else None
                if manager_email:
                    m_ad = find_ad_obj(
                        ad_conn, f"(&(objectClass=user)(mail={manager_email}))"
                    )
                    if m_ad:
                        target_manager_dn = m_ad.distinguishedName.value

            if current_manager_dn != target_manager_dn:
                if target_manager_dn:
                    chg["manager"] = [(MODIFY_REPLACE, [target_manager_dn])]
                    logs.append(f"ACTION : MàJ manager vers '{manager_email}'.")
                elif current_manager_dn:
                    chg["manager"] = [(MODIFY_DELETE, [])]
                    logs.append("ACTION : Supprimer le manager.")

            # Attributs standards
            for attr, col in [
                ("title", COL_FONCTION),
                ("company", COL_STRUCTURE),
                ("department", COL_DEPARTEMENT),
            ]:
                target_val = clean_normalize(row.get(col))
                curr_val = clean_normalize(
                    get_ad_attr(user_ad, attr).value
                    if get_ad_attr(user_ad, attr)
                    else ""
                )
                if target_val and target_val != curr_val:
                    chg[attr] = [(MODIFY_REPLACE, [target_val])]
                    logs.append(
                        f"ACTION : MàJ {attr} : de '{curr_val}' vers '{target_val}'."
                    )

            # Mobile
            rh_phone = phone_lookup.get(
                (
                    clean_normalize(row.get(COL_NOM)),
                    clean_normalize(row.get(COL_PRENOM)),
                )
            )
            rh_mob_norm = validate_format_mobile(rh_phone)
            ad_mob_norm = validate_format_mobile(
                user_ad.mobile.value if "mobile" in user_ad else ""
            )
            if rh_mob_norm and rh_mob_norm != ad_mob_norm:
                ad_fmt = format_phone_for_display(rh_mob_norm)
                chg["mobile"] = [(MODIFY_REPLACE, [ad_fmt])]
                logs.append(f"ACTION : MàJ Mobile vers {ad_fmt}")

            # Intune
            if graph_token:
                logs.extend(check_intune_phone_numbers(graph_token, email, rh_mob_norm))

            # accountExpires
            date_fin = row.get(COL_DATE_FIN_CONTRAT)
            if pd.notna(date_fin):
                try:
                    target_dt = pd.to_datetime(date_fin, dayfirst=True)
                    curr_ft = (
                        user_ad.accountExpires.value
                        if "accountExpires" in user_ad
                        else 0
                    )
                    curr_dt = filetime_to_datetime(curr_ft)
                    if not curr_dt or curr_dt.date() != target_dt.date():
                        chg["accountExpires"] = [
                            (MODIFY_REPLACE, [str(datetime_to_filetime(target_dt))])
                        ]
                        logs.append(
                            f"ACTION : MàJ accountExpires vers {target_dt.strftime('%d/%m/%Y')}."
                        )
                except:
                    pass

            if logs:
                stats["m"] += 1
                modifications_list.append({"name": user_fullname, "changes": logs})
                sys.stdout.write("\r" + " " * 120 + "\r")
                print_info(f"--- Modifications/Alertes : {user_fullname} ---")
                for msg in logs:
                    print_warning(f"  {msg}") if "ALERTE" in msg else print_info(
                        f"  {p_prefix}{msg}"
                    )
                if exec_mode == "production" and chg:
                    ad_conn.modify(dn, chg)
            stats["s"] += 1

        # 5. Finalisation et PDF
        sys.stdout.write("\r" + " " * 120 + "\r")
        print_header("RÉSULTAT DE L'OPÉRATION")
        if modifications_list:
            base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
            pdf_path = os.path.join(
                base_dir, f"Rapport_RH_{datetime.now():%Y%m%d_%H%M}.pdf"
            )
            logo_path = os.path.join(config_dir, "groupe_entis.png")

            generate_pdf_report(
                pdf_path,
                modifications_list,
                exec_mode,
                target_file.name,
                logo_path,
                stats,
            )
            print_success(f"Rapport généré : {pdf_path}")
            webbrowser.open(f"file://{os.path.abspath(pdf_path)}")

            if exec_mode == "production":
                with open(pdf_path, "rb") as f:
                    ctx_dsi.web.get_folder_by_server_relative_url(
                        SP_DOC_LIBRARY_URL
                    ).upload_file(os.path.basename(pdf_path), f.read()).execute_query()

        print_success(
            f"Opération terminée. Traités: {stats['s']}, Modifiés: {stats['m']}, Erreurs: {stats['err']}"
        )

    except Exception as e:
        print_error(f"Erreur critique module RH : {e}")


if __name__ == "__main__":
    print_error("Veuillez lancer ce module via l'orchestrateur 'hub_central.py'.")
