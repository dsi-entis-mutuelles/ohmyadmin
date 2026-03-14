# -*- coding: utf-8 -*-

"""
Auteur : JEL pour Groupe Entis
Rôle   : Assistant interactif pour l'attribution de licences Microsoft 365, la configuration
         des boîtes aux lettres (régionalisation), l'activation MFA et la gestion des
         signatures Letsign-it via Microsoft Graph et l'AD local.

Version : 1.8.9
Date    : 25/12/2024

Historique :
-----------
v1.8.9 (25/12/2024) : Mise en conformité visuelle, simplification des commentaires et du synopsis.
v1.8.8 (10/12/2024) : Intégration de la vérification et de l'activation automatisée de la MFA.
v1.8.0 (05/11/2024) : Ajout de la configuration automatique du fuseau horaire Exchange Online.
"""

import os
import io
import getpass
import pandas as pd
import sys
import json
import asyncio
import time
import re
from pykeepass import PyKeePass
from pykeepass.exceptions import CredentialsError
from ldap3 import Server, Connection, ALL, SIMPLE, MODIFY_REPLACE
from ldap3.core.exceptions import LDAPException
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.users_request_builder import UsersRequestBuilder
from msgraph.generated.models.user import User

# Utilisation de ReferenceCreate pour les ajouts aux groupes
from msgraph.generated.models.reference_create import ReferenceCreate
from msgraph.generated.models.mailbox_settings import MailboxSettings
from msgraph.generated.models.locale_info import LocaleInfo
from kiota_abstractions.api_error import APIError

# ==============================================================================
# CONFIGURATION ET CONSTANTES
# ==============================================================================

BASE_DN = "DC=cetremut,DC=pri"
SHAREPOINT_FILE_URL = "/sites/GLPI/Data/new_users.xlsx"

GROUP_IDS = {
    "business_premium": "cd1bcd92-5fc5-4904-bfdc-a42a3f8cb3e9",
    "exchange_plan_2": "48b9abe9-db9d-4f58-881e-22d2cd421d8c",
    "vpn": "f87449d6-39c1-47c2-8d6b-8cba4f618d09",
    "m365_e3_no_teams": "c3a80a5d-e079-47a7-b247-03d6da0e63fe",
    "only_teams": "89bb8b79-d751-4feb-baa1-cddc71a559cd",
    "mfa": "f122dc47-7276-4848-b5d7-f8ad18e93589",
}

SKU_PART_NUMBERS = {
    "business_premium": "SPB",
    "exchange_plan_2": "EXCHANGEENTERPRISE",
    "m365_e3_no_teams": "O365_w/o Teams Bundle_M3",
    "only_teams": "Microsoft_Teams_EEA_New",
}

SKU_MAPPING = {
    "SPB": "M365 Business Premium",
    "EXCHANGEENTERPRISE": "Exchange Online P2",
    "O365_w/o Teams Bundle_M3": "M365 E3 (sans Teams)",
    "Microsoft_Teams_EEA_New": "Microsoft Teams EEA",
}

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


def load_configs_from_local(config_dir: str):
    print_info("Chargement des référentiels signatures et adresses...")
    try:
        with open(
            os.path.join(config_dir, "Letsign-it.json"), "r", encoding="utf-8-sig"
        ) as f:
            signatures_data = json.load(f)
        with open(
            os.path.join(config_dir, "Adresse_Agences.json"), "r", encoding="utf-8-sig"
        ) as f:
            adresses_data = json.load(f)
        print_success("Configurations locales chargées.")
        return signatures_data, adresses_data
    except Exception as e:
        print_error(f"Erreur de chargement config : {e}")
        return None, None


async def download_excel_from_sharepoint(ctx, file_url):
    try:
        response = await asyncio.to_thread(File.open_binary, ctx, file_url)
        return response.content
    except Exception as e:
        print_error(f"Erreur téléchargement SharePoint : {e}")
        return None


# ==============================================================================
# GESTION DES LICENCES ET CONFIGURATION EXCHANGE
# ==============================================================================


async def get_license_details(graph_client: GraphServiceClient):
    license_details = {}
    try:
        subscribed_skus = await graph_client.subscribed_skus.get()
        if not subscribed_skus or not subscribed_skus.value:
            return {}

        for sku in subscribed_skus.value:
            part_number = sku.sku_part_number.replace("\xa0", " ").strip()
            if part_number in SKU_MAPPING:
                license_details[part_number] = {
                    "total": sku.prepaid_units.enabled,
                    "consumed": sku.consumed_units,
                    "remaining": sku.prepaid_units.enabled - sku.consumed_units,
                    "sku_id": sku.sku_id,
                }
    except Exception as e:
        print_warning(f"Récupération des stocks impossible : {e}")
    return license_details


async def display_license_status(license_details):
    print_header("ÉTAT DES STOCKS LICENCES M365")
    if not license_details:
        return

    table_data = []
    for part_number, details in license_details.items():
        rem = details["remaining"]
        color = Color.GREEN if rem > 5 else (Color.WARNING if rem > 0 else Color.FAIL)
        table_data.append(
            {
                "Licence": SKU_MAPPING.get(part_number, part_number),
                "Total": details["total"],
                "Utilisées": details["consumed"],
                "Restantes": f"{color}{rem}{Color.ENDC}",
            }
        )
    print(pd.DataFrame(table_data).to_markdown(index=False))


def select_users_from_list(users_to_process):
    display_df = users_to_process[["prenom", "nom", "email", "service"]].copy()
    display_df.index.name = "ID"
    print_header("UTILISATEURS EN ATTENTE (A TRAITER)")
    print(display_df.to_markdown())

    while True:
        try:
            user_input = input(
                f"{Color.BLUE}ID à TRAITER (ex: 1,3 ou 'tous') : {Color.ENDC}"
            )
            if user_input.lower() == "tous":
                selected_for_processing = users_to_process
                break
            selected_ids = [int(i.strip()) for i in user_input.split(",")]
            selected_for_processing = users_to_process.loc[selected_ids]
            break
        except:
            print_error("Sélection invalide.")

    remaining = users_to_process.drop(selected_for_processing.index)
    selected_for_ignoring = pd.DataFrame()

    if (
        not remaining.empty
        and input(
            f"{Color.WARNING}❓ Ignorer certains profils restants ? (o/n) : {Color.ENDC}"
        ).lower()
        == "o"
    ):
        while True:
            try:
                user_input = input(
                    f"{Color.BLUE}ID à IGNORER (séparés par virgules) : {Color.ENDC}"
                )
                if not user_input:
                    break
                selected_for_ignoring = remaining.loc[
                    [int(i.strip()) for i in user_input.split(",")]
                ]
                break
            except:
                print_error("ID invalide.")

    return selected_for_processing, selected_for_ignoring


# ==============================================================================
# SÉCURITÉ ET ATTRIBUTS AD (MFA / SIGNATURES)
# ==============================================================================


def check_and_update_signature(
    user_email: str, conn: Connection, signatures_data: dict
) -> bool:
    if not conn or not conn.bound:
        return False
    try:
        conn.search(
            search_base=BASE_DN,
            search_filter=f"(mail={user_email})",
            attributes=[
                "distinguishedName",
                "extensionAttribute2",
                "extensionAttribute3",
            ],
        )
        if not conn.entries:
            return False

        user_entry = conn.entries[0]
        attr2 = (
            user_entry.extensionAttribute2.value
            if "extensionAttribute2" in user_entry
            else None
        )
        attr3 = (
            user_entry.extensionAttribute3.value
            if "extensionAttribute3" in user_entry
            else None
        )

        if attr2 or attr3:
            print_warning(f"Signature déjà configurée (Attr2: {attr2}, Attr3: {attr3})")
            if (
                input(f"{Color.BLUE}Écraser les valeurs ? (o/n) : {Color.ENDC}").lower()
                != "o"
            ):
                return True

        sig_data = [
            {"N°": s["Numero"], "Modèle": s["Signature"]} for s in signatures_data
        ]
        print("\n" + pd.DataFrame(sig_data).to_markdown(index=False))

        choix_sig = int(input(f"\n{Color.WARNING}Numéro de signature : {Color.ENDC}"))
        selected = next((s for s in signatures_data if s["Numero"] == choix_sig), None)

        if selected:
            modifs = {
                "extensionAttribute2": [
                    (MODIFY_REPLACE, [selected["extensionAttribute2"]])
                ],
                "extensionAttribute3": [
                    (MODIFY_REPLACE, [selected["extensionAttribute3"]])
                ],
            }
            conn.modify(user_entry.entry_dn, modifs)
            print_success("Attributs Letsign-it mis à jour dans l'AD.")
            return True
        return False
    except:
        return False


async def check_and_add_to_mfa_group(
    user_email: str, azure_user_id: str, graph_client: GraphServiceClient
) -> bool:
    MFA_GROUP_ID = GROUP_IDS["mfa"]
    try:
        print_info(f"Analyse état MFA pour {user_email}...")

        # Vérification des méthodes enregistrées
        try:
            phones = await graph_client.users.by_user_id(
                azure_user_id
            ).authentication.phone_methods.get()
            auth_apps = await graph_client.users.by_user_id(
                azure_user_id
            ).authentication.microsoft_authenticator_methods.get()
            has_mfa = (phones and phones.value) or (auth_apps and auth_apps.value)
        except:
            has_mfa = False

        if has_mfa:
            print_success("MFA déjà configurée par l'utilisateur.")
            return True

        print_warning("Aucune méthode MFA active.")
        if (
            input(
                f"{Color.WARNING}❓ Forcer la MFA via le groupe de sécurité ? (o/n) : {Color.ENDC}"
            ).lower()
            == "o"
        ):
            try:
                # Utilisation de ReferenceCreate pour pointer vers l'objet utilisateur
                new_member = ReferenceCreate(
                    odata_id=f"https://graph.microsoft.com/v1.0/directoryObjects/{azure_user_id}"
                )
                await graph_client.groups.by_group_id(MFA_GROUP_ID).members.ref.post(
                    new_member
                )
                print_success("Utilisateur ajouté au groupe de sécurité MFA.")
            except APIError as e:
                if "already exist" in str(e).lower():
                    print_info("Déjà membre du groupe MFA.")
                else:
                    raise e
            return True
        return True
    except Exception as e:
        print_error(f"Erreur gestion MFA : {e}")
        return False


# ==============================================================================
# LOGIQUE MÉTIER ET ORCHESTRATION
# ==============================================================================


async def process_user(
    user_data, graph_client, license_details, signatures_data, adresses_data, ad_conn
):
    email = user_data["email"]
    print_header(f"TRAITEMENT : {email}")

    try:
        # Récupération profil Entra ID
        q_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            filter=f"mail eq '{email}' or userPrincipalName eq '{email}'",
            select=["id", "displayName", "usageLocation"],
        )
        res = await graph_client.users.get(
            request_configuration=UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
                query_parameters=q_params
            )
        )

        if not res or not res.value:
            print_warning("Utilisateur introuvable dans Entra ID.")
            return False

        azure_user = res.value[0]
        if not azure_user.usage_location:
            await graph_client.users.by_user_id(azure_user.id).patch(
                User(usage_location="FR")
            )

        # Attribution Licences
        if (
            input(
                f"{Color.WARNING}❓ Gérer les licences pour {email} ? (o/n): {Color.ENDC}"
            ).lower()
            == "o"
        ):
            packs = {
                "1": {
                    "name": "M365 Business Premium",
                    "skus": ["SPB"],
                    "groups": [GROUP_IDS["business_premium"]],
                },
                "2": {
                    "name": "Exchange Online P2",
                    "skus": ["EXCHANGEENTERPRISE"],
                    "groups": [GROUP_IDS["exchange_plan_2"]],
                },
                "3": {
                    "name": "M365 E3 (sans Teams)",
                    "skus": ["O365_w/o Teams Bundle_M3"],
                    "groups": [GROUP_IDS["m365_e3_no_teams"]],
                },
                "4": {
                    "name": "Microsoft Teams EEA",
                    "skus": ["Microsoft_Teams_EEA_New"],
                    "groups": [GROUP_IDS["only_teams"]],
                },
                "5": {
                    "name": "M365 E3 + Teams",
                    "skus": ["O365_w/o Teams Bundle_M3", "Microsoft_Teams_EEA_New"],
                    "groups": [GROUP_IDS["m365_e3_no_teams"], GROUP_IDS["only_teams"]],
                },
            }

            p_table = []
            for k, p in packs.items():
                stock = min(
                    [license_details.get(s, {}).get("remaining", 0) for s in p["skus"]]
                    or [0]
                )
                c = (
                    Color.GREEN
                    if stock > 5
                    else (Color.WARNING if stock > 0 else Color.FAIL)
                )
                p_table.append(
                    {"ID": k, "Pack": p["name"], "Stock": f"{c}{stock}{Color.ENDC}"}
                )

            print("\n" + pd.DataFrame(p_table).to_markdown(index=False))

            choix = input(
                f"\n{Color.WARNING}Choix (ex: 1,4 ou 'n') : {Color.ENDC}"
            ).lower()
            if choix != "n":
                all_groups = {GROUP_IDS["vpn"]}  # VPN inclus par défaut
                for c in [i.strip() for i in choix.split(",")]:
                    if c in packs:
                        all_groups.update(packs[c]["groups"])

                for gid in all_groups:
                    try:
                        new_member = ReferenceCreate(
                            odata_id=f"https://graph.microsoft.com/v1.0/directoryObjects/{azure_user.id}"
                        )
                        await graph_client.groups.by_group_id(gid).members.ref.post(
                            new_member
                        )
                    except APIError as e:
                        if "already exist" not in str(e).lower():
                            raise e

                print_success("Licences et VPN configurés via groupes.")

                # Régionalisation boîte aux lettres (Retry logic car la BAL peut mettre du temps à se provisionner)
                v_key = str(user_data.get("ville", "")).lower().replace(" ", "_")
                agence = adresses_data.get(v_key, {})
                if agence.get("TimeZone"):
                    for attempt in range(10):
                        try:
                            if attempt > 0:
                                await asyncio.sleep(60)
                            print_info(f"Config TimeZone ({attempt + 1}/10)...")
                            settings = MailboxSettings(
                                time_zone=agence["TimeZone"],
                                language=LocaleInfo(locale="fr-FR"),
                            )
                            await graph_client.users.by_user_id(
                                azure_user.id
                            ).mailbox_settings.patch(body=settings)
                            print_success("Fuseau horaire OK.")
                            break
                        except APIError:
                            if attempt == 9:
                                print_warning(
                                    "Retry épuisés. La BAL n'est peut-être pas encore prête."
                                )

        # MFA et Signatures
        if (
            input(f"{Color.WARNING}❓ Vérifier la MFA ? (o/n): {Color.ENDC}").lower()
            == "o"
        ):
            await check_and_add_to_mfa_group(email, azure_user.id, graph_client)

        if (
            input(
                f"{Color.WARNING}❓ Configurer la signature Letsign-it ? (o/n): {Color.ENDC}"
            ).lower()
            == "o"
        ):
            check_and_update_signature(email, ad_conn, signatures_data)

        return True
    except Exception as e:
        print_error(f"Erreur durant le traitement : {e}")
        return False


def write_df_to_sharepoint(ctx, df):
    try:
        buf = io.BytesIO()
        df.to_excel(buf, index=False, engine="openpyxl")
        target = ctx.web.get_folder_by_server_relative_url(
            os.path.dirname(SHAREPOINT_FILE_URL)
        )
        target.upload_file(
            os.path.basename(SHAREPOINT_FILE_URL), buf.getvalue()
        ).execute_query()
        print_success("Registre SharePoint mis à jour.")
    except Exception as e:
        print_error(f"Erreur sauvegarde SharePoint : {e}")


async def run(
    ctx: ClientContext,
    kp,
    tenant_id: str,
    ad_conn: Connection,
    config_dir: str,
    **kwargs,
):
    """
    Point d'entrée synchronisé avec l'orchestrateur.
    kp : session KeePass déjà ouverte transmise par le HUB.
    """
    print_header("SSO : AUTHENTIFICATION MICROSOFT GRAPH")
    try:
        entry = kp.find_entries(title="Azure App Credentials", first=True)
        if not entry:
            raise ValueError("Identifiants Azure introuvables dans le KeePass.")

        credential = ClientSecretCredential(tenant_id, entry.username, entry.password)
        graph_client = GraphServiceClient(
            credentials=credential, scopes=["https://graph.microsoft.com/.default"]
        )
        print_success("Liaison Microsoft Graph établie via SSO.")
    except Exception as e:
        print_error(f"Échec initialisation Graph : {e}")
        return

    license_details = await get_license_details(graph_client)
    await display_license_status(license_details)

    if (
        input(
            f"\n{Color.BLUE}Continuer vers la sélection des profils (o/n) ? {Color.ENDC}"
        ).lower()
        != "o"
    ):
        return

    signatures_data, adresses_data = load_configs_from_local(config_dir)
    excel_content = await download_excel_from_sharepoint(ctx, SHAREPOINT_FILE_URL)

    if not all([signatures_data, adresses_data, excel_content]):
        return

    full_df = pd.read_excel(io.BytesIO(excel_content))
    full_df.reset_index(drop=True, inplace=True)
    full_df["ID"] = full_df.index + 1
    full_df.set_index("ID", inplace=True)

    # Filtrage des comptes AD créés mais non licenciés
    users_to_process = full_df[full_df["Etat"] == "A traiter"].copy()
    if users_to_process.empty:
        print_success("Aucun profil 'A traiter' trouvé dans le registre.")
        return

    to_p, to_i = select_users_from_list(users_to_process)

    # Mise à jour des profils ignorés
    if not to_i.empty:
        for idx in to_i.index:
            full_df.loc[idx, "Etat"] = "Ignoré"
        await asyncio.to_thread(write_df_to_sharepoint, ctx, full_df.reset_index())

    # Traitement des profils sélectionnés
    if not to_p.empty:
        for idx, row in to_p.iterrows():
            if await process_user(
                row,
                graph_client,
                license_details,
                signatures_data,
                adresses_data,
                ad_conn,
            ):
                full_df.loc[idx, "Etat"] = "Traité"

        # Sauvegarde finale
        await asyncio.to_thread(write_df_to_sharepoint, ctx, full_df.reset_index())


if __name__ == "__main__":
    print_error("Ce module doit être piloté par 'hub_central.py'.")
