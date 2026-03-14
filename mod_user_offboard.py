# -*- coding: utf-8 -*-

"""
Auteur : JEL pour Groupe Entis
Rôle   : Gestion complète du processus de départ (Leaver).
         Recherche multicritère (Nom, Prénom, Login), désactivation AD,
         archivage OU, nettoyage licences Entra et suppression matériel Intune & Entra ID.

Version : 1.3.4
Date    : 26/12/2024

Historique :
-----------
v1.3.4 (26/12/2024) : Amélioration recherche croisée (AD/AZ) et ajout traçabilité dans la description AD.
v1.3.3 (26/12/2024) : Correction définitive suppression matériel (getattr pour azureADDeviceId).
v1.3.0 (26/12/2024) : Recherche multicritère interactive (AD & Azure).
"""

import os
import sys
import getpass
import asyncio
import re
from datetime import datetime, timedelta

# Bibliothèques tierces
from pykeepass import PyKeePass
from pykeepass.exceptions import CredentialsError
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.users_request_builder import UsersRequestBuilder
from msgraph.generated.device_management.managed_devices.managed_devices_request_builder import (
    ManagedDevicesRequestBuilder,
)
from msgraph.generated.devices.devices_request_builder import DevicesRequestBuilder
from ldap3 import Server, Connection, ALL, SIMPLE, MODIFY_REPLACE
from ldap3.utils.dn import parse_dn

# ==============================================================================
# CONFIGURATION ET CONSTANTES
# ==============================================================================

BASE_DN = "DC=cetremut,DC=pri"
TARGET_DISABLED_OU = "OU=User,OU=__DESACTIVE__,DC=cetremut,DC=pri"
AUTHORITY_URL = "https://login.microsoftonline.com/"
GRAPH_API_SCOPE = ["https://graph.microsoft.com/.default"]

LICENSE_GROUPS = {
    "cd1bcd92-5fc5-4904-bfdc-a42a3f8cb3e9": "M365 Business Premium",
    "48b9abe9-db9d-4f58-881e-22d2cd421d8c": "Exchange Plan 1",
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
    FAIL = "\033[91m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"


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
# OUTILS DE CONVERSION ET RECHERCHE
# ==============================================================================


def filetime_to_dt(ft):
    if ft is None:
        return "Jamais"
    if hasattr(ft, "strftime"):
        try:
            return ft.strftime("%d/%m/%Y") if 1601 < ft.year < 9000 else "Jamais"
        except:
            return "Jamais"
    try:
        val = int(ft)
        if val == 0 or val >= 9223372036854775807:
            return "Jamais"
        return (datetime(1601, 1, 1) + timedelta(microseconds=val // 10)).strftime(
            "%d/%m/%Y"
        )
    except:
        return "Inconnue"


async def get_intune_devices(graph_client, az_user_id):
    """Récupère les équipements gérés via Microsoft Intune."""
    try:
        res = await graph_client.users.by_user_id(az_user_id).managed_devices.get()
        devices = res.value if res and res.value else []
        pcs = [
            d
            for d in devices
            if d.operating_system and d.operating_system.lower() in ["windows", "macos"]
        ]
        mobiles = [
            d
            for d in devices
            if d.operating_system
            and d.operating_system.lower() in ["ios", "android", "ipados"]
        ]
        return pcs, mobiles
    except:
        return [], []


def find_managed_computers_ad(ad_conn, user_dn):
    """Recherche les PC dont l'utilisateur est responsable dans l'AD."""
    try:
        ad_conn.search(
            BASE_DN,
            f"(&(objectClass=computer)(managedBy={user_dn}))",
            attributes=["cn"],
        )
        return [entry.cn.value for entry in ad_conn.entries]
    except:
        return []


async def search_user_globally(term, ad_conn, graph_client):
    """Effectue une recherche croisée AD et Azure Entra ID plus robuste."""
    results = {}

    # 1. Recherche AD Local (Base de la recherche)
    filter_ad = f"(|(sAMAccountName=*{term}*)(displayName=*{term}*)(mail=*{term}*)(givenName=*{term}*)(sn=*{term}*))"
    ad_conn.search(
        BASE_DN,
        filter_ad,
        attributes=[
            "displayName",
            "mail",
            "userPrincipalName",
            "sAMAccountName",
            "distinguishedName",
            "accountExpires",
            "userAccountControl",
        ],
    )

    for entry in ad_conn.entries:
        # On essaie de déterminer l'email le plus probable
        email_cand = (
            entry.mail.value if entry.mail.value else entry.userPrincipalName.value
        )
        email_key = (
            email_cand.lower()
            if email_cand
            else str(entry.sAMAccountName.value).lower()
        )

        results[email_key] = {
            "email": email_cand if email_cand else "N/A",
            "name": str(entry.displayName.value)
            if entry.displayName.value
            else str(entry.sAMAccountName.value),
            "ad_entry": entry,
            "az_user": None,
            "in_ad": True,
            "in_az": False,
        }

    # 2. Pour chaque résultat AD, on vérifie spécifiquement dans Azure
    for email_key, data in results.items():
        if data["email"] == "N/A":
            continue
        try:
            # On cherche par mail exact ou UPN exact
            q_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
                filter=f"mail eq '{data['email']}' or userPrincipalName eq '{data['email']}'",
                select=["id", "displayName", "mail", "userPrincipalName"],
            )
            config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
                query_parameters=q_params
            )
            az_res = await graph_client.users.get(request_configuration=config)

            if az_res and az_res.value:
                data["az_user"] = az_res.value[0]
                data["in_az"] = True
        except:
            pass

    # 3. Recherche Azure complémentaire (pour les comptes uniquement Cloud qui matchent le terme)
    try:
        q_params_az = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            filter=f"startswith(displayName, '{term}') or startswith(mail, '{term}') or startswith(userPrincipalName, '{term}')",
            select=["id", "displayName", "mail", "userPrincipalName"],
        )
        config_az = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=q_params_az
        )
        az_extra = await graph_client.users.get(request_configuration=config_az)

        if az_extra and az_extra.value:
            for u in az_extra.value:
                email = (u.mail if u.mail else u.user_principal_name).lower()
                if email not in results:
                    results[email] = {
                        "email": email,
                        "name": u.display_name,
                        "ad_entry": None,
                        "az_user": u,
                        "in_ad": False,
                        "in_az": True,
                    }
    except:
        pass

    return list(results.values())


# ==============================================================================
# LOGIQUE DE SORTIE (LEAVER PROCESS)
# ==============================================================================


async def run(ctx, kp, tenant_id, ad_conn, config_dir, **kwargs):
    print_header("PROCESSUS DE SORTIE COLLABORATEUR")

    # 1. AUTHENTIFICATION GRAPH VIA SSO
    try:
        entry = kp.find_entries(title="Azure App Credentials", first=True)
        if not entry:
            entry = kp.find_entries(title="Azure", first=True)
        if not entry:
            raise ValueError("Secrets Azure introuvables.")

        credential = ClientSecretCredential(tenant_id, entry.username, entry.password)
        graph_client = GraphServiceClient(
            credentials=credential, scopes=GRAPH_API_SCOPE
        )
    except Exception as e:
        print_error(f"Échec initialisation Graph : {e}")
        return

    # 2. BOUCLE DE RECHERCHE UTILISATEUR
    selected_profile = None
    while not selected_profile:
        term = input(
            f"\n{Color.BLUE}Recherche collaborateur (Nom, Prénom ou Login) : {Color.ENDC}"
        ).strip()
        if not term:
            return

        print_info(f"Interrogation des annuaires pour '{term}'...")
        candidates = await search_user_globally(term, ad_conn, graph_client)

        if not candidates:
            print_warning("Aucun utilisateur trouvé. Recommencer ?")
            continue

        print(
            f"\n{Color.BOLD}{'#':<3} | {'Nom Complet':<30} | {'Email':<35} | {'Status'}{Color.ENDC}"
        )
        print("-" * 85)
        for i, c in enumerate(candidates):
            status = (
                f"[{'AD' if c['in_ad'] else '--'}] [{'AZ' if c['in_az'] else '--'}]"
            )
            print(
                f"{i + 1:<3} | {str(c['name'])[:30]:<30} | {str(c['email'])[:35]:<35} | {status}"
            )

        choice = (
            input(
                f"\n{Color.BLUE}Sélectionnez un numéro (ou R pour relancer) : {Color.ENDC}"
            )
            .strip()
            .upper()
        )
        if choice == "R":
            continue
        if choice.isdigit() and 0 < int(choice) <= len(candidates):
            selected_profile = candidates[int(choice) - 1]
        else:
            print_error("Sélection invalide.")

    # 3. COLLECTE DES DÉTAILS
    report = []
    email = selected_profile["email"]
    nom_complet = selected_profile["name"].upper()

    pcs_ad = (
        find_managed_computers_ad(ad_conn, selected_profile["ad_entry"].entry_dn)
        if selected_profile["in_ad"]
        else []
    )
    pcs_int, mob_int = ([], [])
    if selected_profile["in_az"]:
        pcs_int, mob_int = await get_intune_devices(
            graph_client, selected_profile["az_user"].id
        )

    # 4. RÉCAPITULATIF
    print_header("PROFIL POUR SORTIE DÉFINITIVE")
    print(f" {Color.BOLD}Utilisateur     :{Color.ENDC} {nom_complet}")
    print(f" {Color.BOLD}Email principal :{Color.ENDC} {email}")

    print(f"\n {Color.CYAN}[ ÉQUIPEMENTS DÉTECTÉS ]{Color.ENDC}")
    all_pcs = set(pcs_ad + [p.device_name for p in pcs_int if p.device_name])
    if not (all_pcs or mob_int):
        print(" - Aucun matériel détecté.")
    else:
        for p in all_pcs:
            print(f" - Ordinateur : {p}")
        for m in mob_int:
            print(f" - Mobile     : {m.device_name} ({m.model})")

    if (
        input(
            f"\n{Color.WARNING}Confirmer le lancement de la procédure ? (o/n) : {Color.ENDC}"
        ).lower()
        != "o"
    ):
        print_warning("Opération annulée.")
        return

    # 5. EXÉCUTION

    # 5.1. ACTIVE DIRECTORY
    if selected_profile["in_ad"]:
        print_header("ACTION 1 : ACTIVE DIRECTORY")
        if (
            input(
                f"{Color.BLUE}Désactiver et archiver le compte AD ? (o/n) : {Color.ENDC}"
            ).lower()
            == "o"
        ):
            u_ad = selected_profile["ad_entry"]

            # --- Traçabilité de la désactivation ---
            now = datetime.now()
            timestamp = now.strftime("%d/%m/%Y à %H:%M")
            # On récupère le nom de l'admin à partir de la connexion AD
            admin_raw = ad_conn.user if hasattr(ad_conn, "user") else "Admin"
            admin_name = admin_raw.split("@")[0].replace("admin.", "").upper()

            desc_str = f"Compte désactivé par {admin_name} le {timestamp}"

            # Mise à jour UAC et Description
            ad_conn.modify(
                u_ad.entry_dn,
                {
                    "userAccountControl": [
                        (MODIFY_REPLACE, [str(int(u_ad.userAccountControl.value) | 2)])
                    ],
                    "description": [(MODIFY_REPLACE, [desc_str])],
                },
            )

            # Déplacement d'OU
            parsed = parse_dn(u_ad.entry_dn)
            rdn = f"{parsed[0][0]}={parsed[0][1]}"
            if TARGET_DISABLED_OU.lower() not in u_ad.entry_dn.lower():
                ad_conn.modify_dn(u_ad.entry_dn, rdn, new_superior=TARGET_DISABLED_OU)
                print_success(
                    "Compte désactivé, descriptif mis à jour et objet déplacé."
                )
            else:
                print_info(
                    "Compte désactivé et descriptif mis à jour (déjà dans l'OU cible)."
                )

            report.append(f"✅ AD Local : Désactivé (Trace: {desc_str}).")

    # 5.2. ENTRA ID (Licences)
    if selected_profile["in_az"]:
        print_header("ACTION 2 : LICENCES ENTRA ID")
        if (
            input(
                f"{Color.BLUE}Retirer les licences M365 Business & Exchange ? (o/n) : {Color.ENDC}"
            ).lower()
            == "o"
        ):
            for g_id, g_name in LICENSE_GROUPS.items():
                try:
                    await (
                        graph_client.groups.by_group_id(g_id)
                        .members.by_directory_object_id(selected_profile["az_user"].id)
                        .ref.delete()
                    )
                    print_info(f"Retrait {g_name} OK.")
                except:
                    pass
            report.append("✅ Entra ID : Nettoyage des groupes de licences.")

    # 5.3. MATÉRIEL (Purge triple couche)
    if all_pcs or mob_int:
        print_header("ACTION 3 : SUPPRESSION DU MATÉRIEL")
        for pc in all_pcs:
            if (
                input(
                    f"{Color.FAIL}Supprimer le poste {pc} (AD/Intune/Azure) ? (o/n) : {Color.ENDC}"
                ).lower()
                == "o"
            ):
                # --- A. AD LOCAL ---
                ad_conn.search(BASE_DN, f"(&(objectClass=computer)(cn={pc}))")
                for entry in ad_conn.entries:
                    ad_conn.delete(entry.entry_dn)

                # --- B. INTUNE ---
                try:
                    q_params = ManagedDevicesRequestBuilder.ManagedDevicesRequestBuilderGetQueryParameters(
                        filter=f"deviceName eq '{pc}'",
                        select=["id", "deviceName", "azureADDeviceId"],
                    )
                    cfg = ManagedDevicesRequestBuilder.ManagedDevicesRequestBuilderGetRequestConfiguration(
                        query_parameters=q_params
                    )
                    dev_res = await graph_client.device_management.managed_devices.get(
                        request_configuration=cfg
                    )

                    if dev_res and dev_res.value:
                        device = dev_res.value[0]
                        azure_id = getattr(
                            device, "azure_a_d_device_id", None
                        ) or getattr(device, "azure_ad_device_id", None)

                        await graph_client.device_management.managed_devices.by_managed_device_id(
                            device.id
                        ).delete()
                        print_success(f"Poste {pc} supprimé d'Intune.")

                        # --- C. ENTRA ID ---
                        if azure_id:
                            dev_entra_params = DevicesRequestBuilder.DevicesRequestBuilderGetQueryParameters(
                                filter=f"deviceId eq '{azure_id}'"
                            )
                            dev_entra_cfg = DevicesRequestBuilder.DevicesRequestBuilderGetRequestConfiguration(
                                query_parameters=dev_entra_params
                            )
                            entra_res = await graph_client.devices.get(
                                request_configuration=dev_entra_cfg
                            )

                            if entra_res and entra_res.value:
                                await graph_client.devices.by_device_id(
                                    entra_res.value[0].id
                                ).delete()
                                print_success(f"Objet Azure purgé.")
                except Exception as e:
                    print_warning(f"Erreur purge Cloud {pc} : {e}")

                report.append(f"✅ Matériel : PC {pc} purgé partout.")

        for m in mob_int:
            if (
                input(
                    f"{Color.FAIL}Supprimer le mobile {m.device_name} d'Intune ? (o/n) : {Color.ENDC}"
                ).lower()
                == "o"
            ):
                try:
                    await graph_client.device_management.managed_devices.by_managed_device_id(
                        m.id
                    ).delete()
                    print_success("Mobile supprimé.")
                    report.append(f"✅ Matériel : Mobile {m.device_name} supprimé.")
                except Exception as e:
                    print_error(f"Échec suppression mobile : {e}")

    # 6. BILAN FINAL
    print_header("RÉCAPITULATIF FINAL")
    for line in report:
        print(f" {line}")
    print_success(f"\nDossier de {nom_complet} traité.")


if __name__ == "__main__":
    print_error("Ce module doit être piloté par 'hub_central.py'.")
