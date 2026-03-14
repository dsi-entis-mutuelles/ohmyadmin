# -*- coding: utf-8 -*-

"""
Synopsis:
    Création automatisée d'utilisateurs Active Directory pour Entis Mutuelles, orchestrée via un hub centralisé.
    Version complète intégrant la logique métier exhaustive et la correction
    de recherche des groupes d'agence dans la structure JSON 'par_ville'.

Processus:
    1. Chargement des configurations (Agences, Groupes, Services).
    2. Lecture SharePoint (Nouveaux arrivants).
    3. Calcul dynamique des identifiants (Contrôle homonymes AD).
    4. Attribution des groupes (Base + Agence via 'par_ville' + Service + VIP).
    5. Gestion des tickets de matériel/téléphonie.
    6. Validation visuelle et écriture SharePoint (Suivi).
    7. Création AD (LDAPS).

Auteur: JEL pour Entis Mutuelles
Version: 2.2.1 (Fix signature run() pour support SSO KeePass via orchestrateur)
"""

import os
import io
import json
import pandas as pd
import secrets
import string
import time
from datetime import datetime
from unidecode import unidecode
from ldap3 import Server, Connection, ALL, SIMPLE, MODIFY_ADD, MODIFY_REPLACE
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.client_request_exception import ClientRequestException
import re


def normalize_service_name(service_name: str) -> str:
    """
    Normalise un nom pour une comparaison fiable.
    Met en minuscules et remplace les séparateurs multiples par un seul underscore.
    """
    if not isinstance(service_name, str):
        return ""
    normalized = unidecode(service_name).lower()
    normalized = re.sub(r"[\s_-]+", "_", normalized)
    return normalized.strip("_")


# --- CONFIGURATION SPÉCIFIQUE AU MODULE ---
BASE_DN = "DC=cetremut,DC=pri"


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


# --- Fonctions utilitaires ---
def remove_special_characters(text):
    if not isinstance(text, str):
        return ""
    return unidecode(text).replace("'", "").replace(" ", "")


def format_name_for_display(name_part):
    return " ".join(word.capitalize() for word in str(name_part).split())


def format_prenom_for_tech(prenom_brut):
    s = unidecode(str(prenom_brut)).strip()
    s = s.replace("_", " ").replace("-", " ")
    parts = [p for p in s.split() if p]
    return "-".join(p.capitalize() for p in parts)


def get_initials(prenom_technique):
    parts = [p for p in str(prenom_technique).split("-") if p]
    return "".join(p[0] for p in parts) if parts else ""


def generate_random_password():
    letters, digits, special_chars = string.ascii_letters, string.digits, "@$!"
    password_list = [
        secrets.choice(letters),
        secrets.choice(digits),
        secrets.choice(special_chars),
        secrets.choice(special_chars),
    ]
    full_charset = letters + digits + special_chars
    password_list += [secrets.choice(full_charset) for _ in range(12)]
    secrets.SystemRandom().shuffle(password_list)
    return "".join(password_list)


def clean_attrs(d: dict) -> dict:
    if not isinstance(d, dict):
        return d
    return {
        k: v
        for k, v in d.items()
        if isinstance(v, (int, bytes)) or (isinstance(v, str) and v.strip())
    }


def get_value_from_possible_keys(data_row, keys, default=""):
    """Parcourt une liste de clés possibles et retourne la valeur de la première clé trouvée."""
    for key in keys:
        if key in data_row and pd.notna(data_row[key]):
            return data_row[key]
    return default


# --- Fonctions de Connexion et de Lecture de Données ---
def load_configs_from_local(config_dir: str):
    print_info("Chargement des fichiers de configuration locaux...")
    try:
        with open(
            os.path.join(config_dir, "Adresse_Agences.json"), "r", encoding="utf-8-sig"
        ) as f:
            lieux = json.load(f)
        with open(
            os.path.join(config_dir, "GD_Agence.json"), "r", encoding="utf-8-sig"
        ) as f:
            groupes_ville = json.load(f)
        with open(
            os.path.join(config_dir, "Groupe_service.json"), "r", encoding="utf-8-sig"
        ) as f:
            groupes_service = json.load(f)
        print_success("Tous les fichiers de configuration ont été chargés.")
        return lieux, groupes_ville, groupes_service
    except Exception as e:
        print_error(f"Erreur lors du chargement des configurations : {e}")
        return None, None, None


def read_arrivants_from_sharepoint(ctx_arrivants):
    if not ctx_arrivants:
        return None
    print_info("Lecture du fichier des nouveaux arrivants...")
    try:
        file_url = "/personal/cs_automate_entis_onmicrosoft_com/Documents/Formulaire nouvel arrivant.xlsx"
        response = File.open_binary(ctx_arrivants, file_url)
        df = pd.read_excel(io.BytesIO(response.content)).fillna("")
        if "ID" not in df.columns:
            df["ID"] = range(1, len(df) + 1)
        print_success("Fichier des nouveaux arrivants lu avec succès.")
        return df
    except Exception as e:
        print_error(f"Erreur lors de la lecture du fichier des arrivants : {e}")
        return None


def write_users_to_sharepoint_excel(users_data, ctx_suivi):
    if not users_data:
        return True
    if not ctx_suivi:
        return False
    print_header("PHASE 2 : ÉCRITURE DANS LE FICHIER DE SUIVI (OBLIGATOIRE)")
    while True:
        try:
            new_users_df = pd.DataFrame(users_data)
            new_users_df["Etat"] = "A traiter"
            new_users_df = new_users_df[
                ["prenom", "nom", "email", "ville", "service", "Etat", "materiel"]
            ]

            suivi_file_url = "/sites/GLPI/Data/new_users.xlsx"

            try:
                response = File.open_binary(ctx_suivi, suivi_file_url)
                existing_df = pd.read_excel(io.BytesIO(response.content))
                updated_df = pd.concat([existing_df, new_users_df], ignore_index=True)
            except Exception:
                print_warning(
                    "Fichier de suivi non trouvé ou vide. Un nouveau fichier sera créé."
                )
                updated_df = new_users_df

            output_buffer = io.BytesIO()
            updated_df.to_excel(output_buffer, index=False, engine="openpyxl")
            target_folder = ctx_suivi.web.get_folder_by_server_relative_url(
                os.path.dirname(suivi_file_url)
            )
            target_folder.upload_file(
                os.path.basename(suivi_file_url), output_buffer.getvalue()
            ).execute_query()
            print_success("Fichier de suivi Excel mis à jour avec succès.")
            return True
        except ClientRequestException as e:
            if "locked for exclusive use" in str(
                e
            ) or "Le fichier est verrouillé" in str(e):
                locked_by_user = "un autre utilisateur"
                try:
                    file = ctx_suivi.web.get_file_by_server_relative_url(suivi_file_url)
                    file.select(["LockedByUser/Title"]).expand(
                        ["LockedByUser"]
                    ).get().execute_query()
                    if file.locked_by_user:
                        locked_by_user = file.locked_by_user.title
                except Exception:
                    pass
                print_error(
                    f"Le fichier Excel de suivi est verrouillé par : {locked_by_user}."
                )
                print_warning(
                    "L'écriture dans ce fichier est OBLIGATOIRE avant la création des comptes."
                )
                choice = input(
                    f"{Color.BLUE}Voulez-vous réessayer (o) ou ABANDONNER (n) ? : {Color.ENDC}"
                ).lower()
                if choice == "o":
                    print_info("Nouvelle tentative dans 5 secondes...")
                    time.sleep(5)
                    continue
                else:
                    print_error(
                        "Opération abandonnée par l'utilisateur. Aucun compte ne sera créé."
                    )
                    return False
            else:
                print_error(
                    f"Erreur inattendue lors de l'écriture sur SharePoint : {e}"
                )
                return False
        except Exception as e:
            print_error(
                f"Erreur critique lors de la préparation du fichier Excel : {e}"
            )
            return False


def display_and_select_users(df):
    display_df = df.copy().rename(
        columns={
            "Heure de début": "Date",
            "Nom": "Demandeur",
            "Nom1": "Nom du salarié",
            "Prénom": "Prénom du salarié",
        }
    )
    if "Date" in display_df.columns:
        display_df["Date"] = pd.to_datetime(
            display_df["Date"], errors="coerce"
        ).dt.strftime("%d-%m-%Y")
    display_columns = ["ID", "Demandeur", "Nom du salarié", "Prénom du salarié", "Date"]
    existing_display_columns = [
        col for col in display_columns if col in display_df.columns
    ]
    display_df = display_df[existing_display_columns].sort_values(
        by="ID", ascending=False
    )

    while True:
        try:
            num_lines_input = input(
                f"{Color.BLUE}Combien de demandes voulez-vous afficher ? (ex: 5, 'tout') : {Color.ENDC}"
            )
            num_lines = (
                len(display_df)
                if num_lines_input.lower() == "tout"
                else int(num_lines_input)
            )
            break
        except ValueError:
            print_error("Veuillez entrer un nombre valide ou 'tout'.")

    displayed_part = display_df.head(num_lines)
    print(displayed_part.to_markdown(index=False))

    valid_ids_on_screen = set(displayed_part["ID"])

    while True:
        user_ids_input = input(
            f"{Color.BLUE}Entrez les ID à créer (uniquement depuis la liste ci-dessus, séparés par virgules) : {Color.ENDC}"
        )
        if not user_ids_input:
            print_warning("Aucun ID saisi. Veuillez réessayer.")
            continue
        try:
            entered_ids = {int(id_str.strip()) for id_str in user_ids_input.split(",")}

            valid_selection = entered_ids.intersection(valid_ids_on_screen)
            invalid_selection = entered_ids.difference(valid_ids_on_screen)

            if invalid_selection:
                print_error(
                    f"Les IDs suivants sont invalides ou ne font pas partie de la liste affichée : {sorted(list(invalid_selection))}"
                )
                if not valid_selection:
                    print_warning(
                        "Aucun ID valide n'a été sélectionné. Veuillez réessayer."
                    )
                continue

            print_success(f"IDs valides sélectionnés : {sorted(list(valid_selection))}")
            selected_users = df[df["ID"].isin(valid_selection)].to_dict("records")
            return selected_users

        except ValueError:
            print_error(
                "Entrée invalide. Veuillez n'entrer que des nombres séparés par des virgules."
            )


def prepare_user_creation_plan(user_data, conn, config_files):
    lieux_data, groupes_ville_json, groupes_service = config_files

    # Mapping flexible des colonnes Excel de votre version originale
    nom_brut = get_value_from_possible_keys(
        user_data, ["Nom2", "Nom1", "Nom du salarié", "Nom salarié", "Nom"]
    )
    prenom_brut = get_value_from_possible_keys(
        user_data, ["Prénom", "Prénom du salarié", "Prenom salarié", "Prenom"]
    )
    ville_brute = get_value_from_possible_keys(
        user_data,
        ["Lieu de travail", "Lieu de Travail", "Lieu", "Ville", "Agence", "Site"],
    )

    if not all(
        [str(nom_brut).strip(), str(prenom_brut).strip(), str(ville_brute).strip()]
    ):
        print_error(
            f"Données manquantes (Nom, Prénom ou Lieu) pour l'ID {user_data.get('ID')}. Ignoré."
        )
        return None

    prenom_display = format_name_for_display(prenom_brut)
    nom_display = str(nom_brut).upper()
    name = f"{prenom_display} {nom_display}"

    # 1. HOMONYMES
    print_info(f"Vérification des homonymes pour {name}...")
    search_filter = (
        f"(&(objectClass=user)(givenName={prenom_display})(sn={nom_display}))"
    )
    conn.search(
        BASE_DN,
        search_filter,
        attributes=["sAMAccountName", "mail", "department", "title"],
    )

    if conn.entries:
        print_warning(f"ATTENTION : Un ou plusieurs homonymes trouvés pour {name} !")
        for entry in conn.entries:
            print(f"  - {Color.BOLD}Login:{Color.ENDC} {entry.sAMAccountName.value}")
            print(f"    {Color.BOLD}Email:{Color.ENDC} {entry.mail.value}")
            print(f"    {Color.BOLD}Service:{Color.ENDC} {entry.department.value}")
            print(f"    {Color.BOLD}Fonction:{Color.ENDC} {entry.title.value}\n")

        if (
            input(
                f"{Color.WARNING}Voulez-vous quand même continuer la création ? (o/n) : {Color.ENDC}"
            ).lower()
            != "o"
        ):
            print_error(f"Création de {name} annulée par l'opérateur.")
            return None
        print_info("Continuation de la création malgré la présence d'homonyme(s).")

    # 2. GÉNÉRATION LOGIN UNIQUE
    prenom_tech = format_prenom_for_tech(prenom_brut)
    print_info(f"Recherche d'un login unique pour {name}...")
    nom_login = remove_special_characters(str(nom_brut)).lower()
    sam_account_name = ""
    is_composite = "-" in prenom_tech

    if is_composite:
        initiales = get_initials(prenom_tech).lower()
        temp_sam = f"{nom_login}_{initiales}"
        conn.search(
            BASE_DN, f"(sAMAccountName={temp_sam})", attributes=["sAMAccountName"]
        )
        if not conn.entries:
            sam_account_name = temp_sam
        else:
            second_part = prenom_tech.split("-")[1].lower()
            for i in range(1, len(second_part) + 1):
                temp_sam_progressive = f"{nom_login}_{initiales}{second_part[:i]}"
                conn.search(
                    BASE_DN,
                    f"(sAMAccountName={temp_sam_progressive})",
                    attributes=["sAMAccountName"],
                )
                if not conn.entries:
                    sam_account_name = temp_sam_progressive
                    break
    else:
        prenom_flat = prenom_tech.lower()
        for i in range(1, len(prenom_flat) + 1):
            temp_sam = f"{nom_login}_{prenom_flat[:i]}"
            conn.search(
                BASE_DN, f"(sAMAccountName={temp_sam})", attributes=["sAMAccountName"]
            )
            if not conn.entries:
                sam_account_name = temp_sam
                break

    if not sam_account_name:
        print_warning(
            f"Tous les logins basés sur le prénom sont pris. Tentative avec suffixe numérique."
        )
        base_sam_fallback = (
            f"{nom_login}_{get_initials(prenom_tech).lower()}"
            if is_composite
            else f"{nom_login}_{prenom_tech.lower()}"
        )
        counter = 1
        while True:
            temp_sam = f"{base_sam_fallback}{counter}"
            conn.search(
                BASE_DN, f"(sAMAccountName={temp_sam})", attributes=["sAMAccountName"]
            )
            if not conn.entries:
                sam_account_name = temp_sam
                break
            counter += 1
    print_success(f"Login unique trouvé : {sam_account_name}")

    # 3. GÉNÉRATION EMAIL UNIQUE
    ville_key = unidecode(str(ville_brute)).replace("'", "").replace(" ", "_").lower()
    agence_info = lieux_data.get(ville_key)
    if not agence_info:
        print_error(
            f"Info d'agence introuvable pour '{ville_brute}' (clé: '{ville_key}'). Ignoré."
        )
        return None

    print_info(f"Recherche d'un email unique pour {name}...")
    domaine_mail = (
        "@"
        + str(
            get_value_from_possible_keys(
                user_data, ["Nom de domaine", "Domaine"], "cetremut.pri"
            )
        ).lower()
    )
    nom_for_email = remove_special_characters(str(nom_brut)).lower()
    mail = ""

    if "mgprev.fr" in domaine_mail:
        base_mail = f"{prenom_tech.lower()}.{nom_for_email}"
        temp_mail = f"{base_mail}{domaine_mail}"
        conn.search(BASE_DN, f"(mail={temp_mail})", attributes=["mail"])
        if not conn.entries:
            mail = temp_mail
        else:
            counter = 2
            while True:
                temp_mail = f"{base_mail}{counter}{domaine_mail}"
                conn.search(BASE_DN, f"(mail={temp_mail})", attributes=["mail"])
                if not conn.entries:
                    mail = temp_mail
                    break
                counter += 1
    else:
        if is_composite:
            initiales = get_initials(prenom_tech).lower()
            base_mail = f"{initiales}.{nom_for_email}"
            temp_mail = f"{base_mail}{domaine_mail}"
            conn.search(BASE_DN, f"(mail={temp_mail})", attributes=["mail"])
            if not conn.entries:
                mail = temp_mail
            else:
                second_part = prenom_tech.split("-")[1].lower()
                for i in range(1, len(second_part) + 1):
                    base_mail_progressive = (
                        f"{initiales}{second_part[:i]}.{nom_for_email}"
                    )
                    temp_mail = f"{base_mail_progressive}{domaine_mail}"
                    conn.search(BASE_DN, f"(mail={temp_mail})", attributes=["mail"])
                    if not conn.entries:
                        mail = temp_mail
                        break
        else:
            prenom_flat_for_email = prenom_tech.lower()
            for i in range(1, len(prenom_flat_for_email) + 1):
                base_mail = f"{prenom_flat_for_email[:i]}.{nom_for_email}"
                temp_mail = f"{base_mail}{domaine_mail}"
                conn.search(BASE_DN, f"(mail={temp_mail})", attributes=["mail"])
                if not conn.entries:
                    mail = temp_mail
                    break

        if not mail:
            base_mail_fallback = (
                f"{get_initials(prenom_tech).lower()}.{nom_for_email}"
                if is_composite
                else f"{prenom_tech.lower()}.{nom_for_email}"
            )
            counter = 2
            while True:
                temp_mail = f"{base_mail_fallback}{counter}{domaine_mail}"
                conn.search(BASE_DN, f"(mail={temp_mail})", attributes=["mail"])
                if not conn.entries:
                    mail = temp_mail
                    break
                counter += 1
    print_success(f"Email unique trouvé : {mail}")

    # 4. GROUPES AD ET ORGANISATION
    user_dn = f"CN={name},OU=User,OU=Account,OU={agence_info['OU']},OU=ENTIS,{BASE_DN}"
    generated_password = generate_random_password()

    computer_type_raw = str(
        get_value_from_possible_keys(user_data, ["Ordinateur", "Matériel"])
    ).lower()
    groupes_a_ajouter = ["g_entis_users", "groupe_cetremut", "G_Entis_TSGW_Sante"]

    if (
        "par_ville" in groupes_ville_json
        and ville_key in groupes_ville_json["par_ville"]
    ):
        print_info(f"Attribution des groupes d'agence pour '{ville_key}'...")
        agence_cfg = groupes_ville_json["par_ville"][ville_key]
        if "groupes" in agence_cfg:
            groupes_a_ajouter.extend([g for g in agence_cfg["groupes"] if g.strip()])
    else:
        print_warning(
            f"Clé '{ville_key}' non trouvée dans la section 'par_ville' de GD_Agence.json."
        )

    # RECHERCHE DE SERVICE
    service_value_excel = str(get_value_from_possible_keys(user_data, ["Service"]))
    normalized_service_to_find = normalize_service_name(service_value_excel)

    found_service = False
    if normalized_service_to_find:
        for service_entry in groupes_service:
            current_service_key_normalized = normalize_service_name(
                service_entry.get("Service", "")
            )
            if current_service_key_normalized == normalized_service_to_find:
                groupes_str = service_entry.get("Groupe", "")
                if groupes_str:
                    groupes_list = [
                        g.strip() for g in groupes_str.split(",") if g.strip()
                    ]
                    groupes_a_ajouter.extend(groupes_list)
                found_service = True
                break

    # Construction du plan
    plan = {
        "user_dn": user_dn,
        "attributes": {
            "sAMAccountName": sam_account_name,
            "userPrincipalName": mail,
            "displayName": name,
            "givenName": prenom_display,
            "sn": nom_display,
            "mail": mail,
            "description": service_value_excel,
            "physicalDeliveryOfficeName": ville_brute,
            "streetAddress": str(agence_info.get("Adresse", "")),
            "l": str(agence_info.get("Ville", "")),
            "postalCode": str(agence_info.get("Codepostal", "")),
            "telephoneNumber": str(agence_info.get("tel", "")),
            "title": get_value_from_possible_keys(user_data, ["Fonction"]),
            "department": service_value_excel,
            "company": "ENTIS Mutuelles"
            if "entis" in domaine_mail
            else "MG PREVOYANCE",
            "unicodePwd": f'"{generated_password}"'.encode("utf-16-le"),
            "userAccountControl": "512",
        },
        "post_creation_attributes": {},
        "groups": list(set(groupes_a_ajouter)),
        "excel_info": {
            "prenom": prenom_display,
            "nom": nom_display,
            "email": mail,
            "ville": ville_brute,
            "service": service_value_excel,
            "materiel": "PC" if "pc" in computer_type_raw else "",
        },
        "final_message": f"Bonjour,\nLe compte pour {name} a bien été créé.\nVoici les informations de connexion:\nLogin : {sam_account_name}\nAdresse email : {mail}\nMot de passe temporaire : {generated_password}\n{"Le mot de passe devra être changé par l'utilisateur lors de sa première connexion." if 'mac' not in computer_type_raw else 'Le mot de passe sera à définir avec le service informatique.'}\nCordialement,\nLe service informatique",
        "sub_tickets": [],
    }

    if "mac" not in computer_type_raw:
        plan["post_creation_attributes"]["pwdLastSet"] = "0"
    date_exp_raw = get_value_from_possible_keys(
        user_data, ["Date de fin de contrat", "Fin de contrat"]
    )
    if pd.notna(date_exp_raw) and str(date_exp_raw):
        try:
            dt_obj = pd.to_datetime(date_exp_raw).to_pydatetime()
            plan["post_creation_attributes"]["accountExpires"] = str(
                int((dt_obj - datetime(1601, 1, 1)).total_seconds() * 10000000)
            )
        except Exception as e:
            print_warning(f"Format date expiration non valide. Ignoré. Erreur: {e}")

    contrat_type = str(
        get_value_from_possible_keys(user_data, ["Type de contrat"])
    ).lower()
    plan["post_creation_attributes"]["employeeType"] = (
        "IM" if "interim" in contrat_type else "Entis"
    )

    # RÉCAPITULATIF VISUEL
    print_header(f"RÉCAPITULATIF POUR {name.upper()}")
    print(f"  {Color.CYAN}{'Informations de connexion':-^78}{Color.ENDC}")
    print(f"  {Color.BOLD}Login AD:{Color.ENDC} {sam_account_name}")
    print(f"  {Color.BOLD}Email:{Color.ENDC} {mail}")
    print(f"  {Color.BOLD}Mot de passe temporaire:{Color.ENDC} {generated_password}")
    print(f"\n  {Color.CYAN}{'Organisation & Contrat':-^78}{Color.ENDC}")
    print(f"  {Color.BOLD}Destination (OU):{Color.ENDC} {agence_info['OU']}")

    contrat_type_brut = str(
        get_value_from_possible_keys(user_data, ["Type de contrat"])
    ).upper()
    date_fin_brute = plan["post_creation_attributes"].get("accountExpires")

    if "CDI" in contrat_type_brut or not date_fin_brute:
        print(f"  {Color.BOLD}Type de contrat:{Color.ENDC} CDI")
        print(f"  {Color.BOLD}Expiration du compte AD:{Color.ENDC} Jamais")
    else:
        dt_obj = pd.to_datetime(
            get_value_from_possible_keys(
                user_data, ["Date de fin de contrat", "Fin de contrat"]
            )
        )
        date_fin_humaine = dt_obj.strftime("%d/%m/%Y")
        print(f"  {Color.BOLD}Type de contrat:{Color.ENDC} {contrat_type_brut}")
        print(f"  {Color.BOLD}Expiration du compte AD:{Color.ENDC} {date_fin_humaine}")

    print(f"\n  {Color.CYAN}{'Appartenance aux groupes':-^78}{Color.ENDC}")
    if not found_service:
        print_warning(f"  Aucun groupe de service trouvé pour '{service_value_excel}'.")

    groupes_tries = sorted(plan["groups"])
    col_width = 38
    for i in range(0, len(groupes_tries), 2):
        col1 = f"  - {groupes_tries[i]}"
        col2 = f"- {groupes_tries[i + 1]}" if (i + 1) < len(groupes_tries) else ""
        print(f"{col1:<{col_width}}{col2}")

    if (
        input(
            f"\n{Color.WARNING}Valider ce plan de création ? (o/n) : {Color.ENDC}"
        ).lower()
        != "o"
    ):
        print_warning("Plan de création annulé.")
        return None

    choix_vip = input(f"{Color.BLUE}S'agit-il d'un VIP ? (o/n) : {Color.ENDC}").lower()
    groupes_a_ajouter.append(
        "GBV_ENTIS_EchangeVIPUsers" if choix_vip == "o" else "GBV_ENTIS_ExchangeUsers"
    )
    plan["groups"] = list(set(groupes_a_ajouter))

    # SOUS-TICKETS
    sub_tickets = []
    if "pc" in computer_type_raw or "mac" in computer_type_raw:
        ticket_line = f"Demande de matériel : {get_value_from_possible_keys(user_data, ['Ordinateur', 'Matériel'])}"
        reseau_info = get_value_from_possible_keys(user_data, ["Réseau"])
        peripherique_info = get_value_from_possible_keys(user_data, ["Périphérique"])
        extras = []
        if reseau_info and str(reseau_info).strip():
            extras.append(f"Réseau: {str(reseau_info).strip()}")
        if peripherique_info and str(peripherique_info).strip():
            extras.append(f"Périphériques: {str(peripherique_info).strip()}")
        if extras:
            ticket_line += f" ({'; '.join(extras)})"
        sub_tickets.append(ticket_line)

    telephone_info = get_value_from_possible_keys(user_data, ["Téléphone"])
    if (
        telephone_info
        and str(telephone_info).strip()
        and str(telephone_info).lower() != "aucun"
    ):
        sub_tickets.append(f"Demande téléphonie : {str(telephone_info).strip()}")

    plan["sub_tickets"] = sub_tickets
    plan["attributes"] = clean_attrs(plan["attributes"])
    plan["post_creation_attributes"] = clean_attrs(plan["post_creation_attributes"])

    return plan


def execute_ad_creation(plan, conn):
    try:
        user_dn = plan["user_dn"]
        conn.add(user_dn, "user", plan["attributes"])
        if conn.result["result"] != 0:
            if conn.result["result"] == 68:
                print_warning(
                    f"L'utilisateur {plan['attributes']['displayName']} existe déjà. Mise à jour des attributs."
                )
                updatable_attrs = {
                    k: v
                    for k, v in plan["attributes"].items()
                    if k not in ["sAMAccountName", "unicodePwd"]
                }
                for attr, value in updatable_attrs.items():
                    conn.modify(user_dn, {attr: [(MODIFY_REPLACE, [value])]})
            else:
                raise Exception(f"Échec de la création de base : {conn.result}")

        for attr, value in plan["post_creation_attributes"].items():
            conn.modify(user_dn, {attr: [(MODIFY_REPLACE, [value])]})

        for group_name in plan["groups"]:
            conn.search(
                BASE_DN,
                f"(&(objectClass=group)(sAMAccountName={group_name}))",
                attributes=["distinguishedName"],
            )
            if conn.entries:
                conn.modify(
                    conn.entries[0].distinguishedName.value,
                    {"member": [(MODIFY_ADD, [user_dn])]},
                )
            else:
                print_warning(f"Groupe '{group_name}' non trouvé.")

        print_success(
            f"Utilisateur {plan['attributes']['displayName']} créé/mis à jour avec succès."
        )
        return True, plan["attributes"]["displayName"], plan["final_message"]
    except Exception as e:
        print_error(
            f"ERREUR CRITIQUE lors de la création de {plan['attributes']['displayName']}: {e}"
        )
        return False, plan["attributes"]["displayName"], None


def run(
    ctx_glpi: ClientContext,
    ctx_arrivants: ClientContext,
    conn: Connection,
    config_dir: str,
    kp=None,
    **kwargs,
):
    """
    Point d'entrée synchronisé avec l'orchestrateur Hub.
    kp : session KeePass ouverte partagée par le Hub.
    kwargs : permet d'absorber d'autres arguments passés par le Hub.
    """
    try:
        print_header("PHASE 1 : PRÉPARATION ET VALIDATION DES PLANS DE CRÉATION")
        config_files = load_configs_from_local(config_dir)
        if not all(config_files):
            return
        if not ctx_arrivants:
            return

        df = read_arrivants_from_sharepoint(ctx_arrivants)
        if df is None:
            return

        users_to_process = display_and_select_users(df)
        if not users_to_process:
            return

        creation_plans = []
        for user_data in users_to_process:
            plan = prepare_user_creation_plan(user_data, conn, config_files)
            if plan:
                creation_plans.append(plan)

        if not creation_plans:
            return

        excel_data = [plan["excel_info"] for plan in creation_plans]
        if not write_users_to_sharepoint_excel(excel_data, ctx_glpi):
            print_error(
                "\nL'écriture dans le fichier de suivi a échoué. Aucun compte créé."
            )
            return

        print_header("PHASE 3 : CRÉATION DES COMPTES DANS ACTIVE DIRECTORY")
        created, failed, messages, sub_ticket_summary = [], [], [], {}
        for plan in creation_plans:
            success, name, msg = execute_ad_creation(plan, conn)
            if success:
                created.append(name)
                if plan.get("sub_tickets"):
                    sub_ticket_summary[name] = plan["sub_tickets"]
                if msg:
                    messages.append(msg)
            else:
                failed.append(name)

        print_header("RAPPORT FINAL DE L'OPÉRATION")
        if created:
            print_success(f"Utilisateurs créés : {', '.join(created)}")
        if failed:
            print_warning(f"Utilisateurs en échec : {', '.join(failed)}")

        if messages:
            print_header("MESSAGES À COPIER POUR LA RÉPONSE")
            print(("\n" + Color.CYAN + "=" * 80 + Color.ENDC + "\n").join(messages))

        if sub_ticket_summary:
            print_header("RÉCAPITULATIF DES SOUS-TICKETS À CRÉER")
            for user_name, tickets in sub_ticket_summary.items():
                print(f"  {Color.BOLD}Pour {user_name}:{Color.ENDC}")
                for ticket in tickets:
                    print(f"    - {ticket}")
                print("")

    except Exception as e:
        print_error(f"Erreur majeure : {e}")


if __name__ == "__main__":
    print_error("Ce module doit être piloté par 'hub_central.py'.")
