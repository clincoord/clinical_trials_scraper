# clinical_trials_scraper.py

import requests
import pandas as pd
import smtplib
from email.message import EmailMessage
from datetime import datetime
import os
# Importar openpyxl para manipulação avançada de Excel
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ========== CONFIGURATION ==========

MAX_RESULTS = 1000
SMTP_SERVER = "smtp-mail.outlook.com"
SMTP_PORT = 587
EMAIL_SENDER = "network@clincoord.org"
EMAIL_PASSWORD = "dcpcknbsykkhfcrb"
EMAIL_RECEIVER = ["rodrigo.lima.cc.ao@gmail.com","tatiana.gomes@clincoord.org" , "Nivison.Nery@clincoord.org" "Joe.coffie@clincoord.org" ]
EMAIL_SUBJECT = "New Clinical Trials Found"
XLSX_FILENAME = "clinical_trials_angola.xlsx"

# ========== FLATTEN FUNCTION ==========

def flatten_json(y, prefix=''):
    out = {}
    if isinstance(y, dict):
        for k, v in y.items():
            # Handle specific cases for lists of objects that need custom flattening
            if k == 'locations' and isinstance(v, list): # CHAVE CORRETA É 'locations'
                if v:
                    locations_info = []
                    for loc in v:
                        facility = loc.get('facility', 'N/A')
                        city = loc.get('city', 'N/A')
                        state = loc.get('state', None) 
                        country = loc.get('country', 'N/A')
                        
                        location_str = f"{facility}, {city}"
                        if state:
                            location_str += f", {state}"
                        location_str += f", {country}"
                        locations_info.append(location_str)
                    out[f"{prefix}combined_locations_string"] = "; ".join(locations_info) 
                else:
                    out[f"{prefix}combined_locations_string"] = "" # No locations found
            elif k == 'centralContacts' and isinstance(v, list):
                if v:
                    first_contact = v[0]
                    out.update(flatten_json(first_contact, f"{prefix}{k}.0."))
            elif k == 'overallOfficials' and isinstance(v, list):
                if v:
                    first_official = v[0]
                    out.update(flatten_json(first_official, f"{prefix}{k}.0."))
            else:
                out.update(flatten_json(v, f"{prefix}{k}."))  # dot notation
    elif isinstance(y, list):
        if all(isinstance(i, (str, int, float, bool)) for i in y):
            out[prefix[:-1]] = ", ".join(map(str, y))  # simple list
        else:
            for i, item in enumerate(y):
                out.update(flatten_json(item, f"{prefix}{i}."))
    else:
        out[prefix[:-1]] = y
    return out

# ========== SCRAPE FUNCTION ==========

def scrape_clinicaltrials_gov_api(max_results=1000):
    url = "https://clinicaltrials.gov/api/v2/studies"
    params = {"pageSize": max_results}
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(url, headers=headers, params=params)

    if response.status_code != 200:
        print(f"Failed to fetch data: HTTP {response.status_code}")
        return []

    try:
        data = response.json()
    except ValueError:
        print("Failed to decode JSON response.")
    except requests.exceptions.JSONDecodeError: 
        print("Failed to decode JSON response due to invalid JSON content.")
        return []

    trials = []
    for study in data.get("studies", []):
        flat_study = flatten_json(study)
        trials.append(flat_study)
    return trials

# ========== EXCEL FORMATTING FUNCTION ==========

def format_excel(filename):
    try:
        workbook = load_workbook(filename)
        sheet = workbook.active

        # 1. Ajustar largura das colunas automaticamente
        for column in sheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try: # handle non-string values
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except TypeError:
                    pass
            adjusted_width = (max_length + 2) # Adiciona um pouco de padding
            if adjusted_width > 100: # Limitar largura máxima para evitar colunas excessivamente largas
                adjusted_width = 100
            sheet.column_dimensions[column_letter].width = adjusted_width

        # 2. Congelar painéis (primeira linha - cabeçalhos)
        sheet.freeze_panes = sheet['A2'] # Congela a partir da célula A2, mantendo a linha 1 visível

        # 3. Adicionar filtro automático aos cabeçalhos
        sheet.auto_filter.ref = sheet.dimensions # Aplica filtro a toda a região de dados

        workbook.save(filename)
        print(f"Excel file '{filename}' formatted successfully.")
    except Exception as e:
        print(f"Error formatting Excel file: {e}")

# ========== EMAIL FUNCTION ==========

def send_email(trials_to_send):
    msg = EmailMessage()
    msg['Subject'] = EMAIL_SUBJECT
    msg['From'] = EMAIL_SENDER
    msg['To'] = ", ".join(EMAIL_RECEIVER)

    if trials_to_send:
        html = f"""
        <html>
        <body>
            <h2>Novos Ensaios Clínicos Encontrados</h2>
            <p>{len(trials_to_send)} novo(s) ensaio(s) clínico(s) foi(ram) encontrado(s) e salvo(s) no Excel.</p>
        </body>
        </html>
        """
    else:
        html = """
        <html>
        <body>
            <h2>Nenhum Novo Ensaio Clínico Encontrado</h2>
            <p>Nenhum ensaio clínico correspondente encontrado no momento.</p>
        </body>
        </html>
        """

    msg.set_content("Este email contém conteúdo HTML. Por favor, visualize-o em um cliente de email compatível com HTML.")
    msg.add_alternative(html, subtype='html')

    if trials_to_send:
        try:
            with open(XLSX_FILENAME, "rb") as f:
                msg.add_attachment(
                    f.read(),
                    maintype="application",
                    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    filename=XLSX_FILENAME
                )
        except FileNotFoundError:
            print(f"Erro: O arquivo Excel '{XLSX_FILENAME}' não foi encontrado. Não é possível anexá-lo ao email.")
    
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
            print("Email enviado com sucesso.")
    except smtplib.SMTPAuthenticationError as e:
        print(f"Erro de Autenticação SMTP: {e}. Verifique o remetente e a senha do seu email.")
    except smtplib.SMTPException as e:
        print(f"Erro SMTP: {e}. Não foi possível enviar o email.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao enviar o email: {e}")


# ========== MAIN FUNCTION ==========

def main():
    trials = scrape_clinicaltrials_gov_api(MAX_RESULTS)

    if trials:
        df = pd.DataFrame(trials)

        # Campos da API mapeados para nomes amigáveis
        columns_map = {
            "protocolSection.identificationModule.nctId": "Trial Registry Number (.gov)",
            "protocolSection.sponsorCollaboratorsModule.leadSponsor.name": "Sponsor Name",
            "protocolSection.sponsorCollaboratorsModule.leadSponsor.class": "Sponsor Type",
            
            "protocolSection.contactsLocationsModule.centralContacts.0.name": "Contact Person",
            "protocolSection.contactsLocationsModule.centralContacts.0.role": "Role",
            "protocolSection.contactsLocationsModule.centralContacts.0.phone": "Phone Number",
            "protocolSection.contactsLocationsModule.centralContacts.0.email": "Email",
            
            "protocolSection.identificationModule.briefTitle": "Trial Name",
            "protocolSection.identificationModule.officialTitle": "Trial/Project Title",
            "protocolSection.designModule.phases": "Trial Phase",
            "protocolSection.statusModule.overallStatus": "Trial Status",
            "protocolSection.conditionsModule.conditions": "Therapeutic Area/Research Category",
            "protocolSection.armsInterventionsModule.interventions.0.name": "Intervention/Investigational Product",
            "protocolSection.statusModule.startDateStruct.date": "Trial Start Date",
            "protocolSection.statusModule.completionDateStruct.date": "Trial End Date",
            
            "protocolSection.contactsLocationsModule.combined_locations_string": "Location", 
            
            # Campos extras personalizados (não vindos da API)
            "custom.num": "Num.",
            "custom.ccsn_previous_sponsorship": "CCSN PI Previous Sponsorship History",
            "custom.feasibility_regions": "Regions with Opened Feasibility",
            "custom.ccsn_potential_sites": "Pontential Sites at CCSN?",
            "custom.ccsn_contact_notes": "CCSN Contact Notes",
            "custom.additional_notes": "Addional Notes/Comments"
        }

        # Seleciona colunas disponíveis e renomeia
        df_filtered = pd.DataFrame()
        for api_key, friendly_name in columns_map.items():
            if api_key in df.columns:
                df_filtered[friendly_name] = df[api_key]
            else:
                df_filtered[friendly_name] = "" 

        # --- FILTRAGEM POR STATUS ---
        allowed_statuses = ["RECRUITING", "NOT_YET_RECRUITING"]
        if "Trial Status" in df_filtered.columns:
            df_filtered = df_filtered[df_filtered["Trial Status"].isin(allowed_statuses)].copy()
            print(f"Filtered to {len(df_filtered)} trials with status 'RECRUITING' or 'NOT_YET_RECRUITING'.")
        else:
            print("Warning: 'Trial Status' column not found for filtering.")
        # --- FIM DA FILTRAGEM ---

        # Reorganiza a ordem das colunas para corresponder ao `desired_order`
        desired_order = list(columns_map.values())
        for col in desired_order:
            if col not in df_filtered.columns:
                df_filtered[col] = "" 

        df_final = df_filtered[desired_order]

        # Salva em Excel
        df_final.to_excel(XLSX_FILENAME, index=False)
        print(f"{len(df_final)} trials saved to '{XLSX_FILENAME}'.")
        
        # CHAMA A FUNÇÃO DE FORMATAÇÃO AQUI
        format_excel(XLSX_FILENAME)

        send_email(df_final.to_dict('records'))
    else:
        send_email([])
        print("No trials found from API call.")

if __name__ == "__main__":
    main()
