import requests
import pandas as pd
import smtplib
from email.message import EmailMessage
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

# ========== CONFIGURATION ==========

SMTP_SERVER = "smtp-mail.outlook.com"
SMTP_PORT = 587
EMAIL_SENDER = "network@clincoord.org"
EMAIL_PASSWORD = ""
EMAIL_RECEIVER = ["rodrigo.lima.cc.ao@gmail.com"]
EMAIL_SUBJECT = "New Clinical Trials Found"
XLSX_FILENAME = "clinical_trials_angola.xlsx"

# ========== FLATTEN FUNCTION ==========

def flatten_json(y, prefix=''):
    out = {}
    if isinstance(y, dict):
        for k, v in y.items():
            if k == 'locations' and isinstance(v, list):
                locations_info = []
                for loc in v:
                    facility = loc.get('facility', 'N/A')
                    city = loc.get('city', 'N/A')
                    state = loc.get('state')
                    country = loc.get('country', 'N/A')
                    parts = [facility, city]
                    if state:
                        parts.append(state)
                    parts.append(country)
                    locations_info.append(", ".join(parts))
                out[f"{prefix}combined_locations_string"] = "; ".join(locations_info)
            elif k in ['centralContacts', 'overallOfficials'] and isinstance(v, list):
                if v:
                    out.update(flatten_json(v[0], f"{prefix}{k}.0."))
            else:
                out.update(flatten_json(v, f"{prefix}{k}."))
    elif isinstance(y, list):
        if all(isinstance(i, (str, int, float, bool)) for i in y):
            out[prefix[:-1]] = ", ".join(map(str, y))
        else:
            for i, item in enumerate(y):
                out.update(flatten_json(item, f"{prefix}{i}."))
    else:
        out[prefix[:-1]] = y
    return out

# ========== SCRAPE FUNCTION WITH GENERATOR ==========

def scrape_clinicaltrials_gov_api():
    url = "https://clinicaltrials.gov/api/v2/studies"
    headers = {"User-Agent": "Mozilla/5.0"}
    page_token = None
    page_num = 1

    while True:
        params = {"pageSize": 1000}
        if page_token:
            params["pageToken"] = page_token

        try:
            response = requests.get(url, headers=headers, params=params, timeout=30)
            response.raise_for_status()
            data = response.json()
        except Exception as e:
            print(f"Erro na página {page_num}: {e}")
            break

        studies = data.get("studies", [])
        if not studies:
            break

        for study in studies:
            yield flatten_json(study)

        print(f"Página {page_num}: {len(studies)} estudos processados.")
        page_num += 1

        page_token = data.get("nextPageToken")
        if not page_token:
            break

# ========== EXCEL FORMATTER ==========

def format_excel(filename):
    try:
        wb = load_workbook(filename)
        ws = wb.active

        for col in ws.columns:
            max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
            max_len = min(max_len + 2, 100)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max_len

        ws.freeze_panes = ws['A2']
        ws.auto_filter.ref = ws.dimensions

        for idx, cell in enumerate(ws[1], start=1):
            if cell.value == "Trial Link":
                for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                    cell = row[0]
                    if cell.value:
                        cell.hyperlink = cell.value
                        cell.font = Font(color="0000FF", underline="single")
                break

        wb.save(filename)
        print(f"Arquivo Excel '{filename}' formatado.")
    except Exception as e:
        print(f"Erro ao formatar Excel: {e}")

# ========== EMAIL ==========

def send_email(trials_count):
    msg = EmailMessage()
    msg['Subject'] = EMAIL_SUBJECT
    msg['From'] = EMAIL_SENDER
    msg['To'] = ", ".join(EMAIL_RECEIVER)

    if trials_count > 0:
        html = f"""
        <html>
        <body>
            <h2>Novos Ensaios Clínicos Encontrados</h2>
            <p>{trials_count} novo(s) ensaio(s) clínico(s) foi(ram) encontrado(s) e salvo(s) no Excel.</p>
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

    msg.set_content("Este email contém conteúdo HTML.")
    msg.add_alternative(html, subtype='html')

    if trials_count > 0:
        with open(XLSX_FILENAME, "rb") as f:
            msg.add_attachment(f.read(), maintype="application",
                               subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               filename=XLSX_FILENAME)
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
            print("Email enviado com sucesso.")
    except Exception as e:
        print(f"Erro ao enviar email: {e}")

# ========== MAIN ==========

def main():
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
        "protocolSection.contactsLocationsModule.combined_locations_string": "Location"
    }

    rows = []
    for flat_study in scrape_clinicaltrials_gov_api():
        row = {v: flat_study.get(k, "") for k, v in columns_map.items()}
        nct = flat_study.get("protocolSection.identificationModule.nctId", "")
        row["Trial Link"] = f"https://clinicaltrials.gov/study/{nct}" if nct else ""
        rows.append(row)

    if rows:
        df = pd.DataFrame(rows)
        df = df[df["Trial Status"].isin(["RECRUITING", "NOT_YET_RECRUITING"])]
        print(f"{len(df)} trials válidos encontrados.")

        df.to_excel(XLSX_FILENAME, index=False)
        format_excel(XLSX_FILENAME)
        send_email(len(df))
    else:
        print("Nenhum estudo retornado.")
        send_email(0)

if __name__ == "__main__":
    main()
