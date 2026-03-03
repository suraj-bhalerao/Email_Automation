import requests
import re
import os
import logging
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from openpyxl import Workbook
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# -------------------------------
# CONFIG
# -------------------------------

TIMEOUT = 3
MAX_WORKERS = 15
DATA_DIR = "data"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# -------------------------------
# SESSION WITH RETRY
# -------------------------------

def create_session():
    session = requests.Session()
    retry = Retry(
        total=2,
        backoff_factor=0.5,
        status_forcelist=[429, 500, 502, 503, 504]
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update(HEADERS)
    return session

session = create_session()

# -------------------------------
# DOMAIN RESOLUTION
# -------------------------------

def generate_possible_domains(company):
    cleaned = re.sub(r"[^a-zA-Z0-9]", "", company.lower())
    return [
        f"https://{cleaned}.com",
        f"https://www.{cleaned}.com",
        f"https://{cleaned}.in",
        f"https://{cleaned}.co"
    ]

def validate_domain(url):
    try:
        response = session.head(url, timeout=TIMEOUT, allow_redirects=True)
        if response.status_code < 400:
            return response.url
    except:
        return None
    return None

def resolve_domain(company):
    for url in generate_possible_domains(company):
        valid = validate_domain(url)
        if valid:
            logging.info(f"Resolved {company} → {valid}")
            return valid
    logging.warning(f"Could not resolve domain for {company}")
    return None

# -------------------------------
# EMAIL VALIDATION
# -------------------------------

def is_valid_email(email):
    email = email.lower()
    local = email.split("@")[0]

    blocked = [
        "support", "info", "contact", "admin",
        "help", "sales", "marketing",
        "noreply", "no-reply", "billing"
    ]

    preferred = [
        "hr", "career", "careers",
        "recruitment", "talent", "hiring"
    ]

    if any(local.startswith(b) for b in blocked):
        return False

    if any(local.startswith(p) for p in preferred):
        return True

    if re.match(r"^[a-z]+([._]?[a-z0-9]+)*$", local):
        return True

    return False

def extract_emails_from_html(html):
    pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
    found = re.findall(pattern, html)
    return {email for email in found if is_valid_email(email)}

# -------------------------------
# SCRAPER
# -------------------------------

def scrape_company(company):
    records = []

    base_url = resolve_domain(company)
    if not base_url:
        return records

    emails = set()

    try:
        response = session.get(base_url, timeout=TIMEOUT)
        if response.status_code != 200:
            return records

        emails |= extract_emails_from_html(response.text)

        # Stop early if found
        if emails:
            for email in emails:
                records.append([company, base_url, email])
            return records

        soup = BeautifulSoup(response.text, "html.parser")

        # Look for relevant pages
        for link in soup.find_all("a", href=True):
            href = link["href"].lower()

            if any(k in href for k in ["contact", "career", "team"]):
                full_url = urljoin(base_url, link["href"])

                try:
                    sub_resp = session.get(full_url, timeout=TIMEOUT)
                    emails |= extract_emails_from_html(sub_resp.text)
                except:
                    continue

        for email in emails:
            records.append([company, base_url, email])

    except Exception as e:
        logging.error(f"Error scraping {company}: {str(e)}")

    return records

# -------------------------------
# SAVE TO EXCEL
# -------------------------------

def save_to_excel(records):
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    wb = Workbook()
    ws = wb.active
    ws.title = "Company Emails"
    ws.append(["Company Name", "Resolved Domain", "Email ID"])

    for record in records:
        ws.append(record)

    file_path = os.path.join(DATA_DIR, "scraped_email.xlsx")
    wb.save(file_path)

    logging.info(f"Saved results to {file_path}")

# -------------------------------
# MAIN
# -------------------------------

if __name__ == "__main__":

    company_list = [
        "3M",
        "ABB",
        "Ajackus",
        "Abbott",
        "Adobe",
        "AMD",
        "American Express",
        "Apple",
        "AstraZeneca",
        "Atlassian",
        "Bain & Company",
        "Bank of America",
        "Barclays",
        "Birlasoft",
        "BigBasket",
        "Bloomberg",
        "Bosch",
        "Boston Consulting Group",
        "BrowserStack",
        "BYJU'S",
        "Bain & Company",
        "Capgemini",
        "Capillary Technologies",
        "CarDekho",
        "Cartesian Consulting",
        "Cashfree",
        "Chargebee",
        "Cigniti Technologies",
        "CISCO Systems",
        "Citibank",
        "Citrix",
        "Cleartax",
        "CleverTap",
        "Coca-Cola",
        "Coforge",
        "Colgate-Palmolive",
        "Coverfox",
        "CRED",
        "Darwinbox",
        "Damco Solutions",
        "Decathlon",
        "Delhivery",
        "Deloitte",
        "Dell Technologies",
        "Dream11",
        "DXC Technology",
        "EPAM Systems",
        "Ericsson",
        "EXL Service",
        "EY",
        "Facile Consulting",
        "Fi Money",
        "Fingent",
        "FirstCry",
        "Flipkart",
        "Foxconn",
        "Fractal Analytics",
        "Freshworks",
        "GE Healthcare",
        "Genpact",
        "GitHub",
        "GlaxoSmithKline",
        "Globant",
        "Goldman Sachs",
        "Google",
        "GoodWorkLabs",
        "Groww",
        "Gupshup",
        "HashedIn",
        "HCL Technologies",
        "Hexaware Technologies",
        "HighRadius",
        "Hitachi",
        "Honda Motor Company",
        "Honeywell",
        "HSBC",
        "Huawei",
        "Hewlett Packard Enterprise",
        "Hyundai Motor Group",
        "IBM",
        "IKEA",
        "Indium Software",
        "Ingram Micro",
        "Infosys",
        "Infra.Market",
        "InMobi",
        "Instamojo",
        "Intel",
        "Intellect Design Arena",
        "JLL",
        "Johnson & Johnson",
        "JPMorgan Chase",
        "Juspay",
        "KPIT Technologies",
        "KPMG",
        "Keka",
        "Larsen & Toubro Infotech",
        "LambdaTest",
        "LeadSquared",
        "Lenovo",
        "LG Electronics",
        "Licious",
        "LoanTap",
        "LTIMindtree",
        "L'Oréal",
        "Mastercard",
        "Media.net",
        "Meesho",
        "Meta",
        "Microsoft",
        "Mindtree",
        "Mphasis",
        "Morgan Stanley",
        "Mu Sigma",
        "Nagarro",
        "Nestlé",
        "Netflix",
        "NetApp",
        "Newgen Software",
        "Nokia",
        "NoBroker",
        "Novartis",
        "Nykaa",
        "OfBusiness",
        "Ola",
        "OnMobile",
        "Open Financial",
        "Oracle",
        "Panasonic",
        "PayNearby",
        "PayPal",
        "Paytm",
        "PayU",
        "PepsiCo",
        "Perfios",
        "Persistent Systems",
        "Pfizer",
        "Philips",
        "PhonePe",
        "Pine Labs",
        "Plum",
        "PolicyBazaar",
        "Postman",
        "Pratilipi",
        "Procter & Gamble",
        "PwC",
        "QBurst",
        "Qualcomm",
        "Qualitest",
        "Razorpay",
        "Red Hat",
        "Route Mobile",
        "Rupeek",
        "Sakson Technologies",
        "Salesforce",
        "Sanofi",
        "SAP",
        "Sasken Technologies",
        "Scaler",
        "Schneider Electric",
        "ServiceNow",
        "ShareChat",
        "Siemens",
        "Siemens India",
        "Signzy",
        "Slice",
        "Snowflake",
        "Sony",
        "Standard Chartered",
        "Stripe",
        "Subex",
        "Suzuki Motor Corporation",
        "Swiggy",
        "Synechron",
        "Tanla Platforms",
        "Tata Consultancy Services",
        "Tech Mahindra",
        "TestVagrant",
        "Testsigma",
        "Texas Instruments",
        "The Boston Consulting Group",
        "The Coca-Cola Company",
        "Thomson Reuters",
        "Thoughtworks",
        "Tiger Analytics",
        "TO THE NEW",
        "Toyota Motor Corporation",
        "Tricentis India",
        "Twilio",
        "Uber",
        "Udaan",
        "Unacademy",
        "Unilever",
        "UpGrad",
        "Urban Company",
        "UST",
        "Valuelabs",
        "Vedantu",
        "VISA",
        "Virtusa",
        "VMware",
        "Walmart",
        "WebEngage",
        "Whatfix",
        "Wipro",
        "Xoriant",
        "Yellow.ai",
        "Yodlee",
        "Zensar Technologies",
        "Zepto",
        "Zeta",
        "Zoho",
        "Zomato"
    ]


    company_list = list(set(company_list))

    all_records = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(scrape_company, company)
                   for company in company_list]

        for future in as_completed(futures):
            result = future.result()
            if result:
                all_records.extend(result)

    save_to_excel(all_records)