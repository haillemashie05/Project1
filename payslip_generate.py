import os
import pandas as pd
from fpdf import FPDF
from email.message import EmailMessage
import smtplib
from dotenv import load_dotenv

# Load .env credentials
load_dotenv()

EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))

# Debug check for environment variables
if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
    print("❌ ERROR: EMAIL_ADDRESS or EMAIL_PASSWORD not set in .env")
    exit()

# Create payslips folder
os.makedirs("payslips", exist_ok=True)

# Load Excel file
file_path = r"C:\Users\uncommonStudent\Documents\Project1\Employee data (1).xlsx"
print("Found file:", os.path.exists(file_path))
if not os.path.exists(file_path):
    print("❌ ERROR: Excel file not found.")
    exit()

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()  # Clean headers

# Payslip generator
def generate_payslip(emp):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    net_salary = emp['Basic Salary'] + emp['Allowances'] - emp['Deductions']
    lines = [
        f"Employee ID: {emp['Employee ID']}",
        f"Name: {emp['Name']}",
        f"Basic Salary: ${emp['Basic Salary']}",
        f"Allowances: ${emp['Allowances']}",
        f"Deductions: ${emp['Deductions']}",
        f"Net Salary: ${net_salary}"
    ]

    for line in lines:
        pdf.cell(200, 10, txt=line, ln=True)

    filename = f"payslips/{emp['Employee ID']}.pdf"
    pdf.output(filename)
    return filename

# Email sender
def send_email(to_email, filename, name):
    msg = EmailMessage()
    msg['Subject'] = "Your Payslip for This Month"
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email
    msg.set_content(f"Dear {name},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nHR")

    with open(filename, "rb") as f:
        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=os.path.basename(filename))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"❌ Email send error for {name}: {e}")
        return False

# Main loop
for _, emp in df.iterrows():
    try:
        payslip_file = generate_payslip(emp)
        sent = send_email(emp['Email'], payslip_file, emp['Name'])
        if sent:
            print(f"✅ Sent payslip to {emp['Name']} at {emp['Email']}")
    except Exception as e:
        print(f"❌ Error processing {emp['Name']}: {e}")
