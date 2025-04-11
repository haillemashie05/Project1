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
import random
def generate_payslip(emp):
    from fpdf import FPDF
    import random

    pdf = FPDF()
    pdf.add_page()

    # === Setup styling ===
    company_color = (0, 102, 204)         # Deep Sky Blue
    header_bg = (0, 102, 204)
    section_bg = (245, 250, 255)
    border_color = (100, 100, 100)
    text_primary = (0, 51, 102)
    earnings_bg = (235, 245, 255)
    net_bg = (220, 255, 220)

    departments = [
        "Research & Development", "Sales & Marketing", "Production", "Quality Assurance",
        "Regulatory Affairs", "Finance", "Human Resources", "IT Support"
    ]
    department = random.choice(departments)
    net_salary = emp['Basic Salary'] + emp['Allowances'] - emp['Deductions']

    # === 🏢 Header: Company Title ===
    pdf.set_fill_color(*header_bg)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 20)
    pdf.cell(0, 15, "Sky Pharmaceuticals", ln=True, align='C', fill=True)

    pdf.set_font("Arial", 'I', 12)
    pdf.cell(0, 10, "Official Payslip Document", ln=True, align='C', fill=True)
    pdf.ln(8)

    # === 👤 Employee Info Section ===
    pdf.set_text_color(*text_primary)
    pdf.set_fill_color(*section_bg)
    pdf.set_draw_color(*border_color)
    pdf.set_font("Arial", 'B', 13)
    pdf.cell(0, 10, "Employee Details", ln=True, fill=True, border=1)

    pdf.set_font("Arial", '', 11)
    fields = [
        ("Employee ID", emp['Employee ID']),
        ("Name", emp['Name']),
        ("Job Title", emp.get('Job Title', 'Not Provided')),
        ("Department", department),
        ("Pay Period", emp.get('Month', 'This Month'))
    ]
    for label, value in fields:
        pdf.cell(0, 10, f"{label}: {value}", ln=True, border=1)
    pdf.ln(5)

    # === 💵 Salary Breakdown Section ===
    pdf.set_font("Arial", 'B', 12)
    pdf.set_fill_color(*earnings_bg)
    pdf.cell(95, 10, "EARNINGS", border=1, align='C', fill=True)
    pdf.cell(95, 10, "DEDUCTIONS", border=1, align='C', fill=True)
    pdf.set_font("Arial", '', 11)
    pdf.set_fill_color(255, 255, 255)
    pdf.ln()
    pdf.cell(95, 10, f"Basic Salary: ${emp['Basic Salary']:.2f}", border=1)
    pdf.cell(95, 10, f"Deductions: ${emp['Deductions']:.2f}", border=1)
    pdf.ln()
    pdf.cell(95, 10, f"Allowances: ${emp['Allowances']:.2f}", border=1)
    pdf.cell(95, 10, "", border=1)
    pdf.ln(10)

    # === ✅ Net Salary Highlight ===
    pdf.set_font("Arial", 'B', 13)
    pdf.set_text_color(0, 102, 0)
    pdf.set_fill_color(*net_bg)
    pdf.cell(0, 12, f"NET SALARY: ${net_salary:.2f}", ln=True, align='C', border=1, fill=True)

    # === 📄 Footer Note ===
    pdf.set_text_color(100, 100, 100)
    pdf.set_font("Arial", 'I', 10)
    pdf.ln(8)
    pdf.cell(0, 10, "This document is system-generated and does not require a signature.", align='C')

    # === Save Payslip ===
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

