# 🧾 Payslip Generator

The **Payslip Generator** is a Python-based automation tool that simplifies HR payroll tasks. It reads structured employee data from an Excel spreadsheet, computes the net salary for each employee, and generates individual PDF payslips. These payslips are then automatically emailed to the respective employees. This project is ideal for small organizations or as a learning exercise for students interested in file handling, PDF generation, and email automation in Python.

The script combines multiple Python libraries to handle data processing, file creation, and secure communication with email servers — all in one streamlined solution.

---

## 📌 Features

This project showcases the integration of various Python modules to achieve a real-world task efficiently:

- 📊 **Data Handling with `pandas`**  
  Loads and processes employee data from an Excel spreadsheet. Strips whitespace from headers and ensures consistent formatting.

- 🧮 **Net Salary Calculation**  
  Computes net salary using the formula:  
  `Net Salary = Basic Salary + Allowances - Deductions`

- 📄 **PDF Generation using `fpdf`**  
  Creates professionally formatted payslip PDFs for each employee with their salary breakdown.

- 📧 **Email Automation with `smtplib` and `email` modules**  
  Sends each generated payslip directly to the employee's email address with a customized message and attachment.

- 🔐 **Secure Configuration with `python-dotenv`**  
  Keeps sensitive credentials (like email address and password) outside of the script using a `.env` file, enhancing security and flexibility.

- 🗂️ **Output Organization**  
  Saves all generated payslips in a `payslips/` directory for easy access and organization.

---

## 🛠 Requirements

Before running the script, make sure you have Python installed and the required packages:

- [`pandas`](https://pandas.pydata.org/): for reading and processing Excel files
- [`fpdf`](https://pyfpdf.github.io/fpdf2/): for generating PDF payslips
- [`python-dotenv`](https://pypi.org/project/python-dotenv/): for managing configuration via environment variables

### ✅ Install dependencies with pip:

```bash
pip install pandas fpdf python-dotenv
