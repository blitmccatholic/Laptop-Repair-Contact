V1.1

Student Invoice Generator

A simple Python tool for generating PDF letters and Outlook drafts for missing or damaged student devices at Thomas More College.

NOTE: Currently uses Outlook Classic instead of the new Outlook


Requirements

Install required libraries before running:

pip install reportlab pywin32


Setup

Place the following files in the same folder as the script:
Logo.png (school logo for the header)
Footer.png (footer image)

Open the script and change the staff name at the top if required:
STAFF_MEMBER = "Your Name"


Run

python student_invoice_generator.py


Features

Generates PDF letters for missing or damaged devices

Add multiple items with costs and automatic total calculation

Includes school header and footer

Automatically opens an Outlook email draft with the PDF attached

Prompts user to choose a save location for the PDF


Output

PDFs are automatically named in the format:
TMC_{Status}Device{Student}.pdf

Example:
TMC_Damaged_Device_John_Smith.pdf


Future Features


Lookup from csv to autofill student name, parent name, parent email
