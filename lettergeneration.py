from pathlib import Path
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

# Define file paths
base_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
excel_path = base_dir / "Participants.xlsx"
pdf_template_path = base_dir / "Participants invitation.pdf"
output_dir = base_dir / "OUTPUT"

# Create output folder if it does not exist
output_dir.mkdir(exist_ok=True)

# Read Excel file into pandas DataFrame
df = pd.read_excel(excel_path, sheet_name="Sheet1")

# Iterate over each row in the DataFrame
for _, record in df.iterrows():
    # Create a temporary PDF with data
    temp_pdf_path = base_dir / "temp.pdf"
    c = canvas.Canvas(str(temp_pdf_path))

    # Set font for participant's name
    c.setFont("Helvetica", 12)

    # Define desired coordinates
    name_x = 82  # X coordinate for Participant's Name
    name_y = 607  # Y coordinate for Participant's Name

    roll_x = 154  # X coordinate for Roll number
    roll_y = 666  # Y coordinate for Roll number

    # Draw participant's name
    c.drawString(name_x, name_y, f"{record[1]}")

    # Change font to bold for roll number
    c.setFont("Helvetica-Bold", 12)

    # Convert the roll number to integer and draw it
    roll_number = int(record[0])
    c.drawString(roll_x, roll_y, str(roll_number))

    # Save the temporary PDF
    c.save()

    # Load the existing PDF template
    template_reader = PdfReader(pdf_template_path)

    # Load the temporary PDF with data
    temp_pdf_reader = PdfReader(temp_pdf_path)

    # Create a PDF writer
    pdf_writer = PdfWriter()

    # Merge data onto the first page of the existing template
    first_page = template_reader.pages[0]
    data_page = temp_pdf_reader.pages[0]

    # Merge the data from the temporary PDF onto the first page of the existing template
    first_page.merge_page(data_page)

    # Add the modified first page to the writer
    pdf_writer.add_page(first_page)

    # Add any remaining pages from the existing template
    for page in template_reader.pages[1:]:
        pdf_writer.add_page(page)

    # Save the merged PDF
    output_path = output_dir / f"{record[1]}-invitation.pdf"
    with open(output_path, "wb") as output_file:
        pdf_writer.write(output_file)

    # Clean up the temporary PDF file
    # Path(temp_pdf_path).unlink()