from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import win32com.client  # For Word to PDF conversion (on Windows only)

def set_font(paragraph, font_name="Calibri", font_size=11, font_bold=False, font_underline=False):
    """Sets the font and size for a given paragraph."""
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = font_bold
        run.font.underline = font_underline
        
        # Ensure the font is set properly for non-default styles
        rPr = run._element.rPr
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), font_name)
        rFonts.set(qn("w:hAnsi"), font_name)
        rFonts.set(qn("w:cs"), font_name)
        rPr.append(rFonts)

def generate_motivation_letter(company_name, company_address, plz, stadt, position):
    # Load the Word template
    save_path = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(save_path, "Anschreiben", "Anschreiben_MarwanHammad.docx")
    
    if not os.path.exists(template_path):
        print(f"Template not found at {template_path}. Please ensure the file exists.")
        return
    doc = Document(template_path)


    # Replace placeholders with the company data
    for paragraph in doc.paragraphs:
        if "[Name_Unternehmen]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[Name_Unternehmen]", company_name)
            set_font(paragraph)
        if "[Straße_Unternehmen]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[Straße_Unternehmen]", company_address)
            set_font(paragraph)
        if "[PLZ]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[PLZ]", plz)
            set_font(paragraph)
        if "[Stadt]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[Stadt]", stadt)
            set_font(paragraph)
        if "[Position]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[Position]", position)
            set_font(paragraph, "Calibri", 14, True)
        if "[position]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[position]", position)
            set_font(paragraph)
    
    # Define the output directory within the same directory as the script
    output_dir = os.path.join(save_path, "Anschreiben")  # "Anschreiben" subfolder

    # Create the output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Save the updated Word document in the output directory
    word_file = os.path.join(output_dir, f"{company_name}.docx")
    doc.save(word_file)
    
    # Convert the Word document to a PDF
    convert_to_pdf(word_file)
    print(f"Motivation letter saved as Word and PDF in: {save_path}")
   
    # Delete the Word document after PDF conversion
    try:
        os.remove(word_file)
        print(f"Temporary Word file deleted: {word_file}")
    except Exception as e:
        print(f"Error deleting the Word file: {e}")


def convert_to_pdf(word_file):
    pdf_file = word_file.replace(".docx", ".pdf")
    word = None
    try:
        # Launch Word application
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(word_file)
        # Save as PDF
        doc.SaveAs(pdf_file, FileFormat=17)  # 17 is the code for PDF format
        doc.Close()
        print(f"PDF saved at: {pdf_file}")
    except Exception as e:
        print(f"An error occurred during Word to PDF conversion: {e}")
    finally:
        # Ensure Word application closes
        if word:
            word.Quit()
    
# generate a motivation letter

# Call the function to generate the motivation letter
if __name__ == "__main__":
    # Get user inputs for the company name and address
    name = input("Enter the company name: ")
    address = input("Enter the company address: ")
    while True:
        plz = input("Enter the company PLZ (numbers only): ")
        if plz.isdigit():
            break
        else:
            print("Invalid input. Please enter numbers only.")
    stadt = input("Enter the company city: ")
    position = input("Enter the Position: ")
    
    # Generate the motivation letter
    generate_motivation_letter(name, address, plz, stadt, position)
