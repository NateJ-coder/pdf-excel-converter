# backend/app.py
# Updated to include:
# - master description list
# - predefined row order
# - file-based year alignment
# - normalized OCR text
# - inclusion of total lines if missing
# - Number formatting for currency columns in Excel
# - Refactored Category Inference with Canonical Descriptions
# - Logging Module instead of print()
# - Pre-initialization of consolidated_items to ensure all MASTER_STRUCTURE descriptions appear
# - Improved Gemini prompt to avoid unnecessary total recalculations
# - Removed responseSchema from Gemini API API call for more flexible parsing
# - Refined Excel generation logic for precise data placement and summary calculations
# - Broadened Gemini prompt to extract data from all relevant financial sections (including notes)
# - Enhanced infer_categories_and_structure to handle more diverse financial line items
# - Improved consolidated_items_by_key population to strictly align with MASTER_STRUCTURE for known items
# - Removed the Summary sheet entirely as per user request
# - Further refined data consolidation to capture all years and granular items
# - Added more robust mapping from Gemini output to MASTER_STRUCTURE
# - Integrated custom AI prompt from frontend for parsing.
# - FIXED: NameError: name 'custom_prompt_text' is not defined in upload_pdfs.
# - FIXED: KeyError: '"Description"' due to f-string and .format() conflict.
# - Modified generate_excel to produce a hierarchical "Description | Year1 | Year2 | ..." table.
# - Addressed duplicate/inflated values by refining canonicalization and mapping.
# - Improved handling of missing values by robust mapping.
# - Fixed line item confusion by ensuring correct value assignment.
# - Trimmed empty category/subcategory rows for cleanliness.
# - Further enhanced number cleaning for robustness (e.g., handling spaces as thousands separators).
# - Reviewed and reinforced canonicalization for "Bank balances" and "Short-term deposits".
# - Added more explicit examples for "Bank balances" and "ABSA" to master structure under Current Assets.
# - Added specific instruction in Gemini prompt for numbers with spaces as thousands separators.
# - **NEW**: Added more explicit examples for "Short-term deposits" for 2020 and 2019 in Gemini prompt.
# - **NEW**: Implemented post-processing logic to ensure "Accumulated deficit" and "Accumulated surplus" are mutually exclusive per year.
# - **NEW**: Integrated Document Refinement Tool with new routes and helper functions.
# - **UPDATED**: Document Refinement Tool now parses DOCX files properly using python-docx.
# - **UPDATED**: Gemini prompt for refinement is more specific about TOC, Definitions, and Director's Table.
# - **UPDATED**: `create_refined_document` now attempts to copy template structure and insert specific content.
# - **UPDATED**: `detect_text_from_pdf` now accepts content directly, removing need to save PDF to disk.
# - **FIXED**: SyntaxError: unterminated string literal in `create_refined_document` for TOC dots.
# - **FIXED**: SyntaxError: '(' was never closed in `create_refined_document` for p.add_run call.
# - **FIXED**: Changed Flask app port to 5000 to match frontend expectation.

import os
import io
import re
import json
import requests
import logging
import copy # Ensure copy is imported for deepcopy
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from google.cloud import vision
from dotenv import load_dotenv
import google.generativeai as genai
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill, numbers
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_LEADER
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


load_dotenv()

# --- Logging Configuration ---
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

if not GEMINI_API_KEY:
    logger.error("Error: GEMINI_API_KEY is not set. Cannot call Gemini API.")
else:
    genai.configure(api_key=GEMINI_API_KEY)

vision_client = vision.ImageAnnotatorClient()

# --- Canonical Descriptions Mapping (Existing) ---
CANONICAL_DESCRIPTIONS = {
    "property, plant and equipment": "Property, plant and equipment",
    "pp&e": "Property, plant and equipment",
    "ppe": "Property, plant and equipment",
    "other financial assets": "Other financial assets",
    "trade and other receivables": "Trade and other receivables",
    "trade receivables": "Trade and other receivables",
    "cash and cash equivalents": "Cash and Cash Equivalents",
    "cash equivalents": "Cash and Cash Equivalents",
    "cash at bank": "Cash and Cash Equivalents",
    "total assets": "Total Assets",
    "total asset": "Total Assets",
    "reserves": "Reserves",
    "accumulated surplus": "Accumulated surplus",
    "accum surplus": "Accumulated surplus",
    "accumulated deficit": "Accumulated deficit",
    "accumulated funds": "Accumulated Funds",
    "total equity": "Total Equity",
    "total equ": "Total Equity",
    "deferred tax liability": "Deferred tax liability",
    "tax liability": "Deferred tax liability",
    "trade and other payables": "Trade and Other Payables",
    "trade payables": "Trade and Other Payables",
    "provisions": "Provisions",
    "bank overdraft": "Bank overdraft",
    "total liabilities": "Total Liabilities",
    "total liability": "Total Liabilities",
    "total equity and liabilities": "Total Equity and Liabilities",
    "total equity & liabilities": "Total Equity and Liabilities",
    "total equity and liab": "Total Equity and Liabilities",
    "revenue": "Revenue",
    "levies received": "Levies received",
    "fines": "Fines",
    "water recovered": "Water recovered",
    "ombudsman levy": "Ombudsman levy",
    "electricity recovered": "Electricity recovered",
    "garage levies": "Garage levies",
    "security levies": "Security levies",
    "other income": "Other income",
    "tower rental": "Tower rental",
    "insurance claims received": "Insurance claims received",
    "special levy": "Special levy",
    "interest received": "Interest received",
    "fair value adjustments": "Fair value adjustments",
    "operating expenses": "Operating expenses",
    "accounting fees": "Accounting fees",
    "bank charges": "Bank charges",
    "csos": "CSOS",
    "cleaning": "Cleaning",
    "depreciation, amortisation and impairments": "Depreciation, amortisation and impairments",
    "depreciation and amortisation": "Depreciation, amortisation and impairments",
    "electricity": "Electricity",
    "employee costs": "Employee costs",
    "garden services": "Garden services",
    "insurance": "Insurance",
    "management fees": "Management fees",
    "other expenses": "Other expenses",
    "petrol and oil": "Petrol and oil",
    "printing and stationery": "Printing and stationery",
    "protective clothing": "Protective clothing",
    "repairs and maintenance": "Repairs and maintenance",
    "security": "Security",
    "investment revenue": "Investment revenue",
    "surplus (deficit) for the year": "Surplus (deficit) for the year",
    "total comprehensive income (loss) for the year": "Total comprehensive income (loss) for the year",
    "cash on hand": "Cash on hand",
    "bank balances": "Bank balances",
    "bank balance": "Bank balances",
    "short-term deposits": "Short-term deposits",
    "short term deposits": "Short-term deposits",
    "amounts received in advance": "Amounts received in advance",
    "deposits received": "Deposits received",
    "legal proceedings": "Legal proceedings",
    "rental income": "Rental Income",
    "auditor's remuneration": "Auditor's remuneration",
    "fees": "Auditor's remuneration",
    "bad debts": "Bad debts",
    "consulting and professional fees": "Consulting and professional fees",
    "interest income": "Interest income",
    "csos levies": "CSOS levies",
    "garbage levies": "Garbage levies",

    "compensation commissioner": "Compensation commissioner",
    "employee costs - salaried staff": "Employee costs - salaried staff",
    "municipal charges": "Municipal charges",
    "electricity - recovered from members": "Electricity - recovered from members",
    "water-recovered from members": "Water - recovered from members",
    "maintenance": "Maintenance",
    "elevator maintenace": "Elevator maintenance",
    "total income": "Total Income",
    "total operating expenses": "Total Operating Expenses",
    "(deficit) surplus for the year": "(Deficit) surplus for the year",
    "surplus before taxation": "Surplus before taxation",
    "adjustments for": "Adjustments for",
    "movements in provisions": "Movements in provisions",
    "changes in working capital": "Changes in working capital",
    "net provisions": "Net provisions",
    "non provision of tax": "Non provision of tax",
    "taxation": "Taxation",
    "cash generated from (used in) operations": "Cash generated from (used in) operations",
    "basic": "Basic (Employee Cost)",
    "uif": "UIF (Employee Cost)",
    "absa": "Bank balances"
}

# Defines the structure and order of items in the final Excel report.
MASTER_STRUCTURE = {
    "Assets": {
        "Non-Current Assets": ["Property, plant and equipment", "Other financial assets"],
        "Current Assets": ["Trade and other receivables", "Cash and Cash Equivalents", "Cash on hand", "Bank balances", "Short-term deposits"],
        "N/A": ["Total Assets"]
    },
    "Equity": {
        "Equity": ["Reserves", "Accumulated surplus", "Accumulated deficit", "Accumulated Funds"],
        "N/A": ["Total Equity"]
    },
    "Liabilities": {
        "Non-Current Liabilities": ["Deferred tax liability"],
        "Current Liabilities": ["Trade and Other Payables", "Provisions", "Amounts received in advance", "Deposits received", "Bank overdraft", "Legal proceedings"],
        "N/A": ["Total Liabilities", "Total Equity and Liabilities"]
    },
    "Comprehensive Income": {
        "Revenue": ["Revenue", "Levies received", "Fines", "Water recovered", "Ombudsman levy", "Electricity recovered", "Garage levies", "Security levies", "Other income", "Tower rental", "Insurance claims received", "Special levy", "Interest received", "Fair value adjustments", "Investment revenue", "Rental Income", "Interest income", "CSOS levies", "Garbage levies"],
        "N/A": ["Total Income"],
        "Operating Expenses": ["Operating expenses", "Accounting fees", "Auditor's remuneration", "Bank charges", "CSOS", "Cleaning", "Depreciation, amortisation and impairments", "Electricity", "Employee costs", "Garden services", "Insurance", "Management fees", "Other expenses", "Petrol and oil", "Printing and stationery", "Protective clothing", "Repairs and maintenance", "Security", "Bad debts", "Consulting and professional fees", "Compensation commissioner", "Employee costs - salaried staff", "Municipal charges", "Electricity - recovered from members", "Water - recovered from members", "Maintenance", "Elevator maintenance", "Basic (Employee Cost)", "UIF (Employee Cost)"],
        "N/A_2": ["Total Operating Expenses"] # Using unique keys for N/A sections
    },
    "Profit/Loss": {
        "N/A": [
            "Surplus (deficit) for the year",
            "Total comprehensive income (loss) for the year",
            "(Deficit) surplus for the year",
            "Surplus before taxation",
            "Taxation",
            "Non provision of tax"
        ]
    },
    "Cash Flow": {
        "Adjustments": ["Adjustments for", "Movements in provisions", "Changes in working capital", "Net provisions"],
        "N/A": ["Cash generated from (used in) operations"]
    }
}

# --- Utility Functions ---

def get_canonical_name(name):
    """Finds the canonical name for a given financial term."""
    return CANONICAL_DESCRIPTIONS.get(name.lower().strip(), name)

def clean_value(value):
    """Cleans and converts a string value to a float, handling various formats."""
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        value = value.strip()
        # Handle negative values in parentheses
        is_negative = value.startswith('(') and value.endswith(')')
        # Remove non-numeric characters except for a single decimal point
        value = re.sub(r'[^\d.]', '', value)
        if value:
            try:
                numeric_value = float(value)
                return -numeric_value if is_negative else numeric_value
            except ValueError:
                return None
    return None

# --- Core Processing Functions ---

def extract_text_from_pdf(pdf_content):
    """Extracts text from a PDF file's content using Google Cloud Vision API."""
    if not vision_client:
        raise ConnectionError("Google Cloud Vision client is not initialized.")
    
    logger.info("Starting OCR for PDF content.")
    request = {
        'input_config': {
            'content': pdf_content,
            'mime_type': 'application/pdf'
        },
        'features': [{'type_': vision.Feature.Type.DOCUMENT_TEXT_DETECTION}],
    }

    try:
        response = vision_client.batch_annotate_files(requests=[request])
        full_text = "".join([
            page_response.full_text_annotation.text
            for page_response in response.responses[0].responses
        ])
        logger.info("Finished OCR for PDF content.")
        return re.sub(r'\s+', ' ', full_text).strip()
    except Exception as e:
        logger.error(f"Error during OCR with Google Cloud Vision: {e}")
        raise

def extract_content_from_docx(docx_content):
    """Extracts text, headings, and tables from a DOCX file's content."""
    logger.info("Extracting content from DOCX file.")
    doc = Document(io.BytesIO(docx_content))
    content = {
        'text': "\n".join([p.text for p in doc.paragraphs]),
        'headings': [],
        'tables': []
    }
    for p in doc.paragraphs:
        if p.style and p.style.name.startswith('Heading'):
            try:
                level = int(p.style.name.split(' ')[-1])
                content['headings'].append({'level': level, 'text': p.text})
            except (ValueError, IndexError):
                pass # Ignore styles that don't conform to "Heading X"
    
    for table in doc.tables:
        table_data = [[cell.text for cell in row.cells] for row in table.rows]
        content['tables'].append(table_data)

    logger.info("Finished extracting content from DOCX.")
    return content

def parse_financial_data_with_gemini(text_content, filename, custom_prompt_text=""):
    """Uses Gemini API to parse financial text and return structured JSON."""
    logger.info(f"Sending text from {filename} to Gemini API for financial data parsing.")
    if not GEMINI_API_KEY:
        raise ValueError("GEMINI_API_KEY is not set.")

    model = genai.GenerativeModel('gemini-1.5-flash-latest')

    system_instruction = """
    You are an expert financial data extractor. Your task is to extract all financial line items and their corresponding numerical values from the provided text.
    - The output must be a valid JSON array of objects.
    - Each object must have two keys: "Description" (string) and "AmountsByYear" (an object).
    - The "AmountsByYear" object should have years as keys (e.g., "2023") and numerical values as values.
    - Parse numbers correctly: remove currency symbols, commas, and handle parentheses for negative numbers. Treat spaces as thousand separators (e.g., "1 234" is 1234).
    - Do NOT calculate or infer any values. Only extract what is explicitly present in the text.
    - Extract data from all sections, including main statements and notes.
    - Example output format: [{"Description": "Revenue", "AmountsByYear": {"2023": 500000, "2022": 450000}}]
    """

    prompt = f"""
    {custom_prompt_text}

    Please extract the financial data from the following text:
    ---
    {text_content}
    ---
    """
    
    try:
        response = model.generate_content(
            prompt,
            generation_config={"response_mime_type": "application/json"},
            system_instruction=system_instruction
        )
        
        # Clean the response text to ensure it's valid JSON
        cleaned_text = response.text.strip()
        if cleaned_text.startswith("```json"):
            cleaned_text = cleaned_text[7:]
        if cleaned_text.endswith("```"):
            cleaned_text = cleaned_text[:-3]

        parsed_data = json.loads(cleaned_text)
        
        if not isinstance(parsed_data, list):
            logger.warning(f"Gemini returned non-list data for {filename}: {type(parsed_data)}")
            return []
            
        logger.info(f"Successfully parsed data from {filename} using Gemini.")
        return parsed_data

    except Exception as e:
        logger.error(f"Error during Gemini API call or JSON parsing for {filename}: {e}")
        # Log the raw response if available for debugging
        if 'response' in locals() and hasattr(response, 'text'):
            logger.error(f"Raw Gemini response: {response.text}")
        return []

def generate_excel_report(all_items, all_years):
    """Generates an Excel workbook from the consolidated financial data."""
    logger.info("Generating Excel report...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Financials"

    # --- Define Styles ---
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    category_font = Font(bold=True)
    subcategory_font = Font(bold=True)
    total_font = Font(bold=True)
    total_border = Border(bottom=Side(style='thin'), top=Side(style='thin'))
    currency_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

    # --- Write Headers ---
    headers = ["Description"] + [str(year) for year in all_years]
    ws.append(headers)
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        ws.column_dimensions[get_column_letter(col_idx)].width = 20 if col_idx > 1 else 50
    
    # --- Write Data ---
    row_idx = 2
    for category, subcategories in MASTER_STRUCTURE.items():
        # Write main category header
        ws.cell(row=row_idx, column=1, value=category).font = category_font
        row_idx += 1

        for subcategory, items in subcategories.items():
            # Write subcategory header if it's not 'N/A'
            if subcategory != "N/A":
                ws.cell(row=row_idx, column=1, value=subcategory).font = subcategory_font
                ws.cell(row=row_idx, column=1).alignment = Alignment(indent=1)
                row_idx += 1

            for item_name in items:
                if item_name in all_items:
                    item_data = all_items[item_name]
                    ws.cell(row=row_idx, column=1, value=item_name).alignment = Alignment(indent=2)
                    for col_idx, year in enumerate(all_years, 2):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.value = item_data.get(str(year))
                        cell.number_format = currency_format
                    
                    # Apply total styling for 'Total' rows
                    if "total" in item_name.lower():
                        for col_idx in range(1, len(headers) + 1):
                            ws.cell(row=row_idx, column=col_idx).font = total_font
                            ws.cell(row=row_idx, column=col_idx).border = total_border
                    
                    row_idx += 1
        
        # Add a blank row between major categories for readability
        row_idx += 1

    logger.info("Excel report generated successfully.")
    return wb

# --- Document Refinement Functions ---

async def generate_refinement_instructions_with_gemini(template_content, data_content):
    """Generates instructions for document refinement using Gemini."""
    logger.info("Generating document refinement instructions with Gemini.")
    if not GEMINI_API_KEY:
        raise ValueError("GEMINI_API_KEY is not set.")

    model = genai.GenerativeModel('gemini-1.5-flash-latest')

    template_text_sample = template_content.get('text', '')[:4000]
    data_headings = data_content.get('headings', [])
    data_tables = data_content.get('tables', [])

    data_headings_str = "\n".join([f"Level {h['level']}: {h['text']}" for h in data_headings])
    data_tables_str = "\n---\n".join(["\n".join([" | ".join(map(str, cell)) for cell in table]) for table in data_tables])

    prompt = f"""
    Analyze the provided template and data document extracts. Your task is to extract specific sections from the template and identify corresponding data from the data document to create a refined, merged document.

    **From the TEMPLATE document text, extract:**
    1.  The complete text of the "DEFINITIONS" section.
    2.  The header row of the "DIRECTORS" or "TRUSTEES" table.

    **From the DATA document content, extract:**
    1.  A structured list of all headings to build a new Table of Contents.
    2.  The data rows from the table that corresponds to the directors/trustees.

    Provide the output as a single, valid JSON object with these keys:
    - "template_definitions_text": A string containing the full text of the template's definitions section.
    - "template_director_table_headers": An array of strings for the director table headers from the template.
    - "data_file_headings_for_toc": An array of objects, each with "level" (integer) and "text" (string), for headings from the data document.
    - "data_file_director_table_data": An array of arrays, where each inner array is a row of data for the director's table from the data document.
    - "summary": A brief summary of the extraction.

    If a section is not found, return an empty string or empty array for that key.

    --- TEMPLATE DOCUMENT EXTRACT ---
    {template_text_sample}
    --- END TEMPLATE DOCUMENT EXTRACT ---

    --- DATA DOCUMENT HEADINGS ---
    {data_headings_str}
    --- END DATA DOCUMENT HEADINGS ---

    --- DATA DOCUMENT TABLES ---
    {data_tables_str}
    --- END DATA DOCUMENT TABLES ---
    """
    
    try:
        response = await model.generate_content_async(
            prompt,
            generation_config={"response_mime_type": "application/json"}
        )
        # It's good practice to clean the response as it might be wrapped in markdown
        cleaned_text = response.text.strip().replace("```json", "").replace("```", "")
        logger.debug(f"Raw JSON from Gemini (refinement): {cleaned_text[:500]}...")
        return json.loads(cleaned_text)
    except Exception as e:
        logger.error(f"Error in Gemini refinement instruction generation: {e}")
        # Return a default structure on error
        return {
            "template_definitions_text": "",
            "template_director_table_headers": [],
            "data_file_headings_for_toc": [],
            "data_file_director_table_data": [],
            "summary": f"Error during generation: {e}"
        }

def create_refined_document(template_content, refinement_data):
    """Creates a new DOCX document by merging a template with extracted data."""
    logger.info("Starting creation of the refined document.")
    template_doc = Document(io.BytesIO(template_content))
    new_doc = Document()

    # Crucially, copy styles from the template to the new document
    for style in template_doc.styles:
        # Check if the style type is a paragraph style (type 1) or table style (type 3)
        if style.type == 1 or style.type == 3:
            # Add the style to the new document if it doesn't already exist
            if style.name not in new_doc.styles:
                new_doc.styles.add_style(style.name, style.type)
            
            # Copy relevant properties for paragraph styles
            if style.type == 1:
                # Ensure base_style exists before assigning
                if style.base_style and style.base_style.name in new_doc.styles:
                    new_doc.styles[style.name].base_style = new_doc.styles[style.base_style.name]
                new_doc.styles[style.name].font.name = style.font.name
                new_doc.styles[style.name].font.size = style.font.size
                new_doc.styles[style.name].paragraph_format.alignment = style.paragraph_format.alignment
                new_doc.styles[style.name].paragraph_format.first_line_indent = style.paragraph_format.first_line_indent
                new_doc.styles[style.name].paragraph_format.left_indent = style.paragraph_format.left_indent
                new_doc.styles[style.name].paragraph_format.right_indent = style.paragraph_format.right_indent
                new_doc.styles[style.name].paragraph_format.space_after = style.paragraph_format.space_after
                new_doc.styles[style.name].paragraph_format.space_before = style.paragraph_format.space_before
                new_doc.styles[style.name].paragraph_format.line_spacing = style.paragraph_format.line_spacing
            # For table styles, you might need to copy specific table properties
            elif style.type == 3:
                # This is more complex and often involves copying XML elements directly
                # For now, we'll rely on default table creation and basic styling.
                pass # Placeholder for more advanced table style copying if needed


    # --- Flags and Data ---
    toc_inserted = False
    definitions_inserted = False
    directors_table_inserted = False
    
    data_headings = refinement_data.get("data_file_headings_for_toc", [])
    definitions_text = refinement_data.get("template_definitions_text", "DEFINITIONS SECTION NOT FOUND.")
    director_headers = refinement_data.get("template_director_table_headers", [])
    director_data = refinement_data.get("data_file_director_table_data", [])

    # --- Iterate and Rebuild Document ---
    # These flags are used to skip copying content from the template once a section has been replaced.
    in_toc_section = False
    in_definitions_section = False
    in_directors_table_section = False


    for element in template_doc.element.body:
        # Check if element is a paragraph
        if element.tag.endswith('p'):
            p = Document(element_part=element).paragraphs[0]
            p_text = p.text.strip() # Use p.text directly for the paragraph content

            # Heuristic to detect TOC section in template
            if "CONTENTS" in p_text.upper() and not toc_inserted:
                in_toc_section = True
                toc_inserted = True
                logger.info("Replacing Table of Contents.")
                new_doc.add_heading("CONTENTS", level=1)
                new_doc.add_paragraph("Note: Page numbers in this Table of Contents are placeholders and may require manual update in Microsoft Word after generation (Ctrl+A, then F9).")
                new_doc.add_paragraph("") # Spacer

                # Add TOC entries based on data file headings
                for heading in data_headings:
                    toc_p = new_doc.add_paragraph()
                    # Attempt to apply the style of the original paragraph if it exists
                    if p.style.name in new_doc.styles:
                        toc_p.style = new_doc.styles[p.style.name]
                    else:
                        toc_p.style = 'Normal' # Fallback to Normal style

                    indent_level = max(0, heading.get('level', 1) - 1) # Adjust for 0-based indent
                    toc_p.paragraph_format.left_indent = Inches(0.25 * indent_level)
                    
                    # Add a right-aligned tab stop with a dot leader
                    tab_stops = toc_p.paragraph_format.tab_stops
                    # Clear existing tab stops to avoid duplicates if copying styles
                    tab_stops.clear_all() 
                    tab_stops.add_tab_stop(Inches(6.0), WD_ALIGN_PARAGRAPH.RIGHT, WD_TAB_LEADER.DOTS)
                    
                    toc_p.add_run(heading.get('text', ''))
                    toc_p.add_run('\t') # Add tab character for page number placeholder

                    # Add a page number field (this will show as 0 until updated in Word)
                    run = toc_p.add_run()
                    r = run._r
                    fldChar = OxmlElement('w:fldChar')
                    fldChar.set(qn('w:fldCharType'), 'begin')
                    r.append(fldChar)

                    instrText = OxmlElement('w:instrText')
                    instrText.set(qn('xml:space'), 'preserve')
                    instrText.text = "PAGE"
                    r.append(instrText)

                    fldChar = OxmlElement('w:fldChar')
                    fldChar.set(qn('w:fldCharType'), 'end')
                    fldChar.set(qn('w:dirty'), 'true') # Mark as dirty to force update
                    r.append(fldChar)
                new_doc.add_page_break()
                continue # Skip adding the original paragraph from template

            elif "DEFINITIONS" in p_text.upper() and not definitions_inserted:
                in_definitions_section = True
                definitions_inserted = True
                logger.info("Replacing Definitions section.")
                new_doc.add_heading("DEFINITIONS", level=1)
                new_doc.add_paragraph(definitions_text)
                continue # Skip adding the original paragraph from template

            # Heuristic to detect end of a section (e.g., a new heading starts)
            # This logic needs to be careful not to skip content that should be copied.
            # If a new heading starts, it means the previous "replaced" section has ended.
            if p.pPr and p.pPr.pStyle and p.pPr.pStyle.val.startswith("Heading"):
                in_toc_section = False
                in_definitions_section = False
                in_directors_table_section = False # Reset this too

            # If we are currently inside a section that has been replaced, skip copying its original content.
            if in_toc_section or in_definitions_section or in_directors_table_section:
                continue
            
            # Copy paragraph to new document if not part of a replaced section
            new_doc.element.body.append(copy.deepcopy(element)) # Copy the raw element, not the parsed paragraph

        # Check if element is a table
        elif element.tag.endswith('tbl'):
            tbl = Document(element_part=element).tables[0] # Get the table object
            # Heuristic to detect the directors table by checking the first row's text
            first_row_text = "".join(cell.text for cell in tbl.rows[0].cells).upper()
            if ("DIRECTOR" in first_row_text or "TRUSTEE" in first_row_text) and not directors_table_inserted:
                in_directors_table_section = True
                directors_table_inserted = True
                logger.info("Replacing Directors table.")
                
                # Create new table with data
                if director_headers and director_data:
                    new_table = new_doc.add_table(rows=1, cols=len(director_headers))
                    new_table.style = 'Table Grid' # Apply a default grid style
                    
                    # Set header row
                    hdr_cells = new_table.rows[0].cells
                    for i, header in enumerate(director_headers):
                        hdr_cells[i].text = header
                        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                    # Add data rows
                    for row_data in director_data:
                        row_cells = new_table.add_row().cells
                        for i, cell_data in enumerate(row_data):
                            if i < len(row_cells): # Ensure we don't go out of bounds
                                row_cells[i].text = str(cell_data)
                                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                                row_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                else:
                    new_doc.add_paragraph("Directors/Trustees table data not found or extracted.")
                continue # Skip copying the original table from template

            # If not the directors table and not in a replaced section, copy the table
            if not in_toc_section and not in_definitions_section and not in_directors_table_section:
                new_doc.element.body.append(copy.deepcopy(element))

    # Save the new document to a memory buffer
    doc_io = io.BytesIO()
    new_doc.save(doc_io)
    doc_io.seek(0)
    logger.info("Refined document created successfully.")
    return doc_io

# --- Flask API Routes ---

@app.route('/upload-and-convert', methods=['POST'])
def upload_and_convert_pdfs():
    """
    API endpoint to upload multiple PDF files, process them, and return a consolidated Excel file.
    """
    if 'files' not in request.files:
        return jsonify({"error": "No files part in the request"}), 400

    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        return jsonify({"error": "No selected files"}), 400

    custom_prompt_text = request.form.get('prompt', '')
    all_extracted_data = []
    all_years = set()

    for file in files:
        if file and file.filename.lower().endswith('.pdf'):
            try:
                logger.info(f"Processing file: {file.filename}")
                pdf_content = file.read()
                text_content = extract_text_from_pdf(pdf_content)
                
                if text_content:
                    parsed_data = parse_financial_data_with_gemini(text_content, file.filename, custom_prompt_text)
                    if parsed_data:
                        all_extracted_data.extend(parsed_data)
                        for item in parsed_data:
                            if 'AmountsByYear' in item and isinstance(item['AmountsByYear'], dict):
                                all_years.update(item['AmountsByYear'].keys())
                else:
                    logger.warning(f"Could not extract text from {file.filename}")

            except Exception as e:
                logger.error(f"Failed to process file {file.filename}: {e}")
                return jsonify({"error": f"An error occurred while processing {file.filename}: {str(e)}"}), 500
    
    if not all_extracted_data:
        return jsonify({"error": "Could not extract any financial data from the provided files."}), 400

    # --- Data Consolidation and Cleaning ---
    sorted_years = sorted(list(all_years), reverse=True)
    consolidated_items = {}

    for item in all_extracted_data:
        if 'Description' not in item or 'AmountsByYear' not in item:
            continue
        
        canonical_name = get_canonical_name(item['Description'])
        if canonical_name not in consolidated_items:
            consolidated_items[canonical_name] = {}
        
        for year, value in item['AmountsByYear'].items():
            cleaned_val = clean_value(value)
            if cleaned_val is not None:
                # If the key already exists, sum the values (handles duplicates)
                consolidated_items[canonical_name][year] = consolidated_items[canonical_name].get(year, 0) + cleaned_val

    # Post-processing for Accumulated Surplus/Deficit
    for year in sorted_years:
        surplus = consolidated_items.get("Accumulated surplus", {}).get(year)
        deficit = consolidated_items.get("Accumulated deficit", {}).get(year)
        if surplus is not None and deficit is not None:
            if surplus > 0 and deficit < 0:
                # Both exist, decide which one to keep. Let's assume surplus is the primary.
                # Ensure the key exists before attempting to delete
                if year in consolidated_items["Accumulated deficit"]:
                    del consolidated_items["Accumulated deficit"][year]
            elif deficit != 0:
                consolidated_items["Accumulated surplus"][year] = deficit
                if year in consolidated_items["Accumulated deficit"]:
                    del consolidated_items["Accumulated deficit"][year]


    # --- Excel Generation ---
    try:
        workbook = generate_excel_report(consolidated_items, sorted_years)
        output_io = io.BytesIO()
        workbook.save(output_io)
        output_io.seek(0)
        
        return send_file(
            output_io,
            as_attachment=True,
            download_name='consolidated_financials.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        logger.error(f"Failed to generate Excel file: {e}")
        return jsonify({"error": "Failed to generate the Excel file."}), 500


@app.route('/refine-document', methods=['POST'])
async def refine_document_route():
    """
    API endpoint to refine a document using a template and a data file.
    """
    if 'template_file' not in request.files or 'data_file' not in request.files:
        return jsonify({"error": "Both a template file and a data file are required."}), 400

    template_file = request.files['template_file']
    data_file = request.files['data_file']

    if not template_file.filename.lower().endswith(('.doc', '.docx')):
        return jsonify({"error": "Template file must be a .doc or .docx file."}), 400

    try:
        template_content = template_file.read()
        data_file_content = data_file.read()

        # Extract content from both documents
        parsed_template = extract_content_from_docx(template_content)
        parsed_data = extract_text_from_file(data_file_content, data_file.filename)
        
        # If data file is docx, use its structured content
        if parsed_data.get('type') == 'docx':
            data_to_process = parsed_data.get('content')
        else: # pdf or txt
            data_to_process = {'text': parsed_data.get('text', ''), 'headings': [], 'tables': []}

        # Get refinement instructions from AI
        refinement_instructions = await generate_refinement_instructions_with_gemini(parsed_template, data_to_process)

        if "Error" in refinement_instructions.get('summary', ''):
             return jsonify({"error": f"AI processing failed: {refinement_instructions['summary']}"}), 500

        # Create the new document
        refined_doc_io = create_refined_document(template_content, refinement_instructions)

        return send_file(
            refined_doc_io,
            as_attachment=True,
            download_name='refined_document.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        logger.error(f"An error occurred during document refinement: {e}", exc_info=True)
        return jsonify({"error": f"An internal error occurred: {str(e)}"}), 500


if __name__ == '__main__':
    # Changed port to 5000 to match frontend expectation
    app.run(debug=True, port=5000)