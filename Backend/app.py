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
# - Removed responseSchema from Gemini API call for more flexible parsing
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
# - Added more explicit examples for "Bank balances" and "Short-term deposits" in Gemini prompt.
# - Added "absa" to canonical descriptions and "ABSA" to master structure under Current Assets.
# - Added specific instruction in Gemini prompt for numbers with spaces as thousands separators.
# - **NEW**: Added more explicit examples for "Short-term deposits" for 2020 and 2019 in Gemini prompt.
# - **NEW**: Implemented post-processing logic to ensure "Accumulated deficit" and "Accumulated surplus" are mutually exclusive per year.
# - **NEW**: Integrated Document Refinement Tool with new routes and helper functions.

import os
import io
import re
import json
import requests
import logging # Import logging module
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill, numbers # Import numbers for formatting
from openpyxl.utils import get_column_letter
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from google.cloud import vision
from dotenv import load_dotenv
import google.generativeai as genai # NEW: Import google.generativeai
from docx import Document # NEW: Import Document for docx creation
from docx.shared import Inches # NEW: Import Inches for docx images (if needed later)
import pandas as pd # Already present, good for Excel/CSV parsing if needed for refinement tool

load_dotenv()

# --- Logging Configuration ---
# Set logging level to DEBUG to see all detailed messages
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

# NEW: Configure Gemini API
if not GEMINI_API_KEY:
    logger.error("Error: GEMINI_API_KEY is not set. Cannot call Gemini API.")
    # You can uncomment the line below for local testing if you prefer to hardcode it,
    # but it's not recommended for production.
    # GEMINI_API_KEY = "YOUR_GEMINI_API_KEY_HERE"
else:
    genai.configure(api_key=GEMINI_API_KEY)


vision_client = vision.ImageAnnotatorClient()

# --- Canonical Descriptions Mapping (Existing) ---
# This dictionary maps various ways a description might appear (from OCR or Gemini)
# to a single, consistent canonical form.
CANONICAL_DESCRIPTIONS = {
    "property, plant and equipment": "Property, plant and equipment",
    "pp&e": "Property, plant and equipment",
    "ppe": "Property, plant and equipment",
    "other financial assets": "Other financial assets",
    "trade and other receivables": "Trade and other receivables",
    "trade receivables": "Trade and other receivables",
    "cash and cash equivalents": "Cash and Cash Equivalents", # Updated for consistency
    "cash equivalents": "Cash and Cash Equivalents",
    "cash at bank": "Cash and Cash Equivalents", # Added alias
    "total assets": "Total Assets",
    "total asset": "Total Assets", # Singular form
    "reserves": "Reserves",
    "accumulated surplus": "Accumulated surplus",
    "accum surplus": "Accumulated surplus",
    "accumulated deficit": "Accumulated deficit", # Added for deficit case
    "accumulated funds": "Accumulated Funds", # Added for consistency
    "total equity": "Total Equity",
    "total equ": "Total Equity",
    "deferred tax liability": "Deferred tax liability",
    "tax liability": "Deferred tax liability",
    "trade and other payables": "Trade and Other Payables", # Updated for consistency
    "trade payables": "Trade and Other Payables",
    "provisions": "Provisions",
    "bank overdraft": "Bank overdraft",
    "total liabilities": "Total Liabilities",
    "total liability": "Total Liabilities", # Singular form
    "total equity and liabilities": "Total Equity and Liabilities",
    "total equity & liabilities": "Total Equity and Liabilities",
    "total equity and liab": "Total Equity and Liabilities",
    # Additional items often found in notes or other statements that might be relevant
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
    "accounting fees": "Accounting fees", # Explicitly added
    "bank charges": "Bank charges",
    "csos": "CSOS",
    "cleaning": "Cleaning",
    "depreciation, amortisation and impairments": "Depreciation, amortisation and impairments",
    "depreciation and amortisation": "Depreciation, amortisation and impairments", # Alias
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
    "bank balance": "Bank balances", # Added alias
    "short-term deposits": "Short-term deposits",
    "short term deposits": "Short-term deposits", # Added alias
    "amounts received in advance": "Amounts received in advance",
    "deposits received": "Deposits received",
    "legal proceedings": "Legal proceedings",
    "rental income": "Rental Income",
    "auditor's remuneration": "Auditor's remuneration",
    "fees": "Auditor's remuneration", # Alias
    "bad debts": "Bad debts",
    "consulting and professional fees": "Consulting and professional fees",
    "interest income": "Interest income",
    "csos levies": "CSOS levies",
    "garbage levies": "Garbage levies", # Typo from PDF "Garage levies"
    "compensation commissioner": "Compensation commissioner",
    "employee costs - salaried staff": "Employee costs - salaried staff",
    "municipal charges": "Municipal charges",
    "electricity - recovered from members": "Electricity - recovered from members",
    "water-recovered from members": "Water - recovered from members",
    "maintenance": "Maintenance",
    "elevator maintenace": "Elevator maintenance", # Typo from PDF
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
    "basic": "Basic (Employee Cost)", # Clarify basic
    "uif": "UIF (Employee Cost)", # Clarify UIF
    "absa": "Bank balances" # NEW: Added alias for ABSA
}


# Master structure for ordered output in Excel, using canonical descriptions
MASTER_STRUCTURE = {
    "Assets": {
        "Non-Current Assets": ["Property, plant and equipment", "Other financial assets"],
        "Current Assets": ["Trade and other receivables", "Cash and Cash Equivalents", "Cash on hand", "Bank balances", "Short-term deposits", "ABSA"], # NEW: Added ABSA here
        "N/A": ["Total Assets"]
    },
    "Equity": {
        "Equity": ["Reserves", "Accumulated surplus", "Accumulated deficit", "Accumulated Funds"], # Added Accumulated Funds
        "N/A": ["Total Equity"] # Total Equity is a direct line item in some PDFs
    },
    "Liabilities": {
        "Non-Current Liabilities": ["Deferred tax liability"],
        "Current Liabilities": ["Trade and Other Payables", "Provisions", "Amounts received in advance", "Deposits received", "Bank overdraft", "Legal proceedings"], # Updated canonical name
        "N/A": ["Total Liabilities"]
    },
    "Revenue": { # New category for Statement of Comprehensive Income items
        "Income": ["Revenue", "Levies received", "Fines", "Water recovered", "Ombudsman levy", "Electricity recovered", "Garage levies", "Security levies", "Other income", "Tower rental", "Insurance claims received", "Special levy", "Interest received", "Fair value adjustments", "Investment revenue", "Rental Income", "Interest income", "CSOS levies", "Garbage levies"],
        "N/A": ["Total Income"] # Added Total Income here
    },
    "Expenses": { # New category for Statement of Comprehensive Income items
        "Operating Expenses": ["Operating expenses", "Accounting fees", "Bank charges", "CSOS", "Cleaning", "Depreciation, amortisation and impairments", "Electricity", "Employee costs", "Garden services", "Insurance", "Management fees", "Other expenses", "Petrol and oil", "Printing and stationery", "Protective clothing", "Repairs and maintenance", "Security", "Auditor's remuneration", "Bad debts", "Consulting and professional fees", "Compensation commissioner", "Employee costs - salaried staff", "Municipal charges", "Electricity - recovered from members", "Water - recovered from members", "Maintenance", "Elevator maintenance", "Basic (Employee Cost)", "UIF (Employee Cost)"],
        "N/A": ["Total Operating Expenses"] # Added Total Operating Expenses here
    },
    "Other Financial Items": { # For items that don't fit neatly above but are financial
        "N/A": ["Surplus (deficit) for the year", "Total comprehensive income (loss) for the year", "Surplus before taxation", "Adjustments for", "Movements in provisions", "Changes in working capital", "Net provisions", "Non provision of tax", "Taxation", "Cash generated from (used in) operations", "(Deficit) surplus for the year"]
    },
    # Removed "Summary" category from MASTER_STRUCTURE as it's no longer needed for display in the main sheet
}


# Helper to extract year range from filename
def extract_years_from_filename(filename):
    """
    Extracts 4-digit years from a filename and returns them sorted in descending order.
    """
    match = re.findall(r'(\d{4})', filename)
    # Ensure all matched years are converted to int and then sorted
    return sorted([int(y) for y in match], reverse=True)

# Preprocess OCR text
def normalize_text(text):
    """
    Normalizes OCR text by replacing multiple spaces with a single space and stripping leading/trailing whitespace.
    """
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

# OCR
def detect_text_from_pdf(pdf_content): # Changed to accept content directly
    """
    Performs OCR on PDF content using Google Cloud Vision API.
    It processes the PDF page by page and returns all detected text.
    """
    logger.info("Starting OCR for PDF content.")
    mime_type = 'application/pdf'

    input_config_content = vision.InputConfig(
        mime_type=mime_type,
        content=pdf_content
    )

    feature = vision.Feature(
        type_=vision.Feature.Type.DOCUMENT_TEXT_DETECTION
    )

    request = vision.AnnotateFileRequest(
        input_config=input_config_content,
        features=[feature]
    )

    try:
        response = vision_client.batch_annotate_files(requests=[request])

        full_text = ""
        for image_response in response.responses[0].responses:
            if image_response.full_text_annotation:
                full_text += image_response.full_text_annotation.text + "\n"

        logger.info("Finished OCR for PDF content.")
        return normalize_text(full_text) # Use normalize_text here
    except Exception as e:
        logger.error(f"Error during OCR with Google Cloud Vision: {e}")
        return None


# NEW: General text extraction function
def extract_text_from_file(file_content, filename):
    """
    Extracts text from various file types.
    Supports PDF (via OCR), TXT, and attempts basic decoding for DOCX, XLSX, CSV.
    """
    logger.info(f"Attempting to extract text from {filename}...")
    if filename.lower().endswith('.pdf'):
        return detect_text_from_pdf(file_content)
    elif filename.lower().endswith('.txt'):
        try:
            return file_content.decode('utf-8')
        except UnicodeDecodeError:
            return file_content.decode('latin-1') # Fallback
    elif filename.lower().endswith(('.doc', '.docx')):
        # For .doc/.docx, a dedicated library like python-docx is needed for proper parsing.
        # This is a basic attempt to read it as text, which might fail or produce garbage.
        logger.warning(f"Full text extraction for {filename} (DOC/DOCX) requires 'python-docx' library and more complex parsing logic. Attempting basic text decode.")
        try:
            return file_content.decode('utf-8')
        except UnicodeDecodeError:
            return file_content.decode('latin-1')
    elif filename.lower().endswith(('.xlsx', '.csv')):
        # For .xlsx/.csv, pandas can read structured data, but direct text extraction is not straightforward.
        # This is a basic attempt to read it as text, which might fail or produce garbage.
        logger.warning(f"Full text extraction for {filename} (XLSX/CSV) requires 'pandas' or 'openpyxl' and specific parsing logic. Attempting basic text decode.")
        try:
            return file_content.decode('utf-8')
        except UnicodeDecodeError:
            return file_content.decode('latin-1')
    else:
        logger.warning(f"Unsupported file type for text extraction: {filename}")
        return None


# Gemini API (Existing parse_with_gemini for PDF to Excel)
def parse_with_gemini(text_content, filename, custom_prompt_text=None):
    """
    Sends the OCR'd text content to the Gemini API for structured data extraction.
    Integrates a custom prompt as additional instructions within the default structured prompt.
    """
    logger.info("Sending text from %s to Gemini API for parsing financial data...", filename)

    if not GEMINI_API_KEY:
        logger.error("Error: GEMINI_API_KEY is not set. Cannot call Gemini API.")
        return []

    # Base prompt structure that ensures JSON output with Description and AmountsByYear
    # Use double curly braces {{ }} for literal curly braces that are part of the JSON example,
    # so they are not misinterpreted as format placeholders.
    base_prompt_template = """
Extract all financial line items and their corresponding numerical values from the entire provided financial statement document.
This document may contain a Statement of Financial Position, Statement of Comprehensive Income, and Notes to the Financial Statements.
For each financial line item that has associated numerical values, provide its "Description" and an "AmountsByYear" object where keys are the year (e.g., "2020", "2019") and values are the corresponding numerical amounts.

Ensure amounts are parsed as numbers (floats or integers).
Handle negative values correctly (e.g., if in parentheses, convert to negative number).
Crucially, only extract line items and their values as they explicitly appear in the document. Do NOT perform any calculations, aggregations, or infer new totals if they are not explicitly present.
Include all granular line items with numerical values, even if they appear in the 'Notes to the Financial Statements' or 'Statement of Comprehensive Income', not just the main Statement of Financial Position.

IMPORTANT: When extracting numbers, treat spaces as thousands separators, not decimal points. For example, "1 234" should be parsed as 1234.00, not 1.234. Only a period (.) should be considered a decimal point.

{custom_instructions}

Provide the output as a JSON array of objects, like this example:
[
    {{"Description": "Property, plant and equipment", "AmountsByYear": {{"2022": 1550000.00, "2021": 1206.00, "2020": 1808.00, "2019": 2410.00}}}},
    {{"Description": "Other financial assets", "AmountsByYear": {{"2022": 276005.00, "2021": 265160.00, "2020": 231407.00, "2019": 231194.00}}}},
    {{"Description": "Trade and other receivables", "AmountsByYear": {{"2022": 479579.00, "2021": 2873397.00, "2020": 2485083.00, "2019": 2074399.00}}}},
    {{"Description": "Cash and Cash Equivalents", "AmountsByYear": {{"2022": 1903085.00, "2021": 327023.00, "2020": 538721.00, "2019": 455864.00}}}},
    {{"Description": "Cash on hand", "AmountsByYear": {{"2022": 345.00, "2021": 345.00, "2020": 345.00, "2019": 345.00}}}},
    {{"Description": "Bank balances", "AmountsByYear": {{"2022": 386098.00, "2021": 307901.00, "2020": 538376.00, "2019": 452857.00}}}},
    {{"Description": "Short-term deposits", "AmountsByYear": {{"2022": 19252.00, "2021": 18777.00, "2020": 19252.00, "2019": 18777.00}}}},
    {{"Description": "Total Assets", "AmountsByYear": {{"2022": 2788359.00, "2021": 3466786.00, "2020": 3257019.00, "2019": 2763867.00}}}},
    {{"Description": "Reserves", "AmountsByYear": {{"2022": 300000.00, "2021": 300000.00, "2020": 300000.00, "2019": 300000.00}}}},
    {{"Description": "Accumulated surplus", "AmountsByYear": {{"2022": -89539.00, "2021": 1730889.00, "2020": 1134806.00, "2019": 1042255.00}}}},
    {{"Description": "Trade and Other Payables", "AmountsByYear": {{"2022": 2302524.00, "2021": 1165346.00, "2020": 1813271.00, "2019": 1421612.00}}}},
    {{"Description": "Provisions", "AmountsByYear": {{"2022": 11319.00, "2021": 8996.00, "2020": 8942.00}}}},
    {{"Description": "Revenue", "AmountsByYear": {{"2022": 2118462.00, "2021": 1551014.00, "2020": 2864966.00, "2019": 2490642.00}}}},
    {{"Description": "Operating expenses", "AmountsByYear": {{"2022": -2000000.00, "2021": -188662.00, "2020": -3019923.00, "2019": -3132967.00}}}}
]

--- START FINANCIAL STATEMENT TEXT ---
{text_content}
--- END FINANCIAL STATEMENT TEXT ---
    """

    # Insert custom instructions if provided
    custom_instructions_placeholder = ""
    if custom_prompt_text:
        custom_instructions_placeholder = f"\n\nAdditional instructions from user for parsing: {custom_prompt_text}\n"

    final_prompt = base_prompt_template.format(
        custom_instructions=custom_instructions_placeholder,
        text_content=text_content
    )

    api_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={GEMINI_API_KEY}"

    headers = {
        'Content-Type': 'application/json'
    }

    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {"text": final_prompt}
                ]
            }
        ],
        "generationConfig": {
            "responseMimeType": "application/json"
        }
    }

    try:
        response = requests.post(api_url, headers=headers, data=json.dumps(payload))
        response.raise_for_status()
        gemini_result = response.json()

        if gemini_result and gemini_result.get('candidates') and gemini_result['candidates'][0].get('content') and gemini_result['candidates'][0]['content'].get('parts'):
            json_string = gemini_result['candidates'][0]['content']['parts'][0]['text']
            logger.debug("Raw JSON string from Gemini API for %s: %s...", filename, json_string[:1000]) # Increased log length
            parsed_data = json.loads(json_string)

            if not isinstance(parsed_data, list):
                logger.warning("Unexpected format from Gemini for %s. Expected a list but got %s. Skipping.", filename, type(parsed_data))
                return []

            logger.info("Successfully parsed data from %s with Gemini API.", filename)
            return parsed_data
        else:
            logger.warning("Gemini API returned an unexpected response for %s: %s", filename, gemini_result)
            if gemini_result.get('promptFeedback'):
                logger.warning("Gemini Prompt Feedback: %s", gemini_result['promptFeedback'])
            if gemini_result.get('candidates') and gemini_result['candidates'][0].get('finishReason'):
                logger.warning("Gemini Finish Reason: %s", gemini_result['candidates'][0]['finishReason'])
            return []
    except requests.exceptions.RequestException as e:
        logger.error("Error calling Gemini API for %s: %s", filename, e)
        if hasattr(e, 'response') and e.response is not None:
            logger.error("Gemini API Response Text (Error): %s", e.response.text)
        return []
    except json.JSONDecodeError as e:
        logger.error("Error decoding JSON from Gemini API response for %s: %s", filename, e)
        if 'response' in locals() and hasattr(response, 'text'):
            logger.error("Raw Gemini response text that failed to decode: %s", response.text)
        return []

# NEW: Gemini API for Document Refinement
async def generate_refinement_instructions_with_gemini(template_text, data_text):
    """
    Generates mapping instructions for document refinement using Gemini AI.
    """
    logger.info("Sending template and data text to Gemini API for refinement instructions...")

    if not GEMINI_API_KEY:
        logger.error("Error: GEMINI_API_KEY is not set. Cannot call Gemini API for refinement.")
        return {"mapping_instructions": [], "summary": "API Key not configured."}

    model = genai.GenerativeModel('gemini-2.0-flash')
    prompt = f"""
    You are an AI assistant specialized in document transformation.
    Given a 'template document' and a 'data document', your goal is to explain
    how to map the relevant information from the 'data document' into the structure
    of the 'template document'.

    Analyze the structure and key fields in the Template Document.
    Analyze the content and key data points in the Data Document.
    Identify potential fields in the Template Document that could be populated by data from the Data Document.

    Template Document Content (first 2000 characters):
    ---
    {template_text[:2000]}
    ---

    Data Document Content (first 2000 characters):
    ---
    {data_text[:2000]}
    ---

    Please provide a JSON object with a 'mapping_instructions' key.
    The value should be a list of objects, where each object describes a mapping.
    Each mapping object should have:
    - 'template_field': A description or key from the template where data should go (e.g., "Invoice Number", "Client Name", "Total Amount").
    - 'data_source': A description or key from the data document where the value comes from (e.g., "Invoice ID from Data", "Customer Name from Data File", "Sum of Line Items").
    - 'example_value': An example of the value from the data document that would be mapped.
    - 'notes': Any specific instructions for inserting or formatting this data (e.g., "Insert directly", "Format as currency", "Extract date only").

    If you cannot find clear mappings, indicate that in the 'summary'.
    Example JSON structure:
    {{
        "mapping_instructions": [
            {{
                "template_field": "Customer Name (from template)",
                "data_source": "Client Name (from data file)",
                "example_value": "John Doe",
                "notes": "Insert directly"
            }},
            {{
                "template_field": "Invoice Total (from template)",
                "data_source": "Total Amount (from data file)",
                "example_value": "123.45",
                "notes": "Ensure currency format"
            }}
        ],
        "summary": "Overall mapping strategy or findings."
    }}
    """
    try:
        response = await model.generate_content_async(
            prompt,
            generation_config={
                "response_mime_type": "application/json",
                "response_schema": {
                    "type": "OBJECT",
                    "properties": {
                        "mapping_instructions": {
                            "type": "ARRAY",
                            "items": {
                                "type": "OBJECT",
                                "properties": {
                                    "template_field": {"type": "STRING"},
                                    "data_source": {"type": "STRING"},
                                    "example_value": {"type": "STRING"},
                                    "notes": {"type": "STRING"}
                                }
                            }
                        },
                        "summary": {"type": "STRING"}
                    }
                }
            }
        )
        json_string = response.candidates[0].content.parts[0].text
        logger.debug(f"Raw JSON string from Gemini (refinement): {json_string[:500]}...")
        return json.loads(json_string)
    except Exception as e:
        logger.error(f"Error generating refinement instructions with Gemini: {e}")
        return {"mapping_instructions": [], "summary": f"Error: {e}"}

# NEW: Function to create refined document (placeholder)
def create_refined_document(template_filename, mapping_instructions):
    """
    Creates a new document based on the template and inserts data based on mappings.
    This is a highly simplified example. A real implementation would need
    to parse the template structure (e.g., identifying placeholders) and
    then insert data. For now, it just creates a simple docx with the instructions.
    """
    logger.info(f"Creating refined document based on template '{template_filename}'...")
    doc = Document()
    doc.add_heading(f'Refined Document Output for {template_filename}', level=1)
    doc.add_paragraph('This document is a placeholder for the refined output.')
    doc.add_paragraph('Below are the AI-generated mapping instructions:')

    if mapping_instructions and mapping_instructions.get("mapping_instructions"):
        for item in mapping_instructions["mapping_instructions"]:
            doc.add_paragraph(f"Template Field: {item.get('template_field', 'N/A')}")
            doc.add_paragraph(f"    Data Source: {item.get('data_source', 'N/A')}")
            doc.add_paragraph(f"    Example Value: {item.get('example_value', 'N/A')}")
            doc.add_paragraph(f"    Notes: {item.get('notes', 'N/A')}")
            doc.add_paragraph("") # Add a blank line for separation
        doc.add_paragraph(f"Summary: {mapping_instructions.get('summary', 'No summary provided.')}")
    else:
        doc.add_paragraph("No specific mapping instructions were generated by the AI.")
        doc.add_paragraph(f"AI Summary: {mapping_instructions.get('summary', 'N/A')}")

    # Save to a BytesIO object
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    logger.info("Refined document (placeholder) generated in memory.")
    return bio


# --- Helper Function for Post-processing and Category Inference (Existing) ---
def infer_categories_and_structure(parsed_items):
    """
    Infers Category and SubCategory for each item based on description keywords and common financial statement structure.
    Returns items with added 'Category' and 'SubCategory' fields.
    Also applies canonical descriptions.
    """
    structured_items = []

    # These keywords are still useful for internal categorization even if not explicitly displayed in Excel
    category_keywords = {
        "Assets": ["assets"],
        "Equity": ["equity", "reserves", "accumulated surplus", "accumulated deficit", "accumulated funds"],
        "Liabilities": ["liabilities", "trade and other payables", "provisions", "bank overdraft", "deferred tax liability"],
        "Revenue": ["revenue", "income", "levies received", "fines", "water recovered", "ombudsman levy", "electricity recovered", "garage levies", "security levies", "tower rental", "insurance claims received", "special levy", "interest received", "fair value adjustments", "investment revenue", "rental income", "interest income", "csos levies", "garbage levies"],
        "Expenses": ["operating expenses", "accounting fees", "bank charges", "csos", "cleaning", "depreciation", "amortisation", "impairments", "electricity", "employee costs", "garden services", "insurance", "management fees", "other expenses", "petrol and oil", "printing and stationery", "protective clothing", "repairs and maintenance", "security", "auditor's remuneration", "bad debts", "consulting and professional fees", "compensation commissioner", "employee costs - salaried staff", "municipal charges", "electricity - recovered from members", "water-recovered from members", "maintenance", "elevator maintenance"],
        "Other Financial Items": ["surplus (deficit) for the year", "total comprehensive income (loss) for the year", "surplus before taxation", "adjustments for", "movements in provisions", "changes in working capital", "net provisions", "non provision of tax", "taxation", "cash generated from (used in) operations", "basic", "uif", "deficit surplus for the year"],
    }

    subcategory_keywords = {
        "Non-Current Assets": ["non-current assets", "property, plant and equipment", "other financial assets"],
        "Current Assets": ["current assets", "trade and other receivables", "cash and cash equivalents", "cash on hand", "bank balances", "short-term deposits", "absa"], # Added 'absa' here
        "Equity": ["reserves", "accumulated surplus", "accumulated deficit", "accumulated funds"],
        "Non-Current Liabilities": ["non-current liabilities", "deferred tax liability"],
        "Current Liabilities": ["current liabilities", "trade and other payables", "provisions", "amounts received in advance", "deposits received", "bank overdraft", "legal proceedings"],
        "Income": ["revenue", "levies received", "fines", "water recovered", "ombudsman levy", "electricity recovered", "garage levies", "security levies", "other income", "tower rental", "insurance claims received", "special levy", "interest received", "fair value gains", "investment revenue", "rental income", "interest income", "csos levies", "garbage levies"],
        "Operating Expenses": ["operating expenses", "accounting fees", "bank charges", "csos", "cleaning", "depreciation", "amortisation", "impairments", "electricity", "employee costs", "garden services", "insurance", "management fees", "other expenses", "petrol and oil", "printing and stationery", "protective clothing", "repairs and maintenance", "security", "auditor's remuneration", "bad debts", "consulting and professional fees", "compensation commissioner", "municipal charges", "maintenance", "elevator maintenance", "basic", "uif"],
        "N/A": [] # For top-level items within a category or summary items
    }

    last_major_section = None

    for item in parsed_items:
        original_description = item.get("Description", "")
        # Normalize and get canonical description
        normalized_desc = normalize_text(original_description).lower()
        canonical_description = CANONICAL_DESCRIPTIONS.get(normalized_desc, original_description)

        # Update the item's description to its canonical form
        item["Description"] = canonical_description

        inferred_category = "Other Financial Items" # Default to a more general category
        inferred_subcategory = "N/A"

        # First, try to identify major categories
        for cat, keywords in category_keywords.items():
            if any(keyword in normalized_desc for keyword in keywords):
                inferred_category = cat
                last_major_section = cat
                break

        # If no explicit category found, use the last major section as a hint
        if inferred_category == "Other Financial Items" and last_major_section:
            inferred_category = last_major_section

        # Then, try to identify subcategories
        for subcat, keywords in subcategory_keywords.items():
            if any(keyword in normalized_desc for keyword in keywords):
                # Refine subcategory based on the main category context
                if (subcat in ["Non-Current Assets", "Current Assets"] and inferred_category == "Assets") or \
                   (subcat == "Equity" and inferred_category == "Equity") or \
                   (subcat in ["Non-Current Liabilities", "Current Liabilities"] and inferred_category == "Liabilities") or \
                   (subcat == "Income" and inferred_category == "Revenue") or \
                   (subcat == "Operating Expenses" and inferred_category == "Expenses"):
                    inferred_subcategory = subcat
                    break
                # Fallback for common items if direct subcategory match isn't found but category is known
                elif inferred_category == "Assets" and ("receivables" in normalized_desc or "cash" in normalized_desc or "bank" in normalized_desc or "deposits" in normalized_desc or "absa" in normalized_desc): # Added 'absa' here
                    inferred_subcategory = "Current Assets"
                elif inferred_category == "Liabilities" and ("payables" in normalized_desc or "provisions" in normalized_desc or "advance" in normalized_desc or "deposits received" in normalized_desc or "legal" in normalized_desc):
                    inferred_subcategory = "Current Liabilities"
                elif inferred_category == "Equity" and ("reserves" in normalized_desc or "surplus" in normalized_desc or "deficit" in normalized_desc or "funds" in normalized_desc): # Added "funds"
                    inferred_subcategory = "Equity"
                elif inferred_category == "Revenue" and ("levies" in normalized_desc or "income" in normalized_desc or "recovered" in normalized_desc or "fines" in normalized_desc or "rental" in normalized_desc):
                    inferred_subcategory = "Income"
                elif inferred_category == "Expenses" and ("fees" in normalized_desc or "charges" in normalized_desc or "depreciation" in normalized_desc or "employee" in normalized_desc or "maintenance" in normalized_desc or "security" in normalized_desc):
                    inferred_subcategory = "Operating Expenses"

        # Specific overrides for common total lines that might not fit subcategories
        if "total assets" in normalized_desc:
            inferred_category = "Assets"
            inferred_subcategory = "N/A"
        elif "total equity" in normalized_desc:
            inferred_category = "Equity"
            inferred_subcategory = "N/A"
        elif "total liabilities" in normalized_desc:
            inferred_category = "Liabilities"
            inferred_subcategory = "N/A"
        elif "total equity and liabilities" in normalized_desc:
            inferred_category = "Other Financial Items" # Categorize this as Other Financial Items
            inferred_subcategory = "N/A"
        elif "total income" in normalized_desc:
            inferred_category = "Revenue"
            inferred_subcategory = "N/A"
        elif "total operating expenses" in normalized_desc:
            inferred_category = "Expenses"
            inferred_subcategory = "N/A"

        item["Category"] = inferred_category
        item["SubCategory"] = inferred_subcategory.replace(" and other", "").replace(" and cash", "")

        structured_items.append(item)

    return structured_items


# --- Helper Function for Excel Generation (Refined for Hierarchical Output) (Existing) ---
def generate_excel(client_name, all_parsed_data):
    """
    Generates a single Excel workbook from the parsed data of multiple PDFs.
    It consolidates all data into one main sheet with a hierarchical structure
    and year-specific columns.
    Returns a BytesIO object containing the Excel file.
    """
    logger.info("Starting Excel generation for client: %s", client_name)
    wb = Workbook()

    # Define styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid") # Indigo
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_aligned_text = Alignment(horizontal="center", vertical="center")
    category_fill = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid") # Light indigo for categories
    subcategory_fill = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid") # Lighter indigo for subcategories
    bold_font = Font(bold=True)
    indent_alignment = Alignment(indent=1) # For descriptions under subcategories

    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # --- Consolidated Data Sheet ---
    main_ws = wb.create_sheet(f"{client_name} - Consolidated Data")

    all_years = set()

    # Apply category inference to all parsed data and collect all years
    processed_all_parsed_data = {}
    for filename, items in all_parsed_data.items():
        processed_items = infer_categories_and_structure(items)
        processed_all_parsed_data[filename] = processed_items

        for item in processed_items:
            if "AmountsByYear" in item and isinstance(item["AmountsByYear"], dict):
                for year_str in item["AmountsByYear"].keys():
                    try:
                        all_years.add(int(year_str))
                    except ValueError:
                        logger.warning("Non-numeric year found: %s in %s", year_str, filename)
                        pass

    sorted_years = sorted(list(all_years), reverse=True) # Sort years descending
    logger.debug(f"Detected and sorted years: {sorted_years}")

    # Main title - merge across Description column + all year columns
    main_ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + len(sorted_years))
    title_cell = main_ws['A1']
    title_cell.value = f"Consolidated Financial Statements for {client_name}"
    title_cell.font = Font(size=16, bold=True, color="4F46E5")
    title_cell.alignment = center_aligned_text
    main_ws.row_dimensions[1].height = 30
    main_ws.append([]) # Spacer

    # Prepare main sheet headers: Description, then years
    main_headers = ["Description"] + [str(year) for year in sorted_years]
    main_ws.append(main_headers)
    for col_num, cell in enumerate(main_ws[3]): # Headers are in row 3
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = center_aligned_text
        main_ws.column_dimensions[get_column_letter(col_num + 1)].width = 20
    main_ws.column_dimensions['A'].width = 40 # Description column width


    # Consolidate data for the main sheet
    # Key: (Canonical Description, Category, SubCategory), Value: {year: amount}
    consolidated_items_by_key = {}

    # Pre-initialize all descriptions from MASTER_STRUCTURE to ensure they appear, even if Gemini misses them
    # The key uses the canonical description and the category/subcategory as defined in MASTER_STRUCTURE
    for category_name, subcats_dict in MASTER_STRUCTURE.items():
        for subcat_name, descs_list in subcats_dict.items():
            for desc in descs_list:
                key = (desc, category_name, subcat_name)
                consolidated_items_by_key[key] = {year: None for year in sorted_years}
    logger.debug(f"Pre-initialized consolidated_items_by_key with {len(consolidated_items_by_key)} entries.")


    # Populate consolidated_items_by_key with actual parsed data, overwriting None values
    for filename, items in processed_all_parsed_data.items():
        logger.debug(f"Processing parsed items from file: {filename}")
        for item in items:
            description = item.get("Description") # This is already canonical from infer_categories_and_structure
            amounts_by_year = item.get("AmountsByYear", {})

            if description:
                # Attempt to find the best matching key in consolidated_items_by_key based on MASTER_STRUCTURE
                target_key = None

                # Priority 1: Find by canonical description within MASTER_STRUCTURE
                # This ensures that if a description is defined in MASTER_STRUCTURE,
                # we use its predefined category/subcategory, overriding Gemini's inference if it differs.
                for master_cat, master_subcats in MASTER_STRUCTURE.items():
                    for master_subcat, master_descs in master_subcats.items():
                        if description in master_descs:
                            target_key = (description, master_cat, master_subcat)
                            break
                    if target_key:
                        break

                if target_key is None:
                    # If description is still not found in MASTER_STRUCTURE,
                    # use the inferred category/subcategory from infer_categories_and_structure
                    # and add it as an "Other" item.
                    inferred_cat = item.get("Category", "Other Financial Items")
                    inferred_subcat = item.get("SubCategory", "N/A")
                    target_key = (description, inferred_cat, inferred_subcat)

                    if target_key not in consolidated_items_by_key:
                        logger.warning("Gemini parsed an item not explicitly in MASTER_STRUCTURE: '%s' (Inferred Cat: '%s', SubCat: '%s'). Adding it.", description, inferred_cat, inferred_subcat)
                        consolidated_items_by_key[target_key] = {year: None for year in sorted_years}

                # Update amounts for the determined target_key
                for year_str, amount in amounts_by_year.items():
                    try:
                        year_int = int(year_str)
                        # Clean amount string: remove all non-digit, non-decimal, non-minus characters first,
                        # then handle parentheses. This is more aggressive and should catch more variations.
                        clean_amount_str = str(amount).strip()

                        # Handle parentheses for negative numbers first
                        if clean_amount_str.startswith('(') and clean_amount_str.endswith(')'):
                            clean_amount_str = '-' + clean_amount_str[1:-1]

                        # Remove all non-numeric characters except for a single decimal point and leading minus sign
                        # This should handle spaces, commas, and other symbols, including non-breaking spaces (\xa0)
                        clean_amount_str = re.sub(r'[^\d.-]+', '', clean_amount_str.replace('\xa0', ''))

                        # Ensure only one decimal point
                        if clean_amount_str.count('.') > 1:
                            parts = clean_amount_str.split('.')
                            clean_amount_str = parts[0] + '.' + ''.join(parts[1:])

                        clean_amount = float(clean_amount_str) if clean_amount_str else None

                        # Ensure we don't overwrite with None if a value already exists
                        # This logic is crucial for combining data from multiple PDFs
                        if clean_amount is not None:
                            # If there's an existing value and the new one is not None,
                            # we need to decide whether to overwrite or sum.
                            # For financial statements, usually the latest or most specific value is preferred,
                            # or if it's truly a duplicate from different sources, we might need a more complex rule.
                            # For now, we'll overwrite, assuming later PDFs might have more complete data or
                            # that the AI should ideally provide the single correct value.
                            # If the issue is "inflated values" (summing duplicates), this overwrite helps.
                            # If it's about missing values, ensuring the value is not None is key.
                            consolidated_items_by_key[target_key][year_int] = clean_amount
                        logger.debug(f"  Mapped '{description}' for year {year_int}: {clean_amount} (from raw '{amount}')")
                    except (ValueError, TypeError) as e:
                        logger.warning(f"  Could not convert amount '{amount}' for year {year_str}, description '{description}' in file {filename}. Setting to None. Error: {e}")
                        consolidated_items_by_key[target_key][year_int] = None

    # --- NEW: Post-processing for Accumulated Surplus/Deficit Mutual Exclusivity (Existing) ---
    surplus_key = ("Accumulated surplus", "Equity", "Equity")
    deficit_key = ("Accumulated deficit", "Equity", "Equity")

    for year in sorted_years:
        surplus_value = None
        deficit_value = None

        # Find the actual keys used in consolidated_items_by_key for surplus/deficit
        # as they might be stored with inferred categories if not strictly matched to MASTER_STRUCTURE initially
        actual_surplus_key = None
        actual_deficit_key = None

        for key in consolidated_items_by_key.keys():
            if key[0] == "Accumulated surplus":
                actual_surplus_key = key
            elif key[0] == "Accumulated deficit":
                actual_deficit_key = key
            if actual_surplus_key and actual_deficit_key:
                break

        if actual_surplus_key and year in consolidated_items_by_key[actual_surplus_key]:
            surplus_value = consolidated_items_by_key[actual_surplus_key][year]

        if actual_deficit_key and year in consolidated_items_by_key[actual_deficit_key]:
            deficit_value = consolidated_items_by_key[actual_deficit_key][year]

        logger.debug(f"Year {year}: Initial Surplus: {surplus_value}, Initial Deficit: {deficit_value}")

        if surplus_value is not None and deficit_value is not None:
            # If both exist, decide which one is truly active for the year
            # A surplus is typically positive. A deficit can be negative or positive on a "deficit" line.
            # We assume if there's a non-zero surplus, the deficit should be ignored, and vice-versa.
            if surplus_value > 0 and deficit_value <= 0: # Positive surplus, or zero/negative deficit
                if actual_deficit_key:
                    consolidated_items_by_key[actual_deficit_key][year] = None
                logger.debug(f"  Year {year}: Kept Surplus, Nullified Deficit.")
            elif deficit_value > 0 and surplus_value <= 0: # Positive deficit (explicit deficit), or zero/negative surplus
                if actual_surplus_key:
                    consolidated_items_by_key[actual_surplus_key][year] = None
                logger.debug(f"  Year {year}: Kept Deficit, Nullified Surplus.")
            elif surplus_value < 0 and deficit_value >= 0: # Negative surplus (implies deficit), and non-negative deficit
                 if actual_surplus_key:
                    consolidated_items_by_key[actual_surplus_key][year] = None
                 logger.debug(f"  Year {year}: Negative surplus implies deficit, nullified surplus.")
            elif deficit_value < 0 and surplus_value >= 0: # Negative deficit (implies surplus), and non-negative surplus
                 if actual_deficit_key:
                    consolidated_items_by_key[actual_deficit_key][year] = None
                 logger.debug(f"  Year {year}: Negative deficit implies surplus, nullified deficit.")
            else:
                # If both are zero or both are positive/negative in an ambiguous way,
                # prioritize surplus if it's non-zero, otherwise deficit.
                if surplus_value != 0:
                    if actual_deficit_key:
                        consolidated_items_by_key[actual_deficit_key][year] = None
                    logger.debug(f"  Year {year}: Ambiguous, prioritized non-zero Surplus.")
                elif deficit_value != 0:
                    if actual_surplus_key:
                        consolidated_items_by_key[actual_surplus_key][year] = None
                    logger.debug(f"  Year {year}: Ambiguous, prioritized non-zero Deficit.")
                else: # Both are None or zero, nothing to do
                    logger.debug(f"  Year {year}: Both surplus and deficit are None/Zero, no action.")
        elif surplus_value is None and deficit_value is not None and deficit_value < 0:
            # If only deficit exists and it's negative, it implies a surplus.
            # Convert this negative deficit to a positive surplus and nullify deficit.
            if actual_surplus_key:
                consolidated_items_by_key[actual_surplus_key][year] = abs(deficit_value)
            if actual_deficit_key:
                consolidated_items_by_key[actual_deficit_key][year] = None
            logger.debug(f"  Year {year}: Converted negative deficit to positive surplus.")
        elif deficit_value is None and surplus_value is not None and surplus_value < 0:
            # If only surplus exists and it's negative, it implies a deficit.
            # Convert this negative surplus to a positive deficit and nullify surplus.
            if actual_deficit_key:
                consolidated_items_by_key[actual_deficit_key][year] = abs(surplus_value)
            if actual_surplus_key:
                consolidated_items_by_key[actual_surplus_key][year] = None
            logger.debug(f"  Year {year}: Converted negative surplus to positive deficit.")

    logger.debug("Post-processing for Accumulated Surplus/Deficit complete.")
    # --- END NEW: Post-processing ---


    # Collect unmapped items for display at the end
    unmapped_items_for_display = []
    for key, year_data in consolidated_items_by_key.items():
        description, category, subcategory = key
        is_explicitly_mapped = False
        for master_cat, master_subcats in MASTER_STRUCTURE.items():
            for master_subcat, master_descs in master_subcats.items():
                if description in master_descs:
                    is_explicitly_mapped = True
                    break
            if is_explicitly_mapped:
                break

        if not is_explicitly_mapped:
            # Exclude specific totals that might be extracted but are not desired in the raw data sheet
            if description.lower() not in ["total assets", "total equity", "total liabilities", "total equity and liabilities", "total income", "total operating expenses"]:
                unmapped_items_for_display.append((description, category, subcategory, year_data))

    unmapped_items_for_display.sort(key=lambda x: x[0]) # Sort alphabetically by description
    logger.debug(f"Unmapped items for display: {len(unmapped_items_for_display)}")


    # Write data to Excel following MASTER_STRUCTURE for hierarchical output
    category_display_order = ["Assets", "Equity", "Liabilities", "Revenue", "Expenses", "Other Financial Items"]

    for category_name in category_display_order:
        subcategories_dict = MASTER_STRUCTURE.get(category_name)
        if not subcategories_dict:
            continue

        # Check if this category has any data before printing its header
        category_has_data = False
        for subcategory_name in sorted(subcategories_dict.keys(), key=lambda x: (0 if x != "N/A" else 1, x)):
            descriptions_list = subcategories_dict.get(subcategory_name)
            if not descriptions_list:
                continue
            for description_from_master in descriptions_list:
                lookup_key = (description_from_master, category_name, subcategory_name)
                year_data = consolidated_items_by_key.get(lookup_key)
                if year_data and any(v is not None for v in year_data.values()):
                    category_has_data = True
                    break
            if category_has_data:
                break

        if category_has_data:
            main_ws.append([]) # Spacer row before new category
            main_ws.append([category_name])
            category_row = main_ws[main_ws.max_row]
            for cell in category_row:
                cell.font = bold_font
                cell.fill = category_fill
                cell.border = thin_border
            main_ws.merge_cells(start_row=category_row[0].row, start_column=1, end_row=category_row[0].row, end_column=1 + len(sorted_years))
            logger.debug(f"Printed category header: {category_name}")

            # Sort subcategories for consistent display (N/A last)
            sorted_subcategories = sorted(subcategories_dict.keys(), key=lambda x: (0 if x != "N/A" else 1, x))

            for subcategory_name in sorted_subcategories:
                descriptions_list = subcategories_dict.get(subcategory_name)
                if not descriptions_list:
                    continue

                # Check if this subcategory has any data before printing its header
                subcategory_has_data = False
                for description_from_master in descriptions_list:
                    lookup_key = (description_from_master, category_name, subcategory_name)
                    year_data = consolidated_items_by_key.get(lookup_key)
                    if year_data and any(v is not None for v in year_data.values()):
                        subcategory_has_data = True
                        break

                if subcategory_has_data and subcategory_name != "N/A":
                    main_ws.append([]) # Spacer row before new subcategory
                    main_ws.append([subcategory_name])
                    subcategory_row = main_ws[main_ws.max_row]
                    for cell in subcategory_row:
                        cell.font = bold_font
                        cell.fill = subcategory_fill
                        cell.border = thin_border
                    main_ws.merge_cells(start_row=subcategory_row[0].row, start_column=1, end_row=subcategory_row[0].row, end_column=1 + len(sorted_years))
                    logger.debug(f"  Printed subcategory header: {subcategory_name}")

                for description_from_master in descriptions_list:
                    lookup_key = (description_from_master, category_name, subcategory_name)
                    year_data = consolidated_items_by_key.get(lookup_key) # Get the consolidated data for this item

                    # Only print the row if there's actual data for it in any year
                    if year_data and any(v is not None for v in year_data.values()):
                        row_values = [description_from_master] # Start with description

                        # Add year values
                        for year in sorted_years:
                            amount = year_data.get(year)
                            row_values.append(amount)

                        main_ws.append(row_values)
                        desc_cell = main_ws[main_ws.max_row][0]

                        # Apply indent to description if it's under a subcategory (and not a top-level total)
                        if category_name in ["Assets", "Equity", "Liabilities", "Revenue", "Expenses", "Other Financial Items"] and subcategory_name != "N/A":
                            desc_cell.alignment = indent_alignment

                        # Apply number formatting and bolding for values
                        for col_idx, value in enumerate(row_values):
                            if col_idx > 0 and value is not None: # Only for value columns (not description)
                                cell = main_ws.cell(row=main_ws.max_row, column=col_idx + 1, value=value)
                                cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1

                        # Apply bold formatting to total lines (across all relevant columns)
                        if description_from_master.lower() in ["total assets", "total equity", "total liabilities", "total income", "total operating expenses", "total equity and liabilities", "accumulated funds"]: # Added accumulated funds to bold
                            for col_idx in range(1, 2 + len(sorted_years)): # Apply bold to Description + all year columns
                                main_ws.cell(row=main_ws.max_row, column=col_idx).font = bold_font
                        logger.debug(f"    Printed data row: {row_values}")
                    else:
                        logger.debug(f"    Skipped printing empty row for: {description_from_master}")


    # Add unmapped items at the very end under an "Additional Items" category
    if unmapped_items_for_display:
        # Check if there's any actual data in unmapped items before printing header
        unmapped_has_data = False
        for _, _, _, year_data in unmapped_items_for_display:
            if any(v is not None for v in year_data.values()):
                unmapped_has_data = True
                break

        if unmapped_has_data:
            main_ws.append([]) # Spacer
            main_ws.append(["Additional Items (Not in Master Structure)"])
            other_header_row = main_ws[main_ws.max_row]
            for cell in other_header_row:
                cell.font = bold_font
                cell.fill = category_fill
                cell.border = thin_border
            main_ws.merge_cells(start_row=other_header_row[0].row, start_column=1, end_row=other_header_row[0].row, end_column=1 + len(sorted_years))
            logger.debug("Printed 'Additional Items' header.")

            for desc, cat, subcat, year_data in unmapped_items_for_display:
                # Only print the row if there's actual data for it in any year
                if year_data and any(v is not None for v in year_data.values()):
                    row_values = [desc]
                    for year in sorted_years:
                        amount = year_data.get(year)
                        row_values.append(amount)
                    main_ws.append(row_values)

                    desc_cell = main_ws[main_ws.max_row][0]
                    desc_cell.alignment = indent_alignment # Indent these as they are "other"

                    for col_idx, value in enumerate(row_values):
                        if col_idx > 0 and value is not None:
                            cell = main_ws.cell(row=main_ws.max_row, column=col_idx + 1, value=value)
                            cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
                    logger.debug(f"    Printed unmapped data row: {row_values}")
                else:
                    logger.debug(f"    Skipped printing empty unmapped row for: {desc}")


    excel_stream = io.BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0)
    logger.info("Excel file generated in memory.")
    return excel_stream

# --- Routes ---

@app.route('/')
def home():
    """
    A simple home route to confirm the server is running.
    """
    return "PDF to Excel Converter Backend is running!"

@app.route('/upload-pdfs', methods=['POST'])
async def upload_pdfs():
    """
    Handles the upload of multiple PDF files from the frontend.
    It saves the files temporarily, performs OCR,
    sends the text to Gemini for parsing, generates Excel,
    and sends the Excel file back for download.
    """
    if 'pdfs' not in request.files:
        logger.error("No PDF files part in the request.")
        return jsonify({"error": "No PDF files part in the request"}), 400

    files = request.files.getlist('pdfs')
    client_name = request.form.get('client_name', 'Unknown Client')
    ai_prompt_text = request.form.get('ai_prompt', None)

    if not files:
        logger.warning("No selected PDF files.")
        return jsonify({"error": "No selected PDF files"}), 400

    all_parsed_data = {}

    for file in files:
        if file.filename == '':
            logger.warning("Skipping file with empty filename.")
            continue
        if file and file.filename.endswith('.pdf'):
            # Read file content directly into memory instead of saving to disk
            pdf_content = file.read()
            logger.info("Received PDF file: %s", file.filename)

            try:
                extracted_text = detect_text_from_pdf(pdf_content) # Pass content directly

                if not extracted_text.strip():
                    logger.warning("Skipped %s due to empty OCR result.", file.filename)
                    continue

                logger.info("OCR result length for %s: %d characters", file.filename, len(extracted_text))

                parsed_items = parse_with_gemini(extracted_text, file.filename, ai_prompt_text)

                if not parsed_items:
                    logger.warning("No items parsed from %s. Check OCR quality or Gemini prompt.", file.filename)

                all_parsed_data[file.filename] = parsed_items

            except Exception as e:
                logger.exception("Error processing %s: %s", file.filename, e) # Use exception for full traceback
                return jsonify({"error": f"Failed to process {file.filename}: {str(e)}"}), 500
        else:
            logger.warning("File %s is not a PDF. Skipping.", file.filename)
            return jsonify({"error": f"File {file.filename} is not a PDF"}), 400

    # No need to clean up uploaded files from disk as they are processed in memory

    if not all_parsed_data:
        logger.error("No data was parsed from any of the uploaded PDFs. Cannot generate Excel.")
        return jsonify({"error": "No data was parsed from any of the uploaded PDFs. Cannot generate Excel."}), 400

    excel_stream = generate_excel(client_name, all_parsed_data)

    try:
        return send_file(
            excel_stream,
            as_attachment=True,
            download_name=f'{client_name}_consolidated_financial_statements_position.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        logger.exception("Error sending Excel file: %s", e)
        return jsonify({"error": f"Failed to send Excel file: {str(e)}"}), 500
    finally:
        pass

@app.route('/refine-document', methods=['POST'])
async def refine_document():
    """
    Handles file uploads for the Document Refinement Tool.
    Extracts text from template and data files,
    uses Gemini to generate mapping instructions,
    and creates a refined document (placeholder for now).
    """
    if 'template_file' not in request.files or 'data_file' not in request.files:
        logger.error("Both template and data files are required for refinement.")
        return jsonify({"error": "Both template and data files are required"}), 400

    template_file = request.files['template_file']
    data_file = request.files['data_file']

    if not template_file.filename or not data_file.filename:
        logger.warning("No selected template or data file for refinement.")
        return jsonify({"error": "No selected template or data file"}), 400

    # Read file contents directly into memory
    template_content = template_file.read()
    data_content = data_file.read()

    logger.info(f"Received template file: {template_file.filename}, data file: {data_file.filename}")

    # Extract text using OCR or simple decoding
    template_text = extract_text_from_file(template_content, template_file.filename)
    data_text = extract_text_from_file(data_content, data_file.filename)

    if not template_text:
        logger.error(f"Could not extract text from template file: {template_file.filename}.")
        return jsonify({"error": f"Could not extract text from template file: {template_file.filename}. Ensure it's a readable format (PDF, TXT, or basic DOCX/XLSX/CSV decode)."}), 400
    if not data_text:
        logger.error(f"Could not extract text from data file: {data_file.filename}.")
        return jsonify({"error": f"Could not extract text from data file: {data_file.filename}. Ensure it's a readable format (PDF, TXT, or basic DOCX/XLSX/CSV decode)."}), 400

    # Generate parsing instructions using Gemini
    mapping_instructions = await generate_refinement_instructions_with_gemini(template_text, data_text)

    # Create the refined document (currently a dummy docx with instructions)
    refined_doc_io = create_refined_document(template_file.filename, mapping_instructions)

    # Determine output filename
    template_name_parts = os.path.splitext(template_file.filename)
    # Default to .docx output for the refined document
    output_filename = f"refined_{template_name_parts[0]}.docx"

    try:
        return send_file(refined_doc_io, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         as_attachment=True, download_name=output_filename)
    except Exception as e:
        logger.exception("Error sending refined document file: %s", e)
        return jsonify({"error": f"Failed to send refined document file: {str(e)}"}), 500


# --- Main execution block ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)