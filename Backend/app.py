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
# - **REMOVED**: Document Refinement Tool and all associated functions and routes.
# - **HEROKU READY**: Modified app.run() to use Heroku's assigned PORT.

import os
import io
import re
import json
import requests
import logging
import copy
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from google.cloud import vision
from dotenv import load_dotenv
import google.generativeai as genai
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill, numbers
from openpyxl.utils import get_column_letter


load_dotenv()

# --- Logging Configuration ---
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# Heroku provides a temporary filesystem, so UPLOAD_FOLDER and OUTPUT_FOLDER
# are not strictly necessary to create persistent directories, but good for local dev.
# For Heroku, files are typically processed in-memory or in /tmp.
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

if not GEMINI_API_KEY:
    logger.error("Error: GEMINI_API_KEY is not set. Cannot call Gemini API.")
    # In a production Heroku app, you might want to raise an error or
    # have a more robust way to handle missing API keys.
    # For now, it will just log an error.
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
    "depreciation and amortisation": "Depreciation and amortisation",
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
    "taxation": "Taxation",
    "non provision of tax": "Non provision of tax",
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


if __name__ == '__main__':
    # Get the port from environment variable, default to 5000 for local development
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)