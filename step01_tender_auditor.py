"""
Filename: step01_tender_auditor.py
Description: [Step 01] Core logic script. Audits the tender against specs without GUI.
Author: Joanna
"""

import pandas as pd
from docx import Document
import os

# Configuration
TENDER_FILE = "mock_tender_document.docx"
SPECS_FILE = "product_specs.xlsx"
OUTPUT_FILE = "Audit_Result_Report.xlsx"

def extract_tender_table(docx_path):
    """Extracts the technical parameter table from Word."""
    if not os.path.exists(docx_path):
        return None
        
    doc = Document(docx_path)
    data = []
    
    for table in doc.tables:
        # Robust logic to find the correct table
        if len(table.rows) > 0:
            header = " ".join([c.text for c in table.rows[0].cells])
            if "Parameter" in header or "Requirement" in header:
                for i, row in enumerate(table.rows):
                    if i == 0: continue
                    # Safe extraction
                    try:
                        data.append({
                            "Parameter Name": row.cells[1].text.strip(),
                            "Tender Requirement": row.cells[2].text.strip()
                        })
                    except IndexError:
                        continue
                break
    return pd.DataFrame(data)

def run_audit():
    print(f"[START] Step 01: Starting Audit Process...")

    # 1. Check Files
    if not os.path.exists(TENDER_FILE) or not os.path.exists(SPECS_FILE):
        print("[ERROR] Input files missing. Please run 'step00_generate_mock_data.py' first.")
        return

    # 2. Load Data
    df_tender = extract_tender_table(TENDER_FILE)
    df_product = pd.read_excel(SPECS_FILE)
    
    if df_tender is None or df_tender.empty:
        print("[ERROR] No table found in Word document.")
        return

    print(f"[OK] Loaded Tender Items: {len(df_tender)}")
    print(f"[OK] Loaded Product Specs: {len(df_product)}")

    # 3. Merge & Audit
    result = pd.merge(df_tender, df_product, on="Parameter Name", how="left")
    
    # Simple Audit Logic
    def check_compliance(row):
        req = str(row.get('Tender Requirement', ''))
        spec = str(row.get('Our Product Spec', ''))
        
        # Logic: If requirement has number, compare numbers
        if "mm" in req and ">=" in req:
            try:
                # Extract number (Simple version)
                req_num = float(''.join(filter(str.isdigit, req)))
                spec_num = float(''.join(filter(str.isdigit, spec)))
                return "PASS" if spec_num >= req_num else "FAIL (Value Mismatch)"
            except:
                pass
        
        return "Manual Review Needed"

    result['Audit Result'] = result.apply(check_compliance, axis=1)

    # 4. Save
    result.to_excel(OUTPUT_FILE, index=False)
    print(f"[DONE] Audit Complete! Report saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    run_audit()