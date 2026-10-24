"""
End-to-End Test and Verification Script.
Creates a sample Bajaj Claim Summary Sheet PDF and verifies extraction & Excel appending.
"""

import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import openpyxl

from extractor import extract_data_from_pdf
from excel_handler import get_excel_info, append_rows_to_excel


def create_sample_bajaj_pdf(pdf_path):
    """Creates a sample PDF identical in text and table structure to the user's INTIMATION.pdf."""
    doc = SimpleDocTemplate(pdf_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []

    # Header
    title_style = ParagraphStyle(
        'HeaderTitle',
        parent=styles['Heading2'],
        fontSize=14,
        leading=16,
        textColor=colors.black
    )
    story.append(Paragraph("<b>Bajaj General Insurance Limited.</b>", title_style))
    story.append(Spacer(1, 10))

    # Claim Summary Sheet Title
    center_style = ParagraphStyle(
        'CenterTitle',
        parent=styles['Normal'],
        fontSize=12,
        leading=14,
        alignment=1,
        fontName="Helvetica-Bold"
    )
    story.append(Paragraph("Claim Summary Sheet", center_style))
    story.append(Spacer(1, 8))

    # Top Metadata Row
    meta_data = [
        ["Date of Intimation", "22-Oct-2026", "Intimation print Date", "23-Oct-2026 09:39:01"]
    ]
    t_meta = Table(meta_data, colWidths=[130, 130, 130, 150])
    t_meta.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.grey),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 5))

    # Section Headers
    section_headers = [
        ["Policy Details", "Claim Intimation/Registration Details"]
    ]
    t_sec = Table(section_headers, colWidths=[260, 280])
    t_sec.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.grey),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke)
    ]))
    story.append(t_sec)

    # Detailed 4-Column Table matching the screenshot
    claim_table_data = [
        ["Policy Reference Number", "OG-27-9910-1827-00011943", "Claim Reference Number", "OC-27-1203-1827-00000303"],
        ["Policy Holder\\Insured Name", "RAHUL PAL", "Name of Caller//Claimant", "GURPREET SINGH"],
        ["Vehicle Registration Number", "PB10KS1232", "Date of Loss", "20-OCT-2026"],
        ["Policy Inception Date", "29-JUL-2026", "Time of Loss", "00:00:00"],
        ["Policy Expiry Date", "28-JUL-2029", "Driver's Name", ""],
        ["Policy Issuance Date", "30-JUL-2026", "Driving License Number", ""],
        ["64VB", "64VB is Cleared.", "Driver Relation with Insured", ""],
        ["Policy Status", "Issued", "Claimant Relationship with", "-"],
        ["Vehicle Make", "MARUTI", "No Claim Bonus", "0"],
        ["Vehicle Model", "BREZZA", "No Claim Bonus Amount", "0"],
        ["Past Endorsement No.", "", "NCB Status", ""],
        ["Engine Number", "K15CNK202315", "Chassis (VIN) Number", "MA3RYHL1STG830860"]
    ]

    t_main = Table(claim_table_data, colWidths=[140, 120, 140, 140])
    t_main.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.grey),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
    ]))
    story.append(t_main)

    doc.build(story)
    print(f"Created sample PDF at: {pdf_path}")


def create_sample_master_excel(excel_path):
    """Creates a sample master Excel file identical to the user's 2.mis-2026-2027-.xlsx."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MIS-2026-2027 (2)"

    # Row 1: Headers matching user screenshot
    headers = [
        "Ref no.", "INS. CO", "PROXIMATE DAYS", "TAT", "STATUS", "", "", "AASSIGN TO",
        "DRIVER RELATIO", "NCB", "IDV", "Policy No", "policy start date", "date of loss",
        "claim no.", "Date of survey", "Name"
    ]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=h)

    # Add dummy rows up to 788 to match user's real sheet
    # Row 785
    ws.cell(row=785, column=1, value=784)
    ws.cell(row=785, column=2, value="BAJAJ")
    ws.cell(row=785, column=11, value="469592/-")
    ws.cell(row=785, column=12, value="80000031250360137266")
    ws.cell(row=785, column=13, value="30-Dec-25")
    ws.cell(row=785, column=14, value="20.10.2026")
    ws.cell(row=785, column=15, value="80000031260390091553")
    ws.cell(row=785, column=16, value="23-Oct-26")
    ws.cell(row=785, column=17, value="TUSHAR SACHDEV")

    # Row 788
    ws.cell(row=788, column=1, value=787)
    ws.cell(row=788, column=2, value="BAJAJ")
    ws.cell(row=788, column=15, value="200504201126110192701")
    ws.cell(row=788, column=16, value="23-Oct-26")
    ws.cell(row=788, column=17, value="Jagdeep Singh")

    # Row 789
    ws.cell(row=789, column=1, value=788)
    ws.cell(row=789, column=2, value="LIBERTY")
    ws.cell(row=789, column=5, value="AP")

    wb.save(excel_path)
    wb.close()
    print(f"Created sample master Excel at: {excel_path}")


def run_test():
    sample_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_data")
    os.makedirs(sample_dir, exist_ok=True)

    pdf_path = os.path.join(sample_dir, "sample_bajaj_intimation.pdf")
    excel_path = os.path.join(sample_dir, "2.mis-2026-2027-.xlsx")

    create_sample_bajaj_pdf(pdf_path)
    create_sample_master_excel(excel_path)

    print("\n--- 1. Testing Excel Inspection ---")
    info = get_excel_info(excel_path, "MIS-2026-2027 (2)")
    print(f"Last Row: {info.get('last_row')}, Last Ref: {info.get('last_ref_no')}, Next Ref: {info.get('next_ref_no')}")
    assert info.get("last_ref_no") == 788, f"Expected 788, got {info.get('last_ref_no')}"
    assert info.get("next_ref_no") == 789, f"Expected 789, got {info.get('next_ref_no')}"

    print("\n--- 2. Testing PDF Extraction ---")
    extracted = extract_data_from_pdf(pdf_path)
    print("Extracted Data:")
    for k, v in extracted.items():
        print(f"  {k}: {v}")

    assert extracted["ins_co"] == "BAJAJ", f"Expected BAJAJ, got {extracted['ins_co']}"
    assert extracted["policy_no"] == "OG-27-9910-1827-00011943", f"Wrong policy_no: {extracted['policy_no']}"
    assert extracted["claim_no"] == "OC-27-1203-1827-00000303", f"Wrong claim_no: {extracted['claim_no']}"
    assert extracted["insured_name"] == "RAHUL PAL", f"Wrong insured_name: {extracted['insured_name']}"
    assert extracted["date_of_loss"] == "20-Oct-26", f"Wrong date_of_loss: {extracted['date_of_loss']}"
    assert extracted["policy_start_date"] == "29-Jul-26", f"Wrong policy_start_date: {extracted['policy_start_date']}"

    print("\n--- 3. Testing Excel Appending ---")
    extracted["ref_no"] = info["next_ref_no"]  # 789
    res = append_rows_to_excel(excel_path, "MIS-2026-2027 (2)", [extracted], create_backup=False)
    print(f"Append Result: {res}")
    assert res["success"] is True

    # Re-inspect excel to verify appended row
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["MIS-2026-2027 (2)"]
    new_row = 790
    print(f"\nRow {new_row} in Excel:")
    print(f"  Col A (Ref no): {ws.cell(row=new_row, column=1).value}")
    print(f"  Col B (INS. CO): {ws.cell(row=new_row, column=2).value}")
    print(f"  Col L (Policy No): {ws.cell(row=new_row, column=12).value}")
    print(f"  Col M (policy start date): {ws.cell(row=new_row, column=13).value}")
    print(f"  Col N (date of loss): {ws.cell(row=new_row, column=14).value}")
    print(f"  Col O (claim no.): {ws.cell(row=new_row, column=15).value}")
    print(f"  Col P (Date of survey): {ws.cell(row=new_row, column=16).value}")
    print(f"  Col Q (Name): {ws.cell(row=new_row, column=17).value}")

    assert ws.cell(row=new_row, column=1).value == 789
    assert ws.cell(row=new_row, column=2).value == "BAJAJ"
    assert ws.cell(row=new_row, column=12).value == "OG-27-9910-1827-00011943"
    assert ws.cell(row=new_row, column=15).value == "OC-27-1203-1827-00000303"
    assert ws.cell(row=new_row, column=17).value == "RAHUL PAL"
    wb.close()

    print("\n=======================================================")
    print("  ALL VERIFICATION TESTS PASSED PERFECTLY (100% ACCURATE)!")
    print("=======================================================\n")


if __name__ == "__main__":
    run_test()
