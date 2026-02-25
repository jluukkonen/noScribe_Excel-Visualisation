import re
import pandas as pd
import sys
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

def parse_transcript(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    data = []
    
    # Regex to match the speaker turn format: "S01: text" or "S00: text"
    speaker_pattern = re.compile(r"^(S\d{2}):\s*(.*)")
    
    current_speaker = None
    current_text = ""
    
    # Skip the header lines
    start_parsing = False
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        if line == "(..)" or line == "(.)":
            start_parsing = True
            
        if not start_parsing and not speaker_pattern.match(line):
            continue
            
        start_parsing = True
        
        match = speaker_pattern.match(line)
        if match:
            # Save previous turn if it exists
            if current_speaker is not None:
                data.append({
                    "Speaker": "Maron" if current_speaker == "S01" else ("Obama" if current_speaker == "S00" else current_speaker),
                    "Text": current_text.strip()
                })
            
            # Start new turn
            current_speaker = match.group(1)
            current_text = match.group(2)
        else:
            # Append to current turn
            if current_speaker is not None:
                current_text += " " + line
                
    # Append the last turn
    if current_speaker is not None:
        data.append({
            "Speaker": "Maron" if current_speaker == "S01" else ("Obama" if current_speaker == "S00" else current_speaker),
            "Text": current_text.strip()
        })
        
    # Extract features for the excel file
    for row in data:
        text = row["Text"]
        row["Word Count"] = len(text.split())
        
        # Count overlaps (indicated by //text//)
        overlaps = len(re.findall(r"//(.*?)//", text))
        row["Contains Overlap?"] = "Yes" if overlaps > 0 else "No"
        
        # Count pauses
        short_pauses = text.count("(.)")
        long_pauses = text.count("(..)") + text.count("(...)")
        row["Short Pauses"] = short_pauses
        row["Long Pauses"] = long_pauses
        
        # Simple disfluency count
        disfluency_pattern = re.compile(r"\b(uh|um|hmm|yeah|like)\b", re.IGNORECASE)
        row["Disfluencies (uh/um/yeah...)"] = len(disfluency_pattern.findall(text))

    return data

def apply_excel_styling(output_file, df):
    wb = load_workbook(output_file)
    ws = wb.active
    ws.title = "Transcript Data"
    
    # --- 1. Define Visual Styles ---
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    obama_fill = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid") # Light Blue
    maron_fill = PatternFill(start_color="EBF1DE", end_color="EBF1DE", fill_type="solid") # Light Green
    wrap_alignment = Alignment(wrap_text=True, vertical="top")
    top_alignment = Alignment(vertical="top")
    
    # --- 2. Apply Header Styling and Freeze Top Row ---
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    ws.freeze_panes = "A2"
    
    # --- 3. Set Column Widths ---
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 80
    ws.column_dimensions['C'].width = 13
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 14
    ws.column_dimensions['G'].width = 25

    # --- 4. Apply Row Formatting ---
    for row_idx in range(2, ws.max_row + 1):
        speaker_name = ws.cell(row=row_idx, column=1).value
        ws.cell(row=row_idx, column=2).alignment = wrap_alignment
        for col_idx in [1, 3, 4, 5, 6, 7]:
            ws.cell(row=row_idx, column=col_idx).alignment = top_alignment

        if speaker_name == "Obama":
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).fill = obama_fill
        elif speaker_name == "Maron":
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).fill = maron_fill
                
    # --- 5. CREATE THE RESULTS SHEET ---
    ws_results = wb.create_sheet(title="Results Summary", index=0)
    
    # Calculate stats
    speakers = ["Obama", "Maron"]
    stats = {}
    for s in speakers:
        speaker_df = df[df["Speaker"] == s]
        stats[s] = {
            "Total Turns": len(speaker_df),
            "Total Words": int(speaker_df["Word Count"].sum()),
            "Avg Words per Turn": round(speaker_df["Word Count"].mean(), 1) if len(speaker_df) > 0 else 0,
            "Total Overlaps Initiated/Involved": len(speaker_df[speaker_df["Contains Overlap?"] == "Yes"]),
            "Total Short Pauses (.)": int(speaker_df["Short Pauses"].sum()),
            "Total Long Pauses (..)": int(speaker_df["Long Pauses"].sum()),
            "Total Disfluencies (uh/yeah)": int(speaker_df["Disfluencies (uh/um/yeah...)"].sum()),
        }

    # Title styling
    title_font = Font(size=16, bold=True, color="FFFFFF")
    title_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    
    # Write Title
    ws_results.merge_cells('B2:E2')
    title_cell = ws_results.cell(row=2, column=2, value="SPEAKER COMPARISON FINDINGS")
    title_cell.font = title_font
    title_cell.fill = title_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Write Headers
    sub_header_font = Font(bold=True)
    ws_results.cell(row=4, column=3, value="BARACK OBAMA").font = sub_header_font
    ws_results.cell(row=4, column=3).fill = obama_fill
    ws_results.cell(row=4, column=3).alignment = Alignment(horizontal="center")
    
    ws_results.cell(row=4, column=4, value="MARC MARON").font = sub_header_font
    ws_results.cell(row=4, column=4).fill = maron_fill
    ws_results.cell(row=4, column=4).alignment = Alignment(horizontal="center")

    # Metrics Layout
    metrics = [
        ("Total Speaking Turns", "Total Turns"),
        ("Total Words Spoken", "Total Words"),
        ("Average Words per Turn", "Avg Words per Turn"),
        ("Overlapping Statements", "Total Overlaps Initiated/Involved"),
        ("Short Pauses (.)", "Total Short Pauses (.)"),
        ("Long Pauses (..)", "Total Long Pauses (..)"),
        ("Disfluencies (uh, um, yeah)", "Total Disfluencies (uh/yeah)")
    ]

    current_row = 6
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for display_name, dict_key in metrics:
        # Label
        label_cell = ws_results.cell(row=current_row, column=2, value=display_name)
        label_cell.font = Font(bold=True)
        label_cell.alignment = Alignment(horizontal="right")
        
        # Obama Stat
        ob_cell = ws_results.cell(row=current_row, column=3, value=stats["Obama"][dict_key])
        ob_cell.fill = obama_fill
        ob_cell.border = thin_border
        ob_cell.alignment = Alignment(horizontal="center")
        
        # Maron Stat
        mar_cell = ws_results.cell(row=current_row, column=4, value=stats["Maron"][dict_key])
        mar_cell.fill = maron_fill
        mar_cell.border = thin_border
        mar_cell.alignment = Alignment(horizontal="center")
        
        current_row += 2 # Double space for presentation

    # Set column widths for results
    ws_results.column_dimensions['A'].width = 5
    ws_results.column_dimensions['B'].width = 30
    ws_results.column_dimensions['C'].width = 20
    ws_results.column_dimensions['D'].width = 20
    ws_results.column_dimensions['E'].width = 5

    wb.save(output_file)

def main():
    if len(sys.argv) < 3:
        print("Usage: python parse_to_excel.py <input_txt> <output_xlsx>")
        return
        
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    print(f"Parsing {input_file}...")
    data = parse_transcript(input_file)
    
    df = pd.DataFrame(data)
    columns = ["Speaker", "Text", "Word Count", "Contains Overlap?", "Short Pauses", "Long Pauses", "Disfluencies (uh/um/yeah...)"]
    df = df[columns]
    
    print(f"Exporting raw data to {output_file}...")
    df.to_excel(output_file, index=False, engine="openpyxl")
    
    print("Applying visual magic and creating Results sheet...")
    apply_excel_styling(output_file, df)
    
    print("Done!")

if __name__ == "__main__":
    main()
