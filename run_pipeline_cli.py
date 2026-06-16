import os
import sys
from dotenv import load_dotenv
from generate_items import run_pipeline

def progress_cb(item_num, total, item_id, level, msg):
    print(f"[{item_num}/{total}] {item_id} ({level}): {msg}")

def main():
    # Load env vars
    if os.path.exists(".env"):
        load_dotenv()
    api_key = os.environ.get("MISTRAL_API_KEY", "")
    if not api_key:
        print("Error: MISTRAL_API_KEY not found in environment.")
        sys.exit(1)
        
    excel_path = r"c:\Users\aryan\OneDrive\Desktop\content team\Life Cycle of a Plant - Living Creatures.xlsx"
    if not os.path.exists(excel_path):
        print(f"Error: Excel file not found at {excel_path}")
        sys.exit(1)
        
    batch_meta = {
        "chapter_name": "Living Creatures: Exploring their Characteristics",
        "chapter_code": "LIV",
        "ncf_cg": "CG-1",
        "competency": "C-1.1",
        "learning_outcome": "Identifies and explains different properties of materials and relates them to their uses.",
        "lo_id": "LO-1.1.a",
        "start_seq": 1,
    }
    
    print("Starting pipeline...")
    with open(excel_path, 'rb') as excel_file:
        docx_bytes, csv_bytes, summary = run_pipeline(
            excel_file=excel_file,
            api_key=api_key,
            batch_meta=batch_meta,
            progress_callback=progress_cb
        )
        
    print(f"\nPipeline complete! Summary: {summary}")
    
    # Save the output docx and csv to output directory in workspace
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    
    docx_out = os.path.join(output_dir, "Science_PAL_G6_Item_Bank_20260611.docx")
    csv_out = os.path.join(output_dir, "Science_PAL_G6_Item_Bank_20260611_log.csv")
    
    with open(docx_out, 'wb') as f:
        f.write(docx_bytes)
    with open(csv_out, 'wb') as f:
        f.write(csv_bytes)
        
    # Also write them to the parent directory to replace the previous delivery files
    parent_docx_out = r"c:\Users\aryan\OneDrive\Desktop\content team\Science_PAL_G6_Item_Bank_20260611.docx"
    with open(parent_docx_out, 'wb') as f:
        f.write(docx_bytes)
        
    print(f"Outputs written to:\n - {docx_out}\n - {csv_out}\n - {parent_docx_out}")

if __name__ == "__main__":
    main()
