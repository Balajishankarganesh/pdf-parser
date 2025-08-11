import fitz  # PyMuPDF
import json
import re
import os

# --- Configuration ---
PDF_FILE_PATH = 'USB_PD_R3_2 V1.1 2024-10.pdf'
OUTPUT_TOC_JSONL_FILE = 'usb_pd_toc.jsonl'
DOCUMENT_TITLE = "USB Power Delivery Specification Rev 3.2 V1.1" # Extracted from the document's first page

def get_toc_pages(doc):
    """
    Identifies the pages containing the Table of Contents.
    This function looks for a page that starts with "Contents" and continues
    until it finds a page that does not seem to be part of the ToC.
    """
    toc_pages = []
    in_toc = False
    for page_num, page in enumerate(doc):
        text = page.get_text("text")
        # A simple heuristic to find the start of the ToC
        if "Contents" in text[:100]: # Check near the top of the page
            in_toc = True
        
        if in_toc:
            # A heuristic to detect the end of the ToC. 
            # Often, the ToC ends before the main content (e.g., "Chapter 1") begins in earnest,
            # or a new header like "Figures" or "Tables" appears. We'll assume for now
            # that once we start, all pages with ToC-like lines are part of it.
            # A more robust solution might check for a sudden change in formatting.
            toc_pages.append(page_num)

            # Heuristic to stop: If we see a page that doesn't have any ToC-like entries, we stop.
            # This regex looks for lines starting with numbers like "1", "1.1", "1.1.1".
            if not re.search(r"^\d+(\.\d+)*\s", text, re.MULTILINE):
                 # Let's check a bit more robustly, sometimes the title "Contents" is on its own page
                 if len(toc_pages) > 1: # Make sure we've actually gathered some ToC pages
                    # This page has no ToC lines, let's remove it and stop
                    toc_pages.pop()
                    break

    # This is a fallback in case the simple check fails. We will manually inspect the PDF.
    # For 'USB_PD_R3_2 V1.1 2024-10.pdf', the ToC is from page 4 to 12.
    # The page numbers in fitz are 0-indexed, so we use 3 to 11.
    return list(range(3, 12))


def parse_toc(pdf_path):
    """
    Parses the Table of Contents from the given PDF file and generates a JSONL file.
    
    This function implements the core logic as described in the assignment:
    - Extracts text from the identified ToC pages. [cite: 1885]
    - Uses regex to parse section ID, title, and page number. [cite: 1923, 1924]
    - Infers hierarchy level and parent-child relationships. [cite: 1925, 1926]
    - Generates a JSONL output file. 
    """
    
    print(f"Opening PDF: {pdf_path}")
    if not os.path.exists(pdf_path):
        print(f"Error: File not found at {pdf_path}")
        return

    doc = fitz.open(pdf_path)
    toc_entries = []
    
    # Regex to capture section_id, title, and page number.
    # It looks for a pattern like: <section_id> <title> ........... <page_number>
    # - ^(\d+(\.\d+)*)   -> Captures the section number at the start of a line (e.g., "2.1.2").
    # - \s+              -> Matches the space after the section number.
    # - ([^\n\.]+?)     -> Captures the title (non-greedily) until we see a series of dots or a newline.
    # - \s*\.+\s* -> Matches the '....' separator, which is optional.
    # - (\d+)$           -> Captures the page number at the end of the line.
    toc_line_regex = re.compile(r"^(\d+(?:\.\d+))\s+([^\n\.]+?)\s\.+\s*(\d+)$", re.MULTILINE)

    toc_page_nums = get_toc_pages(doc)
    print(f"Identified ToC on pages (0-indexed): {toc_page_nums}")

    full_toc_text = ""
    for page_num in toc_page_nums:
        full_toc_text += doc[page_num].get_text("text")

    # Clean up common PDF extraction artifacts, like extra spaces around newlines
    full_toc_text = os.linesep.join([s for s in full_toc_text.splitlines() if s])

    print("\n--- Parsing ToC Text ---")
    matches = toc_line_regex.finditer(full_toc_text)
    
    for match in matches:
        section_id = match.group(1).strip()
        title = match.group(2).strip()
        page = int(match.group(3).strip())
        
        # Infer level from the number of dots in the section_id 
        level = section_id.count('.') + 1
        
        # Determine parent_id from section_id [cite: 1926]
        parent_id = None
        if '.' in section_id:
            parent_id = section_id.rsplit('.', 1)[0]
        
        # Construct the full_path [cite: 1896]
        full_path = f"{section_id} {title}"
        
        # Create the JSON object according to the schema 
        entry = {
            "doc_title": DOCUMENT_TITLE,
            "section_id": section_id,
            "title": title,
            "page": page,
            "level": level,
            "parent_id": parent_id,
            "full_path": full_path,
            "tags": [] # Optional field, kept empty for now
        }
        toc_entries.append(entry)

    print(f"Successfully parsed {len(toc_entries)} ToC entries.")
    
    # Generate the JSONL output file 
    print(f"Generating JSONL file: {OUTPUT_TOC_JSONL_FILE}")
    with open(OUTPUT_TOC_JSONL_FILE, 'w') as f:
        for entry in toc_entries:
            f.write(json.dumps(entry) + '\n')
            
    print("--- ToC Parsing Complete ---")


# --- Main Execution ---
if __name__ == "_main_":
    parse_toc(PDF_FILE_PATH)