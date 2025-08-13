import fitz  # PyMuPDF
import json
import re
import os
import pandas as pd

# --- Configuration ---
PDF_FILE_PATH = 'USB_PD_R3_2_V1.1_2024_10.pdf'
OUTPUT_TOC_JSONL_FILE = 'usb_pd_toc.jsonl'
OUTPUT_SPEC_JSONL_FILE = 'usb_pd_spec.jsonl'
OUTPUT_METADATA_JSONL_FILE = 'usb_pd_metadata.jsonl'
OUTPUT_VALIDATION_REPORT_FILE = 'validation_report.xlsx'
DOCUMENT_TITLE = "USB Power Delivery Specification Rev 3.2 V1.1"


# --- Functions from Part 1 & 2 (with minor improvements) ---

def get_toc_pages(doc):
    """Identifies pages for ToC, Figures, and Tables."""
    # For 'USB_PD_R3_2 V1.1 2024-10.pdf', ToC and lists are from page 4 to 12 (0-indexed 3-11).
    return list(range(3, 12))


def parse_toc(doc, toc_pages):
    """Parses the main Table of Contents for sections."""
    full_toc_text = "".join([doc[page_num].get_text("text") for page_num in toc_pages])
    toc_entries = []
    toc_line_regex = re.compile(r"^(\d+(?:\.\d+))\s+([^\n\.]+?)\s\.+\s*(\d+)$", re.MULTILINE)

    for match in toc_line_regex.finditer(full_toc_text):
        section_id = match.group(1).strip()
        title = match.group(2).strip()
        page = int(match.group(3).strip()) - 1  # Convert to 0-indexed

        entry = {
            "doc_title": DOCUMENT_TITLE, "section_id": section_id, "title": title, "page": page,
            "level": section_id.count('.') + 1,
            "parent_id": section_id.rsplit('.', 1)[0] if '.' in section_id else None,
            "full_path": f"{section_id} {title}", "tags": []
        }
        toc_entries.append(entry)
    return toc_entries


def extract_section_content(doc, toc_entries):
    """Extracts the text content for each section based on the ToC entries."""
    full_spec_entries = []
    sorted_entries = sorted(toc_entries, key=lambda x: x['page'])

    for i, current_entry in enumerate(sorted_entries):
        start_page = current_entry['page']
        end_page = len(doc) - 1
        if i + 1 < len(sorted_entries):
            end_page = sorted_entries[i + 1]['page']

        content_block = "".join([doc[p].get_text("text") for p in range(start_page, end_page + 1)])

        current_title_pattern = re.escape(f"{current_entry['section_id']} {current_entry['title']}")
        start_match = re.search(current_title_pattern, content_block)
        start_index = start_match.start() if start_match else 0

        end_index = len(content_block)
        if i + 1 < len(sorted_entries):
            next_entry = sorted_entries[i + 1]
            if next_entry['page'] <= end_page:
                next_title_pattern = re.escape(f"{next_entry['section_id']} {next_entry['title']}")
                next_match = re.search(next_title_pattern, content_block)
                if next_match:
                    end_index = next_match.start()

        content = content_block[start_index:end_index].strip()

        spec_entry = current_entry.copy()
        spec_entry["text"] = content
        spec_entry["page"] = current_entry['page'] + 1  # Convert back to 1-indexed for output
        full_spec_entries.append(spec_entry)

    return full_spec_entries


# --- New Functions for Part 3 ---

def parse_list_of_tables(doc, toc_pages):
    """Parses the 'List of Tables' from the ToC pages."""
    toc_text = "".join([doc[page_num].get_text("text") for page_num in toc_pages])
    # Regex to find table entries like "Table 6-39. Discover Identity Command Response ......... 295"
    table_regex = re.compile(r"^(Table\s+[\d\w-]+)\s", re.MULTILINE)
    tables_in_toc = re.findall(table_regex, toc_text)
    return tables_in_toc


def count_tables_in_body(spec_entries):
    """Counts table mentions in the parsed text content."""
    # Regex to find "Table X-Y" in the text
    table_regex = re.compile(r"Table\s+\d+[\w-]+\d+")
    found_tables = []
    for entry in spec_entries:
        found_tables.extend(re.findall(table_regex, entry['text']))
    return list(set(found_tables))  # Return unique table mentions


def generate_validation_report(toc_entries, spec_entries, toc_tables, body_tables):
    """Generates an Excel validation report."""
    print("\n--- Generating Validation Report ---")

    # 1. Section Count Comparison
    toc_section_count = len(toc_entries)
    parsed_section_count = len(spec_entries)

    # 2. Section Mismatch/Gap Analysis
    toc_ids = {entry['section_id'] for entry in toc_entries}
    parsed_ids = {entry['section_id'] for entry in spec_entries}

    sections_in_toc_not_parsed = list(toc_ids - parsed_ids)
    sections_parsed_not_in_toc = list(parsed_ids - toc_ids)

    # 3. Table Count Comparison
    toc_table_count = len(toc_tables)
    body_table_count = len(body_tables)

    # 4. Create Summary DataFrame
    summary_data = {
        "Metric": ["Total Sections in ToC", "Total Sections Parsed", "Sections Match", "Total Tables in ToC List",
                   "Total Tables Found in Body"],
        "Value": [toc_section_count, parsed_section_count, "Yes" if toc_section_count == parsed_section_count else "No",
                  toc_table_count, body_table_count]
    }
    summary_df = pd.DataFrame(summary_data)

    # 5. Create Mismatch Details DataFrame
    mismatch_data = {
        "Sections in ToC but Not Parsed": pd.Series(sections_in_toc_not_parsed, dtype='str'),
        "Sections Parsed but Not in ToC": pd.Series(sections_parsed_not_in_toc, dtype='str')
    }
    mismatch_df = pd.DataFrame.from_dict(mismatch_data, orient='index').transpose()

    # 6. Write to Excel file
    with pd.ExcelWriter(OUTPUT_VALIDATION_REPORT_FILE) as writer:
        summary_df.to_excel(writer, sheet_name="Validation Summary", index=False)
        mismatch_df.to_excel(writer, sheet_name="Mismatch Details", index=False)

    print(f"Successfully created validation report: {OUTPUT_VALIDATION_REPORT_FILE}")


# --- Main Execution Logic ---

def main():
    """Main function to run the entire parsing and validation pipeline."""
    print(f"--- Starting Parser for {PDF_FILE_PATH} ---")
    if not os.path.exists(PDF_FILE_PATH):
        print(f"Error: File not found at {PDF_FILE_PATH}")
        return

    doc = fitz.open(PDF_FILE_PATH)

    # Identify pages with ToC, lists of tables/figures
    toc_pages = get_toc_pages(doc)

    # Part 1: Parse ToC for sections and generate JSONL
    toc_entries = parse_toc(doc, toc_pages)
    with open(OUTPUT_TOC_JSONL_FILE, 'w') as f:
        sorted_toc = sorted(toc_entries, key=lambda x: list(map(int, x['section_id'].split('.'))))
        for entry in sorted_toc:
            output_entry = entry.copy()
            output_entry['page'] += 1  # 1-indexed for output
            f.write(json.dumps(output_entry) + '\n')
    print(f"Successfully created {OUTPUT_TOC_JSONL_FILE} with {len(toc_entries)} entries.")

    # Part 2: Parse full document and generate JSONL
    spec_entries = extract_section_content(doc, toc_entries)
    with open(OUTPUT_SPEC_JSONL_FILE, 'w') as f:
        sorted_spec = sorted(spec_entries, key=lambda x: list(map(int, x['section_id'].split('.'))))
        for entry in sorted_spec:
            f.write(json.dumps(entry) + '\n')
    print(f"Successfully created {OUTPUT_SPEC_JSONL_FILE} with {len(spec_entries)} entries.")

    # Generate Metadata file
    metadata = {"doc_title": DOCUMENT_TITLE, "source_file": PDF_FILE_PATH, "total_pages": len(doc)}
    with open(OUTPUT_METADATA_JSONL_FILE, 'w') as f:
        f.write(json.dumps(metadata) + '\n')
    print(f"Successfully created {OUTPUT_METADATA_JSONL_FILE}.")

    # Part 3: Run Validation
    # Get table counts for the report
    toc_tables = parse_list_of_tables(doc, toc_pages)
    body_tables = count_tables_in_body(spec_entries)
    # Generate the report
    generate_validation_report(toc_entries, spec_entries, toc_tables, body_tables)

    print("\n--- Project Complete ---")


if __name__ == "__main__":
    main() 