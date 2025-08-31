# USB Power Delivery (USB PD) Specification Parsing and Structuring System

This project provides a Python-based solution for parsing complex USB Power Delivery (USB PD) specification PDF documents and converting them into structured, machine-readable JSONL formats. The system is designed to intelligently extract the document's hierarchy, content, and metadata, making it suitable for ingestion into vector stores or for use in LLM-based document agents.

## Overview

Technical specifications like the USB PD standard are dense documents containing a mix of text, tables, figures, and hierarchical sections. The primary objective of this project is to build a prototype system that automates the extraction of this information, preserving its logical structure and metadata.

The script processes a raw PDF file and produces:
* A structured representation of the Table of Contents (ToC).
* A structured file of the entire document, section by section, including the text content of each section.
* A validation report to verify the accuracy of the parsing process.

## Deliverables & Output Files

The script generates the following files:

* **usb_pd_toc.jsonl**: A JSONL file where each line represents an entry from the document's Table of Contents.
* **usb_pd_spec.jsonl**: A comprehensive JSONL file containing every section from the document. Each line includes the section's metadata and its full text content.
* **usb_pd_metadata.jsonl**: A simple JSONL file containing high-level metadata about the source document.
* **validation_report.xlsx**: An Excel spreadsheet that validates the parsing job by comparing the ToC against the extracted sections and table counts. It includes a summary and details on any mismatches.

## Prerequisites

Before running the script, you need to have Python 3 installed, along with the following libraries:

* PyMuPDF (for PDF processing)
* pandas (for creating the Excel validation report)
* openpyxl (engine for pandas to write .xlsx files)

You can install all the required libraries with the following command:
bash
pip install PyMuPDF pandas openpyxl


## How to Run the Script

1.  *Place the PDF*: Put the target PDF file, named USB_PD_R3_2 V1.1 2024-10.pdf, in the same directory as the parse_spec.py script.
2.  *Execute the Script*: Open your terminal or command prompt, navigate to the project directory, and run the following command:
    bash
    python parse_spec.py
    
3.  *Check the Output*: The script will process the PDF and generate the four output files listed above in the same directory. The console will print progress updates as it completes each stage of the process.

## Code Structure Explained

The parse_spec.py script is organized into several key functions to ensure reusability and clarity:

* **main()**: The main entry point that controls the entire workflow, from opening the PDF to calling the parsing, extraction, and validation functions in the correct order.

* **parse_toc(doc, toc_pages)**: This function is responsible for parsing the Table of Contents pages. It uses regular expressions to extract the section_id, title, and page number for each entry.

* **extract_section_content(doc, toc_entries)**: The core of the content extraction logic. It uses the ToC entries as a guide to locate the start and end of each section within the document's body and extracts the corresponding text.

* **parse_list_of_tables(doc, toc_pages)**: A helper function for the validation step. It specifically parses the "List of Tables" from the ToC pages to get a count of all tables listed in the document's front matter.

* **count_tables_in_body(spec_entries)**: Complements the function above by iterating through the extracted text of all sections to find and count all mentions of tables in the document's body.

* **generate_validation_report(...)**: This function performs the final validation checks. It compares section counts and table counts from the ToC versus the parsed content, identifies any gaps or mismatches, and writes the results to the validation_report.xlsx file using the pandas library.

## JSONL Schema Overview

The generated JSONL files adhere to a structured schema for easy integration. Each line is a JSON object with the following fields:

| Field | Type | Description |
| :--- | :--- | :--- |
| doc_title | string | The title of the source document for reference. |
| section_id | string | The hierarchical section number (e.g., "2.1.2"). |
| title | string | The title of the section. |
| page | integer| The starting page number of the section. |
| level | integer| The hierarchical depth of the section (e.g., Chapter=1). |
| parent_id | string/null | The section_id of the immediate parent. |
| full_path | string | A combined string of the section ID and title. |
| tags | list | Optional list for semantic tags (currently unused). |
| text | string | **(In usb_pd_spec.jsonl only)** The full text content of the section. |
