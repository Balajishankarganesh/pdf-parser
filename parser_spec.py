# parser_spec.py  (OOP refactor – robust)
# Requires: pymupdf (fitz), pandas, openpyxl
#   pip install pymupdf pandas openpyxl

import fitz
import json
import re
import os
import logging
from dataclasses import dataclass, asdict
from typing import List, Optional, Protocol, Tuple
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


# =========================
# Data Models
# =========================

@dataclass
class Section:
    doc_title: str
    section_id: str
    title: str
    page_start_1idx: int       # 1-indexed page
    level: int
    parent_id: Optional[str]
    full_path: str
    tags: List[str]
    text: Optional[str] = None


# =========================
# Core Interfaces
# =========================

class Parser(Protocol):
    def parse(self, *args, **kwargs): ...


# =========================
# Utilities
# =========================

class PDFLoader:
    def __init__(self, file_path: str):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        self.doc = fitz.open(file_path)

    def get_text(self, pages_0idx: List[int]) -> str:
        return "".join(self.doc[p].get_text("text") for p in pages_0idx)

    def page_text(self, page_0idx: int) -> str:
        return self.doc[page_0idx].get_text("text")

    @property
    def page_count(self) -> int:
        return len(self.doc)


# =========================
# TOC Strategies
# =========================

class TOCStrategy(Protocol):
    def read(self, pdf: PDFLoader, doc_title: str) -> List[Section]: ...


class BookmarkTOCStrategy:
    """Use PDF bookmarks (stable if present)."""
    _secid_from_title = re.compile(r"^(?P<id>\d+(?:\.\d+)*)\s+(?P<title>.+?)\s*$")

    def read(self, pdf: PDFLoader, doc_title: str) -> List[Section]:
        toc = pdf.doc.get_toc(simple=True)  # [level, title, page]
        if not toc:
            return []

        sections: List[Section] = []
        for level, title, page in toc:
            title = title.strip()
            m = self._secid_from_title.match(title)
            if m:
                sec_id = m.group("id")
                sec_title = m.group("title").strip()
            else:
                sec_id = ""
                sec_title = title

            sections.append(
                Section(
                    doc_title=doc_title,
                    section_id=sec_id,
                    title=sec_title,
                    page_start_1idx=page,
                    level=level,
                    parent_id=sec_id.rsplit(".", 1)[0] if sec_id and "." in sec_id else None,
                    full_path=(f"{sec_id} {sec_title}".strip()),
                    tags=[],
                )
            )

        # fill missing ids
        counters = {}
        for s in sections:
            if s.section_id:
                continue
            counters[s.level] = counters.get(s.level, 0) + 1
            s.section_id = f"L{s.level}-{counters[s.level]:03d}"
            s.parent_id = None

        return sections


class RegexTOCStrategy:
    """Fallback ToC parser using regex on ToC pages."""
    def __init__(self, toc_pages_hint: Optional[List[int]] = None):
        self.toc_pages_hint = toc_pages_hint
        self.pattern = re.compile(
            r"^(?P<id>\d+(?:\.\d+)*)\s+(?P<title>[^\n\.]+?)\s(?:\.{3,}|\s\.+\s*)?(?P<page>\d+)\s*$",
            re.MULTILINE
        )

    def _guess_toc_pages(self, pdf: PDFLoader, scan_first_n: int = 40) -> List[int]:
        hits = []
        for i in range(min(scan_first_n, pdf.page_count)):
            txt = pdf.page_text(i)
            if re.search(r"\b(Table of Contents|Contents)\b", txt, flags=re.I):
                hits.append(i)
        if not hits:
            return list(range(0, min(10, pdf.page_count)))
        start = max(hits[0] - 1, 0)
        end = min(hits[0] + 8, pdf.page_count - 1)
        return list(range(start, end + 1))

    def read(self, pdf: PDFLoader, doc_title: str) -> List[Section]:
        pages = self.toc_pages_hint if self.toc_pages_hint else self._guess_toc_pages(pdf)
        text = pdf.get_text(pages)

        sections: List[Section] = []
        for m in self.pattern.finditer(text):
            sec_id = m.group("id").strip()
            title = m.group("title").strip()
            page_1idx = int(m.group("page"))
            sections.append(
                Section(
                    doc_title=doc_title,
                    section_id=sec_id,
                    title=title,
                    page_start_1idx=page_1idx,
                    level=sec_id.count(".") + 1,
                    parent_id=sec_id.rsplit(".", 1)[0] if "." in sec_id else None,
                    full_path=f"{sec_id} {title}",
                    tags=[],
                )
            )
        return sections


class TOCReader:
    def __init__(self, doc_title: str, toc_pages_hint: Optional[List[int]] = None):
        self.doc_title = doc_title
        self.strategies: List[TOCStrategy] = [
            BookmarkTOCStrategy(),
            RegexTOCStrategy(toc_pages_hint),
        ]

    def read(self, pdf: PDFLoader) -> List[Section]:
        for strat in self.strategies:
            sections = strat.read(pdf, self.doc_title)
            if sections:
                logging.info(f"TOC via {strat.__class__.__name__}: {len(sections)} sections")
                return sections
        logging.warning("No TOC found.")
        return []


# =========================
# Spec Parser (page ranges)
# =========================

class SpecParser:
    def extract_sections_text(self, pdf: PDFLoader, sections: List[Section]) -> List[Section]:
        if not sections:
            return sections

        sections_sorted = sorted(sections, key=lambda s: s.page_start_1idx)
        last_page = pdf.page_count

        ranges: List[Tuple[int, int]] = []
        for i, s in enumerate(sections_sorted):
            start = s.page_start_1idx
            end = (sections_sorted[i + 1].page_start_1idx - 1) if (i + 1 < len(sections_sorted)) else last_page
            ranges.append((start, max(start, end)))

        for s, (start, end) in zip(sections_sorted, ranges):
            pages_0idx = list(range(start - 1, end))
            s.text = pdf.get_text(pages_0idx).strip()

        return sections_sorted


# =========================
# Tables
# =========================

class TableParser:
    BODY_TABLE_RE = re.compile(
        r"\bTable\s+[A-Z]?\d+(?:[-\u2012\u2013\u2014\u2015]\d+)?[A-Za-z0-9\-]*",
        flags=re.IGNORECASE
    )

    def list_from_list_of_tables(self, pdf: PDFLoader) -> List[str]:
        pages = []
        for i in range(min(50, pdf.page_count)):
            if re.search(r"\bList of Tables\b", pdf.page_text(i), flags=re.I):
                pages.extend(range(i, min(i + 6, pdf.page_count)))
                break
        if not pages:
            return []

        text = pdf.get_text(pages)
        items = re.findall(r"(Table\s+[A-Z]?\d+(?:[-\u2012\u2013\u2014\u2015]\d+)?)", text, flags=re.I)
        return sorted(set(x.strip() for x in items), key=str.lower)

    def count_in_body(self, sections: List[Section]) -> List[str]:
        found = set()
        for s in sections:
            if not s.text:
                continue
            for m in self.BODY_TABLE_RE.finditer(s.text):
                found.add(m.group(0).strip())
        return sorted(found, key=str.lower)


# =========================
# Reporting
# =========================

class ReportGenerator:
    def __init__(self, output_file: str):
        self.output_file = output_file

    def generate(self, toc: List[Section], spec: List[Section],
                 toc_tables: List[str], body_tables: List[str]) -> None:
        logging.info("Generating validation report...")

        summary_df = pd.DataFrame({
            "Metric": [
                "Total Sections in ToC",
                "Total Sections Parsed (with text)",
                "Sections Match (IDs)",
                "Total Tables in ToC List",
                "Total Tables Found in Body",
            ],
            "Value": [
                len(toc),
                len(spec),
                "Yes" if {s.section_id for s in toc} == {s.section_id for s in spec} else "No",
                len(toc_tables),
                len(body_tables),
            ],
        })

        toc_ids = {s.section_id for s in toc}
        spec_ids = {s.section_id for s in spec}
        mismatch_df = pd.DataFrame({
            "In ToC not Parsed": sorted(toc_ids - spec_ids),
            "Parsed not in ToC": sorted(spec_ids - toc_ids),
        })

        with pd.ExcelWriter(self.output_file) as writer:
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            mismatch_df.to_excel(writer, sheet_name="Mismatches", index=False)

        logging.info(f"Report written: {self.output_file}")


# =========================
# Persistence
# =========================

class JSONLWriter:
    @staticmethod
    def write(filename: str, items: List[object]) -> None:
        with open(filename, "w", encoding="utf-8") as f:
            for it in items:
                if isinstance(it, dict):
                    f.write(json.dumps(it, ensure_ascii=False) + "\n")
                else:
                    f.write(json.dumps(asdict(it), ensure_ascii=False) + "\n")


# =========================
# Orchestrator
# =========================

class USBPDParser:
    def __init__(self,
                 pdf_path: str,
                 doc_title: str,
                 report_path: str = "validation_report.xlsx",
                 toc_pages_hint: Optional[List[int]] = None):
        self.pdf = PDFLoader(pdf_path)
        self.doc_title = doc_title
        self.toc_reader = TOCReader(doc_title, toc_pages_hint=toc_pages_hint)
        self.spec_parser = SpecParser()
        self.table_parser = TableParser()
        self.reporter = ReportGenerator(report_path)

    def run(self):
        toc_sections = self.toc_reader.read(self.pdf)
        if not toc_sections:
            logging.error("No TOC sections found. Aborting.")
            return

        JSONLWriter.write("usb_pd_toc.jsonl", toc_sections)

        spec_sections = self.spec_parser.extract_sections_text(self.pdf, toc_sections)
        JSONLWriter.write("usb_pd_spec.jsonl", spec_sections)

        metadata = {"doc_title": self.doc_title, "pages": self.pdf.page_count}
        JSONLWriter.write("usb_pd_metadata.jsonl", [metadata])

        toc_tables = self.table_parser.list_from_list_of_tables(self.pdf)
        body_tables = self.table_parser.count_in_body(spec_sections)

        self.reporter.generate(toc_sections, spec_sections, toc_tables, body_tables)

if __name__ == "__main__":
    import glob, sys

    cwd = os.getcwd()
    pdf_candidates = glob.glob(os.path.join(cwd, "*.pdf"))

    if not pdf_candidates:
        logging.error("No PDF files found in the current folder: %s", cwd)
        logging.error("Place the PDF (e.g. USB_PD_R3_2_V1.1_2024_10.pdf) into this folder and rerun.")
        sys.exit(1)

    # Prefer a PDF with 'usb' in the filename, otherwise use the only PDF if one exists
    pdf_path = None
    for c in pdf_candidates:
        if "usb" in os.path.basename(c).lower():
            pdf_path = c
            break

    if not pdf_path:
        if len(pdf_candidates) == 1:
            pdf_path = pdf_candidates[0]
        else:
            logging.error("Multiple PDFs found in folder; please set pdf_path manually in the script.")
            logging.error("Found: %s", pdf_candidates)
            sys.exit(1)

    print("Using PDF:", pdf_path)

    parser = USBPDParser(
        pdf_path=pdf_path,
        doc_title="USB Power Delivery Specification Rev 3.2 V1.1",
        report_path="validation_report.xlsx",
        toc_pages_hint=None
    )
    parser.run()



