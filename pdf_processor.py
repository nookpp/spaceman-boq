import os
import re
import io
from typing import List, Dict, Any, Optional
import pymupdf  # PyMuPDF
from PIL import Image

def get_pdf_document(file_source: Any) -> pymupdf.Document:
    """Loads a PyMuPDF document from file path or bytes/file-like buffer."""
    if isinstance(file_source, str):
        return pymupdf.open(file_source)
    elif isinstance(file_source, bytes):
        return pymupdf.open(stream=file_source, filetype="pdf")
    elif hasattr(file_source, "read"):
        data = file_source.read()
        if hasattr(file_source, "seek"):
            file_source.seek(0)
        return pymupdf.open(stream=data, filetype="pdf")
    else:
        raise ValueError("Unsupported file source type")

def get_pdf_summary(doc: pymupdf.Document) -> Dict[str, Any]:
    """Returns general information about the PDF."""
    total_pages = len(doc)
    metadata = doc.metadata or {}
    
    # Quick scan for pages with drawings, specifications, schedules
    page_highlights = []
    for i in range(min(total_pages, 20)):
        text = doc[i].get_text()
        first_few = [l.strip() for l in text.splitlines() if l.strip()][:4]
        if first_few:
            page_highlights.append({"page": i + 1, "preview": " | ".join(first_few)})

    return {
        "total_pages": total_pages,
        "title": metadata.get("title") or "Drawing Document",
        "author": metadata.get("author") or "-",
        "page_highlights": page_highlights
    }

def render_page_image(doc: pymupdf.Document, page_num: int, dpi: int = 150) -> Image.Image:
    """Renders a PDF page to a PIL Image (page_num is 1-indexed)."""
    if page_num < 1 or page_num > len(doc):
        raise IndexError("Page number out of range")
    page = doc[page_num - 1]
    pix = page.get_pixmap(dpi=dpi)
    img_data = pix.tobytes("png")
    return Image.open(io.BytesIO(img_data))

def scan_drawing_schedule_and_codes(doc: pymupdf.Document) -> List[Dict[str, Any]]:
    """
    Scans the document for Joinery/Built-in, Door, and Specification codes.
    Specifically searches for patterns like J1-xx, J2-xx, D1-xx, D2-xx, E1-xx.
    """
    detected_items = []
    code_pattern = re.compile(r'\b([JjDdEe]\d+[-\.]\d{2})\b')
    
    # Priority scan on index pages (like page 8, 1-10) and detail pages
    for i, page in enumerate(doc):
        text = page.get_text()
        matches = code_pattern.findall(text)
        if matches:
            lines = [l.strip() for l in text.splitlines() if l.strip()]
            for m in set(matches):
                code_upper = m.upper().replace('.', '-')
                # Find contextual line mentioning this code
                context_line = ""
                for line in lines:
                    if code_upper in line.upper():
                        context_line = line
                        break
                
                detected_items.append({
                    "code": code_upper,
                    "page": i + 1,
                    "context": context_line or "Found in drawing",
                    "type": "Joinery" if code_upper.startswith("J") else ("Door" if code_upper.startswith("D") else "Electrical")
                })
                
    # Deduplicate while preserving earliest page
    unique_codes = {}
    for item in detected_items:
        c = item["code"]
        if c not in unique_codes:
            unique_codes[c] = item
            
    return sorted(list(unique_codes.values()), key=lambda x: x["code"])

def extract_pages_text(doc: pymupdf.Document, page_numbers: List[int]) -> str:
    """Extracts text from specified pages (1-indexed)."""
    extracted = []
    for p in page_numbers:
        if 1 <= p <= len(doc):
            t = doc[p - 1].get_text()
            extracted.append(f"--- [Page {p}] ---\n{t.strip()}\n")
    return "\n".join(extracted)
