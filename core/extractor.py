import os
import re
from typing import Dict, List, Any, Optional, Tuple
import pymupdf
import docx
import openpyxl

# Lazy singleton for RapidOCR so we only initialize it if a page is scanned
_ocr_engine = None

def get_ocr_engine():
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    return _ocr_engine


def parse_argentine_number(val: Any) -> Optional[float]:
    """
    Parses Argentine formatted accounting numbers:
    Examples:
      '1.234.567,89' -> 1234567.89
      '(1.234.567,89)' -> -1234567.89
      '-1.234.567,89' -> -1234567.89
      '$ 45.000,00' -> 45000.0
      '-' or '—' -> 0.0
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    
    s = str(val).strip()
    if not s:
        return None
    if s in ("-", "—", "–", "-.-", "0", "0,00", "$ -"):
        return 0.0

    # Remove currency symbols and spaces
    s = s.replace("$", "").replace("US$", "").replace("USD", "").strip()
    
    is_negative = False
    if (s.startswith("(") and s.endswith(")")) or s.startswith("-"):
        is_negative = True
        s = s.strip("()- ")

    # Check if it looks like a number
    # Case 1: Argentine format with dots for thousands and comma for decimals (e.g., 1.234,56 or 123.456)
    if re.match(r"^\d{1,3}(\.\d{3})+(,\d+)?$", s):
        s_clean = s.replace(".", "").replace(",", ".")
        try:
            num = float(s_clean)
            return -num if is_negative else num
        except ValueError:
            return None

    # Case 2: Simple number with comma decimal (e.g., 1234,56)
    if re.match(r"^\d+(,\d+)$", s):
        try:
            num = float(s.replace(",", "."))
            return -num if is_negative else num
        except ValueError:
            return None

    # Case 3: English format or raw integer (e.g., 1234567.89 or 123456)
    if re.match(r"^\d+(\.\d+)?$", s):
        try:
            num = float(s)
            return -num if is_negative else num
        except ValueError:
            return None

    return None


def extract_numbers_from_line(line: str) -> List[float]:
    """
    Extracts accounting numbers from a text line (typically 1 or 2 columns: Actual & Comparativo).
    Avoids matching dates (dd/mm/aaaa), CUITs, or Note numbers (Nota 1.3).
    """
    # Remove dates like 31/12/2025 or 30-06-2025
    cleaned = re.sub(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b", " ", line)
    # Remove CUITs
    cleaned = re.sub(r"\b\d{2}-\d{8}-\d\b", " ", cleaned)
    # Remove "Nota X.X" or "Anexo X.X" or "Art. XX" or "Ley XX.XXX" or "Res. P. XXX/XXXX"
    cleaned = re.sub(r"(?i)\b(nota|notas|anexo|anexos|art|artículo|ley|res|resolución|p|cd|n°|nro|t°|f°)\s*\.?\s*\d+([.,/]\d+)*", " ", cleaned)
    # Remove standalone years like 2024, 2025, 2026 when preceded by 'al', 'de', 'ejercicio'
    cleaned = re.sub(r"(?i)\b(de|del|año|ejercicio|al)\s+20[123]\d\b", " ", cleaned)

    # Match accounting tokens: e.g. (1.234.567,89) or 1.234.567,89 or 1.234.567 or -123.456,78
    pattern = r"(\(\s*\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?\s*\)|-?\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\(\s*\d+,\d{1,2}\s*\)|-?\d+,\d{1,2})"
    matches = re.findall(pattern, cleaned)
    results = []
    for m in matches:
        val = parse_argentine_number(m)
        if val is not None:
            results.append(val)
    return results


class DocumentExtractor:
    """Extracts text, pages, tables, and metadata from PDF (native or scanned), DOCX, or XLSX."""

    @staticmethod
    def extract(file_path: str) -> Dict[str, Any]:
        ext = os.path.splitext(file_path)[1].lower()
        file_size = os.path.getsize(file_path)
        file_name = os.path.basename(file_path)

        if ext == ".pdf":
            return DocumentExtractor._extract_pdf(file_path, file_name, file_size)
        elif ext == ".docx":
            return DocumentExtractor._extract_docx(file_path, file_name, file_size)
        elif ext == ".xlsx":
            return DocumentExtractor._extract_xlsx(file_path, file_name, file_size)
        else:
            raise ValueError(f"Formato de archivo no soportado: {ext}")

    @staticmethod
    def _group_pdf_words_into_table_lines(words: List[Tuple]) -> List[str]:
        """
        Groups PyMuPDF words (x0, y0, x1, y1, word, block_no, line_no, word_no)
        by vertical coordinate y0 (within 4.2pt tolerance) so multi-column tables
        (like side-by-side ESP Activo | Pasivo, or Rubro | 2026 | 2025) stay on the
        same line separated by ' | ' when horizontal gap > 14pt.
        """
        if not words:
            return []
        # Sort top-to-bottom (bucketed by 4pt), then left-to-right
        sorted_w = sorted(words, key=lambda w: (round(w[1] / 4.0), w[0]))
        rows: List[List[Tuple]] = []
        cur_row: List[Tuple] = []
        cur_y: Optional[float] = None

        for w in sorted_w:
            y0 = w[1]
            if cur_y is None or abs(y0 - cur_y) <= 4.5:
                cur_row.append(w)
                if cur_y is None:
                    cur_y = y0
                else:
                    cur_y = (cur_y * (len(cur_row) - 1) + y0) / len(cur_row)
            else:
                rows.append(cur_row)
                cur_row = [w]
                cur_y = y0
        if cur_row:
            rows.append(cur_row)

        num_pat = re.compile(r"^\(?-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}\)?$")
        result_lines: List[str] = []
        for r in rows:
            r.sort(key=lambda w: w[0])
            cells: List[str] = []
            cur_tokens: List[str] = [r[0][4]]
            prev_x1: float = r[0][2]
            for w in r[1:]:
                tok = w[4]
                prev_tok = cur_tokens[-1] if cur_tokens else ""
                gap = w[0] - prev_x1
                is_prev_num = bool(num_pat.match(prev_tok))
                is_curr_num = bool(num_pat.match(tok))
                # Split cell if horizontal gap > 14pt, OR if transitioning between an accounting amount and another column
                if (
                    gap > 14.0
                    or (is_prev_num and (is_curr_num or re.match(r"^[A-Za-zÁÉÍÓÚÑáéíóúñ]", tok)))
                    or (is_curr_num and gap > 4.5 and not prev_tok.endswith("$"))
                ):
                    cells.append(" ".join(cur_tokens).strip())
                    cur_tokens = [tok]
                else:
                    cur_tokens.append(tok)
                prev_x1 = max(prev_x1, w[2])
            if cur_tokens:
                cells.append(" ".join(cur_tokens).strip())
            cells = [c for c in cells if c]
            if len(cells) > 1:
                result_lines.append(" | ".join(cells))
            elif cells:
                result_lines.append(cells[0])
        return result_lines

    @staticmethod
    def _extract_pdf(file_path: str, file_name: str, file_size: int) -> Dict[str, Any]:
        doc = pymupdf.open(file_path)
        pages_data = []
        ocr_pages_count = 0
        full_text_parts = []
        structured_lines = []

        for idx, page in enumerate(doc):
            page_num = idx + 1
            words = page.get_text("words")
            raw_text = page.get_text("text").strip()
            used_ocr = False

            # Check if page has embedded table images or sparse native text
            imgs = page.get_images()
            needs_ocr = len(raw_text) < 60 or (len(imgs) >= 1 and len(raw_text) < 850 and not re.search(r"\d{1,3}(?:\.\d{3})+,\d{2}", raw_text))

            if not needs_ocr and words:
                grouped_table_lines = DocumentExtractor._group_pdf_words_into_table_lines(words)
                text = "\n".join(grouped_table_lines)
            else:
                text = raw_text

            if needs_ocr:
                try:
                    pix = page.get_pixmap(dpi=200)
                    img_bytes = pix.tobytes("png")
                    ocr = get_ocr_engine()
                    ocr_res, _ = ocr(img_bytes)
                    if ocr_res:
                        ocr_lines = []
                        for item in ocr_res:
                            box, txt, score = item[0], item[1], item[2]
                            y_top = box[0][1]
                            x_left = box[0][0]
                            x_right = box[1][0]
                            ocr_lines.append((y_top, x_left, x_right, txt))
                        ocr_lines.sort(key=lambda x: (round(x[0] / 12), x[1]))
                        grouped_lines = []
                        current_y = None
                        current_line = []
                        for y_top, x_left, x_right, txt in ocr_lines:
                            if current_y is None or abs(y_top - current_y) <= 14:
                                current_line.append((x_left, x_right, txt))
                                if current_y is None:
                                    current_y = y_top
                            else:
                                current_line.sort(key=lambda x: x[0])
                                grouped_lines.append(" | ".join(t[2] for t in current_line))
                                current_line = [(x_left, x_right, txt)]
                                current_y = y_top
                        if current_line:
                            current_line.sort(key=lambda x: x[0])
                            grouped_lines.append(" | ".join(t[2] for t in current_line))
                        ocr_text = "\n".join(grouped_lines)
                        if len(ocr_text) > len(text):
                            text = ocr_text
                        else:
                            text = text + "\n" + ocr_text
                        used_ocr = True
                        ocr_pages_count += 1
                except Exception as e:
                    if not text:
                        text = f"[Error OCR en página {page_num}: {e}]"

            pages_data.append({
                "page_number": page_num,
                "text": text,
                "used_ocr": used_ocr,
                "char_count": len(text)
            })
            full_text_parts.append(f"--- PÁGINA {page_num} ---\n{text}")
            for line in text.splitlines():
                if line.strip():
                    structured_lines.append({"page": page_num, "text": line.strip()})

        doc.close()
        return {
            "file_name": file_name,
            "file_size_bytes": file_size,
            "file_size_mb": round(file_size / (1024 * 1024), 2),
            "format": "PDF",
            "total_pages": len(pages_data),
            "ocr_pages_count": ocr_pages_count,
            "is_scanned": ocr_pages_count > 0,
            "pages": pages_data,
            "full_text": "\n\n".join(full_text_parts),
            "lines": structured_lines
        }

    @staticmethod
    def _extract_docx(file_path: str, file_name: str, file_size: int) -> Dict[str, Any]:
        d = docx.Document(file_path)
        lines = []
        full_text_parts = []

        for p in d.paragraphs:
            t = p.text.strip()
            if t:
                lines.append({"page": 1, "text": t})
                full_text_parts.append(t)

        for t_idx, table in enumerate(d.tables):
            full_text_parts.append(f"\n[TABLA {t_idx + 1}]")
            for row in table.rows:
                # Avoid duplicate merged cell text
                seen_cells = []
                for cell in row.cells:
                    c_txt = cell.text.strip().replace("\n", " ")
                    if c_txt and (not seen_cells or seen_cells[-1] != c_txt):
                        seen_cells.append(c_txt)
                if seen_cells:
                    row_str = " | ".join(seen_cells)
                    lines.append({"page": 1, "text": row_str})
                    full_text_parts.append(row_str)

        full_text = "\n".join(full_text_parts)
        return {
            "file_name": file_name,
            "file_size_bytes": file_size,
            "file_size_mb": round(file_size / (1024 * 1024), 2),
            "format": "DOCX",
            "total_pages": max(1, len(lines) // 45),
            "ocr_pages_count": 0,
            "is_scanned": False,
            "pages": [{"page_number": 1, "text": full_text, "used_ocr": False, "char_count": len(full_text)}],
            "full_text": full_text,
            "lines": lines
        }

    @staticmethod
    def _extract_xlsx(file_path: str, file_name: str, file_size: int) -> Dict[str, Any]:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        pages_data = []
        full_text_parts = []
        lines = []

        for idx, sheet_name in enumerate(wb.sheetnames):
            ws = wb[sheet_name]
            sheet_lines = []
            for r in range(1, ws.max_row + 1):
                row_vals = []
                for c in range(1, min(20, ws.max_column + 1)):
                    v = ws.cell(r, c).value
                    if v is not None and str(v).strip() != "":
                        if isinstance(v, float):
                            # Format nicely as Argentine number for unified downstream parsing
                            formatted = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            row_vals.append(formatted)
                        else:
                            row_vals.append(str(v).strip().replace("\n", " "))
                if row_vals:
                    line_str = " | ".join(row_vals)
                    sheet_lines.append(line_str)
                    lines.append({"page": idx + 1, "sheet": sheet_name, "text": line_str})

            sheet_text = f"--- HOJA {sheet_name} ---\n" + "\n".join(sheet_lines)
            pages_data.append({
                "page_number": idx + 1,
                "sheet_name": sheet_name,
                "text": sheet_text,
                "used_ocr": False,
                "char_count": len(sheet_text)
            })
            full_text_parts.append(sheet_text)

        return {
            "file_name": file_name,
            "file_size_bytes": file_size,
            "file_size_mb": round(file_size / (1024 * 1024), 2),
            "format": "XLSX",
            "total_pages": len(pages_data),
            "ocr_pages_count": 0,
            "is_scanned": False,
            "pages": pages_data,
            "full_text": "\n\n".join(full_text_parts),
            "lines": lines
        }
