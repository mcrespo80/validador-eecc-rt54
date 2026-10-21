import os
import sys

sys.path.insert(0, r"C:\Users\mcres\validador")
from core.extractor import DocumentExtractor
from core.validator import EECCValidator
from core.report_pdf import generate_audit_report_pdf

def run_tests():
    v = EECCValidator()

    # Test 1: MODELO DE INFORME DE AUDITORÍA.pba.docx (CET. LEO KANNER S.A.)
    p1 = r"C:\Users\mcres\validador\knowledge_base\modelos_consejo\MODELO DE INFORME DE AUDITORÍA.pba.docx"
    ext1 = DocumentExtractor.extract(p1)
    res1 = v.validate(ext1)
    print("=== TEST 1: CET. LEO KANNER S.A. ===")
    print("Summary:", res1["summary"])
    print("Meta:", res1["entity_metadata"])
    for c in res1["checks"]:
        print(f"  [{c['status']}] {c['code']}: {c['title']} -> {c['detail'][:110]}")

    # Test 2: RT 54 Modelo de EECC para EP sin fines de lucro.docx
    p2 = r"C:\Users\mcres\validador\knowledge_base\modelos_consejo\RT 54 Modelo de EECC para EP sin fines de lucro.docx"
    ext2 = DocumentExtractor.extract(p2)
    res2 = v.validate(ext2)
    print("\n=== TEST 2: Modelo RT 54 EP sin fines de lucro ===")
    print("Summary:", res2["summary"])
    print("Meta:", res2["entity_metadata"])
    for c in res2["checks"]:
        print(f"  [{c['status']}] {c['code']}: {c['title']} -> {c['detail'][:110]}")

    # Test 3: Generate PDF Audit Report
    pdf_out = r"C:\Users\mcres\validador\reports\Reporte_Ejemplo_Validacion.pdf"
    generate_audit_report_pdf(res2, pdf_out, logo_path=r"C:\Users\mcres\validador\data\brand\emblem_clean.png")
    print("\nPDF Report generated at:", pdf_out, "Size:", os.path.getsize(pdf_out))

if __name__ == "__main__":
    run_tests()
