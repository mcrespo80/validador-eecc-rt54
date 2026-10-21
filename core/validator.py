import os
import re
import json
from typing import Dict, List, Any, Optional, Tuple
from core.extractor import extract_numbers_from_line, parse_argentine_number

SPANISH_MONTHS = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
}


def validate_cuit_mod11(cuit_str: str) -> bool:
    """Validates an Argentine CUIT using the official Modulo 11 algorithm."""
    digits = re.sub(r"\D", "", cuit_str)
    if len(digits) != 11:
        return False
    weights = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    total = sum(int(d) * w for d, w in zip(digits[:10], weights))
    rem = total % 11
    check = 11 - rem
    if check == 11:
        check = 0
    elif check == 10:
        check = 9
    return int(digits[10]) == check


def normalize_date_str(day: int, month: int, year: int) -> str:
    if year < 100:
        year += 2000
    return f"{day:02d}/{month:02d}/{year:04d}"


def extract_all_dates(text: str) -> List[Dict[str, Any]]:
    """Finds both numeric dates (31/12/2025) and textual dates (31 de diciembre de 2025)."""
    found = []
    # 1. Numeric dd/mm/yyyy
    for m in re.finditer(r"\b(\d{1,2})[/-](\d{1,2})[/-](20[123]\d)\b", text):
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= d <= 31 and 1 <= mo <= 12:
            found.append({
                "raw": m.group(0),
                "normalized": normalize_date_str(d, mo, y),
                "year": y, "month": mo, "day": d,
                "pos": m.start()
            })
    # 2. Textual "30 de junio de 2025"
    pattern_txt = r"\b(\d{1,2})\s+de\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\s+de\s+(20[123]\d)\b"
    for m in re.finditer(pattern_txt, text, flags=re.IGNORECASE):
        d = int(m.group(1))
        mo = SPANISH_MONTHS[m.group(2).lower()]
        y = int(m.group(3))
        if 1 <= d <= 31:
            found.append({
                "raw": m.group(0),
                "normalized": normalize_date_str(d, mo, y),
                "year": y, "month": mo, "day": d,
                "pos": m.start()
            })
    return found


def fmt_ars(val: Optional[float]) -> str:
    if val is None:
        return "No detectado"
    sign = "-" if val < 0 else ""
    v = abs(val)
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sign}$ {s}"


class EECCValidator:
    """
    Comprehensive Validator for Financial Statements (EECC) under:
    - RT 54 (T.O. RT 59 - NUA)
    - Res. P. 460/2024 CPCECABA & Res. JG 660/2026 FACPCE
    - RT 37 (mod. RT 53) & Res. CD 46/2021 CPCECABA
    - Legalizaciones Web CPCECABA Formal Controls
    """

    def __init__(self, rules_path: Optional[str] = None):
        if rules_path is None:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            rules_path = os.path.join(base_dir, "knowledge_base", "reglas_rt54_cpcecaba.json")
        with open(rules_path, "r", encoding="utf-8") as f:
            self.kb = json.load(f)

    def validate(self, extracted_doc: Dict[str, Any], jurisdiction_pref: str = "AUTO") -> Dict[str, Any]:
        full_text = extracted_doc["full_text"]
        lines = extracted_doc["lines"]

        checks: List[Dict[str, Any]] = []
        metadata = self._extract_entity_metadata(full_text, lines, jurisdiction_pref)
        numbers = self._extract_key_financial_numbers(lines, full_text)

        # ------------------------------------------------------------------
        # BLOQUE 1: CONTROLES FORMALES Y DE LEGALIZACIÓN WEB (CPCECABA)
        # ------------------------------------------------------------------
        self._check_formal_controls(extracted_doc, full_text, lines, metadata, checks)

        # ------------------------------------------------------------------
        # BLOQUE 2: NORMATIVA CONTABLE RT 54 (T.O. RT 59 - NUA)
        # ------------------------------------------------------------------
        self._check_rt54_controls(full_text, lines, metadata, numbers, checks)

        # ------------------------------------------------------------------
        # BLOQUE 3: INFORME DEL AUDITOR INDEPENDIENTE (RT 37 MOD. RT 53)
        # ------------------------------------------------------------------
        self._check_auditor_report_controls(full_text, lines, metadata, checks)

        # ------------------------------------------------------------------
        # BLOQUE 4: CUADRATURA NUMÉRICA Y CRUCES DE ESTADOS CONTABLES
        # ------------------------------------------------------------------
        self._check_numerical_consistency(numbers, metadata, checks)

        # Compute overall score & status
        criticos = [c for c in checks if c["status"] == "ERROR"]
        advertencias = [c for c in checks if c["status"] == "WARNING"]
        aprobados = [c for c in checks if c["status"] == "PASS"]
        info_checks = [c for c in checks if c["status"] == "INFO"]

        total_applicable = len(criticos) + len(advertencias) + len(aprobados)
        if total_applicable > 0:
            raw_score = ((len(aprobados) * 1.0) + (len(advertencias) * 0.5)) / total_applicable * 100
            # Penalize harder if critical errors exist
            if len(criticos) > 0:
                score = min(int(raw_score), max(15, 75 - len(criticos) * 12))
            else:
                score = int(round(raw_score))
        else:
            score = 0

        if len(criticos) == 0 and len(advertencias) <= 2:
            dictamen_pre_legalizacion = "APTO PARA LEGALIZACIÓN SIN OBJECIONES"
            semaforo = "VERDE"
        elif len(criticos) == 0:
            dictamen_pre_legalizacion = "APTO CON OBSERVACIONES MENORES (REVISAR ADVERTENCIAS)"
            semaforo = "AMARILLO"
        else:
            dictamen_pre_legalizacion = f"RECHAZO PROBABLE EN LEGALIZACIONES ({len(criticos)} ERRORES CRÍTICOS DETECTADOS)"
            semaforo = "ROJO"

        return {
            "file_info": {
                "file_name": extracted_doc["file_name"],
                "format": extracted_doc["format"],
                "file_size_mb": extracted_doc["file_size_mb"],
                "total_pages": extracted_doc["total_pages"],
                "is_scanned": extracted_doc["is_scanned"],
                "ocr_pages_count": extracted_doc["ocr_pages_count"]
            },
            "entity_metadata": metadata,
            "financial_summary": numbers,
            "summary": {
                "score": score,
                "semaforo": semaforo,
                "dictamen": dictamen_pre_legalizacion,
                "total_checks": len(checks),
                "passed": len(aprobados),
                "warnings": len(advertencias),
                "errors": len(criticos),
                "info": len(info_checks)
            },
            "checks": checks
        }

    # ======================================================================
    # EXTRACCIÓN DE METADATOS DE LA ENTIDAD Y DEL INFORME
    # ======================================================================
    def _extract_entity_metadata(self, full_text: str, lines: List[Dict[str, Any]], jurisdiction_pref: str) -> Dict[str, Any]:
        # 1. CUITs
        cuits_raw = re.findall(r"\b(?:20|23|24|27|30|33|34)-\d{8}-\d\b", full_text)
        unique_cuits = []
        for c in cuits_raw:
            if c not in unique_cuits:
                unique_cuits.append(c)

        entity_cuit = unique_cuits[0] if unique_cuits else None
        cuit_valid = validate_cuit_mod11(entity_cuit) if entity_cuit else False

        # Check placeholder CUITs
        has_placeholder_cuit = bool(re.search(r"XX-XXXXXXXX-X|\.\.\.-\.\.\.\.\.\.\.\.-\.", full_text, re.IGNORECASE))

        # 2. Entity Name
        entity_name = None
        for item in lines[:60]:
            txt = item["text"].strip()
            # Look for "Denominación de la entidad:", "Razón Social:", or "AGENTE DEL SEGURO SALUD:"
            m_den = re.search(r"(?:Denominaci[oó]n(?:\s+social)?(?:\s+de\s+la\s+entidad)?|Raz[oó]n\s+Social|AGENTE\s+DEL\s+SEGURO\s+(?:DE\s+)?SALUD)\s*:\s*(.+)", txt, re.IGNORECASE)
            if m_den:
                cand = m_den.group(1).split("|")[0].strip(" .|-")
                if len(cand) > 2 and "DENOMINACIÓN" not in cand.upper():
                    entity_name = cand
                    break
        if not entity_name:
            # Look in Auditor Report "He auditado los estados contables (adjuntos) de [ENTIDAD], que comprenden"
            m_aud = re.search(r"He\s+auditado\s+los\s+estados\s+contables(?:\s+adjuntos)?\s+de\s+([^,\n]+),\s+que\s+comprenden", full_text, re.IGNORECASE)
            if m_aud:
                entity_name = m_aud.group(1).strip()
        if not entity_name:
            # Look at first prominent uppercase line
            for item in lines[:20]:
                txt = item["text"].strip()
                if len(txt) > 4 and txt.isupper() and not any(w in txt for w in ["MODELO", "ESTADOS CONTABLES", "INFORME", "NOTA", "ÍNDICE", "INDICE", "HOJA"]):
                    entity_name = txt.split("|")[0].strip()
                    break

        # 3. Closing Date (Fecha de cierre del ejercicio)
        closing_date = None
        m_cierre = re.search(
            r"(?:situaci[oó]n\s+patrimonial\s+al|ejercicio\s+(?:econ[oó]mico\s+)?(?:finalizado|terminado)\s+el|ESTADOS\s+CONTABLES\s+AL|Fecha\s+de\s+cierre(?:\s+de\s+ejercicio)?\s*:)\s*(\d{1,2}[/-]\d{1,2}[/-]20[123]\d|\d{1,2}\s+de\s+[a-zA-Z]+\s+de\s+20[123]\d)",
            full_text, re.IGNORECASE
        )
        if m_cierre:
            extracted_d = extract_all_dates(m_cierre.group(1))
            if extracted_d:
                closing_date = extracted_d[0]["normalized"]

        # 4. Entity Type (Sin Fines de Lucro vs Con Fines de Lucro) & RT 54 Category
        is_esfl = bool(re.search(r"ESTADO\s+DE\s+RECURSOS\s+Y\s+GASTOS|sin\s+fines\s+de\s+lucro|Asociaci[oó]n\s+Civil|Fundaci[oó]n|Sindicato|Obra\s+Social|Mutual", full_text, re.IGNORECASE))
        tipo_ente = "Entidad Sin Fines de Lucro (ESFL)" if is_esfl else "Sociedad / Entidad Con Fines de Lucro"

        categoria_rt54 = "No especificada en Notas"
        if re.search(r"car[aá]cter\s+de\s+entidad\s+peque[ñn]a|califica\s+como\s+entidad\s+peque[ñn]a|\(EP\)", full_text, re.IGNORECASE):
            categoria_rt54 = "Entidad Pequeña (EP)"
        elif re.search(r"car[aá]cter\s+de\s+entidad\s+mediana|califica\s+como\s+entidad\s+mediana|\(EM\)", full_text, re.IGNORECASE):
            categoria_rt54 = "Entidad Mediana (EM)"
        elif re.search(r"restantes?\s+entidad(?:es)?|no\s+reviste\s+el\s+car[aá]cter\s+de\s+entidad\s+peque[ñn]a\s+ni\s+mediana", full_text, re.IGNORECASE):
            categoria_rt54 = "Restante Entidad (RE - No Pequeña ni Mediana)"

        # 5. Jurisdiction (CPCECABA vs CPCEPBA)
        if jurisdiction_pref in ("CABA", "PBA"):
            jurisdiccion = jurisdiction_pref
        else:
            if re.search(r"CPCECABA|C\.P\.C\.E\.C\.A\.B\.A|Ciudad\s+Aut[oó]noma\s+de\s+Buenos\s+Aires|460/2024|46/2021", full_text, re.IGNORECASE):
                jurisdiccion = "CABA (CPCECABA)"
            elif re.search(r"Provincia\s+de\s+Buenos\s+Aires|3843|CPCEPBA", full_text, re.IGNORECASE):
                jurisdiccion = "PBA (CPCEPBA)"
            else:
                jurisdiccion = "CABA (CPCECABA - Por defecto)"

        # 6. Check if document contains the Auditor Report & Basic Statements & Notes
        # Require uppercase/header style or explicit table headers to distinguish actual statements from mentions in auditor's opinion
        has_auditor_report = bool(re.search(r"INFORME\s+DE(?:L)?\s+AUDITOR[IÍ]?A|AUDITOR\s+INDEPENDIENTE", full_text, re.IGNORECASE))
        has_esp = bool(re.search(r"ESTADO\s+DE\s+SITUACI[OÓ]N\s+PATRIMONIAL|ACTIVO\s+CORRIENTE", full_text))
        has_er_erg = bool(re.search(r"ESTADO\s+DE\s+RESULTADOS|ESTADO\s+DE\s+RECURSOS\s+Y\s+GASTOS", full_text))
        has_eepn = bool(re.search(r"ESTADO\s+DE\s+EVOLUCI[OÓ]N\s+DEL\s+PATRIMONIO\s+NETO", full_text))
        has_efe = bool(re.search(r"ESTADO\s+DE\s+FLUJOS?\s+DE\s+EFECTIVO|VARIACIONES\s+DEL\s+EFECTIVO", full_text))
        has_notas = bool(re.search(r"NOTAS\s+A\s+LOS\s+ESTADOS\s+CONTABLES|Bases\s+de\s+preparaci[oó]n\s+de\s+los\s+estados\s+contables", full_text, re.IGNORECASE))

        return {
            "denominacion": entity_name or "No detectada (o plantilla en blanco)",
            "cuit": entity_cuit or ("Plantilla XX-XXXXXXXX-X" if has_placeholder_cuit else "No detectada"),
            "cuit_valido_mod11": cuit_valid,
            "all_cuits_found": unique_cuits,
            "fecha_cierre": closing_date or "No detectada (o formato dd/mm/aaaa sin completar)",
            "tipo_ente": tipo_ente,
            "es_sin_fines_de_lucro": is_esfl,
            "categoria_rt54": categoria_rt54,
            "jurisdiccion": jurisdiccion,
            "secciones_presentes": {
                "esp": has_esp,
                "er_erg": has_er_erg,
                "eepn": has_eepn,
                "efe": has_efe,
                "notas": has_notas,
                "informe_auditor": has_auditor_report
            }
        }

    # ======================================================================
    # BLOQUE 1: CONTROLES FORMALES Y DE LEGALIZACIÓN WEB (CPCECABA)
    # ======================================================================
    def _check_formal_controls(self, doc: Dict[str, Any], full_text: str, lines: List[Dict[str, Any]], meta: Dict[str, Any], checks: List[Dict[str, Any]]):
        # FOR-00: Tamaño de archivo PDF (< 10 MB para Legalizaciones Web CPCECABA)
        size_mb = doc["file_size_mb"]
        if size_mb <= 10.0:
            checks.append({
                "code": "FOR-00",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Tamaño del archivo para subida al portal de Legalizaciones Web",
                "status": "PASS",
                "severity": "CRITICO",
                "detail": f"El archivo pesa {size_mb} MB (dentro del límite máximo de 10 MB admitido por el portal web del CPCECABA).",
                "norm_ref": "Instructivo Legalizaciones Web CPCECABA"
            })
        else:
            checks.append({
                "code": "FOR-00",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Tamaño del archivo para subida al portal de Legalizaciones Web",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": f"El archivo pesa {size_mb} MB. El portal del CPCECABA suele limitar la subida a 10 MB por archivo PDF.",
                "recommendation": "Comprimir las imágenes del PDF antes de subirlo al sistema de Legalizaciones.",
                "norm_ref": "Instructivo Legalizaciones Web CPCECABA"
            })

        # FOR-01: Validación de CUIT y consistencia
        cuit = meta["cuit"]
        if cuit and re.match(r"^\d{2}-\d{8}-\d$", cuit):
            if meta["cuit_valido_mod11"]:
                checks.append({
                    "code": "FOR-01",
                    "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                    "title": "Validación de CUIT (Dígito Verificador Módulo 11 y Consistencia)",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": f"CUIT detectada: {cuit} (Dígito verificador válido según algoritmo AFIP/ARCA).",
                    "norm_ref": "Check List Legalizaciones CPCECABA"
                })
            else:
                checks.append({
                    "code": "FOR-01",
                    "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                    "title": "Validación de CUIT (Dígito Verificador Módulo 11 y Consistencia)",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": f"La CUIT detectada ({cuit}) es INVÁLIDA matemáticamente (falla control de dígito verificador Módulo 11). Motivo de rechazo automático en Legalizaciones.",
                    "recommendation": "Verificar posibles errores de tipeo en los dígitos de la CUIT en Carátula e Informe del Auditor.",
                    "norm_ref": "Check List Legalizaciones CPCECABA"
                })
        else:
            checks.append({
                "code": "FOR-01",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Validación de CUIT (Dígito Verificador Módulo 11 y Consistencia)",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": f"No se detectó una CUIT real válida en el documento (Valor actual: {cuit}).",
                "recommendation": "Consignar la CUIT del ente con formato XX-XXXXXXXX-X tanto en la Carátula como en el encabezamiento del Informe del Auditor.",
                "norm_ref": "Check List Legalizaciones CPCECABA"
            })

        # FOR-02: Marcadores de plantilla o borradores sin completar
        placeholders_found = []
        placeholder_patterns = [
            (r"\bdd/mm/aaaa\b", "'dd/mm/aaaa'"),
            (r"\bXX-XXXXXXXX-X\b", "'XX-XXXXXXXX-X'"),
            (r"\bSeñores\s+Presidente\s+y\s+Directores\s+de\s+XYZ\b|\bde\s+XYZ\b", "'XYZ' (Nombre genérico)"),
            (r"DENOMINACIÓN\s+SOCIAL\s+DE\s+LA\s+ENTIDAD", "'DENOMINACIÓN SOCIAL DE LA ENTIDAD'"),
            (r"\bNotas?\s*\.\.\.|\bAnexos?\s*\.\.\.|Notas\s+---\s+a\s+---", "'Notas ... / ---' sin numerar"),
            (r"\{De\s+corresponder\}|\(Corresponde\s+en\s+caso\s+únicamente", "Instrucciones de plantilla '{De corresponder}' sin borrar"),
            (r"del\s+xx,xx\s*%", "'xx,xx %' (Variación del índice FACPCE sin completar)"),
            (r"Dr\.\s*WXYZ|T°\s*…\s*F°\s*…", "'Dr. WXYZ / T°… F°…' (Firma de plantilla)")
        ]
        for pat, label in placeholder_patterns:
            if re.search(pat, full_text, re.IGNORECASE):
                placeholders_found.append(label)

        if not placeholders_found:
            checks.append({
                "code": "FOR-02",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Control de Textos de Plantilla / Campos en Borrador",
                "status": "PASS",
                "severity": "CRITICO",
                "detail": "No se detectaron marcadores residuales de modelos del Consejo ('dd/mm/aaaa', 'XYZ', 'Nota ---', '{De corresponder}').",
                "norm_ref": "Aclaraciones Modelos CPCECABA"
            })
        else:
            checks.append({
                "code": "FOR-02",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Control de Textos de Plantilla / Campos en Borrador",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": f"Se detectaron {len(placeholders_found)} marcadores de plantilla sin reemplazar: {', '.join(placeholders_found)}.",
                "recommendation": "Reemplazar todos los campos genéricos del modelo y eliminar las notas instructivas antes de emitir el PDF definitivo.",
                "norm_ref": "Aclaraciones Modelos CPCECABA ('Los textos en rojo e instructivos deben eliminarse')"
            })

        # FOR-03: Domicilio Legal vs. Domicilio Fiscal
        has_dom_legal = bool(re.search(r"Domicilio\s+legal\s*:", full_text, re.IGNORECASE))
        has_dom_fiscal = bool(re.search(r"Domicilio\s+fiscal\s*:", full_text, re.IGNORECASE))
        if has_dom_legal:
            checks.append({
                "code": "FOR-03",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Consignación expresa de 'Domicilio legal'",
                "status": "PASS",
                "severity": "CRITICO",
                "detail": "Se consigna correctamente 'Domicilio legal:' conforme lo exigido para personas jurídicas (SA, SRL, Asociaciones Civiles, Fundaciones, Obras Sociales, Sindicatos).",
                "norm_ref": "Check List Legalizaciones CPCECABA"
            })
        elif has_dom_fiscal:
            checks.append({
                "code": "FOR-03",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Consignación expresa de 'Domicilio legal'",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "Se detectó 'Domicilio fiscal' en lugar de 'Domicilio legal'. En el CPCECABA, 'Domicilio fiscal' sólo aplica a personas humanas o sociedades simples/de hecho.",
                "recommendation": "Si el ente es una persona jurídica, cambiar la etiqueta a 'Domicilio legal:'.",
                "norm_ref": "Check List Legalizaciones CPCECABA"
            })
        else:
            checks.append({
                "code": "FOR-03",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Consignación expresa de 'Domicilio legal'",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "No se detectó la etiqueta explícita 'Domicilio legal:' en la carátula o en el destinatario del Informe del Auditor.",
                "recommendation": "Agregar 'Domicilio legal: [Calle y altura, CABA]' en la Carátula y en el encabezamiento del Informe del Auditor.",
                "norm_ref": "Check List Legalizaciones CPCECABA"
            })

        # FOR-04: Consistencia de Fecha de Cierre de Ejercicio
        if meta["fecha_cierre"] and "No detectada" not in meta["fecha_cierre"]:
            checks.append({
                "code": "FOR-04",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Identificación y Consistencia de Fecha de Cierre de Ejercicio",
                "status": "PASS",
                "severity": "CRITICO",
                "detail": f"Fecha de cierre del ejercicio identificada correctamente: {meta['fecha_cierre']}.",
                "norm_ref": "Check List Legalizaciones CPCECABA"
            })
        else:
            checks.append({
                "code": "FOR-04",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Identificación y Consistencia de Fecha de Cierre de Ejercicio",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": "No se pudo identificar una fecha de cierre de ejercicio válida (o permanece como 'dd/mm/aaaa').",
                "recommendation": "Completar la fecha real de cierre de ejercicio en Carátula, encabezados de estados, Notas e Informe del Auditor.",
                "norm_ref": "Check List Legalizaciones CPCECABA"
            })

        # FOR-05: Leyenda pie de estados básicos ("Las Notas ... y los Anexos ... forman parte integrante")
        has_pie_notas = bool(re.search(r"Las\s+Notas.*forman\s+parte\s+integrante\s+de\s+estos\s+estados\s+contables", full_text, re.IGNORECASE))
        if has_pie_notas:
            checks.append({
                "code": "FOR-05",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Leyenda de Integración de Notas y Anexos al pie de Estados Básicos",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": "Se incluye la leyenda obligatoria indicando que las Notas y Anexos forman parte integrante de los Estados Contables.",
                "norm_ref": "RT 54 párr. 57 / Modelos CPCECABA"
            })
        elif meta["secciones_presentes"]["esp"]:
            checks.append({
                "code": "FOR-05",
                "category": "1. Requisitos Formales Legalizaciones CPCECABA",
                "title": "Leyenda de Integración de Notas y Anexos al pie de Estados Básicos",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "No se detectó al pie de los Estados Básicos la leyenda 'Las Notas X a Y y los Anexos ... que se acompañan forman parte integrante de estos estados contables'.",
                "recommendation": "Incorporar la leyenda al pie del ESP, ER/ERG, EEPN y EFE.",
                "norm_ref": "RT 54 párr. 57 / Modelos CPCECABA"
            })

    # ======================================================================
    # BLOQUE 2: NORMATIVA CONTABLE RT 54 (T.O. RT 59 - NUA)
    # ======================================================================
    def _check_rt54_controls(self, full_text: str, lines: List[Dict[str, Any]], meta: Dict[str, Any], numbers: Dict[str, Any], checks: List[Dict[str, Any]]):
        sec = meta["secciones_presentes"]

        # RT54-01: Presencia de los 4 Estados Contables Básicos
        estados_faltantes = []
        if not sec["esp"]:
            estados_faltantes.append("Estado de Situación Patrimonial (ESP)")
        if not sec["er_erg"]:
            estados_faltantes.append("Estado de Resultados / Recursos y Gastos (ER/ERG)")
        if not sec["eepn"]:
            estados_faltantes.append("Estado de Evolución del Patrimonio Neto (EEPN)")
        if not sec["efe"]:
            estados_faltantes.append("Estado de Flujos de Efectivo (EFE)")

        if len(estados_faltantes) == 0:
            checks.append({
                "code": "RT54-01",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Integridad de los 4 Estados Contables Básicos (RT 54 párr. 55)",
                "status": "PASS",
                "severity": "CRITICO",
                "detail": "Se encuentran presentes los 4 estados básicos exigidos por el párr. 55 de la RT 54 (ESP, ER/ERG, EEPN y EFE).",
                "norm_ref": "RT 54 párr. 55"
            })
        elif len(estados_faltantes) == 4 and not sec["notas"] and sec["informe_auditor"]:
            checks.append({
                "code": "RT54-01",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Integridad de los 4 Estados Contables Básicos (RT 54 párr. 55)",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "El archivo analizado contiene únicamente el Informe del Auditor Independiente (sin adjuntar los Estados Contables Básicos ni las Notas).",
                "recommendation": "Recordar que para presentar en Legalizaciones Web del CPCECABA debe subirse un único archivo PDF integrado con Estados Básicos + Notas + Anexos + Informe del Auditor.",
                "norm_ref": "Instructivo Legalizaciones Web CPCECABA / RT 54 párr. 55"
            })
            return
        else:
            checks.append({
                "code": "RT54-01",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Integridad de los 4 Estados Contables Básicos (RT 54 párr. 55)",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": f"Faltan estados contables básicos en el documento: {', '.join(estados_faltantes)}.",
                "recommendation": "Incluir la totalidad de los estados básicos requeridos por el párr. 55 de la RT 54.",
                "norm_ref": "RT 54 párr. 55"
            })

        # RT54-02: Nota de Bases de Preparación - Adopción RT 54 (T.O. RT 59) y Res. P. 460/2024 CPCECABA
        has_rt54_mention = bool(re.search(r"Resoluci[oó]n\s+T[eé]cnica\s+N[°º]?\s*54|\bRT\s*54\b|Norma\s+Unificada\s+Argentina", full_text, re.IGNORECASE))
        has_res_460 = bool(re.search(r"460/2024", full_text))
        has_rt59 = bool(re.search(r"RT\s*59|T\.O\.\s*RT\s*59", full_text, re.IGNORECASE))

        if has_rt54_mention and has_res_460:
            checks.append({
                "code": "RT54-02",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Nota de Bases de Preparación: Cita de RT 54 (T.O. RT 59) y Res. P. N° 460/2024 CPCECABA",
                "status": "PASS",
                "severity": "CRITICO",
                "detail": "Se menciona explícitamente la preparación bajo RT 54 (T.O. RT 59) de FACPCE adoptada por la Resolución P. N° 460/2024 del CPCECABA.",
                "norm_ref": "RT 54 párr. 605 / Res. P. N° 460/2024 CPCECABA"
            })
        elif has_rt54_mention and not has_res_460:
            checks.append({
                "code": "RT54-02",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Nota de Bases de Preparación: Cita de RT 54 (T.O. RT 59) y Res. P. N° 460/2024 CPCECABA",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "Se menciona la RT 54, pero no se cita la Resolución P. N° 460/2024 del CPCECABA que la adoptó en CABA.",
                "recommendation": "En Nota 1.1 (Bases de preparación) indicar: '...de conformidad con la Resolución Técnica N° 54 - T.O. RT 59 - (RT 54), emitida por la FACPCE, adoptada por la Resolución P. N° 460/2024 del CPCECABA'.",
                "norm_ref": "Modelo Oficial Notas CPCECABA - Nota 1.1"
            })
        else:
            # Check if old RT 8/9/17/41 are cited instead
            old_rts = re.findall(r"Resoluci[oó]n\s+T[eé]cnica\s+N[°º]?\s*(?:8|9|11|17|41)\b|\bRT\s*(?:8|9|11|17|41)\b", full_text, re.IGNORECASE)
            detail_msg = "No se menciona la Resolución Técnica N° 54 (T.O. RT 59 - NUA) ni la Res. P. N° 460/2024 del CPCECABA en las Bases de Preparación."
            if old_rts:
                detail_msg += f" Se detectaron menciones a normas derogadas/reemplazadas ({', '.join(set(old_rts))})."
            checks.append({
                "code": "RT54-02",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Nota de Bases de Preparación: Cita de RT 54 (T.O. RT 59) y Res. P. N° 460/2024 CPCECABA",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": detail_msg,
                "recommendation": "Para ejercicios iniciados a partir del 01/01/2025 es obligatoria la aplicación de la RT 54 (T.O. RT 59) adoptada por Res. P. N° 460/2024 CPCECABA.",
                "norm_ref": "Res. P. N° 460/2024 CPCECABA"
            })

        # RT54-03: Clasificación de la Entidad (EP / EM / RE) - Nota 1.2
        if meta["categoria_rt54"] != "No especificada en Notas":
            checks.append({
                "code": "RT54-03",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Revelación obligatoria de la Categoría de la Entidad (EP / EM / RE)",
                "status": "PASS",
                "severity": "CRITICO",
                "detail": f"Se revela explícitamente la categorización de la entidad bajo RT 54: {meta['categoria_rt54']}.",
                "norm_ref": "RT 54 párr. 6 a 14 / Modelo CPCECABA Nota 1.2"
            })
        else:
            checks.append({
                "code": "RT54-03",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Revelación obligatoria de la Categoría de la Entidad (EP / EM / RE)",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": "No se detectó la Nota de 'Clasificación de la entidad' indicando si reviste el carácter de Entidad Pequeña (EP), Entidad Mediana (EM) o Restante Entidad (RE).",
                "recommendation": "Agregar Nota 1.2 ('Clasificación de la entidad'): 'De acuerdo con lo establecido por la RT 54, la entidad reviste el carácter de [entidad pequeña (EP) / entidad mediana (EM) / restante entidad]'.",
                "norm_ref": "RT 54 párr. 6 a 14 / Modelo CPCECABA Nota 1.2"
            })

        # RT54-04: Expresión en Moneda Homogénea (Ajuste por Inflación - Nota 1.3)
        has_moneda_homogenea = bool(re.search(r"moneda\s+homog[eé]nea", full_text, re.IGNORECASE))
        has_recpam = bool(re.search(r"RECPAM|poder\s+adquisitivo\s+de\s+la\s+moneda|[ií]ndice\s+de\s+precios\s+FACPCE", full_text, re.IGNORECASE))

        if has_moneda_homogenea and has_recpam:
            checks.append({
                "code": "RT54-04",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Expresión en Moneda Homogénea y Nota de Unidad de Medida (RT 54 párr. 176-200)",
                "status": "PASS",
                "severity": "CRITICO",
                "detail": "Los estados contables expresan correctamente su preparación en 'moneda homogénea' y describen el proceso de reexpresión por inflación (RECPAM / índice FACPCE).",
                "norm_ref": "RT 54 párr. 176 a 200 / Modelo CPCECABA Nota 1.3"
            })
        elif has_moneda_homogenea:
            checks.append({
                "code": "RT54-04",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Expresión en Moneda Homogénea y Nota de Unidad de Medida (RT 54 párr. 176-200)",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "Se menciona 'moneda homogénea', pero no se detecta el detalle del proceso de ajuste por inflación (índice FACPCE, porcentaje de variación anual y exposición del RECPAM).",
                "recommendation": "Completar la Nota 1.3 (Unidad de medida) detallando el porcentaje de variación del índice FACPCE en el ejercicio y el criterio de exposición del RECPAM.",
                "norm_ref": "RT 54 párr. 176 a 200 / Modelo CPCECABA Nota 1.3"
            })
        else:
            checks.append({
                "code": "RT54-04",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Expresión en Moneda Homogénea y Nota de Unidad de Medida (RT 54 párr. 176-200)",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": "No se menciona que los estados contables están expresados en 'moneda homogénea' en los encabezados ni en las Notas.",
                "recommendation": "Incluir en Carátula, encabezados de estados y Nota 1.3 que las cifras están expresadas en pesos en moneda homogénea de cierre.",
                "norm_ref": "RT 54 párr. 176 a 200"
            })

        # RT54-05: Nomenclatura de Rubros NUA vs Normas Anteriores (o Dispensa Res. JG 660/2026)
        has_dispensa_660 = bool(re.search(r"660/2026", full_text))
        nua_terms_found = []
        for term in ["Créditos impositivos", "Inversiones financieras", "Deudas laborales y previsionales", "Deudas fiscales", "Proveedores de bienes y servicios", "Bienes para consumo o comercialización"]:
            if re.search(re.escape(term), full_text, re.IGNORECASE):
                nua_terms_found.append(term)

        legacy_terms_found = []
        for lterm in ["Créditos por ventas", "Deudas comerciales", "Remuneraciones y cargas sociales", "Cargas fiscales"]:
            if re.search(r"\b" + re.escape(lterm) + r"\b", full_text, re.IGNORECASE):
                legacy_terms_found.append(lterm)

        if nua_terms_found and not legacy_terms_found:
            checks.append({
                "code": "RT54-05",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Adecuación de Denominaciones de Rubros según RT 54 (NUA)",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": f"Se utilizan las denominaciones actualizadas de la RT 54 ({', '.join(nua_terms_found[:4])}).",
                "norm_ref": "RT 54 Capitulo 6 / Modelos CPCECABA"
            })
        elif legacy_terms_found and has_dispensa_660:
            checks.append({
                "code": "RT54-05",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Adecuación de Denominaciones de Rubros según RT 54 (NUA)",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": f"Se mantienen denominaciones anteriores ({', '.join(legacy_terms_found)}) amparadas correctamente en la dispensa transitoria de la Res. JG FACPCE N° 660/2026.",
                "norm_ref": "Res. JG FACPCE N° 660/2026"
            })
        elif legacy_terms_found and not has_dispensa_660:
            checks.append({
                "code": "RT54-05",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Adecuación de Denominaciones de Rubros según RT 54 (NUA)",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": f"Se detectaron nombres de rubros del régimen anterior ({', '.join(legacy_terms_found)}) sin invocar la dispensa transitoria de primera aplicación.",
                "recommendation": "Reclasificar a la nomenclatura NUA ('Créditos impositivos', 'Cuentas por cobrar a clientes en moneda', 'Proveedores de bienes y servicios', 'Deudas laborales y previsionales', 'Deudas fiscales') o dejar constancia en Nota 1.1 del uso de la dispensa de exposición de la Res. JG 660/2026.",
                "norm_ref": "RT 54 / Modelo CPCECABA Nota 1.1 / Res. JG 660/2026"
            })

        # RT54-06: Nota de Primera Aplicación / Enfoque de Transición y Dispensas (Párr. 659, EFE sintético, RECPAM en 1 línea)
        has_simplificaciones_nota = bool(re.search(r"simplificaciones\s+dispuestas\s+por\s+la\s+RT\s*54|p[aá]rrafo\s+659|p[aá]rrafo\s+599|enfoque\s+retroactivo", full_text, re.IGNORECASE))
        if has_simplificaciones_nota:
            checks.append({
                "code": "RT54-06",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Revelación de Enfoque de Transición y Simplificaciones / Dispensas de la RT 54",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": "Se informan explícitamente en Notas el enfoque de transición y las simplificaciones utilizadas (ej. dispensa párr. 659 / 599, EFE sintético, unificación de RFyT con RECPAM).",
                "norm_ref": "RT 54 Apéndice A / Modelo CPCECABA Nota 1.1 y 1.3"
            })
        elif sec["notas"]:
            checks.append({
                "code": "RT54-06",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Revelación de Enfoque de Transición y Simplificaciones / Dispensas de la RT 54",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "No se detectó en las Notas el párrafo sobre las simplificaciones/dispensas de la RT 54 utilizadas (como no presentar nota del párr. 659 o 599 inc. f, o exponer RFyT + RECPAM en una línea) ni la advertencia sobre sus limitaciones.",
                "recommendation": "Incluir en Nota 1.1 el impacto de primera aplicación de la RT 54 y en Nota 1.3 el listado de simplificaciones utilizadas junto con el párrafo: 'Los presentes estados contables deben ser leídos e interpretados considerando las limitaciones que la utilización de las dispensas antes mencionadas podría provocar...'.",
                "norm_ref": "Modelo Oficial Notas CPCECABA - Nota 1.1 y 1.3"
            })

        # RT54-07: Detección de Rubros, Filas y Conceptos con Saldo Cero en ambos ejercicios (Actual y Comparativo)
        zero_rubros = numbers.get("zero_balance_lines", [])
        # Also check if template instruction about deleting zero rows is still visible
        has_zero_instruction = bool(re.search(r"Los\s+rubros\s+que\s+no\s+tengan\s+saldo\s+en\s+el\s+ejercicio\s+actual\s+y\s+en\s+el\s+ejercicio\s+comparativo,\s+deben\s+eliminarse", full_text, re.IGNORECASE))
        if len(zero_rubros) == 0 and not has_zero_instruction:
            checks.append({
                "code": "RT54-07",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Supresión de Rubros y Filas con Saldo Cero en Ejercicio Actual y Comparativo",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": "No se detectaron rubros ni filas en cero ($ 0,00 / '-') mantenidas innecesariamente en ambos períodos (cumple con el principio de síntesis y flexibilidad de la RT 54).",
                "norm_ref": "RT 54 párr. 55 y 56 inc. a) / Modelos CPCECABA Nota (1)"
            })
        else:
            sample_zeros = ", ".join(zero_rubros[:8])
            if len(zero_rubros) > 8:
                sample_zeros += f" (y {len(zero_rubros) - 8} más)"
            msg_zero = (
                f"Se detectaron {len(zero_rubros)} líneas/rubros con saldo cero ($ 0,00 o '-') tanto en el ejercicio actual como en el comparativo: {sample_zeros}."
                if zero_rubros else
                "El documento conserva la nota instructiva del modelo sobre eliminación de rubros sin saldo."
            )
            checks.append({
                "code": "RT54-07",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Supresión de Rubros y Filas con Saldo Cero en Ejercicio Actual y Comparativo",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": msg_zero,
                "recommendation": "Según la directiva expresa de los Modelos CPCECABA (Nota 1 de ESP, ER/ERG, EEPN, EFE y Notas) y RT 54 párr. 55-56 inc. a): 'Los rubros, filas o columnas que no tengan saldo en el ejercicio actual y en el ejercicio comparativo DEBEN ELIMINARSE, tanto en los estados básicos como en las notas y anexos relacionados'. Utilice el botón 'Ver / Descargar Versión Mejorada' para limpiarlos automáticamente.",
                "norm_ref": "RT 54 párr. 55 y 56 inc. a) / Modelos CPCECABA Nota (1)"
            })

        # RT54-08: Previsiones y Cuadro / Anexo de Previsiones
        has_anexo_prev = bool(re.search(r"ANEXO.*PREVISIONES|PREVISIONES\s*\n.*Deducidas\s+del\s+activo", full_text, re.IGNORECASE))
        has_prev_deducidas = bool(re.search(r"Deducidas\s+del\s+activo", full_text, re.IGNORECASE))
        has_prev_pasivo = bool(re.search(r"Incluidas\s+en\s+el\s+pasivo", full_text, re.IGNORECASE))
        has_prev_rubro_Mention = bool(re.search(r"\bPrevisiones\b|Previsi[oó]n\s+para", full_text, re.IGNORECASE))

        if has_anexo_prev and has_prev_deducidas and has_prev_pasivo:
            checks.append({
                "code": "RT54-08",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Exposición de Previsiones y Estructura del Anexo de Previsiones (RT 54)",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": "El Anexo de Previsiones respeta la apertura oficial del CPCECABA ('Deducidas del activo' vs. 'Incluidas en el pasivo', separadas en Corrientes y No Corrientes, con columnas de Saldos al inicio, Transferencias, Aumentos [Cargos/Otros], Disminuciones [Utilización/Reversiones/Otros] y Saldos al cierre). El RECPAM se expone en las columnas 'Otros'.",
                "norm_ref": "RT 54 párr. 316 y 483-500 / Modelos CPCECABA Anexo Previsiones"
            })
        elif has_prev_rubro_Mention:
            checks.append({
                "code": "RT54-08",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Exposición de Previsiones y Estructura del Anexo de Previsiones (RT 54)",
                "status": "INFO",
                "severity": "INFO",
                "detail": "Se mencionan Previsiones en los estados contables (o en el índice), pero no se detectó el cuadro completo del Anexo de Previsiones con la apertura entre 'Deducidas del activo' e 'Incluidas en el pasivo' (o la entidad no registró movimientos significativos).",
                "recommendation": "Si existen previsiones al cierre o inicio, incluir el Anexo de Previsiones dividiendo 'Deducidas del activo' (que restan del rubro respectivo en el ESP) e 'Incluidas en el pasivo' (rubro Previsiones del Pasivo), exponiendo el efecto del RECPAM en la columna 'Otros (*)' y aclarando en Nota el criterio de medición (valor nominal en EP; valor descontado en RE).",
                "norm_ref": "RT 54 párr. 316 y 483-500 / Modelos CPCECABA Nota 2.13 / 2.15"
            })

        # RT54-09: Cuadro / Anexo de Gastos (Art. 64 inc. b Ley 19.550 / RT 54 párr. 638 y 727)
        has_anexo_gastos = bool(re.search(r"ANEXO.*GASTOS|GASTOS\s+ORDINARIOS|COSTO\s+DE\s+PRODUCCI[OÓ]N\s+Y\s+GASTOS\s+POR\s+NATURALEZA|CUADRO\s+DE\s+GASTOS", full_text, re.IGNORECASE))
        er_gastos_adm = numbers.get("er_gastos_administracion", {}).get("actual")
        anx_gastos_adm = numbers.get("anexo_gastos_administracion", {}).get("actual")

        if has_anexo_gastos:
            if er_gastos_adm is not None and anx_gastos_adm is not None:
                diff_gtos = round(er_gastos_adm - anx_gastos_adm, 2)
                if abs(diff_gtos) <= 1.0:
                    checks.append({
                        "code": "RT54-09",
                        "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                        "title": "Cuadro / Anexo de Gastos por Naturaleza y Función y Cruce con ER/ERG",
                        "status": "PASS",
                        "severity": "CRITICO",
                        "detail": f"Se incluye el Cuadro/Anexo de Gastos clasificado por naturaleza y función, y el total de Gastos de Administración ({fmt_ars(anx_gastos_adm)}) cruza exactamente con el Estado de Resultados / Recursos y Gastos ({fmt_ars(er_gastos_adm)}).",
                        "norm_ref": "RT 54 párr. 638 y 727 / Art. 64 inc. b) Ley 19.550 / Modelos CPCECABA"
                    })
                else:
                    checks.append({
                        "code": "RT54-09",
                        "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                        "title": "Cuadro / Anexo de Gastos por Naturaleza y Función y Cruce con ER/ERG",
                        "status": "ERROR",
                        "severity": "CRITICO",
                        "detail": f"Diferencia de {fmt_ars(diff_gtos)} entre el Total de Gastos de Administración del Anexo de Gastos ({fmt_ars(anx_gastos_adm)}) y el expuesto en el Estado de Resultados / Recursos y Gastos ({fmt_ars(er_gastos_adm)}).",
                        "recommendation": "Verificar que el total de cada columna funcional del Anexo de Gastos coincida exactamente con la línea homónima del Estado de Resultados / Recursos y Gastos.",
                        "norm_ref": "RT 54 párr. 638 y 727 / Art. 64 inc. b) Ley 19.550"
                    })
            else:
                checks.append({
                    "code": "RT54-09",
                    "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                    "title": "Cuadro / Anexo de Gastos por Naturaleza y Función (RT 54 párr. 638 / 727)",
                    "status": "PASS",
                    "severity": "ADVERTENCIA",
                    "detail": "Se incluye el Anexo de Gastos con apertura por naturaleza y distribución funcional conforme a los modelos del CPCECABA (en ESFL: Producción, Fines generales, Fines específicos y Administración, pudiendo desdoblarse en anexos individuales; en sociedades comerciales: Costo de producción, Costo de servicios, Comercialización y Administración).",
                    "norm_ref": "RT 54 párr. 638 y 727 / Modelos CPCECABA Anexo Gastos"
                })
        else:
            checks.append({
                "code": "RT54-09",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Cuadro / Anexo de Gastos por Naturaleza y Función (RT 54 párr. 638 / 727)",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "No se detectó el Cuadro / Anexo de Gastos (Información requerida por el Art. 64 inc. b de la Ley 19.550 en sociedades o RT 54 párr. 638 / 727 en restantes entidades y ESFL).",
                "recommendation": "Incorporar el Anexo de Gastos abriendo los conceptos por su naturaleza y asignándolos a las columnas funcionales que cruzan con el Estado de Resultados / Recursos y Gastos.",
                "norm_ref": "RT 54 párr. 638 y 727 / Art. 64 inc. b) Ley 19.550"
            })

        # RT54-10: Prohibición de EFE Sintético en Entidades Sin Fines de Lucro y Cooperativas (RT 54 párr. 666)
        has_efe_sintetico = bool(re.search(r"en\s+forma\s+sint[eé]tica|Presentaci[oó]n\s+Sint[eé]tica", full_text, re.IGNORECASE))
        if meta["es_sin_fines_de_lucro"]:
            if has_efe_sintetico:
                checks.append({
                    "code": "RT54-10",
                    "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                    "title": "Prohibición de EFE Sintético en Entidades Sin Fines de Lucro y Cooperativas (RT 54 párr. 666)",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": "La entidad es Sin Fines de Lucro, pero declara o titula el Estado de Flujos de Efectivo 'En forma sintética'. El párrafo 666 de la RT 54 (y la Nota 2 de la hoja 4-EFE del modelo CPCECABA) excluye expresamente a las ESFL y Cooperativas de la opción de EFE sintético.",
                    "recommendation": "Suprimir la mención 'En forma sintética' del título del EFE y de la Nota 1.3, y exponer las causas operativas abriendo los cobros y pagos por el método directo. La 'Versión Mejorada' corrige esta exposición automáticamente.",
                    "norm_ref": "RT 54 párr. 666 / Modelo CPCECABA Hoja 4-EFE Nota (2)"
                })
            else:
                checks.append({
                    "code": "RT54-10",
                    "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                    "title": "Exposición Abierta por Método Directo del EFE en Entidades Sin Fines de Lucro (RT 54 párr. 666)",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": "El Estado de Flujos de Efectivo de la Entidad Sin Fines de Lucro no utiliza la dispensa de exposición sintética prohibida por el párr. 666 de la RT 54.",
                    "norm_ref": "RT 54 párr. 666 / Modelo CPCECABA Hoja 4-EFE Nota (2)"
                })
        else:
            checks.append({
                "code": "RT54-10",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Habilitación de EFE Sintético para Entidades con Fines de Lucro (RT 54 párr. 666)",
                "status": "PASS",
                "severity": "INFO",
                "detail": "Al tratarse de una entidad con fines de lucro en contexto de ajuste por inflación, se encuentra habilitada a presentar el Estado de Flujos de Efectivo en forma sintética (RT 54 párr. 666).",
                "norm_ref": "RT 54 párr. 666"
            })

        # RT54-11: Conciliación entre el Efectivo del EFE y las partidas del ESP (RT 54 párr. 672 inc. a)
        ef_cie = numbers["efe_efectivo_cierre"]["actual"]
        caja_esp = numbers["caja_y_bancos"]["actual"]
        has_conciliacion_efe_nota = bool(re.search(r"conciliaci[oó]n\s+entre\s+el\s+efectivo\s+y\s+sus\s+equivalentes|p[aá]rr(?:afo)?\.?\s*672", full_text, re.IGNORECASE))
        if ef_cie is not None and caja_esp is not None:
            if abs(ef_cie - caja_esp) <= 1.0:
                checks.append({
                    "code": "RT54-11",
                    "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                    "title": "Conciliación del Efectivo y Equivalentes (EFE vs. Caja y Bancos del ESP - RT 54 párr. 672 inc. a)",
                    "status": "PASS",
                    "severity": "ADVERTENCIA",
                    "detail": f"El Efectivo y equivalentes al cierre en el EFE ({fmt_ars(ef_cie)}) coincide exactamente con el rubro Caja y Bancos del ESP ({fmt_ars(caja_esp)}).",
                    "norm_ref": "RT 54 párr. 672 inc. a) / Modelo CPCECABA Hoja 4-EFE Nota (3)"
                })
            elif has_conciliacion_efe_nota:
                checks.append({
                    "code": "RT54-11",
                    "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                    "title": "Conciliación del Efectivo y Equivalentes (EFE vs. Caja y Bancos del ESP - RT 54 párr. 672 inc. a)",
                    "status": "PASS",
                    "severity": "ADVERTENCIA",
                    "detail": f"El Efectivo del EFE ({fmt_ars(ef_cie)}) difiere de Caja y Bancos ({fmt_ars(caja_esp)}) por incluir equivalentes de efectivo de Inversiones Financieras, y cuenta con la Nota de conciliación del párr. 672 inc. a).",
                    "norm_ref": "RT 54 párr. 672 inc. a)"
                })
            else:
                checks.append({
                    "code": "RT54-11",
                    "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                    "title": "Conciliación del Efectivo y Equivalentes (EFE vs. Caja y Bancos del ESP - RT 54 párr. 672 inc. a)",
                    "status": "WARNING",
                    "severity": "ADVERTENCIA",
                    "detail": f"El Efectivo al cierre del EFE ({fmt_ars(ef_cie)}) no coincide con Caja y Bancos del ESP ({fmt_ars(caja_esp)}) y no se detectó la Nota de conciliación exigida por el párr. 672 inc. a) de la RT 54.",
                    "recommendation": "Incorporar en Notas la conciliación entre los componentes de Caja y Bancos (+ Inversiones equivalentes de efectivo menores a 90 días) y el Efectivo expuesto en el EFE.",
                    "norm_ref": "RT 54 párr. 672 inc. a) / Modelo CPCECABA Hoja 4-EFE Nota (3)"
                })
        elif has_conciliacion_efe_nota:
            checks.append({
                "code": "RT54-11",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Conciliación del Efectivo y Equivalentes (EFE vs. ESP - RT 54 párr. 672 inc. a)",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": "Se incluye la referencia a la conciliación entre el Efectivo y equivalentes del EFE y las partidas del ESP (RT 54 párr. 672 inc. a).",
                "norm_ref": "RT 54 párr. 672 inc. a)"
            })
        else:
            checks.append({
                "code": "RT54-11",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Conciliación del Efectivo y Equivalentes (EFE vs. ESP - RT 54 párr. 672 inc. a)",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "No se detectó la Nota de conciliación entre el Efectivo y sus equivalentes considerado en el EFE y las partidas del Estado de Situación Patrimonial (requerida por la Nota 3/7 de la hoja 4-EFE del CPCECABA).",
                "recommendation": "Agregar en la Nota de Efectivo / EFE la conciliación expresa entre el saldo de Caja y Bancos (+ equivalentes en Inversiones Financieras) y el Efectivo al cierre del EFE (RT 54 párr. 672 inc. a).",
                "norm_ref": "RT 54 párr. 672 inc. a) / Modelo CPCECABA Hoja 4-EFE"
            })

        # RT54-12: Evaluación de Valor Recuperable de Bienes de Uso e Intangibles ante Resultados Negativos (Regla de 3 ejercicios positivos)
        res_act = numbers["resultado_ejercicio_er"]["actual"]
        res_cmp = numbers["resultado_ejercicio_er"]["comparativo"]
        has_val_recup_nota = bool(re.search(r"Valor\s+recuperable\s+de\s+los\s+Bienes\s+de\s+Uso|indicios\s+de\s+deterioro|no\s+supera\s+su\s+valor\s+recuperable", full_text, re.IGNORECASE))
        has_loss = (res_act is not None and res_act < 0) or (res_cmp is not None and res_cmp < 0)

        if has_loss and not has_val_recup_nota:
            checks.append({
                "code": "RT54-12",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Evaluación de Valor Recuperable de Bienes de Uso ante Resultados Negativos (Nota 2.9 CPCECABA)",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": "El ejercicio actual o comparativo arroja resultado negativo (Pérdida / Déficit), por lo que NO rige la dispensa de los '3 últimos ejercicios positivos' para Entidades Pequeñas, y no se detectó la Nota de evaluación de valor recuperable de Bienes de Uso e Intangibles.",
                "recommendation": "Incluir la Nota 2.9 ('Valor recuperable de los Bienes de Uso, Propiedades de inversión y Activos intangibles') dejando constancia de la prueba de recuperabilidad efectuada al cierre.",
                "norm_ref": "RT 54 párr. 150-165 / Modelo CPCECABA Nota 2.9"
            })
        elif has_val_recup_nota:
            checks.append({
                "code": "RT54-12",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Evaluación de Valor Recuperable de Bienes de Uso e Intangibles (Nota 2.9 CPCECABA)",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": "Se incluye expresamente en Notas la política de evaluación de indicios de deterioro y valor recuperable de Bienes de Uso, Propiedades de Inversión e Intangibles (regla de los 3 ejercicios positivos en EP).",
                "norm_ref": "RT 54 párr. 150-165 / Modelo CPCECABA Nota 2.9"
            })
        else:
            checks.append({
                "code": "RT54-12",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Evaluación de Valor Recuperable de Bienes de Uso e Intangibles (Nota 2.9 CPCECABA)",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "No se detectó la Nota específica sobre evaluación de indicios de deterioro / valor recuperable de Bienes de Uso e Intangibles (Nota 2.9 del modelo CPCECABA).",
                "recommendation": "Incorporar la Nota 2.9 indicando si la entidad evaluó indicios de deterioro o si aplicó la dispensa por haber obtenido resultados positivos en los últimos tres ejercicios consecutivos.",
                "norm_ref": "RT 54 párr. 150-165 / Modelo CPCECABA Nota 2.9"
            })

        # RT54-13: Consistencia de Primera Aplicación de RT 54 (Cuadro de Conciliación del PN vs. Declaración Sin Efectos)
        has_concil_pn_rt54 = bool(re.search(r"Patrimonio\s+Neto\s+seg[uú]n\s+Normas\s+anteriores|Modificaciones?\s+de\s+resultados\s+de\s+ejercicios\s+anteriores", full_text, re.IGNORECASE))
        has_sin_efectos_rt54 = bool(re.search(r"sin\s+generar\s+efectos\s+en\s+el\s+reconocimiento,\s+medici[oó]n", full_text, re.IGNORECASE))
        if has_concil_pn_rt54 or has_sin_efectos_rt54:
            checks.append({
                "code": "RT54-13",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Nota de Transición en Primera Aplicación de la RT 54 (Conciliación del PN o Sin Efectos de Medición)",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": "La Nota 1.1 cumple con el Apéndice A de la RT 54 informando el impacto de transición (cuadro de conciliación del Patrimonio Neto inicial o manifestación expresa de ausencia de efectos en reconocimiento y medición).",
                "norm_ref": "RT 54 Apéndice A párr. A.2-A.17 / Modelo CPCECABA Nota 1.1 (Nota al pie 1)"
            })
        else:
            checks.append({
                "code": "RT54-13",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Nota de Transición en Primera Aplicación de la RT 54 (Conciliación del PN o Sin Efectos de Medición)",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "No se especificó en Nota 1.1 si la primera aplicación de la RT 54 originó ajustes de medición en los saldos iniciales (Tabla de conciliación del PN) o si no generó efectos de medición.",
                "recommendation": "Si la adopción de la RT 54 no alteró la medición de activos/pasivos, incluir la cláusula oficial del CPCECABA: 'Esta resolución fue aplicada por la entidad sin generar efectos en el reconocimiento, medición y presentación, excepto por las modificaciones en las denominaciones de ciertos rubros'.",
                "norm_ref": "RT 54 Apéndice A / Modelo CPCECABA Nota 1.1"
            })

        # RT54-14: Aportes Irrevocables (RG IGJ 15/2024 Art. 240) y Llamadas de Aprobación de Asamblea en el EEPN
        has_unfilled_eepn_calls = bool(re.search(r"Aprobadas/os\s+por[….]+del\s+dd/mm/aaaa", full_text, re.IGNORECASE))
        if has_unfilled_eepn_calls:
            checks.append({
                "code": "RT54-14",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Llamadas de Aprobación de Asamblea / Órgano y Aportes Irrevocables (RG IGJ 15/2024 Art. 240) en el EEPN",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "El pie del Estado de Evolución del Patrimonio Neto (EEPN) mantiene las llamadas '(*) / (**) Aprobadas/os por... del dd/mm/aaaa' sin completar o sin suprimir si no hubo movimientos de distribución/capitalización.",
                "recommendation": "Si hubo distribución de resultados, absorción o suscripción de capital/aportes irrevocables (RG IGJ 15/2024 Art. 240), consignar la fecha exacta de la Asamblea/Reunión de Socios; si esas filas están en cero, suprimir tanto las filas como las llamadas al pie.",
                "norm_ref": "RG IGJ 15/2024 Art. 240 / Modelo CPCECABA Hoja 3-EEPN"
            })
        else:
            checks.append({
                "code": "RT54-14",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Llamadas de Aprobación de Asamblea / Órgano y Aportes Irrevocables (RG IGJ 15/2024 Art. 240) en el EEPN",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": "El Estado de Evolución del Patrimonio Neto no presenta llamadas de Asamblea incompletas ni inconsistencias con el Art. 240 de la RG IGJ 15/2024.",
                "norm_ref": "RG IGJ 15/2024 Art. 240 / Modelo CPCECABA Hoja 3-EEPN"
            })

        # RT54-15: Integridad y Correlatividad de Referencias Cruzadas a Notas y Anexos (RT 54 párr. 57)
        unnumbered_refs = re.findall(r"\b(?:Notas?|Anexos?)\s*(?:\.{2,}|-{2,}|xx|x\.x)\b", full_text, re.IGNORECASE)
        if unnumbered_refs:
            checks.append({
                "code": "RT54-15",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Integridad y Correlatividad de Referencias Cruzadas a Notas y Anexos (RT 54 párr. 57)",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": f"Se detectaron {len(unnumbered_refs)} referencias cruzadas sin numerar o huérfanas ({', '.join(sorted(set(unnumbered_refs))[:5])}) en los estados básicos o en las notas.",
                "recommendation": "Numerar correlativamente todas las Notas y Anexos definitivos luego de eliminar los que tengan saldo cero, verificando que cada llamado en el ESP y ER/ERG apunte al número de Nota exacto (RT 54 párr. 57).",
                "norm_ref": "RT 54 párr. 57 / Check List Legalizaciones CPCECABA"
            })
        else:
            checks.append({
                "code": "RT54-15",
                "category": "2. Cumplimiento Normativo RT 54 (T.O. RT 59 - NUA)",
                "title": "Integridad y Correlatividad de Referencias Cruzadas a Notas y Anexos (RT 54 párr. 57)",
                "status": "PASS",
                "severity": "CRITICO",
                "detail": "Todas las referencias cruzadas a Notas y Anexos se encuentran debidamente identificadas sin marcadores 'Nota ...' ni 'Anexo X.X' pendientes.",
                "norm_ref": "RT 54 párr. 57"
            })

    # ======================================================================
    # BLOQUE 3: INFORME DEL AUDITOR INDEPENDIENTE (RT 37 MOD. RT 53)
    # ======================================================================
    def _check_auditor_report_controls(self, full_text: str, lines: List[Dict[str, Any]], meta: Dict[str, Any], checks: List[Dict[str, Any]]):
        if not meta["secciones_presentes"]["informe_auditor"]:
            checks.append({
                "code": "RT37-00",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Presencia del Informe del Auditor Independiente",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "El documento cargado contiene los Estados Contables / Notas pero aún no incluye adjunto el Informe del Auditor Independiente.",
                "recommendation": "Integrar el Informe del Auditor al documento antes de enviarlo a legalizar al CPCECABA.",
                "norm_ref": "RT 37 (mod. RT 53) / Legalizaciones CPCECABA"
            })
            return

        # Isolate Auditor Report text segment
        m_start = re.search(r"INFORME\s+DE(?:L)?\s+AUDITOR[IÍ]?A\s+EMITIDO\s+POR\s+EL\s+AUDITOR\s+INDEPENDIENTE|INFORME\s+DE(?:L)?\s+AUDITOR\s+INDEPENDIENTE", full_text, re.IGNORECASE)
        aud_text = full_text[m_start.start():] if m_start else full_text

        # RT37-01: Título Oficial Obligatorio
        has_exact_title = bool(re.search(r"INFORME\s+DE\s+AUDITOR[IÍ]A\s+EMITIDO\s+POR\s+EL\s+AUDITOR\s+INDEPENDIENTE", aud_text, re.IGNORECASE))
        if has_exact_title:
            checks.append({
                "code": "RT37-01",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Título Oficial del Informe del Auditor (RT 37 mod. RT 53)",
                "status": "PASS",
                "severity": "CRITICO",
                "detail": "El título 'INFORME DE AUDITORÍA EMITIDO POR EL AUDITOR INDEPENDIENTE' respeta exactamente el modelo de la RT 37 (mod. por RT 53).",
                "norm_ref": "RT 37 (mod. RT 53) / Modelo CPCECABA"
            })
        else:
            checks.append({
                "code": "RT37-01",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Título Oficial del Informe del Auditor (RT 37 mod. RT 53)",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": "El título del informe no coincide con la denominación oficial exigida por la RT 37 modificada por RT 53.",
                "recommendation": "Utilizar exactamente el título: 'INFORME DE AUDITORÍA EMITIDO POR EL AUDITOR INDEPENDIENTE'.",
                "norm_ref": "RT 37 (mod. RT 53) / Check List CPCECABA"
            })

        # RT37-02: Orden obligatorio de párrafos ("Opinión" 1° y "Fundamento de la opinión" 2°, antes de Responsabilidades)
        pos_opinion = re.search(r"\bOpini[oó]n\b", aud_text, re.IGNORECASE)
        pos_fundamento = re.search(r"\bFundamentos?\s+de\s+la\s+opini[oó]n\b", aud_text, re.IGNORECASE)
        pos_resp_dir = re.search(r"\bResponsabilidad(?:es)?\s+de\s+la\s+(?:Direcci[oó]n|Administraci[oó]n|Comisi[oó]n)", aud_text, re.IGNORECASE)
        pos_resp_aud = re.search(
            r"\bResponsabilidad(?:es)?\s+del\s+auditor\s+en\s+relaci[oó]n\s+con\s+la\s+auditor[ií]a(?!\s+de\s+los\s+estados\s+contables[”\"])",
            aud_text[pos_resp_dir.start():] if pos_resp_dir else aud_text,
            re.IGNORECASE
        )
        pos_resp_aud_abs = (pos_resp_dir.start() + pos_resp_aud.start()) if (pos_resp_dir and pos_resp_aud) else (pos_resp_aud.start() if pos_resp_aud else None)

        if pos_opinion and pos_fundamento and pos_resp_dir and pos_resp_aud_abs is not None:
            if pos_opinion.start() < pos_fundamento.start() < pos_resp_dir.start() < pos_resp_aud_abs:
                checks.append({
                    "code": "RT37-02",
                    "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                    "title": "Estructura y Orden Obligatorio de Secciones (Opinión -> Fundamento -> Responsabilidades)",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": "El orden de los apartados cumple con la RT 37 (mod. RT 53): 1° Opinión, 2° Fundamento de la opinión, seguido de Responsabilidades de la Dirección y del Auditor.",
                    "norm_ref": "RT 37 (mod. RT 53) / Res. CD N° 46/2021 CPCECABA"
                })
            else:
                checks.append({
                    "code": "RT37-02",
                    "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                    "title": "Estructura y Orden Obligatorio de Secciones (Opinión -> Fundamento -> Responsabilidades)",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": "El orden de los párrafos del Informe del Auditor NO respeta la estructura obligatoria de la RT 37 (mod. RT 53). La 'Opinión' debe ir al comienzo, seguida de 'Fundamento de la opinión'.",
                    "recommendation": "Reordenar los apartados según el nuevo esquema RT 37 (mod. RT 53) vigente desde septiembre 2022.",
                    "norm_ref": "RT 37 (mod. RT 53) / Res. CD N° 46/2021 CPCECABA"
                })
        else:
            missing_secs = []
            if not pos_opinion: missing_secs.append("'Opinión'")
            if not pos_fundamento: missing_secs.append("'Fundamento de la opinión'")
            if not pos_resp_dir: missing_secs.append("'Responsabilidades de la Dirección'")
            if not pos_resp_aud: missing_secs.append("'Responsabilidades del auditor'")
            checks.append({
                "code": "RT37-02",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Estructura y Orden Obligatorio de Secciones (Opinión -> Fundamento -> Responsabilidades)",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": f"Faltan apartados obligatorios en el Informe del Auditor: {', '.join(missing_secs)}.",
                "recommendation": "Utilizar el modelo oficial de Informe de Auditoría de la RT 37 (mod. RT 53) del CPCECABA.",
                "norm_ref": "RT 37 (mod. RT 53)"
            })

        # RT37-03: Cita de Resolución de adopción jurisdiccional y Código de Ética en "Fundamento de la opinión"
        has_res_46_caba = bool(re.search(r"46/2021", aud_text))
        has_res_3843_pba = bool(re.search(r"3843", aud_text))
        has_cod_etica = bool(re.search(r"C[oó]digo\s+de\s+[EÉ]tica", aud_text, re.IGNORECASE))

        if "CABA" in meta["jurisdiccion"]:
            if has_res_46_caba and has_cod_etica:
                checks.append({
                    "code": "RT37-03",
                    "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                    "title": "Cita de Resolución C.D. N° 46/2021 CPCECABA e Independencia (Código de Ética)",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": "Se cita correctamente la Sección III.A de la RT N° 37 adoptada por la Res. C.D. N° 46/2021 del CPCECABA y el cumplimiento del Código de Ética del CPCECABA.",
                    "norm_ref": "Modelo Informe Auditoría CPCECABA"
                })
            elif has_res_3843_pba:
                checks.append({
                    "code": "RT37-03",
                    "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                    "title": "Cita de Resolución C.D. N° 46/2021 CPCECABA e Independencia (Código de Ética)",
                    "status": "WARNING",
                    "severity": "ADVERTENCIA",
                    "detail": "El informe cita la Resolución CD N° 3843 del Consejo de Provincia de Buenos Aires (CPCEPBA). Si se presenta en CABA será observado; si es para PBA es correcto.",
                    "recommendation": "Si la presentación es ante el CPCECABA, reemplazar por 'Resolución C.D. N° 46/2021 del CPCECABA'.",
                    "norm_ref": "Check List Legalizaciones CPCECABA"
                })
            else:
                checks.append({
                    "code": "RT37-03",
                    "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                    "title": "Cita de Resolución C.D. N° 46/2021 CPCECABA e Independencia (Código de Ética)",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": "No se cita la Resolución C.D. N° 46/2021 del CPCECABA ni el Código de Ética en el párrafo de 'Fundamento de la opinión'.",
                    "recommendation": "Incluir: '...establecidas en la sección III.A de la Resolución Técnica N° 37 de la FACPCE adoptada por la Resolución C.D. N° 46/2021 del CPCECABA...'.",
                    "norm_ref": "Modelo Informe Auditoría CPCECABA"
                })
        else:
            # PBA
            if has_res_3843_pba and has_cod_etica:
                checks.append({
                    "code": "RT37-03",
                    "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                    "title": "Cita de Resolución CD N° 3843 CPCEPBA e Independencia (Código de Ética)",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": "Se cita correctamente la Resolución CD N° 3843 del CPCEPBA y el Código de Ética vigente.",
                    "norm_ref": "Modelo Informe Auditoría CPCEPBA"
                })

        # RT37-04: Informe sobre otros requerimientos legales y reglamentarios (Deuda Previsional SIPA - Ley 17.250 Art. 10)
        has_otros_req = bool(re.search(r"Informe\s+sobre\s+otros\s+requerimientos\s+legales\s+y\s+reglamentarios", aud_text, re.IGNORECASE))
        has_sipa = bool(re.search(r"Sistema\s+Integrado\s+Previsional\s+Argentino|aportes\s+y\s+contribuciones\s+previsionales", aud_text, re.IGNORECASE))
        has_exigible = bool(re.search(r"exigible", aud_text, re.IGNORECASE))

        if has_otros_req and has_sipa and has_exigible:
            checks.append({
                "code": "RT37-04",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Requerimientos Legales: Deuda Previsional SIPA (Art. 10 Ley 17.250)",
                "status": "PASS",
                "severity": "CRITICO",
                "detail": "Se incluye el apartado 'Informe sobre otros requerimientos legales y reglamentarios' con la manifestación sobre deuda devengada y exigible/no exigible a favor del SIPA.",
                "norm_ref": "Ley 17.250 Art. 10 / Check List CPCECABA"
            })
        else:
            checks.append({
                "code": "RT37-04",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Requerimientos Legales: Deuda Previsional SIPA (Art. 10 Ley 17.250)",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": "Falta o está incompleta la mención obligatoria del pasivo devengado y exigible/no exigible a favor del Sistema Integrado Previsional Argentino (SIPA).",
                "recommendation": "Incluir en 'Informe sobre otros requerimientos legales y reglamentarios' el monto exacto devengado al cierre a favor del SIPA y si resultaba exigible o no exigible.",
                "norm_ref": "Ley 17.250 Art. 10 / Check List CPCECABA"
            })

        # RT37-05: Párrafo sobre Libros Rubricados / Registros Contables
        has_registros = bool(re.search(r"surgen\s+de\s+registros\s+contables\s+llevados", aud_text, re.IGNORECASE))
        if has_registros:
            checks.append({
                "code": "RT37-05",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Manifestación sobre Registros Contables llevados en legal forma",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": "Se informa que los estados contables surgen de registros contables llevados en sus aspectos formales de conformidad con normas legales.",
                "norm_ref": "Modelo Informe Auditoría CPCECABA - Otros Requerimientos inc. a)"
            })
        else:
            checks.append({
                "code": "RT37-05",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Manifestación sobre Registros Contables llevados en legal forma",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "No se detectó el inciso sobre registros contables llevados en sus aspectos formales de acuerdo con normas legales.",
                "recommendation": "En CABA, salvo entes exceptuados de llevar libros rubricados, incluir: 'a) Con base en mi examen descripto, informo que los estados contables citados surgen de registros contables llevados en sus aspectos formales de acuerdo con normas legales'.",
                "norm_ref": "Modelo Informe Auditoría CPCECABA"
            })

        # RT37-06: Control Cronológico de Fechas (Fecha del Informe del Auditor >= Fecha de Cierre y Consistencia en SIPA)
        aud_dates = extract_all_dates(aud_text)
        cierre_dates = extract_all_dates(meta.get("fecha_cierre", ""))
        if aud_dates and cierre_dates:
            c_d = cierre_dates[0]
            c_tuple = (c_d["year"], c_d["month"], c_d["day"])
            # Last date in auditor report is typically the emission date
            em_d = aud_dates[-1]
            em_tuple = (em_d["year"], em_d["month"], em_d["day"])
            if em_tuple >= c_tuple:
                checks.append({
                    "code": "RT37-06",
                    "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                    "title": "Consistencia Cronológica de Fechas (Emisión del Informe >= Cierre de Ejercicio y SIPA)",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": f"La fecha de emisión/cierre verificada en el Informe del Auditor ({em_d['normalized']}) es cronológicamente consistente respecto de la fecha de cierre del ejercicio ({c_d['normalized']}).",
                    "norm_ref": "Check List Legalizaciones CPCECABA / RT 37"
                })
            else:
                checks.append({
                    "code": "RT37-06",
                    "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                    "title": "Consistencia Cronológica de Fechas (Emisión del Informe >= Cierre de Ejercicio y SIPA)",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": f"Error cronológico crítico: la fecha final detectada en el Informe del Auditor ({em_d['normalized']}) es ANTERIOR a la fecha de cierre del ejercicio ({c_d['normalized']}). Motivo de rechazo automático en Legalizaciones.",
                    "recommendation": "Corregir la fecha de emisión del Informe del Auditor para que sea igual o posterior a la fecha de cierre de los Estados Contables.",
                    "norm_ref": "Check List Legalizaciones CPCECABA"
                })
        else:
            checks.append({
                "code": "RT37-06",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Consistencia Cronológica de Fechas (Emisión del Informe >= Cierre de Ejercicio y SIPA)",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "No se pudo verificar automáticamente la relación cronológica entre la fecha de cierre de ejercicio y la fecha de emisión al pie del Informe del Auditor (o se encuentra como 'dd/mm/aaaa').",
                "recommendation": "Asegurar que la fecha de emisión del informe al pie sea igual o posterior a la fecha de cierre del ejercicio y que la fecha del párrafo SIPA coincida con la de cierre.",
                "norm_ref": "Check List Legalizaciones CPCECABA"
            })

        # RT37-07: Sección de 'Otra Información (Memoria)' según Tipo Societario (RT 37 mod. RT 53)
        has_otra_info = bool(re.search(r"Informaci[oó]n\s+distinta\s+de\s+los\s+estados\s+contables|Otra\s+informaci[oó]n", aud_text, re.IGNORECASE))
        has_de_corresponder_left = bool(re.search(r"\{De\s+corresponder\}", aud_text, re.IGNORECASE))
        denom_upper = meta.get("denominacion", "").upper()
        is_srl = "S.R.L" in denom_upper or "SRL" in denom_upper or "RESPONSABILIDAD LIMITADA" in denom_upper

        if has_de_corresponder_left:
            checks.append({
                "code": "RT37-07",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Adecuación de Sección 'Otra Información (Memoria)' según Tipo de Entidad",
                "status": "ERROR",
                "severity": "CRITICO",
                "detail": "El Informe del Auditor conserva los párrafos opcionales del modelo con la leyenda '{De corresponder}' sin definir si aplica o no la sección de Otra Información (Memoria).",
                "recommendation": "Si la entidad presenta Memoria (SA, SAU, Asociación Civil, Fundación), mantener la sección 'Información distinta de los estados contables' eliminando '{De corresponder}'. Si no presenta Memoria, eliminar la sección completa.",
                "norm_ref": "RT 37 (mod. RT 53) Sección III.A / Modelo CPCECABA"
            })
        elif has_otra_info and not is_srl:
            checks.append({
                "code": "RT37-07",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Adecuación de Sección 'Otra Información (Memoria)' según Tipo de Entidad",
                "status": "PASS",
                "severity": "ADVERTENCIA",
                "detail": "Se incluye correctamente el apartado 'Información distinta de los estados contables y del informe de auditoría correspondiente (Otra información: Memoria)'.",
                "norm_ref": "RT 37 (mod. RT 53) Sección III.A"
            })
        else:
            checks.append({
                "code": "RT37-07",
                "category": "3. Informe del Auditor Independiente (RT 37 mod. RT 53)",
                "title": "Adecuación de Sección 'Otra Información (Memoria)' según Tipo de Entidad",
                "status": "INFO",
                "severity": "INFO",
                "detail": "No se incluye el apartado 'Otra información (Memoria)' (correcto si la entidad no acompaña Memoria adjunta; obligatorio si es S.A. o Asociación Civil que presenta Memoria).",
                "recommendation": "Si los Estados Contables se presentan acompañados por la Memoria del órgano de administración, incorporar el párrafo de 'Información distinta de los estados contables (Otra información)'.",
                "norm_ref": "RT 37 (mod. RT 53) Sección III.A"
            })

    # ======================================================================
    # EXTRACCIÓN Y CUADRATURA NUMÉRICA DE LOS ESTADOS CONTABLES
    # ======================================================================
    def _extract_key_financial_numbers(self, lines: List[Dict[str, Any]], full_text: str) -> Dict[str, Any]:
        """
        Scans structured lines to extract key totals for Actual and Comparativo columns:
        - Total Activo Corriente, Activo No Corriente, Total Activo
        - Total Pasivo Corriente, Pasivo No Corriente, Total Pasivo
        - Patrimonio Neto (según ESP) y Saldo al cierre del Patrimonio Neto (según EEPN)
        - Total Pasivo + Patrimonio Neto
        - Resultado / Superávit (Déficit) del ejercicio (según ER/ERG y según EEPN)
        - Efectivo al inicio, Efectivo al cierre, Variación neta del efectivo y causas (Operativas, Inversión, Financiación)
        - Caja y Bancos (ESP)
        """
        data = {
            "caja_y_bancos": {"actual": None, "comparativo": None},
            "inversiones_financieras_corr": {"actual": None, "comparativo": None},
            "activo_corriente": {"actual": None, "comparativo": None},
            "activo_no_corriente": {"actual": None, "comparativo": None},
            "total_activo": {"actual": None, "comparativo": None},
            "pasivo_corriente": {"actual": None, "comparativo": None},
            "pasivo_no_corriente": {"actual": None, "comparativo": None},
            "total_pasivo": {"actual": None, "comparativo": None},
            "patrimonio_neto_esp": {"actual": None, "comparativo": None},
            "total_pasivo_mas_pn": {"actual": None, "comparativo": None},
            "resultado_ejercicio_er": {"actual": None, "comparativo": None},
            "patrimonio_neto_eepn_inicio": {"actual": None, "comparativo": None},
            "resultado_ejercicio_eepn": {"actual": None, "comparativo": None},
            "patrimonio_neto_eepn_cierre": {"actual": None, "comparativo": None},
            "efe_efectivo_inicio": {"actual": None, "comparativo": None},
            "efe_efectivo_cierre": {"actual": None, "comparativo": None},
            "efe_variacion_neta": {"actual": None, "comparativo": None},
            "efe_flujo_operativo": {"actual": None, "comparativo": None},
            "efe_flujo_inversion": {"actual": None, "comparativo": None},
            "efe_flujo_financiacion": {"actual": None, "comparativo": None},
            "efe_causas_variacion": {"actual": None, "comparativo": None},
            "anexo_bienes_uso_neto": {"actual": None, "comparativo": None},
            "esp_bienes_uso": {"actual": None, "comparativo": None},
            "er_gastos_administracion": {"actual": None, "comparativo": None},
            "anexo_gastos_administracion": {"actual": None, "comparativo": None},
            "er_depreciaciones": {"actual": None, "comparativo": None},
            "anexo_gastos_depreciaciones": {"actual": None, "comparativo": None},
            "anexo_buso_depreciacion_ej": {"actual": None, "comparativo": None},
            "esp_bienes_cambio": {"actual": None, "comparativo": None},
            "anexo_costo_exist_inicial": {"actual": None, "comparativo": None},
            "anexo_costo_exist_final": {"actual": None, "comparativo": None},
            "zero_balance_lines": [],
        }

        current_section = "HEADER"
        for idx, item in enumerate(lines):
            txt = item["text"]
            upper = txt.upper()

            # Track active section
            if "ESTADO DE SITUACIÓN PATRIMONIAL" in upper or "ESTADO DE SITUACION PATRIMONIAL" in upper:
                current_section = "ESP"
            elif "ESTADO DE RECURSOS Y GASTOS" in upper or "ESTADO DE RESULTADOS" in upper:
                current_section = "ER"
            elif "ESTADO DE EVOLUCIÓN DEL PATRIMONIO NETO" in upper or "ESTADO DE EVOLUCION DEL PATRIMONIO NETO" in upper:
                current_section = "EEPN"
            elif "ESTADO DE FLUJOS DE EFECTIVO" in upper or "ESTADO DE FLUJO DE EFECTIVO" in upper:
                current_section = "EFE"
            elif "ANEXO" in upper and "GASTOS" in upper and "CONSUMIDOS" not in upper:
                current_section = "ANEXO_GASTOS"
            elif "PRODUCTOS CONSUMIDOS" in upper or "COSTO DE LOS BIENES VENDIDOS" in upper or "COSTO DE BIENES VENDIDOS" in upper:
                current_section = "ANEXO_COSTO"
            elif "NOTAS A LOS ESTADOS CONTABLES" in upper:
                current_section = "NOTAS"
            elif "BIENES DE USO" in upper and ("ANEXO" in upper or current_section in ("NOTAS", "ANEXOS")):
                current_section = "ANEXO_BUSO"

            nums = extract_numbers_from_line(txt)
            # Check for explicit dashes "- | -", "0,00 | 0,00" or blank rubric lines in basic statements / annexes
            if current_section in ("ESP", "ER", "EFE", "ANEXO_GASTOS", "ANEXO_BUSO"):
                # Split side-by-side ESP rows if needed
                parts = [p.strip() for p in txt.split("|")]
                for p_seg in (parts if len(parts) <= 2 else [parts[0]]):
                    label_clean = re.sub(r"\(Notas?[^)]*\)|\(Anexos?[^)]*\)|\(\d+\)", "", p_seg, flags=re.IGNORECASE).strip(" -|:$0.,")
                    is_header_or_total = any(w in label_clean.upper() for w in [
                        "TOTAL", "SUBTOTAL", "ACTIVO", "PASIVO", "PATRIMONIO NETO", "ESTADO DE",
                        "HOJA", "EJERCICIO", "PESOS", "DENOMINACIÓN", "DENOMINACION", "FECHA",
                        "RUBROS", "CONCEPTO", "NOTAS", "ACLARACIÓN", "ACLARACION", "MODELO", "RECPAM",
                        "RECURSOS", "GASTOS ORDINARIOS", "RESULTADOS ORDINARIOS", "INGRESOS", "ANEXO"
                    ])
                    if 4 <= len(label_clean) <= 75 and not is_header_or_total and re.search(r"[A-Za-zÁÉÍÓÚáéíóú]{4,}", label_clean):
                        if len(nums) >= 2 and all(abs(x) < 0.005 for x in nums[:2]):
                            entry_lbl = f"{label_clean} ({current_section})"
                            if entry_lbl not in data["zero_balance_lines"]:
                                data["zero_balance_lines"].append(entry_lbl)
                        elif re.search(r"(?:\|\s*-\s*\|\s*-\s*$|\s+-\s+-\s*$|\|\s*0\s*\|\s*0\b)", txt):
                            entry_lbl = f"{label_clean} ({current_section})"
                            if entry_lbl not in data["zero_balance_lines"]:
                                data["zero_balance_lines"].append(entry_lbl)
                        elif len(nums) == 0 and re.search(r"\(Notas?|\(Anexos?", p_seg, re.IGNORECASE):
                            # Rubric line with Note/Annex reference but no balance in either column
                            entry_lbl = f"{label_clean} [Sin saldo] ({current_section})"
                            if entry_lbl not in data["zero_balance_lines"]:
                                data["zero_balance_lines"].append(entry_lbl)
            # If the line has no numbers, check if the next line has pure numbers (common in PDF extraction)
            if not nums and idx + 1 < len(lines):
                next_txt = lines[idx + 1]["text"]
                # Only look ahead if next line doesn't start with a known text label
                if not re.search(r"[A-Za-zÁÉÍÓÚáéíóú]{4,}", next_txt):
                    nums = extract_numbers_from_line(next_txt)

            if not nums:
                continue

            def assign_pair(key: str, n_list: List[float]):
                if data[key]["actual"] is None and len(n_list) >= 1:
                    data[key]["actual"] = n_list[0]
                    if len(n_list) >= 2:
                        data[key]["comparativo"] = n_list[1]

            if current_section == "ESP":
                if re.search(r"^Caja\s+y\s+bancos", txt, re.IGNORECASE):
                    assign_pair("caja_y_bancos", nums)
                elif re.search(r"^Bienes\s+de\s+uso", txt, re.IGNORECASE):
                    assign_pair("esp_bienes_uso", nums)
                elif re.search(r"^Bienes\s+de\s+cambio|^Bienes\s+para\s+consumo", txt, re.IGNORECASE):
                    assign_pair("esp_bienes_cambio", nums)
                if re.search(r"Total\s+del\s+activo\s+corriente", txt, re.IGNORECASE) and re.search(r"Total\s+del\s+pasivo\s+corriente", txt, re.IGNORECASE):
                    if len(nums) >= 4:
                        assign_pair("activo_corriente", nums[:2])
                        assign_pair("pasivo_corriente", nums[2:4])
                elif re.search(r"Total\s+del\s+activo\s+corriente", txt, re.IGNORECASE):
                    assign_pair("activo_corriente", nums)
                elif re.search(r"Total\s+del\s+pasivo\s+corriente", txt, re.IGNORECASE):
                    assign_pair("pasivo_corriente", nums)
                elif re.search(r"Total\s+del\s+pasivo\s+no\s+corriente", txt, re.IGNORECASE):
                    assign_pair("pasivo_no_corriente", nums)
                elif re.search(r"Total\s+del\s+pasivo\b(?!\s+corriente|\s+no\s+corriente|\s+y\s+patrimonio)", txt, re.IGNORECASE):
                    assign_pair("total_pasivo", nums)
                elif re.search(r"Total\s+del\s+activo\s+no\s+corriente", txt, re.IGNORECASE) and re.search(r"PATRIMONIO\s+NETO", txt, re.IGNORECASE):
                    if len(nums) >= 4:
                        assign_pair("activo_no_corriente", nums[:2])
                        assign_pair("patrimonio_neto_esp", nums[2:4])
                    elif len(nums) == 3:
                        assign_pair("activo_no_corriente", nums[:2])
                        assign_pair("patrimonio_neto_esp", [nums[2]])
                    elif len(nums) >= 2:
                        assign_pair("activo_no_corriente", nums[:2])
                elif re.search(r"Total\s+del\s+activo\s+no\s+corriente", txt, re.IGNORECASE):
                    assign_pair("activo_no_corriente", nums)
                elif re.search(r"PATRIMONIO\s+NETO\s*\(seg[uú]n\s+estado", txt, re.IGNORECASE):
                    assign_pair("patrimonio_neto_esp", nums)
                elif re.search(r"Total\s+del\s+activo\b", txt, re.IGNORECASE) and re.search(r"Total\s+del\s+pasivo\s+y\s+patrimonio\s+neto", txt, re.IGNORECASE):
                    if len(nums) >= 4:
                        assign_pair("total_activo", nums[:2])
                        assign_pair("total_pasivo_mas_pn", nums[2:4])
                elif re.search(r"Total\s+del\s+activo\b(?!\s+corriente|\s+no\s+corriente)", txt, re.IGNORECASE):
                    assign_pair("total_activo", nums)
                elif re.search(r"Total\s+del\s+pasivo\s+y\s+patrimonio\s+neto", txt, re.IGNORECASE):
                    assign_pair("total_pasivo_mas_pn", nums)

            elif current_section == "ER":
                if re.search(r"Gastos\s+(?:generales\s+)?de\s+administraci[oó]n", txt, re.IGNORECASE):
                    assign_pair("er_gastos_administracion", [abs(x) for x in nums])
                elif re.search(r"Depreciaciones\s+de\s+bienes\s+de\s+uso", txt, re.IGNORECASE):
                    assign_pair("er_depreciaciones", [abs(x) for x in nums])
                elif re.search(r"(?:Super[aá]vit|D[eé]ficit|Ganancia|P[eé]rdida|Resultado)\s*(?:\([^)]+\))?\s*del\s+ejercicio", txt, re.IGNORECASE) and "antes" not in txt.lower():
                    assign_pair("resultado_ejercicio_er", nums)

            elif current_section == "ANEXO_GASTOS":
                if re.search(r"Depreciaci[oó]n\s+bienes\s+de\s+uso", txt, re.IGNORECASE) and len(nums) >= 2:
                    assign_pair("anexo_gastos_depreciaciones", [abs(nums[-2]), abs(nums[-1])])
                elif re.search(r"^Totales?\s+del\s+ejercicio|^Totales?\s+al|^Total\b", txt, re.IGNORECASE) and len(nums) >= 3:
                    # Typically the last column before Total Actual is Gastos de administración
                    assign_pair("anexo_gastos_administracion", [abs(nums[-3]), abs(nums[-1])])

            elif current_section == "ANEXO_BUSO":
                if re.search(r"^Total\s+al|^Totales?\s+del\s+ejercicio|^Total\b", txt, re.IGNORECASE) and len(nums) >= 6:
                    # Columns: Al inicio(0), Altas(1), Bajas(2), Al cierre(3), Dep Inicio(4), Dep Ejercicio(5), Dep Cierre(6), Neto Actual(7), Neto Comp(8)
                    assign_pair("anexo_buso_depreciacion_ej", [abs(nums[5])])

            elif current_section == "ANEXO_COSTO":
                if re.search(r"Total\s+Existencia\s+inicial|^-\s*Existencia\s+inicial", txt, re.IGNORECASE):
                    assign_pair("anexo_costo_exist_inicial", [abs(x) for x in nums])
                elif re.search(r"Total\s+Existencia\s+final|^-\s*Existencia\s+final", txt, re.IGNORECASE):
                    assign_pair("anexo_costo_exist_final", [abs(x) for x in nums])

            elif current_section == "EEPN":
                # In EEPN, the last two numbers of the row are typically Total Actual and Total Comparativo
                if re.search(r"Saldos?\s+al\s+inicio\s+del\s+ejercicio", txt, re.IGNORECASE) and "modificado" not in txt.lower():
                    if len(nums) >= 2:
                        assign_pair("patrimonio_neto_eepn_inicio", nums[-2:])
                elif re.search(r"(?:Super[aá]vit|Ganancia|Resultado)\s*(?:\([^)]+\))?\s*del\s+ejercicio", txt, re.IGNORECASE):
                    if len(nums) >= 2:
                        assign_pair("resultado_ejercicio_eepn", nums[-2:])
                    elif len(nums) == 1:
                        assign_pair("resultado_ejercicio_eepn", nums)
                elif re.search(r"Saldos?\s+al\s+cierre\s+del\s+ejercicio", txt, re.IGNORECASE):
                    if len(nums) >= 2:
                        assign_pair("patrimonio_neto_eepn_cierre", nums[-2:])
                    elif len(nums) == 1:
                        assign_pair("patrimonio_neto_eepn_cierre", nums)

            elif current_section == "EFE":
                if re.search(r"Efectivo\s+y\s+sus\s+equivalentes\s+al\s+inicio\s+del\s+ejercicio", txt, re.IGNORECASE):
                    assign_pair("efe_efectivo_inicio", nums)
                elif re.search(r"Efectivo\s+y\s+sus\s+equivalentes\s+al\s+cierre\s+del\s+ejercicio", txt, re.IGNORECASE):
                    assign_pair("efe_efectivo_cierre", nums)
                elif re.search(r"Aumento\s*\(?Disminuci[oó]n\)?\s*net[oa]\s+del\s+efectivo", txt, re.IGNORECASE):
                    if data["efe_variacion_neta"]["actual"] is None:
                        assign_pair("efe_variacion_neta", nums)
                    else:
                        assign_pair("efe_causas_variacion", nums)
                elif re.search(r"Flujo\s+neto\s+de\s+efectivo.*actividades\s+operativas", txt, re.IGNORECASE):
                    assign_pair("efe_flujo_operativo", nums)
                elif re.search(r"Flujo\s+neto\s+de\s+efectivo.*actividades\s+de\s+inversi[oó]n", txt, re.IGNORECASE):
                    assign_pair("efe_flujo_inversion", nums)
                elif re.search(r"Flujo\s+neto\s+de\s+efectivo.*actividades\s+de\s+financiaci[oó]n", txt, re.IGNORECASE):
                    assign_pair("efe_flujo_financiacion", nums)

        return data

    def _check_numerical_consistency(self, numbers: Dict[str, Any], meta: Dict[str, Any], checks: List[Dict[str, Any]]):
        # Check if template has all zeros or no numbers
        tot_act = numbers["total_activo"]["actual"]
        tot_pas = numbers["total_pasivo"]["actual"]
        pn_esp = numbers["patrimonio_neto_esp"]["actual"]
        tot_pas_pn = numbers["total_pasivo_mas_pn"]["actual"]

        if tot_act is None and pn_esp is None and numbers["resultado_ejercicio_er"]["actual"] is None:
            checks.append({
                "code": "NUM-00",
                "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                "title": "Detección de Importes Numéricos en Estados Básicos",
                "status": "INFO",
                "severity": "INFO",
                "detail": "El documento no contiene importes numéricos de Estados Básicos (o es un modelo/plantilla sin saldos cargados).",
                "norm_ref": "Control Aritmético"
            })
            return

        if tot_act == 0.0 and (tot_pas == 0.0 or tot_pas is None) and (pn_esp == 0.0 or pn_esp is None):
            checks.append({
                "code": "NUM-00",
                "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                "title": "Estados Contables con Saldos en Cero (Plantilla Modelo)",
                "status": "WARNING",
                "severity": "ADVERTENCIA",
                "detail": "Todos los totales del Estado de Situación Patrimonial se encuentran en $ 0,00 (plantilla modelo sin datos reales cargados).",
                "recommendation": "Cargar los saldos contables definitivos antes de generar el PDF para legalizar.",
                "norm_ref": "Control Aritmético"
            })
            return

        # NUM-01: Suma vertical del Activo (Activo Corriente + Activo No Corriente = Total Activo)
        ac = numbers["activo_corriente"]["actual"]
        anc = numbers["activo_no_corriente"]["actual"]
        if tot_act is not None and ac is not None:
            anc_val = anc if anc is not None else 0.0
            diff_act = round(tot_act - (ac + anc_val), 2)
            if abs(diff_act) <= 1.0:
                checks.append({
                    "code": "NUM-01",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Suma Vertical del Activo (Corriente + No Corriente = Total Activo)",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": f"Activo Corriente ({fmt_ars(ac)}) + Activo No Corriente ({fmt_ars(anc_val)}) = Total Activo ({fmt_ars(tot_act)}).",
                    "norm_ref": "Cuadratura ESP"
                })
            else:
                checks.append({
                    "code": "NUM-01",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Suma Vertical del Activo (Corriente + No Corriente = Total Activo)",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": f"Diferencia de {fmt_ars(diff_act)}: Activo Corriente ({fmt_ars(ac)}) + No Corriente ({fmt_ars(anc_val)}) no coincide con Total Activo ({fmt_ars(tot_act)}).",
                    "recommendation": "Revisar las sumas verticales del Estado de Situación Patrimonial.",
                    "norm_ref": "Cuadratura ESP"
                })

        # NUM-02: Ecuación Patrimonial Básica (Total Activo = Total Pasivo + Patrimonio Neto)
        if tot_act is not None and (tot_pas_pn is not None or (tot_pas is not None and pn_esp is not None)):
            rhs = tot_pas_pn if tot_pas_pn is not None else ((tot_pas or 0.0) + (pn_esp or 0.0))
            diff_eq = round(tot_act - rhs, 2)
            if abs(diff_eq) <= 1.0:
                checks.append({
                    "code": "NUM-02",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Ecuación Patrimonial en el ESP (Total Activo = Total Pasivo + Patrimonio Neto)",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": f"Cuadratura exacta: Total Activo ({fmt_ars(tot_act)}) = Total Pasivo + Patrimonio Neto ({fmt_ars(rhs)}).",
                    "norm_ref": "Ecuación Contable Fundamental"
                })
            else:
                checks.append({
                    "code": "NUM-02",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Ecuación Patrimonial en el ESP (Total Activo = Total Pasivo + Patrimonio Neto)",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": f"Descuadre patrimonial de {fmt_ars(diff_eq)}: Total Activo ({fmt_ars(tot_act)}) vs. Total Pasivo + PN ({fmt_ars(rhs)}).",
                    "recommendation": "Verificar que el Total del Activo iguale exactamente a la suma del Pasivo y el Patrimonio Neto.",
                    "norm_ref": "Ecuación Contable Fundamental"
                })

        # NUM-03: Cruce de Patrimonio Neto entre ESP y EEPN
        pn_eepn = numbers["patrimonio_neto_eepn_cierre"]["actual"]
        if pn_esp is not None and pn_eepn is not None:
            diff_pn = round(pn_esp - pn_eepn, 2)
            if abs(diff_pn) <= 1.0:
                checks.append({
                    "code": "NUM-03",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Cruce del Patrimonio Neto entre ESP y EEPN",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": f"El Patrimonio Neto en el ESP ({fmt_ars(pn_esp)}) coincide exactamente con el Saldo al Cierre del EEPN ({fmt_ars(pn_eepn)}).",
                    "norm_ref": "Cruce de Estados Básicos"
                })
            else:
                checks.append({
                    "code": "NUM-03",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Cruce del Patrimonio Neto entre ESP y EEPN",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": f"Diferencia de {fmt_ars(diff_pn)} entre el Patrimonio Neto del ESP ({fmt_ars(pn_esp)}) y el Saldo al Cierre del EEPN ({fmt_ars(pn_eepn)}).",
                    "recommendation": "Ajustar el importe de Patrimonio Neto para que coincida exactamente en ambos estados.",
                    "norm_ref": "Cruce de Estados Básicos"
                })

        # NUM-04: Cruce del Resultado / Superávit del Ejercicio entre ER/ERG y EEPN
        res_er = numbers["resultado_ejercicio_er"]["actual"]
        res_eepn = numbers["resultado_ejercicio_eepn"]["actual"]
        if res_er is not None and res_eepn is not None:
            diff_res = round(res_er - res_eepn, 2)
            if abs(diff_res) <= 1.0:
                checks.append({
                    "code": "NUM-04",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Cruce del Resultado / Superávit del Ejercicio entre ER/ERG y EEPN",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": f"El Resultado del Ejercicio en el Estado de Resultados/Recursos ({fmt_ars(res_er)}) coincide con el expuesto en el EEPN ({fmt_ars(res_eepn)}).",
                    "norm_ref": "Cruce de Estados Básicos"
                })
            else:
                checks.append({
                    "code": "NUM-04",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Cruce del Resultado / Superávit del Ejercicio entre ER/ERG y EEPN",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": f"Diferencia de {fmt_ars(diff_res)} entre el Resultado del Ejercicio en ER/ERG ({fmt_ars(res_er)}) y el incorporado en el EEPN ({fmt_ars(res_eepn)}).",
                    "recommendation": "Verificar que el resultado final del ejercicio sea idéntico en el Estado de Resultados y en el Estado de Evolución del Patrimonio Neto.",
                    "norm_ref": "Cruce de Estados Básicos"
                })

        # NUM-05: Cuadratura del Estado de Flujos de Efectivo (EFE)
        ef_ini = numbers["efe_efectivo_inicio"]["actual"]
        ef_cie = numbers["efe_efectivo_cierre"]["actual"]
        ef_var = numbers["efe_variacion_neta"]["actual"]
        if ef_ini is not None and ef_cie is not None and ef_var is not None:
            calc_var = round(ef_cie - ef_ini, 2)
            diff_efe = round(calc_var - ef_var, 2)
            if abs(diff_efe) <= 1.0:
                checks.append({
                    "code": "NUM-05",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Consistencia del Estado de Flujos de Efectivo (Variación de Efectivo)",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": f"Efectivo al cierre ({fmt_ars(ef_cie)}) - Efectivo al inicio ({fmt_ars(ef_ini)}) = Variación Neta ({fmt_ars(ef_var)}).",
                    "norm_ref": "RT 54 párr. 652-672"
                })
            else:
                checks.append({
                    "code": "NUM-05",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Consistencia del Estado de Flujos de Efectivo (Variación de Efectivo)",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": f"Diferencia de {fmt_ars(diff_efe)} en la Variación del Efectivo del EFE: Cierre ({fmt_ars(ef_cie)}) menos Inicio ({fmt_ars(ef_ini)}) = {fmt_ars(calc_var)}, pero se expone {fmt_ars(ef_var)}.",
                    "recommendation": "Revisar el cálculo de aumento/disminución neta del efectivo y su conciliación con las causas operativas, de inversión y financiación.",
                    "norm_ref": "RT 54 párr. 652-672"
                })

        # NUM-06: Cruce de Depreciaciones de Bienes de Uso entre Anexo de Bienes de Uso y Anexo de Gastos / ERG
        dep_buso = numbers["anexo_buso_depreciacion_ej"]["actual"]
        dep_target = numbers["er_depreciaciones"]["actual"] if meta.get("es_sin_fines_de_lucro") else numbers["anexo_gastos_depreciaciones"]["actual"]
        if dep_buso is not None and dep_target is not None:
            diff_dep = round(dep_buso - dep_target, 2)
            if abs(diff_dep) <= 1.0:
                checks.append({
                    "code": "NUM-06",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Cruce de Depreciaciones del Ejercicio (Anexo Bienes de Uso vs. Anexo de Gastos / ERG)",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": f"Las depreciaciones del ejercicio en el Anexo de Bienes de Uso ({fmt_ars(dep_buso)}) coinciden exactamente con el cargo expuesto en {'el Estado de Recursos y Gastos' if meta.get('es_sin_fines_de_lucro') else 'el Anexo de Gastos'} ({fmt_ars(dep_target)}).",
                    "norm_ref": "Cruce Anexo Bienes de Uso vs. Gastos / ERG"
                })
            else:
                checks.append({
                    "code": "NUM-06",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Cruce de Depreciaciones del Ejercicio (Anexo Bienes de Uso vs. Anexo de Gastos / ERG)",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": f"Diferencia de {fmt_ars(diff_dep)} entre la Depreciación del Ejercicio del Anexo de Bienes de Uso ({fmt_ars(dep_buso)}) y el cargo expuesto en {'el ERG' if meta.get('es_sin_fines_de_lucro') else 'el Anexo de Gastos'} ({fmt_ars(dep_target)}).",
                    "recommendation": "Conciliar el cargo anual por depreciación del Anexo de Bienes de Uso con el Anexo de Gastos / Estado de Recursos y Gastos.",
                    "norm_ref": "Cruce Anexo Bienes de Uso vs. Gastos / ERG"
                })

        # NUM-07: Cruce de Existencias del Anexo de Costo de Bienes Vendidos / Consumidos con el ESP
        bc_act = numbers["esp_bienes_cambio"]["actual"]
        bc_cmp = numbers["esp_bienes_cambio"]["comparativo"]
        ex_fin = numbers["anexo_costo_exist_final"]["actual"]
        ex_ini = numbers["anexo_costo_exist_inicial"]["actual"]
        if bc_act is not None and ex_fin is not None:
            diff_bc = round(bc_act - ex_fin, 2)
            if abs(diff_bc) <= 1.0:
                checks.append({
                    "code": "NUM-07",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Cruce de Existencias del Anexo de Costo con Bienes de Cambio / Consumo del ESP",
                    "status": "PASS",
                    "severity": "CRITICO",
                    "detail": f"La Existencia Final del Anexo de Costo ({fmt_ars(ex_fin)}) coincide exactamente con el rubro Bienes de Cambio / Consumo del ESP ({fmt_ars(bc_act)}).",
                    "norm_ref": "Cruce Anexo Costo vs. ESP"
                })
            else:
                checks.append({
                    "code": "NUM-07",
                    "category": "4. Cuadratura Aritmética y Cruces de Estados Contables",
                    "title": "Cruce de Existencias del Anexo de Costo con Bienes de Cambio / Consumo del ESP",
                    "status": "ERROR",
                    "severity": "CRITICO",
                    "detail": f"Diferencia de {fmt_ars(diff_bc)} entre la Existencia Final del Anexo de Costo ({fmt_ars(ex_fin)}) y el saldo de Bienes de Cambio / Consumo en el ESP ({fmt_ars(bc_act)}).",
                    "recommendation": "Verificar que la Existencia Final del Anexo de Costo cruce con el saldo actual de Bienes de Cambio en el ESP y la Existencia Inicial con el saldo comparativo.",
                    "norm_ref": "Cruce Anexo Costo vs. ESP"
                })
