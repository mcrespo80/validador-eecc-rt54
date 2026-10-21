# MIC DIGITAL CONSULTING — Validador y Auto-Remediador de Estados Contables (RT 54 / RT 37 / CPCECABA)

Sistema integral de **Auditoría Pre-Legalización y Generación de Versión Mejorada de Estados Contables (EECC)** desarrollado para cumplir sin objeciones con los controles formales, contables y de auditoría del **Consejo Profesional de Ciencias Económicas de la Ciudad Autónoma de Buenos Aires (CPCECABA)** y del **CPCEPBA**.

---

## 🎯 Normativa Contable y de Auditoría Incorporada

- **Resolución Técnica N° 54 (T.O. RT 59 - NUA)**: Norma Unificada Argentina de Contabilidad.
- **Resolución P. N° 460/2024 CPCECABA**: Adopción de la RT 54 (T.O. RT 59) en el ámbito de la Ciudad Autónoma de Buenos Aires.
- **Resolución JG FACPCE N° 660/2026**: Dispensas transitorias de exposición en primera aplicación.
- **Resolución Técnica N° 37 (modificada por RT 53)**: Nuevo orden obligatorio de secciones del Informe del Auditor Independiente (*1° Opinión*, *2° Fundamento de la opinión*), adoptada por **Res. C.D. N° 46/2021 CPCECABA** y **Res. C.D. N° 3843 CPCEPBA**.
- **Ley N° 17.250 Art. 10 (SIPA)** y **RG IGJ N° 15/2024 Art. 240**.

---

## ✨ Funcionalidades Principales

1. **Ingesta Multi-Formato con OCR Automático (`core/extractor.py`)**:
   - Soporta archivos **PDF nativos**, **PDF escaneados** (mediante motor neural `RapidOCR` local), **Word (`.docx`)** y **Excel (`.xlsx`)**.
2. **Motor de Validación de 29 Reglas (`core/validator.py` & `knowledge_base/reglas_rt54_cpcecaba.json`)**:
   - **Controles Formales Legalizaciones Web (`FOR-00` a `FOR-05`)**: Límite de 10 MB, validación matemática de CUIT (Módulo 11), detección de textos en borrador (`dd/mm/aaaa`, `XYZ`, `{De corresponder}`), consignación expresa de *Domicilio legal* y leyendas de integración de Notas y Anexos.
   - **Cumplimiento RT 54 NUA (`RT54-01` a `RT54-15`)**:
     - Integridad de los 4 Estados Básicos (ESP, ER/ERG, EEPN y EFE).
     - Nota 1.1 de adopción de RT 54 (T.O. RT 59) y Res. P. 460/2024 CPCECABA.
     - Nota 1.2 de Categorización obligatoria (*Entidad Pequeña - EP*, *Entidad Mediana - EM* o *Restante Entidad - RE*).
     - Nota 1.3 de Unidad de Medida (*Moneda Homogénea* e índice FACPCE).
     - Reclasificación de rubros a nomenclatura NUA (*Créditos impositivos*, *Inversiones financieras*, *Proveedores de bienes y servicios*, *Deudas laborales y previsionales*, *Deudas fiscales*).
     - **Supresión obligatoria de rubros, filas y notas con saldo cero en ambos períodos (`RT54-07`)**.
     - **Estructura del Anexo de Previsiones (`RT54-08`)**: Separación entre *Deducidas del Activo* e *Incluidas en el Pasivo*, exposición del RECPAM en columna *"Otros (*)"* y criterio de medición (nominal en EP vs. descontado en RE).
     - **Cuadro / Anexo de Gastos (`RT54-09`)**: Matriz por naturaleza y función (Art. 64 inc. b Ley 19.550 / RT 54 párr. 638 y 727) y cruce con el ER/ERG.
     - **Prohibición de EFE Sintético en Entidades Sin Fines de Lucro y Cooperativas (`RT54-10` — RT 54 párr. 666)**.
     - **Conciliación entre Efectivo del EFE y Caja y Bancos del ESP (`RT54-11` — RT 54 párr. 672 inc. a)**.
     - **Evaluación de Valor Recuperable ante resultados negativos (`RT54-12` — Regla de los 3 ejercicios positivos)**.
     - **Transición de Primera Aplicación (`RT54-13`)**, **Aportes Irrevocables RG IGJ 15/2024 (`RT54-14`)** y **Control de Referencias Cruzadas Huérfanas (`RT54-15`)**.
   - **Auditoría del Informe del Auditor Independiente (`RT37-01` a `RT37-07`)**: Título oficial, orden de párrafos, cita de Res. CD 46/2021 y Código de Ética, deuda SIPA Art. 10 Ley 17.250, libros rubricados, control cronológico de fechas y párrafo de *Otra Información (Memoria)*.
   - **Cuadratura Aritmética y Cruces (`NUM-01` a `NUM-07`)**: Ecuación patrimonial en el ESP, cruce de PN entre ESP y EEPN, cruce de Resultado entre ER/ERG y EEPN, variación del efectivo en el EFE, cruce de depreciaciones con Anexo de Bienes de Uso y cruce de existencias con Anexo de Costo.
3. **Generador Automático de «Versión Mejorada Lista para Presentar en Consejo» (`core/remediator.py`)**:
   - Elimina automáticamente las filas y rubros con saldo cero en ambos ejercicios.
   - Actualiza las denominaciones de rubros al estándar NUA de la RT 54 y limpia marcadores de plantilla.
   - Incorpora las Notas obligatorias RT 54, el Anexo de Previsiones, el Cuadro de Gastos y el Informe del Auditor Independiente RT 37 (mod. RT 53) listo para legalizar en formatos **PDF** y **Word Editable (`.docx`)**.

---

## 🚀 Ejecución Rápida

1. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
   ```
2. Iniciar el servidor web (o ejecutar `iniciar_validador.bat`):
   ```bash
   python app.py
   ```
3. Abrir en el navegador:
   ```
   http://127.0.0.1:8054
   ```

---

**MIC DIGITAL CONSULTING** • `micdigital.com.ar` • `contacto@micdigital.com.ar` • Tel: `156-111-7714`
