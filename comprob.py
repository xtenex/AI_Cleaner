#!/usr/bin/env python3
"""
analizador_metadata.py
Analiza documentos (DOCX/XLSX/PPTX/ODT/PDF) para detectar rastros de IA
y compara metadatos entre dos versiones (original vs limpio).

Uso:
    python analizador_metadata.py detect archivo.docx
    python analizador_metadata.py detect archivo.pdf
    python analizador_metadata.py diff original.docx limpio.docx

Dependencias:
    pip install pypdf
    # ExifTool (opcional, para verificación adicional)
"""

import os
import sys
import zipfile
import re
import json
import hashlib
from pathlib import Path
from datetime import datetime
from xml.etree import ElementTree as ET

# ═══════════════════════════════════════════════════════════
# PATRONES DE DETECCIÓN DE IA
# ═══════════════════════════════════════════════════════════

AI_APPLICATIONS = [
    # Asistentes de IA directos
    r"chatgpt", r"\bgpt[-\s]?\d", r"\bgpt[-\s]?4", r"openai",
    r"claude", r"anthropic", r"copilot", r"bing\s*chat",
    r"midjourney", r"dall[-\s]?e", r"stable[-\s]?diffusion",
    r"gemini", r"bard", r"llama", r"mistral", r"deepseek",
    r"perplexity", r"notion\s*ai", r"grammarly",
    # Librerías de generación de PDF (usadas por pipelines de IA)
    r"reportlab", r"weasyprint", r"wkhtmltopdf", r"pandoc",
    r"python[-\s]?docx", r"fpdf", r"pdfkit", r"pdf[-\s]?creator",
    r"chromium", r"headlesschrome", r"skia[/\s]?pdf",
    r"puppeteer", r"playwright", r"selenium",
    # Herramientas de conversión automatizadas
    r"libreoffice\s*headless", r"unoconv", r"soffice",
    r"aspose", r"spire", r"docx4j", r"apache\s*poi",
    r"onlyoffice", r"documents\.ai", r"pandoc",
    # Generadores genéricos
    r"auto[-\s]?generated", r"machine[-\s]?generated",
    r"ai[-\s]?generated", r"ai[-\s]?assisted",
    r"neural", r"deep[-\s]?learning", r"transformer",
]

AI_CUSTOM_PROP_KEYS = [
    r"ai", r"model", r"prompt", r"seed", r"generation",
    r"diffusion", r"llm", r"neural", r"deep",
    r"chatgpt", r"claude", r"openai", r"anthropic",
    r"midjourney", r"dalle", r"stable",
    r"source[-\s]?model", r"generator[-\s]?type",
]

# Palabras en campos de metadatos que sugieren IA
AI_FIELD_WORDS = [
    "ai", "gpt", "claude", "openai", "anthropic", "copilot",
    "midjourney", "dall-e", "stable diffusion", "gemini",
    "reportlab", "weasyprint", "pandoc", "python",
    "generated", "automated", "script", "api",
]

# ═══════════════════════════════════════════════════════════
# CLASES DE RESULTADO
# ═══════════════════════════════════════════════════════════

class Finding:
    def __init__(self, severity, category, field, value, explanation):
        self.severity = severity      # "CRITICAL", "WARNING", "INFO"
        self.category = category      # "AI_APP", "AI_C2PA", "AI_TIMING", etc.
        self.field = field
        self.value = value
        self.explanation = explanation

    def __str__(self):
        icons = {"CRITICAL": "🔴", "WARNING": "🟡", "INFO": "🔵"}
        icon = icons.get(self.severity, "⚪")
        return f"{icon} [{self.severity}] {self.field}: {self.value}\n" \
               f"   → {self.explanation}"


class AnalysisResult:
    def __init__(self, filepath):
        self.filepath = filepath
        self.findings = []
        self.all_metadata = {}
        self.ai_score = 0

    def add(self, finding):
        self.findings.append(finding)
        if finding.severity == "CRITICAL":
            self.ai_score += 10
        elif finding.severity == "WARNING":
            self.ai_score += 5
        else:
            self.ai_score += 1

    @property
    def verdict(self):
        if self.ai_score >= 20:
            return "🔴 ALTA PROBABILIDAD DE SER CREADO POR IA"
        elif self.ai_score >= 10:
            return "🟡 RASTROS SUGERENTES DE IA"
        elif self.ai_score >= 5:
            return "🟡 POSIBLES INCONSISTENCIAS"
        else:
            return "🟢 NO SE DETECTAN RASTROS CLAROS DE IA"

    def report(self):
        lines = []
        lines.append(f"\n{'═'*60}")
        lines.append(f"  ANÁLISIS: {os.path.basename(self.filepath)}")
        lines.append(f"  Tamaño: {os.path.getsize(self.filepath):,} bytes")
        lines.append(f"  SHA-256: {hashlib.sha256(open(self.filepath,'rb').read()).hexdigest()[:16]}...")
        lines.append(f"{'═'*60}")
        lines.append(f"\n  VEREDICTO: {self.verdict} (score: {self.ai_score})")

        if self.findings:
            lines.append(f"\n  {'─'*50}")
            lines.append(f"  HALLAZGOS ({len(self.findings)}):")
            lines.append(f"  {'─'*50}")
            for f in sorted(self.findings, key=lambda x: {"CRITICAL":0,"WARNING":1,"INFO":2}[x.severity]):
                lines.append(f"  {f}")
                lines.append("")
        else:
            lines.append(f"\n  ✅ No se encontraron rastros de IA.")

        # Mostrar todos los metadatos
        if self.all_metadata:
            lines.append(f"\n  {'─'*50}")
            lines.append(f"  METADATOS COMPLETOS:")
            lines.append(f"  {'─'*50}")
            for section, fields in self.all_metadata.items():
                if fields:
                    lines.append(f"    [{section}]")
                    for k, v in fields.items():
                        lines.append(f"      {k}: {v}")
                    lines.append("")

        lines.append(f"{'═'*60}\n")
        return "\n".join(lines)


# ═══════════════════════════════════════════════════════════
# UTILIDADES
# ═══════════════════════════════════════════════════════════

def check_ai_pattern(text, patterns):
    """Devolver el patrón que coincida, o None."""
    if not text:
        return None
    text_lower = text.lower()
    for p in patterns:
        if re.search(p, text_lower, re.IGNORECASE):
            return p
    return None


def extract_zip_entries(filepath):
    """Extraer todos los archivos de un ZIP (DOCX, ODT, etc.)."""
    entries = {}
    with zipfile.ZipFile(filepath, 'r') as zf:
        for name in zf.namelist():
            try:
                entries[name] = zf.read(name)
            except Exception:
                pass
    return entries


def parse_xml(data):
    """Parsear XML y devolver el root, o None."""
    try:
        return ET.fromstring(data)
    except ET.ParseError:
        return None


def get_xml_text(element, tag):
    """Extraer texto de un elemento XML con namespaces."""
    if element is None:
        return None
    # Buscar en todos los namespaces
    for child in element:
        # Quitar namespace del tag
        local_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if local_tag == tag:
            return (child.text or "").strip()
    return None


# ═══════════════════════════════════════════════════════════
# ANÁLISIS DE DOCX / XLSX / PPTX
# ═══════════════════════════════════════════════════════════

def analyze_ooxml(filepath):
    """Analizar archivo OOXML (DOCX, XLSX, PPTX)."""
    result = AnalysisResult(filepath)
    ext = Path(filepath).suffix.lower()
    
    if ext == ".docx":
        prefix = "word"
    elif ext == ".xlsx":
        prefix = "xl"
    else:
        prefix = "ppt"

    entries = extract_zip_entries(filepath)
    
    # ── 1. docProps/core.xml ──
    core_data = entries.get("docProps/core.xml", b"")
    core = parse_xml(core_data) if core_data else None
    core_fields = {}
    
    if core:
        for child in core:
            local_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            val = (child.text or "").strip()
            core_fields[local_tag] = val
        
        result.all_metadata["core.xml"] = core_fields
        
        # Verificar creator
        creator = core_fields.get("creator", "")
        if check_ai_pattern(creator, AI_APPLICATIONS):
            result.add(Finding("CRITICAL", "AI_APP", "core.xml → creator",
                              creator, f"El campo 'creator' contiene un nombre de herramienta de IA."))
        
        # Verificar lastModifiedBy
        last_mod = core_fields.get("lastModifiedBy", "")
        if check_ai_pattern(last_mod, AI_APPLICATIONS):
            result.add(Finding("CRITICAL", "AI_APP", "core.xml → lastModifiedBy",
                              last_mod, "El campo 'lastModifiedBy' contiene un nombre de herramienta de IA."))
        
        # Revision = 1 o vacío (sin historia de edición)
        revision = core_fields.get("revision", "")
        if revision in ("1", ""):
            result.add(Finding("WARNING", "AI_TIMING", "core.xml → revision",
                              revision or "(vacío)",
                              "Revision=1 o vacío: no hay historia de edición. Un documento humano normalmente tiene varias revisiones."))
        
        # Diferencia entre created y modified
        created = core_fields.get("created", "")
        modified = core_fields.get("modified", "")
        if created and modified:
            try:
                c = datetime.fromisoformat(created.replace("Z", "+00:00"))
                m = datetime.fromisoformat(modified.replace("Z", "+00:00"))
                diff = abs((m - c).total_seconds())
                if diff < 30:
                    result.add(Finding("WARNING", "AI_TIMING", "core.xml → created vs modified",
                                      f"Δ = {diff:.0f}s",
                                      "La diferencia entre creación y modificación es < 30s. Sugiere generación automática."))
                elif diff < 5:
                    result.add(Finding("CRITICAL", "AI_TIMING", "core.xml → created vs modified",
                                      f"Δ = {diff:.0f}s",
                                      "Creación y modificación casi simultáneas (< 5s). Fuerte indicio de generación por script."))
            except (ValueError, TypeError):
                pass

    # ── 2. docProps/app.xml ──
    app_data = entries.get("docProps/app.xml", b"")
    app = parse_xml(app_data) if app_data else None
    app_fields = {}
    
    if app:
        for child in app:
            local_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            val = (child.text or "").strip()
            app_fields[local_tag] = val
        
        result.all_metadata["app.xml"] = app_fields
        
        # Application
        application = app_fields.get("Application", "")
        if not application:
            result.add(Finding("CRITICAL", "AI_APP", "app.xml → Application",
                              "(VACÍO)",
                              "El campo Application está vacío. Esto es una firma clásica de documentos generados por librerías JavaScript (docx.js) o pipelines de IA."))
        elif check_ai_pattern(application, AI_APPLICATIONS):
            result.add(Finding("CRITICAL", "AI_APP", "app.xml → Application",
                              application,
                              "El campo Application identifica una herramienta de IA o librería de generación."))
        
        # Company
        company = app_fields.get("Company", "")
        if check_ai_pattern(company, AI_APPLICATIONS):
            result.add(Finding("WARNING", "AI_APP", "app.xml → Company",
                              company,
                              "El campo Company menciona una herramienta de IA."))
        
        # TotalTime = 0
        total_time = app_fields.get("TotalTime", "")
        if total_time in ("0", ""):
            result.add(Finding("WARNING", "AI_TIMING", "app.xml → TotalTime",
                              total_time or "(vacío)",
                              "Tiempo total de edición = 0. Un documento humano normalmente acumula minutos de edición."))
        
        # Template vacío + Application vacío = muy sospechoso
        template = app_fields.get("Template", "")
        if not application and not template and not total_time:
            result.add(Finding("CRITICAL", "AI_APP", "app.xml → (global)",
                              "app.xml vacío/estriado",
                              "app.xml completamente vacío o sin datos significativos. Firma de generación programática."))

    # ── 3. docProps/custom.xml ──
    custom_data = entries.get("docProps/custom.xml", b"")
    custom = parse_xml(custom_data) if custom_data else None
    custom_fields = {}
    
    if custom:
        for child in custom:
            local_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if local_tag == "property":
                name = child.get("name", "")
                # Buscar valor
                val = ""
                for sub in child:
                    sub_tag = sub.tag.split('}')[-1] if '}' in sub.tag else sub.tag
                    if sub_tag in ("lpwstr", "vt:lpwstr", "i4", "vt:i4", "bool", "vt:bool"):
                        val = (sub.text or "").strip()
                custom_fields[name] = val
        
        result.all_metadata["custom.xml"] = custom_fields
        
        for key, val in custom_fields.items():
            if check_ai_pattern(key, AI_CUSTOM_PROP_KEYS) or check_ai_pattern(val, AI_APPLICATIONS):
                result.add(Finding("CRITICAL", "AI_CUSTOM", f"custom.xml → {key}",
                                  val,
                                  f"Propiedad personalizada con nombre/valor relacionado con IA."))

    # ── 4. C2PA Manifest ──
    c2pa_path = "META-INF/content_credential.c2pa"
    if c2pa_path in entries:
        c2pa_size = len(entries[c2pa_path])
        result.add(Finding("CRITICAL", "AI_C2PA", "META-INF/content_credential.c2pa",
                          f"{c2pa_size:,} bytes",
                          "Manifiesto C2PA (Content Credentials) presente. Este es un estándar de trazabilidad que identifica contenido generado o editado por IA."))
    
    # También buscar en otras ubicaciones
    for name in entries:
        if "c2pa" in name.lower() or "content_credential" in name.lower():
            if name != c2pa_path:
                result.add(Finding("CRITICAL", "AI_C2PA", name,
                                  f"{len(entries[name]):,} bytes",
                                  "Archivo C2PA/Content Credential encontrado en ubicación no estándar."))

    # ── 5. Buscar en [Content_Types].xml referencias a C2PA ──
    ct_data = entries.get("[Content_Types].xml", b"")
    if ct_data:
        ct_str = ct_data.decode("utf-8", errors="ignore")
        if "c2pa" in ct_str.lower():
            result.add(Finding("WARNING", "AI_C2PA", "[Content_Types].xml",
                              "referencia a c2pa",
                              "El Content_Types declara un tipo de contenido C2PA."))

    # ── 6. Buscar en _rels/.rels referencias ──
    rels_data = entries.get("_rels/.rels", b"")
    if rels_data:
        rels_str = rels_data.decode("utf-8", errors="ignore")
        if "c2pa" in rels_str.lower():
            result.add(Finding("WARNING", "AI_C2PA", "_rels/.rels",
                              "relación a c2pa",
                              "El manifiesto C2PA está vinculado como parte del paquete OPC."))

    # ── 7. Buscar en el body del documento (word/document.xml) ──
    body_data = entries.get(f"{prefix}/document.xml", b"")
    if not body_data:
        # Intentar rutas alternativas
        for name in entries:
            if name.endswith("document.xml") and "word" in name:
                body_data = entries[name]
                break
    if body_data:
        body_str = body_data.decode("utf-8", errors="ignore")
        # Buscar comentarios o propiedades incrustadas
        if re.search(r'w:customXml[^>]*>.*?(?:ai|gpt|claude|prompt|generation)', body_str, re.IGNORECASE):
            result.add(Finding("WARNING", "AI_BODY", "document.xml",
                              "customXml con referencias a IA",
                              "Se encontraron elementos customXml en el body con referencias a IA."))

    # ── 8. Buscar archivos adicionales sospechosos ──
    suspicious_files = [n for n in entries if any(k in n.lower() for k in 
                     ["ai_", "_ai", "generation", "prompt", "model_info", "provenance"])]
    for sf in suspicious_files:
        result.add(Finding("WARNING", "AI_FILE", sf,
                          f"{len(entries[sf]):,} bytes",
                          "Archivo interno con nombre relacionado con IA/generación."))

    return result


# ═══════════════════════════════════════════════════════════
# ANÁLISIS DE ODT / ODS / ODP
# ═══════════════════════════════════════════════════════════

def analyze_odt(filepath):
    """Analizar archivo OpenDocument (ODT, ODS, ODP)."""
    result = AnalysisResult(filepath)
    entries = extract_zip_entries(filepath)
    
    # ── meta.xml ──
    meta_data = entries.get("meta.xml", b"")
    meta = parse_xml(meta_data) if meta_data else None
    meta_fields = {}
    
    if meta:
        for child in meta.iter():
            local_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if local_tag in ("creator", "creation-date", "modified-date", "generator",
                            "editing-duration", "editing-cycles", "document-statistic"):
                val = (child.text or "").strip()
                if not val:
                    # Atributos
                    attrs = {k.split('}')[-1]: v for k, v in child.attrib.items()}
                    val = str(attrs) if attrs else ""
                meta_fields[local_tag] = val
        
        result.all_metadata["meta.xml"] = meta_fields
        
        creator = meta_fields.get("creator", "")
        if check_ai_pattern(creator, AI_APPLICATIONS):
            result.add(Finding("CRITICAL", "AI_APP", "meta.xml → creator",
                              creator, "El creador es una herramienta de IA."))
        
        generator = meta_fields.get("generator", "")
        if check_ai_pattern(generator, AI_APPLICATIONS):
            result.add(Finding("CRITICAL", "AI_APP", "meta.xml → generator",
                              generator, "El generador es una herramienta de IA."))
        elif not generator:
            result.add(Finding("WARNING", "AI_APP", "meta.xml → generator",
                              "(vacío)", "Generador vacío: inusual para un documento ODT legítimo."))
        
        # editing-cycles
        cycles = meta_fields.get("editing-cycles", "")
        if cycles in ("0", "1", ""):
            result.add(Finding("WARNING", "AI_TIMING", "meta.xml → editing-cycles",
                              cycles or "(vacío)",
                              "Ciclos de edición = 0 o 1. Sugiere generación sin edición humana."))

    # ── C2PA ──
    for name in entries:
        if "c2pa" in name.lower() or "content_credential" in name.lower():
            result.add(Finding("CRITICAL", "AI_C2PA", name,
                              f"{len(entries[name]):,} bytes",
                              "Manifiesto C2PA presente."))

    # ── content.xml: buscar referencias ──
    content_data = entries.get("content.xml", b"")
    if content_data:
        content_str = content_data.decode("utf-8", errors="ignore")
        if re.search(r'(?:ai[-\s]?generated|gpt|claude|openai|anthropic)', content_str, re.IGNORECASE):
            result.add(Finding("INFO", "AI_BODY", "content.xml",
                              "referencia textual",
                              "Se encontraron términos de IA en el contenido XML."))

    return result


# ═══════════════════════════════════════════════════════════
# ANÁLISIS DE PDF
# ═══════════════════════════════════════════════════════════

def analyze_pdf(filepath):
    """Analizar archivo PDF."""
    result = AnalysisResult(filepath)
    
    try:
        from pypdf import PdfReader
    except ImportError:
        print("  [!] Instala pypdf: pip install pypdf")
        return result
    
    reader = PdfReader(filepath)
    
    # ── Info Dictionary ──
    info = reader.metadata or {}
    info_fields = {}
    for key in ["/Title", "/Author", "/Creator", "/Producer", "/CreationDate", "/ModDate",
                "/Subject", "/Keywords", "/Trapped"]:
        val = info.get(key)
        if val:
            info_fields[key.lstrip("/")] = str(val)
    
    result.all_metadata["Info Dict"] = info_fields
    
    # Verificar Producer
    producer = str(info.get("/Producer", ""))
    if check_ai_pattern(producer, AI_APPLICATIONS):
        result.add(Finding("CRITICAL", "AI_APP", "Producer",
                          producer,
                          "El campo Producer identifica una librería/herramienta asociada a generación por IA."))
    
    # Verificar Creator
    creator = str(info.get("/Creator", ""))
    if check_ai_pattern(creator, AI_APPLICATIONS):
        result.add(Finding("CRITICAL", "AI_APP", "Creator",
                          creator,
                          "El campo Creator identifica una herramienta de IA."))
    
    # Verificar Author
    author = str(info.get("/Author", ""))
    if check_ai_pattern(author, AI_APPLICATIONS):
        result.add(Finding("CRITICAL", "AI_APP", "Author",
                          author,
                          "El campo Author menciona una herramienta de IA."))
    
    # Producer vacío
    if not producer:
        result.add(Finding("WARNING", "AI_APP", "Producer",
                          "(vacío)",
                          "Producer vacío: inusual para un PDF generado por software institucional."))
    
    # ── Fechas ──
    creation_date = str(info.get("/CreationDate", ""))
    mod_date = str(info.get("/ModDate", ""))
    
    if creation_date and mod_date:
        try:
            # Formato PDF: D:YYYYMMDDHHmmSS
            def parse_pdf_date(s):
                s = s.replace("D:", "").replace("Z", "").replace("+00'00'", "")
                s = s[:14]
                return datetime.strptime(s, "%Y%m%d%H%M%S")
            
            c = parse_pdf_date(creation_date)
            m = parse_pdf_date(mod_date)
            diff = abs((m - c).total_seconds())
            
            if diff < 5:
                result.add(Finding("CRITICAL", "AI_TIMING", "CreationDate vs ModDate",
                                  f"Δ = {diff:.0f}s",
                                  "Creación y modificación casi simultáneas (< 5s). Fuerte indicio de generación por script."))
            elif diff < 30:
                result.add(Finding("WARNING", "AI_TIMING", "CreationDate vs ModDate",
                                  f"Δ = {diff:.0f}s",
                                  "Diferencia < 30s entre creación y modificación."))
        except (ValueError, TypeError):
            pass
    
    # ── XMP Metadata ──
    xmp_data = reader.xmp_metadata if hasattr(reader, 'xmp_metadata') else None
    xmp_fields = {}
    
    if xmp_data:
        try:
            for prop in ['xmp:CreatorTool', 'xmp:CreateDate', 'xmp:ModifyDate',
                        'xmp:MetadataDate', 'pdf:Producer', 'dc:creator']:
                val = getattr(xmp_data, prop, None)
                if val:
                    xmp_fields[prop] = str(val)
            
            result.all_metadata["XMP"] = xmp_fields
            
            creator_tool = xmp_fields.get("xmp:CreatorTool", "")
            if check_ai_pattern(creator_tool, AI_APPLICATIONS):
                result.add(Finding("CRITICAL", "AI_APP", "XMP → CreatorTool",
                                  creator_tool,
                                  "XMP CreatorTool identifica una herramienta de IA."))
        except Exception:
            pass
    
    # ── Buscar C2PA en el PDF ──
    # Leer bytes crudos para buscar firma C2PA
    with open(filepath, 'rb') as f:
        pdf_bytes = f.read()
    
    if b"content_credential" in pdf_bytes or b"C2PA" in pdf_bytes or b"jumbf" in pdf_bytes.lower():
        result.add(Finding("CRITICAL", "AI_C2PA", "PDF (bytes crudos)",
                          "firma C2PA/JUMBF detectada",
                          "Se detectaron bytes de un manifiesto C2PA/JUMBF embebido en el PDF."))
    
    # ── Estructura del PDF ──
    # Contar objetos y verificar patrones
    try:
        num_pages = len(reader.pages)
        result.all_metadata["Estructura"] = {"Páginas": str(num_pages)}
        
        # Verificar si hay incremental updates (re-saves)
        # Un PDF generado por IA normalmente no tiene incremental updates
        if b"/Prev" not in pdf_bytes[:5000]:
            result.add(Finding("INFO", "AI_STRUCTURE", "PDF → estructura",
                              "sin incremental updates",
                              "El PDF no tiene actualizaciones incrementales. Un documento editado por un humano normalmente las tiene."))
    except Exception:
        pass
    
    # ── Fuentes (señal débil) ──
    try:
        fonts = set()
        for page in reader.pages[:5]:  # Primeras 5 páginas
            if page.get("/Resources"):
                res = page["/Resources"]
                if res.get("/Font"):
                    for font_name in res["/Font"]:
                        try:
                            f_obj = res["/Font"][font_name]
                            base = f_obj.get("/BaseFont", b"")
                            if isinstance(base, bytes):
                                base = base.decode("latin-1")
                            fonts.add(str(base))
                        except Exception:
                            pass
        if fonts:
            result.all_metadata["Fuentes"] = ", ".join(sorted(fonts)[:10])
    except Exception:
        pass
    
    return result


# ═══════════════════════════════════════════════════════════
# MODO DIFF: Comparar dos archivos
# ═══════════════════════════════════════════════════════════

def extract_all_metadata(filepath):
    """Extraer TODOS los metadatos de un archivo como dict plano."""
    ext = Path(filepath).suffix.lower()
    meta = {}
    
    if ext in (".docx", ".xlsx", ".pptx"):
        entries = extract_zip_entries(filepath)
        
        # core.xml
        core_data = entries.get("docProps/core.xml", b"")
        if core_data:
            core = parse_xml(core_data)
            if core:
                for child in core:
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    meta[f"core.{tag}"] = (child.text or "").strip()
        
        # app.xml
        app_data = entries.get("docProps/app.xml", b"")
        if app_data:
            app = parse_xml(app_data)
            if app:
                for child in app:
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    meta[f"app.{tag}"] = (child.text or "").strip()
        
        # custom.xml
        custom_data = entries.get("docProps/custom.xml", b"")
        if custom_data:
            custom = parse_xml(custom_data)
            if custom:
                for child in custom:
                    if child.tag.split('}')[-1] == "property":
                        name = child.get("name", "?")
                        val = ""
                        for sub in child:
                            if sub.text:
                                val = (sub.text or "").strip()
                                break
                        meta[f"custom.{name}"] = val
        
        # C2PA
        c2pa = "META-INF/content_credential.c2pa"
        if c2pa in entries:
            meta["c2pa.size"] = str(len(entries[c2pa]))
            meta["c2pa.sha256"] = hashlib.sha256(entries[c2pa]).hexdigest()[:16]
        
        # Lista de archivos internos
        meta["_files"] = ",".join(sorted(entries.keys()))
        
    elif ext in (".odt", ".ods", ".odp"):
        entries = extract_zip_entries(filepath)
        
        meta_data = entries.get("meta.xml", b"")
        if meta_data:
            meta_str = meta_data.decode("utf-8", errors="ignore")
            # Extraer campos con regex
            patterns = {
                "creator": r'<dc:creator>(.*?)</dc:creator>',
                "generator": r'<generator>(.*?)</generator>',
                "creation-date": r'<dcterms:created[^>]*>(.*?)</dcterms:created>',
                "modified-date": r'<dcterms:modified[^>]*>(.*?)</dcterms:modified>',
                "editing-cycles": r'<meta:editing-cycles>(.*?)</meta:editing-cycles>',
            }
            for key, pat in patterns.items():
                m = re.search(pat, meta_str)
                if m:
                    meta[f"meta.{key}"] = m.group(1).strip()
        
        # C2PA
        for name in entries:
            if "c2pa" in name.lower():
                meta[f"c2pa.{name}"] = f"{len(entries[name])} bytes"
        
        meta["_files"] = ",".join(sorted(entries.keys()))
        
    elif ext == ".pdf":
        try:
            from pypdf import PdfReader
            reader = PdfReader(filepath)
            info = reader.metadata or {}
            for key in ["/Title", "/Author", "/Creator", "/Producer", 
                       "/CreationDate", "/ModDate", "/Subject", "/Keywords"]:
                val = info.get(key)
                if val:
                    meta[f"pdf.{key.lstrip('/')}"] = str(val)
            
            # XMP
            try:
                xmp = reader.xmp_metadata
                if xmp:
                    for prop in ['xmp:CreatorTool', 'xmp:CreateDate', 'xmp:ModifyDate']:
                        val = getattr(xmp, prop, None)
                        if val:
                            meta[f"xmp.{prop}"] = str(val)
            except Exception:
                pass
            
            # C2PA
            with open(filepath, 'rb') as f:
                pdf_bytes = f.read()
            if b"content_credential" in pdf_bytes or b"C2PA" in pdf_bytes:
                meta["c2pa"] = "presente"
                
        except ImportError:
            meta["_error"] = "pypdf no instalado"
    
    return meta


def diff_metadata(file_a, file_b, label_a="ORIGINAL", label_b="LIMPIO"):
    """Comparar metadatos entre dos archivos y mostrar diferencias."""
    meta_a = extract_all_metadata(file_a)
    meta_b = extract_all_metadata(file_b)
    
    all_keys = sorted(set(list(meta_a.keys()) + list(meta_b.keys())))
    
    # Separar _files
    files_a = set(meta_a.pop("_files", "").split(",")) if meta_a.get("_files") else set()
    files_b = set(meta_b.pop("_files", "").split(",")) if meta_b.get("_files") else set()
    
    lines = []
    lines.append(f"\n{'═'*70}")
    lines.append(f"  COMPARACIÓN DE METADATOS")
    lines.append(f"  {label_a}: {os.path.basename(file_a)}")
    lines.append(f"  {label_b}: {os.path.basename(file_b)}")
    lines.append(f"{'═'*70}")
    
    # Diferencias en archivos internos
    if files_a != files_b:
        removed = files_a - files_b
        added = files_b - files_a
        if removed:
            lines.append(f"\n  📁 ARCHIVOS INTERNOS REMOVIDOS:")
            for f in sorted(removed):
                lines.append(f"     - {f}")
        if added:
            lines.append(f"\n  📁 ARCHIVOS INTERNOS AÑADIDOS:")
            for f in sorted(added):
                lines.append(f"     + {f}")
    
    # Diferencias en campos
    lines.append(f"\n  {'─'*70}")
    lines.append(f"  {'CAMPO':<35} {'ORIGINAL':<25} {'LIMPIO':<25}")
    lines.append(f"  {'─'*70}")
    
    changed = 0
    for key in all_keys:
        if key.startswith("_"):
            continue
        val_a = meta_a.get(key, "—")
        val_b = meta_b.get(key, "—")
        
        if val_a != val_b:
            changed += 1
            # Truncar valores largos
            va = val_a if len(val_a) <= 23 else val_a[:20] + "..."
            vb = val_b if len(val_b) <= 23 else val_b[:20] + "..."
            
            # Marcar tipo de cambio
            if val_a == "—" or not val_a:
                marker = "🆕"  # Nuevo
            elif val_b == "—" or not val_b:
                marker = "🗑️"  # Eliminado
            else:
                marker = "✏️"  # Modificado
            
            lines.append(f"  {marker} {key:<35} {va:<25} {vb:<25}")
    
    if changed == 0:
        lines.append(f"\n  ✅ No hay diferencias en los metadatos.")
    else:
        lines.append(f"\n  Total de cambios: {changed}")
    
    # Resumen de impacto
    lines.append(f"\n  {'─'*70}")
    lines.append(f"  RESUMEN DEL IMPACTO:")
    
    # Verificar si se eliminaron rastros de IA
    ai_fields_a = {k: v for k, v in meta_a.items() if check_ai_pattern(v, AI_APPLICATIONS) or check_ai_pattern(k, AI_CUSTOM_PROP_KEYS)}
    ai_fields_b = {k: v for k, v in meta_b.items() if check_ai_pattern(v, AI_APPLICATIONS) or check_ai_pattern(k, AI_CUSTOM_PROP_KEYS)}
    
    if ai_fields_a:
        lines.append(f"  🔴 Rastros de IA en ORIGINAL: {len(ai_fields_a)}")
        for k, v in ai_fields_a.items():
            lines.append(f"       {k} = {v}")
    
    if ai_fields_b:
        lines.append(f"  🔴 Rastros de IA en LIMPIO: {len(ai_fields_b)}")
        for k, v in ai_fields_b.items():
            lines.append(f"       {k} = {v}")
    elif ai_fields_a:
        lines.append(f"  🟢 Rastros de IA en LIMPIO: 0 (todos eliminados)")
    
    if not ai_fields_a and not ai_fields_b:
        lines.append(f"  ℹ️  No se detectaron rastros de IA en ninguno.")
    
    # Verificar C2PA
    c2pa_a = any("c2pa" in k.lower() for k in meta_a)
    c2pa_b = any("c2pa" in k.lower() for k in meta_b)
    if c2pa_a and not c2pa_b:
        lines.append(f"  🟢 C2PA: presente en ORIGINAL → eliminado en LIMPIO")
    elif c2pa_a and c2pa_b:
        lines.append(f"  🔴 C2PA: presente en AMBOS")
    
    lines.append(f"\n{'═'*70}\n")
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════
# CLI PRINCIPAL
# ═══════════════════════════════════════════════════════════

def main():
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Analizador de metadatos para detectar rastros de IA en documentos",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  %(prog)s detect informe.docx
  %(prog)s detect informe.pdf --json
  %(prog)s detect carpeta/ -r
  %(prog)s diff original.docx limpio.docx
  %(prog)s diff original.pdf limpio.pdf --label-a "Antes" --label-b "Después"
        """
    )
    
    subparsers = parser.add_subparsers(dest="mode", required=True)
    
    # Subcomando: detect
    detect_p = subparsers.add_parser("detect", help="Detectar rastros de IA en un archivo")
    detect_p.add_argument("path", help="Archivo o carpeta a analizar")
    detect_p.add_argument("-r", "--recursive", action="store_true", help="Procesar recursivamente")
    detect_p.add_argument("--json", action="store_true", help="Salida en JSON")
    detect_p.add_argument("--verbose", "-v", action="store_true", help="Mostrar todos los metadatos")
    
    # Subcomando: diff
    diff_p = subparsers.add_parser("diff", help="Comparar metadatos entre dos archivos")
    diff_p.add_argument("original", help="Archivo original (sin limpiar)")
    diff_p.add_argument("clean", help="Archivo limpio (procesado)")
    diff_p.add_argument("--label-a", default="ORIGINAL", help="Etiqueta para el archivo original")
    diff_p.add_argument("--label-b", default="LIMPIO", help="Etiqueta para el archivo limpio")
    
    args = parser.parse_args()
    
    SUPPORTED_EXTS = {".docx", ".xlsx", ".pptx", ".odt", ".ods", ".odp", ".pdf"}
    
    if args.mode == "detect":
        path = Path(args.path)
        
        if path.is_file():
            files = [path]
        elif path.is_dir():
            if args.recursive:
                files = [f for f in path.rglob("*") if f.suffix.lower() in SUPPORTED_EXTS]
            else:
                files = [f for f in path.iterdir() if f.suffix.lower() in SUPPORTED_EXTS]
        else:
            print(f"[✗] Ruta no encontrada: {args.path}")
            sys.exit(1)
        
        if not files:
            print("[✗] No se encontraron archivos soportados")
            sys.exit(1)
        
        results = []
        
        for f in files:
            ext = f.suffix.lower()
            try:
                if ext in (".docx", ".xlsx", ".pptx"):
                    r = analyze_ooxml(str(f))
                elif ext in (".odt", ".ods", ".odp"):
                    r = analyze_odt(str(f))
                elif ext == ".pdf":
                    r = analyze_pdf(str(f))
                else:
                    continue
                results.append(r)
            except Exception as e:
                print(f"[✗] Error analizando {f.name}: {e}")
                continue
        
        # Salida
        if args.json:
            output = []
            for r in results:
                output.append({
                    "file": str(r.filepath),
                    "verdict": r.verdict,
                    "score": r.ai_score,
                    "findings": [
                        {
                            "severity": f.severity,
                            "category": f.category,
                            "field": f.field,
                            "value": f.value,
                            "explanation": f.explanation
                        } for f in r.findings
                    ],
                    "metadata": r.all_metadata
                })
            print(json.dumps(output, indent=2, ensure_ascii=False))
        else:
            for r in results:
                if args.verbose or r.findings:
                    print(r.report())
                elif not r.findings:
                    print(f"\n  ✅ {os.path.basename(r.filepath)}: {r.verdict}")
    
    elif args.mode == "diff":
        if not os.path.exists(args.original):
            print(f"[✗] No encontrado: {args.original}")
            sys.exit(1)
        if not os.path.exists(args.clean):
            print(f"[✗] No encontrado: {args.clean}")
            sys.exit(1)
        
        print(diff_metadata(args.original, args.clean, args.label_a, args.label_b))


if __name__ == "__main__":
    main()   
