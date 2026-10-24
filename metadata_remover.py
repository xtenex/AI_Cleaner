#!/usr/bin/env python3
"""
limpiador_metadata.py
Herramienta para eliminar rastros de IA y editar metadatos
de documentos e imágenes para hacerlos pasar como creados
por Microsoft Office o LibreOffice/OpenOffice.

Dependencias:
    pip install python-docx Pillow
    # ExifTool debe estar instalado en el sistema (recomendado)
    # Windows: choco install exiftool  o descargar de exiftool.org
"""

import os
import sys
import shutil
import subprocess
import zipfile
import re
from datetime import datetime, timedelta
from pathlib import Path
from io import BytesIO

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────

# Opciones de "disfraz" para documentos
PROFILES = {
    "word": {
        "author": "Juan Pérez",
        "last_modified_by": "Juan Pérez",
        "application": "Microsoft Office Word",
        "company": "Microsoft Corporation",
    },
    "libreoffice": {
        "author": "Juan Pérez",
        "last_modified_by": "Juan Pérez",
        "application": "LibreOffice",
        "company": "The Document Foundation",
    },
    "openoffice": {
        "author": "Juan Pérez",
        "last_modified_by": "Juan Pérez",
        "application": "OpenOffice.org",
        "company": "Sun Microsystems",
    },
}

# Extensiones soportadas
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".tiff", ".bmp", ".gif"}
DOC_EXTS = {".docx", ".xlsx", ".pptx"}
ODT_EXTS = {".odt", ".ods", ".odp"}
PDF_EXTS = {".pdf"}

# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────

def has_exiftool() -> bool:
    """Verificar si ExifTool está disponible."""
    return shutil.which("exiftool") is not None


def random_date(start_days_ago=30, end_days_ago=1):
    """Generar una fecha aleatoria en el pasado."""
    days_ago = end_days_ago + (start_days_ago - end_days_ago) * 0.5  # medio
    return datetime.now() - timedelta(days=days_ago, hours=3, minutes=42)


def random_date_modified():
    """Fecha de modificación más reciente."""
    return datetime.now() - timedelta(hours=5, minutes=17)


# ─────────────────────────────────────────────
# IMÁGENES: Eliminar metadatos de IA
# ─────────────────────────────────────────────

def clean_image(path: str, preserve_icc: bool = True, in_place: bool = True) -> bool:
    """
    Elimina TODOS los metadatos de una imagen:
    - C2PA / JUMBF (Content Credentials)
    - XMP (incluye prompts de IA, model info)
    - IPTC (DigitalSourceType, etc.)
    - EXIF completo
    - PNG text chunks
    - CBOR
    
    Usa ExifTool para edición lossless. Si no está disponible,
    usa Pillow como fallback (re-codifica).
    """
    path = str(path)
    
    if has_exiftool():
        # ExifTool: edición lossless, elimina TODO
        cmd = [
            "exiftool",
            "-all=",                    # elimina todos los metadatos
            "-JUMBF:all=",              # C2PA manifests
            "-CBOR:all=",               # CBOR data (nuevos formatos)
            "-overwrite_original",
        ]
        if preserve_icc:
            # Preservar perfil ICC para no alterar colores
            cmd += ["-tagsfromfile", "@", "-ICCProfile"]
        cmd.append(path)
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  [!] ExifTool warning: {result.stderr.strip()}")
        print(f"  [✓] Metadatos eliminados con ExifTool (lossless)")
        return True
    else:
        # Fallback: Pillow (re-codifica, puede perder calidad en JPEG)
        from PIL import Image, ImageOps
        
        with Image.open(path) as im:
            fmt = im.format
            icc = im.info.get("icc_profile")
            im = ImageOps.exif_transpose(im)
            
            save_options = {"format": fmt}
            if icc:
                save_options["icc_profile"] = icc
            if fmt == "JPEG":
                save_options["quality"] = 95
            if fmt == "WEBP":
                save_options["quality"] = 95
            
            output = path if in_place else path + "_clean"
            im.save(output, **save_options)
        
        print(f"  [✓] Metadatos eliminados con Pillow (re-codificado)")
        print(f"  [!] Nota: instala ExifTool para evitar re-codificación")
        return True


# ─────────────────────────────────────────────
# DOCX / XLSX / PPTX: Editar metadatos
# ─────────────────────────────────────────────

def clean_docx(path: str, profile: str = "word", in_place: bool = True) -> bool:
    """
    Edita metadatos de archivos .docx/.xlsx/.pptx
    para hacerlos pasar como creados por Office o LibreOffice.
    """
    from docx import Document
    import docx
    
    path = str(path)
    profile_data = PROFILES.get(profile, PROFILES["word"])
    
    # Para .docx usamos python-docx directamente
    ext = Path(path).suffix.lower()
    
    if ext == ".docx":
        doc = docx.Document(path)
        props = doc.core_properties
    elif ext == ".xlsx":
        # openpyxl para Excel
        from openpyxl import load_workbook
        wb = load_workbook(path)
        props = wb.properties
    elif ext == ".pptx":
        from pptx import Presentation
        prs = Presentation(path)
        props = prs.core_properties
    else:
        print(f"  [!] Formato no soportado: {ext}")
        return False
    
    # Aplicar perfil
    created = random_date()
    modified = random_date_modified()
    
    props.author = profile_data["author"]
    props.last_modified_by = profile_data["last_modified_by"]
    props.created = created
    props.modified = modified
    props.revision = 3  # simular ediciones
    
    # Para docx, también limpiar campos que delatan IA
    if ext == ".docx":
        props.comments = None
        props.category = None
        props.keywords = None
        props.content_status = None
        props.identifier = None
        props.version = None
    
    # Eliminar propiedades personalizadas que puedan delatar
    try:
        if hasattr(props, 'custom_properties') and props.custom_properties:
            # python-docx no permite borrar fácilmente, pero podemos
            # trabajar a nivel de zip
            _remove_custom_props_docx(path, in_place)
    except Exception:
        pass
    
    # Guardar
    output = path if in_place else path + "_clean"
    
    if ext == ".docx":
        doc.save(output)
    elif ext == ".xlsx":
        wb.save(output)
    elif ext == ".pptx":
        prs.save(output)
    
    # Editar app.xml para cambiar el "Application" string
    _patch_app_xml(output, profile_data["application"])
    
    print(f"  [✓] Metadatos editados → perfil '{profile}'")
    print(f"      Autor: {profile_data['author']}")
    print(f"      App:   {profile_data['application']}")
    print(f"      Creado: {created.strftime('%Y-%m-%d %H:%M')}")
    return True


def _patch_app_xml(docx_path: str, application_name: str):
    """
    Patea el archivo word/app.xml dentro del .docx para
    cambiar el campo 'Application' que python-docx no expone.
    """
    import tempfile
    
    try:
        # Leer el zip
        with zipfile.ZipFile(docx_path, 'r') as zf:
            app_xml_path = None
            for name in zf.namelist():
                if 'app.xml' in name and 'docProps' in name:
                    app_xml_path = name
                    break
            
            if not app_xml_path:
                return
            
            app_xml = zf.read(app_xml_path).decode('utf-8')
            
            # Reemplazar el nombre de la aplicación
            # Formato típico: <Application>Microsoft Office Word</Application>
            app_xml = re.sub(
                r'(<Application>).*?(</Application>)',
                rf'\g<1>{application_name}\g<2>',
                app_xml
            )
            
            # También limpiar "Template" si existe
            app_xml = re.sub(
                r'<Template>.*?</Template>',
                '',
                app_xml
            )
            
            # Reescribir el zip
            tmp = docx_path + ".tmp"
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as new_zf:
                for item in zf.infolist():
                    if item.filename == app_xml_path:
                        new_zf.writestr(item, app_xml.encode('utf-8'))
                    else:
                        new_zf.writestr(item, zf.read(item.filename))
            
            os.replace(tmp, docx_path)
    except Exception as e:
        print(f"  [!] No se pudo editar app.xml: {e}")


def _remove_custom_props_docx(path: str, in_place: bool):
    """Eliminar customProperties.xml si existe."""
    tmp = path + ".tmp"
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as new_zf:
                for item in zf.infolist():
                    if 'customProperties' in item.filename:
                        continue  # omitir
                    new_zf.writestr(item, zf.read(item.filename))
        os.replace(tmp, path)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)


# ─────────────────────────────────────────────
# ODT / ODS / ODP: Editar metadatos
# ─────────────────────────────────────────────

def clean_odt(path: str, profile: str = "libreoffice", in_place: bool = True) -> bool:
    """
    Edita metadatos de archivos ODT/ODS/ODP
    (formato OpenDocument = zip con XML).
    """
    path = str(path)
    profile_data = PROFILES.get(profile, PROFILES["libreoffice"])
    
    created = random_date().strftime("%Y-%m-%dT%H:%M:%S")
    modified = random_date_modified().strftime("%Y-%m-%dT%H:%M:%S")
    
    tmp = path + ".tmp"
    
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as new_zf:
                for item in zf.infolist():
                    data = zf.read(item.filename)
                    
                    if item.filename == 'meta.xml':
                        xml = data.decode('utf-8')
                        
                        # Reemplazar creador
                        xml = re.sub(
                            r'<dc:creator>.*?</dc:creator>',
                            f'<dc:creator>{profile_data["author"]}</dc:creator>',
                            xml
                        )
                        # Reemplazar generador
                        xml = re.sub(
                            r'<generator>.*?</generator>',
                            f'<generator>{profile_data["application"]}</generator>',
                            xml
                        )
                        # Fechas
                        xml = re.sub(
                            r'<dcterms:created[^>]*>.*?</dcterms:created>',
                            f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>',
                            xml
                        )
                        xml = re.sub(
                            r'<dcterms:modified[^>]*>.*?</dcterms:modified>',
                            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified>',
                            xml
                        )
                        # Eliminar metadata de IA si existe
                        xml = re.sub(
                            r'<meta:generation-info>.*?</meta:generation-info>',
                            '',
                            xml,
                            flags=re.DOTALL
                        )
                        data = xml.encode('utf-8')
                    
                    new_zf.writestr(item, data)
        
        output = path if in_place else path + "_clean"
        os.replace(tmp, output)
        
        print(f"  [✓] Metadatos ODT editados → perfil '{profile}'")
        print(f"      Generador: {profile_data['application']}")
        return True
        
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        print(f"  [!] Error editando ODT: {e}")
        return False


# ─────────────────────────────────────────────
# PDF: Editar metadatos
# ─────────────────────────────────────────────

def clean_pdf(path: str, profile: str = "word", in_place: bool = True) -> bool:
    """
    Edita metadatos de PDF usando pypdf.
    """
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        print("  [!] Instala pypdf: pip install pypdf")
        return False
    
    path = str(path)
    profile_data = PROFILES.get(profile, PROFILES["word"])
    
    reader = PdfReader(path)
    writer = PdfWriter()
    
    # Copiar páginas
    for page in reader.pages:
        writer.add_page(page)
    
    # Setear metadatos
    writer.add_metadata({
        "/Title": "Documento",
        "/Author": profile_data["author"],
        "/Creator": profile_data["application"],
        "/Producer": profile_data["application"],
        "/CreationDate": f"D:{random_date().strftime('%Y%m%d%H%M%S')}",
        "/ModDate": f"D:{random_date_modified().strftime('%Y%m%d%H%M%S')}",
    })
    
    output = path if in_place else path + "_clean"
    with open(output, 'wb') as f:
        writer.write(f)
    
    print(f"  [✓] Metadatos PDF editados → perfil '{profile}'")
    return True


# ─────────────────────────────────────────────
# VERIFICACIÓN
# ─────────────────────────────────────────────

def verify(path: str):
    """Verificar metadatos restantes usando ExifTool."""
    if not has_exiftool():
        print("  [!] ExifTool no disponible para verificación")
        return
    
    result = subprocess.run(
        ["exiftool", "-G1", "-a", "-s", str(path)],
        capture_output=True, text=True
    )
    
    lines = [l for l in result.stdout.strip().split('\n') if l.strip()]
    
    # Filtrar líneas relevantes
    ai_indicators = [
        'c2pa', 'jumbf', 'prompt', 'negativeprompt', 
        'digitalsource', 'trainedalgorithmic', 'aiprompt',
        'generation', 'model', 'seed', 'sampler', 'steps',
        'cfgscale', 'lora', 'controlnet', 'vae', 'chatgpt',
        'midjourney', 'dall', 'stable.diffusion'
    ]
    
    suspicious = []
    for line in lines:
        lower = line.lower()
        if any(ind in lower for ind in ai_indicators):
            suspicious.append(line)
    
    if suspicious:
        print(f"  [⚠] Rastros de IA detectados:")
        for s in suspicious:
            print(f"      {s}")
    else:
        print(f"  [✓] No se detectan rastros de IA en los metadatos")
    
    # Mostrar resumen
    if len(lines) <= 10:
        print(f"  Metadatos restantes:")
        for l in lines:
            print(f"    {l}")
    else:
        print(f"  ({len(lines)} campos de metadatos presentes)")


# ─────────────────────────────────────────────
# CLI PRINCIPAL
# ─────────────────────────────────────────────

def main():
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Limpia metadatos de IA y edita documentos",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  %(prog)s foto.jpg                    # Limpiar imagen
  %(prog)s documento.docx              # Limpiar DOCX (perfil Word)
  %(prog)s doc.docx -p libreoffice     # Limpiar DOCX (perfil LibreOffice)
  %(prog)s archivo.odt -p openoffice   # Limpiar ODT (perfil OpenOffice)
  %(prog)s carpeta/ -r                 # Procesar carpeta recursivamente
  %(prog)s foto.jpg -v                 # Limpiar y verificar
  %(prog)s doc.docx --author "María López"  # Autor personalizado
        """
    )
    
    parser.add_argument("path", help="Archivo o carpeta a procesar")
    parser.add_argument("-p", "--profile", choices=PROFILES.keys(),
                       default="word", help="Perfil de aplicación (default: word)")
    parser.add_argument("-r", "--recursive", action="store_true",
                       help="Procesar recursivamente")
    parser.add_argument("-v", "--verify", action="store_true",
                       help="Verificar metadatos después de limpiar")
    parser.add_argument("-o", "--output", default=None,
                       help="Ruta de salida (default: in-place)")
    parser.add_argument("--author", default=None,
                       help="Autor personalizado (sobrescribe perfil)")
    parser.add_argument("--no-icc", action="store_true",
                       help="No preservar perfil ICC en imágenes")
    
    args = parser.parse_args()
    
    # Sobrescribir autor si se especifica
    if args.author:
        for profile in PROFILES.values():
            profile["author"] = args.author
            profile["last_modified_by"] = args.author
    
    # Recopitar archivos
    files = []
    path = Path(args.path)
    
    if path.is_file():
        files = [path]
    elif path.is_dir():
        if args.recursive:
            for f in path.rglob("*"):
                if f.suffix.lower() in IMAGE_EXTS | DOC_EXTS | ODT_EXTS | PDF_EXTS:
                    files.append(f)
        else:
            for f in path.iterdir():
                if f.suffix.lower() in IMAGE_EXTS | DOC_EXTS | ODT_EXTS | PDF_EXTS:
                    files.append(f)
    else:
        print(f"[✗] Ruta no encontrada: {args.path}")
        sys.exit(1)
    
    if not files:
        print("[✗] No se encontraron archivos soportados")
        sys.exit(1)
    
    print(f"\n{'='*50}")
    print(f"  Limpieza de metadatos | {len(files)} archivo(s)")
    print(f"  Perfil: {args.profile}")
    print(f"{'='*50}\n")
    
    for f in files:
        ext = f.suffix.lower()
        print(f"  → {f.name}")
        
        try:
            if ext in IMAGE_EXTS:
                clean_image(str(f), preserve_icc=not args.no_icc)
            elif ext in DOC_EXTS:
                clean_docx(str(f), profile=args.profile)
            elif ext in ODT_EXTS:
                clean_odt(str(f), profile=args.profile)
            elif ext in PDF_EXTS:
                clean_pdf(str(f), profile=args.profile)
            else:
                print(f"  [!] Extensión no soportada: {ext}")
                continue
        except Exception as e:
            print(f"  [✗] Error: {e}")
            continue
        
        if args.verify:
            verify(str(f))
        print()
    
    print(f"{'='*50}")
    print(f"  Completado. {len(files)} archivo(s) procesados.")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()
