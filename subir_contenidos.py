#!/usr/bin/env python3
"""
Subidor automático de contenidos a Compartir Conocimientos (Santillana).
Lee un Excel con la lista de archivos, los busca en la carpeta Completo/,
y los sube aplicando los mismos metadatos a todo el lote.
"""

import sys
import time
import json
import re
from pathlib import Path
from datetime import datetime

try:
    import requests
    from tqdm import tqdm
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    print("Instalando dependencias...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "requests", "tqdm", "openpyxl"])
    import requests
    from tqdm import tqdm
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment


# ── Constantes ──────────────────────────────────────────────────────────────

API_BASE = "https://compartirconocimientos-pe.santillana.com"

# Mapeo extensión → type_guid
# CTTY_05=PDF  CTTY_08=HTML Interactivo  CTTY_12=Office
TYPE_GUID = {
    ".pdf":  "CTTY_05",
    ".zip":  "CTTY_08",   # ZIP = HTML Interactivo (paquete HTML)
    ".docx": "CTTY_12",
    ".doc":  "CTTY_12",
    ".pptx": "CTTY_12",
    ".xlsx": "CTTY_12",
    ".html": "CTTY_08",
    ".htm":  "CTTY_08",
}

# is_teacher_only según "disponible para"
DISPONIBLE_MAP = {
    "1": (0, "Docentes y Estudiantes"),
    "2": (1, "Solo Docentes"),
}

IDIOMAS_MAP = {
    "1": "es",
    "2": "en",
    "3": "pt",
}


# ── Sesión HTTP ─────────────────────────────────────────────────────────────

def build_session(token: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "Authorization": token.strip(),
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://publisher.compartirconocimientos-pe.santillana.com",
        "Referer": "https://publisher.compartirconocimientos-pe.santillana.com/",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/149.0.0.0 Safari/537.36 Edg/149.0.0.0"
        ),
    })
    return s


def api_get(session, path: str, params=None):
    url = f"{API_BASE}{path}"
    r = session.get(url, params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def api_post(session, path: str, json_data=None, files=None, headers_extra=None):
    url = f"{API_BASE}{path}"
    if files:
        # Multipart upload: quitar Content-Type para que requests lo ponga con boundary
        hdrs = {k: v for k, v in session.headers.items() if k.lower() != "content-type"}
        if headers_extra:
            hdrs.update(headers_extra)
        r = requests.post(url, files=files, headers=hdrs, timeout=120)
    else:
        r = session.post(url, json=json_data, timeout=30)
    r.raise_for_status()
    return r.json()


def api_put(session, path: str, json_data: dict):
    url = f"{API_BASE}{path}"
    r = session.put(url, json=json_data, timeout=30)
    r.raise_for_status()
    return r.json()


# ── Menús interactivos ──────────────────────────────────────────────────────

def menu(titulo: str, opciones: list[dict], key_label="name", key_value="guid") -> str:
    """
    Muestra un menú numerado y devuelve el valor del campo key_value elegido.
    opciones: lista de dicts con al menos key_label y key_value.
    """
    print(f"\n{titulo}")
    print("─" * 50)
    for i, op in enumerate(opciones, 1):
        print(f"  {i:>2}. {op[key_label]}")
    print()
    while True:
        entrada = input("Elige número > ").strip()
        if entrada.isdigit() and 1 <= int(entrada) <= len(opciones):
            elegido = opciones[int(entrada) - 1]
            print(f"  ✓ Seleccionado: {elegido[key_label]}")
            return elegido[key_value]
        print(f"  Ingresa un número entre 1 y {len(opciones)}")


def menu_simple(titulo: str, opciones: dict) -> tuple:
    """opciones: {clave: (valor, etiqueta)}. Devuelve (valor, etiqueta)."""
    print(f"\n{titulo}")
    print("─" * 50)
    for k, (_, etiqueta) in opciones.items():
        print(f"  {k}. {etiqueta}")
    print()
    while True:
        entrada = input("Elige número > ").strip()
        if entrada in opciones:
            valor, etiqueta = opciones[entrada]
            print(f"  ✓ Seleccionado: {etiqueta}")
            return valor, etiqueta
        print(f"  Opción no válida.")


# ── Fetch de catálogos desde la API ─────────────────────────────────────────

def fetch_education_levels(session) -> list[dict]:
    try:
        data = api_get(session, "/api/cms/education-levels")
        items = data.get("data", data)
        if isinstance(items, list):
            return [{"guid": x.get("guid", x.get("education_level_guid")),
                     "name": x.get("name", x.get("education_level_name", str(x)))}
                    for x in items]
    except Exception:
        pass
    # Fallback con los valores conocidos
    return [
        {"guid": "00000000-0000-1000-0000-000000000038", "name": "Inicial"},
        {"guid": "00000000-0000-1000-0000-000000000039", "name": "Primaria"},
        {"guid": "00000000-0000-1000-0000-000000000040", "name": "Secundaria"},
    ]


_YEARS_FALLBACK = {
    # Primaria (nivel 039)
    "00000000-0000-1000-0000-000000000039": [
        {"guid": "00000000-0000-1000-0000-000000000119", "name": "1.er grado - Primaria"},
        {"guid": "00000000-0000-1000-0000-000000000120", "name": "2.do grado - Primaria"},
        {"guid": "00000000-0000-1000-0000-000000000121", "name": "3.er grado - Primaria"},
        {"guid": "00000000-0000-1000-0000-000000000122", "name": "4.to grado - Primaria"},
        {"guid": "00000000-0000-1000-0000-000000000123", "name": "5.to grado - Primaria"},
        {"guid": "00000000-0000-1000-0000-000000000124", "name": "6.to grado - Primaria"},
    ],
    # Secundaria (nivel 040)
    "00000000-0000-1000-0000-000000000040": [
        {"guid": "00000000-0000-1000-0000-000000000126", "name": "1.er año - Secundaria"},
        {"guid": "00000000-0000-1000-0000-000000000127", "name": "2.do año - Secundaria"},
        {"guid": "00000000-0000-1000-0000-000000000128", "name": "3.er año - Secundaria"},
        {"guid": "00000000-0000-1000-0000-000000000129", "name": "4.to año - Secundaria"},
        {"guid": "00000000-0000-1000-0000-000000000130", "name": "5.to año - Secundaria"},
    ],
}


def fetch_education_years(session, level_guid: str) -> list[dict]:
    try:
        data = api_get(session, "/api/cms/education-years",
                       params={"education_level_guid": level_guid})
        items = data.get("data", data)
        if isinstance(items, list) and items:
            return [{"guid": x.get("guid", x.get("education_year_guid")),
                     "name": x.get("name", x.get("education_year_name", str(x)))}
                    for x in items]
    except Exception:
        pass
    # Fallback con GUIDs conocidos
    return _YEARS_FALLBACK.get(level_guid, [])


def fetch_disciplines(session) -> list[dict]:
    try:
        data = api_get(session, "/api/cms/disciplines")
        items = data.get("data", data)
        if isinstance(items, list):
            return [{"guid": x.get("guid", x.get("discipline_guid")),
                     "name": x.get("name", x.get("discipline_name", str(x)))}
                    for x in items]
    except Exception:
        pass
    return []


def fetch_collections(session, search: str = "") -> list[dict]:
    data = api_get(session, "/api/cms/collections",
                   params={"pageSize": 100, "search": search})
    cols = data.get("data", {}).get("collections", [])
    return [{"guid": c["guid"], "name": c["collection"]} for c in cols]


# ── Lógica de subida ─────────────────────────────────────────────────────────

def crear_contenido(session, nombre_visible: str, erp_id: str, type_guid: str,
                    is_teacher_only: int) -> str:
    """POST /api/cms/contents → devuelve el guid del contenido creado."""
    payload = {
        "description": "",
        "erp_id": erp_id,
        "is_available_offline": 1,
        "is_teacher_only": is_teacher_only,
        "langs": [],
        "name": nombre_visible,
        "type_guid": type_guid,
    }
    resp = api_post(session, "/api/cms/contents", json_data=payload)
    return resp["data"]["guid"]


def obtener_upload_info(session, guid: str) -> dict:
    """GET /api/cms/contents/{guid}/content/upload → token y endpoint."""
    resp = api_get(session, f"/api/cms/contents/{guid}/content/upload")
    upload = resp["data"]["data"]["upload"]
    return {"token": upload["token"], "endpoint": upload["endpoint"]}


def subir_archivo(session, endpoint: str, upload_token: str, filepath: Path) -> bool:
    """POST multipart al endpoint de files-storage."""
    mime = _mime(filepath)
    with open(filepath, "rb") as f:
        files = {
            "file": (filepath.name, f, mime),
            "token": (None, upload_token),
        }
        resp = api_post(session, endpoint, files=files)
    return resp.get("status") == "success"


def _mime(path: Path) -> str:
    ext = path.suffix.lower()
    return {
        ".pdf":  "application/pdf",
        ".zip":  "application/zip",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".html": "text/html",
    }.get(ext, "application/octet-stream")


def actualizar_metadata(session, guid: str, nombre_visible: str, erp_id: str,
                        type_guid: str, is_teacher_only: int,
                        education_levels: list, education_years: list,
                        disciplines: list, langs: list,
                        collections: list) -> bool:
    payload = {
        "collections":          collections,
        "customTags":           [],
        "dependencies":         [],
        "description":          "",
        "didacticTypes":        [],
        "disciplines":          disciplines,
        "educationLevels":      education_levels,
        "educationYears":       education_years,
        "encoded_transcription_url": None,
        "erp_id":               erp_id,
        "guid":                 guid,
        "is_available_offline": 1,
        "is_downloadable":      0,
        "is_public":            0,
        "is_teacher_only":      is_teacher_only,
        "langs":                langs,
        "learningObjectives":   [],
        "mobile_friendly":      1,
        "name":                 nombre_visible,
        "publications":         [],
        "status":               "active",
        "tags":                 [],
        "topics":               [],
        "transcription_bundle": None,
        "transcription_url":    "",
        "type_guid":            type_guid,
    }
    resp = api_put(session, f"/api/cms/contents/{guid}", json_data=payload)
    return resp.get("status") == "success"


# ── Lectura del Excel ────────────────────────────────────────────────────────

def leer_excel(xlsx_path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    archivos = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        nombre_visible, nombre_archivo, ruta_dest, tipo, carpeta, carpeta_fuente = row
        if not nombre_archivo:
            continue
        ext = Path(str(nombre_archivo)).suffix.lower()
        erp_id = Path(str(nombre_archivo)).stem
        type_guid = TYPE_GUID.get(ext)
        if not type_guid:
            tqdm.write(f"  [AVISO] Extensión desconocida ignorada: {nombre_archivo}")
            continue
        archivos.append({
            "nombre_visible":  str(nombre_visible).strip(),
            "nombre_archivo":  str(nombre_archivo).strip(),
            "erp_id":          erp_id,
            "type_guid":       type_guid,
            "carpeta":         str(carpeta).strip(),
            "tipo":            str(tipo).strip(),
            "carpeta_fuente":  str(carpeta_fuente).strip() if carpeta_fuente else "sin_categoria",
        })
    return archivos


# ── Log de resultados ────────────────────────────────────────────────────────

def guardar_log(resultados: list[dict], ruta_log: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultados"
    headers = ["Archivo", "Nombre visible", "GUID", "Estado", "Error", "URL viewer"]
    fill_h = PatternFill("solid", fgColor="1F4E79")
    font_h = Font(bold=True, color="FFFFFF")
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = fill_h
        c.font = font_h
        c.alignment = Alignment(horizontal="center")

    fill_ok  = PatternFill("solid", fgColor="C6EFCE")
    fill_err = PatternFill("solid", fgColor="FFC7CE")
    for i, r in enumerate(resultados, 2):
        fill = fill_ok if r["estado"] == "OK" else fill_err
        valores = [r["archivo"], r["nombre_visible"], r.get("guid", ""),
                   r["estado"], r.get("error", ""), r.get("url_viewer", "")]
        for col, val in enumerate(valores, 1):
            c = ws.cell(row=i, column=col, value=val)
            c.fill = fill

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 38
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 50
    ws.column_dimensions["F"].width = 70
    wb.save(ruta_log)


# ── Main ────────────────────────────────────────────────────────────────────

def main():
    print("=" * 65)
    print("  Subidor de contenidos — Compartir Conocimientos Santillana")
    print("=" * 65)

    # ── 1. Token JWT ────────────────────────────────────────────────
    print("\nPega tu token JWT (header Authorization completo, ej: eyJ...).")
    token = input("Token > ").strip()
    if not token:
        sys.exit("No se ingresó token.")
    # Agregar prefijo Bearer si no lo tiene
    if not token.startswith("Bearer ") and not token.startswith("eyJ"):
        sys.exit("Token con formato inesperado.")
    auth_header = token if token.startswith("Bearer ") else token  # el API lo recibe sin prefix
    session = build_session(auth_header)

    # ── 2. Rutas (auto-detectadas desde donde se ejecuta el script) ─
    # Estructura esperada:
    #   <raiz>/              ← aquí corre el script
    #   <raiz>/OUTPUT/       ← contiene el Excel .xlsx
    #   <raiz>/OUTPUT/Completo/  ← contiene todos los archivos a subir
    raiz = Path.cwd()
    output_dir   = raiz / "OUTPUT"
    completo_dir = output_dir / "Completo"

    print(f"\n  Directorio raíz   : {raiz}")
    print(f"  Carpeta OUTPUT    : {output_dir}")
    print(f"  Carpeta Completo  : {completo_dir}")

    if not output_dir.is_dir():
        sys.exit(f"\n[ERROR] No se encontró la carpeta OUTPUT en {raiz}\n"
                 f"Ejecuta el script desde la carpeta raíz del proyecto.")
    if not completo_dir.is_dir():
        sys.exit(f"\n[ERROR] No se encontró OUTPUT/Completo en {raiz}")

    # Buscar Excel en OUTPUT/
    excels = list(output_dir.glob("*.xlsx"))
    if not excels:
        sys.exit(f"\n[ERROR] No se encontró ningún .xlsx en {output_dir}")
    if len(excels) == 1:
        xlsx_path = excels[0]
        print(f"  Excel detectado   : {xlsx_path.name}")
    else:
        print("\nVarios Excel encontrados en OUTPUT/:")
        for i, p in enumerate(excels, 1):
            print(f"  {i}. {p.name}")
        idx = int(input("Elige número > ").strip()) - 1
        xlsx_path = excels[idx]

    # ── 3. Leer Excel ───────────────────────────────────────────────
    print(f"\nLeyendo {xlsx_path.name}...")
    archivos = leer_excel(xlsx_path)
    print(f"  {len(archivos)} entradas encontradas en el Excel.")

    # Verificar cuáles existen físicamente
    faltantes = []
    for a in archivos:
        ruta_arch = completo_dir / a["nombre_archivo"]
        a["ruta"] = ruta_arch
        if not ruta_arch.exists():
            faltantes.append(a["nombre_archivo"])

    if faltantes:
        print(f"\n  [AVISO] {len(faltantes)} archivo(s) no encontrados en Completo/:")
        for f in faltantes[:10]:
            print(f"    - {f}")
        if len(faltantes) > 10:
            print(f"    ... y {len(faltantes)-10} más")
        cont = input("\n¿Continuar de todas formas? (s/n) > ").strip().lower()
        if cont != "s":
            sys.exit("Cancelado.")
        archivos = [a for a in archivos if a["ruta"].exists()]
        print(f"  Se subirán {len(archivos)} archivos encontrados.")

    # ── 4. Selección de metadatos ────────────────────────────────────
    print("\n" + "=" * 65)
    print("  Define los metadatos del lote completo")
    print("=" * 65)

    # Etapa / Education Level
    print("\nObteniendo etapas educativas...")
    levels = fetch_education_levels(session)
    level_guid = menu("Etapa educativa", levels)

    # Año / Education Year
    print("\nObteniendo años/series...")
    years = fetch_education_years(session, level_guid)
    if years:
        year_guid = menu("Año / Serie", years)
    else:
        print("  No se pudo obtener la lista de años. Ingresa el GUID manualmente.")
        year_guid = input("GUID año > ").strip()

    # Asignatura / Discipline
    print("\nObteniendo asignaturas...")
    disciplines = fetch_disciplines(session)
    if disciplines:
        disc_guid = menu("Asignatura / Disciplina", disciplines)
    else:
        print("  No se pudo obtener la lista. Ingresa el GUID manualmente.")
        disc_guid = input("GUID asignatura > ").strip()

    # Idioma
    idioma_val, idioma_label = menu_simple(
        "Idioma",
        {"1": ("es", "Español"), "2": ("en", "Inglés"), "3": ("pt", "Portugués")}
    )

    # Colección
    print("\nBuscando colecciones (escribe parte del nombre para filtrar):")
    busq = input("Búsqueda colección > ").strip()
    collections_list = fetch_collections(session, search=busq)
    if not collections_list:
        print("  Sin resultados. Ingresa el GUID de la colección manualmente.")
        col_guid = input("GUID colección > ").strip()
        col_label = col_guid
    else:
        col_guid = menu("Colección", collections_list)
        col_label = next((c["name"] for c in collections_list if c["guid"] == col_guid), col_guid)

    # ── 5. Disponible para — por archivo con defaults inteligentes ───
    TYPE_LABEL = {
        "CTTY_05": "PDF",
        "CTTY_08": "HTML Interactivo",
        "CTTY_12": "Office",
        "CTTY_13": "ZIP",
    }
    DISP_LABEL = {0: "Docentes y Estudiantes", 1: "Solo Docentes"}

    import re as _re

    def default_disp(archivo: dict) -> int:
        # ZIP con patrón de unidad (U01, U1, U02...U08) → ambos
        # Todo lo demás → solo docentes
        es_zip = archivo["nombre_archivo"].lower().endswith(".zip")
        tiene_unidad = bool(_re.search(r'[Uu]\d{1,2}', archivo["nombre_archivo"]))
        return 0 if (es_zip and tiene_unidad) else 1

    # Asignar defaults
    for a in archivos:
        a["is_teacher_only"] = default_disp(a)

    # Mostrar agrupado por carpeta fuente para que el usuario confirme o cambie
    from collections import Counter
    conteo_fuentes = Counter(a["carpeta_fuente"] for a in archivos)
    fuentes_unicas = sorted(conteo_fuentes.keys())

    print("\n" + "=" * 70)
    print("  DISPONIBLE PARA — confirma o cambia el default por tipo de recurso")
    print("=" * 70)
    print("  Regla automática: ZIP con número de unidad (U01-U08) → ambos.")
    print("  Todo lo demás → Solo Docentes.")
    print()
    print("  Presiona Enter para aceptar el default, o escribe 1/2 para cambiar:")
    print("    1 = Docentes y Estudiantes   2 = Solo Docentes")
    print()

    for fuente in fuentes_unicas:
        # Calcular default de la fuente (mayoría)
        vals = [a["is_teacher_only"] for a in archivos if a["carpeta_fuente"] == fuente]
        default_fuente = max(set(vals), key=vals.count)  # valor más frecuente
        default_str = "1" if default_fuente == 0 else "2"
        default_txt = DISP_LABEL[default_fuente]
        n = conteo_fuentes[fuente]

        resp = input(f"  [{fuente}] ({n} archivos)  default={default_txt}  > ").strip()
        if resp == "":
            nuevo = default_fuente
        elif resp == "1":
            nuevo = 0
        elif resp == "2":
            nuevo = 1
        else:
            print(f"    Entrada no válida, se usa el default: {default_txt}")
            nuevo = default_fuente

        # Aplicar a todos los archivos de esta fuente
        for a in archivos:
            if a["carpeta_fuente"] == fuente:
                a["is_teacher_only"] = nuevo

    # ── 6. Previsualización detallada ────────────────────────────────
    level_name = next((l["name"] for l in levels if l["guid"] == level_guid), level_guid)
    year_name  = next((y["name"] for y in years  if y["guid"] == year_guid),  year_guid) if years else year_guid
    disc_name  = next((d["name"] for d in disciplines if d["guid"] == disc_guid), disc_guid) if disciplines else disc_guid

    print("\n" + "=" * 100)
    print("  DETALLE COMPLETO — así se subirá cada archivo")
    print("=" * 100)
    print(f"  {'#':<4} {'Nombre visible':<38} {'Archivo':<36} {'Tipo':<18} {'Disponible para'}")
    print("  " + "─" * 96)
    for i, a in enumerate(archivos, 1):
        tipo_label = TYPE_LABEL.get(a["type_guid"], a["type_guid"])
        disp_label = DISP_LABEL[a["is_teacher_only"]]
        nombre_vis = a["nombre_visible"][:37]
        nombre_arc = a["nombre_archivo"][:35]
        print(f"  {i:<4} {nombre_vis:<38} {nombre_arc:<36} {tipo_label:<18} {disp_label}")
    print("=" * 100)

    print(f"\n  METADATOS COMUNES A TODOS:")
    print(f"    Etapa           : {level_name}")
    print(f"    Año / Serie     : {year_name}")
    print(f"    Asignatura      : {disc_name}")
    print(f"    Idioma          : {idioma_label}")
    print(f"    Colección       : {col_label}")
    print(f"\n  Total: {len(archivos)} archivos")
    print("=" * 100)

    conf = input("\n¿Iniciar subida? (s/n) > ").strip().lower()
    if conf != "s":
        sys.exit("Cancelado por el usuario.")

    # ── 7. Subida ────────────────────────────────────────────────────
    resultados = []
    barra = tqdm(archivos, desc="Subiendo", unit="arch", ncols=80)

    for item in barra:
        nombre         = item["nombre_visible"]
        erp_id         = item["erp_id"]
        type_g         = item["type_guid"]
        filepath       = item["ruta"]
        is_teacher_only = item["is_teacher_only"]
        barra.set_postfix_str(filepath.name[:35])

        resultado = {"archivo": filepath.name, "nombre_visible": nombre,
                     "estado": "ERROR", "guid": "", "error": "", "url_viewer": ""}
        try:
            # Paso 1: crear contenido
            guid = crear_contenido(session, nombre, erp_id, type_g, is_teacher_only)
            resultado["guid"] = guid

            # Paso 2: obtener token de subida
            upload_info = obtener_upload_info(session, guid)

            # Paso 3: subir archivo
            ok_upload = subir_archivo(
                session,
                upload_info["endpoint"],
                upload_info["token"],
                filepath,
            )
            if not ok_upload:
                raise RuntimeError("El endpoint de upload no devolvió success")

            # Esperar brevemente a que el storage procese
            time.sleep(1)

            # Paso 4: actualizar metadata
            ok_meta = actualizar_metadata(
                session, guid, nombre, erp_id, type_g, item["is_teacher_only"],
                education_levels=[level_guid],
                education_years=[year_guid],
                disciplines=[disc_guid],
                langs=[idioma_val],
                collections=[col_guid],
            )
            if not ok_meta:
                raise RuntimeError("Error al actualizar metadatos (PUT)")

            viewer_url = f"https://publisher.compartirconocimientos-pe.santillana.com/content/{guid}/1"
            resultado["estado"] = "OK"
            resultado["url_viewer"] = viewer_url

        except Exception as e:
            resultado["error"] = str(e)
            tqdm.write(f"\n  [ERROR] {filepath.name}: {e}")

        resultados.append(resultado)

    # ── 7. Log de resultados ─────────────────────────────────────────
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = raiz / f"log_subida_{ts}.xlsx"
    guardar_log(resultados, log_path)

    ok_count  = sum(1 for r in resultados if r["estado"] == "OK")
    err_count = len(resultados) - ok_count

    print(f"\n{'=' * 65}")
    print(f"  Subida completada.")
    print(f"  OK      : {ok_count}")
    print(f"  Errores : {err_count}")
    print(f"  Log     : {log_path}")
    print("=" * 65)


if __name__ == "__main__":
    main()
