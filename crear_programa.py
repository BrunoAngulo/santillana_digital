#!/usr/bin/env python3
"""
Crea módulos y recursos en el LMS de Santillana a partir de la estructura local.

Flujo por producto:
  1. Detecta carpetas bajo PRIM/ que tienen OUTPUT/resumen_global.xlsx
  2. Lee el Excel: A=Nombre visible, B=Nombre archivo, C=Ruta, D=Tipo,
                   E=Carpeta contenedora (módulo), F=Carpeta fuente, G=Name (sección)
  3. Aplica fix de nombres:
       - "Unidad XX" → "Name Unidad XX"   (col A empieza con "Unidad")
       - nombre == stem archivo → "Name nombre"  (col A igual al stem de col B)
  4. Busca el curso en el LMS (búsqueda interactiva)
  5. Por cada módulo del Excel (orden: Documentos → Unidad 00 → Unidad 01 …):
       a. Crea una lección (módulo) si no existe ya
       b. Por cada sección dentro del módulo (col G, orden original del Excel):
            - Crea la sección con el nombre de col G
            - Busca cada archivo en CMS por ERP ID y lo agrega a esa sección

Uso:
  python crear_programa.py
  python crear_programa.py "C:\\...\\SUBIDAS\\PRIM"
"""

import re
import sys
import time
from datetime import datetime
from pathlib import Path

try:
    import requests
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from tqdm import tqdm
except ImportError:
    print("Instalando dependencias...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "requests", "openpyxl", "tqdm"])
    import requests
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from tqdm import tqdm


API_BASE = "https://compartirconocimientos-pe.santillana.com"

CONTENT_TYPES = [
    "CTTY_04", "CTTY_01", "CTTY_08", "CTTY_07", "CTTY_06",
    "CTTY_12", "CTTY_05", "CTTY_03", "CTTY_20", "CTTY_13",
]

CARPETAS_EXCLUIR = {"completo"}  # carpetas que no se convierten en módulo

# Mapeo carpeta fuente (col F) → nombre de sección visible en el LMS
SECCION_NOMBRE = {
    "dif_lect":  "Dificultades de lectoescritura",
    "doc_curr":  "Documentos curriculares",
    "est_lect":  "Estrategias de lectura",
    "guia_met":  "Guía metodológica",
    "lam_didac": "Láminas didácticas",
    "libmed":    "Libromedia",
    "mat_imp":   "Módulo de material imprimible",
    "mod_ie":    "Instrumentos de evaluación",
    "senpai":    "Herramientas cooperativas",
}


# ── Sesión HTTP ──────────────────────────────────────────────────────────────

def build_session(token: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "Authorization": token.strip(),
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*",
        "Origin":   "https://publisher.compartirconocimientos-pe.santillana.com",
        "Referer":  "https://publisher.compartirconocimientos-pe.santillana.com/",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/149.0.0.0 Safari/537.36 Edg/149.0.0.0"
        ),
    })
    return s


def api_get(session, path: str, params=None):
    r = session.get(f"{API_BASE}{path}", params=params, timeout=30)
    try:
        r.raise_for_status()
    except requests.HTTPError as e:
        raise requests.HTTPError(f"{e} — {r.text[:300]}", response=r) from None
    return r.json()


def api_post(session, path: str, payload: dict):
    r = session.post(f"{API_BASE}{path}", json=payload, timeout=30)
    try:
        r.raise_for_status()
    except requests.HTTPError as e:
        raise requests.HTTPError(f"{e} — {r.text[:300]}", response=r) from None
    return r.json()


# ── Detección de productos ───────────────────────────────────────────────────

def detectar_productos(base_path: Path) -> list[Path]:
    """Devuelve subcarpetas de base_path que contienen OUTPUT/resumen_global.xlsx."""
    return sorted(
        p for p in base_path.iterdir()
        if p.is_dir() and (p / "OUTPUT" / "resumen_global.xlsx").exists()
    )


# ── Lectura y normalización del Excel ────────────────────────────────────────

def _fix_nombre(nombre_visible: str, nombre_archivo: str, name_serie: str) -> str:
    """
    Si nombre_visible empieza con "Unidad" o es igual al stem del archivo,
    antepone el nombre de la serie: "Dificultades de lectoescritura Unidad 01".
    """
    v     = nombre_visible.strip()
    stem  = Path(nombre_archivo).stem if nombre_archivo else ""
    serie = name_serie.strip() if name_serie else ""

    if serie and (v.lower().startswith("unidad") or v == stem):
        return f"{serie} {v}"
    return v


def leer_excel_producto(xlsx_path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    archivos = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        cols = list(row) + [None] * 8
        nombre_visible = str(cols[0] or "").strip()
        nombre_archivo = str(cols[1] or "").strip()
        # cols[2] ruta_destino  — no se usa aquí
        # cols[3] tipo          — no se usa aquí
        modulo         = str(cols[4] or "").strip()   # Carpeta contenedora (módulo LMS)
        carpeta_fuente = str(cols[5] or "").strip()   # Código de sección (dif_lect, doc_curr, ...)
        name_serie     = str(cols[6] or "").strip()   # Name: prefijo para nombres de contenido
        seccion        = SECCION_NOMBRE.get(carpeta_fuente, carpeta_fuente)  # nombre visible de la sección

        if not nombre_archivo or not modulo:
            continue

        erp_id       = Path(nombre_archivo).stem
        nombre_final = _fix_nombre(nombre_visible, nombre_archivo, name_serie)

        archivos.append({
            "nombre_visible": nombre_final,
            "erp_id":         erp_id,
            "modulo":         modulo,
            "seccion":        seccion,
        })
    return archivos


def agrupar_por_modulo(archivos: list[dict]) -> dict[str, dict[str, list[dict]]]:
    """
    Devuelve {modulo → {seccion → [archivos]}}.
    Excluye módulos en CARPETAS_EXCLUIR.
    Ordena módulos: Documentos primero, luego Unidad numericamente.
    Las secciones mantienen el orden original del Excel.
    """
    estructura: dict[str, dict[str, list[dict]]] = {}
    secciones_orden: dict[str, list[str]] = {}

    for a in archivos:
        modulo  = a["modulo"]
        seccion = a["seccion"]

        if modulo.lower() in CARPETAS_EXCLUIR:
            continue

        if modulo not in estructura:
            estructura[modulo] = {}
            secciones_orden[modulo] = []

        if seccion not in estructura[modulo]:
            estructura[modulo][seccion] = []
            secciones_orden[modulo].append(seccion)

        estructura[modulo][seccion].append(a)

    def sort_key_modulo(nombre: str) -> tuple:
        lower = nombre.lower()
        if "documentos" in lower:
            return (0, 0)
        m = re.search(r'\d+', nombre)
        return (1, int(m.group()) if m else 999)

    resultado: dict[str, dict[str, list[dict]]] = {}
    for modulo in sorted(estructura.keys(), key=sort_key_modulo):
        resultado[modulo] = {
            sec: estructura[modulo][sec]
            for sec in secciones_orden[modulo]
        }

    return resultado


# ── API: cursos ───────────────────────────────────────────────────────────────

def buscar_cursos(session, search: str) -> list[dict]:
    data = api_get(session, "/api/lms/courses", params={
        "offset": 0, "page": 0, "pageSize": 20,
        "isEditorial": 1, "search": search,
    })
    return data.get("data", {}).get("courses", [])


def get_course_items(session, course_guid: str) -> list[dict]:
    data = api_get(session, f"/api/front/courses/{course_guid}/items")
    return data.get("data", {}).get("items", [])


def crear_lesson(session, course_guid: str, nombre_modulo: str) -> str | None:
    """Crea una lección (módulo) en el curso. Devuelve su lesson_guid."""
    html_name = f'<p><span style="font-size: 24px;">{nombre_modulo}</span></p>'
    payload = {
        "evaluation_period_id": "annual",
        "name":               html_name,
        "number_of_sessions": 1,
        "type":               "teacher",
    }
    data = api_post(session, f"/api/front/courses/{course_guid}/lessons", payload)
    d = data.get("data", {})
    return d.get("lesson_guid") or d.get("guid")


def crear_seccion(session, lesson_guid: str, nombre_seccion: str) -> str | None:
    """Crea una sección dentro de la lección. Devuelve su guid (parent_guid)."""
    nombre_seccion = nombre_seccion.strip() or "Recursos"
    payload = {
        "lesson_guid":       lesson_guid,
        "section":           nombre_seccion,
        "section_type_guid": "default",
    }
    r = session.post(f"{API_BASE}/api/front/lesson-items", json=payload, timeout=30)
    try:
        r.raise_for_status()
    except requests.HTTPError as e:
        tqdm.write(f"      [ERROR HTTP] crear_seccion: {e} — {r.text[:200]}")
        return None
    data = r.json()
    guid = (data.get("data") or {}).get("guid")
    if not guid:
        tqdm.write(f"      [WARN] crear_seccion sin guid — resp: {str(data)[:200]}")
    return guid


# ── API: contenidos CMS ───────────────────────────────────────────────────────

def buscar_contenido_por_erp(session, erp_id: str) -> dict | None:
    params = [
        ("offset", 0), ("page", 0), ("pageSize", 5),
        ("search", erp_id), ("isEditorial", 1),
    ]
    for t in CONTENT_TYPES:
        params.append(("type[]", t))

    r = session.get(f"{API_BASE}/api/cms/contents", params=params, timeout=30)
    r.raise_for_status()
    contents = r.json().get("data", {}).get("contents", [])

    # Coincidencia exacta primero
    for c in contents:
        if c.get("erp_id", "").lower() == erp_id.lower():
            return c
    return contents[0] if contents else None


def agregar_contenido(session, lesson_guid: str, parent_guid: str,
                      content: dict, nombre_visible: str) -> str | None:
    """Agrega el contenido a la sección. Devuelve el item_guid creado o None si falla."""
    payload = {
        "content_guid":       content["guid"],
        "description":        "",
        "disciplines":        [d["discipline_guid"]        for d in content.get("disciplines",     [])],
        "educationLevels":    [e["education_level_guid"]   for e in content.get("educationLevels",  [])],
        "educationYears":     [a["education_year_guid"]    for a in content.get("educationYears",   [])],
        "include_in_gradebook": 0,
        "is_embed":           1,
        "item_for":           "all",
        "learningObjectives": [],
        "lesson_guid":        lesson_guid,
        "name":               nombre_visible,
        "parent_guid":        parent_guid,
        "status":             "published",
        "teacher_notes":      None,
    }
    data = api_post(session, "/api/front/lesson-items", payload)
    return data.get("data", {}).get("guid") or None


# ── Log Excel ────────────────────────────────────────────────────────────────

_FILL_HDR   = PatternFill("solid", fgColor="1F4E79")
_FONT_HDR   = Font(bold=True, color="FFFFFF", size=11)
_FILL_OK    = PatternFill("solid", fgColor="C6EFCE")
_FILL_ERR   = PatternFill("solid", fgColor="FFC7CE")
_FILL_MISS  = PatternFill("solid", fgColor="FFEB9C")   # amarillo: NOT FOUND
_FILL_SKIP  = PatternFill("solid", fgColor="D9D9D9")


def guardar_log(entradas: list[dict], log_path: Path):
    """
    Genera un Excel con una fila por cada contenido procesado.
    Columnas: Producto | Módulo | Sección | ERP ID | Nombre visible |
              Estado | GUID CMS | Nombre en CMS | GUID item LMS | Detalle
    Colores: VERDE=CREADO, ROJO=ERROR, AMARILLO=NOT_FOUND, GRIS=SKIP
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Log"

    headers = [
        "Producto", "Módulo", "Sección", "ERP ID", "Nombre visible",
        "Estado", "GUID CMS", "Nombre en CMS", "GUID item LMS", "Detalle",
    ]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = _FILL_HDR
        c.font = _FONT_HDR
        c.alignment = Alignment(horizontal="center")

    for i, e in enumerate(entradas, 2):
        ws.cell(row=i, column=1,  value=e.get("producto",       ""))
        ws.cell(row=i, column=2,  value=e.get("modulo",         ""))
        ws.cell(row=i, column=3,  value=e.get("seccion",        ""))
        ws.cell(row=i, column=4,  value=e.get("erp_id",         ""))
        ws.cell(row=i, column=5,  value=e.get("nombre_visible", ""))
        ws.cell(row=i, column=6,  value=e.get("estado",         ""))
        ws.cell(row=i, column=7,  value=e.get("guid_cms",       ""))
        ws.cell(row=i, column=8,  value=e.get("nombre_cms",     ""))
        ws.cell(row=i, column=9,  value=e.get("guid_item_lms",  ""))
        ws.cell(row=i, column=10, value=e.get("detalle",        ""))

        estado = e.get("estado", "")
        if estado == "CREADO":
            fill = _FILL_OK
        elif estado == "NOT_FOUND":
            fill = _FILL_MISS
        elif estado == "SKIP":
            fill = _FILL_SKIP
        else:
            fill = _FILL_ERR

        for col in range(1, 11):
            ws.cell(row=i, column=col).fill = fill

    anchos = {"A": 25, "B": 30, "C": 35, "D": 30, "E": 45,
              "F": 12, "G": 38, "H": 45, "I": 38, "J": 40}
    for letra, ancho in anchos.items():
        ws.column_dimensions[letra].width = ancho

    wb.save(log_path)
    print(f"  Log guardado: {log_path.resolve()}")


# ── Selección interactiva ─────────────────────────────────────────────────────

def seleccionar(titulo: str, opciones: list[str]) -> int:
    print(f"\n  {titulo}:")
    for i, op in enumerate(opciones, 1):
        print(f"    {i}. {op}")
    while True:
        try:
            idx = int(input(f"  Selección (1-{len(opciones)}) > ").strip())
            if 1 <= idx <= len(opciones):
                return idx - 1
        except (ValueError, EOFError):
            pass
        print("  Opción inválida.")


# ── Extracción de texto de HTML ───────────────────────────────────────────────

def html_a_texto(html: str) -> str:
    return re.sub(r'<[^>]+>', '', html or "").strip()


# ── Procesamiento de un producto ──────────────────────────────────────────────

def procesar_producto(session, producto_path: Path) -> dict:
    xlsx = producto_path / "OUTPUT" / "resumen_global.xlsx"
    print(f"\n{'─' * 62}")
    print(f"  Producto : {producto_path.name}")

    archivos   = leer_excel_producto(xlsx)
    estructura = agrupar_por_modulo(archivos)
    modulos    = list(estructura.keys())

    print(f"  Módulos  : {modulos}")

    # Buscar curso
    print()
    search_term = input(f"  Búsqueda del curso para '{producto_path.name}' > ").strip()
    cursos = buscar_cursos(session, search_term)

    if not cursos:
        print("  [AVISO] No se encontraron cursos. Saltando.")
        return {"producto": producto_path.name, "error": "sin curso", "entradas_log": []}

    if len(cursos) == 1:
        curso = cursos[0]
        print(f"  Curso: {curso['name']}  [{curso['guid']}]")
    else:
        idx = seleccionar(
            "Selecciona el curso",
            [f"{c['name']} — {c.get('education_year_name','?')} — {c.get('discipline_name','?')}"
             for c in cursos],
        )
        curso = cursos[idx]

    course_guid = curso["guid"]

    # Módulos ya existentes (para no duplicar)
    items_existentes   = get_course_items(session, course_guid)
    modulos_existentes = {html_a_texto(it.get("lesson_name", "")).lower()
                          for it in items_existentes}

    modulos_creados    = 0
    modulos_salteados  = 0
    secciones_creadas  = 0
    contenidos_ok      = 0
    contenidos_error   = 0
    entradas_log: list[dict] = []

    def _log(modulo, seccion, archivo, estado, content=None, item_guid=None, detalle=""):
        entradas_log.append({
            "producto":       producto_path.name,
            "modulo":         modulo,
            "seccion":        seccion,
            "erp_id":         archivo.get("erp_id", "") if archivo else "",
            "nombre_visible": archivo.get("nombre_visible", "") if archivo else "",
            "estado":         estado,
            "guid_cms":       content["guid"]  if content else "",
            "nombre_cms":     content.get("name", "") if content else "",
            "guid_item_lms":  item_guid or "",
            "detalle":        detalle,
        })

    barra = tqdm(modulos, desc=f"  {producto_path.name}", unit="módulo", ncols=80)

    for modulo in barra:
        barra.set_postfix_str(modulo[:30])

        if modulo.lower() in modulos_existentes:
            tqdm.write(f"    [SKIP] Módulo ya existe: {modulo}")
            modulos_salteados += 1
            # Registrar todos sus contenidos como SKIP en el log
            for seccion, archivos_sec in estructura[modulo].items():
                for archivo in archivos_sec:
                    _log(modulo, seccion, archivo, "SKIP", detalle="Módulo ya existía")
            continue

        # Crear módulo
        lesson_guid = crear_lesson(session, course_guid, modulo)
        if not lesson_guid:
            tqdm.write(f"    [ERROR] No se pudo crear módulo: {modulo}")
            for seccion, archivos_sec in estructura[modulo].items():
                for archivo in archivos_sec:
                    _log(modulo, seccion, archivo, "ERROR", detalle="Fallo al crear módulo")
            continue
        tqdm.write(f"    [+] Módulo: {modulo}  [{lesson_guid}]")
        modulos_creados += 1

        # Por cada sección dentro del módulo
        for seccion, archivos_sec in estructura[modulo].items():
            parent_guid = crear_seccion(session, lesson_guid, seccion)
            if not parent_guid:
                tqdm.write(f"      [ERROR] No se pudo crear sección '{seccion}'")
                for archivo in archivos_sec:
                    _log(modulo, seccion, archivo, "ERROR",
                         detalle=f"Fallo al crear sección '{seccion}'")
                continue
            tqdm.write(f"      [+] Sección: {seccion}  [{parent_guid}]")
            secciones_creadas += 1

            # Agregar contenidos a la sección
            for archivo in archivos_sec:
                content = buscar_contenido_por_erp(session, archivo["erp_id"])
                if not content:
                    tqdm.write(f"        [NOT FOUND] {archivo['erp_id']}")
                    contenidos_error += 1
                    _log(modulo, seccion, archivo, "NOT_FOUND",
                         detalle="ERP ID no encontrado en CMS")
                else:
                    item_guid = agregar_contenido(
                        session, lesson_guid, parent_guid,
                        content, archivo["nombre_visible"],
                    )
                    if item_guid:
                        tqdm.write(f"        [+] {archivo['nombre_visible']}")
                        contenidos_ok += 1
                        _log(modulo, seccion, archivo, "CREADO",
                             content=content, item_guid=item_guid)
                    else:
                        tqdm.write(f"        [ERROR] {archivo['erp_id']}")
                        contenidos_error += 1
                        _log(modulo, seccion, archivo, "ERROR",
                             content=content, detalle="PUT lesson-item falló")

                time.sleep(0.15)   # evitar rate-limit

    return {
        "producto":          producto_path.name,
        "curso":             curso["name"],
        "modulos_creados":   modulos_creados,
        "modulos_salteados": modulos_salteados,
        "secciones_creadas": secciones_creadas,
        "contenidos_ok":     contenidos_ok,
        "contenidos_error":  contenidos_error,
        "entradas_log":      entradas_log,
    }


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 62)
    print("  Creador de programa LMS - Santillana Digital")
    print("=" * 62)
    print()

    if len(sys.argv) > 1:
        base_str = sys.argv[1].strip().strip('"').strip("'")
    else:
        base_str = input(
            "Ruta base PRIM (ej: C:\\...\\SUBIDAS\\PRIM) > "
        ).strip().strip('"').strip("'")

    base_path = Path(base_str)
    if not base_path.is_dir():
        sys.exit(f"\n[ERROR] La carpeta no existe: {base_path}")

    token = input("\nToken de autorización (Bearer ...) > ").strip()
    if not token:
        sys.exit("No se ingresó token.")

    session   = build_session(token)
    productos = detectar_productos(base_path)

    if not productos:
        sys.exit("No se encontraron carpetas con OUTPUT/resumen_global.xlsx.")

    print(f"\nProductos detectados: {len(productos)}")
    for i, p in enumerate(productos, 1):
        print(f"  {i}. {p.name}")

    selec = input("\n¿Procesar todos? [s] o números separados por comas > ").strip().lower()
    if selec not in ("", "s"):
        nums = [int(x.strip()) - 1 for x in selec.split(",") if x.strip().isdigit()]
        productos = [productos[i] for i in nums if 0 <= i < len(productos)]

    resultados = [procesar_producto(session, p) for p in productos]

    # ── Log Excel ────────────────────────────────────────────────────
    todas_entradas = []
    for r in resultados:
        todas_entradas.extend(r.get("entradas_log", []))

    if todas_entradas:
        ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = base_path / f"log_programa_{ts}.xlsx"
        print()
        guardar_log(todas_entradas, log_path)

    # ── Resumen final ────────────────────────────────────────────────
    print(f"\n{'=' * 62}")
    print("  RESUMEN FINAL")
    print("=" * 62)
    for r in resultados:
        if "error" in r:
            print(f"  {r['producto']:<20} ERROR: {r['error']}")
        else:
            print(f"  {r['producto']}")
            print(f"    Curso             : {r['curso']}")
            print(f"    Módulos creados   : {r['modulos_creados']}")
            print(f"    Módulos saltados  : {r['modulos_salteados']}")
            print(f"    Secciones creadas : {r['secciones_creadas']}")
            print(f"    Contenidos OK     : {r['contenidos_ok']}")
            print(f"    Contenidos error  : {r['contenidos_error']}")
    print("=" * 62)


if __name__ == "__main__":
    main()
