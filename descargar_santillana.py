#!/usr/bin/env python3
"""
Descargador de documentos de Santillana Digital (SantiVaContigoDocs).
Navega 3 niveles: productos -> carpetas -> archivos PDF.
"""

import os
import re
import sys
import time
import getpass
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs

try:
    import requests
    from bs4 import BeautifulSoup
    from tqdm import tqdm
except ImportError:
    print("Instalando dependencias necesarias...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "beautifulsoup4", "tqdm"])
    import requests
    from bs4 import BeautifulSoup
    from tqdm import tqdm


BASE_URL = "https://digital.santillana.com.pe"
DOCS_URL = f"{BASE_URL}/Documentos/SantiVaContigoDocs"
OUTPUT_DIR = Path("SantillanaDigital_Docs")


def sanitize_name(name: str) -> str:
    """Elimina caracteres no válidos para nombres de carpeta/archivo."""
    name = name.strip()
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    name = re.sub(r'\s+', ' ', name)
    return name or "sin_nombre"


def build_session(cookie_str: str) -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Referer": BASE_URL,
    })
    # Parsear cookies del string pegado (formato: key=value; key2=value2)
    for part in cookie_str.split(";"):
        part = part.strip()
        if "=" in part:
            key, _, value = part.partition("=")
            session.cookies.set(key.strip(), value.strip(), domain="digital.santillana.com.pe")
    return session


def get_soup(session: requests.Session, url: str) -> BeautifulSoup | None:
    """Realiza GET y devuelve BeautifulSoup, con reintento en caso de error."""
    for intento in range(3):
        try:
            resp = session.get(url, timeout=30)
            resp.raise_for_status()
            return BeautifulSoup(resp.text, "html.parser")
        except requests.RequestException as e:
            if intento < 2:
                time.sleep(2 ** intento)
            else:
                print(f"\n  [ERROR] No se pudo obtener {url}: {e}")
                return None


def obtener_productos(session: requests.Session) -> list[dict]:
    """Nivel 1: lista todos los productos desde SantiVaContigoDocs."""
    print(f"\n[1/3] Obteniendo lista de productos desde {DOCS_URL} ...")
    soup = get_soup(session, DOCS_URL)
    if not soup:
        return []

    productos = []
    for a in soup.select("a[href*='/Documentos/Contenido']"):
        href = a["href"]
        nombre_tag = a.find_next("p", class_="titulo-contenido")
        nombre = nombre_tag.get_text(strip=True) if nombre_tag else href
        url_completa = urljoin(BASE_URL, href)
        productos.append({"nombre": nombre, "url": url_completa})

    print(f"  Encontrados {len(productos)} productos.")
    return productos


def obtener_carpetas(session: requests.Session, producto: dict) -> list[dict]:
    """Nivel 2: carpetas/secciones dentro de un producto."""
    soup = get_soup(session, producto["url"])
    if not soup:
        return []

    carpetas = []
    for a in soup.select("a[href*='/Documentos/Archivo']"):
        href = a["href"]
        nombre_tag = a.find_next("p", class_="titulo-contenido")
        nombre = nombre_tag.get_text(strip=True) if nombre_tag else href
        url_completa = urljoin(BASE_URL, href)
        carpetas.append({"nombre": nombre, "url": url_completa})
    return carpetas


def obtener_archivos(session: requests.Session, carpeta: dict) -> list[dict]:
    """Nivel 3: links directos a los archivos (PDF, etc.)."""
    soup = get_soup(session, carpeta["url"])
    if not soup:
        return []

    archivos = []
    # Los archivos reales apuntan a /api/accesoblob/... o similares
    for a in soup.select("a[href]"):
        href = a["href"]
        if "/api/accesoblob/" in href or href.lower().endswith((".pdf", ".docx", ".xlsx", ".pptx", ".zip")):
            nombre_tag = a.find_next("p", class_="titulo-contenido")
            nombre = nombre_tag.get_text(strip=True) if nombre_tag else ""
            # Deducir extensión desde la URL si no hay nombre claro
            parsed = urlparse(href)
            ext = Path(parsed.path).suffix or ".pdf"
            if not nombre:
                nombre = Path(parsed.path).stem
            nombre_archivo = sanitize_name(nombre) + ext
            archivos.append({"nombre": nombre_archivo, "url": href})
    return archivos


def descargar_archivo(session: requests.Session, url: str, destino: Path) -> bool:
    """Descarga un archivo con barra de progreso interna."""
    if destino.exists():
        return True  # ya descargado, saltar

    destino.parent.mkdir(parents=True, exist_ok=True)
    tmp = destino.with_suffix(".tmp")

    try:
        with session.get(url, stream=True, timeout=60) as resp:
            resp.raise_for_status()
            total = int(resp.headers.get("content-length", 0))
            with open(tmp, "wb") as f, tqdm(
                total=total,
                unit="B",
                unit_scale=True,
                unit_divisor=1024,
                desc=f"    {destino.name[:50]}",
                leave=False,
                ncols=90,
            ) as bar:
                for chunk in resp.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        bar.update(len(chunk))
        tmp.rename(destino)
        return True
    except requests.RequestException as e:
        print(f"\n  [ERROR] Descarga fallida {url}: {e}")
        if tmp.exists():
            tmp.unlink()
        return False


def main():
    print("=" * 60)
    print("  Descargador de Santillana Digital - SantiVaContigoDocs")
    print("=" * 60)
    print("\nPega tu cookie de sesión (se ocultará al escribir).")
    print("Puedes obtenerla desde DevTools > Application > Cookies.")
    print("Formato: .AspNet.ApplicationCookie=...; otra=valor\n")

    cookie_str = getpass.getpass("Cookie: ")
    if not cookie_str.strip():
        print("No se ingresó ninguna cookie. Saliendo.")
        sys.exit(1)

    session = build_session(cookie_str)

    # ── Nivel 1: productos ──────────────────────────────────────────
    productos = obtener_productos(session)
    if not productos:
        print("No se encontraron productos. Verifica la cookie o la URL.")
        sys.exit(1)

    total_archivos_descargados = 0
    total_archivos_fallidos = 0

    # ── Barra de progreso de productos ─────────────────────────────
    print(f"\n[2/3] Procesando productos ...\n")
    barra_productos = tqdm(productos, desc="Productos", unit="prod", ncols=90)

    for producto in barra_productos:
        nombre_producto = sanitize_name(producto["nombre"])
        barra_productos.set_postfix_str(nombre_producto[:40])
        dir_producto = OUTPUT_DIR / nombre_producto

        # ── Nivel 2: carpetas ────────────────────────────────────────
        carpetas = obtener_carpetas(session, producto)
        if not carpetas:
            continue

        barra_carpetas = tqdm(carpetas, desc=f"  Carpetas", unit="carp", leave=False, ncols=90)
        for carpeta in barra_carpetas:
            nombre_carpeta = sanitize_name(carpeta["nombre"])
            barra_carpetas.set_postfix_str(nombre_carpeta[:35])
            dir_carpeta = dir_producto / nombre_carpeta

            # ── Nivel 3: archivos ────────────────────────────────────
            archivos = obtener_archivos(session, carpeta)
            if not archivos:
                continue

            for archivo in archivos:
                destino = dir_carpeta / archivo["nombre"]
                ok = descargar_archivo(session, archivo["url"], destino)
                if ok:
                    total_archivos_descargados += 1
                else:
                    total_archivos_fallidos += 1

    # ── Resumen final ───────────────────────────────────────────────
    print(f"\n{'=' * 60}")
    print(f"  Descarga completada.")
    print(f"  Archivos descargados : {total_archivos_descargados}")
    print(f"  Errores              : {total_archivos_fallidos}")
    print(f"  Guardados en         : {OUTPUT_DIR.resolve()}")
    print("=" * 60)


if __name__ == "__main__":
    main()
