import json
import re
import time
from datetime import datetime, timezone
from urllib.parse import urljoin
import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_DOMAIN = "https://www.farmaciasahumada.cl"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "es-CL,es;q=0.9",
}

def limpiar_precio(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (int, float)):
        v = int(round(valor))
        return v if v > 500 else None
    solo_digitos = re.sub(r"[^0-9]", "", str(valor))
    if not solo_digitos:
        return None
    v = int(solo_digitos)
    return v if v > 500 else None

def construir_url_sfcc(sku, url_opcional=""):
    url_opt = str(url_opcional or "").strip()
    if url_opt.startswith("http"):
        return url_opt.split("?")[0]
    sku_limpio = re.sub(r"[^0-9A-Za-z_-]", "", str(sku).strip())
    return f"{BASE_DOMAIN}/producto-{sku_limpio}.html"

def extraer_datos_ahumada(sku, url_opcional=""):
    url_objetivo = construir_url_sfcc(sku, url_opcional)
    try:
        resp = requests.get(url_objetivo, headers=HEADERS, timeout=18, allow_redirects=True)
        url_final = resp.url.split("?")[0] if resp.url else url_objetivo
        if resp.status_code != 200:
            return url_final, None, None

        soup = BeautifulSoup(resp.text, "html.parser")
        imagen_url = None
        precio_oferta = None
        precio_normal = None

        for script in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(script.string or "{}")
                items = data if isinstance(data, list) else [data]
                for item in items:
                    if item.get("@type") == "Product":
                        img = item.get("image")
                        if isinstance(img, list) and img:
                            imagen_url = img[0]
                        elif isinstance(img, str):
                            imagen_url = img

                        offers = item.get("offers", {})
                        if isinstance(offers, list) and offers:
                            offers = offers[0]
                        if isinstance(offers, dict):
                            p = limpiar_precio(offers.get("lowPrice") or offers.get("price"))
                            if p:
                                precio_oferta = p
            except Exception:
                continue

        if not imagen_url:
            img_el = (
                soup.select_one(".pdp-carousel img")
                or soup.select_one(".product-detail img.d-block")
                or soup.select_one("img[alt*='large image number 0']")
                or soup.select_one(".primary-images img")
            )
            if img_el:
                src = img_el.get("data-src") or img_el.get("src")
                if src:
                    imagen_url = urljoin(BASE_DOMAIN, src)

        if not imagen_url:
            og_img = soup.find("meta", property="og:image")
            if og_img and og_img.get("content"):
                imagen_url = urljoin(BASE_DOMAIN, og_img["content"].strip())

        precios_detectados = []
        for sel in [
            ".prices .sales .value",
            ".price .sales .value",
            ".loyalty-price .value",
            ".sales .value[content]",
            "[data-pid] .sales .value",
        ]:
            for el in soup.select(sel):
                val = limpiar_precio(el.get("content") or el.get_text())
                if val:
                    precios_detectados.append(val)

        if precios_detectados:
            precio_oferta = min(precios_detectados)

        if not precio_oferta:
            for sel in [".price .value", ".product-price .value", "[itemprop='price']"]:
                el = soup.select_one(sel)
                if el:
                    val = limpiar_precio(el.get("content") or el.get_text())
                    if val:
                        precio_normal = val
                        break

        return url_final, imagen_url, (precio_oferta or precio_normal)
    except Exception as e:
        print(f"Error SKU {sku}: {e}")
        return url_objetivo, None, None

def recoger_lista_columnas(row, columnas):
    valores = []
    for col in columnas:
        if col in row and pd.notna(row[col]):
            txt = str(row[col]).strip()
            if txt and txt.lower() != "nan":
                for parte in txt.split(","):
                    p = parte.strip()
                    if p and p not in valores:
                        valores.append(p)
    return valores

def main():
    df = pd.read_excel("matriz_rutas_farmacia.xlsx", sheet_name="Matriz_Productos")
    df = df[df["Activo"].astype(str).str.upper().str.strip() == "SI"]

    productos = []
    for _, row in df.iterrows():
        sku = str(row.get("SKU", "")).strip()
        if sku.endswith(".0"):
            sku = sku[:-2]
        if not sku or sku.lower() == "nan":
            continue

        url_opt = str(row.get("URL_Opcional", "") or "").strip()
        if url_opt.lower() == "nan":
            url_opt = ""
        precio_respaldo = limpiar_precio(row.get("Precio_Respaldo_CLP")) or 0

        pieles = recoger_lista_columnas(row, ["Piel_1 (Principal)", "Piel_2 (Opcional)", "Piel_3 (Opcional)", "Pieles_Compatibles (Auto)", "Pieles_Compatibles"])
        necesidades = recoger_lista_columnas(row, ["Necesidad_1 (Principal)", "Necesidad_2 (Opcional)", "Necesidad_3 (Opcional)", "Necesidad_4 (Opcional)", "Necesidades_Compatibles (Auto)", "Necesidades_Compatibles"])

        url_canonica, img_live, precio_live = extraer_datos_ahumada(sku, url_opt)
        time.sleep(0.9)

        precio_final = precio_live if precio_live else precio_respaldo
        fuente = "sfcc_en_vivo" if (img_live or precio_live) else "respaldo_excel"

        productos.append({
            "sku": sku,
            "modulo": str(row.get("Modulo", "Dermocosmetica")).strip(),
            "paso": str(row.get("Paso_Rutina", "")).strip(),
            "presupuesto": str(row.get("Presupuesto", "Medio")).strip(),
            "marca": str(row.get("Marca", "")).strip(),
            "nombre": str(row.get("Nombre_Producto", "")).strip(),
            "url": url_canonica,
            "pieles": pieles,
            "necesidades": necesidades,
            "textura": str(row.get("Textura", "Ligero")).strip(),
            "beneficio": str(row.get("Para_Que_Sirve", "")).strip(),
            "cuando": str(row.get("Cuando_Usar", "Mañana y noche")).strip(),
            "precio": precio_final,
            "imagen": img_live or "",
            "fuente": fuente
        })
        print(f"[{fuente}] SKU {sku} -> ${precio_final} | pieles={pieles}")

    salida = {
        "ultima_actualizacion": datetime.now(timezone.utc).astimezone().isoformat(),
        "total_productos": len(productos),
        "productos": productos
    }
    with open("catalogo.json", "w", encoding="utf-8") as f:
        json.dump(salida, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    main()
