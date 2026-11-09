# -*- coding: utf-8 -*-
"""
CARGADOR DE MENSAJES DE WHATSAPP  ->  TIPIFICACIONES

Reemplaza al formulario de Kobo: el departamento pega el mensaje del freelance en la
página /cargar, el sistema lo convierte en filas (con la misma lógica de emparejamiento
que usa el reporte) y las guarda en la hoja de cálculo.

Se conecta a server.py con UNA línea al final del archivo (ver INSTALACION al pie):
    from cargador import init_cargador
    init_cargador(app, globals())

Variables de entorno (Render -> Environment):
    GOOGLE_SERVICE_ACCOUNT_JSON  contenido completo del JSON de la cuenta de servicio
    SPREADSHEET_ID               el ID de la hoja BD TEST (está en su URL)
Sin ellas la página funciona en modo "solo vista previa" (no guarda).
"""
import os
import re
import json
import time
import threading
import unicodedata
import logging
from datetime import datetime

from flask import Blueprint, jsonify, request, redirect, url_for

log = logging.getLogger(__name__)

HOJA_TIPIFICACIONES = os.environ.get('HOJA_TIPIFICACIONES', 'TIPIFICACIONES')
HOJA_MANUAL = os.environ.get('HOJA_MANUAL', 'MANUAL_WHATSAPP')
# Mientras el importador de Kobo siga activo, borra y reescribe TIPIFICACIONES con
# (Kobo + MANUAL_WHATSAPP). Por eso se guarda en MANUAL_WHATSAPP (permanente) y, además,
# directo en TIPIFICACIONES (para verlo al instante). Cuando Kobo se apague, poner
# ESCRIBIR_MANUAL=0 en Render y se guarda solo en TIPIFICACIONES.
ESCRIBIR_MANUAL = os.environ.get('ESCRIBIR_MANUAL', '1') != '0'

ENCABEZADOS = ["ID_Envio", "Fecha", "Horario", "Empresa", "Codigo_Tienda", "Sucursal",
               "Sedes Damasco", "Sede Daka", "Producto", "Marca", "Cantidad", "Precio",
               "Factura", "Visitas"]

# ---------------------------------------------------------------------------
# TIENDAS (código -> (empresa, nombre de sucursal tal como va en TIPIFICACIONES))
# ---------------------------------------------------------------------------
def _mk():
    t = {}
    daka = ["CHACAO", "EL RECREO", "EL MARQUES", "SAN MARTIN", "CANDELARIA", "BOLEITA", "LAS MERCEDES",
            "LA TRINIDAD", "CARRIZAL", "SAN DIEGO", "NAGUANAGUA", "CENTRO VALENCIA", "GUACARA", "AV BRANGER",
            "PUERTO CABELLO", "MARACAY REDOMA DEL AVION", "MARACAY CENTRO", "SAN FELIPE",
            "BARQUISIMETO AV VENEZUELA", "BARQUISIMETO CENTRO", "PUNTO FIJO", "LA LIMPIA", "MARACAIBO",
            "PUERTO ORDAZ", "EL TIGRE", "PTO LA CRUZ", "LECHERIA", "ACARIGUA", "VALLE DE LA PASCUA",
            "MARGARITA", "MATURIN", "VALERA", "SAN CRISTOBAL", "CHARALLAVE", "AV 5 DE JULIO", "CAGUA"]
    for i, n in enumerate(daka, 1):
        t[f"BW-{i}"] = ("ddaka", "DAKA " + n)
    mm = ["LOS SIMBOLOS", "LOS CORTIJOS", "MULTIMAX CANDELARIA", "PLAZA VENEZUELA", "EL HATILLO", "EXPRESS",
          "GUATIRE", "CHARALLAVE", "CARRIZAL", "EXPRESS CHACAO", "MAÑONGO", "GUACARA", "MORON", "AV BOLIVAR",
          "MARACAY", "CAGUA", "AVIADORES", "LA VICTORIA", "CUMANA", "LECHERIA", "PUERTO LA CRUZ", "LA LIMPIA",
          "SAMBIL MARACAIBO", "MARACAIBO 5 DE JULIO", "CABIMAS", "MARGARITA CC TRAKI", "MARGARITA PAMPATAR",
          "PTO ORDAZ", "SAN FELIX", "CIUDAD BOLIVAR", "MATURIN", "SAN FERNANDO DE APURE", "PUNTO FIJO", "COCO",
          "BARINAS", "BARQUISIMETO CC SAMBIL", "BARQUISIMETO AV LARA", "CABUDARE", "SAN CRISTOBAL", "GUANARE",
          "SAN FELIPE", "VALERA", "VIGIA", "MERIDA", "CC GLOBAL MARACAY"]
    for i, n in enumerate(mm, 1):
        t[f"CW-{i}"] = ("mmultimax", "MULTIMAX " + n)
    dm = ["TRINIDAD", "C.C.C.T", "CANDELARIA", "NUEVA GRANADA", "LAS MERCEDES", "CATIA CALLE COLOMBIA",
          "CATIA MUEBLERIA", "CATIA ANTIGUA", "SAN MARTIN PLAZA CAPUCHINO", "SAN MARTIN ESQUINA ANGELITOS",
          "SABANA GRANDE", "PARAISO", "CALIFORNIA", "AV BOLIVAR VALENCIA", "VALENCIA CENTRO", "LOS TEQUES",
          "GUATIRE", "LA GUAIRA", "CAGUA", "MARACAY AV BOLIVAR", "BARQUISIMETO AV VENEZUELA",
          "BARQUISIMETO CC RECREO", "LECHERIA", "PUERTO LA CRUZ", "PUERTO ORDAZ", "VALERA", "MATURIN",
          "SAN CRISTOBAL", "MARACAIBO", "SAN FELIPE", "CENTRO MARACAY", "SAN MARTIN", "CHARALLAVE", "BARINAS",
          "ACARIGUA", "MERIDA", "LA LIMPIA", "CARRIZAL", "VALLE DE LA PASCUA", "CUMANA", "CAGUA", "SAN DIEGO",
          "MARGARITA", "CUA", "SAMBIL MARGARITA", "TORRE DAMASCO", "TUCUPITA", "CORO", "PUNTO FIJO", "APURE",
          "LAS DELICIAS MARACAIBO", "EL TIGRE", "PALO NEGRO", "CABUDARE"]
    for i, n in enumerate(dm, 1):
        t[f"AW-{i}"] = ("ddamasco", n + " - DAMASCO")
    t["AW-61"] = ("ddamasco", "PLAZA DE TORO - DAMASCO")
    return t


TIENDAS = _mk()
GID_POR_EMPRESA = {}      # se llena en init_cargador con los GID de server.py


def _nombre_corto(emp, suc):
    return suc.replace(" - DAMASCO", "") if emp == "ddamasco" else suc


# ---------------------------------------------------------------------------
# PARSER DEL MENSAJE (misma lógica que la herramienta de Apps Script)
# ---------------------------------------------------------------------------
MARCAS = ["HYUNDAI", "DAEWOO", "BREMEN", "SANSUI", "RIBELLE", "DA+CO", "DACO", "SYON", "MIDEA", "CLX", "CONDESA",
          "FRIGILUX", "MILATTI", "KUCCE", "VETRUX", "ITOK", "MUNDO BLANCO", "XIAOMI", "YALA", "BISTRO", "OSLO",
          "COMPAC", "COMPAQ", "HP", "LG", "SAMSUNG", "OSTER", "CUISINART", "HAMILTON BEACH", "BLACK DECKER",
          "BLACK+DECKER", "PROCTOR SILEX", "PHILIPS", "WHIRLPOOL", "ELECTROLUX", "MABE", "JBL", "SONY", "TCL",
          "HISENSE", "LENOVO", "DELL", "ASUS", "ACER", "APPLE", "MOTOROLA", "INFINIX", "TECNO", "ZTE",
          "UNIVERSAL", "MIRAGE", "ARTCOOL", "BETTY CROCKER", "DAMASCO", "ROYAL", "RCA", "NINJA", "KITCHENAID"]
_MARCAS_ORD = sorted(set(MARCAS), key=lambda m: -len(m))


def _sin_tildes(s):
    return re.sub(r'[̀-ͯ]', '', unicodedata.normalize('NFD', s))


def _titulo(s):
    return re.sub(r'(^|[\s/(+\-])([a-zñ])', lambda m: m.group(1) + m.group(2).upper(), s.lower())


def horario_wa(txt):
    """'(10:00 am a 11:00 am)' -> '10_00am_11_00am' (mismo formato que Kobo)."""
    t = re.sub(r'^\(?[\s.]+', '(', txt)
    m = re.search(r'\(?\s*(\d{1,2})(?::\d{2})?\s*(am|pm|a\.m\.|p\.m\.)?\s*(?:a|-|–|hasta)\s*(\d{1,2})(?::\d{2})?\s*(am|pm|a\.m\.|p\.m\.)\s*\)?', t, re.I)
    if not m:
        return ""
    h1, h2 = int(m.group(1)), int(m.group(3))
    s2 = m.group(4).replace('.', '').lower()
    s1 = m.group(2).replace('.', '').lower() if m.group(2) else s2
    if not m.group(2) and h1 > h2 and h1 != 12:
        s1 = "am" if s2 == "pm" else s2

    def f(h, s):
        return (("0" + str(h)) if (h < 10 and s == "am") else str(h)) + "_00" + s
    return f(h1, s1) + "_" + f(h2, s2)


def normalizar_producto(desc):
    t = " " + re.sub(r'\s+', ' ', _sin_tildes(desc.strip())) + " "
    t = re.sub(r'\bcompac\b', 'Compaq', t, flags=re.I)
    t = re.sub(r'\bsanduchera\b', 'Sandwichera', t, flags=re.I)
    t = re.sub(r'\bhydrojet\b|\bhidrojet\b', 'Hidrojet', t, flags=re.I)
    t = re.sub(r'\bmini\s+ups\b', 'UPS', t, flags=re.I)
    t = re.sub(r'\bcorneta\s+altavoz\b|\baltavoz\b', 'Corneta', t, flags=re.I)
    marca = ""
    upper = t.upper()
    for m in _MARCAS_ORD:
        pat = r'(^|[^A-Z0-9])' + re.escape(m) + r'(?![A-Z0-9])'
        if re.search(pat, upper):
            marca = m
            t = re.sub(r'(^|[^A-Za-z0-9])' + re.escape(m) + r'(?![A-Za-z0-9])', r'\1 ', t, count=1, flags=re.I)
            break
    t = re.sub(r'(\d)\.\s*(?=mil\b)', r'\1 ', t, flags=re.I)
    t = re.sub(r'\b(\d{1,3})\.(\d{3})\b', r'\1\2', t)
    t = re.sub(r'\b(\d+(?:[.,]\d+)?)\s*mil\s*btu\b',
               lambda m: str(int(round(float(m.group(1).replace(',', '.')) * 1000))) + " BTU", t, flags=re.I)
    t = re.sub(r'\b(\d{4,6})\s*btu\b', r'\1 BTU', t, flags=re.I)
    t = re.sub(r'\b(\d+)\s*(?:pies?|ft|feet)\b', r'\1 Pies', t, flags=re.I)
    t = re.sub(r'\b(\d+)\s*(?:l|lt|lts|litros?)\b', r'\1 Litros', t, flags=re.I)
    t = re.sub(r'\b(\d+)\s*(?:k|kg|kgs|kilos?)\b', r'\1 Kg', t, flags=re.I)
    t = re.sub(r'\b(\d+)\s*(?:w|watts?)\b', r'\1W', t, flags=re.I)
    t = re.sub(r'\b(\d+)\s*(?:tsz|tz|tazas?)\b', r'\1 Tazas', t, flags=re.I)
    t = re.sub(r'\b(\d+)\s*piezas?\b', r'\1 Piezas', t, flags=re.I)
    t = re.sub(r'\b(\d+)\s*(?:"|\'\'|”|“|pulgadas?|pulg|plgs?)', r'\1 Pulgadas', t, flags=re.I)
    t = re.sub(r'\bdoble\s+tina\b', 'Doble Tina', t, flags=re.I)
    t = re.sub(r'\baa\b|\ba\s*/\s*a\b|\ba\s+a\b|\ba\.a\.?\b|\bac\b|\ba/c\b', 'Aire Acondicionado', t, flags=re.I)
    if re.search(r'\bsplit\b', t, re.I) and re.search(r'\bbtu\b', t, re.I) and not re.search(r'aire acondicionado', t, re.I):
        t = "Aire Acondicionado " + t
    if re.search(r'aire acondicionado', t, re.I):
        t = re.sub(r'\b(\d{4,5})\b(?!\s*(?:btu|w\b))', r'\1 BTU', t, count=1, flags=re.I)
    t = re.sub(r'\btv\b', 'Televisor', t, flags=re.I)
    if re.search(r'pulgadas', t, re.I) and re.search(r'\bsmart\b', t, re.I) and not re.search(r'televisor', t, re.I):
        t = "Televisor " + t
    t = re.sub(r'\bnevera(s)?\b', lambda m: 'Refrigerador' + (m.group(1) or ''), t, flags=re.I)
    t = re.sub(r'\btostiarepas\b', 'Tostiarepas', t, flags=re.I)
    t = re.sub(r'\s+', ' ', t).strip()
    nombre = _titulo(t)
    nombre = re.sub(r'\bBtu\b', 'BTU', nombre)
    nombre = re.sub(r'\bUps\b', 'UPS', nombre)
    nombre = re.sub(r'\bTv\b', 'TV', nombre)
    nombre = re.sub(r'(\d)w\b', lambda m: m.group(1) + 'W', nombre, flags=re.I)
    nombre = re.sub(r'\bDe De\b', 'De', nombre)
    nombre = re.sub(r'\bKg\.$', 'Kg', nombre)
    nombre = re.sub(r'\.$', '', nombre)
    return {"producto": nombre, "marca": (_titulo(marca) if marca else "")}


_UNIDADES = re.compile(r'^(?:pulg|pulgadas?|"|\'\'|btu|mil|kg|kgs|k\b|l\b|lt|lts|litros?|w\b|watts?|pies?|ft|tazas?|tz|piezas?|pzas?|v\b|hz)', re.I)
_SIN_VENTAS = re.compile(r'sin\s+ventas?|no\s+se\s+observ|sin\s+productos?|no\s+hubo\s+ventas?', re.I)


def _es_linea_horario(l):
    return bool(re.search(r'\b(am|pm)\b', l, re.I) and re.match(r'^\(?[\s.]*\d', l) and not re.match(r'^\d+\s*[-.)]', l))


def parsear_mensajes(msg):
    """Devuelve una lista de bloques, uno por horario:
    {horario, visitas, prod_decl, fact_decl, facturas:[{nro, productos:[{cant, texto, producto, marca}]}],
     sin_ventas, sin_cant, avisos}"""
    lineas = [l.strip() for l in re.split(r'\r?\n', msg or '') if l.strip()]
    bloques, b = [], None

    def nuevo(h):
        return {"horario": h, "visitas": "", "prod_decl": None, "fact_decl": None, "facturas": [],
                "sin_ventas": False, "sin_cant": 0, "avisos": []}
    for l in lineas:
        if re.match(r'^EDM\b', l, re.I):
            continue
        if _es_linea_horario(l):
            h = horario_wa(l)
            if h:
                b = nuevo(h)
                bloques.append(b)
                continue
        if b is None:
            b = nuevo("")
            b["avisos"].append("Hay texto antes de la primera hora: " + l[:40])
            bloques.append(b)
        m = re.match(r'^visitas?\s*:?\s*(\d+)', l, re.I)
        if m:
            b["visitas"] = int(m.group(1)); continue
        m = re.match(r'^productos?\s*:?\s*(\d+)', l, re.I)
        if m:
            b["prod_decl"] = int(m.group(1)); continue
        m = re.match(r'^facturas?\s*:?\s*(\d+)', l, re.I)
        if m:
            b["fact_decl"] = int(m.group(1)); continue
        m = re.match(r'^(\d+)\s*[-.)]\s*(.+)$', l)
        mb = None if m else re.match(r'^[-•*–—]\s*(.+)$', l)
        if m or mb:
            nro = int(m.group(1)) if m else len(b["facturas"]) + 1
            cuerpo = m.group(2) if m else mb.group(1)
            items = [x.strip() for x in re.split(r'\s*\+\s+|\s+\+\s*', cuerpo) if x.strip()]
            prods = []
            for it in items:
                q, d = 1, it
                mq = re.match(r'^(\d{1,2})\s*([^\d\s.,].*)$', it)
                if mq and not _UNIDADES.match(mq.group(2)):
                    q, d = int(mq.group(1)), mq.group(2)
                else:
                    b["sin_cant"] += 1
                n = normalizar_producto(d)
                prods.append({"cant": q, "texto": it, "producto": n["producto"], "marca": n["marca"]})
            b["facturas"].append({"nro": nro, "productos": prods})
            continue
        if _SIN_VENTAS.search(l):
            b["sin_ventas"] = True
    for b in bloques:
        nf = len(b["facturas"])
        npd = sum(len(f["productos"]) for f in b["facturas"])
        un = sum(p["cant"] for f in b["facturas"] for p in f["productos"])
        if not b["horario"]:
            b["avisos"].append("No pude leer el horario (ejemplo: 10:00 am a 11:00 am).")
        if b["visitas"] == "":
            b["avisos"].append("Falta el número de visitas.")
        if nf == 0 and not b["sin_ventas"]:
            b["avisos"].append("No encontré productos en esta hora.")
        if b["fact_decl"] is not None and b["fact_decl"] != nf:
            b["avisos"].append(f"El mensaje dice {b['fact_decl']} facturas y conté {nf}.")
        if b["prod_decl"] is not None and b["prod_decl"] not in (npd, un):
            b["avisos"].append(f"El mensaje dice {b['prod_decl']} productos y conté {npd} ({un} unidades).")
    return bloques


def horario_de_hora(h24):
    """8 -> '08_00am_09_00am' (hora de inicio en 24 h)."""
    h24 = int(h24)
    def p(x):
        x %= 24
        s = "am" if x < 12 else "pm"
        y = x % 12 or 12
        return f"{y}:00 {s}"
    return horario_wa(f"{p(h24)} a {p(h24 + 1)}")


def _aplicar_manuales(bloques, hora, visitas, avisos_gen):
    """Horario y visitas elegidos a mano en la página. Con UN solo bloque mandan sobre lo
    leído del mensaje; con varios, solo completan los bloques a los que les falta el dato."""
    try:
        hor = horario_de_hora(hora) if hora not in (None, '') else ''
    except Exception:
        hor = ''
    try:
        vis = int(visitas) if visitas not in (None, '') else None
        if vis is not None and vis < 0:
            vis = None
    except Exception:
        vis = None
    if not hor and vis is None:
        return
    unico = len(bloques) == 1
    for b in bloques:
        if hor and (unico or not b['horario']):
            if b['horario'] and b['horario'] != hor:
                b['avisos'].append("El mensaje decía otro horario; se usó el que elegiste arriba.")
            b['horario'] = hor
            b['avisos'] = [x for x in b['avisos'] if not x.startswith(("No pude leer el horario", "Hay texto antes de la primera hora"))]
        if vis is not None and (unico or b['visitas'] == ""):
            if b['visitas'] != "" and b['visitas'] != vis:
                b['avisos'].append("El mensaje decía otras visitas; se usaron las que escribiste arriba.")
            b['visitas'] = vis
            b['avisos'] = [x for x in b['avisos'] if not x.startswith("Falta el número de visitas")]
    if not unico:
        avisos_gen.append("Elegiste horario o visitas, pero el mensaje trae varias horas: solo se aplicó a las que no traían ese dato.")


# ---------------------------------------------------------------------------
# GOOGLE SHEETS (cuenta de servicio)
# ---------------------------------------------------------------------------
_sheets_lock = threading.Lock()
_sesion = {"s": None, "t": 0}


def sheets_configurado():
    return bool(os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON', '').strip() and os.environ.get('SPREADSHEET_ID', '').strip())


def _sesion_google():
    with _sheets_lock:
        if _sesion["s"] is None:
            from google.oauth2 import service_account
            from google.auth.transport.requests import AuthorizedSession
            info = json.loads(os.environ['GOOGLE_SERVICE_ACCOUNT_JSON'])
            cred = service_account.Credentials.from_service_account_info(
                info, scopes=['https://www.googleapis.com/auth/spreadsheets'])
            _sesion["s"] = AuthorizedSession(cred)
        return _sesion["s"]


def _api(metodo, ruta, **kw):
    base = f"https://sheets.googleapis.com/v4/spreadsheets/{os.environ['SPREADSHEET_ID'].strip()}"
    r = _sesion_google().request(metodo, base + ruta, timeout=40, **kw)
    if r.status_code >= 300:
        raise RuntimeError(f"Google Sheets respondió {r.status_code}: {r.text[:300]}")
    return r.json() if r.text else {}


def _leer_hoja(nombre, rango="A:N"):
    from urllib.parse import quote
    j = _api('GET', f"/values/{quote(nombre + '!' + rango)}?valueRenderOption=FORMATTED_VALUE")
    return j.get('values', [])


def _ids_hojas():
    j = _api('GET', "?fields=sheets.properties(sheetId,title)")
    return {s['properties']['title']: s['properties']['sheetId'] for s in j.get('sheets', [])}


def _asegurar_hoja(nombre):
    ids = _ids_hojas()
    if nombre in ids:
        return ids[nombre]
    _api('POST', ":batchUpdate", json={"requests": [{"addSheet": {"properties": {"title": nombre}}}]})
    from urllib.parse import quote
    _api('PUT', f"/values/{quote(nombre + '!A1')}?valueInputOption=RAW", json={"values": [ENCABEZADOS]})
    return _ids_hojas()[nombre]


def _borrar_envios(nombre_hoja, ids_envio):
    """Borra las filas cuyo ID_Envio esté en ids_envio. Devuelve cuántas borró."""
    filas = _leer_hoja(nombre_hoja, "A:A")
    sid = _ids_hojas().get(nombre_hoja)
    if sid is None:
        return 0
    borrar = [i for i, f in enumerate(filas) if i > 0 and f and f[0] in ids_envio]
    if not borrar:
        return 0
    rangos, ini, prev = [], borrar[0], borrar[0]
    for i in borrar[1:]:
        if i == prev + 1:
            prev = i
        else:
            rangos.append((ini, prev)); ini = prev = i
    rangos.append((ini, prev))
    reqs = [{"deleteDimension": {"range": {"sheetId": sid, "dimension": "ROWS", "startIndex": a, "endIndex": b + 1}}}
            for a, b in reversed(rangos)]
    _api('POST', ":batchUpdate", json={"requests": reqs})
    return len(borrar)


def _agregar_filas(nombre_hoja, filas):
    from urllib.parse import quote
    _api('POST', f"/values/{quote(nombre_hoja + '!A:N')}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
         json={"values": filas})


# ---------------------------------------------------------------------------
# INIT: registra las rutas usando las funciones que ya existen en server.py
# ---------------------------------------------------------------------------
def init_cargador(app, G):
    """G = globals() de server.py (para reutilizar el emparejamiento y las bases de precios)."""
    import pandas as pd
    from rapidfuzz import process as rf_process, fuzz

    GID_POR_EMPRESA.update({"ddaka": G['GID_BASE_DAKA'], "ddamasco": G['GID_BASE_DAMASCO'],
                            "mmultimax": G['GID_BASE_MULTIMAX']})
    bp = Blueprint('cargador', __name__)

    def logueado():
        return G['is_logged_in']()

    def _maestro(emp):
        """(df_maestro, prices_numeric, mapeo_precios, mapeo_marcas) igual que filtrar_datos_reporte."""
        _, dfm = G['fetch_data'](gid_base=GID_POR_EMPRESA[emp])
        prices = dfm['PRECIO DE REFERENCIA'].astype(str).str.replace(' ', '').str.replace(',', '.')
        pn = pd.to_numeric(prices, errors='coerce').fillna(0.0)
        return (dfm, pn, dict(zip(dfm['PRODUCTO'].astype(str), pn)),
                dict(zip(dfm['PRODUCTO'].astype(str), dfm['MARCA'])))

    def _fecha_ddmmyyyy(f):
        f = (f or '').strip()
        if re.match(r'^\d{4}-\d{2}-\d{2}$', f):
            y, m, d = f.split('-'); return f"{d}/{m}/{y}"
        if re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', f):
            d, m, y = f.split('/'); return f"{int(d):02d}/{int(m):02d}/{y}"
        return ''

    @bp.route('/cargar')
    def pagina():
        if not logueado():
            return redirect(url_for('login_page'))
        from flask import send_file
        return send_file(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'cargar.html'), mimetype='text/html')

    @bp.route('/idm_logo.png')
    def logo():
        from flask import send_file
        return send_file(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'idm_logo.png'), mimetype='image/png')

    @bp.route('/api/cargar/config')
    def config():
        if not logueado():
            return jsonify({'success': False, 'error': 'No autorizado'}), 401
        tiendas = [{'codigo': c, 'empresa': e, 'nombre': _nombre_corto(e, s)} for c, (e, s) in TIENDAS.items()]
        return jsonify({'success': True, 'tiendas': tiendas, 'guarda': sheets_configurado(),
                        'ia': bool(G['_gemini_clave']())})

    def _a_fecha(txt):
        t = str(txt or '').strip()
        m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', t)
        try:
            if m:
                return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).date()
            m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})', t)
            if m:
                return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))).date()
        except ValueError:
            pass
        return None

    @bp.route('/api/cargar/recientes')
    def recientes():
        """Todas las horas cargadas de los últimos `dias` días (por defecto 10), de la más nueva a la más vieja.
        Incluye las que vinieron de Kobo y las cargadas desde esta página."""
        if not logueado():
            return jsonify({'success': False, 'error': 'No autorizado'}), 401
        try:
            try:
                dias = max(1, min(60, int(request.args.get('dias', '10'))))
            except ValueError:
                dias = 10
            if sheets_configurado():
                v = _leer_hoja(HOJA_TIPIFICACIONES, "A:N")
                filas = [list(r) + [''] * (14 - len(r)) for r in v[1:]]
            else:
                dfc, _ = G['fetch_data'](gid_base=G['GID_BASE_DAKA'])
                filas = dfc.fillna('').astype(str).values.tolist()
            por_nombre = {s: c for c, (e, s) in TIENDAS.items()}
            grupos, orden = {}, []
            for r in filas:
                if not r[0]:
                    continue
                g = grupos.get(r[0])
                if g is None:
                    cod = re.sub(r'^([A-Z]{2})[\s*_.\-]*0*(\d+)$', r'\1-\2', str(r[4]).strip().upper())   # AW*35 -> AW-35
                    if cod not in TIENDAS:
                        cod = por_nombre.get(r[5]) or str(r[4])
                    g = grupos[r[0]] = {'id': r[0], 'fecha': str(r[1])[:10], 'horario': r[2], 'empresa': r[3],
                                        'codigo': cod, 'tienda': r[5], 'prods': 0, 'visitas': r[13]}
                    orden.append(r[0])
                if str(r[8]).strip().lower() not in ('sin productos', ''):
                    g['prods'] += 1
            todas = [grupos[i] for i in orden]
            hoy = datetime.now().date()
            fechas = [d for d in (_a_fecha(g['fecha']) for g in todas) if d]
            # la ventana se mide hacia atrás desde la fecha más reciente cargada (o hoy, lo que sea mayor)
            tope = max([hoy] + fechas) if fechas else hoy
            sel = []
            for i, g in enumerate(todas):
                d = _a_fecha(g['fecha'])
                if d is None or (tope - d).days < dias or i >= len(todas) - 60:
                    sel.append(g)
            out = sel[-5000:]
            out.reverse()          # lo último que entró primero
            return jsonify({'success': True, 'cargas': out, 'dias': dias})
        except Exception as e:
            app.logger.error("recientes: %s", e, exc_info=True)
            return jsonify({'success': False, 'error': str(e)}), 500

    def _analizar_bloques(empresa, bloques):
        """Empareja todos los productos de todos los bloques en UNA sola pasada por el matcher."""
        filas = []
        for bi, b in enumerate(bloques):
            for f in b['facturas']:
                for p in f['productos']:
                    filas.append({'bi': bi, 'factura': f['nro'], 'cant': p['cant'], 'texto': p['texto'],
                                  'Producto': p['producto'], 'Marca': p['marca'] or float('nan'), 'Cantidad': p['cant']})
        if not filas:
            return filas
        df = pd.DataFrame(filas)
        dfm, pn, mp, mm = _maestro(empresa)
        df2, _u = G['matchear_productos'](df[['Producto', 'Marca', 'Cantidad']].copy(), dfm, pn, mp, mm)
        for i, f in enumerate(filas):
            r = df2.iloc[i]
            correcto = str(r['PRODUCTO_CORRECTO'])
            ia = str(r.get('REVISION_IA', '') or '')
            sin = correcto in ('POR CLASIFICAR', 'Vacío', 'PRODUCTO POR IDENTIFICAR')
            f['producto_final'] = f['Producto'] if sin else correcto
            f['marca_final'] = (str(r['MARCA_MAESTRA']) if not sin and str(r['MARCA_MAESTRA']) != 'Desconocida'
                                else (f['Marca'] if isinstance(f['Marca'], str) else 'N/D'))
            f['precio'] = 0 if sin else float(r['PRECIO_MAESTRO'])
            # Resguardo extra (el reporte usa el mejor parecido aunque la marca no cuadre):
            # si el freelance nombró una marca y el producto encontrado es de OTRA, se revisa a mano.
            marca_dif = False
            if not sin and isinstance(f['Marca'], str) and f['Marca'].strip():
                g_txt = G['obtener_grupo_marca'](f['Marca'])
                g_ref = G['obtener_grupo_marca'](str(r['MARCA_MAESTRA']))
                texto_ref = G['estandarizar_texto'](correcto)
                if g_txt != g_ref and G['remover_tildes'](f['Marca']).lower() not in texto_ref:
                    marca_dif = True
            if sin:
                f['estado'], f['nota'] = 'REVISAR', 'Sin coincidencia clara en la base'
            elif marca_dif:
                f['estado'], f['nota'] = 'REVISAR', f"Escribió {f['Marca']} pero lo más parecido es de otra marca"
            elif ia.startswith('IA DUDOSA'):
                f['estado'], f['nota'] = 'REVISAR', 'La IA dudó entre varias opciones'
            elif ia.startswith('IA:'):
                f['estado'], f['nota'] = 'OK', 'Emparejado con IA'
            else:
                f['estado'], f['nota'] = 'OK', ''
        return filas

    @bp.route('/api/cargar/analizar', methods=['POST'])
    def analizar():
        if not logueado():
            return jsonify({'success': False, 'error': 'No autorizado'}), 401
        try:
            d = request.get_json(force=True)
            cod = (d.get('codigo') or '').strip().upper()
            fecha = _fecha_ddmmyyyy(d.get('fecha'))
            if cod not in TIENDAS:
                return jsonify({'success': False, 'error': 'Elige una tienda de la lista.'}), 400
            if not fecha:
                return jsonify({'success': False, 'error': 'Elige la fecha.'}), 400
            if not (d.get('mensaje') or '').strip():
                return jsonify({'success': False, 'error': 'Pega el mensaje.'}), 400
            emp, suc = TIENDAS[cod]
            bloques = parsear_mensajes(d['mensaje'])
            if not bloques:
                return jsonify({'success': False, 'error': 'No encontré ninguna hora en el mensaje.'}), 400
            vistos, avisos_gen = {}, []
            _aplicar_manuales(bloques, d.get('hora'), d.get('visitas'), avisos_gen)
            for i, b in enumerate(bloques):
                if b['horario'] in vistos:
                    avisos_gen.append(f"La hora {b['horario']} aparece dos veces en lo pegado: se guardará la última.")
                    bloques[vistos[b['horario']]]['duplicada'] = True
                vistos[b['horario']] = i
            filas = _analizar_bloques(emp, bloques)
            salida = []
            for bi, b in enumerate(bloques):
                salida.append({
                    'horario': b['horario'], 'visitas': b['visitas'], 'avisos': b['avisos'],
                    'duplicada': bool(b.get('duplicada')), 'sin_ventas': b['sin_ventas'] and not b['facturas'],
                    'id_envio': f"WA-{cod.replace('-', '')}-{fecha[6:]}{fecha[3:5]}{fecha[:2]}-{b['horario']}",
                    'filas': [{'factura': f['factura'], 'cant': f['cant'], 'original': f['texto'],
                               'producto': f['producto_final'], 'marca': f['marca_final'], 'precio': f['precio'],
                               'estado': f['estado'], 'nota': f['nota']} for f in filas if f['bi'] == bi]})
            return jsonify({'success': True, 'codigo': cod, 'fecha': fecha, 'tienda': _nombre_corto(emp, suc),
                            'empresa': emp, 'bloques': salida, 'avisos': avisos_gen})
        except Exception as e:
            app.logger.error("analizar: %s", e, exc_info=True)
            return jsonify({'success': False, 'error': f'No pude analizar el mensaje: {e}'}), 500

    @bp.route('/api/cargar/buscar_producto')
    def buscar_producto():
        if not logueado():
            return jsonify({'success': False, 'error': 'No autorizado'}), 401
        emp = request.args.get('empresa', '')
        q = (request.args.get('q') or '').strip()
        if emp not in GID_POR_EMPRESA or len(q) < 2:
            return jsonify({'success': True, 'resultados': []})
        dfm, pn, mp, mm = _maestro(emp)
        nombres = list(mp.keys())
        top = rf_process.extract(G['estandarizar_texto'](q), [G['estandarizar_texto'](n) for n in nombres],
                                 scorer=fuzz.token_set_ratio, limit=8)
        return jsonify({'success': True, 'resultados': [
            {'producto': nombres[i], 'marca': ('' if pd.isna(mm.get(nombres[i])) else str(mm.get(nombres[i]))),
             'precio': mp[nombres[i]]} for _, _, i in top]})

    @bp.route('/api/cargar/guardar', methods=['POST'])
    def guardar():
        if not logueado():
            return jsonify({'success': False, 'error': 'No autorizado'}), 401
        if not sheets_configurado():
            return jsonify({'success': False, 'error': 'Falta conectar la hoja de Google (ver instrucciones).'}), 503
        try:
            d = request.get_json(force=True)
            cod = (d.get('codigo') or '').strip().upper()
            fecha = _fecha_ddmmyyyy(d.get('fecha'))
            if cod not in TIENDAS or not fecha:
                return jsonify({'success': False, 'error': 'Tienda o fecha inválida.'}), 400
            emp, suc = TIENDAS[cod]
            filas, ids = [], []
            for b in d.get('bloques', []):
                if b.get('duplicada') or not b.get('horario'):
                    continue
                hid = f"WA-{cod.replace('-', '')}-{fecha[6:]}{fecha[3:5]}{fecha[:2]}-{b['horario']}"
                ids.append(hid)
                comunes = [hid, fecha, b['horario'], emp, cod, suc,
                           suc if emp == 'ddamasco' else '', suc if emp == 'ddaka' else '']
                vis = b.get('visitas', '')
                if not b.get('filas'):
                    filas.append(comunes + ['Sin productos', '', '', '', '', vis])
                    continue
                for f in b['filas']:
                    filas.append(comunes + [str(f['producto']).strip(), str(f.get('marca') or 'N/D'),
                                            int(f['cant']), float(f.get('precio') or 0), f"WA{f['factura']}", vis])
            if not filas:
                return jsonify({'success': False, 'error': 'No hay nada para guardar.'}), 400
            reemplazadas = 0
            hojas = [HOJA_TIPIFICACIONES] + ([HOJA_MANUAL] if ESCRIBIR_MANUAL else [])
            if ESCRIBIR_MANUAL:
                _asegurar_hoja(HOJA_MANUAL)
            for h in hojas:
                n = _borrar_envios(h, set(ids))
                reemplazadas = max(reemplazadas, n)
                _agregar_filas(h, filas)
            # la caché de 10 min del reporte se vence para ver lo nuevo al instante
            try:
                with G['_cache_lock']:
                    G['_cache_campo_time'] = 0.0
                    G['_cache_campo_data'] = None
            except Exception:
                pass
            return jsonify({'success': True, 'horas': len(ids), 'filas': len(filas), 'reemplazadas': reemplazadas})
        except Exception as e:
            app.logger.error("guardar: %s", e, exc_info=True)
            return jsonify({'success': False, 'error': str(e)}), 500

    def _num(x, ent=False):
        t = re.sub(r'[^0-9.,]', '', str(x or ''))
        if not t:
            return 0
        if ',' in t and '.' in t:
            t = t.replace(',', '') if t.rfind('.') > t.rfind(',') else t.replace('.', '').replace(',', '.')
        elif ',' in t:
            t = t.replace(',', '.') if len(t.split(',')[-1]) != 3 else t.replace(',', '')
        try:
            v = float(t)
        except ValueError:
            return 0
        return int(v) if ent else v

    @bp.route('/api/cargar/hora')
    def hora():
        """Filas de una hora ya guardada (para editarla)."""
        if not logueado():
            return jsonify({'success': False, 'error': 'No autorizado'}), 401
        hid = (request.args.get('id') or '').strip()
        if not hid.startswith('WA-'):
            return jsonify({'success': False, 'error': 'Esta hora vino del formulario de Kobo y no se puede editar aquí.'}), 400
        if not sheets_configurado():
            return jsonify({'success': False, 'error': 'Falta conectar la hoja de Google.'}), 503
        try:
            v = _leer_hoja(HOJA_TIPIFICACIONES, "A:N")
            rs = [list(r) + [''] * (14 - len(r)) for r in v[1:] if r and r[0] == hid]
            if not rs:
                return jsonify({'success': False, 'error': 'No encontré esa hora en la hoja.'}), 404
            cod = rs[0][4]
            emp = TIENDAS.get(cod, (rs[0][3], ''))[0]
            filas = []
            for r in rs:
                if str(r[8]).strip().lower() in ('sin productos', ''):
                    continue
                m = re.search(r'(\d+)', str(r[12]))
                filas.append({'factura': int(m.group(1)) if m else len(filas) + 1, 'cant': _num(r[10], True) or 1,
                              'original': 'Ya guardado', 'producto': str(r[8]), 'marca': str(r[9] or 'N/D'),
                              'precio': _num(r[11]), 'estado': 'OK', 'nota': ''})
            return jsonify({'success': True, 'codigo': cod, 'empresa': emp, 'tienda': _nombre_corto(emp, rs[0][5]),
                            'fecha': str(rs[0][1])[:10], 'horario': rs[0][2], 'visitas': rs[0][13], 'id_envio': hid,
                            'filas': filas})
        except Exception as e:
            app.logger.error("hora: %s", e, exc_info=True)
            return jsonify({'success': False, 'error': str(e)}), 500

    @bp.route('/api/cargar/borrar', methods=['POST'])
    def borrar():
        if not logueado():
            return jsonify({'success': False, 'error': 'No autorizado'}), 401
        hid = ((request.get_json(force=True) or {}).get('id') or '').strip()
        if not hid.startswith('WA-'):
            return jsonify({'success': False, 'error': 'Solo se pueden borrar horas cargadas desde esta página.'}), 400
        if not sheets_configurado():
            return jsonify({'success': False, 'error': 'Falta conectar la hoja de Google.'}), 503
        try:
            n = 0
            for h in [HOJA_TIPIFICACIONES] + ([HOJA_MANUAL] if ESCRIBIR_MANUAL else []):
                n = max(n, _borrar_envios(h, {hid}))
            try:
                with G['_cache_lock']:
                    G['_cache_campo_time'] = 0.0
                    G['_cache_campo_data'] = None
            except Exception:
                pass
            return jsonify({'success': True, 'filas': n})
        except Exception as e:
            app.logger.error("borrar: %s", e, exc_info=True)
            return jsonify({'success': False, 'error': str(e)}), 500

    app.register_blueprint(bp)


# ---------------------------------------------------------------------------
# INSTALACION (resumen)
#  1) Subir a GitHub (todos en la carpeta principal): cargador.py, cargar.html, idm_logo.png
#  2) En server.py, antes de  if __name__ == '__main__':  agregar
#         from cargador import init_cargador
#         init_cargador(app, globals())
#  3) requirements.txt: agregar  google-auth>=2.0.0  y  requests>=2.28.0
#  4) Render -> Environment: GOOGLE_SERVICE_ACCOUNT_JSON  y  SPREADSHEET_ID
# ---------------------------------------------------------------------------
