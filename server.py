import os
import re
import io
import time
import unicodedata
import urllib.request
import threading
from datetime import datetime
import pandas as pd

# Global variables for hybrid in-memory caching
_cache_campo_data = None
_cache_campo_time = 0.0
_cache_maestro_data = {}
_cache_maestro_time = {}
# FIX CONCURRENCIA: gunicorn corre con 4 threads sobre el mismo proceso (ver Procfile),
# y estos diccionarios/variables de caché son estado compartido mutable. Sin un lock,
# dos requests simultáneos (ej. dos personas generando un reporte a la vez justo cuando
# la caché expira, o alguien usando /api/sync mientras otro genera un reporte) podían
# pisarse la escritura uno al otro o leer un estado a medio escribir. Se usa RLock (no
# Lock simple) porque es más seguro por defecto si en el futuro alguna función bajo el
# lock terminara llamando a otra que también lo adquiere.
_cache_lock = threading.RLock()
from flask import Flask, jsonify, request, send_file, render_template, redirect, url_for, session, after_this_request
from werkzeug.security import generate_password_hash, check_password_hash
from rapidfuzz import fuzz
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import PieChart, Reference, BarChart
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter

app = Flask(__name__, template_folder='templates', static_folder='static')

import logging
log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'server.log')
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(logging.Formatter(
    '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
))
app.logger.addHandler(file_handler)
app.logger.setLevel(logging.INFO)

# FIX SEGURIDAD: antes la secret_key estaba fija en el código ('idm-secret-key-2026'),
# visible para cualquiera con acceso al repo. Con eso, cualquiera podía forjar/firmar
# cookies de sesión válidas para la app (por ejemplo, entrar como 'admin' sin
# contraseña). Ahora se toma de la variable de entorno FLASK_SECRET_KEY.
# Si no está configurada, se genera una aleatoria al arrancar (más segura que una fija,
# pero invalida las sesiones activas en cada reinicio del servidor) y se deja un aviso
# en el log para que se configure la variable de entorno en el hosting real.
import secrets
_secret_key_env = os.environ.get('FLASK_SECRET_KEY')
if _secret_key_env:
    app.secret_key = _secret_key_env
else:
    app.secret_key = secrets.token_hex(32)
    app.logger.warning(
        "FLASK_SECRET_KEY no está configurada; se generó una clave aleatoria temporal. "
        "Las sesiones activas se invalidarán en cada reinicio del servidor. "
        "Configura FLASK_SECRET_KEY en el entorno de despliegue para evitar esto."
    )

import json

USERS_FILE = os.path.join(os.path.dirname(__file__), 'users.json')

# FIX SEGURIDAD: antes users.json guardaba las contraseñas en texto plano. Ahora se
# guardan como hash (scrypt, vía werkzeug), y el login/creación de usuarios pasa por las
# funciones de abajo en vez de comparar strings directamente.

def _es_hash(valor):
    """True si `valor` ya es un hash de werkzeug (formato nuevo), False si es texto
    plano (usuario creado antes de este fix, o dato corrupto)."""
    return isinstance(valor, str) and valor.startswith(('pbkdf2:', 'scrypt:'))

def verificar_password(valor_guardado, password_ingresada):
    """
    Compara la contraseña ingresada contra lo guardado en users.json, soportando tanto
    hashes (formato nuevo) como texto plano (usuarios que existían antes de este fix,
    para no invalidar sus accesos de golpe — se migran a hash automáticamente la
    primera vez que inician sesión con éxito, ver login_page()).
    """
    if valor_guardado is None:
        return False
    if _es_hash(valor_guardado):
        try:
            return check_password_hash(valor_guardado, password_ingresada)
        except Exception:
            return False
    return valor_guardado == password_ingresada

def load_users():
    if not os.path.exists(USERS_FILE):
        default_users = {
            "admin": generate_password_hash("admin123"),
            "idm": generate_password_hash("idm2026")
        }
        try:
            with open(USERS_FILE, 'w', encoding='utf-8') as f:
                json.dump(default_users, f, indent=4)
        except Exception:
            pass
        return default_users
    try:
        with open(USERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {"admin": generate_password_hash("admin123"), "idm": generate_password_hash("idm2026")}

def save_users(users_dict):
    try:
        with open(USERS_FILE, 'w', encoding='utf-8') as f:
            json.dump(users_dict, f, indent=4)
        return True
    except Exception:
        return False

USERS = load_users()

# --- PARÁMETROS DE PROYECCIÓN (editables desde /admin/parametros) ---
# Antes estos 3 porcentajes por empresa estaban escritos literal en el código
# de generate_observations (if is_daka: ecommerce_pct = 0.20 ...). Cualquier
# ajuste requería que un desarrollador editara server.py y volviera a
# desplegar. Ahora viven en un archivo editable desde la web, con los mismos
# valores de siempre como default (no cambia nada hasta que alguien los
# edite a propósito).
PARAMETROS_FILE = os.path.join(os.path.dirname(__file__), 'parametros.json')

DEFAULT_PARAMETROS = {
    "ddaka": {"ecommerce_pct": 0.20, "al_mayor_pct": 0.20, "no_visibles_pct": 0.10},
    "ddamasco": {"ecommerce_pct": 0.10, "al_mayor_pct": 0.20, "no_visibles_pct": 0.10},
    "dmultimax": {"ecommerce_pct": 0.10, "al_mayor_pct": 0.20, "no_visibles_pct": 0.10},
}

def load_parametros():
    if not os.path.exists(PARAMETROS_FILE):
        try:
            with open(PARAMETROS_FILE, 'w', encoding='utf-8') as f:
                json.dump(DEFAULT_PARAMETROS, f, indent=4)
        except Exception:
            pass
        return dict(DEFAULT_PARAMETROS)
    try:
        with open(PARAMETROS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # Por si el archivo quedara incompleto (ej. alguien lo edito a mano y
        # se olvido una empresa): completar con los defaults, nunca fallar.
        for empresa, valores in DEFAULT_PARAMETROS.items():
            if empresa not in data:
                data[empresa] = dict(valores)
            else:
                for clave, val in valores.items():
                    data[empresa].setdefault(clave, val)
        return data
    except Exception:
        return dict(DEFAULT_PARAMETROS)

def save_parametros(parametros_dict):
    try:
        with open(PARAMETROS_FILE, 'w', encoding='utf-8') as f:
            json.dump(parametros_dict, f, indent=4)
        return True
    except Exception:
        return False

def obtener_parametros_proyeccion(empresa):
    """
    Devuelve (ecommerce_pct, al_mayor_pct, no_visibles_pct, has_projections)
    para la empresa dada, leyendo parametros.json (editable desde
    /admin/parametros). Antes este mismo bloque if/elif estaba repetido y
    hardcodeado en 3 lugares distintos (get_projected_totals, generate_report,
    generate_observations) — unificado acá para que los tres lean siempre el
    mismo valor, y una edición desde la web se refleje en el Excel y el TXT
    por igual, no solo en uno de los dos.
    """
    empresa_norm = empresa.lower().strip()
    parametros = load_parametros()
    if empresa_norm in parametros:
        p = parametros[empresa_norm]
        return p['ecommerce_pct'], p['al_mayor_pct'], p['no_visibles_pct'], True
    return 0.0, 0.0, 0.0, False

def is_logged_in():
    return 'user' in session

# --- GOOGLE SHEETS GIDS ---
GID_TIPIFICACIONES = "1109198771"
GID_BASE_DAKA = "1240474880"
GID_BASE_DAMASCO = "841459536"
GID_BASE_MULTIMAX = "2089283830"

# --- PALETA FIJA DE COLORES POR CATEGORÍA (gráfico de torta del Excel) ---
# FIX (organización/estética de gráficas): antes el gráfico de categorías no
# tenía colores asignados, así que Excel le ponía su paleta por defecto —
# impredecible, y sin relación con los colores ya usados en los otros
# gráficos de la misma hoja (naranja/azul para Promo/Fuera de Promo). Con
# esto, cada categoría tiene SIEMPRE el mismo color en todos los reportes
# (útil para comparar reportes de distintas fechas/tiendas a simple vista),
# y ninguno choca con el naranja/azul ya usados para Promo.
# --- PALETA DE COLORES DEL GRÁFICO DE CATEGORÍAS (gráfico de torta del Excel) ---
# FIX (organización/estética de gráficas, ajustado a pedido del cliente para
# calzar EXACTO con los reportes que ya hacen a mano): el reporte de referencia
# no usa un color fijo por nombre de categoría — usa los 6 colores de acento
# por defecto de Excel ("tema Office"), asignados por POSICIÓN en la tabla
# (que está ordenada por venta descendente), repitiendo con una variante más
# oscura para las categorías 7 en adelante. Se replican acá los valores hex
# exactos (calculados a partir del tema real de un archivo de referencia).
PALETA_CATEGORIAS_POSICION = [
    "4472C4",  # accent1 - azul
    "ED7D31",  # accent2 - naranja
    "A5A5A5",  # accent3 - gris
    "FFC000",  # accent4 - dorado
    "5B9BD5",  # accent5 - celeste
    "70AD47",  # accent6 - verde
    "264478",  # accent1 oscuro (lumMod 60%)
    "9E480E",  # accent2 oscuro
    "636363",  # accent3 oscuro
    "997300",  # accent4 oscuro
    "255E91",  # accent5 oscuro
]

# --- DATA URLS ---
URL_TIPIFICACIONES = f"https://docs.google.com/spreadsheets/d/e/2PACX-1vTfq81DhLQ_8jkbFIAs7OWaO7qkYRis350TTRz_BbbsVucVw4K87Ai0YgiynRIQG1CqRJv9i1V6oEDo/pub?gid={GID_TIPIFICACIONES}&single=true&output=csv"

# --- HELPER FUNCTIONS ---
def round_half_up(n):
    return int(n + 0.5) if n >= 0 else int(n - 0.5)

def remover_tildes(texto):
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize('NFD', texto)
    texto = re.sub(r'[\u0300-\u036f]', '', texto)
    return texto

def estandarizar_texto(texto):
    if pd.isna(texto):
        return ""
    t = remover_tildes(str(texto).lower().strip())
    # Convertir decimales (comas y puntos seguidos de 1 o 2 dígitos) al estándar con punto '.'
    t = re.sub(r'(\d+)[,.](\d{1,2})(?![0-9])', r'\1.\2', t)
    # Limpiar separadores de miles (comas y puntos seguidos de 3 dígitos)
    t = re.sub(r'(\d+)[,.](\d{3})(?![0-9])', r'\1\2', t)
    
    # Mapeo de doble tina -> lavadora semiautomatica
    t = t.replace("semi automatica", "semiautomatica")\
         .replace("doble tina", "semiautomatica")\
         .replace("dobletina", "semiautomatica")\
         .replace("dos tinas", "semiautomatica")\
         .replace("dostinas", "semiautomatica")
    t = re.sub(r'\bsemi\b', 'semiautomatica', t)
    t = re.sub(r'\bauto\b', 'automatica', t)
    
    # Mapeo de sinónimos de Base TV
    t = t.replace("base de tv ajustable", "base tv")\
         .replace("base de tv fija", "base tv")\
         .replace("base de tv movil", "base tv")\
         .replace("base de tv móvil", "base tv")\
         .replace("base de televisor", "base tv")\
         .replace("soporte de tv", "base tv")\
         .replace("soporte para tv", "base tv")\
         .replace("base para tv", "base tv")\
         .replace("base de tv", "base tv")\
         .replace("soporte tv", "base tv")\
         .replace("para tv", "base tv")
         
    t = t.replace("nevera", "refrigerador")
    t = t.replace("televisor", "tv")
    t = t.replace("telefono movil", "celular").replace("telefono", "celular")
    t = t.replace("colchón", "colchon").replace("colchon easy go", "colchon")
    
    # Freezer, frezer y congelador es lo mismo
    t = t.replace("freezer", "congelador").replace("frezer", "congelador")
    
    # Audífonos y auriculares es lo mismo
    t = t.replace("auriculares", "audifonos").replace("auricular", "audifonos")
    t = re.sub(r'\baudifono\b', 'audifonos', t)
    
    # A/A, AA y aire acondicionado es lo mismo
    t = t.replace("aire acondicionado", "a/a").replace("a/c", "a/a").replace("splint", "split")
    t = re.sub(r'\baa\b', 'a/a', t)
    t = re.sub(r'\bac\b', 'a/a', t)
    
    # Platos y vajillas son sinónimos
    t = t.replace("platos", "vajilla").replace("plato", "vajilla").replace("vajillas", "vajilla")
    
    # Map screen sizes robustly to pulg (supporting both single and double quotes)
    t = t.replace('"', ' pulg ').replace("'", ' pulg ')
    t = t.replace("pulgadas", ' pulg ').replace("pulgada", ' pulg ').replace(" pulg", ' pulg ').replace("pulg", ' pulg ')

    # Synonyms mapping
    t = t.replace("tosty arepa", "tostiarepa")\
         .replace("tostyarepa", "tostiarepa")\
         .replace("reloj inteligente", "reloj smart")\
         .replace("relojes inteligentes", "relojes smart")
         
    t = t.replace("airfryer", "freidora de aire")\
         .replace("air fryer", "freidora de aire")\
         .replace("airfryers", "freidora de aire")\
         .replace("air fryers", "freidora de aire")\
         .replace("freidora/aire", "freidora de aire")\
         .replace("freidora de aires", "freidora de aire")
         
    t = re.sub(r'\bmicroonda\b', 'microondas', t)

    t = t.replace("refrigerador exhibidor", "refrigerador comercial")\
         .replace("nevera exhibidora", "refrigerador comercial")\
         .replace("vitrina exhibidora", "refrigerador comercial")\
         .replace("exhibidor", "refrigerador comercial")\
         .replace("exhibidora", "refrigerador comercial")
    
    # Estandarizar BTUs (convertir 12btu -> 12000btu, 12k -> 12000 btu, 12.000 btu -> 12000btu, etc.)
    t = re.sub(r'\s+btu', 'btu', t)
    t = re.sub(r'\b(5|6|8|9|12|18|24|36)btu\b', lambda m: f"{int(m.group(1))*1000}btu", t)
    t = re.sub(r'\b(\d+)\s*k\s*(?:btu)?\b', lambda m: str(int(m.group(1)) * 1000) + ' btu', t)
         
    # If it is like 'tv 32', add ' pulg '
    if "tv" in t and "pulg" not in t:
        t = re.sub(r'\btv\s+(\d+)\b', r'tv \1 pulg', t)

    if "tv" in t and "smart" not in t and "analog" not in t:
        t = t + " smart"
        
    t = re.sub(r'(\d+)\s*h\b', r'\1 hornillas', t)
    t = t.replace("hornilla", "hornillas")
    
    # Estandarizar unidades de medida
    t = re.sub(r'(\d+(?:\.\d+)?)\s*(?:l|lt|lts|litros|litro)\b', r'\1 litros', t)
    t = re.sub(r'(\d+(?:\.\d+)?)\s*(?:kg|kgs|kilogramos|kilogramo)\b', r'\1 kg', t)
    t = re.sub(r'(\d+(?:\.\d+)?)\s*(?:pies|pie|ft|cu\s*ft)\b', r'\1 pies', t)
    t = re.sub(r'(\d+(?:\.\d+)?)\s*(?:pulg|pulgadas|pulgada)\b', r'\1 pulg', t)
    t = re.sub(r'(\d+)\s*gb\b', r'\1gb', t)
    
    # Congelador 99L es lo mismo que 100L
    t = re.sub(r'\b99\s*litros\b', '100 litros', t)
    
    return re.sub(r'\s+', ' ', t).strip()

def obtener_capacidad_normalizada(texto, prod_type):
    # Find all floats/ints in the text
    raw_nums = []
    for m in re.findall(r'\b\d+(?:\.\d+)?\b', texto.lower()):
        try:
            raw_nums.append(float(m))
        except ValueError:
            pass
            
    caps = []
    if prod_type in ['NEVERA', 'CONGELADOR']:
        for val in raw_nums:
            if 1.5 <= val <= 26.0:
                # Feet to liters
                caps.append(val * 28.3168)
            elif 30.0 <= val <= 850.0:
                # Liters
                caps.append(val)
    elif prod_type == 'MICROONDAS':
        for val in raw_nums:
            if 0.5 <= val <= 5.0:
                # Cubic feet to liters
                caps.append(val * 28.3168)
            elif 10.0 <= val <= 60.0:
                # Liters
                caps.append(val)
    elif prod_type == 'TV':
        for val in raw_nums:
            if 22.0 <= val <= 110.0:
                # Inches
                caps.append(val)
    elif prod_type == 'A/A':
        for val in raw_nums:
            if 1.0 <= val <= 5.0:
                # Tons to BTU
                caps.append(val * 12000.0)
            elif 5000.0 <= val <= 65000.0:
                # BTU
                caps.append(val)
    else:
        # Fallback to any number
        caps = raw_nums
    return caps

def obtener_grupos_palabras_clave(texto_limpio):
    groups = [
        ['microondas', 'microonda'],
        ['licuadora', 'licuadoras', 'nutribullet'],
        ['freidora', 'airfryer', 'air fryer', 'fryer'],
        ['cafetera', 'espumador'],
        ['horno'],
        ['tostador', 'tostadora', 'sanduchera', 'sandwichera', 'waflera', 'panini', 'tostiarepa'],
        ['olla', 'arrocera', 'multiolla', 'instant pot'],
        ['ventilador'],
        ['plancha'],
        ['exprimidor', 'extractor'],
        ['batidora'],
        ['picatodo', 'picadora', 'procesador'],
        ['nevera', 'refrigerador', 'minibar', 'exhibidora'],
        ['congelador', 'freezer'],
        ['tv', 'televisor', 'pantalla'],
        ['aire', 'split', 'pisotecho', 'a/a'],
        ['dispensador'],
        ['corneta', 'altavoz', 'parlante', 'sonido'],
        ['base', 'soporte']
    ]
    matched_group_indices = []
    for idx, g in enumerate(groups):
        if any(w in texto_limpio for w in g):
            matched_group_indices.append(idx)
    return matched_group_indices

def check_keyword_group_match_indices(query_groups, option_groups):
    for idx in query_groups:
        if idx not in option_groups:
            return False
    return True

def extraer_spec_numerica(texto_norm):
    numeros = set()
    # Find all numbers with units (supporting both integers and decimals)
    matches = re.findall(r'(\d+(?:\.\d+)?)\s*(btu|kg|litros|pies|gb|pulg)', texto_norm)
    for num_str, unit in matches:
        numeros.add(num_str)
        try:
            val = float(num_str)
            if unit == 'pies':
                # Convert to liters
                liters_val = val * 28.3168
                numeros.add(str(round(liters_val)))
                numeros.add(f"{liters_val:.1f}")
            elif unit == 'litros':
                # Convert to pies
                pies_val = val / 28.3168
                numeros.add(f"{pies_val:.1f}")
                numeros.add(str(round(pies_val)))
        except Exception:
            pass
    return numeros

def parse_horario(horario_str):
    if not isinstance(horario_str, str):
        return None, None
    match = re.match(r'(\d+)_(\d+)(am|pm)_(\d+)_(\d+)(am|pm)', horario_str.lower().strip())
    if not match:
        return None, None
    h1, m1, p1, h2, m2, p2 = match.groups()
    
    def to_24h(h, m, p):
        h = int(h)
        m = int(m)
        if p == 'pm' and h != 12:
            h += 12
        elif p == 'am' and h == 12:
            h = 0
        return h + m/60.0

    return to_24h(h1, m1, p1), to_24h(h2, m2, p2)

def formatear_horario_lindo(horario_str):
    if not isinstance(horario_str, str):
        return str(horario_str)
    match = re.match(r'(\d+)_(\d+)(am|pm)_(\d+)_(\d+)(am|pm)', horario_str.lower().strip())
    if match:
        h1, m1, p1, h2, m2, p2 = match.groups()
        return f"{h1}:{m1} {p1.upper()} a {h2}:{m2} {p2.upper()}"
    return horario_str

def calcular_time_pct(opening_hour, limit_hour, studied_hours):
    op_h = int(opening_hour)
    lim_h = int(limit_hour)
    
    unstudied = [h for h in range(op_h, lim_h) if h not in studied_hours]
    if not unstudied:
        return 0.0
        
    time_pct = 0.0
    
    # 1. Morning block (h < 12) - 5% per hour
    morning_unstudied = [h for h in unstudied if h < 12]
    time_pct += 0.05 * len(morning_unstudied)
    
    # 2. Middle block (12 <= h < 17) - 5% per hour
    middle_unstudied = [h for h in unstudied if 12 <= h < 17]
    time_pct += 0.05 * len(middle_unstudied)
    
    # 3. Evening block (h >= 17)
    # - 5pm to 6pm (h=17, lim_h=18) = 15%
    # - 5pm to 7pm/8pm/9pm (h=17, lim_h>=19) = 30%
    # - 6pm to 7pm/8pm/9pm (h=18, lim_h>=19) = 15%
    # - Any hour outside this table (like h>=21 if lim_h>21) adds 5% per hour
    if 17 in unstudied:
        if lim_h == 18:
            time_pct += 0.15
        elif lim_h >= 19:
            time_pct += 0.30
            if lim_h > 21:
                time_pct += 0.05 * (lim_h - 21)
    elif 18 in unstudied:
        if lim_h >= 19:
            time_pct += 0.15
            if lim_h > 21:
                time_pct += 0.05 * (lim_h - 21)
    else:
        # 17 and 18 are both studied
        # Any late unstudied hour >= 19 adds 5% per hour
        evening_unstudied_late = [h for h in unstudied if h >= 19]
        time_pct += 0.05 * len(evening_unstudied_late)
        
    return time_pct

def format_hour_12h(h_val):
    h_int = int(h_val)
    m_int = int(round((h_val - h_int) * 60))
    suffix = 'PM' if h_int >= 12 else 'AM'
    h_12 = h_int - 12 if h_int > 12 else h_int
    if h_12 == 0:
        h_12 = 12
    return f"{h_12}:{m_int:02d}{suffix}"

def parse_input_hour(h_str):
    if not isinstance(h_str, str):
        return 24.0
    h_str = h_str.lower().strip()
    m = re.match(r'(\d+):(\d+)\s*(am|pm)?', h_str)
    if m:
        h, mins, p = m.groups()
        h = int(h)
        mins = int(mins)
        if p == 'pm' and h != 12:
            h += 12
        elif p == 'am' and h == 12:
            h = 0
        return h + mins/60.0

def parse_date(date_str):
    if not date_str or not isinstance(date_str, str):
        return None
    from datetime import datetime
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None

def get_projected_totals(df_study, opening_hour_val, limit_hour_val, empresa):
    # Group hourly data
    t1_data = []
    for hor, group in df_study.groupby('Horario', sort=False):
        prod_sold = int(group['Cantidad'].sum())
        invoices = group.dropna(subset=['Factura']).groupby(['Horario', 'Factura']).ngroups
        visits = group['Visitas'].dropna().max()
        if pd.isna(visits):
            visits = 0
        t1_data.append({
            'Horario': hor,
            'Facturas': invoices,
            'Visitas/Clientes': int(visits)
        })
    df_t1 = pd.DataFrame(t1_data)
    
    if df_t1.empty:
        return 0, 0
    
    # Hours times
    hours_times = []
    for hor in df_t1['Horario'].unique():
        t1, t2 = parse_horario(hor)
        if t1 is not None:
            hours_times.extend([t1, t2])
            
    studied_hours = set()
    for hor in df_t1['Horario'].unique():
        t1, t2 = parse_horario(hor)
        if t1 is not None and t2 is not None:
            for h in range(int(t1), int(t2)):
                studied_hours.add(h)
                
    time_pct = calcular_time_pct(opening_hour_val, limit_hour_val, studied_hours)
    
    # Projected visits
    d_general = df_t1['Visitas/Clientes'].sum()
    d_tiempo = round_half_up(d_general * time_pct)
    projected_visits = d_general + d_tiempo
    
    # Projected sales
    large_invoices_sum_val = 0.0
    large_invoices_set = set()
    df_inv = df_study.dropna(subset=['Factura'])
    for (hor, inv), group in df_inv.groupby(['Horario', 'Factura']):
        inv_sum = group['VENTA_TOTAL'].sum()
        if inv_sum >= 1000.0:
            large_invoices_sum_val += float(inv_sum)
            large_invoices_set.add((hor, inv))
            
    total_conteo_val = df_study['VENTA_TOTAL'].sum()
    neto_val = total_conteo_val - large_invoices_sum_val
    tiempo_val = round_half_up(neto_val * time_pct)
    
    ecommerce_pct, al_mayor_pct, no_visibles_pct, has_projections = obtener_parametros_proyeccion(empresa)

    if has_projections:
        total_general_conteo_val = round_half_up(neto_val + tiempo_val)
        ecommerce_val = round_half_up(total_general_conteo_val * ecommerce_pct)
        mayor_val = round_half_up(total_general_conteo_val * al_mayor_pct)
        novisibles_val = round_half_up(total_general_conteo_val * no_visibles_pct)
        projected_sales = round_half_up(total_conteo_val + tiempo_val + ecommerce_val + mayor_val + novisibles_val)
    else:
        projected_sales = round_half_up(neto_val + tiempo_val)
        
    return projected_sales, projected_visits

def find_previous_study_df(df_all, current_date_str, empresa, sucursal, col_sucursal, limit_hour):
    current_dt = parse_date(current_date_str)
    if not current_dt:
        return None
        
    # Filter by company and sucursal
    df_same_store = df_all[
        (df_all['Empresa'].astype(str).str.lower().str.strip() == empresa.lower().strip()) &
        (df_all[col_sucursal].astype(str).str.lower().str.strip() == sucursal.lower().strip())
    ]
    
    # Find all dates and parse them
    past_studies = []
    for date_val, group in df_same_store.groupby('Fecha'):
        dt = parse_date(date_val)
        if dt and dt < current_dt:
            days_diff = (current_dt - dt).days
            if days_diff <= 60:
                past_studies.append((dt, date_val, group))
                
    if not past_studies:
        return None
        
    # Sort past studies by date descending to get the most recent one
    past_studies.sort(key=lambda x: x[0], reverse=True)
    df_prev = past_studies[0][2]
    
    # Apply the same hour filters as current study
    def is_before_close(row):
        t1, t2 = parse_horario(row['Horario'])
        if t1 is None:
            return True
        return t2 <= limit_hour
        
    df_prev_filtered = df_prev[df_prev.apply(is_before_close, axis=1)].copy()
    return df_prev_filtered

def contains_word(text, word):
    return any(w == word for w in re.findall(r'[a-z0-9+/]+', text.lower()))

def categorizar_producto(nombre_producto):
    nombre = remover_tildes(str(nombre_producto).lower().strip())
    
    # 1. A/A
    if any(k in nombre for k in ['aire', 'a/a', 'split', 'pisotecho', 'piso techo']) and 'freidora' not in nombre and 'fryer' not in nombre:
        return 'A/A'
    # 2. TV
    elif any(k in nombre for k in ['tv', 'televisor', 'pantalla']) and 'base' not in nombre and 'soporte' not in nombre:
        return 'TV'
    # 3. LAVADO
    elif any(k in nombre for k in ['lavadora', 'secadora', 'doble tina', 'semiautomatica']):
        return 'LAVADO'
    # 4. CONGELADOR
    elif any(k in nombre for k in ['congelador', 'freezer']):
        return 'CONGELADOR'
    # 5. NEVERA
    elif any(k in nombre for k in ['refrigerador', 'nevera']):
        return 'NEVERA'
    # 6. COCINA
    elif any(k in nombre for k in ['tope a gas', 'tope electrico', 'tope dual', 'cocina a gas', 'campana', 'estufa', 'cocina', 'sarten', 'cubierto', 'vajilla', 'plato', 'vaso', 'utensilio']) or ('olla' in nombre and 'arrocera' not in nombre):
        return 'COCINA'
    # 7. COMPUTACION
    elif any(k in nombre for k in ['laptop', 'computadora', 'monitor', 'auriculares', 'audifonos', 'audifono', 'teclado', 'mouse', 'raton', 'router', 'modem']) or contains_word(nombre, 'pc') or contains_word(nombre, 'ups'):
        if any(x in nombre for x in ['morral', 'bolso', 'mochila', 'funda']):
            return 'OTROS'
        return 'COMPUTACION'
    # 8. SONIDO
    elif any(k in nombre for k in ['corneta', 'altavoz', 'barra de sonido', 'sonido', 'parlante', 'subwoofer']):
        return 'SONIDO'
    # 9. TELEFONIA
    elif any(k in nombre for k in ['celular', 'telefono', 'smartphone', 'movil']):
        return 'TELEFONIA'
    # 10. ELECTRODOMESTICOS
    elif any(k in nombre for k in [
        'licuadora', 'freidora', 'airfryer', 'air fryer', 'tostador', 'tostadora', 'cafetera', 'sanduchera',
        'sandwichera', 'waflera', 'batidora', 'procesador', 'arrocera', 'nutribullet', 'hervidor', 'tetera',
        'extractor', 'exprimidor', 'multiolla', 'instant pot', 'parrilla', 'parrillera', 'panini', 'afilador', 'balanza',
        'plancha', 'aspiradora', 'secador', 'rizador', 'rasuradora', 'recortadora', 'cepillo de dientes',
        'ventilador', 'humificador', 'humidificador', 'picatodo', 'crepera', 'deshidratador', 'cotufera',
        'espumador', 'yogurtera', 'hidrojet', 'microonda', 'olla de presion', 'olla presion',
        'dispensador', 'horno tostador', 'horno electrico', 'horno freidora', 'horno freidor', 'tosta horno'
    ]):
        return 'ELECTRODOMESTICOS'
    else:
        return 'OTROS'

def is_own_brand(marca_db, empresa):
    emp = empresa.lower().strip()
    brand = remover_tildes(str(marca_db)).upper().strip()
    
    if 'daka' in emp:
        own_brands = ['HYUNDAI', 'DAEWOO', 'SANSUI', 'RIBELLE', 'BM', 'BREMEN']
    elif 'damasco' in emp:
        own_brands = ['DA+CO', 'DACO', 'DAMASCO']
    elif 'multimax' in emp:
        own_brands = ['CLX', 'CONDESA', 'FRIGILUX', 'MILATTI', 'KUCCE', 'VETRUX', 'MUNDO BLANCO']
    else:
        own_brands = []
        
    return any(ob in brand for ob in own_brands)

def obtener_grupo_marca(marca):
    m = remover_tildes(str(marca)).lower().strip()
    m = m.replace("+", "").replace(" ", "").replace("-", "")
    if any(k in m for k in ['damasco', 'daco']):
        return 'damasco'
    if any(k in m for k in ['multimax', 'clx', 'condesa', 'frigilux', 'milatti', 'kucce', 'vetrux', 'mundoblanco']):
        return 'multimax'
    if 'daka' in m:
        return 'daka'
    return m
def normalize_brands_for_matching(text):
    t = text.lower()
    t = re.sub(r'\bda\+co\b', 'damasco', t)
    t = re.sub(r'\bdaco\b', 'damasco', t)
    return t

def limpiar_descripcion_producto(texto):
    """
    Limpia la descripción de un producto ya emparejado (PRODUCTO_CORRECTO) para
    el reporte, removiendo códigos internos de SKU/modelo que no aportan nada
    en un reporte gerencial — igual a como el equipo lo hace a mano hoy (ver
    bitácora, pedido explícito de calcar los reportes manuales).

    Solo se aplica DESPUÉS del matching, sobre el texto que se muestra —
    nunca sobre el texto que usa el motor de comparación (buscar_coincidencia_
    tecnica sigue comparando contra el texto original completo, sin tocar).

    Preserva specs técnicas relevantes (capacidad, voltaje, tamaño, BTU, etc.)
    y nombres de línea de producto reales (ej. "GALAXY A57", "X9B") — solo
    descarta tokens que mezclan letras y números de forma típica de un
    código de catálogo/modelo interno (ej. "DMTL-MS12-W1", "MCL2440ESBB0").
    """
    if not isinstance(texto, str) or not texto.strip():
        return texto

    t = texto.strip()

    # Normalizar "DA+CO" (como se ve la marca repetida en el propio nombre)
    # a "DAMASCO", igual que en los reportes de referencia.
    t = re.sub(r'\bDA\+CO\b', 'DAMASCO', t, flags=re.IGNORECASE)

    # Remover códigos puramente numéricos con guión (ej. "086-545979")
    t = re.sub(r'\b\d{2,4}-\d{4,8}\b', '', t)

    # Unidades/specs conocidas que NUNCA se deben confundir con un código de
    # modelo, aunque mezclen letra y número. Cubre varios formatos reales
    # encontrados en las 3 listas maestras:
    #   - simple:      "220V", "12000BTU", "1.5L"
    #   - rango:        "13-25MM", "12-18KG"
    #   - dimensión x:  "92X30", "100X200CM"
    #   - dimensión LW: "L100*W100", "L100*W100*H50"
    #   - combo con +:  "256GB+12GB", "8+512GB" (RAM+almacenamiento de celulares)
    UNIDAD = r'(BTU|KGS?|LTS?|LT|L|ML|V|VA|WATT|W|HZ|GB|MB|TB|MM|CM|M|"|\'|HRS?|H|PZAS?|PZ|PCS|PIES?|K)'
    NUM = r'\d+(?:[.,]\d+)?'
    patron_spec = re.compile(
        rf'^('
        rf'{NUM}{UNIDAD}?'                                  # numero + unidad opcional
        rf'|{NUM}-{NUM}{UNIDAD}?'                            # rango: 13-25MM
        rf'|{NUM}[xX]{NUM}{UNIDAD}?'                         # dimension: 92X30, 100X200CM
        rf'|{NUM}{UNIDAD}?\+{NUM}{UNIDAD}?'                  # combo: 256GB+12GB, 8+512GB
        rf'|[LWHD]{NUM}(?:[*xX][LWHD]?{NUM}){{1,2}}'         # L100*W100, D140*H200
        rf')$',
        re.IGNORECASE
    )

    palabras_limpias = []
    for palabra in t.split():
        base = palabra.strip('.,;:()')
        tiene_letra = any(c.isalpha() for c in base)
        tiene_digito = any(c.isdigit() for c in base)
        es_codigo_candidato = tiene_letra and tiene_digito and len(base) >= 5
        if es_codigo_candidato and not patron_spec.match(base):
            continue  # descartar: parece código de modelo/SKU, no un spec conocido
        palabras_limpias.append(palabra)

    resultado = ' '.join(palabras_limpias)
    resultado = re.sub(r'\s{2,}', ' ', resultado).strip(' -,')

    # Si al limpiar el código quedó la palabra "modelo" colgada sola al
    # final (ej. "Lámpara decorativa modelo" — el código que la seguía se
    # descartó arriba), se saca también: ningún nombre de producto real
    # termina en la palabra suelta "modelo".
    resultado = re.sub(r'\s+modelo$', '', resultado, flags=re.IGNORECASE).strip()

    # Salvaguarda: si la limpieza dejó el texto vacío o casi vacío (ej. un
    # producto cuyo nombre completo era un código), se devuelve el original
    # para no perder la referencia del producto por completo.
    if len(resultado) < 3:
        return t
    return resultado


def buscar_coincidencia_tecnica(fila, universo_maestro):
    producto_campo = fila['Producto']
    marca_campo = remover_tildes(str(fila['Marca'])).lower().strip() if not pd.isna(fila['Marca']) else ""
    
    if pd.isna(producto_campo):
        return "Vacío"
    
    campo_limpio = estandarizar_texto(producto_campo)
    
    if marca_campo and marca_campo not in campo_limpio:
        campo_limpio = f"{campo_limpio} {marca_campo}"

    categoria_campo = categorizar_producto(campo_limpio)

    # Determine product type for capacity matching
    prod_type = None
    if categoria_campo == 'NEVERA':
        prod_type = 'NEVERA'
    elif categoria_campo == 'CONGELADOR':
        prod_type = 'CONGELADOR'
    elif 'microonda' in campo_limpio or 'microondas' in campo_limpio:
        prod_type = 'MICROONDAS'
    elif categoria_campo == 'TV':
        prod_type = 'TV'
    elif categoria_campo == 'A/A':
        prod_type = 'A/A'

    query_caps = obtener_capacidad_normalizada(campo_limpio, prod_type)
    query_kw_groups = obtener_grupos_palabras_clave(campo_limpio)

    def buscar_en_candidatos(candidatos_cands):
        if not candidatos_cands:
            return "POR CLASIFICAR", -1
            
        mejor_opcion = "POR CLASIFICAR"
        mejor_puntaje = -1
        
        key_terms = ['arrocera', 'freidora', 'licuadora', 'cafetera', 'horno', 'nevera', 'refrigerador', 'congelador', 'lavadora', 'secadora', 'olla', 'tope', 'campana', 'aire', 'televisor', 'tv']
        query_keys = [k for k in key_terms if k in campo_limpio]

        # Precompute query-specific variables once before the loop
        q_norm = normalize_brands_for_matching(campo_limpio)
        query_is_mount = 'base' in campo_limpio or 'soporte' in campo_limpio
        query_has_led = "led" in str(producto_campo).lower()
        query_has_semi = any(x in campo_limpio for x in ['semi', 'doble tina', 'dos tinas', 'semiautomatica'])

        for opcion in candidatos_cands:
            # Enforce keyword group matches using precomputed indices
            if not check_keyword_group_match_indices(query_kw_groups, opcion.get('kw_groups', [])):
                continue

            # Use precomputed brand-normalized text
            o_norm = opcion.get('brand_norm', opcion['normalizado'])
            score = fuzz.token_set_ratio(q_norm, o_norm)
            
            # Avoid category switching
            categoria_opcion = opcion.get('categoria', 'POR CLASIFICAR')
            if categoria_opcion != categoria_campo:
                score = 0
                
            # Prevent matching mounts/bases with the main product (e.g. Base TV with TV)
            option_is_mount = opcion.get('is_mount', False)
            if query_is_mount != option_is_mount:
                score = 0
                
            # Prevent confusing Microondas with Horno electrico
            query_is_microondas = 'microonda' in campo_limpio or 'microondas' in campo_limpio
            option_is_microondas = 'microonda' in opcion['normalizado'] or 'microondas' in opcion['normalizado']
            query_is_horno = 'horno' in campo_limpio and not query_is_microondas
            option_is_horno = 'horno' in opcion['normalizado'] and not option_is_microondas
            if (query_is_microondas and option_is_horno) or (query_is_horno and option_is_microondas):
                score = 0
                
            # Key terms penalty
            for k in query_keys:
                if k not in opcion['normalizado']:
                    score -= 40
                    
            # Prioritize Smart TV
            if categoria_campo == 'TV':
                is_smart_opcion = opcion.get('is_smart', False)
                if is_smart_opcion:
                    score += 5
                
                # Exclude LED TV candidates unless LED is explicitly in the user's field product query
                is_basic_led = opcion.get('is_basic_led', False)
                if not query_has_led and is_basic_led:
                    score -= 30
                
            # Wash machine match constraints
            if categoria_campo == 'LAVADO':
                option_has_semi = opcion.get('has_semi', False)
                if query_has_semi != option_has_semi:
                    score -= 40
                
            if "smart" in campo_limpio:
                if "analog" in opcion['normalizado'] or "smart" not in opcion['normalizado']:
                    score = 0  
            if "monitor" in campo_limpio:
                if "silla" in opcion['normalizado'] or "gamer" in opcion['normalizado'] or "gaming" in opcion['normalizado']:
                    score = 0
                    
            # Nevera without capacity prioritizes price between 300 and 600
            if categoria_campo == 'NEVERA':
                if not query_caps:
                    price_val = opcion.get('precio', 0.0)
                    if 300.0 <= price_val <= 600.0:
                        score += 15
                        
            # Capacity score adjustment (with support for unit conversion liters <-> feet)
            if query_caps:
                op_caps = opcion.get('capacidades', [])
                if op_caps:
                    tolerance = 1.5
                    if prod_type in ['NEVERA', 'CONGELADOR']:
                        tolerance = 5.0
                    elif prod_type == 'A/A':
                        tolerance = 1000.0
                    
                    if any(abs(q - o) <= tolerance for q in query_caps for o in op_caps):
                        score += 30
                    else:
                        if prod_type in ['NEVERA', 'CONGELADOR', 'MICROONDAS']:
                            min_diff_std = min(abs(q - o) for q in query_caps for o in op_caps) / 28.3168
                            score -= min(35.0, min_diff_std * 6.0)
                        elif prod_type == 'TV':
                            min_diff_std = min(abs(q - o) for q in query_caps for o in op_caps)
                            score -= min(35.0, min_diff_std * 1.5)
                        elif prod_type == 'A/A':
                            min_diff_std = min(abs(q - o) for q in query_caps for o in op_caps) / 1000.0
                            score -= min(35.0, min_diff_std * 1.0)
                        else:
                            min_diff_std = min(abs(q - o) for q in query_caps for o in op_caps)
                            score -= min(35.0, min_diff_std * 1.5)
                else:
                    score -= 20

            if score > mejor_puntaje:
                mejor_puntaje = score
                mejor_opcion = opcion['original']
                
        return mejor_opcion, mejor_puntaje

    # 1. Intentar primero filtrando estrictamente por marca
    opciones_filtradas = []
    if marca_campo:
        marca_campo_grupo = obtener_grupo_marca(marca_campo)
        for p in universo_maestro:
            cand_brand_grupo = p['marca_grupo']
            if cand_brand_grupo == marca_campo_grupo or marca_campo_grupo in p['normalizado'] or p['marca_normalizada'] in campo_limpio:
                opciones_filtradas.append(p)

    res_opcion, res_score = "POR CLASIFICAR", -1
    if opciones_filtradas:
        res_opcion, res_score = buscar_en_candidatos(opciones_filtradas)

    # 2. Si no se encontró nada con la marca, el puntaje es bajo, o hay especificación de capacidad, buscar en todo el universo maestro
    if res_score < 75 or res_opcion == "POR CLASIFICAR" or query_caps:
        res_opcion_global, res_score_global = buscar_en_candidatos(universo_maestro)
        if res_score_global > res_score:
            res_opcion, res_score = res_opcion_global, res_score_global

    if res_score > 0 and res_opcion != "POR CLASIFICAR":
        return res_opcion
    return "POR CLASIFICAR"

def check_promo_status(producto_maestro, precio):
    nombre = estandarizar_texto(producto_maestro)
    
    # 1. Neveras
    if any(k in nombre for k in ['refrigerador', 'nevera', 'freezer', 'congelador vertical']):
        if 'ejec' in nombre or 'ejecutiva' in nombre or 'minibar' in nombre:
            if '2p' in nombre or '2 pies' in nombre:
                return precio <= 140.0, 'Neveras (EJEC 2P)'
            elif '3p' in nombre or '3 pies' in nombre:
                return precio <= 170.0, 'Neveras (EJEC 3P)'
            elif '4p' in nombre or '4 pies' in nombre:
                return precio <= 230.0, 'Neveras (EJEC 4P)'
            elif '5p' in nombre or '5 pies' in nombre:
                return precio <= 250.0, 'Neveras (EJEC 5P)'
            else:
                return precio <= 170.0, 'Neveras (EJEC 3P)'
        
        pies = None
        m_p = re.search(r'(\d+)\s*(?:p|pies|p c)', nombre)
        if m_p:
            pies = int(m_p.group(1))
        else:
            m_l = re.search(r'(\d+)\s*(?:litros|lts|l)', nombre)
            if m_l:
                lits = int(m_l.group(1))
                pies = round(lits / 28.3)
        
        if pies:
            limits = {
                7: 400.0, 8: 395.0, 9: 600.0, 10: 610.0,
                11: 625.0, 12: 750.0, 13: 800.0
            }
            p_cap = max(7, min(13, pies))
            limit = limits.get(p_cap, 800.0)
            return precio <= limit, f'Neveras (TF {p_cap}P)'
        return precio <= 500.0, 'Neveras'
        
    # 2. Lavado
    elif any(k in nombre for k in ['lavadora', 'secadora']):
        is_semi = any(k in nombre for k in ['semi', 'doble tina', 'dos tinas', 'semiautomatica'])
        kg = None
        m_kg = re.search(r'(\d+)\s*(?:kg|k)', nombre)
        if m_kg:
            kg = int(m_kg.group(1))
            
        if kg:
            if is_semi:
                limits = {
                    5: 150.0, 6: 180.0, 7: 200.0, 8: 220.0,
                    9: 230.0, 10: 250.0, 11: 275.0, 12: 400.0,
                    13: 380.0, 15: 380.0
                }
                closest_kg = min(limits.keys(), key=lambda k: abs(k - kg))
                limit = limits[closest_kg]
                return precio <= limit, f'Lavado (SEMI-AUTO {closest_kg}KG)'
            else:
                limits = {
                    7: 280.0, 8: 330.0, 9: 330.0, 10: 340.0,
                    11: 350.0, 12: 430.0
                }
                closest_kg = min(limits.keys(), key=lambda k: abs(k - kg))
                limit = limits[closest_kg]
                return precio <= limit, f'Lavado (AUTOMATICA {closest_kg}KG)'
        return precio <= 300.0, 'Lavado'
        
    # 3. Aires Acondicionados
    elif any(k in nombre for k in ['aire', 'a/a', 'split', 'piso techo']) and 'freidora' not in nombre and 'enfriador' not in nombre:
        btu = None
        m_btu = re.search(r'(\d+)\s*(?:btu|mil)', nombre)
        if m_btu:
            val = int(m_btu.group(1))
            if val > 1000:
                btu = round(val / 1000)
            else:
                btu = val
        else:
            if '3 ton' in nombre:
                btu = 36
            elif '5 ton' in nombre:
                btu = 60
                
        is_split = 'split' in nombre
        is_pisotecho = 'piso techo' in nombre or 'pisotecho' in nombre
        is_inverter = 'inverter' in nombre or 'inv' in nombre
        is_110v = '110v' in nombre or '110 v' in nombre
        is_ventana = 'ventana' in nombre
        
        if is_pisotecho:
            if btu and btu >= 48:
                return precio <= 1750.0, 'A/A (PISO TECHO 5 TON)'
            return precio <= 1200.0, 'A/A (PISO TECHO 3 TON)'
            
        if is_split:
            if btu == 12:
                if is_inverter:
                    return precio <= 360.0, 'A/A (SPLIT 12 MIL INV 220V)'
                if is_110v:
                    return precio <= 330.0, 'A/A (SPLIT 12 MIL 110V)'
                return precio <= 310.0, 'A/A (SPLIT 12 MIL 220V)'
            elif btu == 18:
                return precio <= 470.0, 'A/A (SPLIT 18 MIL 220V)'
            elif btu == 24:
                return precio <= 550.0, 'A/A (SPLIT 24 MIL 220V)'
            return precio <= 500.0, 'A/A (Split)'
            
        # Ventana
        if btu == 5:
            return precio <= 150.0, 'A/A (VENTANA 5 MIL 110V)'
        elif btu == 8:
            return precio <= 200.0, 'A/A (VENTANA 8 MIL 110V)'
        elif btu == 12:
            if is_110v:
                return precio <= 270.0, 'A/A (VENTANA 12 MIL 110V)'
            return precio <= 275.0, 'A/A (VENTANA 12 MIL 220V)'
        elif btu == 18:
            return precio <= 400.0, 'A/A (VENTANA 18 MIL 220V)'
        elif btu == 24:
            return precio <= 450.0, 'A/A (VENTANA 24 MIL 220V)'
        if is_ventana:
            return precio <= 300.0, 'A/A (Ventana)'
        return False, 'No Cesta Básica'
        
    # 4. Televisores
    elif any(k in nombre for k in ['tv', 'televisor']) and 'base' not in nombre and 'soporte' not in nombre:
        size = None
        m_s = re.search(r'(\d+)\s*(?:"|\'|pulgadas|pulg)', nombre)
        if m_s:
            size = int(m_s.group(1))
            
        is_smart = any(k in nombre for k in ['smart', 'google tv', 'android tv', 'tizen', 'qled'])
        
        if size:
            if size == 32:
                if is_smart:
                    return precio <= 160.0, 'TV 32" SMART'
                return precio <= 100.0, 'TV 32" LED SLIM'
            elif size == 40:
                return precio <= 230.0, 'TV 40" SMART'
            elif size == 43:
                return precio <= 260.0, 'TV 43" SMART'
            elif size == 50:
                return precio <= 395.0, 'TV 50" SMART'
            elif size == 55:
                return precio <= 420.0, 'TV 55" SMART'
            elif size == 58:
                return precio <= 460.0, 'TV 58" SMART'
            elif size == 65:
                return precio <= 655.0, 'TV 65" SMART'
            elif size == 70:
                return precio <= 730.0, 'TV 70" SMART'
            elif size == 75:
                return precio <= 830.0, 'TV 75" SMART'
            limits = {
                32: 160.0, 40: 230.0, 43: 260.0, 50: 395.0,
                55: 420.0, 58: 460.0, 65: 655.0, 70: 730.0, 75: 830.0
            }
            closest_size = min(limits.keys(), key=lambda s: abs(s - size))
            return precio <= limits[closest_size], f'TV {closest_size}" SMART'
        return precio <= 400.0, 'TV Smart'
        
    # 5. Cocina
    elif any(k in nombre for k in ['cocina', 'tope a gas', 'tope electrico', 'campana']):
        is_campana = 'campana' in nombre
        is_gas = 'gas' in nombre
        is_elec = 'electrico' in nombre or 'electrica' in nombre
        is_dual = 'dual' in nombre
        is_tope = 'tope' in nombre
        h = None
        m_h = re.search(r'(\d+)\s*(?:h|hornillas|hornilla)', nombre)
        if m_h:
            h = int(m_h.group(1))
            
        if is_campana:
            return precio <= 180.0, 'Cocina (CAMPANA/COCINA 60CM)'
            
        if is_tope:
            if is_gas:
                if h == 5:
                    return precio <= 180.0, 'Cocina (TOPE A GAS 5H)'
                return precio <= 150.0, 'Cocina (TOPE A GAS 4H)'
            elif is_elec:
                if h == 5:
                    return precio <= 390.0, 'Cocina (TOPE ELECTRICO 5H)'
                return precio <= 210.0, 'Cocina (TOPE ELECTRICO 4H)'
            elif is_dual:
                return precio <= 220.0, 'Cocina (TOPE DUAL 4H)'
        return precio <= 250.0, 'Cocina (COCINA A GAS 4H)'
        
    # 6. Congelador
    elif 'congelador' in nombre or 'freezer' in nombre:
        lits = None
        m_l = re.search(r'(\d+)\s*(?:l|litros|lts)', nombre)
        if m_l:
            lits = int(m_l.group(1))
        if lits:
            limits = {100: 220.0, 142: 270.0, 200: 340.0}
            closest_l = min(limits.keys(), key=lambda l: abs(l - lits))
            return precio <= limits[closest_l], f'Congelador (HORIZONTAL {closest_l}LT)'
        return precio <= 250.0, 'Congelador'
        
    return False, 'No Cesta Básica'

def fetch_data(gid_base=None, force_sync=False):
    global _cache_campo_data, _cache_campo_time, _cache_maestro_data, _cache_maestro_time
    
    # Helper to fetch sheets from Google published CSV
    if gid_base is None:
        gid_base = GID_BASE_DAKA
        
    base_dir = os.path.dirname(os.path.abspath(__file__))
    cache_campo = os.path.join(base_dir, "df_campo_cache.csv")
    cache_maestro = os.path.join(base_dir, f"df_maestro_{gid_base}_cache.csv")
    
    df_campo = None
    df_maestro = None
    now = time.time()
    
    # FIX CONCURRENCIA: todo el ciclo de lectura/escritura de la caché (memoria y
    # disco) queda protegido por el mismo lock que usa /api/sync, para que dos
    # requests no puedan pisarse la caché compartida al mismo tiempo.
    with _cache_lock:
        # 1. Load or fetch df_campo
        # First check memory cache
        if not force_sync and _cache_campo_data is not None and (now - _cache_campo_time < 600.0):
            df_campo = _cache_campo_data
            
        if df_campo is None:
            try:
                req1 = urllib.request.Request(
                    URL_TIPIFICACIONES, 
                    headers={'User-Agent': 'Mozilla/5.0'}
                )
                with urllib.request.urlopen(req1, timeout=10) as response:
                    df_campo = pd.read_csv(io.StringIO(response.read().decode('utf-8')))
                df_campo.to_csv(cache_campo, index=False)
                _cache_campo_data = df_campo
                _cache_campo_time = now
            except Exception as e:
                app.logger.warning(f"Error fetching df_campo from Google Sheets: {e}. Trying to load from disk cache...")
                if os.path.exists(cache_campo):
                    try:
                        df_campo = pd.read_csv(cache_campo)
                        _cache_campo_data = df_campo
                        _cache_campo_time = now  # Avoid spamming network if it fails
                    except Exception:
                        raise e
                else:
                    raise e
                    
        # 2. Load or fetch df_maestro
        # First check memory cache
        if not force_sync and gid_base in _cache_maestro_data and (now - _cache_maestro_time.get(gid_base, 0.0) < 600.0):
            df_maestro = _cache_maestro_data[gid_base]
            
        if df_maestro is None:
            try:
                url_base = f"https://docs.google.com/spreadsheets/d/e/2PACX-1vTfq81DhLQ_8jkbFIAs7OWaO7qkYRis350TTRz_BbbsVucVw4K87Ai0YgiynRIQG1CqRJv9i1V6oEDo/pub?gid={gid_base}&single=true&output=csv"
                req2 = urllib.request.Request(
                    url_base, 
                    headers={'User-Agent': 'Mozilla/5.0'}
                )
                with urllib.request.urlopen(req2, timeout=10) as response:
                    df_maestro = pd.read_csv(io.StringIO(response.read().decode('utf-8')))
                df_maestro.to_csv(cache_maestro, index=False)
                _cache_maestro_data[gid_base] = df_maestro
                _cache_maestro_time[gid_base] = now
            except Exception as e:
                app.logger.warning(f"Error fetching df_maestro for {gid_base}: {e}. Trying to load from disk cache...")
                if os.path.exists(cache_maestro):
                    try:
                        df_maestro = pd.read_csv(cache_maestro)
                        _cache_maestro_data[gid_base] = df_maestro
                        _cache_maestro_time[gid_base] = now
                    except Exception:
                        raise e
                else:
                    raise e
                
    # Clean df_maestro: exclude rows with empty, NaN, or <= 0 price
    if 'PRECIO DE REFERENCIA' in df_maestro.columns:
        prices = df_maestro['PRECIO DE REFERENCIA'].astype(str).str.replace(' ', '').str.replace(',', '.')
        prices_numeric = pd.to_numeric(prices, errors='coerce').fillna(0.0)
        df_maestro = df_maestro[prices_numeric > 0.0].copy()
        
    return df_campo, df_maestro

# --- LÓGICA COMPARTIDA ENTRE /api/generate Y /api/generate_observations ---
# (antes esto estaba duplicado casi línea por línea en ambos endpoints; ver informe de
# auditoría, sección 4.2. Se separó en dos funciones -filtrar y luego matchear- en vez
# de una sola, porque /api/generate necesita insertar su propio paso de ordenamiento
# por hora + asignación de WS2_ROW entre el filtrado y el matching, y ese paso no le
# aplica a /api/generate_observations.)

def filtrar_datos_reporte(fecha, empresa_input, sucursal, hora_apertura_str, hora_cierre_str):
    """
    Mapea la empresa recibida al identificador interno, carga los datos (Google Sheets
    o caché) y filtra df_campo por fecha, empresa, sucursal y hora de cierre.

    No hace el matching difuso de productos contra la base maestra (eso lo hace
    `matchear_productos`, aparte), para que el caller pueda insertar pasos propios
    entre el filtrado y el matching si lo necesita.

    Devuelve un dict con el contexto necesario para continuar, o None si no hay
    registros para los parámetros dados (el caller debe responder con el 404
    correspondiente).
    """
    # Mapeamos los nombres limpios recibidos a los valores reales de la hoja de cálculo
    emp_clean = empresa_input.lower().strip()
    if 'daka' in emp_clean:
        empresa = 'ddaka'
        gid_base = GID_BASE_DAKA
    elif 'damasco' in emp_clean:
        empresa = 'ddamasco'
        gid_base = GID_BASE_DAMASCO
    elif 'multimax' in emp_clean:
        empresa = 'dmultimax'
        gid_base = GID_BASE_MULTIMAX
    else:
        empresa = emp_clean
        gid_base = GID_BASE_DAKA

    # Load sheets
    df_campo, df_maestro = fetch_data(gid_base=gid_base)

    # Mapeos de precios y marcas
    prices = df_maestro['PRECIO DE REFERENCIA'].astype(str).str.replace(' ', '').str.replace(',', '.')
    prices_numeric = pd.to_numeric(prices, errors='coerce').fillna(0.0)
    mapeo_precios = dict(zip(df_maestro['PRODUCTO'].astype(str), prices_numeric))
    mapeo_marcas = dict(zip(df_maestro['PRODUCTO'].astype(str), df_maestro['MARCA']))

    # --- FIX RENDIMIENTO ---
    # Filtramos antes de correr el matching difuso (que es lo costoso) en vez de
    # después. Ver informe de auditoría, sección 4.1.

    # Filtrar por fecha
    df_filtered = df_campo[df_campo['Fecha'].astype(str).str.strip() == str(fecha).strip()]

    # Filtrar por empresa
    target_emp = 'mmultimax' if empresa == 'dmultimax' else empresa
    df_filtered = df_filtered[df_filtered['Empresa'].astype(str).str.lower().str.strip() == target_emp.lower().strip()]

    # Filtrar por sucursal
    col_sucursal = 'Sucursal'
    if empresa.lower() == 'ddaka':
        old_col = 'Sede Daka'
    elif empresa.lower() == 'ddamasco':
        old_col = 'Sedes Damasco'
    else:
        old_col = 'Sedes Multimax' if 'Sedes Multimax' in df_campo.columns else ('Sede Multimax' if 'Sede Multimax' in df_campo.columns else 'Sede Daka')

    def row_matches(row):
        suc_val = row.get('Sucursal')
        old_val = row.get(old_col)

        target_clean = str(sucursal).lower().strip()

        if suc_val and not pd.isna(suc_val):
            suc_clean = str(suc_val).lower().strip()
            if suc_clean not in ["código no registrado", "codigo no registrado", "nan", "none", ""]:
                if suc_clean == target_clean:
                    return True

        if old_val and not pd.isna(old_val):
            old_clean = str(old_val).lower().strip().replace('_', ' ')
            if old_clean in target_clean or target_clean in old_clean:
                return True

        return False

    df_filtered = df_filtered[df_filtered.apply(row_matches, axis=1)]

    if df_filtered.empty:
        return None

    # Filtrar por hora de cierre
    limit_hour = parse_input_hour(hora_cierre_str)
    opening_hour = parse_input_hour(hora_apertura_str)

    def is_before_close(row):
        t1, t2 = parse_horario(row['Horario'])
        if t1 is None:
            return True
        return t2 <= limit_hour

    df_filtered = df_filtered[df_filtered.apply(is_before_close, axis=1)].copy()

    if df_filtered.empty:
        return None

    # FIX (bug encontrado durante pruebas): las filas placeholder "Sin productos" (una
    # hora sin ventas, presente solo para registrar Visitas) tienen 'Cantidad' en NaN.
    # Sin este fillna, VENTA_TOTAL = PRECIO_MAESTRO * NaN = NaN, y más abajo tanto
    # round_half_up(NaN) (en la proyección) como int(NaN) (al escribir la hoja 'Datos
    # Detallados') lanzan ValueError y el reporte falla cada vez que una hora del
    # estudio no tuvo ventas.
    df_filtered['Cantidad'] = df_filtered['Cantidad'].fillna(0)

    return {
        'df_filtered': df_filtered,
        'df_campo': df_campo,
        'df_maestro': df_maestro,
        'prices_numeric': prices_numeric,
        'mapeo_precios': mapeo_precios,
        'mapeo_marcas': mapeo_marcas,
        'empresa': empresa,
        'target_emp': target_emp,
        'gid_base': gid_base,
        'col_sucursal': col_sucursal,
        'limit_hour': limit_hour,
        'opening_hour': opening_hour,
    }


def matchear_productos(df_filtered, df_maestro, prices_numeric, mapeo_precios, mapeo_marcas):
    """
    Construye el universo maestro estandarizado a partir de df_maestro y corre el
    matching difuso (buscar_coincidencia_tecnica) sobre df_filtered, agregando las
    columnas PRODUCTO_CORRECTO, PRECIO_MAESTRO, MARCA_MAESTRA, VENTA_TOTAL y CATEGORIA.

    Devuelve (df_filtered_con_matching, universo_maestro).
    """
    # Estandarizar base de datos maestra
    universo_maestro = []
    for idx, row in df_maestro.iterrows():
        prod = str(row['PRODUCTO'])
        marca = str(row['MARCA']) if not pd.isna(row['MARCA']) else ""
        prod_norm = estandarizar_texto(prod)
        marca_norm = remover_tildes(marca).lower().strip()
        price_val = float(prices_numeric.loc[idx])

        row_cat = categorizar_producto(prod_norm)
        prod_type_cand = None
        if row_cat == 'NEVERA':
            prod_type_cand = 'NEVERA'
        elif row_cat == 'CONGELADOR':
            prod_type_cand = 'CONGELADOR'
        elif 'microonda' in prod_norm or 'microondas' in prod_norm:
            prod_type_cand = 'MICROONDAS'
        elif row_cat == 'TV':
            prod_type_cand = 'TV'
        elif row_cat == 'A/A':
            prod_type_cand = 'A/A'

        universo_maestro.append({
            'original': prod,
            'normalizado': prod_norm,
            'marca': marca,
            'marca_normalizada': marca_norm,
            'marca_grupo': obtener_grupo_marca(marca),
            'numeros': extraer_spec_numerica(prod_norm),
            'precio': price_val,
            'categoria': row_cat,
            'brand_norm': normalize_brands_for_matching(prod_norm),
            'is_mount': 'base' in prod_norm or 'soporte' in prod_norm,
            'is_smart': any(x in prod_norm for x in ['smart', 'google tv', 'android tv', 'tizen', 'webos', 'roku']),
            'has_led': "led" in prod.lower() or "led" in prod_norm,
            'is_basic_led': ("led" in prod.lower() or "led" in prod_norm) and not any(x in prod.lower() for x in ['smart', 'google tv', 'android tv', 'tizen', 'webos', 'roku']),
            'has_semi': any(x in prod_norm for x in ['semi', 'doble tina', 'dos tinas', 'semiautomatica']),
            'capacidades': obtener_capacidad_normalizada(prod_norm, prod_type_cand),
            'kw_groups': obtener_grupos_palabras_clave(prod_norm)
        })

    # Matching corre solo sobre df_filtered (el subconjunto del reporte), no sobre todo df_campo
    df_filtered = df_filtered.copy()

    # FIX (cosmético, hallado durante las pruebas de la sesión 2): las filas "Sin
    # productos" (una hora sin ventas, existen solo para registrar Visitas) no
    # deberían pasar por el matching difuso — a veces "empataban" por casualidad con
    # un producto real sin sentido (ej. "SILLA COMEDOR") en la columna de detalle del
    # Excel. No afectaba ningún total (Cantidad ya es 0 para estas filas gracias al
    # fillna de filtrar_datos_reporte), pero se veía mal. Se detectan aparte y se les
    # asigna un valor fijo sin pasar por el matcher.
    es_sin_productos = df_filtered['Producto'].astype(str).str.strip().str.lower() == 'sin productos'

    def _match_row(row):
        if es_sin_productos.loc[row.name]:
            return 'SIN PRODUCTOS'
        return buscar_coincidencia_tecnica(row, universo_maestro)

    df_filtered['PRODUCTO_CORRECTO'] = df_filtered.apply(_match_row, axis=1)
    df_filtered['PRECIO_MAESTRO'] = df_filtered['PRODUCTO_CORRECTO'].map(mapeo_precios).fillna(0.0).apply(round_half_up)
    df_filtered['MARCA_MAESTRA'] = df_filtered['PRODUCTO_CORRECTO'].map(mapeo_marcas).fillna("Desconocida")
    df_filtered['VENTA_TOTAL'] = df_filtered['PRECIO_MAESTRO'] * df_filtered['Cantidad']
    df_filtered['CATEGORIA'] = df_filtered['PRODUCTO_CORRECTO'].apply(categorizar_producto)

    return df_filtered, universo_maestro


# --- AUTHENTICATION ROUTES ---
@app.route('/login', methods=['GET', 'POST'])
def login_page():
    if request.method == 'POST':
        global USERS
        USERS = load_users()  # Reload to get newly added users
        
        if request.is_json:
            data = request.json
            username = data.get('username')
            password = data.get('password')
        else:
            username = request.form.get('username')
            password = request.form.get('password')

        stored = USERS.get(username)
        if verificar_password(stored, password):
            # Migración automática: si el usuario todavía tenía la contraseña en texto
            # plano (de antes de este fix), la reemplazamos por su hash ahora que
            # sabemos que es correcta. Transparente para el usuario, no requiere que
            # nadie resetee su contraseña.
            if not _es_hash(stored):
                USERS[username] = generate_password_hash(password)
                save_users(USERS)
            session['user'] = username
            if request.is_json:
                return jsonify({'success': True})
            return redirect(url_for('home'))
        else:
            error_msg = "Usuario o contraseña incorrectos"
            if request.is_json:
                return jsonify({'success': False, 'error': error_msg}), 401
            return render_template('login.html', error=error_msg)
            
    if is_logged_in():
        return redirect(url_for('home'))
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('login_page'))

def is_admin():
    return session.get('user') == 'admin'

@app.route('/admin/users', methods=['GET', 'POST'])
def admin_users():
    if not is_logged_in():
        return redirect(url_for('login_page'))
    if not is_admin():
        return redirect(url_for('home'))
        
    global USERS
    USERS = load_users()
    
    error = None
    success = None
    
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            new_username = request.form.get('new_username', '').strip()
            new_password = request.form.get('new_password', '').strip()
            
            if not new_username or not new_password:
                error = "El usuario y contraseña no pueden estar vacíos"
            elif new_username in USERS:
                error = f"El usuario '{new_username}' ya existe"
            else:
                USERS[new_username] = generate_password_hash(new_password)
                if save_users(USERS):
                    success = f"Usuario '{new_username}' creado con éxito"
                else:
                    error = "Error al guardar el usuario en el archivo"
        elif action == 'delete':
            delete_username = request.form.get('delete_username', '').strip()
            if delete_username == 'admin':
                error = "No se puede eliminar el usuario 'admin'"
            elif delete_username not in USERS:
                error = "El usuario no existe"
            else:
                del USERS[delete_username]
                if save_users(USERS):
                    success = f"Usuario '{delete_username}' eliminado con éxito"
                else:
                    error = "Error al guardar los cambios"
                    
    return render_template('admin_users.html', users=USERS, error=error, success=success)

@app.route('/admin/parametros', methods=['GET', 'POST'])
def admin_parametros():
    if not is_logged_in():
        return redirect(url_for('login_page'))
    if not is_admin():
        return redirect(url_for('home'))

    parametros = load_parametros()
    error = None
    success = None

    if request.method == 'POST':
        nuevos = {}
        try:
            for empresa in DEFAULT_PARAMETROS.keys():
                nuevos[empresa] = {}
                for clave in ('ecommerce_pct', 'al_mayor_pct', 'no_visibles_pct'):
                    campo = f'{empresa}_{clave}'
                    valor_raw = request.form.get(campo, '').strip().replace(',', '.')
                    valor = float(valor_raw)
                    if valor < 0 or valor > 100:
                        raise ValueError(f'{campo} fuera de rango')
                    # Se ingresa como porcentaje (ej. 20) y se guarda como
                    # fracción (0.20), que es como lo usa generate_report /
                    # generate_observations.
                    nuevos[empresa][clave] = round(valor / 100.0, 4)
        except (ValueError, TypeError):
            error = "Todos los valores deben ser números entre 0 y 100 (el porcentaje, sin el símbolo %)."
        else:
            if save_parametros(nuevos):
                parametros = nuevos
                success = "Parámetros actualizados con éxito. Se aplican al próximo reporte que se genere."
            else:
                error = "Error al guardar los parámetros en el archivo."

    return render_template('admin_parametros.html', parametros=parametros, error=error, success=success)

# --- API ENDPOINTS ---
@app.route('/')
def home():
    if not is_logged_in():
        return redirect(url_for('login_page'))
    return render_template('index.html')

@app.route('/api/options', methods=['GET'])
def get_options():
    if not is_logged_in():
        return jsonify({'success': False, 'error': 'No autorizado. Por favor inicie sesión.'}), 401
    try:
        df_campo, _ = fetch_data()
        dates = df_campo['Fecha'].dropna().astype(str).str.strip().unique().tolist()
        
        # We will build branches mapping: { date: { company_name: [branches] } }
        branches = {}
        for d in dates:
            df_date = df_campo[df_campo['Fecha'].astype(str).str.strip() == d]
            branches[d] = {
                'Daka': [],
                'Damasco': [],
                'Multimax': []
            }
            
            companies_raw = df_date['Empresa'].dropna().astype(str).str.strip().unique().tolist()
            for company in companies_raw:
                comp_lower = company.lower()
                if comp_lower == 'ddaka':
                    clean_name = 'Daka'
                elif comp_lower == 'ddamasco':
                    clean_name = 'Damasco'
                elif comp_lower == 'dmultimax' or 'multimax' in comp_lower:
                    clean_name = 'Multimax'
                else:
                    clean_name = company
                
                df_date_comp = df_date[df_date['Empresa'] == company]
                raw_branches = df_date_comp['Sucursal'].dropna().astype(str).str.strip().unique().tolist()
                
                comp_branches = []
                for b in raw_branches:
                    b_clean = b.strip()
                    if b_clean.lower() not in ["código no registrado", "codigo no registrado", "nan", "none", ""]:
                        if b_clean not in comp_branches:
                            comp_branches.append(b_clean)
                
                if comp_lower == 'ddaka':
                    old_col = 'Sede Daka'
                elif comp_lower == 'ddamasco':
                    old_col = 'Sedes Damasco'
                else:
                    old_col = 'Sedes Multimax' if 'Sedes Multimax' in df_campo.columns else ('Sede Multimax' if 'Sede Multimax' in df_campo.columns else 'Sede Daka')
                    
                if old_col in df_campo.columns:
                    old_vals = df_date_comp[old_col].dropna().astype(str).str.strip().unique().tolist()
                    for val in old_vals:
                        val_clean = val.strip().replace('_', ' ').upper()
                        if val_clean and val_clean.lower() not in ["nan", "none", ""]:
                            mapped = val_clean
                            if comp_lower == 'ddamasco':
                                mapped = f"{val_clean} - DAMASCO"
                            elif comp_lower == 'ddaka':
                                mapped = f"DAKA {val_clean}"
                            elif comp_lower == 'dmultimax':
                                mapped = f"MULTIMAX {val_clean}"
                                
                            exists = False
                            for existing in comp_branches:
                                if val_clean.lower() in existing.lower() or existing.lower() in val_clean.lower():
                                    exists = True
                                    break
                            if not exists:
                                comp_branches.append(mapped)
                
                if clean_name not in branches[d]:
                    branches[d][clean_name] = []
                branches[d][clean_name].extend(comp_branches)
                branches[d][clean_name] = sorted(list(set(branches[d][clean_name])))
                
        return jsonify({
            'success': True,
            # Fix: sorted() ordenaba "DD/MM/YYYY" alfabéticamente (ej. "01/07/2026" antes que
            # "30/06/2026"), no cronológicamente. Ordenamos parseando la fecha real.
            'dates': sorted(dates, key=lambda d: parse_date(d) or datetime.min),
            'companies': ['Daka', 'Damasco', 'Multimax'],
            'branches': branches
        })
    except Exception as e:
        # Fix: antes se devolvía el traceback completo de Python al cliente (se veía en el
        # toast de error). Ahora se loguea en el servidor y al usuario le llega un mensaje limpio.
        app.logger.error("Exception in get_options", exc_info=True)
        return jsonify({'success': False, 'error': f"No se pudieron cargar los parámetros: {str(e)}"}), 400

@app.route('/api/generate', methods=['POST'])
def generate_report():
    if not is_logged_in():
        return jsonify({'success': False, 'error': 'No autorizado. Por favor inicie sesión.'}), 401
    try:
        data = request.json
        fecha = data.get('fecha')
        empresa_input = data.get('empresa', '').strip()
        sucursal = data.get('sucursal')
        hora_cierre_str = data.get('hora_cierre', '19:00')
        hora_apertura_str = data.get('hora_apertura', '09:00')

        # (fix limpieza de código: esta carga + filtrado + matching estaba duplicada
        # casi línea por línea con /api/generate_observations; ver informe de
        # auditoría sección 4.2. Ahora ambos endpoints usan las mismas dos funciones.)
        ctx = filtrar_datos_reporte(fecha, empresa_input, sucursal, hora_apertura_str, hora_cierre_str)
        if ctx is None:
            return jsonify({'success': False, 'error': 'No se encontraron registros para los parámetros seleccionados.'}), 404

        df_filtered = ctx['df_filtered']
        df_maestro = ctx['df_maestro']
        prices_numeric = ctx['prices_numeric']
        mapeo_precios = ctx['mapeo_precios']
        mapeo_marcas = ctx['mapeo_marcas']
        empresa = ctx['empresa']
        col_sucursal = ctx['col_sucursal']
        limit_hour = ctx['limit_hour']
        opening_hour = ctx['opening_hour']

        # Sort chronologically by hour (earliest to latest) — esto es específico de
        # /api/generate: WS2_ROW debe reflejar el orden final en que las filas se
        # escriben en la hoja "Datos Detallados", así que el ordenamiento va antes del
        # matching (que solo agrega columnas, no reordena).
        def get_hora_orden(horario_str):
            t1, _ = parse_horario(horario_str)
            return t1 if t1 is not None else 999.0
        df_filtered['HORA_ORDEN'] = df_filtered['Horario'].apply(get_hora_orden)
        df_filtered = df_filtered.sort_values(by='HORA_ORDEN', ascending=True).drop(columns=['HORA_ORDEN']).copy()
        df_filtered['WS2_ROW'] = range(2, 2 + len(df_filtered))

        df_filtered, universo_maestro = matchear_productos(df_filtered, df_maestro, prices_numeric, mapeo_precios, mapeo_marcas)

        # Estandarizar promo y marcas
        def apply_promo(row):
            is_promo, category_spec = check_promo_status(row['PRODUCTO_CORRECTO'], row['PRECIO_MAESTRO'])
            return is_promo
            
        df_filtered['ES_PROMO'] = df_filtered.apply(apply_promo, axis=1)
        
        # Append " (PROMO)" to PRODUCTO_CORRECTO if ES_PROMO is True
        def append_promo_tag(row):
            prod = str(row['PRODUCTO_CORRECTO'])
            if row['ES_PROMO']:
                if not prod.endswith(" (PROMO)"):
                    return prod + " (PROMO)"
            return prod
        df_filtered['PRODUCTO_CORRECTO'] = df_filtered.apply(append_promo_tag, axis=1)

        # FIX (descripciones de producto, ajustado a pedido del cliente para
        # calzar con los reportes que ya hacen a mano): se limpia el texto
        # ACÁ, después de que ya se usó para buscar precio/marca/categoría y
        # para decidir si es promo (todo eso ya quedó calculado en columnas
        # separadas arriba) — así la limpieza nunca puede romper esas
        # búsquedas internas, solo afecta lo que se ve en el reporte.
        df_filtered['PRODUCTO_CORRECTO'] = df_filtered['PRODUCTO_CORRECTO'].apply(limpiar_descripcion_producto)

        df_filtered['ES_MARCA_PROPIA'] = df_filtered.apply(lambda r: is_own_brand(r['MARCA_MAESTRA'], empresa), axis=1)
        df_filtered['VENTA_PROMO'] = df_filtered['VENTA_TOTAL'].where(df_filtered['ES_PROMO'], 0.0)
        df_filtered['VENTA_FUERA_PROMO'] = df_filtered['VENTA_TOTAL'].where(~df_filtered['ES_PROMO'], 0.0)
        df_filtered['VENTA_MARCA_PROPIA'] = df_filtered['VENTA_TOTAL'].where(df_filtered['ES_MARCA_PROPIA'], 0.0)
        df_filtered['CANT_PROMO'] = df_filtered['Cantidad'].where(df_filtered['ES_PROMO'], 0)
        df_filtered['CANT_FUERA_PROMO'] = df_filtered['Cantidad'].where(~df_filtered['ES_PROMO'], 0)
        df_filtered['CANT_MARCA_PROPIA'] = df_filtered['Cantidad'].where(df_filtered['ES_MARCA_PROPIA'], 0)

        # Convert text columns in df_filtered to uppercase and remove tildes in Pandas
        text_cols = ['Producto', 'Marca', 'PRODUCTO_CORRECTO', 'MARCA_MAESTRA', 'CATEGORIA', 'Horario', 'Factura']
        for col in text_cols:
            if col in df_filtered.columns:
                df_filtered[col] = df_filtered[col].apply(lambda x: remover_tildes(str(x)).upper() if not pd.isna(x) else x)

        # --- CALCULATIONS ---
        # Table 1: Hourly metrics
        t1_data = []
        for hor, group in df_filtered.groupby('Horario', sort=False):
            prod_sold = int(group['Cantidad'].sum())
            invoices = group.dropna(subset=['Factura']).groupby(['Horario', 'Factura']).ngroups
            visits = group['Visitas'].dropna().max()
            if pd.isna(visits):
                visits = 0
            t1_data.append({
                'Horario': hor,
                'Productos Vendidos': prod_sold,
                'Facturas': invoices,
                'Visitas/Clientes': int(visits)
            })
        df_t1 = pd.DataFrame(t1_data)
        
        # Table 2: Product sales consolidation
        df_t2 = df_filtered.groupby(['PRODUCTO_CORRECTO', 'PRECIO_MAESTRO'])['Cantidad'].sum().reset_index()
        df_t2['Venta Total'] = df_t2['PRECIO_MAESTRO'] * df_t2['Cantidad']
        df_t2.columns = ['Producto', 'Precio', 'Cantidad', 'Venta Total']
        df_t2 = df_t2.sort_values(by='Cantidad', ascending=True).reset_index(drop=True)
        
        # Large sales discount
        large_invoices_sum = 0.0
        large_invoices = set()
        df_inv = df_filtered.dropna(subset=['Factura'])
        for (hor, inv), group in df_inv.groupby(['Horario', 'Factura']):
            inv_sum = (group['PRECIO_MAESTRO'] * group['Cantidad']).sum()
            if inv_sum >= 1000.0:
                large_invoices_sum += float(inv_sum)
                large_invoices.add((hor, inv))

        # Projections
        hours_times = []
        for hor in df_t1['Horario'].unique():
            t1, t2 = parse_horario(hor)
            if t1 is not None:
                hours_times.extend([t1, t2])
                
        start_study = min(hours_times) if hours_times else opening_hour
        end_study = max(hours_times) if hours_times else 17.0

        studied_hours = set()
        for hor in df_t1['Horario'].unique():
            t1, t2 = parse_horario(hor)
            if t1 is not None and t2 is not None:
                for h in range(int(t1), int(t2)):
                    studied_hours.add(h)
                    
        time_pct = calcular_time_pct(opening_hour, limit_hour, studied_hours)
        ecommerce_pct, al_mayor_pct, no_visibles_pct, has_projections = obtener_parametros_proyeccion(empresa)
            
        total_pct = time_pct + ecommerce_pct + al_mayor_pct + no_visibles_pct

        # Venta proyectada por fila (excluye facturas grandes)
        def calc_venta_proyectada(row):
            hor = row['Horario']
            inv = row['Factura']
            venta_tot = row['VENTA_TOTAL']
            if not pd.isna(inv) and (hor, inv) in large_invoices:
                return float(venta_tot)
            else:
                proj_mult = (1.0 + time_pct) * (1.0 + ecommerce_pct + al_mayor_pct + no_visibles_pct)
                return float(round_half_up(venta_tot * proj_mult))

        df_filtered['VENTA_PROYECTADA'] = df_filtered.apply(calc_venta_proyectada, axis=1)

        all_categories = [
            'A/A', 'TV', 'LAVADO', 'CONGELADOR', 'ELECTRODOMESTICOS', 
            'COCINA', 'SONIDO', 'NEVERA', 'COMPUTACION', 'TELEFONIA', 'OTROS'
        ]
        cat_sales = df_filtered.groupby('CATEGORIA')['VENTA_TOTAL'].sum().to_dict()
        df_t4_data = []
        for cat in all_categories:
            df_t4_data.append({
                'Categoría': cat,
                'Venta Total': float(cat_sales.get(cat, 0.0))
            })
        df_t4 = pd.DataFrame(df_t4_data)
        df_t4 = df_t4.sort_values(by='Venta Total', ascending=False).reset_index(drop=True)
        # --- GENERATE EXCEL IN MEMORY ---
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "EDM"
        ws1.views.sheetView[0].showGridLines = True
        
        font_title = Font(name="Segoe UI", size=16, bold=True, color="1F497D")
        font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="595959")
        font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        font_subtitles = Font(name="Segoe UI", size=11, bold=True, color="000000")
        font_data = Font(name="Segoe UI", size=10)
        font_bold = Font(name="Segoe UI", size=10, bold=True)
        
        fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        fill_zebra = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        fill_totals = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid") # Warm yellow for totals
        fill_subtitles = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # Light blue-gray for subtitles/headers
        
        thin_border = Border(
            left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
        )
        double_bottom_border = Border(
            left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000')
        )

        # Title block
        empresa_cased = empresa_input.title()
        sucursal_cased = sucursal.replace('_', ' ').title()
        ws1['A1'] = f"Reporte EDM {empresa_cased} {sucursal_cased}"
        ws1['A1'].font = font_title
        ws1['A2'] = f"Parámetros: Fecha: {fecha} | Empresa: {empresa_input.upper()} | Sucursal: {sucursal.upper()} | Apertura: {hora_apertura_str} | Cierre: {hora_cierre_str}"
        ws1['A2'].font = font_subtitle
        
        # --- TABLA 1: Rendimiento por Horario ---
        t1_title_row = 4
        ws1.cell(row=t1_title_row, column=2, value=f"EDM {empresa_input.upper()} {sucursal.upper().replace('_', ' ')} {fecha}")
        ws1.merge_cells(start_row=t1_title_row, start_column=2, end_row=t1_title_row, end_column=7)
        for col_idx in range(2, 8):
            c = ws1.cell(row=t1_title_row, column=col_idx)
            c.font = font_header
            c.fill = fill_header
            c.border = thin_border
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[t1_title_row].height = 26
        
        headers_t1 = ['HORA', 'DESCRIPCION', 'PRECIO', 'CANTIDAD', 'VENTA T', 'VISITAS']
        for col_idx, h in enumerate(headers_t1, start=2):
            cell = ws1.cell(row=t1_title_row+1, column=col_idx, value=h)
            cell.font = font_subtitles
            cell.fill = fill_subtitles
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws1.row_dimensions[t1_title_row+1].height = 24
        
        # Sort hours chronologically
        def hor_sort_key(h_str):
            t1, t2 = parse_horario(h_str)
            return t1 if t1 is not None else 0.0
        
        sorted_horarios = sorted(df_filtered['Horario'].dropna().unique(), key=hor_sort_key)
        
        t1_start_row = t1_title_row + 2
        current_row = t1_start_row
        
        for hor in sorted_horarios:
            group = df_filtered[df_filtered['Horario'] == hor]
            visits = group['Visitas'].dropna().max()
            if pd.isna(visits):
                visits = 0
            
            # Filter valid product sales (exclude "sin productos")
            valid_sales = group[group['Producto'].astype(str).str.lower().str.strip() != 'sin productos']
            
            start_slot_row = current_row
            
            if valid_sales.empty:
                ws1.cell(row=current_row, column=3, value="-").alignment = Alignment(horizontal="left")
                ws1.cell(row=current_row, column=4, value="")
                ws1.cell(row=current_row, column=5, value="")
                venta_cell = ws1.cell(row=current_row, column=6, value=0)
                venta_cell.number_format = '$#,##0'
                venta_cell.alignment = Alignment(horizontal="right")
                
                # Apply cell styles
                for col_idx in range(2, 8):
                    c = ws1.cell(row=current_row, column=col_idx)
                    c.font = font_data
                    c.border = thin_border
                    if current_row % 2 == 1:
                        c.fill = fill_zebra
                
                ws1.cell(row=current_row, column=2, value=formatear_horario_lindo(hor)).alignment = Alignment(horizontal="center", vertical="center")
                ws1.cell(row=current_row, column=7, value=int(visits)).alignment = Alignment(horizontal="center", vertical="center")
                
                current_row += 1
            else:
                for _, r_item in valid_sales.iterrows():
                    ws2_row = int(r_item['WS2_ROW'])
                    ws1.cell(row=current_row, column=3, value=f"='Datos Detallados'!F{ws2_row}").alignment = Alignment(horizontal="left")
                    
                    price_cell = ws1.cell(row=current_row, column=4, value=f"='Datos Detallados'!I{ws2_row}")
                    price_cell.number_format = '$#,##0'
                    price_cell.alignment = Alignment(horizontal="right")
                    
                    qty_cell = ws1.cell(row=current_row, column=5, value=f"='Datos Detallados'!J{ws2_row}")
                    qty_cell.alignment = Alignment(horizontal="center")
                    
                    venta_cell = ws1.cell(row=current_row, column=6, value=f"=ROUND(D{current_row}*E{current_row}, 0)")
                    venta_cell.number_format = '$#,##0'
                    venta_cell.alignment = Alignment(horizontal="right")
                    
                    # Style cells
                    for col_idx in range(2, 8):
                        c = ws1.cell(row=current_row, column=col_idx)
                        c.font = font_data
                        c.border = thin_border
                        if current_row % 2 == 1:
                            c.fill = fill_zebra
                            
                    current_row += 1
                
                end_slot_row = current_row - 1
                
                # Merge HORA and VISITAS columns
                if end_slot_row > start_slot_row:
                    ws1.merge_cells(start_row=start_slot_row, start_column=2, end_row=end_slot_row, end_column=2)
                    ws1.merge_cells(start_row=start_slot_row, start_column=7, end_row=end_slot_row, end_column=7)
                
                hora_cell = ws1.cell(row=start_slot_row, column=2, value=formatear_horario_lindo(hor))
                hora_cell.alignment = Alignment(horizontal="center", vertical="center")
                
                visits_cell = ws1.cell(row=start_slot_row, column=7, value=int(visits))
                visits_cell.alignment = Alignment(horizontal="center", vertical="center")
                
                for r in range(start_slot_row, end_slot_row + 1):
                    ws1.cell(row=r, column=2).font = font_data
                    ws1.cell(row=r, column=2).border = thin_border
                    ws1.cell(row=r, column=7).font = font_data
                    ws1.cell(row=r, column=7).border = thin_border
        
        t1_end_hour_row = current_row - 1
        
        # TOTAL GENERAL
        total_general_row = current_row
        ws1.merge_cells(start_row=total_general_row, start_column=2, end_row=total_general_row, end_column=4)
        ws1.cell(row=total_general_row, column=2, value="TOTAL GENERAL:").font = font_bold
        ws1.cell(row=total_general_row, column=2).alignment = Alignment(horizontal="left", vertical="center")
        
        ws1.cell(row=total_general_row, column=5, value=f"=SUM(E{t1_start_row}:E{t1_end_hour_row})").font = font_bold
        ws1.cell(row=total_general_row, column=6, value=f"=SUM(F{t1_start_row}:F{t1_end_hour_row})").font = font_bold
        ws1.cell(row=total_general_row, column=6).number_format = '$#,##0'
        ws1.cell(row=total_general_row, column=7, value=f"=SUM(G{t1_start_row}:G{t1_end_hour_row})").font = font_bold
        
        for col_idx in range(2, 8):
            c = ws1.cell(row=total_general_row, column=col_idx)
            c.fill = fill_totals
            c.border = thin_border
            
        current_row += 1
        
        # PROYECCION EN TIEMPO
        tiempo_row = current_row
        ws1.merge_cells(start_row=tiempo_row, start_column=2, end_row=tiempo_row, end_column=4)
        ws1.cell(row=tiempo_row, column=2, value=f"PROYECCION EN TIEMPO DE {format_hour_12h(end_study)} A {format_hour_12h(limit_hour)} ({int(time_pct*100)}%)").font = font_data
        ws1.cell(row=tiempo_row, column=2).alignment = Alignment(horizontal="left", vertical="center")
        
        ws1.cell(row=tiempo_row, column=5, value=f"=ROUND(E{total_general_row} * {time_pct}, 0)").font = font_data
        ws1.cell(row=tiempo_row, column=6, value="-").font = font_data
        ws1.cell(row=tiempo_row, column=6).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(row=tiempo_row, column=7, value=f"=ROUND(G{total_general_row} * {time_pct}, 0)").font = font_data
        
        for col_idx in range(2, 8):
            c = ws1.cell(row=tiempo_row, column=col_idx)
            c.border = thin_border
            
        current_row += 1
        
        if has_projections:
            # Ecommerce
            ecommerce_row = current_row
            ws1.merge_cells(start_row=ecommerce_row, start_column=2, end_row=ecommerce_row, end_column=4)
            ws1.cell(row=ecommerce_row, column=2, value=f"VENTAS ECOMMERCE/DELIVERY ({int(ecommerce_pct*100)}%)").font = font_data
            ws1.cell(row=ecommerce_row, column=2).alignment = Alignment(horizontal="left", vertical="center")
            
            ws1.cell(row=ecommerce_row, column=5, value=f"=ROUND(E{total_general_row} * {ecommerce_pct}, 0)").font = font_data
            ws1.cell(row=ecommerce_row, column=6, value="-").font = font_data
            ws1.cell(row=ecommerce_row, column=6).alignment = Alignment(horizontal="center", vertical="center")
            ws1.cell(row=ecommerce_row, column=7, value="").font = font_data
            
            for col_idx in range(2, 8):
                c = ws1.cell(row=ecommerce_row, column=col_idx)
                c.border = thin_border
                
            current_row += 1
            
            # Mayor
            mayor_row = current_row
            ws1.merge_cells(start_row=mayor_row, start_column=2, end_row=mayor_row, end_column=4)
            ws1.cell(row=mayor_row, column=2, value=f"VENTAS AL MAYOR ({int(al_mayor_pct*100)}%)").font = font_data
            ws1.cell(row=mayor_row, column=2).alignment = Alignment(horizontal="left", vertical="center")
            
            ws1.cell(row=mayor_row, column=5, value=f"=ROUND(E{total_general_row} * {al_mayor_pct}, 0)").font = font_data
            ws1.cell(row=mayor_row, column=6, value="-").font = font_data
            ws1.cell(row=mayor_row, column=6).alignment = Alignment(horizontal="center", vertical="center")
            ws1.cell(row=mayor_row, column=7, value="").font = font_data
            
            for col_idx in range(2, 8):
                c = ws1.cell(row=mayor_row, column=col_idx)
                c.border = thin_border
                
            current_row += 1
            
            # No Visibles
            novisibles_row = current_row
            ws1.merge_cells(start_row=novisibles_row, start_column=2, end_row=novisibles_row, end_column=4)
            ws1.cell(row=novisibles_row, column=2, value=f"VENTA DE PRODUCTOS NO VISIBLES ({int(no_visibles_pct*100)}%)").font = font_data
            ws1.cell(row=novisibles_row, column=2).alignment = Alignment(horizontal="left", vertical="center")
            
            ws1.cell(row=novisibles_row, column=5, value=f"=ROUND(E{total_general_row} * {no_visibles_pct}, 0)").font = font_data
            ws1.cell(row=novisibles_row, column=6, value="-").font = font_data
            ws1.cell(row=novisibles_row, column=6).alignment = Alignment(horizontal="center", vertical="center")
            ws1.cell(row=novisibles_row, column=7, value="").font = font_data
            
            for col_idx in range(2, 8):
                c = ws1.cell(row=novisibles_row, column=col_idx)
                c.border = thin_border
                
            current_row += 1
            
            # TOTAL PROYECTADO
            t1_total_proyectado_row = current_row
            ws1.merge_cells(start_row=t1_total_proyectado_row, start_column=2, end_row=t1_total_proyectado_row, end_column=4)
            ws1.cell(row=t1_total_proyectado_row, column=2, value="TOTAL PROYECTADO ARTICULOS Y VISITAS").font = font_bold
            ws1.cell(row=t1_total_proyectado_row, column=2).alignment = Alignment(horizontal="left", vertical="center")
            
            ws1.cell(row=t1_total_proyectado_row, column=5, value=f"=E{total_general_row}+E{tiempo_row}+E{ecommerce_row}+E{mayor_row}+E{novisibles_row}").font = font_bold
            ws1.cell(row=t1_total_proyectado_row, column=6, value="-").font = font_bold
            ws1.cell(row=t1_total_proyectado_row, column=6).alignment = Alignment(horizontal="center", vertical="center")
            ws1.cell(row=t1_total_proyectado_row, column=7, value=f"=G{total_general_row}+G{tiempo_row}").font = font_bold
            
            for col_idx in range(2, 8):
                c = ws1.cell(row=t1_total_proyectado_row, column=col_idx)
                c.fill = fill_totals
                c.border = double_bottom_border
                
            current_row += 1
        else:
            # TOTAL PROYECTADO
            t1_total_proyectado_row = current_row
            ws1.merge_cells(start_row=t1_total_proyectado_row, start_column=2, end_row=t1_total_proyectado_row, end_column=4)
            ws1.cell(row=t1_total_proyectado_row, column=2, value="TOTAL PROYECTADO ARTICULOS Y VISITAS").font = font_bold
            ws1.cell(row=t1_total_proyectado_row, column=2).alignment = Alignment(horizontal="left", vertical="center")
            
            ws1.cell(row=t1_total_proyectado_row, column=5, value=f"=E{total_general_row}+E{tiempo_row}").font = font_bold
            ws1.cell(row=t1_total_proyectado_row, column=6, value="-").font = font_bold
            ws1.cell(row=t1_total_proyectado_row, column=6).alignment = Alignment(horizontal="center", vertical="center")
            ws1.cell(row=t1_total_proyectado_row, column=7, value=f"=G{total_general_row}+G{tiempo_row}").font = font_bold
            
            for col_idx in range(2, 8):
                c = ws1.cell(row=t1_total_proyectado_row, column=col_idx)
                c.fill = fill_totals
                c.border = double_bottom_border
                
            current_row += 1
            
        # Invoices and effectiveness block
        total_invoices = df_filtered.dropna(subset=['Factura']).groupby(['Horario', 'Factura']).ngroups
        bottom_only_border = Border(bottom=Side(style='thin', color='000000'))
        no_border = Border()
        
        # 1. TOTAL FACTURAS
        tf_row = current_row
        ws1.merge_cells(start_row=tf_row, start_column=3, end_row=tf_row, end_column=5)
        ws1.cell(row=tf_row, column=3, value="TOTAL FACTURAS").font = font_bold
        ws1.cell(row=tf_row, column=3).alignment = Alignment(horizontal="center", vertical="center")
        tf_val_cell = ws1.cell(row=tf_row, column=6, value=total_invoices)
        tf_val_cell.font = font_bold
        tf_val_cell.alignment = Alignment(horizontal="center", vertical="center")
        for col_idx in range(2, 8):
            ws1.cell(row=tf_row, column=col_idx).border = no_border
        current_row += 1
        
        # 2. PROYECCION EN TIEMPO / TOTAL PROYECCION
        if has_projections:
            total_proj_pct = time_pct + ecommerce_pct + al_mayor_pct + no_visibles_pct
            pt_label = f"PROYECCION ({int(round(total_proj_pct*100))}%)"
        else:
            total_proj_pct = time_pct
            pt_label = f"PROYECCION EN TIEMPO DE {format_hour_12h(end_study)} A {format_hour_12h(limit_hour)} ({int(round(total_proj_pct*100))}%)"

        pt_row = current_row
        ws1.merge_cells(start_row=pt_row, start_column=3, end_row=pt_row, end_column=5)
        ws1.cell(row=pt_row, column=3, value=pt_label).font = font_data
        ws1.cell(row=pt_row, column=3).alignment = Alignment(horizontal="center", vertical="center")
        pt_val_cell = ws1.cell(row=pt_row, column=6, value=f"=ROUND(F{tf_row} * {total_proj_pct}, 0)")
        pt_val_cell.font = font_data
        pt_val_cell.alignment = Alignment(horizontal="center", vertical="center")
        for col_idx in range(2, 8):
            if col_idx in [3, 4, 5, 6]:
                ws1.cell(row=pt_row, column=col_idx).border = bottom_only_border
            else:
                ws1.cell(row=pt_row, column=col_idx).border = no_border
        current_row += 1
        
        # 3. NUMERO DE FACTURAS PROYECTADAS
        nfp_row = current_row
        ws1.merge_cells(start_row=nfp_row, start_column=3, end_row=nfp_row, end_column=5)
        ws1.cell(row=nfp_row, column=3, value="NUMERO DE FACTURAS PROYECTADAS").font = font_bold
        ws1.cell(row=nfp_row, column=3).alignment = Alignment(horizontal="center", vertical="center")
        nfp_val_cell = ws1.cell(row=nfp_row, column=6, value=f"=F{tf_row}+F{pt_row}")
        nfp_val_cell.font = font_bold
        nfp_val_cell.alignment = Alignment(horizontal="center", vertical="center")
        for col_idx in range(2, 8):
            ws1.cell(row=nfp_row, column=col_idx).border = no_border
        current_row += 1
        
        # 4. EFECTIVIDAD
        eff_row = current_row
        ws1.merge_cells(start_row=eff_row, start_column=3, end_row=eff_row, end_column=5)
        ws1.cell(row=eff_row, column=3, value="EFECTIVIDAD").font = font_bold
        ws1.cell(row=eff_row, column=3).alignment = Alignment(horizontal="center", vertical="center")
        
        eff_cell = ws1.cell(row=eff_row, column=6, value=f"=IF(G{t1_total_proyectado_row}>0, ROUND(F{nfp_row}/G{t1_total_proyectado_row}, 2), 0)")
        eff_cell.font = font_bold
        eff_cell.alignment = Alignment(horizontal="center", vertical="center")
        eff_cell.number_format = '0%'
        for col_idx in range(2, 8):
            ws1.cell(row=eff_row, column=col_idx).border = no_border
            
        current_row += 2 # Dejamos espacio para la Tabla 2
        
        # --- TABLA 2: Consolidado de Ventas por Producto ---
        t2_title_row = current_row
        ws1.cell(row=t2_title_row, column=3, value=f"CONSOLIDADO EDM {empresa_input.upper()} {sucursal.upper().replace('_', ' ')}")
        ws1.merge_cells(start_row=t2_title_row, start_column=3, end_row=t2_title_row, end_column=6)
        for col_idx in range(3, 7):
            c = ws1.cell(row=t2_title_row, column=col_idx)
            c.font = font_header
            c.fill = fill_header
            c.border = thin_border
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[t2_title_row].height = 26
        
        headers_t2 = ['DESCRIPCION', 'PRECIO', 'CANTIDAD', 'VENTA T']
        for col_idx, h in enumerate(headers_t2, start=3):
            cell = ws1.cell(row=t2_title_row+1, column=col_idx, value=h)
            cell.font = font_subtitles
            cell.fill = fill_subtitles
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws1.row_dimensions[t2_title_row+1].height = 24
        
        # Map PRODUCTO_CORRECTO to its first row index in Datos Detallados (starting at 2)
        prod_to_row_idx = {}
        for i, f_row in enumerate(df_filtered.itertuples()):
            p_name = getattr(f_row, 'PRODUCTO_CORRECTO')
            if p_name not in prod_to_row_idx:
                prod_to_row_idx[p_name] = i + 2
        raw_last_row = len(df_filtered) + 1

        t2_data_start = t2_title_row + 2
        for idx, row in df_t2.iterrows():
            r_idx = t2_data_start + idx
            p_name = row['Producto']
            ws2_row = prod_to_row_idx.get(p_name, 2)
            
            ws1.cell(row=r_idx, column=3, value=f"='Datos Detallados'!F{ws2_row}").alignment = Alignment(horizontal="left")
            ws1.cell(row=r_idx, column=4, value=f"='Datos Detallados'!I{ws2_row}").number_format = '$#,##0'
            ws1.cell(row=r_idx, column=4).alignment = Alignment(horizontal="right")
            ws1.cell(row=r_idx, column=5, value=f"=SUMIF('Datos Detallados'!F$2:F${raw_last_row}, C{r_idx}, 'Datos Detallados'!J$2:J${raw_last_row})").alignment = Alignment(horizontal="right")
            ws1.cell(row=r_idx, column=6, value=f"=ROUND(D{r_idx} * E{r_idx}, 0)").number_format = '$#,##0'
            ws1.cell(row=r_idx, column=6).alignment = Alignment(horizontal="right")
            for col_idx in range(3, 7):
                c = ws1.cell(row=r_idx, column=col_idx)
                c.font = font_data
                c.border = thin_border
                if r_idx % 2 == 1:
                    c.fill = fill_zebra
                    
        t2_last_prod_row = t2_data_start + len(df_t2) - 1
        
        # TOTAL CONTEO (TABLA 2)
        current_t2_row = t2_last_prod_row + 1
        ws1.cell(row=current_t2_row, column=3, value="TOTAL CONTEO").font = font_bold
        ws1.cell(row=current_t2_row, column=5, value=f"=SUM(E{t2_data_start}:E{t2_last_prod_row})").font = font_bold
        ws1.cell(row=current_t2_row, column=6, value=f"=ROUND(SUM(F{t2_data_start}:F{t2_last_prod_row}), 0)").font = font_bold
        ws1.cell(row=current_t2_row, column=6).number_format = '$#,##0'
        for col_idx in range(3, 7):
            ws1.cell(row=current_t2_row, column=col_idx).fill = fill_totals
        for col_idx in range(3, 7):
            ws1.cell(row=current_t2_row, column=col_idx).border = thin_border
        total_conteo_t2 = current_t2_row
        current_t2_row += 1
        
        # DESCUENTO VENTA GRANDE
        ws1.cell(row=current_t2_row, column=3, value="DESCUENTO EN PROYECCION POR VENTA GRANDE").font = font_bold
        ws1.cell(row=current_t2_row, column=6, value=float(round_half_up(large_invoices_sum))).font = font_bold
        ws1.cell(row=current_t2_row, column=6).number_format = '$#,##0'
        for col_idx in range(3, 7):
            ws1.cell(row=current_t2_row, column=col_idx).border = thin_border
        descuento_row_t2 = current_t2_row
        current_t2_row += 1
        
        # TOTAL NETO
        ws1.cell(row=current_t2_row, column=3, value="TOTAL NETO PARA PROYECCION DE TIEMPO").font = font_bold
        ws1.cell(row=current_t2_row, column=6, value=f"=ROUND(F{total_conteo_t2}-F{descuento_row_t2}, 0)").font = font_bold
        ws1.cell(row=current_t2_row, column=6).number_format = '$#,##0'
        for col_idx in range(3, 7):
            ws1.cell(row=current_t2_row, column=col_idx).fill = fill_totals
        for col_idx in range(3, 7):
            ws1.cell(row=current_t2_row, column=col_idx).border = thin_border
        neto_row_t2 = current_t2_row
        current_t2_row += 1
        
        # TIEMPO
        ws1.cell(row=current_t2_row, column=3, value=f"TIEMPO DE {format_hour_12h(end_study)} A {format_hour_12h(limit_hour)}").font = font_data
        ws1.cell(row=current_t2_row, column=5, value=f"+ {int(time_pct*100)}%").font = font_data
        ws1.cell(row=current_t2_row, column=6, value=f"=ROUND(F{neto_row_t2} * {time_pct}, 0)").font = font_data
        ws1.cell(row=current_t2_row, column=6).number_format = '$#,##0'
        for col_idx in range(3, 7):
            ws1.cell(row=current_t2_row, column=col_idx).border = thin_border
        tiempo_row_t2 = current_t2_row
        current_t2_row += 1
        
        # TOTAL GENERAL CONTEO (TABLA 2)
        ws1.cell(row=current_t2_row, column=3, value="TOTAL GENERAL CONTEO").font = font_bold
        ws1.cell(row=current_t2_row, column=6, value=f"=ROUND(F{neto_row_t2}+F{tiempo_row_t2}, 0)").font = font_bold
        ws1.cell(row=current_t2_row, column=6).number_format = '$#,##0'
        border_total_general = Border(
            left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'), bottom=Side(style='double' if not has_projections else 'thin', color='000000')
        )
        for col_idx in range(3, 7):
            c = ws1.cell(row=current_t2_row, column=col_idx)
            c.fill = fill_totals
            c.border = border_total_general
        total_general_conteo_row_t2 = current_t2_row
        
        if has_projections:
            current_t2_row += 2
            
            # Extras (Daka/Damasco)
            ws1.cell(row=current_t2_row, column=3, value="VENTAS ECOMMERCE/DELIVERY").font = font_data
            ws1.cell(row=current_t2_row, column=5, value=f"+ {int(ecommerce_pct*100)}%").font = font_data
            ws1.cell(row=current_t2_row, column=6, value=f"=ROUND(F{total_general_conteo_row_t2} * {ecommerce_pct}, 0)").font = font_data
            ws1.cell(row=current_t2_row, column=6).number_format = '$#,##0'
            for col_idx in range(3, 7):
                ws1.cell(row=current_t2_row, column=col_idx).border = thin_border
            ecommerce_val_row = current_t2_row
            current_t2_row += 1
            
            ws1.cell(row=current_t2_row, column=3, value="VENTAS AL MAYOR").font = font_data
            ws1.cell(row=current_t2_row, column=5, value=f"+ {int(al_mayor_pct*100)}%").font = font_data
            ws1.cell(row=current_t2_row, column=6, value=f"=ROUND(F{total_general_conteo_row_t2} * {al_mayor_pct}, 0)").font = font_data
            ws1.cell(row=current_t2_row, column=6).number_format = '$#,##0'
            for col_idx in range(3, 7):
                ws1.cell(row=current_t2_row, column=col_idx).border = thin_border
            mayor_val_row = current_t2_row
            current_t2_row += 1
            
            ws1.cell(row=current_t2_row, column=3, value="VENTA DE PRODUCTOS NO VISIBLES").font = font_data
            ws1.cell(row=current_t2_row, column=5, value=f"+ {int(no_visibles_pct*100)}%").font = font_data
            ws1.cell(row=current_t2_row, column=6, value=f"=ROUND(F{total_general_conteo_row_t2} * {no_visibles_pct}, 0)").font = font_data
            ws1.cell(row=current_t2_row, column=6).number_format = '$#,##0'
            for col_idx in range(3, 7):
                ws1.cell(row=current_t2_row, column=col_idx).border = thin_border
            novisibles_val_row = current_t2_row
            current_t2_row += 1
            
            # TOTAL PROYECCION DE VENTA (TABLA 2 FINAL)
            ws1.cell(row=current_t2_row, column=3, value="TOTAL PROYECCION DE VENTA").font = font_bold
            ws1.cell(row=current_t2_row, column=5, value=f"{int(total_pct*100)}%").font = font_bold
            ws1.cell(row=current_t2_row, column=6, value=f"=ROUND(F{total_conteo_t2}+F{tiempo_row_t2}+F{ecommerce_val_row}+F{mayor_val_row}+F{novisibles_val_row}, 0)").font = font_bold
            ws1.cell(row=current_t2_row, column=6).number_format = '$#,##0'
            for col_idx in range(3, 7):
                ws1.cell(row=current_t2_row, column=col_idx).fill = fill_totals
            for col_idx in range(3, 7):
                ws1.cell(row=current_t2_row, column=col_idx).border = double_bottom_border

        # --- TABLA 3: Resumen Promo / Fuera de Promo / Marcas ---
        # Ubicada al centro superior: Columnas I-L, filas 2-4
        ws1.cell(row=2, column=9, value=f"{empresa_input.upper()} {sucursal.upper().replace('_', ' ')}")
        ws1.merge_cells(start_row=2, start_column=9, end_row=2, end_column=12)
        for col_idx in range(9, 13):
            c = ws1.cell(row=2, column=col_idx)
            c.font = font_header
            c.fill = fill_header
            c.border = thin_border
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[2].height = 26
        
        headers_t3 = ['PROMO', 'FUERA DE PROMO', 'MARCAS', 'TOTAL']
        for idx, h in enumerate(headers_t3, start=9):
            cell = ws1.cell(row=3, column=idx, value=h)
            cell.font = font_subtitles
            cell.fill = fill_subtitles
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws1.row_dimensions[3].height = 24
        
        raw_last_row = len(df_filtered) + 1
        ws1.cell(row=4, column=9, value=f"=SUMIF('Datos Detallados'!N$2:N${raw_last_row}, \"SI\", 'Datos Detallados'!K$2:K${raw_last_row})").number_format = '$#,##0'
        ws1.cell(row=4, column=10, value=f"=SUMIF('Datos Detallados'!N$2:N${raw_last_row}, \"NO\", 'Datos Detallados'!K$2:K${raw_last_row})").number_format = '$#,##0'
        ws1.cell(row=4, column=11, value=f"=SUMIF('Datos Detallados'!O$2:O${raw_last_row}, \"SI\", 'Datos Detallados'!K$2:K${raw_last_row})").number_format = '$#,##0'
        ws1.cell(row=4, column=12, value=f"=I4+J4").number_format = '$#,##0'
        
        for col_idx in range(9, 13):
            c = ws1.cell(row=4, column=col_idx)
            c.font = font_bold
            c.alignment = Alignment(horizontal="right")
            c.border = thin_border
            c.fill = fill_zebra
        ws1.row_dimensions[4].height = 22

        # --- TABLA 4: Ventas por Categoría ---
        # Ubicada a la derecha superior: Columnas N-O, filas 2-15
        ws1.cell(row=2, column=14, value=f"{empresa_input.upper()} {sucursal.upper().replace('_', ' ')}")
        ws1.merge_cells(start_row=2, start_column=14, end_row=2, end_column=15)
        for col_idx in range(14, 16):
            c = ws1.cell(row=2, column=col_idx)
            c.font = font_header
            c.fill = fill_header
            c.border = thin_border
            c.alignment = Alignment(horizontal="center", vertical="center")
        
        headers_t4 = ['CATEGORIAS', 'VALOR']
        for idx, h in enumerate(headers_t4, start=14):
            cell = ws1.cell(row=3, column=idx, value=h)
            cell.font = font_subtitles
            cell.fill = fill_subtitles
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws1.row_dimensions[3].height = 24
        
        data_start_t4 = 4
        for idx, row in df_t4.iterrows():
            r_idx = data_start_t4 + idx
            ws1.cell(row=r_idx, column=14, value=row['Categoría']).alignment = Alignment(horizontal="left")
            ws1.cell(row=r_idx, column=15, value=f"=SUMIF('Datos Detallados'!H$2:H${raw_last_row}, N{r_idx}, 'Datos Detallados'!K$2:K${raw_last_row})").number_format = '$#,##0'
            ws1.cell(row=r_idx, column=15).alignment = Alignment(horizontal="right")
            for col_idx in range(14, 16):
                c = ws1.cell(row=r_idx, column=col_idx)
                c.font = font_bold if col_idx == 14 else font_data
                c.border = thin_border
                if r_idx % 2 == 1:
                    c.fill = fill_zebra
                    
        t4_end_row = data_start_t4 + len(df_t4) - 1
        
        ws1.cell(row=t4_end_row+1, column=14, value="TOTAL").font = font_bold
        ws1.cell(row=t4_end_row+1, column=15, value=f"=SUM(O{data_start_t4}:O{t4_end_row})").font = font_bold
        ws1.cell(row=t4_end_row+1, column=15).number_format = '$#,##0'
        for col_idx in range(14, 16):
            ws1.cell(row=t4_end_row+1, column=col_idx).fill = fill_totals
            ws1.cell(row=t4_end_row+1, column=col_idx).border = double_bottom_border

        # Adjust widths
        for col in ws1.columns:
            col_letter = get_column_letter(col[0].column)
            if col_letter in ['H', 'M', 'P']:
                ws1.column_dimensions[col_letter].width = 4
                continue
            max_len = 0
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws1.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 30)
        ws1.column_dimensions['A'].width = 4
        ws1.column_dimensions['B'].width = 18
        ws1.column_dimensions['C'].width = 35
        ws1.column_dimensions['D'].width = 12
        ws1.column_dimensions['E'].width = 12
        ws1.column_dimensions['F'].width = 15
        ws1.column_dimensions['G'].width = 12
        ws1.column_dimensions['H'].width = 4
        ws1.column_dimensions['I'].width = 15
        ws1.column_dimensions['J'].width = 18
        ws1.column_dimensions['K'].width = 15
        ws1.column_dimensions['L'].width = 15
        ws1.column_dimensions['M'].width = 4
        ws1.column_dimensions['N'].width = 25
        ws1.column_dimensions['O'].width = 15

        # --- ENCABEZADO DE SECCIÓN (separa visualmente las tablas de los gráficos) ---
        ws1.cell(row=5, column=9, value="GRÁFICAS DE DESEMPEÑO")
        ws1.merge_cells(start_row=5, start_column=9, end_row=5, end_column=12)
        c_sec = ws1.cell(row=5, column=9)
        c_sec.font = font_subtitles
        c_sec.fill = fill_subtitles
        c_sec.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[5].height = 22

        # --- ADD CHARTS (Nativos) ---
        # 1. Bar Chart of Promo, Fuera de Promo, Marcas
        chart_bar = BarChart()
        chart_bar.type = "col"
        chart_bar.style = 2
        chart_bar.title = f"{empresa_input.upper()} {sucursal.upper().replace('_', ' ')}"
        chart_bar.legend = None # No legend for single-series bar chart as in Imagen 1
        
        # I4:K4 data (row 4 values, row 3 headers)
        data_bar = Reference(ws1, min_col=9, max_col=11, min_row=4, max_row=4)
        cats_bar = Reference(ws1, min_col=9, max_col=11, min_row=3, max_row=3)
        chart_bar.add_data(data_bar, from_rows=True)
        chart_bar.set_categories(cats_bar)
        
        from openpyxl.chart.series import DataPoint
        # NOTA (ajustado a pedido del cliente, para calzar exacto con el reporte
        # de referencia que ya usan): las 3 barras van todas del mismo azul
        # (accent1 del tema de Excel) — así es como sale en el reporte que
        # hacen a mano, sin distinguir colores por barra. No se le agrega
        # coloreado individual a propósito.
            
        # Remove gridlines
        chart_bar.y_axis.majorGridlines = None
        chart_bar.x_axis.majorGridlines = None
        
        chart_bar.dataLabels = DataLabelList()
        chart_bar.dataLabels.showSerName = False
        chart_bar.dataLabels.showCatName = False
        chart_bar.dataLabels.showVal = True
        chart_bar.dataLabels.showPercent = False
        chart_bar.width = 11
        chart_bar.height = 8.5
        ws1.add_chart(chart_bar, "I6")

        # 2. Pie Chart of Promo vs Fuera de Promo
        chart_pie_promo = PieChart()
        chart_pie_promo.title = "FUERA DE PROMO / PROMO"
        
        # I4:J4 data, I3:J3 headers (Col 9 to 10)
        data_pie = Reference(ws1, min_col=9, max_col=10, min_row=4, max_row=4)
        cats_pie = Reference(ws1, min_col=9, max_col=10, min_row=3, max_row=3)
        chart_pie_promo.add_data(data_pie, from_rows=True)
        chart_pie_promo.set_categories(cats_pie)
        
        # Assign individual slice colors: #4472C4 (Promo) y #ED7D31 (Fuera de Promo)
        # (ajustado a pedido del cliente — el reporte de referencia usa el azul
        # para Promo y el naranja para Fuera de Promo, no al revés)
        slice_colors = ["4472C4", "ED7D31"]
        series_pie = chart_pie_promo.series[0]
        for idx, color in enumerate(slice_colors):
            dp = DataPoint(idx=idx)
            dp.graphicalProperties.solidFill = color
            series_pie.dPt.append(dp)

        chart_pie_promo.dataLabels = DataLabelList()
        chart_pie_promo.dataLabels.showSerName = False
        chart_pie_promo.dataLabels.showCatName = False
        chart_pie_promo.dataLabels.showVal = False
        chart_pie_promo.dataLabels.showPercent = True
        chart_pie_promo.width = 11
        chart_pie_promo.height = 8.5
        ws1.add_chart(chart_pie_promo, "I24")

        # 3. Pie Chart of Categories
        chart_pie_cat = PieChart()
        chart_pie_cat.title = f"{empresa_input.upper()} {sucursal.upper().replace('_', ' ')}"
        chart_pie_cat.legend = None  # No legend, labels are directly on/outside slices
        
        # O4:O{t4_end_row} data, N4:N{t4_end_row} categories (Col 15 and 14)
        data_cat = Reference(ws1, min_col=15, min_row=3, max_row=t4_end_row)
        cats_cat = Reference(ws1, min_col=14, min_row=4, max_row=t4_end_row)
        chart_pie_cat.add_data(data_cat, titles_from_data=True)
        chart_pie_cat.set_categories(cats_cat)

        # FIX (organización/estética de gráficas, ajustado a pedido del cliente):
        # antes este gráfico no tenía colores propios (paleta por defecto de
        # Excel, impredecible). Se probó primero con un color fijo por NOMBRE
        # de categoría, pero el reporte de referencia usa los 6 colores de
        # acento del tema por POSICIÓN en la tabla (que está ordenada por
        # venta) — se replica ese comportamiento exacto acá, ciclando la
        # paleta si hay más de 11 categorías con datos.
        series_cat = chart_pie_cat.series[0]
        for idx, row in df_t4.iterrows():
            color = PALETA_CATEGORIAS_POSICION[idx % len(PALETA_CATEGORIAS_POSICION)]
            dp = DataPoint(idx=idx)
            dp.graphicalProperties.solidFill = color
            series_cat.dPt.append(dp)
        
        # Show category name and percentage with leader lines on/outside slices
        chart_pie_cat.dataLabels = DataLabelList()
        chart_pie_cat.dataLabels.showSerName = False
        chart_pie_cat.dataLabels.showCatName = True
        chart_pie_cat.dataLabels.showVal = False
        chart_pie_cat.dataLabels.showPercent = True
        chart_pie_cat.dataLabels.showLeaderLines = True
        
        # Sized nicely to give Excel space for leader lines and prevent label overlap
        chart_pie_cat.width = 12
        chart_pie_cat.height = 10.5
        # FIX (organización/estética de gráficas): antes empezaba en la fila 2,
        # mientras que el gráfico de barras de al lado empieza en la fila 6 —
        # quedaban desalineados verticalmente. Ahora los dos arrancan en la
        # misma fila.
        ws1.add_chart(chart_pie_cat, "Q6")

        # Sheet 2: Raw data
        ws2 = wb.create_sheet(title="Datos Detallados")
        ws2.views.sheetView[0].showGridLines = True
        
        headers_t6 = ['ID Envio', 'Fecha', 'Horario', 'Producto Campo', 'Marca Campo', 'Producto Estandarizado', 'Marca DB', 'Categoría', 'Precio ($)', 'Cantidad', 'Venta Total ($)', 'Factura', 'Visitas', 'Es Promo', 'Marca Propia', 'Venta Proyectada ($)']
        for col_idx, h in enumerate(headers_t6, start=1):
            cell = ws2.cell(row=1, column=col_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center")
        ws2.row_dimensions[1].height = 24
        
        raw_row_idx = 2
        for idx, row in df_filtered.iterrows():
            ws2.cell(row=raw_row_idx, column=1, value=row['ID_Envio']).alignment = Alignment(horizontal="center")
            ws2.cell(row=raw_row_idx, column=2, value=row['Fecha']).alignment = Alignment(horizontal="center")
            ws2.cell(row=raw_row_idx, column=3, value=formatear_horario_lindo(row['Horario'])).alignment = Alignment(horizontal="center")
            ws2.cell(row=raw_row_idx, column=4, value=row['Producto']).alignment = Alignment(horizontal="left")
            ws2.cell(row=raw_row_idx, column=5, value=row['Marca']).alignment = Alignment(horizontal="left")
            ws2.cell(row=raw_row_idx, column=6, value=row['PRODUCTO_CORRECTO']).alignment = Alignment(horizontal="left")
            ws2.cell(row=raw_row_idx, column=7, value=row['MARCA_MAESTRA']).alignment = Alignment(horizontal="left")
            ws2.cell(row=raw_row_idx, column=8, value=row['CATEGORIA']).alignment = Alignment(horizontal="left")
            ws2.cell(row=raw_row_idx, column=9, value=round_half_up(float(row['PRECIO_MAESTRO']))).number_format = '$#,##0'
            ws2.cell(row=raw_row_idx, column=9).alignment = Alignment(horizontal="right")
            ws2.cell(row=raw_row_idx, column=10, value=int(row['Cantidad'])).alignment = Alignment(horizontal="right")
            ws2.cell(row=raw_row_idx, column=11, value=f"=ROUND(I{raw_row_idx} * J{raw_row_idx}, 0)").number_format = '$#,##0'
            ws2.cell(row=raw_row_idx, column=11).alignment = Alignment(horizontal="right")
            
            fact_val = row['Factura']
            if not pd.isna(fact_val):
                try:
                    fact_val = int(float(fact_val))
                except Exception:
                    pass
            ws2.cell(row=raw_row_idx, column=12, value=fact_val).alignment = Alignment(horizontal="center")
            ws2.cell(row=raw_row_idx, column=13, value=row['Visitas']).alignment = Alignment(horizontal="center")
            ws2.cell(row=raw_row_idx, column=14, value="SÍ" if row['ES_PROMO'] else "NO").alignment = Alignment(horizontal="center")
            ws2.cell(row=raw_row_idx, column=15, value="SÍ" if row['ES_MARCA_PROPIA'] else "NO").alignment = Alignment(horizontal="center")
            ws2.cell(row=raw_row_idx, column=16, value=round_half_up(float(row['VENTA_PROYECTADA']))).number_format = '$#,##0'
            ws2.cell(row=raw_row_idx, column=16).alignment = Alignment(horizontal="right")
            
            for col_idx in range(1, 17):
                c = ws2.cell(row=raw_row_idx, column=col_idx)
                c.font = font_data
                c.border = thin_border
                if raw_row_idx % 2 == 1:
                    c.fill = fill_zebra
            raw_row_idx += 1
 
        for col in ws2.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws2.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 30)

        # --- Pestaña de DESCUENTO si aplica ---
        if large_invoices:
            ws_desc = wb.create_sheet(title="DESCUENTO")
            ws_desc.views.sheetView[0].showGridLines = True
            
            # Title block in row 1
            ws_desc.merge_cells("A1:E1")
            for col in range(1, 6):
                c = ws_desc.cell(row=1, column=col)
                c.fill = fill_header
                c.border = thin_border
            title_text = f"DESCUENTOS POR VENTAS GRANDES {empresa_input.upper()} {sucursal.upper().replace('_', ' ')} {fecha}"
            title_cell = ws_desc.cell(row=1, column=1, value=title_text)
            title_cell.font = font_header
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws_desc.row_dimensions[1].height = 26
            
            headers_desc = ['HORA', 'DESCRIPCION', 'PRECIO', 'CANTIDAD', 'VENTA T']
            for col_idx, h in enumerate(headers_desc, start=1):
                cell = ws_desc.cell(row=2, column=col_idx, value=h)
                cell.font = font_subtitles
                cell.fill = fill_subtitles
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
            ws_desc.row_dimensions[2].height = 24
            
            # Gather and sort large invoices
            def sort_key(item):
                hor, inv = item
                t1, t2 = parse_horario(hor)
                t1_val = t1 if t1 is not None else 0.0
                try:
                    inv_val = float(inv)
                except ValueError:
                    inv_val = str(inv)
                return (t1_val, inv_val)
            sorted_large_invoices = sorted(list(large_invoices), key=sort_key)
            
            row_idx = 3
            totalizer_rows = []
            fill_grand_total = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
            
            for hor, inv in sorted_large_invoices:
                invoice_items = df_filtered[(df_filtered['Horario'] == hor) & (df_filtered['Factura'] == inv)]
                start_row = row_idx
                end_row = row_idx + len(invoice_items) - 1
                
                # Write each product
                for i, (_, item) in enumerate(invoice_items.iterrows()):
                    ws_desc.cell(row=row_idx, column=2, value=item['PRODUCTO_CORRECTO']).alignment = Alignment(horizontal="left")
                    price_cell = ws_desc.cell(row=row_idx, column=3, value=round_half_up(float(item['PRECIO_MAESTRO'])))
                    price_cell.number_format = '$#,##0'
                    price_cell.alignment = Alignment(horizontal="right")
                    
                    qty_cell = ws_desc.cell(row=row_idx, column=4, value=int(item['Cantidad']))
                    qty_cell.alignment = Alignment(horizontal="center")
                    
                    venta_cell = ws_desc.cell(row=row_idx, column=5, value=f"=C{row_idx}*D{row_idx}")
                    venta_cell.number_format = '$#,##0'
                    venta_cell.alignment = Alignment(horizontal="right")
                    
                    # Style cells
                    for col in range(1, 6):
                        c = ws_desc.cell(row=row_idx, column=col)
                        c.font = font_data
                        c.border = thin_border
                    
                    row_idx += 1
                
                # Merge HORA column
                if len(invoice_items) > 1:
                    ws_desc.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
                
                # Set value in merged cell
                hora_cell = ws_desc.cell(row=start_row, column=1, value=formatear_horario_lindo(hor))
                hora_cell.alignment = Alignment(horizontal="center", vertical="center")
                
                # Apply styles (font, borders) to all merged cells in column 1
                for r in range(start_row, end_row + 1):
                    ws_desc.cell(row=r, column=1).font = font_data
                    ws_desc.cell(row=r, column=1).border = thin_border
                
                # Totalizer row for this invoice
                ws_desc.cell(row=row_idx, column=4, value="TOTALIZA:").alignment = Alignment(horizontal="right")
                tot_formula = ws_desc.cell(row=row_idx, column=5, value=f"=SUM(E{start_row}:E{end_row})")
                tot_formula.number_format = '$#,##0'
                tot_formula.alignment = Alignment(horizontal="right")
                
                totalizer_rows.append(row_idx)
                
                # Style the totalizer row
                ws_desc.row_dimensions[row_idx].height = 20
                for col in range(1, 6):
                    c = ws_desc.cell(row=row_idx, column=col)
                    c.font = font_bold
                    c.fill = fill_totals
                    c.border = thin_border
                
                row_idx += 1
                
            # Grand Total row at the bottom
            ws_desc.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
            for col in range(1, 5):
                c = ws_desc.cell(row=row_idx, column=col)
                c.fill = fill_grand_total
                c.border = double_bottom_border
                
            gt_label = ws_desc.cell(row=row_idx, column=1, value="TOTAL DESCUENTO POR VENTAS GRANDES")
            gt_label.font = font_bold
            gt_label.alignment = Alignment(horizontal="center", vertical="center")
            
            # E{row_idx} is the sum of all individual totalizers
            if totalizer_rows:
                gt_formula_str = "=" + "+".join(f"E{r}" for r in totalizer_rows)
            else:
                gt_formula_str = "=0"
            gt_val = ws_desc.cell(row=row_idx, column=5, value=gt_formula_str)
            gt_val.font = font_bold
            gt_val.fill = fill_grand_total
            gt_val.border = double_bottom_border
            gt_val.number_format = '$#,##0'
            gt_val.alignment = Alignment(horizontal="right")
            
            ws_desc.row_dimensions[row_idx].height = 22
            
            # Set dimensions
            ws_desc.column_dimensions['A'].width = 20
            ws_desc.column_dimensions['B'].width = 45
            ws_desc.column_dimensions['C'].width = 12
            ws_desc.column_dimensions['D'].width = 12
            ws_desc.column_dimensions['E'].width = 15

        # Convert all string cell values to uppercase and strip accents
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    val = cell.value
                    if isinstance(val, str):
                        if not val.startswith("="):
                            cell.value = remover_tildes(val).upper()

        # Save workbook to memory stream
        output = io.BytesIO()
        wb.save(output)
        wb.close()
        
        filename = f"Reporte_Auditoria_{empresa.upper()}_{sucursal.upper()}_{fecha.replace('/', '-')}.xlsx"
        
        # Force garbage collection to free RAM on Render
        import gc
        gc.collect()
        
        output.seek(0)
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        import gc
        gc.collect()
        # Fix: el traceback completo ya no se envía al cliente (se veía en el toast de
        # error); se loguea en el servidor y el usuario recibe un mensaje claro.
        app.logger.error("Exception in generate_report", exc_info=True)
        return jsonify({'success': False, 'error': f"No se pudo generar el reporte: {str(e)}"}), 400

@app.route('/api/generate_observations', methods=['POST'])
def generate_observations():
    if not is_logged_in():
        return jsonify({'success': False, 'error': 'No autorizado. Por favor inicie sesión.'}), 401
    try:
        data = request.json
        fecha = data.get('fecha')
        empresa_input = data.get('empresa', '').strip()
        sucursal = data.get('sucursal')
        hora_cierre_str = data.get('hora_cierre', '19:00')
        hora_apertura_str = data.get('hora_apertura', '09:00')

        # (fix limpieza de código: mismo criterio que en /api/generate — ver informe de
        # auditoría sección 4.2.)
        ctx = filtrar_datos_reporte(fecha, empresa_input, sucursal, hora_apertura_str, hora_cierre_str)
        if ctx is None:
            return jsonify({'success': False, 'error': 'No se encontraron registros para los parámetros seleccionados.'}), 404

        df_filtered = ctx['df_filtered']
        df_campo = ctx['df_campo']
        df_maestro = ctx['df_maestro']
        prices_numeric = ctx['prices_numeric']
        mapeo_precios = ctx['mapeo_precios']
        mapeo_marcas = ctx['mapeo_marcas']
        empresa = ctx['empresa']
        target_emp = ctx['target_emp']
        col_sucursal = ctx['col_sucursal']
        limit_hour = ctx['limit_hour']
        opening_hour = ctx['opening_hour']

        df_filtered, universo_maestro = matchear_productos(df_filtered, df_maestro, prices_numeric, mapeo_precios, mapeo_marcas)

        df_filtered['ES_MARCA_PROPIA'] = df_filtered.apply(lambda r: is_own_brand(r['MARCA_MAESTRA'], empresa), axis=1)
        df_filtered['VENTA_MARCA_PROPIA'] = df_filtered['VENTA_TOTAL'].where(df_filtered['ES_MARCA_PROPIA'], 0.0)
        df_filtered['CANT_MARCA_PROPIA'] = df_filtered['Cantidad'].where(df_filtered['ES_MARCA_PROPIA'], 0)
        
        # Promo vs Non-promo observed sales
        df_filtered['ES_PROMO'] = df_filtered.apply(lambda r: check_promo_status(r['PRODUCTO_CORRECTO'], r['PRECIO_MAESTRO'])[0], axis=1)
        
        # Append " (PROMO)" to PRODUCTO_CORRECTO if ES_PROMO is True
        def append_promo_tag(row):
            prod = str(row['PRODUCTO_CORRECTO'])
            if row['ES_PROMO']:
                if not prod.endswith(" (PROMO)"):
                    return prod + " (PROMO)"
            return prod
        df_filtered['PRODUCTO_CORRECTO'] = df_filtered.apply(append_promo_tag, axis=1)

        # FIX (descripciones de producto, mismo criterio que en /api/generate):
        # limpiar acá, después de que ya se usó el texto crudo para las
        # búsquedas internas (precio, marca, categoría, promo).
        df_filtered['PRODUCTO_CORRECTO'] = df_filtered['PRODUCTO_CORRECTO'].apply(limpiar_descripcion_producto)

        df_filtered['VENTA_PROMO'] = df_filtered['VENTA_TOTAL'].where(df_filtered['ES_PROMO'], 0.0)
        df_filtered['VENTA_FUERA_PROMO'] = df_filtered['VENTA_TOTAL'].where(~df_filtered['ES_PROMO'], 0.0)

        # Table 1: Hourly metrics
        t1_data = []
        for hor, group in df_filtered.groupby('Horario', sort=False):
            prod_sold = int(group['Cantidad'].sum())
            invoices = group.dropna(subset=['Factura']).groupby(['Horario', 'Factura']).ngroups
            visits = group['Visitas'].dropna().max()
            if pd.isna(visits):
                visits = 0
            t1_data.append({
                'Horario': hor,
                'Productos Vendidos': prod_sold,
                'Facturas': invoices,
                'Visitas/Clientes': int(visits)
            })
        df_t1 = pd.DataFrame(t1_data)

        # Projections
        studied_hours = set()
        for hor in df_t1['Horario'].unique():
            t1, t2 = parse_horario(hor)
            if t1 is not None and t2 is not None:
                for h in range(int(t1), int(t2)):
                    studied_hours.add(h)
                    
        time_pct = calcular_time_pct(opening_hour, limit_hour, studied_hours)
        ecommerce_pct, al_mayor_pct, no_visibles_pct, has_projections = obtener_parametros_proyeccion(empresa)
        
        # Calculate effectiveness metrics
        c_general = df_t1['Facturas'].sum()
        d_general = df_t1['Visitas/Clientes'].sum()
        
        total_proj_pct = time_pct + ecommerce_pct + al_mayor_pct + no_visibles_pct if has_projections else time_pct
        c_proyeccion = round_half_up(c_general * total_proj_pct)
        c_total_proyectado = c_general + c_proyeccion
        
        d_proyeccion = round_half_up(d_general * time_pct)
        d_total_proyectado = d_general + d_proyeccion
        
        projected_effectiveness = (c_total_proyectado / d_total_proyectado * 100.0) if d_total_proyectado > 0 else 0.0
        observed_effectiveness = (c_general / d_general * 100.0) if d_general > 0 else 0.0
        
        # Own brand metrics
        total_sales = df_filtered['VENTA_TOTAL'].sum()
        own_brand_sales = df_filtered['VENTA_MARCA_PROPIA'].sum()
        pct_sales = (own_brand_sales / total_sales * 100.0) if total_sales > 0 else 0.0
        
        # Promo vs Non-promo sales (observed)
        promo_sales = df_filtered['VENTA_PROMO'].sum()
        non_promo_sales = df_filtered['VENTA_FUERA_PROMO'].sum()
        pct_promo = (promo_sales / total_sales * 100.0) if total_sales > 0 else 0.0
        pct_non_promo = (non_promo_sales / total_sales * 100.0) if total_sales > 0 else 0.0
        
        all_categories = [
            'A/A', 'TV', 'LAVADO', 'CONGELADOR', 'ELECTRODOMESTICOS', 
            'COCINA', 'SONIDO', 'NEVERA', 'COMPUTACION', 'TELEFONIA', 'OTROS'
        ]
        cat_sales_dict = df_filtered.groupby('CATEGORIA')['VENTA_TOTAL'].sum().to_dict()
        df_t4_data = []
        for cat in all_categories:
            df_t4_data.append({
                'Categoría': cat,
                'Venta Total': float(cat_sales_dict.get(cat, 0.0))
            })
        df_t4 = pd.DataFrame(df_t4_data)
        df_t4 = df_t4.sort_values(by='Venta Total', ascending=False).reset_index(drop=True)
        
        top_cats = []
        total_cat_sales = df_t4['Venta Total'].sum()
        # Filter active categories (Venta Total > 0)
        df_t4_active = df_t4[df_t4['Venta Total'] > 0]
        if df_t4_active.empty:
            df_t4_active = df_t4
        for idx, row in df_t4_active.head(5).iterrows():
            cat_name = row['Categoría']
            cat_val = row['Venta Total']
            pct = (cat_val / total_cat_sales * 100.0) if total_cat_sales > 0 else 0.0
            top_cats.append(f"{cat_name} con {int(round_half_up(pct))}%")
            
        if len(top_cats) > 2:
            categorias_str = f"{top_cats[0]}, seguida de {', '.join(top_cats[1:-1])} y {top_cats[-1]}"
        elif len(top_cats) == 2:
            categorias_str = f"{top_cats[0]}, seguida de {top_cats[1]}"
        elif len(top_cats) == 1:
            categorias_str = top_cats[0]
        else:
            categorias_str = "ninguna"
            
        # Top 3 Products by units sold
        df_t2 = df_filtered.groupby(['PRODUCTO_CORRECTO', 'PRECIO_MAESTRO'])['Cantidad'].sum().reset_index()
        df_t2.columns = ['Producto', 'Precio', 'Cantidad']
        df_top_products = df_t2.sort_values(by='Cantidad', ascending=False).head(3)
        
        prod_items = []
        for idx, row in df_top_products.iterrows():
            prod_items.append(f"{row['Producto']} (${int(round_half_up(row['Precio']))})")
        if len(prod_items) > 2:
            productos_str = f"{prod_items[0]}, {prod_items[1]} y {prod_items[2]}"
        elif len(prod_items) == 2:
            productos_str = f"{prod_items[0]} y {prod_items[1]}"
        elif len(prod_items) == 1:
            productos_str = prod_items[0]
        else:
            productos_str = "ninguno"
            
        # Large sale discount
        large_invoices_sum = 0.0
        df_inv = df_filtered.dropna(subset=['Factura'])
        for (hor, inv), group in df_inv.groupby(['Horario', 'Factura']):
            inv_sum = group['VENTA_TOTAL'].sum()
            if inv_sum >= 1000.0:
                large_invoices_sum += float(inv_sum)
                
        if large_invoices_sum > 0:
            descuento_str = f"Se observó un descuento en proyección por venta grande de ${int(round_half_up(large_invoices_sum))}"
        else:
            descuento_str = "No se observó descuento en proyección por venta grande"
            
        # Variation compared to previous audit
        curr_sales, curr_visits = get_projected_totals(df_filtered, opening_hour, limit_hour, empresa)
        # NOTA (fix): antes se pasaba `empresa_input` (ej. "Damasco"), que nunca coincide
        # con el valor crudo de la columna 'Empresa' del Sheet (ej. "ddamasco"), por lo que
        # esta búsqueda nunca encontraba el estudio anterior. Usamos `target_emp`, que es el
        # mismo identificador ya traducido que se usa para filtrar df_filtered arriba.
        df_prev_filtered = find_previous_study_df(df_campo, fecha, target_emp, sucursal, col_sucursal, limit_hour)
        
        if df_prev_filtered is not None and not df_prev_filtered.empty:
            # df_prev_filtered viene de df_campo "crudo" (sin las columnas de matching, ya
            # que ahora el matching solo corre sobre df_filtered por rendimiento). Como
            # get_projected_totals necesita 'VENTA_TOTAL', calculamos el matching aquí,
            # pero solo sobre estas ~decenas/cientos de filas del estudio anterior, no
            # sobre todo el histórico — se sigue evitando el costo original.
            df_prev_filtered = df_prev_filtered.copy()
            df_prev_filtered['PRODUCTO_CORRECTO'] = df_prev_filtered.apply(lambda r: buscar_coincidencia_tecnica(r, universo_maestro), axis=1)
            df_prev_filtered['PRECIO_MAESTRO'] = df_prev_filtered['PRODUCTO_CORRECTO'].map(mapeo_precios).fillna(0.0).apply(round_half_up)
            df_prev_filtered['VENTA_TOTAL'] = df_prev_filtered['PRECIO_MAESTRO'] * df_prev_filtered['Cantidad']
            prev_sales, prev_visits = get_projected_totals(df_prev_filtered, opening_hour, limit_hour, empresa)
            sales_var = ((curr_sales - prev_sales) / prev_sales * 100.0) if prev_sales > 0 else 0.0
            visits_var = ((curr_visits - prev_visits) / prev_visits * 100.0) if prev_visits > 0 else 0.0
            
            sales_sign = "+" if sales_var >= 0 else ""
            visits_sign = "+" if visits_var >= 0 else ""
            sales_var_rounded = int(round_half_up(sales_var))
            visits_var_rounded = int(round_half_up(visits_var))
            
            variacion_str = f"ventas proyectadas {sales_sign}{sales_var_rounded}% ({sales_var:+.1f}%), visitas proyectadas {visits_sign}{visits_var_rounded}% ({visits_var:+.1f}%)"
        else:
            variacion_str = "No se encontró estudio previo en los últimos dos meses"
            
        # Format opening and closing hours beautifully in 12h
        apertura_12h = format_hour_12h(opening_hour)
        cierre_12h = format_hour_12h(limit_hour)
        
        obs_lines = []
        obs_lines.append(f"-Horario de apertura y cierre: {apertura_12h} - {cierre_12h}")
        obs_lines.append(f"-Efectividad: {int(round_half_up(projected_effectiveness))}%")
        obs_lines.append(f"-Marca propia representa el {int(round_half_up(pct_sales))}% de las ventas observadas en el estudio")
        obs_lines.append(f"-Productos en promocion representan el {int(round_half_up(pct_promo))}% de las ventas, mientras que los productos fuera de promo representan el {int(round_half_up(pct_non_promo))}% de las ventas observadas.")
            
        obs_lines.append(f"-Categorias con mayor representación: {categorias_str}")
        obs_lines.append(f"-Productos más vendidos: {productos_str}")
        obs_lines.append(f"-{descuento_str}")
        obs_lines.append(f"-{variacion_str}")
        
        # Average visits / Traffic criteria
        avg_visits_val = int(round_half_up(df_t1['Visitas/Clientes'].mean()))
        if avg_visits_val >= 80:
            trafico_criterio = "tráfico alto"
        elif avg_visits_val >= 50:
            trafico_criterio = "tráfico medio"
        else:
            trafico_criterio = "tráfico bajo"
            
        # Hour with maximum visits
        max_visits = int(df_t1['Visitas/Clientes'].max())
        max_rows = df_t1[df_t1['Visitas/Clientes'] == max_visits]
        if not max_rows.empty:
            max_horario_raw = max_rows.iloc[0]['Horario']
            max_horario_formatted = formatear_horario_lindo(max_horario_raw)
        else:
            max_horario_formatted = "N/A"
            
        obs_lines.append(f"-El tráfico en tienda fue {trafico_criterio}, registrando un promedio de {avg_visits_val} visitas por hora.")
        obs_lines.append(f"-La hora con el mayor número de clientes fue de {max_horario_formatted}, registrando {max_visits} visitas.")
        
        obs_text = "\n".join(obs_lines) + "\n"
        obs_text = remover_tildes(obs_text).upper()
        
        # Write to memory stream
        output = io.BytesIO()
        output.write(obs_text.encode('utf-8'))
        
        filename = f"Observaciones_Auditoria_{empresa.upper()}_{sucursal.upper()}_{fecha.replace('/', '-')}.txt"
        
        # Force garbage collection to free RAM on Render
        import gc
        gc.collect()
        
        output.seek(0)
        return send_file(
            output,
            mimetype="text/plain",
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        import gc
        gc.collect()
        # Fix: mismo criterio que en /api/generate — traceback al log, mensaje limpio al usuario.
        app.logger.error("Exception in generate_observations", exc_info=True)
        return jsonify({'success': False, 'error': f"No se pudieron generar las observaciones: {str(e)}"}), 400

@app.route('/api/sync', methods=['POST'])
def sync_database():
    if not is_logged_in():
        return jsonify({'success': False, 'error': 'No autorizado.'}), 401
    try:
        global _cache_campo_data, _cache_campo_time, _cache_maestro_data, _cache_maestro_time
        
        base_dir = os.path.dirname(os.path.abspath(__file__))
        now = time.time()
        
        # FIX CONCURRENCIA: mismo lock que usa fetch_data(), para que un /api/sync no
        # pueda pisar la caché a mitad de una lectura que esté haciendo otro request.
        with _cache_lock:
            # 1. Fetch df_campo exactly once
            cache_campo_path = os.path.join(base_dir, "df_campo_cache.csv")
            req1 = urllib.request.Request(
                URL_TIPIFICACIONES, 
                headers={'User-Agent': 'Mozilla/5.0'}
            )
            with urllib.request.urlopen(req1, timeout=10) as response:
                df_campo = pd.read_csv(io.StringIO(response.read().decode('utf-8')))
            df_campo.to_csv(cache_campo_path, index=False)
            _cache_campo_data = df_campo
            _cache_campo_time = now
            
            # 2. Fetch each master database exactly once
            for gid in [GID_BASE_DAKA, GID_BASE_DAMASCO, GID_BASE_MULTIMAX]:
                cache_maestro_path = os.path.join(base_dir, f"df_maestro_{gid}_cache.csv")
                url_base = f"https://docs.google.com/spreadsheets/d/e/2PACX-1vTfq81DhLQ_8jkbFIAs7OWaO7qkYRis350TTRz_BbbsVucVw4K87Ai0YgiynRIQG1CqRJv9i1V6oEDo/pub?gid={gid}&single=true&output=csv"
                req2 = urllib.request.Request(
                    url_base, 
                    headers={'User-Agent': 'Mozilla/5.0'}
                )
                with urllib.request.urlopen(req2, timeout=10) as response:
                    df_maestro = pd.read_csv(io.StringIO(response.read().decode('utf-8')))
                df_maestro.to_csv(cache_maestro_path, index=False)
                
                # Clean df_maestro for the in-memory cache
                df_maestro_clean = df_maestro.copy()
                if 'PRECIO DE REFERENCIA' in df_maestro_clean.columns:
                    prices = df_maestro_clean['PRECIO DE REFERENCIA'].astype(str).str.replace(' ', '').str.replace(',', '.')
                    prices_numeric = pd.to_numeric(prices, errors='coerce').fillna(0.0)
                    df_maestro_clean = df_maestro_clean[prices_numeric > 0.0].copy()
                    
                _cache_maestro_data[gid] = df_maestro_clean
                _cache_maestro_time[gid] = now
            
        return jsonify({'success': True, 'message': 'Base de datos sincronizada con éxito.'})
    except Exception as e:
        # Fix: mismo criterio — sin traceback en la respuesta al cliente.
        app.logger.error("Exception in sync_database", exc_info=True)
        return jsonify({'success': False, 'error': f"Error de conexión con Google Sheets: {str(e)}"}), 400

if __name__ == '__main__':
    # FIX SEGURIDAD: debug=True fijo exponía el debugger interactivo de Werkzeug
    # (ejecución arbitraria de código en el navegador) si este archivo se corre
    # directamente en un servidor accesible desde afuera. En producción real esto no
    # aplicaba porque el Procfile usa gunicorn (que no pasa por este bloque), pero
    # quedaba activo por defecto para cualquiera que corriera `python server.py` en un
    # entorno mal configurado. Ahora el modo debug requiere activarse a propósito.
    debug_mode = os.environ.get('FLASK_DEBUG', '0') == '1'
    app.run(host='0.0.0.0', port=5000, debug=debug_mode)
