import pandas as pd
import numpy as np
from datetime import datetime
import math
import os
import random
from collections import defaultdict
import unicodedata

EXCEL_PATH = './Encuesta de trabajo de grado_ Tesis trafico Perestrelo-Balbuena.xlsx'
SHEET_NAME = 'Form Responses 1'
OUTPUT_TXT = 'simulacion_resumen.txt'

# -------------------
# Configuraciones
# -------------------

DISTANCIA_DEFAULT_KM = 10.0

ESTILO_MAP = {
    'Agresivo': {'base': 60, 'sigma': 15},
    'Prudente': {'base': 45, 'sigma': 10},
    'Lento': {'base': 35, 'sigma': 8},
}

CONDICION_EFFECT = {
    'Buena': 1.0,
    'Regular': 0.90,
    'Mala': 0.80,
}

CLIMA_EFFECT = {
    'Soleado': 1.0,
    'Nublado': 0.95,
    'Lluvioso': 0.85,
}

TIPO_CAPACIDAD = {
    'carro': (1, 4),
    'moto': (1, 2),
    'wawa': (1, 32),
    'taxi': (1, 4),
}

LIMITES_VIA = {
    "autopista": 70,
    "troncal": 60,
    "urbana": 40,
    "interseccion": 15
}

HORA_SALIDA_MEDIA_DEFAULT = 7.5
HORA_SALIDA_SIGMA_H = 0.5

np.random.seed(42)

# -------------------
# Funciones auxiliares
# -------------------

def hourfloat_to_str(hf):
    try:
        hh = int(math.floor(hf))
        mm = int(round((hf - hh) * 60))
        if mm == 60:
            hh += 1
            mm = 0
        return f"{hh:02d}:{mm:02d}"
    except:
        return "00:00"

def normalize_text(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = s.replace("¿", "").replace("?", "")
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s.lower().strip()

def clasificar_tipo_via(zona):
    if pd.isna(zona):
        return "urbana"
    zona = zona.lower()
    autopista_keywords = ["la guaira", "guaira", "petare", "guarenas", "cota mil", "francisco fajardo", "gran cacique"]
    troncal_keywords = ["los teques", "san antonio", "carretera vieja"]
    urbana_keywords = ["chacao", "altamira", "sebucán", "los chorros", "urbina", "macaracuay", "terrazas", "castellana", "florida"]
    if any(k in zona for k in autopista_keywords):
        return "autopista"
    if any(k in zona for k in troncal_keywords):
        return "troncal"
    return "urbana"

def aplicar_limite_legal(velocidad, tipo_via, interseccion=False):
    if interseccion:
        return min(velocidad, LIMITES_VIA["interseccion"])
    return min(velocidad, LIMITES_VIA.get(tipo_via, 40))

# -------------------
# Leer Excel
# -------------------

if not os.path.exists(EXCEL_PATH):
    raise FileNotFoundError(f"No se encontró el archivo Excel en {EXCEL_PATH}.")

df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
df.columns = [str(c).strip() for c in df.columns]
N = len(df)
print(f'Filas leídas: {N}')

# -------------------
# Detectar columnas importantes
# -------------------

col_anio_entrada = "¿En qué año entró a la universidad?"
col_estilo = next((c for c in df.columns if 'estilo' in normalize_text(c)), None)
col_condicion = next((c for c in df.columns if 'condic' in normalize_text(c)), None)
col_clima = next((c for c in df.columns if 'clima' in normalize_text(c)), None)
col_tipo = "¿Cuál es su principal medio de transporte a la UNIMET?"
col_zona = '¿De qué parte de Caracas vienes?'
COL_ENTRADA = '¿Cúal entrada sueles tomar para entrar en la UNIMET?'
COL_COLEGIO = '¿Ha encontrado cola del Colegio Integral el Ávila en el horario seleccionado previamente?'
COL_TRAFICO = '¿En que horas consideras que hay tráfico en los accesos a la UNIMET?'

# -------------------
# Completar columnas faltantes
# -------------------

if col_estilo is None:
    df['Estilo de manejo'] = np.random.choice(list(ESTILO_MAP.keys()), size=N, p=[0.2,0.6,0.2])
    col_estilo = 'Estilo de manejo'
else:
    df[col_estilo] = df[col_estilo].fillna('Prudente')

if col_condicion is None:
    df['Condicion'] = np.random.choice(list(CONDICION_EFFECT.keys()), size=N, p=[0.7,0.2,0.1])
    col_condicion = 'Condicion'
else:
    df[col_condicion] = df[col_condicion].fillna('Buena')

if col_clima is None:
    df['Clima'] = np.random.choice(list(CLIMA_EFFECT.keys()), size=N, p=[0.6,0.25,0.15])
    col_clima = 'Clima'
else:
    df[col_clima] = df[col_clima].fillna('Soleado')

# -------------------
# Normalizar columna de tipo de transporte
# -------------------

def normalizar_tipo(row):
    tipo_raw = str(row[col_tipo]).strip().lower()
    
    # Caso especial: "carro propio / moto propia"
    if "carro propio / moto propia" in tipo_raw:
        return pd.Series({"Tipo de Carro": "carro propio", "carro/moto/wawa/taxi": "carro"})
    
    # Caso especial: "a pie"
    elif "a pie" in tipo_raw:
        return pd.Series({"Tipo de Carro": "wawa", "carro/moto/wawa/taxi": "wawa"})
    
    # Caso general: normalizar según palabras clave
    elif "carro" in tipo_raw:
        return pd.Series({"Tipo de Carro": row[col_tipo], "carro/moto/wawa/taxi": "carro"})
    elif "moto" in tipo_raw:
        return pd.Series({"Tipo de Carro": row[col_tipo], "carro/moto/wawa/taxi": "moto"})
    elif "wawa" in tipo_raw:
        return pd.Series({"Tipo de Carro": row[col_tipo], "carro/moto/wawa/taxi": "wawa"})
    elif "taxi" in tipo_raw or "ridery" in tipo_raw or "yummy" in tipo_raw:
        return pd.Series({"Tipo de Carro": row[col_tipo], "carro/moto/wawa/taxi": "taxi"})
    else:
        return pd.Series({"Tipo de Carro": row[col_tipo], "carro/moto/wawa/taxi": "carro"})

df[['Tipo de Carro', 'carro/moto/wawa/taxi']] = df.apply(normalizar_tipo, axis=1)

# -------------------
# Procesamiento fila por fila
# -------------------

results = []

for _, row in df.iterrows():

    # Año
    try:
        anio_raw = int(float(row.get(col_anio_entrada, 2020)))
    except:
        anio_raw = 2020
    if anio_raw <= 2020:
        anio = 2020
    elif anio_raw in [2021,2022,2023,2024,2025]:
        anio = anio_raw
    else:
        anio = 2025

    # Estilo, clima, condición, tipo
    estilo = str(row.get(col_estilo, 'Prudente'))
    estilo_key = next((k for k in ESTILO_MAP.keys() if k.lower() in estilo.lower()), 'Prudente')
    base = ESTILO_MAP[estilo_key]['base']
    sigma = ESTILO_MAP[estilo_key]['sigma']

    condicion = str(row.get(col_condicion, 'Buena'))
    cond_key = next((k for k in CONDICION_EFFECT.keys() if k.lower() in condicion.lower()), 'Buena')

    clima = str(row.get(col_clima, 'Soleado'))
    clima_key = next((k for k in CLIMA_EFFECT.keys() if k.lower() in clima.lower()), 'Soleado')

    tipo_simple = row['carro/moto/wawa/taxi']

    # Velocidad simulada
    vel_base = np.random.normal(base, sigma)
    vel_base = max(10.0, vel_base)
    vel_prom = vel_base * CLIMA_EFFECT[clima_key]
    vel_inicial = vel_prom * 0.9
    vel_max = vel_prom * 1.6 * CONDICION_EFFECT[cond_key]
    vel_max = max(vel_prom, vel_max)

    # Tipo de vía
    zona_text = str(row.get(col_zona, "Desconocido"))
    tipo_via = clasificar_tipo_via(zona_text)

    vel_via = aplicar_limite_legal(vel_prom, tipo_via)
    vel_interseccion = aplicar_limite_legal(vel_prom, tipo_via, interseccion=True)
    vel_legal = vel_via * 0.95 + vel_interseccion * 0.05

    # Ajuste por tipo de vehículo
    if tipo_simple == "moto":
        vel_legal *= 1.05
    elif tipo_simple == "wawa":
        vel_legal *= 0.9

    # Tiempo de viaje
    distancia_km = DISTANCIA_DEFAULT_KM
    tiempo_h = distancia_km / max(vel_legal, 5.0)

    # Hora salida / llegada
    h_salida = float(np.random.normal(HORA_SALIDA_MEDIA_DEFAULT, HORA_SALIDA_SIGMA_H))
    h_llegada = h_salida + tiempo_h
    demora_min = max(0, (h_llegada - 8.0) * 60)
    voy_tarde = 'Sí' if h_llegada > 8.0 else 'No'

    # Comportamiento
    reaccion_pct = np.random.normal(50, 15)
    reaccion_pct = max(1, min(reaccion_pct, 100))
    desvio_pct = np.random.normal(18, 7)
    desvio_pct = max(1, min(desvio_pct, 100))
    cambio_vel_pct = round(np.random.normal(5.5, 2), 1)

    # Personas y acompañantes
    num_personas_fila = 1
    if "te traen en carro" in str(row[col_tipo]).lower():
        num_personas_fila = random.randint(2, TIPO_CAPACIDAD['carro'][1])
    elif tipo_simple in ["taxi", "ridery", "yummy"]:
        num_personas_fila = 2
    elif tipo_simple == "wawa":
        num_personas_fila = random.randint(1, TIPO_CAPACIDAD['wawa'][1])
    acompañado = 'Sí' if num_personas_fila > 1 else 'No'

    # Información final
    info = {
        'Año': anio,
        'Número de personas': num_personas_fila,
        'Velocidad': round(vel_legal,1),
        'Velocidad inicial': round(vel_inicial,1),
        'Velocidad máximo': round(vel_max,1),
        'Velocidad promedio': round(vel_legal,1),
        'Ubicación': zona_text,
        'Tipo de vía': tipo_via,
        'Zona donde vive': zona_text,
        'Entrada preferida': row.get(COL_ENTRADA, 'Desconocida'),
        'Estilo de manejo': estilo_key,
        'reacción a situación externa (%)': round(reaccion_pct,1),
        'desvío (%)': round(desvio_pct,1),
        'cambio de velocidad (%)': cambio_vel_pct,
        'Demora estimada (min)': round(demora_min,1),
        'Tipo de Carro': row['Tipo de Carro'],
        'carro/moto/wawa/taxi': tipo_simple,
        'Sincrónico/automático': row.get('Transmisión', 'Automático'),
        'buenas/malas condiciones': cond_key,
        'Retraso': voy_tarde,
        'Viene acompañado': acompañado,
        'Hora de salir de la casa': hourfloat_to_str(h_salida),
        'Hora de llegar a la Unimet': hourfloat_to_str(h_llegada),
        'Colegio Integral el Ávila': row.get(COL_COLEGIO, 'No'),
        'tráfico estimado por los encuestados': row.get(COL_TRAFICO, 'No responde'),
    }

    results.append(info)

# -------------------
# Ajuste por año
# -------------------

TARGET_COUNTS = {2020:96,2021:35,2022:30,2023:66,2024:126,2025:101}
year_groups = defaultdict(list)
for r in results:
    year_groups[r['Año']].append(r)

adjusted_results = []
for year, target_count in TARGET_COUNTS.items():
    group = year_groups.get(year, [])
    if len(group) == 0:
        base = random.choice(results)
        new_group = [base.copy() for _ in range(target_count)]
        for g in new_group:
            g['Año'] = year
        adjusted_results.extend(new_group)
    elif len(group) == target_count:
        adjusted_results.extend(group)
    elif len(group) > target_count:
        adjusted_results.extend(random.sample(group, target_count))
    else:
        needed = target_count - len(group)
        replicated = random.choices(group, k=needed)
        adjusted_results.extend(group + replicated)

results = adjusted_results

# -------------------
# Escribir TXT
# -------------------

with open(OUTPUT_TXT, 'w', encoding='utf-8') as f:
    for r in results:
        line = ' | '.join([f"{k}: {v}" for k,v in r.items()])
        f.write(line + '\n')

print(f"✔ Archivo generado correctamente: {OUTPUT_TXT}")
