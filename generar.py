"""
Script: generar_txt_simulacion.py
Lee la hoja "Form Responses 1" del Excel subido, genera variables derivadas y escribe un único archivo .txt resumen
por persona para usarse en simulaciones.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import math
import os
import random
from collections import defaultdict
import unicodedata

EXCEL_PATH = 'C:\\Users\\andre\\Documents\\GitHub\\TxtSimulacion\\Simulaction-TXT\\Encuesta de trabajo de grado_ Tesis trafico Perestrelo-Balbuena.xlsx'
SHEET_NAME = 'Form Responses 1'
OUTPUT_TXT = 'simulacion_resumen.txt'

# Distancia típica (km)
DISTANCIA_DEFAULT_KM = 10.0

# Modelos
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
    'wawa': (1, 20),
    'taxi': (1, 4),
}

HORA_SALIDA_MEDIA_DEFAULT = 7.5
HORA_SALIDA_SIGMA_H = 0.5
HORA_SALIDA_UNIMET_MEDIA = 15.75
HORA_SALIDA_UNIMET_SIGMA_H = 0.5

np.random.seed(42)

def hourfloat_to_str(hf):
    # hf puede ser NaN o valor extraño; protegerse
    try:
        hh = int(math.floor(hf))
        mm = int(round((hf - hh) * 60))
        if mm == 60:
            hh += 1
            mm = 0
        return f"{hh:02d}:{mm:02d}"
    except Exception:
        return "00:00"

def sample_normal_hour(mean_h, sigma_h):
    return float(np.random.normal(loc=mean_h, scale=sigma_h))

def normalize_text(s: str) -> str:
    """Quita tildes, signos de interrogación y pasa a minúsculas."""
    if not isinstance(s, str):
        return ""
    s = s.replace("¿", "").replace("?", "")
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')  # remove accents
    return s.lower().strip()

# --- Cargar datos ---
print('Leyendo Excel:', EXCEL_PATH)
if not os.path.exists(EXCEL_PATH):
    raise FileNotFoundError(f"No se encontró el archivo Excel en {EXCEL_PATH}.")

try:
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
except Exception as e:
    raise RuntimeError(f"Error leyendo la hoja '{SHEET_NAME}': {e}")

# conservar nombres originales pero limpiarlos visualmente
df.columns = [str(c).strip() for c in df.columns]
N = len(df)
print(f'Filas leídas: {N}')

# -------------------------------------------------------
# DETECCIÓN ROBUSTA DE LA COLUMNA DEL AÑO DE ENTRADA
# -------------------------------------------------------
col_anio_entrada = None
for c in df.columns:
    n = normalize_text(c)
    # buscamos palabras núcleo: "ano" (año), "entro"/"entr" (entró/entró), "univers" (universidad)
    if "ano" in n and ("entro" in n or "entr" in n) and "univers" in n:
        col_anio_entrada = c
        break

if col_anio_entrada is None:
    print("\nColumnas detectadas en el Excel:")
    for c in df.columns:
        print(f" - {c!r}")
    raise RuntimeError("No se encontró la columna del año de entrada a la universidad (se buscó 'año' + 'entr/entro' + 'univers').")

print(f"Columna detectada correctamente: '{col_anio_entrada}'")

# -------------------------------------------------------
# Detectar otras columnas relevantes (flexible)
# -------------------------------------------------------
col_estilo = next((c for c in df.columns if 'estilo' in normalize_text(str(c))), None)
col_condicion = next((c for c in df.columns if 'condic' in normalize_text(str(c))), None)
col_clima = next((c for c in df.columns if 'clima' in normalize_text(str(c))), None)
col_tipo = next((c for c in df.columns if any(t in normalize_text(str(c)) for t in ['tipo','veh','carro','moto','taxi'])), None)
col_ubicacion = next((c for c in df.columns if 'ubic' in normalize_text(str(c))), None)
col_zona = next((c for c in df.columns if 'zona' in normalize_text(str(c))), None)
col_hora_salida = next((c for c in df.columns if 'salida' in normalize_text(str(c)) and 'casa' in normalize_text(str(c))), None)
col_personas = next((c for c in df.columns if 'person' in normalize_text(str(c)) or 'acompa' in normalize_text(str(c))), None)

# Manejo de columnas faltantes y limpieza
if col_estilo is None:
    df['Estilo de manejo'] = np.random.choice(list(ESTILO_MAP.keys()), size=N, p=[0.2,0.6,0.2])
    col_estilo = 'Estilo de manejo'
else:
    df[col_estilo] = df[col_estilo].fillna('Prudente')

if col_condicion is None:
    df['Condicion'] = np.random.choice(['Buena','Regular','Mala'], size=N, p=[0.7,0.2,0.1])
    col_condicion = 'Condicion'
else:
    df[col_condicion] = df[col_condicion].fillna('Buena')

if col_clima is None:
    df['Clima'] = np.random.choice(['Soleado','Nublado','Lluvioso'], size=N, p=[0.6,0.25,0.15])
    col_clima = 'Clima'
else:
    df[col_clima] = df[col_clima].fillna('Soleado')

if col_tipo is None:
    df['Tipo de Carro'] = np.random.choice(list(TIPO_CAPACIDAD.keys()), size=N, p=[0.6,0.2,0.05,0.15])
    col_tipo = 'Tipo de Carro'
else:
    df[col_tipo] = df[col_tipo].fillna('carro')

if col_personas is None:
    df['NumPersonasFila'] = 1
    col_personas = 'NumPersonasFila'
else:
    # Asegurar numericidad
    try:
        df[col_personas] = pd.to_numeric(df[col_personas], errors='coerce')
    except:
        pass
    df['NumPersonasFila'] = df[col_personas].fillna(1)

# -------------------------------------------------------
# Horas: manejo correcto de nulos (NO usar fillna(array))
# -------------------------------------------------------

# Hora de salida de casa
if col_hora_salida is None:
    df['HoraSalidaCasa_h'] = np.random.normal(loc=HORA_SALIDA_MEDIA_DEFAULT, scale=HORA_SALIDA_SIGMA_H, size=N)
else:
    parsed = pd.to_datetime(df[col_hora_salida], errors='coerce')
    df['HoraSalidaCasa_h'] = parsed.dt.hour + parsed.dt.minute / 60.0
    mask_missing = df['HoraSalidaCasa_h'].isna()
    if mask_missing.any():
        df.loc[mask_missing, 'HoraSalidaCasa_h'] = np.random.normal(
            loc=HORA_SALIDA_MEDIA_DEFAULT,
            scale=HORA_SALIDA_SIGMA_H,
            size=mask_missing.sum()
        )

# Hora de salida de la Unimet (acceso/unimet)
col_unimet = next((c for c in df.columns if 'unimet' in normalize_text(str(c))), None)
if col_unimet is None:
    df['HoraSalidaUnimet_h'] = np.random.normal(loc=HORA_SALIDA_UNIMET_MEDIA, scale=HORA_SALIDA_UNIMET_SIGMA_H, size=N)
else:
    parsed = pd.to_datetime(df[col_unimet], errors='coerce')
    df['HoraSalidaUnimet_h'] = parsed.dt.hour + parsed.dt.minute / 60.0
    mask_missing_unimet = df['HoraSalidaUnimet_h'].isna()
    if mask_missing_unimet.any():
        df.loc[mask_missing_unimet, 'HoraSalidaUnimet_h'] = np.random.normal(
            loc=HORA_SALIDA_UNIMET_MEDIA,
            scale=HORA_SALIDA_UNIMET_SIGMA_H,
            size=mask_missing_unimet.sum()
        )

# -------------------------------------------------------
# Construir resultados por fila
# -------------------------------------------------------

results = []

for idx, row in df.iterrows():
    # Año basado en la columna detectada
    try:
        # intentar convertir a entero de forma segura
        anio_raw = int(float(row.get(col_anio_entrada, 2020)))
    except Exception:
        anio_raw = 2020

    if anio_raw <= 2020:
        anio = 2020
    elif anio_raw in [2021, 2022, 2023, 2024, 2025]:
        anio = anio_raw
    else:
        anio = 2025

    # Lectura segura de campos
    estilo = str(row.get(col_estilo, 'Prudente'))
    condicion = str(row.get(col_condicion, 'Buena'))
    clima = str(row.get(col_clima, 'Soleado'))
    tipo = str(row.get(col_tipo, 'carro')).lower()
    ubicacion = str(row.get(col_ubicacion, 'Desconocido')) if col_ubicacion else 'Desconocido'
    zona = str(row.get(col_zona, 'Desconocido')) if col_zona else 'Desconocido'

    estilo_key = next((k for k in ESTILO_MAP.keys() if k.lower() in estilo.lower()), 'Prudente')
    base = ESTILO_MAP[estilo_key]['base']
    sigma = ESTILO_MAP[estilo_key]['sigma']

    vel_prom_sample = float(np.random.normal(loc=base, scale=sigma))
    vel_prom_sample = max(10.0, vel_prom_sample)
    clima_key = next((k for k in CLIMA_EFFECT.keys() if k.lower() in clima.lower()), 'Soleado')
    vel_prom = vel_prom_sample * CLIMA_EFFECT[clima_key]

    vel_inicial = float(np.random.uniform(10.0, min(vel_prom, 40.0)))
    cond_key = next((k for k in CONDICION_EFFECT.keys() if k.lower() in condicion.lower()), 'Buena')
    vel_max = vel_prom_sample * (1 + np.random.uniform(0.15,0.6)) * CONDICION_EFFECT[cond_key]
    vel_max = max(vel_prom, vel_max)

    # ajustes por tipo
    if 'moto' in tipo:
        vel_prom *= 1.05
        vel_max *= 1.05
    elif 'wawa' in tipo:
        vel_prom *= 0.9
        vel_max *= 0.9

    # tiempo de trayecto
    distancia_km = DISTANCIA_DEFAULT_KM
    tiempo_h = distancia_km / max(vel_prom, 5.0)

    # horas como floats (asegurar que existan)
    try:
        h_salida = float(row.get('HoraSalidaCasa_h', np.random.normal(loc=HORA_SALIDA_MEDIA_DEFAULT, scale=HORA_SALIDA_SIGMA_H)))
    except Exception:
        h_salida = float(np.random.normal(loc=HORA_SALIDA_MEDIA_DEFAULT, scale=HORA_SALIDA_SIGMA_H))

    h_llegada = h_salida + tiempo_h

    demora_min = max(0, (h_llegada - 8.0) * 60)  # minutos por encima de las 8:00
    voy_tarde = 'Sí' if h_llegada > 8.0 else 'No'

    cambio_vel_pct = float(abs(np.random.normal(loc=5.0, scale=3.0)))

    # desvío y reacción
    if estilo_key == 'Agresivo':
        desvio = np.random.choice(['Alto', 'Medio', 'Bajo'], p=[0.25,0.5,0.25])
        reaccion = np.random.choice(['Rápida','Moderada','Lenta'], p=[0.6,0.3,0.1])
    elif estilo_key == 'Prudente':
        desvio = np.random.choice(['Alto', 'Medio', 'Bajo'], p=[0.05,0.25,0.7])
        reaccion = np.random.choice(['Rápida','Moderada','Lenta'], p=[0.2,0.6,0.2])
    else:
        desvio = np.random.choice(['Alto', 'Medio', 'Bajo'], p=[0.02,0.18,0.8])
        reaccion = np.random.choice(['Rápida','Moderada','Lenta'], p=[0.1,0.5,0.4])

    # personas
    try:
        num_personas_fila = int(row.get('NumPersonasFila', 1))
    except Exception:
        num_personas_fila = 1
    acompañado = 'Sí' if num_personas_fila > 1 else 'No'

    # tráfico estimado
    col_trafico = next((c for c in df.columns if 'traf' in normalize_text(str(c))), None)
    trafico_est = str(row[col_trafico]) if col_trafico and pd.notna(row.get(col_trafico)) else np.random.choice(['Bajo','Medio','Alto'], p=[0.4,0.45,0.15])

    # transmisión
    col_transm = next((c for c in df.columns if 'trans' in normalize_text(str(c))), None)
    transm = str(row[col_transm]) if col_transm and pd.notna(row.get(col_transm)) else np.random.choice(['Automático','Sincrónico'], p=[0.6,0.4])

    # Hora salida Unimet (proteger)
    try:
        h_salida_unimet = float(row.get('HoraSalidaUnimet_h', np.random.normal(loc=HORA_SALIDA_UNIMET_MEDIA, scale=HORA_SALIDA_UNIMET_SIGMA_H)))
    except Exception:
        h_salida_unimet = float(np.random.normal(loc=HORA_SALIDA_UNIMET_MEDIA, scale=HORA_SALIDA_UNIMET_SIGMA_H))

    info = {
        'Año': anio,
        'Número de personas': num_personas_fila,
        'Velocidad': round(vel_prom,1),
        'Velocidad inicial': round(vel_inicial,1),
        'Velocidad máximo': round(vel_max,1),
        'Velocidad promedio': round(vel_prom,1),
        'Ubicación': ubicacion,
        'Zona donde vive': zona,
        'Entrada preferida': row.get('Entrada preferida', 'Desconocida'),
        'Comportamiento': row.get('Comportamiento',''),
        'Estilo de manejo': estilo_key,
        'reacción a situación externa': reaccion,
        'desvío': desvio,
        'cambio de velocidad (%)': round(cambio_vel_pct,1),
        'Demora estimada (min)': round(demora_min,1),
        'Tipo de Carro': tipo,
        'carro/moto/wawa/taxi': tipo,
        'Sincrónico/automático': transm,
        'buenas/malas condiciones': cond_key,
        'Retraso': voy_tarde,
        'Viene acompañado': acompañado,
        'probabilidad': None,
        'Hora de salir de la casa': hourfloat_to_str(h_salida),
        'Hora de llegar a la Unimet': hourfloat_to_str(h_llegada),
        'Hora de salir de la Unimet': hourfloat_to_str(h_salida_unimet),
        'Colegio Integral el Ávila': row.get('Colegio Integral el Ávila','No'),
        'tráfico estimado por los encuestados': trafico_est,
    }

    results.append(info)

# -------------------------------------------------------
# AJUSTAR CANTIDADES POR AÑO (OPCIÓN B)
# -------------------------------------------------------

TARGET_COUNTS = {
    2020: 96,
    2021: 35,
    2022: 30,
    2023: 66,
    2024: 126,
    2025: 101
}

year_groups = defaultdict(list)
for r in results:
    year_groups[r['Año']].append(r)

adjusted_results = []

for year, target_count in TARGET_COUNTS.items():
    group = year_groups.get(year, [])

    if len(group) == 0:
        print(f"⚠️ No hay registros del año {year}. Se generarán por copia.")
        base = random.choice(results)
        new_group = [base.copy() for _ in range(target_count)]
        for g in new_group:
            g['Año'] = year
        adjusted_results.extend(new_group)
        continue

    if len(group) == target_count:
        adjusted_results.extend(group)
    elif len(group) > target_count:
        adjusted_results.extend(random.sample(group, target_count))
    else:
        needed = target_count - len(group)
        replicated = random.choices(group, k=needed)
        adjusted_results.extend(group + replicated)

results = adjusted_results

print("\n✔ Cantidades finales por año:")
for y in TARGET_COUNTS:
    print(f"{y}: {sum(1 for r in results if r['Año']==y)} registros")

# -------------------------------------------------------
# ESCRIBIR TXT FINAL
# -------------------------------------------------------
print(f'\nGenerando archivo: {OUTPUT_TXT}')
with open(OUTPUT_TXT, 'w', encoding='utf-8') as f:
    for r in results:
        parts = [f"{k}: {v}" for k,v in r.items()]
        line = ' | '.join(parts)
        f.write(line + '\n')

print("\n✔ Hecho. Archivo generado correctamente.\n")
