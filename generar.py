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

EXCEL_PATH = 'C:\\Users\\andre\\Documents\\GitHub\\TxtSimulacion\\Simulaction-TXT\\Encuesta de trabajo de grado_ Tesis trafico Perestrelo-Balbuena.xlsx'
SHEET_NAME = 'Form Responses 1'
OUTPUT_TXT = 'simulacion_resumen.txt'

# Distancia típica (km) para cálculo de tiempo de viaje si no hay campo de distancia
DISTANCIA_DEFAULT_KM = 10.0

# Modelos (puedes ajustar probabilidades o valores por defecto)
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

# Hora de salida, si no hay columna de hora: media por defecto (en horas flotante, ej 7.5 = 7:30)
HORA_SALIDA_MEDIA_DEFAULT = 7.5
HORA_SALIDA_SIGMA_H = 0.5

HORA_SALIDA_UNIMET_MEDIA = 15.75  # 15:45 -> 15.75 horas
HORA_SALIDA_UNIMET_SIGMA_H = 0.5


np.random.seed(42)  # reproducibilidad por defecto

def hourfloat_to_str(hf):
    hh = int(math.floor(hf))
    mm = int(round((hf - hh) * 60))
    if mm == 60:
        hh += 1
        mm = 0
    return f"{hh:02d}:{mm:02d}"


def sample_normal_hour(mean_h, sigma_h):
    return float(np.random.normal(loc=mean_h, scale=sigma_h))


# --- Cargar datos ---
print('Leyendo Excel:', EXCEL_PATH)
if not os.path.exists(EXCEL_PATH):
    raise FileNotFoundError(f"No se encontró el archivo Excel en {EXCEL_PATH}.\nColoca el archivo en esa ruta o cambia EXCEL_PATH en el script.")

try:
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
except Exception as e:
    raise RuntimeError(f"Error leyendo la hoja '{SHEET_NAME}' del Excel: {e}")

df.columns = [str(c).strip() for c in df.columns]
N = len(df)
print(f'Filas leídas: {N}')

col_estilo = next((c for c in df.columns if 'estilo' in c.lower()), None)
col_condicion = next((c for c in df.columns if 'condic' in c.lower()), None)
col_clima = next((c for c in df.columns if 'clima' in c.lower()), None)
col_tipo = next((c for c in df.columns if any(t in c.lower() for t in ['tipo', 'veh', 'carro', 'moto', 'taxi'])), None)
col_ubicacion = next((c for c in df.columns if 'ubic' in c.lower()), None)
col_zona = next((c for c in df.columns if 'zona' in c.lower()), None)
col_hora_salida = next((c for c in df.columns if 'salida' in c.lower() and 'casa' in c.lower()), None)
col_hora_llegada = next((c for c in df.columns if 'llegad' in c.lower()), None)
col_personas = next((c for c in df.columns if 'person' in c.lower() or 'acompa' in c.lower()), None)

if col_estilo is None:
    print('No se encontró columna "estilo". Se asignará aleatorio con probabilidades por defecto.')
    estilos = np.random.choice(list(ESTILO_MAP.keys()), size=N, p=[0.2, 0.6, 0.2])
    df['Estilo de manejo'] = estilos
    col_estilo = 'Estilo de manejo'
else:
    df[col_estilo] = df[col_estilo].fillna('Prudente')

if col_condicion is None:
    print('No se encontró columna "condición". Asignando "Buena" por defecto con algo de variabilidad.')
    df['Condicion'] = np.random.choice(['Buena', 'Regular', 'Mala'], size=N, p=[0.7, 0.2, 0.1])
    col_condicion = 'Condicion'
else:
    df[col_condicion] = df[col_condicion].fillna('Buena')

if col_clima is None:
    print('No se encontró columna "clima". Se simulará clima por persona (Soleado/Nublado/Lluvioso).')
    df['Clima'] = np.random.choice(['Soleado', 'Nublado', 'Lluvioso'], size=N, p=[0.6, 0.25, 0.15])
    col_clima = 'Clima'
else:
    df[col_clima] = df[col_clima].fillna('Soleado')

if col_tipo is None:
    print('No se encontró columna de tipo de vehículo. Se asignará aleatoriamente.')
    df['Tipo de Carro'] = np.random.choice(list(TIPO_CAPACIDAD.keys()), size=N, p=[0.6, 0.2, 0.05, 0.15])
    col_tipo = 'Tipo de Carro'
else:
    df[col_tipo] = df[col_tipo].fillna('carro')

if col_ubicacion is None:
    df['Ubicación'] = df.get(col_ubicacion, 'Desconocido')
    col_ubicacion = 'Ubicación'

if col_zona is None:
    df['Zona'] = df.get(col_zona, 'Desconocido')
    col_zona = 'Zona'

# Personas por fila: si existe columna la usamos, si no asumimos 1 persona por fila (después agrupamos en vehículos)
if col_personas is None:
    df['NumPersonasFila'] = 1
    col_personas = 'NumPersonasFila'
else:
    df[col_personas] = df[col_personas].fillna(1)
    df['NumPersonasFila'] = df[col_personas]

# Hora de salida de casa: si existe, intentar parsear; si no, generar normal
if col_hora_salida is None:
    df['HoraSalidaCasa_h'] = np.random.normal(loc=HORA_SALIDA_MEDIA_DEFAULT, scale=HORA_SALIDA_SIGMA_H, size=N)
else:
    # intentar extraer hora si la columna es de tipo fecha/hora o texto
    try:
        parsed = pd.to_datetime(df[col_hora_salida], errors='coerce')
        mask = parsed.notna()
        df['HoraSalidaCasa_h'] = np.nan
        df.loc[mask, 'HoraSalidaCasa_h'] = parsed[mask].dt.hour + parsed[mask].dt.minute / 60.0
        # para nulos, muestrear
        n_missing = df['HoraSalidaCasa_h'].isna().sum()
        if n_missing > 0:
            df.loc[df['HoraSalidaCasa_h'].isna(), 'HoraSalidaCasa_h'] = np.random.normal(loc=HORA_SALIDA_MEDIA_DEFAULT, scale=HORA_SALIDA_SIGMA_H, size=n_missing)
    except Exception:
        df['HoraSalidaCasa_h'] = np.random.normal(loc=HORA_SALIDA_MEDIA_DEFAULT, scale=HORA_SALIDA_SIGMA_H, size=N)

# Hora de salida de la UNIMET (si existe, use; si no, muestree around media)
col_unimet = next((c for c in df.columns if 'unimet' in c.lower()), None)
if col_unimet is None:
    df['HoraSalidaUnimet_h'] = np.random.normal(loc=HORA_SALIDA_UNIMET_MEDIA, scale=HORA_SALIDA_UNIMET_SIGMA_H, size=N)
else:
    try:
        parsed = pd.to_datetime(df[col_unimet], errors='coerce')
        mask = parsed.notna()
        df['HoraSalidaUnimet_h'] = np.nan
        df.loc[mask, 'HoraSalidaUnimet_h'] = parsed[mask].dt.hour + parsed[mask].dt.minute / 60.0
        n_missing = df['HoraSalidaUnimet_h'].isna().sum()
        if n_missing > 0:
            df.loc[df['HoraSalidaUnimet_h'].isna(), 'HoraSalidaUnimet_h'] = np.random.normal(loc=HORA_SALIDA_UNIMET_MEDIA, scale=HORA_SALIDA_UNIMET_SIGMA_H, size=n_missing)
    except Exception:
        df['HoraSalidaUnimet_h'] = np.random.normal(loc=HORA_SALIDA_UNIMET_MEDIA, scale=HORA_SALIDA_UNIMET_SIGMA_H, size=N)

results = []

for idx, row in df.iterrows():
    estilo = str(row[col_estilo]) if col_estilo in row else 'Prudente'
    condicion = str(row[col_condicion]) if col_condicion in row else 'Buena'
    clima = str(row[col_clima]) if col_clima in row else 'Soleado'
    tipo = str(row[col_tipo]).lower() if col_tipo in row else 'carro'
    ubicacion = str(row[col_ubicacion]) if col_ubicacion in row else 'Desconocido'
    zona = str(row[col_zona]) if col_zona in row else 'Desconocido'

    # Base de velocidad según estilo
    estilo_key = next((k for k in ESTILO_MAP.keys() if k.lower() in estilo.lower()), 'Prudente')
    base = ESTILO_MAP[estilo_key]['base']
    sigma = ESTILO_MAP[estilo_key]['sigma']

    # velocidad promedio muestreada por persona
    vel_prom_sample = float(np.random.normal(loc=base, scale=sigma))
    vel_prom_sample = max(10.0, vel_prom_sample)  # límite inferior razonable

    # aplicar efecto clima
    clima_key = next((k for k in CLIMA_EFFECT.keys() if k.lower() in clima.lower()), 'Soleado')
    vel_prom = vel_prom_sample * CLIMA_EFFECT[clima_key]

    # velocidad inicial (por ejemplo cuando sale de casa)
    vel_inicial = float(np.random.uniform(10.0, min(vel_prom, 40.0)))

    # velocidad máxima estimada según estilo y condición del carro
    cond_key = next((k for k in CONDICION_EFFECT.keys() if k.lower() in condicion.lower()), 'Buena')
    vel_max = vel_prom_sample * (1 + np.random.uniform(0.15, 0.6)) * CONDICION_EFFECT[cond_key]
    vel_max = max(vel_prom, vel_max)

    # ajuste por tipo de vehículo (motos suelen tener mayor aceleración pero menor promedio en ciudad)
    if 'moto' in tipo:
        vel_prom *= 1.05
        vel_max *= 1.05
    elif 'wawa' in tipo:
        vel_prom *= 0.9
        vel_max *= 0.9

    # tiempo de trayecto (horas) = distancia / velocidad_prom (velocidad en km/h)
    distancia_km = DISTANCIA_DEFAULT_KM
    tiempo_h = distancia_km / max(vel_prom, 5.0)

    # horas como floats
    h_salida = float(row.get('HoraSalidaCasa_h', np.random.normal(loc=HORA_SALIDA_MEDIA_DEFAULT, scale=HORA_SALIDA_SIGMA_H)))
    h_llegada = h_salida + tiempo_h

    # demora estimada: si llegada > 8.0 (ejemplo) se considera 'voy tarde' (esto es configurable)
    demora_min = max(0, (h_llegada - (8.0)) * 60)  # minutos por encima de las 8:00
    voy_tarde = 'Sí' if h_llegada > 8.0 else 'No'

    # cambio de velocidad (porcentaje estimado)
    cambio_vel_pct = float(np.abs(np.random.normal(loc=5.0, scale=3.0)))  # ejemplo

    # desvío y reacción a situación externa (simple categorización aleatoria con dependencia del estilo)
    if estilo_key == 'Agresivo':
        desvio = np.random.choice(['Alto', 'Medio', 'Bajo'], p=[0.25,0.5,0.25])
        reaccion = np.random.choice(['Rápida','Moderada','Lenta'], p=[0.6,0.3,0.1])
    elif estilo_key == 'Prudente':
        desvio = np.random.choice(['Alto', 'Medio', 'Bajo'], p=[0.05,0.25,0.7])
        reaccion = np.random.choice(['Rápida','Moderada','Lenta'], p=[0.2,0.6,0.2])
    else:  # Lento
        desvio = np.random.choice(['Alto', 'Medio', 'Bajo'], p=[0.02,0.18,0.8])
        reaccion = np.random.choice(['Rápida','Moderada','Lenta'], p=[0.1,0.5,0.4])

    # acompaÃ±ado (si la fila indica n personas, consideramos si >1)
    num_personas_fila = int(row.get('NumPersonasFila', 1))
    acompañado = 'Sí' if num_personas_fila > 1 else 'No'

    # probabilidad / asignación de personas a vehículos: aquí sólo informamos la probabilidad (se puede agrupar después)
    prob_asignacion = None

    # tráfico estimado por encuestado (si existe columna la usamos, si no, simulamos)
    col_trafico = next((c for c in df.columns if 'tráfico' in c.lower() or 'traf' in c.lower()), None)
    if col_trafico is None:
        trafico_est = np.random.choice(['Bajo','Medio','Alto'], p=[0.4,0.45,0.15])
    else:
        trafico_est = str(row[col_trafico]) if pd.notna(row[col_trafico]) else np.random.choice(['Bajo','Medio','Alto'])

    # Tipo transmisión (sincrónico/automático) aleatorio si no existe
    col_transm = next((c for c in df.columns if 'trans' in c.lower()), None)
    if col_transm is None:
        transm = np.random.choice(['Automático','Sincrónico'], p=[0.6,0.4])
    else:
        transm = str(row[col_transm]) if pd.notna(row[col_transm]) else np.random.choice(['Automático','Sincrónico'])

    # buenas/malas condiciones ya representadas por condicion

    # formato final por persona (diccionario)
    info = {
        'Año': datetime.now().year,
        'Número de personas': num_personas_fila,
        'Velocidad': round(vel_prom,1),
        'Velocidad inicial': round(vel_inicial,1),
        'Velocidad máximo': round(vel_max,1),
        'Velocidad promedio': round(vel_prom,1),
        'Ubicación': ubicacion,
        'Zona donde vive': zona,
        'Entrada preferida': row.get('Entrada preferida', 'Desconocida') if 'Entrada preferida' in df.columns else row.get('Entrada', 'Desconocida'),
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
        'probabilidad': prob_asignacion,
        'Hora de salir de la casa': hourfloat_to_str(h_salida),
        'Hora de llegar a la Unimet': hourfloat_to_str(h_llegada),
        'Hora de salir de la Unimet': hourfloat_to_str(float(row.get('HoraSalidaUnimet_h', np.random.normal(loc=HORA_SALIDA_UNIMET_MEDIA, scale=HORA_SALIDA_UNIMET_SIGMA_H)))),
        'Colegio Integral el Ávila': row.get('Colegio Integral el Ávila', 'No') if 'Colegio Integral el Ávila' in df.columns else 'No',
        'tráfico estimado por los encuestados': trafico_est,
    }

    results.append(info)

# --- Escribir TXT de salida ---
print(f'Generando archivo de salida: {OUTPUT_TXT} (registros: {len(results)})')
with open(OUTPUT_TXT, 'w', encoding='utf-8') as f:
    for r in results:
        # una línea por persona, campos separados por ' | '
        parts = [f"{k}: {v}" for k, v in r.items()]
        line = ' | '.join(parts)
        f.write(line + '\n')

print('Hecho. Archivo generado:', OUTPUT_TXT)
print('Puedes ajustar parámetros en la sección "Parámetros de simulación" del script.')
