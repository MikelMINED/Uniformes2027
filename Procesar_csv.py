# ============================================================
# Procesar_csv.py
# ------------------------------------------------------------
# Pipeline de generación de datos para el dashboard
# "Logística de Uniformes 2027 | MINEDUCYT".
#
# Qué hace:
#   1. Pide al usuario un CSV de matrícula (corte actual).
#   2. Detecta automáticamente el corte anterior por nombre.
#   3. Calcula nuevos ingresos y traslados por NIE.
#   4. Cruza con "Tallado Sí.xlsx" para marcar tallado (Sí/No).
#   5. Actualiza 3 JSON en 03_dashboard/:
#        - resumen_diario.json
#        - detalle_cambios.json
#        - detalle_nuevos.json
#   6. Hace git add/commit/push del repo.
#
# Dependencias:
#   pip install pandas openpyxl
#
# Estructura esperada:
#   RAIZ/
#     01_originales/
#        matricula_al_DDMMYYYY_normalizado.csv   (cortes)
#        Tallado Sí.xlsx                          (nie, tallado)
#     03_dashboard/                               (salida JSON)
#     Procesar_csv.py                             (este script)
# ============================================================

import os
import re
import json
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from pathlib import Path
import pandas as pd
import subprocess

# ============================================
# CONFIGURACIÓN DE ENTORNO
# ============================================
RAIZ = Path(__file__).parent.resolve()
ORIGINALES = RAIZ / "01_originales"
DASHBOARD = RAIZ / "03_dashboard"
TALLADO_XLSX = ORIGINALES / "Tallado Sí.xlsx"

COLUMNAS = [
    'est_nie', 'est_nombre1', 'est_nombre2', 'est_nombre3',
    'est_apellido1', 'est_apellido2', 'est_apellido3',
    'GRADO_RECOD_BTO', 'CÓDIGO_CE', 'NOMBRE_CE', 'SECTOR',
    'DEPTO_CE', 'NOMBRE_MUNICIPIO_CE', 'NOMBRE_DISTRITO_CE', 'DIRECCIÓN_CE'
]

def extraer_fecha(nombre_archivo):
    match = re.search(r'_al_(\d{8})_normalizado', nombre_archivo)
    if match:
        fecha_str = match.group(1)
        return datetime.strptime(fecha_str, "%d%m%Y")
    return None

def cargar_archivos():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    ruta_actual = filedialog.askopenfilename(
        initialdir=ORIGINALES,
        title="Selecciona el archivo CSV del corte actual",
        filetypes=[("Archivos CSV", "*.csv")]
    )

    if not ruta_actual:
        print("Operación cancelada.")
        return None, None, None

    archivo_actual = os.path.basename(ruta_actual)
    fecha_actual_obj = extraer_fecha(archivo_actual)

    if not fecha_actual_obj:
        print("El archivo seleccionado no cumple el patrón de nomenclatura esperado.")
        return None, None, None

    fecha_str_iso = fecha_actual_obj.strftime("%Y-%m-%d")

    archivos_validos = []
    for f in os.listdir(ORIGINALES):
        if f.endswith(".csv"):
            f_fecha = extraer_fecha(f)
            if f_fecha:
                archivos_validos.append((f_fecha, f))

    archivos_validos.sort(key=lambda x: x[0])

    ruta_anterior = None
    for i, (fecha, nombre) in enumerate(archivos_validos):
        if nombre == archivo_actual:
            if i > 0:
                ruta_anterior = ORIGINALES / archivos_validos[i - 1][1]
            break

    print(f"Archivo actual: {archivo_actual}")
    if ruta_anterior:
        print(f"Archivo anterior: {ruta_anterior.name}")
    else:
        print("Archivo anterior: No existe (primer registro).")

    return Path(ruta_actual), ruta_anterior, fecha_str_iso

# ============================================
# TALLADO
# ============================================
def cargar_tallados():
    """Devuelve un set con los NIE que tienen 'si' en la columna tallado."""
    if not TALLADO_XLSX.exists():
        print(f"⚠ No se encontró '{TALLADO_XLSX.name}'. Se asumirá 'No' para todos.")
        return set()

    try:
        df = pd.read_excel(TALLADO_XLSX, dtype=str)
    except Exception as e:
        print(f"⚠ Error al leer '{TALLADO_XLSX.name}': {e}. Se asumirá 'No' para todos.")
        return set()

    df.columns = [str(c).strip().lower() for c in df.columns]

    if 'nie' not in df.columns or 'tallado' not in df.columns:
        print(f"⚠ El Excel de tallado debe tener columnas 'nie' y 'tallado'. Encontradas: {list(df.columns)}")
        return set()

    mask = df['tallado'].fillna('').astype(str).str.strip().str.lower().isin(['si', 'sí', 's'])
    tallados = set(df.loc[mask, 'nie'].fillna('').astype(str).str.strip())
    tallados.discard('')
    print(f"Tallados cargados desde Excel: {len(tallados)}")
    return tallados

def marcar_tallado(df, tallados):
    """Añade columna 'tallado' con 'Sí' / 'No' a un DataFrame."""
    if df is None:
        return None
    df = df.copy()
    df['tallado'] = df['est_nie'].fillna('').astype(str).str.strip().isin(tallados).map({True: 'Sí', False: 'No'})
    return df

def calcular_totales_tallado(df_actual, tallados):
    """Devuelve (n_tallados, n_no_tallados) sobre la matrícula actual."""
    serie = df_actual['est_nie'].fillna('').astype(str).str.strip()
    n_tallados = int(serie.isin(tallados).sum())
    total = len(df_actual)
    return n_tallados, total - n_tallados

# ============================================
# PROCESAMIENTO
# ============================================
def procesar_datos(ruta_actual, ruta_anterior):
    print("\nCargando archivo actual...")
    df_actual = pd.read_csv(ruta_actual, usecols=COLUMNAS, dtype=str, encoding='utf-8', low_memory=False)

    cols_nombres = ['est_nombre1', 'est_nombre2', 'est_nombre3', 'est_apellido1', 'est_apellido2', 'est_apellido3']
    df_actual['nombre_completo'] = df_actual[cols_nombres].fillna('').agg(' '.join, axis=1).str.replace(r'\s+', ' ', regex=True).str.strip()
    df_actual = df_actual.drop(columns=cols_nombres)

    total_actual = len(df_actual)
    total_publico = len(df_actual[df_actual['SECTOR'].str.upper().str.contains('PÚBLICO|OFICIAL', na=False)])
    total_privado = len(df_actual[df_actual['SECTOR'].str.upper().str.contains('PRIVADO', na=False)])

    print(f"Total matrícula: {total_actual}")
    print(f"Público: {total_publico} | Privado: {total_privado}")

    if not ruta_anterior:
        print("\nFase de Inicialización: No hay archivo anterior para comparar.")
        return total_actual, total_publico, total_privado, None, None, df_actual

    print("\nCargando archivo anterior para comparación...")
    cols_anterior = ['est_nie', 'CÓDIGO_CE', 'NOMBRE_CE', 'SECTOR', 'DEPTO_CE', 'NOMBRE_MUNICIPIO_CE', 'NOMBRE_DISTRITO_CE', 'DIRECCIÓN_CE']
    df_anterior = pd.read_csv(ruta_anterior, usecols=cols_anterior, dtype=str, encoding='utf-8', low_memory=False)

    df_nuevos = df_actual[~df_actual['est_nie'].isin(df_anterior['est_nie'])].copy()
    print(f"Nuevos ingresos detectados: {len(df_nuevos)}")

    df_comun = df_actual.merge(df_anterior, on='est_nie', suffixes=('_destino', '_origen'))
    df_cambios = df_comun[df_comun['CÓDIGO_CE_destino'] != df_comun['CÓDIGO_CE_origen']].copy()
    print(f"Cambios de centro detectados: {len(df_cambios)}")

    return total_actual, total_publico, total_privado, df_cambios, df_nuevos, df_actual

# ============================================
# EXPORTACIÓN
# ============================================
def exportar_json(fecha_str, total_actual, total_publico, total_privado,
                  df_cambios, df_nuevos, total_tallados, total_no_tallados):
    DASHBOARD.mkdir(parents=True, exist_ok=True)

    ruta_resumen = DASHBOARD / "resumen_diario.json"
    if ruta_resumen.exists():
        with open(ruta_resumen, 'r', encoding='utf-8') as f:
            resumen = json.load(f)
    else:
        resumen = []

    resumen = [r for r in resumen if r['fecha'] != fecha_str]

    nuevo_resumen = {
        "fecha": fecha_str,
        "total": int(total_actual),
        "publico": int(total_publico),
        "privado": int(total_privado),
        "cambios": int(len(df_cambios)) if df_cambios is not None else 0,
        "nuevos": int(len(df_nuevos)) if df_nuevos is not None else 0,
        "tallados": int(total_tallados),
        "no_tallados": int(total_no_tallados),
    }
    resumen.append(nuevo_resumen)
    resumen.sort(key=lambda x: x['fecha'], reverse=True)

    with open(ruta_resumen, 'w', encoding='utf-8') as f:
        json.dump(resumen, f, indent=2, ensure_ascii=False)

    if df_cambios is not None:
        ruta_cambios = DASHBOARD / "detalle_cambios.json"
        hist_cambios = {}
        if ruta_cambios.exists():
            with open(ruta_cambios, 'r', encoding='utf-8') as f:
                hist_cambios = json.load(f)

        hist_cambios[fecha_str] = df_cambios.fillna('').to_dict(orient='records')
        with open(ruta_cambios, 'w', encoding='utf-8') as f:
            json.dump(hist_cambios, f, indent=2, ensure_ascii=False)

    if df_nuevos is not None:
        ruta_nuevos = DASHBOARD / "detalle_nuevos.json"
        hist_nuevos = {}
        if ruta_nuevos.exists():
            with open(ruta_nuevos, 'r', encoding='utf-8') as f:
                hist_nuevos = json.load(f)

        hist_nuevos[fecha_str] = df_nuevos.fillna('').to_dict(orient='records')
        with open(ruta_nuevos, 'w', encoding='utf-8') as f:
            json.dump(hist_nuevos, f, indent=2, ensure_ascii=False)

    print(f"\nArchivos JSON actualizados exitosamente en: {DASHBOARD}")

def sincronizar_con_github():
    print("\nSincronizando con GitHub (vía Git local)...")
    try:
        repo_path = str(RAIZ)
        subprocess.run(["git", "-C", repo_path, "add", "03_dashboard/", "index.html", "Procesar_csv.py"], check=True)
        mensaje_commit = f"Actualización automática de JSON {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        subprocess.run(["git", "-C", repo_path, "commit", "-m", mensaje_commit], check=True)
        subprocess.run(["git", "-C", repo_path, "push"], check=True)
        print("Datos subidos exitosamente a GitHub")
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar comando Git: {e}")
    except Exception as e:
        print(f"Error inesperado al sincronizar con GitHub: {e}")

# ============================================
# MAIN
# ============================================
if __name__ == "__main__":
    actual, anterior, fecha_str = cargar_archivos()
    if actual:
        t_act, t_pub, t_pri, df_cam, df_nue, df_actual = procesar_datos(actual, anterior)

        tallados = cargar_tallados()
        df_cam = marcar_tallado(df_cam, tallados)
        df_nue = marcar_tallado(df_nue, tallados)
        n_tall, n_no_tall = calcular_totales_tallado(df_actual, tallados)

        print(f"Tallados en matrícula actual: {n_tall} | No tallados: {n_no_tall}")

        exportar_json(fecha_str, t_act, t_pub, t_pri, df_cam, df_nue, n_tall, n_no_tall)
        sincronizar_con_github()