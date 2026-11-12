
"""
CODEBRAN - Proyecto Final de Programación II
Bodega y cuantificación de materiales de obra
================================================================

Universidad Mariano Gálvez de Guatemala
Facultad de Ingeniería, Ingeniería Civil, Sección "B"
Antigua Guatemala

Integrantes:
  - Obed Barrientos
  - Mayelin Gisel García
  - Dervin Velásquez
  - Manfred Urízar Ruano

Programa de consola con la misma lógica que la página web:
  0. Cuentas: al entrar, cada persona inicia sesión o crea su cuenta con su
     correo electrónico y contraseña (sin administrador). La contraseña se
     guarda cifrada (PBKDF2-SHA256 con sal) en "cuentas_programa.json".
  1. Inventario de bodega (agregar, retirar, corregir, quitar, kardex),
     con unidades de medida según el tipo de material y conversión entre ellas.
     El inventario es independiente: los cálculos no lo leen ni lo descuentan.
  2. Cálculo de materiales para elementos de concreto reforzado
     (losa, viga, columna, zapata, muro, cimiento corrido, solera, mocheta),
     con el grado de hierro según la medida de la varilla.
  3. Muros: block (mampostería reforzada, AGIES NSE 7.4, NTG 41054) y
     muros de contención de concreto reforzado, concreto ciclópeo, gaviones
     y mampostería de piedra, con revisión de estabilidad (AGIES NSE 2.1).
  4. Carreteras y drenaje: pavimento rígido, bordillo fundido,
     cuneta tipo L, cuneta tipo V y canal (con capacidad hidráulica).
  5. Tablas de normativa guatemalteca: acero, cemento, block, mortero
     y tipos de muro.
  6. Memoria de cálculo en Excel (.xlsx) con fórmulas, sección típica e
     isométrico, en la carpeta "memorias" (sin instalar librerías).
  7. Historial de cálculos guardados y exportación a CSV.

Los datos se guardan en "datos_proyecto.json", junto al programa.
Las cantidades son valores de referencia para un estimado; el diseño
final debe verificarlo un ingeniero estructural.

Ejecutar:  python proyecto_final.py
"""

import csv
import getpass
import hashlib
import json
import math
import os
import re
import secrets
import struct
import unicodedata
import zipfile
import zlib
from datetime import datetime

UNIVERSIDAD = "Universidad Mariano Gálvez de Guatemala"
CARRERA = "Facultad de Ingeniería, Ingeniería Civil, Sección «B», Antigua Guatemala"
INTEGRANTES = ["Obed Barrientos", "Mayelin Gisel García", "Dervin Velásquez", "Manfred Urízar Ruano"]

# ============================================================
# PARTE 1: DATOS DE REFERENCIA
# ============================================================
DESPERDICIO = 0.05      # 5 % de desperdicio
VOLUMEN_SECO = 1.54     # factor de volumen seco de la mezcla de concreto
KG_SACO = 42.5          # peso de un saco de cemento (kg)
VOLUMEN_SACO = 0.0283   # volumen de un saco de cemento (m³)
LARGO_VARILLA = 6       # largo comercial de una varilla (m)
KG_M3_CAL = 500         # densidad aparente aproximada de la cal hidratada (kg/m³)

PESO_VARILLA = {"#2": 0.249, "#3": 0.560, "#4": 0.994, "#5": 1.552,
                "#6": 2.235, "#7": 3.042, "#8": 3.973}   # kg/m
EQUIVALENCIA_VARILLA = {"#2": "1/4", "#3": "3/8", "#4": "1/2", "#5": "5/8",
                        "#6": "3/4", "#7": "7/8", "#8": "1"}

# Acero de refuerzo, COGUANOR NTG 36011 (ASTM A615): diámetro, área,
# grado usual según la medida y uso típico.
VARILLAS = {
    "#2": {"mm": 6.35,  "area": 0.32, "grado": 40, "uso": "Estribos, eslabones, amarres (se vende lisa)"},
    "#3": {"mm": 9.53,  "area": 0.71, "grado": 40, "uso": "Losas, soleras, mochetas, cimiento corrido, muros de block"},
    "#4": {"mm": 12.70, "area": 1.27, "grado": 40, "uso": "Vigas, columnas, zapatas, barras de amarre de pavimento"},
    "#5": {"mm": 15.88, "area": 1.98, "grado": 60, "uso": "Vigas y columnas de mayor carga"},
    "#6": {"mm": 19.05, "area": 2.85, "grado": 60, "uso": "Columnas y vigas de edificios; pasajuntas de losas de 15 cm"},
    "#7": {"mm": 22.23, "area": 3.88, "grado": 60, "uso": "Edificios y puentes; pasajuntas de losas de 18 cm"},
    "#8": {"mm": 25.40, "area": 5.07, "grado": 60, "uso": "Edificios y puentes; pasajuntas de losas de 20 a 25 cm"},
}
GRADOS = {40: {"fy": 2810, "psi": "40,000", "mpa": 280}, 60: {"fy": 4200, "psi": "60,000", "mpa": 420}}

RESISTENCIAS = {
    "1": {"nombre": "2000 PSI (140 kg/cm²)", "cemento": 1, "arena": 3,   "piedrin": 5,   "agua_cemento": 0.70,
          "uso": "Cimiento ciclópeo, banquetas, fundiciones sin carga"},
    "2": {"nombre": "2500 PSI (175 kg/cm²)", "cemento": 1, "arena": 2,   "piedrin": 4,   "agua_cemento": 0.62,
          "uso": "Bordillos, cunetas, soleras y mochetas de vivienda"},
    "3": {"nombre": "3000 PSI (210 kg/cm²)", "cemento": 1, "arena": 2,   "piedrin": 3,   "agua_cemento": 0.55,
          "uso": "Losas, vigas, columnas, zapatas, canales (uso estructural común)"},
    "4": {"nombre": "4000 PSI (280 kg/cm²)", "cemento": 1, "arena": 1.5, "piedrin": 2.5, "agua_cemento": 0.46,
          "uso": "Pavimento rígido (MR ≈ 4.5 MPa), elementos muy cargados"},
}

# Morteros de pegado (ASTM C270 / AGIES NSE 7.4) y concreto fluido para celdas.
MORTEROS = {
    "M": {"nombre": "Mortero tipo M", "cemento": 1, "cal": 0.25, "arena": 3,   "piedrin": 0, "agua_cemento": 0.75,
          "factor": 1.25, "kgcm2": 175, "uso": "Muros de carga, cimientos y muros bajo el nivel del suelo"},
    "S": {"nombre": "Mortero tipo S", "cemento": 1, "cal": 0.5,  "arena": 4.5, "piedrin": 0, "agua_cemento": 0.85,
          "factor": 1.25, "kgcm2": 125, "uso": "Mampostería reforzada en general (el más usado)"},
    "N": {"nombre": "Mortero tipo N", "cemento": 1, "cal": 1,    "arena": 6,   "piedrin": 0, "agua_cemento": 1.00,
          "factor": 1.25, "kgcm2": 50,  "uso": "Muros interiores sin carga y tabiques"},
}
GROUT = {"nombre": "Concreto fluido 1 : 2¼ : 2 (piedrín 3/8)", "cemento": 1, "arena": 2.25, "piedrin": 2, "agua_cemento": 0.70}

# Block de concreto, COGUANOR NTG 41054 (resistencia sobre área neta a 28 días).
CLASES_BLOCK = {
    "A": {"promedio": 133.3, "individual": 113.0, "absorcion": 10, "color": "Azul",  "uso": "Muros de carga, mampostería reforzada"},
    "B": {"promedio": 100.0, "individual": 85.0,  "absorcion": 15, "color": "Rojo",  "uso": "Vivienda de un nivel con carga moderada"},
    "C": {"promedio": 66.0,  "individual": 56.0,  "absorcion": 20, "color": "Verde", "uso": "Muros divisorios y cercos sin carga"},
}
CELDA = {9: 0.0087, 14: 0.0142, 19: 0.0205}   # área interior de una celda del block (m²)

TIPOS_MURO = [
    ("Mampostería reforzada de block", "AGIES NSE 7.4",
     "Block clase A con refuerzo vertical en celdas rellenas (máx. a cada 1.20 m) y soleras de humedad, intermedia y corona."),
    ("Mampostería confinada", "AGIES NSE 4 y NSE 7.4",
     "Muro de block o ladrillo con columnas (mochetas) y soleras de concreto que lo confinan. Común en vivienda de 1 y 2 niveles."),
    ("Ladrillo de barro cocido", "NSE 7.4",
     "Ladrillo tayuyo o perforado, confinado o reforzado. Muy usado en Antigua y el occidente del país."),
    ("Muro de concreto reforzado", "AGIES NSE 7.1 (ACI 318)",
     "Muro fundido con doble cama de acero: muros de corte, sótanos, tanques y cisternas."),
    ("Muros de contención", "AGIES NSE 2.1 (geotecnia)",
     "De gravedad (piedra o concreto ciclópeo), en voladizo (concreto reforzado) o de gaviones. Necesitan estudio de suelos y drenaje."),
    ("Muros no estructurales", "NTG 41055",
     "Tabiques de block de 9 cm, tablayeso o paneles. Solo dividen ambientes; deben anclarse a la estructura."),
    ("Adobe y bajareque", "No recomendado en obra nueva",
     "Sistemas de tierra muy vulnerables en sismos; en obra nueva se sustituyen por mampostería reforzada o confinada."),
]

ELEMENTOS = {
    "losa":    {"nombre": "Losa",    "hierro_kg_m3": 80,  "varilla": "#3", "piedrin_tamano": "3/4 pulg",
                "medidas": ("Largo", "Ancho", "Espesor"),
                "recomendacion": "Varilla #3 (3/8 pulg) en ambos sentidos, a cada 15 o 20 cm."},
    "viga":    {"nombre": "Viga",    "hierro_kg_m3": 120, "varilla": "#4", "piedrin_tamano": "3/4 pulg",
                "medidas": ("Largo", "Base", "Peralte"),
                "recomendacion": "Varillas longitudinales #4 o #5; estribos #2 o #3 a cada 10-20 cm."},
    "columna": {"nombre": "Columna", "hierro_kg_m3": 150, "varilla": "#4", "piedrin_tamano": "3/4 pulg",
                "medidas": ("Lado a", "Lado b", "Altura"),
                "recomendacion": "Varillas longitudinales #4 (o #5 si la carga es alta); estribos #2 o #3 a cada 10-15 cm."},
    "zapata":  {"nombre": "Zapata",  "hierro_kg_m3": 70,  "varilla": "#4", "piedrin_tamano": "1 pulg",
                "medidas": ("Largo", "Ancho", "Peralte"),
                "recomendacion": "Varilla #4 en ambos sentidos, a cada 15-25 cm."},
    "muro":    {"nombre": "Muro",    "hierro_kg_m3": 60,  "varilla": "#3", "piedrin_tamano": "3/4 a 1 pulg",
                "medidas": ("Largo", "Espesor", "Altura"),
                "recomendacion": "Varilla #3 vertical y horizontal, a cada 20-30 cm."},
    "cimiento": {"nombre": "Cimiento corrido", "hierro_kg_m3": 45, "varilla": "#3", "piedrin_tamano": "3/4 a 1 pulg",
                 "medidas": ("Largo", "Ancho", "Peralte"),
                 "recomendacion": "3 varillas #3 corridas con eslabones #2 a cada 0.20 m (vivienda de un nivel, NSE 4)."},
    "solera":  {"nombre": "Solera",  "hierro_kg_m3": 110, "varilla": "#3", "piedrin_tamano": "1/2 pulg",
                "medidas": ("Largo", "Ancho", "Alto"),
                "recomendacion": "4 varillas #3 con estribos #2 a cada 0.15-0.20 m. Soleras de humedad, intermedia y corona."},
    "mocheta": {"nombre": "Mocheta", "hierro_kg_m3": 100, "varilla": "#3", "piedrin_tamano": "1/2 pulg",
                "medidas": ("Lado a", "Lado b", "Altura"),
                "recomendacion": "2 a 4 varillas #3 con eslabones o estribos #2 a cada 0.15-0.20 m; en esquinas y a los lados de vanos."},
}

ARCHIVO_DATOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "datos_proyecto.json")


# ============================================================
# PARTE 2: LÓGICA DE CÁLCULO (sin entradas ni impresiones)
# ============================================================
def texto_grado(varilla):
    """'#4' -> 'Grado 40 (fy 2,810 kg/cm²); Grado 60 en marcos sismorresistentes'."""
    g = VARILLAS[varilla]["grado"]
    texto = f"Grado {g} (fy {GRADOS[g]['fy']:,} kg/cm²)"
    if g == 40 and varilla != "#2":
        texto += "; Grado 60 en marcos sismorresistentes"
    return texto


def calcular_volumen(largo, ancho, espesor):
    """Volumen de concreto en m³ de un elemento."""
    return largo * ancho * espesor


def calcular_materiales(volumen, resistencia, elemento, varilla=None):
    """Devuelve un diccionario con las cantidades de cada material."""
    varilla = varilla or elemento["varilla"]
    volumen_total = volumen * (1 + DESPERDICIO)
    suma = resistencia["cemento"] + resistencia["arena"] + resistencia["piedrin"]

    vol_cemento = VOLUMEN_SECO * (resistencia["cemento"] / suma) * volumen_total
    sacos = vol_cemento / VOLUMEN_SACO
    arena = VOLUMEN_SECO * (resistencia["arena"] / suma) * volumen_total
    piedrin = VOLUMEN_SECO * (resistencia["piedrin"] / suma) * volumen_total
    agua = sacos * KG_SACO * resistencia["agua_cemento"]
    hierro = volumen_total * elemento["hierro_kg_m3"]
    kg_por_varilla = PESO_VARILLA[varilla] * LARGO_VARILLA
    piezas = math.ceil(hierro / kg_por_varilla)

    return {
        "volumen_total": volumen_total,
        "sacos": sacos,
        "sacos_comprar": math.ceil(sacos),
        "arena_m3": arena,
        "piedrin_m3": piedrin,
        "cal_m3": 0.0,
        "agua_litros": agua,
        "hierro_kg": hierro,
        "varillas": piezas,
        "varilla": varilla,
        "acero": [{"varilla": varilla, "piezas": piezas, "kg": hierro, "metros": hierro / PESO_VARILLA[varilla]}],
    }


def mezcla(volumen_total, r):
    """Cemento, cal, arena, piedrín y agua de una mezcla (concreto o mortero)."""
    f = r.get("factor", VOLUMEN_SECO)
    suma = r["cemento"] + r.get("cal", 0) + r["arena"] + r["piedrin"]
    sacos = f * r["cemento"] / suma * volumen_total / VOLUMEN_SACO
    return {"sacos": sacos, "cal_m3": f * r.get("cal", 0) / suma * volumen_total,
            "arena_m3": f * r["arena"] / suma * volumen_total,
            "piedrin_m3": f * r["piedrin"] / suma * volumen_total,
            "agua_litros": sacos * KG_SACO * r["agua_cemento"]}


def resumir_materiales(mezclas, aceros):
    """Junta varias mezclas [(nombre, volumen, receta)] y el acero [(varilla, metros, uso)]."""
    m = {"sacos": 0.0, "cal_m3": 0.0, "arena_m3": 0.0, "piedrin_m3": 0.0, "agua_litros": 0.0,
         "hierro_kg": 0.0, "varillas": 0, "acero": [], "partes": []}
    for nombre, vol, receta in mezclas:
        if vol <= 0:
            continue
        vt = vol * (1 + DESPERDICIO)
        q = mezcla(vt, receta)
        m["partes"].append({"nombre": nombre, "vol": vol, "q": q, "receta": receta})
        for k in ("sacos", "cal_m3", "arena_m3", "piedrin_m3", "agua_litros"):
            m[k] += q[k]
    metros = {}
    for varilla, mts, _uso in aceros:
        if mts > 0:
            metros[varilla] = metros.get(varilla, 0) + mts
    for varilla in PESO_VARILLA:
        if varilla not in metros:
            continue
        kg = metros[varilla] * PESO_VARILLA[varilla]
        piezas = math.ceil(metros[varilla] * (1 + DESPERDICIO) / LARGO_VARILLA)
        m["acero"].append({"varilla": varilla, "metros": metros[varilla], "kg": kg, "piezas": piezas})
        m["hierro_kg"] += kg
        m["varillas"] += piezas
    m["sacos_comprar"] = math.ceil(m["sacos"] - 1e-9)
    return m


def manning(area, perimetro, pendiente, n):
    """Caudal (m³/s) con la fórmula de Manning."""
    radio = area / perimetro
    return area * radio ** (2 / 3) * math.sqrt(pendiente) / n, radio


def revisar_velocidad(v, avisos):
    if v < 0.6:
        avisos.append(f"Velocidad {v:.2f} m/s: menor que 0.6 m/s, puede sedimentar. Aumenta la pendiente.")
    elif v > 4.5:
        avisos.append(f"Velocidad {v:.2f} m/s: mayor que 4.5 m/s, puede erosionar el concreto.")


def pasajunta_para(espesor):
    """Diámetro de pasajunta ≈ espesor / 8."""
    d = espesor * 1000 / 8
    return "#6" if d <= 19.6 else "#7" if d <= 22.8 else "#8"


# ---------- Muro de block ----------
def calcular_muro_block(L, H, vanos, blk, clase, mort, sep, varilla, ns, res):
    w, hs, av = blk / 100, 0.20, []
    a_block = L * H - vanos - ns * L * hs
    if a_block <= 0:
        return {"error": "Las puertas, ventanas y soleras ocupan todo el muro."}
    n_block = math.ceil(a_block * 12.5 * 1.05)
    v_mort = a_block * 7.375 * 0.01 * w * 1.10
    n_cel = math.floor(L / sep + 1e-9) + 1
    v_grout = n_cel * max(0, H - ns * hs) * CELDA[blk]
    v_sol = ns * L * w * hs
    n_long = 4 if w >= 0.14 else 2
    n_est = (math.floor(L / 0.20) + 1) * ns
    l_est = 2 * (w - 0.05) + 2 * (hs - 0.05) + 0.12 if w >= 0.14 else (hs - 0.05) + 0.12
    if sep > 1.20:
        av.append(f"NSE 7.4 pide refuerzo vertical a no más de 1.20 m (tienes {sep:.2f} m).")
    if blk == 9:
        av.append("El block de 9 cm es para tabiques que no cargan (NTG 41055).")
    if blk != 9 and clase == "C":
        av.append("El block clase C no es para muros de carga; usa clase A.")
    if blk != 9 and clase == "B":
        av.append("Clase B solo para vivienda de un nivel con poca carga.")
    if mort == "N" and blk != 9:
        av.append("El mortero tipo N es para muros sin carga; usa tipo S o M.")
    m = resumir_materiales(
        [("Soleras", v_sol, res), (GROUT["nombre"], v_grout, GROUT), (MORTEROS[mort]["nombre"], v_mort, MORTEROS[mort])],
        [(varilla, n_cel * (H + 0.40), "Refuerzo vertical"), ("#3", n_long * L * 1.05 * ns, "Soleras"), ("#2", n_est * l_est, "Estribos")])
    return {"m": m, "avisos": av, "volumen": v_sol + v_grout, "hidra": None,
            "extras": [(f"Block {blk}x19x39 clase {clase}", f"{n_block}", "unidades"),
                       ("Mortero de pegado", f"{v_mort:.3f}", "m³"),
                       ("Concreto fluido en celdas", f"{v_grout:.3f}", f"m³ ({n_cel} celdas)"),
                       ("Concreto de soleras", f"{v_sol:.3f}", "m³")]}


# ---------- Muros de contención (AGIES NSE 2.1): empuje del suelo y estabilidad ----------
G_CONC, G_CICLO, G_MAMP = 2.4, 2.3, 2.2      # t/m³: concreto reforzado, concreto ciclópeo, mampostería de piedra
VACIOS_PIEDRA = 0.35                          # piedra suelta = volumen sólido / 0.65
MORT_PIEDRA = {
    "3": {"nombre": "Mortero 1 : 3 cemento-arena", "cemento": 1, "cal": 0, "arena": 3, "piedrin": 0, "agua_cemento": 0.80, "factor": 1.25},
    "4": {"nombre": "Mortero 1 : 4 cemento-arena", "cemento": 1, "cal": 0, "arena": 4, "piedrin": 0, "agua_cemento": 0.90, "factor": 1.25},
}
SUELO_DEFECTO = {"gs": 1.8, "phi": 30, "q": 0.5, "Df": 0.6, "mu": 0.55, "qa": 15, "kh": 0.15}


def error_suelo(s):
    if not 10 <= s["phi"] <= 45:
        return "El ángulo de fricción del relleno debe estar entre 10° y 45°."
    if not 1.2 <= s["gs"] <= 2.4:
        return "El peso unitario del relleno debe estar entre 1.2 y 2.4 t/m³."
    if not 0.2 <= s["mu"] <= 1:
        return "La fricción base-suelo (tan δ) debe estar entre 0.20 y 1.00."
    if s["kh"] >= 0.5 or math.atan(s["kh"]) >= math.radians(s["phi"]):
        return "El coeficiente sísmico kh es demasiado alto para ese ángulo de fricción (usa kh < 0.50)."
    return None


def empuje_suelo(s, H):
    """Rankine (estático) y Mononobe-Okabe (sismo), respaldo vertical y relleno horizontal."""
    phi = math.radians(s["phi"])
    Ka = math.tan(math.pi / 4 - phi / 2) ** 2
    o = {"Ka": Ka, "Kp": 1 / Ka, "Ps": 0.5 * Ka * s["gs"] * H * H, "Pq": Ka * s["q"] * H,
         "Pp": 0.5 / Ka * s["gs"] * s["Df"] ** 2, "KAE": Ka, "dPe": 0.0}
    if s["kh"] > 0:
        th = math.atan(s["kh"])
        r = 1 + math.sqrt(math.sin(phi) * math.sin(phi - th) / math.cos(th))
        o["KAE"] = math.cos(phi - th) ** 2 / (math.cos(th) * r * r)
        o["dPe"] = 0.5 * s["gs"] * H * H * (o["KAE"] - Ka)
    return o


def estabilidad(s, H, B, pesos):
    """pesos: [(nombre, W t/m, brazo desde la punta m, altura del centroide m)]."""
    E = empuje_suelo(s, H)
    W = sum(p[1] for p in pesos)
    Mr = sum(p[1] * p[2] for p in pesos)
    Mi = sum(p[1] * p[3] for p in pesos)
    Pa, Ma = E["Ps"] + E["Pq"], E["Ps"] * H / 3 + E["Pq"] * H / 2

    def caso(Ph, M, lv, ld, qadm, elim, etxt):
        xR = (Mr - M) / W
        e = B / 2 - xR
        ae = abs(e)
        if xR <= 0 or xR >= B:
            qmax, qmin = math.inf, 0.0
        elif ae <= B / 6:
            qmax, qmin = W / B * (1 + 6 * ae / B), W / B * (1 - 6 * ae / B)
        else:
            qmax, qmin = 2 * W / (3 * (xR if e > 0 else B - xR)), 0.0
        FSv, FSd = Mr / M, (s["mu"] * W + E["Pp"]) / Ph
        return {"FSv": FSv, "FSd": FSd, "e": e, "qmax": qmax, "qmin": qmin, "lv": lv, "ld": ld, "qadm": qadm,
                "elim": elim, "etxt": etxt, "okV": FSv >= lv, "okD": FSd >= ld, "okE": ae <= elim + 1e-9, "okQ": qmax <= qadm}

    est = caso(Pa, Ma, 2.0, 1.5, s["qa"], B / 6, "B/6")
    sis = None
    if s["kh"] > 0:
        sis = caso(Pa + E["dPe"] + s["kh"] * W, Ma + E["dPe"] * 0.6 * H + s["kh"] * Mi, 1.5, 1.1, 1.33 * s["qa"], B / 4, "B/4")
    av = []
    for c, t in ((est, ""), (sis, " con sismo")):
        if c is None:
            continue
        if not c["okV"]:
            av.append(f"Volteo{t}: FS = {c['FSv']:.2f}, menor que {c['lv']:.1f}. Alarga la base o el talón.")
        if not c["okD"]:
            av.append(f"Deslizamiento{t}: FS = {c['FSd']:.2f}, menor que {c['ld']:.1f}. Alarga la base, "
                      "profundiza el desplante o agrega un dentellón bajo la base.")
        if not c["okE"]:
            av.append(f"Excentricidad{t}: e = {abs(c['e']):.3f} m, mayor que {c['etxt']} = {c['elim']:.3f} m. Alarga la base.")
        if not c["okQ"]:
            av.append(f"Presión en el suelo{t}: {c['qmax']:.2f} t/m², mayor que la admisible de {c['qadm']:.2f} t/m².")
    return {"E": E, "W": W, "Mr": Mr, "Ma": Ma, "B": B, "est": est, "sis": sis, "avisos": av}


def acero_flexion(Mu, d, fc, fy):
    """cm² por metro para Mu (t·m/m) y d (cm); NaN si la sección no alcanza."""
    Rn = Mu * 1e5 / (0.9 * 100 * d * d)
    k = 1 - 2 * Rn / (0.85 * fc)
    return math.nan if k < 0 else 0.85 * fc / fy * (1 - math.sqrt(k)) * 100 * d


def sep_para(area, As, maximo=0.30):
    return max(0.05, min(maximo, math.floor(area / As * 100 / 2.5 + 1e-9) * 2.5 / 100))


def lloraderos(h, L, Df, largo_tubo):
    filas = max(1, math.ceil((h - Df) / 1.5 - 1e-9))
    n = filas * (math.floor(L / 1.5 + 1e-9) + 1)
    return filas, n, math.ceil(n * largo_tubo / 6 - 1e-9)


def fc_de(res):
    m = re.search(r"\((\d+)\s*kg", res["nombre"])
    return int(m.group(1)) if m else 210


def calcular_contencion_reforzado(L, H, t1, t2, B, b1, hz, v, s, res):
    hp, talon, tri = H - hz, B - b1 - t2, t2 - t1
    err = error_suelo(s)
    if err:
        return {"error": err}
    if hp <= 0:
        return {"error": "La zapata es tan gruesa como el muro."}
    if tri < 0:
        return {"error": "La pantalla no puede ser más gruesa en la corona que en la base."}
    if talon < 0:
        return {"error": "La punta más la pantalla son más anchas que la zapata."}
    pesos = [("Pantalla, parte rectangular", t1 * hp * G_CONC, b1 + tri + t1 / 2, hz + hp / 2),
             ("Pantalla, parte inclinada", tri * hp / 2 * G_CONC, b1 + tri * 2 / 3, hz + hp / 3),
             ("Zapata", B * hz * G_CONC, B / 2, hz / 2),
             ("Relleno sobre el talón", talon * hp * s["gs"], B - talon / 2, hz + hp / 2)]
    S = estabilidad(s, H, B, [p for p in pesos if p[1] > 0])
    av, Ka = S["avisos"], S["E"]["Ka"]
    fc, fyV, fy4 = fc_de(res), GRADOS[VARILLAS[v]["grado"]]["fy"], GRADOS[VARILLAS["#4"]["grado"]]["fy"]
    Wp = pesos[0][1] + pesos[1][1]
    Ms = Ka * s["gs"] * hp ** 3 / 6 + Ka * s["q"] * hp * hp / 2
    Vs = Ka * s["gs"] * hp * hp / 2 + Ka * s["q"] * hp
    dK = S["E"]["KAE"] - Ka if s["kh"] > 0 else 0
    Me = 0.5 * s["gs"] * hp * hp * dK * 0.6 * hp + s["kh"] * Wp * hp * 0.45
    Ve = 0.5 * s["gs"] * hp * hp * dK + s["kh"] * Wp
    Mu, Vu = 1.6 * Ms + Me, 1.6 * Vs + Ve
    d = t2 * 100 - 5 - VARILLAS[v]["mm"] / 20
    As_min, Asf = 0.0015 * 100 * t2 * 100, acero_flexion(Mu, d, fc, fyV)
    if math.isnan(Asf):
        return {"error": f"La pantalla de {t2:.2f} m no resiste el momento en su base (Mu = {Mu:.2f} t·m/m). Aumenta t2."}
    As = max(Asf, As_min)
    sv = sep_para(VARILLAS[v]["area"], As)
    if VARILLAS[v]["area"] / As < 0.10:
        av.append(f"Con varilla {v} habría que ponerla a menos de 10 cm. Usa una varilla más gruesa o aumenta t2.")
    phiVc = 0.75 * 0.53 * math.sqrt(fc) * 100 * d / 1000
    if Vu > phiVc:
        av.append(f"Cortante en la base de la pantalla: Vu = {Vu:.2f} t, mayor que φVc = {phiVc:.2f} t. Aumenta t2.")
    tm = (t1 + t2) / 2
    sh = sep_para(0.71, 0.0020 * 100 * tm * 100 / 2)
    qm = S["est"]["qmax"] if math.isfinite(S["est"]["qmax"]) else 0
    dz, As_minz = hz * 100 - 7.5 - 0.635, 0.0018 * 100 * hz * 100
    AsP = acero_flexion(1.6 * qm * b1 * b1 / 2, dz, fc, fy4)
    AsT = acero_flexion((1.2 * hz * G_CONC + 1.6 * (hp * s["gs"] + s["q"])) * talon * talon / 2, dz, fc, fy4)
    if math.isnan(AsP) or math.isnan(AsT):
        return {"error": f"La zapata de {hz:.2f} m no resiste la flexión de la punta o del talón. Aumenta su espesor."}
    sp, st = sep_para(1.27, max(AsP, As_minz)), sep_para(1.27, max(AsT, As_minz))
    nV, lV = math.floor(L / sv + 1e-9) + 1, hp - 0.05 + hz - 0.075 + 0.30
    nVe = math.floor(L / 0.30 + 1e-9) + 1
    nH = 2 * (math.floor(hp / sh + 1e-9) + 1)
    nP, nT = math.floor(L / sp + 1e-9) + 1, math.floor(L / st + 1e-9) + 1
    nLz = 2 * (math.floor((B - 0.15) / 0.30 + 1e-9) + 1)
    vZ, vP = B * hz * L, tm * hp * L
    filas, n_ll, tubos = lloraderos(hp, L, s["Df"], t2 + 0.1)
    m = resumir_materiales([("Zapata", vZ, res), ("Pantalla", vP, res)],
                           [(v, nV * lV, "Pantalla, cara del relleno"), ("#3", nVe * (hp + 0.35), "Pantalla, cara exterior"),
                            ("#3", nH * L * 1.05, "Pantalla, horizontal"), ("#4", nP * (B + 0.15), "Zapata, punta"),
                            ("#4", nT * (talon + t2 + 0.30), "Zapata, talón"), ("#3", nLz * L * 1.05, "Zapata, longitudinal")])
    return {"m": m, "avisos": av, "volumen": vZ + vP, "hidra": None, "estab": S,
            "extras": [("Armado de la pantalla", f"{v} a cada {sv:.3f} m", "(cara del relleno)"),
                       ("Horizontal de la pantalla", f"#3 a cada {sh:.3f} m", "en cada cara"),
                       ("Zapata", f"#4 a cada {sp:.3f} m abajo y {st:.3f} m arriba", ""),
                       ("Formaleta de la pantalla", f"{(hp + math.hypot(hp, tri)) * L:.2f}", "m²"),
                       ("Lloraderos PVC 3 pulg", f"{n_ll}", f"unidades ({filas} fila(s), {tubos} tubos de 6 m)"),
                       ("Filtro de grava", f"{0.30 * hp * L:.3f}", "m³"),
                       ("Relleno compactado sobre el talón", f"{talon * hp * L:.3f}", "m³"),
                       ("Excavación", f"{(B + 0.6) * (hz + s['Df']) * L:.3f}", "m³")]}


def calcular_gravedad(L, H, c, B, hc, pz, pp, s, receta, g, corto, cmin, pp_rango, piedra_nombre, mezcla_nombre):
    """Muro de gravedad (concreto ciclópeo o mampostería de piedra) con respaldo vertical sobre un cimiento."""
    hb, Bc, inc = H - hc, B + pz, B - c
    err = error_suelo(s)
    if err:
        return {"error": err}
    if hb <= 0:
        return {"error": "El cimiento es tan alto como el muro."}
    if c > B:
        return {"error": "La corona no puede ser más ancha que la base del muro."}
    if not 0 < pp < 100:
        return {"error": "El porcentaje de piedra debe estar entre 0 y 100 %."}
    pesos = [("Cimiento", Bc * hc * g, Bc / 2, hc / 2), ("Cuerpo, parte rectangular", c * hb * g, Bc - c / 2, hc + hb / 2),
             ("Cuerpo, parte inclinada", inc * hb / 2 * g, pz + inc * 2 / 3, hc + hb / 3)]
    S = estabilidad(s, H, Bc, [p for p in pesos if p[1] > 0])
    av = S["avisos"]
    if c < cmin:
        av.append(f"Corona de {c:.2f} m: en {corto} se recomienda al menos {cmin:.2f} m.")
    if B / H < 0.45:
        av.append(f"La base mide {B / H:.2f} H; en muros de gravedad suele ser de 0.5 a 0.7 H.")
    if not pp_rango[0] <= pp <= pp_rango[1]:
        av.append(f"En {corto} lo usual es {pp_rango[0]} a {pp_rango[1]} % de piedra.")
    V = (Bc * hc + (B + c) / 2 * hb) * L
    Vp = V * pp / 100
    piedra = Vp / (1 - VACIOS_PIEDRA) * (1 + DESPERDICIO)
    filas, n_ll, tubos = lloraderos(hb, L, s["Df"], (B + c) / 2 + 0.1)
    m = resumir_materiales([(mezcla_nombre, V - Vp, receta)], [])
    return {"m": m, "avisos": av, "volumen": V, "hidra": None, "estab": S, "V": V,
            "extras": [(piedra_nombre, f"{piedra:.2f}", "m³ (suelta)"), (mezcla_nombre, f"{V - Vp:.3f}", "m³"),
                       ("Lloraderos PVC 3 pulg", f"{n_ll}", f"unidades ({filas} fila(s), {tubos} tubos de 6 m)"),
                       ("Filtro de grava", f"{0.30 * hb * L:.3f}", "m³"),
                       ("Excavación", f"{(Bc + 0.6) * (hc + s['Df']) * L:.3f}", "m³")]}


def calcular_ciclopeo(L, H, c, B, hc, pz, pp, s, res):
    return calcular_gravedad(L, H, c, B, hc, pz, pp, s, res, G_CICLO, "concreto ciclópeo", 0.30, (25, 50),
                             "Piedra bola", "Concreto del ciclópeo")


def calcular_mamposteria(L, H, c, B, hc, pz, pp, mort, s):
    r = calcular_gravedad(L, H, c, B, hc, pz, pp, s, MORT_PIEDRA[mort], G_MAMP, "mampostería de piedra", 0.40, (60, 75),
                          "Piedra de cantera o de río", MORT_PIEDRA[mort]["nombre"])
    if "error" not in r:
        r["volumen"] = 0.0
    return r


def anchos_gavion(n, Bb, Bt):
    return [max(1, math.floor((Bb if n == 1 else Bb - (Bb - Bt) * i / (n - 1)) + 0.5)) for i in range(n)]   # redondeo igual que JavaScript


def calcular_gaviones(L, n, hc, Bb, Bt, gp, s):
    err = error_suelo(s)
    if err:
        return {"error": err}
    if n > 12:
        return {"error": "Más de 12 hiladas no es práctico en gaviones; divide el talud en terrazas."}
    if Bt > Bb:
        return {"error": "La hilada superior no puede ser más ancha que la de base."}
    if Bb < 1 or Bt < 1:
        return {"error": "Las cajas de gavión miden 1 m de ancho: cada hilada necesita al menos 1 m."}
    if not 2 <= gp <= 3:
        return {"error": "El peso unitario de la piedra debe estar entre 2.0 y 3.0 t/m³."}
    anchos = anchos_gavion(n, Bb, Bt)
    B, H, gg = anchos[0], n * hc, gp * 0.70
    pesos = [(f"Hilada {i + 1}", w * hc * gg, B - w / 2, hc * (i + 0.5)) for i, w in enumerate(anchos)]
    S = estabilidad(s, H, B, pesos)
    av = S["avisos"]
    if B < 0.5 * H:
        av.append(f"La base mide {B / H:.2f} H; en gaviones suele ser de 0.5 a 0.8 H.")
    por_c = math.ceil(L / 2 - 1e-9)
    cajas = sum(por_c * w for w in anchos)
    V = sum(w * hc * L for w in anchos)
    malla = cajas * (2 * (2 + 3 * hc) + hc)
    alambre = malla * 1.25 * 0.08
    piedra = V * 0.70 / (1 - VACIOS_PIEDRA) * (1 + DESPERDICIO)
    geo = (H + B) * L * 1.10
    m = resumir_materiales([], [])
    return {"m": m, "avisos": av, "volumen": 0.0, "hidra": None, "estab": S, "V": V,
            "extras": [("Anchos de hilada (abajo a arriba)", ", ".join(f"{w}" for w in anchos), "m"),
                       (f"Gaviones de 2 x 1 x {hc:g} m", f"{cajas}", "cajas"),
                       ("Malla hexagonal 8 x 10 cm", f"{malla:.2f}", "m²"),
                       ("Piedra de relleno 10 a 25 cm", f"{piedra:.2f}", "m³ (suelta)"),
                       ("Alambre galvanizado de amarre", f"{alambre:.1f}", "kg"),
                       ("Geotextil no tejido", f"{geo:.2f}", "m²")]}


# ---------- Pavimento rígido ----------
def calcular_pavimento(L, A, e, carriles, sj, eb, pasajuntas, res):
    av = []
    vol = L * A * e
    panos = math.ceil(L / sj - 1e-9)
    n_t, n_l, ancho_c = max(0, panos - 1), max(0, carriles - 1), A / carriles
    v_pj = pasajunta_para(e)
    m_pj = n_t * math.ceil(A / 0.30 - 1e-9) * 0.45 if pasajuntas else 0
    m_am = n_l * math.ceil(L / 0.75 - 1e-9) * 0.75
    base = L * (A + 0.60) * eb
    if e < 0.15:
        av.append("Espesor menor que 15 cm: es el mínimo usual para pavimento rígido.")
    max_j = min(4.5, 24 * e)
    if sj > max_j + 1e-9:
        av.append(f"Juntas a cada {sj:.2f} m: máximo recomendado {max_j:.2f} m (24 × espesor, ≤ 4.5 m).")
    rel = max(sj, ancho_c) / min(sj, ancho_c)
    if rel > 1.25:
        av.append(f"Losas de {sj:.2f} × {ancho_c:.2f} m (relación {rel:.2f}); conviene ≤ 1.25.")
    if ancho_c > 4.5:
        av.append(f"Carril de {ancho_c:.2f} m: agrega otra junta longitudinal (máx. 4.5 m).")
    if res is not RESISTENCIAS["4"]:
        av.append("Para pavimento se recomienda concreto de 4,000 psi (MR ≈ 4.5 MPa).")
    m = resumir_materiales([("Losas de pavimento", vol, res)],
                           [(v_pj, m_pj, "Pasajuntas lisas"), ("#4", m_am, "Barras de amarre")])
    return {"m": m, "avisos": av, "volumen": vol, "hidra": None,
            "extras": [("Losas (paños)", f"{panos * carriles}", f"de {sj:.2f} x {ancho_c:.2f} m"),
                       ("Base granular compactada", f"{base:.3f}", f"m³ (suelto ≈ {base * 1.3:.3f} m³)"),
                       ("Corte y sello de juntas", f"{n_t * A + n_l * L:.2f}", "m lineales"),
                       ("Membrana de curado", f"{0.20 * L * A:.0f}", "litros (0.20 L/m²)"),
                       ("Formaleta lateral", f"{2 * L:.2f}", "m lineales")]}


# ---------- Bordillo ----------
def calcular_bordillo(L, b1, b2, h, sj, armado, res):
    av = []
    area = (b1 + b2) / 2 * h
    vol = area * L
    n_esl, l_esl = math.floor(L / 0.30) + 1, (h - 0.15) + 0.16
    if h < 0.30:
        av.append("Un bordillo de menos de 30 cm queda muy bajo sobre el pavimento.")
    if b2 > b1:
        av.append("La base superior es mayor que la inferior: revisa las medidas.")
    aceros = [("#3", 2 * L * 1.05, "Longitudinal"), ("#2", n_esl * l_esl, "Eslabones")] if armado else []
    m = resumir_materiales([("Bordillo", vol, res)], aceros)
    return {"m": m, "avisos": av, "volumen": vol, "hidra": None,
            "extras": [("Formaleta (dos caras)", f"{2 * h * L:.2f}", "m²"),
                       ("Juntas de dilatación", f"{max(0, math.ceil(L / sj - 1e-9) - 1)}", "unidades")]}


# ---------- Cuneta tipo L ----------
def calcular_cuneta_l(L, a, t, b, h, sx_pct, T, s_pct, n, armado, res):
    av, sx, s = [], sx_pct / 100, s_pct / 100
    area = a * (t + a * sx / 2) + b * h
    vol = area * L
    q = 0.376 / n * sx ** (5 / 3) * math.sqrt(s) * T ** (8 / 3)    # Izzard (HEC-22)
    y = T * sx
    v = q / (T * y / 2)
    if y > h - t:
        av.append(f"El agua sube {y:.2f} m y rebasa el bordillo ({h - t:.2f} m visibles).")
    revisar_velocidad(v, av)
    n_esl, l_esl = math.floor(L / 0.30) + 1, (h - 0.075) + (a + b - 0.075) + 0.10
    aceros = [("#3", 3 * L * 1.05, "Longitudinal"), ("#2", n_esl * l_esl, "Eslabones en L")] if armado else []
    m = resumir_materiales([("Cuneta L", vol, res)], aceros)
    return {"m": m, "avisos": av, "volumen": vol, "hidra": {"Q": q, "V": v, "y": y},
            "extras": [("Juntas a cada 3 m", f"{max(0, math.ceil(L / 3) - 1)}", "unidades")]}


# ---------- Cuneta tipo V ----------
def calcular_cuneta_v(L, B, h, t, s_pct, n, res):
    av, s = [], s_pct / 100
    lado = math.sqrt((B / 2) ** 2 + h ** 2)
    vol = 2 * lado * t * L
    area = B * h / 2
    q, _r = manning(area, 2 * lado, s, n)
    v = q / area
    revisar_velocidad(v, av)
    if (B / 2) / h < 2:
        av.append("Taludes muy parados; junto a la calzada se recomienda 3 : 1 o más tendido.")
    m = resumir_materiales([("Cuneta V", vol, res)], [])
    return {"m": m, "avisos": av, "volumen": vol, "hidra": {"Q": q, "V": v, "y": h},
            "extras": [("Área revestida", f"{2 * lado * L:.2f}", "m²"),
                       ("Juntas a cada 3 m", f"{max(0, math.ceil(L / 3) - 1)}", "unidades")]}


# ---------- Canal ----------
def calcular_canal(L, b, H, z, tw, tf, y, s_pct, n, varilla, sep, res):
    av, s, k = [], s_pct / 100, math.sqrt(1 + z * z)
    area_c = (b + 2 * tw) * tf + 2 * tw * H if z == 0 else b * tf + 2 * tw * H * k
    vol = area_c * L
    area = (b + z * y) * y
    q, _r = manning(area, b + 2 * y * k, s, n)
    v = q / area
    froude = v / math.sqrt(9.81 * area / (b + 2 * z * y))
    if y >= H:
        av.append("El tirante es igual o mayor que la pared: el canal se rebalsa.")
    elif H - y < max(0.10, 0.2 * y):
        av.append(f"Bordo libre de {H - y:.2f} m; se recomienda al menos {max(0.10, 0.2 * y):.2f} m.")
    revisar_velocidad(v, av)
    pd = b + 2 * H * k
    aceros = [(varilla, (math.ceil(pd / sep) + 1) * L * 1.05, "Longitudinal"),
              (varilla, (math.floor(L / sep) + 1) * (pd + 0.60), "Transversal")]
    m = resumir_materiales([("Canal", vol, res)], aceros)
    return {"m": m, "avisos": av, "volumen": vol,
            "hidra": {"Q": q, "V": v, "y": y, "Fr": froude},
            "extras": [("Formaleta (dos caras)", f"{4 * H * k * L:.2f}", "m²"),
                       ("Juntas a cada 3 m", f"{max(0, math.ceil(L / 3) - 1)}", "unidades")]}


def sin_acentos(texto):
    """'Piedrín' -> 'piedrin' (para comparar nombres)."""
    t = unicodedata.normalize("NFD", str(texto))
    return "".join(c for c in t if unicodedata.category(c) != "Mn").lower().strip()


# ---------- Unidades de medida ----------
LB = 0.45359237   # kg por libra
SINONIMOS_UNIDAD = {
    "m3": ("m3", "m³", "m^3", "mt3", "metrocubico", "metroscubicos"),
    "yd3": ("yd3", "yd³", "yarda3", "yardas3", "yardacubica", "yardascubicas"),
    "pie3": ("pie3", "pie³", "pies3", "pies³", "ft3", "piecubico", "piescubicos"),
    "l": ("l", "lt", "lts", "litro", "litros"),
    "gal": ("gal", "gl", "galon", "galones"),
    "kg": ("kg", "kgs", "kilo", "kilos", "kilogramo", "kilogramos"),
    "lb": ("lb", "lbs", "libra", "libras"),
    "qq": ("qq", "quintal", "quintales"),
    "t": ("t", "ton", "tons", "tonelada", "toneladas"),
    "m": ("m", "ml", "mt", "mts", "metro", "metros", "metrolineal", "metroslineales"),
    "pie": ("pie", "pies", "ft"),
    "conteo": ("saco", "sacos", "bulto", "bultos", "bolsa", "bolsas", "pieza", "piezas", "pza", "pzas", "pz",
               "varilla", "varillas", "unidad", "unidades", "u"),
    "ciento": ("ciento", "cientos"),
    "millar": ("millar", "millares"),
}
# unidad -> (dimensión, valor en la unidad base de esa dimensión)
DIM_UNIDAD = {"m3": ("vol", 1), "yd3": ("vol", 0.764555), "pie3": ("vol", 0.0283168), "l": ("vol", 0.001),
              "gal": ("vol", 0.0037854), "kg": ("masa", 1), "lb": ("masa", LB), "qq": ("masa", 100 * LB),
              "t": ("masa", 1000), "m": ("long", 1), "pie": ("long", 0.3048),
              "conteo": ("conteo", 1), "ciento": ("conteo", 100), "millar": ("conteo", 1000)}
BASE_DIM = {"vol": "m3", "masa": "kg", "long": "m", "conteo": "conteo"}

# Unidades que se ofrecen según el tipo de material: (valor guardado, descripción)
U = {
    "sacos": ("sacos", "Sacos de 42.5 kg"), "bolsas": ("sacos", "Sacos o bolsas"), "kg": ("kg", "Kilogramos"),
    "lb": ("lb", "Libras"), "qq": ("quintales", "Quintales (100 lb)"), "t": ("toneladas", "Toneladas (1,000 kg)"),
    "m3": ("m³", "Metros cúbicos"), "yd3": ("yd³", "Yardas cúbicas"), "pie3": ("pie³", "Pies cúbicos"),
    "l": ("litros", "Litros"), "gal": ("galones", "Galones"), "cub": ("cubetas", "Cubetas de 5 galones"),
    "tonel": ("toneles", "Toneles de 55 galones"), "var": ("varillas", "Varillas de 6 m"),
    "m": ("metros", "Metros lineales"), "pie": ("pies", "Pies"), "m2": ("m²", "Metros cuadrados"),
    "u": ("unidades", "Unidades"), "ciento": ("cientos", "Cientos"), "millar": ("millares", "Millares"),
    "pz": ("piezas", "Piezas"), "pt": ("pies tablares", "Pies tablares"), "vara": ("varas", "Varas (0.84 m)"),
    "hojas": ("hojas", "Hojas o láminas"), "rollos": ("rollos", "Rollos"), "cajas": ("cajas", "Cajas"),
    "tubos": ("tubos", "Tubos"), "juegos": ("juegos", "Juegos"), "glob": ("global", "Global"),
}
TIPOS_MATERIAL = [
    ("cal y aglomerantes en polvo", r"\bcal\b|\byeso\b|pegablock|pegamix", ["kg", "lb", "qq", "bolsas", "t"]),
    ("cemento", r"cement", ["sacos", "kg", "lb", "qq", "t"]),
    ("acero de refuerzo", r"varilla|hierro|acero|pasajunta|barra", ["var", "qq", "kg", "lb", "t", "m"]),
    ("alambre y clavos", r"alambre|clavo|tornillo|perno|grapa", ["lb", "kg", "qq", "rollos", "cajas"]),
    ("agregados (arena, piedrín, grava, selecto)",
     r"arena|piedr|grava|gravill|selecto|balasto|base granular|tierra|relleno|piedra", ["m3", "yd3", "pie3"]),
    ("block, ladrillo y adoquín", r"block|bloque|ladrillo|adoquin|tayuyo|tabicon", ["u", "ciento", "millar", "m2"]),
    ("láminas, mallas y paneles", r"lamina|tablayeso|malla|geotextil|plastico|nylon|polietileno|panel",
     ["hojas", "m2", "rollos", "pz", "m"]),
    ("madera y formaleta", r"madera|tabla|regla|parral|duela|plywood|triplay|tablon|formaleta|pino",
     ["pz", "pt", "vara", "m", "hojas"]),
    ("tubería", r"tubo|tuberia|pvc|ducto|poliducto", ["tubos", "m", "pie", "pz"]),
    ("líquidos (pintura, sellador, aditivo, membrana)",
     r"pintura|sellador|membrana|aditivo|impermeab|curador|aceite|desmoldante|thinner|solvente|agua|resina|emulsion|asfalto",
     ["gal", "l", "cub", "tonel"]),
]
TIPO_GENERAL = ("materiales en general", None,
                ["pz", "u", "kg", "lb", "m", "m2", "m3", "l", "gal", "rollos", "cajas", "juegos", "glob"])


def tipo_material(nombre):
    """Devuelve (nombre del tipo, regex, claves de unidades) según el nombre del material."""
    n = sin_acentos(nombre)
    for t in TIPOS_MATERIAL:
        if re.search(t[1], n):
            return t
    return TIPO_GENERAL


def unidades_para(nombre):
    return [U[k] for k in tipo_material(nombre)[2]]


def tipo_unidad(unidad):
    u = re.sub(r"\(.*\)", "", sin_acentos(unidad))
    u = re.sub(r"[\s.]+", "", u)
    for clave, nombres in SINONIMOS_UNIDAD.items():
        if u in nombres:
            return clave
    return "otra"


def factor_base(tu, eq):
    """Cuántas unidades del cálculo vale 1 unidad 'tu'. eq: valor de 1 unidad del cálculo por dimensión."""
    if tu not in DIM_UNIDAD:
        return 0
    dim, valor = DIM_UNIDAD[tu]
    base = eq.get(BASE_DIM[dim])
    return valor / base if base else 0


def equiv_material(nombre):
    """Equivalencias propias del material: un saco de cemento pesa 42.5 kg, una varilla mide 6 m, etc."""
    if "cement" in sin_acentos(nombre):
        return {"conteo": 1, "kg": KG_SACO}
    for v in PESO_VARILLA:
        if es_varilla(nombre, v):
            return {"conteo": 1, "kg": PESO_VARILLA[v] * LARGO_VARILLA, "m": LARGO_VARILLA}
    return None


def convertir_unidad(cantidad, de, a, nombre):
    """Convierte entre unidades; devuelve None si no se puede."""
    if sin_acentos(de).replace(" ", "") == sin_acentos(a).replace(" ", ""):
        return cantidad
    td, ta = tipo_unidad(de), tipo_unidad(a)
    if td not in DIM_UNIDAD or ta not in DIM_UNIDAD:
        return None
    if DIM_UNIDAD[td][0] == DIM_UNIDAD[ta][0]:
        return cantidad * DIM_UNIDAD[td][1] / DIM_UNIDAD[ta][1]
    eq = equiv_material(nombre)
    if not eq:
        return None
    fd, fa = factor_base(td, eq), factor_base(ta, eq)
    return cantidad * fd / fa if fd and fa else None


def es_varilla(nombre, varilla):
    n = sin_acentos(nombre)
    if "varilla" not in n and "hierro" not in n:
        return False

    def token(t):
        return re.search(r"(^|[^0-9/])" + re.escape(t) + r"([^0-9/]|$)", n) is not None

    pulg = EQUIVALENCIA_VARILLA[varilla]
    return token(varilla[1:]) or ("/" in pulg and token(pulg))


# ============================================================
# PARTE 3: DATOS GUARDADOS (archivo JSON)
# ============================================================
def cargar_datos():
    if os.path.exists(ARCHIVO_DATOS):
        try:
            with open(ARCHIVO_DATOS, "r", encoding="utf-8") as f:
                datos = json.load(f)
                datos.setdefault("inventario", {})
                datos.setdefault("movimientos", [])
                datos.setdefault("calculos", [])
                return datos
        except (json.JSONDecodeError, OSError):
            print("  Aviso: no se pudo leer el archivo de datos; se empieza vacío.")
    return {"inventario": {}, "movimientos": [], "calculos": []}


def guardar_datos(datos):
    with open(ARCHIVO_DATOS, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)


def registrar_movimiento(datos, clave, delta, tipo, nota=""):
    datos["movimientos"].append({
        "material": clave, "delta": delta, "tipo": tipo, "nota": nota,
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
    })


# ============================================================
# PARTE 3B: CUENTAS DE USUARIO (cada persona crea la suya con su correo)
# ============================================================
ARCHIVO_CUENTAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cuentas_programa.json")
RX_CORREO = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]{2,}$")


def leer_cuentas():
    if os.path.exists(ARCHIVO_CUENTAS):
        try:
            with open(ARCHIVO_CUENTAS, "r", encoding="utf-8") as f:
                d = json.load(f)
                if isinstance(d.get("usuarios"), list):
                    return d
        except (json.JSONDecodeError, OSError):
            print("  Aviso: no se pudo leer el archivo de cuentas.")
    return {"usuarios": []}


def guardar_cuentas(c):
    with open(ARCHIVO_CUENTAS, "w", encoding="utf-8") as f:
        json.dump(c, f, ensure_ascii=False, indent=2)


def huella_clave(clave, sal):
    """La contraseña no se guarda tal cual: solo su huella (PBKDF2-SHA256 con sal)."""
    return hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), bytes.fromhex(sal), 120_000).hex()


def pedir_clave(mensaje):
    try:
        return getpass.getpass(mensaje)
    except (EOFError, KeyboardInterrupt):
        raise
    except Exception:          # algunas consolas no permiten ocultar lo que se escribe
        return input(mensaje)


def validar_clave(c1, c2):
    if len(c1) < 6:
        return "La contraseña debe tener al menos 6 caracteres."
    if c1 != c2:
        return "Las contraseñas no coinciden."
    return ""


def buscar_cuenta(cuentas, correo):
    return next((u for u in cuentas["usuarios"] if u["correo"] == correo), None)


def crear_cuenta(cuentas):
    print("\n  ===== CREAR MI CUENTA =====")
    print("  Escribe tu correo y elige una contraseña.")
    nombre = pedir_texto("  Nombre completo: ")
    correo = input("  Correo electrónico: ").strip().lower()
    if not RX_CORREO.match(correo):
        print("  Escribe un correo electrónico válido.")
        return None
    if buscar_cuenta(cuentas, correo):
        print("  Ya existe una cuenta con ese correo. Inicia sesión.")
        return None
    c1 = pedir_clave("  Contraseña: ")
    e = validar_clave(c1, pedir_clave("  Repetir contraseña: "))
    if e:
        print("  " + e)
        return None
    sal = secrets.token_hex(16)
    u = {"nombre": nombre, "correo": correo, "sal": sal, "huella": huella_clave(c1, sal),
         "creado": datetime.now().strftime("%Y-%m-%d %H:%M"), "ultimo": ""}
    cuentas["usuarios"].append(u)
    guardar_cuentas(cuentas)
    print(f"  Cuenta creada. ¡Bienvenido, {nombre}!")
    return u


def iniciar_sesion():
    """Pantalla de entrada: iniciar sesión o crear cuenta. Devuelve la cuenta o None para salir."""
    cuentas = leer_cuentas()
    fallos = 0
    while True:
        encabezado()
        if not cuentas["usuarios"]:
            print("  Primera vez en esta computadora: crea tu cuenta con tu correo.")
            u = crear_cuenta(cuentas)
            if u:
                break
            continue
        print("  1. Iniciar sesión")
        print("  2. Crear cuenta nueva")
        print("  0. Salir")
        op = input("  Opción: ").strip()
        if op == "0":
            return None
        if op == "2":
            u = crear_cuenta(cuentas)
            if u:
                break
            continue
        if op != "1":
            print("  Opción no válida.")
            continue
        correo = input("  Correo electrónico: ").strip().lower()
        clave = pedir_clave("  Contraseña: ")
        u = buscar_cuenta(cuentas, correo)
        if u and huella_clave(clave, u["sal"]) == u["huella"]:
            break
        fallos += 1
        print("  Correo o contraseña incorrectos.")
        if fallos >= 5:
            print("  Demasiados intentos. Vuelve a abrir el programa.")
            return None
    u["ultimo"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    guardar_cuentas(cuentas)
    return u


def menu_cuenta(_datos, usuario):
    print("\n  ===== MI CUENTA =====")
    print(f"  Nombre: {usuario['nombre']}")
    print(f"  Correo: {usuario['correo']}")
    print("  Cuenta guardada en esta computadora; la contraseña no se guarda tal cual, solo su huella cifrada.")
    if input("  ¿Cambiar mi contraseña? (s/n): ").strip().lower() != "s":
        return
    cuentas = leer_cuentas()
    u = buscar_cuenta(cuentas, usuario["correo"])
    if not u or huella_clave(pedir_clave("  Contraseña actual: "), u["sal"]) != u["huella"]:
        print("  La contraseña actual no es correcta.")
        return
    c1 = pedir_clave("  Nueva contraseña: ")
    e = validar_clave(c1, pedir_clave("  Repetir: "))
    if e:
        print("  " + e)
        return
    u["sal"] = secrets.token_hex(16)
    u["huella"] = huella_clave(c1, u["sal"])
    guardar_cuentas(cuentas)
    print("  Contraseña cambiada.")


# ============================================================
# PARTE 4: ENTRADA DE DATOS (validación)
# ============================================================
def pedir_numero(mensaje, permitir_cero=False, entero=False, defecto=None):
    while True:
        texto = input(mensaje + (f"[{defecto:g}] " if defecto is not None else "")).strip().replace(",", ".")
        if texto == "" and defecto is not None:
            return int(defecto) if entero else float(defecto)
        try:
            valor = float(texto)
        except ValueError:
            print("  Escribe un número válido.")
            continue
        if entero and valor != int(valor):
            print("  Escribe un número entero.")
            continue
        if valor < 0 or (valor == 0 and not permitir_cero):
            print("  El número debe ser mayor que 0." if not permitir_cero else "  El número no puede ser negativo.")
            continue
        return int(valor) if entero else valor


def pedir_texto(mensaje, obligatorio=True):
    while True:
        texto = input(mensaje).strip()
        if texto or not obligatorio:
            return texto
        print("  Este dato es obligatorio.")


def pedir_opcion(mensaje, opciones, defecto):
    """opciones: lista de (valor, texto). Devuelve el valor elegido."""
    for i, (_v, t) in enumerate(opciones, 1):
        print(f"    {i}. {t}")
    while True:
        texto = input(f"  {mensaje} (Enter = {dict(opciones)[defecto]}): ").strip()
        if texto == "":
            return defecto
        if texto.isdigit() and 1 <= int(texto) <= len(opciones):
            return opciones[int(texto) - 1][0]
        print("  Escribe el número de una opción.")


def pedir_resistencia(defecto):
    return RESISTENCIAS[pedir_opcion("Resistencia del concreto",
                                     [(k, r["nombre"]) for k, r in RESISTENCIAS.items()], defecto)]


def elegir_material(datos):
    inv = datos["inventario"]
    if not inv:
        print("  La bodega está vacía.")
        return None
    claves = list(inv.keys())
    for i, c in enumerate(claves, 1):
        m = inv[c]
        print(f"  {i}. {m['nombre']}: {m['cantidad']:g} {m['unidad']}")
    n = pedir_numero("  Número del material: ", entero=True)
    if 1 <= n <= len(claves):
        return claves[n - 1]
    print("  Ese número no está en la lista.")
    return None


# ============================================================
# PARTE 5: MENÚ DE INVENTARIO
# ============================================================
def ver_inventario(datos):
    inv = datos["inventario"]
    print("\n  ---------- EN BODEGA ----------")
    if not inv:
        print("  La bodega está vacía.")
        return
    print(f"  {'Material':<26}{'Cantidad':>10}  Unidad")
    for m in inv.values():
        aviso = "  (agotado)" if m["cantidad"] <= 0 else ""
        print(f"  {m['nombre']:<26}{m['cantidad']:>10g}  {m['unidad']}{aviso}")


CATALOGO_MATERIALES = [
    ("Cemento y cal", ["Cemento UGC 4060 PSI", "Cemento 5000 PSI", "Cemento resistente a sulfatos",
                       "Cemento de mampostería", "Cal hidratada", "Pegablock"]),
    ("Agregados", ["Arena de río", "Arena amarilla", "Piedrín 3/4", "Piedrín 3/8", "Piedrín 1", "Grava",
                   "Material selecto", "Base granular", "Piedra bola"]),
    ("Acero de refuerzo", ["Varilla #2", "Varilla #3", "Varilla #4", "Varilla #5", "Varilla #6", "Varilla #7",
                           "Varilla #8", "Varilla #8 lisa (pasajuntas)", "Alambre de amarre", "Clavo de 3 pulg",
                           "Clavo de 2 1/2 pulg"]),
    ("Block, ladrillo y adoquín", ["Block 14x19x39 clase A", "Block 19x19x39 clase A", "Block 9x19x39",
                                   "Ladrillo tayuyo", "Adoquín"]),
    ("Madera y formaleta", ["Tabla para formaleta", "Regla de pino", "Parales", "Plywood"]),
    ("Láminas y mallas", ["Lámina galvanizada", "Malla electrosoldada", "Tablayeso", "Nylon (polietileno)"]),
    ("Tubería", ["Tubo PVC 4 pulg", "Tubo PVC 2 pulg", "Tubo PVC 1/2 pulg"]),
    ("Líquidos", ["Sellador de juntas", "Membrana de curado", "Aditivo para concreto", "Impermeabilizante",
                  "Pintura", "Desmoldante"]),
]


def elegir_nombre_material(datos):
    """Lista de selección: primero el grupo, luego el material. Siempre hay opción de escribir otro."""
    grupos = []
    en_bodega = sorted((m["nombre"] for m in datos["inventario"].values()), key=sin_acentos)
    if en_bodega:
        grupos.append(("Ya en bodega", en_bodega))
    for nombre_g, lista in CATALOGO_MATERIALES:
        nuevos = [n for n in lista if sin_acentos(n) not in datos["inventario"]]
        if nuevos:
            grupos.append((nombre_g, nuevos))
    opciones = [(i, g[0]) for i, g in enumerate(grupos)] + [("otro", "Otro material (escribir el nombre)")]
    g = pedir_opcion("Tipo de material", opciones, opciones[0][0])
    if g == "otro":
        return pedir_texto("  Nombre del material: ")
    lista = grupos[g][1]
    return pedir_opcion("Material", [(n, n) for n in lista] + [("otro", "Otro (escribir el nombre)")], lista[0])


def agregar_material(datos):
    nombre = elegir_nombre_material(datos)
    if nombre == "otro":
        nombre = pedir_texto("  Nombre del material: ")
    print(f"  Material: {nombre}")
    clave = sin_acentos(nombre)
    inv = datos["inventario"]
    existe = inv.get(clave)
    print(f"  Unidades para {tipo_material(nombre)[0]}:")
    opciones = list(unidades_para(nombre))
    if existe and existe["unidad"] not in [v for v, _t in opciones]:
        opciones.insert(0, (existe["unidad"], f"{existe['unidad']} (como está en bodega)"))
    opciones.append(("__otra", "Otra unidad"))
    unidad = pedir_opcion("Unidad", opciones, existe["unidad"] if existe else opciones[0][0])
    if unidad == "__otra":
        unidad = pedir_texto("  Escribe la unidad: ", obligatorio=False) or "piezas"
    cantidad = pedir_numero(f"  Cantidad ({unidad}): ")
    if existe:
        convertida = convertir_unidad(cantidad, unidad, existe["unidad"], existe["nombre"])
        if convertida is None:
            print(f"  «{existe['nombre']}» se lleva en {existe['unidad']} y no se puede convertir desde {unidad}.")
            return
        if unidad != existe["unidad"]:
            print(f"  {cantidad:g} {unidad} = {convertida:.4g} {existe['unidad']}")
        cantidad = round(convertida, 6)
        existe["cantidad"] = round(existe["cantidad"] + cantidad, 6)
        print(f"  «{existe['nombre']}» ya existía: se sumaron {cantidad:g} {existe['unidad']}.")
    else:
        inv[clave] = {"nombre": nombre, "cantidad": cantidad, "unidad": unidad}
        print(f"  «{nombre}» agregado al inventario.")
    registrar_movimiento(datos, clave, cantidad, "entrada")
    guardar_datos(datos)


def retirar_material(datos):
    clave = elegir_material(datos)
    if clave is None:
        return
    m = datos["inventario"][clave]
    q = pedir_numero(f"  ¿Cuánto retiras? (hay {m['cantidad']:g} {m['unidad']}): ")
    if q > m["cantidad"]:
        print(f"  No se puede: solo hay {m['cantidad']:g} {m['unidad']}.")
        return
    nota = pedir_texto("  Destino o nota (opcional): ", obligatorio=False)
    m["cantidad"] = round(m["cantidad"] - q, 6)
    registrar_movimiento(datos, clave, -q, "salida", nota)
    guardar_datos(datos)
    print(f"  Listo. Quedan {m['cantidad']:g} {m['unidad']} de {m['nombre']}.")


def corregir_material(datos):
    clave = elegir_material(datos)
    if clave is None:
        return
    m = datos["inventario"][clave]
    nueva = pedir_numero("  Nueva cantidad (puede ser 0): ", permitir_cero=True)
    delta = round(nueva - m["cantidad"], 6)
    m["cantidad"] = nueva
    if delta:
        registrar_movimiento(datos, clave, delta, "ajuste")
    guardar_datos(datos)
    print(f"  {m['nombre']} ahora tiene {nueva:g} {m['unidad']}.")


def quitar_material(datos):
    clave = elegir_material(datos)
    if clave is None:
        return
    nombre = datos["inventario"][clave]["nombre"]
    if pedir_texto(f"  ¿Quitar «{nombre}»? (s/n): ").lower() == "s":
        del datos["inventario"][clave]
        datos["movimientos"] = [x for x in datos["movimientos"] if x["material"] != clave]
        guardar_datos(datos)
        print(f"  «{nombre}» quitado del inventario.")


def ver_kardex(datos):
    movs = datos["movimientos"][-20:]
    print("\n  ---------- ÚLTIMOS MOVIMIENTOS ----------")
    if not movs:
        print("  Sin movimientos.")
        return
    for x in reversed(movs):
        m = datos["inventario"].get(x["material"], {"nombre": x["material"], "unidad": ""})
        signo = "+" if x["delta"] > 0 else ""
        nota = f"  ({x['nota']})" if x.get("nota") else ""
        print(f"  {x['fecha']}  {x['tipo']:<8} {m['nombre']:<20} {signo}{x['delta']:g} {m['unidad']}{nota}")


def menu_inventario(datos):
    opciones = {
        "1": ("Ver inventario", ver_inventario),
        "2": ("Agregar material", agregar_material),
        "3": ("Retirar material", retirar_material),
        "4": ("Corregir conteo", corregir_material),
        "5": ("Quitar material", quitar_material),
        "6": ("Ver movimientos (kardex)", ver_kardex),
    }
    while True:
        print("\n  ===== INVENTARIO =====")
        for k, (texto, _) in opciones.items():
            print(f"  {k}. {texto}")
        print("  0. Volver")
        op = input("  Opción: ").strip()
        if op == "0":
            return
        if op in opciones:
            opciones[op][1](datos)
        else:
            print("  Opción no válida.")


# ============================================================
# PARTE 6: MENÚ DE CÁLCULO (elementos de concreto)
# ============================================================
def mostrar_resultado(volumen, res, elem, m):
    v = m["varilla"]
    print("\n  ---------- ESTIMADO ----------")
    print(f"  Elemento:        {elem['nombre']}")
    print(f"  Resistencia:     {res['nombre']}  ({res['cemento']:g} : {res['arena']:g} : {res['piedrin']:g})")
    print(f"  Volumen:         {volumen:.3f} m³  (con 5 % de desperdicio: {m['volumen_total']:.3f} m³)")
    print(f"  Cemento:         {m['sacos_comprar']} sacos  ({m['sacos']:.1f} calculados de {KG_SACO} kg)")
    print(f"  Arena:           {m['arena_m3']:.2f} m³")
    print(f"  Piedrín:         {m['piedrin_m3']:.2f} m³  (tamaño {elem['piedrin_tamano']})")
    print(f"  Agua:            {m['agua_litros']:.0f} litros")
    print(f"  Hierro:          {m['hierro_kg']:.0f} kg  (≈ {m['varillas']} varillas {v} de {LARGO_VARILLA} m)")
    print(f"  Grado de hierro: {texto_grado(v)}  (Ø {VARILLAS[v]['mm']} mm, {VARILLAS[v]['area']} cm²)")
    print(f"  Recomendación:   {elem['recomendacion']}")
    print("  Valores de referencia. El diseño final debe verificarlo un ingeniero estructural.")


def guardar_calculo(datos, elemento, volumen, res, m, varilla_texto, medidas=None, params=None):
    if input("\n  ¿Guardar este cálculo? (s/n): ").strip().lower() != "s":
        return ""
    obra = pedir_texto("  Obra o proyecto (opcional): ", obligatorio=False) or "Sin obra"
    datos["calculos"].append({
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"), "obra": obra,
        "elemento": elemento, "medidas": medidas, "params": params, "volumen": round(volumen, 4),
        "resistencia": res["nombre"], "sacos": m["sacos_comprar"],
        "arena_m3": round(m["arena_m3"], 3), "piedrin_m3": round(m["piedrin_m3"], 3),
        "agua_litros": round(m["agua_litros"]), "hierro_kg": round(m["hierro_kg"], 1),
        "varillas": m["varillas"], "varilla": varilla_texto,
    })
    guardar_datos(datos)
    print("  Cálculo guardado.")
    return obra


def menu_calculo(datos):
    print("\n  ===== CALCULAR MATERIALES =====")
    claves = list(ELEMENTOS.keys())
    for i, k in enumerate(claves, 1):
        print(f"  {i}. {ELEMENTOS[k]['nombre']}")
    n = pedir_numero("  Elemento: ", entero=True)
    if not 1 <= n <= len(claves):
        print("  Ese número no está en la lista.")
        return
    clave = claves[n - 1]
    elem = ELEMENTOS[clave]

    print("  1. Por medidas   2. Volumen conocido")
    if input("  Cómo ingresar el concreto: ").strip() == "2":
        volumen = pedir_numero("  Volumen de concreto (m³): ")
        medidas = None
    else:
        a, b, c = elem["medidas"]
        d1 = pedir_numero(f"  {a} (m): ")
        d2 = pedir_numero(f"  {b} (m): ")
        d3 = pedir_numero(f"  {c} (m): ")
        volumen = calcular_volumen(d1, d2, d3)
        medidas = [d1, d2, d3]

    res = pedir_resistencia("3")
    varilla = pedir_opcion("Varilla principal",
                           [(v, f"{v} ({EQUIVALENCIA_VARILLA[v]} pulg), {texto_grado(v)}") for v in PESO_VARILLA],
                           elem["varilla"])

    m = calcular_materiales(volumen, res, elem, varilla)
    mostrar_resultado(volumen, res, elem, m)
    obra = guardar_calculo(datos, elem["nombre"], volumen, res, m, varilla, medidas=medidas)
    ofrecer_excel(lambda: excel_elemento(clave, medidas, volumen, res, varilla, m, obra))


# ============================================================
# PARTE 7: MUROS (BLOCK Y CONTENCIÓN) Y CARRETERAS
# ============================================================
def mostrar_formula(nombre, res, r):
    m = r["m"]
    print(f"\n  ---------- {nombre.upper()} ----------")
    if res is SIN_CONCRETO:
        print(f"  Volumen del muro: {r['V']:.3f} m³")
    else:
        print(f"  Concreto:        {r['volumen']:.3f} m³  ({res['nombre']})")
    for p in m["partes"]:
        rec = p["receta"]
        dos = f"{rec['cemento']:g}" + (f" : {rec['cal']:g}" if rec.get("cal") else "") + f" : {rec['arena']:g}" + \
              (f" : {rec['piedrin']:g}" if rec["piedrin"] else "")
        print(f"    - {p['nombre']:<38} {p['vol']:.3f} m³  ({dos}), {p['q']['sacos']:.1f} sacos")
    if m["sacos"] > 0:
        print(f"  Cemento:         {m['sacos_comprar']} sacos  ({m['sacos']:.1f} calculados de {KG_SACO} kg)")
        print(f"  Arena:           {m['arena_m3']:.2f} m³")
    if m["piedrin_m3"] > 0:
        print(f"  Piedrín:         {m['piedrin_m3']:.2f} m³")
    if m["cal_m3"] > 0:
        print(f"  Cal hidratada:   {m['cal_m3'] * KG_M3_CAL:.0f} kg  (≈ {m['cal_m3']:.3f} m³)")
    if m["agua_litros"] > 0:
        print(f"  Agua:            {m['agua_litros']:.0f} litros")
    for a in m["acero"]:
        print(f"  Varilla {a['varilla']:<3}     {a['metros']:.2f} m = {a['kg']:.1f} kg ≈ {a['piezas']} varillas de 6 m, "
              f"Grado {VARILLAS[a['varilla']]['grado']}")
    if m["hierro_kg"] > 0:
        print(f"  Alambre de amarre ≈ {m['hierro_kg'] * 0.03:.1f} kg (3 % del acero)")
    for concepto, cantidad, unidad in r["extras"]:
        print(f"  {concepto:<30} {cantidad} {unidad}")
    if r.get("estab"):
        print("  Estabilidad por metro de muro:")
        mostrar_estabilidad(r["estab"])
    h = r.get("hidra")
    if h:
        extra = f", Froude {h['Fr']:.2f} ({'subcrítico' if h['Fr'] < 1 else 'supercrítico'})" if "Fr" in h else ""
        print(f"  Capacidad:       Q = {h['Q'] * 1000:.1f} L/s, V = {h['V']:.2f} m/s, tirante {h['y']:.3f} m{extra}")
    if r["avisos"]:
        print("  Revisión:")
        for a in r["avisos"]:
            print(f"    ! {a}")
    else:
        print("  Revisión:        cumple las reglas de referencia que revisa el programa.")
    print("  Valores de referencia. El diseño final debe revisarlo un ingeniero civil colegiado.")


def terminar_formula(datos, nombre, res, r, params, tipo):
    if "error" in r:
        print("  " + r["error"])
        return
    mostrar_formula(nombre, res, r)
    obra = guardar_calculo(datos, nombre, r["volumen"], res, r["m"],
                           ", ".join(a["varilla"] for a in r["m"]["acero"]), params=params)
    ofrecer_excel(lambda: excel_formula(tipo, nombre, params, res, r, obra))


def pedir_suelo():
    print("  --- Relleno y suelo de cimentación (Enter = valor sugerido) ---")
    d = SUELO_DEFECTO
    return {"gs": pedir_numero("  Peso unitario del relleno (t/m³) ", defecto=d["gs"]),
            "phi": pedir_numero("  Ángulo de fricción del relleno (°) ", defecto=d["phi"]),
            "q": pedir_numero("  Sobrecarga sobre el relleno (t/m²) ", permitir_cero=True, defecto=d["q"]),
            "Df": pedir_numero("  Relleno frente a la base, desplante (m) ", permitir_cero=True, defecto=d["Df"]),
            "mu": pedir_numero("  Fricción base-suelo, tan δ ", defecto=d["mu"]),
            "qa": pedir_numero("  Capacidad soporte admisible (t/m²) ", defecto=d["qa"]),
            "kh": pedir_numero("  Coeficiente sísmico kh (0 = sin sismo) ", permitir_cero=True, defecto=d["kh"])}


def mostrar_estabilidad(S):
    def linea(nombre, c):
        q = f"{c['qmax']:.2f}" if math.isfinite(c["qmax"]) else "sin apoyo"
        ok = lambda b: "cumple" if b else "NO CUMPLE"
        print(f"  {nombre}: volteo FS {c['FSv']:.2f} (≥ {c['lv']:.1f}, {ok(c['okV'])}); deslizamiento FS {c['FSd']:.2f} "
              f"(≥ {c['ld']:.1f}, {ok(c['okD'])})")
        print(f"  {' ' * len(nombre)}  e = {abs(c['e']):.3f} m (≤ {c['etxt']} = {c['elim']:.3f}, {ok(c['okE'])}); "
              f"q máx = {q} t/m² (≤ {c['qadm']:.2f}, {ok(c['okQ'])})")
    E = S["E"]
    print(f"  Empuje: Ka = {E['Ka']:.3f}, Pa = {E['Ps'] + E['Pq']:.2f} t/m; peso ΣW = {S['W']:.2f} t/m; "
          f"Mr = {S['Mr']:.2f}, Ma = {S['Ma']:.2f} t·m/m")
    linea("Estático", S["est"])
    if S["sis"]:
        print(f"  Sismo (Mononobe-Okabe): KAE = {E['KAE']:.3f}, ΔPAE = {E['dPe']:.2f} t/m a 0.6 H")
        linea("Con sismo", S["sis"])


SIN_CONCRETO = {"nombre": "No aplica"}


def menu_muros(datos):
    while True:
        print("\n  ===== MUROS =====")
        print("  1. Muro de block (mampostería reforzada, NSE 7.4)")
        print("  2. Muro de contención de concreto reforzado (voladizo)")
        print("  3. Muro de contención de concreto ciclópeo")
        print("  4. Muro de gaviones")
        print("  5. Muro de mampostería de piedra")
        print("  0. Volver")
        op = input("  Opción: ").strip()
        if op == "0":
            return
        if op == "1":
            menu_muro_block(datos)
        elif op == "2":
            print("\n  ===== MURO DE CONTENCIÓN DE CONCRETO REFORZADO (AGIES NSE 2.1, ACI 318) =====")
            L = pedir_numero("  Largo del muro (m) ", defecto=10)
            H = pedir_numero("  Altura total, del fondo de la zapata a la corona (m) ", defecto=3)
            t1 = pedir_numero("  Pantalla en la corona (m) ", defecto=0.25)
            t2 = pedir_numero("  Pantalla en la base (m) ", defecto=0.35)
            B = pedir_numero("  Ancho de la zapata (m) ", defecto=2.4)
            b1 = pedir_numero("  Punta, zapata al frente de la pantalla (m) ", defecto=0.6)
            hz = pedir_numero("  Espesor de la zapata (m) ", defecto=0.40)
            v = pedir_opcion("Varilla vertical de la pantalla", [("#4", "#4 (1/2 pulg)"), ("#5", "#5 (5/8 pulg)"), ("#6", "#6 (3/4 pulg)")], "#4")
            s = pedir_suelo()
            res = pedir_resistencia("3")
            r = calcular_contencion_reforzado(L, H, t1, t2, B, b1, hz, v, s, res)
            terminar_formula(datos, "Muro de contención de concreto reforzado", res, r,
                             dict(L=L, H=H, t1=t1, t2=t2, B=B, b1=b1, hz=hz, v=v, **s), "contRef")
        elif op in ("3", "5"):
            cic = op == "3"
            print("\n  ===== " + ("MURO DE CONTENCIÓN DE CONCRETO CICLÓPEO" if cic else "MURO DE MAMPOSTERÍA DE PIEDRA") + " (AGIES NSE 2.1) =====")
            L = pedir_numero("  Largo del muro (m) ", defecto=10)
            H = pedir_numero("  Altura total, con cimiento (m) ", defecto=3)
            c = pedir_numero("  Ancho de la corona (m) ", defecto=0.40)
            B = pedir_numero("  Base del cuerpo (m) ", defecto=2.0)
            hc = pedir_numero("  Alto del cimiento (m) ", defecto=0.50)
            pz = pedir_numero("  Saliente del cimiento al frente (m) ", permitir_cero=True, defecto=0.20)
            pp = pedir_numero("  Piedra bola (%) " if cic else "  Piedra (%) ", defecto=40 if cic else 70)
            mort = None if cic else pedir_opcion("Mortero de pegado", [("3", "1 : 3 cemento-arena"), ("4", "1 : 4 cemento-arena")], "3")
            s = pedir_suelo()
            if cic:
                res = pedir_resistencia("1")
                r = calcular_ciclopeo(L, H, c, B, hc, pz, pp, s, res)
            else:
                res = SIN_CONCRETO
                r = calcular_mamposteria(L, H, c, B, hc, pz, pp, mort, s)
            terminar_formula(datos, "Muro de contención de concreto ciclópeo" if cic else "Muro de mampostería de piedra", res, r,
                             dict(L=L, H=H, c=c, B=B, hc=hc, pz=pz, pp=pp, mort=mort, **s), "contCiclo" if cic else "contMamp")
        elif op == "4":
            print("\n  ===== MURO DE GAVIONES (cajas de 2 x 1 m, malla 8 x 10 cm) =====")
            L = pedir_numero("  Largo del muro (m) ", defecto=10)
            n = pedir_numero("  Número de hiladas ", entero=True, defecto=3)
            hc = pedir_opcion("Alto de cada caja", [(1.0, "1.0 m"), (0.5, "0.5 m")], 1.0)
            Bb = pedir_numero("  Ancho de la hilada de base (m) ", defecto=3)
            Bt = pedir_numero("  Ancho de la hilada superior (m) ", defecto=1)
            gp = pedir_numero("  Peso unitario de la piedra (t/m³) ", defecto=2.6)
            s = pedir_suelo()
            r = calcular_gaviones(L, n, hc, Bb, Bt, gp, s)
            terminar_formula(datos, "Muro de gaviones", SIN_CONCRETO, r, dict(L=L, n=n, hc=hc, Bb=Bb, Bt=Bt, gp=gp, **s), "gavion")
        else:
            print("  Opción no válida.")


def menu_muro_block(datos):
    print("\n  ===== MURO DE BLOCK (NSE 7.4, NTG 41054) =====")
    L = pedir_numero("  Largo del muro (m) ", defecto=6)
    H = pedir_numero("  Alto del muro (m) ", defecto=2.6)
    vanos = pedir_numero("  Área de puertas y ventanas (m²) ", permitir_cero=True, defecto=0)
    blk = pedir_opcion("Block", [(14, "14x19x39 cm (carga)"), (19, "19x19x39 cm (carga)"), (9, "9x19x39 cm (tabique)")], 14)
    clase = pedir_opcion("Clase del block", [(k, f"Clase {k}, {c['promedio']:g} kg/cm²") for k, c in CLASES_BLOCK.items()], "A")
    mort = pedir_opcion("Mortero", [(k, f"{m['nombre']} {m['cemento']}:{m['cal']:g}:{m['arena']:g}") for k, m in MORTEROS.items()], "S")
    sep = pedir_numero("  Refuerzo vertical a cada (m) ", defecto=0.80)
    varilla = pedir_opcion("Varilla vertical", [("#3", "#3 (3/8 pulg)"), ("#4", "#4 (1/2 pulg)")], "#3")
    ns = pedir_numero("  Número de soleras ", permitir_cero=True, entero=True, defecto=3)
    res = pedir_resistencia("3")
    r = calcular_muro_block(L, H, vanos, blk, clase, mort, sep, varilla, ns, res)
    terminar_formula(datos, "Muro de block", res, r, {"L": L, "H": H, "vanos": vanos, "blk": blk, "clase": clase,
                                                      "mort": mort, "s": sep, "v": varilla, "ns": ns}, "muro")


def menu_carreteras(datos):
    while True:
        print("\n  ===== CARRETERAS Y DRENAJE =====")
        print("  1. Pavimento rígido")
        print("  2. Bordillo fundido")
        print("  3. Cuneta tipo L")
        print("  4. Cuneta tipo V")
        print("  5. Canal")
        print("  0. Volver")
        op = input("  Opción: ").strip()
        if op == "0":
            return
        if op == "1":
            L = pedir_numero("  Largo del tramo (m) ", defecto=100)
            A = pedir_numero("  Ancho de calzada (m) ", defecto=6)
            e = pedir_numero("  Espesor de losa (m) ", defecto=0.18)
            c = pedir_numero("  Carriles ", entero=True, defecto=2)
            sj = pedir_numero("  Juntas transversales a cada (m) ", defecto=3)
            eb = pedir_numero("  Espesor de base granular (m) ", permitir_cero=True, defecto=0.15)
            pj = pedir_opcion("¿Pasajuntas?", [(True, "Sí"), (False, "No")], True)
            res = pedir_resistencia("4")
            terminar_formula(datos, "Pavimento rígido", res, calcular_pavimento(L, A, e, c, sj, eb, pj, res),
                             {"L": L, "A": A, "e": e, "c": c, "sj": sj, "eb": eb, "pj": pj}, "pavimento")
        elif op == "2":
            L = pedir_numero("  Largo (m) ", defecto=100)
            b1 = pedir_numero("  Base inferior (m) ", defecto=0.20)
            b2 = pedir_numero("  Base superior (m) ", defecto=0.15)
            h = pedir_numero("  Altura total (m) ", defecto=0.45)
            sj = pedir_numero("  Juntas a cada (m) ", defecto=3)
            ar = pedir_opcion("¿Con armado (2 #3 y eslabones #2)?", [(True, "Sí"), (False, "No")], True)
            res = pedir_resistencia("2")
            terminar_formula(datos, "Bordillo fundido", res, calcular_bordillo(L, b1, b2, h, sj, ar, res),
                             {"L": L, "b1": b1, "b2": b2, "h": h, "sj": sj, "ar": ar}, "bordillo")
        elif op == "3":
            L = pedir_numero("  Largo (m) ", defecto=100)
            a = pedir_numero("  Ancho de cuneta (m) ", defecto=0.50)
            t = pedir_numero("  Espesor de cuneta (m) ", defecto=0.12)
            b = pedir_numero("  Espesor de bordillo (m) ", defecto=0.15)
            h = pedir_numero("  Altura de bordillo (m) ", defecto=0.45)
            sx = pedir_numero("  Pendiente transversal (%) ", defecto=6)
            T = pedir_numero("  Ancho de agua permitido (m) ", defecto=1.0)
            s = pedir_numero("  Pendiente longitudinal (%) ", defecto=1)
            n = pedir_numero("  n de Manning ", defecto=0.016)
            ar = pedir_opcion("¿Con armado?", [(True, "Sí"), (False, "No")], False)
            res = pedir_resistencia("2")
            terminar_formula(datos, "Cuneta tipo L", res, calcular_cuneta_l(L, a, t, b, h, sx, T, s, n, ar, res),
                             {"L": L, "a": a, "t": t, "b": b, "h": h, "Sx": sx, "T": T, "S": s, "n": n, "ar": ar}, "cunetaL")
        elif op == "4":
            L = pedir_numero("  Largo (m) ", defecto=100)
            B = pedir_numero("  Ancho superior (m) ", defecto=0.80)
            h = pedir_numero("  Profundidad (m) ", defecto=0.20)
            t = pedir_numero("  Espesor del revestimiento (m) ", defecto=0.10)
            s = pedir_numero("  Pendiente longitudinal (%) ", defecto=1)
            n = pedir_numero("  n de Manning ", defecto=0.015)
            res = pedir_resistencia("2")
            terminar_formula(datos, "Cuneta tipo V", res, calcular_cuneta_v(L, B, h, t, s, n, res),
                             {"L": L, "B": B, "hv": h, "t": t, "S": s, "n": n}, "cunetaV")
        elif op == "5":
            L = pedir_numero("  Largo (m) ", defecto=50)
            b = pedir_numero("  Ancho de fondo (m) ", defecto=0.60)
            H = pedir_numero("  Altura de pared (m) ", defecto=0.70)
            z = pedir_numero("  Talud z (0 = rectangular) ", permitir_cero=True, defecto=0)
            tw = pedir_numero("  Espesor de pared (m) ", defecto=0.15)
            tf = pedir_numero("  Espesor de fondo (m) ", defecto=0.15)
            y = pedir_numero("  Tirante de diseño (m) ", defecto=0.50)
            s = pedir_numero("  Pendiente (%) ", defecto=0.5)
            n = pedir_numero("  n de Manning ", defecto=0.014)
            v = pedir_opcion("Varilla", [("#3", "#3 (3/8 pulg)"), ("#4", "#4 (1/2 pulg)")], "#3")
            sep = pedir_numero("  Varillas a cada (m) ", defecto=0.20)
            res = pedir_resistencia("3")
            terminar_formula(datos, "Canal", res, calcular_canal(L, b, H, z, tw, tf, y, s, n, v, sep, res),
                             {"L": L, "b": b, "H": H, "z": z, "tw": tw, "tf": tf, "y": y, "S": s, "n": n, "v": v, "sep": sep}, "canal")
        else:
            print("  Opción no válida.")


# ============================================================
# PARTE 8: NORMATIVA
# ============================================================
def ver_normativa(_datos=None):
    print("\n  ===== ACERO DE REFUERZO (NTG 36011 / ASTM A615) =====")
    print(f"  {'Var.':<5}{'Pulg':<6}{'Ø mm':>7}{'cm²':>7}{'kg/m':>8}{'kg/6m':>8}  {'Grado':<12}Uso")
    for v, x in VARILLAS.items():
        grado = f"{x['grado']}" + (" o 60" if x["grado"] == 40 and v != "#2" else "")
        print(f"  {v:<5}{EQUIVALENCIA_VARILLA[v]:<6}{x['mm']:>7.2f}{x['area']:>7.2f}{PESO_VARILLA[v]:>8.3f}"
              f"{PESO_VARILLA[v] * LARGO_VARILLA:>8.2f}  {grado:<12}{x['uso']}")
    for g, d in GRADOS.items():
        print(f"  Grado {g}: fy = {d['fy']:,} kg/cm² ({d['psi']} psi, {d['mpa']} MPa)")
    print("  Grado 40: vivienda de 1-2 niveles (NSE 4). Grado 60: marcos sismorresistentes (NSE 7.1), desde la #5.")
    print("  Hierro comercial (no normado): no usar en elementos estructurales.")

    print("\n  ===== CEMENTO (NTG 41095 / ASTM C1157, saco de 42.5 kg) =====")
    print("  Uso general UGC (tipo GU) ........ 4,060 psi (28 MPa): losas, vigas, columnas, pavimentos")
    print("  Estructural (GU o HE) ............ 5,000 psi (35 MPa): alta resistencia, prefabricados")
    print("  Resistente a sulfatos (MS/HS) .... cunetas, canales, alcantarillas, suelos agresivos")
    print("  Cemento de mampostería (C91) ..... solo morteros y repellos, no concreto estructural")
    print("  Nota: 4,060 psi es la resistencia del cemento en cubos de mortero, no la del concreto.")
    print("  Concreto según dosificación:")
    for r in RESISTENCIAS.values():
        print(f"    {r['nombre']:<24} {r['cemento']:g} : {r['arena']:g} : {r['piedrin']:g}  a/c {r['agua_cemento']}  -> {r['uso']}")

    print("\n  ===== BLOCK DE CONCRETO (NTG 41054) =====")
    for k, c in CLASES_BLOCK.items():
        print(f"  Clase {k}: {c['promedio']:g} kg/cm² promedio (mín. individual {c['individual']:g}), absorción ≤ {c['absorcion']} %, "
              f"marca {c['color']}. {c['uso']}.")
    print("  Medidas reales: 9x19x39, 14x19x39, 19x19x39 cm; 12.5 block por m² con sisa de 1 cm.")

    print("\n  ===== MORTERO DE PEGADO (NSE 7.4 / ASTM C270) =====")
    for k, m in MORTEROS.items():
        print(f"  Tipo {k}: {m['cemento']} : {m['cal']:g} : {m['arena']:g}  ({m['kgcm2']} kg/cm²)  {m['uso']}")

    print("\n  ===== TIPOS DE MUROS EN GUATEMALA =====")
    for nombre, norma, desc in TIPOS_MURO:
        print(f"  * {nombre} [{norma}]\n      {desc}")

    print("\n  ===== PAVIMENTO RÍGIDO Y DRENAJE (DGC 2001, AASHTO 93) =====")
    print("  Concreto MR ≥ 4.5 MPa (≈ f'c 4,000 psi); losa de 15 a 25 cm; juntas ≤ 24 x espesor y ≤ 4.5 m.")
    print("  Pasajuntas lisas Ø ≈ e/8, 45 cm a cada 30 cm; barras de amarre #4 de 75 cm a cada 75 cm.")
    print("  Bordillo típico 15/20 x 45 cm; cunetas y canales de 2,500-3,000 psi con juntas a cada 3 m.")
    print("  Velocidad del agua entre 0.6 y 4.5 m/s. Manning n: concreto 0.013-0.017, piedra 0.025.")


# ============================================================
# PARTE 9: HISTORIAL Y EXPORTACIÓN A CSV
# ============================================================
def ver_calculos(datos):
    cal = datos["calculos"]
    print("\n  ===== CÁLCULOS GUARDADOS =====")
    if not cal:
        print("  Aún no hay cálculos guardados.")
        return
    tot = {"sacos": 0, "arena_m3": 0, "piedrin_m3": 0, "hierro_kg": 0}
    for i, c in enumerate(cal, 1):
        print(f"  {i}. [{c.get('obra', 'Sin obra')}] {c['elemento']} {c['volumen']:.3f} m³, {c['resistencia']}: "
              f"{c['sacos']} sacos, {c['arena_m3']:.2f} m³ arena, {c['piedrin_m3']:.2f} m³ piedrín, {c['hierro_kg']:.0f} kg hierro")
        for k in tot:
            tot[k] += c[k]
    print(f"\n  TOTAL: {tot['sacos']} sacos, {tot['arena_m3']:.2f} m³ arena, "
          f"{tot['piedrin_m3']:.2f} m³ piedrín, {tot['hierro_kg']:.0f} kg hierro")


def exportar_csv(datos):
    base = os.path.dirname(ARCHIVO_DATOS)
    ruta_inv = os.path.join(base, "inventario.csv")
    ruta_cal = os.path.join(base, "calculos.csv")
    with open(ruta_inv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Material", "Cantidad", "Unidad"])
        for m in datos["inventario"].values():
            w.writerow([m["nombre"], m["cantidad"], m["unidad"]])
    with open(ruta_cal, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Fecha", "Obra", "Elemento", "Volumen m3", "Resistencia", "Sacos",
                    "Arena m3", "Piedrin m3", "Agua L", "Hierro kg", "Varillas"])
        for c in datos["calculos"]:
            w.writerow([c["fecha"], c.get("obra", ""), c["elemento"], c["volumen"], c["resistencia"], c["sacos"],
                        c["arena_m3"], c["piedrin_m3"], c.get("agua_litros", ""), c["hierro_kg"],
                        f"{c.get('varillas', '')} {c.get('varilla', '')}".strip()])
    print(f"  Archivos creados:\n  {ruta_inv}\n  {ruta_cal}")


# ============================================================
# PARTE 10: DIBUJOS (sección típica e isométrico)
# Se dibujan en SVG (nítido en Excel 2016/365) y también en PNG
# (copia simple sin textos para programas que no leen SVG).
# ============================================================
PALETA = {"ink": "#17202B", "steel": "#2C5F8D", "rust": "#A8432A", "conc": "#C4CAD1", "conc-d": "#A3ACB6",
          "conc-dd": "#87909B", "sand": "#D1AE68", "stone": "#9AA3AD", "stone-d": "#6F7A86", "muted": "#56616F",
          "grid": "#ECEEF1", "blanco": "#FFFFFF"}


def f2(x):
    return f"{x:,.2f}"


def color_hex(c):
    return PALETA.get(c, c)


def tono(c, f):
    """Aclara (f > 1) u oscurece (f < 1) un color."""
    h = color_hex(c).lstrip("#")
    rgb = [int(h[i:i + 2], 16) for i in (0, 2, 4)]
    rgb = [v + (255 - v) * (f - 1) if f >= 1 else v * f for v in rgb]
    return "#" + "".join(f"{max(0, min(255, round(v))):02X}" for v in rgb)


def esc_xml(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;"))


class Lienzo:
    """Guarda figuras en coordenadas de pantalla y las entrega como SVG o como PNG."""

    def __init__(self, ancho, alto):
        self.W, self.H, self.items = ancho, alto, []

    def poly(self, pts, fill, stroke="ink", sw=1.2, alfa=1.0, alfa_borde=1.0):
        self.items.append(("poly", [(float(x), float(y)) for x, y in pts], fill, stroke, sw, alfa, alfa_borde))

    def linea(self, pts, col, w=1.0):
        self.items.append(("linea", [(float(x), float(y)) for x, y in pts], col, w))

    def punto(self, x, y, r, fill):
        self.items.append(("punto", float(x), float(y), r, fill))

    def texto(self, x, y, txt, size=14, anchor="start", bold=False, col="ink", halo=False):
        self.items.append(("texto", float(x), float(y), str(txt), size, anchor, bold, col, halo))

    def svg(self):
        s = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{self.W}" height="{self.H}" viewBox="0 0 {self.W} {self.H}">',
             f'<rect width="{self.W}" height="{self.H}" fill="#FFFFFF"/>']
        for x in range(0, self.W, 20):
            s.append(f'<path d="M{x} 0V{self.H}" stroke="{PALETA["grid"]}"/>')
        for y in range(0, self.H, 20):
            s.append(f'<path d="M0 {y}H{self.W}" stroke="{PALETA["grid"]}"/>')
        for it in self.items:
            if it[0] == "poly":
                _, pts, fill, stroke, sw, a, ab = it
                p = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
                st = f' stroke="{color_hex(stroke)}" stroke-width="{sw}" stroke-opacity="{ab}" stroke-linejoin="round"' if stroke else ""
                s.append(f'<polygon points="{p}" fill="{color_hex(fill)}" fill-opacity="{a}"{st}/>')
            elif it[0] == "linea":
                _, pts, col, w = it
                d = "L".join(f"{x:.1f},{y:.1f}" for x, y in pts)
                s.append(f'<path d="M{d}" stroke="{color_hex(col)}" stroke-width="{w}" fill="none" stroke-linecap="round" stroke-linejoin="round"/>')
            elif it[0] == "punto":
                _, x, y, r, fill = it
                s.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{r}" fill="{color_hex(fill)}"/>')
            else:
                _, x, y, txt, size, anchor, bold, col, halo = it
                extra = ' paint-order="stroke" stroke="#FFFFFF" stroke-width="4"' if halo else ""
                peso = ' font-weight="700"' if bold else ""
                s.append(f'<text x="{x:.1f}" y="{y:.1f}" font-family="Times New Roman, Times, serif" font-size="{size}"'
                         f'{peso} text-anchor="{anchor}" fill="{color_hex(col)}"{extra}>{esc_xml(txt)}</text>')
        s.append("</svg>")
        return "".join(s)

    def png(self, k=2):
        """Dibujo en PNG hecho a mano (sin librerías): figuras y líneas, sin textos."""
        W, H = self.W * k, self.H * k
        buf = bytearray(b"\xff" * (W * H * 3))

        def rgb(c):
            h = color_hex(c).lstrip("#")
            return bytes(int(h[i:i + 2], 16) for i in (0, 2, 4))

        def rellenar(pts, col, a=1.0):
            pts = [(x * k, y * k) for x, y in pts]
            c = rgb(col)
            y0, y1 = max(0, int(min(p[1] for p in pts))), min(H - 1, int(max(p[1] for p in pts)) + 1)
            n = len(pts)
            for y in range(y0, y1 + 1):
                yc, xs = y + 0.5, []
                for i in range(n):
                    (xa, ya), (xb, yb) = pts[i], pts[(i + 1) % n]
                    if (ya <= yc < yb) or (yb <= yc < ya):
                        xs.append(xa + (yc - ya) * (xb - xa) / (yb - ya))
                xs.sort()
                for j in range(0, len(xs) - 1, 2):
                    xa, xb = max(0, int(round(xs[j]))), min(W, int(round(xs[j + 1])))
                    if xb <= xa:
                        continue
                    i0 = (y * W + xa) * 3
                    if a >= 0.999:
                        buf[i0:i0 + (xb - xa) * 3] = c * (xb - xa)
                    else:
                        for i in range(i0, i0 + (xb - xa) * 3, 3):
                            for q in range(3):
                                buf[i + q] = int(buf[i + q] * (1 - a) + c[q] * a)

        def trazo(pts, col, w):
            w = max(1.0, w) / 2
            for (xa, ya), (xb, yb) in zip(pts, pts[1:]):
                L = math.hypot(xb - xa, yb - ya) or 1
                nx, ny = -(yb - ya) / L * w, (xb - xa) / L * w
                rellenar([(xa + nx, ya + ny), (xb + nx, yb + ny), (xb - nx, yb - ny), (xa - nx, ya - ny)], col)

        for x in range(0, self.W, 20):
            trazo([(x, 0), (x, self.H)], "grid", 1)
        for y in range(0, self.H, 20):
            trazo([(0, y), (self.W, y)], "grid", 1)
        for it in self.items:
            if it[0] == "poly":
                _, pts, fill, stroke, sw, a, ab = it
                rellenar(pts, fill, a)
                if stroke and ab > 0.5:
                    trazo(pts + [pts[0]], stroke, sw)
            elif it[0] == "linea":
                trazo(it[1], it[2], it[3])
            elif it[0] == "punto":
                _, x, y, r, fill = it
                rellenar([(x + r * math.cos(t / 8 * math.tau), y + r * math.sin(t / 8 * math.tau)) for t in range(8)], fill)
        filas = b"".join(b"\x00" + bytes(buf[y * W * 3:(y + 1) * W * 3]) for y in range(H))

        def trozo(tipo, datos):
            return struct.pack(">I", len(datos)) + tipo + datos + struct.pack(">I", zlib.crc32(tipo + datos) & 0xFFFFFFFF)
        return (b"\x89PNG\r\n\x1a\n" + trozo(b"IHDR", struct.pack(">IIBBBBB", W, H, 8, 2, 0, 0, 0))
                + trozo(b"IDAT", zlib.compress(filas, 9)) + trozo(b"IEND", b""))


def dib_cota(L, A, B, txt, anchor="middle", dy=0):
    L.linea([A, B], "steel", 1)
    L.punto(A[0], A[1], 2.5, "steel")
    L.punto(B[0], B[1], 2.5, "steel")
    mx, my = (A[0] + B[0]) / 2, (A[1] + B[1]) / 2
    tx = mx + (-6 if anchor == "end" else 6 if anchor == "start" else 0)
    ty = my + (16 if dy > 0 else -6 if dy < 0 else 5)
    L.texto(tx, ty, txt, 14, anchor, halo=True)


def dib_seccion(o):
    """Sección 2D a escala (misma lógica que la página web)."""
    W, H, pad, padx, ex = 520, 300, 62, 96, o.get("ex", 1)
    todos = [p for s in o["shapes"] for p in s["pts"]] + list(o.get("water") or [])
    xs, ys = [p[0] for p in todos], [p[1] * ex for p in todos]
    minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
    k = min((W - 2 * padx) / max(maxx - minx, 1e-6), (H - 2 * pad) / max(maxy - miny, 1e-6))
    ox, oy = (W - (maxx - minx) * k) / 2 - minx * k, H / 2 + (maxy + miny) * k / 2

    def P(p):
        return (ox + p[0] * k, oy - p[1] * ex * k)
    L = Lienzo(W, H)
    for sh in o["shapes"]:
        L.poly([P(p) for p in sh["pts"]], sh["fill"], "ink", 1.3)
    if o.get("water"):
        L.poly([P(p) for p in o["water"]], "steel", None, alfa=0.35)
    for ln in o.get("lines", []):
        L.linea([P(p) for p in ln["pts"]], ln["s"], ln.get("w", 1))
    for b in o.get("bars", []):
        x, y = P(b)
        L.punto(x, y, 4, "rust")
    for c in o.get("cotas", []):
        A, B = P(c["a"]), P(c["b"])
        dib_cota(L, (A[0] + c["dx"], A[1] + c["dy"]), (B[0] + c["dx"], B[1] + c["dy"]), c["txt"], c.get("anchor", "middle"), c["dy"])
    L.texto(14, 24, o["titulo"], 15, bold=True)
    if o.get("nota"):
        L.texto(14, H - 12, o["nota"], 12.5, col="muted")
    return L


def dib_isometrico(cfg, D, titulo, largo, nota, suelo=False):
    """Isométrico: cada figura de la sección se extruye a lo largo de D."""
    W, H, ex = 560, 330, cfg.get("ex", 1)
    c30, s30 = math.cos(math.pi / 6), 0.5

    def P(x, y, z):
        return ((x - y) * c30, (x + y) * s30 - z * ex)

    def area(pts):
        return abs(sum(a[0] * b[1] - b[0] * a[1] for a, b in zip(pts, pts[1:] + pts[:1]))) / 2
    shapes = cfg["shapes"]
    tx, tz = [p[0] for s in shapes for p in s["pts"]], [p[1] for s in shapes for p in s["pts"]]
    caja = (max(tx) - min(tx)) * (max(tz) - min(tz)) or 1
    piezas = [s for s in shapes if len(shapes) > 6 and area(s["pts"]) < caja * 0.02]
    solidos = sorted([s for s in shapes if s not in piezas],
                     key=lambda s: (min(p[1] for p in s["pts"]), min(p[0] for p in s["pts"])))
    caras = []
    for sh in solidos:
        pts, n = sh["pts"], len(sh["pts"])
        a_s = sum(pts[i][0] * pts[(i + 1) % n][1] - pts[(i + 1) % n][0] * pts[i][1] for i in range(n))
        sg, alfa, propias = (1 if a_s >= 0 else -1), (0.35 if suelo and sh["fill"] == "sand" else 1.0), []
        for i in range(n):
            A, B = pts[i], pts[(i + 1) % n]
            dx, dz = B[0] - A[0], (B[1] - A[1]) * ex
            Ln = math.hypot(dx, dz) or 1
            nx, nz = sg * dz / Ln, -sg * dx / Ln
            if nx + nz <= 1e-6:
                continue
            propias.append(((A[0] + B[0]) / 2 + D / 2 + (A[1] + B[1]) / 2 * ex,
                            [(A[0], 0, A[1]), (B[0], 0, B[1]), (B[0], D, B[1]), (A[0], D, A[1])],
                            tono(sh["fill"], 0.8 + 0.32 * max(0, nz) + (-0.04 if nx > 0.7 else 0)), alfa))
        propias.sort(key=lambda c: c[0])
        caras += [(c[1], c[2], c[3]) for c in propias] + [([(x, D, z) for x, z in pts], color_hex(sh["fill"]), alfa)]
    caras += [([(x, D, z) for x, z in s["pts"]], color_hex(s["fill"]), 1.0) for s in piezas]
    todos = [P(*q) for c in caras for q in c[0]]
    minx, maxx = min(p[0] for p in todos), max(p[0] for p in todos)
    miny, maxy = min(p[1] for p in todos), max(p[1] for p in todos)
    n_ab = sum(1 for c in cfg.get("cotas", []) if c["dy"] != 0 and abs(c["a"][1] - c["b"][1]) < 1e-9)
    padT, padB = 48, 40 + 22 * max(1, n_ab)
    k = min((W - 180) / max(maxx - minx, 1e-6), (H - padT - padB) / max(maxy - miny, 1e-6))
    ox = (W - (maxx - minx) * k) / 2 - minx * k
    oy = padT + ((H - padT - padB) - (maxy - miny) * k) / 2 - miny * k

    def Q(x, y, z):
        p = P(x, y, z)
        return (ox + p[0] * k, oy + p[1] * k)
    L = Lienzo(W, H)
    for pts3, fill, alfa in caras:
        L.poly([Q(*q) for q in pts3], fill, "ink", 1.1, alfa, 0.35 if alfa < 1 else 1)
    if cfg.get("water") and not suelo:
        L.poly([Q(x, D, z) for x, z in cfg["water"]], "steel", None, alfa=0.35)
    for ln in cfg.get("lines", []):
        L.linea([Q(x, D, z) for x, z in ln["pts"]], ln["s"], max(0.8, ln.get("w", 1) * 0.8))
    for b in cfg.get("bars", []):
        a1, a2 = Q(b[0], D, b[1]), Q(b[0], D * 1.06, b[1])
        L.linea([a1, a2], "rust", 2.4)
        L.punto(a1[0], a1[1], 2.6, "rust")
    zmin, abajo = min(tz), sum(1 for c in cfg.get("cotas", []) if c["dy"] > 0 and abs(c["a"][1] - c["b"][1]) < 1e-9)
    for c in cfg.get("cotas", []):
        a, b, dx, dy = c["a"], c["b"], c["dx"], c["dy"]
        if dy < 0 and abs(a[1] - b[1]) < 1e-9:
            abajo += 1
            a, b, dy = (a[0], zmin), (b[0], zmin), 4 + 22 * abajo
        A, B = Q(a[0], D, a[1]), Q(b[0], D, b[1])
        A, B = (A[0] + dx, A[1] + dy), (B[0] + dx, B[1] + dy)
        mx, my = (A[0] + B[0]) / 2, (A[1] + B[1]) / 2
        an = c.get("anchor", "middle")
        L.linea([A, B], "steel", 1)
        L.texto(mx + (-6 if an == "end" else 6 if an == "start" else 0), my + (16 if dy > 0 else -6 if dy < 0 else 5), c["txt"], 13, an, halo=True)
    sol = [p for s in solidos if not (suelo and s["fill"] == "sand") for p in s["pts"]]
    if sol and largo:
        xm, zm = max(p[0] for p in sol), min(p[1] for p in sol)
        L1, L2 = Q(xm, 0, zm), Q(xm, D, zm)
        L1, L2 = (L1[0] + 12, L1[1] + 10), (L2[0] + 12, L2[1] + 10)
        L.linea([L1, L2], "steel", 1.2)
        L.punto(L1[0], L1[1], 2.5, "steel")
        L.punto(L2[0], L2[1], 2.5, "steel")
        L.texto((L1[0] + L2[0]) / 2 + 10, (L1[1] + L2[1]) / 2 + 16, largo, 13, halo=True)
    L.texto(14, 24, titulo, 15, bold=True)
    if nota:
        L.texto(14, H - 12, nota, 12, col="muted")
    return L


# ---------- secciones de cada cálculo (mismos datos que la página web) ----------
def rect(x0, y0, x1, y1, fill="conc"):
    return {"pts": [(x0, y0), (x1, y0), (x1, y1), (x0, y1)], "fill": fill}


def seccion_elemento(clave, d):
    e, (a, b, c) = ELEMENTOS[clave], d
    shapes, lines, bars, cotas, ex = [], [], [], [], 1

    def estribo(x0, y0, x1, y1):
        return {"pts": [(x0, y0), (x1, y0), (x1, y1), (x0, y1), (x0, y0)], "s": "rust", "w": 1.6}
    if clave in ("losa", "zapata"):
        ancho, esp, rec = (b if clave == "losa" else a), c, (0.025 if clave == "losa" else 0.075)
        ex = max(1, min(10, ancho / esp / 4))
        shapes.append(rect(0, 0, ancho, esp))
        if clave == "zapata":
            cw = min(0.4, ancho * 0.3)
            shapes.append(rect(ancho / 2 - cw / 2, esp, ancho / 2 + cw / 2, esp * 1.9, "conc-d"))
        n = min(40, math.floor((ancho - 2 * rec) / 0.15) + 1)
        bars += [(rec + (ancho - 2 * rec) * (0.5 if n == 1 else i / (n - 1)), rec + 0.006) for i in range(n)]
        lines.append({"pts": [(rec, rec + 0.016), (ancho - rec, rec + 0.016)], "s": "rust", "w": 2})
        cotas += [{"a": (0, 0), "b": (ancho, 0), "txt": ("Ancho = " if clave == "losa" else "Largo = ") + f2(ancho) + " m", "dx": 0, "dy": 18},
                  {"a": (ancho, 0), "b": (ancho, esp), "txt": ("e = " if clave == "losa" else "h = ") + f2(esp) + " m", "dx": 14, "dy": 0, "anchor": "start"}]
        titulo = "Sección típica de " + e["nombre"].lower()
        nota = f"Malla de varilla {e['varilla']} en ambos sentidos, recubrimiento {round(rec * 100)} cm" + (f"; escala vertical × {ex:.1f}" if ex > 1 else "") + "."
    elif clave in ("columna", "mocheta"):
        r = 0.04 if clave == "columna" else 0.025
        shapes.append(rect(0, 0, a, b))
        lines.append(estribo(r, r, a - r, b - r))
        bars += [(r + .012, r + .012), (a - r - .012, r + .012), (r + .012, b - r - .012), (a - r - .012, b - r - .012)]
        if a > 0.35:
            bars += [(a / 2, r + .012), (a / 2, b - r - .012)]
        if b > 0.35:
            bars += [(r + .012, b / 2), (a - r - .012, b / 2)]
        cotas += [{"a": (0, 0), "b": (a, 0), "txt": f"a = {f2(a)} m", "dx": 0, "dy": 18},
                  {"a": (a, 0), "b": (a, b), "txt": f"b = {f2(b)} m", "dx": 14, "dy": 0, "anchor": "start"}]
        titulo = f"Sección típica de {e['nombre'].lower()} (corte horizontal)"
        nota = f"Varillas longitudinales {e['varilla']} y estribo #2; altura {f2(c)} m."
    elif clave == "muro":
        shapes.append(rect(0, 0, b, c))
        lines.append({"pts": [(b / 2, 0.05), (b / 2, c - 0.05)], "s": "rust", "w": 2})
        y = 0.15
        while y < c - 0.05:
            bars.append((b / 2 + 0.012, y))
            y += 0.25
        cotas += [{"a": (0, 0), "b": (b, 0), "txt": f"t = {f2(b)} m", "dx": 0, "dy": 18},
                  {"a": (b, 0), "b": (b, c), "txt": f"H = {f2(c)} m", "dx": 14, "dy": 0, "anchor": "start"}]
        titulo, nota = "Sección típica de muro (corte vertical)", f"Varilla {e['varilla']} vertical al centro y horizontal a cada 0.25 m (ilustrativo)."
    else:   # viga, cimiento corrido, solera
        r = 0.075 if clave == "cimiento" else 0.04 if clave == "viga" else 0.025
        shapes.append(rect(0, 0, b, c))
        if clave == "cimiento":
            bars += [(r, r), (b / 2, r), (b - r, r)]
            lines.append({"pts": [(r - 0.015, r + 0.03), (b - r + 0.015, r + 0.03)], "s": "rust", "w": 1.6})
            shapes.append(rect(b / 2 - 0.07, c, b / 2 + 0.07, c + min(0.4, c * 1.2), "conc-d"))
        else:
            lines.append(estribo(r, r, b - r, c - r))
            bars += [(r + .012, r + .012), (b - r - .012, r + .012), (r + .012, c - r - .012), (b - r - .012, c - r - .012)]
            if clave == "viga" and b >= 0.25:
                bars.append((b / 2, r + .012))
        cotas += [{"a": (0, 0), "b": (b, 0), "txt": f"b = {f2(b)} m", "dx": 0, "dy": 18},
                  {"a": (b, 0), "b": (b, c), "txt": f"h = {f2(c)} m", "dx": 14, "dy": 0, "anchor": "start"}]
        titulo = "Sección típica de " + e["nombre"].lower()
        nota = ("3 varillas #3 con eslabón #2; encima, block de cimentación." if clave == "cimiento"
                else f"Varillas {e['varilla']} en las esquinas con estribo #2; largo {f2(a)} m.")
    return {"shapes": shapes, "lines": lines, "bars": bars, "cotas": cotas, "ex": ex, "titulo": titulo, "nota": nota}


def bloque_elemento(clave, d):
    """El elemento como bloque (largo × alto, con el ancho como profundidad) para el isométrico."""
    a, b, c = d
    if clave == "columna" or clave == "mocheta":
        largo_x, alto, prof = a, c, b
    else:
        largo_x, alto, prof = a, c, b
    vis = max(largo_x, alto, prof) * 0.07
    alto_v, prof_v = max(alto, vis), max(prof, vis)
    lines = []
    if ELEMENTOS[clave].get("armado", "") == "malla" or clave in ("losa", "zapata"):
        lines.append({"pts": [(largo_x * 0.03, alto_v * 0.3), (largo_x * 0.97, alto_v * 0.3)], "s": "rust", "w": 2})
    elif clave in ("columna", "mocheta", "muro"):
        for f in (0.2, 0.5, 0.8):
            lines.append({"pts": [(largo_x * f, alto_v * 0.03), (largo_x * f, alto_v * 0.97)], "s": "rust", "w": 2})
    else:
        for f in (0.2, 0.8):
            lines.append({"pts": [(largo_x * 0.03, alto_v * f), (largo_x * 0.97, alto_v * f)], "s": "rust", "w": 2})
    n1, n2, n3 = ELEMENTOS[clave]["medidas"]
    cotas = [{"a": (0, 0), "b": (largo_x, 0), "txt": f"{n1} {f2(a)} m", "dx": 0, "dy": 18},
             {"a": (0, 0), "b": (0, alto_v), "txt": f"{n3} {f2(c)} m", "dx": -14, "dy": 0, "anchor": "end"}]
    return {"shapes": [rect(0, 0, largo_x, alto_v)], "lines": lines, "cotas": cotas}, prof_v, f"{n2} {f2(b)} m"


def piedras(x0, x1, xb, y0, y1, r):
    out, fila, y = [], 0, y0 + r * 1.2
    while y < y1 - r * 0.9:
        xf = x0 + (x1 - x0) * (y - y0) / (y1 - y0)
        x = xf + r * 1.3 + (fila % 2) * r * 0.9
        while x < xb - r * 1.1:
            pts = []
            for k in range(7):
                ang = k / 7 * 2 * math.pi + fila
                rr = r * (0.78 + 0.22 * math.sin(k * 2.3 + x * 7))
                pts.append((x + rr * math.cos(ang), y + rr * 0.8 * math.sin(ang)))
            out.append({"pts": pts, "fill": "stone"})
            if len(out) > 180:
                return out
            x += r * 2.3
        y += r * 2.1
        fila += 1
    return out


def diag_empuje(x, H, p):
    w, qf = 0.3 * H, (p["q"] / (p["q"] + p["gs"] * H) if p["q"] > 0 else 0)
    return [(x, H), (x + w * qf, H), (x + w, 0), (x, 0)]


def seccion_formula(tipo, p):
    """Sección (o elevación del muro de block) de cada cálculo de Muros y Carreteras."""
    if tipo == "muro":
        L, H, w, ns, hs = p["L"], p["H"], p["blk"] / 100, p["ns"], 0.20
        shapes, lines, sol = [rect(0, 0, L, H)], [], []
        if ns >= 1:
            sol.append(0)
        if ns >= 3:
            sol += [(H - hs) * k / (ns - 1) for k in range(1, ns - 1)]
        if ns >= 2:
            sol.append(H - hs)
        cursos = round(H / 0.20)
        muchos = (L / 0.40) * cursos > 1500
        for i in range(1, cursos):
            y = i * 0.20
            lines.append({"pts": [(0, y), (L, y)], "s": "conc-dd", "w": 0.8})
            if not muchos:
                x = (i % 2) * 0.20
                while x < L:
                    if x > 0:
                        lines.append({"pts": [(x, y), (x, min(H, y + 0.20))], "s": "conc-dd", "w": 0.8})
                    x += 0.40
        shapes += [rect(0, y, L, y + hs, "conc-d") for y in sol]
        n_cel = math.floor(L / p["s"] + 1e-9) + 1
        lines += [{"pts": [(min(L - 0.1, 0.1 + c * p["s"]), -0.1), (min(L - 0.1, 0.1 + c * p["s"]), H + 0.15)], "s": "rust", "w": 2} for c in range(n_cel)]
        return {"shapes": shapes, "lines": lines, "titulo": "Muro de block, elevación",
                "nota": f"Celdas con refuerzo {p['v']} a cada {f2(p['s'])} m; soleras en gris oscuro. Espesor {round(w * 100)} cm.",
                "cotas": [{"a": (0, 0), "b": (L, 0), "txt": f"L = {f2(L)} m", "dx": 0, "dy": 16},
                          {"a": (0, 0), "b": (0, H), "txt": f"H = {f2(H)} m", "dx": -14, "dy": 0, "anchor": "end"}]}
    if tipo == "contRef":
        B, H, hz, b1, t1, t2 = p["B"], p["H"], p["hz"], p["b1"], p["t1"], p["t2"]
        tri, xb, hp = t2 - t1, b1 + t2, H - hz
        Df, ext, r = min(p["Df"], hp), max(0.5, 0.25 * H), 0.06
        shapes = [rect(xb, 0, B + ext, H, "sand"),
                  {"pts": [(-ext * 0.6, 0), (b1, 0), (b1 + tri * Df / hp, hz + Df), (-ext * 0.6, hz + Df)], "fill": "sand"},
                  rect(0, 0, B, hz), {"pts": [(b1, hz), (xb, hz), (xb, H), (b1 + tri, H)], "fill": "conc"}]
        yl = hz + Df + 0.15
        xl = b1 + tri * (yl - hz) / hp
        lines = [{"pts": [(xb - r, H - 0.05), (xb - r, r), (min(B - r, xb + 0.45), r)], "s": "rust", "w": 2.2},
                 {"pts": [(b1 + r, hz + 0.05), (b1 + tri + r, H - 0.05)], "s": "rust", "w": 1.6},
                 {"pts": [(r, r + 0.035), (B - r, r + 0.035)], "s": "rust", "w": 1.6},
                 {"pts": [(max(r, b1 - 0.3), hz - r), (B - r, hz - r)], "s": "rust", "w": 1.6},
                 {"pts": [(xl - 0.05, yl - 0.04), (xb + 0.05, yl + 0.04)], "s": "steel", "w": 3},
                 {"pts": [(xb + 0.04, hz + 0.1), (xb + 0.04, H - 0.1)], "s": "stone-d", "w": 4}]
        return {"shapes": shapes, "lines": lines, "water": diag_empuje(B + ext, H, p), "titulo": "Muro en voladizo, sección",
                "nota": f"Rojo: acero {p['v']} en la cara del relleno. Azul: empuje y lloradero. Gris: filtro.",
                "cotas": [{"a": (-ext * 0.6, 0), "b": (-ext * 0.6, H), "txt": f"H = {f2(H)}", "dx": -12, "dy": 0, "anchor": "end"},
                          {"a": (0, 0), "b": (B, 0), "txt": f"B = {f2(B)}", "dx": 0, "dy": 18},
                          {"a": (b1 + tri, H), "b": (xb, H), "txt": f"t₁ = {f2(t1)}", "dx": 0, "dy": -14}]}
    if tipo in ("contCiclo", "contMamp"):
        cic = tipo == "contCiclo"
        Bc, H, hc = p["B"] + p["pz"], p["H"], p["hc"]
        hb, inc = H - hc, p["B"] - p["c"]
        Df, ext = min(p["Df"], hb), max(0.5, 0.25 * H)
        cuerpo, r = ("conc", min(0.14, p["c"] / 3)) if cic else ("conc-d", min(0.2, p["c"] / 2.6))
        shapes = [rect(Bc, 0, Bc + ext, H, "sand"),
                  {"pts": [(-ext * 0.6, 0), (p["pz"], 0), (p["pz"] + inc * Df / hb, hc + Df), (-ext * 0.6, hc + Df)], "fill": "sand"},
                  rect(0, 0, Bc, hc, cuerpo), {"pts": [(p["pz"], hc), (Bc, hc), (Bc, H), (p["pz"] + inc, H)], "fill": cuerpo}]
        shapes += piedras(0, 0, Bc, 0, hc, r) + piedras(p["pz"], p["pz"] + inc, Bc, hc, H, r)
        yl = hc + Df + 0.15
        xl = p["pz"] + inc * (yl - hc) / hb
        return {"shapes": shapes, "water": diag_empuje(Bc + ext, H, p),
                "lines": [{"pts": [(xl - 0.05, yl - 0.04), (Bc + 0.05, yl + 0.04)], "s": "steel", "w": 3},
                          {"pts": [(Bc + 0.04, hc + 0.1), (Bc + 0.04, H - 0.1)], "s": "stone-d", "w": 4}],
                "cotas": [{"a": (-ext * 0.6, 0), "b": (-ext * 0.6, H), "txt": f"H = {f2(H)}", "dx": -12, "dy": 0, "anchor": "end"},
                          {"a": (0, 0), "b": (Bc, 0), "txt": f"Bc = {f2(Bc)}", "dx": 0, "dy": 18},
                          {"a": (p["pz"] + inc, H), "b": (Bc, H), "txt": f"c = {f2(p['c'])}", "dx": 0, "dy": -14}],
                "titulo": ("Muro de concreto ciclópeo" if cic else "Muro de mampostería de piedra") + ", sección",
                "nota": "Azul: empuje del relleno y lloradero. Gris: filtro de grava."}
    if tipo == "gavion":
        hc, a = p["hc"], anchos_gavion(p["n"], p["Bb"], p["Bt"])
        B, H = a[0], p["n"] * p["hc"]
        ext, Df = max(0.5, 0.25 * H), min(p["Df"], H)
        shapes = [rect(B, 0, B + ext, H, "sand"), rect(-ext * 0.6, 0, 0, Df, "sand")]
        lines = []
        for i, w in enumerate(a):
            y0, y1 = i * hc, (i + 1) * hc
            shapes.append(rect(B - w, y0, B, y1, "stone"))
            for kk in range(w):
                x0 = B - w + kk
                if kk > 0:
                    lines.append({"pts": [(x0, y0), (x0, y1)], "s": "ink", "w": 1.3})
                lines += [{"pts": [(x0, y0), (x0 + 1, y1)], "s": "stone-d", "w": 0.8}, {"pts": [(x0 + 1, y0), (x0, y1)], "s": "stone-d", "w": 0.8}]
        lines.append({"pts": [(B + 0.04, H), (B + 0.04, -0.04), (0, -0.04)], "s": "steel", "w": 2})
        return {"shapes": shapes, "lines": lines, "water": diag_empuje(B + ext, H, p), "titulo": "Muro de gaviones, sección",
                "nota": f"Cajas de {f2(hc)} m de alto. Azul: empuje y geotextil.",
                "cotas": [{"a": (-ext * 0.6, 0), "b": (-ext * 0.6, H), "txt": f"H = {f2(H)}", "dx": -12, "dy": 0, "anchor": "end"},
                          {"a": (0, 0), "b": (B, 0), "txt": f"B = {f2(B)}", "dx": 0, "dy": 18},
                          {"a": (B - a[-1], H), "b": (B, H), "txt": f"{a[-1]} m", "dx": 0, "dy": -14}]}
    if tipo == "pavimento":
        A, e, eb, c = p["A"], p["e"], p["eb"], p["c"]
        ex = max(1, min(8, A / 6 / max(e + eb, 0.1) * 1.2))
        shapes = [rect(-0.3, -eb, A + 0.3, 0, "sand")] if eb > 0 else []
        shapes += [rect(A * i / c, 0, A * (i + 1) / c, e) for i in range(c)]
        lines = [{"pts": [(A * j / c - 0.375, e / 2), (A * j / c + 0.375, e / 2)], "s": "rust", "w": 3} for j in range(1, c)]
        lines.append({"pts": [(-0.6, -eb), (A + 0.6, -eb)], "s": "stone-d", "w": 1.5})
        cotas = [{"a": (0, e), "b": (A, e), "txt": f"A = {f2(A)} m", "dx": 0, "dy": -16},
                 {"a": (A, 0), "b": (A, e), "txt": f"e = {f2(e)} m", "dx": 14, "dy": 0, "anchor": "start"}]
        if eb > 0:
            cotas.append({"a": (-0.3, -eb), "b": (-0.3, 0), "txt": f"base {f2(eb)} m", "dx": -12, "dy": 0, "anchor": "end"})
        return {"shapes": shapes, "lines": lines, "ex": ex, "cotas": cotas, "titulo": "Pavimento rígido, sección transversal",
                "nota": f"Escala vertical × {ex:.1f}. Barra de amarre #4 en cada junta longitudinal."}
    if tipo == "bordillo":
        b1, b2, h = p["b1"], p["b2"], p["h"]
        bars = [(min(b1, b2) / 2, 0.07), (min(b1, b2) / 2, h - 0.07)] if p["ar"] else []
        return {"shapes": [{"pts": [(0, 0), (b1, 0), (b2, h), (0, h)], "fill": "conc"}], "bars": bars,
                "lines": [{"pts": [(-0.25, h - 0.05), (0, h - 0.05)], "s": "stone-d", "w": 1.5}, {"pts": [(b1, h - 0.15), (b1 + 0.35, h - 0.15)], "s": "ink", "w": 2}],
                "cotas": [{"a": (0, 0), "b": (b1, 0), "txt": f"b₁ = {f2(b1)}", "dx": 0, "dy": 20},
                          {"a": (0, h), "b": (b2, h), "txt": f"b₂ = {f2(b2)}", "dx": 0, "dy": -14},
                          {"a": (0, 0), "b": (0, h), "txt": f"h = {f2(h)}", "dx": -14, "dy": 0, "anchor": "end"}],
                "titulo": "Bordillo, sección", "nota": "Cara inclinada hacia la calle; línea negra = nivel de pavimento."}
    if tipo == "cunetaL":
        sx, b, a, T, t, h = p["Sx"] / 100, p["b"], p["a"], p["T"], p["t"], p["h"]
        xr = b + max(a, T) + 0.2
        return {"shapes": [rect(0, 0, b, h), {"pts": [(b, 0), (b + a, 0), (b + a, t + a * sx), (b, t)], "fill": "conc"}],
                "water": [(b, t), (b + T, t + T * sx), (b, t + T * sx)],
                "lines": [{"pts": [(b + a, t + a * sx), (xr, t + (xr - b) * sx)], "s": "ink", "w": 2}],
                "cotas": [{"a": (b, 0), "b": (b + a, 0), "txt": f"a = {f2(a)}", "dx": 0, "dy": 20},
                          {"a": (0, 0), "b": (0, h), "txt": f"h = {f2(h)}", "dx": -14, "dy": 0, "anchor": "end"},
                          {"a": (0, h), "b": (b, h), "txt": f"b = {f2(b)}", "dx": 0, "dy": -14},
                          {"a": (b, t + T * sx), "b": (b + T, t + T * sx), "txt": f"T = {f2(T)}", "dx": 0, "dy": -14}],
                "titulo": "Cuneta tipo L, sección", "nota": "Azul: agua con el ancho permitido T. Línea negra: pavimento."}
    if tipo == "cunetaV":
        B, h = p["B"], p["hv"]
        lado = math.sqrt(B * B / 4 + h * h)
        tv = p["t"] * lado / (B / 2)
        return {"shapes": [{"pts": [(-B / 2, 0), (0, -h), (B / 2, 0), (B / 2, -tv), (0, -h - tv), (-B / 2, -tv)], "fill": "conc"}],
                "water": [(-B / 2, 0), (0, -h), (B / 2, 0)],
                "lines": [{"pts": [(-B / 2 - 0.3, 0), (-B / 2, 0)], "s": "ink", "w": 2}, {"pts": [(B / 2, 0), (B / 2 + 0.3, 0)], "s": "stone-d", "w": 1.5}],
                "cotas": [{"a": (-B / 2, 0), "b": (B / 2, 0), "txt": f"B = {f2(B)}", "dx": 0, "dy": -14},
                          {"a": (B / 2 + 0.05, -h), "b": (B / 2 + 0.05, 0), "txt": f"h = {f2(h)}", "dx": 14, "dy": 0, "anchor": "start"}],
                "titulo": "Cuneta tipo V, sección", "nota": f"Revestimiento de {round(p['t'] * 100)} cm. Azul: sección llena."}
    if tipo == "canal":
        b, H, z, tw, tf = p["b"], p["H"], p["z"], p["tw"], p["tf"]
        kk = math.sqrt(1 + z * z)
        if z == 0:
            pts = [(-tw, -tf), (b + tw, -tf), (b + tw, H), (b, H), (b, 0), (0, 0), (0, H), (-tw, H)]
        else:
            pts = [(-z * H, H), (0, 0), (b, 0), (b + z * H, H), (b + z * H + tw * kk, H), (b + tw, -tf), (-tw, -tf), (-z * H - tw * kk, H)]
        y = min(p["y"], H)
        return {"shapes": [{"pts": pts, "fill": "conc"}], "water": [(-z * y, y), (0, 0), (b, 0), (b + z * y, y)],
                "cotas": [{"a": (0, 0), "b": (b, 0), "txt": f"b = {f2(b)}", "dx": 0, "dy": -12},
                          {"a": (b + z * H + tw * kk, 0), "b": (b + z * H + tw * kk, H), "txt": f"H = {f2(H)}", "dx": 14, "dy": 0, "anchor": "start"},
                          {"a": (-z * H - tw * kk - 0.02, 0), "b": (-z * H - tw * kk - 0.02, y), "txt": f"y = {f2(p['y'])}", "dx": -12, "dy": 0, "anchor": "end"}],
                "titulo": "Canal " + ("rectangular" if z == 0 else "trapezoidal") + ", sección", "nota": "Azul: agua con el tirante de diseño."}
    return None


def dibujos_elemento(clave, medidas, volumen):
    d = medidas or [volumen ** (1 / 3)] * 3
    e = ELEMENTOS[clave]
    sec = dib_seccion(seccion_elemento(clave, d))
    cfg, prof, prof_txt = bloque_elemento(clave, d)
    iso = dib_isometrico(cfg, prof, f"Vista isométrica: {e['nombre'].lower()}", prof_txt,
                         "Esquema ilustrativo, sin escala en lo delgado.")
    cap_iso = (f"Vista isométrica: el volumen es {' × '.join(e['medidas'])} = {' × '.join(f'{x:g}' for x in d)} = {volumen:.3f} m³ (paso 1 del cálculo)."
               if medidas else f"Vista isométrica ilustrativa: el volumen fue dado directamente ({volumen:.3f} m³), sin medidas.")
    return [(sec, "Sección típica: muestra el área de concreto y dónde va el acero. Las cotas son las medidas del cálculo.", "Sección típica"),
            (iso, cap_iso, "Isométrico")]


def dibujos_formula(tipo, nombre, p):
    cfg = seccion_formula(tipo, p)
    if not cfg:
        return []
    elev = tipo == "muro"
    xs = [q[0] for s in cfg["shapes"] for q in s["pts"]]
    ys = [q[1] * cfg.get("ex", 1) for s in cfg["shapes"] for q in s["pts"]]
    S = max(max(xs) - min(xs), max(ys) - min(ys))
    Dreal = p["blk"] / 100 if elev else p["L"]
    Dm = max(Dreal, S * 0.05) if elev else min(Dreal, S * 1.3)
    largo = f"Espesor {f2(Dreal)} m" if elev else f"L = {f2(Dreal)} m" + (" (se dibuja un tramo)" if Dm < Dreal - 1e-9 else "")
    iso = dib_isometrico(cfg, Dm, "Vista isométrica: " + nombre.lower(), largo,
                         "Muro completo con sus soleras y el refuerzo en las celdas." if elev else f"La sección se repite a lo largo de {f2(Dreal)} m.",
                         suelo=tipo in ("contRef", "contCiclo", "contMamp", "gavion"))
    sec = dib_seccion(cfg)
    cap_sec = ("Elevación del muro" if elev else "Sección típica") + ": con estas medidas se calculan las áreas y volúmenes de cada parte."
    cap_iso = ("Vista isométrica: área del muro × espesor del block; soleras y celdas rellenas son las partes de concreto de la tabla de mezclas."
               if elev else f"Vista isométrica: volumen de cada parte = área de su sección × largo ({f2(Dreal)} m). Esos volúmenes están en la tabla de mezclas.")
    return [(sec, cap_sec, "Sección típica"), (iso, cap_iso, "Isométrico")]


# ============================================================
# PARTE 11: MEMORIA DE CÁLCULO EN EXCEL (.xlsx con fórmulas)
# Se arma con la librería estándar de Python (zipfile), sin instalar nada.
# Las celdas amarillas son datos; todo lo demás se calcula con fórmulas.
# ============================================================
XL = {"FIJA": 0, "TIT": 1, "ENC": 2, "TXT": 3, "N2": 4, "N3": 5, "N0": 6, "DATO": 7, "RES": 8, "RES0": 9,
      "NOTA": 10, "SUB": 11, "N4": 12, "DATOTXT": 13, "FTXT": 14}
XL_ESTILOS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              '<numFmts count="2"><numFmt numFmtId="164" formatCode="0.000"/><numFmt numFmtId="165" formatCode="0.0000"/></numFmts>'
              '<fonts count="7"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="15"/><color rgb="FF17202B"/><name val="Calibri"/></font>'
              '<font><sz val="11"/><color rgb="FF1F4E9A"/><name val="Calibri"/></font><font><b/><sz val="12"/><color rgb="FF2C5F8D"/><name val="Calibri"/></font><font><i/><sz val="10"/><color rgb="FF56616F"/><name val="Calibri"/></font><font><sz val="10"/><color rgb="FF56616F"/><name val="Consolas"/></font></fonts>'
              '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDCE3EA"/></patternFill></fill>'
              '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF4CC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFBE38A"/></patternFill></fill></fills>'
              '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFB9C0C8"/></left><right style="thin"><color rgb="FFB9C0C8"/></right><top style="thin"><color rgb="FFB9C0C8"/></top><bottom style="thin"><color rgb="FFB9C0C8"/></bottom><diagonal/></border></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="15">'
              '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
              '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
              '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>'
              '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
              '<xf numFmtId="2" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
              '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
              '<xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
              '<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
              '<xf numFmtId="2" fontId="1" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>'
              '<xf numFmtId="1" fontId="1" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>'
              '<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
              '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
              '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
              '<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1"/></xf>'
              '<xf numFmtId="0" fontId="6" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
              '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
XL_ENC = "Universidad Mariano Gálvez de Guatemala · Facultad de Ingeniería · Ingeniería Civil"


def xl_col(n):
    s, n = "", n + 1
    while n > 0:
        n, m = divmod(n - 1, 26)
        s = chr(65 + m) + s
    return s


class Hoja:
    """Filas de celdas. Una celda es texto, número o dict {v, f, s, id, m}; las fórmulas usan {nombre}."""

    def __init__(self, nombre, anchos):
        self.nombre, self.anchos, self.filas, self.refs, self.imagenes = nombre, anchos, [], {}, []

    def fila(self, celdas=()):
        r = len(self.filas) + 1
        lista = []
        for i, c in enumerate(celdas):
            if c is None:
                lista.append(None)
                continue
            if not isinstance(c, dict):
                c = {"v": c}
            c = dict(c)
            if c.get("id"):
                self.refs[c["id"]] = f"{xl_col(i)}{r}"
            lista.append(c)
        for c in lista:
            if c and c.get("f"):
                c["f"] = re.sub(r"\{([^}]+)\}", lambda m: self.refs[m.group(1)], c["f"])
        self.filas.append(lista)
        return r

    def ref(self, k):
        return self.refs[k]

    def xml(self):
        uniones, filas_xml = [], []
        for i, f in enumerate(self.filas):
            r, lineas = i + 1, 1
            celdas = []
            for j, c in enumerate(f):
                if not c:
                    continue
                if c.get("m", 1) > 1:
                    uniones.append(f"{xl_col(j)}{r}:{xl_col(j + c['m'] - 1)}{r}")
                v, s = c.get("v"), c.get("s")
                if s is None:
                    s = XL["N2"] if (isinstance(v, (int, float)) and not isinstance(v, bool)) or c.get("f") else XL["TXT"]
                if isinstance(v, str) and not c.get("f") and s in (XL["TXT"], XL["NOTA"], XL["FTXT"], XL["DATOTXT"], XL["ENC"]):
                    w = sum(self.anchos[j + k] if j + k < len(self.anchos) else 9 for k in range(c.get("m", 1)))
                    cpl = max(4, w * (1.2 if s in (XL["NOTA"], XL["FTXT"]) else 1.0) - 1)
                    lineas = max(lineas, sum(max(1, math.ceil(len(seg) / cpl)) for seg in v.split("\n")))
                ref = f"{xl_col(j)}{r}"
                if c.get("f"):
                    val = "" if v is None else (f"<v>{esc_xml(v)}</v>" if isinstance(v, str) else (f"<v>{v}</v>" if math.isfinite(v) else ""))
                    t_str = ' t="str"' if isinstance(v, str) else ""
                    celdas.append(f'<c r="{ref}" s="{s}"{t_str}><f>{esc_xml(c["f"])}</f>{val}</c>')
                elif isinstance(v, bool) or v is None or v == "":
                    celdas.append(f'<c r="{ref}" s="{s}"/>')
                elif isinstance(v, (int, float)):
                    celdas.append(f'<c r="{ref}" s="{s}"><v>{v}</v></c>' if math.isfinite(v) else f'<c r="{ref}" s="{s}"/>')
                else:
                    celdas.append(f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{esc_xml(v)}</t></is></c>')
            fija = len(f) == 1 and f[0] and f[0].get("s") == XL["FIJA"]
            alto = f' ht="{round(lineas * 13.6 + 4)}" customHeight="1"' if lineas > 1 else (' ht="15" customHeight="1"' if fija else "")
            filas_xml.append(f'<row r="{r}"{alto}>{"".join(celdas)}</row>')
        cols = "<cols>" + "".join(f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>' for i, w in enumerate(self.anchos)) + "</cols>"
        mc = (f'<mergeCells count="{len(uniones)}">' + "".join(f'<mergeCell ref="{u}"/>' for u in uniones) + "</mergeCells>") if uniones else ""
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>'
                f'<sheetFormatPr defaultRowHeight="15"/>{cols}<sheetData>{"".join(filas_xml)}</sheetData>{mc}'
                '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup orientation="portrait" fitToWidth="1" fitToHeight="0"/>'
                + ('<drawing r:id="rId1"/>' if self.imagenes else "") + "</worksheet>")


def libro_xlsx(hoja, ruta):
    """Escribe el .xlsx: hoja con fórmulas y los dibujos (SVG con copia PNG)."""
    ns_rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    tipo_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                   '<Default Extension="png" ContentType="image/png"/><Default Extension="svg" ContentType="image/svg+xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                   + ('<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' if hoja.imagenes else "")
                   + "</Types>")
        z.writestr("_rels/.rels", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{ns_rel}"><Relationship Id="rId1" Type="{tipo_rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                   f'<sheets><sheet name="{esc_xml(hoja.nombre[:31])}" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels",
                   f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{ns_rel}"><Relationship Id="rId1" Type="{tipo_rel}/worksheet" Target="worksheets/sheet1.xml"/>'
                   f'<Relationship Id="rId2" Type="{tipo_rel}/styles" Target="styles.xml"/></Relationships>')
        z.writestr("xl/styles.xml", XL_ESTILOS)
        z.writestr("xl/worksheets/sheet1.xml", hoja.xml())
        if hoja.imagenes:
            EMU, anclas, rels = 9525, [], []
            for j, im in enumerate(hoja.imagenes):
                z.writestr(f"xl/media/image{j + 1}.png", im["png"])
                z.writestr(f"xl/media/image{j + 1}.svg", im["svg"])
                rels.append(f'<Relationship Id="rId{2 * j + 1}" Type="{tipo_rel}/image" Target="../media/image{j + 1}.png"/>'
                            f'<Relationship Id="rId{2 * j + 2}" Type="{tipo_rel}/image" Target="../media/image{j + 1}.svg"/>')
                cx, cy = im["w"] * EMU, im["h"] * EMU
                anclas.append(
                    f'<xdr:oneCellAnchor><xdr:from><xdr:col>{im["col"]}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{im["row"]}</xdr:row><xdr:rowOff>{4 * EMU}</xdr:rowOff></xdr:from>'
                    f'<xdr:ext cx="{cx}" cy="{cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{j + 2}" name="{esc_xml(im["nombre"])}" descr="{esc_xml(im["nombre"])}"/>'
                    '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
                    f'<xdr:blipFill><a:blip r:embed="rId{2 * j + 1}"><a:extLst><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}">'
                    f'<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId{2 * j + 2}"/></a:ext></a:extLst></a:blip>'
                    '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
                    f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                    '<a:ln w="9525"><a:solidFill><a:srgbClr val="B9C0C8"/></a:solidFill></a:ln></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>')
            z.writestr("xl/worksheets/_rels/sheet1.xml.rels",
                       f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{ns_rel}"><Relationship Id="rId1" Type="{tipo_rel}/drawing" Target="../drawings/drawing1.xml"/></Relationships>')
            z.writestr("xl/drawings/drawing1.xml",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                       'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                       + "".join(anclas) + "</xdr:wsDr>")
            z.writestr("xl/drawings/_rels/drawing1.xml.rels",
                       f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{ns_rel}">' + "".join(rels) + "</Relationships>")


def xl_encabezado(H, titulo, sub):
    H.fila([{"v": titulo, "s": XL["TIT"]}])
    H.fila([{"v": XL_ENC, "s": XL["NOTA"], "m": 5}])
    H.fila([{"v": sub, "s": XL["NOTA"], "m": 5}])
    H.fila([{"v": "Las celdas amarillas son datos: si las cambias, Excel recalcula todo. La columna «Fórmula en Excel» muestra la fórmula "
                  "de cada resultado. Los dibujos muestran de dónde salen las medidas del cálculo.", "s": XL["NOTA"], "m": 5}])
    H.fila()


def xl_dibujos(H, dibujos, columnas):
    if not dibujos:
        return
    H.fila([{"v": "DIBUJOS: SECCIÓN TÍPICA E ISOMÉTRICO", "s": XL["SUB"]}])
    fila0, alto = len(H.filas), 0
    for (lienzo, _cap, nombre), (col, w, _m) in zip(dibujos, columnas):
        h = round(w * lienzo.H / lienzo.W)
        alto = max(alto, h)
        H.imagenes.append({"svg": lienzo.svg(), "png": lienzo.png(), "col": col, "row": fila0, "w": w, "h": h, "nombre": nombre})
    for _ in range(math.ceil(alto / 20) + 1):
        H.fila([{"v": "", "s": XL["FIJA"]}])
    caps = [None] * 6
    for (_l, cap, _n), (col, _w, m) in zip(dibujos, columnas):
        caps[col] = {"v": cap, "s": XL["NOTA"], "m": m}
    H.fila(caps)
    H.fila()


def xl_constantes(H):
    H.fila([{"v": "CONSTANTES", "s": XL["SUB"]}])
    H.fila([{"v": t, "s": XL["ENC"]} for t in ("Concepto", "Valor", "Unidad", "Nota")])
    H.fila(["Desperdicio", {"v": DESPERDICIO, "s": XL["DATO"], "id": "desp"}, "fracción", "5 % de desperdicio"])
    H.fila(["Factor de volumen seco (concreto)", {"v": VOLUMEN_SECO, "s": XL["DATO"], "id": "fseco"}, "", "Volumen suelto ÷ volumen compacto"])
    H.fila(["Volumen de un saco de cemento", {"v": VOLUMEN_SACO, "s": XL["DATO"], "id": "vsaco"}, "m³", f"Saco de {KG_SACO} kg"])
    H.fila(["Peso de un saco de cemento", {"v": KG_SACO, "s": XL["DATO"], "id": "kgsaco"}, "kg", ""])
    H.fila(["Largo comercial de la varilla", {"v": LARGO_VARILLA, "s": XL["DATO"], "id": "lvar"}, "m", ""])
    H.fila()


def xl_paso(H, ident, concepto, f, v, unidad, desc, estilo=None):
    visible = "=" + re.sub(r"\{([^}]+)\}", lambda m: H.refs.get(m.group(1), "?"), f)
    H.fila([concepto, {"f": f, "v": v, "s": XL["N3"] if estilo is None else estilo, "id": ident}, unidad, {"v": desc, "s": XL["TXT"]},
            {"f": "_xlfn.FORMULATEXT({" + ident + "})", "v": visible, "s": XL["FTXT"]}])


def slug(texto):
    return re.sub(r"[^a-z0-9]+", "-", sin_acentos(texto)).strip("-")[:60] or "memoria"


def ruta_memoria(nombre, obra):
    carpeta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "memorias")
    os.makedirs(carpeta, exist_ok=True)
    base = "memoria_" + slug(nombre) + ("_" + slug(obra) if obra and obra != "Sin obra" else "") + "_" + datetime.now().strftime("%Y%m%d-%H%M%S")
    return os.path.join(carpeta, base + ".xlsx")


def excel_elemento(clave, medidas, volumen, res, varilla, m, obra=""):
    """Memoria de losa, viga, columna, zapata, muro, cimiento, solera o mocheta, con todas las fórmulas."""
    e = ELEMENTOS[clave]
    H = Hoja("Memoria de cálculo", [40, 16, 11, 46, 34])
    xl_encabezado(H, f"Memoria de cálculo: {e['nombre']}" + (f", {obra}" if obra and obra != "Sin obra" else ""),
                  f"Concreto {res['nombre']}. Fecha: {datetime.now().strftime('%d/%m/%Y')}.")
    xl_dibujos(H, dibujos_elemento(clave, medidas, volumen), [(0, 470, 3), (3, 500, 2)])
    H.fila([{"v": "DATOS DE ENTRADA", "s": XL["SUB"]}])
    H.fila([{"v": t, "s": XL["ENC"]} for t in ("Dato", "Valor", "Unidad", "Nota")])
    if medidas:
        for i, (nom, val) in enumerate(zip(e["medidas"], medidas)):
            H.fila([nom, {"v": val, "s": XL["DATO"], "id": f"d{i}"}, "m", ""])
    else:
        H.fila(["Volumen de concreto", {"v": volumen, "s": XL["DATO"], "id": "vol"}, "m³", "Dado por el usuario"])
    H.fila(["Resistencia del concreto", {"v": res["nombre"], "s": XL["DATOTXT"]}, "", ""])
    H.fila(["Proporción de cemento", {"v": res["cemento"], "s": XL["DATO"], "id": "pc"}, "partes", "Dosificación por volumen"])
    H.fila(["Proporción de arena", {"v": res["arena"], "s": XL["DATO"], "id": "pa"}, "partes", ""])
    H.fila(["Proporción de piedrín", {"v": res["piedrin"], "s": XL["DATO"], "id": "pp"}, "partes", f"Tamaño {e['piedrin_tamano']}"])
    H.fila(["Relación agua / cemento", {"v": res["agua_cemento"], "s": XL["DATO"], "id": "ac"}, "", ""])
    H.fila(["Cuantía de acero de referencia", {"v": e["hierro_kg_m3"], "s": XL["DATO"], "id": "cuantia"}, "kg/m³", ""])
    H.fila(["Varilla", {"v": f"{varilla} ({EQUIVALENCIA_VARILLA[varilla]} pulg)", "s": XL["DATOTXT"]}, "", texto_grado(varilla)])
    H.fila(["Peso de la varilla", {"v": PESO_VARILLA[varilla], "s": XL["DATO"], "id": "pvar"}, "kg/m", ""])
    H.fila()
    xl_constantes(H)
    H.fila([{"v": "CÁLCULO PASO A PASO", "s": XL["SUB"]}])
    H.fila([{"v": t, "s": XL["ENC"]} for t in ("Concepto", "Resultado", "Unidad", "Descripción", "Fórmula en Excel")])
    if medidas:
        xl_paso(H, "V", "1. Volumen geométrico", "{d0}*{d1}*{d2}", volumen, "m³", "V = " + " × ".join(e["medidas"]))
    else:
        xl_paso(H, "V", "1. Volumen", "{vol}", volumen, "m³", "Volumen dado")
    suma = res["cemento"] + res["arena"] + res["piedrin"]
    xl_paso(H, "Vt", "2. Volumen con desperdicio", "{V}*(1+{desp})", m["volumen_total"], "m³", "Vt = V × (1 + desperdicio)")
    xl_paso(H, "suma", "3. Suma de proporciones", "{pc}+{pa}+{pp}", suma, "partes", "cemento + arena + piedrín", XL["N2"])
    xl_paso(H, "Vc", "4. Volumen de cemento", "{fseco}*{pc}/{suma}*{Vt}", m["sacos"] * VOLUMEN_SACO, "m³", "Vc = 1.54 × cemento / suma × Vt", XL["N4"])
    xl_paso(H, "sacos", "5. Sacos de cemento calculados", "{Vc}/{vsaco}", m["sacos"], "sacos", "Sacos = Vc / volumen de un saco", XL["N2"])
    xl_paso(H, "sacosC", "6. Sacos de cemento a comprar", "ROUNDUP({sacos},0)", m["sacos_comprar"], "sacos", "Se redondea hacia arriba", XL["RES0"])
    xl_paso(H, "arena", "7. Arena", "{fseco}*{pa}/{suma}*{Vt}", m["arena_m3"], "m³", "1.54 × arena / suma × Vt", XL["RES"])
    xl_paso(H, "piedrin", "8. Piedrín", "{fseco}*{pp}/{suma}*{Vt}", m["piedrin_m3"], "m³", "1.54 × piedrín / suma × Vt", XL["RES"])
    xl_paso(H, "agua", "9. Agua", "{sacos}*{kgsaco}*{ac}", m["agua_litros"], "litros", "Sacos × 42.5 kg × a/c", XL["RES"])
    xl_paso(H, "hierro", "10. Acero de refuerzo", "{Vt}*{cuantia}", m["hierro_kg"], "kg", "Vt × cuantía de referencia", XL["RES"])
    xl_paso(H, "kgvar", "11. Peso de una varilla", "{pvar}*{lvar}", PESO_VARILLA[varilla] * LARGO_VARILLA, "kg", "Peso por metro × largo", XL["N2"])
    xl_paso(H, "varillas", "12. Varillas a comprar", "ROUNDUP({hierro}/{kgvar},0)", m["varillas"], "varillas", "Hierro / peso de una varilla, hacia arriba", XL["RES0"])
    H.fila()
    H.fila([{"v": "RESUMEN DE MATERIALES", "s": XL["SUB"]}])
    H.fila([{"v": t, "s": XL["ENC"]} for t in ("Material", "Cantidad", "Unidad")])
    H.fila([f"Cemento (sacos de {KG_SACO} kg)", {"f": "{sacosC}", "v": m["sacos_comprar"], "s": XL["RES0"]}, "sacos"])
    H.fila(["Arena", {"f": "{arena}", "v": m["arena_m3"], "s": XL["RES"]}, "m³"])
    H.fila([f"Piedrín ({e['piedrin_tamano']})", {"f": "{piedrin}", "v": m["piedrin_m3"], "s": XL["RES"]}, "m³"])
    H.fila(["Agua", {"f": "{agua}", "v": m["agua_litros"], "s": XL["RES"]}, "litros"])
    H.fila([f"Hierro {varilla}", {"f": "{hierro}", "v": m["hierro_kg"], "s": XL["RES"]}, "kg"])
    H.fila([f"Varillas {varilla} de {LARGO_VARILLA} m", {"f": "{varillas}", "v": m["varillas"], "s": XL["RES0"]}, "varillas"])
    H.fila()
    H.fila([{"v": "Recomendación de armado: " + e["recomendacion"], "s": XL["NOTA"], "m": 5}])
    H.fila([{"v": "Valores de referencia para un estimado. El diseño final debe verificarlo un ingeniero estructural.", "s": XL["NOTA"], "m": 5}])
    ruta = ruta_memoria(e["nombre"], obra)
    libro_xlsx(H, ruta)
    return ruta


SUELO_ETIQ = [("gs", "Peso unitario del relleno", "t/m³"), ("phi", "Ángulo de fricción del relleno", "°"), ("q", "Sobrecarga sobre el relleno", "t/m²"),
              ("Df", "Relleno frente a la base (desplante)", "m"), ("mu", "Fricción base-suelo (tan δ)", ""), ("qa", "Capacidad soporte admisible", "t/m²"),
              ("kh", "Coeficiente sísmico kh (0 = sin sismo)", "")]
ENTRADAS = {
    "muro": [("L", "Largo del muro", "m"), ("H", "Alto del muro", "m"), ("vanos", "Área de puertas y ventanas", "m²"), ("blk", "Block (ancho)", "cm"),
             ("clase", "Clase del block", ""), ("mort", "Mortero de pegado (tipo)", ""), ("s", "Refuerzo vertical a cada", "m"),
             ("v", "Varilla vertical", ""), ("ns", "Número de soleras", "")],
    "contRef": [("L", "Largo del muro", "m"), ("H", "Altura total (fondo de zapata a corona)", "m"), ("t1", "Pantalla en la corona", "m"),
                ("t2", "Pantalla en la base", "m"), ("B", "Ancho de la zapata", "m"), ("b1", "Punta (zapata al frente de la pantalla)", "m"),
                ("hz", "Espesor de la zapata", "m"), ("v", "Varilla vertical de la pantalla", "")] + SUELO_ETIQ,
    "contCiclo": [("L", "Largo del muro", "m"), ("H", "Altura total con cimiento", "m"), ("c", "Ancho de la corona", "m"), ("B", "Base del cuerpo", "m"),
                  ("hc", "Alto del cimiento", "m"), ("pz", "Saliente del cimiento al frente", "m"), ("pp", "Piedra bola", "%")] + SUELO_ETIQ,
    "contMamp": [("L", "Largo del muro", "m"), ("H", "Altura total con cimiento", "m"), ("c", "Ancho de la corona", "m"), ("B", "Base del cuerpo", "m"),
                 ("hc", "Alto del cimiento", "m"), ("pz", "Saliente del cimiento al frente", "m"), ("pp", "Piedra", "%"),
                 ("mort", "Mortero (1 : 3 o 1 : 4)", "")] + SUELO_ETIQ,
    "gavion": [("L", "Largo del muro", "m"), ("n", "Número de hiladas", ""), ("hc", "Alto de cada caja", "m"), ("Bb", "Ancho de la hilada de base", "m"),
               ("Bt", "Ancho de la hilada superior", "m"), ("gp", "Peso unitario de la piedra", "t/m³")] + SUELO_ETIQ,
    "pavimento": [("L", "Largo del tramo", "m"), ("A", "Ancho de calzada", "m"), ("e", "Espesor de losa", "m"), ("c", "Carriles", ""),
                  ("sj", "Juntas transversales a cada", "m"), ("eb", "Espesor de base granular", "m"), ("pj", "Pasajuntas", "")],
    "bordillo": [("L", "Largo", "m"), ("b1", "Base inferior", "m"), ("b2", "Base superior", "m"), ("h", "Altura total", "m"),
                 ("sj", "Juntas a cada", "m"), ("ar", "Con armado (2 #3 y eslabones #2)", "")],
    "cunetaL": [("L", "Largo", "m"), ("a", "Ancho de cuneta", "m"), ("t", "Espesor de cuneta", "m"), ("b", "Espesor de bordillo", "m"),
                ("h", "Altura de bordillo", "m"), ("Sx", "Pendiente transversal", "%"), ("T", "Ancho de agua permitido", "m"),
                ("S", "Pendiente longitudinal", "%"), ("n", "n de Manning", ""), ("ar", "Con armado", "")],
    "cunetaV": [("L", "Largo", "m"), ("B", "Ancho superior", "m"), ("hv", "Profundidad", "m"), ("t", "Espesor del revestimiento", "m"),
                ("S", "Pendiente longitudinal", "%"), ("n", "n de Manning", "")],
    "canal": [("L", "Largo", "m"), ("b", "Ancho de fondo", "m"), ("H", "Altura de pared", "m"), ("z", "Talud z (0 = rectangular)", ""),
              ("tw", "Espesor de pared", "m"), ("tf", "Espesor de fondo", "m"), ("y", "Tirante de diseño", "m"), ("S", "Pendiente", "%"),
              ("n", "n de Manning", ""), ("v", "Varilla", ""), ("sep", "Varillas a cada", "m")],
}


def excel_formula(tipo, nombre, params, res, r, obra=""):
    """Memoria de Muros y Carreteras: datos, revisión, mezclas y acero con fórmulas, y dibujos."""
    m = r["m"]
    H = Hoja("Memoria de cálculo", [44, 15, 12, 46, 34, 11, 11, 11, 11, 13, 13, 13, 13, 13, 13])
    xl_encabezado(H, f"Memoria de cálculo: {nombre}" + (f", {obra}" if obra and obra != "Sin obra" else ""),
                  ("" if res is SIN_CONCRETO else f"Concreto {res['nombre']}. ") + f"Fecha: {datetime.now().strftime('%d/%m/%Y')}.")
    xl_dibujos(H, dibujos_formula(tipo, nombre, params), [(0, 500, 3), (3, 530, 2)])
    H.fila([{"v": "DATOS DE ENTRADA", "s": XL["SUB"]}])
    H.fila([{"v": t, "s": XL["ENC"]} for t in ("Dato", "Valor", "Unidad")])
    for k, etiqueta, unidad in ENTRADAS.get(tipo, []):
        if k not in params or params[k] is None:
            continue
        v = params[k]
        if isinstance(v, bool):
            H.fila([etiqueta, {"v": "Sí" if v else "No", "s": XL["DATOTXT"]}, unidad])
        elif isinstance(v, (int, float)):
            H.fila([etiqueta, {"v": v, "s": XL["DATO"]}, unidad])
        else:
            H.fila([etiqueta, {"v": str(v), "s": XL["DATOTXT"]}, unidad])
    if res is not SIN_CONCRETO:
        H.fila(["Resistencia del concreto", {"v": res["nombre"], "s": XL["DATOTXT"]}, ""])
    H.fila()
    H.fila([{"v": "REVISIÓN", "s": XL["SUB"]}])
    for a in (r["avisos"] or []):
        H.fila([{"v": "⚠ " + a, "s": XL["NOTA"], "m": 5}])
    if not r["avisos"]:
        H.fila([{"v": "✓ Cumple las reglas de referencia que revisa esta calculadora.", "s": XL["NOTA"], "m": 5}])
    S = r.get("estab")
    if S:
        H.fila()
        H.fila([{"v": "Estabilidad por metro de muro", "s": XL["SUB"]}])
        H.fila([{"v": t, "s": XL["ENC"]} for t in ("Revisión", "Valor", "Estado")])
        for nom_caso, c in (("", S["est"]), (" con sismo", S["sis"])):
            if not c:
                continue
            q = f"{c['qmax']:.2f}" if math.isfinite(c["qmax"]) else "sin apoyo"
            for txt, val, ok in ((f"FS volteo{nom_caso} (≥ {c['lv']:.1f})", f"{c['FSv']:.2f}", c["okV"]),
                                 (f"FS deslizamiento{nom_caso} (≥ {c['ld']:.1f})", f"{c['FSd']:.2f}", c["okD"]),
                                 (f"Excentricidad{nom_caso} (≤ {c['etxt']} = {c['elim']:.3f} m)", f"{abs(c['e']):.3f} m", c["okE"]),
                                 (f"q máx{nom_caso} t/m² (≤ {c['qadm']:.2f})", q, c["okQ"])):
                H.fila([txt, val, "✓ Cumple" if ok else "⚠ No cumple"])
    h = r.get("hidra")
    if h:
        H.fila()
        H.fila([{"v": "Capacidad hidráulica", "s": XL["SUB"]}])
        H.fila(["Caudal Q", round(h["Q"] * 1000, 2), "L/s"])
        H.fila(["Velocidad", round(h["V"], 3), "m/s"])
        H.fila(["Tirante", round(h["y"], 3), "m"])
        if "Fr" in h:
            H.fila(["Número de Froude", round(h["Fr"], 3), "subcrítico" if h["Fr"] < 1 else "supercrítico"])
    H.fila()
    xl_constantes(H)
    if m["partes"]:
        H.fila([{"v": "MEZCLAS (concreto, mortero y concreto fluido)", "s": XL["SUB"]}])
        H.fila([{"v": t, "s": XL["ENC"]} for t in ("Parte", "Volumen neto m³", "Cemento", "Cal", "Arena", "Piedrín", "Factor seco", "a/c",
                                                   "Vt m³", "Sacos", "Cal m³", "Arena m³", "Piedrín m³", "Agua L")])
        desde = len(H.filas) + 1
        for x in m["partes"]:
            rc, i = x["receta"], len(H.filas) + 1
            suma = f"(C{i}+D{i}+E{i}+F{i})"
            H.fila([x["nombre"], {"v": x["vol"], "s": XL["DATO"]}, {"v": rc["cemento"], "s": XL["DATO"]}, {"v": rc.get("cal", 0), "s": XL["DATO"]},
                    {"v": rc["arena"], "s": XL["DATO"]}, {"v": rc["piedrin"], "s": XL["DATO"]}, {"v": rc.get("factor", VOLUMEN_SECO), "s": XL["DATO"]},
                    {"v": rc["agua_cemento"], "s": XL["DATO"]},
                    {"f": f"B{i}*(1+{{desp}})", "v": x["vol"] * (1 + DESPERDICIO), "s": XL["N3"]},
                    {"f": f"G{i}*C{i}/{suma}*I{i}/{{vsaco}}", "v": x["q"]["sacos"], "s": XL["N2"]},
                    {"f": f"G{i}*D{i}/{suma}*I{i}", "v": x["q"]["cal_m3"], "s": XL["N3"]},
                    {"f": f"G{i}*E{i}/{suma}*I{i}", "v": x["q"]["arena_m3"], "s": XL["N3"]},
                    {"f": f"G{i}*F{i}/{suma}*I{i}", "v": x["q"]["piedrin_m3"], "s": XL["N3"]},
                    {"f": f"J{i}*{{kgsaco}}*H{i}", "v": x["q"]["agua_litros"], "s": XL["N2"]}])
        hasta = len(H.filas)
        H.fila([{"v": "Totales", "s": XL["ENC"]}, {"f": f"SUM(B{desde}:B{hasta})", "s": XL["N3"]}, None, None, None, None, None, None,
                {"f": f"SUM(I{desde}:I{hasta})", "s": XL["N3"]},
                {"f": f"SUM(J{desde}:J{hasta})", "v": m["sacos"], "s": XL["RES"], "id": "tsacos"},
                {"f": f"SUM(K{desde}:K{hasta})", "v": m["cal_m3"], "s": XL["N3"], "id": "tcal"},
                {"f": f"SUM(L{desde}:L{hasta})", "v": m["arena_m3"], "s": XL["RES"], "id": "tarena"},
                {"f": f"SUM(M{desde}:M{hasta})", "v": m["piedrin_m3"], "s": XL["RES"], "id": "tpiedrin"},
                {"f": f"SUM(N{desde}:N{hasta})", "v": m["agua_litros"], "s": XL["RES"], "id": "tagua"}])
        H.fila([{"v": "Fórmulas de cada fila: Vt = V × (1 + desperdicio); sacos = factor × cemento / suma × Vt / volumen del saco; "
                      "cal, arena y piedrín = factor × proporción / suma × Vt; agua = sacos × 42.5 × a/c.", "s": XL["NOTA"], "m": 8}])
        H.fila()
    if m["acero"]:
        H.fila([{"v": "ACERO DE REFUERZO POR DIÁMETRO", "s": XL["SUB"]}])
        H.fila([{"v": t, "s": XL["ENC"]} for t in ("Varilla", "Metros", "Peso kg/m", "Grado", f"Varillas de {LARGO_VARILLA} m", "kg")])
        s0 = len(H.filas) + 1
        for a in m["acero"]:
            i = len(H.filas) + 1
            H.fila([{"v": a["varilla"], "s": XL["TXT"]}, {"v": a["metros"], "s": XL["DATO"]}, {"v": PESO_VARILLA[a["varilla"]], "s": XL["DATO"]},
                    f"Grado {VARILLAS[a['varilla']]['grado']}", {"f": f"ROUNDUP(B{i}*(1+{{desp}})/{{lvar}},0)", "v": a["piezas"], "s": XL["RES0"]},
                    {"f": f"B{i}*C{i}", "v": a["kg"], "s": XL["N2"]}])
        s1 = len(H.filas)
        H.fila([{"v": "Totales", "s": XL["ENC"]}, {"f": f"SUM(B{s0}:B{s1})", "s": XL["N2"]}, None, None,
                {"f": f"SUM(E{s0}:E{s1})", "v": m["varillas"], "s": XL["RES0"], "id": "tvarillas"},
                {"f": f"SUM(F{s0}:F{s1})", "v": m["hierro_kg"], "s": XL["RES"], "id": "thierro"}])
        H.fila([{"v": f"Varillas = metros × (1 + desperdicio) / {LARGO_VARILLA} m, redondeado hacia arriba; kg = metros × peso por metro. "
                      "Alambre de amarre ≈ 3 % del peso del acero.", "s": XL["NOTA"], "m": 6}])
        H.fila()
    H.fila([{"v": "RESUMEN DE MATERIALES", "s": XL["SUB"]}])
    H.fila([{"v": t, "s": XL["ENC"]} for t in ("Material", "Cantidad", "Unidad", "Nota", "Fórmula en Excel")])
    if m["partes"]:
        xl_paso(H, "rsacos", f"Cemento (sacos de {KG_SACO} kg)", "ROUNDUP(ROUND({tsacos},6),0)", m["sacos_comprar"], "sacos", "Total de sacos hacia arriba", XL["RES0"])
        xl_paso(H, "rarena", "Arena", "{tarena}", m["arena_m3"], "m³", "", XL["RES"])
        if m["piedrin_m3"] > 0:
            xl_paso(H, "rpiedrin", "Piedrín", "{tpiedrin}", m["piedrin_m3"], "m³", "", XL["RES"])
        if m["cal_m3"] > 0:
            xl_paso(H, "rcal", "Cal hidratada", "{tcal}*" + str(KG_M3_CAL), m["cal_m3"] * KG_M3_CAL, "kg", f"m³ de cal × {KG_M3_CAL} kg/m³", XL["RES"])
        xl_paso(H, "ragua", "Agua", "{tagua}", m["agua_litros"], "litros", "", XL["RES"])
    if m["acero"]:
        xl_paso(H, "rhierro", "Hierro", "{thierro}", m["hierro_kg"], "kg", "", XL["RES"])
        xl_paso(H, "ralambre", "Alambre de amarre", "{thierro}*0.03", m["hierro_kg"] * 0.03, "kg", "3 % del peso del acero", XL["RES"])
        xl_paso(H, "rvar", f"Varillas de {LARGO_VARILLA} m", "{tvarillas}", m["varillas"], "varillas", "", XL["RES0"])
    if r.get("extras"):
        H.fila()
        H.fila([{"v": "OTROS MATERIALES Y TRABAJOS", "s": XL["SUB"]}])
        H.fila([{"v": t, "s": XL["ENC"]} for t in ("Concepto", "Cantidad", "Unidad")])
        for concepto, cantidad, unidad in r["extras"]:
            try:
                val = float(str(cantidad).replace(",", ""))
            except ValueError:
                val = str(cantidad)
            H.fila([concepto, val, {"v": unidad, "s": XL["TXT"], "m": 2}])
    H.fila()
    H.fila([{"v": "Valores de referencia para un estimado. El diseño final debe revisarlo un ingeniero civil colegiado.", "s": XL["NOTA"], "m": 5}])
    ruta = ruta_memoria(nombre, obra)
    libro_xlsx(H, ruta)
    return ruta


def ofrecer_excel(generar):
    """Pregunta si se descarga la memoria en Excel y la crea en la carpeta «memorias»."""
    if input("  ¿Descargar la memoria de cálculo en Excel (con fórmulas y dibujos)? (s/n): ").strip().lower() != "s":
        return
    try:
        ruta = generar()
        print(f"  Memoria creada:\n  {ruta}")
    except OSError as e:
        print(f"  No se pudo crear el archivo: {e}")


# ============================================================
# PARTE 12: PROGRAMA PRINCIPAL
# ============================================================
def encabezado():
    print("\n==============================================================")
    print(f"  {UNIVERSIDAD.upper()}")
    print(f"  {CARRERA}")
    print("  Integrantes: " + ", ".join(INTEGRANTES))
    print("==============================================================")
    print("  CODEBRAN · PROYECTO FINAL: BODEGA Y CUANTIFICACIÓN DE OBRA")


def main():
    usuario = iniciar_sesion()
    if not usuario:
        print("  ¡Hasta pronto!")
        return
    print(f"  Sesión iniciada como {usuario['nombre']} ({usuario['correo']}).")
    datos = cargar_datos()
    opciones = {
        "1": ("Inventario", menu_inventario),
        "2": ("Calcular materiales (losa, viga, columna, zapata, muro, cimiento, solera, mocheta)", menu_calculo),
        "3": ("Muros: block y contención (reforzado, ciclópeo, gaviones, mampostería)", menu_muros),
        "4": ("Carreteras: pavimento rígido, bordillo, cunetas L y V, canal", menu_carreteras),
        "5": ("Normativa y materiales en Guatemala", ver_normativa),
        "6": ("Ver cálculos guardados", ver_calculos),
        "7": ("Exportar a CSV (Excel)", exportar_csv),
        "8": ("Mi cuenta y contraseña", lambda d: menu_cuenta(d, usuario)),
    }
    while True:
        encabezado()
        print(f"  Sesión: {usuario['nombre']} ({usuario['correo']})")
        print(f"  {len(datos['inventario'])} materiales en bodega, {len(datos['calculos'])} cálculos guardados")
        for k, (texto, _) in opciones.items():
            print(f"  {k}. {texto}")
        print("  0. Salir")
        op = input("  Opción: ").strip()
        if op == "0":
            print(f"  Datos guardados. ¡Hasta pronto, {usuario['nombre']}!")
            break
        if op in opciones:
            opciones[op][1](datos)
        else:
            print("  Opción no válida.")


if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, EOFError):
        print("\n  Programa cerrado.")
