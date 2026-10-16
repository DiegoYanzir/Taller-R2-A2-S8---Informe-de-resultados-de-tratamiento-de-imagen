# -*- coding: utf-8 -*-
"""
tratamiento_imagen.py
---------------------------------------------------------------------------
R2-A2-S8 - Informe de resultados de tratamiento de imagen
Percepcion Computacional - Especializacion en Inteligencia Artificial

Ciclo de trabajo: Diagnostico -> Decision tecnica -> Implementacion -> Evaluacion

1. Carga las cinco capturas del mismo objeto (referencia, oscura, desenfocada,
   fondo, angulo) y calcula metricas de calidad comparables.
2. Compara el almacenamiento JPG frente a PNG.
3. Trata la imagen problematica (baja iluminacion) con correccion gamma
   adaptativa y CLAHE sobre el canal L de CIELAB, seguido de un filtro bilateral, y la compara con alternativas
   mas simples (ajuste lineal y ecualizacion global).
4. Vuelve a medir, genera las figuras comparativas y guarda metricas.json.

Uso (local, Colab o Jupyter):
    python tratamiento_imagen.py --imagenes imagenes --salida resultados
En Colab: suba la carpeta 'imagenes' y ejecute  !python tratamiento_imagen.py
---------------------------------------------------------------------------
"""

import argparse                      # Lectura de argumentos de linea de comandos
import json                          # Exportacion de metricas a JSON
import os                            # Manejo de rutas y tamanos de archivo

import cv2                           # OpenCV: lectura, color, filtros, bordes
import numpy as np                   # Operaciones numericas sobre matrices
import matplotlib                    # Motor de graficas
matplotlib.use("Agg")                # Backend sin ventana (sirve en Colab y en servidor)
import matplotlib.pyplot as plt      # Interfaz de graficas

# ---------------------------------------------------------------------------
# Configuracion general
# ---------------------------------------------------------------------------
CONDICIONES = {                      # Nombre de archivo -> condicion de captura
    "referencia": "Referencia",
    "oscura": "Baja iluminación",
    "desenfocada": "Desenfoque / movimiento",
    "fondo": "Fondo complejo",
    "angulo": "Otra perspectiva",
}
EXTENSIONES = (".jpg", ".jpeg", ".png")  # Formatos aceptados al buscar cada archivo
ANCHO_ANALISIS = 1024                # Ancho comun para que las metricas sean comparables
UMBRAL_OSCURO = 50                   # Intensidad bajo la cual un pixel se considera subexpuesto
UMBRAL_SATURADO = 205                # Intensidad sobre la cual un pixel se considera sobreexpuesto
CANNY_BAJO, CANNY_ALTO = 50, 150     # Umbrales fijos de Canny (simulan el detector posterior)

# Hiperparametros del tratamiento propuesto
GAMMA_MIN = 0.35                     # Gamma minima permitida (evita lavar la imagen)
CLAHE_CLIP = 2.0                     # clipLimit: tope de amplificacion del contraste local
CLAHE_TILES = (8, 8)                 # tileGridSize: rejilla de regiones contextuales
BILAT_D = 5                          # Diametro del vecindario del filtro bilateral (px)
BILAT_SIGMA_COLOR = 25               # sigmaColor: tolerancia en diferencia de intensidad
BILAT_SIGMA_SPACE = 5                # sigmaSpace: alcance espacial de la gaussiana

# Colores de las graficas (antes / despues)
AZUL, NARANJA, GRIS = "#2a78d6", "#eb6834", "#52514e"


# ---------------------------------------------------------------------------
# 1. Carga y normalizacion
# ---------------------------------------------------------------------------
def buscar_archivo(carpeta, nombre):
    """Devuelve la ruta de 'nombre' con cualquier extension valida, o None."""
    for ext in EXTENSIONES:                                   # Recorre las extensiones admitidas
        ruta = os.path.join(carpeta, nombre + ext)            # Construye la ruta candidata
        if os.path.isfile(ruta):                              # Si el archivo existe...
            return ruta                                       # ...se devuelve
    return None                                               # Ninguna coincidencia


def cargar_imagen(ruta):
    """Lee una imagen en BGR (convencion de OpenCV) y valida la lectura."""
    datos = np.fromfile(ruta, dtype=np.uint8)                 # Lee bytes (soporta rutas con tildes)
    imagen = cv2.imdecode(datos, cv2.IMREAD_COLOR)            # Decodifica a matriz BGR de 8 bits
    if imagen is None:                                        # Falla de decodificacion
        raise IOError(f"No se pudo leer la imagen: {ruta}")
    return imagen


def redimensionar(imagen, ancho=ANCHO_ANALISIS):
    """Escala a un ancho comun conservando la relacion de aspecto."""
    alto_orig, ancho_orig = imagen.shape[:2]                  # Dimensiones originales
    if ancho_orig == ancho:                                   # Ya tiene el tamano objetivo
        return imagen.copy()
    escala = ancho / ancho_orig                               # Factor de escala
    interp = cv2.INTER_AREA if escala < 1 else cv2.INTER_CUBIC  # AREA reduce sin aliasing
    return cv2.resize(imagen, (ancho, int(round(alto_orig * escala))), interpolation=interp)


# ---------------------------------------------------------------------------
# 2. Metricas de calidad
# ---------------------------------------------------------------------------
def estimar_ruido(gris):
    """Desviacion estandar del ruido por el metodo rapido de Immerkaer (1996)."""
    h, w = gris.shape                                         # Alto y ancho
    mascara = np.array([[1, -2, 1], [-2, 4, -2], [1, -2, 1]], np.float64)  # Anula estructura lineal
    resp = cv2.filter2D(gris.astype(np.float64), -1, mascara)  # Respuesta del operador
    suma = np.abs(resp[1:-1, 1:-1]).sum()                     # Excluye el borde de la imagen
    return float(np.sqrt(np.pi / 2) * suma / (6 * (w - 2) * (h - 2)))


def entropia(gris):
    """Entropia de Shannon del histograma (bits): riqueza de niveles de gris."""
    hist = cv2.calcHist([gris], [0], None, [256], [0, 256]).ravel()  # Histograma de 256 bins
    p = hist / hist.sum()                                     # Probabilidades
    p = p[p > 0]                                              # Evita log(0)
    return float(-(p * np.log2(p)).sum())


def calcular_metricas(imagen_bgr):
    """Calcula las metricas usadas en el informe sobre una imagen BGR."""
    gris = cv2.cvtColor(imagen_bgr, cv2.COLOR_BGR2GRAY)       # Luminancia: 0.299R+0.587G+0.114B
    bordes = cv2.Canny(gris, CANNY_BAJO, CANNY_ALTO)          # Mapa binario de bordes
    return {
        "brillo": float(gris.mean()),                         # Media de intensidad (0-255)
        "contraste": float(gris.std()),                       # Desviacion estandar (contraste RMS)
        "nitidez": float(cv2.Laplacian(gris, cv2.CV_64F).var()),  # Varianza del Laplaciano
        "entropia": entropia(gris),                           # Bits por pixel
        "ruido": estimar_ruido(gris),                         # Sigma de ruido estimada
        "pct_oscuros": float((gris < UMBRAL_OSCURO).mean() * 100),     # % subexpuesto
        "pct_saturados": float((gris > UMBRAL_SATURADO).mean() * 100),  # % sobreexpuesto
        "densidad_bordes": float((bordes > 0).mean() * 100),  # % de pixeles de borde
    }


def diagnosticar(nombre, m, ref):
    """Asigna problema y adecuacion comparando contra la imagen de referencia."""
    problemas = []                                            # Lista de defectos detectados
    if m["brillo"] < 0.6 * ref["brillo"] or m["pct_oscuros"] > 40:
        problemas.append("subexposición")                     # Exposicion insuficiente
    if m["contraste"] < 0.6 * ref["contraste"]:
        problemas.append("bajo contraste")                    # Rango dinamico comprimido
    if m["nitidez"] < 0.35 * ref["nitidez"]:
        problemas.append("pérdida de nitidez")                # Desenfoque significativo
    if m["ruido"] > 1.5 * ref["ruido"] and nombre != "fondo":
        problemas.append("ruido elevado")                     # Ruido de sensor (ISO alto)
    if nombre == "fondo":
        problemas.append("distractores en el fondo")          # No lo capta una metrica global
    if nombre == "angulo":
        problemas.append("cambio geométrico de vista")        # Variacion de pose / escala
    if nombre == "referencia":
        return "Ninguno (línea base)", "Sí"
    if not problemas:
        return "Sin defecto métrico relevante", "Sí"
    criticos = {"subexposición", "bajo contraste", "pérdida de nitidez", "ruido elevado"}
    adecuada = "No" if criticos.intersection(problemas) else "Condicional"
    return ", ".join(problemas).capitalize(), adecuada


# ---------------------------------------------------------------------------
# 3. Tratamiento propuesto y alternativas
# ---------------------------------------------------------------------------
def ajuste_lineal(imagen, alpha=1.2, beta=25):
    """Alternativa 1: g(x) = alpha*f(x) + beta (transformacion global)."""
    return cv2.convertScaleAbs(imagen, alpha=alpha, beta=beta)  # Satura a [0, 255]


def ecualizacion_global(imagen):
    """Alternativa 2: ecualizacion de histograma global sobre L (CIELAB)."""
    lab = cv2.cvtColor(imagen, cv2.COLOR_BGR2LAB)             # Separa luminancia de croma
    l, a, b = cv2.split(lab)                                  # Canales L, a, b
    l = cv2.equalizeHist(l)                                   # CDF global como transferencia
    return cv2.cvtColor(cv2.merge((l, a, b)), cv2.COLOR_LAB2BGR)


def gamma_adaptativa(l):
    """Elige gamma para llevar la luminancia media a 0.5: gamma = ln(0.5)/ln(media)."""
    media = max(float(l.mean()) / 255.0, 1e-3)                # Media normalizada en (0, 1]
    gamma = np.log(0.5) / np.log(media)                       # Exponente que mapea media -> 0.5
    return float(np.clip(gamma, GAMMA_MIN, 1.0))              # Solo aclara; limita la expansion


def corregir_gamma(l, gamma):
    """Transformacion de potencia s = 255*(r/255)^gamma mediante tabla de busqueda."""
    lut = (np.power(np.arange(256) / 255.0, gamma) * 255).round().astype(np.uint8)  # 256 entradas
    return cv2.LUT(l, lut)                                    # Aplica la tabla pixel a pixel


def aplicar_clahe(imagen, clip=CLAHE_CLIP, tiles=CLAHE_TILES, gamma=False):
    """CLAHE sobre el canal L (opcionalmente precedido de gamma) sin alterar el tono."""
    lab = cv2.cvtColor(imagen, cv2.COLOR_BGR2LAB)             # BGR -> CIELAB
    l, a, b = cv2.split(lab)                                  # Solo se modifica L
    if gamma:                                                 # Etapa global: levanta las sombras
        l = corregir_gamma(l, gamma_adaptativa(l))
    clahe = cv2.createCLAHE(clipLimit=clip, tileGridSize=tiles)  # Configura CLAHE
    l_eq = clahe.apply(l)                                     # Ecualizacion local limitada
    return cv2.cvtColor(cv2.merge((l_eq, a, b)), cv2.COLOR_LAB2BGR)  # Regresa a BGR


def tratamiento_propuesto(imagen):
    """Pipeline: gamma adaptativa(L) -> CLAHE(L) -> filtro bilateral."""
    realzada = aplicar_clahe(imagen, gamma=True)              # Exposicion global + contraste local
    return cv2.bilateralFilter(realzada, BILAT_D,             # Promedia solo pixeles similares
                               BILAT_SIGMA_COLOR, BILAT_SIGMA_SPACE)


# ---------------------------------------------------------------------------
# 4. Figuras
# ---------------------------------------------------------------------------
def rgb(imagen):
    """Matplotlib espera RGB; OpenCV entrega BGR."""
    return cv2.cvtColor(imagen, cv2.COLOR_BGR2RGB)


def estilo_ejes(ax):
    """Ejes discretos: sin marco superior/derecho y texto gris."""
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    ax.tick_params(colors=GRIS, labelsize=8)
    ax.grid(axis="y", color="#e6e5e0", linewidth=0.6)
    ax.set_axisbelow(True)


def figura_conjunto(imagenes, metricas, ruta):
    """Figura 1: las cinco capturas con sus metricas principales."""
    fig, ejes = plt.subplots(2, 3, figsize=(12, 7))           # Rejilla 2x3 (una celda libre)
    ejes.ravel()[-1].axis("off")
    for ax, (nombre, img) in zip(ejes.ravel(), imagenes.items()):
        m = metricas[nombre]
        ax.imshow(rgb(img))
        ax.set_title(CONDICIONES[nombre], fontsize=10)
        ax.set_xlabel(f"Brillo {m['brillo']:.1f} | Contraste {m['contraste']:.1f}\n"
                      f"Nitidez {m['nitidez']:.1f}", fontsize=8, color=GRIS)
        ax.set_xticks([]); ax.set_yticks([])
    fig.tight_layout()
    fig.savefig(ruta, dpi=200)
    plt.close(fig)


def figura_histogramas(imagenes, ruta):
    """Figura 2: histogramas de intensidad (multiplos pequenos, misma escala)."""
    fig, ejes = plt.subplots(2, 3, figsize=(12, 5.6), sharey=True)
    ejes.ravel()[-1].axis("off")
    for ax, (nombre, img) in zip(ejes.ravel(), imagenes.items()):
        gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        hist = cv2.calcHist([gris], [0], None, [256], [0, 256]).ravel()
        ax.fill_between(np.arange(256), hist / hist.sum() * 100, color=AZUL, alpha=0.35, linewidth=0)
        ax.plot(hist / hist.sum() * 100, color=AZUL, linewidth=1.2)
        ax.set_title(CONDICIONES[nombre], fontsize=10)
        ax.set_xlim(0, 255)
        ax.set_xlabel("Intensidad", fontsize=8, color=GRIS)
        estilo_ejes(ax)
    for ax in ejes[:, 0]:
        ax.set_ylabel("% de píxeles", fontsize=8, color=GRIS)
    fig.tight_layout()
    fig.savefig(ruta, dpi=200)
    plt.close(fig)


def figura_antes_despues(variantes, ruta):
    """Figura 3: original, alternativas y tratamiento propuesto."""
    filas = int(np.ceil(len(variantes) / 3))                  # Rejilla de 3 columnas
    fig, ejes = plt.subplots(filas, 3, figsize=(13, 3.6 * filas))
    for ax in ejes.ravel()[len(variantes):]:                  # Oculta celdas sobrantes
        ax.axis("off")
    for ax, (titulo, img) in zip(ejes.ravel(), variantes.items()):
        ax.imshow(rgb(img))
        ax.set_title(titulo, fontsize=10)
        ax.axis("off")
    fig.tight_layout()
    fig.savefig(ruta, dpi=200)
    plt.close(fig)


def figura_histograma_tratamiento(original, tratada, ruta):
    """Figura 4: histograma y funcion de distribucion acumulada antes/despues."""
    fig, (h_ax, c_ax) = plt.subplots(1, 2, figsize=(11, 3.4))
    for img, etiqueta, color in ((original, "Antes", AZUL), (tratada, "Después", NARANJA)):
        gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        hist = cv2.calcHist([gris], [0], None, [256], [0, 256]).ravel()
        hist = hist / hist.sum()
        h_ax.plot(hist * 100, color=color, linewidth=1.6, label=etiqueta)
        c_ax.plot(np.cumsum(hist) * 100, color=color, linewidth=2, label=etiqueta)
    h_ax.set_title("Histograma de intensidad", fontsize=10)
    h_ax.set_ylabel("% de píxeles", fontsize=8, color=GRIS)
    c_ax.set_title("Distribución acumulada (CDF)", fontsize=10)
    c_ax.set_ylabel("% acumulado", fontsize=8, color=GRIS)
    for ax in (h_ax, c_ax):
        ax.set_xlim(0, 255)
        ax.set_xlabel("Intensidad", fontsize=8, color=GRIS)
        ax.legend(frameon=False, fontsize=8)
        estilo_ejes(ax)
    fig.tight_layout()
    fig.savefig(ruta, dpi=200)
    plt.close(fig)


def figura_bordes(original, tratada, ruta):
    """Figura 5: respuesta de Canny (umbrales fijos) antes y despues."""
    fig, ejes = plt.subplots(1, 2, figsize=(10, 4.2))
    for ax, img, titulo in zip(ejes, (original, tratada), ("Bordes - original", "Bordes - tratada")):
        gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        ax.imshow(cv2.Canny(gris, CANNY_BAJO, CANNY_ALTO), cmap="gray")
        ax.set_title(titulo, fontsize=10)
        ax.axis("off")
    fig.tight_layout()
    fig.savefig(ruta, dpi=200)
    plt.close(fig)


def figura_pipeline(ruta):
    """Figura del flujo de procesamiento (diagrama conceptual)."""
    pasos = ["Captura\n(oscura.jpg)", "Diagnóstico\nbrillo, contraste,\nnitidez, ruido",
             "BGR → CIELAB\n(separar L)", f"Gamma adaptativa\nen L (γ ≥ {GAMMA_MIN})",
             f"CLAHE en L\nclip={CLAHE_CLIP}, {CLAHE_TILES[0]}×{CLAHE_TILES[1]}", "LAB → BGR",
             f"Filtro bilateral\nd={BILAT_D}, σc={BILAT_SIGMA_COLOR}, σs={BILAT_SIGMA_SPACE}",
             "Re-medición\ny comparación"]
    fig, ax = plt.subplots(figsize=(11, 4.2))
    ax.set_xlim(0, 8); ax.set_ylim(0, 4.2); ax.axis("off")
    pos = [(i * 2 + 0.2, 2.5) for i in range(4)] + [((3 - i) * 2 + 0.2, 0.3) for i in range(4)]  # Recorrido en S
    for i, (texto, (x, y)) in enumerate(zip(pasos, pos)):
        resalte = i in (3, 4, 6)                                  # Etapas de transformacion
        ax.add_patch(plt.Rectangle((x, y), 1.6, 1.4, facecolor="#fde9e0" if resalte else "#eaf1fb",
                                   edgecolor=NARANJA if resalte else AZUL, linewidth=1.4))
        ax.text(x + 0.8, y + 0.7, texto, ha="center", va="center", fontsize=10)
    for i in range(len(pasos) - 1):                               # Flechas entre etapas
        (x0, y0), (x1, y1) = pos[i], pos[i + 1]
        if y0 != y1:                                              # Bajada a la segunda fila
            inicio, fin = (x0 + 0.8, y0), (x1 + 0.8, y1 + 1.4)
        elif x1 > x0:                                             # Avance hacia la derecha
            inicio, fin = (x0 + 1.6, y0 + 0.7), (x1, y1 + 0.7)
        else:                                                     # Avance hacia la izquierda
            inicio, fin = (x0, y0 + 0.7), (x1 + 1.6, y1 + 0.7)
        ax.annotate("", xy=fin, xytext=inicio, arrowprops=dict(arrowstyle="->", color=GRIS, lw=1.4))
    fig.tight_layout()
    fig.savefig(ruta, dpi=200)
    plt.close(fig)


# ---------------------------------------------------------------------------
# 5. Programa principal
# ---------------------------------------------------------------------------
def comparar_formatos(imagen, carpeta):
    """Guarda la imagen en JPG (calidad 90) y PNG y compara tamano y fidelidad."""
    ruta_jpg = os.path.join(carpeta, "comparacion.jpg")
    ruta_png = os.path.join(carpeta, "comparacion.png")
    cv2.imwrite(ruta_jpg, imagen, [cv2.IMWRITE_JPEG_QUALITY, 90])  # Compresion con perdida (DCT)
    cv2.imwrite(ruta_png, imagen)                             # Compresion sin perdida (DEFLATE)
    jpg = cv2.imread(ruta_jpg)                                # Relee para medir la perdida
    alto, ancho = imagen.shape[:2]
    return {
        "resolucion": f"{ancho} × {alto}",
        "jpg_kb": os.path.getsize(ruta_jpg) / 1024,
        "png_kb": os.path.getsize(ruta_png) / 1024,
        "psnr_jpg": float(cv2.PSNR(imagen, jpg)),             # dB frente al original
    }


def ejecutar(carpeta_imagenes, carpeta_salida):
    """Ejecuta todo el ciclo y devuelve un diccionario con los resultados."""
    os.makedirs(carpeta_salida, exist_ok=True)
    faltan = [n for n in CONDICIONES if buscar_archivo(carpeta_imagenes, n) is None]
    if faltan:                                                # Mensaje claro antes de procesar
        raise FileNotFoundError(
            f"Faltan {len(faltan)} imagen(es) en {os.path.abspath(carpeta_imagenes)}: "
            + ", ".join(f + ".jpg" for f in faltan)
            + ". Se aceptan .jpg, .jpeg o .png (convierta antes las fotos .HEIC del iPhone).")
    originales, imagenes = {}, {}
    for nombre in CONDICIONES:                                # Carga en orden fijo
        ruta = buscar_archivo(carpeta_imagenes, nombre)
        originales[nombre] = cargar_imagen(ruta)
        imagenes[nombre] = redimensionar(originales[nombre])  # Ancho comun para comparar

    # --- Diagnostico ---
    metricas = {n: calcular_metricas(img) for n, img in imagenes.items()}
    for n in metricas:
        h, w, c = originales[n].shape
        metricas[n].update({"ancho": w, "alto": h, "canales": c, "megapixeles": w * h / 1e6})
        metricas[n]["problema"], metricas[n]["adecuada"] = diagnosticar(n, metricas[n], metricas["referencia"])

    formatos = comparar_formatos(imagenes["referencia"], carpeta_salida)

    # --- Tratamiento de la imagen problematica (baja iluminacion) ---
    original = imagenes["oscura"]
    variantes = {
        "Original (oscura)": original,
        "Ajuste lineal α=1.2, β=25": ajuste_lineal(original),
        "Ecualización global": ecualizacion_global(original),
        "Solo CLAHE": aplicar_clahe(original),
        "Gamma + CLAHE": aplicar_clahe(original, gamma=True),
        "Gamma + CLAHE + bilateral (propuesta)": tratamiento_propuesto(original),
    }
    tratadas = {k: calcular_metricas(v) for k, v in variantes.items()}
    tratada = variantes["Gamma + CLAHE + bilateral (propuesta)"]
    cv2.imwrite(os.path.join(carpeta_salida, "oscura_tratada.png"), tratada)

    # --- Figuras ---
    figura_conjunto(imagenes, metricas, os.path.join(carpeta_salida, "fig1_conjunto.png"))
    figura_histogramas(imagenes, os.path.join(carpeta_salida, "fig2_histogramas.png"))
    figura_pipeline(os.path.join(carpeta_salida, "fig3_pipeline.png"))
    figura_antes_despues(variantes, os.path.join(carpeta_salida, "fig4_antes_despues.png"))
    figura_histograma_tratamiento(original, tratada, os.path.join(carpeta_salida, "fig5_histograma.png"))
    figura_bordes(original, tratada, os.path.join(carpeta_salida, "fig6_bordes.png"))

    resultados = {"metricas": metricas, "formatos": formatos, "tratamiento": tratadas,
                  "parametros": {"gamma": gamma_adaptativa(cv2.cvtColor(original, cv2.COLOR_BGR2LAB)[:, :, 0]),
                                 "gamma_min": GAMMA_MIN, "clipLimit": CLAHE_CLIP, "tileGridSize": list(CLAHE_TILES),
                                 "d": BILAT_D, "sigmaColor": BILAT_SIGMA_COLOR,
                                 "sigmaSpace": BILAT_SIGMA_SPACE, "ancho_analisis": ANCHO_ANALISIS}}
    with open(os.path.join(carpeta_salida, "metricas.json"), "w", encoding="utf-8") as f:
        json.dump(resultados, f, ensure_ascii=False, indent=2)
    return resultados


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Diagnóstico y tratamiento de imágenes con OpenCV")
    base = os.path.dirname(os.path.abspath(__file__))         # Rutas relativas al script, no al IDE
    parser.add_argument("--imagenes", default=os.path.join(base, "imagenes"), help="Carpeta con las cinco capturas")
    parser.add_argument("--salida", default=os.path.join(base, "resultados"), help="Carpeta de figuras y métricas")
    args = parser.parse_args()
    if hasattr(os.sys.stdout, "reconfigure"):
        os.sys.stdout.reconfigure(encoding="utf-8")              # Tildes en consolas de Windows
    r = ejecutar(args.imagenes, args.salida)
    print(f"{'Imagen':<12}{'Brillo':>8}{'Contr.':>8}{'Nitidez':>10}{'Ruido':>7}  Adecuada")
    for n, m in r["metricas"].items():
        print(f"{n:<12}{m['brillo']:>8.1f}{m['contraste']:>8.1f}{m['nitidez']:>10.1f}{m['ruido']:>7.2f}  {m['adecuada']}")
