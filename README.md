# Informe de resultados de tratamiento de imagen

Taller R2-A2-S8 · Percepción Computacional 2026

Diagnóstico → decisión técnica → implementación → evaluación, con OpenCV en Python.

## Estructura
- `imagenes/` → coloque aquí sus cinco fotos: `referencia.jpg`, `oscura.jpg`, `desenfocada.jpg`, `fondo.jpg`, `angulo.jpg`
- `tratamiento_imagen.py` → métricas, comparación JPG/PNG, tratamiento (gamma adaptativa + CLAHE en L + bilateral) y figuras
- `generar_informe.py` → ejecuta el análisis y genera `Informe_tratamiento_imagen.docx` (APA 7)
- `resultados/` → figuras, `metricas.json`, `oscura_tratada.png` (se crea al ejecutar)

## Uso
```bash
pip install opencv-python numpy matplotlib python-docx
python tratamiento_imagen.py          # solo análisis y figuras
python generar_informe.py             # análisis + informe .docx
```
En Google Colab: suba la carpeta `imagenes/` y los dos scripts, y ejecute `!python generar_informe.py`.
