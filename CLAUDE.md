# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

University assignment (Percepción Computacional, R2-A2-S8): diagnose capture-quality problems in five photos of the same object, treat the problematic one with OpenCV, and produce a Word report in APA 7th edition (Spanish). Requirements come from `Guia_Semana_8_Tratamiento_de_Imagen.pdf` (read it with `pdftotext -enc UTF-8 -layout`). The guide asks for the student's **own** photos and for a PDF deliverable. Don't fabricate metrics or swap in internet images.

## Commands

```bash
pip install opencv-python numpy matplotlib python-docx
python tratamiento_imagen.py      # analysis only -> resultados/ (figures, metricas.json, oscura_tratada.png)
python generar_informe.py         # runs the analysis, then writes Informe_tratamiento_imagen.docx
python generar_informe.py --imagenes <dir> --resultados <dir> --salida <file.docx>
```

- On Windows consoles, set `PYTHONIOENCODING=utf-8` when you print results (variant names contain α/β).
- There are no tests or linters. To check a change, regenerate the report, export it to PDF through Word COM (LibreOffice/soffice isn't installed), and render pages with PyMuPDF (`import pymupdf`). `pdftoppm` isn't available.
- If the .docx is open in Word, saving raises `PermissionError`. Write to another name with `--salida`.

## Architecture

- **`tratamiento_imagen.py`** does all the computation. It expects `imagenes/{referencia,oscura,desenfocada,fondo,angulo}.{jpg,jpeg,png}` (the names and order come from the `CONDICIONES` dict). Default paths resolve relative to the script, not the current directory. Before measuring, every image is resized to `ANCHO_ANALISIS` (1024 px) so the scale-dependent metrics (Laplacian variance, Canny density) are comparable. `ejecutar()` returns, and also writes to `metricas.json`, a dict with keys `metricas`, `formatos`, `tratamiento` and `parametros`.
- **The treatment pipeline** targets the dark image only: adaptive gamma on L (CIELAB), then CLAHE on L, then a bilateral filter. The gamma stage exists because CLAHE alone barely brightens a severely underexposed image, and the report text cites this. `ejecutar()` also computes the alternatives (linear α/β, global equalization, CLAHE only, Gamma + CLAHE) so the report can justify the choice. Hyperparameters are module-level constants, and the pipeline diagram labels are built from them.
- **`generar_informe.py`** imports `tratamiento_imagen as ti` and builds the docx with python-docx:
  - **Placeholder mode:** if any image is missing it still produces a full report, filling values with `[valor]` and marking figures as pending. The `Datos` wrapper returns `None`, which formatters such as `v()`, `var_pct()` and `ratio_pct()` turn into placeholders.
  - **Data-dependent text:** some narrative is chosen by the results, such as the improvement verdict, the per-metric interpretations in Tabla 6, and the "more edges than reference" caveat (triggered when edge density is above 1.3× the reference). Keep any new prose consistent with the actual numbers.
  - **Keyed lookups:** the report fetches treatment variants by exact name, and `CLAVE_PROPUESTA` must match the key used in `ti.ejecutar()`. Renaming a variant in one file requires renaming it in the other.
  - **Embedded code:** code snippets are pulled from `tratamiento_imagen.py` by function name via `ast` (`fuente_funcion`). Appendix A embeds the whole file, so the comments there are part of the deliverable.
  - **Cover-page fields:** student, course, object and device are constants at the top of the file.
- **APA 7 formatting** lives in helpers: `configurar_estilos` sets Times New Roman 12, double spacing and heading levels 1–3; `tabla_apa` draws horizontal rules only and keeps rows together; `figura` and `etiqueta` produce a bold number with an italic title plus a *Nota*; `referencia` applies a hanging indent. `add_runs` supports inline `**bold**`, `*italic*`, `` `code` ``, `^{sup}` and `_{sub}` markup.
- Figure colors use the blue `#2a78d6` for before and the orange `#eb6834` for after.
