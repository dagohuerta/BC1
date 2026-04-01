# Retail ROI Python Project

Proyecto listo para abrirse en Visual Studio Code, PyCharm o cualquier IDE compatible con Python.

## Qué incluye
- Motor Python para reproducir el cálculo del workbook ROI.
- CLI para ejecutar el modelo desde terminal.
- App simple en Streamlit para correr el modelo vía navegador.
- `requirements.txt` y `pyproject.toml` para instalar dependencias.
- Estructura `src/` lista para empaquetado.

## Estructura

```text
retail-roi-python-project/
├─ app.py
├─ pyproject.toml
├─ requirements.txt
├─ README.md
├─ .gitignore
├─ src/
│  └─ retail_roi_model/
│     ├─ __init__.py
│     ├─ cli.py
│     └─ engine.py
├─ tests/
│  └─ test_finance_helpers.py
├─ data/
│  └─ .gitkeep
└─ examples/
   └─ .gitkeep
```

## Requisitos
- Python 3.10 o superior

## Instalación

Nota: si tienes múltiples versiones de Python, usa `python3.10` (o `python3.11`) en lugar de `python` para garantizar compatibilidad con el proyecto.

### Opción 1: con requirements
```bash
python3.10 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

En Windows PowerShell:
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### Opción 2: editable package
```bash
pip install -e .
```

## Cómo correr el modelo por terminal
```bash
python -m retail_roi_model.cli "./data/TU_ARCHIVO.xlsm" --out roi_output.json
```

O si instalaste el script:
```bash
retail-roi "./data/TU_ARCHIVO.xlsm" --out roi_output.json
```

## Cómo levantar la app web
```bash
streamlit run app.py
```

## Inputs que considera el motor
El motor está preparado para leer, como mínimo, las hojas siguientes del workbook:
- Financial Input
- Module Selection
- Admin - Sheet_Row
- Benefit Input
- Value Benefit
- Adoption Input
- Timeline Input
- Investment Input
- Forward Financials
- ROI Output - Modular
- ROI Output - Total
- P&L and Cash Flow
- Proforma Financials

### Bloques clave de input
1. **Financial Input**
   - Estado de resultados base
   - Balance principal
   - Annual inventory turns
   - Tax rate
   - Inventory carrying cost
   - Growth rates
   - Discount rate / interest rate
   - Ecommerce mix
   - Quarterly revenue and mix
   - Business breakdown

2. **Module Selection**
   - Cliente
   - Tipo de retailer
   - Módulos seleccionados
   - Tipo de assessment

3. **Benefit Input**
   - % increased sales
   - % inventory reduction
   - % margin improvement
   - % labor reduction
   - % logistics reduction

4. **Value Benefit**
   - Benchmarks complementarios por módulo

5. **Adoption Input / Timeline Input**
   - Curvas de adopción
   - Secuencia de go-live
   - Timing por quarter

6. **Investment Input**
   - Software fees
   - Software maintenance
   - Oracle services
   - Client services
   - 3rd party integrator
   - Hardware capex / maintenance
   - Hosting

## Notas
- No incluye el workbook dentro del zip para que puedas poner el archivo del cliente que corresponda.
- Coloca tu `.xlsm` dentro de `data/` o indica la ruta completa al correrlo.
