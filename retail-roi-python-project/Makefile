.PHONY: all venv install test run-cli run-streamlit clean

PYTHON := $(shell command -v python3.10 2>/dev/null || command -v python3.11 2>/dev/null || command -v python3.9 2>/dev/null || command -v python3 2>/dev/null || command -v python 2>/dev/null)
VENV_DIR := .venv
PIP := $(VENV_DIR)/bin/pip
PYTHON_VENV := $(VENV_DIR)/bin/python
STREAMLIT := $(VENV_DIR)/bin/streamlit
PYTEST := $(VENV_DIR)/bin/pytest

all: install test

venv:
	@if [ ! -d "$(VENV_DIR)" ]; then \
		echo "Creando entorno virtual $(VENV_DIR)"; \
		$(PYTHON) -m venv $(VENV_DIR); \
	else \
		echo "Entorno virtual ya existe"; \
	fi

install: venv
	$(PIP) install --upgrade pip
	$(PIP) install -r requirements.txt
	$(PIP) install -e .

test: install
	$(PYTEST) -q

# Uso: make run-cli ARGS="./data/TU_ARCHIVO.xlsm --out roi_output.json"
run-cli:
	$(PYTHON_VENV) -m retail_roi_model.cli $(ARGS)

# Uso: make run-streamlit
run-streamlit:
	$(STREAMLIT) run app.py

clean:
	rm -rf $(VENV_DIR) __pycache__ *.pyc
