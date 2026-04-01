#!/bin/bash
# Inicia el proyecto asegurando que no haya procesos colgando en 8501 antes.
set -e
cd "$(dirname "$0")"
# Mata procesos existentes en 8501
existing=$(lsof -ti tcp:8501 || true)
if [ -n "$existing" ]; then
  echo "Matando procesos existentes en puerto 8501: $existing"
  echo "$existing" | xargs -r kill -9
fi
# Activa virtualenv y arranca Streamlit
source .venv/bin/activate
streamlit run app.py --server.port 8501 --server.headless true
