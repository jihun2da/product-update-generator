#!/bin/sh
set -e
mkdir -p .streamlit
if [ -n "$SECRETS_TOML" ]; then
  printf '%s' "$SECRETS_TOML" > .streamlit/secrets.toml
  fi
  exec streamlit run app.py --server.port ${PORT:-8501} --server.address 0.0.0.0
  
