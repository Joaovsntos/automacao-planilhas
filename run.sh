#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
python src/preparar_entradas.py
python src/consolidar.py
