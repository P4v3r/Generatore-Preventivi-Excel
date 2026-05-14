#!/bin/bash
# Genera preventivi con setup automatico

set -e

echo "📦 Installazione dipendenze..."
pip install -r requirements.txt

echo "🚀 Esecuzione generatore..."
python genera_preventivi.py "$@"

echo "✅ Fatto!"