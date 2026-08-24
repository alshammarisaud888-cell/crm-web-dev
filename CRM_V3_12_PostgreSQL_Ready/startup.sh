#!/usr/bin/env bash
set -e
export APP_ENV=azure
export APP_HOST=0.0.0.0
export APP_PORT="${SERVER_PORT:-${PORT:-8000}}"
python main.py
