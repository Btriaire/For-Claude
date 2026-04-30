#!/bin/bash
source /opt/render/project/src/.venv/bin/activate
exec gunicorn app:app --bind 0.0.0.0:$PORT
