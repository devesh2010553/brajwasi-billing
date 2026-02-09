#!/bin/bash
export FLASK_ENV=production
gunicorn app:app --bind 0.0.0.0:$PORT
