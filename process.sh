#!/bin/bash
if [ -z "$1" ]; then echo "Error: you must specify the XLS file"; exit 1; fi
./venv/bin/python process.py "$1"
