#!/usr/bin/env bash
set -o errexit  

apt-get update && apt-get install -y wkhtmltopdf
pip install -r requirements.txt
