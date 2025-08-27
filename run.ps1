$ErrorActionPreference = 'Stop'
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python src\fetch_outlook.py $args
