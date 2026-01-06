#!/bin/bash
# MacOS/Linux Launcher

echo "====================================================="
echo "     Transol VMS - Mac/Linux Launcher"
echo "====================================================="
echo ""

# Get directory of this script
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
# Move to Project Root (2 levels up)
cd "$SCRIPT_DIR/../.."

# 1. Check Python
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python 3 is not installed."
    exit 1
fi

# 2. Check/Setup Venv
if [ ! -d "venv" ]; then
    echo "[INFO] First time setup: Creating virtual environment..."
    python3 -m venv venv
    
    echo "[INFO] Installing dependencies..."
    source venv/bin/activate
    pip install --upgrade pip
    if [ -f "requirements.txt" ]; then
        pip install -r requirements.txt
    else
        echo "[WARNING] requirements.txt not found."
    fi
else
    echo "[INFO] Virtual environment found. Activating..."
    source venv/bin/activate
fi

# 3. Database
echo "[INFO] Checking database..."
python manage.py migrate

# 4. Start
echo ""
echo "[INFO] Starting Server..."
echo "[INFO] Opening browser in 3 seconds..."

(sleep 3 && open "http://127.0.0.1:8000/") &
python manage.py runserver
