@echo off
echo Setting up virtual environment...

if not exist venv (
    python -m venv venv
)

venv\Scripts\activate

echo Installing dependencies...
pip install -r requirements.txt

echo Starting Django development server...
python manage.py runserver
