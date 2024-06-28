@echo off

rem Activate virtual environment (if necessary)
call venv\Scripts\activate

rem Start Django development server in a new terminal window
start cmd /k python manage.py runserver

rem Delay to ensure server starts before running process_tasks
timeout /t 5

rem Run process_tasks in another new terminal window
start cmd /k python manage.py process_tasks
