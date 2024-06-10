@echo off

:: Check if venv folder exists
if exist venv (
  echo "Virtual environment found."
) else (
  echo "Virtual environment not found."
  echo.
  CHOICE /C YNC /M "Do you want to create one? (Yes/No/Cancel)"
  IF ERRORLEVEL 3 GOTO cancel
  IF ERRORLEVEL 2 GOTO no
  IF ERRORLEVEL 1 (
    echo "Creating virtual environment..."
    python -m venv venv
    echo "Virtual environment created."
  )
)

:yes
  echo "Activating virtual environment..."
  call .\venv\Scripts\activate

  echo "Installing requirements..."
  pip install -r requirements.txt

  echo "Running main.py..."
  python main.py
  goto end

:no
  echo Running without virtual environment.
  python main.py
  goto end

:cancel
  echo Operation cancelled.
  goto end

:end
echo Script completed.
pause
