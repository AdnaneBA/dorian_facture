@echo off
title Validation Factures
echo ================================================
echo   Validation Factures - Demarrage...
echo ================================================
echo.

REM Vérifie si Python est installé
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo Telechargez Python sur https://www.python.org/downloads/
    echo Cochez bien "Add Python to PATH" lors de l'installation.
    pause
    exit /b
)

REM Installe les dépendances si besoin
echo Installation des dependances (premiere fois uniquement)...
pip install -r requirements.txt --quiet

echo.
echo Ouverture de l'application dans votre navigateur...
echo Pour fermer l'application, fermez cette fenetre.
echo.

REM Lance Streamlit et ouvre le navigateur
streamlit run app.py --server.headless false --browser.gatherUsageStats false

pause
