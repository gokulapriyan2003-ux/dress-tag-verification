@echo off
title Dress Tag & Master Sheet Verifier
echo ========================================================
echo       Starting Dress Tag & Master Sheet Verifier
echo ========================================================
echo.
cd /d D:\dress_tag_automation
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501 --server.enableCORS false --server.enableXsrfProtection false --server.maxUploadSize 500
pause
