@echo off
title Attendance Dashboard
cd /d "D:\Spine HR 2"
echo Starting Attendance Dashboard...
echo.
echo Open your browser at: http://localhost:8501
echo.
python -m streamlit run dashboard.py
pause
