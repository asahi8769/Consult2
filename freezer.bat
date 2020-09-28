@ECHO ON
title freeze Start

cd /D %~dp0\

%~dp0\venv\Scripts\python.exe freezer.py

cmd.exe