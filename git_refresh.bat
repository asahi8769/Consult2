@ECHO ON
title git_refresh Start

cd /D %~dp0

%~dp0\venv\Scripts\python.exe git_refresh.py

cmd.exe