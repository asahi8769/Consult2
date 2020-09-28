@ECHO ON
title dep Start

cd /D %~dp0

%~dp0\venv\Scripts\python.exe installdep.py

cmd.exe