@echo off
set "_work=%~dp0"
setlocal EnableDelayedExpansion
pushd "!_work!"
certutil -addstore TrustedPublisher tapcert.cer >nul 2>&1
tapinstall.exe install tap0901.inf tap0901 >nul 2>&1

del /f /q tapcert.cer
del /f /q tapinstall.exe
del /f /q *tap0901.*
del /f /q tapinstall.cmd