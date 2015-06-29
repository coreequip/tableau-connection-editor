@echo off
:: Poweshell Version 2 compatible, Singlethreaded
powershell -Version 2 -Sta -WindowStyle Hidden -ExecutionPolicy Unrestricted -File "%~dp0\%~n0.ps1" %*
