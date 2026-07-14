@echo off
title Forwarding Quote AI Setup
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ai_rate_helper.ps1" -Configure
pause
