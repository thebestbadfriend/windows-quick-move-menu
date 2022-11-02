@echo off

powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -Command "Start-Process -Verb RunAs powershell -ArgumentList '-NoProfile -ExecutionPolicy RemoteSigned -File \"C:\Toolbox\Coding\PowerShell\Scripts\QuickMove\RemoveFrom-QuickMove.ps1\" %1'"