@echo off

echo インストールを開始します。よろしければ何かのキーを押してください...
pause > nul
echo;
cd /d %~dp0
powershell -NoProfile -ExecutionPolicy Unrestricted packages\register.ps1 -verb runas
echo;
echo 何かのキーを押して終了してください...
pause > nul