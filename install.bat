@echo off

echo �C���X�g�[�����J�n���܂��B��낵����Ή����̃L�[�������Ă�������...
pause > nul
echo;
cd /d %~dp0
powershell -NoProfile -ExecutionPolicy Unrestricted packages\register.ps1 -verb runas
echo;
echo �����̃L�[�������ďI�����Ă�������...
pause > nul