@echo off

echo Windows Defender によるフルスキャンを行います...
echo この画面を終了させないでください。スリープは問題ありません。
"C:\Program Files\Windows Defender\MpCmdRun.exe" Scan -ScanType 2 -Trace

exit 0