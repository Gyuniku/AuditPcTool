## ディレクトリ設定
# 自動レポーティングツールが格納されたディレクトリ名
$TOOL_DIR_NAME = "\AutoReporting" 
# 実行するファイル名
$FILE_NAME = "\auto_reporting.ps1"
# このファイルのあるディレクトリ
$CURRENT_DIR = (Split-Path -Parent $MyInvocation.MyCommand.Path)
# 定期実行スクリプトが入ったディレクトリ
$COPY_FROM = $CURRENT_DIR + $TOOL_DIR_NAME
# 配置するディレクトリ
$COPY_TO = $HOME

# ロガー読み込み
. ($CURRENT_DIR+"\AutoReporting\components\mypss-master\Get-Logger.ps1")
# ロガー設定
$logger = Get-Logger -Logfile ((Split-Path -Parent $CURRENT_DIR) + "\install.log")
$logger.info.Invoke("インストールスクリプトを実行しています...")

## タスクスケジューラ設定
# タスクの場所
$TASKPATH = "\"
# タスク名
$LOGGING_TASKNAME = "DevPC-Logging"
$SCAN_TASKNAME = "Windows Defender Full Scan"
$ONCE_SCAN_TASKNAME = "Windows Defender Full Scan(Once)"
# 実行アカウント
$USER = "SYSTEM"
$USER_TCS = "USER"
# 実行権限
$RUNLEVEL = "Highest"
# トリガー設定
# レポーティングツールは月曜日～金曜日の13:00:00に起動する
$LOGGING_TRIGGER = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "13:00:00"
# フルスキャンは毎月第四木曜日の12:30:00に起動する(10月配布として11月22日(木) 12:30:00開始)
$SCAN_TRIGGER = New-ScheduledTaskTrigger -Weekly -At "2018/11/22 12:30:00" -DaysOfWeek Thursday -WeeksInterval 4
# スキャン履歴を埋めるため翌日の12:30:00にフルスキャンを一度だけ起動する
$ONCE_SCAN_DATETIME = ((Get-Date).AddDays(1)).ToString("yyyy/MM/dd HH:mm:ss")
$ONCE_SCAN_TRIGGER = New-ScheduledTaskTrigger -Once -At $ONCE_SCAN_DATETIME
# 行う操作
# ログスクリプトをPowerShellで実行する(実行時にポリシーオプションを付与)
$LOGGING_ACTION = New-ScheduledTaskAction -Execute "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument (" -ExecutionPolicy RemoteSigned " + $COPY_TO + $TOOL_DIR_NAME + $FILE_NAME) -WorkingDirectory ($COPY_TO + $TOOL_DIR_NAME)
# フルスキャン用バッチを実行する
$SCAN_ACTION = New-ScheduledTaskAction -Execute ('"' + $COPY_TO + $TOOL_DIR_NAME + '\defender\scan.bat"')
# 設定　スケジュールされた時刻にタスクを開始できなかった場合、すぐにタスクを実行する
$SETTING = New-ScheduledTaskSettingsSet -StartWhenAvailable

## メイン関数
# フォルダをコピーする。フォルダがすでに存在する場合は破棄する
try {
    if(Test-Path ($COPY_TO + $TOOL_DIR_NAME)) {
        Remove-Item ($COPY_TO + $TOOL_DIR_NAME) -Force -Recurse -ErrorAction Stop
    }
    # フォルダをコピー
    Copy-Item -Path $COPY_FROM -Destination $COPY_TO -Recurse -ErrorAction Stop
    $logger.info.Invoke("フォルダをコピーしました。")
} catch [Exception] {
    $logger.error.Invole("フォルダのコピーに失敗しました。")
    exit
}

# タスクスケジューラにレポーティングツール用タスクが登録されているか確認する。すでに登録されている場合は削除
$task = Get-ScheduledTask | Where-Object {$_.TaskName -like $LOGGING_TASKNAME }
if ($task) {
    Unregister-ScheduledTask -TaskName $LOGGING_TASKNAME -Confirm:$false
    $logger.info.Invoke("すでにタスクスケジューラに登録されていたため、レポーティングツール用タスクを削除しました。")
} else {
    $logger.info.Invoke("タスクスケジューラにレポーティングツール用のタスクが登録されていません。登録を行います。")
}

# レポーティングツール用タスクをタスクスケジューラに登録する
try {
    Register-ScheduledTask -TaskPath $TASKPATH -TaskName $LOGGING_TASKNAME -User $USER -RunLevel $RUNLEVEL -Trigger $LOGGING_TRIGGER -Action $LOGGING_ACTION -ErrorAction Stop | Out-Null
} catch [Exception] {
    $logger.error.Invoke("タスクスケジューラへの登録に失敗しました。管理者権限を持つ状態で実行しているか確認してください。")
    exit
}

$logger.info.Invoke("レポーティングツール用タスクの登録が完了しました。タスクを実行します...")

# レポーティングツール用タスク実行 
Get-ScheduledTask -TaskName $LOGGING_TASKNAME | Start-ScheduledTask -ErrorAction Stop
$logger.info.Invoke("レポーティングツール用タスクの初回実行が完了しました。")

# セキュリティスキャン用タスク(定期実行用・一度だけ実行用)がすでにタスクスケジューラに登録されている場合は削除
$task = Get-ScheduledTask | Where-Object {$_.TaskName -like $SCAN_TASKNAME }
if ($task) {
    Unregister-ScheduledTask -TaskName $SCAN_TASKNAME -Confirm:$false
    $logger.info.Invoke("すでにタスクスケジューラに登録されていたため、セキュリティスキャン用タスクを削除しました。")
} else {
    $logger.info.Invoke("タスクスケジューラにセキュリティスキャン用タスクが登録されていません。登録を行います。")
}
$task = Get-ScheduledTask | Where-Object {$_.TaskName -like $ONCE_SCAN_TASKNAME }
if ($task) {
    Unregister-ScheduledTask -TaskName $ONCE_SCAN_TASKNAME -Confirm:$false
} else {
}

# セキュリティスキャン用タスク(定期実行用・一度だけ実行用)をタスクスケジューラに登録
try {
    Register-ScheduledTask -TaskPath $TASKPATH -TaskName $SCAN_TASKNAME -Trigger $SCAN_TRIGGER -Action $SCAN_ACTION -Settings $SETTING -ErrorAction Stop | Out-Null
    Register-ScheduledTask -TaskPath $TASKPATH -TaskName $ONCE_SCAN_TASKNAME -Trigger $ONCE_SCAN_TRIGGER -Action $SCAN_ACTION -Settings $SETTING -ErrorAction Stop | Out-Null
} catch [Exception] {
    $logger.error.Invoke("タスクスケジューラへの登録に失敗しました。管理者権限を持つ状態で実行しているか確認してください。")
    exit
}
$logger.info.Invoke("セキュリティスキャン用タスクの登録が完了しました。")

$logger.info.Invoke("インストールが完了しました。")