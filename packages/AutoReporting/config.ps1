### レポーティングツールの設定モジュール

## バージョン設定
# このスクリプトのバージョン
$script_version = "1.1"

## PowerShell設定
# 例外をキャッチする設定(stop)
$ErrorActionPreference = "stop"

## ロガー設定
# ロガーのパス(カレントディレクトリからの相対パス)
$path_logger = "\components\mypss-master\Get-Logger.ps1"
# ログファイルのパス
$path_logfile = "\log\reporting.log"

## MongoDB設定
# MongoDBドライバのホームディレクトリ(カレントディレクトリからの相対パス)
$path_mongodriver_dir = "\components\CSharpDriver-1.11.0"
# MongoDBのURL
$mongo_url = "76492d1116743f0423413b16050a5345MgB8AEcAeQByADUAdwBHADkAOABXAFQAaQByAGwAcwBuAGIAegBCADkAdQBNAGcAPQA9AHwAZQA5ADcAZABhAGUAZQA5ADAAMwAwAGMAYQBhADEAOAA4ADMAYgA1AGIANwA3AGIANgBmADIANQAzADYAZAAwADYA
YgAxADUAMQBmADkAYgA0AGIANgA1AGUAMAA0AGYAZgA5ADIANwA1ADIAOAAzAGMAZgAyAGIAMwAyAGMAYgA4ADgANwBjADIAYgBkADMANQA2AGYAOAA5AGYAZQBmADUANAA3AGQAZgAzAGMAOQA3ADUAMQBiAGYANAA1AGIAMgA3ADMAMwA0ADgANgA3AGUA
YgAyADgAMQAyAGEAYQBmADkAZQBlADgAOAA3AGYAYgBhADMAOQA4ADYAMQBmAA=="
# データベース名
$mongo_db_name = "76492d1116743f0423413b16050a5345MgB8AGYASQBDAGQAcwBCAFEAUQB5AFMAYgBBADgAUAA0AHUAVgBsAEgAawA5AGcAPQA9AHwANgA2AGMANAAzADQAMgBjADQANgAzADEAMABlADgAMQBiAGUAMgA2AGYAYwA2AGEAOAA5ADIANgA5ADMANwAyAA=="
# Windowsの状態を格納するコレクション名
$mongo_coll_log = "76492d1116743f0423413b16050a5345MgB8AC8ARABIAFEAUAB5AEYASABHAFQAeABzAG8ANABOAGcAZQBiADQASwAwAHcAPQA9AHwAYwBjADMANAAxADIAMAA2ADMAYwA1AGEAYwA2ADUANAAzAGIAOQA2ADQAMwAzADYAZgA2AGYAYwA2ADAAZQA2AA=="
# Windows Defenderの脅威検出ログを格納するコレクション名
$mongo_coll_threat = "76492d1116743f0423413b16050a5345MgB8AE0ARgBTAEUAKwBCAGQAQgBLADkALwBvAGMAUwBoAHcANABRAHUAcQBOAHcAPQA9AHwAZABiADcAMgBjAGQAYQAyAGMAMwA3ADYAOQBiADEAOAAyAGMAOQA2ADAANwA5AGMAZgAyAGQAMAAzADAAOQBkAA=="

## 関数定義
# セキュアストリングをコンバートする
function SecureStringToPlainString($secure_string){
    $decrypt = ConvertTo-SecureString -String $secure_string -Key (1..16)
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($decrypt)
    $plain_string = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)

    # $BSTRを削除
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

    return $plain_string
}

## セットアップ処理
# このファイルのあるディレクトリパス
$current_dir = (Split-Path -Parent $MyInvocation.MyCommand.Path)

# MongoDB用ドライバ読み込み
Add-Type -Path ($current_dir + $path_mongodriver_dir + "\MongoDB.Bson.dll")
Add-Type -Path ($current_dir + $path_mongodriver_dir + "\MongoDB.Driver.dll")
# MongoDBに接続
$client = New-Object MongoDB.Driver.MongoClient((SecureStringToPlainString $mongo_url))
$server = $client.GetServer()
$db = $server.GetDatabase((SecureStringToPlainString $mongo_db_name))
# Windowsのログを収集するカラム
$coll_log = $db.GetCollection((SecureStringToPlainString $mongo_coll_log))
# Windows Defenderの脅威検出ログを収集するカラム
$coll_threat = $db.GetCollection((SecureStringToPlainString $mongo_coll_threat))

# ロガー読み込み
. ($current_dir + $path_logger)
# ロガー設定
$logger = Get-Logger -Logfile ($current_dir + $path_logfile)

# JSON -> HashTable用シリアライザ設定
Add-Type -AssemblyName System.Web.Extensions
$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
