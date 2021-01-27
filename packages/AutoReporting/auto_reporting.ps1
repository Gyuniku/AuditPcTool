##### Windows Update, Windows Defender, 導入済ソフトウェア一覧をMongoDBに送信 #####

# 設定モジュール読み込み・初期セットアップ実行
. ".\config.ps1"

## 関数定義
# Windows Updateのログを取得して配列で返す関数
function create_update_logs {
    $logger.info.Invoke("Windows Update Logs getting...")
    # WMIを使用してUpdateのログを取得し、必要なものを抜き出す
    $update_logs = Get-WMIObject Win32_QuickFixEngineering | select  Description, HotFixID, InstalledOn
    $logger.info.Invoke("Completed. Obtained Windows Update logs : " + $update_logs.Count)

    # データを整形
    $update_val = New-Object System.Collections.ArrayList
    foreach($log in $update_logs) {
        $temp = [ordered]@{}
        $temp.Add("Description", $log.Description)
        $temp.Add("HotFixID", $log.HotFixID)
        $temp.Add("InstalledOn", $log.InstalledOn)
        $update_val.Add($temp) | Out-Null
    }

    return $update_val
}

# Windows Defenderの状態を取得して連想配列で返す関数
function create_defender_logs {
    $logger.info.Invoke("Windows Defender Logs getting...")
    # Windows DefenderのログをJSON形式で取得
    $defender_logs = Get-MpComputerStatus | ConvertTo-Json -Compress

    # 扱いやすくするためにHashTableにする
    $defender_hash = $serializer.Deserialize($defender_logs, [System.Collections.Hashtable])
    # 不要な情報を削除
    $defender_hash.Remove("CimClass")
    $defender_hash.Remove("CimInstanceProperties")
    $defender_hash.Remove("CimSystemProperties")
    $logger.info.Invoke("Completed. Obtained Windows Defender status columns : " + $defender_hash.Count)

    return $defender_hash
}

# インストール済ソフトウェア一覧を取得して配列で返す関数
function create_software_list {
    $logger.info.Invoke("Installed Software list getting...")
    # ソフトウェア一覧を取得
    $software_list = Get-ChildItem -Path('HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall','HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall')  |  % { Get-ItemProperty $_.PsPath | Select-Object DisplayName, DisplayVersion, Publisher} | ? { $_.DisplayName -ne $null }
    $logger.info.Invoke("Completed. Installed Softwares list : " + $software_list.Count)

    # 取得したソフトウェア一覧の中身から必要なデータを1件ずつ抽出
    $softwares_val = New-Object System.Collections.ArrayList
    foreach($software in $software_list) {
        $temp = [ordered]@{}
        $temp.Add("DisplayName", $software.DisplayName)

        # バージョン情報や発行者がある場合は追加し、ない場合は空文字にする。ソフトウェア名が取得できないのはありえないので不要
        if ($software.DisplayVersion) {
            $temp.Add("DisplayVersion", $software.DisplayVersion)
        } else {
            $temp.Add("DisplayVersion", "")
        }
        if ($software.Publisher) {
            $temp.Add("Publisher", $software.Publisher)
        } else {
            $temp.Add("Publisher", "")
        }

        $softwares_val.Add($temp) | Out-Null
    }

    return $softwares_val
}

# Windows Defenderの脅威検出ログを取得し、取得できた場合は専用のコレクションに挿入する関数
function post_threats_log {
    $logger.info.Invoke("Windows Defender Threat Logs getting...")
    # Windows Defenderの脅威検出ログを取得
    $threat_logs = @(Get-MpThreatDetection)

    # 脅威検出ログが1件以上の場合は処理を行い、0件の場合はログを出力して終了
    if ($threat_logs) {
        # すでに登録されていたドキュメントの数を持つ変数
        $registered_count = 0

        $logger.info.Invoke("Completed. Threat Logs : " + $threat_logs.Length)
        $logger.info.Invoke("Postting to MongoDB...")

        # ログの内容を一件ずつ整形し専用コレクションに登録する
        # MongoDBの専用コレクションに"DetectionID"と"LastThreatStatusChangeTime"で一意性制約をかけているため、すでに追加されており、更新がないコレクションは登録しない
        foreach($log in $threat_logs) {
            try {
                # Detection IDは余計な括弧を削除
                $DetectionID = (($log.DetectionID) -replace "[{|}]", "")

                # 必要な要素を追加
                $ThreatInfo = [ordered]@{}
                $ThreatInfo.Add("ActionSuccess", $log.ActionSuccess)
                $ThreatInfo.Add("AdditionalActionsBitMask", $log.AdditionalActionsBitMask)
                $ThreatInfo.Add("AMProductVersion", $log.AMProductVersion)
                $ThreatInfo.Add("CleaningActionID", $log.CleaningActionID)
                $ThreatInfo.Add("CurrentThreatExecutionStatusID", $log.CurrentThreatExecutionStatusID)
                $ThreatInfo.Add("DetectionID", $DetectionID)
                $ThreatInfo.Add("DetectionSourceTypeID", $log.DetectionSourceTypeID)
                $ThreatInfo.Add("DomainUser", $log.DomainUser)
                $ThreatInfo.Add("InitialDetectionTime", $log.InitialDetectionTime)
                $ThreatInfo.Add("LastThreatStatusChangeTime", $log.LastThreatStatusChangeTime)
                $ThreatInfo.Add("ProcessName", $log.ProcessName)
                $ThreatInfo.Add("RemediationTime", $log.RemediationTime)
                $ThreatInfo.Add("Resources", $log.Resources)
                $ThreatInfo.Add("ThreatID", $log.ThreatID)
                $ThreatInfo.Add("ThreatStatusErrorCode", $log.ThreatStatusErrorCode)
                $ThreatInfo.Add("ThreatStatusID", $log.ThreatStatusID)

                # 専用コレクション用の登録用ドキュメントを作成
                [MongoDB.Bson.BsonDocument]$doc_threat = [ordered]@{
                    DetectionID = $DetectionID
                    LastThreatStatusChangeTime = $log.LastThreatStatusChangeTime
                    ComputerName = $computer_name
                    ThreatInfo = $ThreatInfo
                }

                $coll_threat.Insert($doc_threat)
            } catch [MongoDB.Driver.MongoDuplicateKeyException] {
                # 一意制約違反が発生した場合は"DetectionID"と"LastThreatStatusChangeTime", "ThreatID"をログに出力
                $logger.info.Invoke("Document has been registered. DetectionID = " + $log.DetectionID + ", InitialDetectionTime = " + $log.InitialDetectionTime + ", ThreatID = " + $log.ThreatID.ToString())
                # すでに登録されていたドキュメントの数をインクリメント
                $registered_count++
            }
        }

        $logger.info.Invoke("Posted to MongoDB. Posted Threat Logs count : " + (($threat_logs.Length) - $registered_count))
    } else {
        # 脅威検出ログが取得できなかった場合はログを出力して終了
        $logger.info.Invoke("Completed. There are no results.")
    }
}

## メイン関数
$logger.info.Invoke("---------- Auto Reporting START ----------")

# 現在日付を取得
$now = (Get-Date).ToUniversalTime()
$logger.info.Invoke("Timestamp : " + $now)
# 環境変数からコンピュータ名を取得
$computer_name = $ENV:COMPUTERNAME
$logger.info.Invoke("ComputerName : " + $computer_name)

# Windows Updateのログを取得
$update_val = create_update_logs
# Windows Defenderのログを取得
$defender_hash = create_defender_logs
# インストールされているソフトウェア一覧を取得
$softwares_val = create_software_list

# win_logカラム用の登録用ドキュメントを作成
$status = [ordered]@{
    TimeStamp = $now
    Update = @($update_val)
    Defender = $defender_hash
    Software = @($softwares_val)
}
# 生成したデータをMongoDB用のドキュメントに変換
[MongoDB.Bson.BsonDocument]$doc_log = [ordered]@{
    ComputerName = $computer_name
    ScriptVersion = $script_version
    Status = $status
}

# MongoDBのWindowsログ用コレクションに書き込み
$logger.info.Invoke("Postting to MongoDB...")
$coll_log.Insert($doc_log)
$logger.info.Invoke("Posted to MongoDB")

# Windows Defenderの脅威検出ログを取得し、MongoDBの脅威検出ログ用コレクションに書き込み
post_threats_log

$logger.info.Invoke("---------- Auto Reporting END ----------")
