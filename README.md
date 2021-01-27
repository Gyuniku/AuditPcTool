# 概要
PC監査ツール  
PCの月次報告に用いるデータを自動で社内サーバに送信します。  
また、Windows Defenderによる毎月のフルスキャンの登録を行います。  

# インストール方法
1. 「install.bat」を右クリックし「管理者として実行」をクリックしてください。
1. インストール画面が起動しますので、任意のキーを押して続行してください。
1. 「インストールが完了しました。」と表示されたことを確認し、任意のキーを押して終了してください。インストールはこれで完了です。

# 送信されるデータについて
以下のデータが社内サーバ上のデータベースに送信されます。
* Windows Updateの適用履歴
Windows Updateの適用状況を取得します。
* Windows Defenderの状況
セキュリティソフトであるWindows Defenderのウイルス定義ファイルのバージョンやスキャン履歴などを取得します。
* Windows Defenderが検出した脅威ログ
ウイルスが発見された場合や処理された場合に記録されるWindows Defenderの脅威ログを取得します。
* インストールされているソフトウェアの一覧
インストーラを使用して導入したソフトウェアの一覧を取得します。

