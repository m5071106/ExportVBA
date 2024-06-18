@ECHO OFF
REM バッチ実行フォルダへ移動
CD /D %~dp0
REM 文字コードをUTF-8に設定
CHCP 65001
REM Excelフォルダ内の全ファイルと同じ名前のフォルダをVBAフォルダ配下に作成
FOR %%f IN (.\Excel\*.xlsm) DO IF NOT EXIST ".\VBA\%%~nf\" MKDIR ".\VBA\%%~nf\"
REM Excelフォルダ内の全ファイルに対してVBAスクリプトをエクスポート
FOR %%f IN (.\Excel\*.xlsm) DO CScript ExportVBA.vbs %~dp0"%%f" ".\VBA\%%~nf\"
