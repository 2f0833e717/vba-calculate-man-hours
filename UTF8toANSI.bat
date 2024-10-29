@echo off
REM 参考
REM https://www.shegolab.jp/entry/windows-conv-text-utf8

:UTF-8 -> Shift_JIS
setlocal enabledelayedexpansion

REM フォルダパスを設定（指定のフォルダパス内の.txtを読み込む）
set folderPath=C:\Users\user\work\000_memo\_bk\yyyymm

for %%f in (%folderPath%\*.txt %folderPath%\*.csv) do (
    echo %%~ff| findstr /| /e /i ".txt .csv"
    if !ERRORLEVEL! equ 0 (
        powershell -nop -c "&{[IO.File]::WriteAllText($args[0], [IO.File]::ReadAllText($args[0], [Text.Encoding]::UTF8), [TEXT.Encoding]::GetEncoding(932))}" \"%%~ff\" \"%%~ff\"
    )
)