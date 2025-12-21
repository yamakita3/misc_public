@echo off
setlocal enabledelayedexpansion

echo ==== exif_gps_simple.bat 開始 ====
echo 引数1 = "%~1"
echo.

if "%~1"=="" (
    echo 画像ファイルをこのBATにドラッグ&ドロップしてください。
    pause
    exit /b
)

set "INPUT_FILE=%~f1"
set "OUTPUT_FILE=%~dpn1_exif.txt"
set "OUTPUT_CSV=%~dpn1_exif_all.csv"
set "EXIFTOOL=%~dp0exiftool.exe"

echo 入力ファイル : "%INPUT_FILE%"
echo 出力ファイル : "%OUTPUT_FILE%"
echo 出力CSV     : "%OUTPUT_CSV%"
echo 使用ExifTool : "%EXIFTOOL%"
echo.

REM ▼ EXIF全文をそのまま出力
"%EXIFTOOL%" "%INPUT_FILE%" > "%OUTPUT_FILE%"
REM "%EXIFTOOL%" -all -csv -n -c "%.8f" "%INPUT_FILE%" > "%OUTPUT_CSV%"
REM 2) GPSを度（少数）でCSV出力[web:97][web:102][web:111]
"%EXIFTOOL%" -c "%.8f" -n -charset EXIF=UTF8 ^
  -GPSLatitude -GPSLongitude -FileName ^
  -csv "%INPUT_FILE%" > "%OUTPUT_CSV%"

echo 完了しました。
echo 出力: "%OUTPUT_FILE%"
echo CSV: "%OUTPUT_CSV%"
echo.
pause
