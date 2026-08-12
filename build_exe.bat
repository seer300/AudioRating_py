@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/3] 检查/安装 PyInstaller ...
python -m pip install -U pyinstaller pygame openpyxl
if errorlevel 1 (
  echo pip 安装失败，请确认已激活 Python 3.12 环境。
  pause
  exit /b 1
)

echo [2/3] 开始打包（单目录版，启动更快、更稳）...
REM --noconsole: 不弹出黑框；若调试打包问题可改成 --console
REM --collect-all pygame: 把 pygame 依赖资源一并打进包
REM music 文件夹会复制到 exe 同级，便于你事后替换 wav

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name "盲听音频评分工具" ^
  --collect-all pygame ^
  --collect-all openpyxl ^
  --hidden-import openpyxl ^
  --hidden-import pygame ^
  blind_listen_score.py

if errorlevel 1 (
  echo 打包失败。
  pause
  exit /b 1
)

echo [3/3] 复制 music 音频目录到发布目录 ...
if exist "dist\盲听音频评分工具\music" rmdir /s /q "dist\盲听音频评分工具\music"
xcopy /E /I /Y "music" "dist\盲听音频评分工具\music" >nul

echo.
echo 完成！发布目录：
echo   %cd%\dist\盲听音频评分工具\
echo.
echo 把整个「盲听音频评分工具」文件夹拷到目标电脑即可运行。
echo 目标电脑无需安装 Python；请确保有音频输出设备。
echo.
pause
