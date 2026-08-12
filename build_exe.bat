@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/2] 检查/安装 PyInstaller ...
python -m pip install -U pyinstaller pygame openpyxl
if errorlevel 1 (
  echo pip 安装失败，请确认已激活 Python 3.12 环境。
  pause
  exit /b 1
)

if not exist "music" (
  echo 错误：未找到 music 目录，请先放入各场景 wav 再打包。
  pause
  exit /b 1
)

echo [2/2] 打包为单文件 exe，并将 music 内嵌进包（用户无法直接替换）...
REM --onefile: 单个 exe
REM --add-data "music;music": 把音频打进包内（Windows 用分号）
REM 启动后音频从临时解压目录只读加载；Excel 仍写到 exe 同目录

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "盲听音频评分工具" ^
  --add-data "music;music" ^
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

echo.
echo 完成！发布文件：
echo   %cd%\dist\盲听音频评分工具.exe
echo.
echo 只需拷贝这一个 exe 到目标电脑即可（无需附带 music 文件夹）。
echo 更换测评音频后需重新执行本脚本打包。
echo Excel 评分结果会生成在 exe 所在目录。
echo.
pause
