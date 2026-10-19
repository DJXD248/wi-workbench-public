@echo off
chcp 936 >/dev/null 2>&1
cd /d "%~dp0"

if not exist "作业指导书工作台.exe" (
    if not exist "发布包\作业指导书工作台.exe" (
        echo 错误: 未找到 作业指导书工作台.exe
        echo 请将发布包解压到此目录后双击此 bat 启动。
        pause
        exit /b 1
    )
    start "" "发布包\作业指导书工作台.exe"
) else (
    start "" "作业指导书工作台.exe"
)
exit /b
