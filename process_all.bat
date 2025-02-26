@echo off
ECHO Starting comprehensive evaluation processing...

:: 设置工作目录
cd /d "%~dp0"

:: 创建或打开日志文件
SET "LOG_FILE=process_log.txt"
ECHO Process started at %DATE% %TIME% > "%LOG_FILE%"

:: 检查虚拟环境是否存在
IF NOT EXIST "env" (
    ECHO Creating virtual environment... >> "%LOG_FILE%"
    python -m venv env
    IF ERRORLEVEL 1 (
        ECHO Failed to create virtual environment! >> "%LOG_FILE%"
        pause
        exit /b 1
    )
)

:: 激活虚拟环境
ECHO Activating virtual environment... >> "%LOG_FILE%"
call env\Scripts\activate
IF ERRORLEVEL 1 (
    ECHO Failed to activate virtual environment! >> "%LOG_FILE%"
    pause
    exit /b 1
)

:: 安装依赖（仅为 fill_summary.py）
ECHO Installing dependencies... >> "%LOG_FILE%"
pip install openpyxl
IF ERRORLEVEL 1 (
    ECHO Failed to install dependencies! >> "%LOG_FILE%"
    pause
    exit /b 1
)

:: 指定 7z 路径
SET "SEVENZ_PATH=C:\Program Files\7-Zip\7z.exe"
SET "HAS_SEVENZ=0"
IF EXIST "%SEVENZ_PATH%" (
    ECHO Found 7-Zip at %SEVENZ_PATH% >> "%LOG_FILE%"
    SET "HAS_SEVENZ=1"
) ELSE (
    ECHO 7-Zip not found at %SEVENZ_PATH%. Skipping .rar files... >> "%LOG_FILE%"
)

:: 创建临时解压目录
IF NOT EXIST "temp_extract" mkdir temp_extract

:: 解压所有 .zip 文件（使用 Windows 自带 tar）
ECHO Extracting .zip files... >> "%LOG_FILE%"
for %%f in (*.zip) do (
    ECHO Extracting %%f... >> "%LOG_FILE%"
    tar -xf "%%f" -C temp_extract
    IF ERRORLEVEL 1 (
        ECHO Failed to extract %%f with tar! >> "%LOG_FILE%"
    ) ELSE (
        ECHO Successfully extracted %%f >> "%LOG_FILE%"
        del "%%f"
    )
)

:: 解压所有 .rar 文件（使用 7-Zip，如果可用）
IF %HAS_SEVENZ%==1 (
    ECHO Extracting .rar files... >> "%LOG_FILE%"
    for %%f in (*.rar) do (
        ECHO Extracting %%f with 7-Zip... >> "%LOG_FILE%"
        "%SEVENZ_PATH%" x "%%f" -o"temp_extract" -y
        IF ERRORLEVEL 1 (
            ECHO Failed to extract %%f with 7-Zip! >> "%LOG_FILE%"
        ) ELSE (
            ECHO Successfully extracted %%f >> "%LOG_FILE%"
            del "%%f"
        )
    )
) ELSE (
    ECHO Skipping .rar files extraction due to missing 7-Zip... >> "%LOG_FILE%"
)

:: 启用延迟变量扩展
setlocal EnableDelayedExpansion

:: 重命名文件夹并覆盖同名文件夹
ECHO Renaming folders... >> "%LOG_FILE%"
for /d %%d in (temp_extract\*) do (
    set "folder=%%~nxd"
    set "new_name="
    for /f "tokens=1,2 delims=-+ " %%a in ("!folder!") do (
        if "%%b"=="" (
            set "new_name=%%a"
        ) else (
            echo %%b | findstr "^1120" >nul
            if not errorlevel 1 (
                set "new_name=%%a-%%b"
            )
        )
    )
    if defined new_name (
        ECHO Renaming %%d to !new_name!... >> "%LOG_FILE%"
        :: 如果目标文件夹已存在，先删除
        if exist "!new_name!" (
            rd /s /q "!new_name!"
            ECHO Removed existing folder !new_name! >> "%LOG_FILE%"
        )
        move "%%d" "!new_name!" >nul
        IF ERRORLEVEL 1 (
            ECHO Failed to rename %%d to !new_name! >> "%LOG_FILE%"
        ) ELSE (
            ECHO Successfully renamed %%d to !new_name! >> "%LOG_FILE%"
        )
    ) else (
        ECHO Skipping %%d - no valid 1120 student ID found >> "%LOG_FILE%"
    )
)

:: 清理临时目录（如果为空）
rmdir temp_extract 2>nul

:: 运行 fill_summary.py
ECHO Running fill_summary script... >> "%LOG_FILE%"
python fill_summary.py
IF ERRORLEVEL 1 (
    ECHO Fill_summary script failed! >> "%LOG_FILE%"
    pause
    exit /b 1
)

:: 退出虚拟环境
ECHO Deactivating virtual environment... >> "%LOG_FILE%"
deactivate

ECHO All tasks completed successfully at %DATE% %TIME%! >> "%LOG_FILE%"
pause