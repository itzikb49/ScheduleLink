@echo off
REM ═══════════════════════════════════════════════
REM  ScheduleLink Post-Build Script
REM  Copies DLLs + .addin to Revit Addins folders
REM ═══════════════════════════════════════════════

set REVIT_VERSIONS=2023 2024 2025 2026

REM Parameters from MSBuild:
REM %~1 = TargetPath  (full path to DLL)
REM %~2 = TargetDir   (output directory)
REM %~3 = TargetName  (assembly name without extension)
REM %~4 = ProjectDir  (project directory)

set TARGET_PATH=%~1
set TARGET_DIR=%~2
set TARGET_NAME=%~3
set PROJECT_DIR=%~4

echo.
echo ============================================
echo  ScheduleLink Post-Build Copy
echo ============================================
echo  DLL:     %TARGET_PATH%
echo  OutDir:  %TARGET_DIR%
echo  Project: %PROJECT_DIR%
echo ============================================

for %%V in (%REVIT_VERSIONS%) do (
    echo.
    echo --- Revit %%V ---

    if not exist "C:\ProgramData\Autodesk\Revit\Addins\%%V" mkdir "C:\ProgramData\Autodesk\Revit\Addins\%%V"

    if exist "%TARGET_PATH%" (
        xcopy /Y /Q "%TARGET_PATH%" "C:\ProgramData\Autodesk\Revit\Addins\%%V\"
        echo   [OK] ScheduleLink.dll
    )

    if exist "%PROJECT_DIR%ScheduleLink.addin" (
        xcopy /Y /Q "%PROJECT_DIR%ScheduleLink.addin" "C:\ProgramData\Autodesk\Revit\Addins\%%V\"
        echo   [OK] ScheduleLink.addin
    )

    if exist "%TARGET_DIR%%TARGET_NAME%.pdb" (
        xcopy /Y /Q "%TARGET_DIR%%TARGET_NAME%.pdb" "C:\ProgramData\Autodesk\Revit\Addins\%%V\"
        echo   [OK] PDB
    )

    if exist "%TARGET_DIR%EPPlus.dll" (
        xcopy /Y /Q "%TARGET_DIR%EPPlus.dll" "C:\ProgramData\Autodesk\Revit\Addins\%%V\"
        echo   [OK] EPPlus.dll
    ) else (
        echo   [WARN] EPPlus.dll not found
    )
)

echo.
echo ============================================
echo  Post-build completed: %REVIT_VERSIONS%
echo ============================================

exit /b 0
