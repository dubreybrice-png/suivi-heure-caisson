@echo off
REM Helper clasp - Suivi Heure Caisson
set DEPLOY_ID=AKfycbxNE4sqgXKZvfSUQsFP02rziShhSYtzATFSk5Q_6wGiPE6ALpia5GJFV-VMsSkMfms2nQ

if "%1"=="push" (
  cmd /c "clasp push --force"
  goto :eof
)

if "%1"=="deploy" (
  if "%DEPLOY_ID%"=="" (
    echo DEPLOY_ID non defini. Definissez-le dans ce fichier.
    goto :eof
  )
  cmd /c "clasp redeploy %DEPLOY_ID% -d stable"
  goto :eof
)

if "%1"=="pushdeploy" (
  cmd /c "clasp push --force"
  if "%DEPLOY_ID%"=="" (
    echo Push effectue. DEPLOY_ID non defini, redeploy ignore.
    goto :eof
  )
  cmd /c "clasp redeploy %DEPLOY_ID% -d stable"
  goto :eof
)

if "%1"=="open" (
  if "%DEPLOY_ID%"=="" (
    cmd /c "clasp open"
  ) else (
    cmd /c "clasp open-web-app %DEPLOY_ID%"
  )
  goto :eof
)

echo Usage: clasp-helper.bat [push^|deploy^|pushdeploy^|open]
