@echo off
REM Helper clasp - Suivi Heure Caisson
set DEPLOY_ID=AKfycbw5BuzvcNbCMqIDWaywk9rEeciykmIWkkFmklx5j55S2SvDCxWYwtHUVJ7OoSnHndLfXA

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
