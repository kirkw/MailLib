@echo off
REM MailLib Environment Setup Script
REM This script sets up environment variables for testing MailLib
REM 
REM PRIVATE VERSION - Edit this file with your actual credentials
REM PUBLIC VERSION - Use this as a template for setting up environment variables

echo Setting up MailLib environment variables...
echo.

REM ===========================================
REM PRIVATE CONFIGURATION - REPLACE WITH YOUR VALUES
REM ===========================================
set MAILLIB_SMTP_HOST=smtp.gmail.com
set MAILLIB_SMTP_PORT=587
set MAILLIB_SMTP_USERNAME=your-email@gmail.com
set MAILLIB_SMTP_PASSWORD=your-app-password
set MAILLIB_FROM_EMAIL=sender@yourdomain.com
set MAILLIB_TO_EMAIL=recipient@example.com

REM ===========================================
REM PRIVATE CONFIGURATION - REPLACE WITH YOUR VALUES
REM ===========================================
set MAILLIB_SMTP_HOST=smtp.office365.com
set MAILLIB_SMTP_PORT=587
set MAILLIB_SMTP_USERNAME=dummy-username@example.com
set MAILLIB_SMTP_PASSWORD=dummy-password-123
set MAILLIB_FROM_EMAIL=dummy-sender@example.com
set MAILLIB_TO_EMAIL=dummy-recipient@example.com

REM ===========================================
REM DISPLAY CURRENT SETTINGS
REM ===========================================
echo Current MailLib Environment Variables:
echo =====================================
echo MAILLIB_SMTP_HOST=%MAILLIB_SMTP_HOST%
echo MAILLIB_SMTP_PORT=%MAILLIB_SMTP_PORT%
echo MAILLIB_SMTP_USERNAME=%MAILLIB_SMTP_USERNAME%
echo MAILLIB_SMTP_PASSWORD=***HIDDEN***
echo MAILLIB_FROM_EMAIL=%MAILLIB_FROM_EMAIL%
echo MAILLIB_TO_EMAIL=%MAILLIB_TO_EMAIL%
echo.

REM ===========================================
REM VALIDATE SETTINGS
REM ===========================================
if "%MAILLIB_SMTP_HOST%"=="" (
    echo ERROR: MAILLIB_SMTP_HOST is not set
    goto :error
)

if "%MAILLIB_SMTP_USERNAME%"=="" (
    echo ERROR: MAILLIB_SMTP_USERNAME is not set
    goto :error
)

if "%MAILLIB_FROM_EMAIL%"=="" (
    echo ERROR: MAILLIB_FROM_EMAIL is not set
    goto :error
)

if "%MAILLIB_TO_EMAIL%"=="" (
    echo ERROR: MAILLIB_TO_EMAIL is not set
    goto :error
)

echo Environment variables set successfully!
echo.
echo To test MailLib, run:
echo   cscript TestMailLib.vbs
echo.
echo To make these environment variables permanent, add them to your system environment variables.
echo.
pause
goto :end

:error
echo.
echo Please edit this batch file and set your actual SMTP credentials.
echo.
echo Common SMTP settings:
echo   Gmail: smtp.gmail.com:587 (use App Password)
echo   Outlook: smtp-mail.outlook.com:587
echo   Yahoo: smtp.mail.yahoo.com:587
echo.
pause
exit /b 1

:end 