@ECHO OFF
SETLOCAL

REM === How to use === 
REM - Double click this *.bat file and then follow the prompts

REM - The serial number of the Cloud Credentials to be deleted
::SET CONTAINER_SERIAL_NUMBER=140-616616616
SET CONTAINER_SERIAL_NUMBER=140-1888294169

REM - CodeMeter utility path
SET PATH_TO_CMU32="C:\Program Files (x86)\CodeMeter\Runtime\bin\cmu32.exe"

REM - Message to user
ECHO === WARNING - Please review the below before you continue  ===
ECHO.
ECHO - This script will delete the following specified CmCloud Credentials from this device
ECHO - Container serial number: %CONTAINER_SERIAL_NUMBER%
ECHO.
ECHO - The CodeMeter runtime (installed as part of the Echoview installation) must be installed before running this script
ECHO - Please contact support@echoview.com for further instruction if required
ECHO. 

REM - Obtain user choice
ECHO - Do you want to continue? [Y/[N]]
SET /P ANSWER=

IF /I "%ANSWER%" == "Y" (
  GOTO :continue
) ELSE (
  ECHO - Script will close on continue
  PAUSE
  EXIT /B
)

:continue
ECHO - You chose to continue
ECHO.
REM - CmCloud deletion command
%PATH_TO_CMU32% --delete-cmcloud-credentials --serial %CONTAINER_SERIAL_NUMBER%

ECHO.
ECHO - Please review the above output
ECHO - Script complete and will close on continue
PAUSE

ENDLOCAL
@ECHO OFF
