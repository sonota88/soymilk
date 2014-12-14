@echo off

echo ######## dump ########

set JJS=C:\path\to\jdk\bin\jrunscript.exe
set LIBO_DIR=C:\Program Files (x86)\LibreOffice 4

set CLASSPATH=%LIBO_DIR%\program\classes\unoil.jar
set CLASSPATH=%CLASSPATH%;%LIBO_DIR%\URE\java\juh.jar
set CLASSPATH=%CLASSPATH%;%LIBO_DIR%\URE\java\jurt.jar
set CLASSPATH=%CLASSPATH%;%LIBO_DIR%\URE\java\ridl.jar
rem echo %CLASSPATH%

set ODS_FULLPATH=%1

%JJS% -encoding utf8 -classpath "%CLASSPATH%" -f dump_ods.js "%ODS_FULLPATH%"

echo ######## import ########

set CLASSPATH=mysql-connector-java-5.1.33.jar

%JJS% -encoding utf8 -classpath "%CLASSPATH%" -f import.js "%ODS_FULLPATH%"
