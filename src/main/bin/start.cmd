@echo off
if not "%JAVA_HOME%"=="" goto st
set JAVA_HOME=%cd%\jdk
set path=.;%JAVA_HOME%\bin;%path%
set CLASSPATH=.;%JAVA_HOME%\lib;%CLASSPATH%
goto st
:st
java -jar webcollection.jar %cd%\result\
pause
exit