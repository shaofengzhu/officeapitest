set _targetdir=%~d0\officeapi
rd /s /q %_targetdir%
call node %SRCROOT%\richapi\Tools\buildofficeapinpm.js %_targetdir%
pushd %_targetdir%
call npm link
cd %~dp0
call npm link @microsoft/office-api
call node_modules\.bin\webpack 
popd