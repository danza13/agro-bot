@echo off
set "output=output.txt"
if exist "%output%" del "%output%"

for %%f in (*) do (
    echo ----- %%f ----- >> "%output%"
    type "%%f" >> "%output%"
    echo. >> "%output%"
)
echo Завершено! Дані записано у %output%
pause
