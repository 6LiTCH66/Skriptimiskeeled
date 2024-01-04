@echo off
setlocal EnableDelayedExpansion

set "input=%~1"
set "output="

set "total=0"
set "count_ä=0"
set "count_ö=0"
set "count_ü=0"
set "count_õ=0"
set "count_š=0"
set "count_ž=0"

:replace
if "!input!"=="" goto end
set "char=!input:~0,1!"
set "rest=!input:~1!"

if "!char!"=="ä" (
    set "output=!output!&auml;"
    set /a count_ä+=1
    set /a total+=1
    goto process_next
)

if "!char!"=="ö" (
    set "output=!output!&ouml;"
    set /a count_ö+=1
    set /a total+=1
    goto process_next
)

if "!char!"=="ü" (
    set "output=!output!&uuml;"
    set /a count_ü+=1
    set /a total+=1
    goto process_next
)

if "!char!"=="õ" (
    set "output=!output!&otilde;"
    set /a count_õ+=1
    set /a total+=1
    goto process_next
)

if "!char!"=="š" (
    set "output=!output!&scaron;"
    set /a count_š+=1
    set /a total+=1
    goto process_next
)

if "!char!"=="ž" (
    set "output=!output!&zcaron;"
    set /a count_ž+=1
    set /a total+=1
    goto process_next
)

set "output=!output!!char!"

:process_next
set "input=!rest!"
goto replace

:end
echo !output!
if !total! gtr 0 (
    echo Vahetatud:
    if !count_ä! gtr 0 echo ä: !count_ä!
    if !count_ö! gtr 0 echo ö: !count_ö!
    if !count_ü! gtr 0 echo ü: !count_ü!
    if !count_õ! gtr 0 echo õ: !count_õ!
    if !count_š! gtr 0 echo š: !count_š!
    if !count_ž! gtr 0 echo ž: !count_ž!
    echo Kokku: !total!
) else (
    echo Ei leidnud ühtegi täpitähte.
)
