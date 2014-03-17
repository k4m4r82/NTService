cls
echo. install komponen NT Service Control Module
pause
copy NTSVC.ocx %systemroot%\system32
regsvr32 /s %systemroot%\system32\NTSVC.ocx
