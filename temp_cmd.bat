
    @echo off

    rem Define the share name of the printer you want to switch to
    set "target_printer=55EXP_Purolator"

    rem Check if the printer with the target share name exists and set it as default
    powershell -Command "Get-WmiObject -Query \"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\" | Invoke-WmiMethod -Name SetDefaultPrinter"

    echo Hello! From the Purolator Printer.

    rem Execute BarTender command
    "C:\Seagull\BarTender 7.10\Standard\bartend.exe" /f=C:\BTAutomation\XI6.btw /p /x
    