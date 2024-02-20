strCommand = "cmd /c scrcpy.exe --tcpip=192.168.100.125:39875"

For Each Arg In WScript.Arguments
    strCommand = strCommand & " """ & replace(Arg, """", """""""""") & """"
Next

CreateObject("Wscript.Shell").Run strCommand, 0, false
