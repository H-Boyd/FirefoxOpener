Set WshShell = CreateObject("WScript.Shell")

ReDim arr(WScript.Arguments.Count-1)
For i = 0 To WScript.Arguments.Count-1
  arr(i) = WScript.Arguments(i)
Next

Arguments =  Join(arr)

cmds=WshShell.RUN("python3 C:\Users\Harry\Documents\AutomationScriptsAndTweaks\CustomScripts\FirefoxRedirector\OpenFirefox.py " + Arguments,0, True)

rem MsgBox("C:\Users\Harry\Documents\AutomationScriptsAndTweaks\CustomScripts\FirefoxRedirector\OpenFirefox.py " + WScript.Arguments(0))

Set WshShell = Nothing