' Zagon Otroske premedikacije v locenem oknu brez orodnih vrstic brskalnika.
' Deluje iz katerekoli mape (tudi z USB kljuca), ker poisce HTML v svoji mapi.
Option Explicit
Dim sh, fso, mapa, html, url, args
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
mapa = fso.GetParentFolderName(WScript.ScriptFullName)
html = fso.BuildPath(mapa, "OtroskaPremedikacija.html")
url  = "file:///" & Replace(html, "\", "/")
args = " --app=""" & url & """ --window-size=560,760"

On Error Resume Next
sh.Run "msedge" & args, 1, False
If Err.Number <> 0 Then
  Err.Clear
  sh.Run "chrome" & args, 1, False
  If Err.Number <> 0 Then
    Err.Clear
    sh.Run """" & html & """", 1, False
  End If
End If
