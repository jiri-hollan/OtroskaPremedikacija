' Zagon Otroske premedikacije v locenem oknu brez orodnih vrstic brskalnika.
' Deluje iz katerekoli mape (tudi z USB kljuca), ker poisce HTML v svoji mapi.
Option Explicit
Dim sh, fso, mapa, html, url, args
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
mapa = fso.GetParentFolderName(WScript.ScriptFullName)
html = fso.BuildPath(mapa, "OtroskaPremedikacija.html")
url  = PotVUrl(html)
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

' Pretvori pot v URL: znake izven ASCII (s, c, z) in presledke zakodira v UTF-8 (%C5%A1 ...)
Function PotVUrl(pot)
  Dim st, bajti, i, b, c, izhod
  Set st = CreateObject("ADODB.Stream")
  st.Type = 2 : st.Charset = "utf-8" : st.Open
  st.WriteText Replace(pot, "\", "/")
  st.Position = 0 : st.Type = 1 : st.Position = 3   ' preskoci UTF-8 BOM
  bajti = st.Read : st.Close
  izhod = ""
  For i = 1 To LenB(bajti)
    b = AscB(MidB(bajti, i, 1))
    c = Chr(b)
    If (b >= 48 And b <= 57) Or (b >= 65 And b <= 90) Or (b >= 97 And b <= 122) Or InStr("/:-_.~", c) > 0 Then
      izhod = izhod & c
    Else
      izhod = izhod & "%" & Right("0" & Hex(b), 2)
    End If
  Next
  PotVUrl = "file:///" & izhod
End Function
