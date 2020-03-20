' Writer : JWLK
' GitHub : https://github.com/JWLK-HASLO
' Email : jewels@haslo.co

loginId = "[ID-String]"
loginPw = "[PW-String]#"

Dim objIE : Set objIE = CreateObject("InternetExplorer.Application")
Dim WshShell : Set WshShell = WScript.CreateObject ("WScript.Shell")

Function waitIE(obj)
     Do While obj.Busy = True Or obj.ReadyState <> 4
          WScript.Sleep 100
     Loop
End Function

objIE.Navigate "[Site-URL]"
objIE.Visible = True
WshShell.AppActivate objIE 'Click IE Browser < This Need To Tab Control

With objIE.Document
  .getElementByid("[InputID-TagID]").Value = loginId
  .getElementByid("[InputPW-TagID]").Value = loginPw
  .getElementByid("[EnterButton-TagID]").Click
End With
waitIE(objIE)
