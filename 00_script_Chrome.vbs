' Writer : JWLK
' GitHub : https://github.com/JWLK-HASLO
' Email : jewels@haslo.co

'Find App Using Registry
function readFromRegistry (strRegistryKey, strDefault)
    Dim WSHShell, value
    On Error Resume Next
    Set WSHShell = CreateObject ("WScript.Shell")
    value = WSHShell.RegRead (strRegistryKey)

    if err.number <> 0 then
        readFromRegistry= strDefault
    else
        readFromRegistry=value
    end if

    set WSHShell = nothing
end function

'Chrome Browser URL Start
function OpenWithChrome(strURL)
    Dim strChrome
    Dim WShellChrome
    strChrome = readFromRegistry ( "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\Path", "") 
    if (strChrome = "") then
        strChrome = "chrome.exe"
    else
        strChrome = strChrome & "\chrome.exe"
    end if
    Set WShellChrome = CreateObject("WScript.Shell")
    strChrome = """" & strChrome & """" & " " & strURL
    WShellChrome.Run strChrome, 1, false
end function


'Start Script
OpenWithChrome "[Site-URL]"

