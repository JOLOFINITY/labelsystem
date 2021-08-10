Sub HTTPDownload( myURL, myPath )
    Dim i, objFile, objFSO, objHTTP, strFile, strMsg
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )
    If objFSO.FolderExists( myPath ) Then
        strFile = objFSO.BuildPath( myPath, Mid( myURL, InStrRev( myURL, "/" ) + 1 ) )
    ElseIf objFSO.FolderExists( Left( myPath, InStrRev( myPath, "\" ) - 1 ) ) Then
        strFile = myPath
    Else
        WScript.Echo "ERROR: Target folder not found."
        Exit Sub
    End If
    Set objFile = objFSO.OpenTextFile( strFile, ForWriting, True )
    Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
    objHTTP.Open "GET", myURL, False
    objHTTP.Send
    For i = 1 To LenB( objHTTP.ResponseBody )
        objFile.Write Chr( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
    Next
    objFile.Close( )
End Sub

HTTPDownload "https://raw.githubusercontent.com/JOLOFINITY/labelsystem/main/update.zip", "C:\Users\Administrator\Desktop\testing Script\update.zip"

Msgbox "Update Successfully Updated",0,"Download Completed"

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'Label System.exe'")

For Each objProcess in colProcessList
    objProcess.Terminate()
Next

Dim Fso
Set Fso = WScript.CreateObject("Scripting.FileSystemObject")

Msgbox "Script will now update the Local App.",0,"Information"

If Fso.FileExists("Label System.exe") Then 
 Fso.DeleteFile "Label System.exe"
End If

Fso.MoveFile "update.zip", "Label System.exe"

Msgbox "Local App. successfully update.",0,"Information"

Dim objShell
Set objShell = WScript.CreateObject( "WScript.Shell" )
objShell.Run("""Label System.exe""")
Set objShell = Nothing
