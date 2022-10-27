'Start loop
Do

Dim strDriveLetter
Dim intDriveLetter
'Check for error
On Error Resume Next
Set fs = CreateObject("Scripting.FileSystemObject")
strDriveLetter = ""
'Loops through all available drives, chooses first CD drive
For intDriveLetter = Asc("A") To Asc("Z")
Err.Clear
'Windows indicates CD drive type by returning 4 when queried
If fs.GetDrive(Chr(intDriveLetter)).DriveType = 4 Then
If Err.Number = 0 Then
strDriveLetter = Chr(intDriveLetter)
Exit For
End If
End If
Next

'Calls windows media player
Set oWMP = CreateObject("WMPlayer.OCX.7" )
Set colCDROMs = oWMP.cdromCollection

'For loops to close and open cd drive

For d = 0 to colCDROMs.Count - 1
colCDROMs.Item(d).Eject
Next 

For d = 0 to colCDROMs.Count - 1
colCDROMs.Item(d).Eject
Next 

'End loop
loop

