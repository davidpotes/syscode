On Error Resume Next
Const ForWriting = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objDictionary = CreateObject("Scripting.Dictionary")
strDomain = "dc.gtnexus.com" ' Your Domain (e.g., monster.com)
strPCsFile = "Service_Tag.txt"
strPath = "C:\scripts\" ' Create this folder
strWorkstationID = "C:\scripts\Service_Tag.txt"
i = 0

'#########
Set objFSO3 = CreateObject("Scripting.FileSystemObject")

If objFSO3.FileExists(strPath & strPCsFile) Then
 Set objFolder = objFSO3.GetFile(strPath & strPCsFile)

Else
  Set objFSOf = CreateObject("Scripting.FileSystemObject")
  If objFSOf.FolderExists(strPath) Then
   Set objFolder = objFSOf.GetFolder(strPath)
  Else
   Set objFolder = objFSO3.CreateFolder(strPath)
  End If
 Set objFile1 = objFSO3.CreateTextFile(strPath & strPCsFile)
 objFile1.Close

End If
'#########

Wscript.Echo "This program will return the Dell Service Tag on remote computer(s)"
strMbox = MsgBox("Would you like info for the entire domain:" & strDomain & "?",4,"Hostname")

If strMbox = 6 Then
'Set objPCTXTFile = objFSO.OpenTextFile(strPath & strPCsFile, ForWriting, True)
Set objDomain = GetObject("WinNT://" & strDomain) ' Note LDAP does not work
objDomain.Filter = Array("Computer")
For Each pcObject In objDomain 
objDictionary.Add i, pcObject.Name
'Wscript.Echo objDictionary(i)
i = i + 1

Next

For each DomainPC in objDictionary
strComputer = objDictionary.Item(DomainPC)
GetWorkstationID()
Next

Set objFilesystem = Nothing

WScript.echo "Finished Scanning Network check : " & strPath

Else
strHost = InputBox("Enter the computer you wish to find the Dell Service Tag from","Hostname"," ")
GetWorkstationID2()
End If

Sub GetWorkstationID()

On Error Resume Next

Set objWMIservice = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
set colitems = objWMIservice.ExecQuery("Select * from Win32_BIOS",,48)
For each objitem in colitems
 
      strWriteFile = strComputer &"   Dell Service Tag: " & objitem.serialnumber

Next
'#############

Set objFileSystem = CreateObject("Scripting.fileSystemObject")
Set objOutputFile = objFileSystem.OpenTextFile(strPath & strPCsFile, ForWriting, True)
objOutputFile.WriteLine(strWriteFile)
objOutputFile.Close

'#############
End Sub

Sub GetWorkstationID2()

On Error Resume Next
Wscript.echo strComputer & ": " & objitem.serialnumber
Set objWMIservice = GetObject("winmgmts:\\" & strHost & "\root\cimv2")
set colitems = objWMIservice.ExecQuery("Select * from Win32_BIOS",,48)
For each objitem in colitems

      Wscript.echo "Computer: " & strHost & "   Dell Service Tag: " & objitem.serialnumber
   strWriteFile = "Computer: " & strHost &"   Dell Service Tag: " & objitem.serialnumber

   Next
'%%%%%%%%%%

Set objFileSystem = CreateObject("Scripting.fileSystemObject")
Set objOutputFile = objFileSystem.OpenTextFile(strPath & strPCsFile, ForWriting, True)
objOutputFile.WriteLine(strWriteFile)
objOutputFile.Close

'%%%%%%%%%%
End Sub 