Function HandleDetermineDirectory(ExpectedSubDir As String, Optional ServerRootDir As String) As String
Dim FinalDir As String

If ServerRootDir = "" Then
   FinalDir = CurrentProject.Path
Else
  FinalDir = ServerRootDir
End If

If Right(FinalDir, 1) <> "\" Then
  FinalDir = FinalDir & "\"
End If

'If ExpectedSubDir is an absolute path, uses it instead
If Mid(ExpectedSubDir, 2, 2) = ":\" Or Left(ExpectedSubDir, 2) = "\\" Then
   FinalDir = ExpectedSubDir
Else
  FinalDir = FinalDir & ExpectedSubDir
End If


If Right(FinalDir, 1) <> "\" Then
  FinalDir = FinalDir & "\"
End If

'Make the directory
Call MakeCheckPath(FinalDir)

HandleDetermineDirectory = FinalDir

End Function


Function DoesFileExist(sPath As String, filename As String) As Boolean

DoesFileExist = False

Dim FSad, FOad, fcad, fi
Set FSad = CreateObject("Scripting.FileSystemObject")
Set FOad = FSad.GetFolder(sPath)
Set fcad = FOad.Files
For Each fi In fcad
 If fi.Name = filename Then
     DoesFileExist = True
 End If
Next

End Function

Function RetrieveFile(sPath As String, filename As String) As Object

Dim FSad, FOad, fcad, fi

Set RetrieveFile = Nothing

Set FSad = CreateObject("Scripting.FileSystemObject")
Set FOad = FSad.GetFolder(sPath)
Set fcad = FOad.Files
For Each fi In fcad
 If fi.Name = filename Then
     Set RetrieveFile = fi
     Exit Function
 End If
Next


End Function


Sub DeleteFile(sPath As String, filename As String)
    
Dim FSad, FOad, fcad, fi
Set FSad = CreateObject("Scripting.FileSystemObject")
Set FOad = FSad.GetFolder(sPath)
Set fcad = FOad.Files
For Each fi In fcad
 If fi.Name = filename Then
     fi.Delete
 End If
Next

End Sub


Sub MakeCheckPath(sPath As String)
Dim fso, folder, f1, sFolders, sf
Dim DirExists As Boolean


Dim pathComponents() As String
Dim CheckPath As String
Dim i As Integer
Dim firstComp As Integer
Dim iniPath As String
iniPath = sPath
If Left(iniPath, 2) = "\\" Then
    sPath = Right(sPath, Len(sPath) - 2)
End If
pathComponents = Split(sPath, "\")
CheckPath = ""
firstComp = 0
If Left(iniPath, 2) = "\\" Then
   pathComponents(0) = "\\" & pathComponents(0)
   sPath = "\\" & sPath
   CheckPath = CheckPath & pathComponents(0) & "\"
   firstComp = 1
End If



For i = firstComp To UBound(pathComponents)

  CheckPath = CheckPath & pathComponents(i) & "\"
  If Not FileFolderExists(CheckPath) Then
      MkDir CheckPath
      'CheckPath = CheckPath
  End If

Next




End Sub


Public Function FileFolderExists(strFullPath As String) As Boolean
'Author       : Ken Puls (www.excelguru.ca)
'Macro Purpose: Check if a file or folder exists

    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    
EarlyExit:
    On Error GoTo 0

End Function


Sub CreateSubFolder(sPath As String, FolderName As String)

Dim fso, folder, f1, sFolders, sf
Dim DirExists As Boolean

Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(sPath)
Set sFolders = folder.subFolders

DirExists = False
If Not sFolders Is Nothing Then
    For Each f1 In sFolders
      If f1.Name = FolderName Then
         DirExists = True
      End If
    Next
End If

If Not DirExists Then
   fso.CreateFolder (sPath & "\" & FolderName)
End If


End Sub
