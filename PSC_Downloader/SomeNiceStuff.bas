Attribute VB_Name = "SomeNiceStuff"
Option Explicit

Public Function SafeFileName(ByVal oldFileName As String, ByVal FileExt As String) As String
  Dim ServDir  As String
  Dim BadChars As String
  Dim BadChar  As String
  Dim x        As Long
  BadChars = "\/:*?""<>|"
  ServDir = oldFileName
  For x = 1 To Len(BadChars)
    BadChar = Mid$(BadChars, x, 1)
    ServDir = Replace$(ServDir, BadChar, "_")
  Next
  If Len(ServDir) + 1 + Len(FileExt) > 255 Then 'FAT32 and up file path length limitation
    ServDir = Left$(ServDir, 255 - 1 - Len(FileExt) - 1)
  End If
  SafeFileName = ServDir & "." & FileExt
End Function

Public Function DoesDirectoryExist(DirPath As String) As Boolean
    DoesDirectoryExist = Dir(DirPath, vbDirectory) <> ""
End Function

Public Function FileExist(fname As String) As Boolean
Dim f As Long
f = FreeFile
On Error GoTo nofile
    Open fname For Input As #f
    Close #f
    FileExist = True
    Exit Function
nofile:
    FileExist = False
End Function
Public Sub MoveFiles()
'Create Desktop Directory
'Copy The Files
'Delete The old Files
End Sub
