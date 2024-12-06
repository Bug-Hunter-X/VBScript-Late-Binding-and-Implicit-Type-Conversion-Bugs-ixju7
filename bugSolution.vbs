Option Explicit

'Early Binding for better performance and error checking
Dim objFSO As Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Explicit type declaration
Dim strFilePath As String
strFilePath = "C:\\test.txt"

'Check if file exists before attempting operations
If objFSO.FileExists(strFilePath) Then
  'Explicit type conversion
  Dim intFileSize As Long
  intFileSize = CLng(objFSO.GetFile(strFilePath).Size)
  MsgBox "File size: " & intFileSize & " bytes"
Else
  MsgBox "File not found!"
End If

Set objFSO = Nothing