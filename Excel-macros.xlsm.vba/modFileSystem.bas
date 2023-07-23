Attribute VB_Name = "modFileSystem"
Option Explicit
' TODO error handling
Public Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim txt_file As Integer
    ' Determine the next available file number to be used by the FileOpen function
    txt_file = FreeFile
    
    Open filePath For Output As txt_file
    Print #txt_file, content
    Close txt_file
End Sub

Public Sub AppendTextFile(ByVal filePath As String, ByVal content As String)
    Dim txt_file As Integer
    ' Determine the next available file number to be used by the FileOpen function
    txt_file = FreeFile
    
    Open filePath For Append As txt_file
    Print #txt_file, content
    Close txt_file
End Sub

Public Function ReadTextFile(ByVal filePath As String) As String
    Dim txt_file As Integer
    Dim content As String

    ' Determine the next available file number to be used by the FileOpen function
    txt_file = FreeFile
    
    Open filePath For Input As txt_file
    content = Input(LOF(txt_file), txt_file)
    Close txt_file
    
    ' Return
    ReadTextFile = content
End Function

