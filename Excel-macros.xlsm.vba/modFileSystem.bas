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


Public Function FolderExists(ByVal fullPath As String) As Boolean
    If fullPath = vbNullString Then
        FolderExists = False
        Exit Function
    End If
    
    If Dir(fullPath, vbDirectory) = vbNullString Then
        FolderExists = False
        Exit Function
    End If
    
    FolderExists = True
End Function


Public Function FileExists(ByVal fullPath As String) As Boolean
    If fullPath = vbNullString Then
        FileExists = False
        Exit Function
    End If
    
    If Dir(fullPath) = vbNullString Then
        FileExists = False
        Exit Function
    End If
    
    FileExists = True
End Function

Public Sub DeleteFile(ByVal fullPath As String)
    If Not FileExists(fullPath) Then
        Exit Sub
    End If
    
    Call Kill(fullPath)
End Sub

Public Sub DeleteFolder(ByVal fullPath As String)
    If Not FolderExists(fullPath) Then
        Exit Sub
    End If

    Call RmDir(fullPath)
End Sub


Public Function ListFolder(ByVal sPath As String, Optional ByVal sFilter As String) As Collection

    Dim result As New Collection

    Dim sFile As String
    Dim nCounter As Long

    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If

    If sFilter = "" Then
        sFilter = "*.*"
    End If

    'call with path "initializes" the dir function and returns the first file
    sFile = Dir(sPath & sFilter, vbDirectory)

    'call it until there is no filename returned
    Do While sFile <> ""
        
        If sFile <> "." And sFile <> ".." Then
            Call result.Add(sFile)
        End If
        
        'subsequent calls without param return next file
        sFile = Dir
    Loop

    'return the array of file names
    Set ListFolder = result

End Function


' See https://learn.microsoft.com/en-us/office/vba/api/excel.application.getsaveasfilename for more details about the filter values
Public Function GetNewFilePathFromUser(Optional ByVal initialName As String, Optional ByVal filter As String = "Text Files (*.txt), *.txt") As String
    Dim result As Variant
    result = Application.GetSaveAsFilename(InitialFileName:=initialName, fileFilter:=filter)
    
    If result = False Then
        GetNewFilePathFromUser = ""
        Exit Function
    End If
    
    GetNewFilePathFromUser = result
End Function

Public Function GetExistingFilePathFromUser(Optional ByVal windowTitle As String) As String
    Dim result As Variant
    result = Application.GetOpenFilename(title:=windowTitle)
    
    If result = False Then
        GetExistingFilePathFromUser = ""
        Exit Function
    End If
    
    GetExistingFilePathFromUser = result
End Function
