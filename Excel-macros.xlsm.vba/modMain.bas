Attribute VB_Name = "modMain"
Option Explicit ' Forces all variables to be defined

' Comments start with a quote

' Each sub is considered a macro and it can be assigned to a button/ui element
' It doesn't return any value
Public Sub Main()
    ' Press F5 to run the code from the current sub

    ' View -> Immediate Window to see the results
    Debug.Print ("Hello World")
    ' Select all and Delete to clear the immediate window
    
    ' Strings use double quotes, use & for concatenation
    Debug.Print ("4 + 7 = " & MyAdd(4, 7))

    modLogger.ActiveLogLevel = modLogger.LL_WARN

    ' Call modLogger.LogMsg("Stop the program", LL_FATAL)
    Call modLogger.LogMsg("Hello", LL_DEBUG)
    Call modLogger.LogMsg("Hello", LL_INFO)
    Call modLogger.LogMsg("Hello", LL_WARN)
    Call modLogger.LogMsg("Hello", LL_ERROR)


    ' using call makes it clear it's a sub
    Call DeclareVariables
    Call Operators
    Call EarlyReturn
    Call HandleErrors
    Call ControlStructures
    Call CollectionAndDictionary
    Call WorkingWithFiles
    Call EnvironmentVariables

    ' For 3rd party libraries see https://github.com/sancarn/awesome-vba

End Sub

' functions return values, use ByVal until you need ByRef
' functions can be used in regular excel cells
Public Function MyAdd(ByVal a As Long, ByVal b As Long) As Long
    Dim result As Long
    
    result = a + b

    ' Return uses the name of the function
    MyAdd = result
End Function


Private Sub DeclareVariables()
    Debug.Print ("---- DECLARE VARIABLES ----")
    Dim x As Long
    x = 42
    
    Const PI = 3.1415926535
    
    ' intialized as false
    Dim flag As Boolean
    flag = True ' or False

    Debug.Print (x & " " & flag & " " & PI)

    Dim p1 As clsPerson
    ' For objects use Set
    Set p1 = modFactory.CreatePerson("Mihai", "Vilcu", 1990)
    
    Debug.Print (p1.FullName & " is " & p1.Age & " years old")

    ' For all data types see https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary

End Sub

Private Sub Operators()
    Debug.Print ("---- OPERATORS ----")
    
    Debug.Print ("7 = 10 " & (7 = 10)) ' equal
    Debug.Print ("7 <> 10 " & (7 <> 10)) ' not equal
    Debug.Print ("Not 7 = 10 " & (Not 7 = 10)) ' negation
    Debug.Print ("7 < 10 " & (7 < 10))
    Debug.Print ("7 > 10 " & (7 > 10))
    Debug.Print ("7 <= 10 " & (7 <= 10))
    Debug.Print ("7 >= 10 " & (7 >= 10))

    Debug.Print ("7 + 12 " & (7 + 12))
    Debug.Print ("7 - 12 " & (7 - 12))
    Debug.Print ("7 * 12 " & (7 * 12))
    Debug.Print ("7 / 12 " & (7 / 12))
    Debug.Print ("70 \ 12 " & (70 \ 12)) ' integer division
    Debug.Print ("7 Mod 12 " & (7 Mod 12))
    Debug.Print ("2 ^ 4 " & (2 ^ 4)) ' exponential
    Debug.Print ("-3 " & (-3)) ' negation

    Debug.Print ("7 < 12 And 12 > 40 " & (7 < 12 And 12 > 40))
    Debug.Print ("7 < 12 Or 12 > 40 " & (7 < 12 Or 12 > 40))

    Debug.Print ("7 < 12 Or DoesNotHaveShortCircuit() " & (7 < 12 Or DoesNotHaveShortCircuit())) ' no short circuit, use guard clauses instead (see EarlyReturn)

End Sub

Private Function DoesNotHaveShortCircuit()
    Debug.Print ("DoesNotHaveShortCircuit called")
    DoesNotHaveShortCircuit = True
End Function

Private Sub EarlyReturn()
    Debug.Print ("---- EARY RETURN/GUARD CLAUSES ----")

    Dim n As Integer
    n = modRandom.RandInt(1, 50)

    If n < 25 Then
        Debug.Print ("Early return for " & n)
        Exit Sub ' Works for functions as well
    End If

    Debug.Print ("No early return for " & n)
End Sub

Private Sub HandleErrors()
    Debug.Print ("---- HANDLE ERRORS ----")
    On Error GoTo ProcessOnError
    
    Dim p2 As clsPerson
    ' For objects use Set, you can pass parameters by name
    Set p2 = modFactory.CreatePerson(FirstName:="John", LastName:="Doe", Yob:=-6)
    Call p2.PrintInfo
    
    Exit Sub

ProcessOnError:
    Call MsgBox("modMain: Number = " & err.Number & " " & err.Description)
End Sub

Private Sub ControlStructures()
    Debug.Print ("---- CONTROL STRUCTURES ----")
    
    Dim n As Integer
    n = modRandom.RandInt(1, 500)
    
    If n <> 3 Then ' not equal
        Debug.Print ("n is not 3")
    Else
        Debug.Print ("n is something else " & n)
    End If
    
    Dim i As Integer
    For i = 0 To 6 Step 2
        Debug.Print ("i = " & i) ' 0, 2, 4, 6
    Next
    
    ' List of primitives
    Dim myList As New Collection
    myList.Add ("A")
    myList.Add ("B")
    myList.Add ("C")
    
    Dim item As Variant
    For Each item In myList
        Debug.Print ("item = " & item)
    Next
    
    ' List of objects
    Dim allPersons As New Collection
    Call allPersons.Add(modFactory.CreatePerson("John", "Doe", 1900))
    Call allPersons.Add(modFactory.CreatePerson("Michael", "Smith", 1950))
    Call allPersons.Add(modFactory.CreatePerson("Maria", "Doe", 1980))

    Dim currentPerson As clsPerson
    For Each currentPerson In allPersons
        Call currentPerson.PrintInfo
    Next
    
End Sub

Private Sub CollectionAndDictionary()
    Debug.Print ("---- COLLECTION AND DICTIONARY ----")
    
    Dim nums As New Collection

    Call nums.Add(100)
    Call nums.Add(150)
    Call nums.Add(200)

    Dim n As Variant
    For Each n In nums
        Debug.Print ("n = " & n)
    Next

    ' Need to add microsoft scripting runtime in references
    ' For cross platform support use https://github.com/VBA-tools/VBA-Dictionary
    Dim codeToName As New Dictionary

    Call codeToName.Add(100, "INFO")
    Call codeToName.Add(200, "OK")
    Call codeToName.Add(300, "REDIRECT")
    Call codeToName.Add(400, "CLIENT ERROR")
    Call codeToName.Add(500, "SERVER ERROR")
    
    If codeToName.Exists(200) Then
        Call codeToName.Add(202, "CREATED")
    End If

    Dim key As Variant
    For Each key In codeToName
        Debug.Print ("key = " & key & " = " & codeToName.item(key))
    Next
    
    
End Sub


Private Sub WorkingWithFiles()
    Debug.Print ("---- FILES AND FOLDERS ----")
    Call modFileSystem.WriteTextFile("C:\tmp\result.txt", "Hello from VBA")
    Call modFileSystem.WriteTextFile("C:\tmp\result.txt", "Hello from VBA2") ' overwrite the existing content
    
    Call modFileSystem.AppendTextFile("C:\tmp\result.txt", "This text is appended") ' new line is added automatically
    Call modFileSystem.AppendTextFile("C:\tmp\result.txt", "More appended text")
    
    Dim content As String
    content = modFileSystem.ReadTextFile("C:\tmp\result.txt")
    Debug.Print (content)
    
    
    If Not modFileSystem.FolderExists("C:\tmp\vba") Then
        Call MkDir("C:\tmp\vba")
        Debug.Print ("Created C:\tmp\vba")
    End If
    
    If Not modFileSystem.FolderExists("C:\tmp\vba\sub") Then
        Call MkDir("C:\tmp\vba\sub")
        Debug.Print ("Created C:\tmp\vba\sub")
    End If
    
    If Not modFileSystem.FolderExists("C:\tmp\vba\sub\folder") Then
        Call MkDir("C:\tmp\vba\sub\folder")
        Debug.Print ("Created C:\tmp\vba\sub\folder")
    End If

    Call modFileSystem.WriteTextFile("C:\tmp\vba\sub\folder\temp.txt", "Temp file")
    Debug.Print ("Created C:\tmp\vba\sub\folder\temp.txt")
    
    Call modFileSystem.Rename("C:\tmp\vba\sub\folder\temp.txt", "C:\tmp\vba\sub\folder\temp-renamed.txt")
    Debug.Print ("Renamed C:\tmp\vba\sub\folder\temp.txt -> C:\tmp\vba\sub\folder\temp-renamed.txt")

    Call modFileSystem.DeleteFile("C:\tmp\vba\sub\folder\temp-renamed.txt")
    Debug.Print ("Deleted C:\tmp\vba\sub\folder\temp-renamed.txt")

    ' Folder should be empty before delete
    Call modFileSystem.DeleteFolder("C:\tmp\vba\sub\folder")
    Debug.Print ("Deleted C:\tmp\vba\sub\folder")

    Dim resultPath As String
    resultPath = "C:\tmp\vba\calc-result.txt"

    If modFileSystem.FileExists(resultPath) Then
        Call modFileSystem.AppendTextFile(resultPath, "Adding to existing results")
        Debug.Print ("Added information to exsting file")
    Else
        Call modFileSystem.WriteTextFile(resultPath, "First line of the file")
        Debug.Print ("Created result file")
    End If

    Dim folderContent As Collection
    Set folderContent = modFileSystem.ListFolder("C:\tmp\vba", "*") ' use *.txt to filter by file type

    Debug.Print ("Listing C:\tmp\vba, found " & folderContent.Count & " items")
    Dim item As Variant
    For Each item In folderContent
        Debug.Print ("    " & item)
    Next
        
    Dim selectedFile As String
    selectedFile = modFileSystem.GetNewFilePathFromUser("report.txt")
    Debug.Print ("New file " & selectedFile)
    
    selectedFile = modFileSystem.GetExistingFilePathFromUser("Select report")
    Debug.Print ("Existing file " & selectedFile)

    selectedFile = modFileSystem.GetFolderPathFromUser("Select working folder")
    Debug.Print ("Folder " & selectedFile)


End Sub

Private Sub EnvironmentVariables()
    Dim homeFolder As String
    homeFolder = Environ("USERPROFILE")
    Debug.Print ("home = " & homeFolder)
End Sub

