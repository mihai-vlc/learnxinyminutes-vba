Attribute VB_Name = "clsPerson"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public FirstName As String
Public LastName As String
Private p_yob As Integer
Event YobChanged(ByVal newValue As Integer)



Public Property Get FullName() As String
    ' Me is the reference to the current instance
    FullName = Me.FirstName & " " & Me.LastName
End Property

Public Property Get Age() As Integer
    Age = Year(Now()) - Me.Yob
End Property

Public Property Get Yob() As Integer
    Yob = p_yob
End Property

Public Property Let Yob(ByVal val As Integer)
    If val < 0 Then
        Call err.Raise(vbObjectError + 9000, "person", "Invalid year of birth")
    End If
    p_yob = val
    
    RaiseEvent YobChanged(val)
End Property


Public Sub PrintInfo()
    Debug.Print (Me.FullName & " is " & Me.Age & " years old")
End Sub

