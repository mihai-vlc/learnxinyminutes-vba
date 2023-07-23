Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private WithEvents activeUser As clsPerson
Attribute activeUser.VB_VarHelpID = -1

Private Sub Worksheet_Activate()
    Set activeUser = modFactory.CreatePerson("John", "Doe", 1990)
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    activeUser.Yob = modRandom.RandInt(1990, 2010)
End Sub

Private Sub activeUser_YobChanged(ByVal newValue As Integer)
    Debug.Print ("Yob changed to " & newValue)
End Sub
