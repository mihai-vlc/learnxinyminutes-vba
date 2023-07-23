Attribute VB_Name = "modLogger"
Option Explicit

Public Enum LogLevel
    LL_DEBUG = 0
    LL_INFO = 1
    LL_WARN = 2
    LL_ERROR = 3
    LL_FATAL = 4
End Enum

Public ActiveLogLevel As LogLevel

Public Sub LogMsg(ByVal msg As String, Optional ByVal level As LogLevel)

    If level >= ActiveLogLevel Then
        Debug.Print (GetLogLevelName(level) & " : " & msg)
    End If

    If level = LL_FATAL Then
        End ' End execution of ALL instructions and clear variables
    End If

End Sub

Private Function GetLogLevelName(ByVal level As LogLevel)
    Dim result As String

    Select Case level
        Case LL_DEBUG
            result = "DEBUG"
        Case LL_INFO
            result = "INFO"
        Case LL_WARN
            result = "WARN"
        Case LL_ERROR
            result = "ERROR"
        Case LL_FATAL
            result = "FATAL"
        Case Else
            result = "INVALID"
    End Select
    
    GetLogLevelName = result
End Function

