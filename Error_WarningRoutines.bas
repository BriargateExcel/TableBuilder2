Attribute VB_Name = "Error_WarningRoutines"
Option Explicit
' Changes
' 9/15/19
'       Changed from ErrorFilesClass to ErrorFileClass
' 9/20/19
'   Changed RaiseError and DisplayError to functions
'       so they don't show up in the list of executable routines
'   In DisplayError changed ActiveWorkbook to ThisWorkbook
' 9/22/19
'   Changed ReportError to function
' 10/6/19
'   Deleted performance/debug code in ReportError
'   Added ReportWarning
'   Changed module name to Error_WarningRoutines
'   Deleted reference to CloseErrorFile

' This module provides the error handling routines
' See the example usage at the end of the module

Private Const Module_Name As String = "Error_WarningRoutines."

Private pErrorFile As MessageFileClass
Private Const pErrorStreamName As String = "Error Messages"

Private pWarningFile As MessageFileClass
Private Const pWarningStreamName As String = "Warning Messages"

Private SourceOfError As String

Public Function ReportWarning( _
       ByVal WarningMsg As String, _
       ParamArray Args() As Variant)

    ' This routine writes a warning message to the warning file
    
    Const RoutineName As String = Module_Name & "ReportWarning"
    On Error GoTo ErrorHandler
    
    If pWarningFile Is Nothing Then
        Set pWarningFile = New MessageFileClass
    End If
    
    Dim WarningMessage As String
    WarningMessage = "Non-Fatal Warning" & vbCrLf & WarningMsg & vbCrLf
    
    Dim I As Long
    For I = 0 To IIf(UBound(Args, 1) Mod 2 = 0, UBound(Args, 1) - 2, UBound(Args, 1) - 1) Step 2
        WarningMessage = WarningMessage & Args(I) & " = " & Args(I + 1) & vbCrLf
    Next I

    pWarningFile.WriteMessageLine WarningMessage, pWarningStreamName
    
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' ReportWarning

Public Property Get ErrorsFound() As Boolean
    ErrorsFound = Not pErrorFile Is Nothing
End Property

Public Function RaiseError( _
       ByVal ErrorNo As Long, _
       ByVal Src As String, _
       ByVal Proc As String, _
       ByVal Desc As String, _
       ParamArray Args() As Variant)

    ' https://excelmacromastery.com/vba-error-handling/
    ' Reraises an error and adds line number and current procedure name
    ' Adds a list of parameter names and corresponding parameter values
    ' One name and value per line

    ' Add procedure to source
    SourceOfError = SourceOfError & vbCrLf & Proc
    ReportError SourceOfError
    
    ' Check if procedure where error occurs has line numbers
    ' Add error line number if present
    '    If Erl <> 0 Then
    '        SourceOfError = vbCrLf & "Line no: " & Erl
    '    End If
    '
    '    Dim I As Long
    '    For I = 1 To IIf(UBound(Args, 1) Mod 2 = 2, UBound(Args, 1), UBound(Args, 1) - 1) Step 2
    '        SourceOfError = SourceOfError & Args(I) & " = " & Args(I + 1) & vbCrLf
    '    Next I

    ' If the code stops here,
    ' make sure DisplayError is placed in the top most Sub
    Err.Raise ErrorNo, SourceOfError, Desc

End Function ' RaiseError

Public Function DisplayError(ByVal ProcName As String)

    ' Writes the error to the error file when it reaches the topmost sub

    Dim Msg As String
    Dim ErrorNumber As Long
    On Error Resume Next
    Msg = "The following exception was raised: " & vbCrLf & _
          "Description: " & Err.Description & vbCrLf & _
          "VBA Project: " & ThisWorkbook.VBProject.Name & vbCrLf & _
          SourceOfError & vbCrLf & ProcName
    ErrorNumber = Err.Number
    On Error GoTo 0
    If ErrorNumber <> 0 Then
    Msg = "The following exception was raised: " & vbCrLf & _
          "Description: " & Err.Description & vbCrLf & _
          SourceOfError & vbCrLf & ProcName
    End If

    ReportError Msg
    
End Function ' DisplayError

Public Function ReportError( _
       ByVal ErrMsg As String, _
       ParamArray Args() As Variant)

    ' This routine writes an error message to the error file
    
    Const RoutineName As String = Module_Name & "ReportError"
    On Error GoTo ErrorHandler
    
    If pErrorFile Is Nothing Then
        Set pErrorFile = New MessageFileClass
    End If
    
    Dim ErrorMessage As String
    ErrorMessage = "Fatal Error Message" & vbCrLf & ErrMsg & vbCrLf
    
    Dim I As Long
    For I = 0 To IIf(UBound(Args, 1) Mod 2 = 0, UBound(Args, 1) - 2, UBound(Args, 1) - 1) Step 2
        ErrorMessage = ErrorMessage & Args(I) & " = " & Args(I + 1) & vbCrLf
    Next I

    pErrorFile.WriteMessageLine ErrorMessage, pErrorStreamName
    
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' ReportError

Public Function CloseErrorFile() As Boolean
    ' Declared as a function to keep it of the Alt-F8 list of executable routines

    Set pErrorFile = Nothing
    Set pWarningFile = Nothing
    SourceOfError = vbNullString

End Function ' CloseErrorFile

'Private Sub SubErrorRaiseProcess(ByVal Parameter As String)
'
'    ' This routine tests the error raise process
'
'    Const RoutineName As String = Module_Name & "SubErrorRaiseProcess"
'    On Error GoTo ErrorHandler
'
'    ReportError "Error Message", _
'        "Parameter", Parameter, _
'        "Param 1", 1, _
'        "Param 2", 2
'
'    Dim Test As Long
'    Test = 1 / 0
'
'Done:
'    Exit Sub
'ErrorHandler:
'    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
'End Sub ' SubErrorRaiseProcess
'
'Public Sub TestErrorRaiseProcess()
'
'    ' This routine tests the error raise process
'
'    Const RoutineName As String = Module_Name & "TestErrorRaiseProcess"
'    On Error GoTo ErrorHandler
'
'    SubErrorRaiseProcess "First"
'    SubErrorRaiseProcess "Second"
'
'    CloseErrorFile
'
'Done:
'    CloseErrorFile
'    Exit Sub
'ErrorHandler:
'    DisplayError RoutineName
'    CloseErrorFile
'End Sub ' TestErrorRaiseProcess











