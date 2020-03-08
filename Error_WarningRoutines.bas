Attribute VB_Name = "Error_WarningRoutines"
Option Explicit
'@Folder "Common"
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
' 10/29/19
'   Added ability get message folder path

' This module provides the error handling routines
' See the example usage at the end of the module

Private Const Module_Name As String = "ErrorRoutines."

Private Const pErrorStreamName As String = "Error Messages"
Private Const pWarningStreamName As String = "Warning Messages"

Private Type ErrorWarningType
    ErrorFile As MessageFileClass
    WarningFile As MessageFileClass
    SourceOfError As String
End Type

Private This As ErrorWarningType

Public Property Get WarningMessageFolderPath()
    WarningMessageFolderPath = This.WarningFile.MessageFolderPath
End Property

Public Sub ReportWarning( _
       ByVal WarningMsg As String, _
       ParamArray Args() As Variant)

    ' This routine writes a warning message to the warning file
    
    Const RoutineName As String = Module_Name & "ReportWarning"
    On Error GoTo ErrorHandler
    
    If This.WarningFile Is Nothing Then
        Set This.WarningFile = New MessageFileClass
    End If
    
    Dim WarningMessage As String
    WarningMessage = "Non-Fatal Warning" & vbCrLf & WarningMsg & vbCrLf
    
    Dim I As Long
    For I = 0 To IIf(UBound(Args, 1) Mod 2 = 0, UBound(Args, 1) - 2, UBound(Args, 1) - 1) Step 2
        WarningMessage = WarningMessage & Args(I) & " = " & Args(I + 1) & vbCrLf
    Next I

    This.WarningFile.WriteMessageLine WarningMessage, pWarningStreamName
    
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub      ' ReportWarning

Public Property Get ErrorsFound() As Boolean
    ErrorsFound = Not This.ErrorFile Is Nothing
End Property

Public Sub RaiseError( _
       ByVal ErrorNo As Long, _
       ByVal Src As String, _
       ByVal proc As String, _
       ByVal Desc As String, _
       ParamArray Args() As Variant)

    ' https://excelmacromastery.com/vba-error-handling/
    ' Reraises an error and adds line number and current procedure name
    ' Adds a list of parameter names and corresponding parameter values
    ' One name and value per line

    ' Add procedure to source
    This.SourceOfError = This.SourceOfError & vbCrLf & proc
    ReportError This.SourceOfError
    
    ' Check if procedure where error occurs has line numbers
    ' Add error line number if present
'    If Erl <> 0 Then
'        This.SourceOfError = vbCrLf & "Line no: " & Erl
'    End If
'
'    Dim I As Long
'    For I = 1 To IIf(UBound(Args, 1) Mod 2 = 2, UBound(Args, 1), UBound(Args, 1) - 1) Step 2
'        This.SourceOfError = This.SourceOfError & Args(I) & " = " & Args(I + 1) & vbCrLf
'    Next I

    ' If the code stops here,
    ' make sure DisplayError is placed in the top most Sub
    Err.Raise ErrorNo, This.SourceOfError, Desc

End Sub      ' RaiseError

Public Sub DisplayError(ByVal ProcName As String)

    ' Writes the error to the error file when it reaches the topmost sub

    Dim Msg As String
    Msg = "The following exception was raised: " & vbCrLf & _
          "Description: " & Err.Description & vbCrLf & _
          "VBA Project: " & ThisWorkbook.VBProject.Name & vbCrLf & _
          This.SourceOfError & vbCrLf & ProcName

    ReportError Msg
    
End Sub      ' DisplayError

Public Sub ReportError( _
       ByVal ErrMsg As String, _
       ParamArray Args() As Variant)

    ' This routine writes an error message to the error file
    
    Const RoutineName As String = Module_Name & "ReportError"
    On Error GoTo ErrorHandler
    
    If This.ErrorFile Is Nothing Then
        Set This.ErrorFile = New MessageFileClass
    End If
    
    Dim ErrorMessage As String
    ErrorMessage = "Fatal Error Message" & vbCrLf & ErrMsg & vbCrLf
    
    Dim I As Long
    For I = 0 To IIf(UBound(Args, 1) Mod 2 = 0, UBound(Args, 1) - 2, UBound(Args, 1) - 1) Step 2
        ErrorMessage = ErrorMessage & Args(I) & " = " & Args(I + 1) & vbCrLf
    Next I

    This.ErrorFile.WriteMessageLine ErrorMessage, pErrorStreamName
    
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub      ' ReportError

Public Sub CloseErrorFile()
    ' Declared as a function to keep it of the Alt-F8 list of executable routines

    Set This.ErrorFile = Nothing
    Set This.WarningFile = Nothing
    This.SourceOfError = vbNullString

End Sub      ' CloseErrorFile

