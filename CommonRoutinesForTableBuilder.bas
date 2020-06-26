Attribute VB_Name = "CommonRoutinesForTableBuilder"
Option Explicit
'@Folder "Common"
Private Const Module_Name As String = "CommonRoutinesForTableBuilder."

Private Const CommandBarName As String = "Table Builder"

Private Type CodeType
    FormCanceled As Boolean
'    FormDeleted As Boolean
    
    ModuleList As Dictionary
    ModuleTable As ListObject
    
    PathFolder As Dictionary
    PathTable As ListObject
    Path As String
    
    ReferencesList As Dictionary
    ReferencesTable As ListObject
    
    Project As VBProject
    
    Workbook As Workbook
    Worksheet As Worksheet
    
    CustomBar As CommandBar
End Type

Private This As CodeType

Public Sub Auto_Open()
' https://bettersolutions.com/vba/ribbon/face-ids-2003.htm for FaceIDs
    
    Dim NewButton As CommandBarButton
    
    On Error Resume Next
    CommandBars(CommandBarName).Delete
    On Error GoTo 0
    
    Set This.CustomBar = CommandBars.Add(Name:=CommandBarName)
    
    BuildButton "BuildModules", "Build Modules", 81
    
    This.CustomBar.Visible = True
    
End Sub ' Auto_Open

Private Sub BuildButton( _
    ByVal RoutineToExecute As String, _
    ByVal Caption As String, _
    ByVal FaceID As Long)

    ' Build one button on the command bar
    
    Const RoutineName As String = Module_Name & "BuildButton"
    On Error GoTo ErrorHandler
    
    Dim NewButton As CommandBarButton
    
    Set NewButton = This.CustomBar.Controls.Add(msoControlButton)
    NewButton.OnAction = RoutineToExecute
    NewButton.Caption = Caption
    NewButton.FaceID = FaceID
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildButton

Public Sub Auto_Close()
    On Error Resume Next
    CommandBars(CommandBarName).Delete
    On Error GoTo 0
End Sub ' Auto_Close

Public Function TryGetFile( _
    ByVal Path As String, _
    ByRef Contents As String _
    ) As Boolean

       ' Checks to see if there is a file
       ' Returns the contents if it exists
    
    Const RoutineName As String = Module_Name & "TryGetFile"
    On Error GoTo ErrorHandler
    
    TryGetFile = True
    
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    If FSO.FileExists(Path) Then
        Dim Stream As Scripting.TextStream
        Set Stream = FSO.OpenTextFile(Path, ForReading)
        
        Contents = Stream.ReadAll()
        
        Stream.Close
    Else
        TryGetFile = False
        GoTo Done
    End If
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryGetFile

Public Sub BuildApplicationUniqueDeclarations( _
    ByVal StreamFile As MessageFileClass, _
    ByVal StreamName As String, _
    ByVal TableName As String, _
    ByVal FileExtension As String)

    ' Adds any application unique declarations
    
    Const RoutineName As String = Module_Name & "BuildApplicationUniqueDeclarations"
    On Error GoTo ErrorHandler
    
    Dim Path As String
    Path = DesktopFolder & Application.PathSeparator & "Modules" & Application.PathSeparator & _
        "Application_Unique_Code" & Application.PathSeparator & TableName & "Declarations" & FileExtension
    
    CopyFile StreamFile, StreamName, Path, "' No application specific declarations found" & vbCrLf
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildApplicationUniqueDeclarations

Public Sub BuildApplicationUniqueRoutines( _
    ByVal StreamFile As MessageFileClass, _
    ByVal StreamName As String, _
    ByVal TableName As String, _
    ByVal FileExtension As String)

    ' Adds any application unique code
    
    Const RoutineName As String = Module_Name & "BuildApplicationUniqueRoutines"
    On Error GoTo ErrorHandler
    
    Dim Path As String
    Path = DesktopFolder & Application.PathSeparator & "Modules" & Application.PathSeparator & _
        "Application_Unique_Code" & Application.PathSeparator & TableName & FileExtension
    
    CopyFile StreamFile, StreamName, Path, "' No application unique routines found" & vbCrLf
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildApplicationUniqueDeclarations

Public Sub CopyFile( _
    ByVal StreamFile As MessageFileClass, _
    ByVal StreamName As String, _
    ByVal Path As String, _
    ByVal NothingFoundMessage As String)

    ' Adds any application unique information
    
    Const RoutineName As String = Module_Name & "CopyFile"
    On Error GoTo ErrorHandler
    
    Dim Contents As String
    Dim Line As String
    
    If TryGetFile(Path, Contents) Then
        '
        ' Code separator
        '
        
        Line = _
            "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
            "'                                                  '" & vbCrLf & _
            "'   Start of application specific code             '" & vbCrLf & _
            "'                                                  '" & vbCrLf & _
            "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        StreamFile.WriteMessageLine Line, StreamName
    
        StreamFile.WriteMessageLine Contents, StreamName
        '
        ' Code separator
        '
        
        Line = _
            "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
            "'                                                  '" & vbCrLf & _
            "'    End of application specific code              '" & vbCrLf & _
            "'                                                  '" & vbCrLf & _
            "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        StreamFile.WriteMessageLine Line, StreamName
    Else
        Line = NothingFoundMessage
        StreamFile.WriteMessageLine Line, StreamName
        GoTo Done
    End If
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' CopyFile
