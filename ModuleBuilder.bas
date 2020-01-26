Attribute VB_Name = "ModuleBuilder"
Option Explicit

Private Const Module_Name As String = "ModuleBuilder."

Private Const Quote As String = """"

Public Sub ModuleBuilder( _
    ByVal DetailsDict As Dictionary, _
    ByVal TableName As String, _
    ByVal ClassName As String)

    ' This routine builds the basic module

    Const RoutineName As String = Module_Name & "ClassBuilder"
    On Error GoTo ErrorHandler
    
    Dim StreamName As String
    StreamName = TableName & ".bas"
    
    Dim StreamFile As MessageFileClass
    Set StreamFile = New MessageFileClass
    
    Dim Line As String
    
    '
    ' Declarations and public column properties
    '
    
    Line = PrintString( _
        "Attribute VB_Name = qq%1qq" & vbCrLf & _
        "Option Explicit" & vbCrLf & _
        vbCrLf & _
        "' Built on " & Now() & vbCrLf & _
        "' Built By Briargate Excel Table Builder" & vbCrLf & _
        "' See BriargateExcel.com for details" & vbCrLf & _
        vbCrLf & _
        "Private Const Module_Name As String = qq%1.qq" & vbCrLf & vbCrLf & _
        "Private pInitialized As Boolean", _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName, "Modules", True

    Line = PrintString( _
        "Private p%1Dict As Dictionary" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "'   Start of application specific declarations     '" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
        vbCrLf & _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "'    End of application specific declarations      '" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    BuildConstantsAndProperties StreamFile, StreamName, DetailsDict, TableName

    '
    ' Get Dictionary
    '
    Line = PrintString( _
        "Public Property Get %1Dictionary() As Dictionary" & vbCrLf & _
        "   Set %1Dictionary = p%1Dict" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' Get Initialized
    '
    
    Line = PrintString( _
        "Public Property Get %1Initialized() As Boolean" & vbCrLf & _
        "   %1Initialized = pInitialized" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' Reset
    '
    
    Line = PrintString( _
        "Public Sub %1Reset()" & vbCrLf & _
        "    pInitialized = False" & vbCrLf & _
        "    Set p%1Dict = Nothing" & vbCrLf & _
        "End Sub" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' HeaderWidth
    '
    
    Line = PrintString( _
        "Public Property Get %1HeaderWidth() As Long" & vbCrLf & _
        "    %1HeaderWidth = pHeaderWidth" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' Headers
    '
    
    BuildHeaders StreamFile, StreamName, DetailsDict, TableName

    '
    ' Initialize
    '
    
    Line = PrintString( _
        "Public Sub %1Initialize()" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & qq%1Initializeqq" & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName

    Line = PrintString( _
        "    Dim  %1 As %2" & vbCrLf & _
        "    Set %1 = New %2" & vbCrLf & _
        vbCrLf & _
        "    Set p%1Dict = New Dictionary" & vbCrLf & _
        "    If Table.TryCopyTableToDictionary(%1, %1Table, p%1Dict) Then" & vbCrLf & _
        "        pInitialized = True" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        ReportError qqError copying %1 tableqq, qqRoutineqq, RoutineName" & vbCrLf & _
        "        pInitialized = False" & vbCrLf & _
        "        GoTo Done" & vbCrLf & _
        "    End If" & vbCrLf, _
        TableName, ClassName)
     StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "Done:" & vbCrLf & _
        "    Exit Sub" & vbCrLf & _
        "ErrorHandler:" & vbCrLf & _
        "    ReportError " & Quote & "Exception raised" & Quote & ", _" & vbCrLf & _
        "                " & Quote & "Routine" & Quote & ", RoutineName, _" & vbCrLf & _
        "                " & Quote & "Error Number" & Quote & ", Err.Number, _" & vbCrLf & _
        "                " & Quote & "Error Description" & Quote & ", Err.Description" & vbCrLf & _
        "    RaiseError Err.Number, Err.Source, RoutineName, Err.Description" & vbCrLf & _
        "End Sub ' " & TableName & "Initialize" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' TryCopyDictionaryToArray
    '
    
    Line = PrintString( _
    "Public Function %1TryCopyDictionaryToArray( _" & vbCrLf & _
    "    ByVal Dict As Dictionary, _" & vbCrLf & _
    "    ByRef Ary As Variant _" & vbCrLf & _
    "    ) As Boolean" & vbCrLf & _
    vbCrLf & _
    "    Const RoutineName As String = Module_Name & qq%1TryCopyDictionaryToArrayqq" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    vbCrLf & _
    "    %1TryCopyDictionaryToArray = True" & vbCrLf & _
    vbCrLf & _
    "    If Dict.Count = 0 Then" & vbCrLf & _
    "        ReportError qqError copying %1 dictionary to array,qq, qqRoutineqq, RoutineName" & vbCrLf & _
    "        %1TryCopyDictionaryToArray = False" & vbCrLf & _
    "        GoTo Done" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    Dim I As Long" & vbCrLf & _
    "    I = 1" & vbCrLf, _
    TableName)
    StreamFile.WriteMessageLine Line, StreamName

    RecordToArray StreamFile, StreamName, DetailsDict, ClassName
    
    Line = _
        "Done:" & vbCrLf & _
        "    Exit Function" & vbCrLf & _
        "ErrorHandler:" & vbCrLf & _
        "    ReportError " & Quote & "Exception raised" & Quote & ", _" & vbCrLf & _
        "                " & Quote & "Routine" & Quote & ", RoutineName, _" & vbCrLf & _
        "                " & Quote & "Error Number" & Quote & ", Err.Number, _" & vbCrLf & _
        "                " & Quote & "Error Description" & Quote & ", Err.Description" & vbCrLf & _
        "    RaiseError Err.Number, Err.Source, RoutineName, Err.Description" & vbCrLf & _
        "End Function ' " & TableName & "TryCopyDictionaryToArray" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "'         The routines that follow may need        '" & vbCrLf & _
        "'        changes depending on the application      '" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' Get Table
    '
    
    Line = PrintString( _
        "Public Property Get %1Table() As ListObject" & vbCrLf & _
        vbCrLf & _
        "    ' Change the table reference if the table is in another workbook" & vbCrLf & _
        vbCrLf & _
        "    Set %1Table = %1Sheet.ListObjects(qq%1Tableqq)" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' TryCopyArrayToDictionary
    '
    
    Line = PrintString( _
        "Public Function %1TryCopyArrayToDictionary( _" & vbCrLf & _
        "       ByVal Ary As Variant, _" & vbCrLf & _
        "       ByRef Dict As Dictionary)" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & qq%1TryCopyArrayToDictionaryqq" & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    %1TryCopyArrayToDictionary = True" & vbCrLf & _
        vbCrLf & _
        "    Dim I As Long" & vbCrLf & _
        vbCrLf & _
        "    Set Dict = New Dictionary" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName

    ArrayToRecord StreamFile, StreamName, DetailsDict, ClassName, TableName

    Line = _
        "Done:" & vbCrLf & _
        "    Exit Function" & vbCrLf & _
        "ErrorHandler:" & vbCrLf & _
        "    ReportError " & Quote & "Exception raised" & Quote & ", _" & vbCrLf & _
        "                " & Quote & "Routine" & Quote & ", RoutineName, _" & vbCrLf & _
        "                " & Quote & "Error Number" & Quote & ", Err.Number, _" & vbCrLf & _
        "                " & Quote & "Error Description" & Quote & ", Err.Description" & vbCrLf & _
        "    RaiseError Err.Number, Err.Source, RoutineName, Err.Description" & vbCrLf & _
        "End Function ' " & TableName & "TryCopyArrayToDictionary" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' FormatArrayAndWorksheet
    '
    
    Line = PrintString( _
        "Public Sub %1FormatArrayAndWorksheet( _" & vbCrLf & _
        "    ByRef Ary as Variant, _" & vbCrLf & _
        "    ByVal Table As ListObject)" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & qq%1FormatArrayAndWorksheetqq" & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    ' Array and Table formatting goes here" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "Done:" & vbCrLf & _
        "    Exit Sub" & vbCrLf & _
        "ErrorHandler:" & vbCrLf & _
        "    ReportError " & Quote & "Exception raised" & Quote & ", _" & vbCrLf & _
        "                " & Quote & "Routine" & Quote & ", RoutineName, _" & vbCrLf & _
        "                " & Quote & "Error Number" & Quote & ", Err.Number, _" & vbCrLf & _
        "                " & Quote & "Error Description" & Quote & ", Err.Description" & vbCrLf & _
        "    RaiseError Err.Number, Err.Source, RoutineName, Err.Description" & vbCrLf & _
        "End Sub ' " & TableName & "FormatArrayAndWorksheet" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' End of generated code comment
    '

    Line = _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "'             End of Generated code                '" & vbCrLf & _
        "'            Start unique code here                '" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Set StreamFile = Nothing

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' ModuleBuilder

Private Sub BuildConstantsAndProperties( _
        ByVal StreamFile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal DetailsDict As Dictionary, _
        ByVal TableName As String)

    ' This routine builds Get and Let Properties
    
    Const RoutineName As String = Module_Name & "BuildConstantsAndProperties"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    Dim Counter As Long
    
    Dim Entry As Variant
    For Each Entry In DetailsDict.Keys
        Counter = Counter + 1
        
        Line = PrintString( _
            "Private Const p%1Column As Long = " & Counter, _
            DetailsDict.Item(Entry).VariableName)
        StreamFile.WriteMessageLine Line, StreamName
    Next Entry
    
    Line = _
        "Private Const pHeaderWidth As Long = " & Counter & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Counter = 0
    For Each Entry In DetailsDict.Keys
        Counter = Counter + 1
        
        Line = PrintString( _
            "Public Property Get %1%2Column() As Long" & vbCrLf & _
            "    %1%2Column = p%2Column" & vbCrLf & _
            "End Property" & vbCrLf, _
            TableName, DetailsDict.Item(Entry).VariableName)
        StreamFile.WriteMessageLine Line, StreamName
    Next Entry
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildConstantsAndProperties

Private Sub RecordToArray( _
        ByVal StreamFile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal DetailsDict As Dictionary, _
        ByVal ClassName As String)

    ' This routine builds an array from details
    
    Const RoutineName As String = Module_Name & "RecordToArray"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "    Dim Record As %1" & vbCrLf & _
        "    Dim Entry As Variant" & vbCrLf & _
        "    For Each Entry In Dict.Keys" & vbCrLf & _
        "        Set Record = Dict.Item(Entry)" & vbCrLf, _
        ClassName)
    StreamFile.WriteMessageLine Line, StreamName
    
    Dim Entry As Variant
    Dim I As Long
    For Each Entry In DetailsDict.Keys
        I = I + 1
        Line = PrintString( _
            "        Ary(I, p%1Column) = Record.%1", _
            DetailsDict.Item(Entry).VariableName)
        StreamFile.WriteMessageLine Line, StreamName
    Next Entry
    
    Line = vbCrLf & "        I = I + 1"
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = "    Next Entry" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' RecordToArray

Private Sub BuildHeaders( _
        ByVal StreamFile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal DetailsDict As Dictionary, _
        ByVal TableName As String)

    ' This routine builds the Headers property
    
    Const RoutineName As String = Module_Name & "RecordToArray"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "Public Property Get %1Headers() As Variant", _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = PrintString( _
        "    %1Headers = Array( _" & vbCrLf & _
        "        qq%2qq, _" & vbCrLf, _
        TableName, DetailsDict.Items(0).ColumnHeader)
        
    Dim I As Long
    For I = 1 To DetailsDict.Count - 2
        Line = PrintString(Line & _
            "        qq%1qq, _" & vbCrLf, _
            DetailsDict.Items(I).ColumnHeader)
    Next I
    
    Line = PrintString(Line & _
        "        qq%1qq)", _
        DetailsDict.Items(DetailsDict.Count - 1).ColumnHeader)
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' RecordToArray

Private Sub ArrayToRecord( _
        ByVal StreamFile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal DetailsDict As Dictionary, _
        ByVal ClassName As String, _
        ByVal TableName As String)

    ' This routine builds a dictionary from an array
    
    Const RoutineName As String = Module_Name & "ArrayToRecord"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    Line = PrintString( _
        "    Dim Key As String" & vbCrLf & _
        "    Dim Record as %1" & vbCrLf & vbCrLf & _
        "    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then" & vbCrLf & _
        "        For I = 1 To UBound(Ary, 1)" & vbCrLf & _
        "            ' May have to change the key to generate unique keys" & vbCrLf & _
        "            Key = Ary(I, p%2Column)" & vbCrLf & vbCrLf & _
        "            If Dict.Exists(Key) Then" & vbCrLf & _
        "                ReportWarning " & Quote & "Duplicate keyqq, qqRoutineqq, RoutineName, qqKeyqq, Key" & vbCrLf & _
        "                " & TableName & "TryCopyArrayToDictionary = False" & vbCrLf & _
        "                GoTo Done" & vbCrLf & _
        "            Else" & vbCrLf & _
        "                Set Record = New " & ClassName & vbCrLf, _
        ClassName, DetailsDict.Items(0).VariableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    Dim Entry As Variant
    For Each Entry In DetailsDict.Keys
        If DetailsDict.Item(Entry).VariableType = "Boolean" Then
            Line = PrintString( _
                "                Record.%1 = IIf(Ary(I, p%1Column) = qqYesqq , True,False)", _
                DetailsDict.Item(Entry).VariableName)
        Else
            Line = PrintString( _
                "                Record.%1 = Ary(I, p%1Column)", _
                DetailsDict.Item(Entry).VariableName)
        End If
        StreamFile.WriteMessageLine Line, StreamName
    Next Entry
    
    StreamFile.WriteBlankMessageLines StreamName
    
    Line = _
        "                Dict.Add Key, Record" & vbCrLf & _
        "            End If" & vbCrLf & _
        "        Next I" & vbCrLf & vbCrLf & _
        "    Else" & vbCrLf & _
        "        Dict.Add Ary, Ary" & vbCrLf & _
        "    End If" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' ArrayToRecord


