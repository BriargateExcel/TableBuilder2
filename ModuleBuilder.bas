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
    
    Line = _
        "Attribute VB_Name = " & Quote & TableName & Quote & vbCrLf & _
        "Option Explicit" & vbCrLf & _
        vbCrLf & _
        "' Built on " & Now() & vbCrLf & _
        "' Built By Briargate Excel Table Builder" & vbCrLf & _
        "' See BriargateExcel.com for details" & vbCrLf & _
        vbCrLf & _
        "Private Const Module_Name As String = " & Quote & TableName & "." & Quote & vbCrLf & vbCrLf & _
        "Private pInitialized As Boolean"
    StreamFile.WriteMessageLine Line, StreamName, "Modules", True

    Line = _
        "Private p" & TableName & "Dict As Dictionary" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    BuildConstantsAndProperties StreamFile, StreamName, DetailsDict, TableName

    '
    ' Get Table
    '
    
    Line = _
        "Public Property Get " & TableName & "Table() As ListObject" & vbCrLf & _
        "    Set " & TableName & "Table = " & TableName & "Sheet.ListObjects(" _
        & Quote & TableName & "Table" & Quote & ")" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' Get Dictionary
    '
    Line = _
        "Public Property Get " & TableName & "Dictionary() As Dictionary" & vbCrLf & _
        "   Set " & TableName & "Dictionary = p" & TableName & "Dict" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' Get Initialized
    '
    
    Line = _
        "Public Property Get " & TableName & "Initialized() As Boolean" & vbCrLf & _
        "   " & TableName & "Initialized = pInitialized" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' Reset
    '
    
    Line = _
        "Public Sub " & TableName & "Reset()" & vbCrLf & _
        "    pInitialized = False" & vbCrLf & _
        "    Set p" & TableName & "Dict = Nothing" & vbCrLf & _
        "End Sub" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' HeaderWidth
    '
    
    Line = _
        "Public Property Get " & TableName & "HeaderWidth() As Long" & vbCrLf & _
        "    " & TableName & "HeaderWidth = pHeaderWidth" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' Headers
    '
    
    BuildHeaders StreamFile, StreamName, DetailsDict, TableName

    '
    ' Initialize
    '
    
    Line = _
        "Public Sub " & TableName & "Initialize()" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & " & Quote & TableName & "Initialize" & Quote & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    pInitialized = True" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "    Dim  " & TableName & " As " & ClassName & vbCrLf & _
        "    Set " & TableName & " = New " & ClassName & vbCrLf & _
        vbCrLf & _
        "    Set p" & TableName & "Dict = New Dictionary" & vbCrLf & _
        "    If Table.TryCopyTableToDictionary(" & TableName & "," & TableName & "Table, p" & TableName & "Dict) Then" & vbCrLf & _
        "        ' Success; do nothing" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        MsgBox " & Quote & "Error copying " & TableName & " table" & Quote & vbCrLf & _
        "        pInitialized = False" & vbCrLf & _
        "        GoTo Done" & vbCrLf & _
        "    End If" & vbCrLf
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
    
    Line = _
    "Public Function " & TableName & "TryCopyDictionaryToArray( _" & vbCrLf & _
    "    ByVal Dict As Dictionary, _" & vbCrLf & _
    "    ByRef Ary As Variant _" & vbCrLf & _
    "    ) As Boolean" & vbCrLf & _
    vbCrLf & _
    "    Const RoutineName As String = Module_Name & " & Quote & TableName & "TryCopyDictionaryToArray" & Quote & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    vbCrLf & _
    "    " & TableName & "TryCopyDictionaryToArray = True" & vbCrLf & _
    vbCrLf & _
    "    If Dict.Count = 0 Then" & vbCrLf & _
    "        ReportError " & Quote & "Error copying " & TableName & " dictionary to array," & Quote & ", " & Quote & "Routine" & Quote & ", RoutineName" & vbCrLf & _
    "        " & TableName & "TryCopyDictionaryToArray = False" & vbCrLf & _
    "        GoTo Done" & vbCrLf & _
    "    End If" & vbCrLf & _
    "    Dim I As Long" & vbCrLf & _
    "    I = 1" & vbCrLf
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

    '
    ' TryCopyArrayToDictionary
    '
    
    Line = _
        "Public Function " & TableName & "TryCopyArrayToDictionary( _" & vbCrLf & _
        "       ByVal Ary As Variant, _" & vbCrLf & _
        "       ByRef Dict As Dictionary)" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & " & Quote & TableName & "TryCopyArrayToDictionary" & Quote & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    " & TableName & "TryCopyArrayToDictionary = True" & vbCrLf & _
        vbCrLf & _
        "    Dim I As Long" & vbCrLf & _
        vbCrLf & _
        "    Set Dict = New Dictionary" & vbCrLf
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
    ' FormatWorksheet
    '
    
    Line = _
        "Public Sub " & TableName & "FormatWorksheet(ByVal Table As ListObject)" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & " & Quote & TableName & "FormatWorksheet" & Quote & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    ' Worksheet formatting goes here" & vbCrLf
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
        "End Sub ' " & TableName & "FormatWorksheet" & vbCrLf
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
        
        Line = "Private Const p" & DetailsDict.Item(Entry).VariableName & "Column As Long = " & Counter
        StreamFile.WriteMessageLine Line, StreamName
    Next Entry
    
    Line = _
        "Private Const pHeaderWidth As Long = " & Counter & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Counter = 0
    For Each Entry In DetailsDict.Keys
        Counter = Counter + 1
        
        Line = _
            "Public Property Get " & TableName & DetailsDict.Item(Entry).VariableName & "Column() As Long" & vbCrLf & _
            "    " & TableName & DetailsDict.Item(Entry).VariableName & "Column = p" & DetailsDict.Item(Entry).VariableName & "Column" & vbCrLf & _
            "End Property" & vbCrLf
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
    
    Line = _
        "    Dim Record As " & ClassName & vbCrLf & _
        "    Dim Entry As Variant" & vbCrLf & _
        "    For Each Entry In Dict.Keys" & vbCrLf & _
        "        Set Record = Dict.Item(Entry)" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Dim Entry As Variant
    Dim I As Long
    For Each Entry In DetailsDict.Keys
        I = I + 1
        Line = _
            "        Ary(I, p" & DetailsDict.Item(Entry).VariableName & "Column) = Record." & DetailsDict.Item(Entry).VariableName
        StreamFile.WriteMessageLine Line, StreamName
    Next Entry
    
    Line = vbCrLf & "        I = I + 1"
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = "    Next Entry"
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
    
    Line = "Public Property Get " & TableName & "Headers() As Variant"
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "    " & TableName & "Headers = Array(" & Quote & DetailsDict.Items(0).ColumnHeader & Quote
    
    Dim I As Long
    For I = 1 To DetailsDict.Count - 1
        Line = Line & ", " & Quote & DetailsDict.Items(I).ColumnHeader & Quote
    Next I
    
    Line = Line & ")"
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
    Line = _
        "    Dim Key As String" & vbCrLf & _
        "    Dim Record as " & ClassName & vbCrLf & vbCrLf & _
        "    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then" & vbCrLf & _
        "        For I = 1 To UBound(Ary, 1)" & vbCrLf & _
        "            Key = Ary(I, p" & DetailsDict.Items(0).VariableName & "Column)' May have to change this to generate unique keys" & vbCrLf & vbCrLf & _
        "            If Dict.Exists(Key) Then" & vbCrLf & _
        "                MsgBox " & Quote & "Duplicate key" & Quote & vbCrLf & _
        "                " & TableName & "TryCopyArrayToDictionary = False" & vbCrLf & _
        "                GoTo Done" & vbCrLf & _
        "            Else" & vbCrLf & _
        "                Set Record = New " & ClassName & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Dim Entry As Variant
    For Each Entry In DetailsDict.Keys
        If DetailsDict.Item(Entry).VariableType = "Boolean" Then
            Line = _
                "                Record." & DetailsDict.Item(Entry).VariableName & " = IIf(Ary(I, p" & DetailsDict.Item(Entry).VariableName & "Column) = " & Quote & "Yes" & Quote & ", True,False)"
        Else
            Line = _
                "                Record." & DetailsDict.Item(Entry).VariableName & " = Ary(I, p" & DetailsDict.Item(Entry).VariableName & "Column)"
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
        "    End If" & vbCrLf & vbCrLf & _
        "    '    Array formatting goes here" & vbCrLf
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


