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
    ' Declarations
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

    BuildConstants StreamFile, StreamName, DetailsDict

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
    ' Reset
    '
    
    Line = _
        "Public Sub " & TableName & "Reset()" & vbCrLf & _
        "    pInitialized = False" & vbCrLf & _
        "    Set p" & TableName & "Dict = Nothing" & vbCrLf & _
        "End Sub" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' TryCopyTableToDictionary
    '
    
    Line = _
        "Public Function " & TableName & "TryCopyTableToDictionary( " & "_" & vbCrLf & _
        "    ByVal Tbl As ListObject, _" & vbCrLf & _
        "    Optional ByRef Dict As Dictionary _" & vbCrLf & _
        "    ) As Boolean" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "    Const RoutineName As String = Module_Name & " & Quote & TableName & "TryCopyTableToDictionary" & Quote & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        "    " & TableName & "TryCopyTableToDictionary = True" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "    Dim Ary As Variant" & vbCrLf & _
        "    Ary = Tbl.DataBodyRange" & vbCrLf & _
        "    If Err.Number <> 0 Then" & vbCrLf & _
        "        MsgBox " & Quote & "The " & TableName & " table is empty" & Quote & vbCrLf & _
        "        " & TableName & "TryCopyTableToDictionary = False" & vbCrLf & _
        "        GoTo Done" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    Err.Clear" & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "    Dim ThisDict As Dictionary" & vbCrLf & _
        "    If Dict Is Nothing Then" & vbCrLf & _
        "        Set ThisDict = New Dictionary" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        Set ThisDict = p" & TableName & "Dict" & vbCrLf & _
        "    End If" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "    If " & TableName & "TryCopyArrayToDictionary(Ary, ThisDict) Then" & vbCrLf & _
        "        ' Success; do nothing" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        ReportError " & Quote & "Error copying array to dictionary" & Quote & ", " & Quote & "Routine" & Quote & ", RoutineName" & vbCrLf & _
        "        " & TableName & "TryCopyTableToDictionary = False" & vbCrLf & _
        "        GoTo Done" & vbCrLf & _
        "    End If" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "    Set Dict = ThisDict" & vbCrLf & _
        vbCrLf & _
        "Done:" & vbCrLf & _
        "    Exit Function" & vbCrLf & _
        "ErrorHandler:" & vbCrLf & _
        "    ReportError " & Quote & "Exception raised" & Quote & ", _" & vbCrLf & _
        "                " & Quote & "Routine" & Quote & ", RoutineName, _" & vbCrLf & _
        "                " & Quote & "Error Number" & Quote & ", Err.Number, _" & vbCrLf & _
        "                " & Quote & "Error Description" & Quote & ", Err.Description" & vbCrLf & _
        "    RaiseError Err.Number, Err.Source, RoutineName, Err.Description" & vbCrLf & _
        "End Function ' " & TableName & "TryCopyTableToDictionary" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' TryCopyDictionaryToTable
    '
    
    Line = _
        "Public Function " & TableName & "TryCopyDictionaryToTable( _" & vbCrLf & _
        "       ByVal Dict As Dictionary, _" & vbCrLf & _
        "       Optional ByVal Table As ListObject = Nothing, _" & vbCrLf & _
        "       Optional TableCorner As Range = Nothing, _" & vbCrLf & _
        "       Optional TableName As String _" & vbCrLf & _
        "       ) As Boolean" & vbCrLf & _
        vbCrLf & _
        "    ' This routine copies a dictionary to a table" & vbCrLf & _
        "    ' If Dict is nothing then use default dictionary" & vbCrLf & _
        "    ' If Table is nothing then build a table using TableCorner and TableName" & vbCrLf & _
        "    ' if Table and TableCorner are both Nothing then use TableDetailsTable" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & " & Quote & TableName & "TryCopyDictionaryToTable" & Quote & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    " & TableName & "TryCopyDictionaryToTable = True" & vbCrLf & _
        vbCrLf & _
        "    If Not pInitialized Then " & TableName & "Initialize" & vbCrLf & _
        vbCrLf & _
        "    Dim ClassName As " & ClassName & vbCrLf & _
        "    Set ClassName = New " & ClassName & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    Line = _
        "    '    FormatColumnAsText pFirstColumn, Table, TableCorner" & vbCrLf & _
        vbCrLf & _
        "    If Table.TryCopyDictionaryToTable(ClassName, Dict, Table, TableCorner, TableName) Then" & vbCrLf & _
        "        ' Success; do " & vbCrLf & _
        "    Else" & vbCrLf & _
        "        MsgBox " & Quote & "Error copying " & TableName & " dictionary to table" & Quote & vbCrLf & _
        "        " & TableName & "TryCopyDictionaryToTable = False" & vbCrLf & _
        "    End If" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "Done:" & vbCrLf & _
        "    Exit Function" & vbCrLf & _
        "ErrorHandler:" & vbCrLf & _
        "    ReportError " & Quote & "Exception raised" & Quote & ", _" & vbCrLf & _
        "                " & Quote & "Routine" & Quote & ", RoutineName, _" & vbCrLf & _
        "                " & Quote & "Error Number" & Quote & ", Err.Number, _" & vbCrLf & _
        "                " & Quote & "Error Description" & Quote & ", Err.Description" & vbCrLf & _
        "    RaiseError Err.Number, Err.Source, RoutineName, Err.Description" & vbCrLf & _
        "End Function ' " & TableName & "TryCopyDictionaryToTable" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' HeaderWidth
    '
    
    Line = _
        "Private Property Get " & TableName & "HeaderWidth() As Long" & vbCrLf & _
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
        "Private Sub " & TableName & "Initialize()" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & " & Quote & TableName & "Initialize" & Quote & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    pInitialized = True" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "    Dim  " & TableName & " As " & ClassName & vbCrLf & _
        "    Set " & TableName & " = New " & ClassName & vbCrLf & vbCrLf & _
        "    Set p" & TableName & "Dict = New Dictionary" & vbCrLf & _
        "    If " & TableName & "TryCopyTableToDictionary(" & TableName & "Table, p" & TableName & "Dict) Then" & vbCrLf & _
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
    ' CopyDictionaryToArray
    '
    
    Line = _
    "Private Sub " & TableName & "CopyDictionaryToArray( _" & vbCrLf & _
    "    ByVal DetailsDict As Dictionary, _" & vbCrLf & _
    "       ByRef Ary As Variant)" & vbCrLf & _
    vbCrLf & _
    "    Const RoutineName As String = Module_Name & " & Quote & TableName & "CopyDictionaryToArray" & Quote & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    vbCrLf & _
    "    Dim I As Long" & vbCrLf & _
    "    I = 1" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    RecordToArray StreamFile, StreamName, DetailsDict, ClassName

    Line = _
        "Done:" & vbCrLf & _
        "    Exit Sub" & vbCrLf & _
        "ErrorHandler:" & vbCrLf & _
        "    ReportError " & Quote & "Exception raised" & Quote & ", _" & vbCrLf & _
        "                " & Quote & "Routine" & Quote & ", RoutineName, _" & vbCrLf & _
        "                " & Quote & "Error Number" & Quote & ", Err.Number, _" & vbCrLf & _
        "                " & Quote & "Error Description" & Quote & ", Err.Description" & vbCrLf & _
        "    RaiseError Err.Number, Err.Source, RoutineName, Err.Description" & vbCrLf & _
        "End Sub ' " & TableName & "CopyDictionaryToArray" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    '
    ' TryCopyArrayToDictionary
    '
    
    Line = _
        "Private Function " & TableName & "TryCopyArrayToDictionary( _" & vbCrLf & _
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

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' ModuleBuilder

Private Sub BuildConstants( _
        ByVal StreamFile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal DetailsDict As Dictionary)

    ' This routine builds Get and Let Properties
    
    Const RoutineName As String = Module_Name & "BuildConstants"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    Dim Counter As Long
    
    Dim Entry As Variant
    For Each Entry In DetailsDict.Keys
        Counter = Counter + 1
        
        Line = "Private Const p" & DetailsDict.Item(Entry).VariableName & "Column As Long = " & Counter
        StreamFile.WriteMessageLine Line, StreamName
    Next Entry
    
    Line = "Private Const pHeaderWidth As Long = " & Counter
    StreamFile.WriteMessageLine Line, StreamName
    
    StreamFile.WriteBlankMessageLines StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildConstants

Private Sub RecordToArray( _
        ByVal StreamFile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal DetailsDict As Dictionary, _
        ByVal ClassName As String)

    ' This routine builds an array from details
    
    Const RoutineName As String = Module_Name & "RecordToArray"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = "    Dim Record As " & ClassName & vbCrLf & _
        "    Dim Entry As Variant" & vbCrLf & _
        "    For Each Entry In DetailsDict.Keys" & vbCrLf & _
        "        Set Record = DetailsDict.Item(Entry)" & vbCrLf
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
    
    Line = "Private Property Get " & TableName & "Headers() As Variant"
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
        "            Key = Ary(I, p" & DetailsDict.Items(0).VariableName & "Column)" & vbCrLf & vbCrLf & _
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


