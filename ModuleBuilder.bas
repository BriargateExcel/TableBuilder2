Attribute VB_Name = "ModuleBuilder"
Option Explicit

Private Const Module_Name As String = "ModuleBuilder."

Private Const Quote As String = """"

Public Sub ModuleBuilder( _
    ByVal DetailsTable As ListObject, _
    ByVal BasicsTable As ListObject)

    ' This routine builds the basic module

    Const RoutineName As String = Module_Name & "ClassBuilder"
    On Error GoTo ErrorHandler
    
    Dim DetailsDict As Dictionary
    If TableDetails.TryCopyTableToDictionary(DetailsTable, DetailsDict) Then
        ' Success; do nothing
    Else
        ReportError "Error copying Table to dictionary", "Routine", RoutineName
    End If
    
    Dim BasicDict As Dictionary
    If TableBasics.TryCopyTableToDictionary(BasicsTable, BasicDict) Then
        ' Success; do nothing
    Else
        ReportError "Error copying TableBasics to dictionary", "Routine", RoutineName
    End If
    
    Dim StreamName As String
    StreamName = BasicDict.Items(0).TableName & ".bas"
    
    Dim StreamFile As MessageFileClass
    Set StreamFile = New MessageFileClass
    
    Dim Line As String
    
    Line = _
        "Attribute VB_Name = " & Quote & BasicDict.Items(0).TableName & Quote & vbCrLf & _
        "Option Explicit" & vbCrLf & _
        vbCrLf & _
        "' Built on " & Now() & vbCrLf & _
        "' Built By Briargate Excel Table Builder" & vbCrLf & _
        "' See BriargateExcel.com for details" & vbCrLf & _
        vbCrLf & _
        "Private Const Module_Name As String = " & Quote & BasicDict.Items(0).TableName & "." & Quote & vbCrLf & vbCrLf & _
        "Private pInitialized As Boolean"
    StreamFile.WriteMessageLine Line, StreamName, "Basic Modules"

    Line = _
        "Private p" & BasicDict.Items(0).TableName & "Dict As Dictionary" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    BuildConstants StreamFile, StreamName, DetailsDict

    Line = _
        "Public Property Get " & BasicDict.Items(0).TableName & "Table() As ListObject" & vbCrLf & _
        "    Set " & BasicDict.Items(0).TableName & "Table = " & BasicDict.Items(0).TableName & "Sheet.ListObjects(" _
        & Quote & BasicDict.Items(0).TableName & "Table" & Quote & ")" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "Public Sub Reset" & BasicDict.Items(0).TableName & "()" & vbCrLf & _
        "    pInitialized = False" & vbCrLf & _
        "    Set p" & BasicDict.Items(0).TableName & "Dict = Nothing" & vbCrLf & _
        "End Sub" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "Public Function TryCopyTableToDictionary( " & "_" & vbCrLf & _
        "    ByVal Tbl As ListObject, _" & vbCrLf & _
        "    Optional ByRef Dict As Dictionary _" & vbCrLf & _
        "    ) As Boolean" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "    Const RoutineName As String = Module_Name & " & Quote & "TryCopyTableToDictionary" & Quote & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        "    TryCopyTableToDictionary = True" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "    Dim Ary As Variant" & vbCrLf & _
        "    Ary = Tbl.DataBodyRange" & vbCrLf & _
        "    If Err.Number <> 0 Then" & vbCrLf & _
        "        MsgBox " & Quote & "The " & BasicDict.Items(0).TableName & " table is empty" & Quote & vbCrLf & _
        "        TryCopyTableToDictionary = False" & vbCrLf & _
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
        "        Set ThisDict = pTableDetailsDict" & vbCrLf & _
        "    End If" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "    If " & BasicDict.Items(0).TableName & ".TryCopyArrayToDictionary(Ary, ThisDict) Then" & vbCrLf & _
        "        ' Success; do nothing" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        ReportError " & Quote & "Error copying array to dictionary" & Quote & ", " & Quote & "Routine" & Quote & ", RoutineName" & vbCrLf & _
        "        TryCopyTableToDictionary = False" & vbCrLf & _
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
        "End Function ' TryCopyTableToDictionary" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    BuildHeaders StreamFile, StreamName, DetailsDict

    Line = _
        "Public Property Get TableDetailsDictionary() As Dictionary" & vbCrLf & _
        "   Set TableDetailsDictionary = pTableDetailsDict" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "Public Property Get TableDetailsHeaderWidth() As Long" & vbCrLf & _
        "    TableDetailsHeaderWidth = pHeaderWidth" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
            "Public Property Get TableDetailsInitialized() As Boolean" & vbCrLf & _
            "    TableDetailsInitialized = pInitialized" & vbCrLf & _
            "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "Public Sub Initialize()" & vbCrLf & _
        vbCrLf & _
        "    ' This routine loads the dictionary" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & " & Quote & "Initialize" & Quote & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    pInitialized = True" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "    Dim TableDetails As TableDetails_Table" & vbCrLf & _
        "    Set TableDetails = New TableDetails_Table" & vbCrLf & _
        "    Set pTableDetailsDict = New Dictionary" & vbCrLf & _
        "    If Table.TryCopyTableToDictionary(TableDetails, TableDetailsTable, pTableDetailsDict) Then" & vbCrLf & _
        "        ' Success; do nothing" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        MsgBox " & Quote & "Error copying TableDetails table" & Quote & vbCrLf & _
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
        "End Sub ' Initialize" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
    "Public Sub CopyDictionaryToArray( _" & vbCrLf & _
    "    ByVal DetailsDict As Dictionary, _" & vbCrLf & _
    "       ByRef Ary As Variant)" & vbCrLf & _
    vbCrLf & _
    "    ' loads TableDetails Dict into Ary" & vbCrLf & _
    vbCrLf & _
    "    Const RoutineName As String = Module_Name & " & Quote & "CopyDictionaryToArray" & Quote & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    vbCrLf & _
    "    Dim I As Long" & vbCrLf & _
    "    I = 1" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    RecordToArray StreamFile, StreamName, DetailsDict, BasicDict

    Line = _
        "Done:" & vbCrLf & _
        "    Exit Sub" & vbCrLf & _
        "ErrorHandler:" & vbCrLf & _
        "    ReportError " & Quote & "Exception raised" & Quote & ", _" & vbCrLf & _
        "                " & Quote & "Routine" & Quote & ", RoutineName, _" & vbCrLf & _
        "                " & Quote & "Error Number" & Quote & ", Err.Number, _" & vbCrLf & _
        "                " & Quote & "Error Description" & Quote & ", Err.Description" & vbCrLf & _
        "    RaiseError Err.Number, Err.Source, RoutineName, Err.Description" & vbCrLf & _
        "End Sub ' CopyDictionaryToArray" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "Public Function TryCopyArrayToDictionary( _" & vbCrLf & _
        "       ByVal Ary As Variant, _" & vbCrLf & _
        "       ByRef Dict As Dictionary)" & vbCrLf & _
        vbCrLf & _
        "    ' Copy TableDetails array to dictionary" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & " & Quote & "TryCopyArrayToDictionary" & Quote & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    TryCopyArrayToDictionary = True" & vbCrLf & _
        vbCrLf & _
        "    Dim I As Long" & vbCrLf & _
        vbCrLf & _
        "    Set Dict = New Dictionary" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    ArrayToRecord StreamFile, StreamName, DetailsDict, BasicDict

    Line = _
        "Done:" & vbCrLf & _
        "    Exit Function" & vbCrLf & _
        "ErrorHandler:" & vbCrLf & _
        "    ReportError " & Quote & "Exception raised" & Quote & ", _" & vbCrLf & _
        "                " & Quote & "Routine" & Quote & ", RoutineName, _" & vbCrLf & _
        "                " & Quote & "Error Number" & Quote & ", Err.Number, _" & vbCrLf & _
        "                " & Quote & "Error Description" & Quote & ", Err.Description" & vbCrLf & _
        "    RaiseError Err.Number, Err.Source, RoutineName, Err.Description" & vbCrLf & _
        "End Function ' TryCopyArrayToDictionary" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName

    Line = _
        "Public Function TryCopyDictionaryToTable( _" & vbCrLf & _
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
        "    Const RoutineName As String = Module_Name & " & Quote & "TryCopyDictionaryToTable" & Quote & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    TryCopyDictionaryToTable = True" & vbCrLf & _
        vbCrLf & _
        "    If Not pInitialized Then TableDetails.Initialize" & vbCrLf & _
        vbCrLf & _
        "    Dim ClassName As " & BasicDict.Items(0).ClassName & vbCrLf & _
        "    Set ClassName = New " & BasicDict.Items(0).ClassName & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    Line = _
        "    '    FormatColumnAsText pFirstColumn, Table, TableCorner" & vbCrLf & _
        vbCrLf & _
        "    If Table.TryCopyDictionaryToTable(ClassName, Dict, Table, TableCorner, TableName) Then" & vbCrLf & _
        "        ' Success; do " & vbCrLf & _
        "    Else" & vbCrLf & _
        "        MsgBox " & Quote & "Error copying " & BasicDict.Items(0).TableName & " dictionary to table" & Quote & vbCrLf & _
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
        "End Function ' TryCopyDictionaryToTable" & vbCrLf
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
        ByVal BasicDict As Dictionary)

    ' This routine builds an array from details
    
    Const RoutineName As String = Module_Name & "RecordToArray"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = "    Dim Record As " & BasicDict.Items(0).ClassName & vbCrLf & _
        "    Dim Entry As Variant" & vbCrLf & _
        "    For Each Entry In DetailsDict.Keys" & vbCrLf & _
        "        Set Record = DetailsDict.Item(Entry)" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Dim Entry As Variant
    Dim I As Long
    For Each Entry In DetailsDict.Keys
        I = I + 1
        
        Line = "        Ary(I, p" & DetailsDict.Item(Entry).VariableName & "Column) = Record." & DetailsDict.Item(Entry).VariableName
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
        ByVal DetailsDict As Dictionary)

    ' This routine builds the Headers property
    
    Const RoutineName As String = Module_Name & "RecordToArray"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = "Public Property Get Headers() As Variant"
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = "    Headers = Array(" & Quote & DetailsDict.Items(0).VariableName & Quote
    
    Dim I As Long
    For I = 1 To DetailsDict.Count - 1
        Line = Line & ", " & Quote & DetailsDict.Items(I).VariableName & Quote
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
        ByVal BasicDict As Dictionary)

    ' This routine builds a dictionary from an array
    
    Const RoutineName As String = Module_Name & "ArrayToRecord"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    Line = _
        "    Dim Key As String" & vbCrLf & _
        "    Dim Record as " & BasicDict.Items(0).ClassName & vbCrLf & vbCrLf & _
        "    For I = 1 To UBound(Ary, 1)" & vbCrLf & _
        "        Key = Ary(I, p" & DetailsDict.Items(0).VariableName & "Column)" & vbCrLf & vbCrLf & _
        "        If Dict.Exists(Key) Then" & vbCrLf & _
        "            MsgBox " & Quote & "Duplicate key" & Quote & vbCrLf & _
        "            TryCopyArrayToDictionary = False" & vbCrLf & _
        "            GoTo Done" & vbCrLf & _
        "        Else" & vbCrLf & _
        "            Set Record = New " & BasicDict.Items(0).ClassName & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Dim Entry As Variant
    For Each Entry In DetailsDict.Keys
        If DetailsDict.Item(Entry).VariableType = "Boolean" Then
        
'        Record.Formatted = IIf(Ary(I, pFormattedColumn) = "Yes", True, False)
        
            Line = "            Record." & DetailsDict.Item(Entry).VariableName & " = IIf(Ary(I, p" & DetailsDict.Item(Entry).VariableName & "Column) = " & Quote & "Yes" & Quote & ", True,False)"
        
        
        
        Else
            Line = "            Record." & DetailsDict.Item(Entry).VariableName & " = Ary(I, p" & DetailsDict.Item(Entry).VariableName & "Column)"
        End If
        StreamFile.WriteMessageLine Line, StreamName
    Next Entry
    
    StreamFile.WriteBlankMessageLines StreamName
    
    Line = _
        "            Dict.Add Key, Record" & vbCrLf & _
        "        End If" & vbCrLf & _
        "    Next I" & vbCrLf & vbCrLf & _
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


