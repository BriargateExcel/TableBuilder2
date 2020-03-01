Attribute VB_Name = "ModuleBuilder"
Option Explicit

Private Const Module_Name As String = "ModuleBuilder."

Public Sub ModuleBuilder( _
    ByVal DetailsDict As Dictionary, _
    ByVal TableName As String, _
    ByVal ClassName As String)

    ' This routine builds the basic module

    Const RoutineName As String = Module_Name & "ClassBuilder"
    On Error GoTo ErrorHandler
    
    Dim StreamName As String
    StreamName = TableName & ".bas"
    
    Dim Streamfile As MessageFileClass
    Set Streamfile = New MessageFileClass
    
    Dim Line As String
    
    '
    ' Module name and initial declarations
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
        "Private Type %1Type" & vbCrLf & _
        "    Initialized as Boolean" & vbCrLf & _
        "    Dict as Dictionary" & vbCrLf & _
        "End Type" & vbCrLf & _
        vbCrLf & _
        "Private This as %1Type" & vbCrLf, _
        TableName)
    Streamfile.WriteMessageLine Line, StreamName, "Modules", True

    '
    ' Declarations separator
    '
    
    Line = _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "'   Start of application specific declarations     '" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
    Streamfile.WriteMessageLine Line, StreamName

        
    '
    ' Application specific declarations
    '
    
    BuildApplicationUniqueDeclarations Streamfile, StreamName, TableName, ".bas"
        
    '
    ' Declarations separator
    '
    
    Line = _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "'    End of application specific declarations      '" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
    Streamfile.WriteMessageLine Line, StreamName

    '
    ' Constants and properties
    '
    
    BuildConstantsAndProperties Streamfile, StreamName, DetailsDict, TableName

    '
    ' Property Get Dictionary
    '
    Line = PrintString( _
        "Public Property Get %1Dictionary() As Dictionary" & vbCrLf & _
        "   Set %1Dictionary = This.Dict" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    Streamfile.WriteMessageLine Line, StreamName

    '
    ' Property Get Initialized
    '
    
    Line = PrintString( _
        "Public Property Get %1Initialized() As Boolean" & vbCrLf & _
        "   %1Initialized = This.Initialized" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    Streamfile.WriteMessageLine Line, StreamName

    '
    ' Sub Reset
    '
    
    Line = PrintString( _
        "Public Sub %1Reset()" & vbCrLf & _
        "    This.Initialized = False" & vbCrLf & _
        "    Set This.Dict = Nothing" & vbCrLf & _
        "End Sub" & vbCrLf, _
        TableName)
    Streamfile.WriteMessageLine Line, StreamName

    '
    ' Property Get HeaderWidth
    '
    
    Line = PrintString( _
        "Public Property Get %1HeaderWidth() As Long" & vbCrLf & _
        "    %1HeaderWidth = pHeaderWidth" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    Streamfile.WriteMessageLine Line, StreamName

    '
    ' Constants for table columns
    '
    
    BuildColumnConstants Streamfile, StreamName, DetailsDict, TableName

    '
    ' Sub Initialize
    '
    
    Line = PrintString( _
        "Public Sub %1Initialize()" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & qq%1Initializeqq" & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf, _
        TableName)
    Streamfile.WriteMessageLine Line, StreamName

    Line = PrintString( _
        "    Dim  %1 As %2" & vbCrLf & _
        "    Set %1 = New %2" & vbCrLf & _
        vbCrLf & _
        "    Set This.Dict = New Dictionary" & vbCrLf & _
        "    If Table.TryCopyTableToDictionary(%1, %1Table, This.Dict) Then" & vbCrLf & _
        "        This.Initialized = True" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        ReportError qqError copying %1 tableqq, qqRoutineqq, RoutineName" & vbCrLf & _
        "        This.Initialized = False" & vbCrLf & _
        "        GoTo Done" & vbCrLf & _
        "    End If" & vbCrLf, _
        TableName, ClassName)
     Streamfile.WriteMessageLine Line, StreamName

    SubEnding Streamfile, StreamName, TableName & "Initialize"

    '
    ' Function TryCopyDictionaryToArray
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
    Streamfile.WriteMessageLine Line, StreamName

    RecordToArray Streamfile, StreamName, DetailsDict, ClassName, TableName
    
    '
    ' Function CheckExists
    '
    
    BuildCheckExists DetailsDict, TableName, Streamfile, StreamName
    
    '
    ' TryCopyArrayToDictionary
    '
    
    Line = PrintString( _
        "Public Function %1TryCopyArrayToDictionary( _" & vbCrLf & _
        "       ByVal Ary As Variant, _" & vbCrLf & _
        "       ByRef Dict As Dictionary _" & vbCrLf & _
        "       ) As Boolean" & vbCrLf & _
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
    Streamfile.WriteMessageLine Line, StreamName

    ArrayToRecord Streamfile, StreamName, DetailsDict, ClassName, TableName

    FunctionEnding Streamfile, StreamName, TableName & "TryCopyArrayToDictionary"
    
    '
    ' Sub FormatArrayAndWorksheet
    '
    
    Line = PrintString( _
        "Public Sub %1FormatArrayAndWorksheet( _" & vbCrLf & _
        "    ByRef Ary as Variant, _" & vbCrLf & _
        "    ByVal Table As ListObject)" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & qq%1FormatArrayAndWorksheetqq" & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf, _
        TableName)
    Streamfile.WriteMessageLine Line, StreamName
    
    FormatDetails DetailsDict, Streamfile, StreamName

    SubEnding Streamfile, StreamName, TableName & "FormatArrayAndWorksheet"
    
    '
    ' Property Get Routines
    '
    
    BuildPropertyGetRoutines Streamfile, StreamName, DetailsDict, TableName
    
    '
    ' Separator for routines potentially needing changes
    '
    
    Line = _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "'         The routines that follow may need        '" & vbCrLf & _
        "'        changes depending on the application      '" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
    Streamfile.WriteMessageLine Line, StreamName

    '
    ' Property Get Table
    '
    
    Line = PrintString( _
        "Public Property Get %1Table() As ListObject" & vbCrLf & _
        vbCrLf & _
        "    ' Change the table reference if the table is in another workbook" & vbCrLf & _
        vbCrLf & _
        "    Set %1Table = %1Sheet.ListObjects(qq%1Tableqq)" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    Streamfile.WriteMessageLine Line, StreamName

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
    Streamfile.WriteMessageLine Line, StreamName

    '
    ' End of generated code comment
    '

    BuildApplicationUniqueRoutines Streamfile, StreamName, TableName, ".bas"
    
    '
    ' Wrapup
    '

    Set Streamfile = Nothing

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
        ByVal Streamfile As MessageFileClass, _
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
        Streamfile.WriteMessageLine Line, StreamName
    Next Entry
    
    Line = _
        "Private Const pHeaderWidth As Long = " & Counter & vbCrLf
    Streamfile.WriteMessageLine Line, StreamName
    
    Counter = 0
    For Each Entry In DetailsDict.Keys
        Counter = Counter + 1
        
        Line = PrintString( _
            "Public Property Get %1%2Column() As Long" & vbCrLf & _
            "    %1%2Column = p%2Column" & vbCrLf & _
            "End Property" & vbCrLf, _
            TableName, DetailsDict.Item(Entry).VariableName)
        Streamfile.WriteMessageLine Line, StreamName
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
        ByVal Streamfile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal DetailsDict As Dictionary, _
        ByVal ClassName As String, _
        ByVal TableName As String)

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
    Streamfile.WriteMessageLine Line, StreamName
    
    Dim Entry As Variant
    Dim I As Long
    For Each Entry In DetailsDict.Keys
        I = I + 1
        Line = PrintString( _
            "        Ary(I, p%1Column) = Record.%1", _
            DetailsDict.Item(Entry).VariableName)
        Streamfile.WriteMessageLine Line, StreamName
    Next Entry
    
    Line = vbCrLf & "        I = I + 1"
    Streamfile.WriteMessageLine Line, StreamName
    
    Line = "    Next Entry" & vbCrLf
    Streamfile.WriteMessageLine Line, StreamName
    
    FunctionEnding Streamfile, StreamName, TableName & "TryCopyDictionaryToArray"

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' RecordToArray

Private Sub BuildColumnConstants( _
        ByVal Streamfile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal DetailsDict As Dictionary, _
        ByVal TableName As String)

    ' This routine builds the constants defining the table columns
    
    Const RoutineName As String = Module_Name & "BuildColumnConstants"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "Public Property Get %1Headers() As Variant", _
        TableName)
    Streamfile.WriteMessageLine Line, StreamName
    
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
    Streamfile.WriteMessageLine Line, StreamName
    
    Line = "End Property" & vbCrLf
    Streamfile.WriteMessageLine Line, StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildColumnConstants

Private Sub BuildCheckExists( _
    ByVal Dict As Dictionary, _
    ByVal TableName As String, _
    ByVal Streamfile As MessageFileClass, _
    ByVal StreamName As String)

    ' Builds the CheckExists routine
    
    Const RoutineName As String = Module_Name & "BuildCheckExists"
    On Error GoTo ErrorHandler
    
    Dim Entry As Variant
    Dim TD As TableDetails_Table
    Dim Found As Boolean
    Dim Key As String
    For Each Entry In Dict
        Set TD = Dict(Entry)
        If TD.Key = "Key" Then
            Key = TD.VariableName
            Found = True
            Exit For
        End If
    Next Entry
    
    If Not Found Then GoTo Done
    
    Dim Line As String
    
    Line = PrintString( _
        "Public Function Check%1Exists(ByVal %1 As String) As Boolean _" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & qqCheck%1Existsqq" & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    If Not This.Initialized Then %2Initialize" & vbCrLf, _
        Key, TableName)
    Streamfile.WriteMessageLine Line, StreamName
    
    Line = PrintString( _
        "    If %1 = vbNullString Then" & vbCrLf & _
        "        Check%1Exists = True" & vbCrLf & _
        "        Exit Function" & vbCrLf & _
        "    End If" & vbCrLf & _
        vbCrLf & _
        "    Check%1Exists = This.Dict.Exists(%1)" & vbCrLf, _
        Key, TableName)
    Streamfile.WriteMessageLine Line, StreamName

    FunctionEnding Streamfile, StreamName, "Check" & Key & "Exists"
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub      ' BuildCheckExists

Private Sub BuildKeyArray( _
    ByVal Dict As Dictionary, _
    ByRef Ary As Variant)

    ' Populates Ary with the keys in order
    
    Const RoutineName As String = Module_Name & "BuildKeyArray"
    On Error GoTo ErrorHandler
    
    Dim Entry As Variant
    Dim TD As TableDetails_Table
    Dim Count As Long
    For Each Entry In Dict
        Set TD = Dict(Entry)
        If Left$(TD.Key, 3) = "Key" Then
            Count = Count + 1
        End If
    Next Entry
    
    If Count = 0 Then
        ReDim Ary(1 To 1)
        Ary(1) = "None"
        GoTo Done
    Else
        ReDim Ary(1 To Count)
    End If
    
    Dim Key As String
    For Each Entry In Dict
        Set TD = Dict(Entry)
        Key = TD.Key
        If Left$(Key, 3) = "Key" Then
            If Len(Key) > 3 Then
                Ary(Right$(Key, 1)) = TD.VariableName
            Else
                ' This is the only Key
                Ary(1) = TD.VariableName
                GoTo Done
            End If
        End If
    Next Entry
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildKeyArray

Private Sub ArrayToRecord( _
        ByVal Streamfile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal DetailsDict As Dictionary, _
        ByVal ClassName As String, _
        ByVal TableName As String)

    ' This routine builds a dictionary from an array
    
    Const RoutineName As String = Module_Name & "ArrayToRecord"
    On Error GoTo ErrorHandler
    
    Dim Ary As Variant
    BuildKeyArray DetailsDict, Ary
    
    Dim Line As String
    
    If UBound(Ary, 1) = 1 Or Ary(1) = "None" Then
        BuildNoneOrOneKey Ary(1), TableName, ClassName, Streamfile, StreamName
    Else
        BuildMoreThanOneKey Ary, TableName, ClassName, Streamfile, StreamName
    End If
    
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
        Streamfile.WriteMessageLine Line, StreamName
    Next Entry

    Streamfile.WriteBlankMessageLines StreamName
    
    Line = _
        "                Dict.Add Key, Record" & vbCrLf & _
        "            End If" & vbCrLf & _
        "        Next I" & vbCrLf & vbCrLf & _
        "    Else" & vbCrLf & _
        "        Dict.Add Ary, Ary" & vbCrLf & _
        "    End If" & vbCrLf
    Streamfile.WriteMessageLine Line, StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' ArrayToRecord

Private Sub BuildNoneOrOneKey( _
    ByVal Key As String, _
    ByVal TableName As String, _
    ByVal ClassName As String, _
    ByVal Streamfile As MessageFileClass, _
    ByVal StreamName As String)

    ' Build the ArrayToDictionary code for none or one key
    
    Const RoutineName As String = Module_Name & "BuildNoneOrOneKey"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "    Dim Key As String" & vbCrLf & _
        "    Dim Record as %1" & vbCrLf & vbCrLf & _
        "    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then" & vbCrLf & _
        "        For I = 1 To UBound(Ary, 1)" & vbCrLf & _
        IIf(Key = "None", _
            "            Key = Ary(I, 1)" & vbCrLf, _
            "            Key = Ary(I, p%2Column)" & vbCrLf) & _
        vbCrLf & _
        "            If Dict.Exists(Key) Then" & vbCrLf & _
        "                ReportWarning qqDuplicate keyqq, qqRoutineqq, RoutineName, qqKeyqq, Key" & vbCrLf & _
        "                %3TryCopyArrayToDictionary = False" & vbCrLf & _
        "                GoTo Done" & vbCrLf & _
        "            Else" & vbCrLf & _
        "                Set Record = New " & ClassName & vbCrLf, _
        ClassName, Key, TableName)
    Streamfile.WriteMessageLine Line, StreamName
        
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildNoneOrOneKey

Private Sub BuildMoreThanOneKey( _
    ByVal Ary As Variant, _
    ByVal TableName As String, _
    ByVal ClassName As String, _
    ByVal Streamfile As MessageFileClass, _
    ByVal StreamName As String)

    ' Build the ArrayToDictionary code more than one key
    
    Const RoutineName As String = Module_Name & "BuildMoreThanOneKey"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "    Dim Key As String" & vbCrLf & _
        "    Dim Record as %1" & vbCrLf & vbCrLf & _
        "    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then" & vbCrLf & _
        "        For I = 1 To UBound(Ary, 1)" & vbCrLf, _
        ClassName)
    Streamfile.WriteMessageLine Line, StreamName
    
    Line = PrintString("            Key = qq|qq _", Ary(1))
        
    Dim I As Long
    For I = 1 To UBound(Ary, 1) - 1
        Line = Line & PrintString(vbCrLf & "                & Ary(I, p%1Column) & qq|qq _", Ary(I))
    Next I
    
    Line = Line & PrintString(vbCrLf & "                & Ary(I, p%1Column) & qq|qq" & vbCrLf, Ary(UBound(Ary, 1)))
    Streamfile.WriteMessageLine Line, StreamName
    
    Line = PrintString( _
        "            If Dict.Exists(Key) Then" & vbCrLf & _
        "                ReportWarning qqDuplicate keyqq, qqRoutineqq, RoutineName, qqKeyqq, Key" & vbCrLf & _
        "                %1TryCopyArrayToDictionary = False" & vbCrLf & _
        "                GoTo Done" & vbCrLf & _
        "            Else" & vbCrLf & _
        "                Set Record = New " & ClassName & vbCrLf, _
        TableName)
    Streamfile.WriteMessageLine Line, StreamName
        
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildMoreThanOneKey

Private Sub FormatDetails( _
    ByVal DetailsDict As Dictionary, _
    ByVal Streamfile As MessageFileClass, _
    ByVal StreamName As String)

    ' Build the format routine calls
    
    Const RoutineName As String = Module_Name & "FormatDetails"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    Dim Entry As Variant
    Dim TD As TableDetails_Table
    Dim Vbl As String
    For Each Entry In DetailsDict
        Set TD = DetailsDict(Entry)
        Vbl = TD.VariableName
        
        Select Case TD.Format
        Case "CLIN"
            Line = PrintString("    CleanCLINData Table, p%1Column", Vbl)
        Case "Dollar"
            Line = PrintString("    CleanDollars Table, p%1Column", Vbl)
        Case "EmpNum"
            Line = PrintString("    CleanEmployeeData Ary, Table, p%1Column", Vbl)
        Case "Month"
            Line = PrintString("    CleanMonthData Table, p%1Column", Vbl)
        Case "TwoDecimal"
            Line = PrintString("    CleanTwoDecimalData Table, p%1Column", Vbl)
        Case "Week"
            Line = PrintString("    CleanWeekData Table, p%1Column", Vbl)
        End Select
        
        If TD.Format <> vbNullString Then Streamfile.WriteMessageLine Line, StreamName
    Next Entry
    
    Streamfile.WriteBlankMessageLines StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' FormatDetails

Private Sub BuildPropertyGetRoutines( _
        ByVal Streamfile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal DetailsDict As Dictionary, _
        ByVal TableName As String)

    ' Build the Property Get routines
    
    Const RoutineName As String = Module_Name & "BuildPropertyGetRoutines"
    On Error GoTo ErrorHandler
    
    Dim Entry As Variant
    Dim FoundAKey As Boolean
    Dim Details As TableDetails_Table
    Dim Key As String
    Dim KeyType As String
    For Each Entry In DetailsDict
        Set Details = DetailsDict(Entry)
        If Details.Key = "Key" Then
            Key = Details.VariableName
            KeyType = Details.VariableType
            FoundAKey = True
            Exit For
        End If
    Next Entry
    
    If Not FoundAKey Then GoTo Done
    
    Dim Target As String ' The variable name to be fetched
    Dim TargetType As String
    For Each Entry In DetailsDict
        Set Details = DetailsDict(Entry)
        If Details.VariableName <> Key Then
            Target = Details.VariableName
            TargetType = Details.VariableType
            BuildOnePropertyGetRoutine Streamfile, StreamName, Key, KeyType, Target, TargetType, TableName
        End If
    Next Entry
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildPropertyGetRoutines

Private Sub BuildOnePropertyGetRoutine( _
        ByVal Streamfile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal Key As String, _
        ByVal KeyType As String, _
        ByVal Target As String, _
        ByVal TargetType As String, _
        TableName As String)

    ' Build one Property Get routine
    
    Const RoutineName As String = Module_Name & "BuildOnePropertyGetRoutine"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    Line = PrintString( _
        "Public Property Get Get%1From%2(ByVal %2 As %3) As %4" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & qqGet%1From%2qq" & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
        vbCrLf & _
        "    If Not This.Initialized Then %5Initialize" & vbCrLf & _
        vbCrLf & _
        "    If Check%2Exists(%2) Then" & vbCrLf & _
        "        Get%1From%2 = This.Dict(%2).%1" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        ReportError qqUnrecognized %2qq, qqRoutineqq, RoutineName" & vbCrLf & _
        "    End If" & vbCrLf, _
        Target, Key, KeyType, TargetType, TableName)
        
    Streamfile.WriteMessageLine Line, StreamName
    
    PropertyEnding Streamfile, StreamName, PrintString("Get%1From%2", Target, Key)
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildOnePropertyGetRoutine

Private Sub SubEnding( _
    ByVal Streamfile As MessageFileClass, _
    ByVal StreamName As String, _
    ByVal SubName As String)

    ' The standard end for subs
    
    Const RoutineName As String = Module_Name & "SubEnding"
    On Error GoTo ErrorHandler
    
    RoutineEnding Streamfile, StreamName, SubName, "Sub"
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' SubEnding

Private Sub FunctionEnding( _
    ByVal Streamfile As MessageFileClass, _
    ByVal StreamName As String, _
    ByVal FunctionName As String)

    ' The standard end for functions
    
    Const RoutineName As String = Module_Name & "FunctionEnding"
    On Error GoTo ErrorHandler
    
    RoutineEnding Streamfile, StreamName, FunctionName, "Function"

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' FunctionEnding

Private Sub PropertyEnding( _
    ByVal Streamfile As MessageFileClass, _
    ByVal StreamName As String, _
    ByVal SubName As String)

    ' The standard end for Property
    
    Const RoutineName As String = Module_Name & "PropertyEnding"
    On Error GoTo ErrorHandler
    
    RoutineEnding Streamfile, StreamName, SubName, "Property"
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' PropertyEnding

Private Sub RoutineEnding( _
    ByVal Streamfile As MessageFileClass, _
    ByVal StreamName As String, _
    ByVal SubName As String, _
    ByVal RoutineType As String)

    ' Creates a standard routine ending
    
    Const RoutineName As String = Module_Name & "RoutineEnding"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "Done:" & vbCrLf & _
        "    Exit %1" & vbCrLf & _
        "ErrorHandler:" & vbCrLf & _
        "    ReportError qqException raisedqq, _" & vbCrLf & _
        "                qqRoutineqq, RoutineName, _" & vbCrLf & _
        "                qqError Numberqq, Err.Number, _" & vbCrLf & _
        "                qqError Descriptionqq, Err.Description" & vbCrLf & _
        "    RaiseError Err.Number, Err.Source, RoutineName, Err.Description" & vbCrLf & _
        "End %1 ' %2" & vbCrLf, _
        RoutineType, SubName)
    Streamfile.WriteMessageLine Line, StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' RoutineEnding


