Attribute VB_Name = "ModuleBuilder"
Option Explicit
'@Folder "Builder"
Private Const Module_Name As String = "ModuleBuilder."

Private Type ModuleData
    StreamFile As MessageFileClass
    StreamName As String
    TableName As String
    ClassName As String
    FileName As String
    WorksheetName As String
    ExternalTableName As String
    DetailsDict As Dictionary
    BasicDict As Dictionary
End Type ' ModuleData

Private This As ModuleData

' Order:
' Front end
' Application specific declarations
' Constants and properties
' Constants for table columns
' Get Routines
' Get Dictionary
' Get Table
' Get Initialized
' Initialize
' Reset
' Get Headerwidth
' Dictionary to Array
' Array to Dictionary
' Check Exists
' Format array and worksheet
' Application unique routines

Public Sub ModuleBuilder( _
    ByVal DetailsDict As Dictionary, _
    ByVal BasicDict As Dictionary)

    ' This routine builds the basic module

    Const RoutineName As String = Module_Name & "ClassBuilder"
    On Error GoTo ErrorHandler
    
    Set This.StreamFile = New MessageFileClass
    
    This.TableName = BasicDict.Items(0).TableName
    This.ClassName = This.TableName & "_Table"
    
    This.StreamName = This.TableName & ".bas"
    
    This.FileName = BasicDict.Items(0).FileName
    
    This.WorksheetName = BasicDict.Items(0).WorksheetName
    
    This.ExternalTableName = BasicDict.Items(0).ExternalTableName
    
    Set This.BasicDict = BasicDict
    
    Set This.DetailsDict = DetailsDict
    
    Dim Line As String
    
    ' Module name and initial declarations
    BuildFrontEnd
    
    ' Application specific declarations
    BuildApplicationUniqueDeclarations This.StreamFile, This.StreamName, This.TableName, ".bas"
        
    ' Constants and properties
    BuildConstantsAndProperties

    ' Constants for table columns
    BuildColumnConstants

    ' Property Get Routines
    BuildPropertyGetRoutines
    
    ' Property Get Dictionary
    Line = PrintString( _
        "Public Property Get %1Dictionary() As Dictionary" & vbCrLf & _
        "   Set %1Dictionary = This.Dict" & vbCrLf & _
        "End Property" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName

    ' Property Get Table
    BuildGetTable

    ' Property Get Initialized
    Line = PrintString( _
        "Public Property Get %1Initialized() As Boolean" & vbCrLf & _
        "   %1Initialized = This.Initialized" & vbCrLf & _
        "End Property" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName

    ' Sub Initialize
    BuildInitialize
    
    ' Sub Reset
    Line = PrintString( _
        "Public Sub %1Reset()" & vbCrLf & _
        "    This.Initialized = False" & vbCrLf & _
        "    Set This.Dict = Nothing" & vbCrLf & _
        "End Sub" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName

    ' Property Get HeaderWidth
    Line = PrintString( _
        "Public Property Get %1HeaderWidth() As Long" & vbCrLf & _
        "    %1HeaderWidth = pHeaderWidth" & vbCrLf & _
        "End Property" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Function TryCopyDictionaryToArray
    BuildDictionaryToArray

    ' TryCopyArrayToDictionary
    BuildArrayToDictionary
    
    ' Function CheckExists
    BuildCheckExists
    
    ' Sub FormatArrayAndWorksheet
    BuildFormatArrayAndWorksheet
    
    ' Application unique routines
    BuildApplicationUniqueRoutines This.StreamFile, This.StreamName, This.TableName, ".bas"
    
    ' Wrapup
    Set This.StreamFile = Nothing

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' ModuleBuilder

Private Sub BuildFrontEnd()

    ' The first few lines of a module
    
    Const RoutineName As String = Module_Name & "BuildFrontEnd"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
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
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName, "Modules", True
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildFrontEnd

Private Sub BuildConstantsAndProperties()

    ' This routine builds Get and Let Properties
    
    Const RoutineName As String = Module_Name & "BuildConstantsAndProperties"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    Dim Counter As Long
    
    Dim Entry As Variant
    For Each Entry In This.DetailsDict.Keys
        Counter = Counter + 1
        
        Line = PrintString( _
            "Private Const p%1Column As Long = " & Counter, _
            This.DetailsDict.Item(Entry).VariableName)
        This.StreamFile.WriteMessageLine Line, This.StreamName
    Next Entry
    
    Line = _
        "Private Const pHeaderWidth As Long = " & Counter & vbCrLf
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Counter = 0
    For Each Entry In This.DetailsDict.Keys
        Counter = Counter + 1
        
        Line = PrintString( _
            "Public Property Get %1%2Column() As Long" & vbCrLf & _
            "    %1%2Column = p%2Column" & vbCrLf & _
            "End Property" & vbCrLf, _
            This.TableName, This.DetailsDict.Item(Entry).VariableName)
        This.StreamFile.WriteMessageLine Line, This.StreamName
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

Private Sub BuildDictionaryToArray()

    ' This routine builds an array from details
    
    Const RoutineName As String = Module_Name & "BuildDictionaryToArray"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
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
    This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Line = PrintString( _
        "    Dim Record As %1" & vbCrLf & _
        "    Dim Entry As Variant" & vbCrLf & _
        "    For Each Entry In Dict.Keys" & vbCrLf & _
        "        Set Record = Dict.Item(Entry)" & vbCrLf, _
        This.ClassName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Dim Entry As Variant
    Dim I As Long
    For Each Entry In This.DetailsDict.Keys
        I = I + 1
        Line = PrintString( _
            "        Ary(I, p%1Column) = Record.%1", _
            This.DetailsDict.Item(Entry).VariableName)
        This.StreamFile.WriteMessageLine Line, This.StreamName
    Next Entry
    
    Line = vbCrLf & "        I = I + 1"
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Line = "    Next Entry" & vbCrLf
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    BuildFunctionEnding This.TableName & "TryCopyDictionaryToArray"

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildDictionaryToArray

Private Sub BuildColumnConstants()

    ' This routine builds the constants defining the table columns
    
    Const RoutineName As String = Module_Name & "BuildColumnConstants"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "Public Property Get %1Headers() As Variant", _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' First Headers
    Dim StartPoint As Long
    If This.DetailsDict.Count Mod 2 = 0 Then
        Line = PrintString( _
            "    %1Headers = Array( _" & vbCrLf & _
            "        qq%2qq, _" & vbCrLf, _
            This.TableName, This.DetailsDict.Items(0).ColumnHeader)
        StartPoint = 1
    Else
        Line = PrintString( _
            "    %1Headers = Array( _" & vbCrLf & _
            "        qq%2qq, qq%3qq, _" & vbCrLf, _
            This.TableName, This.DetailsDict.Items(0).ColumnHeader, This.DetailsDict.Items(1).ColumnHeader)
        StartPoint = 2
    End If
    ' Remaining Headers
    Dim I As Long
    For I = StartPoint To This.DetailsDict.Count - 2 Step 2
        Line = PrintString(Line & _
            "        qq%1qq, qq%2qq, _" & vbCrLf, _
            This.DetailsDict.Items(I).ColumnHeader, This.DetailsDict.Items(I + 1).ColumnHeader)
    Next I
    
    ' End the last Header
    If This.DetailsDict.Count Mod 1 = 0 Then
    Line = PrintString(Line & _
        "        qq%1qq)", _
        This.DetailsDict.Items(This.DetailsDict.Count - 1).ColumnHeader)
    Else
    Line = Line & "        )"
    End If
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Line = "End Property" & vbCrLf
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildColumnConstants

Private Sub BuildCheckExists()

    ' Builds the CheckExists routine
    
    Const RoutineName As String = Module_Name & "BuildCheckExists"
    On Error GoTo ErrorHandler
    
    Dim Entry As Variant
    Dim TD As TableDetails_Table
    Dim Found As Boolean
    Dim Key As String
    Dim KeyName As String
    For Each Entry In This.DetailsDict
        Set TD = This.DetailsDict(Entry)
        If TD.Key = "Key" Then
            Key = TD.VariableName
            KeyName = TD.ColumnHeader
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
        Key, This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Line = PrintString( _
        "    If %1 = vbNullString Then" & vbCrLf & _
        "        Check%1Exists = True" & vbCrLf & _
        "        Exit Function" & vbCrLf & _
        "    End If" & vbCrLf & _
        vbCrLf & _
        "    Check%1Exists = This.Dict.Exists(%1)" & vbCrLf, _
        Key, This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName

    Line = "Check" & Key & "Exists"
    BuildFunctionEnding Line, KeyName, Key
    
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

Private Sub BuildArrayToRecord()
    
    ' This routine builds a dictionary from an array
    
    Const RoutineName As String = Module_Name & "BuildArrayToRecord"
    On Error GoTo ErrorHandler
    
    Dim Ary As Variant
    BuildKeyArray This.DetailsDict, Ary
    
    Dim Line As String
    
    If UBound(Ary, 1) = 1 Or Ary(1) = "None" Then
        BuildNoneOrOneKey Ary(1)
    Else
        BuildMoreThanOneKey Ary
    End If
    
    Dim Entry As Variant
    For Each Entry In This.DetailsDict.Keys
        If This.DetailsDict.Item(Entry).VariableType = "Boolean" Then
            Line = PrintString( _
                "                Record.%1 = IIf(Ary(I, p%1Column) = qqYesqq , True,False)", _
                This.DetailsDict.Item(Entry).VariableName)
        Else
            Line = PrintString( _
                "                Record.%1 = Ary(I, p%1Column)", _
                This.DetailsDict.Item(Entry).VariableName)
        End If
        This.StreamFile.WriteMessageLine Line, This.StreamName
    Next Entry

    This.StreamFile.WriteBlankMessageLines This.StreamName
    
    Line = _
        "                Dict.Add Key, Record" & vbCrLf & _
        "            End If" & vbCrLf & _
        "        Next I" & vbCrLf & vbCrLf & _
        "    Else" & vbCrLf & _
        "        Dict.Add Ary, Ary" & vbCrLf & _
        "    End If" & vbCrLf
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildArrayToRecord

Private Sub BuildNoneOrOneKey(ByVal Key As String)

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
        "            If This.detailsdict.Exists(Key) Then" & vbCrLf & _
        "                ReportWarning qqDuplicate keyqq, qqRoutineqq, RoutineName, qqKeyqq, Key" & vbCrLf & _
        "                %3TryCopyArrayToDictionary = False" & vbCrLf & _
        "                GoTo Done" & vbCrLf & _
        "            Else" & vbCrLf & _
        "                Set Record = New " & This.ClassName & vbCrLf, _
        This.ClassName, Key, This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
        
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildNoneOrOneKey

Private Sub BuildMoreThanOneKey(ByVal Ary As Variant)

    ' Build the ArrayToDictionary code more than one key
    
    Const RoutineName As String = Module_Name & "BuildMoreThanOneKey"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "    Dim Key As String" & vbCrLf & _
        "    Dim Record as %1" & vbCrLf & vbCrLf & _
        "    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then" & vbCrLf & _
        "        For I = 1 To UBound(Ary, 1)" & vbCrLf, _
        This.ClassName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Line = PrintString("            Key = qq|qq _", Ary(1))
        
    Dim I As Long
    For I = 1 To UBound(Ary, 1) - 1
        Line = Line & PrintString(vbCrLf & "                & Ary(I, p%1Column) & qq|qq _", Ary(I))
    Next I
    
    Line = Line & PrintString(vbCrLf & "                & Ary(I, p%1Column) & qq|qq" & vbCrLf, Ary(UBound(Ary, 1)))
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Line = PrintString( _
        "            If Dict.Exists(Key) Then" & vbCrLf & _
        "                ReportWarning qqDuplicate keyqq, qqRoutineqq, RoutineName, qqKeyqq, Key" & vbCrLf & _
        "                %1TryCopyArrayToDictionary = False" & vbCrLf & _
        "                GoTo Done" & vbCrLf & _
        "            Else" & vbCrLf & _
        "                Set Record = New " & This.ClassName & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
        
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildMoreThanOneKey

Private Sub BuildFormatDetails()

    ' Build the format routine calls
    
    Const RoutineName As String = Module_Name & "BuildFormatDetails"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    Dim Entry As Variant
    Dim TD As TableDetails_Table
    Dim Vbl As String
    For Each Entry In This.DetailsDict
        Set TD = This.DetailsDict(Entry)
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
        
        If TD.Format <> vbNullString Then This.StreamFile.WriteMessageLine Line, This.StreamName
    Next Entry
    
    This.StreamFile.WriteBlankMessageLines This.StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildFormatDetails

Private Sub BuildPropertyGetRoutines()

    ' Build the Property Get routines
    
    Const RoutineName As String = Module_Name & "BuildPropertyGetRoutines"
    On Error GoTo ErrorHandler
    
    Dim Entry As Variant
    Dim FoundAKey As Boolean
    Dim Details As TableDetails_Table
    Dim Key As String
    Dim KeyType As String
    Dim KeyName As String
    For Each Entry In This.DetailsDict
        Set Details = This.DetailsDict(Entry)
        If Details.Key = "Key" Then
            Key = Details.VariableName
            KeyType = Details.VariableType
            KeyName = Details.ColumnHeader
            FoundAKey = True
            Exit For
        End If
    Next Entry
    
    If Not FoundAKey Then GoTo Done
    
    Dim Target As String ' The variable name to be fetched
    Dim TargetType As String
    For Each Entry In This.DetailsDict
        Set Details = This.DetailsDict(Entry)
        If Details.VariableName <> Key Then
            Target = Details.VariableName
            TargetType = Details.VariableType
            BuildOnePropertyGetRoutine Key, KeyType, KeyName, Target, TargetType
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
        ByVal Key As String, _
        ByVal KeyType As String, _
        ByVal KeyName As String, _
        ByVal Target As String, _
        ByVal TargetType As String)

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
        "        ReportError qqUnrecognized %2qq, _" & vbCrLf & _
        "            qqRoutineqq, RoutineName, _" & vbCrLf & _
        "            qq%6qq, %2" & vbCrLf & _
        "    End If" & vbCrLf, _
        Target, Key, KeyType, TargetType, This.TableName, KeyName)
        
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Line = PrintString("Get%1From%2", Target, Key)
    BuildPropertyEnding Line, KeyName, Key
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildOnePropertyGetRoutine

Private Sub BuildInitialize()

    ' Used for lower level routines
    
    Const RoutineName As String = Module_Name & "BuildInitialize"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "Public Sub %1Initialize()" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & qq%1Initializeqq" & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf & _
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
        This.TableName, This.ClassName)
     This.StreamFile.WriteMessageLine Line, This.StreamName

    BuildSubEnding This.TableName & "Initialize"
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildInitialize

Private Sub BuildArrayToDictionary()

    ' Build TryCopyArrayToDictionary
    
    Const RoutineName As String = Module_Name & "BuildArrayToDictionary"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
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
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName

    BuildArrayToRecord
    
    BuildFunctionEnding This.TableName & "TryCopyArrayToDictionary"
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildArrayToDictionary

Private Sub BuildFormatArrayAndWorksheet()

    ' Used for lower level routines
    
    Const RoutineName As String = Module_Name & "BuildFormatArrayAndWorksheet"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "Public Sub %1FormatArrayAndWorksheet( _" & vbCrLf & _
        "    ByRef Ary as Variant, _" & vbCrLf & _
        "    ByVal Table As ListObject)" & vbCrLf & _
        vbCrLf & _
        "    Const RoutineName As String = Module_Name & qq%1FormatArrayAndWorksheetqq" & vbCrLf & _
        "    On Error GoTo ErrorHandler" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    BuildFormatDetails

    BuildSubEnding This.TableName & "FormatArrayAndWorksheet"
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildFormatArrayAndWorksheet

Private Sub BuildSubEnding( _
    ByVal SubName As String, _
    ParamArray Args() As Variant)

    ' The standard end for subs
    
    Const RoutineName As String = Module_Name & "BuildSubEnding"
    On Error GoTo ErrorHandler
    
    BuildRoutineEnding SubName, "Sub", Args
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildSubEnding

Private Sub BuildFunctionEnding( _
    ByVal FunctionName As String, _
    ParamArray Args() As Variant)

    ' The standard end for functions
    
    Const RoutineName As String = Module_Name & "BuildFunctionEnding"
    On Error GoTo ErrorHandler
    
    BuildRoutineEnding FunctionName, "Function", Args

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildFunctionEnding

Private Sub BuildPropertyEnding( _
    ByVal SubName As String, _
    ParamArray Args() As Variant)

    ' The standard end for Property
    
    Const RoutineName As String = Module_Name & "BuildPropertyEnding"
    On Error GoTo ErrorHandler
    
    BuildRoutineEnding SubName, "Property", Args
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildPropertyEnding

Private Sub BuildRoutineEnding( _
    ByVal SubName As String, _
    ByVal RoutineType As String, _
    ParamArray Args() As Variant)

    ' Creates a standard routine ending
    
    Const RoutineName As String = Module_Name & "BuildRoutineEnding"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Debug.Assert UBound(Args(0), 1) Mod 2 <> 0
    
    If UBound(Args(0), 1) = 0 Then
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
        This.StreamFile.WriteMessageLine Line, This.StreamName
    Else
        Line = PrintString( _
            "Done:" & vbCrLf & _
            "    Exit %1" & vbCrLf & _
            "ErrorHandler:" & vbCrLf & _
            "    ReportError qqException raisedqq, _" & vbCrLf & _
            "                qqRoutineqq, RoutineName, _" & vbCrLf & _
            "                qqError Numberqq, Err.Number, _" & vbCrLf & _
            "                qqError Descriptionqq, Err.Description", _
            RoutineType)
            
            Dim I As Long
            For I = 0 To IIf(UBound(Args(0), 1) Mod 2 = 0, UBound(Args(0), 1) - 2, UBound(Args(0), 1) - 1) Step 2
                Line = Line & "' _" & vbCrLf & _
                    PrintString("                qq" & Args(0)(I) & "qq, " & Args(0)(I + 1))
            Next I
            Line = Line & vbCrLf
            This.StreamFile.WriteMessageLine Line, This.StreamName
            
        Line = PrintString( _
            "    RaiseError Err.Number, Err.Source, RoutineName, Err.Description" & vbCrLf & _
            "End %1 ' %2" & vbCrLf, _
            RoutineType, SubName)
            This.StreamFile.WriteMessageLine Line, This.StreamName
    End If
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildRoutineEnding

Private Sub BuildGetTable()

    ' Builds the Get Table routine
    
    Const RoutineName As String = Module_Name & "BuildGetTable"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    ' todo: check FileName for Excel file then build as below
    If This.FileName = vbNullString Then
        Line = PrintString( _
            "Public Property Get %1Table() As ListObject" & vbCrLf & _
            "    Set %1Table = %1Sheet.ListObjects(qq%1Tableqq)" & vbCrLf & _
            "End Property" & vbCrLf, _
            This.TableName)
    Else
        Line = PrintString( _
            "Public Property Get %1Table() As ListObject" & vbCrLf & _
            "'    Table not in this workbook" & vbCrLf & _
            "End Property" & vbCrLf, _
            This.TableName)
    End If
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildGetTable

'Public Property Get TestTable() As ListObject
'    Dim FileName As String
'    FileName = GetDataFilesFolder & Application.PathSeparator & WorkbookName
'
'    Set Wkbk = Workbooks.Open(FileName:=FileName, UpdateLinks:=0, ReadOnly:=True)
'
'    Dim Wksht As Worksheet
'    Set Wksht = Wkbk.Worksheets(WorksheetName)
'
'    Set TestTable = Wksht.ListObjects(TableName)
'End Property
'
'Public Sub TestInitialize()
'
'    Const RoutineName As String = Module_Name & "TestInitialize"
'    On Error GoTo ErrorHandler
'    Dim CalendarRates As CalendarRates_Table
'    Set CalendarRates = New CalendarRates_Table
'
'    Dim Dict As Dictionary
'    Set Dict = New Dictionary
'    If Table.TryCopyTableToDictionary(CalendarRates, TestTable, Dict) Then
'        Initialized = True
'    Else
'        ReportError "Error copying CalendarRates table", "Routine", RoutineName
'        Initialized = False
'        GoTo Done
'    End If
'
'Done:
'    Wkbk.Close
'    Exit Sub
'ErrorHandler:
'    Wkbk.Close
'    ReportError "Exception raised", _
'                "Routine", RoutineName, _
'                "Error Number", Err.Number, _
'                "Error Description", Err.Description
'    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
'End Sub ' TestInitialize
'
'
