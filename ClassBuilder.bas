Attribute VB_Name = "ClassBuilder"
Option Explicit

Private Const Module_Name As String = "ClassBuilder."

Private Const Quote As String = """"

Public Sub ClassBuilder( _
    ByVal DetailsDict As Dictionary, _
    ByVal TableName As String, _
    ByVal ClassName As String)

    ' This routine builds the class module

    Const RoutineName As String = Module_Name & "ClassBuilder"
    On Error GoTo ErrorHandler
    
    Dim StreamName As String
    StreamName = ClassName & ".cls"
    
    Dim StreamFile As MessageFileClass
    Set StreamFile = New MessageFileClass
    
    Dim Line As String
    
    Line = _
        "VERSION 1.0 CLASS" & vbCrLf & _
        "BEGIN" & vbCrLf & _
        "  MultiUse = -1  'True" & vbCrLf & _
        "End" & vbCrLf & _
        "Attribute VB_Name = " & Quote & ClassName & Quote & vbCrLf & _
        "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
        "Attribute VB_Creatable = False" & vbCrLf & _
        "Attribute VB_PredeclaredId = False" & vbCrLf & _
        "Attribute VB_Exposed = False" & vbCrLf & _
        "Option Explicit" & vbCrLf & _
        "Implements iTable" & vbCrLf & _
        vbCrLf & _
        "' Built on " & Now() & vbCrLf & _
        "' Built By Briargate Excel Table Builder" & vbCrLf & _
        "' See BriargateExcel.com for details" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName, "Class Modules"
    
    Dim Entry As Variant
    
    For Each Entry In DetailsDict.Keys
        Line = "Private p" & DetailsDict.Item(Entry).VariableName & " As " & DetailsDict.Item(Entry).VariableType
        StreamFile.WriteMessageLine Line, StreamName
    Next Entry
    
    StreamFile.WriteBlankMessageLines StreamName
    
    For Each Entry In DetailsDict.Keys
        BuildProperties StreamFile, StreamName, DetailsDict.Item(Entry)
    Next Entry
    
    Line = _
        "Public Property Get iTable_Headers() As Variant" & vbCrLf & _
        "    iTable_Headers = " & TableName & ".Headers" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Get iTable_LocalDictionary() As Dictionary" & vbCrLf & _
        "    Set iTable_LocalDictionary = " & TableName & "Dictionary" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Get iTable_HeaderWidth() As Long" & vbCrLf & _
        "    iTable_HeaderWidth = " & TableName & "HeaderWidth" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Get iTable_Initialized() As Boolean" & vbCrLf & _
        "    iTable_Initialized = " & TableName & "Initialized" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Get iTable_LocalTable() As ListObject" & vbCrLf & _
        "    Set iTable_LocalTable = " & TableName & "Table" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Sub iTable_Initialize()" & vbCrLf & _
        "    TableDetails.Initialize" & vbCrLf & _
        "End Sub" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean" & vbCrLf & _
        "    iTable_TryCopyArrayToDictionary = " & TableName & ".TryCopyArrayToDictionary(Ary, Dict)" & vbCrLf & _
        "End Function" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Sub iTable_CopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant)" & vbCrLf & _
        "    " & TableName & ".CopyDictionaryToArray Dict, Ary" & vbCrLf & _
        "End Sub" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Function TryCopyDictionaryToTable( _" & vbCrLf & _
        "       Optional ByVal Dict As Dictionary = Nothing, _" & vbCrLf & _
        "       Optional ByVal Table As ListObject = Nothing, _" & vbCrLf & _
        "       Optional TableCorner As Range = Nothing, _" & vbCrLf & _
        "       Optional TableName As String)" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "    TryCopyDictionaryToTable = " & TableName & ".TryCopyDictionaryToTable(Dict, Table, TableCorner, TableName)" & vbCrLf & _
        "End Function" & vbCrLf
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
End Sub ' ClassBuilder

Private Sub BuildProperties( _
        ByVal StreamFile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal Record As TableDetails_Table)

    ' This routine builds Get and Let Properties
    
    Const RoutineName As String = Module_Name & "BuildProperties"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = _
        "Public Property Get " & Record.VariableName & "() as " & Record.VariableType & vbCrLf & _
        "    " & Record.VariableName & " = p" & Record.VariableName & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Let " & Record.VariableName & "(ByVal Param as " & Record.VariableType & ")" & vbCrLf & _
        "    p" & Record.VariableName & " = Param" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildProperties


