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
    
    '
    ' Declarations
    '

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
    StreamFile.WriteMessageLine Line, StreamName, "Modules", True
    
    Dim Entry As Variant
    
    '
    ' Constants
    '

    For Each Entry In DetailsDict.Keys
        Line = "Private p" & DetailsDict.Item(Entry).VariableName & " As " & DetailsDict.Item(Entry).VariableType
        StreamFile.WriteMessageLine Line, StreamName
    Next Entry
    StreamFile.WriteBlankMessageLines StreamName
    
    '
    ' Added for headcount tool
    '

    Line = _
        "Private pEntry As String ' Added for headcount tool" & vbCrLf & _
        "Private pRecord As iTable ' Added for headcount tool" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' Properties
    '

    For Each Entry In DetailsDict.Keys
        BuildProperties StreamFile, StreamName, DetailsDict.Item(Entry)
    Next Entry
        
    '
    ' Local Dictionary
    '

    Line = _
        "Public Property Get iTable_LocalDictionary() As Dictionary" & vbCrLf & _
        "    Set iTable_LocalDictionary = " & TableName & "Dictionary" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' HeaderWidth
    '

    Line = _
        "Public Property Get iTable_HeaderWidth() As Long" & vbCrLf & _
        "    iTable_HeaderWidth = " & TableName & "HeaderWidth" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' Headers
    '

    Line = _
        "Public Property Get iTable_Headers() As Variant" & vbCrLf & _
        "    iTable_Headers = " & TableName & "Headers" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' Get Initialized
    '

    Line = _
        "Public Property Get iTable_Initialized() As Boolean" & vbCrLf & _
        "    iTable_Initialized = " & TableName & "Initialized" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' Local Table
    '

    Line = _
        "Public Property Get iTable_LocalTable() As ListObject" & vbCrLf & _
        "    Set iTable_Localtable = " & TableName & "Table" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' Initialize
    '

    Line = _
        "Public Sub iTable_Initialize()" & vbCrLf & _
        "    " & TableName & "Initialize" & vbCrLf & _
        "End Sub" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' TryCopyArrayToDictionary
    '

    Line = _
        "Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean" & vbCrLf & _
        "    iTable_TryCopyArrayToDictionary = " & TableName & "TryCopyArrayToDictionary(Ary, Dict)" & vbCrLf & _
        "End Function" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' TryCopyDictionaryToArray
    '

    Line = _
        "Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean" & vbCrLf & _
        "    iTable_TryCopyDictionaryToArray = " & TableName & "TryCopyDictionaryToArray(Dict, Ary)" & vbCrLf & _
        "End Function" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' FormatWorksheet
    '

    Line = _
        "Public Sub iTable_FormatWorksheet(ByVal Table As ListObject)" & vbCrLf & _
        "    " & TableName & "FormatWorksheet Table" & vbCrLf & _
        "End Sub" & vbCrLf
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
    
    Line = _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "'          Start of headcount unique code          '" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Get iTable_EmployeeNumber(ByVal Entry As String) As String" & vbCrLf & _
        "    If Entry <> pEntry Then Set pRecord = " & TableName & "Dictionary.Item(Entry)" & vbCrLf & _
        "    iTable_EmployeeNumber = pRecord.EmployeeNumber" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Get iTable_Month(ByVal Entry As String) As Date" & vbCrLf & _
        "    If Entry <> pEntry Then Set pRecord = " & TableName & "Dictionary.Item(Entry)" & vbCrLf & _
        "    iTable_Month = pRecord.Month" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Get iTable_ControlAccount(ByVal Entry As String) As String" & vbCrLf & _
        "    If Entry <> pEntry Then Set pRecord = " & TableName & "Dictionary.Item(Entry)" & vbCrLf & _
        "    iTable_ControlAccount = pRecord.ControlAccount" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Get iTable_ChargeNumber(ByVal Entry As String) As String" & vbCrLf & _
        "    If Entry <> pEntry Then Set pRecord = " & TableName & "Dictionary.Item(Entry)" & vbCrLf & _
        "    iTable_ChargeNumber = pRecord.ChargeNumber" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Get iTable_BudEPs(ByVal Entry As String) As Single" & vbCrLf & _
        "    If Entry <> pEntry Then Set pRecord = " & TableName & "Dictionary.Item(Entry)" & vbCrLf & _
        "    iTable_BudEPs = pRecord.BudEPs" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Get iTable_EstEPs(ByVal Entry As String) As Single" & vbCrLf & _
        "    If Entry <> pEntry Then Set pRecord = " & TableName & "Dictionary.Item(Entry)" & vbCrLf & _
        "    iTable_EstEPs = pRecord.EstEPs" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "Public Property Get iTable_ActHrs(ByVal Entry As String) As Single" & vbCrLf & _
        "    If Entry <> pEntry Then Set pRecord = " & TableName & "Dictionary.Item(Entry)" & vbCrLf & _
        "    iTable_ActHrs = pRecord.ActHrs" & vbCrLf & _
        "End Property" & vbCrLf
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = _
        "''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
        "'                                                  '" & vbCrLf & _
        "'             End of headcount unique code         '" & vbCrLf & _
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


