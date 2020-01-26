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

    Line = PrintString( _
        "VERSION 1.0 CLASS" & vbCrLf & _
        "BEGIN" & vbCrLf & _
        "  MultiUse = -1  'True" & vbCrLf & _
        "End" & vbCrLf & _
        "Attribute VB_Name = qq%1qq" & vbCrLf & _
        "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
        "Attribute VB_Creatable = False" & vbCrLf & _
        "Attribute VB_PredeclaredId = False" & vbCrLf & _
        "Attribute VB_Exposed = False" & vbCrLf & _
        "Option Explicit" & vbCrLf & _
        "Implements iTable" & vbCrLf & _
        vbCrLf & _
        "' Built on " & Now() & vbCrLf & _
        "' Built By Briargate Excel Table Builder" & vbCrLf & _
        "' See BriargateExcel.com for details" & vbCrLf, _
        ClassName)
        
    StreamFile.WriteMessageLine Line, StreamName, "Modules", True
    
    Dim Entry As Variant
    
    '
    ' Constants
    '

    For Each Entry In DetailsDict.Keys
        Line = PrintString( _
            "Private p%1 As %2", _
            DetailsDict.Item(Entry).VariableName, DetailsDict.Item(Entry).VariableType)
        StreamFile.WriteMessageLine Line, StreamName
    Next Entry
    StreamFile.WriteBlankMessageLines StreamName
    
    Line = PrintString( _
        "Private p%1Dict As Dictionary" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' Declaration separator
    '
    
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

    '
    ' Properties
    '

    For Each Entry In DetailsDict.Keys
        BuildProperties StreamFile, StreamName, DetailsDict.Item(Entry)
    Next Entry
            
    '
    ' Local Dictionary
    '
    
    Line = PrintString( _
        "Public Property Get iTable_LocalDictionary() As Dictionary" & vbCrLf & _
        "    Set iTable_LocalDictionary = %1Dictionary" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' HeaderWidth
    '

    Line = PrintString( _
        "Public Property Get iTable_HeaderWidth() As Long" & vbCrLf & _
        "    iTable_HeaderWidth = %1HeaderWidth" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' Headers
    '

    Line = PrintString( _
        "Public Property Get iTable_Headers() As Variant" & vbCrLf & _
        "    iTable_Headers = %1Headers" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' Get Initialized
    '

    Line = PrintString( _
        "Public Property Get iTable_Initialized() As Boolean" & vbCrLf & _
        "    iTable_Initialized = %1Initialized" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' Local Table
    '

    Line = PrintString( _
        "Public Property Get iTable_LocalTable() As ListObject" & vbCrLf & _
        "    Set iTable_Localtable = %1Table" & vbCrLf & _
        "End Property" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' Local Name
    '

    Line = PrintString( _
        "Public Property Get iTable_LocalName() As String" & vbCrLf & _
        "    iTable_LocalName = qq%1qq" & vbCrLf & _
        "End Property" & vbCrLf, _
        ClassName)
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' Initialize
    '

    Line = PrintString( _
        "Public Sub iTable_Initialize()" & vbCrLf & _
        "    %1Initialize" & vbCrLf & _
        "End Sub" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' TryCopyArrayToDictionary
    '

    Line = PrintString( _
        "Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean" & vbCrLf & _
        "    iTable_TryCopyArrayToDictionary = %1TryCopyArrayToDictionary(Ary, Dict)" & vbCrLf & _
        "End Function" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' TryCopyDictionaryToArray
    '

    Line = PrintString( _
        "Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean" & vbCrLf & _
        "    iTable_TryCopyDictionaryToArray = %1TryCopyDictionaryToArray(Dict, Ary)" & vbCrLf & _
        "End Function" & vbCrLf, _
        TableName)
    StreamFile.WriteMessageLine Line, StreamName
    
    '
    ' FormatArrayAndWorksheet
    '

    Line = PrintString( _
        "Public Sub iTable_FormatArrayAndWorksheet( _" & vbCrLf & _
        "    ByRef Ary as Variant, _" & vbCrLf & _
        "    ByVal Table As ListObject)" & vbCrLf & _
        "    %1FormatArrayAndWorksheet Ary, Table" & vbCrLf & _
        "End Sub ' FormatArrayAndWorksheet" & vbCrLf, _
        TableName)
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
End Sub ' ClassBuilder

Private Sub BuildProperties( _
        ByVal StreamFile As MessageFileClass, _
        ByVal StreamName As String, _
        ByVal Record As TableDetails_Table)

    ' This routine builds Get and Let Properties
    
    Const RoutineName As String = Module_Name & "BuildProperties"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "Public Property Get %1() as %2" & vbCrLf & _
        "    %1 = p%1" & vbCrLf & _
        "End Property" & vbCrLf, _
        Record.VariableName, Record.VariableType)
    StreamFile.WriteMessageLine Line, StreamName
    
    Line = PrintString( _
        "Public Property Let %1(ByVal Param as %2)" & vbCrLf & _
        "    p%1 = Param" & vbCrLf & _
        "End Property" & vbCrLf, _
        Record.VariableName, Record.VariableType)
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


