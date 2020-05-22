Attribute VB_Name = "ClassBuilder"
Option Explicit
'@Folder "Builder"
Private Const Module_Name As String = "ClassBuilder."

Private Type PrivateData
    StreamFile As MessageFileClass
    StreamName As String
    TableName As String
    ClassName As String
    FileName As String
    WorksheetName As String
    ExternalTableName As String
    DetailsDict As Dictionary
    BasicDict As Dictionary
End Type ' PrivateData

Private This As PrivateData

' Order:
' Front end
' Private variables
' Application specific declarations
' Properties
' Local dictionary
' Headerwidth
' Headers
' Get Initialized
' Initialize
' Local table
' Local Name
' Array to Dictionary
' Dictionary to Array
' Format array and worksheet
' CreateKey
' Application specific routines

Public Sub ClassBuilder( _
    ByVal DetailsDict As Dictionary, _
    ByVal BasicDict As Dictionary)

    ' This routine builds the class module

    Const RoutineName As String = Module_Name & "ClassBuilder"
    On Error GoTo ErrorHandler
    
    ' Load PrivateType
    Set This.StreamFile = New MessageFileClass
    
    Set This.BasicDict = BasicDict
    
    Set This.DetailsDict = DetailsDict
    
    This.TableName = This.BasicDict.Items(0).TableName
    This.ClassName = This.TableName & "_Table"
    
    This.StreamName = This.ClassName & ".cls"
    
    This.FileName = This.BasicDict.Items(0).FileName
    ' End of loading PrivateType
    
    Dim Line As String
    
    ' Declarations
    BuildFrontEnd

    ' Private variables
    Dim Entry As Variant
    
    Line = PrintString("Private Type PrivateType")
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    For Each Entry In This.DetailsDict.Keys
        Line = PrintString( _
            "    %1 As %2", _
            This.DetailsDict.Item(Entry).VariableName, This.DetailsDict.Item(Entry).VariableType)
        This.StreamFile.WriteMessageLine Line, This.StreamName
    Next Entry
    
    Line = PrintString("End Type ' PrivateType" & vbCrLf, This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Line = PrintString("Private This as PrivateType" & vbCrLf)
    This.StreamFile.WriteMessageLine Line, This.StreamName
            
    ' Application specific declarations
    BuildApplicationUniqueDeclarations This.StreamFile, This.StreamName, This.TableName, ".cls"
        
    ' Properties
    For Each Entry In DetailsDict.Keys
        BuildProperties This.DetailsDict.Item(Entry)
    Next Entry
            
    ' Local Dictionary
    Line = PrintString( _
        "Public Property Get iTable_LocalDictionary() As Dictionary" & vbCrLf & _
        "    Set iTable_LocalDictionary = %1.Dict" & vbCrLf & _
        "End Property ' LocalDictionary" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' HeaderWidth
    Line = PrintString( _
        "Public Property Get iTable_HeaderWidth() As Long" & vbCrLf & _
        "    iTable_HeaderWidth = %1.HeaderWidth" & vbCrLf & _
        "End Property ' HeaderWidth" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Headers
    Line = PrintString( _
        "Public Property Get iTable_Headers() As Variant" & vbCrLf & _
        "    iTable_Headers = %1.Headers" & vbCrLf & _
        "End Property ' Headers" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Get Initialized
    Line = PrintString( _
        "Public Property Get iTable_Initialized() As Boolean" & vbCrLf & _
        "    iTable_Initialized = %1.Initialized" & vbCrLf & _
        "End Property ' Initialized" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Initialize
    Line = PrintString( _
        "Public Sub iTable_Initialize()" & vbCrLf & _
        "    %1.Initialize" & vbCrLf & _
        "End Sub ' Initialize" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Local Table
    Line = PrintString( _
        "Public Property Get iTable_LocalTable() As ListObject" & vbCrLf & _
        "    Set iTable_Localtable = %1.SpecificTable" & vbCrLf & _
        "End Property ' LocalTable" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Local Name
    Line = PrintString( _
        "Public Property Get iTable_LocalName() As String" & vbCrLf & _
        "    iTable_LocalName = qq%1qq" & vbCrLf & _
        "End Property ' LocalName" & vbCrLf, _
        This.ClassName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' TryCopyArrayToDictionary
    Line = PrintString( _
        "Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean" & vbCrLf & _
        "    iTable_TryCopyArrayToDictionary = %1.TryCopyArrayToDictionary(Ary, Dict)" & vbCrLf & _
        "End Function ' TryCopyArrayToDictionary" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' TryCopyDictionaryToArray
    Line = PrintString( _
        "Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean" & vbCrLf & _
        "    iTable_TryCopyDictionaryToArray = %1.TryCopyDictionaryToArray(Dict, Ary)" & vbCrLf & _
        "End Function ' TryCopyDictionaryToArray" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' FormatArrayAndWorksheet
    Line = PrintString( _
        "Public Sub iTable_FormatArrayAndWorksheet( _" & vbCrLf & _
        "    ByRef Ary as Variant, _" & vbCrLf & _
        "    ByVal Table As ListObject)" & vbCrLf & _
        "    %1.FormatArrayAndWorksheet Ary, Table" & vbCrLf & _
        "End Sub ' FormatArrayAndWorksheet" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' CreateKey
    Line = PrintString( _
        "Public Property Get iTable_CreateKey(ByVal Record As iTable) As String" & vbCrLf & _
        "    iTable_CreateKey = %1.CreateKey(Record)" & vbCrLf & _
        "End Property ' CreateKey" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' IsDictionary function
    Dim DatabaseExists As String
    DatabaseExists = IIf(Right(This.FileName, 6) = ".accdb", "True", "False")
    
    Line = PrintString( _
        "Public Property Get iTable_IsDatabase() As Boolean" & vbCrLf & _
        "    iTable_IsDatabase = %1" & vbCrLf & _
        "End Property ' IsDictionary" & vbCrLf, _
        DatabaseExists)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Database Name
    Line = PrintString( _
        "Public Property Get iTable_DatabaseName() As String" & vbCrLf & _
        "    iTable_DatabaseName = qq%1qq" & vbCrLf & _
        "End Property ' DatabaseName" & vbCrLf, _
        This.FileName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Database Table Name
    Line = PrintString( _
        "Public Property Get iTable_DatabaseTableName() As String" & vbCrLf & _
        "    iTable_DatabaseTableName = qq%1qq" & vbCrLf & _
        "End Property ' DatabaseTableName" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Add any unique code
    BuildApplicationUniqueRoutines This.StreamFile, This.StreamName, This.TableName, ".cls"
    
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
End Sub ' ClassBuilder


Private Sub BuildFrontEnd()

    ' Builds the front matter
    
    Const RoutineName As String = Module_Name & "BuildFrontEnd"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
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
        This.ClassName)
        
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

Private Sub BuildProperties(ByVal Record As TableDetails_Table)

    ' This routine builds Get and Let Properties
    
    Const RoutineName As String = Module_Name & "BuildProperties"
    On Error GoTo ErrorHandler
    
    Dim Line As String
    
    Line = PrintString( _
        "Public Property Get %1() as %2" & vbCrLf & _
        "    %1 = This.%1" & vbCrLf & _
        "End Property ' %1" & vbCrLf, _
        Record.VariableName, Record.VariableType)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Line = PrintString( _
        "Public Property Let %1(ByVal Param as %2)" & vbCrLf & _
        "    This.%1 = Param" & vbCrLf & _
        "End Property ' %1" & vbCrLf, _
        Record.VariableName, Record.VariableType)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildProperties

