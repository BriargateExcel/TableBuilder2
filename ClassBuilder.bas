Attribute VB_Name = "ClassBuilder"
Option Explicit
'@Folder "Builder"
Private Const Module_Name As String = "ClassBuilder."

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
' Application specific routines

Public Sub ClassBuilder( _
    ByVal DetailsDict As Dictionary, _
    ByVal BasicDict As Dictionary)

    ' This routine builds the class module

    Const RoutineName As String = Module_Name & "ClassBuilder"
    On Error GoTo ErrorHandler
    
    Set This.StreamFile = New MessageFileClass
    
    Set This.BasicDict = BasicDict
    
    Set This.DetailsDict = DetailsDict
    
    This.TableName = This.BasicDict.Items(0).TableName
    This.ClassName = This.TableName & "_Table"
    
    This.StreamName = This.ClassName & ".cls"
    
    This.FileName = This.BasicDict.Items(0).FileName
    
    Dim Line As String
    
    ' Declarations
    BuildFrontEnd

    ' Private variables
    Dim Entry As Variant
    
    Line = PrintString("Private Type %1Type", This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    For Each Entry In This.DetailsDict.Keys
        Line = PrintString( _
            "    %1 As %2", _
            This.DetailsDict.Item(Entry).VariableName, This.DetailsDict.Item(Entry).VariableType)
        This.StreamFile.WriteMessageLine Line, This.StreamName
    Next Entry
    
    Line = PrintString("End Type" & vbCrLf, This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Line = PrintString("Private This as %1Type" & vbCrLf, This.TableName)
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
        "    Set iTable_LocalDictionary = %1Dictionary" & vbCrLf & _
        "End Property" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' HeaderWidth
    Line = PrintString( _
        "Public Property Get iTable_HeaderWidth() As Long" & vbCrLf & _
        "    iTable_HeaderWidth = %1HeaderWidth" & vbCrLf & _
        "End Property" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Headers
    Line = PrintString( _
        "Public Property Get iTable_Headers() As Variant" & vbCrLf & _
        "    iTable_Headers = %1Headers" & vbCrLf & _
        "End Property" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Get Initialized
    Line = PrintString( _
        "Public Property Get iTable_Initialized() As Boolean" & vbCrLf & _
        "    iTable_Initialized = %1Initialized" & vbCrLf & _
        "End Property" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Initialize
    Line = PrintString( _
        "Public Sub iTable_Initialize()" & vbCrLf & _
        "    %1Initialize" & vbCrLf & _
        "End Sub" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Local Table
    If This.FileName = vbNullString Then
        Line = PrintString( _
            "Public Property Get iTable_LocalTable() As ListObject" & vbCrLf & _
            "    Set iTable_Localtable = %1Table" & vbCrLf & _
            "End Property" & vbCrLf, _
            This.TableName)
    Else
        Line = PrintString( _
            "Public Property Get iTable_LocalTable() As ListObject" & vbCrLf & _
            "End Property" & vbCrLf, _
            This.TableName)
    End If
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' Local Name
    Line = PrintString( _
        "Public Property Get iTable_LocalName() As String" & vbCrLf & _
        "    iTable_LocalName = qq%1qq" & vbCrLf & _
        "End Property" & vbCrLf, _
        This.ClassName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' TryCopyArrayToDictionary
    Line = PrintString( _
        "Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean" & vbCrLf & _
        "    iTable_TryCopyArrayToDictionary = %1TryCopyArrayToDictionary(Ary, Dict)" & vbCrLf & _
        "End Function" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' TryCopyDictionaryToArray
    Line = PrintString( _
        "Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean" & vbCrLf & _
        "    iTable_TryCopyDictionaryToArray = %1TryCopyDictionaryToArray(Dict, Ary)" & vbCrLf & _
        "End Function" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    ' FormatArrayAndWorksheet
    Line = PrintString( _
        "Public Sub iTable_FormatArrayAndWorksheet( _" & vbCrLf & _
        "    ByRef Ary as Variant, _" & vbCrLf & _
        "    ByVal Table As ListObject)" & vbCrLf & _
        "    %1FormatArrayAndWorksheet Ary, Table" & vbCrLf & _
        "End Sub ' FormatArrayAndWorksheet" & vbCrLf, _
        This.TableName)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
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
        "End Property" & vbCrLf, _
        Record.VariableName, Record.VariableType)
    This.StreamFile.WriteMessageLine Line, This.StreamName
    
    Line = PrintString( _
        "Public Property Let %1(ByVal Param as %2)" & vbCrLf & _
        "    This.%1 = Param" & vbCrLf & _
        "End Property" & vbCrLf, _
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

