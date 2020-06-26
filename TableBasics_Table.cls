VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableBasics_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iTable

' Built on 6/12/2020 3:16:22 PM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Type PrivateType
    TableName As String
    FileName As String
    WorksheetName As String
    ExternalTableName As String
    Skip As String
End Type ' PrivateType

Private This As PrivateType

' No application specific declarations found

Public Property Get TableName() As String
    TableName = This.TableName
End Property ' TableName

Public Property Let TableName(ByVal Param As String)
    This.TableName = Param
End Property ' TableName

Public Property Get FileName() As String
    FileName = This.FileName
End Property ' FileName

Public Property Let FileName(ByVal Param As String)
    This.FileName = Param
End Property ' FileName

Public Property Get WorksheetName() As String
    WorksheetName = This.WorksheetName
End Property ' WorksheetName

Public Property Let WorksheetName(ByVal Param As String)
    This.WorksheetName = Param
End Property ' WorksheetName

Public Property Get ExternalTableName() As String
    ExternalTableName = This.ExternalTableName
End Property ' ExternalTableName

Public Property Let ExternalTableName(ByVal Param As String)
    This.ExternalTableName = Param
End Property ' ExternalTableName

Public Property Get Skip() As String
    Skip = This.Skip
End Property ' Skip

Public Property Let Skip(ByVal Param As String)
    This.Skip = Param
End Property ' Skip

Public Property Get iTable_LocalDictionary() As Dictionary
    Set iTable_LocalDictionary = TableBasics.Dict
End Property ' LocalDictionary

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = TableBasics.HeaderWidth
End Property ' HeaderWidth

Public Property Get iTable_Headers() As Variant
    iTable_Headers = TableBasics.Headers
End Property ' Headers

Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = TableBasics.Initialized
End Property ' Initialized

Public Sub iTable_Initialize()
    TableBasics.Initialize
End Sub ' Initialize

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = TableBasics.SpecificTable
End Property ' LocalTable

Public Property Get iTable_LocalName() As String
    iTable_LocalName = "TableBasics_Table"
End Property ' LocalName

Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = TableBasics.TryCopyArrayToDictionary(Ary, Dict)
End Function ' TryCopyArrayToDictionary

Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean
    iTable_TryCopyDictionaryToArray = TableBasics.TryCopyDictionaryToArray(Dict, Ary)
End Function ' TryCopyDictionaryToArray

Public Sub iTable_FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)
    TableBasics.FormatArrayAndWorksheet Ary, Table
End Sub ' FormatArrayAndWorksheet

Public Property Get iTable_CreateKey(ByVal Record As iTable) As String
    iTable_CreateKey = TableBasics.CreateKey(Record)
End Property ' CreateKey

Public Property Get iTable_IsDatabase() As Boolean
    iTable_IsDatabase = False
End Property ' IsDictionary

Public Property Get iTable_DatabaseName() As String
    iTable_DatabaseName = ""
End Property ' DatabaseName

Public Property Get iTable_DatabaseTableName() As String
    iTable_DatabaseTableName = "TableBasics"
End Property ' DatabaseTableName

' No application unique routines found

