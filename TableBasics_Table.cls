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

' Built on 3/23/2020 11:18:26 AM
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
End Property

Public Property Let TableName(ByVal Param As String)
    This.TableName = Param
End Property

Public Property Get FileName() As String
    FileName = This.FileName
End Property

Public Property Let FileName(ByVal Param As String)
    This.FileName = Param
End Property

Public Property Get WorksheetName() As String
    WorksheetName = This.WorksheetName
End Property

Public Property Let WorksheetName(ByVal Param As String)
    This.WorksheetName = Param
End Property

Public Property Get ExternalTableName() As String
    ExternalTableName = This.ExternalTableName
End Property

Public Property Let ExternalTableName(ByVal Param As String)
    This.ExternalTableName = Param
End Property

Public Property Get Skip() As String
    Skip = This.Skip
End Property

Public Property Let Skip(ByVal Param As String)
    This.Skip = Param
End Property

Public Property Get iTable_LocalDictionary() As Dictionary
    Set iTable_LocalDictionary = TableBasics.Dict
End Property

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = TableBasics.HeaderWidth
End Property

Public Property Get iTable_Headers() As Variant
    iTable_Headers = TableBasics.Headers
End Property

Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = TableBasics.Initialized
End Property

Public Sub iTable_Initialize()
    TableBasics.Initialize
End Sub

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = TableBasics.SpecificTable
End Property

Public Property Get iTable_LocalName() As String
    iTable_LocalName = "TableBasics_Table"
End Property

Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = TableBasics.TryCopyArrayToDictionary(Ary, Dict)
End Function

Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean
    iTable_TryCopyDictionaryToArray = TableBasics.TryCopyDictionaryToArray(Dict, Ary)
End Function

Public Sub iTable_FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)
    TableBasics.FormatArrayAndWorksheet Ary, Table
End Sub ' FormatArrayAndWorksheet

Public Property Get iTable_CreateKey(ByVal Record As iTable) As String
    iTable_CreateKey = TableBasics.CreateKey(Record)
End Property

' No application unique routines found

