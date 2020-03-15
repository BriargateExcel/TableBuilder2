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

' Built on 3/15/2020 10:39:08 AM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Type TableBasicsType
    TableName As String
    FileName As String
    WorksheetName As String
    ExternalTableName As String
End Type

Private This As TableBasicsType

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

Public Property Get iTable_LocalDictionary() As Dictionary
    Set iTable_LocalDictionary = TableBasicsDictionary
End Property

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = TableBasicsHeaderWidth
End Property

Public Property Get iTable_Headers() As Variant
    iTable_Headers = TableBasicsHeaders
End Property

Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = TableBasicsInitialized
End Property

Public Sub iTable_Initialize()
    TableBasicsInitialize
End Sub

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = TableBasicsTable
End Property

Public Property Get iTable_LocalName() As String
    iTable_LocalName = "TableBasics_Table"
End Property

Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = TableBasicsTryCopyArrayToDictionary(Ary, Dict)
End Function

Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean
    iTable_TryCopyDictionaryToArray = TableBasicsTryCopyDictionaryToArray(Dict, Ary)
End Function

Public Sub iTable_FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)
    TableBasicsFormatArrayAndWorksheet Ary, Table
End Sub ' FormatArrayAndWorksheet

' No application unique routines found

