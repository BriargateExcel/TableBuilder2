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

' Built on 12/24/2019 10:55:28 AM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private pSheetName As String
Private pTableName As String
Private pClassName As String

Public Property Get SheetName() As String
    SheetName = pSheetName
End Property

Public Property Let SheetName(ByVal Param As String)
    pSheetName = Param
End Property

Public Property Get TableName() As String
    TableName = pTableName
End Property

Public Property Let TableName(ByVal Param As String)
    pTableName = Param
End Property

Public Property Get ClassName() As String
    ClassName = pClassName
End Property

Public Property Let ClassName(ByVal Param As String)
    pClassName = Param
End Property

Public Property Get iTable_Headers() As Variant
    iTable_Headers = TableBasics.Headers
End Property

Public Property Get iTable_LocalDictionary() As Dictionary
    Set iTable_LocalDictionary = TableBasicsDictionary
End Property

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = TableBasicsHeaderWidth
End Property

Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = TableBasicsInitialized
End Property

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = TableBasicsTable
End Property

Public Sub iTable_Initialize()
    TableDetails.Initialize
End Sub

Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = TableBasics.TryCopyArrayToDictionary(Ary, Dict)
End Function

Public Sub iTable_CopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant)
    TableBasics.CopyDictionaryToArray Dict, Ary
End Sub

Public Function TryCopyDictionaryToTable( _
       Optional ByVal Dict As Dictionary = Nothing, _
       Optional ByVal Table As ListObject = Nothing, _
       Optional TableCorner As Range = Nothing, _
       Optional TableName As String)

    TryCopyDictionaryToTable = TableBasics.TryCopyDictionaryToTable(Dict, Table, TableCorner, TableName)
End Function


