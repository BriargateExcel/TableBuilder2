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

' Built on 2/2/2020 10:28:50 AM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private pTableName As String

Private pTableBasicsDict As Dictionary

''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                  '
'   Start of application specific declarations     '
'                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                  '
'    End of application specific declarations      '
'                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TableName() As String
    TableName = pTableName
End Property

Public Property Let TableName(ByVal Param As String)
    pTableName = Param
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

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = TableBasicsTable
End Property

Public Property Get iTable_LocalName() As String
    iTable_LocalName = "TableBasics_Table"
End Property

Public Sub iTable_Initialize()
    TableBasicsInitialize
End Sub

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

''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                  '
'             End of Generated code                '
'            Start unique code here                '
'                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''

