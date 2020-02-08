VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableDetails_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iTable

' Built on 2/8/2020 11:26:08 AM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private pColumnHeader As String
Private pVariableName As String
Private pVariableType As String
Private pKey As String
Private pFormat As String

Private pTableDetailsDict As Dictionary

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

Public Property Get ColumnHeader() As String
    ColumnHeader = pColumnHeader
End Property

Public Property Let ColumnHeader(ByVal Param As String)
    pColumnHeader = Param
End Property

Public Property Get VariableName() As String
    VariableName = pVariableName
End Property

Public Property Let VariableName(ByVal Param As String)
    pVariableName = Param
End Property

Public Property Get VariableType() As String
    VariableType = pVariableType
End Property

Public Property Let VariableType(ByVal Param As String)
    pVariableType = Param
End Property

Public Property Get Key() As String
    Key = pKey
End Property

Public Property Let Key(ByVal Param As String)
    pKey = Param
End Property

Public Property Get Format() As String
    Format = pFormat
End Property

Public Property Let Format(ByVal Param As String)
    pFormat = Param
End Property

Public Property Get iTable_LocalDictionary() As Dictionary
    Set iTable_LocalDictionary = TableDetailsDictionary
End Property

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = TableDetailsHeaderWidth
End Property

Public Property Get iTable_Headers() As Variant
    iTable_Headers = TableDetailsHeaders
End Property

Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = TableDetailsInitialized
End Property

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = TableDetailsTable
End Property

Public Property Get iTable_LocalName() As String
    iTable_LocalName = "TableDetails_Table"
End Property

Public Sub iTable_Initialize()
    TableDetailsInitialize
End Sub

Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = TableDetailsTryCopyArrayToDictionary(Ary, Dict)
End Function

Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean
    iTable_TryCopyDictionaryToArray = TableDetailsTryCopyDictionaryToArray(Dict, Ary)
End Function

Public Sub iTable_FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)
    TableDetailsFormatArrayAndWorksheet Ary, Table
End Sub ' FormatArrayAndWorksheet

''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                  '
'             End of Generated code                '
'            Start unique code here                '
'                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''

