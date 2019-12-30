VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get HeaderWidth() As Long
End Property

Public Property Get LocalDictionary() As Dictionary
End Property

Public Property Get LocalTable() As ListObject
End Property

Public Property Get Headers() As Variant
End Property

Public Property Get Initialized() As Boolean
End Property

Public Sub Initialize()
End Sub

Public Sub CopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant)
End Sub

Public Function TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
End Function

' An example of how to create a Class module that uses this interface:
'Option Explicit
'Implements iTable
'
'Private Const Module_Name As String = "CAMTable."
'
'Private pCAM As String
'Private pDeletionFlag As String
'
'Public Property Get CAM() As String
'    CAM = pCAM
'End Property
'
'Public Property Let CAM(ByVal Param As String)
'    pCAM = Param
'End Property
'
'Public Property Get DeletionFlag() As String
'    DeletionFlag = pDeletionFlag
'End Property
'
'Public Property Let DeletionFlag(ByVal Param As String)
'    pDeletionFlag = Param
'End Property
'
'Public Property Get iTable_Headers() As Variant
'    iTable_Headers = CAMHeaders
'End Property
'
'Public Property Get iTable_LocalDictionary() As Dictionary
'    Set iTable_LocalDictionary = CAMDictionary
'End Property
'
'Public Property Get iTable_HeaderWidth() As Long
'    iTable_HeaderWidth = CAMHeaderWidth
'End Property
'
'Public Property Get iTable_Initialized() As Boolean
'    iTable_Initialized = CAMInitialized
'End Property
'
'Public Property Get iTable_LocalTable() As ListObject
'    Set iTable_LocalTable = CAMTable
'End Property
'
'Public Sub iTable_Initialize()
'    CAMInitialize
'End Sub
'
'Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
'    iTable_TryCopyArrayToDictionary = TryCopyCAMArrayToDictionary(Ary, Dict)
'End Function
'
'Public Sub iTable_CopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant)
'    CopyCAMDictionaryToArray Dict, Ary
'End Sub
'
'Public Function TryCopyDictionaryToTable( _
'    Optional ByVal Dict As Dictionary = Nothing, _
'    Optional ByVal Table As ListObject = Nothing, _
'    Optional TableCorner As Range = Nothing, _
'    Optional TableName As String)
'
'    TryCopyDictionaryToTable = TryCopyCAMDictionaryToTable(Dict, Table, TableCorner, TableName)
'End Function
'

' An example of how to build a code module that supports this interface
'Public Property Get CAMHeaders() As Variant
'    CAMHeaders = Array("CAM", "Candidate for Deletion")
'End Property
'
'Public Property Get CAMDictionary() As Dictionary
'    Set CAMDictionary = pCAMDict
'End Property
'
'Public Property Get CAMHeaderWidth() As Long
'    CAMHeaderWidth = pHeaderWidth
'End Property
'
'Public Property Get CAMInitialized() As Boolean
'    CAMInitialized = pInitialized
'End Property
'
'Public Property Get CAMTable() As ListObject
'    Set CAMTable = CAMsSheet.ListObjects(pCAMTableName)
'End Property
'
'Public Sub CAMInitialize()
'
'    ' This routine loads the CAM dictionary
'
'    Const RoutineName As String = Module_Name & "Initialize"
'    On Error GoTo ErrorHandler
'
'    pInitialized = True
'
'    Dim CAM As CAMTable
'    Set CAM = New CAMTable
'    Set pCAMDict = New Dictionary
'    If TryCopyTableToDictionary(CAM, CAMTable, pCAMDict) Then
'        ' Success; do nothing
'    Else
'        ReportError "Error copying CAM table", "Routine", RoutineName
'        pInitialized = False
'        GoTo Done
'    End If
'
'Done:
'    Exit Sub
'ErrorHandler:
'    ReportError "Exception raised", _
'                "Routine", RoutineName, _
'                "Error Number", Err.Number, _
'                "Error Description", Err.Description
'    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
'End Sub ' Initialize
'
'Public Sub CopyCAMDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant)
'
'    ' loads Dict into Ary
'
'    Const RoutineName As String = Module_Name & "CopyCAMDictionaryToArray"
'    On Error GoTo ErrorHandler
'
'    Dim I As Long
'    I = 1
'
'    Dim Record As CAMTable
'    Dim Entry As Variant
'    For Each Entry In Dict.Keys
'        Set Record = Dict.Item(Entry)
'
'        Ary(I, pCAMColumn) = Record.CAM
'        Ary(I, pDeletionFlagColumn) = Record.DeletionFlag
'
'        I = I + 1
'    Next Entry
'
'Done:
'    Exit Sub
'ErrorHandler:
'    ReportError "Exception raised", _
'                "Routine", RoutineName, _
'                "Error Number", Err.Number, _
'                "Error Description", Err.Description
'    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
'End Sub ' CopyCAMDictionaryToArray
'
'Public Function TryCopyCAMArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary)
'
'    ' Copy CAM array to dictionary
'
'    Const RoutineName As String = Module_Name & "TryCopyCAMArrayToDictionary"
'    On Error GoTo ErrorHandler
'
'    TryCopyCAMArrayToDictionary = True
'
'    Dim I As Long
'
'    Set Dict = New Dictionary
'
'    Dim Record As CAMTable
'
'    For I = 1 To UBound(Ary, 1)
'        If Dict.Exists(Ary(I, 1)) Then
'            ReportError "Duplicate key", "Routine", RoutineName
'            TryCopyCAMArrayToDictionary = False
'            GoTo Done
'        Else
'            Set Record = New CAMTable
'
'            Record.CAM = Ary(I, pCAMColumn)
'            Record.DeletionFlag = Ary(I, pDeletionFlagColumn)
'
'            Dict.Add Ary(I, pCAMColumn), Record
'        End If
'    Next I
'Done:
'    Exit Function
'ErrorHandler:
'    ReportError "Exception raised", _
'                "Routine", RoutineName, _
'                "Error Number", Err.Number, _
'                "Error Description", Err.Description
'    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
'End Function ' TryCopyCAMArrayToDictionary
'
'Public Function TryCopyCAMDictionaryToTable( _
'    ByVal CAMDict As Dictionary, _
'    Optional ByVal Table As ListObject = Nothing, _
'    Optional TableCorner As Range = Nothing, _
'    Optional TableName As String _
'    ) As Boolean
'
'    ' This routine copies a CAM dictionary to a table
'    ' If Dict is nothing then use pCAMDict dictionary
'    ' If Table is nothing then build a table using TableCorner and TableName
'    ' if Table and TableCorner are both Nothing then use CAMTable
'
'    Const RoutineName As String = Module_Name & "TryCopyCAMDictionaryToTable"
'    On Error GoTo ErrorHandler
'
'    TryCopyCAMDictionaryToTable = True
'
'    If Not pInitialized Then CAMInitialize
'
'    Dim CAM As CAMTable
'    Set CAM = New CAMTable
'
'    If TryCopyDictionaryToTable(CAM, CAMDict, Table, TableCorner, TableName) Then
'        ' Success; do nothing
'    Else
'        ReportError "Error copying CAM dictionary to table", "Routine", RoutineName
'    End If
'
'
'Done:
'    Exit Function
'ErrorHandler:
'    ReportError "Exception raised", _
'                "Routine", RoutineName, _
'                "Error Number", Err.Number, _
'                "Error Description", Err.Description
'    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
'End Function ' TryCopyCAMDictionaryToTable
'





