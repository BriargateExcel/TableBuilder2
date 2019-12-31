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

' Built on 12/31/2019 12:09:32 PM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private pColumnHeader As String
Private pVariableName As String
Private pFormatted As Boolean
Private pVariableType As String

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

Public Property Get Formatted() As Boolean
    Formatted = pFormatted
End Property

Public Property Let Formatted(ByVal Param As Boolean)
    pFormatted = Param
End Property

Public Property Get VariableType() As String
    VariableType = pVariableType
End Property

Public Property Let VariableType(ByVal Param As String)
    pVariableType = Param
End Property

