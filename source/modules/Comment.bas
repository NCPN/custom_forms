Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Comment
' Level:        Framework class
' Version:      1.00
'
' Description:  Comment object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_Comment As String
Private m_MaxLength As Integer

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Integer)
    If IsNumeric(Value) Then
        m_ID = Value
    End If
End Property

Public Property Get ID() As Integer
    ID = m_ID
End Property

Public Property Let Comment(Value As String)
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_Comment = Value
    End If
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let MaxLength(Value As Integer)
    If IsNumeric(Value) Then
        m_MaxLength = Value
    End If
End Property

Public Property Get MaxLength() As Integer
    MaxLength = m_MaxLength
End Property

'---------------------
' Methods
'---------------------