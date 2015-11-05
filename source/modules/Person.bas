Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Person
' Level:        Framework class
' Version:      1.00
'
' Description:  Person object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_FirstName As String
Private m_LastName As String
Private m_Name As String
Private m_Email As String
Private m_Role As String

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Integer)
    m_ID = Value
End Property

Public Property Get ID() As Integer
    ID = m_ID
End Property

Public Property Let FirstName(Value As String)
    m_FirstName = Value
End Property

Public Property Get FirstName() As String
    FirstName = m_FirstName
End Property

Public Property Let LastName(Value As String)
    m_LastName = Value
End Property

Public Property Get LastName() As String
    LastName = m_LastName
End Property

Public Property Let name(Value As String)
    m_Name = Value
End Property

Public Property Get name() As String
    name = m_Name
End Property

Public Property Let Email(Value As String)
    m_Email = Value
End Property

Public Property Get Email() As String
    Email = m_Email
End Property

Public Property Let Role(Value As String)
    Select Case Value
        Case "Observer"
        Case "Recorder"
        Case "DataEntry"
        Case "DataVerify"
        Case "PhotoDownload"
        Case "Photographer"
        Case "DataCertify"
    End Select
    m_Role = Value
End Property

Public Property Get Role() As String
    Role = m_Role
End Property