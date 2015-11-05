Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Location
' Level:        Framework class
' Version:      1.00
'
' Description:  Location object related properties, Locations, functions & procedures
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:   -
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_LocationTYpe As String

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

Public Property Let LocationType(Value As String)
    m_LocationTYpe = Value
End Property

Public Property Get LocationType() As String
    LocationType = m_LocationTYpe
End Property