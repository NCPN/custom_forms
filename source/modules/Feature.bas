Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Feature
' Level:        Framework class
' Version:      1.00
'
' Description:  Feature object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_Name As String
Private m_Description As String
Private m_Directions As String
Private m_Sequence As Integer

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

'Public Property Let FeatureID(Value as integer)
'   m_FeatureID = Value
'End Property

'Public Property Get FeatureID()
'   FeatureID = m_FeatureID
'End Property

Public Property Let name(Value As String)
    m_Name = Value
End Property

Public Property Get name() As String
    name = m_Name
End Property

'Public Property Let Feature(Value as string)
'   m_Feature = Value
'End Property

'Public Property Get Feature()
'   Feature = m_Feature
'End Property

Public Property Let description(Value As String)
    m_Description = Value
End Property

Public Property Get description() As String
    description = m_Description
End Property

Public Property Let Directions(Value As String)
    m_Directions = Value
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let Sequence(Value As Integer)
    m_Sequence = Value
End Property

Public Property Get Sequence() As Integer
    Sequence = m_Sequence
End Property