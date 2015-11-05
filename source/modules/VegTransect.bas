Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        VegTransect
' Level:        Framework class
' Version:      1.00
'
' Description:  VegTransect object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_LocationID As Integer
Private m_EventID As Integer
Private m_Number As Integer
Private m_TransectType As String
Private m_SampleDate As Integer
Private m_ObserverID As Integer
Private m_RecorderID As Integer
Private m_Observer As String
Private m_Recorder As String

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

Public Property Let LocationID(Value As Integer)
    m_LocationID = Value
End Property

Public Property Get LocationID() As Integer
    LocationID = m_LocationID
End Property

Public Property Let EventID(Value As Integer)
    m_EventID = Value
End Property

Public Property Get EventID() As Integer
    EventID = m_EventID
End Property

Public Property Let Number(Value As Integer)
    m_Number = Value
End Property

Public Property Get Number() As Integer
    Number = m_Number
End Property

Public Property Let TransectType(Value As String)
    m_TransectType = Value
End Property

Public Property Get TransectType() As String
    TransectType = m_TransectType
End Property

Public Property Let SampleDate(Value As Integer)
    m_SampleDate = Value
End Property

Public Property Get SampleDate() As Integer
    SampleDate = m_SampleDate
End Property

Public Property Let ObserverID(Value As Integer)
    m_ObserverID = Value
End Property

Public Property Get ObserverID() As Integer
    ObserverID = m_ObserverID
End Property

Public Property Let Observer(Value As String)
    m_Observer = Value
End Property

Public Property Get Observer() As String
    Observer = m_Observer
End Property

Public Property Let RecorderID(Value As Integer)
    m_RecorderID = Value
End Property

Public Property Get RecorderID() As Integer
    RecorderID = m_RecorderID
End Property

Public Property Let Recorder(Value As String)
    m_Recorder = Value
End Property

Public Property Get Recorder() As String
    Recorder = m_Recorder
End Property