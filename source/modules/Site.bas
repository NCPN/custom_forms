Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Site
' Level:        Framework class
' Version:      1.00
'
' Description:  Site object related properties, events, functions & procedures
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
Private m_Code As String
Private m_Park As String
Private m_Description As String
Private m_Directions As String
Private m_SiteID As Integer
Private m_LocationID As Integer
Private m_ObserverID As Integer
Private m_RecorderID As Integer
Private m_Observer As String
Private m_Recorder As String
Private m_CommentID As Integer
Private m_Comment As String





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

'Public Property Let EventID(Value as integer)
'   m_EventID = Value
'End Property

'Public Property Get EventID()
'   EventID = m_EventID
'End Property


Public Property Let name(Value As String)
    m_Name = Value
End Property

Public Property Get name() As String
    name = m_Name
End Property

Public Property Let Code(Value As String)
    m_Code = Value
End Property

Public Property Get Code() As String
    Code = m_Code
End Property

Public Property Let Park(Value As String)
    m_Park = Value
End Property

Public Property Get Park() As String
    Park = m_Park
End Property

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

Public Property Let SiteID(Value As Integer)
    m_SiteID = Value
End Property

Public Property Get SiteID() As Integer
    SiteID = m_SiteID
End Property

Public Property Let LocationID(Value As Integer)
    m_LocationID = Value
End Property

Public Property Get LocationID() As Integer
    LocationID = m_LocationID
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

'---------------------
'change to comment object instead??
'---------------------
Public Property Let CommentID(Value As Integer)
    m_CommentID = Value
End Property

Public Property Get CommentID() As Integer
    CommentID = m_CommentID
End Property

Public Property Let Comment(Value As String)
    m_Comment = Value
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property


'---------------------
' Site
'---------------------
'---------------------
' Methods
'---------------------