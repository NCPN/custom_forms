Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Photo
' Level:        Framework class
' Version:      1.00
'
' Description:  Photo object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
' =================================

'    [ID] [smallint] IDENTITY(1,1) NOT NULL,
'    [PhotographerID] [int] NULL,
'    [DownloadByID] [int] NULL,
'    [EntryByID] [int] NOT NULL,
'    [VerifyByID] [int] NULL,
'    [LastUpdateByID] [int] NOT NULL,
'    [PhotoType] [nvarchar](2) NOT NULL,
'    [PhotographerFacing] [nvarchar](2) NOT NULL,
'    [PhotographerLocation] [nvarchar](15) NOT NULL,
'    [SubjectLocation] [nvarchar](10) NULL,
'    [PhotoLabel] [nvarchar](8) NOT NULL,
'    [DigitalFilename] [nvarchar](15) NOT NULL,
'    [NCPNImageName] [nvarchar](15) NOT NULL,
'    [IsReplacement] [bit] NOT NULL,
'    [IsCloseup] [bit] NOT NULL,
'    [InActive] [bit] NOT NULL,
'    [TakenDate] [datetime] NOT NULL,
'    [DownloadDate] [datetime] NOT NULL,
'    [EntryDate] [timestamp] NOT NULL,
'    [VerifyDate] [datetime] NOT NULL,
'    [LastUpdate] [datetime] NOT NULL,

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_PhotoType As String
Private m_fileName As String
Private m_PhotographerLocation As Location
Private m_SubjectLocation As Location
Private m_Photographer As Person
Private m_Downloader As Person
Private m_Enterer As Person
Private m_Verifier As Person

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

Public Property Let PhotoType(Value As String)
    m_PhotoType = Value
End Property

Public Property Get PhotoType() As String
    PhotoType = m_PhotoType
End Property

Public Property Let fileName(Value As String)
    m_fileName = Value
End Property

Public Property Get fileName() As String
    fileName = m_fileName
End Property

Public Property Let PhotographerLocation(Value As Location)
    m_PhotographerLocation = Value
End Property

Public Property Get PhotographerLocation() As Location
    PhotographerLocation = m_PhotographerLocation
End Property

Public Property Let SubjectLocation(Value As Location)
    m_SubjectLocation = Value
End Property

Public Property Get SubjectLocation() As Location
    SubjectLocation = m_SubjectLocation
End Property