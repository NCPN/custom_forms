Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Species
' Level:        Framework class
' Version:      1.00
'
' Description:  Species form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 10/30/2015
' References:   -
' Revisions:    BLC - 10/30/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_Name As String
Private m_COFamily As String
Private m_UTFamily As String
Private m_WYFamily As String
Private m_COName As String
Private m_UTName As String
Private m_WYName As String
Private m_LUCode As String 'lookup code

'---------------------
' Events
'---------------------
Public Event Selected()
Public Event Initialize()
Public Event Terminate()

'---------------------
' Properties
'---------------------
Public Property Let name(Value As String)
    m_Name = Value
End Property

Public Property Get name() As String
    name = m_Name
End Property

Public Property Let COFamily(Value As String)
    m_COFamily = Value
End Property

Public Property Get COFamily() As String
    COFamily = m_COFamily
End Property

Public Property Let UTFamily(Value As String)
    m_UTFamily = Value
End Property

Public Property Get UTFamily() As String
    UTFamily = m_UTFamily
End Property

Public Property Let WYFamily(Value As String)
    m_WYFamily = Value
End Property

Public Property Get WYFamily() As String
    WYFamily = m_WYFamily
End Property

Public Property Let COName(Value As String)
    m_COName = Value
End Property

Public Property Get COName() As String
    COName = m_COName
End Property

Public Property Let UTName(Value As String)
    m_UTName = Value
End Property

Public Property Get UTName() As String
    UTName = m_UTName
End Property

Public Property Let WYName(Value As String)
    m_WYName = Value
End Property

Public Property Get WYName() As String
    WYName = m_WYName
End Property

Public Property Let LUCode(Value As String)
    m_LUCode = Value
End Property

Public Property Get LUCode() As String
    LUCode = m_LUCode
End Property



'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Class_Initialize
' Description:  Class initialization (starting) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

    MsgBox "Initializing...", vbOKOnly


Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[Species class])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' Sub:          Class_Terminate
' Description:  Class termination (closing) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler
Exit_Sub:
    Exit Sub
    
    MsgBox "Terminating...", vbOKOnly
    
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[Species class])"
    End Select
    Resume Exit_Sub
End Sub


' ---------------------------------
' Sub:          SetHeaderColor
' Description:  Set header color event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub SetHeaderColor()
On Error GoTo Err_Handler
Exit_Sub:
    
    MsgBox "SetHeaderColor...", vbOKOnly
    
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[Species class])"
    End Select
    Resume Exit_Sub
End Sub