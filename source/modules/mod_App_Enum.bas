Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_Enum
' Level:        Application module
' Version:      1.00
' Description:  enum functions & procedures specific to this application
'
' Source/date:  Bonnie Campbell, 11/5/2015
' Revisions:    BLC - 11/5/2015  - 1.00 - initial version
' =================================

' ---------------------------------
'   Enums
' ---------------------------------

Public Enum TaskType
    '[_First] = 1
    Site
    Feature
    Photo
    Transect
    Plot
    '[_Last] = 5
End Enum

Public Enum PhotoType
    Feature = 0
    Transect = 1
    Overview = 2
    Reference = 3
    'other
    Animals = 4
    Plants = 5
    Cultural = 6
    Scenic = 7
    Disturbance = 8
    Weather = 9
    Fieldwork = 10
    Other = 11
End Enum

Public Enum DirectionFacing
    US
    DS
    RR
    RL
End Enum

Public Enum TransducerType
    US
    DS
    Air
End Enum

Public Enum WentworthClassSize
    s
    FG
    MG
    CG
    SP
    LP
    SC
    LC
    B
    BED
End Enum

Public Enum River
    CAC
    CBC
    Green
    GAC
    GBC
    Gunnison
    Yampa
End Enum

Public Enum Park
    BLCA
    CANY
    DINO
End Enum

Public Enum SlopeChangeCause
    Debris
    Ground
    Rock
    Veg
    Water
End Enum

' Ashareef, August 22, 2014
' http://stackoverflow.com/questions/25445422/array-in-an-enumeration
Public Function CreateEnum()
    Dim db As Database
    Dim rs As Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("MYENUMS", dbOpenSnapshot)

    Dim m As Module
    Dim s As String
    Set m = Modules("myEnumsModule")

    s = "Option Compare Database"
    s = s & vbNewLine & "Option Explicit"
    s = s & vbNewLine
    s = s & vbNewLine & "Public Enum MyEnums"
    With rs
        Do Until .EOF
            s = s & vbNewLine & vbTab & .Fields("MYENUM") & " = " & rs.Fields("MYENUM_ID")
            .MoveNext
        Loop
    End With
    s = s & vbNewLine & "End Enum"

    Call m.DeleteLines(1, m.CountOfLines)
    Call m.AddFromString(s)
End Function

' ---------------------------------
' SUB:          UpdateContacts
' Description:  Update the contacts table from the linked table
' Parameters:
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, November 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 11/5/2015  - initial version
' ---------------------------------
Public Sub UpdateContacts()

On Error GoTo Err_Handler
    
    Dim df As DirectionFacing
    
    
    ' Chip Pearson, March 12, 2008
    ' http://www.cpearson.com/excel/Enums.aspx
    ' you can combine enums for combo types as in direction facing
    df = DS + RL
    
    PhotoType.Feature
    PhotoType.Overview

    
    
       
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateContacts[mod_Linked_Data])"
    End Select
    Resume Exit_Sub
End Sub