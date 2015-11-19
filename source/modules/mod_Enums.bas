Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Enums
' Level:        Application module
' Version:      1.00
' Description:  enum functions & procedures specific to this application
'
' Source/date:  Bonnie Campbell, 11/5/2015
' Revisions:    BLC - 11/5/2015  - 1.00 - initial version
' =================================

'-----------------------------
'  Functions
'-----------------------------

' ---------------------------------
' FUNCTION:     CreateEnums
' Description:  Create application specific enums based on Enum table
' Notes:
' you can combine enums for combo types as in direction facing
'    Dim df As DirectionFacing
'    df = DS + RL
'
'    PhotoType.Feature
'    PhotoType.Overview
'
' for more information see the following reference
'   Chip Pearson, March 12, 2008
'   http://www.cpearson.com/excel/Enums.aspx
' Parameters:
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Ashareef, August 22, 2014
' http://stackoverflow.com/questions/25445422/array-in-an-enumeration
' Adapted:      Bonnie Campbell, November 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 11/5/2015  - initial version
' ---------------------------------
Public Function CreateEnums(Optional EnumType As String)
On Error GoTo Err_Handler
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("Enum", dbOpenSnapshot)

    Dim m As Module
    Dim s As String, PrevEnumType As String, strEnumType As String
    Set m = Modules("mod_App_Enum")

    'clear current enums
    Call m.DeleteLines(1, m.CountOfLines)

    PrevEnumType = ""
    
    s = "Option Compare Database"
    s = s & vbNewLine & "Option Explicit"
    s = s & vbNewLine
    
    s = s & vbNewLine & "' ================================="
    s = s & vbNewLine & "' MODULE:       mod_App_Enum"
    s = s & vbNewLine & "' Level:        Application module"
    s = s & vbNewLine & "' Version:      1.00"
    s = s & vbNewLine & "' Description:  enum functions & procedures specific to this application"
    s = s & vbNewLine & "'"
    s = s & vbNewLine & "' Source/date:  Bonnie Campbell, 11/5/2015"
    s = s & vbNewLine & "' Revisions:    BLC - 11/5/2015  - 1.00 - initial version"
    s = s & vbNewLine & "' =================================" & vbNewLine
    
    With rs
    
        .Sort = "EnumType, ID"
        
        Do Until .EOF
            
            'handle first enum
            If .Fields("EnumType") <> PrevEnumType Then
                
                'handle plurals
                If Right(.Fields("EnumType"), 1) = "s" Or Right(.Fields("EnumType"), 1) = "y" Then
                    strEnumType = .Fields("EnumType")
                Else
                    strEnumType = .Fields("EnumType") & "s"
                End If
            
                s = s & vbNewLine & "'-----------------------------"
                s = s & vbNewLine & "'  " & strEnumType
                s = s & vbNewLine & "'-----------------------------"
                
                s = s & vbNewLine & "Public Enum " & .Fields("EnumType")
                If PrevEnumType = "" Then PrevEnumType = .Fields("EnumType")
            End If
            
            s = s & vbNewLine & vbTab & .Fields("Label") & " = " & rs.Fields("ID")
            .MoveNext
            
            If Not .EOF Then
                If .Fields("EnumType") <> PrevEnumType Then
                    s = s & vbNewLine & "End Enum" & vbNewLine
                    PrevEnumType = .Fields("EnumType")
                    
                    'handle plurals
                    If Right(.Fields("EnumType"), 1) = "s" Or Right(.Fields("EnumType"), 1) = "y" Then
                        strEnumType = .Fields("EnumType")
                    Else
                        strEnumType = .Fields("EnumType") & "s"
                    End If
                    
                    'handle remaining enums
                    s = s & vbNewLine & "'-----------------------------"
                    s = s & vbNewLine & "'  " & strEnumType
                    s = s & vbNewLine & "'-----------------------------"
                    
                    s = s & vbNewLine & "Public Enum " & .Fields("EnumType")
                End If
            End If
        
        Loop
        s = s & vbNewLine & "End Enum"
    End With
    
    'Call m.DeleteLines(1, m.CountOfLines)
    Call m.AddFromString(s)

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateEnums[mod_Enums])"
    End Select
    Resume Exit_Function
End Function