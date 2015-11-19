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

'-----------------------------
'  SlopeChangeCauses
'-----------------------------
Public Enum SlopeChangeCause
    Debris = 55
    Ground = 56
    Rock = 57
    Veg = 58
    Water = 59
End Enum

'-----------------------------
'  PhotoTypes
'-----------------------------
Public Enum PhotoType
    Feature = 1
    Transect = 2
    Overview = 3
    Reference = 4
    Animals = 5
    Plants = 6
    Cultural = 7
    Scenic = 8
    Disturbance = 9
    Weather = 10
    Fieldwork = 11
    Other = 12
End Enum

'-----------------------------
'  DirectionFacings
'-----------------------------
Public Enum DirectionFacing
    US = 13
    DS = 14
    RR = 15
    RL = 16
End Enum

'-----------------------------
'  TransducerTypes
'-----------------------------
Public Enum TransducerType
    US = 17
    DS = 18
    Air = 19
End Enum

'-----------------------------
'  Rivers
'-----------------------------
Public Enum River
    CAC = 20
    CBC = 21
    Green = 22
    GAC = 23
    GBC = 24
    Gunnison = 25
    Yampa = 26
End Enum

'-----------------------------
'  WentworthClassSizes
'-----------------------------
Public Enum WentworthClassSize
    s = 27
    FG = 28
    MG = 29
    CG = 30
    SP = 31
    LP = 32
    SC = 33
    LC = 34
    B = 35
    BED = 36
End Enum

'-----------------------------
'  TaskTypes
'-----------------------------
Public Enum TaskType
    Site = 37
    Feature = 38
    Photo = 39
    Transect = 40
    Plot = 41
End Enum

'-----------------------------
'  ActionTypes
'-----------------------------
Public Enum ActionType
    Sample = 42
    DataEntry = 43
    Verification = 44
    Download = 45
    Change = 46
End Enum

'-----------------------------
'  Status
'-----------------------------
Public Enum Status
    Opened = 47
    InProgress = 48
    Completed = 49
    Deferred = 50
End Enum

'-----------------------------
'  Priority
'-----------------------------
Public Enum Priority
    Critical = 51
    High = 52
    Medium = 53
    Low = 54
End Enum