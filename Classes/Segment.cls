VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Segment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private pResellerType As resellerType

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get resellerType() As resellerType
    Set resellerType = pResellerType
End Property

Public Property Let resellerType(value As resellerType)
    Set pResellerType = value
    Call pResellerType.AddSegment(Me)
End Property
