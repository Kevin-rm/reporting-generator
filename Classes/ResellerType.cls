VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResellerType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private pSegments As Collection

Private Sub Class_Initialize()
    Set pSegments = New Collection
End Sub

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get segments() As Collection
    Set segments = pSegments
End Property

Public Sub AddSegment(Segment As Segment)
    Call pSegments.Add(Segment)
End Sub
