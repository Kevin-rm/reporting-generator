VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Reseller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private pZone As String
Private pSegment As Segment
Private pCategory As String
Private pCity As String
Private pAddress As String
Private pContact As String

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get Zone() As String
    Zone = pZone
End Property

Public Property Let Zone(value As String)
    pZone = value
End Property

Public Property Get Segment() As Segment
    Set Segment = pSegment
End Property

Public Property Let Segment(value As Segment)
    Set pSegment = value
End Property

Public Property Get Category() As String
    Category = pCategory
End Property

Public Property Let Category(value As String)
    pCategory = value
End Property

Public Property Get City() As String
    City = pCity
End Property

Public Property Let City(value As String)
    pCity = value
End Property

Public Property Get Address() As String
    Address = pAddress
End Property

Public Property Let Address(value As String)
    pAddress = value
End Property

Public Property Get Contact() As String
    Contact = pContact
End Property

Public Property Let Contact(value As String)
    pContact = value
End Property
