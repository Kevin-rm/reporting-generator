Attribute VB_Name = "MainModule"
Option Explicit

' Normes sur les numéros des erreurs :
' - À partir de 800 pour celles destinées aux utilisateurs
' - À partir de 700 pour les erreurs internes et autres

Public generatedPassword As String

' Coordonnées de la cellule \"B04 livrées\"
Public startRowIndex As Long
Public startColumnIndex As Long

Public monthNames() As String
Public lastDataRow As Long
Public resellerTypeCollection As Collection

Private formConfigSheet As Worksheet

Sub HandleUnexpectedError()
    Call MsgBox( _
        "Une erreur inattendue s'est produite." & vbCrLf & _
        "Description : " & Err.Description & vbCrLf & _
        "Source : " & Err.Source, _
        vbCritical _
    )
End Sub

Sub InitGeneratedPassword()
    Static initialized As Boolean
    
    If initialized Then Exit Sub
    generatedPassword = "uWpadfrH9NqmC5Pvyn3MwFjGRZJ6s8DhgtQeX7Ac4bKk2EYLBx"
    
    initialized = True
End Sub

Sub InitStartRowIndexAndStartColumnIndex()
    Static initialized As Boolean
    
    If initialized Then Exit Sub
    
    startRowIndex = 4
    startColumnIndex = 10
    
    initialized = True
End Sub

Sub InitMonthNames()
    Static initialized As Boolean
    
    If initialized Then Exit Sub
    Dim i As Integer
    ReDim monthNames(0 To 11)
    
    For i = 1 To 12
        monthNames(i - 1) = monthName(i)
    Next i
    
    initialized = True
End Sub

Sub FindLastDataRow(recapSheet As Worksheet, Optional update = False)
    Static initialized As Boolean
    
    If initialized And update = False Then Exit Sub
    lastDataRow = recapSheet.Cells(recapSheet.Rows.Count, "C").End(xlUp).row
    
    initialized = True
End Sub

Sub ShowResellerForm()
    Call loadFormConfigs
    Call ResellerForm.Show
End Sub

Private Sub InitFormConfigSheet()
    Static initialized As Boolean
    
    If initialized Then Exit Sub
    Set formConfigSheet = ThisWorkbook.Sheets("config-formulaire")
    
    initialized = True
End Sub

Private Sub loadFormConfigs()
    On Error GoTo ErrorHandler
    
    Dim start As Integer
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim col As Long
    Dim row As Long
    Dim resellerType As resellerType
    Dim segmentName As String
    Dim Segment As Segment

    Call InitFormConfigSheet
    Set resellerTypeCollection = New Collection
    
    start = 4
    
    lastColumn = formConfigSheet.Cells(start, formConfigSheet.Columns.Count).End(xlToLeft).Column
    lastRow = formConfigSheet.Cells(formConfigSheet.Rows.Count, 1).End(xlUp).row
    If lastRow < start Then lastRow = start
    
    For col = 1 To lastColumn
        Set resellerType = New resellerType
        resellerType.Name = formConfigSheet.Cells(start, col).value
        
        For row = start + 1 To lastRow
            segmentName = formConfigSheet.Cells(row, col).value
            If segmentName <> "" Then
                Set Segment = New Segment
                Segment.Name = segmentName
                
                Call resellerType.AddSegment(Segment)
            End If
        Next row

        Call resellerTypeCollection.Add(resellerType)
    Next col

    With ResellerForm.TypeComboBox
        .Clear

        For Each resellerType In resellerTypeCollection
            .AddItem resellerType.Name
        Next
    End With
    Call FillComboBox(ResellerForm.ZoneComboBox, "B1")
    
    Exit Sub
ErrorHandler:
    Call HandleUnexpectedError
    
    Exit Sub
End Sub

Private Sub FillComboBox(comboBox As comboBox, cellAddress As String)
    Dim i As Integer
    Dim items() As String
    
    items = Split(formConfigSheet.Range(cellAddress), ",")
    With comboBox
        .Clear
        
        For i = LBound(items) To UBound(items)
            .AddItem Trim(items(i))
        Next i
        
        .ListIndex = 0
    End With
End Sub
