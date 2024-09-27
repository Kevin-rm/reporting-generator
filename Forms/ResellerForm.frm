VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResellerForm 
   Caption         =   "Formulaire revendeur"
   ClientHeight    =   7716
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4932
   OleObjectBlob   =   "ResellerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ResellerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TypeComboBox_Change()
    With Me
        .SegmentComboBox.Clear
        
        Dim segmentCollection As Collection
        Dim Segment As Segment
        
        Set segmentCollection = GetSegmentsByResellerTypeName(.TypeComboBox.value)
        For Each Segment In segmentCollection
            .SegmentComboBox.AddItem Segment.Name
        Next Segment
    End With
End Sub

Private Sub SaveButton_Click()
    On Error GoTo ErrorHandler
    
    Call ValidateData
    Call AddDataToSheets(Me)
    
    Call MsgBox("Données ajoutées avec succès !", vbInformation)
    Call Unload(Me)

    Exit Sub
ErrorHandler:
    Call MsgBox("Erreur : " & Err.Description, vbCritical)
    Err.Clear
End Sub

Private Sub CancelButton_Click()
    Call Unload(Me)
End Sub

Private Sub AssertNotBlank(value As String, errorMessageIfNotValid As String)
    value = Trim(value)
    
    If value = "" Then
        Call Err.Raise(vbObjectError + 1000, Description:=errorMessageIfNotValid)
    End If
End Sub

Private Sub ValidateData()
    Call AssertNotBlank(NameTextBox.value, "Le champ Nom est obligatoire.")
    Call AssertNotBlank(TypeComboBox.value, "Le champ Type est obligatoire.")
    Call AssertNotBlank(SegmentComboBox.value, "Le champ Segment est obligatoire.")
    Call AssertNotBlank(CityTextBox.value, "Le champ Ville est obligatoire.")
    Call AssertNotBlank(AdressTextBox.value, "Le champ Adresse est obligatoire.")
    Call AssertNotBlank(ContactTextBox.value, "Le champ Contact est obligatoire.")
End Sub

Private Sub AddDataToSheets(form As ResellerForm)
    Dim wsName As Variant
    Dim wsNames() As String
    Dim ws As Worksheet
    Dim newRow As Long
    Dim monthAutoFillCols As Variant
    Dim col As Variant
    Dim colsWithMediumWeightLeftBorders As Variant
    Dim previousCell As Range
    Dim currentCell As Range
    
    Call InitMonthNames
    Call FindLastDataRow(ThisWorkbook.Sheets("RECAP"))
    
    wsNames = Split("RECAP," & Join(monthNames, ","), ",")
    
    monthAutoFillCols = Array(2, 12, 15, 18, 21, 24, 27, 28) ' B, L, O, R, U, X, AA, AB
    colsWithMediumWeightLeftBorders = Array(10, 13, 16, 19, 22, 25, 28, 29, 38, 41)
    newRow = lastDataRow + 1

    For Each wsName In wsNames
        Set ws = ThisWorkbook.Sheets(wsName)
        ws.Rows(newRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
        If wsName = "RECAP" Then
            Call AutoFillCell(newRow, 2, previousCell, currentCell, ws) ' B
            
            ' J jusqu'à AN
            For col = 10 To 40
                Call AutoFillCell(newRow, col, previousCell, currentCell, ws)
            Next col
        Else
            For Each col In monthAutoFillCols
                Call AutoFillCell(newRow, col, previousCell, currentCell, ws)
            Next col
        End If
        
        Set previousCell = Nothing
        Set currentCell = Nothing
        
        ws.Cells(newRow, 1).value = form.ZoneComboBox.value
        ws.Cells(newRow, 3).value = form.NameTextBox.value
        ws.Cells(newRow, 4).value = form.TypeComboBox.value
        ws.Cells(newRow, 5).value = form.SegmentComboBox.value
        ws.Cells(newRow, 7).value = form.CityTextBox.value
        ws.Cells(newRow, 8).value = form.AdressTextBox.value
        ws.Cells(newRow, 9).value = form.ContactTextBox.value
        
        With ws.Range(ws.Cells(newRow, 1), ws.Cells(newRow, 40))
            .Borders.LineStyle = xlContinuous
            .Borders.weight = xlThin
        End With
        
        For Each col In colsWithMediumWeightLeftBorders
            With ws.Cells(newRow, col).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .weight = xlMedium
            End With
        Next col
        
        Set ws = Nothing
    Next wsName
    
    Call UpdatePivotTablesDataRange
End Sub

Private Sub AutoFillCell(ByVal newRow As Long, ByVal columnIndex As Variant, previousCell As Range, currentCell As Range, ws As Worksheet)
    Set previousCell = ws.Cells(newRow - 1, columnIndex)
    Set currentCell = ws.Cells(newRow, columnIndex)

    Call previousCell.AutoFill(Destination:=ws.Range(previousCell, currentCell), Type:=xlFillDefault)
End Sub

Private Sub UpdatePivotTablesDataRange()
    Dim recapSheet As Worksheet
    Dim TCDSheet As Worksheet
    Dim pivotTable As pivotTable
    
    Set TCDSheet = ThisWorkbook.Sheets("TCD")
    Set recapSheet = ThisWorkbook.Sheets("RECAP")
    Call FindLastDataRow(recapSheet, True)
    
    Call InitStartRowIndexAndStartColumnIndex
    For Each pivotTable In TCDSheet.PivotTables
        Call pivotTable.ChangePivotCache(ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=recapSheet.Range("A" & startRowIndex - 2 & ":AB" & lastDataRow) _
        ))
        
        Call pivotTable.RefreshTable
    Next pivotTable
End Sub

Private Function GetSegmentsByResellerTypeName(resellerTypeName) As Collection
    Dim resellerType As resellerType
    
    Set GetSegmentsByResellerTypeName = Nothing
    For Each resellerType In resellerTypeCollection
        If resellerType.Name = resellerTypeName Then
            Set GetSegmentsByResellerTypeName = resellerType.segments
            
            Exit For
        End If
    Next resellerType
End Function
