VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResellerForm 
   Caption         =   "Formulaire revendeur"
   ClientHeight    =   8568.001
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
    Call AssertNotBlank(AddressTextBox.value, "Le champ Adresse est obligatoire.")
    Call AssertNotBlank(ContactTextBox.value, "Le champ Contact est obligatoire.")
End Sub

Private Sub AddDataToSheets(form As ResellerForm)
    Dim wsName As Variant
    Dim wsNames() As String
    Dim ws As Worksheet
    Dim recapSheet As Worksheet
    Dim newRow As Long
    Dim monthAutoFillCols As Variant
    Dim col As Variant
    Dim colsWithMediumWeightLeftBorders As Variant
    Dim previousCell As Range
    Dim currentCell As Range
    
    Call InitMonthNames
    Call InitStartRowIndexAndStartColumnIndex
    Set recapSheet = ThisWorkbook.Sheets("RECAP")
    Call FindLastDataRow(recapSheet)
    
    wsNames = Split("RECAP," & Join(monthNames, ","), ",")
    
    monthAutoFillCols = Array(1, 2, 15, 18, 21, 24, 27, 30, 31) ' A, B, O, R, U, X, AA, AD, AE
    colsWithMediumWeightLeftBorders = Array(13, 16, 19, 22, 25, 28, 31, 32, 41, 44)
    newRow = lastDataRow + 1
    
    Call InitGeneratedPassword
    For Each wsName In wsNames
        Set ws = ThisWorkbook.Sheets(wsName)
        
        Call ws.Unprotect(generatedPassword)
        Call ws.Rows(newRow).Insert(Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove)
        
        If wsName = "RECAP" Then
            Call AutoFillCell(newRow, 1, previousCell, currentCell, ws) ' A (Zone)
            Call AutoFillCell(newRow, 2, previousCell, currentCell, ws) ' B (distributeur)
            
            ' M jusqu'à AQ
            For col = startColumnIndex To 43
                Call AutoFillCell(newRow, col, previousCell, currentCell, ws)
            Next col
        Else
            For Each col In monthAutoFillCols
                Call AutoFillCell(newRow, col, previousCell, currentCell, ws)
            Next col
        End If
        
        Set previousCell = Nothing
        Set currentCell = Nothing
        
        ws.Cells(newRow, 3).value = form.CodeTextBox.value
        ws.Cells(newRow, 4).value = form.NameTextBox.value
        ws.Cells(newRow, 5).value = form.TypeComboBox.value
        ws.Cells(newRow, 6).value = form.SegmentComboBox.value
        ws.Cells(newRow, 7).value = form.CityTextBox.value
        ws.Cells(newRow, 8).value = form.AddressTextBox.value
        ws.Cells(newRow, 9).value = form.ContactTextBox.value
        ws.Cells(newRow, 10).value = Date
        ws.Cells(newRow, 11).value = "Actif"
        ws.Cells(newRow, 12).value = form.VolumeTextBox.value
        
        With ws.Range(ws.Cells(newRow, 1), ws.Cells(newRow, 44))
            .Borders.LineStyle = xlContinuous
            .Borders.weight = xlThin
        End With
        
        For Each col In colsWithMediumWeightLeftBorders
            With ws.Cells(newRow, col).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .weight = xlMedium
            End With
        Next col
        
        ws.Columns("A:L").AutoFit
        Call CalculateSum(ws, startRowIndex - 1, newRow + 1)
        
        If wsName <> "RECAP" Then Call ws.Protect(password:=generatedPassword)
        
        Set ws = Nothing
    Next wsName
    
    Call UpdatePivotTablesDataRange(recapSheet)
End Sub

Private Sub AutoFillCell(ByVal newRow As Long, ByVal columnIndex As Variant, previousCell As Range, currentCell As Range, ws As Worksheet)
    Set previousCell = ws.Cells(newRow - 1, columnIndex)
    Set currentCell = ws.Cells(newRow, columnIndex)

    Call previousCell.AutoFill(Destination:=ws.Range(previousCell, currentCell), Type:=xlFillDefault)
End Sub

Private Sub UpdatePivotTablesDataRange(recapSheet As Worksheet)
    Dim TCDSheet As Worksheet
    Dim pivotTable As pivotTable
    
    Set TCDSheet = ThisWorkbook.Sheets("TCD")
    
    Call FindLastDataRow(recapSheet, True)
    Call TCDSheet.Unprotect(generatedPassword)
    
    For Each pivotTable In TCDSheet.PivotTables
        Call pivotTable.ChangePivotCache(ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=recapSheet.Range("A" & startRowIndex - 2 & ":AE" & lastDataRow) _
        ))
        
        Call pivotTable.RefreshTable
    Next pivotTable
    
    Call recapSheet.Protect(generatedPassword)
    Call TCDSheet.Protect(generatedPassword)
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
