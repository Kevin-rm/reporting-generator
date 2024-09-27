Attribute VB_Name = "VolumeRecapWbGeneratorModule"
Option Explicit

Public volumeRecapWorksheet As Worksheet

Sub InitVolumeRecapWorksheet()
    Static initialized As Boolean
    
    If initialized Then Exit Sub
    Set volumeRecapWorksheet = ThisWorkbook.Sheets("RECAP volume")
    
    initialized = True
End Sub

Sub CreateVolumeRecapWorkbook()
    On Error GoTo ErrorHandler

    Dim newWorkbook As Workbook
    Dim distributorWorkbook As Workbook
    Dim newSheet As Worksheet
    Dim newSheetName As String
    Dim folderPath As String
    Dim filesFound As Boolean
    Dim file As String
    Dim fileNameWithoutExtension As String
    Dim lastDashPosition As Integer
    Dim savePath As String
    
    Call InitWelcomeWorksheet
    Call SetDestinationFolderPath
    Call InitVolumeRecapWorksheet
    Call InitMonthNames
    
    folderPath = destinationFolderPath
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    savePath = folderPath & "RECAP VOLUME.xlsx"
    If Dir(savePath) <> "" Then
        Call MsgBox("Le fichier de récapitulation de volume """ & savePath & """ existe déjà. Veuillez le renommer ou le supprimer avant de continuer.", vbExclamation)
        Exit Sub
    End If
    
    Set newWorkbook = Workbooks.Add
    
    file = Dir(folderPath & "*.xlsm")
    Do While file <> ""
        filesFound = True
    
        Set distributorWorkbook = Workbooks.Open(folderPath & file)
        
        Call volumeRecapWorksheet.Copy(After:=newWorkbook.Sheets(newWorkbook.Sheets.Count))
        
        fileNameWithoutExtension = Left(file, InStrRev(file, ".") - 1)
        lastDashPosition = InStrRev(fileNameWithoutExtension, "-")
        
        If lastDashPosition > 0 Then
            newSheetName = Trim(Mid(fileNameWithoutExtension, lastDashPosition + 1))
        Else
            newSheetName = fileNameWithoutExtension
        End If
        
        Set newSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count)
        newSheet.Name = newSheetName

        Call GenerateCurrentWorksheetContent(distributorWorkbook, newSheet)
        Call distributorWorkbook.Close(False)
    
        file = Dir
    Loop
    
    If Not filesFound Then
        Call MsgBox("Aucun fichier de reporting n'a été trouvé dans le dossier spécifié.", vbExclamation)
        Call newWorkbook.Close(False)
        
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    newWorkbook.Sheets(1).Delete
    Application.DisplayAlerts = True

    Call newWorkbook.SaveAs(fileName:=savePath, FileFormat:=xlOpenXMLWorkbook)
    Call newWorkbook.Close(SaveChanges:=False)
    
    Exit Sub
ErrorHandler:
    Select Case Err.Number - vbObjectError
        Case 801, 802, 803, 804, 805
            Call MsgBox("Erreur: " & Err.Description, vbCritical)
        
        Case Else
            Call HandleUnexpectedError
    End Select
    
    If Not newWorkbook Is Nothing Then Call newWorkbook.Close(False)
    Set newWorkbook = Nothing
    
    Err.Clear
    
    Exit Sub
End Sub

Private Sub GenerateCurrentWorksheetContent(distributorWorkbook As Workbook, currentWorksheet As Worksheet)
    Dim ws As Worksheet
    Dim index As Integer
    Dim targetRow As Long
    
    targetRow = 6
    
    Set ws = distributorWorkbook.Sheets(monthNames(1))
    Call FindLastDataRow(ws, True)
    
    Call ws.Range("C3:C" & lastDataRow).Copy(Destination:=currentWorksheet.Cells(targetRow, 1))
    Call currentWorksheet.Columns("A").AutoFit
    
    For index = LBound(monthNames) To UBound(monthNames)
        Set ws = distributorWorkbook.Sheets(monthNames(index))
        
        Call ws.Range("AB3:AB" & lastDataRow).Copy
        Call currentWorksheet.Cells(targetRow, index + 2).PasteSpecial(Paste:=xlPasteValues)
    Next index

    lastDataRow = currentWorksheet.Cells(currentWorksheet.Rows.Count, "A").End(xlUp).row
    currentWorksheet.Cells(targetRow, "N").Formula = "=IFERROR(SUM(B" & targetRow & ":M" & targetRow & "), """")"
    
    Call currentWorksheet.Range("N" & targetRow & ":N" & lastDataRow).FillDown
    Call FormatDataRange(currentWorksheet.Range("A" & targetRow & ":N" & lastDataRow))
End Sub

Private Sub FormatDataRange(dataRange As Range)
    With dataRange
        Call ApplyBorder(.Borders(xlEdgeBottom), xlMedium)
        Call ApplyBorder(.Borders(xlEdgeTop), xlMedium)
        Call ApplyBorder(.Borders(xlEdgeLeft), xlMedium)
        Call ApplyBorder(.Borders(xlEdgeRight), xlMedium)
        
        Call ApplyBorder(.Borders(xlInsideVertical), xlMedium)
        Call ApplyBorder(.Borders(xlInsideHorizontal), xlThin)
        
        .Font.Size = 9
        .numberFormat = "#,##0.00"
    End With
End Sub

Private Sub ApplyBorder(border As border, weight As XlBorderWeight)
    With border
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = weight
    End With
End Sub
