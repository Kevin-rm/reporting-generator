VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Distributor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private pWorkbook As Workbook
Private pResellers As Collection

Private Sub Class_Initialize()
    Set pResellers = New Collection
End Sub

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get Resellers() As Collection
    Set Resellers = pResellers
End Property

Public Sub AddReseller(Reseller As Reseller)
    Call pResellers.Add(Reseller)
End Sub

Public Sub LoadResellersFromSheet(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim Reseller As Reseller
    Dim resellerType As resellerType
    Dim Segment As Segment
    
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    For i = 2 To lastRow
        Set Reseller = New Reseller
        
        Set Segment = New Segment
        Segment.Name = ws.Cells(i, 4).value
        
        Set resellerType = New resellerType
        resellerType.Name = ws.Cells(i, 3).value
        Segment.resellerType = resellerType
        
        Reseller.Zone = ws.Cells(i, 1).value
        Reseller.Name = ws.Cells(i, 2).value
        Reseller.Segment = Segment
        Reseller.Category = ws.Cells(i, 5).value
        Reseller.City = ws.Cells(i, 6).value
        Reseller.Address = ws.Cells(i, 7).value
        Reseller.Contact = ws.Cells(i, 8).value
        
        Call Me.AddReseller(Reseller)
    Next i
End Sub

Public Sub CreateWorkbook( _
    ByVal destinationFolderPath As String, _
    ByVal mainModuleFileName As String, _
    ByVal resellerFormFileName As String, _
    ByVal resellerTypeFileName As String, _
    ByVal segmentFileName As String _
)
    On Error GoTo ErrorHandler

    Dim fileName As String
    Dim i As Integer
    Dim sheet As Worksheet
    Dim recapSheet As Worksheet
    Dim currentYear As Integer
    
    currentYear = Year(Date)
    fileName = destinationFolderPath & "\REPORTING " & currentYear & " - " & pName & ".xlsm"
    If fso.FileExists(fileName) Then
        Call Err.Raise( _
            vbObjectError + 800, _
            "CreateWorkbook", "Le fichier """ & fileName & """ existe déjà. La régénération du fichier a été annulée." _
        )
    End If
  
    Set pWorkbook = Workbooks.Add
     
    For Each sheet In ThisWorkbook.Sheets
        If Not sheet Is welcomeWorksheet And Not sheet Is volumeRecapWorksheet Then ' On ignore les sheets "Accueil" et "RECAP volume"
            Call sheet.Copy(After:=pWorkbook.Sheets(pWorkbook.Sheets.Count))
        End If
    Next sheet
    
    Call ImportComponent(resellerTypeFileName)
    Call ImportComponent(segmentFileName)
    Call ImportComponent(resellerFormFileName)
    Call ImportComponent(mainModuleFileName)
    
    ' Code sensible: change la macro affectée à la forme qui affiche le formulaire
    pWorkbook.Sheets("Formulaire") _
        .Shapes("Group 1") _
        .GroupItems("Rounded Rectangle 2") _
        .OnAction = "'" & pWorkbook.Name & "'!ShowResellerForm"
    
    Set recapSheet = pWorkbook.Sheets("RECAP")
    Call GenerateWorkbookContent(currentYear, recapSheet)
    Call CreatePivotTables(recapSheet)
    
    Application.DisplayAlerts = False
    For i = 1 To 2
        pWorkbook.Sheets(1).Delete
    Next i
    Application.DisplayAlerts = True
    
    Call pWorkbook.SaveAs(fileName:=fileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled)
    Call pWorkbook.Close(SaveChanges:=False)
    
    Exit Sub
ErrorHandler:
    If Not pWorkbook Is Nothing Then
        Call pWorkbook.Close(SaveChanges:=False)
        Set pWorkbook = Nothing
    End If
    
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub CreatePivotTables(recapSheet As Worksheet)
    Dim TCDWorksheet As Worksheet
    Dim pivotCache As pivotCache
    Dim gasCylinderFormatPivotTable As pivotTable
    
    Set TCDWorksheet = pWorkbook.Sheets("TCD")
    Set pivotCache = pWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=recapSheet.Range("A2:AB" & lastDataRow - 1) _
    )
    
    Call PivotTableCreationHelper("Ville", pivotCache, TCDWorksheet, "A1")
    Call PivotTableCreationHelper("Type", pivotCache, TCDWorksheet, "E1")
    Call PivotTableCreationHelper("Segment", pivotCache, TCDWorksheet, "I1")
    
    Set gasCylinderFormatPivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=TCDWorksheet.Range("M1"), _
        TableName:="TCDFormatBouteille" _
    )
    With gasCylinderFormatPivotTable
        Call AddDataField(gasCylinderFormatPivotTable, "Livrées", "B04 vendues")
        Call AddDataField(gasCylinderFormatPivotTable, "Livrées2", "B09 vendues")
        Call AddDataField(gasCylinderFormatPivotTable, "Livrées3", "B12 vendues")
        Call AddDataField(gasCylinderFormatPivotTable, "Livrées4", "RC12 vendues")
        Call AddDataField(gasCylinderFormatPivotTable, "Livrées5", "B39 vendues")
        Call AddDataField(gasCylinderFormatPivotTable, "Livrées6", "RC39 vendues")
        
        .DataPivotField.caption = "PAR FORMAT DE BOUTEILLE"
    End With
End Sub

Private Sub PivotTableCreationHelper( _
    ByVal pivotRowFieldName As String, _
    pivotCache As pivotCache, _
    TCDWorksheet As Worksheet, _
    startCellString As String _
)
    Dim pivotTable As pivotTable

    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=TCDWorksheet.Range(startCellString), _
        TableName:="TCD" & pivotRowFieldName _
    )
    With pivotTable
        .PivotFields(UCase(pivotRowFieldName)).Orientation = xlRowField
    
        Call AddDataField(pivotTable, "Volume", "Volume total")
        Call AddDataField(pivotTable, "Volume", "Pourcentage", "0.00%", xlPercentOfTotal)
        
        .DataPivotField.caption = "PAR " & UCase(pivotRowFieldName)
    End With
End Sub

Private Sub AddDataField( _
    pivotTable As pivotTable, _
    fieldName As String, _
    caption As String, _
    Optional numberFormat As String = "#,##0", _
    Optional calculation As XlPivotFieldCalculation = xlNoAdditionalCalculation _
)
    With pivotTable.PivotFields(fieldName)
        .Orientation = xlDataField
        .Function = xlSum
        .calculation = calculation
        .caption = caption
        .numberFormat = numberFormat
    End With
End Sub

Private Sub GenerateWorkbookContent(ByVal currentYear As Integer, recapSheet As Worksheet)
    Dim templateSheet As Worksheet
    Dim inventorySheet As Worksheet
    Dim rowIndex As Integer
    Dim i As Integer
    Dim Reseller As Reseller

    Set templateSheet = pWorkbook.Sheets("template")
    Set inventorySheet = pWorkbook.Sheets("Inventaire")
    
    With inventorySheet
        .Range("A6").value = DateSerial(currentYear, 1, 1)
        .Range("A6").AutoFill Destination:=.Range("A6:A17"), Type:=xlFillValues
        
        Dim cell As Range
        For Each cell In .Range("H6:H17").Cells
            cell.Errors(xlOmittedCells).Ignore = True
        Next cell
    End With
    
    rowIndex = startRowIndex
    For i = 1 To pResellers.Count
        Set Reseller = pResellers(i)
        
        If i < pResellers.Count Then
            templateSheet.Rows(rowIndex).Insert Shift:=xlDown
            recapSheet.Rows(rowIndex).Insert Shift:=xlDown
        End If
        Call InsertResellerRow(Reseller, templateSheet, rowIndex)
        Call InsertResellerRow(Reseller, recapSheet, rowIndex)
        
        rowIndex = rowIndex + 1
    Next i
    
    lastDataRow = rowIndex - 1
    
    Call FillConsignementAndVolumeCells(templateSheet, rowIndex, lastDataRow)
    Call CreateMonthSheets(templateSheet)
    Call DoRecap(recapSheet, rowIndex, lastDataRow)
End Sub

Private Sub InsertResellerRow(Reseller As Reseller, ws As Worksheet, ByVal currentRowIndexValue As Integer)
    Dim Segment As Segment
    
    Set Segment = Reseller.Segment
    With ws
        .Cells(currentRowIndexValue, 1).value = Reseller.Zone
        .Cells(currentRowIndexValue, 2).value = pName
        .Cells(currentRowIndexValue, 3).value = Reseller.Name
        .Cells(currentRowIndexValue, 4).value = Segment.resellerType.Name
        .Cells(currentRowIndexValue, 5).value = Segment.Name
        .Cells(currentRowIndexValue, 6).value = Reseller.Category
        .Cells(currentRowIndexValue, 7).value = Reseller.City
        .Cells(currentRowIndexValue, 8).value = Reseller.Address
        .Cells(currentRowIndexValue, 9).value = Reseller.Contact
    End With
End Sub

Private Sub FillConsignementAndVolumeCells( _
    templateSheet As Worksheet, _
    ByVal currentRowIndexValue As Integer, _
    ByVal lastRowIndex As Integer _
)
    Dim initialConsignementCell As Range
    Dim initialVolumeCell As Range
    Dim i As Integer
    Dim colOffset As Integer
    
    Set initialConsignementCell = templateSheet.Cells(startRowIndex, startColumnIndex + 2)
    initialConsignementCell.Formula = "=IFERROR(J" & startRowIndex & " - K" & startRowIndex & ", """")"
    
    With initialConsignementCell
        If lastRowIndex > startRowIndex Then
            .AutoFill Destination:=templateSheet.Range(.Address, templateSheet.Cells(lastRowIndex, .Column))
        End If
    
        For i = 1 To 5
            colOffset = (startColumnIndex + 2) + (i * 3)
            
            .Copy
            templateSheet.Cells(startRowIndex, colOffset).PasteSpecial Paste:=xlPasteFormulas
            If lastRowIndex > startRowIndex Then
                templateSheet.Range(.Address, templateSheet.Cells(lastRowIndex, colOffset)).FillDown
            End If
        Next i
    End With
    
    ' Remplissement des cellules de la colonne \"Volume\"
    Set initialVolumeCell = templateSheet.Cells(startRowIndex, colOffset + 1)
    initialVolumeCell.Formula = "=J" & startRowIndex & " * 4 + " & _
                   "M" & startRowIndex & " * 9 + " & _
                   "P" & startRowIndex & " * 12.5 + " & _
                   "S" & startRowIndex & " * 39 + " & _
                   "V" & startRowIndex & " * 12.5 + " & _
                   "Y" & startRowIndex & " * 39"
            
    If lastRowIndex > startRowIndex Then
        templateSheet.Range(initialVolumeCell.Address, templateSheet.Cells(lastRowIndex, colOffset + 1)).FillDown
    End If
    
    Call CalculateSum(templateSheet, currentRowIndexValue)
    Call ReformatSheet(templateSheet)
End Sub

Private Sub CreateMonthSheets(templateSheet As Worksheet)
    Dim monthName As Variant
    Dim newSheet As Worksheet
    
    For Each monthName In monthNames
        templateSheet.Copy After:=pWorkbook.Sheets(pWorkbook.Sheets.Count)
        
        Set newSheet = pWorkbook.Sheets(pWorkbook.Sheets.Count)
        newSheet.Name = monthName
    Next monthName
End Sub

Private Sub DoRecap( _
    recapSheet As Worksheet, _
    ByVal currentRowIndexValue As Integer, _
    ByVal lastRowIndex As Integer _
)
    Dim col As Integer
    Dim i As Integer
    Dim formulaString As String
    Dim monthSheet As Worksheet
    
    For col = startColumnIndex To 40
        If col >= 38 Then ' Les présentoirs ne se cumulent pas mais prennent la valeur du mois actuel
            formulaString = "=INDIRECT(TEXT(TODAY(), ""mmmm"") & ""!L"" & ROW() & ""C"" & " & col & ", FALSE)"
        Else
            ' \"startRowIndex - 1\" car on a déjà reformaté les months sheets
            formulaString = "=IFERROR(SUM('" & monthNames(0) & ":" & monthNames(11) & "'!" & recapSheet.Cells(startRowIndex - 1, col).Address(False, False) & "), """")"
        End If
        
        recapSheet.Cells(startRowIndex, col).Formula = formulaString
        
        If lastRowIndex > startRowIndex Then
            recapSheet.Cells(startRowIndex, col).AutoFill Destination:=recapSheet.Range(recapSheet.Cells(startRowIndex, col), recapSheet.Cells(lastRowIndex, col)), Type:=xlFillValues
        End If
    Next col
    
    Call CalculateSum(recapSheet, currentRowIndexValue)
    Call ReformatSheet(recapSheet)
End Sub

Private Sub CalculateSum(ws As Worksheet, ByVal currentRowIndexValue As Integer)
    With ws
        .Cells(currentRowIndexValue, startColumnIndex).Formula = "=SUM(J" & startRowIndex & ":J" & currentRowIndexValue - 1 & ")"
        .Range(.Cells(currentRowIndexValue, startColumnIndex), .Cells(currentRowIndexValue, 40)).FillRight
    End With
End Sub

Private Sub ReformatSheet(ws As Worksheet)
    With ws
        .Rows(startRowIndex - 1).Delete
        .Columns("A:I").AutoFit
    End With
End Sub

Private Sub ImportComponent(componentFileName As String)
    Call pWorkbook _
            .VBProject _
            .VBComponents _
            .Import(componentFileName)
End Sub
