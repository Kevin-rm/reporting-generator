Attribute VB_Name = "DistributorWbGeneratorModule"
Option Explicit

Public welcomeWorksheet As Worksheet

Private mainModuleFileName As String
Private resellerFormFileName As String
Private resellerTypeFileName As String
Private segmentFileName As String

Sub InitWelcomeWorksheet()
    Static initialized As Boolean
    
    If initialized Then Exit Sub
    Set welcomeWorksheet = ThisWorkbook.Sheets("Accueil")
    
    initialized = True
End Sub

Sub CreateWorkbooksForDistributors()
    On Error GoTo ErrorHandler
    
    Call InitStartRowIndexAndStartColumnIndex
    Call InitMonthNames
    Call InitWelcomeWorksheet
    Call InitVolumeRecapWorksheet
    
    Call SetDestinationFolderPath
    Call SetDatabaseWorkbookPath
    
    mainModuleFileName = ExportComponent("MainModule", "module")
    resellerFormFileName = ExportComponent("ResellerForm", "userform")
    resellerTypeFileName = ExportComponent("ResellerType", "class")
    segmentFileName = ExportComponent("Segment", "class")
    
    Dim databaseWorkbook As Workbook
    Dim ws As Worksheet
    Dim distributor As distributor
    
    Set databaseWorkbook = Workbooks.Open(databaseWorkbookPath)
    For Each ws In databaseWorkbook.Worksheets
        Set distributor = New distributor
        distributor.Name = ws.Name
        
        Call distributor.LoadResellersFromSheet(ws)
        Call distributor.CreateWorkbook( _
            destinationFolderPath, _
            mainModuleFileName, _
            resellerFormFileName, _
            resellerTypeFileName, _
            segmentFileName _
        )
    Next ws
    
    Call databaseWorkbook.Close(False)
    Call KillComponentFiles
    
    Exit Sub
ErrorHandler:
    Select Case Err.Number - vbObjectError
        Case 801, 802, 803, 804, 805
            Call MsgBox("Erreur: " & Err.Description, vbCritical)
            
        Case 800
            Call MsgBox(Err.Description, vbExclamation)
            Resume Next
        
        Case Else
            Call HandleUnexpectedError
    End Select
    
    If Not databaseWorkbook Is Nothing Then Call databaseWorkbook.Close(False)
    Set databaseWorkbook = Nothing
    
    Call KillComponentFiles
    
    Err.Clear
    
    Exit Sub
End Sub

Private Sub KillComponentFiles()
    If mainModuleFileName = "" Or _
        resellerFormFileName = "" Or _
        resellerTypeFileName = "" Or _
        segmentFileName = "" Then Exit Sub
        
    Call Kill(mainModuleFileName)
    Call Kill(resellerFormFileName)
    Call Kill(resellerTypeFileName)
    Call Kill(segmentFileName)
End Sub

Private Function ExportComponent(componentName As String, componentType As String) As String
    Dim temporaryFileName As String
    Dim temporaryFileExtension As String
    
    temporaryFileName = Environ("Temp") & "\" & componentName
    Select Case LCase(Trim(componentType))
        Case "module"
            temporaryFileExtension = "bas"
        Case "userform"
            temporaryFileExtension = "frm"
        Case "class"
            temporaryFileExtension = "cls"
        Case Else
            Err.Raise vbObjectError + 700, , "La valeur de l'argument componentType """ & componentType & """ est non reconnue"
    End Select
    temporaryFileName = temporaryFileName & "." & temporaryFileExtension

    Call ThisWorkbook _
        .VBProject _
        .VBComponents(componentName) _
        .Export(temporaryFileName)

    ExportComponent = temporaryFileName
End Function
