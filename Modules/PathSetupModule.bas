Attribute VB_Name = "PathSetupModule"
Option Explicit

Public fso As Object
Public destinationFolderPath As String
Public databaseWorkbookPath As String

Sub SetDestinationFolderPath()
    Call InitFso

    destinationFolderPath = Trim(welcomeWorksheet.Range("J12").value)
    If destinationFolderPath = "" Then
        Call Err.Raise( _
            vbObjectError + 801, _
            Description:="Le chemin du dossier d'emplacement des reportings ne peut pas �tre vide." _
        )
    End If
    
    If IsRelativePath(destinationFolderPath) Then Call ConvertRelativePathToAbsolute(destinationFolderPath)
    If fso.FileExists(destinationFolderPath) Then
        Call Err.Raise( _
            vbObjectError + 802, _
            Description:="Le chemin sp�cifi� """ & destinationFolderPath & """ correspond � un fichier, pas � un dossier." _
        )
    ElseIf Not fso.FolderExists(destinationFolderPath) Then
        Call Err.Raise( _
            vbObjectError + 803, _
            Description:="Le dossier de destination des reportings donn� """ & destinationFolderPath & """ n'existe pas." _
        )
    End If
End Sub

Sub SetDatabaseWorkbookPath()
    Call InitFso

    databaseWorkbookPath = Trim(welcomeWorksheet.Range("J15").value)
    If databaseWorkbookPath = "" Then
        Call Err.Raise( _
            vbObjectError + 804, _
            Description:="Le chemin du fichier de base de donn�es ne peut pas �tre vide." _
        )
    End If
    
    If IsRelativePath(databaseWorkbookPath) Then Call ConvertRelativePathToAbsolute(databaseWorkbookPath)
    If Not fso.FileExists(databaseWorkbookPath) Then
        Call Err.Raise( _
            vbObjectError + 805, _
            Description:="Le fichier de base de donn�es """ & databaseWorkbookPath & """ n'existe pas." _
        )
    End If
End Sub

Private Sub InitFso()
    Static initialized As Boolean
    
    If initialized Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    initialized = True
End Sub

Private Sub ConvertRelativePathToAbsolute(relativePath As String)
    relativePath = fso.BuildPath(ThisWorkbook.path, relativePath)
End Sub

Private Function IsRelativePath(path As String) As Boolean
    If Left(path, 1) = "/" Then
        IsRelativePath = False
    ElseIf Len(path) > 2 And Mid(path, 2, 1) = ":" Then
        IsRelativePath = False
    Else
        IsRelativePath = True
    End If
End Function
