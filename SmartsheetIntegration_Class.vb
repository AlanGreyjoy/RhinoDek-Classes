Imports Smartsheet.Api
Imports Smartsheet.Api.Models
Imports Smartsheet.Api.OAuth

Public Class SmartSheetIntegration

    Dim token As Token
    Dim settings As New RhinoDekSettings
    Dim smartSheet As SmartsheetClient

    Public Sub New()
        token = New Token
        token.AccessToken = settings.SmartSheetToken
        smartSheet = New SmartsheetBuilder().SetAccessToken(token.AccessToken).Build()
    End Sub

    'Test Create Folder
    Public Sub CreateFolderHome(folderName As String)
        Dim folder As Folder
        folder = New Folder.CreateFolderBuilder(folderName).Build()
        folder = smartSheet.HomeResources.FolderResources.CreateFolder(folder)
    End Sub

    'Test Create Workspace
    Public Sub CreateWorkSpace(workspaceName As String)
        Dim workspace As Workspace
        workspace = New Workspace.CreateWorkspaceBuilder(workspaceName).Build()
        workspace = smartSheet.WorkspaceResources.CreateWorkspace(workspace)
    End Sub


    'Get all sheets in given workspace
    Public Sub GetSheetsInWorkspace(workspaceID As Int64)

        Dim workspace As Workspace
        workspace = smartSheet.WorkspaceResources.GetWorkspace(workspaceID, Nothing, Nothing)

        Dim folder As Folder
        folder = smartSheet.FolderResources.GetFolder(5398922303694724, Nothing)

        Dim sheets As List(Of Sheet)
        sheets = folder.Sheets

        For Each sheet As Sheet In sheets
            If sheet.Name = "ALAN S. INPUT" Then
                Dim id
                id = sheet.Id
                Dim getSheet = smartSheet.SheetResources.GetSheet(id, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                Dim row As Long
                row = getSheet.Rows(0).Id.Value
                smartSheet.SheetResources.RowResources.AttachmentResources.AttachFile(id, row, "N:\CUSTOM\M-CUS-28256\M-CUS-28256-TV.pdf", Nothing)
            End If
        Next

    End Sub

    'Get the users rfd que
    Public Function GetUserRFD(workspaceID As Int64, sheetName As String)

        Dim workspace As Workspace
        workspace = smartSheet.WorkspaceResources.GetWorkspace(workspaceID, Nothing, Nothing)

        Dim folder As Folder
        folder = smartSheet.FolderResources.GetFolder(5398922303694724, Nothing)

        Dim sheets As List(Of Sheet)
        sheets = folder.Sheets

        For Each sheet As Sheet In sheets
            If sheet.Name = sheetName Then
                Dim id = sheet.Id
                Dim getSheet = smartSheet.SheetResources.GetSheet(id, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                Dim rows As List(Of Row)
                rows = getSheet.Rows
                Return rows
            End If
        Next

    End Function

    'Get the row in the sheet
    Public Function GetRow(workspaceID As Int64, sheetName As String, rowID As Long)
        Dim workspace As Workspace
        workspace = smartSheet.WorkspaceResources.GetWorkspace(workspaceID, Nothing, Nothing)

        Dim folder As Folder
        folder = smartSheet.FolderResources.GetFolder(5398922303694724, Nothing)

        Dim sheets As List(Of Sheet)
        sheets = folder.Sheets

        For Each sheet As Sheet In sheets
            If sheet.Name = sheetName Then
                Dim id = sheet.Id
                Dim getSheet = smartSheet.SheetResources.GetSheet(id, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                Dim foundRow = smartSheet.SheetResources.RowResources.GetRow(getSheet.Id, rowID, Nothing, Nothing)
                Return foundRow
            End If
        Next
    End Function

End Class
