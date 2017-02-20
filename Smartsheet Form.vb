Imports System.IO
Imports System.Windows.Forms
Imports Smartsheet.Api
Imports Smartsheet.Api.Models
Imports Smartsheet.Api.OAuth
Imports Telerik.WinControls.UI

Public Class SmartsheetForm

    Dim rowID As Long
    Dim selectedRow As Row
    Dim sapPartnumber As String = ""
    Dim rowIDs As New List(Of Long)
    Dim g_row As Row
    Dim g_rowID As Long
    Dim g_columnID As Long
    Dim g_Sheet As Sheet
    Dim g_sheetID As Long
    Dim g_status As String
    Dim g_workspaceID As Long
    Dim g_rfdSheetID As Long
    Dim g_workspace As Workspace
    Dim g_smartSheet As SmartsheetClient
    Dim g_token As New Token
    Dim g_attachmentLink As String
    Dim settings As New RhinoDekSettings


    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Close()
    End Sub

    'Debug button
    Private Sub btnSmartSheet_Click(sender As Object, e As EventArgs)
        Dim clsSmartSheet As SmartSheetIntegration
        clsSmartSheet = New SmartSheetIntegration()
        'clsSmartSheet.CreateFolderHome("Test Folder")
        'clsSmartSheet.CreateWorkSpace("Workspace created from inside Rhino Test")
        clsSmartSheet.GetSheetsInWorkspace(3083654818752388)
    End Sub

    'Page Loag
    Private Sub SmartsheetForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim userSource = GetSources("\\surfer\public\cad\rhinodek\database\smartsheetusers.txt")
        If IsNothing(userSource) Then
            MsgBox("Your Smartsheet User text file is missing....")
            Return
        End If

        cbUsers.AutoCompleteCustomSource = userSource
        For Each user In userSource
            cbUsers.Items.Add(user)
        Next

        g_token.AccessToken = settings.SmartSheetToken
        g_smartSheet = New SmartsheetBuilder().SetAccessToken(g_token.AccessToken).Build()
        g_workspace = g_smartSheet.WorkspaceResources.GetWorkspace(settings.rfd_workspaceID, Nothing, Nothing)
        g_sheetID = settings.rfd_sheetid

        panelJobInfo.Visible = False

    End Sub

    'Get User Text file
    Public Function GetSources(file As String)
        Dim collection As New AutoCompleteStringCollection
        Try
            Using reader As New StreamReader(file)
                While Not reader.EndOfStream
                    collection.Add(reader.ReadLine())
                End While
            End Using
            Return collection
        Catch ex As Exception
            MsgBox("The SmarsheetUsers.txt file is missing from public/cad/rhinodek/database/smartsheetusers.txt!")
        End Try
    End Function

    Private Sub cbUsers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbUsers.SelectedIndexChanged
        If listJobQue.Items.Count > 0 Then
            listJobQue.Items.Clear()
        End If
        If rowIDs.Count > 0 Then
            rowIDs.Clear()
        End If

        Dim ss As New SmartSheetIntegration()
        ss.smartSheet = g_smartSheet
        Dim result As SearchResult
        result = ss.SearchQue(settings.rfd_sheetid, cbUsers.Text)

        For Each searchResult As SearchResultItem In result.Results
            If searchResult.ObjectType = SearchObjectType.ROW Then
                Dim rowID As Long
                rowID = searchResult.ObjectId
                rowIDs.Add(rowID)
                Dim row As Row = ss.GetRow(rowID)
                If row.Cells(2).Value = cbUsers.Text Then
                    Dim item As ListViewDataItem = New ListViewDataItem
                    Dim jobImage As New LightVisualElement()
                    jobImage.Image = My.Resources.appbar1
                    listJobQue.Items.Add(item)
                    item(1) = row.Cells(6).Value
                    If row.Cells(1).Value = "Red" Then
                        item.ForeColor = Drawing.Color.Red
                    End If
                    If row.Cells(1).Value = "Green" Then
                        item.ForeColor = Drawing.Color.Green
                    End If
                End If
            End If
        Next

        lblJobQueNumber.Text = listJobQue.Items.Count

    End Sub


    ''' <summary>
    ''' Fill out the textboxes with if from the retrived row in the smartsheet.
    ''' 
    ''' GetRow(workspaceID, sheetname, rowID)
    '''     workspaceID: The id of the Workspace that contain all the RFD's
    '''     sheetname: The name of the designers sheet ie:"ALAN S. INPUT"
    '''     rowID: The id of the row you want to to get
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub listJobQue_ItemMouseClick(sender As Object, e As Telerik.WinControls.UI.ListViewItemEventArgs) Handles listJobQue.ItemMouseClick
        panelJobInfo.Visible = True

        sapPartnumber = e.Item.Text
        rowID = e.ListViewElement.SelectedIndex
        Dim ss = New SmartSheetIntegration()
        ss.smartSheet = g_smartSheet
        Dim gotRow As Row = ss.GetRow(rowIDs(e.ListViewElement.SelectedIndex))
        Dim status = gotRow.Cells(1).Value

        If status = "Red" Then
            comboStatus.SelectedIndex = 3
        End If
        If status = "Green" Then
            comboStatus.SelectedIndex = 0
        End If
        If status = "Yellow" Then
            comboStatus.SelectedIndex = 1
        End If
        If status = "Blue" Then
            comboStatus.SelectedIndex = 2
        End If

        lblDueDate.Text = gotRow.Cells(4).Value
        lblSubmittedBy.Text = gotRow.Cells(3).Value
        lblCustomerType.Text = gotRow.Cells(5).Value
        lblCustomerName.Text = gotRow.Cells(7).Value
        lblMakeModelYear.Text = gotRow.Cells(8).Value
        lblNotes.Text = gotRow.Cells(14).Value
        lblSAPPN.Text = gotRow.Cells(6).Value

        Dim attachment As PaginatedResult(Of Attachment) = g_smartSheet.SheetResources.RowResources.AttachmentResources.ListAttachments(g_sheetID, rowIDs(e.ListViewElement.SelectedIndex), Nothing)
        Dim rfdAttachment As New Attachment()

        rfdAttachment = g_smartSheet.SheetResources.AttachmentResources.GetAttachment(g_sheetID, attachment.Data(0).Id)

        If rfdAttachment IsNot Nothing Then
            Dim button As New RadButton()
            button.Text = "PDF"
            button.Size = New Drawing.Size(30, 30)
            button.Dock = DockStyle.Left
            g_attachmentLink = rfdAttachment.Url
            AddHandler button.Click, AddressOf AttachmentClick
            panelAttachments.Controls.Add(button)
        End If

        g_row = gotRow
        g_rowID = gotRow.Id
        g_columnID = gotRow.Cells(1).ColumnId
        g_sheetID = gotRow.SheetId

        lblColumnID.Text = gotRow.Cells(2).ColumnId
        lblRowID.Text = gotRow.Id
        lblSheetId.Text = gotRow.SheetId

    End Sub

    Private Sub AttachmentClick(sender As Object, e As EventArgs)
        Process.Start(g_attachmentLink)
    End Sub


    ''' <summary>
    ''' Update the smartsheet for the given user
    ''' and the given workspaceid
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Dim ss As New SmartSheetIntegration()

        Dim cell As New List(Of Cell)
        cell.Add(New Cell.UpdateCellBuilder(g_columnID, g_status).Build())

        Dim row As Row = New Row.UpdateRowBuilder(g_rowID).SetCells(cell).Build()

        ss.SetRow(row, g_sheetID)
    End Sub

    Private Sub comboStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comboStatus.SelectedIndexChanged
        If comboStatus.SelectedIndex = 3 Then
            g_status = "Red"
        End If
        If comboStatus.SelectedIndex = 0 Then
            g_status = "Green"
        End If
        If comboStatus.SelectedIndex = 1 Then
            g_status = "Yellow"
        End If
        If comboStatus.SelectedIndex = 2 Then
            g_status = "Blue"
        End If
    End Sub
End Class
