Imports System.Data.SQLite
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
    Dim g_columnID_Status As Long
    Dim g_columnID_StartDate As Long
    Dim g_columnID_EndDate As Long
    Dim g_columnID_Time As Long
    Dim g_columnID_NumberOfParts As Long
    Dim g_Sheet As Sheet
    Dim g_sheetID As Long
    Dim g_status As String
    Dim g_workspaceID As Long
    Dim g_rfdSheetID As Long
    Dim g_workspace As Workspace
    Dim g_smartSheet As SmartsheetClient
    Dim g_token As New Token
    Dim g_attachmentLink As String
    Dim g_partNumber As String
    Dim g_pdfPath As String
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
        If panelAttachments.Controls.Count > 0 Then
            panelAttachments.Controls.Clear()
        End If
        If panelJobInfo.Visible Then
            panelJobInfo.Visible = False
        End If

        Dim wait As Boolean = True

        While wait
            Me.Cursor = Cursors.WaitCursor
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
            wait = False
        End While
        Me.Cursor = Cursors.Arrow

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
        Try
            Dim wait As Boolean = True
            While wait
                Me.Cursor = Cursors.WaitCursor
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
                tbNotes.Text = gotRow.Cells(14).Value
                lblSAPPN.Text = gotRow.Cells(6).Value

                'Start Date, End Date, Time, Number of Parts
                If gotRow.Cells(16).Value IsNot Nothing Then
                    tbStartDate.Text = gotRow.Cells(16).Value
                Else
                    tbStartDate.Text = ""
                End If
                If gotRow.Cells(17).Value IsNot Nothing Then
                    tbEndDate.Text = gotRow.Cells(17).Value
                Else
                    tbEndDate.Text = ""
                End If
                If gotRow.Cells(18).Value IsNot Nothing Then
                    tbTotalTime.Text = gotRow.Cells(18).Value
                Else
                    tbTotalTime.Text = ""
                End If
                If gotRow.Cells(19).Value IsNot Nothing Then
                    tbNumberofParts.Text = gotRow.Cells(19).Value
                Else
                    tbNumberofParts.Text = ""
                End If

                'Teak, Logo, Laser, Swapout Checkboxes
                If gotRow.Cells(9).Value = "True" Then
                    checkTeak.Checked = True
                ElseIf gotRow.Cells(9).Value = Nothing Then
                    checkTeak.Checked = False
                End If
                If gotRow.Cells(10).Value = "True" Then
                    checkLogo.Checked = True
                ElseIf gotRow.Cells(10).Value = Nothing Then
                    checkLogo.Checked = False
                End If
                If gotRow.Cells(11).Value = "True" Then
                    checkLaser.Checked = True
                ElseIf gotRow.Cells(11).Value = Nothing Then
                    checkLaser.Checked = False
                End If
                If gotRow.Cells(13).Value = "True" Then
                    checkSwapout.Checked = True
                ElseIf gotRow.Cells(13).Value = Nothing Then
                    checkSwapout.Checked = False
                End If

                Dim attachment As PaginatedResult(Of Attachment) = g_smartSheet.SheetResources.RowResources.AttachmentResources.ListAttachments(g_sheetID, rowIDs(e.ListViewElement.SelectedIndex), Nothing)
                Dim rfdAttachment As New Attachment()

                If attachment.TotalCount > 0 Then
                    For index As Integer = 0 To attachment.TotalCount - 1
                        rfdAttachment = g_smartSheet.SheetResources.AttachmentResources.GetAttachment(g_sheetID, attachment.Data(index).Id)
                        If rfdAttachment IsNot Nothing Then
                            Dim button As New RadButton()
                            'button.Text = rfdAttachment.MimeType
                            button.Size = New Drawing.Size(30, 30)
                            button.Dock = DockStyle.Left
                            button.ButtonElement.ImageAlignment = Drawing.ContentAlignment.MiddleCenter
                            button.Name = rfdAttachment.Url
                            button.ButtonElement.ButtonFillElement.ForeColor = Drawing.Color.Transparent
                            button.ButtonElement.ButtonFillElement.BackColor = Drawing.Color.Transparent
                            button.ButtonElement.Margin = New Padding(2, 0, 2, 0)
                            If rfdAttachment.MimeType = "application/pdf" Then
                                button.Image = My.Resources.pdf32
                            End If
                            If rfdAttachment.MimeType = "image/png" Then
                                button.Image = My.Resources.png32
                            End If
                            AddHandler button.Click, AddressOf AttachmentClick
                            panelAttachments.Controls.Add(button)
                            'Debugger.Break()
                        End If
                    Next
                End If

                g_row = gotRow
                g_rowID = gotRow.Id
                g_columnID_Status = gotRow.Cells(1).ColumnId
                g_columnID_StartDate = gotRow.Cells(16).ColumnId
                g_columnID_EndDate = gotRow.Cells(17).ColumnId
                g_columnID_Time = gotRow.Cells(18).ColumnId
                g_columnID_NumberOfParts = gotRow.Cells(19).ColumnId
                g_sheetID = gotRow.SheetId

                wait = False
            End While
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            Me.Cursor = Cursors.Arrow
            MsgBox("There was an error getting the job details. Error: " + ex.Message)
            Return
        End Try
    End Sub

    Private Sub AttachmentClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim button As RadButton = CType(sender, RadButton)
        Process.Start(button.Name)
    End Sub


    ''' <summary>
    ''' Update the smartsheet for the given user
    ''' and the given workspaceid
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Dim ss As New SmartSheetIntegration()

        Dim startDate As New Date()
        startDate = tbStartDate.Text

        Dim endDate As New Date()
        endDate = tbEndDate.Text

        Dim cell As New List(Of Cell)
        cell.Add(New Cell.UpdateCellBuilder(g_columnID_Status, g_status).Build())
        cell.Add(New Cell.UpdateCellBuilder(g_columnID_NumberOfParts, tbNumberofParts.Text).Build())
        cell.Add(New Cell.UpdateCellBuilder(g_columnID_Time, tbTotalTime.Text).Build())
        cell.Add(New Cell.UpdateCellBuilder(g_columnID_StartDate, startDate).Build())
        cell.Add(New Cell.UpdateCellBuilder(g_columnID_EndDate, endDate).Build())

        Dim row As Row = New Row.UpdateRowBuilder(g_rowID).SetCells(cell).Build()

        ss.smartSheet = g_smartSheet
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

    Private Sub panelJobInfo_Paint(sender As Object, e As PaintEventArgs) Handles panelJobInfo.Paint

    End Sub

    Private Sub btnAddAttachment_Click(sender As Object, e As EventArgs) Handles btnAddAttachment.Click
        Dim fileOpen As New OpenFileDialog()
        Dim filePaths() As String
        Dim folderPrefix = GetFolder(lblSAPPN.Text.Substring(2, 3))
        Dim partPath = lblSAPPN.Text.Substring(0, 11)

        fileOpen.InitialDirectory = String.Format("\\surfer\cnc\{0}\{1}\", folderPrefix, partPath)
        fileOpen.Multiselect = True
        fileOpen.RestoreDirectory = True

        Dim ss As New SmartSheetIntegration()

        If fileOpen.ShowDialog = DialogResult.OK Then
            filePaths = fileOpen.FileNames
        Else
            Return
        End If

        Try
            Dim wait As Boolean = True
            While wait
                Me.Cursor = Cursors.WaitCursor
                For Each file As String In filePaths
                    ss.AddAttachment(g_smartSheet, g_sheetID, g_rowID, file)
                Next
                wait = False
            End While
        Catch ex As Exception
            MsgBox("There was an error " + ex.Message)
            Me.Cursor = Cursors.Arrow
            Return
        End Try

        Me.Cursor = Cursors.Arrow
        MsgBox("All attachments have been uploaded to your smartsheet.")
        RefreshAttachmentList()

    End Sub

    Public Sub RefreshAttachmentList()
        If panelAttachments.Controls.Count > 0 Then
            panelAttachments.Controls.Clear()
        End If
        Dim attachment As PaginatedResult(Of Attachment) = g_smartSheet.SheetResources.RowResources.AttachmentResources.ListAttachments(g_sheetID, rowIDs(listJobQue.SelectedIndex), Nothing)
        Dim rfdAttachment As New Attachment()

        If attachment.TotalCount > 0 Then
            For index As Integer = 0 To attachment.TotalCount - 1
                rfdAttachment = g_smartSheet.SheetResources.AttachmentResources.GetAttachment(g_sheetID, attachment.Data(index).Id)
                If rfdAttachment IsNot Nothing Then
                    Dim button As New RadButton()
                    'button.Text = rfdAttachment.MimeType
                    button.Size = New Drawing.Size(30, 30)
                    button.Dock = DockStyle.Left
                    button.ButtonElement.ImageAlignment = Drawing.ContentAlignment.MiddleCenter
                    button.Name = rfdAttachment.Url
                    button.ButtonElement.ButtonFillElement.ForeColor = Drawing.Color.Transparent
                    button.ButtonElement.ButtonFillElement.BackColor = Drawing.Color.Transparent
                    button.ButtonElement.Margin = New Padding(2, 0, 2, 0)
                    If rfdAttachment.MimeType = "application/pdf" Then
                        button.Image = My.Resources.pdf32
                    End If
                    If rfdAttachment.MimeType = "image/png" Then
                        button.Image = My.Resources.png32
                    End If
                    AddHandler button.Click, AddressOf AttachmentClick
                    panelAttachments.Controls.Add(button)
                    'Debugger.Break()
                End If
            Next
        End If
    End Sub


    'GET PROJECT FOLDER
    Private Function GetFolder(prefix)
        Dim officeDB As String = "Data Source=\\surfer\public\CAD\RhinoDek\SQLITE\MainDatabase.db"
        Dim myQury As String = String.Format("SELECT * FROM ProjectFolders WHERE Prefix LIKE '{0}'", prefix)
        Dim conn As New SQLiteConnection(officeDB, True)
        Dim command As New SQLiteCommand(myQury, conn)
        Try
            conn.Open()
            Dim reader As SQLiteDataReader
            reader = command.ExecuteReader()
            reader.Read()

            If reader.HasRows Then
                Return reader.Item("Name").ToString
                reader.Close()
                conn.Close()
            Else
                reader.Close()
                conn.Close()
                Return Nothing
            End If
        Catch ex As Exception
            conn.Close()
            MsgBox(ex.Message)
        End Try
    End Function



End Class
