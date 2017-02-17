Imports System.IO
Imports System.Windows.Forms
Imports Smartsheet.Api
Imports Smartsheet.Api.Models
Imports Smartsheet.Api.OAuth


Public Class SmartsheetForm

    Dim rowID As Long
    Dim selectedRow As Row
    Dim sapPartnumber As String = ""
    Dim rowIDs As New List(Of Long)



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
        Dim gotRows = ss.GetUserRFD(3083654818752388, cbUsers.Text)
        For Each row As Row In gotRows
            rowIDs.Add(row.Id)
            listJobQue.Items.Add(row.Cells(6).Value)
        Next
    End Sub

    'Selected Item Clicked on (You clicked an item in the list)
    Private Sub listJobQue_ItemMouseClick(sender As Object, e As Telerik.WinControls.UI.ListViewItemEventArgs) Handles listJobQue.ItemMouseClick
        sapPartnumber = e.Item.Text
        Dim ss = New SmartSheetIntegration()
        Dim gotRow As Row = ss.GetRow(3083654818752388, cbUsers.Text, rowIDs(e.ListViewElement.SelectedIndex))

        dateTimePicker.Text = gotRow.Cells(4).Value
        comboStatus.Text = gotRow.Cells(1).Value
        tbSubmittedBy.Text = gotRow.Cells(3).Value
        tbCustomerType.Text = gotRow.Cells(5).Value
        tbCustomerName.Text = gotRow.Cells(7).Value
        tbMakeModelYear.Text = gotRow.Cells(8).Value
        tbNotes.Text = gotRow.Cells(9).Value
        tbSAPPartNumber.Text = gotRow.Cells(6).Value

    End Sub
End Class
