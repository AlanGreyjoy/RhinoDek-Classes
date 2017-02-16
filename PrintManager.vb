Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports Rhino
Imports Rhino.Commands
Imports Rhino.Geometry
Imports RhinoScript4
Imports System.Windows.Forms.SaveFileDialog
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Windows.Forms.Form
Imports Bullzip
Imports System.Drawing
Imports System.Data.SQLite
Imports AutoItX3Lib
Imports System.Threading

Public Class PrintManager

    'Global Dims
    Dim topLeftPartViewPoint, bottomRightPartViewPoint
    Dim topleftTableViewPoint, bottomRightTableViewPoint
    Dim topLeftPartViewPointCoor, bottomRightPartViewPointCoor
    Dim topLeftTableViewPointCoor, bottomRightTableViewPointCoor
    Dim GlobalSeaDekPartNumber
    Dim globalColorAddedBool As Boolean
    Dim globalOverallMM
    Dim globalOverallMMStrToInt
    Dim noPreview As Boolean

    Dim settings As New RhinoDekSettings()

    Dim globalFileName
    Dim origFileName

    Dim p1SB As New StringBuilder()
    Dim p2SB As New StringBuilder()
    Dim p3SB As New StringBuilder()
    Dim p4SB As New StringBuilder()

    'Start Rhinoscriptsyntax
    Dim obj As Object = Rhino.RhinoApp.GetPlugInObject("RhinoScript")
    Dim rs As RhinoScript4.IRhinoScript = CType(obj, RhinoScript4.IRhinoScript)

    Dim dbaPartNumber, partViewFileName, tableViewPartName, topSheetColor, bottomSheetColor

    Dim au3 As New AutoItX3

    'Reset the Part View and Talbe View Stringbuilders to empty. Needed!
    Public Function ResetPoints() As Result

        p1SB.Clear()
        p2SB.Clear()
        p3SB.Clear()
        p4SB.Clear()

        Return Result.Success
    End Function

    'START UP - The main function for this form. It starts everything we need at the forms first run
    Public Function StartPrintManager() As Result

        topLeftPartViewPoint = rs.ObjectsByName("TOP LEFT PART VIEW POINT")
        bottomRightPartViewPoint = rs.ObjectsByName("BOTTOM RIGHT PART VIEW POINT")
        topleftTableViewPoint = rs.ObjectsByName("TOP LEFT TABLE VIEW POINT")
        bottomRightTableViewPoint = rs.ObjectsByName("BOTTOM RIGHT TABLE VIEW POINT")

        If IsDBNull(topLeftPartViewPoint) Then
            rs.Print("The print points are not setup!")
            TBDPanel.Visible = False
            ColorPanel.Visible = False
            rhinoPic.Visible = True
            PrintingPointsLabel.Visible = True
            Return Result.Cancel
        Else
            rhinoPic.Visible = False
            PrintingPointsLabel.Visible = False

            topLeftPartViewPointCoor = rs.PointCoordinates(topLeftPartViewPoint(0))
            bottomRightPartViewPointCoor = rs.PointCoordinates(bottomRightPartViewPoint(0))

            topLeftTableViewPointCoor = rs.PointCoordinates(topleftTableViewPoint(0))
            bottomRightTableViewPointCoor = rs.PointCoordinates(bottomRightTableViewPoint(0))
        End If

        globalOverallMM = rs.ObjectsByName("OVERALL MM INT")
        If IsDBNull(globalOverallMM) Then
            rs.Print("There is no BOM in the drawing!")
        Else
            rs.Print("Found BOM :)")
            Dim overallMMText = rs.TextObjectText(globalOverallMM(0))
            globalOverallMMStrToInt = CInt(overallMMText)
        End If


        'TBD Printers
        For Each printer In PrinterSettings.InstalledPrinters
            tbdPrinterComboBox.Items.Add(printer)
        Next printer
        'COLOR Printers
        For Each printer In PrinterSettings.InstalledPrinters
            colorPrinterComboBox.Items.Add(printer)
        Next printer

        Dim seaDekPN = rs.ObjectsByName("SEADEK PART NUMBER")
        Dim seaDekPNText = rs.TextObjectText(seaDekPN(0))
        GlobalSeaDekPartNumber = seaDekPNText

        globalColorAddedBool = False

        Dim getFileName = rs.ObjectsByName("FILE NAME")
        Dim fileNameText = rs.TextObjectText(getFileName(0))
        globalFileName = fileNameText
        origFileName = fileNameText

        Dim projectColorSuffix = rs.GetDocumentData("ProjectInfo", "ProjectColorSuffix")
        Dim projectTopColor = rs.GetDocumentData("ProjectInfo", "ProjectTopColor")
        Dim projectMiddleColor = rs.GetDocumentData("ProjectInfo", "ProjectMiddleColor")
        Dim projectBottomColor = rs.GetDocumentData("ProjectInfo", "ProjectBottomColor")

        If IsDBNull(projectColorSuffix) Then
        Else
            colorSuffixTextbox.Text = projectColorSuffix
        End If
        If IsDBNull(projectTopColor) Then
        Else
            TopSheetColorTextbox.Text = projectTopColor
        End If
        If IsDBNull(projectMiddleColor) Then
        Else
            tbMiddleSheetColor.Text = projectMiddleColor
        End If
        If IsDBNull(projectBottomColor) Then
        Else
            BottomSheetColorTextbox.Text = projectBottomColor
        End If


        Return Result.Success
    End Function

    'Print TBD's
    Private Sub PrintTBD_Click(sender As Object, e As EventArgs) Handles PrintTBD.Click
        ResetPoints()

        If tbdPrinterComboBox.Text = "" Then
            rs.MessageBox("You must select a printer first",, "RhinoDek - Printer Error")
            Return
        End If

        ResetPoints()

        If tbdPartViewCheckbox.Checked Then
            If tbdPrinterComboBox.Text = "Bullzip PDF Printer" Then
                If noPreview = True Then
                    PrintTBDPDFPartView()
                Else
                    Dim newThread As New Thread(AddressOf GetPrintWindow)
                    newThread.Start()
                    PrintTBDPDFPartView()
                End If
            Else
                If noPreview = True Then
                    PrintPartView()
                Else
                    Dim newThread As New Thread(AddressOf GetPrintWindow)
                    newThread.Start()
                    PrintPartView()
                End If
            End If
        End If

        If tbdTableViewCheckbox.Checked Then
            If tbdPrinterComboBox.Text = "Bullzip PDF Printer" Then
                Dim newThread As New Thread(AddressOf GetPrintWindow)
                newThread.Start()
                PrintTBDPDFTableView()
            Else
                Dim newThread As New Thread(AddressOf GetPrintWindow)
                newThread.Start()
                PrintTableView()
            End If

        End If

    End Sub

    'Print Colors
    Private Sub PrintColorButton_Click(sender As Object, e As EventArgs) Handles PrintColorButton.Click
        ResetPoints()
        If colorPrinterComboBox.Text = "" Then
            rs.MessageBox("You must select a printer first",, "RhinoDek - Printer Error")
            Return
        End If
        ResetPoints()
        If colorPartViewCheckbox.Checked AndAlso colorTableViewCheckbox.Checked Then
            MsgBox("Unable to do select both right now.")

            Return
        End If

        If colorPartViewCheckbox.Checked Then
            If colorPrinterComboBox.Text = "Bullzip PDF Printer" Then
                PrintColorPDFPartView()
            Else
                PrintPartViewColor()
            End If
        End If

        If colorTableViewCheckbox.Checked Then
            If colorPrinterComboBox.Text = "Bullzip PDF Printer" Then
                PrintColorPDFTableView()
            Else
                PrintTableViewColor()
            End If
        End If

    End Sub

    Private Sub ClosePrintsFormButton_Click(sender As Object, e As EventArgs) Handles ClosePrintsFormButton.Click
        ResetPoints()
        ResetValues()
        Close()
    End Sub

    'Reset Project Text details to before color information was added
    Public Function ResetValues() As Result

        Dim fileName, tvFileName, seaDekPN
        Dim topSheetColor
        Dim middleSheetColorObj
        Dim bottomSheetColor
        Dim oneSheetColorObject

        topSheetColor = rs.ObjectsByName("TOP SHEET COLOR")
        middleSheetColorObj = rs.ObjectsByName("MIDDLE SHEET COLOR")
        bottomSheetColor = rs.ObjectsByName("BOTTOM SHEET COLOR")
        oneSheetColorObject = rs.ObjectsByName("ONE SHEET COLOR")

        Try
            seaDekPN = rs.ObjectsByName("SEADEK PART NUMBER")
            For Each obj In seaDekPN
                rs.TextObjectText(obj, GlobalSeaDekPartNumber)
            Next


            fileName = rs.ObjectsByName("FILE NAME")
            For Each obj In fileName
                rs.TextObjectText(obj, "%<filename(""3"")>%")
            Next


            tvFileName = rs.ObjectsByName("TV FILE NAME")
            For Each obj In tvFileName
                rs.TextObjectText(obj, "%<filename(""3"")>%")
            Next

            If IsDBNull(topSheetColor) Then
                rs.Print("No top sheet color")
            Else
                For Each obj In topSheetColor
                    rs.TextObjectText(obj, "TBD")
                Next obj
            End If

            If IsDBNull(bottomSheetColor) Then
                rs.Print("No bottom sheet color")
            Else
                For Each obj In bottomSheetColor
                    rs.TextObjectText(obj, "TBD")
                Next obj
            End If

            If IsDBNull(middleSheetColorObj) Then
                rs.Print("No middle sheet to reset.")
            Else
                For Each obj In middleSheetColorObj
                    rs.TextObjectText(obj, "TBD")
                Next
            End If

            If IsDBNull(oneSheetColorObject) Then
                rs.Print("No one sheet color!")
            Else
                For Each obj In oneSheetColorObject
                    rs.TextObjectText(obj, "TBD")
                Next obj
            End If

            globalColorAddedBool = False

        Catch ex As Exception
            MsgBox(ex.ToString + System.Environment.NewLine + System.Environment.NewLine + "Please report this error to Chad Bryant.")
            Return Result.Cancel
        End Try

        Return Result.Success

    End Function

    'Set and Print Part View
    Public Function PrintPartView() As Result

        Dim i As Integer = 0
        While i < 2
            p1SB.Append(topLeftPartViewPointCoor(i))
            p1SB.Append(",")
            i = i + 1
        End While
        p1SB.Append(topLeftPartViewPointCoor(2))

        Dim ii As Integer = 0
        While ii < 2
            p2SB.Append(bottomRightPartViewPointCoor(ii))
            p2SB.Append(",")
            ii = ii + 1
        End While
        p2SB.Append(bottomRightPartViewPointCoor(2))
        Dim strPrinter = Chr(34) & tbdPrinterComboBox.Text & Chr(34)
        If noPreview Then
            Rhino.RhinoApp.RunScript("-Print Setup View ViewportArea=Window " + p1SB.ToString() + " " + p2SB.ToString() + " " + "Enter Enter Destination Printer " + strPrinter + " Enter Enter GO", True)
        Else
            Rhino.RhinoApp.RunScript("-Print Setup View ViewportArea=Window " + p1SB.ToString() + " " + p2SB.ToString() + " " + "Enter Enter Destination Printer " + strPrinter + " Enter Enter Preview", True)
        End If
        ResetPoints()
        Return Result.Success
    End Function

    Private Sub tbdPartViewCheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles tbdPartViewCheckbox.CheckedChanged
        If tbdTableViewCheckbox.Checked Then
            tbdTableViewCheckbox.Checked = False
        Else
            tbdPartViewCheckbox.Checked = True
        End If
    End Sub

    Private Sub tbdTableViewCheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles tbdTableViewCheckbox.CheckedChanged
        If tbdPartViewCheckbox.Checked Then
            tbdPartViewCheckbox.Checked = False
        Else
            tbdTableViewCheckbox.Checked = True
        End If
    End Sub

    Private Sub colorPartViewCheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles colorPartViewCheckbox.CheckedChanged
        If colorTableViewCheckbox.Checked Then
            colorTableViewCheckbox.Checked = False
        Else
            colorPartViewCheckbox.Checked = True
        End If
    End Sub

    Private Sub colorTableViewCheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles colorTableViewCheckbox.CheckedChanged
        If colorPartViewCheckbox.Checked Then
            colorPartViewCheckbox.Checked = False
        Else
            colorTableViewCheckbox.Checked = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub PrintManager_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        settings.FormOpen = True
        settings.Save()

        For Each f As Form In Application.OpenForms
            If Not f.InvokeRequired Then
                ' Can access the form directly.
                'Get main form , use main form
                If f.Tag = "main" Then
                    Dim fcast As New MainMenu '<< whatever your form name
                    fcast = f
                    fcast.printsButton.BackColor = Color.FromArgb(255, 198, 70)
                    fcast.breadcrumbLabel.Text = "Print Manager"
                End If

            End If
        Next

        If settings.FastPrint = True Then
            FastPrintsLabel.Visible = True
            FastPrintButton.Visible = True
            noPreview = True
        Else
            FastPrintsLabel.Visible = False
            FastPrintButton.Visible = False
            noPreview = False
        End If
        If settings.PrintPreview = False Then
            noPreview = False
        End If
        If settings.PrintPreview = True Then
            noPreview = True
        End If

    End Sub

    Private Sub PrintManager_Closed(sender As Object, e As EventArgs) Handles MyBase.Closed
        settings.FormOpen = False
        settings.Save()
        For Each f As Form In Application.OpenForms

            If Not f.InvokeRequired Then
                ' Can access the form directly.
                'Get main form , use main form
                If f.Tag = "main" Then
                    Dim fcast As New MainMenu '<< whatever your form name
                    fcast = f
                    fcast.printsButton.BackColor = Color.Transparent
                    fcast.breadcrumbLabel.Text = "Home - Dashboard"
                End If

            End If

        Next
        ResetValues()
    End Sub

    Private Sub TBDPDFRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles TBDPDFRadioButton.CheckedChanged
        Dim settings As New RhinoDekSettings()
        tbdPrinterComboBox.Text = settings.PDF_PRINTER
    End Sub

    Private Sub TBDPaperRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles TBDPaperRadioButton.CheckedChanged
        Dim settings As New RhinoDekSettings()
        tbdPrinterComboBox.Text = settings.PAPER_PRINTER
    End Sub

    Private Sub ColorPDFRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles ColorPDFRadioButton.CheckedChanged
        Dim settings As New RhinoDekSettings()
        colorPrinterComboBox.Text = settings.PDF_PRINTER
    End Sub

    Private Sub ColorPaperRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles ColorPaperRadioButton.CheckedChanged
        Dim settings As New RhinoDekSettings()
        colorPrinterComboBox.Text = settings.PAPER_PRINTER
    End Sub

    Private Sub FastPrintButton_Click(sender As Object, e As EventArgs) Handles FastPrintButton.Click

        'FastPrints
        Dim settings As New RhinoDekSettings()
        Dim flag As Boolean
        If settings.FastPrint = True Then

            tbdPrinterComboBox.Text = settings.PDF_PRINTER
            If PrintTBDPDFPartView() = Result.Success Then
                If PrintTBDPDFTableView() = Result.Success Then
                    colorPrinterComboBox.Text = settings.PDF_PRINTER
                    If PrintColorPDFPartView() = Result.Success Then
                        If PrintColorPDFTableView() = Result.Success Then
                            MsgBox("ALL PDF's have been made!")
                        End If
                    End If
                End If
            End If
        Else
        End If

    End Sub

    'Grab input from textbox and search the database
    Dim oldColorSuffix
    Dim currentColorSuffix
    Dim officeDB As String = "Data Source=\\surfer\public\CAD\RhinoDek\SQLITE\MainDatabase.db"
    Private Sub colorSuffixTextbox_TextChanged(sender As Object, e As EventArgs) Handles colorSuffixTextbox.TextChanged
        If (colorSuffixTextbox.Text.Length >= 5) Then
            currentColorSuffix = colorSuffixTextbox.Text
            GetDoublePly(colorSuffixTextbox.Text)
        End If
    End Sub

    'Search for colors in the color database
    'Get Triple Ply Color Info
    Private Sub GetTriplePly(suffix)
        Dim myQury As String = String.Format("SELECT * FROM TriplePly WHERE ColorSuffix LIKE '{0}'", suffix)
        Dim conn As New SQLiteConnection(officeDB, True)
        Dim command As New SQLiteCommand(myQury, conn)

        conn.Open()
        Dim reader As SQLiteDataReader
        reader = command.ExecuteReader()
        reader.Read()
        Try

            If reader.HasRows Then
                TopSheetColorTextbox.Text = reader.Item("TopSheetColor").ToString
                tbMiddleSheetColor.Text = reader.Item("MiddleSheetColor").ToString
                BottomSheetColorTextbox.Text = reader.Item("BottomSheetColor").ToString

                reader.Close()
                conn.Close()
            Else
                reader.Close()
                conn.Close()
                Return
            End If

        Catch ex As Exception
            reader.Close()
            conn.Close()
            MsgBox(ex.Message)
        End Try
    End Sub

    'GET DOUBLE PLY COLOR INFO
    Private Sub GetDoublePly(suffix)
        Dim myQury As String = String.Format("SELECT * FROM DoublePly WHERE ColorSuffix LIKE '{0}'", suffix)
        Dim conn As New SQLiteConnection(officeDB, True)
        Dim command As New SQLiteCommand(myQury, conn)

        conn.Open()
        Dim reader As SQLiteDataReader
        reader = command.ExecuteReader()
        reader.Read()
        Try

            If reader.HasRows Then
                TopSheetColorTextbox.Text = reader.Item("TopSheetColor").ToString
                BottomSheetColorTextbox.Text = reader.Item("BottomSheetColor").ToString

                reader.Close()
                conn.Close()
            Else
                reader.Close()
                conn.Close()
                GetSinglePly(suffix)
            End If

        Catch ex As Exception
            reader.Close()
            conn.Close()
            MsgBox(ex.Message)
        End Try

    End Sub

    'GET SINGLE PLY COLOR INFO
    Private Sub GetSinglePly(suffix)
        Dim myQury As String = String.Format("SELECT * FROM SinglePly WHERE ColorSuffix LIKE '{0}'", suffix)
        Dim conn As New SQLiteConnection(officeDB, True)
        Dim command As New SQLiteCommand(myQury, conn)

        conn.Open()
        Dim reader As SQLiteDataReader
        reader = command.ExecuteReader()
        reader.Read()
        Try

            If reader.HasRows Then
                TopSheetColorTextbox.Text = reader.Item("TopSheetColor").ToString
                reader.Close()
                conn.Close()
                reader.Close()
                conn.Close()
            Else
                reader.Close()
                conn.Close()
                GetTriplePly(suffix)
            End If

        Catch ex As Exception
            reader.Close()
            conn.Close()
            MsgBox(ex.Message)
        End Try
    End Sub

    'Set and Print Table View
    Public Function PrintTableView() As Result

        Dim a As Integer = 0
        While a < 2
            p3SB.Append(topLeftTableViewPointCoor(a))
            p3SB.Append(",")
            a = a + 1
        End While
        p3SB.Append(topLeftTableViewPointCoor(2))

        Dim aa As Integer = 0
        While aa < 2
            p4SB.Append(bottomRightTableViewPointCoor(aa))
            p4SB.Append(",")
            aa = aa + 1
        End While

        p4SB.Append(bottomRightTableViewPointCoor(2))
        Dim strPrinter = Chr(34) & tbdPrinterComboBox.Text & Chr(34)
        If noPreview Then
            Rhino.RhinoApp.RunScript("-Print Setup View ViewportArea=Window " + p3SB.ToString() + " " + p4SB.ToString() + " " + "Enter Enter Destination Printer " + strPrinter + " Enter Enter GO", True)
        Else
            Rhino.RhinoApp.RunScript("-Print Setup View ViewportArea=Window " + p3SB.ToString() + " " + p4SB.ToString() + " " + "Enter Enter Destination Printer " + strPrinter + " Enter Enter Preview", True)
        End If
        ResetPoints()
        Return Result.Success
    End Function

    Sub GetPrintWindow()
        If au3.WinWait("Print Setup", "", 5) = 0 Then
            MessageBox.Show("No Print Window?")
        End If

        If au3.WinActive("Print Setup") Then
            au3.ControlClick("Print Setup", "", "[CLASSNN:Button9]", "LEFT", 1)
        Else
            au3.WinActivate("Print Setup")
            au3.ControlClick("Print Setup", "", "[CLASSNN:Button9]", "LEFT", 1)
        End If
    End Sub

    'Print TBD Part View PDF
    Public Function PrintTBDPDFPartView() As Result
        SetPDFPath(False, True, "PartView")
        'PrintPartView()
        Return Result.Success
    End Function

    'Print TBD Table View PDF
    Public Function PrintTBDPDFTableView() As Result
        SetPDFPath(False, True, "TableView")
        'PrintTableView()
        Return Result.Success
    End Function

    'Print Color PDF Part View
    Dim partnumberAdded = False
    Public Function PrintColorPDFPartView() As Result

        Dim seaDekPartNumberObject, fileNameObject, tableViewFileNameObject

        seaDekPartNumberObject = rs.ObjectsByName("SEADEK PART NUMBER")
        fileNameObject = rs.ObjectsByName("FILE NAME")
        tableViewFileNameObject = rs.ObjectsByName("TV FILE NAME")

        If partnumberAdded = False Then
            Dim getSeaDekPN = rs.TextObjectText(seaDekPartNumberObject(0))
            GlobalSeaDekPartNumber = getSeaDekPN

            For Each obj In seaDekPartNumberObject
                rs.TextObjectText(obj, getSeaDekPN + "-" + colorSuffixTextbox.Text)
            Next

            Dim getFileName = rs.TextObjectText(fileNameObject(0))

            For Each obj In fileNameObject
                rs.TextObjectText(obj, getFileName + "-" + colorSuffixTextbox.Text)
            Next
            For Each obj In tableViewFileNameObject
                rs.TextObjectText(obj, getFileName + "-" + colorSuffixTextbox.Text)
            Next
            partnumberAdded = True
        Else
            ResetValues()
            Dim getSeaDekPN = rs.TextObjectText(seaDekPartNumberObject(0))
            GlobalSeaDekPartNumber = getSeaDekPN

            For Each obj In seaDekPartNumberObject
                rs.TextObjectText(obj, getSeaDekPN + "-" + colorSuffixTextbox.Text)
            Next

            Dim getFileName = rs.TextObjectText(fileNameObject(0))

            For Each obj In fileNameObject
                rs.TextObjectText(obj, getFileName + "-" + colorSuffixTextbox.Text)
            Next
            For Each obj In tableViewFileNameObject
                rs.TextObjectText(obj, getFileName + "-" + colorSuffixTextbox.Text)
            Next
        End If

        Dim topColorObject
        Dim middleColorObject
        Dim bottomColorObject
        Dim oneSheetColorObject

        oneSheetColorObject = rs.ObjectsByName("ONE SHEET COLOR")
        topColorObject = rs.ObjectsByName("TOP SHEET COLOR")
        middleColorObject = rs.ObjectsByName("MIDDLE SHEET COLOR")
        bottomColorObject = rs.ObjectsByName("BOTTOM SHEET COLOR")

        'Single Ply
        If tbMiddleSheetColor.Text.Length = 0 And BottomSheetColorTextbox.Text.Length = 0 Then
            If IsDBNull(oneSheetColorObject) Then
                Return Result.Cancel
            Else
                For Each obj In oneSheetColorObject
                    rs.TextObjectText(obj, TopSheetColorTextbox.Text)
                Next
                globalColorAddedBool = True
            End If
        End If

        'Double Ply
        If TopSheetColorTextbox.Text.Length >= 1 And BottomSheetColorTextbox.Text.Length >= 1 And tbMiddleSheetColor.Text.Length = 0 Then
            If IsDBNull(topColorObject) And IsDBNull(bottomColorObject) Then
                Return Result.Cancel
            Else
                For Each obj In topColorObject
                    rs.TextObjectText(obj, TopSheetColorTextbox.Text)
                Next
                For Each obj In bottomColorObject
                    rs.TextObjectText(obj, BottomSheetColorTextbox.Text)
                Next
                globalColorAddedBool = True
            End If
        End If

        'Triple Ply
        If TopSheetColorTextbox.Text.Length >= 1 And tbMiddleSheetColor.Text.Length >= 1 And BottomSheetColorTextbox.Text.Length >= 1 Then
            If IsDBNull(topColorObject) And IsDBNull(bottomColorObject) And IsDBNull(middleColorObject) Then
                Return Result.Cancel
            Else
                For Each obj In topColorObject
                    rs.TextObjectText(obj, TopSheetColorTextbox.Text)
                Next
                For Each obj In middleColorObject
                    rs.TextObjectText(obj, tbMiddleSheetColor.Text)
                Next
                For Each obj In bottomColorObject
                    rs.TextObjectText(obj, BottomSheetColorTextbox.Text)
                Next
                globalColorAddedBool = True
            End If
        End If

        SetPDFPath(True, False, "PartView")

        Return Result.Success
    End Function

    'Print Color PDF Table View
    Public Function PrintColorPDFTableView() As Result
        SetPDFPath(True, False, "TableView")
        'PrintTableViewColor()
        Return Result.Success
    End Function

    'Print Color Part View (NOT PDF - Paper)
    Public Function PrintPartViewColor() As Result
        Dim newThread As New Thread(AddressOf GetPrintWindow)
        newThread.Start()

        Dim i As Integer = 0
        While i < 2
            p1SB.Append(topLeftPartViewPointCoor(i))
            p1SB.Append(",")
            i = i + 1
        End While
        p1SB.Append(topLeftPartViewPointCoor(2))

        Dim ii As Integer = 0
        While ii < 2
            p2SB.Append(bottomRightPartViewPointCoor(ii))
            p2SB.Append(",")
            ii = ii + 1
        End While
        p2SB.Append(bottomRightPartViewPointCoor(2))

        Dim strPrinter = Chr(34) & colorPrinterComboBox.Text & Chr(34)
        If noPreview Then
            Rhino.RhinoApp.RunScript("-Print Setup View ViewportArea=Window " + p1SB.ToString() + " " + p2SB.ToString() + " " + "Enter Enter Destination Printer " + strPrinter + " Enter Enter GO", True)
        Else
            Rhino.RhinoApp.RunScript("-Print Setup View ViewportArea=Window " + p1SB.ToString() + " " + p2SB.ToString() + " " + "Enter Enter Destination Printer " + strPrinter + " Enter Enter Preview", True)
        End If
        ResetPoints()

        Return Result.Success
    End Function

    'Print Color Talbe View (NOT PDF - Paper)
    Public Function PrintTableViewColor() As Result
        Dim newThread As New Thread(AddressOf GetPrintWindow)
        newThread.Start()

        Dim a As Integer = 0
        While a < 2
            p3SB.Append(topLeftTableViewPointCoor(a))
            p3SB.Append(",")
            a = a + 1
        End While
        p3SB.Append(topLeftTableViewPointCoor(2))

        Dim aa As Integer = 0
        While aa < 2
            p4SB.Append(bottomRightTableViewPointCoor(aa))
            p4SB.Append(",")
            aa = aa + 1
        End While
        p4SB.Append(bottomRightTableViewPointCoor(2))
        Dim strPrinter = Chr(34) & colorPrinterComboBox.Text & Chr(34)
        If noPreview Then
            Rhino.RhinoApp.RunScript("-Print Setup View ViewportArea=Window " + p3SB.ToString() + " " + p4SB.ToString() + " " + "Enter Enter Destination Printer " + strPrinter + " Enter Enter GO", True)
        Else
            Rhino.RhinoApp.RunScript("-Print Setup View ViewportArea=Window " + p3SB.ToString() + " " + p4SB.ToString() + " " + "Enter Enter Destination Printer " + strPrinter + " Enter Enter Preview", True)
        End If
        ResetPoints()

        Return Result.Success
    End Function





    'Set PDF Path
    Public Function SetPDFPath(ByRef color As Boolean, ByRef tbd As Boolean, ByRef view As String) As Result

        Dim settings As New RhinoDekSettings()
        Dim getProjectFolder = rs.GetDocumentData("ProjectInfo", "ProjectFolder")

        Dim fileNameObject = rs.ObjectsByName("FILE NAME")
        Dim fileNameText = rs.TextObjectText(fileNameObject(0))

        Dim tbdPDFpath = "\\surfer\cnc\" + getProjectFolder + "\" + origFileName + "\"
        Dim colorPDFPath = "\\surfer\cnc\" + getProjectFolder + "\" + origFileName + "\" + origFileName + "-" + colorSuffixTextbox.Text + "\"

        If color = False Then
            If tbd = True Then
                If view = "PartView" Then

                    Dim userINI As String
                    userINI = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\user.ini"
                    Dim userSettings As String
                    userSettings = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\settings.ini"

                    If Not Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\") Then
                        Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\")
                    End If
                    If Not File.Exists(userINI) Then
                        MsgBox("Creating user.ini file")
                        File.Create(userINI)
                    End If
                    If Not File.Exists(userSettings) Then
                        MsgBox("Creating usersettings.ini file")
                        File.Create(userSettings)
                    End If

                    Dim pdf As New PDFPrinterSettings()
                    pdf.SetValue("ShowPDF", "no")
                    Dim output As String = tbdPDFpath + origFileName + "-" + "PV" + ".pdf"
                    pdf.SetValue("Output", "")
                    pdf.SetValue("Output", output)

                    If settings.BullZipNewFolder = True Then
                        pdf.SetValue("confirmnewfolder", "no")
                    End If
                    If settings.BullZipOverwrite = True Then
                        pdf.SetValue("confirmoverwrite", "no")
                    End If
                    If settings.BullZipSaveDialog = True Then
                        pdf.SetValue("showsettings", "never")
                        pdf.SetValue("showsaveas", "never")
                    End If

                    pdf.WriteSettings(True)

                    PrintPartView()

                ElseIf view = "TableView" Then

                    Dim userINI As String
                    userINI = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\user.ini"

                    If Not Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\") Then
                        Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\")
                    End If
                    If Not File.Exists(userINI) Then
                        File.Create(userINI)
                    End If

                    Dim pdf As New PDFPrinterSettings()
                    pdf.SetValue("ShowPDF", "no")
                    pdf.SetValue("Output", "")
                    pdf.SetValue("output", tbdPDFpath + origFileName + "-" + "TV" + ".pdf")
                    If settings.BullZipNewFolder = True Then
                        pdf.SetValue("confirmnewfolder", "no")
                    End If
                    If settings.BullZipOverwrite = True Then
                        pdf.SetValue("confirmoverwrite", "no")
                    End If
                    If settings.BullZipSaveDialog = True Then
                        pdf.SetValue("showsettings", "never")
                        pdf.SetValue("showsaveas", "never")
                    End If
                    pdf.WriteSettings(True)

                    PrintTableView()

                End If
            End If
        End If


        If color = True Then
            If tbd = False Then
                If view = "PartView" Then

                    Dim userINI As String
                    userINI = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\user.ini"

                    If Not Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\") Then
                        Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\")
                    End If
                    If Not File.Exists(userINI) Then
                        File.Create(userINI)
                    End If

                    Dim pdf As New PDFPrinterSettings()
                    pdf.SetValue("Output", "")
                    pdf.SetValue("output", colorPDFPath + origFileName + "-" + colorSuffixTextbox.Text + "-" + "PV" + ".pdf")
                    pdf.SetValue("ShowPDF", "no")
                    If settings.BullZipNewFolder = True Then
                        pdf.SetValue("confirmnewfolder", "no")
                    End If
                    If settings.BullZipOverwrite = True Then
                        pdf.SetValue("confirmoverwrite", "no")
                    End If
                    If settings.BullZipSaveDialog = True Then
                        pdf.SetValue("showsettings", "never")
                        pdf.SetValue("showsaveas", "never")
                    End If
                    pdf.WriteSettings(True)

                    PrintPartViewColor()

                ElseIf view = "TableView" Then

                    Dim userINI As String
                    userINI = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\user.ini"

                    If Not Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\") Then
                        Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\PDF Writer\Bullzip PDF Printer\")
                    End If
                    If Not File.Exists(userINI) Then
                        File.Create(userINI)
                    End If

                    Dim pdf As New PDFPrinterSettings()
                    pdf.SetValue("Output", "")
                    pdf.SetValue("output", colorPDFPath + origFileName + "-" + colorSuffixTextbox.Text + "-" + "TV" + ".pdf")
                    pdf.SetValue("ShowPDF", "no")
                    If settings.BullZipNewFolder = True Then
                        pdf.SetValue("confirmnewfolder", "no")
                    End If
                    If settings.BullZipOverwrite = True Then
                        pdf.SetValue("confirmoverwrite", "no")
                    End If
                    If settings.BullZipSaveDialog = True Then
                        pdf.SetValue("showsettings", "never")
                        pdf.SetValue("showsaveas", "never")
                    End If
                    pdf.WriteSettings(True)

                    PrintTableViewColor()

                End If
            End If
        End If

        Return Result.Success
    End Function

End Class
