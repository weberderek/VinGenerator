Imports Microsoft.Office.Interop
Public Class Form1
    'directory that contains the Work Order Template Folder 
    Private TemplateFilePath = "G:\Shared drives\Public\Work_Order_Generator\"
    'Private TemplateFilePath = "C:\Users\derek.weber\source\repos\WorkOrder\WorkOrder\"
    Private logFile = TemplateFilePath + "WorkOrders\log.txt"
    Private NumFiles = 0
    Private filesCreated = 0
    Private strFileName As String

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Main(strFileName)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        MsgBox("Done")
        Application.Exit()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        fd.Filter = "excel files (*.xlsx*)|*.xlsx*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    Private Sub SubAssembly(prefix As String, vins As List(Of String), dates As List(Of String), fileName As String)
        Dim xlsapp As Excel.Application
        Dim xlsworkbook As Excel.Workbook
        Dim xlsworksheet As Excel.Worksheet
        xlsapp = New Excel.Application
        xlsworkbook = xlsapp.Workbooks.Open(TemplateFilePath + "\work order templates\" + fileName)
        xlsworksheet = xlsworkbook.Worksheets(2)
        Dim numRows = xlsworksheet.UsedRange.Rows.Count - 1
        Dim currentRow = 2
        Dim xlscell As Excel.Range

        Dim PSN As String
        Dim WOPROD As String
        Dim SOPROD As String
        Dim temp As String
        Dim DirectoryPathFolder As String
        DirectoryPathFolder = TemplateFilePath + "WorkOrders\Subassemblies\"
        Dim DirectoryPathFile As String
        For i As Integer = 0 To vins.Count() - 1
            xlsworksheet = xlsworkbook.Worksheets(1)
            Dim ch15 As Char = vins.ElementAt(i)(14)
            Dim ch16 As Char = vins.ElementAt(i)(15)
            Dim ch17 As Char = vins.ElementAt(i)(16)
            Dim row As String = (i + 2).ToString()
            PSN = ch15 + ch16 + ch17

            WOPROD = "WO-PROD-" + PSN
            SOPROD = "SO-" + PSN
            'move variables and assign to cells
            xlscell = xlsworksheet.Range("a" + row)
            xlscell.Value = prefix + WOPROD
            xlscell = xlsworksheet.Range("b" + row)
            xlscell.Value = prefix + vins.ElementAt(i)
            xlscell = xlsworksheet.Range("c2")
            temp = xlscell.Value
            xlscell = xlsworksheet.Range("c" + row)
            xlscell.Value = temp
            xlscell = xlsworksheet.Range("d2")
            temp = xlscell.Value
            xlscell = xlsworksheet.Range("d" + row)
            xlscell.Value = temp
            xlscell = xlsworksheet.Range("e2")
            temp = xlscell.Value
            xlscell = xlsworksheet.Range("e" + row)
            xlscell.Value = temp
            xlscell = xlsworksheet.Range("f" + row)
            xlscell.Value = dates.ElementAt(i)
            xlscell = xlsworksheet.Range("g2")
            temp = xlscell.Value
            xlscell = xlsworksheet.Range("g" + row)
            xlscell.Value = temp

            'creates all work instructions for vin
            For j As Integer = 2 To numRows + 1
                xlsworksheet = xlsworkbook.Worksheets(2)
                xlscell = xlsworksheet.Range("a" + currentRow.ToString())
                xlscell.Value = prefix + vins.ElementAt(i)
                xlscell = xlsworksheet.Range("b" + j.ToString())
                temp = xlscell.Value
                xlscell = xlsworksheet.Range("b" + currentRow.ToString())
                xlscell.Value = temp
                xlscell = xlsworksheet.Range("c" + j.ToString())
                temp = xlscell.Value
                xlscell = xlsworksheet.Range("c" + currentRow.ToString())
                xlscell.Value = temp
                xlscell = xlsworksheet.Range("d" + j.ToString())
                temp = xlscell.Value
                xlscell = xlsworksheet.Range("d" + currentRow.ToString())
                xlscell.Value = temp
                xlscell = xlsworksheet.Range("e" + j.ToString())
                temp = xlscell.Value
                xlscell = xlsworksheet.Range("e" + currentRow.ToString())
                xlscell.Value = temp
                xlscell = xlsworksheet.Range("h" + j.ToString())
                temp = xlscell.Value
                xlscell = xlsworksheet.Range("h" + currentRow.ToString())
                xlscell.Value = temp
                xlscell = xlsworksheet.Range("i" + j.ToString())
                temp = xlscell.Value
                xlscell = xlsworksheet.Range("i" + currentRow.ToString())
                xlscell.Value = temp
                xlscell = xlsworksheet.Range("j" + j.ToString())
                temp = xlscell.Value
                xlscell = xlsworksheet.Range("j" + currentRow.ToString())
                xlscell.Value = temp
                xlscell = xlsworksheet.Range("k" + j.ToString())
                temp = xlscell.Value
                xlscell = xlsworksheet.Range("k" + currentRow.ToString())
                xlscell.Value = temp
                currentRow = currentRow + 1
            Next
            ''last save and close
            filesCreated = filesCreated + 1
            BackgroundWorker1.ReportProgress((filesCreated * 100) / NumFiles)

        Next
        DirectoryPathFile = DirectoryPathFolder + prefix
        xlsworkbook.SaveAs(DirectoryPathFile)
        xlsworkbook.Close()

    End Sub

    Sub VinWorkOrder(vins As List(Of String), dates As List(Of String))
        Dim PSN As String
        Dim WOPROD As String
        Dim SOPROD As String
        Dim DirectoryPathFolder As String
        DirectoryPathFolder = TemplateFilePath + "WorkOrders"
        My.Computer.FileSystem.CreateDirectory(DirectoryPathFolder)
        For i As Integer = 0 To vins.Count() - 1
            Dim ch15 As Char = vins.ElementAt(i)(14)
            Dim ch16 As Char = vins.ElementAt(i)(15)
            Dim ch17 As Char = vins.ElementAt(i)(16)
            PSN = ch15 + ch16 + ch17

            WOPROD = "WO-PROD-" + PSN
            SOPROD = "SO-" + PSN
            Dim xlsApp As Excel.Application
            Dim xlsWorkBook As Excel.Workbook
            Dim xlsWorkSheet As Excel.Worksheet
            Dim xlsCell As Excel.Range

            ' Initialise Excel Object
            xlsApp = New Excel.Application

            ' Open test Excel spreadsheet
            xlsWorkBook = xlsApp.Workbooks.Open(TemplateFilePath + "\WORK ORDER TEMPLATES\MAIN BUILD-WORK ORDER.xlsx")
            ' Open worksheet 
            xlsWorkSheet = xlsWorkBook.Worksheets(1)
            '

            'Move Variables and Assign to Cells
            xlsCell = xlsWorkSheet.Range("A2")
            xlsCell.Value = WOPROD
            xlsCell = xlsWorkSheet.Range("B2")
            xlsCell.Value = vins.ElementAt(i)
            xlsCell = xlsWorkSheet.Range("F2")
            xlsCell.Value = dates.ElementAt(i)
            xlsCell = xlsWorkSheet.Range("H2")
            xlsCell.Value = SOPROD
            xlsCell = xlsWorkSheet.Range("J2")
            xlsCell.Value = PSN

            'LAST ROW 

            xlsWorkSheet = xlsWorkBook.Worksheets(2)
            Dim lastRow As Integer = xlsWorkSheet.UsedRange.Rows.Count
            'This is next emptyRow in the sheet
            Dim emptyRow As Integer = lastRow + 1

            Dim lastRow1 As String
            lastRow1 = lastRow
            Dim lastcell As String
            lastcell = "A" + lastRow1
            Dim newrange1 As String
            newrange1 = "A2:" + lastcell
            'MessageBox.Show(newrange1)
            xlsCell = xlsWorkSheet.Range(newrange1)
            xlsCell.Value = vins.ElementAt(i)

            'CASES AND JOIN DB
            Dim DirectoryPathFile As String
            DirectoryPathFolder = TemplateFilePath + "WorkOrders\"
            My.Computer.FileSystem.CreateDirectory(DirectoryPathFolder)
            'Last Save and Close
            DirectoryPathFile = TemplateFilePath + "WorkOrders\" + vins.ElementAt(i)
            xlsWorkBook.SaveAs(DirectoryPathFile)
            xlsWorkBook.Close()
            filesCreated = filesCreated + 1
            BackgroundWorker1.ReportProgress((filesCreated * 100) / NumFiles)
        Next
        DirectoryPathFolder = TemplateFilePath + "WorkOrders\Subassemblies\"
        My.Computer.FileSystem.CreateDirectory(DirectoryPathFolder)


    End Sub

    Function VinIsValid(vin As String) As Boolean
        If (vin.Length <> 17) Then
            Return False
        End If

        Return True
    End Function

    Sub Main(path As String)
        Dim xlsApp As Excel.Application
        Dim xlsWorkBook As Excel.Workbook
        Dim xlsWorkSheet As Excel.Worksheet
        Dim vins As New List(Of String)
        Dim dates As New List(Of String)

        Try
            Dim DirectoryPathFolder = TemplateFilePath + "WorkOrders"
            My.Computer.FileSystem.CreateDirectory(DirectoryPathFolder)
            xlsApp = New Excel.Application
            xlsWorkBook = xlsApp.Workbooks.Open(path)
            xlsWorkSheet = xlsWorkBook.Worksheets(1)

            Dim Range = xlsWorkSheet.UsedRange

            For rcnt = 1 To Range.Rows.Count
                Dim obj = CType(Range.Cells(rcnt, 1), Excel.Range)
                Dim obj1 = CType(Range.Cells(rcnt, 2), Excel.Range)
                If (obj.Value IsNot Nothing And obj1.Value IsNot Nothing) Then
                    If (VinIsValid(obj.Value) And IsDate(obj1.Value)) Then
                        vins.Add(obj.Value)
                        dates.Add(CDate(obj1.Value).ToString("yyyy-MM-dd"))
                        NumFiles = NumFiles + 21
                    Else
                        IO.File.AppendAllText(logFile, "vin or date is invalid for vin: " + obj.Value + Environment.NewLine)
                    End If

                End If
            Next

            xlsWorkBook.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
            Application.Exit()
        Finally

        End Try
        If (vins.Count = 0) Then
            MsgBox("Either the input file was empty or there were no valid vin date combinations")
            Application.Exit()
        End If
        Try
            VinWorkOrder(vins, dates)
            SubAssembly("5THWHEEL-", vins, dates, "5THWHEEL - WORK ORDER.xlsx")
            SubAssembly("ACC-", vins, dates, "AC COMPRESSOR - WORK ORDER.xlsx")
            SubAssembly("AIRT-", vins, dates, "AIR TANK - WORK ORDER.xlsx")
            SubAssembly("PREX-", vins, dates, "BATTERY PACK - WORK ORDER.xlsx")
            SubAssembly("BUMP-", vins, dates, "BUMPER - WORK ORDER.xlsx")
            SubAssembly("CABP-", vins, dates, "CAB PREP - WORK ORDER.xlsx")
            'SubAssembly("COMP-", vins, dates, "COMPRESSOR-HEATER - WORK ORDER.xlsx")
            SubAssembly("DCDC-", vins, dates, "DCDC - WORK ORDER.xlsx")
            SubAssembly("EAXLE-", vins, dates, "E-AXLE - WORK ORDER.xlsx")
            SubAssembly("EXP-", vins, dates, "EXPANSION TANK - WORK ORDER.xlsx")
            SubAssembly("FBEAM-", vins, dates, "FRONT BEAM SUSPENSION - WORK ORDER.xlsx")
            SubAssembly("FAXLE-", vins, dates, "FRONT-AXLE - WORK ORDER.xlsx")
            SubAssembly("FUP-", vins, dates, "FUP - WORK ORDER.xlsx")
            SubAssembly("HVH-", vins, dates, "HV HEATER - WORK ORDER.xlsx")
            SubAssembly("INV-", vins, dates, "INVERTER - WORK ORDER.xlsx")
            SubAssembly("LVBAT-", vins, dates, "LV BATTERY BOX - WORK ORDER.xlsx")
            SubAssembly("PNEL-", vins, dates, "PNEUMATIC LINES - WORK ORDER.xlsx")
            SubAssembly("PNEV-", vins, dates, "PNEUMATIC VALVES - WORK ORDER.xlsx")
            SubAssembly("TAIL-", vins, dates, "TAILLIGHT - WORK ORDER.xlsx")
            SubAssembly("TAXLE-", vins, dates, "TAG AXLE - WORK ORDER.xlsx")
            SubAssembly("TCOIL-", vins, dates, "TRAILER COIL - WORK ORDER.xlsx")
        Catch ex As Exception
            MsgBox(ex.ToString)
            Application.Exit()
        Finally

        End Try
    End Sub


    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
    End Sub
End Class
