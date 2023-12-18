Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop

Module Module1
    Public FOLDERPATH As String = Path.Combine(System.IO.Directory.GetCurrentDirectory, "temp")
    Public TIME As String = DateTime.Now.ToString("MMddyyyy_HHmmss")
    Public LOGPATH As String = Path.Combine(FOLDERPATH, "log" & TIME & ".txt")
    Public TEMPMP As String = Path.Combine(FOLDERPATH, "temp")                      ' The main CSV is processed and stored as a CSV temporarily

    Public SOURCEDS As New Data.DataSet
    Public ERRORDS As New Data.DataSet
    Public ISOLATEDDS As New Data.DataSet
    Public FILTEREDDS As New Data.DataSet

    Public Sub viewTable(ByRef grid As DataGridView, ByRef sourceTable As System.Data.DataTable)
        ' Shows a datatable on a datagridview
        ' @param grid Datagridview to show the data on
        ' @param sourceTable Table to show 

        Dim temp As Long = sourceTable.Rows.Count
        Dim j As Long = 100

        If temp = 0 Then
            Exit Sub
        ElseIf temp < j Then
            j = temp
        End If

        Dim tempTable As New Data.DataTable
        tempTable.Columns.Add("ID", GetType(Long))
        For i = 1 To CInt(Form1.TextBox2.Text)
            tempTable.Columns.Add("C" & i, GetType(Byte))
        Next

        For i = 0 To j - 1
            tempTable.Rows.Add(sourceTable.Rows(i).ItemArray)
        Next

        grid.DataSource = tempTable
    End Sub
    Public Sub ExportCSV(ByRef datatable As Data.DataTable, ByVal path As String)
        ' Exports a datatable into a CSV File
        Dim csvText As String = String.Join(Environment.NewLine, datatable.Rows.Cast(Of DataRow)().[Select](Function(x) String.Join(",", x.ItemArray)))
        File.WriteAllText(path, csvText)
    End Sub

    Public Function FileBrowseAndShow(ByVal listBox As ListBox) As String
        ' Browses for a file and adds it to a listbox
        ' @param listBox {ListBox} Listbox to add to
        ' @returns {String} Filepath

        Dim openFileDialog As New OpenFileDialog() With {
            .Filter = "Excel Files (.xls, .xlsx, .xlsb)|*.xls*",
            .FilterIndex = 1,
            .Multiselect = False
            }

        FileBrowseAndShow = ""

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Get the selected file path
            Dim filePath As String = openFileDialog.FileName

            ' Do something with the file...
            If listBox.Items.Count > 0 Then listBox.Items.Clear()
            listBox.Items.Add(filePath)

            FileBrowseAndShow = filePath
        End If

    End Function

    Public Function Initiate(ByVal path As String) As Boolean
        ' Initiates. Calculates total rows.
        ' @return success = 0, there was something wrong with the processing
        Dim csvData As String = File.ReadAllText(path & ".csv")

        Dim mainTable As New Data.DataTable
        Dim errorTable As New Data.DataTable

        Using csvParser As New Microsoft.VisualBasic.FileIO.TextFieldParser(New StringReader(csvData))
            csvParser.Delimiters = New String() {","}
            csvParser.HasFieldsEnclosedInQuotes = False

            ' Initiating tables
            mainTable.Columns.Add("ID", GetType(Double))
            mainTable.PrimaryKey = New DataColumn() {mainTable.Columns("ID")}

            errorTable.Columns.Add("ID", GetType(Double))


            For i = 1 To CInt(Form1.TextBox2.Text)
                mainTable.Columns.Add("C" & i, GetType(Byte))
            Next

            Dim id As Long = 1

            ' Treating the first row differently, to account for any errors in column count
            Dim rowFields As String() = csvParser.ReadFields()
            If rowFields.Length <> CByte(Form1.TextBox2.Text) Then
                MsgBox("Record length in the main CSV is " & rowFields.Length)
                Initiate = False
                Exit Function
            Else
                Dim newRow As DataRow = mainTable.NewRow()
                newRow.ItemArray = {CStr(id)}.Concat(rowFields).ToArray
                mainTable.Rows.Add(newRow)
                id += 1
            End If

            ' Looping through all other rows
            While Not csvParser.EndOfData
                Try
                    Dim newRow As DataRow = mainTable.NewRow()
                    rowFields = csvParser.ReadFields()
                    If rowFields.Length <> CByte(Form1.TextBox2.Text) Then Throw New Exception
                    newRow.ItemArray = {CStr(id)}.Concat(rowFields).ToArray
                    mainTable.Rows.Add(newRow)
                Catch
                    errorTable.Rows.Add(id)
                End Try
                id += 1
            End While

            mainTable.AcceptChanges()
            errorTable.AcceptChanges()

        End Using

        SOURCEDS.Tables.Add(mainTable)
        ERRORDS.Tables.Add(errorTable)

        'Dim item As New fileItem With {
        '    .total = mainTable.Rows.Count,
        '    .text = "Main File (" & .total & ")",
        '    .history = "Loaded " & Form1.ListBox1.Items(0)
        '    }

        Initiate = True

    End Function

    Public Function generateBinFile(ByRef item As FileItem, ByRef sourceTable As Data.DataTable, ByRef destDS As Data.DataSet) As Boolean
        ' Filters a datatable based on the commaseparated integers, or single integer. Here, it is stored in the custom item.searched
        ' In place updates the item.total property
        ' @param sourceTable Source table to filter
        ' @param item Custom object
        ' @param destDS Destination dataset, where the table is added to
        ' @return 1 -> success 0 -> Faile (filtered row == 0)

        Dim matches As String() = item.searched.Split(",")
        Dim arrayToWrite() As String
        ReDim arrayToWrite(0 To 999)
        Dim ind_ As Integer
        ind_ = 0


        Dim filteredTable As New Data.DataTable()
        Dim filteredRows() As DataRow

        filteredTable.Columns.Add("ID", GetType(Double))
        filteredTable.PrimaryKey = New DataColumn() {filteredTable.Columns("ID")}

        For i = 1 To CInt(Form1.TextBox2.Text)
            filteredTable.Columns.Add("C" & i, GetType(Byte))
        Next

        ' Comma separated search value
        For Each match As String In matches ' Each value
            filteredRows = sourceTable.AsEnumerable().Where(Function(row) row.ItemArray.Skip(1).Contains(CByte(match.Trim))).ToArray()

            If filteredRows.Length = 0 Then
                MsgBox("No row found for " & match.Trim)
                Continue For
            End If

            Dim tempTable As Data.DataTable = filteredRows.CopyToDataTable()
            tempTable.PrimaryKey = New DataColumn() {tempTable.Columns("ID")}

            filteredTable.Merge(tempTable)
        Next

        Dim totalRows As Double = filteredTable.Rows.Count

        If totalRows = 0 Then
            generateBinFile = False
            Exit Function
        End If

        ' Processing of the filtered table
        Dim tableCount As Integer = destDS.Tables.Count
        filteredTable.TableName = "Table" & CStr(tableCount + 1)

        '' Appending previous tables data
        If destDS.Tables.Count > 0 Then filteredTable.Load(destDS.Tables(tableCount - 1).CreateDataReader)

        Dim dv As DataView = filteredTable.DefaultView
        dv.Sort = "ID"
        filteredTable = dv.ToTable
        filteredTable.AcceptChanges()
        destDS.Tables.Add(filteredTable)

        item.Total = totalRows
        generateBinFile = True

    End Function

    Public Sub WriteLog()
        ' Writes a log file of the changes
        Dim final As String = ""
        For Each item In Form1.ListBox2.Items
            final = final & vbNewLine & item.history
        Next

        Using writer As New System.IO.StreamWriter(LOGPATH, True)
            writer.WriteLine(final)
        End Using


    End Sub

    Public Function saveDirectory(ByVal item As FileItem) As Boolean
        ' Browses for a file and moves the temporary file to that location selected by the user
        ' @param item item.Path is the file path to move
        '

        Dim saveFileDialog As New SaveFileDialog() With {
            .FileName = "output",
            .DefaultExt = ".csv",
            .Filter = "CSV Files (*.csv)|*.csv",
            .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            }


        ' Show the dialog box and get the result
        Dim result As DialogResult = saveFileDialog.ShowDialog()

        If result = DialogResult.OK Then
            ' Get the selected file path
            Dim filePath As String = saveFileDialog.FileName

            ' Get the selected folder path
            Dim folderPath As String = Path.GetDirectoryName(filePath)

            ' Get the selected file name without extension
            Dim fileName As String = Path.GetFileNameWithoutExtension(filePath)

            ' New path for the log file
            Dim logPath As String = Path.Combine(folderPath, fileName & ".txt")

            ' Save the file to the selected location with the specified name
            Try
                File.Copy(item.Path, filePath, True)
                File.Copy(item.LogPath, logPath, True)
            Catch ex As Exception
            End Try

        End If

        Return True
    End Function

    Public Sub ExportLog(ByRef item As FileItem)
        Using writer As New StreamWriter(item.LogPath)
            writer.WriteLine("-- Log --")
            writer.WriteLine(item.History)
        End Using

    End Sub

    Public Function splitMainFile(ByVal mainFilePath As String) As Boolean
        ' Splits the main excel into 2 separate CSVs
        ' For pattern finder Design A
        ' @param mainFilePath Path of the main file

        'Try
        ' Starting Excel
        Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
            Dim workbook As Microsoft.Office.Interop.Excel.Workbook
            Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet
            excelApp.DisplayAlerts = False

            ' Loading file
            workbook = excelApp.Workbooks.Open(mainFilePath)

            ' Loading Main WS
            worksheet = workbook.Sheets(1)

            '' Get the range of used cells in the worksheet
            Dim usedRange As Excel.Range = worksheet.UsedRange
            Dim rowCount As Integer = usedRange.Rows.Count
            Dim columnCount As Integer = usedRange.Columns.Count

            '' Processing
            If Regex.IsMatch(worksheet.Cells(1, 1).value, "[a-zA-Z]") Then
                worksheet.Rows(1).delete
            End If

            Dim colCount As Integer = worksheet.UsedRange.Columns.Count

            If Regex.IsMatch(worksheet.Cells(1, 1).value, "[,]") AndAlso colCount > 1 Then
                worksheet.Columns(1).delete
            ElseIf colCount > 1 Then
            Else
                MsgBox("The expected file is supposed to contain each number in a different column in the Excel!")
                Exit Function
            End If

            Form1.ProgressBar1.Value = 15

            '' Exporting
            Try
                File.Delete(TEMPMP & ".csv")
            Catch ex As Exception
            Finally
                worksheet.SaveAs(TEMPMP, Excel.XlFileFormat.xlCSV)
            End Try

            ' Wrapping up
            workbook.Close()
            excelApp.Quit()
            workbook = Nothing
            excelApp = Nothing

            Return True
        'Catch
        '    MsgBox("Try closing the Excel file if it is open")
        '    Return False
        'End Try

    End Function


End Module
