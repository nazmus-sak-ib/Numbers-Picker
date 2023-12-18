Imports System.IO
Imports System.Text.RegularExpressions

Public Class Form1
    Private Sub browseVisuals()
        ' VIsuals after browse button is pressed (everything resets)
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        Button3.Enabled = False
        Button6.Enabled = False
        GroupBox2.Enabled = False
        GroupBox3.Enabled = False
        GroupBox4.Enabled = False
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Browse
        browseVisuals()

        Dim path As String = FileBrowseAndShow(ListBox1)

        If path <> "" Then
            Button2.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
        End If

    End Sub
    Public Function warning1() As Boolean
        warning1 = False
        If Not IsNumeric(TextBox2.Text) Then
            MsgBox("Enter the record length")
        ElseIf Not IsNumeric(TextBox3.Text) Then
            MsgBox("Enter the record length")
        ElseIf ListBox1.Items.Count = 0 Then
            MsgBox("Please select a csv file first")
        Else
            warning1 = True
        End If
    End Function
    Public Sub initiateVisuals()
        ' Visual tweaks after initiation
        TextBox1.Enabled = True
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        Button2.Enabled = False

        Button3.Enabled = True
        GroupBox1.Enabled = True
        GroupBox3.Enabled = True

    End Sub
    Private Sub Initiate_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Initiate
        ProgressBar1.Maximum = 100
        ProgressBar1.Value = 8

        If Not warning1() Then Exit Sub
        ProgressBar1.Value = 14
        Dim success As Boolean                                  ' If processing is successful

        success = splitMainFile(ListBox1.Items(0))
        ProgressBar1.Value = 30

        If Not success Then
            MsgBox("File processing failed!")
            ProgressBar1.Value = 0
            Exit Sub
        End If

        ProgressBar1.Value = 36

        success = Initiate(TEMPMP)

        ProgressBar1.Value = 80

        If success Then
            initiateVisuals()
            ProgressBar1.Value = 85
            viewTable(DataGridView3, SOURCEDS.Tables(0))
            TextBox5.Text = SOURCEDS.Tables(0).Rows.Count

            ProgressBar1.Value = 90

            viewTable(DataGridView4, ERRORDS.Tables(0))
            ProgressBar1.Value = 95
            TextBox7.Text = ERRORDS.Tables(0).Rows.Count
            ProgressBar1.Value = 100
        Else
            ProgressBar1.Value = 0
        End If

    End Sub
    Public Function Warning2() As Boolean
        Warning2 = True

        If TextBox1.Text.ToString = "" Then
            MsgBox("Enter a single integer, or comma separated integers!")
            Warning2 = False
        Else
            For Each num As String In TextBox1.Text.ToString.Split(",")
                If Not IsNumeric(num) Then
                    MsgBox("Enter a single integer, or comma separated integers!")
                    Warning2 = False
                    Exit Function
                ElseIf CInt(num) > CInt(TextBox3.Text) Then
                    MsgBox("Lookup value cannot be greater than " & TextBox1.Text.ToString)
                    Warning2 = False
                    Exit Function
                End If
            Next
        End If
    End Function
    Public Sub isolateVisuals()
        ' Visuals after isolate button is pressed
        If ListBox2.Items.Count > 0 Then
            Button3.Enabled = False
            TextBox1.Text = ""
            TextBox1.Enabled = False

            GroupBox2.Enabled = True

            TextBox4.Enabled = True
            Button6.Enabled = True

        Else
            Button3.Enabled = True
            TextBox1.Enabled = True

            GroupBox2.Enabled = False

            TextBox4.Enabled = False
            Button6.Enabled = False
        End If
    End Sub
    Private Sub Isolate_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Isolate button
        If Not Warning2() Then Exit Sub

        ' Custom object
        Dim item As New FileItem With {
            .searched = TextBox1.Text
            }

        If Not generateBinFile(item, SOURCEDS.Tables(0), ISOLATEDDS) Then
            Exit Sub
        End If

        viewTable(DataGridView2, ISOLATEDDS.Tables.Item(ISOLATEDDS.Tables.Count - 1))

        ListBox2.Items.Add(item)
        isolateVisuals()

    End Sub

    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        If ListBox3.SelectedIndex = -1 Then
            Button8.Enabled = False
        Else
            Button8.Enabled = True
        End If
    End Sub

    Private Sub OpenCSVFiltered_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ' Open Filtered bin CSV
        If ListBox3.SelectedItems.Count = 1 Then
            ListBox3.SelectedItem.setPath
            ExportCSV(FILTEREDDS.Tables(ListBox3.SelectedIndex), ListBox3.SelectedItem.path)

            Try
                Process.Start(ListBox3.SelectedItem.path)
            Catch ex As Exception
                MsgBox("Selected file does not exist!")
            End Try

        End If
    End Sub

    Private Sub StepBackFilteredBin_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ' Step back Filtered bin
        If ListBox3.Items.Count > 0 Then
            Dim tempPath As String = ListBox3.Items(ListBox3.Items.Count - 1).path

            Try
                System.IO.File.Delete(tempPath)
            Catch ex As Exception
            End Try

            ListBox3.Items.RemoveAt(ListBox3.Items.Count - 1)
            FILTEREDDS.Tables.RemoveAt(FILTEREDDS.Tables.Count - 1)

            Dim tempCount As Integer = FILTEREDDS.Tables.Count
            If tempCount > 0 Then
                viewTable(DataGridView1, FILTEREDDS.Tables(tempCount - 1))
            Else
                DataGridView1.DataSource = Nothing
            End If

            searchVisuals()

        End If
    End Sub

    Private Sub Delete_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ' Delete
        Dim deleteTable As Data.DataTable
        Dim copyTable As Data.DataTable

        If FILTEREDDS.Tables.Count > 0 Then
            deleteTable = FILTEREDDS.Tables(FILTEREDDS.Tables.Count - 1)

            Dim response = MsgBox("Proceed to delete " & deleteTable.Rows.Count & " rows?", vbYesNo, "Confirm")
            If response = vbNo Then Exit Sub

            Dim mainTable As Data.DataTable = ISOLATEDDS.Tables(ISOLATEDDS.Tables.Count - 1)

            copyTable = mainTable.Copy
            copyTable.PrimaryKey = New DataColumn() {copyTable.Columns("ID")}

            Dim regex As New Regex("(\d+)")
            Dim match As Match = regex.Match(mainTable.TableName)
            Dim newN As Integer = CInt(match.Groups(1).Value) + 1

            copyTable.TableName = "Table" & CStr(newN)

            For Each row As DataRow In deleteTable.Rows
                copyTable.Rows.Find(row("ID")).Delete()
            Next

            copyTable.AcceptChanges()
            ISOLATEDDS.Tables.Add(copyTable)

            ' Updating Basket bin
            Dim prevItem As FileItem = ListBox2.Items(ListBox2.Items.Count - 1)
            Dim item As New FileItem With {
                .History = prevItem.History
            }

            ' Updating history of new item from all filters
            For Each tempItem As FileItem In ListBox3.Items
                item.History &= vbNewLine _
                                & tempItem.History
            Next

            item.searched = "Executed Deletion"
            item.Total = copyTable.Rows.Count

            item.History &= vbNewLine & "Finally " & copyTable.Rows.Count

            item.setPath()

            Dim exportTable As Data.DataTable = copyTable.Copy
            exportTable.PrimaryKey = Nothing
            exportTable.Columns.Remove("ID")

            ExportCSV(exportTable, item.Path)
            ExportLog(item)
            saveDirectory(item)
            deleteVisuals(item)

            FILTEREDDS = New Data.DataSet
            viewTable(DataGridView1, ISOLATEDDS.Tables(ISOLATEDDS.Tables.Count - 1))

        ElseIf FILTEREDDS.Tables.Count = 0 And ISOLATEDDS.Tables.Count > 0 Then
            Dim item As FileItem = ListBox2.Items(ListBox2.Items.Count - 1)
            item.setPath()
            ExportCSV(ISOLATEDDS.Tables(ISOLATEDDS.Tables.Count - 1), item.Path)
            ExportLog(item)
            saveDirectory(item)

            FILTEREDDS = New Data.DataSet
        End If

    End Sub
    Private Sub deleteVisuals(ByRef item As FileItem)
        ' After delete button is pressed
        GroupBox4.Enabled = True
        ListBox4.Items.Add(item)
        ListBox2.Items.Add(item)
        ListBox3.Items.Clear()
        'GroupBox3.Enabled = False
        TextBox6.Text = ""
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub StepBackIsolated_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Step back Isolated bin
        Dim tempPath As String = ListBox2.Items(0).path

        Try
            File.Delete(tempPath)
        Catch ex As Exception
        End Try

        ListBox2.Items.RemoveAt(ListBox2.Items.Count - 1)
        ISOLATEDDS.Tables.RemoveAt(ISOLATEDDS.Tables.Count - 1)
        DataGridView2.DataSource = Nothing
        GroupBox2.Enabled = False

        ' Also cleaning fitered data/listbox
        FILTEREDDS.Tables.Clear()
        ListBox3.Items.Clear()
        isolateVisuals()

    End Sub

    Private Sub IsolatedCSVOpen_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Open Isolated bin CSV
        If ListBox2.SelectedIndex <> -1 Then
            ListBox2.SelectedItem.setPath()
            ExportCSV(ISOLATEDDS.Tables(ListBox2.SelectedIndex), ListBox2.SelectedItem.Path)

            Try
                Process.Start(ListBox2.SelectedItem.path)
            Catch ex As Exception
                MsgBox("Selected file does not exist!")
            End Try
        End If
    End Sub


    Private Sub OpenSavedFilesBin_Click(sender As Object, e As EventArgs) Handles Button10.Click
        ' Open Saved Files bin
        If ListBox4.SelectedIndex <> -1 Then
            Try
                Process.Start(ListBox4.SelectedItem.path)
            Catch ex As Exception
                MsgBox("Selected file does not exist!")
            End Try
        End If
    End Sub
    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
        If ListBox4.SelectedIndices.Count = 0 Then
            Button10.Enabled = False
        Else
            Button10.Enabled = True
        End If
    End Sub

    Public Function Warning4() As Boolean
        ' Similar as Warning2
        Warning4 = True

        If TextBox4.Text.ToString = "" Then
            MsgBox("Enter a single integer, or comma separated integers!")
            Warning4 = False
        Else
            For Each num As String In TextBox4.Text.ToString.Split(",")
                If Not IsNumeric(num) Then
                    MsgBox("Enter a single integer, or comma separated integers!")
                    Warning4 = False
                    Exit Function
                ElseIf CInt(num) > CInt(TextBox3.Text) Then
                    MsgBox("Lookup value cannot be greater than " & TextBox4.Text.ToString)
                    Warning4 = False
                    Exit Function
                End If
            Next
        End If
    End Function

    Private Sub Search_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ' Search button for filtering
        If Not Warning4() Then Exit Sub

        ' Custom object
        Dim item As New FileItem With {
            .searched = TextBox4.Text
            }

        If Not generateBinFile(item, ISOLATEDDS.Tables(ISOLATEDDS.Tables.Count - 1), FILTEREDDS) Then
            MsgBox("No result found!")
            Exit Sub
        End If


        ListBox3.Items.Add(item)

        searchVisuals()

    End Sub
    Public Sub searchVisuals()
        ' Visuals after search button for filtering is pressed

        Dim tempCount As Integer = FILTEREDDS.Tables.Count
        If tempCount > 0 Then
            viewTable(DataGridView1, FILTEREDDS.Tables.Item(tempCount - 1))
            'GroupBox3.Enabled = True
            TextBox6.Text = FILTEREDDS.Tables.Item(tempCount - 1).Rows.Count
            TextBox4.Text = ""
        Else
            'GroupBox3.Enabled = False
            TextBox6.Text = 0
        End If


    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        ' Exit menu item click
        Dim response = MsgBox("Do you want to exit the program?", vbOKCancel, "Exit")
        If response = vbOK Then Me.Close()
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        If ListBox2.SelectedIndex = -1 Then
            Button5.Enabled = False
        Else
            Button5.Enabled = True
        End If
    End Sub


    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        ' Help menu item click
        FormHelp.Show()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Listbox management
        ListBox2.DisplayMember = "text"
        ListBox3.DisplayMember = "text"
        ListBox4.DisplayMember = "text"

        ' File management
        Try
            System.IO.Directory.CreateDirectory(FOLDERPATH)
        Catch ex As Exception
        End Try

    End Sub
    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        ' Delete all temp files
        For Each deleteFile In Directory.GetFiles(FOLDERPATH, "*.csv", System.IO.SearchOption.TopDirectoryOnly)
            File.Delete(deleteFile)
        Next
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub
End Class
