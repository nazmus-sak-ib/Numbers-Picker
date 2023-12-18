Public Class FileItem
    ' File item to display on a listbox/listview
    Private Property _total As Long        ' Total item in the associated gridview/datatable
    Private Property _searched As String        ' The value that was searched for (comma separated integers, or single integer)
    Public Property Path As String      ' CSV Path
    Public Property LogPath As String   ' Log file path
    Public Property History As String    ' Full history, initiated with a value, and can be updated (appended) anytime
    Public Property text As String    ' Display text on listbox
    Public Property time As Date

    Public Sub New()
        time = DateTime.Now
    End Sub

    Public Property Total As Long     ' Total lines
        Get
            Return _total
        End Get
        Set(value As Long)
            _total = value
            text = text & " (Size: " & _total & ")"
            History = History & vbNewLine & time.ToString("MM/dd/yyyy_HH:mm:ss") & " - " & text
        End Set
    End Property
    Public Sub setPath()         ' Sets the output file path
        Path = IO.Path.Combine(FOLDERPATH, "temp_" & time.ToString("MMddyyyy_HHmmss") & ".csv")
        LogPath = IO.Path.Combine(FOLDERPATH, "temp_" & time.ToString("MMddyyyy_HHmmss") & ".txt")
    End Sub

    Public Property searched As String    ' The value that was searched for (comma separated integers, or single integer)
        Get
            Return _searched
        End Get
        Set(value As String)
            _searched = value
            text = _searched
        End Set
    End Property


End Class