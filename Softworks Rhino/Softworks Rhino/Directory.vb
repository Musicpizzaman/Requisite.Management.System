Public Class frmDirectory
    Dim selectedDB As String
    Dim Options() As String
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btnSignOut.Click

        Me.Close()
        login.Show()

    End Sub

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        Dim message, title, defaultValue As String
        Dim myValue As Object
        ' Set prompt.
        message = "Enter Project Name"
        ' Set title.
        title = "Create New Project"
        defaultValue = ""   ' Set default value.

        ' Display message, title, and default value.
        myValue = InputBox(message, title, defaultValue)
        
        ' If user has clicked Cancel, set myValue to defaultValue 
        If myValue Is "" Then myValue = defaultValue
        If Not myValue Is "" Then
            createDatabase(myValue)
        End If
        ListBox1.Items.Clear()
        Dim counter = 0
        ReDim Options(My.Computer.FileSystem.GetFiles(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "/projects/",
    Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.mdb").Count)
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(
    My.Computer.FileSystem.SpecialDirectories.MyDocuments & "/projects/",
    Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.mdb")
            ListBox1.Items.Add(foundFile)
            Options(counter) = (foundFile)
            counter = counter + 1
        Next
    End Sub

    Private Sub Directory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim counter = 0
        ReDim Options(My.Computer.FileSystem.GetFiles(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "/projects/",
    Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.mdb").Count)
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(
    My.Computer.FileSystem.SpecialDirectories.MyDocuments & "/projects/",
    Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.mdb")
            ListBox1.Items.Add(foundFile)
            Options(counter) = (foundFile)
            counter = counter + 1
        Next
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Selection(selectedDB)
        Me.Hide()
        ReqView.Show()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        selectedDB = Options(ListBox1.SelectedIndex)
    End Sub

End Class