Public Class ReqView
    'When a req is selected the associated code will appear 
    Private Sub ReqView_shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        btnProjComplete.Enabled = False
        CheckedListBox1.CheckOnClick = False
        Dim load As String(,)
        Loader(load)
        For j As Integer = 0 To load.GetLength(1) - 1
            If Not load(0, j) = Nothing Then
                ListBox1.Items.Add(load(0, j))
            End If
        Next
        For j As Integer = 0 To load.GetLength(1) - 1
            If Not load(1, j) = Nothing Then
                ListBox2.Items.Add(load(1, j))
            End If
        Next
        For j As Integer = 0 To load.GetLength(1) - 1
            If Not load(3, j) = Nothing Then
                ListBox4.Items.Add(load(3, j))
            End If
        Next
    End Sub
    Private Sub ReqView_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ReturnToDirectoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReturnToDirectoryToolStripMenuItem.Click
        Me.Close()
        frmDirectory.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        Newreq.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        NewTestCase.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        NewDesignElem.Show()
    End Sub

    Private Sub ListBox1_Doubleclick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If Not ListBox1.SelectedItem = Nothing Then
            EditMode(ListBox1.SelectedIndex)
            Me.Close()
            Newreq.Show()
        End If

    End Sub
End Class