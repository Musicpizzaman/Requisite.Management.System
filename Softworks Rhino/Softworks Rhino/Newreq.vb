Public Class Newreq


    'only the people who can approve new requirements are the only ones who can add reqs. 

    Private Sub txtName_Click(sender As Object, e As EventArgs) Handles txtName.Click
        txtName.SelectAll()
    End Sub


    Private Sub Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim editmode As Boolean
        checkEditMode(editmode)
        If editmode = False Then
            If txtName.Text = String.Empty Or txtPath.Text = String.Empty Or cboEntity.Items.Count = 0 Or cboTime.Items.Count = 0 Or cboQFD.Items.Count = 0 Then
                MsgBox("Fill in all fields in the form", MsgBoxStyle.OkOnly, )
                Exit Sub
            End If
            checkSpecEntry(txtName.Text)
            Dim unique As Boolean
            Checkunique(unique)
            If unique = True Then
                newreqentry(txtName.Text, txtPath.Text, cboEntity.Text, cboTime.Text, cboQFD.Text, rtbDiscription.Text)
                Me.Hide()
                ReqView.Show()
            Else
                Exit Sub
            End If
        Else
            Dim load As String()
            loadspec(load)
            txtName.Text = load(1)
            txtPath.Text = load(6)
            cboEntity.SelectedText = load(3)
            cboTime.SelectedText = load(4)
            cboQFD.SelectedText = load(5)
            rtbDiscription.Text = load(2)
            clearEditMode(editmode)
        End If


    End Sub
    'Or txtName.Text = "Trait Name" 
End Class