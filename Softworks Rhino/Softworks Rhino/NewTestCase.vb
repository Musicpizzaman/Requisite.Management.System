﻿Public Class NewTestCase


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = String.Empty Or TextBox1.Text = "Design Element name" Or TextBox2.Text = String.Empty Or TextBox2.Text = "Enter URL here" Or RichTextBox1.Text = "Enter Design Element description" Or RichTextBox1.Text = String.Empty Then
            MsgBox("Fill in all fields in the form", MsgBoxStyle.OkOnly, )
            Exit Sub
        End If
        checkTCEntry(TextBox1.Text)
        Dim unique As Boolean
        Checkunique(unique)
        If unique = True Then
            newTCentry(TextBox1.Text, TextBox2.Text, RichTextBox1.Text, ComboBox1.SelectedIndex)
            Me.Hide()
            ReqView.Show()
        Else
            Exit Sub
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        TextBox1.SelectAll()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        TextBox2.SelectAll()
    End Sub

    Private Sub NewTestCase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim load As String(,)
        Loader(load)
        For j As Integer = 0 To load.GetLength(1) - 1
            If Not load(0, j) = Nothing Then
                ComboBox1.Items.Add(load(0, j))
            End If
        Next
    End Sub
End Class