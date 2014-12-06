Imports Microsoft.SqlServer.Server
Imports Microsoft.SqlServer
Public Class login
    Dim User() As User
    Private Sub SetDefault(ByVal Button1 As Button)
        Me.AcceptButton = Button1
    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        TextBox1.SelectAll()
    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        TextBox1.SelectAll()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\..\LogIn.mdb")
        Dim dt As New DataTable
        Dim adapter As OleDb.OleDbDataAdapter
        Dim cmd1 As New OleDb.OleDbCommand("SELECT COUNT(*) FROM UserInfo", conn)
        Dim totalRows As Integer
        Dim currentRow As Integer
        Try
            If (conn.State = ConnectionState.Closed) Then
                conn.Open()
            End If
            adapter = New OleDb.OleDbDataAdapter("SELECT * FROM UserInfo", conn)
            adapter.Fill(dt)
            totalRows = CInt(cmd1.ExecuteScalar()) - 1
            currentRow = 0
            ReDim User(totalRows)
            While currentRow <= totalRows
                User(currentRow) = New User(dt.Rows(currentRow).Item(1), dt.Rows(currentRow).Item(2), dt.Rows(currentRow).Item(3))
                currentRow = currentRow + 1
            End While
        Catch ex As Exception
            MsgBox("Error with IO")
        Finally
            conn.Close()
        End Try

        For i As Integer = 0 To totalRows
            If TextBox1.Text = User(i).getname And TextBox2.Text = User(i).getpw Then
                Usersend(User(i).getname)
                frmDirectory.Show()
                Me.Hide()
                Exit Sub
            End If
        Next
        MsgBox("Incorrect information", MsgBoxStyle.OkOnly)

    End Sub

    Private Sub login_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

End Class
