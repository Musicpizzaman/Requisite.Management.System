Public Class Filecarry
    Private SelectedDB As String

    Public Sub New(n As String)
        SelectedDB = n
    End Sub
    Public ReadOnly Property getSelectedDB As String
        Get
            Return SelectedDB
        End Get
    End Property
End Class
