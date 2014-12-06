Public Class User
    'James Haley
    Private name As String
    Private password As String
    Private usertype As String

    Public Sub New(n As String, pw As String, type As String)
        name = n
        password = pw
        usertype = Type
    End Sub
    Public ReadOnly Property getname As String
        Get
            Return name
        End Get
    End Property
    Public ReadOnly Property getpw As String
        Get
            Return password
        End Get
    End Property
    Public ReadOnly Property getusertype As String
        Get
            Return usertype
        End Get
    End Property
End Class
