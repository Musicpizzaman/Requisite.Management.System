Imports System
Imports System.IO
Module Depository
    Dim carryFilename As String
    Dim CarryUsername As String
    Dim carryUnique As Boolean
    Dim modeEdit As Boolean
    Dim editselection As String
    Public Sub Usersend(user As String)
        CarryUsername = user
    End Sub
    Public Sub createDatabase(filename As String)
        If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\projects\" & filename & ".mdb") Then
            MsgBox("File name already exists. Please enter a new filename.")
            Exit Sub
        End If
        If Not My.Computer.FileSystem.DirectoryExists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\projects\") Then
            My.Computer.FileSystem.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\projects\")
        End If
        Dim accessApp As New ADOX.Catalog
        Try
            Dim sCreateString As String
            sCreateString = _
              "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\projects\" & filename & ".mdb"
            accessApp.Create(sCreateString)
        Catch Excep As System.Runtime.InteropServices.COMException
        Finally
            accessApp = Nothing
        End Try
        Dim strTestCase As String
        Dim strDeploy As String
        Dim strSpec As String
        Dim strDesign As String
        Dim strTCSpec As String
        Dim strTCDel As String
        Dim strSpecDE As String
        Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\projects\" & filename & ".mdb")
        strTestCase = "CREATE TABLE TestCase (TestCaseID int PRIMARY KEY  NOT NULL, Name varchar(255), Description varchar(255), Link varchar(255), DateLastChanged datetime, AuthorOfLastChange varchar(255))"
        strDesign = "CREATE TABLE Design (DesignID int PRIMARY KEY  NOT NULL, Name varchar(255), Description varchar(255), Link varchar(255), DateLastChanged datetime, AuthorOfLastChange varchar(255))"
        strDeploy = "CREATE TABLE Deploy (DeployID int PRIMARY KEY  NOT NULL, Name varchar(255), Description varchar(255), Link varchar(255), DateLastChanged datetime, AuthorOfLastChange varchar(255))"
        strSpec = "CREATE TABLE Spec (SpecID int PRIMARY KEY  NOT NULL, Name varchar(255), Description varchar(255), Type varchar(255), Effort varchar(255), QFD Varchar(10), Path Varchar(255), ApprovalDate varchar(255), ApprovingUser varchar(255), DateLastChanged datetime, AuthorOfLastChange varchar(255))"
        strTCSpec = "CREATE TABLE TCSpec (TestCaseID int, SpecId int, CONSTRAINT pk_tcspec PRIMARY KEY (TestCaseID, SpecID), CONSTRAINT fk_testcase1 FOREIGN KEY (TestCaseID) REFERENCES TestCase (TestCaseID) ON UPDATE CASCADE ON DELETE CASCADE, CONSTRAINT fk_Spec1 FOREIGN KEY (SpecID) REFERENCES Spec (SpecID) ON UPDATE CASCADE ON DELETE CASCADE)"
        strTCDel = "CREATE TABLE TCDel (TestCaseID int, DeployID int, CONSTRAINT pk_tcdel PRIMARY KEY (TestCaseID, DeployID), CONSTRAINT fk_testcase2 FOREIGN KEY (TestCaseID) REFERENCES TestCase (TestCaseID) ON UPDATE CASCADE ON DELETE CASCADE, CONSTRAINT fk_Deploy FOREIGN KEY (DeployID) REFERENCES Deploy (DeployID) ON UPDATE CASCADE ON DELETE CASCADE)"
        strSpecDE = "CREATE TABLE SpecDE(SpecID int, DesignID int, CONSTRAINT pk_specde PRIMARY KEY (DesignID, SpecID), CONSTRAINT fk_design FOREIGN KEY (DesignID) REFERENCES Design (DesignID) ON UPDATE CASCADE ON DELETE CASCADE, CONSTRAINT fk_Spec2 FOREIGN KEY (SpecID) REFERENCES Spec (SpecID) ON UPDATE CASCADE ON DELETE CASCADE)"
        Dim testCaseCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand(strTestCase, myConn)
        Dim designCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand(strDesign, myConn)
        Dim deployCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand(strDeploy, myConn)
        Dim specCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand(strSpec, myConn)
        Dim tCSpecCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand(strTCSpec, myConn)
        Dim tCDelCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand(strTCDel, myConn)
        Dim specDECommand As OleDb.OleDbCommand = New OleDb.OleDbCommand(strSpecDE, myConn)
        Try
            If (myConn.State = ConnectionState.Closed) Then
                myConn.Open()
            End If
            testCaseCommand.ExecuteNonQuery()
            designCommand.ExecuteNonQuery()
            deployCommand.ExecuteNonQuery()
            specCommand.ExecuteNonQuery()
            tCSpecCommand.ExecuteNonQuery()
            tCDelCommand.ExecuteNonQuery()
            specDECommand.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        Finally
            If (myConn.State = ConnectionState.Open) Then
                myConn.Close()
            End If
        End Try
    End Sub
    Public Sub Selection(database As String)
        carryFilename = database
    End Sub
    Public Function Loader(ByRef Loaded As String(,)) As String(,)
        Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & carryFilename)
        Dim dtSpec As New DataTable
        Dim dtDesign As New DataTable
        Dim dtDeploy As New DataTable
        Dim dtTestCase As New DataTable
        Dim dtTCSPEC As New DataTable
        Dim dtTCDEL As New DataTable
        Dim dtSPECDE As New DataTable
        Dim adapter As OleDb.OleDbDataAdapter
        Dim cmdSpec As New OleDb.OleDbCommand("SELECT COUNT(*) FROM Spec", myConn)
        Dim cmdDesign As New OleDb.OleDbCommand("SELECT COUNT(*) FROM Design", myConn)
        Dim cmdDeploy As New OleDb.OleDbCommand("SELECT COUNT(*) FROM Deploy", myConn)
        Dim cmdTestCase As New OleDb.OleDbCommand("SELECT COUNT(*) FROM TestCase", myConn)
        Dim cmdTCSpec As New OleDb.OleDbCommand("SELECT COUNT(*) FROM TCSpec", myConn)
        Dim cmdTCDel As New OleDb.OleDbCommand("SELECT COUNT(*) FROM TCDel", myConn)
        Dim cmdSpecDE As New OleDb.OleDbCommand("SELECT COUNT(*) FROM SpecDE", myConn)
        Dim totalRowsSpec As Integer
        Dim currentRowSpec As Integer
        Dim totalRowsDesign As Integer
        Dim currentRowDesign As Integer
        Dim totalRowsDeploy As Integer
        Dim currentRowDeploy As Integer
        Dim totalRowsTestCase As Integer
        Dim currentRowTestCase As Integer
        Dim totalRowsTCSPEC As Integer
        Dim currentRowTCSPEC As Integer
        Dim totalRowsTCDEL As Integer
        Dim currentRowTCDEL As Integer
        Dim totalRowsSPECDE As Integer
        Dim currentRowSPECDE As Integer
        Try
            If (myConn.State = ConnectionState.Closed) Then
                myConn.Open()
            End If
            totalRowsSpec = CInt(cmdSpec.ExecuteScalar())
            currentRowSpec = 0
            totalRowsDesign = CInt(cmdDesign.ExecuteScalar())
            currentRowDesign = 0
            totalRowsDeploy = CInt(cmdDeploy.ExecuteScalar())
            currentRowDeploy = 0
            totalRowsTestCase = CInt(cmdTestCase.ExecuteScalar())
            currentRowTestCase = 0
            totalRowsTCSPEC = CInt(cmdTCSpec.ExecuteScalar())
            currentRowTCSPEC = 0
            totalRowsTCDEL = CInt(cmdTCDel.ExecuteScalar())
            currentRowTCDEL = 0
            totalRowsSPECDE = CInt(cmdSpecDE.ExecuteScalar())
            currentRowSPECDE = 0
            ReDim Loaded(7, totalRowsDeploy + totalRowsDesign + totalRowsSpec + totalRowsSPECDE + totalRowsTCDEL + totalRowsTCSPEC + totalRowsTestCase)
            adapter = New OleDb.OleDbDataAdapter("SELECT * FROM Spec", myConn)
            adapter.Fill(dtSpec)
            adapter = New OleDb.OleDbDataAdapter("SELECT * FROM Design", myConn)
            adapter.Fill(dtDesign)
            adapter = New OleDb.OleDbDataAdapter("SELECT * FROM Deploy", myConn)
            adapter.Fill(dtDeploy)
            adapter = New OleDb.OleDbDataAdapter("SELECT * FROM TestCase", myConn)
            adapter.Fill(dtTestCase)
            adapter = New OleDb.OleDbDataAdapter("SELECT * FROM TCSpec", myConn)
            adapter.Fill(dtTCSPEC)
            adapter = New OleDb.OleDbDataAdapter("SELECT * FROM TCDel", myConn)
            adapter.Fill(dtTCDEL)
            adapter = New OleDb.OleDbDataAdapter("SELECT * FROM SpecDE", myConn)
            adapter.Fill(dtSPECDE)
            While currentRowSpec < totalRowsSpec
                Loaded(0, currentRowSpec) = dtSpec.Rows(currentRowSpec).Item(1)
                currentRowSpec = currentRowSpec + 1
            End While
            While currentRowDesign < totalRowsDesign
                Loaded(1, currentRowDesign) = dtDesign.Rows(currentRowDesign).Item(1)
                currentRowDesign = currentRowDesign + 1
            End While
            While currentRowDeploy < totalRowsDeploy
                Loaded(2, currentRowDeploy) = dtDeploy.Rows(currentRowDeploy).Item(1)
                currentRowDeploy = currentRowDeploy + 1
            End While
            While currentRowTestCase < totalRowsTestCase
                Loaded(3, currentRowTestCase) = dtTestCase.Rows(currentRowTestCase).Item(1)
                currentRowTestCase = currentRowTestCase + 1
            End While
            While currentRowTCSPEC < totalRowsTCSPEC
                Loaded(4, currentRowTCSPEC) = dtTCSPEC.Rows(currentRowTCSPEC).Item(1)
                currentRowTCSPEC = currentRowTCSPEC + 1
            End While
            While currentRowTCDEL < totalRowsTCDEL
                Loaded(5, currentRowTCDEL) = dtTCDEL.Rows(currentRowTCDEL).Item(1)
                currentRowTCDEL = currentRowTCDEL + 1
            End While
            While currentRowSPECDE < totalRowsSPECDE
                Loaded(6, currentRowSPECDE) = dtSPECDE.Rows(currentRowSPECDE).Item(1)
                currentRowSPECDE = currentRowSPECDE + 1
            End While
        Catch ex As Exception
            MsgBox("Error with Database IO")
        Finally
            myConn.Close()
        End Try
        Return Loaded
    End Function
    Public Sub newreqentry(name As String, path As String, Entity As String, completion As String, QFD As String, description As String)
        If carryUnique = False Then
            Exit Sub
        End If
        Dim Itemcount As Integer
        Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & carryFilename)
        Dim cmdSpec As New OleDb.OleDbCommand("SELECT COUNT(*) FROM Spec", myConn)
        Try
            If (myConn.State = ConnectionState.Closed) Then
                myConn.Open()
            End If
            Itemcount = CInt(cmdSpec.ExecuteScalar())
            Dim str As String = "INSERT INTO Spec (SpecID, Name, Description, Type, Effort, QFD, Path, DateLastChanged, AuthorofLastChange) VALUES ('" & Itemcount & "', '" & name & "', '" & description & "', '" & Entity & "', '" & completion & "', '" & QFD & "', '" & path & "', '" & Now.Month & "/" & Now.Day & "/" & Now.Year & "', '" & CarryUsername & "')"
            Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(str, myConn)
            command.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        Finally
            If (myConn.State = ConnectionState.Open) Then
                myConn.Close()
            End If
        End Try
    End Sub
    Public Sub checkSpecEntry(name As String)
        Dim check As String
        Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & carryFilename)
        Dim cmdSpec As New OleDb.OleDbCommand("SELECT COUNT(*) FROM Spec", myConn)
        Try
            If (myConn.State = ConnectionState.Closed) Then
                myConn.Open()
            End If
            Dim str As String = "SELECT COUNT(*) FROM Spec WHERE Name='" & name & "';"
            Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(str, myConn)
            check = command.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        Finally
            If (myConn.State = ConnectionState.Open) Then
                myConn.Close()
            End If
        End Try
        If Not check = 0 Then
            MsgBox("That name is already used. Please try again.")
            carryUnique = False
            Exit Sub
        Else
            carryUnique = True
        End If
    End Sub
    Public Function Checkunique(ByRef test As String)
        test = carryUnique
        Return test
    End Function
    Public Sub checkDEEntry(name As String)
        Dim check As String
        Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & carryFilename)
        Dim cmdSpec As New OleDb.OleDbCommand("SELECT COUNT(*) FROM Design", myConn)
        Try
            If (myConn.State = ConnectionState.Closed) Then
                myConn.Open()
            End If
            Dim str As String = "SELECT COUNT(*) FROM Design WHERE Name='" & name & "';"
            Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(str, myConn)
            check = command.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        Finally
            If (myConn.State = ConnectionState.Open) Then
                myConn.Close()
            End If
        End Try
        If Not check = 0 Then
            MsgBox("That name is already used. Please try again.")
            carryUnique = False
            Exit Sub
        Else
            carryUnique = True
        End If
    End Sub
    Public Sub newDEentry(name As String, path As String, description As String, reference As String)
        If carryUnique = False Then
            Exit Sub
        End If
        Dim Itemcount As Integer
        Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & carryFilename)
        Dim cmdDesign As New OleDb.OleDbCommand("SELECT COUNT(*) FROM Design", myConn)
        Try
            If (myConn.State = ConnectionState.Closed) Then
                myConn.Open()
            End If
            Itemcount = CInt(cmdDesign.ExecuteScalar())
            Dim str As String = "INSERT INTO Design (DesignID, Name, Description, Link, DateLastChanged, AuthorOfLastChange) VALUES ('" & Itemcount & "', '" & name & "', '" & description & "', '" & path & "', '" & Now.Month & "/" & Now.Day & "/" & Now.Year & "', '" & CarryUsername & "')"
            Dim command1 As OleDb.OleDbCommand = New OleDb.OleDbCommand(str, myConn)
            command1.ExecuteNonQuery()
            str = "INSERT INTO SpecDE (SpecID, DesignID) VALUES ('" & reference & "', '" & Itemcount & "')"
            Dim command2 As OleDb.OleDbCommand = New OleDb.OleDbCommand(str, myConn)
            command2.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        Finally
            If (myConn.State = ConnectionState.Open) Then
                myConn.Close()
            End If
        End Try
    End Sub
    Public Sub checkDeployEntry(name As String)
        Dim check As String
        Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & carryFilename)
        Dim cmdDeploy As New OleDb.OleDbCommand("SELECT COUNT(*) FROM Deploy", myConn)
        Try
            If (myConn.State = ConnectionState.Closed) Then
                myConn.Open()
            End If
            Dim str As String = "SELECT COUNT(*) FROM Deploy WHERE Name='" & name & "';"
            Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(str, myConn)
            check = command.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        Finally
            If (myConn.State = ConnectionState.Open) Then
                myConn.Close()
            End If
        End Try
        If Not check = 0 Then
            MsgBox("That name is already used. Please try again.")
            carryUnique = False
            Exit Sub
        Else
            carryUnique = True
        End If
    End Sub
    Public Sub newDeployentry(name As String, path As String, description As String)
        If carryUnique = False Then
            Exit Sub
        End If
        Dim Itemcount As Integer
        Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & carryFilename)
        Dim cmdDeploy As New OleDb.OleDbCommand("SELECT COUNT(*) FROM Deploy", myConn)
        Try
            If (myConn.State = ConnectionState.Closed) Then
                myConn.Open()
            End If
            Itemcount = CInt(cmdDeploy.ExecuteScalar())
            Dim str As String = "INSERT INTO Deploy ((DeployID, Name, Description, Link, DateLastChanged, AuthorOfLastChange) VALUES ('" & Itemcount & "', '" & name & "', '" & description & "', '" & path & "', '" & Now.Month & "/" & Now.Day & "/" & Now.Year & "', '" & CarryUsername & "')"
            Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(str, myConn)
            command.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        Finally
            If (myConn.State = ConnectionState.Open) Then
                myConn.Close()
            End If
        End Try
    End Sub
    Public Sub checkTCEntry(name As String)
        Dim check As String
        Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & carryFilename)
        Dim cmdTestCase As New OleDb.OleDbCommand("SELECT COUNT(*) FROM TestCase", myConn)
        Try
            If (myConn.State = ConnectionState.Closed) Then
                myConn.Open()
            End If
            Dim str As String = "SELECT COUNT(*) FROM TestCase WHERE Name='" & name & "';"
            Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(str, myConn)
            check = command.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        Finally
            If (myConn.State = ConnectionState.Open) Then
                myConn.Close()
            End If
        End Try
        If Not check = 0 Then
            MsgBox("That name is already used. Please try again.")
            carryUnique = False
            Exit Sub
        Else
            carryUnique = True
        End If
    End Sub
    Public Sub newTCentry(name As String, path As String, description As String, reference As String)
        If carryUnique = False Then
            Exit Sub
        End If
        Dim Itemcount As Integer
        Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & carryFilename)
        Dim cmdTestCase As New OleDb.OleDbCommand("SELECT COUNT(*) FROM TestCase", myConn)

        Try
            If (myConn.State = ConnectionState.Closed) Then
                myConn.Open()
            End If

            Itemcount = CInt(cmdTestCase.ExecuteScalar())
            Dim str As String = "INSERT INTO TestCase (TestCaseID, Name, Description, Link, DateLastChanged, AuthorOfLastChange) VALUES ('" & Itemcount & "', '" & name & "', '" & description & "', '" & path & "', '" & Now.Month & "/" & Now.Day & "/" & Now.Year & "', '" & CarryUsername & "')"
            Dim command1 As OleDb.OleDbCommand = New OleDb.OleDbCommand(str, myConn)
            command1.ExecuteNonQuery()
            str = "INSERT INTO TCspec (TestCaseID, SpecID) VALUES ('" & Itemcount & "', '" & reference & "')"
            Dim command2 As OleDb.OleDbCommand = New OleDb.OleDbCommand(str, myConn)
            command2.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        Finally
            If (myConn.State = ConnectionState.Open) Then
                myConn.Close()
            End If
        End Try
    End Sub
    Public Sub EditMode(Selection As String)
        modeEdit = True
        editselection = Selection
    End Sub
    Public Sub clearEditMode(Selection As String)
        modeEdit = False
        editselection = Selection
    End Sub
    Public Function checkEditMode(ByRef editMode As Boolean)
        editMode = modeEdit
        Return editMode
    End Function
    Public Function loadspec(ByRef loaded As String())
        Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & carryFilename)
        Dim dtSpec As New DataTable
        Dim adapter As OleDb.OleDbDataAdapter
        Try
            If (myConn.State = ConnectionState.Closed) Then
                myConn.Open()
            End If
            ReDim loaded(11)
            adapter = New OleDb.OleDbDataAdapter("SELECT * FROM Spec WHERE SpecID='" & editselection & "';", myConn)
            adapter.Fill(dtSpec)
            loaded(0) = dtSpec.Rows(0).Item(0)
            loaded(1) = dtSpec.Rows(0).Item(1)
            loaded(2) = dtSpec.Rows(0).Item(2)
            loaded(3) = dtSpec.Rows(0).Item(3)
            loaded(4) = dtSpec.Rows(0).Item(4)
            loaded(5) = dtSpec.Rows(0).Item(5)
            loaded(6) = dtSpec.Rows(0).Item(6)
            loaded(7) = dtSpec.Rows(0).Item(7)
            loaded(8) = dtSpec.Rows(0).Item(8)
            loaded(9) = dtSpec.Rows(0).Item(9)
            loaded(10) = dtSpec.Rows(0).Item(10)
        Catch ex As Exception
            MsgBox("Error with Database IO")
        Finally
            myConn.Close()
        End Try
        Return loaded
    End Function

End Module
