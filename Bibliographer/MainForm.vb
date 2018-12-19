Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class MainForm

    ' database connectivity used throughout MainForm
    Dim dbConnection As OleDbConnection
    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Documents.accdb;Persist Security Info=True"

    ' local variables to store information about documents on screen
    Dim allDocIDs As New List(Of String) ' list of all DocIDs in the database
    Dim docIDIndex As Integer = 0 ' index into allDocIDs of the currently-viewed DocID (or -1 if a new one); starting at 0 or first DocID
    Dim displayedPersonIDs As New List(Of String) ' list of all PersonIDs for the currently viewed DocID

    ' After the form loads, create the database connection that will be used through the program
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            dbConnection = New OleDbConnection(connectionString)
        Catch ex As Exception
            MessageBox.Show("Can't open the Documents database: " & ex.Message)
        Finally
            dbConnection.Close()
        End Try
    End Sub

    ' Update button in the topright of the SearchPanel switches to the UpdatePanel
    Private Sub UpdateButton_Click(sender As Object, e As EventArgs) Handles UpdateButton.Click
        SearchPanel.Visible = False
        UpdatePanel.Visible = True

        ' When Update button is hit and switching to UpdatePanel, update displayed record
        UpdateAllDocIDs()
        UpdateDisplayedPersonIDs()
        UpdateDisplayedRecord()
    End Sub

    ' Done button in the bottomright of the UpdatePanel returns to the SearchPanel
    Private Sub DoneButton_Click(sender As Object, e As EventArgs) Handles DoneButton.Click
        UpdatePanel.Visible = False
        SearchPanel.Visible = True
    End Sub

    ' returns the currently viewed DocID or "" if it's a new one (we won't get a DocID for a new document until it's commited to the database)
    Private Function GetCurrentDocID() As String
        If docIDIndex >= 0 And docIDIndex < allDocIDs.Count Then ' if index is in range, return DocID at that index
            Return allDocIDs(docIDIndex)
        Else ' otherwise it's likely a new index
            Return ""
        End If
    End Function

    ' This repopulates allDocIDs with all DocIDs from the database
    Private Sub UpdateAllDocIDs()
        allDocIDs.Clear() ' wipe out existing ones first
        Try ' attempt to query the Document table for all DocIDs
            dbConnection.Open()
            Dim cmd As OleDbCommand = New OleDbCommand("SELECT DocID FROM Document", dbConnection)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            While reader.Read()
                allDocIDs.Add(reader("DocID").ToString())
            End While
        Catch ex As Exception
            MessageBox.Show("Error reading DocIDs: " & ex.Message)
        Finally
            dbConnection.Close()
        End Try
    End Sub

    ' This repopulates displayedPersonIDs, a list of PersonIDs associated with the currently viewed DocID
    Private Sub UpdateDisplayedPersonIDs()
        displayedPersonIDs.Clear()
        ' if index is -1 it's a new document and no author known yet
        If docIDIndex >= 0 Then
            Try ' attempt to query the Author junction table for all PersonIDs associated with the current DocID
                dbConnection.Open()
                Dim cmd As OleDbCommand = New OleDbCommand("SELECT PersonID FROM Author WHERE DocID=" & GetCurrentDocID(), dbConnection)
                Dim reader As OleDbDataReader = cmd.ExecuteReader()
                While reader.Read()
                    displayedPersonIDs.Add(reader("PersonID").ToString())
                End While
            Catch ex As Exception
                MessageBox.Show("Error reading PersonIDs: " & ex.Message)
            Finally
                dbConnection.Close()
            End Try
        End If
    End Sub

    ' This updates the controls on the UpdatePanel to refect the currently-viewed record
    Private Sub UpdateDisplayedRecord()

        ' if index into allDocIDs is out of bounds, wrap around to the other end of the list
        If docIDIndex < 0 Then
            docIDIndex = allDocIDs.Count - 1
        ElseIf docIDIndex > allDocIDs.Count - 1 Then
            docIDIndex = 0
        End If

        ' update the "Document" portion (everything but the authors) by pulling its record from the database
        Try
            dbConnection.Open()
            Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM Document WHERE DocID=" & allDocIDs(docIDIndex), dbConnection)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            While reader.Read()
                ' each of the quoted strings ("DocType") is the name of a field in the Document table
                DocTypeComboBox.SelectedItem = reader("DocType").ToString()
                DocTitleTextBox.Text = reader("DocTitle").ToString()
                SectionTitleTextBox.Text = reader("SectionTitle").ToString()
                CityTextBox.Text = reader("City").ToString()
                StateTextBox.Text = reader("State").ToString()
                PublisherTextBox.Text = reader("Publisher").ToString()
                DocYearTextBox.Text = reader("DocYear").ToString()
                DocMonthTextBox.Text = reader("DocMonth").ToString()
                DocDayTextBox.Text = reader("DocDay").ToString()
                StartPageTextBox.Text = reader("StartPage").ToString()
                EndPageTextBox.Text = reader("EndPage").ToString()
                URLTextBox.Text = reader("URL").ToString()
                VolumeTextBox.Text = reader("VolumeNum").ToString()
                IssueTextBox.Text = reader("IssueNum").ToString()
                ' update progress text ("1/5") between nav buttons
                ProgressLabel.Text = String.Format("{0} of {1}", docIDIndex + 1, allDocIDs.Count)
            End While
        Catch ex As Exception
            MessageBox.Show("Error reading Document record (" & docIDIndex & "): " & ex.Message)
        Finally
            dbConnection.Close()
        End Try

    End Sub

End Class
