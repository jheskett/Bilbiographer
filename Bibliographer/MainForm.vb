' Bibliographer.vb
' CIS410 Fall 2018 Final Project
' Members: Jeff Heskett, Phillip Foss, Ian Shaw
' Purpose: This project will allow a user to store and retrieve bibliographic references for documents by author or year
'
' This is the main form that primarily has an SearchPanel and its controls to look up existing document refrences,
' and UpdatePanel to add, update and delete document references.

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
        AuthorNames.ShareDBConnection(dbConnection) ' share dbConnection with AuthorNames (ideally this db connection should be its own static module)
    End Sub

    ' Update button in the topright of the SearchPanel switches to the UpdatePanel
    Private Sub UpdateButton_Click(sender As Object, e As EventArgs) Handles UpdateButton.Click
        SearchPanel.Visible = False
        UpdatePanel.Visible = True

        ' When Update button is hit and switching to UpdatePanel, update displayed record
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

        ' before doing anything, make sure we have the most up-to-date DocIDs
        UpdateAllDocIDs()

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

        ' before displaying authors for this document, populate displayedPersonIDs with all PersonIDs associated with this DocID
        UpdateDisplayedPersonIDs()

        ' if there's any PersonID tied to this DocID, fill the textboxes with the first author's name
        UpdateFrontAuthor()

    End Sub

    ' click of First button on UpdatePanel to show first record
    Private Sub FirstNavButton_Click(sender As Object, e As EventArgs) Handles FirstNavButton.Click
        docIDIndex = 0
        UpdateDisplayedRecord()
    End Sub

    ' click of Prev button on UpdatePanel to display previous record
    Private Sub PrevNavButton_Click(sender As Object, e As EventArgs) Handles PrevNavButton.Click
        docIDIndex -= 1
        UpdateDisplayedRecord()
    End Sub

    ' click of Next button on UpdatePanel to display next record
    Private Sub NextNavButton_Click(sender As Object, e As EventArgs) Handles NextNavButton.Click
        docIDIndex += 1
        UpdateDisplayedRecord()
    End Sub

    ' click of Last button on UpdatePanel to display last record
    Private Sub LastNavButton_Click(sender As Object, e As EventArgs) Handles LastNavButton.Click
        docIDIndex = allDocIDs.Count - 1
        UpdateDisplayedRecord()
    End Sub

    ' click of New button will empty the form to wait for a new document to be entered
    Private Sub NewButton_Click(sender As Object, e As EventArgs) Handles NewButton.Click
        ' index Of -1 means it's a new record (going next or prev will wrap around to a valid record)
        docIDIndex = -1
        ' reset all fields
        DocTypeComboBox.SelectedIndex = 0 ' Select Default For DocType
        DocTitleTextBox.Text = ""
        SectionTitleTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        PublisherTextBox.Text = ""
        DocYearTextBox.Text = ""
        DocMonthTextBox.Text = ""
        DocDayTextBox.Text = ""
        StartPageTextBox.Text = ""
        EndPageTextBox.Text = ""
        URLTextBox.Text = ""
        VolumeTextBox.Text = ""
        IssueTextBox.Text = ""
        ' update progress text to show it's beyond last record: "6/5"
        ProgressLabel.Text = String.Format("{0} of {1}", allDocIDs.Count + 1, allDocIDs.Count)
        ' empty author stuff
        displayedPersonIDs.Clear()
        UpdateFrontAuthor()
        ' move focus to first empty field (AuthorFirstName)
        AuthorFirstNameTextBox.Select()
    End Sub

    ' click of Save button on UpdatePanel will either insert the document into the database (if docIDIndex is -1)
    ' or update an existing document (if docIDIndex>=0) with the filled in textbox contents
    Private Sub SaveButton_Click(sender As Object, e As EventArgs) Handles SaveButton.Click

        ' don't allow a document to save if an author, title or year have not been entered
        If AuthorFirstNameTextBox.Text.Trim() = "" Or DocTitleTextBox.Text.Trim() = "" Or DocYearTextBox.Text.Trim() = "" Then
            MessageBox.Show("All documents must have an author, title and year.")
            Return
        End If

        ' first update the Document table (remember no PersonID is in this table; doing soon after)
        If docIDIndex = -1 Then ' new document, needs to be INSERT into Document table
            InsertNewDocument()
        Else ' updating existing document, needs to be UPDATEd in Document table
            UpdateExistingDocument()
        End If

        ' if the name in the UpdatePanel's author textboxes are not in displayedPersonIDs, then add it to the list of personIDs
        ' this can potentially add new names to PersonIDs; which is okay they'll be cleaned up after Save and they won't be associated until saved
        Dim frontPersonID = AuthorNames.GetPersonIDByName(AuthorFirstNameTextBox.Text, AuthorMITextBox.Text, AuthorLastNameTextBox.Text)
        If frontPersonID <> "" And Not displayedPersonIDs.Contains(frontPersonID) Then
            If displayedPersonIDs.Count = 0 Then ' if there is no authors in displayedPersonIDs list
                displayedPersonIDs.Add(frontPersonID) ' then add them
            Else
                displayedPersonIDs(0) = frontPersonID ' otherwise replacing first name that was in that box
            End If
        End If

        ' next update Author table for the author(s) (the Person table already added all new authors when getting PersonIDs for them)
        UpdateAuthorTable()

        ' finally update the displayed record to reflect what's in the database (should be no apparent change to user)
        UpdateDisplayedRecord()
    End Sub

    ' for use when saving potentially empty number fields to a database: return text as an integer or a null
    Private Function NumberOrNull(text As String) As Object
        If IsNumeric(text) Then
            Return CInt(text)
        Else
            Return DBNull.Value
        End If
    End Function


    ' called by SaveButton click; creates a new record in the Document table from textbox contents
    Private Sub InsertNewDocument()
        Try
            dbConnection.Open()
            Dim cmd As OleDbCommand = New OleDbCommand("INSERT INTO Document(DocType, DocTitle, SectionTitle, City, State, Publisher, DocYear, DocMonth, DocDay, StartPage, EndPage, URL, VolumeNum, IssueNum) " +
                            "VALUES (@DocType, @DocTitle, @SectionTitle, @City, @State, @Publisher, @DocYear, @DocMonth, @DocDay, @StartPage, @EndPage, @URL, @VolumeNum, @IssueNum)", dbConnection)
            ' using parameters to prevent SQL injection from user input
            With cmd.Parameters
                .Add(New OleDbParameter("@DocType", DocTypeComboBox.SelectedItem))
                .Add(New OleDbParameter("@DocTitle", DocTitleTextBox.Text))
                .Add(New OleDbParameter("@SectionTitle", SectionTitleTextBox.Text))
                .Add(New OleDbParameter("@City", CityTextBox.Text))
                .Add(New OleDbParameter("@State", StateTextBox.Text))
                .Add(New OleDbParameter("@Publisher", PublisherTextBox.Text))
                .Add(New OleDbParameter("@DocYear", NumberOrNull(DocYearTextBox.Text)))
                .Add(New OleDbParameter("@DocMonth", NumberOrNull(DocMonthTextBox.Text)))
                .Add(New OleDbParameter("@DocDay", NumberOrNull(DocDayTextBox.Text)))
                .Add(New OleDbParameter("@StartPage", NumberOrNull(StartPageTextBox.Text)))
                .Add(New OleDbParameter("@EndPage", NumberOrNull(EndPageTextBox.Text)))
                .Add(New OleDbParameter("@URL", URLTextBox.Text))
                .Add(New OleDbParameter("@VolumeNum", NumberOrNull(VolumeTextBox.Text)))
                .Add(New OleDbParameter("@IssueNum", NumberOrNull(IssueTextBox.Text)))
            End With
            ' execute the INSERT
            cmd.ExecuteNonQuery()
            ' since a new DocID was just created, and we can't be sure what it is, need to update them
            dbConnection.Close()
            UpdateAllDocIDs()
            docIDIndex = allDocIDs.Count - 1 ' last record is new
        Catch ex As Exception
            MessageBox.Show("Error creating new record in Document table: " + ex.Message)
        Finally
            dbConnection.Close()
        End Try
    End Sub

    ' called by SaveButton click; updates an existing record in the Document table from textbox contents
    Private Sub UpdateExistingDocument()
        Try
            dbConnection.Open()
            Dim cmd As OleDbCommand = New OleDbCommand("UPDATE Document " &
                            "SET DocType=@DocType, DocTitle=@DocTitle, SectionTitle=@SectionTitle, City=@City, State=@State, Publisher=@Publiser, DocYear=@DocYear, DocMonth=@DocMonth, DocDay=@DocDay, StartPage=@StartPage, EndPage=@EndPage, URL=@URL, VolumeNum=@VolumeNum, IssueNum=@IssueNum " &
                            "WHERE DocID=" + GetCurrentDocID(), dbConnection)
            ' using parameters to prevent SQL injection from user input
            With cmd.Parameters
                .Add(New OleDbParameter("@DocType", DocTypeComboBox.SelectedItem))
                .Add(New OleDbParameter("@DocTitle", DocTitleTextBox.Text))
                .Add(New OleDbParameter("@SectionTitle", SectionTitleTextBox.Text))
                .Add(New OleDbParameter("@City", CityTextBox.Text))
                .Add(New OleDbParameter("@State", StateTextBox.Text))
                .Add(New OleDbParameter("@Publisher", PublisherTextBox.Text))
                .Add(New OleDbParameter("@DocYear", NumberOrNull(DocYearTextBox.Text)))
                .Add(New OleDbParameter("@DocMonth", NumberOrNull(DocMonthTextBox.Text)))
                .Add(New OleDbParameter("@DocDay", NumberOrNull(DocDayTextBox.Text)))
                .Add(New OleDbParameter("@StartPage", NumberOrNull(StartPageTextBox.Text)))
                .Add(New OleDbParameter("@EndPage", NumberOrNull(EndPageTextBox.Text)))
                .Add(New OleDbParameter("@URL", URLTextBox.Text))
                .Add(New OleDbParameter("@VolumeNum", NumberOrNull(VolumeTextBox.Text)))
                .Add(New OleDbParameter("@IssueNum", NumberOrNull(IssueTextBox.Text)))
            End With
            ' execute the INSERT
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Error updating record in Document table: " + ex.Message)
        Finally
            dbConnection.Close()
        End Try
    End Sub

    ' click of the Add More Authors button pops up a dialog to enter more authors
    Private Sub MoreAuthorsButton_Click(sender As Object, e As EventArgs) Handles MoreAuthorsButton.Click

        Dim popup As MoreAuthorsForm = New MoreAuthorsForm() ' create a form

        ' if the name in the UpdatePanel's author textboxes are not in displayedPersonIDs, then add it to the list
        ' this can potentially add new names to PersonIDs; which is okay they'll be cleaned up after Save and they won't be associated until saved
        Dim frontPersonID = AuthorNames.GetPersonIDByName(AuthorFirstNameTextBox.Text, AuthorMITextBox.Text, AuthorLastNameTextBox.Text)
        If frontPersonID <> "" And Not displayedPersonIDs.Contains(frontPersonID) Then
            If displayedPersonIDs.Count = 0 Then ' if there is no authors in displayedPersonIDs list
                displayedPersonIDs.Add(frontPersonID) ' then add them
            Else
                displayedPersonIDs(0) = frontPersonID ' otherwise replacing first name that was in that box
            End If
        End If

        ' update textboxes in the popup form to show known names associated for this document
        popup.SetPersonIDs(displayedPersonIDs)

        Dim result As DialogResult = popup.ShowDialog() ' show the form

        If result = DialogResult.OK Then ' if they clicked OK then accept names entered in MoreAuthors popup
            ' update list of PersonIDs associated with this document from the MoreAuthors popup
            displayedPersonIDs = popup.GetPersonIDs()
            ' and update the UpdatePanel's display of the first author (if any)
            UpdateFrontAuthor()
        End If

    End Sub

    ' this updates the "front author"s first, middle and last name in UpdatePanel to the first displayedPersonID, if any
    Private Sub UpdateFrontAuthor()
        If displayedPersonIDs.Count > 0 Then ' if there's at least one author
            Dim person = AuthorNames.GetPersonByPersonID(displayedPersonIDs(0))
            AuthorFirstNameTextBox.Text = person.firstName
            AuthorMITextBox.Text = person.middleInit
            AuthorLastNameTextBox.Text = person.lastName
            If displayedPersonIDs.Count > 1 Then ' if there's more than 2 then change button to "+x More Authors"
                MoreAuthorsButton.Text = "+" & displayedPersonIDs.Count - 1 & " More Authors"
            Else
                MoreAuthorsButton.Text = "Add More Authors"
            End If
        Else ' there are no authors in displayedPersonIDs
            AuthorFirstNameTextBox.Text = ""
            AuthorMITextBox.Text = ""
            AuthorLastNameTextBox.Text = ""
            MoreAuthorsButton.Text = "Add More Authors"
        End If
    End Sub

    ' this is called during a save and will update the Author junction table with the displayed PersonIDs with the current DocID
    ' in three steps: 1) delete every record with the currently viewed DocID, 2) add all PersonID+DocID pairs for this DocID, and
    ' 3) delete any orphaned PersonIDs from Person that are no longer associated with a DocID in Author table
    Private Sub UpdateAuthorTable()
        ' step 1: delete all rows in Author table that have the current DocID
        Try
            dbConnection.Open()
            Dim cmd As OleDbCommand = New OleDbCommand("DELETE FROM Author WHERE DocID=" & GetCurrentDocID(), dbConnection)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Error updating Authors step 1: " & ex.Message) ' "step 1" intentionally ambiguous: don't want to tell user there was a problem deleting stuff when they're trying to save
            dbConnection.Close()
            Return ' if there were problems, don't stick around to try the rest of the update
        End Try

        ' step 2: add to Author all PersonID+DocID pairs for the current DocID
        Try
            ' the following builds a statement like "INSERT IGNORE INTO Author VALUES (1,2), (3,4), (3, 2)" where each ordered pair is a PersonID,DocID
            For Each personID As String In displayedPersonIDs
                Dim cmd As OleDbCommand = New OleDbCommand("INSERT INTO AUTHOR VALUES (" & personID & ", " & GetCurrentDocID() & ");", dbConnection)
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            MessageBox.Show("Error updating Authors step 2: " & ex.Message)
            dbConnection.Close()
            Return ' again, if any problems, don't stick around; leave immediately
        End Try

        ' step 3: remove any PersonIDs from Person table that are no longer in Author table (meaning no longer associated with a document)
        Try
            Dim cmd As OleDbCommand = New OleDbCommand("DELETE * FROM PERSON WHERE Person.PersonID IN (SELECT Person.PersonID FROM Person LEFT JOIN Author ON Person.PersonID = Author.PersonID WHERE Author.DocID IS NULL);", dbConnection)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Error updating Authors step 3: " & ex.Message)
            ' this was last step so can close db in finally instead of catch
        Finally
            dbConnection.Close()
        End Try
    End Sub

    ' Name, title and year are required fields. when the text in those fields change, enable the Save button only if all three fields have a value
    Private Sub RequiredFields_TextChanged(sender As Object, e As EventArgs) Handles DocYearTextBox.TextChanged, DocTitleTextBox.TextChanged, AuthorFirstNameTextBox.TextChanged
        SaveButton.Enabled = Not (AuthorFirstNameTextBox.Text.Trim() = "" Or DocTitleTextBox.Text.Trim() = "" Or DocYearTextBox.Text.Trim() = "")
    End Sub

    ' When a doc type is chosen (combo box changes), update controls to be more specific to that type
    Private Sub DocTypeComboBox_SelectedValueChanged(sender As Object, e As EventArgs) Handles DocTypeComboBox.SelectedValueChanged
        Select Case DocTypeComboBox.SelectedItem.ToString()
            Case "Book"
                DocTitleLabel.Text = "Book Title"
                SectionTitleLabel.Text = "Book Section"
            Case "Journal"
                DocTitleLabel.Text = "Journal Title"
                SectionTitleLabel.Text = "Article Title"
            Case Else
                DocTitleLabel.Text = "Conference Name"
                SectionTitleLabel.Text = "Paper Title"
        End Select
    End Sub
End Class
