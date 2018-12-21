' MainForm.SearchPanel.vb
' CIS410 Fall 2018 Final Project
' Members: Jeff Heskett, Phillip Foss, Ian Shaw
' Purpose: This project will allow a user to store and retrieve bibliographic references for documents by author or year
'
' This is a Partial class of MainForm and contains the code for the SearchPanel in a separate file. It is not a distinct
' class but merely SearchPanel-related code split from MainForm.vb into a separate file so the code is better organized.
'
' For clarity, do not define any class-wide properties in this file; use the main MainForm for that.

Imports System.Data.OleDb

Partial Public Class MainForm

    ' Update button in the topright of the SearchPanel switches to the UpdatePanel
    Private Sub UpdateButton_Click(sender As Object, e As EventArgs) Handles UpdateButton.Click
        SearchPanel.Visible = False
        UpdatePanel.Visible = True

        ' When Update button is hit and switching to UpdatePanel, update displayed record
        UpdateDisplayedRecord()
    End Sub

    ' Search button to the right of the SearchTextBox will search for all documents with the text
    Private Sub SearchButton_Click(sender As Object, e As EventArgs) Handles SearchButton.Click
        UpdateAllDocuments()
        Dim searchText As String = SearchTextBox.Text.Trim.ToUpper() ' making search text case insensitive
        ResultsRichTextBox.Text = ""
        For Each document As Document In allDocuments
            If document.docTitle.ToUpper().IndexOf(searchText) >= 0 Then
                AddDocumentToResults(document)
            Else
                For Each author As Person In document.authors
                    If author.firstName.ToUpper().IndexOf(searchText) >= 0 Or author.lastName.ToUpper().IndexOf(searchText) >= 0 Then
                        AddDocumentToResults(document)
                        Exit For ' only want to add document once if a single author found
                    End If
                Next
            End If
        Next
        SearchTextBox.Select() ' return focus back to the SearchTextBox
    End Sub

    ' adds the given document to the richtextbox where search results are listed (here is where it can be formatted)
    Private Sub AddDocumentToResults(document As Document)
        ResultsRichTextBox.Text &= document.docTitle & vbCrLf & vbCrLf
    End Sub

    ' this returns the Document in the allDocuments list from its docID, if it exists; if not it returns null
    Private Function GetDocumentByDocID(docID As String) As Document
        For Each document As Document In allDocuments
            If document.docID = docID Then
                Return document
            End If
        Next
        Return New Document()
    End Function

    ' this populates the allDocuments list with all documents (and their authors) from the database
    Private Sub UpdateAllDocuments()
        allDocuments.Clear() ' wipe old data first
        Try
            dbConnection.Open()
            ' this joins the People and Document table with the Author junction table to query every Author expanded to its People and Document records (M:M relationship)
            Dim cmd As OleDbCommand = New OleDbCommand("SELECT Author.PersonID, Author.DocID, FirstName, MiddleInit, LastName, DocType, DocTitle, SectionTitle, City, State, Publisher, DocYear, DocMonth, DocDay, StartPage, EndPage, URL, VolumeNum, IssueNum FROM ((Person INNER JOIN Author ON Person.PersonID = Author.PersonID) INNER JOIN Document ON Author.DocID = Document.DocID)", dbConnection)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            Dim count = 0
            While reader.Read()
                Dim docID As String = reader("DocID").ToString()
                Dim document As Document = GetDocumentByDocID(docID)
                ' if this document doesn't exist in allDocuments, then add it (for now just the Document fields)
                If document.docID = "" Then
                    count += 1
                    document.docID = docID
                    document.docType = reader("DocType").ToString()
                    document.docTitle = reader("DocTitle").ToString()
                    document.sectionTitle = reader("SectionTitle").ToString()
                    document.city = reader("City").ToString()
                    document.state = reader("State").ToString()
                    document.publisher = reader("Publisher").ToString()
                    document.docYear = reader("DocYear").ToString()
                    document.docMonth = reader("DocMonth").ToString()
                    document.docDay = reader("DocDay").ToString()
                    document.startPage = reader("StartPage").ToString()
                    document.endPage = reader("EndPage").ToString()
                    document.url = reader("URL").ToString()
                    document.volumeNum = reader("VolumeNum").ToString()
                    document.issueNum = reader("IssueNum").ToString()
                    document.authors = New List(Of Person)
                    allDocuments.Add(document)
                End If
                ' this query result is one author, create a Person to reflect this author
                Dim author As Person = New Person()
                author.personID = reader("PersonID").ToString()
                author.firstName = reader("FirstName").ToString()
                author.middleInit = reader("MiddleInit").ToString()
                author.lastName = reader("LastName").ToString()
                ' now add this Person to the document's list of authors
                document.authors.Add(author)
            End While
        Catch ex As Exception
            MessageBox.Show("Error querying Documents: " & ex.Message)
        Finally
            dbConnection.Close()
        End Try
    End Sub

    ' clicking the X button will reset the searchbox and click the button to force a search of all documents
    Private Sub SearchResetButton_Click(sender As Object, e As EventArgs) Handles SearchResetButton.Click
        SearchTextBox.Text = ""
        SearchButton_Click(SearchButton, e) ' click the SearchButton
    End Sub

    ' hitting Enter in the SearchTextBox is made equivalent to clicking the search button to do the search
    Private Sub SearchTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles SearchTextBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            SearchButton_Click(SearchButton, e) ' click the SearchButton
            e.Handled = True ' stop the annoying error sound when hitting Enter in a textbox
            e.SuppressKeyPress = True ' this seems needed too
        End If
    End Sub

End Class
