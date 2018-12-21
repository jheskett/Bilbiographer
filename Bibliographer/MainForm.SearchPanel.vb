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
        If APARadioButton.Checked Then ' APA
            ResultsRichTextBox.Text &= GetAPAFormat(document) & vbCrLf & vbCrLf
        ElseIf MLARadioButton.Checked Then ' MLA
            ResultsRichTextBox.Text &= GetMLAFormat(document) & vbCrLf & vbCrLf
        Else ' IEEE format
            ResultsRichTextBox.Text &= GetIEEEFormat(document) & vbCrLf & vbCrLf
        End If

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
            ' default sort is by author name
            Dim sortBy As String = "ORDER BY LastName, FirstName, DocYear DESC, DocTitle"
            If SortYearRadioButton.Checked Then ' if sorting by year, move it to first order sort
                sortBy = "ORDER BY DocYear DESC, LastName, FirstName, DocTitle"
            End If
            ' this joins the People and Document table with the Author junction table to query every Author expanded to its People and Document records (M:M relationship)
            Dim cmd As OleDbCommand = New OleDbCommand("SELECT Author.PersonID, Author.DocID, FirstName, MiddleInit, LastName, DocType, DocTitle, SectionTitle, City, State, Publisher, DocYear, DocMonth, DocDay, StartPage, EndPage, URL, VolumeNum, IssueNum FROM ((Person INNER JOIN Author ON Person.PersonID = Author.PersonID) INNER JOIN Document ON Author.DocID = Document.DocID) " & sortBy, dbConnection)
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
        SearchButton.PerformClick() ' click the SearchButton
    End Sub

    ' hitting Enter in the SearchTextBox is made equivalent to clicking the search button to do the search
    Private Sub SearchTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles SearchTextBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            SearchButton.PerformClick()
            e.Handled = True ' stop the annoying error sound when hitting Enter in a textbox
            e.SuppressKeyPress = True ' this seems needed too
        End If
    End Sub


    ' if the sort radio buttons (Author or Year) are checked, the sort order changes
    Private Sub RadioButton_Click(sender As Object, e As EventArgs) Handles SortYearRadioButton.Click, SortAuthorRadioButton.Click, APARadioButton.Click, MLARadioButton.Click, IEEERadioButton.Click
        SearchButton.PerformClick()
    End Sub

    ' returns the given document as an APA-formatted string
    Private Function GetAPAFormat(document As Document) As String
        Dim output As String = ""

        output &= GetAPAAuthorsList(document.authors) ' start with list of authors

        Select Case document.docType

            Case "Book"
                output &= " (" & document.docYear & "). " ' then (Year)
                If document.sectionTitle <> "" Then ' then section of book if there is one
                    output &= document.sectionTitle & ". "
                End If
                output &= document.docTitle & ". " ' title of book
                If document.startPage <> "" And document.endPage <> "" Then ' pages
                    output &= "(pp. " & document.startPage & "-" & document.endPage & "). "
                ElseIf document.startPage <> "" Then
                    output &= "(p. " & document.startPage & "). "
                End If
                If document.city <> "" And document.state <> "" Then ' publisher location
                    output &= document.city & ", " & document.state & ": "
                ElseIf document.city <> "" Then
                    output &= document.city & ": "
                End If
                If document.publisher <> "" Then ' publisher
                    output &= document.publisher & "."
                End If

            Case "Journal"
                output &= " (" & document.docYear & "). "
                If document.sectionTitle <> "" Then
                    output &= document.sectionTitle & ". "
                End If
                output &= document.docTitle & ". "
                If document.volumeNum <> "" Then
                    output &= "vol. " & document.volumeNum
                    If document.issueNum <> "" Then
                        output &= "(" & document.issueNum & ")"
                    End If
                    output &= ". "
                End If
                If document.startPage <> "" And document.endPage <> "" Then ' pages
                    output &= document.startPage & "-" & document.endPage & ". "
                ElseIf document.startPage <> "" Then
                    output &= document.startPage & ". "
                End If
                If document.url <> "" Then
                    If document.url.Substring(0, 3).ToUpper() = "DOI" Then
                        output &= document.url
                    Else
                        output &= "Retrieved from " & document.url
                    End If
                End If

            Case Else ' Conference
                If document.docMonth <> "" And document.docDay <> "" Then
                    output &= "(" & document.docYear & ", " & DateAndTime.MonthName(CInt(document.docMonth), True) & " " & document.docDay & "). "
                ElseIf document.docMonth <> "" Then
                    output &= "(" & document.docYear & ", " & DateAndTime.MonthName(CInt(document.docMonth), True) & "). "
                Else
                    output &= "(" & document.docYear & "). "
                End If
                If document.sectionTitle <> "" Then
                    output &= document.sectionTitle & ". "
                End If
                output &= "Presented at " & document.docTitle & ". "
                If document.city <> "" Then
                    output &= document.city & ", "
                End If
                If document.publisher <> "" And document.state <> "" Then
                    output &= document.state & ": " & document.publisher & ". "
                ElseIf document.state <> "" Then
                    output &= document.state & ". "
                End If
                If document.url <> "" Then
                    If document.url.Substring(0, 3).ToUpper() = "DOI" Then
                        output &= document.url
                    Else
                        output &= "Retrieved from " & document.url
                    End If
                End If

        End Select

        Return output
    End Function

    ' returns a list of authors in APA format (LastName, F.M.)
    Private Function GetAPAAuthorsList(authors As List(Of Person)) As String
        Dim output As String = ""
        For Each author As Person In authors
            If author.lastName.Length > 0 Then
                output &= author.lastName
                If author.firstName.Length > 0 Then
                    output &= ", " & author.firstName.Substring(0, 1) & "." ' only using firt initial of first name in APA
                    If author.middleInit.Length > 0 Then
                        output &= author.middleInit & "." ' let whole middle initial add (in case two letters like John RR Tolkien)
                    End If
                End If
            ElseIf author.firstName.Length > 0 Then ' if no last name given (goes by one-word psuedonym) use whole first name and ignore middle
                output &= author.firstName
            End If
            output &= ", "
        Next
        ' trim trailing ", " if there's any names added
        If output.Length > 1 Then
            output = output.Remove(output.Length - 2)
        End If
        Return output
    End Function

    Private Function GetMLAFormat(document As Document) As String
        Dim output As String = GetMLAAuthorsList(document.authors) & ". "
        output = output.Replace("..", ".") ' any middleinits may leave extra . so remove them if so

        If document.sectionTitle <> "" Then
            output &= """" & document.sectionTitle & "."" "
        End If
        output &= document.docTitle

        Select Case document.docType
            Case "Book"
                output &= ". "
                If document.publisher <> "" Then
                    output &= document.publisher & ", "
                End If
                If document.startPage <> "" Then
                    output &= document.docYear & ", "
                    If document.startPage <> "" And document.endPage <> "" Then ' pages
                        output &= "pp. " & document.startPage & "-" & document.endPage & ". "
                    Else
                        output &= "p. " & document.startPage & ". "
                    End If
                Else
                    output &= document.docYear & "."
                End If
            Case "Journal"
                If document.volumeNum <> "" Then
                    output &= "vol. " & document.volumeNum
                    If document.issueNum <> "" Then
                        output &= ", no. " & document.issueNum & ", "
                    End If
                    output &= ". "
                End If
                If document.startPage <> "" And document.endPage <> "" Then ' pages
                    output &= "pp. " & document.startPage & "-" & document.endPage & ". "
                ElseIf document.startPage <> "" Then
                    output &= "p. " & document.startPage & ". "
                End If
                If document.url <> "" Then
                    output &= document.url & "."
                End If
            Case Else ' Conference
                output &= ", "
                If document.docMonth <> "" Then
                    output &= DateAndTime.MonthName(CInt(document.docMonth), True) & " "
                    If document.docDay <> "" Then
                        output &= document.docDay & " "
                    End If
                End If
                output &= document.docYear & ", "
                If document.publisher <> "" Then
                    output &= document.publisher & ". "
                End If
                If document.url <> "" Then
                    output &= document.url & "."
                End If
        End Select


        Return output
    End Function

    ' MLA lists like LastName, FirstName M. (up to two names, more than that get an et al after first author)
    Private Function GetMLAAuthorsList(authors As List(Of Person)) As String
        If authors.Count = 1 Then
            Return GetSingleAuthorName(authors(0))
        ElseIf authors.Count = 2 Then
            Return GetSingleAuthorName(authors(0)) & ", and " & GetSingleAuthorName(authors(1))
        ElseIf authors.Count > 2 Then
            Return GetSingleAuthorName(authors(0)) & ", et al."
        Else
            Return ""
        End If
    End Function

    ' take a single Person and returns a name formatted as LastName, FirstName M.
    Private Function GetSingleAuthorName(author As Person) As String
        If author.lastName = "" Then
            Return author.firstName
        ElseIf author.middleInit <> "" Then
            Return author.lastName & ", " & author.firstName & " " & author.middleInit & "."
        Else
            Return author.lastName & ", " & author.firstName
        End If
    End Function

    Private Function GetIEEEFormat(document As Document) As String
        Dim output As String = ""

        output &= GetIEEEAuthorsList(document.authors) & ", "

        If document.sectionTitle <> "" Then
            output &= """" & document.sectionTitle & "."" "
        End If
        output &= document.docTitle & ", "
        If document.city <> "" And document.state <> "" And document.publisher <> "" Then
            output &= document.city & ", " & document.state & ": " & document.publisher & ", "
        ElseIf document.city <> "" And document.publisher <> "" Then
            output &= document.city & ": " & document.publisher & ", "
        Else
            If document.publisher <> "" Then
                output &= document.publisher & ", "
            End If
        End If
        If document.volumeNum <> "" Then
            output &= "vol. " & document.volumeNum
            If document.issueNum <> "" Then
                output &= ", no. " & document.issueNum & ", "
            End If
            output &= ", "
        End If
        If document.startPage <> "" And document.endPage <> "" Then ' pages
            output &= "pp. " & document.startPage & "-" & document.endPage & ". "
        ElseIf document.startPage <> "" Then
            output &= "p. " & document.startPage & ". "
        End If
        If document.url <> "" Then
            output &= document.url & ". "
        End If
        output &= document.docYear

        Return output
    End Function

    ' returns a string of author names in IEEE format (F. M. LastName)
    Private Function GetIEEEAuthorsList(authors As List(Of Person)) As String
        Dim output As String = ""
        For Each author As Person In authors
            If author.lastName <> "" Then
                output &= author.firstName.Substring(0, 1) & ". "
                If author.middleInit <> "" Then
                    output &= author.middleInit & ". "
                End If
                output &= author.lastName
            Else
                output &= author.firstName
            End If
            output &= ", "
        Next
        ' trim trailing ", " if there's any names added
        If output.Length > 1 Then
            output = output.Remove(output.Length - 2)
        End If
        Return output
    End Function

End Class
