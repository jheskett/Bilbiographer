' Bibliographer.vb
' CIS410 Fall 2018 Final Project
' Members: Jeff Heskett, Phillip Foss, Ian Shaw
' Purpose: This project will allow a user to store and retrieve bibliographic references for documents by author or year
'
' This is the main form that primarily has an SearchPanel and its controls to look up existing document refrences,
' and UpdatePanel to add, update and delete document references.
'
' For SearchPanel code, see MainForm.SearchPanel.vb
' For UpdatePanel code, see MainForm.UpdatePanel.vb

Imports System.Data.OleDb

' structure to store a record from the Person table
Structure Person
    Dim personID As String
    Dim firstName As String
    Dim middleInit As String
    Dim lastName As String
End Structure

' structure to store a document and its authors
Structure Document
    Dim authors As List(Of Person)
    Dim docID As String
    Dim docType As String
    Dim docTitle As String
    Dim sectionTitle As String
    Dim city As String
    Dim state As String
    Dim publisher As String
    Dim docYear As String
    Dim docMonth As String
    Dim docDay As String
    Dim startPage As String
    Dim endPage As String
    Dim url As String
    Dim volumeNum As String
    Dim issueNum As String
End Structure

Partial Public Class MainForm

    ' database connectivity used throughout MainForm
    Dim dbConnection As OleDbConnection
    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Documents.accdb;Persist Security Info=True"

    ' class-wide variables to store information about documents on screen
    Dim allDocIDs As New List(Of String) ' list of all DocIDs in the database
    Dim docIDIndex As Integer = 0 ' index into allDocIDs of the currently-viewed DocID (or -1 if a new one); starting at 0 or first DocID
    Dim displayedPersonIDs As New List(Of String) ' list of all PersonIDs for the currently viewed DocID

    Dim allDocuments As New List(Of Document) ' list of all documents stored in the Document structure (which includes authors)

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

        ' start everything off with an empty search to list all documents
        SearchButton.PerformClick()
    End Sub


End Class
