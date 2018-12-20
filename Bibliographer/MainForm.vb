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

Partial Public Class MainForm

    ' database connectivity used throughout MainForm
    Dim dbConnection As OleDbConnection
    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Documents.accdb;Persist Security Info=True"

    ' class-wide variables to store information about documents on screen
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
End Class
