' AuthorNames.vb
' CIS410 Fall 2018 Final Project
' Members: Jeff Heskett, Phillip Foss, Ian Shaw
' Purpose: This project will allow a user to store and retrieve bibliographic references for documents by author or year
'
' This static module is for saving and recalling the names of PersonIDs without querying the database every time. Additionally,
' if a request is made for the PersonID of a firstName,MI,lastName that doesn't exist, this module will create a PersonID for that name

' structure to store a record in the Person table
Imports System.Data.OleDb

Structure Person
    Dim personID As String
    Dim firstName As String
    Dim middleInit As String
    Dim lastName As String
End Structure

Module AuthorNames

    ' this list of Person structs will be a copy of the People table
    Dim personData As List(Of Person) = New List(Of Person)

    ' connection information created in MainForm, shared in ShareDBConnection
    Dim dbConnection As OleDbConnection

    ' This should be called once the MainForm has a connection defined; this module will use the same one
    Public Sub ShareDBConnection(connection As OleDbConnection)
        dbConnection = connection
        ' While here, pull in the names in the database
        UpdatePersonIDs()
    End Sub

    ' This populates personData with all PersonIDs and their fields from the database
    Private Sub UpdatePersonIDs()
        personData.Clear() ' wiping list will require re-creating every Person
        Try
            dbConnection.Open()
            Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM Person", dbConnection)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            While reader.Read()
                Dim person As Person = New Person() ' creating a new person that will be added to the list
                person.personID = reader("PersonID").ToString()
                person.firstName = reader("FirstName").ToString()
                person.middleInit = reader("MiddleInit").ToString()
                person.lastName = reader("LastName").ToString()
                personData.Add(person) ' add this new person to the list
            End While
        Catch ex As Exception
            MessageBox.Show("Couldn't read from Person table in database: " & ex.Message)
        Finally
            dbConnection.Close()
        End Try
    End Sub

    ' This function takes a first, middle and last name and returns the PersonID of that person, creating a new
    ' PersonID if they don't exist yet. (Actually, it has the database create the new PersonID (autonum), by
    ' inserting the new person and then finding the PersonID the database assigned to them)
    Public Function GetPersonIDByName(firstName As String, middleInit As String, lastName As String) As String

        ' first trim any whitespace before/after the names
        firstName = firstName.Trim()
        middleInit = middleInit.Trim()
        lastName = lastName.Trim()

        ' all names must have a first name to be valid; if no first name, return "" for an invalid PersonID
        If firstName.Length = 0 Then
            Return ""
        End If

        ' look for an existing person with the precisely same name
        For Each person As Person In personData
            If person.firstName = firstName And person.middleInit = middleInit And person.lastName = lastName Then
                Return person.personID ' if someone was found with same name, return their PersonID
            End If
        Next

        ' if we reached here, a match wasn't found. create a new one by insert the new name into the database


    End Function

End Module
