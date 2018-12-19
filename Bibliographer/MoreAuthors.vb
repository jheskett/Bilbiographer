' AuthorNames.vb
' CIS410 Fall 2018 Final Project
' Members: Jeff Heskett, Phillip Foss, Ian Shaw
' Purpose: This project will allow a user to store and retrieve bibliographic references for documents by author or year
'
' This is a popup dialog where the user can enter the names of up to seven authors that contributed on a document.

Public Class MoreAuthorsForm

    ' when passed a list of PersonIDs, it will fill in the seven name textboxes on the form
    Public Sub SetPersonIDs(personIDs As List(Of String))
        Dim authors As List(Of Person) = New List(Of Person)
        ' first create a list of Person structs to reflect the list of PersonIDs
        For index = 0 To 6
            If index < personIDs.Count Then ' while there's names in given list, get their Person struct and add to a list
                authors.Add(GetPersonByPersonID(personIDs(index)))
            Else
                authors.Add(New Person()) ' pad list with people without a name
            End If
        Next
        ' next populate all 21 text fields (7x(first+middle+last)); this is rather ugly
        FirstName1TextBox.Text = authors(0).firstName
        MiddleInit1TextBox.Text = authors(0).middleInit
        LastName1TextBox.Text = authors(0).lastName
        FirstName2TextBox.Text = authors(1).firstName
        MiddleInit2TextBox.Text = authors(1).middleInit
        LastName2TextBox.Text = authors(1).lastName
        FirstName3TextBox.Text = authors(2).firstName
        MiddleInit3TextBox.Text = authors(2).middleInit
        LastName3TextBox.Text = authors(2).lastName
        FirstName4TextBox.Text = authors(3).firstName
        MiddleInit4TextBox.Text = authors(3).middleInit
        LastName4TextBox.Text = authors(3).lastName
        FirstName5TextBox.Text = authors(4).firstName
        MiddleInit5TextBox.Text = authors(4).middleInit
        LastName5TextBox.Text = authors(4).lastName
        FirstName6TextBox.Text = authors(5).firstName
        MiddleInit6TextBox.Text = authors(5).middleInit
        LastName6TextBox.Text = authors(5).lastName
        FirstName7TextBox.Text = authors(6).firstName
        MiddleInit7TextBox.Text = authors(6).middleInit
        LastName7TextBox.Text = authors(6).lastName
    End Sub

    ' this returns a List of PersonIDs that were in the list of names
    Public Function GetPersonIDs() As List(Of String)
        Dim authors As List(Of String) = New List(Of String)

        ' now convert every name in the textbox (that has something in first name textbox) and create a list of PersonIDs from those
        ' name #1
        Dim personID As String = AuthorNames.GetPersonIDByName(FirstName1TextBox.Text, MiddleInit1TextBox.Text, LastName1TextBox.Text)
        If personID <> "" Then
            authors.Add(personID)
        End If
        ' name #2
        personID = AuthorNames.GetPersonIDByName(FirstName2TextBox.Text, MiddleInit2TextBox.Text, LastName2TextBox.Text)
        If personID <> "" Then
            authors.Add(personID)
        End If
        ' name #3
        personID = AuthorNames.GetPersonIDByName(FirstName3TextBox.Text, MiddleInit3TextBox.Text, LastName3TextBox.Text)
        If personID <> "" Then
            authors.Add(personID)
        End If
        ' name #4
        personID = AuthorNames.GetPersonIDByName(FirstName4TextBox.Text, MiddleInit4TextBox.Text, LastName4TextBox.Text)
        If personID <> "" Then
            authors.Add(personID)
        End If
        ' name #5
        personID = AuthorNames.GetPersonIDByName(FirstName5TextBox.Text, MiddleInit5TextBox.Text, LastName5TextBox.Text)
        If personID <> "" Then
            authors.Add(personID)
        End If
        ' name #6
        personID = AuthorNames.GetPersonIDByName(FirstName6TextBox.Text, MiddleInit6TextBox.Text, LastName6TextBox.Text)
        If personID <> "" Then
            authors.Add(personID)
        End If
        ' name #7
        personID = AuthorNames.GetPersonIDByName(FirstName7TextBox.Text, MiddleInit7TextBox.Text, LastName7TextBox.Text)
        If personID <> "" Then
            authors.Add(personID)
        End If

        Return authors
    End Function


End Class