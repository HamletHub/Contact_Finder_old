Public Class FAddKeywords
    ReadOnly sql As New SQLiteControl 'Declare SQL variable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim AList As New ArrayList 'Declare array list

        Try
            For Each keyword As String In RichTextBox1.Lines 'For each keyword...
                For Each town As String In RichTextBox2.Lines 'And for each town...
                    sql.AddParam("@keyword", keyword & " in " & town & " ct") 'Generate SQL parameter
                    sql.ExecQuery("SELECT * FROM GoogleDoneKeywords WHERE keywords LIKE @keyword;") 'Execute query

                    If sql.RecordCount = 0 Then Dim value As String = keyword & " in " & town & " ct" : AList.Add(value) 'If SQL records count is 0, that means the same keyword does not exist in the database, and we can add it to the list
                Next

            Next
        Catch ex As Exception
        End Try

        Dim rnd As New Random() 'Declare random
        Dim Stringarray() As Object = AList.ToArray()
        Dim ShuffledItems = Stringarray.OrderBy(Function() rnd.Next).ToArray() 'Shuffle keywords

        For Each value In ShuffledItems
            sql.AddParam("@keyword", value)
            sql.ExecQuery("INSERT INTO PendingKeywords (BKeywords) " &
                                          "VALUES (@keyword);") 'Add shuffled keywords to the database
        Next

        Form1.Show() 'Show previous form
        Me.Hide() 'Hide this form
    End Sub
End Class