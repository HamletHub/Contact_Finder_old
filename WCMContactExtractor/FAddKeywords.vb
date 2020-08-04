Public Class FAddKeywords
    ReadOnly sql As New SQLiteControl
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim AList As New ArrayList

        Try
            For Each keyword As String In RichTextBox1.Lines
                For Each town As String In RichTextBox2.Lines
                    sql.AddParam("@keyword", keyword & " in " & town & " ct")
                    sql.ExecQuery("SELECT * FROM GoogleDoneKeywords WHERE keywords LIKE @keyword;")

                    If sql.RecordCount = 0 Then
                        Dim value As String = keyword & " in " & town & " ct"
                        AList.Add(value)
                    End If

                Next

            Next
        Catch ex As Exception
        End Try

        Dim rnd As New Random()
        Dim Stringarray() As Object = AList.ToArray()
        Dim ShuffledItems = Stringarray.OrderBy(Function() rnd.Next).ToArray()

        For Each value In ShuffledItems
            sql.AddParam("@keyword", value)
            sql.ExecQuery("INSERT INTO PendingKeywords (BKeywords) " &
                                          "VALUES (@keyword);")
        Next


        Form1.Show()
        Me.Hide()
    End Sub
End Class