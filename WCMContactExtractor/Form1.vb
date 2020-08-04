Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not RichTextBox1.Text = "" Then
            If Not ComboBox1.SelectedItem = "" Then FDoWork.Show() : Me.Hide() 'Show form
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Danbury")
        ComboBox1.Items.Add("Darien")
        ComboBox1.Items.Add("Milford")
        ComboBox1.Items.Add("New Canaan")
        ComboBox1.Items.Add("Newtown")
        ComboBox1.Items.Add("Norwalk")
        ComboBox1.Items.Add("Ridgefield")
        ComboBox1.Items.Add("Stamford")
        ComboBox1.Items.Add("Westport")

        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FAddKeywords.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FWorkSpace.Show()
        Me.Hide()
    End Sub
End Class
