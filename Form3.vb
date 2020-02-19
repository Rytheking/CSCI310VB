Public Class Form3
    Public numtot As Integer
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        numtot = Convert.ToInt32(TextBox2.Text) - Convert.ToInt32(TextBox3.Text)
        Me.Close()
    End Sub
End Class