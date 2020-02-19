Public Class Form2
    Structure booktype
        Public id As Integer
        Public title As String
        Public author As String
        Public genre As String
        Public price As Double
        Public num As Integer
        Public value As Double
    End Structure
    Dim temp As smalltownLib.booktype
    Private Sub form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.AutoSizeMode = False
        Label9.Text = smalltownLib.s(0).id
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles addBut.Click
        temp.id = newIDBox.Text
        temp.title = newTitleBox.Text
        temp.author = newAuthorBox.Text
        temp.genre = newGenreBox.Text
        temp.price = newPriceBox.Text
        temp.num = newNumBox.Text
        temp.value = smalltownLib.s(smalltownLib.numbooks + 1).price * smalltownLib.s(smalltownLib.numbooks + 1).price
        smallTownLibrary_v2.smalltownLib.s(smalltownLib.numbooks) = temp
        Me.Close()
    End Sub
    Function getid() As Integer
        Return temp.id
    End Function
    Function gettitle() As Integer
        Return temp.title
    End Function
    Function getauthor() As Integer
        Return temp.author
    End Function
    Function getgenre() As Integer
        Return temp.genre
    End Function
    Function getprice() As Integer
        Return temp.price
    End Function
    Function getnum() As Integer
        Return temp.num
    End Function

    Private Sub cancelBut_Click(sender As Object, e As EventArgs) Handles cancelBut.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label9.Text = newIDBox.Text
    End Sub
End Class