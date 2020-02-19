Imports System.IO

Public Class smalltownLib
    Structure booktype
        Public id As Integer
        Public title As String
        Public author As String
        Public genre As String
        Public price As Double
        Public num As Integer
        Public value As Double
    End Structure
    Const maxbook = 50
    Public s(maxbook), tem(maxbook), temp(maxbook) As booktype
    Public numbooks, numfiltered As Integer
    Dim totalAccum As Integer = 0
    Dim filtersetting As Integer
    Dim i As Integer
    Dim fname As String
    Dim fsize As Integer = 14
    Sub showem(i As Integer)
        If (filtersetting = 0) Then
            bookNumLab.Text = s(i).id
            titleLab.Text = s(i).title
            authorLab.Text = s(i).author
            genreLab.Text = s(i).genre
            priceLab.Text = "$" + (s(i).price).ToString
            numCopiesLab.Text = s(i).num
            whatisi.Text = "Current Book: " + (i + 1).ToString
            total.Text = totalAccum
            uniqueLab.Text = numbooks
            valLab.Text = s(i).value
        Else
            bookNumLab.Text = tem(i).id
            titleLab.Text = tem(i).title
            authorLab.Text = tem(i).author
            genreLab.Text = tem(i).genre
            priceLab.Text = "$" + (tem(i).price).ToString
            numCopiesLab.Text = tem(i).num
            whatisi.Text = "Current Book: " + (i + 1).ToString
            total.Text = "N/A"
            uniqueLab.Text = numfiltered
            valLab.Text = tem(i).value
        End If

    End Sub
    Private Sub quitBut_Click(sender As Object, e As EventArgs) Handles quitBut.Click
        Application.Exit()
    End Sub
    Private Sub openBut_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Try
            Dim result As DialogResult = OpenFileDialog1.ShowDialog
            Dim currentrow As String()
            Dim i As Integer = 0
            Dim infile As String
            If (result <> DialogResult.Cancel) Then
                infile = OpenFileDialog1.FileName
                Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(infile)
                    MyReader.TextFieldType = FileIO.FieldType.Delimited
                    MyReader.SetDelimiters("|")
                    While Not MyReader.EndOfData
                        currentrow = MyReader.ReadFields
                        s(i).id = currentrow(0)
                        s(i).title = currentrow(1)
                        s(i).author = currentrow(2)
                        s(i).genre = currentrow(3)
                        s(i).price = currentrow(4)
                        s(i).num = currentrow(5)
                        s(i).value = s(i).num * s(i).price
                        totalAccum += s(i).num
                        i += 1
                    End While
                    numbooks = i
                    uniqueLab.Text = numbooks
                    total.Text = totalAccum
                    i = 0
                    showem(i)
                End Using
            End If
        Catch ex As Exception

            Dim Msg, Style, Title, Help, Ctxt, Response, MyString
        Msg = "that was invalid"
        Style = vbOK
        Title = "Invalid data"
        Help = "DEMO.HLP"
        Ctxt = 1000
        MsgBox(Msg, Style, Title)

        End Try


    End Sub
    Private Sub swapem(ByRef a As booktype, ByRef b As booktype)
        Dim temp As booktype
        temp = a
        a = b
        b = temp
    End Sub
    Private Sub report_Click(sender As Object, e As EventArgs) Handles ReportToolStripMenuItem.Click
        Using mywriter As New StreamWriter("C:\Users\ryane\OneDrive\Documents\School\Spring 2020\CSCI_310\report.txt")
            Dim i As Integer = 0
            Dim temp As String
            mywriter.WriteLine(" Book # Title       Value ")
            For i = 0 To 26
                temp += "-"
            Next
            mywriter.WriteLine(temp)
            Dim line, one, two, three As String
            For i = 0 To numbooks - 1
                one = "  " + (s(i).id).ToString
                two = s(i).title
                three = s(i).value
                mywriter.Write(one.PadRight(7))
                mywriter.Write(two.PadRight(13))
                mywriter.WriteLine(three.PadRight(26))
            Next i
            mywriter.Close()
        End Using
    End Sub
    Private Sub prevEntry_Click(sender As Object, e As EventArgs) Handles prevEntry.Click
        If filtersetting = 0 Then
            If (i.Equals(0)) Then
                i = numbooks - 1
                showem(i)
            Else
                i -= 1
                showem(i)
            End If
        Else
            If (i.Equals(0)) Then
                i = numfiltered - 1
                showem(i)
            Else
                i -= 1
                showem(i)
            End If
        End If

    End Sub
    Public Sub form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.AutoSizeMode = False
        bookNumLab.Text = ""
        titleLab.Text = ""
        authorLab.Text = ""
        genreLab.Text = ""
        priceLab.Text = ""
        numCopiesLab.Text = ""
        whatisi.Text = ""
        total.Text = ""
        showem(0)
        Me.ForeColor = Color.Black
        Me.BackColor = Color.Gray
        fname = "Arial"
        fsize = 14
        setfonts()
    End Sub
    Private Sub nextEntry_Click(sender As Object, e As EventArgs) Handles nextEntry.Click
        If filtersetting = 0 Then
            If (i.Equals(numbooks - 1)) Then
                i = 0
                showem(i)
            Else
                i += 1
                showem(i)

            End If
        Else
            If (i.Equals(numfiltered - 1) Or i >= numfiltered) Then
                i = 0
                showem(i)
            Else
                i += 1
                showem(i)

            End If
        End If

    End Sub
    Private Sub Clear_Click(sender As Object, e As EventArgs)
        bookNumLab.Text = ""
        titleLab.Text = ""
        authorLab.Text = ""
        genreLab.Text = ""
        priceLab.Text = ""
        numCopiesLab.Text = ""
        whatisi.Text = "Books: "
        total.Text = "Total Books: "
    End Sub
    Private Sub ByBookToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ByBookToolStripMenuItem.Click
        Dim index, j As Integer
        For j = 0 To numbooks - 2
            For index = 0 To numbooks - 2
                If (s(index).id > s(index + 1).id) Then
                    swapem(s(index), s(index + 1))
                End If
            Next
        Next
        i = 0
        showem(i)
    End Sub
    Private Sub ByAuthorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ByAuthorToolStripMenuItem.Click
        Dim index, j As Integer
        For j = 0 To numbooks - 2
            For index = 0 To numbooks - 2
                If (s(index).author > s(index + 1).author) Then
                    swapem(s(index), s(index + 1))
                End If
            Next
        Next
        i = 0
        showem(i)
    End Sub
    Private Sub setfonts()
        Me.Font = New Font(fname, fsize, FontStyle.Regular)
    End Sub
    Private Sub BothToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BothToolStripMenuItem.Click

    End Sub
    Private Sub BlackToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BlackToolStripMenuItem.Click
        Me.BackColor = Color.Black
        PictureBox1.BackColor = Color.Black
        PictureBox2.BackColor = Color.Black
        PictureBox3.BackColor = Color.Black
        prevEntry.BackColor = Color.Black
        nextEntry.BackColor = Color.Black
    End Sub
    Private Sub IndigoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IndigoToolStripMenuItem.Click
        Me.BackColor = Color.Indigo
        PictureBox1.BackColor = Color.Indigo
        PictureBox2.BackColor = Color.Indigo
        PictureBox3.BackColor = Color.Indigo
        prevEntry.BackColor = Color.Indigo
        nextEntry.BackColor = Color.Indigo
    End Sub
    Private Sub VioletToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VioletToolStripMenuItem.Click
        Me.BackColor = Color.Violet
        PictureBox1.BackColor = Color.Violet
        PictureBox2.BackColor = Color.Violet
        PictureBox3.BackColor = Color.Violet
        prevEntry.BackColor = Color.Violet
        nextEntry.BackColor = Color.Violet
    End Sub
    Private Sub BlueToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BlueToolStripMenuItem.Click
        Me.BackColor = Color.Blue
        PictureBox1.BackColor = Color.Blue
        PictureBox2.BackColor = Color.Blue
        PictureBox3.BackColor = Color.Blue
        prevEntry.BackColor = Color.Blue
        nextEntry.BackColor = Color.Blue
    End Sub
    Private Sub GreenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GreenToolStripMenuItem.Click
        Me.BackColor = Color.Green
        PictureBox1.BackColor = Color.Green
        PictureBox2.BackColor = Color.Green
        PictureBox3.BackColor = Color.Green
        prevEntry.BackColor = Color.Green
        nextEntry.BackColor = Color.Green
    End Sub
    Private Sub YellowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles YellowToolStripMenuItem.Click
        Me.BackColor = Color.Yellow
        PictureBox1.BackColor = Color.Yellow
        PictureBox2.BackColor = Color.Yellow
        PictureBox3.BackColor = Color.Yellow
        prevEntry.BackColor = Color.Yellow
        nextEntry.BackColor = Color.Yellow
    End Sub
    Private Sub OrangeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OrangeToolStripMenuItem.Click
        Me.BackColor = Color.Orange
        PictureBox1.BackColor = Color.Orange
        PictureBox2.BackColor = Color.Orange
        PictureBox3.BackColor = Color.Orange
        prevEntry.BackColor = Color.Orange
        nextEntry.BackColor = Color.Orange
    End Sub
    Private Sub RedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedToolStripMenuItem.Click
        Me.BackColor = Color.Red
        PictureBox1.BackColor = Color.Red
        PictureBox2.BackColor = Color.Red
        PictureBox3.BackColor = Color.Red
        prevEntry.BackColor = Color.Red
        nextEntry.BackColor = Color.Red
    End Sub
    Private Sub WhiteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WhiteToolStripMenuItem.Click
        Me.BackColor = Color.White
        PictureBox1.BackColor = Color.White
        PictureBox2.BackColor = Color.White
        PictureBox3.BackColor = Color.White
        prevEntry.BackColor = Color.White
        nextEntry.BackColor = Color.White
    End Sub
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        Dim result As DialogResult = SaveFileDialog1.ShowDialog
        Dim outf As String
        If result <> DialogResult.Cancel Then
            outf = SaveFileDialog1.FileName
            Using mywriter As New StreamWriter(outf)
                Dim i As Integer = 0
                Dim temp As String
                For i = 0 To numbooks - 1
                    temp = s(i).id.ToString + "|" + s(i).title + "|" + s(i).author + "|" + s(i).genre + "|" + s(i).price.ToString + "|" + s(i).num.ToString
                    mywriter.WriteLine(temp)
                Next i
                mywriter.Close()
            End Using
        End If
    End Sub

    Private Sub DelimitedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DelimitedToolStripMenuItem.Click
        Using mywriter As New StreamWriter("C:\Users\ryane\OneDrive\Documents\School\Spring 2020\CSCI_310\libdel.txt")
            Dim i As Integer = 0
            Dim temp As String
            mywriter.WriteLine(temp)
            Dim line, one, two, three As String
            For i = 0 To numbooks - 1
                one = "  " + (s(i).id).ToString
                two = s(i).title
                three = s(i).value
                temp = one + "|" + two + "|" + three
                mywriter.WriteLine(temp)
            Next i
            mywriter.Close()
        End Using
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        Application.Exit()
    End Sub
    Private Sub ArialToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ArialToolStripMenuItem.Click
        fname = "Arial"
        setfonts()
    End Sub
    Private Sub TrebuchetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TrebuchetToolStripMenuItem.Click
        fname = "Trebuchet MS"
        setfonts()
    End Sub

    Private Sub ComicSansToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ComicSansToolStripMenuItem.Click
        fname = "Comic Sans MS"
        setfonts()
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        fsize = 8
        setfonts()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        fsize = 10
        setfonts()
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        fsize = 12
        setfonts()
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        fsize = 14
        setfonts()
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        fsize = 16
        setfonts()
    End Sub

    Private Sub ToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem7.Click
        fsize = 18
        setfonts()
    End Sub

    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem8.Click
        fsize = 20
        setfonts()
    End Sub
    Private Sub GeorggiaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GeorggiaToolStripMenuItem.Click
        fname = "Georgia"
        setfonts()
    End Sub
    Private Sub BlackToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles BlackToolStripMenuItem1.Click

        Me.ForeColor = Color.Black
    End Sub
    Private Sub IndigoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles IndigoToolStripMenuItem1.Click

        Me.ForeColor = Color.Indigo
    End Sub
    Private Sub VioletToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles VioletToolStripMenuItem1.Click

        Me.ForeColor = Color.Violet
    End Sub
    Private Sub BlueToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles BlueToolStripMenuItem1.Click

        Me.ForeColor = Color.Blue
    End Sub
    Private Sub GreenToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles GreenToolStripMenuItem1.Click

        Me.ForeColor = Color.Green
    End Sub
    Private Sub YellowToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles YellowToolStripMenuItem1.Click

        Me.ForeColor = Color.Yellow
    End Sub
    Private Sub OrangeToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles OrangeToolStripMenuItem1.Click

        Me.ForeColor = Color.Orange
    End Sub
    Private Sub RedToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RedToolStripMenuItem1.Click

        Me.ForeColor = Color.Red
    End Sub

    Private Sub WhiteToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles WhiteToolStripMenuItem1.Click
        Me.ForeColor = Color.White
    End Sub

    Private Sub ResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem.Click
        Me.ForeColor = Color.Black
        Me.BackColor = Color.Gray
        fname = "Arial"
        fsize = 14
        setfonts()
    End Sub
    Private Sub findBookTitle(ftitle As String)
        For i = 0 To numbooks
            If s(i).title = ftitle Then
                showem(i)
                Return
            End If
        Next
    End Sub
    Private Sub findBookID(fid As Integer)
        For i = 0 To numbooks
            If s(i).id = fid Then
                showem(i)
                Return
            End If
        Next
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedIndex = 0 Then
            Dim temp As Integer = Convert.ToInt32(TextBox1.Text)
            findBookID(temp)
        ElseIf ListBox1.SelectedIndex = 1 Then
            Dim temp As String = TextBox1.Text
            findBookTitle(temp)
        End If
    End Sub
    Private Sub ByBookNumberToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ByBookNumberToolStripMenuItem.Click
        Try
            Dim temp As Integer = Convert.ToInt32(InputBox("Input Integer", "Find Book by ID"))
            findBookID(temp)

        Catch ex As Exception

            Dim Msg, Style, Title, Help, Ctxt, Response, MyString
            Msg = "that was invalid"
            Style = vbOK
            Title = "Invalid data"
            Help = "DEMO.HLP"
            Ctxt = 1000
            MsgBox(Msg, Style, Title)

        End Try



    End Sub

    Private Sub ByBookTitleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ByBookTitleToolStripMenuItem.Click
        Dim temp As String = InputBox("Input Title", "Find Book by Title")
        findBookTitle(temp)
    End Sub

    Private Sub ByBookAuthorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ByBookAuthorToolStripMenuItem.Click
        Try
            Dim temp As String = InputBox("Input Author", "Filter Books by Author")
            filtersetting = 1
            Dim index As Integer = 0
            Dim accum As Integer = 0
            currentfiltersetting.Text = temp
            While accum <> numbooks
                If s(accum).author = temp Then
                    tem(index).id = s(accum).id
                    tem(index).title = s(accum).title
                    tem(index).author = s(accum).author
                    tem(index).genre = s(accum).genre
                    tem(index).price = s(accum).price
                    tem(index).num = s(accum).num
                    tem(index).value = s(accum).num * s(accum).price
                    index += 1
                End If
                accum += 1
            End While
            numfiltered = index
            Label9.Text = numfiltered
        Catch ex As Exception

            Dim Msg, Style, Title, Help, Ctxt, Response, MyString
        Msg = "that was invalid"
        Style = vbOK
        Title = "Invalid data"
        Help = "DEMO.HLP"
        Ctxt = 1000
        MsgBox(Msg, Style, Title)

        End Try
    End Sub

    Private Sub InsertToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InsertToolStripMenuItem.Click
        Try
            Dim myDialog As New Form2
            numbooks += 1
            myDialog.ShowDialog()
            totalAccum += s(numbooks).num
            showem(numbooks)
        Catch ex As Exception
            Dim Msg, Style, Title, Help, Ctxt, Response, MyString
            Msg = "that was invalid"
            Style = vbOK
            Title = "Invalid data"
            Help = "DEMO.HLP"
            Ctxt = 1000
            MsgBox(Msg, Style, Title)
        End Try
    End Sub

    Private Sub UpdateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpdateToolStripMenuItem.Click
        Dim myDialog As New Form3
        myDialog.ShowDialog()
        Dim index As Integer = 0
        Dim accum As Integer = 0
        Dim idCheck As Integer
        Try
            idCheck = myDialog.TextBox1.Text
            While accum <> numbooks
                If s(accum).id = idCheck Then
                    If (s(accum).num <= myDialog.numtot) Then
                    End If
                    totalAccum -= s(accum).num
                    s(accum).num += myDialog.numtot
                    If s(accum).num <= 0 Then
                        ifZero(accum)
                    Else
                        totalAccum += s(accum).num
                        Exit While
                    End If
                End If
                    accum += 1
            End While
            showem(accum)
        Catch ex As Exception

            Dim Msg, Style, Title, Help, Ctxt, Response, MyString
            Msg = "that was invalid"
            Style = vbOK
            Title = "Invalid data"
            Help = "DEMO.HLP"
            Ctxt = 1000
            MsgBox(Msg, Style, Title)

        End Try
    End Sub

    Private Sub ColorsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ColorsToolStripMenuItem.Click

    End Sub

    Private Sub StyleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StyleToolStripMenuItem.Click

    End Sub

    Private Sub ByGenreToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ByGenreToolStripMenuItem.Click
        Dim temp As String = InputBox("Input Genre", "Filter Books by Genre")
        Try
            filtersetting = 1
            currentfiltersetting.Text = temp
            Dim index As Integer = 0
            Dim accum As Integer = 0
            currentfiltersetting.Text = temp
            While accum <> numbooks
                If s(accum).genre = temp Then
                    tem(index).id = s(accum).id
                    tem(index).title = s(accum).title
                    tem(index).author = s(accum).author
                    tem(index).genre = s(accum).genre
                    tem(index).price = s(accum).price
                    tem(index).num = s(accum).num
                    tem(index).value = s(accum).num * s(accum).price
                    index += 1
                End If
                accum += 1
            End While
            numfiltered = index
            Label9.Text = numfiltered
        Catch ex As Exception

            Dim Msg, Style, Title, Help, Ctxt, Response, MyString
            Msg = "that was invalid"
            Style = vbOK
            Title = "Invalid data"
            Help = "DEMO.HLP"
            Ctxt = 1000
            MsgBox(Msg, Style, Title)

        End Try
    End Sub
    Private Sub ifZero(indexrem As Integer)
        For i = 0 To numbooks - 1
            If (i <> indexrem) Then
                temp(i) = s(i)
            End If
            s = temp
        Next
    End Sub

    Private Sub ResetToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem1.Click
        currentfiltersetting.Text = "None"
        filtersetting = 0
    End Sub
End Class
