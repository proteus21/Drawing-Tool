Public Class Translation_text

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Translation_Message.ListBox2.Items.Add(TextBox2.Text)
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.TextBox2.Clear()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Me.Width = 775 Then
            Me.Width = 439
            Button3.Text = ">>"
        Else '
            Me.Width = 775
            Button3.Text = "<<"
        End If

    End Sub

    Private Sub Translation_text_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label5.Text = Translation.ComboBox1.Text & " /" & Translation.ComboBox2.Text & " /" & Translation.ComboBox3.Text
        'MsgBox(Translation.ComboBox1.Text)
    End Sub
    ' Funkcja do wyszukiwania teksu w richbox i tekstbox
    Public Overloads Function SearchText(ByVal textToFind As String, Optional ByVal startPosition As Integer = 0, Optional ByVal endPosition As Integer = 0, Optional ByVal highlightText As Boolean = True, Optional ByVal matchCase As Boolean = False) As Integer
        '
        'Contains the return value of the search. IT it returns -1, then a match was not found.
        Dim i As Integer

        If endPosition < 1 Then

            If Not matchCase Then

                textToFind = textToFind.ToLower

                Dim temp As String = TextBox1.Text.ToLower

                i = temp.IndexOf(textToFind, startPosition, Me.Text.Length)

            Else

                i = TextBox1.Text.IndexOf(textToFind, startPosition, Me.Text.Length)

            End If

        Else

            If matchCase = False Then

                textToFind = textToFind.ToLower

                Dim temp As String = TextBox1.Text.ToLower

                i = temp.IndexOf(textToFind, startPosition, endPosition)

            Else

                i = TextBox1.Text.IndexOf(textToFind, startPosition, endPosition)

            End If

        End If

        If i > -1 Then

            If highlightText Then

                TextBox1.Focus()

                TextBox1.SelectionStart = i

                TextBox1.SelectionLength = textToFind.Length

            End If

        End If
        '
        'Returns the position the text was found at, otherwise it will report -1, which means that the search string was not found.
        Return i

    End Function

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.TextBox3.Clear()
    End Sub

 
End Class