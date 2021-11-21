Imports System.Environment

Public Class HelpForm

    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel1.Click
        RichTextBox1.Clear()
        RichTextBox1.AppendText("CORRECT FORMATTING OF THE TEXT" & vbLf)
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.SelectionStart = RichTextBox1.Find("CORRECT FORMATTING OF THE TEXT")
        Dim bfont As New System.Drawing.Font(RichTextBox1.Font, FontStyle.Bold)
        RichTextBox1.SelectionFont = bfont
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox1.AppendText("The best solution for the correct formatting of text in multiple lines is to use a colon at the end of sentences." & vbLf)
        RichTextBox1.AppendText("Language Translations should be separated by a slash." & vbLf)
        RichTextBox1.AppendText("Whereas the numerical values should be inserted after the pause." & vbLf)
        RichTextBox1.AppendText("A value of units must be in square brackets." & vbLf)
        RichTextBox1.AppendText("Full signature of the device should be split up  a character '|'." & vbLf)
      
        Dim startingPoint As Integer = -1
        Dim sum As Integer
        Do
            startingPoint = RichTextBox1.Find("The best", startingPoint + 1, RichTextBoxFinds.None)
            If (startingPoint >= 0) Then
                RichTextBox1.SelectionStart = startingPoint
                For start As Integer = 3 To 7
                    Dim licz As Integer = Len(RichTextBox1.Lines(start))
                    sum = licz + sum
                Next
                RichTextBox1.SelectionLength = sum + 4
                'RichTextBox1.SelectionLength = "Full signature of the device should be split up  a character '|'.".Length
                RichTextBox1.SelectionColor = System.Drawing.Color.Blue
                RichTextBox1.SelectionFont = New System.Drawing.Font("Tahoma", 10.0F)
            End If
        Loop Until startingPoint < 0

        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText("DANE TECHNICZNE / TECHNISCHE DATEN / TECHNICAL DATA:" & vbLf) '& Environment.NewLine& vbLf
        RichTextBox1.AppendText("Motoreduktor / Stirnradgetriebemotor / Gearmotor:" & vbLf)
        RichTextBox1.AppendText("R97DRE100LC4BE5HR|IS|TF|AS7W;" & vbCrLf)
        RichTextBox1.AppendText("Moc znamionowa / Nennleistung / Rated Power  -  3 [kW];" & vbCrLf)
        RichTextBox1.AppendText("Prêdkoœæ wyjœciowa / Abtriebsdrehzahl / Output speed  -  20 [1/min];" & vbCrLf)
        RichTextBox1.AppendText("Moment wyjœciowy / Abtriebsmoment / Output torque  -  1420 [Nm];" & vbCrLf)
        RichTextBox1.AppendText("Prze³o¿enie / Getriebeuebersetzung / Gear ratio  - 72,17;" & vbCrLf)
        RichTextBox1.AppendText("Napiêcie znamionowe / Spannung / Voltage  -  400 [V];" & vbCrLf)
        RichTextBox1.AppendText("Czêstotliwoœæ / Frequenz / Frequency  -  50 [Hz];" & vbCrLf)
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText("A short text has  to be according to a schema. " & vbLf)
        RichTextBox1.SelectionStart = RichTextBox1.Find("A short text has  to be according to a schema.")
        Dim bfonts As New System.Drawing.Font(RichTextBox1.Font, FontStyle.Bold)
        RichTextBox1.SelectionFont = bfonts
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Left
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText("Motoreduktor / Stirnradgetriebemotor / Gearmotor" & vbLf)
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText("A multi notes if you have a lot of a place it better to split on two separate." & vbLf)
    End Sub

    Private Sub HelpForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        RichTextBox1.Clear()
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText("THE MOST IMPORTANT INFORMATION" & vbLf) '& Environment.NewLine& vbLf
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.SelectionStart = RichTextBox1.Find("THE MOST IMPORTANT INFORMATION")
        Dim bfont As New System.Drawing.Font(RichTextBox1.Font, FontStyle.Bold)
        RichTextBox1.SelectionFont = bfont
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Center
    End Sub

    Private Sub ToolStripLabel3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel3.Click
        ' Zamkniêcie formularza
        Me.Close()
    End Sub



    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click
        RichTextBox1.Clear()
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText("CORRECT FORMATTING TEXT OF THE PARTLIST " & vbLf)
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.SelectionStart = RichTextBox1.Find("CORRECT FORMATTING TEXT OF THE PARTLIST ")
        Dim bfont As New System.Drawing.Font(RichTextBox1.Font, FontStyle.Bold)
        RichTextBox1.SelectionFont = bfont
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText("A short text has  to be according to a schema. " & vbLf)
        RichTextBox1.SelectionStart = RichTextBox1.Find("A short text has  to be according to a schema.")
        Dim bfonts As New System.Drawing.Font(RichTextBox1.Font, FontStyle.Bold)
        RichTextBox1.SelectionFont = bfonts
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Left
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText("Motoreduktor / Stirnradgetriebemotor / Gearmotor" & vbLf)
        RichTextBox1.SelectionStart = RichTextBox1.Find("Motoreduktor / Stirnradgetriebemotor / Gearmotor")
        RichTextBox1.SelectionColor = System.Drawing.Color.Blue
        RichTextBox1.SelectionFont = New System.Drawing.Font("Tahoma", 10.0F)
        RichTextBox1.AppendText(vbLf)
        RichTextBox1.AppendText("What else?.Language Translations should be separated by a slash." & vbLf)
        RichTextBox1.AppendText("That is it" & vbLf)
    End Sub

End Class