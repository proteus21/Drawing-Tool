Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.IO
Imports Excel
Imports Excel.ExcelReaderFactory
Imports ICSharpCode.SharpZipLib

Public Class ConvertCSV

    Partial Public Class Form1
        Inherits Form
        'Public result As New DataSet()
        Public Sub New()
            InitializeComponent()
        End Sub




        'Private Sub getExcelData(ByVal file__1 As String)

        '    If file__1.EndsWith(".xlsx") Then
        '        ' Reading from a binary Excel file (format; *.xlsx)
        '        Dim stream As FileStream = File.Open(file__1, FileMode.Open, FileAccess.Read)
        '        Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream)
        '        result = excelReader.AsDataSet()
        '        excelReader.Close()
        '    End If

        '    If file__1.EndsWith(".xls") Then
        '        ' Reading from a binary Excel file ('97-2003 format; *.xls)
        '        Dim stream As FileStream = File.Open(file__1, FileMode.Open, FileAccess.Read)
        '        Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateBinaryReader(stream)
        '        result = excelReader.AsDataSet()
        '        excelReader.Close()
        '    End If

        '    Dim items As New List(Of String)()
        '    For i As Integer = 0 To result.Tables.Count - 1
        '        items.Add(result.Tables(i).TableName.ToString())
        '    Next
        '    ComboBox1.DataSource = items

        'End Sub

        'Sub converToCSV(ByVal ind As Integer)
        '    ' sheets in excel file becomes tables in dataset
        '    'result.Tables[0].TableName.ToString(); // to get sheet name (table name)

        '    Dim a As String = ""
        '    Dim row_no As Integer = 0

        '    While row_no < result.Tables(ind).Rows.Count
        '        For i As Integer = 0 To result.Tables(ind).Columns.Count - 1
        '            a += result.Tables(ind).Rows(row_no)(i).ToString() & ","
        '        Next
        '        row_no += 1
        '        a += vbLf
        '    End While
        '    Dim output As String = TextBox2.Text + "\" + TextBox3.Text & ".csv"
        '    Dim csv As New IO.StreamWriter(output, False)
        '    csv.Write(a)
        '    csv.Close()

        '    MessageBox.Show("File converted succussfully")

        '    TextBox1.Text = ""
        '    TextBox2.Text = ""
        '    TextBox3.Text = ""
        '    ComboBox1.DataSource = Nothing
        '    Return
        'End Sub

        Private Sub InitializeComponent()
            Me.Button1 = New System.Windows.Forms.Button
            Me.Button2 = New System.Windows.Forms.Button
            Me.TextBox1 = New System.Windows.Forms.TextBox
            Me.TextBox2 = New System.Windows.Forms.TextBox
            Me.TextBox3 = New System.Windows.Forms.TextBox
            Me.ComboBox1 = New System.Windows.Forms.ComboBox
            Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
            Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
            Me.SuspendLayout()
            '
            'Button1
            '
            Me.Button1.Location = New System.Drawing.Point(373, 66)
            Me.Button1.Name = "Button1"
            Me.Button1.Size = New System.Drawing.Size(80, 24)
            Me.Button1.TabIndex = 0
            Me.Button1.Text = "Button1"
            Me.Button1.UseVisualStyleBackColor = True
            '
            'Button2
            '
            Me.Button2.Location = New System.Drawing.Point(377, 151)
            Me.Button2.Name = "Button2"
            Me.Button2.Size = New System.Drawing.Size(75, 26)
            Me.Button2.TabIndex = 1
            Me.Button2.Text = "Button2"
            Me.Button2.UseVisualStyleBackColor = True
            '
            'TextBox1
            '
            Me.TextBox1.Location = New System.Drawing.Point(181, 64)
            Me.TextBox1.Name = "TextBox1"
            Me.TextBox1.Size = New System.Drawing.Size(161, 20)
            Me.TextBox1.TabIndex = 2
            '
            'TextBox2
            '
            Me.TextBox2.Location = New System.Drawing.Point(188, 148)
            Me.TextBox2.Name = "TextBox2"
            Me.TextBox2.Size = New System.Drawing.Size(143, 20)
            Me.TextBox2.TabIndex = 3
            '
            'TextBox3
            '
            Me.TextBox3.Location = New System.Drawing.Point(192, 241)
            Me.TextBox3.Name = "TextBox3"
            Me.TextBox3.Size = New System.Drawing.Size(149, 20)
            Me.TextBox3.TabIndex = 4
            '
            'ComboBox1
            '
            Me.ComboBox1.FormattingEnabled = True
            Me.ComboBox1.Location = New System.Drawing.Point(188, 108)
            Me.ComboBox1.Name = "ComboBox1"
            Me.ComboBox1.Size = New System.Drawing.Size(117, 21)
            Me.ComboBox1.TabIndex = 5
            '
            'OpenFileDialog1
            '
            Me.OpenFileDialog1.FileName = ""
            '
            'Form1
            '
            Me.ClientSize = New System.Drawing.Size(494, 356)
            Me.Controls.Add(Me.ComboBox1)
            Me.Controls.Add(Me.TextBox3)
            Me.Controls.Add(Me.TextBox2)
            Me.Controls.Add(Me.TextBox1)
            Me.Controls.Add(Me.Button2)
            Me.Controls.Add(Me.Button1)
            Me.Name = "Form1"
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents Button1 As System.Windows.Forms.Button
        Friend WithEvents Button2 As System.Windows.Forms.Button
        Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
        Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
        Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
        Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
        Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
        Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    End Class






    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim result As DialogResult = Me.FolderBrowserDialog1.ShowDialog()
        Dim foldername As String = ""
        If result = DialogResult.OK Then
            foldername = Me.FolderBrowserDialog1.SelectedPath
        End If

        TextBox2.Text = foldername
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim fileName As String = ""
        fileName = TextBox3.Text

        If fileName = "" Then
            MessageBox.Show("Enter Valid file name")
            Return
        End If

        Dim a As String = ""
        Dim row_no As Integer = 0
        Dim ind As Integer = ComboBox1.SelectedIndex
        While row_no < result.Tables(ind).Rows.Count
            For i As Integer = 0 To result.Tables(ind).Columns.Count - 1
                a += result.Tables(ind).Rows(row_no)(i).ToString() & ","
            Next
            row_no += 1
            a += vbLf
        End While
        Dim output As String = TextBox2.Text + "\" + TextBox3.Text & ".csv"
        Dim csv As New IO.StreamWriter(output, False)
        csv.Write(a)
        csv.Close()

        MessageBox.Show("File converted succussfully")

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.DataSource = Nothing

    End Sub
    Public result As New DataSet()
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Chosen_File As String
        With OpenFileDialog1
            .FileName = ""
            .InitialDirectory = "C:\"
            .Title = "Select Excel file"
            .DefaultExt = ".xls"
            .Filter = "Excel File  (*.xls) | *.xls"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

                Chosen_File = OpenFileDialog1.FileName
            End If
        End With
        
        If Chosen_File = String.Empty Then
            Return
        End If
        TextBox1.Text = Chosen_File
        Dim file__1 As String = TextBox1.Text
        'getExcelData(TextBox1.Text)
        If TextBox1.Text.EndsWith(".xlsx") Then
            ' Reading from a binary Excel file (format; *.xlsx)
            Dim stream As FileStream = File.Open(TextBox1.Text, FileMode.Open, FileAccess.Read)
            Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream)
            result = excelReader.AsDataSet()
            excelReader.Close()
        End If

        If TextBox1.Text.EndsWith(".xls") Then
            ' Reading from a binary Excel file ('97-2003 format; *.xls)
            Dim stream As FileStream = File.Open(TextBox1.Text, FileMode.Open, FileAccess.Read)
            Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateBinaryReader(stream)
            result = excelReader.AsDataSet()
            excelReader.Close()
        End If

        Dim items As New List(Of String)()
        For i As Integer = 0 To result.Tables.Count - 1
            items.Add(result.Tables(i).TableName.ToString())
        Next
        ComboBox1.DataSource = items

    End Sub

End Class