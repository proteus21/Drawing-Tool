Imports System.Text
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.Data.SQLite
Imports System.ComponentModel
Imports System.Drawing

Imports System.Windows.Forms
Public Class DataBaseSet
    Public DirectoryW As String = My.Application.Info.DirectoryPath
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        With OpenFileDialog1
            .InitialDirectory = "C:\"
            .Title = "Select Excel file"
            .DefaultExt = ".idw"
            .Filter = "Excel CSV File  (*.csv) | *.csv"
            .ShowDialog()
        End With
        If OpenFileDialog1.FileName <> "" Then

            Dim dt As New Data.DataTable()
            Dim line As String = Nothing
            Dim i As Integer = 0
            Dim ascii As System.Text.Encoding = Encoding.ASCII
            Dim unicode As System.Text.Encoding = Encoding.Unicode
            ' Using sr As New IO.StreamReader("C:\Documents and Settings\bkonefal\Moje dokumenty\Baza\g.csv", System.Text.UTF8Encoding.UTF8)
            Using sr As New IO.StreamReader(OpenFileDialog1.FileName, System.Text.UTF8Encoding.UTF8)
                ' Perform the conversion from one encoding to the other.
                line = sr.ReadLine()
                Do While line IsNot Nothing
                    Dim data() As String = line.Split(","c)
                    If data.Length > 0 Then
                        If i = 0 Then
                            Dim item
                            For Each item In data
                                dt.Columns.Add(New DataColumn())
                            Next item
                            i += 1
                        End If
                        Dim row As DataRow = dt.NewRow()
                        row.ItemArray = data
                        dt.Rows.Add(row)
                    End If
                    line = sr.ReadLine()
                Loop
            End Using

            'Using cn As New SqlClient.SqlConnection(ConfigurationManager.ConnectionStrings("ConsoleApplication3.Properties.Settings.daasConnectionString").ConnectionString)
            Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
                cn.Open()
                Dim SQLcommand As SQLite.SQLiteCommand
                SQLcommand = cn.CreateCommand
                For Each dRow As DataRow In dt.Rows

                    'Insert Record into TranslateBase
                    SQLcommand.CommandText = "INSERT INTO TranslateBase(PL,DE,EN,HU,RU,CZ,SLOV) VALUES ('" + dRow(0) + "','" + dRow(1) + "','" + dRow(2) + "','" + dRow(3) + "','" + dRow(4) + "','" + dRow(5) + "','" + dRow(6) + "')"
                    ''Update Last Created Record in Foo
                    ' SQLcommand.CommandText = "UPDATE TranslateBase SET LN = 'New LN', PL = 'New PL' WHERE id = last_insert_rowid()"
                    ' SQLcommand.CommandText = "INSERT INTO TranslateBase (PL,DE,EN,HU,RUS) VALUES ('" + dRow(0) + " ','" + dRow(1) + "','" + dRow(2) + "','" + dRow(3) + "','" + dRow(4) + "')"
                    'SQLcommand.CommandText = "INSERT INTO TranslateBase(PL,DE,EN,HU,RUS,CZ,SLOV) VALUES ('This is a title','" + dRow(0) + "','','','','','')"

                    SQLcommand.ExecuteNonQuery()
                Next

                SQLcommand.Dispose()
                cn.Close()
            End Using
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim dt As New Data.DataTable()
        'Using cn As New SqlClient.SqlConnection(ConfigurationManager.ConnectionStrings("ConsoleApplication3.Properties.Settings.daasConnectionString").ConnectionString)
        Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
            cn.Open()
            Dim SQLcommand As SQLite.SQLiteCommand
            SQLcommand = cn.CreateCommand

            ''Delete Last Created Record from TranslateBase
            SQLcommand.CommandText = "DELETE FROM TranslateBase"
            SQLcommand.ExecuteNonQuery()
            SQLcommand.Dispose()
            cn.Close()
        End Using
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ConvertCSV.Show()
    End Sub



    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TabControl1.SelectedIndex = 3
        Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
            cn.Open()
            Dim SQLcommand As SQLite.SQLiteCommand
            SQLcommand = cn.CreateCommand
            'Insert Record into TranslateBase
            If RadioButton1.Checked = True Then
                ' sortowanie w górê
                SQLcommand.CommandText = "Select* From TranslateBase ORDER BY PL ASC"
            Else
                ' sortowanie w dó³
                SQLcommand.CommandText = "Select* From TranslateBase ORDER BY PL DESC"
            End If

            Dim lrd As IDataReader = SQLcommand.ExecuteReader()
            ' Dim SQLreader As System.Data.SqlClient.SqlDataReader = SQLcommand.ExecuteReader()
            DataGridView3.Rows.Clear()
            'Tracks the current record number
            Dim n As Integer = 0
            Dim i As Integer
            While lrd.Read()


                n = DataGridView3.Rows.Add()
                DataGridView3.Rows.Item(n).Cells(0).Value = lrd.Item(0).ToString
                DataGridView3.Rows.Item(n).Cells(1).Value = lrd.Item(1).ToString
                DataGridView3.Rows.Item(n).Cells(2).Value = lrd.Item(2).ToString
                DataGridView3.Rows.Item(n).Cells(3).Value = lrd.Item(3).ToString
                DataGridView3.Rows.Item(n).Cells(4).Value = lrd.Item(4).ToString
                DataGridView3.Rows.Item(n).Cells(5).Value = lrd.Item(5).ToString
                DataGridView3.Rows.Item(n).Cells(6).Value = lrd.Item(6).ToString
                DataGridView3.Rows.Item(n).Cells(7).Value = lrd.Item(7).ToString



            End While

            'SQLcommand.ExecuteNonQuery()
            SQLcommand.Dispose()
            'Next
            cn.Close()
            'SQLcommand.ExecuteNonQuery()
            SQLcommand.Dispose()
        End Using
    End Sub

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        'MsgBox(TabControl1.TabCount)
        If TabControl1.SelectedIndex = 2 Or TabControl1.SelectedIndex = 1 Then

            Dim i As Int32              'General counter
            Dim Row As DataRow
            ' ToolStripStatusLabel1.Text = "Connecting..."
            'Open a connection
            Dim Con As New SQLite.SQLiteConnection
            Con.ConnectionString = "Data Source=" & DirectoryW & "\TranslateBase.s3db;"
            Con.Open()
            ' Sort data 
            Try
                'Fill a DataSet object
                Dim Adapter As New SQLite.SQLiteDataAdapter("SELECT * FROM TranslateBase", Con)
                Dim ds As New DataSet
                Adapter.FillSchema(ds, SchemaType.Mapped, "TranslateBase")
                Adapter.Fill(ds, "TranslateBase")
                ' Assign the DataSet as the DataSource for the BindingSource.

                BindingSource1.DataSource = ds.Tables("TranslateBase")
                'BindingNavigator1.BindingSource = BindingSource1
                TextBox1.DataBindings.Clear()
                TextBox2.DataBindings.Clear()
                TextBox3.DataBindings.Clear()
                TextBox4.DataBindings.Clear()
                TextBox5.DataBindings.Clear()
                TextBox6.DataBindings.Clear()
                TextBox7.DataBindings.Clear()
                TextBox8.DataBindings.Clear()


                TextBox1.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "LN"))
                TextBox2.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "PL"))
                TextBox3.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "DE"))
                TextBox4.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "EN"))
                TextBox5.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "HU"))
                TextBox6.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "RU"))
                TextBox7.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "CZ"))
                TextBox8.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "SLOV"))
            Finally
                Con.Dispose()
                'Close the connection
                Con.Close()
            End Try
        End If
    End Sub
    Public ds As New DataSet("TranslateBase")
    Public Con As System.Data.SQLite.SQLiteConnection
    Public objtabela As DataTable
    Public objwiersz As DataRow
    Public nr_rekordu As Integer = 0
    'Public wiersz As Integer
    Private Sub odczyt_do_edycji(ByVal nr_rekordu As Integer)
        'MsgBox(TabControl1.TabCount)
        If TabControl1.TabIndex = 2 Then

            Dim i As Int32              'General counter
            Dim Row As DataRow
            ' ToolStripStatusLabel1.Text = "Connecting..."
            'Open a connection
            Dim Con As New SQLite.SQLiteConnection
            Con.ConnectionString = "Data Source=" & DirectoryW & "\TranslateBase.s3db;"
            Con.Open()
            ' Sort data 
            Try
                'Fill a DataSet object
                Dim Adapter As New SQLite.SQLiteDataAdapter("SELECT * FROM TranslateBase", Con)
                Dim ds As New DataSet
                Adapter.FillSchema(ds, SchemaType.Mapped, "TranslateBase")
                Adapter.Fill(ds, "TranslateBase")
                ' Assign the DataSet as the DataSource for the BindingSource.

                BindingSource1.DataSource = ds.Tables("TranslateBase")
                objtabela = ds.Tables("TranslateBase")

            Finally
                ' Con.Dispose()
                'Close the connection
                Con.Close()
            End Try

        End If
        If objtabela.Rows.Count = 0 Then
            wyswietl_rekord(-1)
            Exit Sub
        Else
            wyswietl_rekord(nr_rekordu)
        End If
    End Sub

    Private Sub wyswietl_rekord(ByVal wiersz As Integer)
        If wiersz >= 0 Then
            'Dim objtabela As DataTable
            ' Dim objwiersz As DataRow

            objwiersz = objtabela.Rows(wiersz)

            ''BindingNavigator1.BindingSource = BindingSource1
            'TextBox1.DataBindings.Clear()
            'TextBox2.DataBindings.Clear()
            'TextBox3.DataBindings.Clear()
            'TextBox4.DataBindings.Clear()
            'TextBox5.DataBindings.Clear()
            'TextBox6.DataBindings.Clear()
            'TextBox7.DataBindings.Clear()
            'TextBox8.DataBindings.Clear()

            TextBox1.Text = objwiersz.Item(0)
            TextBox2.Text = objwiersz.Item(1)
            TextBox3.Text = objwiersz.Item(2)
            TextBox4.Text = objwiersz.Item(3)
            TextBox5.Text = objwiersz.Item(4)
            TextBox6.Text = objwiersz.Item(5)
            TextBox7.Text = objwiersz.Item(6)
            TextBox8.Text = objwiersz.Item(7)
            'TextBox1.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "LN"))
            'TextBox2.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "PL"))
            'TextBox3.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "DE"))
            'TextBox4.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "EN"))
            'TextBox5.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "HU"))
            'TextBox6.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "RU"))
            'TextBox7.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "CZ"))
            'TextBox8.DataBindings.Add(New Binding("text", BindingSource1.DataSource, "SLOV"))


        End If
    End Sub


    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        ' uwaga
        odczyt_do_edycji(nr_rekordu)

        If nr_rekordu < objtabela.Rows.Count - 1 Then
            nr_rekordu += 1
            wyswietl_rekord(nr_rekordu)
        Else
            wyswietl_rekord(nr_rekordu)
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        'odczyt_do_edycji(0)

        If nr_rekordu > 0 Then
            nr_rekordu -= 1
            wyswietl_rekord(nr_rekordu)
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        odczyt_do_edycji(0)
        If nr_rekordu >= 0 Then
            nr_rekordu = 0
            wyswietl_rekord(nr_rekordu)
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        odczyt_do_edycji(0)
        If nr_rekordu >= 0 Then
            nr_rekordu = objtabela.Rows.Count - 1
            wyswietl_rekord(nr_rekordu)
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        connect_to_base()
        If Con.State = ConnectionState.Open Then
            If ile_rekordow() > 0 Then
                Dim pytanie As DialogResult
                pytanie = MessageBox.Show("Do you want delate current records?", "Delate data", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
                If pytanie = System.Windows.Forms.DialogResult.Yes Then
                    Dim strSqlite As String
                    strSqlite = "delete from TranslateBase " & "where LN=" & TextBox1.Text
                    Try
                        Dim objzapytanie As New System.Data.SQLite.SQLiteCommand(strSqlite, Con)
                        objzapytanie.ExecuteNonQuery()
                        MessageBox.Show("NO connection to the base", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        odczyt_do_edycji(0)
                    Catch ex As Exception
                        MessageBox.Show("Error while data were updating", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Finally
                        Con.Close()
                    End Try
                End If
            End If
        End If
    End Sub
    Private Sub connect_to_base()

        ' ToolStripStatusLabel1.Text = "Connecting..."
        'Open a connection
        Try
            Con = New System.Data.SQLite.SQLiteConnection
            Con.ConnectionString = "Data Source=" & DirectoryW & "\TranslateBase.s3db;"
            Con.Open()
        Catch ex As Exception
            MessageBox.Show("Error while data were updating", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ' Sort data 

        'Fill a DataSet object
        'Dim Adapter As New SQLite.SQLiteDataAdapter("SELECT * FROM TranslateBase", Con)
        'Dim ds As New DataSet
        'Adapter.FillSchema(ds, SchemaType.Mapped, "TranslateBase")
        'Adapter.Fill(ds, "TranslateBase")

        'Close the connection
        ' Con.Close()
    End Sub
    Private Function ile_rekordow() As Integer
        Dim strSqlite As String
        ds.Clear()
        strSqlite = "select*from TranslateBase"
        Dim objzapytanie As New System.Data.SQLite.SQLiteDataAdapter(strSqlite, Con)
        If Con.State = ConnectionState.Open Then
            Try
                objzapytanie.Fill(ds, "TranslateBase")
                Return ds.Tables("TranslateBase").Rows.Count
            Catch ex As Exception
                MessageBox.Show("Can't read a base", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Function

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        ' edycja 
        Dim StrWpis As String
        Dim StrSqlite As String
        If TextBox2.Text.Length = 0 Then
            MessageBox.Show("Fill in data missing", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox2.Focus()
            Exit Sub
        End If
        If TextBox3.Text.Length = 0 Then
            MessageBox.Show("Fill in data missing", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox3.Focus()
            Exit Sub
        End If
        If TextBox4.Text.Length = 0 Then
            MessageBox.Show("Fill in data missing", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Focus()
            Exit Sub
        End If
        If TextBox5.Text.Length = 0 Then
            MessageBox.Show("Fill in data missing", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox5.Focus()
            Exit Sub
        End If
        If TextBox6.Text.Length = 0 Then
            MessageBox.Show("Fill in data missing", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox6.Focus()
            Exit Sub
        End If

        If TextBox7.Text.Length = 0 Then
            MessageBox.Show("Fill in data missing", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox7.Focus()
            Exit Sub
        End If
        If TextBox8.Text.Length = 0 Then
            MessageBox.Show("Fill in data missing", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox8.Focus()
            Exit Sub
        End If
        connect_to_base()

        StrSqlite = "insert into Translatebase (LN,PL,DE,EN,HU,RU,CZ,SLOV) values(" & TextBox1.Text & "," & "'" & TextBox2.Text & "'" & "," & "'" & TextBox3.Text & "'" & "," & "'" & TextBox4.Text & "'" & "," & "'" & TextBox5.Text & "'" & "," & "'" & TextBox6.Text & "'" & "," & "'" & TextBox7.Text & "'" & "," & "'" & TextBox8.Text & "'" & ")"
        If Con.State = ConnectionState.Open Then
            Try
                Dim objzapytanie As New System.Data.SQLite.SQLiteCommand(StrSqlite, Con)
                objzapytanie.ExecuteNonQuery()
                odczyt_do_edycji(nr_rekordu)
                MessageBox.Show("Record was updated", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As DataException
                MessageBox.Show("Error while data were writing", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Con.Close()
            End Try
        End If
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        ' open and select tabpage 3
        TabControl1.SelectedIndex = 3
        ' open sql data base

        Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
            cn.Open()
            Dim SQLcommand As New SQLite.SQLiteCommand

            ' For Each dRow As DataRow In dt.Rows
            'Insert Record into TranslateBase


            SQLcommand = cn.CreateCommand

            '  Dim dt As New Data.DataTable()
            SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL like '" & TextBox9.Text & "'"
            'SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL='" + defDescription(0).ToString + "' "
            'SQLcommand.CommandText = "SELECT PL,DE FROM TranslateBase"
            Dim lrd As IDataReader = SQLcommand.ExecuteReader()
            ' Dim SQLreader As System.Data.SqlClient.SqlDataReader = SQLcommand.ExecuteReader()
            DataGridView3.Rows.Clear()
            'Tracks the current record number
            Dim n As Integer = 0
            Dim i As Integer
            While lrd.Read()

                'MsgBox(lrd.Item(1).ToString)
                n = DataGridView3.Rows.Add()
                DataGridView3.Rows.Item(n).Cells(0).Value = lrd.Item(0).ToString
                DataGridView3.Rows.Item(n).Cells(1).Value = lrd.Item(1).ToString
                DataGridView3.Rows.Item(n).Cells(2).Value = lrd.Item(2).ToString
                DataGridView3.Rows.Item(n).Cells(3).Value = lrd.Item(3).ToString
                DataGridView3.Rows.Item(n).Cells(4).Value = lrd.Item(4).ToString
                DataGridView3.Rows.Item(n).Cells(5).Value = lrd.Item(5).ToString
                DataGridView3.Rows.Item(n).Cells(6).Value = lrd.Item(6).ToString
                DataGridView3.Rows.Item(n).Cells(7).Value = lrd.Item(7).ToString



            End While

            'SQLcommand.ExecuteNonQuery()
            SQLcommand.Dispose()
            'Next
            cn.Close()
        End Using
        ToolStripStatusLabel2.Text = DataGridView3.RowCount - 1
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        Dim i As Int32              'General counter
        Dim Row As DataRow
        ' ToolStripStatusLabel1.Text = "Connecting..."
        'Open a connection
        Dim Con As New SQLite.SQLiteConnection
        Con.ConnectionString = "Data Source=" & DirectoryW & "\TranslateBase.s3db;"
        Con.Open()
        ' Sort data 

        'Fill a DataSet object
        Dim Adapter As New SQLite.SQLiteDataAdapter("SELECT * FROM TranslateBase", Con)
        Dim ds As New DataSet
        Adapter.FillSchema(ds, SchemaType.Mapped, "TranslateBase")
        Adapter.Fill(ds, "TranslateBase")

        'Display some data
        Debug.Print("Total Rows: " & ds.Tables("TranslateBase").Rows.Count)
        DataGridView1.Rows.Clear()
        i = 1           'Tracks the current record number
        For Each Row In ds.Tables("TranslateBase").Rows
            Dim n As Integer = DataGridView1.Rows.Add()
            DataGridView1.Rows.Item(n).Cells(0).Value = Row("LN")
            DataGridView1.Rows.Item(n).Cells(1).Value = Row("PL")
            DataGridView1.Rows.Item(n).Cells(2).Value = Row("DE")
            DataGridView1.Rows.Item(n).Cells(3).Value = Row("EN")
            DataGridView1.Rows.Item(n).Cells(4).Value = Row("HU")
            DataGridView1.Rows.Item(n).Cells(5).Value = Row("RU")
            DataGridView1.Rows.Item(n).Cells(6).Value = Row("CZ")
            DataGridView1.Rows.Item(n).Cells(7).Value = Row("SLOV")

            n = DataGridView2.Rows.Add()
            DataGridView2.Rows.Item(n).Cells(0).Value = Row("LN")
            DataGridView2.Rows.Item(n).Cells(1).Value = Row("PL")
            DataGridView2.Rows.Item(n).Cells(2).Value = Row("DE")
            DataGridView2.Rows.Item(n).Cells(3).Value = Row("EN")
            DataGridView2.Rows.Item(n).Cells(4).Value = Row("HU")
            DataGridView2.Rows.Item(n).Cells(5).Value = Row("RU")
            DataGridView2.Rows.Item(n).Cells(6).Value = Row("CZ")
            DataGridView2.Rows.Item(n).Cells(7).Value = Row("SLOV")
            'MsgBox(i & " Name: " & Row("LN") & " City: " & Row("PL") & " State: " & Row("ENG"))
            i += 1
        Next

        'Close the connection
        Con.Close()
        ' Add any initialization after the InitializeComponent() call.
        ToolStripStatusLabel4.Text = DataGridView1.RowCount - 1
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Dim StrWpis As String
            Dim StrSqlite As String
            connect_to_base()
            ' find all records which are empty
            StrSqlite = "SELECT * FROM TranslateBase where PL like ''"
            If Con.State = ConnectionState.Open Then
                Try
                    Dim objzapytanie As New System.Data.SQLite.SQLiteCommand(StrSqlite, Con)

                    Dim lrd As IDataReader = objzapytanie.ExecuteReader()


                    ' Dim SQLreader As System.Data.SqlClient.SqlDataReader = SQLcommand.ExecuteReader()
                    ' DataGridView3.Rows.Clear()
                    'Tracks the current record number
                    Dim n As Integer = 0
                    Dim kl As Integer = 0
                    Dim i As Integer
                    While lrd.Read()
                        kl += 1
                    End While
                    If kl = 0 Then
                        MessageBox.Show("There aren't empty records", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                    'SQLcommand.ExecuteNonQuery()
                    objzapytanie.Dispose()
                    If kl > 0 Then
                        ' delete all empty records
                        lrd.NextResult()
                        StrSqlite = "Delete FROM TranslateBase where PL=''"
                        objzapytanie = New System.Data.SQLite.SQLiteCommand(StrSqlite, Con)

                        objzapytanie.ExecuteNonQuery()
                        objzapytanie.Dispose()
                        'objzapytanie.ExecuteNonQuery()
                        MessageBox.Show("Record was updated", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If




                Catch ex As DataException
                    MessageBox.Show("Error while data were updating", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    Con.Close()
                End Try
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Dim StrWpis As String
            Dim StrSqlite As String
            connect_to_base()
            ' find all records which are empty

            StrSqlite = "DELETE from TranslateBase WHERE LN NOT IN (select  MIN(LN) FROM TranslateBase GROUP BY PL)"
            If Con.State = ConnectionState.Open Then
                Try
                    Dim objzapytanie As New System.Data.SQLite.SQLiteCommand(StrSqlite, Con)

                    objzapytanie.ExecuteNonQuery()
                    objzapytanie.Dispose()

                    'objzapytanie.ExecuteNonQuery()
                    MessageBox.Show("Record was updated.Deleted the same records", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)



                Catch ex As DataException
                    MessageBox.Show("Error while data were updating", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    Con.Close()
                End Try
            End If
        Else
        End If
    End Sub

    Private Sub SaveEditedRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveEditedRecordToolStripMenuItem.Click
        ' edycja 
        Dim StrWpis As String
        Dim StrSqlite As String
        Dim i, j As Integer
        i = DataGridView2.CurrentCell.RowIndex
        'MsgBox(DataGridView2.Item(0, i).Value)
        If DataGridView2.Item(1, i).Value = "" Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If DataGridView2.Item(2, i).Value = "" Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If DataGridView2.Item(3, i).Value = "" Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If DataGridView2.Item(4, i).Value = "" Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If DataGridView2.Item(5, i).Value = "" Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If DataGridView2.Item(6, i).Value = "" Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If DataGridView2.Item(7, i).Value = "" Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Exit Sub
        End If
        connect_to_base()
        StrWpis = "PL='" & DataGridView2.Item(1, i).Value & "',DE='" & DataGridView2.Item(2, i).Value & "',EN='" & DataGridView2.Item(3, i).Value & "',HU='" & DataGridView2.Item(4, i).Value & "', RU='" & DataGridView2.Item(5, i).Value & "',CZ='" & DataGridView2.Item(6, i).Value & "',SLOV='" & DataGridView2.Item(7, i).Value & "'"
        StrSqlite = " update TranslateBase set " & StrWpis & " where LN=" & DataGridView2.Item(0, i).Value
        If Con.State = ConnectionState.Open Then
            Try
                Dim objzapytanie As New System.Data.SQLite.SQLiteCommand(StrSqlite, Con)
                objzapytanie.ExecuteNonQuery()
                odczyt_do_edycji(nr_rekordu)
                MessageBox.Show("Record was updated", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As DataException
                MessageBox.Show("Error while data were updating", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Con.Close()
            End Try
        End If
    End Sub

    Private Sub RefreshRecordsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshRecordsToolStripMenuItem.Click
        Dim i As Int32              'General counter
        Dim Row As DataRow
        ' ToolStripStatusLabel1.Text = "Connecting..."
        'Open a connection
        Dim Con As New SQLite.SQLiteConnection
        Con.ConnectionString = "Data Source=" & DirectoryW & "\TranslateBase.s3db;"
        Con.Open()
        ' Sort data 

        'Fill a DataSet object
        Dim Adapter As New SQLite.SQLiteDataAdapter("SELECT * FROM TranslateBase", Con)
        Dim ds As New DataSet
        Adapter.FillSchema(ds, SchemaType.Mapped, "TranslateBase")
        Adapter.Fill(ds, "TranslateBase")

        'Display some data
        Debug.Print("Total Rows: " & ds.Tables("TranslateBase").Rows.Count)
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        i = 1           'Tracks the current record number
        For Each Row In ds.Tables("TranslateBase").Rows
            Dim n As Integer = DataGridView1.Rows.Add()
            DataGridView1.Rows.Item(n).Cells(0).Value = Row("LN")
            DataGridView1.Rows.Item(n).Cells(1).Value = Row("PL")
            DataGridView1.Rows.Item(n).Cells(2).Value = Row("DE")
            DataGridView1.Rows.Item(n).Cells(3).Value = Row("EN")
            DataGridView1.Rows.Item(n).Cells(4).Value = Row("HU")
            DataGridView1.Rows.Item(n).Cells(5).Value = Row("RU")
            DataGridView1.Rows.Item(n).Cells(6).Value = Row("CZ")
            DataGridView1.Rows.Item(n).Cells(7).Value = Row("SLOV")

            n = DataGridView2.Rows.Add()
            DataGridView2.Rows.Item(n).Cells(0).Value = Row("LN")
            DataGridView2.Rows.Item(n).Cells(1).Value = Row("PL")
            DataGridView2.Rows.Item(n).Cells(2).Value = Row("DE")
            DataGridView2.Rows.Item(n).Cells(3).Value = Row("EN")
            DataGridView2.Rows.Item(n).Cells(4).Value = Row("HU")
            DataGridView2.Rows.Item(n).Cells(5).Value = Row("RU")
            DataGridView2.Rows.Item(n).Cells(6).Value = Row("CZ")
            DataGridView2.Rows.Item(n).Cells(7).Value = Row("SLOV")
            'MsgBox(i & " Name: " & Row("LN") & " City: " & Row("PL") & " State: " & Row("ENG"))
            i += 1
        Next

        'Close the connection
        Con.Close()
        ' Add any initialization after the InitializeComponent() call.
        ToolStripStatusLabel4.Text = DataGridView1.RowCount - 1
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ' edycja 
        Dim StrWpis As String
        Dim StrSqlite As String

        If TextBox2.Text.Length = 0 Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox2.Focus()
            Exit Sub
        End If
        If TextBox3.Text.Length = 0 Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox3.Focus()
            Exit Sub
        End If
        If TextBox4.Text.Length = 0 Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Focus()
            Exit Sub
        End If
        If TextBox5.Text.Length = 0 Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox5.Focus()
            Exit Sub
        End If
        If TextBox6.Text.Length = 0 Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox6.Focus()
            Exit Sub
        End If

        If TextBox7.Text.Length = 0 Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox7.Focus()
            Exit Sub
        End If
        If TextBox8.Text.Length = 0 Then
            MessageBox.Show("Incorrect entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox8.Focus()
            Exit Sub
        End If
        connect_to_base()
        StrWpis = "PL='" & TextBox2.Text & "',DE='" & TextBox3.Text & "',EN='" & TextBox4.Text & "',HU='" & TextBox5.Text & "', RU='" & TextBox6.Text & "',CZ='" & TextBox7.Text & "',SLOV='" & TextBox8.Text & "'"
        StrSqlite = " update TranslateBase set " & StrWpis & " where LN=" & TextBox1.Text
        If Con.State = ConnectionState.Open Then
            Try
                Dim objzapytanie As New System.Data.SQLite.SQLiteCommand(StrSqlite, Con)
                objzapytanie.ExecuteNonQuery()
                odczyt_do_edycji(nr_rekordu)
                MessageBox.Show("Record was updated", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As DataException
                MessageBox.Show("Error while data were updating", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Con.Close()
            End Try
        End If


    End Sub


    Private Sub DeleteRecordsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteRecordsToolStripMenuItem.Click
        connect_to_base()
        If Con.State = ConnectionState.Open Then
            Dim isd As Integer
            isd = DataGridView2.CurrentCell.RowIndex
            If isd > 0 Then
                Dim pytanie As DialogResult
                pytanie = MessageBox.Show("Do you want delate current records?", "Delate data", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
                If pytanie = System.Windows.Forms.DialogResult.Yes Then
                    Dim strSqlite As String

                    strSqlite = " Delete FROM TranslateBase where LN=" & DataGridView2.Item(0, isd).Value
                    Try
                        Dim objzapytanie As New System.Data.SQLite.SQLiteCommand(strSqlite, Con)
                        objzapytanie.ExecuteNonQuery()

                        MessageBox.Show("Record was updated.Deleted the selected record", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Dim i As Int32              'General counter
                        Dim Row As DataRow
                        ' ToolStripStatusLabel1.Text = "Connecting..."
                        'Open a connection

                        'Fill a DataSet object
                        Dim Adapter As New SQLite.SQLiteDataAdapter("SELECT * FROM TranslateBase", Con)
                        Dim ds As New DataSet
                        Adapter.FillSchema(ds, SchemaType.Mapped, "TranslateBase")
                        Adapter.Fill(ds, "TranslateBase")

                        'Display some data
                        Debug.Print("Total Rows: " & ds.Tables("TranslateBase").Rows.Count)
                        DataGridView1.Rows.Clear()
                        DataGridView2.Rows.Clear()
                        i = 1           'Tracks the current record number
                        For Each Row In ds.Tables("TranslateBase").Rows
                            Dim n As Integer = DataGridView1.Rows.Add()
                            DataGridView1.Rows.Item(n).Cells(0).Value = Row("LN")
                            DataGridView1.Rows.Item(n).Cells(1).Value = Row("PL")
                            DataGridView1.Rows.Item(n).Cells(2).Value = Row("DE")
                            DataGridView1.Rows.Item(n).Cells(3).Value = Row("EN")
                            DataGridView1.Rows.Item(n).Cells(4).Value = Row("HU")
                            DataGridView1.Rows.Item(n).Cells(5).Value = Row("RU")
                            DataGridView1.Rows.Item(n).Cells(6).Value = Row("CZ")
                            DataGridView1.Rows.Item(n).Cells(7).Value = Row("SLOV")

                            n = DataGridView2.Rows.Add()
                            DataGridView2.Rows.Item(n).Cells(0).Value = Row("LN")
                            DataGridView2.Rows.Item(n).Cells(1).Value = Row("PL")
                            DataGridView2.Rows.Item(n).Cells(2).Value = Row("DE")
                            DataGridView2.Rows.Item(n).Cells(3).Value = Row("EN")
                            DataGridView2.Rows.Item(n).Cells(4).Value = Row("HU")
                            DataGridView2.Rows.Item(n).Cells(5).Value = Row("RU")
                            DataGridView2.Rows.Item(n).Cells(6).Value = Row("CZ")
                            DataGridView2.Rows.Item(n).Cells(7).Value = Row("SLOV")
                            'MsgBox(i & " Name: " & Row("LN") & " City: " & Row("PL") & " State: " & Row("ENG"))
                            i += 1
                        Next

                    Catch ex As Exception
                        MessageBox.Show("Error while data were updating", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Finally
                        Con.Close()
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub Button9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        ' uwaga
        ' uwaga
        Dim temp_text1 As String
        odczyt_do_edycji(0)
        If nr_rekordu >= 0 Then
            nr_rekordu = objtabela.Rows.Count - 1
            wyswietl_rekord(nr_rekordu)
            temp_text1 = TextBox1.Text
            TextBox1.Text = 1 + temp_text1
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
        End If


    End Sub

    Private Sub TextBox1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox1.MouseDoubleClick
        '  odczyt_do_edycji(0)'
        Dim i As Int32              'General counter
        Dim Row As DataRow
        ' ToolStripStatusLabel1.Text = "Connecting..."
        'Open a connection
        Dim Con As New SQLite.SQLiteConnection
        Con.ConnectionString = "Data Source=" & DirectoryW & "\TranslateBase.s3db;"
        Con.Open()
        ' Sort data 
        Try
            'Fill a DataSet object
            Dim Adapter As New SQLite.SQLiteDataAdapter("SELECT * FROM TranslateBase", Con)
            Dim ds As New DataSet
            Adapter.FillSchema(ds, SchemaType.Mapped, "TranslateBase")
            Adapter.Fill(ds, "TranslateBase")
            ' Assign the DataSet as the DataSource for the BindingSource.

            BindingSource1.DataSource = ds.Tables("TranslateBase")
            objtabela = ds.Tables("TranslateBase")

        Finally
            ' Con.Dispose()
            'Close the connection
            Con.Close()
        End Try


        If nr_rekordu >= 0 Then
            nr_rekordu = TextBox1.Text
            wyswietl_rekord(nr_rekordu)

        End If
    End Sub


End Class