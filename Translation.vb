Imports System.Text
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports Excel


Public Class Translation
    Public DirectoryW As String = My.Application.Info.DirectoryPath
    Public textW As String

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataBaseSet.Show()
    End Sub
    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, ByVal value As T) As T
        target = value
        Return value
    End Function

    Private Sub Translation_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Dim i As Int32              'General counter
        Dim columnv As DataColumn

        'Open a connection
        Dim Con As New SQLite.SQLiteConnection
        Con.ConnectionString = "Data Source=" & DirectoryW & "\TranslateBase.s3db;"
        Con.Open()

        'Fill a DataSet object
        Dim Adapter As New SQLite.SQLiteDataAdapter("SELECT * FROM TranslateBase", Con)
        Dim ds As New DataSet
        Adapter.FillSchema(ds, SchemaType.Mapped, "TranslateBase")
        Adapter.Fill(ds, "TranslateBase")

        'Display some data
        Debug.Print("Total Rows: " & ds.Tables("TranslateBase").Rows.Count)
        'DataGridView1.Rows.Clear()
        ' clear all rows in comboboxes
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()

        i = 1           'Tracks the current record number
        For i = 2 To ds.Tables("TranslateBase").Columns.Count - 1
            'For Each columnv In ds.Tables("TranslateBase").Columns
            ' Append text to comboboxes
            Dim sName As String
            Select Case ds.Tables("TranslateBase").Columns.Item(i).ColumnName
                Case "DE"
                    sName = " - German"
                Case "EN"
                    sName = "  - English"
                Case "HU"
                    sName = " - Hungarian"
                Case "RU"
                    sName = " - Russian"
                Case "CZ"
                    sName = " - Czech "
                Case "SLOV"
                    sName = " - Slovak"
            End Select
            ComboBox2.Items.Add(ds.Tables("TranslateBase").Columns.Item(i).ColumnName & sName)
            ComboBox3.Items.Add(ds.Tables("TranslateBase").Columns.Item(i).ColumnName & sName)
            'i += 1
        Next
        ComboBox1.Items.Add("PL - Polish")
        ' Set item in comboboxes
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 1
        'Close the connection
        Con.Close()
        If ComboBox1.SelectedIndex >= 0 Then
            ComboBox1.BackColor = Color.Silver
        End If
        If ComboBox2.SelectedIndex >= 0 Then
            ComboBox2.BackColor = Color.LightCoral
        End If
        If ComboBox3.SelectedIndex >= 0 Then
            ComboBox3.BackColor = Color.LemonChiffon
        End If
    End Sub
    Public ThisApplication As Inventor.Application
    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        If Form1.TextBox13.Text = "" Then
            MessageBox.Show("Select folder to save up a file. You can set up its in settings application", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            GoTo EndOfLoop7
        End If

        Dim folder As String = Nothing
        Dim listTab(1) As String
        Dim progressbar1_value As Integer = 0
        Dim progressbar1_value_max As Integer = 0
        ProgressBar1.Minimum = 0
        For di As Integer = 0 To Form1.ListView1.Items.Count - 1
            If Form1.ListView1.Items(di).Selected = False Then
                If Form1.ListView1.Items(di).Text.Contains(".idw") = True Then
                    progressbar1_value_max += 1
                End If
            End If
        Next
        Dim oInvApp As Inventor.Application
        'Launch Inventor
        oInvApp = CreateObject("Inventor.Application")
        oInvApp.Visible = True
        Dim oDoc As DrawingDocument
        ' Maksymalna liczba plików do analizy
        ToolStripStatusLabel2.Text = progressbar1_value_max
        ' calculation the number of appeared file
        ProgressBar1.Maximum = progressbar1_value_max
        Dim d_start As Integer = 0
        If CheckBox1.Checked = True Then
            With OpenFileDialog1
                .InitialDirectory = "C:\"
                .Title = "Select Inventor file"
                .DefaultExt = ".idw"
                .Filter = "Inventor File  (*.idw) | *.idw"
                .ShowDialog()
                d_start = 1
                ProgressBar1.Maximum = 1
            End With
            ToolStripStatusLabel2.Text = 1
        End If
        For di As Integer = 0 To Form1.ListView1.Items.Count + d_start - 1
            If d_start <> 1 Then

                If Form1.ListView1.Items(di).Selected = False Then

                    If Form1.ListView1.Items(di).Text.Contains(".idw") = True Then
                        listTab(0) = CStr(Form1.TreeView1.SelectedNode.Tag)
                        Form1.ListView1.Items(di).Tag = Form1.ListView1.Items(di).Text
                        folder = listTab(0) + "\" + CStr(Form1.ListView1.Items(di).Tag)
                        OpenFileDialog1.FileName = folder
                        progressbar1_value += 1
                        ProgressBar1.Value = progressbar1_value
                    Else
                        ' MessageBox.Show("Don't exist any file with extension '.idw", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                End If
            End If

            'If progressbar1_value_max = 0 And folder = "" Then
            '    MessageBox.Show("Don't exist any file with extension '.idw", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            'End If

            If OpenFileDialog1.FileName <> "" Then
                Dim labelPath As String
                Dim txtName As String = OpenFileDialog1.FileName
                Dim OpenFileName As String() = txtName.Split(New Char() {("\"c)})
                Dim vOpenFileName As Integer = OpenFileName.GetUpperBound(0)
                TextBox1.Text = OpenFileName(vOpenFileName)
                Dim openFilelenght As Integer = Len(OpenFileName(vOpenFileName))
                Dim labelPathLenght As Integer = Len(OpenFileDialog1.FileName)

                If labelPathLenght < 110 Then
                    labelPath = Microsoft.VisualBasic.Left(OpenFileDialog1.FileName, labelPathLenght - openFilelenght)
                Else
                    Dim openFilelenght_0 As Integer = Len(OpenFileName(0))
                    Dim openFilelenght_1 As Integer = Len(OpenFileName(1))
                    Dim openFilelenght_2 As Integer = Len(OpenFileName(2))
                    Dim openFilelenght_3 As Integer = Len(OpenFileName(3))
                    Dim openFilelenght_4 As Integer = Len(OpenFileName(4))

                    labelPath = Microsoft.VisualBasic.Left(OpenFileDialog1.FileName, labelPathLenght - openFilelenght - openFilelenght_4 - openFilelenght_3) & "..."
                End If
                Label6.Text = labelPath


                ProgressBar2.Minimum = 0

                ToolStripStatusLabel1.Text = "Connecting with a Database and INVENTOR "
                ' powiêkszanie rozmaru okna i wyœwitlanie paska pastêpu
                ' increase window size and display statusbar
                If Me.Height = 539 Then
                    Me.Height = 646
                End If

                'Set Startup options.
                oInvApp.GeneralOptions.ShowStartupDialog = False


                Dim oDocOpenOptions As NameValueMap
                oDocOpenOptions = oInvApp.TransientObjects.CreateNameValueMap()

                ' Set the option to defer the drawing update.
                Call oDocOpenOptions.Add("DeferUpdates", True)

                ' Open the document, suppressing the warning dialog.
                oInvApp.SilentOperation = True
                'Dim oDoc As DrawingDocument
                ' sprawdzenie rozszerzenia pliku
                Dim cIdw As Integer = OpenFileName.GetUpperBound(0)
                Dim extFile As Integer = Len(OpenFileName(cIdw))
                Dim l_3 As Integer = extFile - 3
                Dim LFileName As String = Microsoft.VisualBasic.Right(OpenFileName(cIdw), extFile - l_3)
                If LFileName = "idw" Then
                    oDoc = oInvApp.Documents.OpenWithOptions(OpenFileDialog1.FileName, oDocOpenOptions, True)
                Else
                    GoTo EndOfLoop6
                End If
                'oDoc = oInvApp.Documents.Open(OpenFileDialog1.FileName, True)
                ' oDoc = ThisApplication.ActiveDocument
                Dim oSheet As Inventor.Sheet
                oSheet = oDoc.ActiveSheet

                Dim oView As DrawingView
                '   oView = oDoc.ActiveSheet.DrawingViews.Item(1)

                'If oDoc.ActiveSheet.DrawingViews.Count > 0 Then MsgBox("Number of Views = " & oDoc.ActiveSheet.DrawingViews.Count)
                Dim oTitleBlock As TitleBlock = oSheet.TitleBlock

                'MsgBox("Sheet Name: " & oSheet.Name)
                'MsgBox("Number of Parts = " & oDoc.ActiveSheet.PartsLists.Count)
                'MsgBox("View Name: " & oView.Name)
                'MsgBox("View Item: " & oDoc.ActiveSheet.DrawingViews.Item(1).Name)
                Dim w1 As Integer

                Dim wTitleblock_progress As Integer = 0
                'Dim oSheet As Sheet
                'Dim oTitleBlock As TitleBlock

                'if you want to use the default template then set UseDefaultTemplate = True
                'if you want to use a custom template set the path and filename of sTemplatePart and UseDefaultTemaplte = False
                'declare the environment
                ' Path to my application DirectoryW
                Dim DirectoryW As String = My.Application.Info.DirectoryPath
                Dim oApp As Inventor.Application
                ' Dim oCurrentDoc As Document
                ' Dim oNewDoc As Document
                ' Dim UseDefaultTemplate As Boolean
                ' Dim sCurrentFileName As String
                'Dim sTemplateDrawing As String
                Dim obox As Inventor.TextBox
                Dim oboxD As Inventor.TextBox
                Dim ts As String = Nothing
                Dim create_data As String = Nothing
                Dim NR As String = Nothing
                Dim Autor As String = Nothing
                Dim Sprawdzil As String = Nothing
                Dim Zatwierdzil As String = Nothing
                Dim check_data As String = Nothing
                Dim approved_data As String = Nothing
                Dim sheet_number As String = Nothing ' zatwierdzenie
                Dim sheet_size As String = Nothing
                Dim thema As String = Nothing
                Dim sheet_Numb As String = Nothing
                Dim Articel_number As String = Nothing
                Dim scale As String = Nothing
                Dim masa As String = Nothing
                Dim version_Numb As String = Nothing
                '
                ' Progressbar
                Dim css As Integer = 0
                Dim xgg As Integer = 0
                Dim xcc As Integer = 0
                Dim xbb As Integer = 0
                Dim xss As Integer = 0
                Dim oTBD As Integer = 0
                Dim oATBDSTC As Integer = 0
                Dim xPartlist As Integer = 0
                Dim oASSC As Integer = 0
                Dim oSSSC As Integer = 0
                Dim oSSDC As Integer = 0
                Dim oADNC As Integer = 0
                Dim oAPIPLRC As Integer = 0
                Dim tevTainT As Integer = 0

                ' sketchedsymbols progressbar
                Dim xwer As Integer = 0
                Dim xwer1 As Integer = 0
                Dim xwer2 As Integer = 0
                Dim xwer3 As Integer = 0
                Dim xwer4 As Integer = 0
                Dim xwer5 As Integer = 0
                Dim xwer6 As Integer = 0
                Dim xwer7 As Integer = 0
                Dim xwer8 As Integer = 0
                ' revision table
                Dim xrev As Integer = 0
                ' drawingnotes
                Dim xDN As Integer = 0
                ' dodawanie danych do tabeli list( arraylist)
                Dim xwer11 As Integer = 0
                Dim xwer21 As Integer = 0
                Dim xwer31 As Integer = 0
                Dim xwer41 As Integer = 0
                Dim xwer51 As Integer = 0
                Dim xwer61 As Integer = 0
                Dim xwer71 As Integer = 0
                Dim xwer81 As Integer = 0

                ' Liczenie liczby wpisów dla sketchsymbols- progressbar
                If RadioButton9.Checked = True Then
                    oTBD = oDoc.TitleBlockDefinitions.Count
                    oATBDSTC = oDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes.Count
                Else
                    oTBD = 0
                    oATBDSTC = 0
                End If
                If RadioButton4.Checked = True Then
                    oASSC = oDoc.ActiveSheet.SketchedSymbols.Count
                    oSSSC = oDoc.Sheets(1).SketchedSymbols.Count
                    oSSDC = oDoc.SketchedSymbolDefinitions.Count
                Else
                    oASSC = 0
                    oSSSC = 0
                    oSSDC = 0
                End If
                '  Drawing notes
                If RadioButton3.Checked = True Then
                    oADNC = oDoc.ActiveSheet.DrawingNotes.Count
                    'oSSSC = oDoc.Sheets(1).SketchedSymbols.Count

                Else
                    oADNC = 0
                    '  oSSSC = 0

                End If
                ' part list
                If RadioButton2.Checked = True Then
                    oAPIPLRC = oDoc.ActiveSheet.PartsLists.Item(1).PartsListRows.Count
                    'oSSSC = oDoc.Sheets(1).SketchedSymbols.Count
                    oSSDC = oDoc.SketchedSymbolDefinitions.Count
                Else
                    oAPIPLRC = 0
                    '  oSSSC = 0

                End If
                ' revision table
                If RadioButton8.Checked = True Then
                    tevTainT = 5
                Else
                    tevTainT = 0
                End If

                If oDoc.ActiveSheet.DrawingViews.Count <> 0 Then
                    If oDoc.ActiveSheet.SketchedSymbols.Count <> 0 And RadioButton4.Checked = True Then
                        For xwer = 1 To oDoc.Sheets(1).SketchedSymbols.Count
                            '  nameN += 1
                            '  MsgBox(oDoc.Sheets(1).SketchedSymbols.Item(wer).Name)
                            For tSketch As Integer = 1 To oDoc.ActiveSheet.SketchedSymbols.Item(xwer).Definition.Sketch.TextBoxes.Count
                                If xwer = 1 Then xwer1 = oDoc.ActiveSheet.SketchedSymbols.Item(xwer).Definition.Sketch.TextBoxes.Count

                                If xwer = 2 Then xwer2 = oDoc.ActiveSheet.SketchedSymbols.Item(xwer).Definition.Sketch.TextBoxes.Count
                                If xwer = 3 Then xwer3 = oDoc.ActiveSheet.SketchedSymbols.Item(xwer).Definition.Sketch.TextBoxes.Count
                                If xwer = 4 Then xwer4 = oDoc.ActiveSheet.SketchedSymbols.Item(xwer).Definition.Sketch.TextBoxes.Count
                                If xwer = 5 Then xwer5 = oDoc.ActiveSheet.SketchedSymbols.Item(xwer).Definition.Sketch.TextBoxes.Count
                                If xwer = 6 Then xwer6 = oDoc.ActiveSheet.SketchedSymbols.Item(xwer).Definition.Sketch.TextBoxes.Count
                                If xwer = 7 Then xwer7 = oDoc.ActiveSheet.SketchedSymbols.Item(xwer).Definition.Sketch.TextBoxes.Count
                                If xwer = 8 Then xwer8 = oDoc.ActiveSheet.SketchedSymbols.Item(xwer).Definition.Sketch.TextBoxes.Count
                            Next
                        Next
                    End If
                    'ProgressBar2.Maximum = (oDoc.TitleBlockDefinitions.Count + oDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes.Count + oDoc.ActiveSheet.SketchedSymbols.Count + oDoc.Sheets(1).SketchedSymbols.Count + oDoc.SketchedSymbolDefinitions.Count + oDoc.SketchedSymbolDefinitions.Count + oDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes.Count + oDoc.ActiveSheet.DrawingNotes.Count + oDoc.ActiveSheet.PartsLists.Item(1).PartsListRows.Count + xwer1 + xwer2 + xwer3 + xwer4 + xwer5 + xwer6 + xwer7 + xwer8) ' _ oDoc.SketchedSymbolDefinitions.Count
                    ' If RadioButton4.Checked = True Or RadioButton1.Checked = True Or RadioButton2.Checked = True Or RadioButton3.Checked = True Or RadioButton8.Checked = True Or RadioButton9.Checked = True Then
                    ProgressBar2.Maximum = (oTBD + oATBDSTC + oASSC + oSSDC + oSSDC + oATBDSTC + oADNC + oAPIPLRC + tevTainT + xwer1 + xwer2 + xwer3 + xwer4 + xwer5 + xwer6 + xwer7 + xwer8 + xrev) ' _ oDoc.SketchedSymbolDefinitions.Count
                    'Dim obox As Inventor.TextBox
                    progressbar1_value = 1

                    ' Progressbar, ca³kowita wartoœæ wystêpuj¹cych elementów do t³umaczenia
                    If Not oDoc.ActiveSheet.TitleBlock Is Nothing And RadioButton9.Checked = True Then
                        ProgressBar1.Value = progressbar1_value
                        For Each obox In oDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes
                            'MsgBox(obox.Text)

                            Select Case obox.Text
                                Case "<TS> "
                                    ts = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<TS>"
                                    ts = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<TS>  <TS>"
                                    ts = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                    'oDoc.TitleBlock.SetPromptResultText(obox, "ccc")
                                Case "<DATA UTWORZENIA>"
                                    create_data = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<NR>"
                                    NR = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<AUTOR>"
                                    Autor = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<SPRAWDZI£>"
                                    Sprawdzil = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<ZATWIERDZI£>"
                                    Zatwierdzil = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<DATA SPRAWDZENIA>"
                                    check_data = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<DATA ZATWIERDZENIA>"
                                    approved_data = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<Numer arkusza>"
                                    sheet_number = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<Rozmiar arkusza>" 'format
                                    sheet_size = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<TEMAT>"
                                    thema = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<liczba arkuszy>"
                                    sheet_Numb = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<NR_ARTYKULU>"
                                    Articel_number = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "Skala"
                                    scale = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<MASA\P1;>"
                                    masa = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                                Case "<NUMER WERSJI>"
                                    version_Numb = oDoc.ActiveSheet.TitleBlock.GetResultText(obox)
                            End Select
                            ' progressbar
                            wTitleblock_progress += 1
                            ProgressBar2.Value = wTitleblock_progress
                        Next
                    End If
                End If


                'Select Case oInvApp.ActiveDocumentType

                '    Case DocumentTypeEnum.kAssemblyDocumentObject, DocumentTypeEnum.kDrawingDocumentObject

                '        sCurrentFileName = oInvApp.ActiveDocument.FullFileName

                '        If sCurrentFileName = "" Then
                '            MsgBox("The active file must first be saved", vbInformation, "Warning")
                '            Exit Sub
                '        End If

                '        'if you want to use the default template then set UseDefaultTemaplte = True
                '        'if you want to use a custom template set the path and filename of sTemplatePart and UseDefaultTemaplte = False

                '        UseDefaultTemplate = False
                '        sTemplateDrawing = "C:\Documents and Settings\bkonefal\Pulpit\Templates\TS_PL_DE_EN.idw" 'Change this path if necessary

                '        Select Case UseDefaultTemplate
                '            Case True
                '                oNewDoc = oInvApp.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, oInvApp.FileManager.GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject), True)
                '            Case False
                '                oNewDoc = oInvApp.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, sTemplateDrawing, True)
                '        End Select
                'End Select

                'Dim oTitleBlockDef As TitleBlockDefinition
                ''        On Error GoTo Errorhandler
                'oTitleBlockDef = oDoc.TitleBlockDefinitions.Item(7) ' this is our standard title block
                'Set Active Application
                'Dim oApp As Application
                oApp = ThisApplication
                'Get the active document and make sure it's a drawing.

                If oDoc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject And oDoc.DocumentSubType.DocumentSubTypeID <> "{BBF9FDF1-52DC-11D0-8C04-0800090BE8EC}" Then
                    MsgBox("Must be in Drawing Document to run the Title Block Update Tool! ")
                    Exit Sub
                End If
                'Check for and Remedy Deferred Updates
                If oDoc.DrawingSettings.DeferUpdates = True Then
                    ' RemoveDeferralForm.Show()
                End If

                'Save the names of the title block definition used for each sheet
                'and delete the title blocks.
                'Dim colDefNames As New Collection
                'Dim oTargetSheet As Sheet
                'For Each oTargetSheet In oDoc.Sheets
                '    'Activate the sheet.
                '    oTargetSheet.Activate()
                '    'Save the name of the title block definition used for this sheet.
                '    colDefNames.Add(oDoc.ActiveSheet.TitleBlock.Definition.Name)
                '    'Delete the title block.
                '    oTargetSheet.TitleBlock.Delete()
                'Next
                'Copy the new definitions from the source into the target document.

                ''Dim oSourceSS As SketchedSymbolDefinitions
                ''oSourceSS = oSourceDoc.SketchedSymbolDefinitions
                'Call oSourceTB.Item(12).CopyTo(oDoc, True)

                ''Call oSourceSS.Item("Revision Block").CopyTo(oDoc, True)

                ''Add the title blocks to the sheets.
                'Dim i As Integer





                oTitleBlock = oSheet.TitleBlock


                ' Dim obox As Inventor.TextBox
                'For Each obox In oDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes
                '    If obox.Text = "" Then
                '        oTitleBlock.SetPromptResultText(obox, oSheet.Name)
                '    End If
                'Next

                Dim sPromptStrings(39) As String
                sPromptStrings(1) = "Konstruowa³ Gezeichn./Design."
                sPromptStrings(2) = "Sprawdzi³ Geprueft/Checked"
                sPromptStrings(13) = create_data
                sPromptStrings(0) = "Nazwisko" & Chr(10) & " Name/Name"
                sPromptStrings(22) = Articel_number
                sPromptStrings(21) = sheet_Numb
                sPromptStrings(20) = sheet_size
                sPromptStrings(18) = approved_data
                sPromptStrings(19) = sheet_number
                sPromptStrings(24) = "Zatwierdzi³ Bestaetigt" & Chr(10) & "Approved"
                sPromptStrings(17) = check_data
                sPromptStrings(23) = scale
                If scale = Nothing Then
                    sPromptStrings(22) = ""
                Else
                    sPromptStrings(22) = scale
                End If
                sPromptStrings(25) = masa
                sPromptStrings(3) = "Data" & Chr(10) & " Datum/Date"
                sPromptStrings(15) = Sprawdzil
                sPromptStrings(8) = "Ciê¿ar" & Chr(10) & "Gewicht" & Chr(10) & "Weight"
                sPromptStrings(6) = "Format: Blattgroesse:" & Chr(10) & "Size:"
                sPromptStrings(7) = "Skala" & Chr(10) & "Masstab" & Chr(10) & "Scale"
                sPromptStrings(9) = "Zmiana:" & Chr(10) & "Index:" & Chr(10) & "Revision:"
                sPromptStrings(10) = "Nr rysunku/ Zeichnungs-Nr./ Drawing-No."
                sPromptStrings(4) = "Nazwa z³o¿enia/czêœci/ / Benennung / Name of assembly/part drawing"
                sPromptStrings(16) = Zatwierdzil
                sPromptStrings(26) = "Nr.artyku³u:" & Chr(10) & "Artikel-Nr " & Chr(10) & "Article-N"
                sPromptStrings(5) = "Iloœæ ark.:" & Chr(10) & "Blatt-Gesamt" & Chr(10) & "No.of Sheets"
                sPromptStrings(14) = Autor

                sPromptStrings(27) = "Dzia³" & Chr(10) & "Abteilung/Division"
                If ts Is Nothing Then
                    sPromptStrings(12) = ""
                Else
                    sPromptStrings(12) = ts ' skala
                End If
                sPromptStrings(11) = NR ' numer art

                sPromptStrings(28) = "Ogólne tolerancje wykonania: Allgemeintoleranzen:" & Chr(10) & "General tolerances"
                sPromptStrings(29) = version_Numb
                sPromptStrings(30) = "JEDN." & Chr(10) & "1:EIN" & Chr(10) & "ITEM"
                sPromptStrings(31) = "MATERIA£" & Chr(10) & "WERKSTOFF" & Chr(10) & "MATERIAL"
                sPromptStrings(32) = "NR-RYSUNKU" & Chr(10) & "ZCHN.-NR" & Chr(10) & "DRAWING NO"
                sPromptStrings(33) = "NAZWA" & Chr(10) & " BENENNUNG" & Chr(10) & "DESCRIPTION"
                sPromptStrings(34) = "SZT." & Chr(10) & "STK." & Chr(10) & "QTY."
                sPromptStrings(35) = "POZ." & Chr(10) & "POS." & Chr(10) & "ITEM"
                sPromptStrings(36) = "CALK." & Chr(10) & "GESAMT" & Chr(10) & "TOTAL"
                sPromptStrings(37) = "NR-NORMY" & Chr(10) & "STANDARD" & Chr(10) & "STANDARD"
                sPromptStrings(38) = "MASA/MASSE/WEIGHT"
                sPromptStrings(39) = "Arkusz-Nr.:" & Chr(10) & "Blatt-Nr.:" & Chr(10) & "Drg.sheet No.:"
                If sheet_Numb = Nothing Then
                    sPromptStrings(21) = ""
                Else
                    sPromptStrings(21) = sheet_Numb
                End If
                If Articel_number = Nothing Then
                    sPromptStrings(22) = ""
                Else
                    sPromptStrings(22) = Articel_number
                End If


                Dim oTitleBlock1
                Dim oTitleBlock2
                Dim w As Integer = 0

                If Not oDoc.ActiveSheet.TitleBlock Is Nothing And RadioButton9.Checked = True Then
                    If oDoc.ActiveSheet.TitleBlock.Name <> "TS DE PL EN N S" Then
                        oDoc.ActiveSheet.TitleBlock.Delete()
                        '                        Set actTempl = ThisApplication.Documents.Open("C:\VAULT\Inventor 2011\Bin\Macros\Standard.idw", False)
                        oDoc.Activate()
                        Dim oSourceDoc As Inventor.DrawingDocument

                        oSourceDoc = oInvApp.Documents.Open(DirectoryW & "\Templates\TS_PL_DE_EN.idw", False)
                        ' path to my application DirectoryW

                        Dim oSourceTB As TitleBlockDefinitions
                        oSourceTB = oSourceDoc.TitleBlockDefinitions
                        Call oSourceTB.Item(12).CopyTo(oDoc, True) ', Multilanguage table is number 12
                        For w = 1 To oDoc.TitleBlockDefinitions.Count
                            'MsgBox(oDoc.TitleBlockDefinitions.Item(w).Name)
                            If oDoc.TitleBlockDefinitions.Item(w).Name = "MultiLanguage" Then
                                '  oTitleBlock2 = oSourceTB.Item(w).CopyTo(oDoc, True)
                                oTitleBlock2 = oDoc.TitleBlockDefinitions.Item(w).Name
                                ' Call oDoc.TitleBlockDefinitions.Item(w).CopyTo(oDoc, True)
                                'oTitleBlock2 = oDoc.TitleBlockDefinitions.Item(w).CopyTo(oDoc, True)
                                oTitleBlock1 = oDoc.ActiveSheet.AddTitleBlock(oTitleBlock2, , sPromptStrings)
                                'Set oTitleBlock1 = actDocd.ActiveSheet.AddTitleBlock(actDocd.TitleBlockDefinitions.Item(w), , sPromptStrings)
                                Exit For
                            Else

                            End If '
                            ' progressbar
                            ProgressBar2.Value = w + wTitleblock_progress
                        Next w
                    End If
                End If


                If Not oDoc.ActiveSheet.SketchedSymbols Is Nothing And RadioButton4.Checked = True Then
                    If oDoc.ActiveSheet.SketchedSymbols.Count > 0 Then
                        oDoc.Activate()
                        ' Add new template
                        Dim oSourceDoc As Inventor.DrawingDocument
                        oSourceDoc = oInvApp.Documents.Open(DirectoryW & "\Templates\TS_PL_DE_EN.idw", False)
                        Dim oSourceSketch As SketchedSymbolDefinitions
                        oSourceSketch = oSourceDoc.SketchedSymbolDefinitions
                        Dim SketchSname(12) As String
                        'Dim OpositionT(,) As Double
                        ' declare some kind an array for various position
                        Dim OpositionT_00 As Double
                        Dim OpositionT_01 As Double
                        Dim OpositionT_10 As Double
                        Dim OpositionT_11 As Double
                        Dim OpositionT_20 As Double
                        Dim OpositionT_21 As Double
                        Dim OpositionT_30 As Double
                        Dim OpositionT_31 As Double
                        Dim OpositionT_40 As Double
                        Dim OpositionT_41 As Double
                        Dim OpositionT_50 As Double
                        Dim OpositionT_51 As Double
                        Dim OpositionT_60 As Double
                        Dim OpositionT_61 As Double
                        Dim OpositionT_70 As Double
                        Dim OpositionT_71 As Double
                        Dim OpositionT_80 As Double
                        Dim OpositionT_81 As Double
                        Dim OpositionT_90 As Double
                        Dim OpositionT_91 As Double
                        Dim OpositionT_100 As Double
                        Dim OpositionT_101 As Double
                        Dim OpositionT_110 As Double
                        Dim OpositionT_111 As Double
                        Dim oSketchdel As Integer = 0
                        Dim faza As String
                        Dim xS As Integer = 0
                        'Dim oPosition As Point2d
                        Dim fazaT As Integer = 0
                        'For oSketchdel = 1 To SketchSname.Rank - 1
                        For oSketchdel = 1 To oDoc.Sheets(1).SketchedSymbols.Count

                            'ReDim OpositionT(xS, 1)
                            Select Case oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                'Select Case SketchSname(oSketchdel - 1)
                                Case "Tolerancja otworów"

                                    OpositionT_00 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_01 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(0) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    ' oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                                Case "Tolerancja ogólna wykonania"
                                    OpositionT_10 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_11 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(1) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    'oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                                Case "Tolerancja ogólna dla konstrukcji spawanych"
                                    OpositionT_20 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_21 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(2) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    ' oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                                Case "Tolerancja otworów pasowanych"
                                    OpositionT_30 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_31 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(3) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    'oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                                Case "Spoiny nie oznaczone"
                                    OpositionT_40 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_41 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(4) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    ' oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                                Case "Krawêdzie"
                                    OpositionT_50 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_51 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(5) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    For fazaT = 1 To oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Definition.Sketch.TextBoxes.Count

                                        Select Case oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Definition.Sketch.TextBoxes.Item(fazaT).Text
                                            Case "FAZA"

                                                faza = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Definition.Sketch.TextBoxes.Item(fazaT))
                                        End Select
                                    Next
                                    ' oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                                Case "Momenty Dokrêcania Œrub"
                                    OpositionT_60 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_61 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(6) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    ' oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                                Case "Kierunek ruchu"
                                    OpositionT_70 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_71 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(7) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    'oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                                Case "40 HMT"
                                    OpositionT_80 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_81 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(8) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    ' oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                                Case "Cynkowaæ"
                                    OpositionT_90 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_91 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(9) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    ' oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                                Case "ko³o ³añcuchowe"
                                    OpositionT_100 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_101 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(10) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    ' oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                                Case "ko³o zêbate"
                                    OpositionT_110 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.X
                                    OpositionT_111 = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Position.Y
                                    SketchSname(11) = oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Name
                                    ' oDoc.ActiveSheet.SketchedSymbols.Item(oSketchdel).Delete()
                            End Select
                            xss += 1
                            ' progressbar
                            ProgressBar2.Value = w + wTitleblock_progress + xss
                        Next
                        Dim oSketchdelW As Integer
                        Dim XD As Integer
                        ' Delate appeared symbols. The skechedsymbols are updated All the time
                        For oSketchdelW = 0 To SketchSname.GetUpperBound(0) - 1
                            If Not SketchSname(oSketchdelW) Is Nothing Then
                                XD += 1
                            End If
                        Next
Endofloop:
                        For cS As Integer = 1 To XD
                            If oDoc.Sheets(1).SketchedSymbols.Count > 0 Then
                                '  oDoc.sheets(1).SketchedSymbols.Item(cs).
                                oDoc.Sheets(1).SketchedSymbols.Item(cS).Delete()
                                oDoc.Sheets(1).Update()
                                XD = oDoc.Sheets(1).SketchedSymbols.Count
                                If XD = 0 Then
                                    Exit For
                                End If
                                GoTo EndOfLoop
                            End If
                            css += 1
                            ProgressBar2.Value = w + wTitleblock_progress + xS + css
                        Next
                        ' Delate all sketcheddefinitions. Clean up all
Endofloop1:
                        For xG As Integer = 1 To oDoc.SketchedSymbolDefinitions.Count
                            If oDoc.SketchedSymbolDefinitions.Count > 0 Then
                                oDoc.SketchedSymbolDefinitions.Item(xG).Delete()
                                GoTo EndOfLoop1

                            End If
                            ProgressBar2.Value = w + wTitleblock_progress + xss + xG + css
                        Next


                        'For nSketSym As Integer = 1 To oDoc.SketchedSymbolDefinitions.Count
                        ' Copy sketchsymbol from templets into current document
                        ' Kopiwanie wystêpuj¹cego na rysunku pocz¹tkowym symbolu
                        For osketchw As Integer = 1 To oSourceSketch.Count
                            If SketchSname(osketchw - 1) = "Tolerancja otworów" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If

                            If SketchSname(osketchw - 1) = "Tolerancja ogólna wykonania" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If
                            If SketchSname(osketchw - 1) = "Tolerancja ogólna dla konstrukcji spawanych" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If

                            If SketchSname(osketchw - 1) = "Tolerancja otworów pasowanych" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If

                            If SketchSname(osketchw - 1) = "Spoiny nie oznaczone" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If

                            If SketchSname(osketchw - 1) = "Krawêdzie" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If

                            If SketchSname(osketchw - 1) = "Momenty Dokrêcania Œrub" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If

                            If SketchSname(osketchw - 1) = "Kierunek ruchu" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If

                            If SketchSname(osketchw - 1) = "40 HMT" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If

                            If SketchSname(osketchw - 1) = "Cynkowaæ" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If

                            If SketchSname(osketchw - 1) = "ko³o ³añcuchowe" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If

                            If SketchSname(osketchw - 1) = "ko³o zêbate" And Not SketchSname(osketchw - 1) Is Nothing Then
                                Call oSourceSketch.Item(osketchw).CopyTo(oDoc, True)
                            Else
                            End If


                        Next
                        Dim oTitleSketch1

                        ' copy all sketchsymbols to activesheet
                        ' Definiowanie promptstringów dla wszystkich symboli
                        Dim ws As Integer
                        '"Tolerancja otworów" - Holes tolerances
                        Dim sPromptStrings1(2) As String
                        sPromptStrings1(0) = "Tolerancja wykonania nieoznaczonych otworów"
                        sPromptStrings1(2) = "Manufacturing tolerance of unmarked holes"
                        sPromptStrings1(1) = "Alle nicht bezeichneten Bohrungen" & Chr(10) & "alle umbemaßten Schweißnähte"

                        '"Tolerancja ogólna wykonania" -  manufacturing general tolerances 
                        Dim sPromptStrings2(7) As String
                        sPromptStrings2(0) = "Freimasztoleranzen fuer Bearbeitung"
                        sPromptStrings2(1) = "Middle"
                        sPromptStrings2(2) = "Mittle"
                        sPromptStrings2(3) = "Stopien dokladnoœci"
                        sPromptStrings2(4) = "General tolerances for treatment"
                        sPromptStrings2(5) = "Degree of accuracy"
                        sPromptStrings2(6) = "Genauigkeitsgrad"
                        sPromptStrings2(7) = "Tolerancja ogolna" & Chr(10) & "wykonania"
                        '"Tolerancja ogólna dla konstrukcji spawanych"
                        Dim sPromptStrings3(5) As String
                        sPromptStrings3(0) = "Genauigkeitsgrad" '"Allgemeintoleranzen fuer Schweiszkonstruktionen"
                        sPromptStrings3(1) = "Degree of accuracy" '"General tolerances for welding construction"
                        sPromptStrings3(2) = "Stopien dokladnoœci" ' '"Mittle"
                        sPromptStrings3(3) = "Tolerancja ogolna" & Chr(10) & "dla konstrukcji spawanych" '
                        sPromptStrings3(4) = "Allgemeintoleranzen" & Chr(10) & " fuer Schweiszkonstruktionen"
                        sPromptStrings3(5) = "General tolerances" & Chr(10) & " for welding construction" '
                        ' sPromptStrings3(6) = 
                        ' sPromptStrings3(7) = 
                        '"Tolerancja otworów pasowanych"
                        Dim sPromptStrings4(2) As String
                        sPromptStrings4(1) = "Alle nicht bezeichneten" & Chr(10) & "Bohrungen alle umbemaßten Schweißnähte"
                        sPromptStrings4(2) = "Tolerance of fitting holes"
                        sPromptStrings4(0) = "Tolerancja wykonania" & Chr(10) & " otworów pasowanych"

                        '"Spoiny nie oznaczone"
                        Dim sPromptStrings5(3) As String
                        sPromptStrings5(0) = "Unmarked seam welding"
                        sPromptStrings5(1) = "Alle umbemaßten Schweißnähte"
                        sPromptStrings5(2) = "Spoiny nieoznaczone spawaæ:"
                        sPromptStrings5(3) = "SPOINA"

                        ' Case "Krawêdzie"
                        Dim sPromptStrings6(3) As String
                        sPromptStrings6(0) = "Krawedzie zalamac"
                        sPromptStrings6(1) = "Kanten gebrochen"
                        sPromptStrings6(2) = "Chamfer"
                        sPromptStrings6(3) = faza

                        '"Momenty Dokrêcania Œrub"
                        Dim sPromptStrings7(2) As String
                        sPromptStrings7(1) = "Momenty dokrêcania dla œrub ze ³bem np: DIN 912/933 itd./" & Chr(10) & "Anzugsmomente fuer Schrauben mit Kopf z.B. DIN 912/933 usw./" & Chr(10) & "Tightening torque for screws with head e.g. DIN 912/933 etc."
                        sPromptStrings7(2) = "/Klasa œrub" & Chr(10) & "/Festigkeitsklasse" & Chr(10) & "/Strenght category"
                        sPromptStrings7(0) = "Momenty dokrêcania œrub podane w Nm dla typowych gwintów ze œrednim" & Chr(10) & "wspó³czynnikiem  tarcia równym 0,12 (g³adka powierzchnia smarowana lub sucha)/" & Chr(10) & "Angeben in Nm fuer Regelgewinde bei einer mittleren Gleitreibungszahl  von 0,12" & Chr(10) & "(gute Oberflaeche geschmiert oder trocken)/" & Chr(10) & "Tightening torque data in Nm for regular type of screw thread with an average friction " & Chr(10) & "factor of 0,12 (smooth surface lubricated or not)"

                        '"Kierunek ruchu"
                        Dim sPromptStrings8(0) As String
                        sPromptStrings8(0) = "Kierunek transportu/" & Chr(10) & "Foerderrichtung/" & Chr(10) & "Conveying direction"
                        '"40 HMT"
                        Dim sPromptStrings9(1) As String
                        'sPromptStrings9(0) = "Uwaga / Bemerkung / Note:" & Chr(10) & "40HMT - materia³ ulepszony cieplnie 32÷40 HRC" & Chr(10) & "40HMT - Vergüten 32÷40 HRC" & Chr(10) & "40HMT - quenched material 32÷40 HRC"
                        sPromptStrings9(0) = "Uwaga / Bemerkung / Note:"
                        sPromptStrings9(1) = "40HMT - materia³ ulepszony cieplnie 32÷40 HRC" & "/" & Chr(10) & "40HMT - Vergüten 32÷40 HRC" & "/" & Chr(10) & "40HMT - quenched material 32÷40 HRC"
                        '  sPromptStrings9(2) = 
                        ' sPromptStrings9(3) = "40HMT - quenched material 32÷40 HRC"
                        '"Cynkowaæ" - Galvanized
                        Dim sPromptStrings10(1) As String
                        sPromptStrings10(0) = "Postac handlowa:" & Chr(10) & "Handelsform:" & Chr(10) & "Commercial form:"
                        sPromptStrings10(1) = "Ocynk. / Verzinkt / Galvanized"
                        ' "ko³o ³añcuchowe"
                        Dim sPromptStrings11(8) As String
                        sPromptStrings11(0) = "Liczba zêbów/Zähnezahl/Number of teeth"
                        sPromptStrings11(1) = "Podzia³ka / Graduierung / Graduation"
                        sPromptStrings11(2) = "Typ ³añcucha/Kettentyp/Chain type"
                        sPromptStrings11(3) = "Œrednica podzia³owa/Wirkdurchmesser/Pitch diameter"
                        sPromptStrings11(4) = ""
                        sPromptStrings11(5) = ""
                        sPromptStrings11(6) = ""
                        sPromptStrings11(7) = ""
                        sPromptStrings11(8) = "Dane ko³a ³añcuchowego/Kettenrad Daten/Sprocket data"

                        '"ko³o zêbate"
                        Dim sPromptStrings12(8) As String
                        sPromptStrings12(0) = "Liczba zêbów/Zähnezahl/Number of teeth"
                        sPromptStrings12(1) = "Modu³/Modul/Module"
                        sPromptStrings12(2) = "Szerokoœæ zêba/Zahnbreit/Tooth width"
                        sPromptStrings12(3) = "Œrednica podzia³owa/Wirkdurchmesser/Pitch diameter"
                        sPromptStrings12(4) = ""
                        sPromptStrings12(5) = ""
                        sPromptStrings12(6) = ""
                        sPromptStrings12(7) = ""
                        sPromptStrings12(8) = "Dane ko³a zêbatego / Stirnzahnrad Daten / Sprocket data"
                        Dim oSSDef As SketchedSymbolDefinition
                        Dim oTG As TransientGeometry
                        oTG = oInvApp.TransientGeometry

                        ' add all parameters to sketchdefinition
                        ' Dodawanie symbolu w oparciu o punkt wstawienia i domyœlny prompt
                        For ws = 1 To oDoc.SketchedSymbolDefinitions.Count
                            Select Case oDoc.SketchedSymbolDefinitions.Item(ws).Name
                                Case "Tolerancja otworów"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_00, OpositionT_01), 0, 1, sPromptStrings1)

                                Case "Tolerancja ogólna wykonania"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_10, OpositionT_11), 0, 1, sPromptStrings2)
                                Case "Tolerancja ogólna dla konstrukcji spawanych"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_20, OpositionT_21), 0, 1, sPromptStrings3)
                                Case "Tolerancja otworów pasowanych"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_30, OpositionT_31), 0, 1, sPromptStrings4)
                                Case "Spoiny nie oznaczone"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_40, OpositionT_41), 0, 1, sPromptStrings5)
                                Case "Krawêdzie"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_50, OpositionT_51), 0, 1, sPromptStrings6)
                                Case "Momenty Dokrêcania Œrub"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_60, OpositionT_61), 0, 1, sPromptStrings7)
                                Case "Kierunek ruchu"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_70, OpositionT_71), 0, 1, sPromptStrings8)
                                Case "40 HMT"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_80, OpositionT_81), 0, 1, sPromptStrings9)
                                Case "Cynkowaæ"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_90, OpositionT_91), 0, 1, sPromptStrings10)
                                Case "ko³o ³añcuchowe"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_100, OpositionT_101), 0, 1, sPromptStrings11)
                                Case "ko³o zêbate"
                                    oSSDef = oDoc.SketchedSymbolDefinitions.Item(ws)
                                    oTitleSketch1 = oDoc.ActiveSheet.SketchedSymbols.Add(oSSDef, oTG.CreatePoint2d(OpositionT_110, OpositionT_111), 0, 1, sPromptStrings12)
                            End Select
                            xcc += 1
                            ProgressBar2.Value = w + wTitleblock_progress + xss + xgg + css + xcc
                        Next
                    End If
                End If


                ' element w inventorze, który musi byc zaznaczony i wstawiony wymaga selecta. Tzn. Musimy go znale¿æ i zaznaczyæ. Dlatego stosujemy pêtle.
                ' W ten sposób nie dzia³a  oTitleBlock1 = oDoc.ActiveSheet.AddTitleBlock(oDoc.TitleBlockDefinitions.Item(12), , sPromptStrings)
                ' ani ten    Call oTargetSheet.AddTitleBlock(oDoc.TitleBlockDefinitions.Item(12), , sPromptStrings)
                If RadioButton9.Checked = True Then
                    Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
                        cn.Open()
                        Dim xkl As Integer = 0
                        Dim tableTs As String()
                        Dim SQLcommand As New SQLite.SQLiteCommand
                        For Each oboxD In oDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes
                            xbb += 1
                            xkl += 1
                            ProgressBar2.Value = w + wTitleblock_progress + xss + xgg + css + xcc + xbb
                            ' Rows to test a name of textboxes
                            ' MsgBox(oboxD.Text)
                            Select Case oboxD.Text
                                Case "NazwiskoName/Name"
                                    'MsgBox(oDoc.ActiveSheet.TitleBlock.GetResultText(oboxD))
                            End Select
                            ' For Each dRow As DataRow In dt.Rows
                            'Insert Record into TranslateBase
                            ' --------------
                            ' Wyjatki w translacji. Nie wiadomo dlaczego tak czyta.
                            Dim defDescription2 As String() = oboxD.Text.Split(New Char() {"/"c, ":"c})
                            If "Nazwa z³o¿enia/czêœci/ / Benennung / Name of assembly/part drawing" = oboxD.Text Then
                                defDescription2(0) = "Nazwa z³o¿enia/czêœci/"
                            End If
                            If "Konstruowa³ Gezeichn./Design." = oboxD.Text Then
                                defDescription2(0) = "Konstruowa³"
                            End If
                            If "Sprawdzi³ Geprueft/Checked" = oboxD.Text Then
                                defDescription2(0) = "Sprawdzi³"
                            End If
                            If "Sprawdzi³ Geprueft/Checked" = oboxD.Text Then
                                defDescription2(0) = "Sprawdzi³"
                            End If
                            If "SZT.STK.QTY." = oboxD.Text Then
                                defDescription2(0) = "SZT."
                            End If
                            If "POZ.POS.ITEM" = oboxD.Text Then
                                defDescription2(0) = "POZ."
                            End If
                            If "Iloœæ ark.:Blatt-GesamtNo.of sheets" = oboxD.Text Then
                                defDescription2(0) = "Iloœæ ark"
                            End If

                            If "<TS>" = oboxD.Text Then

                                '
                                Dim defDescriptionTS As String() = ts.Split(New Char() {"/"c})

                                'Translation_Message.Show()
                                'Translation_Message.RichTextBox1.Clear()
                                'Translation_Message.RichTextBox1.AppendText(invPLCell3.Value)

                                tableTs = Nothing
                                '   MsgBox("s")
                                Dim n_s_1 = 0
                                Dim n_s_2 = 0
                                Dim n_s_3 = 0
                                Dim n_s_4 = 0
                                Dim n_s_5 = 0
                                Dim ng As Integer
                                Dim ngh As Integer
                                Dim Iss
                                Dim ch11
                                Dim ch21
                                Dim mv1 As Integer
                                Dim mvv1 As Integer
                                Dim remem1 As String = Nothing
                                Dim s1 As Integer = 0
                                Dim hs1 As Integer = 0
                                For gNum As Integer = 0 To defDescriptionTS.GetUpperBound(0)
                                    remem1 = Nothing
                                    n_s_1 = 0
                                    n_s_2 = 0
                                    n_s_4 = 0
                                    ngh = 0
                                    If defDescriptionTS(gNum) <> "" Then
                                        For Iss = 1 To Len(defDescriptionTS(gNum))
                                            n_s_3 = Mid(defDescriptionTS(gNum), Iss, 1)
                                            If n_s_3 = " " Then
                                                n_s_1 += 1
                                            End If
                                            If n_s_3 <> " " Then
                                                n_s_2 += 1
                                            End If
                                            If n_s_1 > n_s_2 And n_s_2 < 2 Then
                                                n_s_4 = n_s_1 '- 1
                                                'ch1 = Mid(descriptDNFormTextSlas(gNum), I + fg, 1)
                                                ' Else
                                                'ch1 = Mid(descriptDNFormTextSlas(gNum), I, 1)
                                            End If

                                        Next
                                        For n_s_5 = 1 To Len(defDescriptionTS(gNum))
                                            ch11 = Mid(defDescriptionTS(gNum), n_s_5 + n_s_4, 1)
                                            If n_s_5 > 2 Then
                                                ch21 = Mid(defDescriptionTS(gNum), n_s_5 + n_s_4 - 1, 1)
                                            End If

                                            If ch11 = " " And n_s_5 > 3 And Len(defDescriptionTS(gNum)) >= n_s_5 Then
                                                ngh += 1
                                            End If
                                            If ch11 + ch21 = "  " Then
                                                ' pozycja z dwoma chr (32). 
                                                If mv1 < 3 Then
                                                    mv1 += 1
                                                    mvv1 = n_s_5 + n_s_4
                                                End If
                                                If mv1 > 0 Then
                                                    ''   hs1 = 1
                                                    ' remem1 = Nothing

                                                End If
                                            End If
                                            If mv1 > 0 Or ngh < 2 Then
                                                If ch11 <> " " Then
                                                    ch11 = Mid(defDescriptionTS(gNum), n_s_5 + n_s_4, 1)
                                                    If ngh >= 1 Then
                                                        ngh = 0
                                                    End If
                                                Else
                                                    If ngh > 1 Then
                                                        'ch11 = Mid(defDescriptionPartlist(gNum), n_s_5 + n_s_4, 1)
                                                        ch11 = Nothing
                                                    End If
                                                    Dim ch31 = Nothing
                                                    If Len(defDescriptionTS(gNum)) > n_s_5 + n_s_4 Then
                                                        ch31 = Mid(defDescriptionTS(gNum), n_s_5 + n_s_4 + 1, 1)
                                                    End If
                                                    If ngh >= 1 And Len(defDescriptionTS(gNum)) <> n_s_5 + n_s_4 And ch31 <> " " Then
                                                        'ch11 = Mid(defDescriptionPartlist(gNum), n_s_5 + n_s_4, 1)
                                                        '   MsgBox(Len(defDescriptionPartlist(gNum)))
                                                        ch11 = " "
                                                    Else
                                                        ch11 = Nothing
                                                    End If
                                                    If ngh > 0 And Len(defDescriptionTS(gNum)) = n_s_5 + n_s_4 Then
                                                        ch11 = Nothing
                                                    End If
                                                End If

                                                ' remem = remem & ch1 '+ Chr(10)
                                            End If

                                            Dim zlozenie As String = Nothing
                                            If n_s_5 <= Len(defDescriptionTS(gNum)) Then
                                                'If iv = Len(descriptDNFormTextSlas(gNum)) Then s += 1 '
                                                ReDim Preserve tableTs(gNum + s1 + hs1)
                                                'If ch1 = descriptDNFormTextSlas(gNum).ToUpper Then
                                                Translation_Message.TextBox1.Focus()
                                                'remem1 = remem1 & ch11
                                                ' jezeli tekst bêdzie zawiera³ / wtedy mo¿e byæ w tabeli wiecej ni¿ 3 stringi
                                                ' wtedy poprzez pêtle dopisuje do stringa 3 pozosta³e informacje
                                                '---
                                                If Len(defDescriptionTS(gNum)) <> n_s_5 Or ch11 <> "" Then ' And ch11 <> " " Then

                                                    remem1 = remem1 & ch11
                                                    'End If
                                                    If defDescriptionTS.GetUpperBound(0) > 3 Then

                                                        If gNum = 2 And n_s_5 = Len(defDescriptionTS(gNum)) - 1 Then
                                                            For vbn As Integer = 3 To defDescriptionTS.GetUpperBound(0)
                                                                zlozenie = zlozenie & "/" & defDescriptionTS(vbn)
                                                            Next
                                                            remem1 = remem1 + Chr(32) + zlozenie
                                                        End If
                                                    End If
                                                Else
                                                    If ch11 = "" Then
                                                        remem1 = remem1 '& ch11
                                                    End If
                                                End If
                                                '-------
                                                ' Analizowany tekst zapisuje do tabeli i generuje tekst do txt_PartLIst
                                                tableTs(gNum + s1 + hs1) = remem1
                                                'End If


                                            End If

                                        Next

                                    End If
                                Next gNum
                                defDescription2(0) = tableTs(0).ToString & Chr(32)
                            End If


                            SQLcommand = cn.CreateCommand

                            '  Dim dt As New Data.DataTable()
                            SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL like '" & defDescription2(0) & "' "
                            'SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL='" + defDescription(0).ToString + "' "
                            'SQLcommand.CommandText = "SELECT PL,DE FROM TranslateBase"
                            Dim lrd As IDataReader = SQLcommand.ExecuteReader()
                            ' Dim SQLreader As System.Data.SqlClient.SqlDataReader = SQLcommand.ExecuteReader()
                            ' Next
                            Dim f As Integer
                            While lrd.Read()

                                Dim sName As String
                                Dim sName3 As String
                                Select Case ComboBox2.Text.ToString
                                    Case "PL- Polish"
                                        sName = "PL"
                                    Case "DE - German"
                                        sName = "DE"
                                    Case "EN  - English"
                                        sName = "EN"
                                    Case "HU - Hungarian"
                                        sName = "HU"
                                    Case "RU - Russian"
                                        sName = "RU"
                                    Case "CZ - Czech"
                                        sName = "CZ"
                                    Case "SLOV - Slovian"
                                        sName = "SLOV"
                                End Select
                                Select Case ComboBox3.Text.ToString
                                    Case "PL- Polish"
                                        sName3 = "PL"
                                    Case "DE - German"
                                        sName3 = "DE"
                                    Case "EN  - English"
                                        sName3 = "EN"
                                    Case "HU - Hungarian"
                                        sName3 = "HU"
                                    Case "RU - Russian"
                                        sName3 = "RU"
                                    Case "CZ - Czech"
                                        sName3 = "CZ"
                                    Case "SLOV - Slovian"
                                        sName3 = "SLOV"
                                End Select
                                Dim sname1 As String
                                If ComboBox1.Text = "PL - Polish" Then sname1 = "PL"
                                Dim txt_ As String
                                Dim txt_1 As String
                                Dim text_Tr As String
                                'MsgBox(lrd.GetValue(f))
                                txt_ = lrd(sname1.ToString) & lrd(sName.ToString) & lrd(sName3.ToString)
                                Select Case oboxD.Text
                                    'Formatowanie


                                    Case "Nazwisko/Name/Name"
                                        ' "Nazwisko" & Chr(10) & " Name/Name"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & "/" & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Konstruowa³ Gezeichn./Design."
                                        '"Konstruowa³ Gezeichn./Design."
                                        txt_ = lrd(sname1.ToString) & Chr(32) & lrd(sName.ToString) & "/" & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Sprawdzi³ Geprueft/Checked"
                                        '"Sprawdzi³ Geprueft/Checked"
                                        txt_ = lrd(sname1.ToString) & Chr(32) & lrd(sName.ToString) & "/" & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Zatwierdzi³/ Bestaetigt/ Approved"
                                        '"Zatwierdzi³ Bestaetigt" & Chr(10) & "Approved"
                                        txt_ = lrd(sname1.ToString) & Chr(32) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Data/ Datum/Date"
                                        '"Data" & Chr(10) & " Datum/Date"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & "/" & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Ciê¿ar/Gewicht/Weight"
                                        ' "Ciê¿ar" & Chr(10) & "Gewicht" & Chr(10) & "Weight"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Format: Blattgroesse: Size:"
                                        ' "Format: Blattgroesse:" & Chr(10) & "Size:"
                                        txt_ = lrd(sname1.ToString) & ":" & Chr(10) & lrd(sName.ToString) & ":" & Chr(10) & lrd(sName3.ToString) & ":"
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Skala/ Masstab/Scale"
                                        ' "Skala" & Chr(10) & "Masstab" & Chr(10) & "Scale"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Zmiana:Index:Revision:"
                                        '"Zmiana:" & Chr(10) & "Index:" & Chr(10) & "Revision:"
                                        txt_ = lrd(sname1.ToString) & ":" & Chr(10) & lrd(sName.ToString) & ":" & Chr(10) & lrd(sName3.ToString) & ":"
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Nazwa z³o¿enia/czêœci/ / Benennung / Name of assembly/part drawing"
                                        ' "Nazwa z³o¿enia/czêœci/ / Benennung / Name of assembly/part drawing"
                                        txt_ = lrd(sname1.ToString) & "/" & lrd(sName.ToString) & "/" & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Iloœæ ark.:Blatt-GesamtNo.of sheets"
                                        '"Iloœæ ark.:" & Chr(10) & "Blatt-Gesamt" & Chr(10) & "No.of Sheets"
                                        txt_ = lrd(sname1.ToString) & ".:" & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)
                                    Case "Nr rysunku / Zeichnungs-Nr. / Drawing-No."
                                        ' "Nr rysunku / Zeichnungs-Nr. / Drawing-No."
                                        txt_ = lrd(sname1.ToString) & "/" & lrd(sName.ToString) & "/" & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "<TS> "
                                        ts = oDoc.ActiveSheet.TitleBlock.GetResultText(oboxD)
                                        txt_ = lrd(sname1.ToString) & "/" & Chr(32) & lrd(sName.ToString) & "/" & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)
                                    Case "<TS>"
                                        Dim textWF As String = Nothing
                                        ts = oDoc.ActiveSheet.TitleBlock.GetResultText(oboxD)


                                  

                                        'txt_1 = lrd("PL") & "/" & Chr(32) & lrd("DE") & Chr(32) & "/" & Chr(32) & lrd("EN")
                                        'txt_ = lrd(sname1.ToString) & "/" & lrd(sName.ToString) & "/" & lrd(sName3.ToString)
                                        'text_Tr = Replace(ts, txt_1, txt_)
                                        ts = tableTs(0) & Chr(32) & "/" & tableTs(1) & Chr(32) & "/" & tableTs(2) & Chr(32)
                                        'txt_PartLIstD = lrd("PL") & "/" & lrd(sName.ToString) & "/" & lrd(sName3.ToString)
                                        txt_1 = lrd("PL") & "/" & lrd("DE") & "/" & lrd("EN")
                                        txt_ = lrd(sname1.ToString) & "/" & Chr(32) & lrd(sName.ToString) & "/" & Chr(32) & lrd(sName3.ToString) & Chr(32)
                                        text_Tr = Replace(ts, txt_1, txt_)
                                        If text_Tr = ts Then
                                            Translation_Message.ListBox2.Items.Clear()
                                            Translation_Message.ListBox3.Items.Clear()
                                            Translation_text.TextBox3.Clear()
                                            Translation_text.TextBox3.AppendText(ts & Chr(10))

                                            ' dodawanie formatki z zapisywaniem nowych danych.
                                            ' Musi byc dodana druga formatka, aby dane zapisaæ do tej piwerwszej i z tej pierwszej mia³a pobraæ pêtla.

                                            Translation_text.ShowDialog()

                                            Translation_text.Hide()

                                            For VlistBox1 As Integer = 0 To Translation_Message.ListBox2.Items.Count - 1

                                                Translation_Message.ListBox3.Items.Add(Translation_Message.ListBox2.Items(VlistBox1).ToString)
                                                'MsgBox("2")
                                            Next '
                                            textWF = ""
                                            If textWF = "" Then
                                                For VlistBox2 As Integer = 0 To Translation_Message.ListBox3.Items.Count - 1

                                                    textWF = textWF + Translation_Message.ListBox3.Items(VlistBox2).ToString
                                                    'MsgBox("2")
                                                Next
                                            End If
                                            text_Tr = textWF
                                        End If
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, text_Tr)
                                    Case "Nr.artyku³u: Artikel-Nr: Article-No:"
                                        ' "Nr.artyku³u:" & Chr(10) & "Artikel-Nr " & Chr(10) & "Article-N"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Dzia³/Abteilung/Division"
                                        '"Dzia³" & Chr(10) & "Abteilung/Division"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & "/" & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Ogólne tolerancje wykonania: Allgemeintoleranzen: General tolerances:"
                                        ' "Ogólne tolerancje wykonania: Allgemeintoleranzen:" & Chr(10) & "General tolerances"
                                        txt_ = lrd(sname1.ToString) & ":" & Chr(10) & lrd(sName.ToString) & ":" & Chr(10) & lrd(sName3.ToString) & ":"
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "JEDN./ 1 EIN/ ITEM"
                                        '"JEDN." & Chr(10) & "1:EIN" & Chr(10) & "ITEM"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)
                                    Case "MATERIA£/ WERKSTOFF/ MATERIAL"
                                        ' "MATERIA£" & Chr(10) & "WERKSTOFF" & Chr(10) & "MATERIAL"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "NR-RYSUNKU/ ZCHN.-NR/ DRAWING NO"
                                        '"NR-RYSUNKU" & Chr(10) & "ZCHN.-NR" & Chr(10) & "DRAWING NO"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "NAZWA/ BENENNUNG/ DESCRIPTION"
                                        ' "NAZWA" & Chr(10) & " BENENNUNG" & Chr(10) & "DESCRIPTION"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "SZT.STK.QTY."
                                        '"SZT." & Chr(10) & "STK." & Chr(10) & "QTY."
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "POZ.POS.ITEM"
                                        ' "POZ." & Chr(10) & "POS." & Chr(10) & "ITEM"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "CALK./ GESAMT/ TOTAL"
                                        ' "CALK." & Chr(10) & "GESAMT" & Chr(10) & "TOTAL"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "NR-NORMY/STANDARD/STANDARD"
                                        ' "NR-NORMY" & Chr(10) & "STANDARD" & Chr(10) & "STANDARD"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "MASA/MASSE/WEIGHT"
                                        ' "MASA/MASSE/WEIGHT"
                                        txt_ = lrd(sname1.ToString) & "/" & lrd(sName.ToString) & "/" & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)

                                    Case "Arkusz-Nr.:Blatt-Nr.:Drg.sheet No.:"
                                        ' "Arkusz-Nr.:" & Chr(10) & "Blatt-Nr.:" & Chr(10) & "Drg.sheet No.:"
                                        txt_ = lrd(sname1.ToString) & Chr(10) & lrd(sName.ToString) & Chr(10) & lrd(sName3.ToString)
                                        Call oDoc.ActiveSheet.TitleBlock.SetPromptResultText(oboxD, txt_)


                                End Select
                                ' Call oDoc.Sheets(1).TitleBlock.SetPromptResultText(obox, txt_)
                                ' MsgBox(Convert.ToString(lrd("DE")))
                                'txt_update_description.Text = SQLreader("description")
                                f += 0
                            End While
                            ' Translation modul for a text in Multilanguage table.
                            ' Modu³ t³umaczenia tekstu dla tabelki multilanguage
                            ' For Each obox In oDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes
                            'MsgBox(obox.Text)
                            SQLcommand.Dispose()
                            w1 += 1
                            'Next
                        Next
                        cn.Close()
                    End Using
                End If
                '----------------------------------
                'Revision table
                '-----------------------------------
                If oDoc.ActiveSheet.RevisionTables.Count > 0 And RadioButton8.Checked = True Then
                    Dim titleR As String = oDoc.ActiveSheet.RevisionTables.Item(1).Title 'historia zmian
                    Dim descriptionR As String = oDoc.ActiveSheet.RevisionTables.Item(1).RevisionTableColumns.Item(2).Title 'discription
                    Dim DataR As String = oDoc.ActiveSheet.RevisionTables.Item(1).RevisionTableColumns.Item(3).Title 'date
                    Dim nameR As String = oDoc.ActiveSheet.RevisionTables.Item(1).RevisionTableColumns.Item(4).Title 'name
                    'Dim titlR As String() = titleR.Split(New Char() {"/"c})
                    Dim titlR(1) As String
                    titlR(0) = "HISTORIA ZMIAN"
                    Dim descriptR As String() = descriptionR.Split(New Char() {"/"c})
                    Dim datR As String() = DataR.Split(New Char() {"/"c})
                    Dim namR As String() = nameR.Split(New Char() {"/"c})
                    ' declere all data into one table
                    Dim tTable(3) As String
                    tTable(0) = titlR(0)
                    tTable(1) = descriptR(0)
                    tTable(2) = datR(0)
                    tTable(3) = namR(0)


                    Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
                        cn.Open()
                        Dim SQLcommand As New SQLite.SQLiteCommand
                        For tabT As Integer = 0 To tTable.GetUpperBound(0)
                            ' progressbar
                            xrev += 1
                            ProgressBar2.Value = w + wTitleblock_progress + xss + xgg + css + xcc + xbb + xrev
                            'Insert Record into TranslateBase

                            SQLcommand = cn.CreateCommand

                            '  Dim dt As New Data.DataTable()
                            SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL like '" & tTable(tabT) & Chr(32) & "' "

                            Dim lrd As IDataReader = SQLcommand.ExecuteReader()
                            ' Dim SQLreader As System.Data.SqlClient.SqlDataReader = SQLcommand.ExecuteReader()
                            ' Next
                            Dim f As Integer
                            While lrd.Read()

                                Dim sName As String = Nothing
                                Dim sName3 As String = Nothing
                                Select Case ComboBox2.Text.ToString
                                    Case "PL- Polish"
                                        sName = "PL"
                                    Case "DE - German"
                                        sName = "DE"
                                    Case "EN  - English"
                                        sName = "EN"
                                    Case "HU - Hungarian"
                                        sName = "HU"
                                    Case "RU - Russian"
                                        sName = "RU"
                                    Case "CZ - Czech"
                                        sName = "CZ"
                                    Case "SLOV - Slovian"
                                        sName = "SLOV"
                                End Select
                                Select Case ComboBox3.Text.ToString
                                    Case "PL- Polish"
                                        sName3 = "PL"
                                    Case "DE - German"
                                        sName3 = "DE"
                                    Case "EN  - English"
                                        sName3 = "EN"
                                    Case "HU - Hungarian"
                                        sName3 = "HU"
                                    Case "RU - Russian"
                                        sName3 = "RU"
                                    Case "CZ - Czech"
                                        sName3 = "CZ"
                                    Case "SLOV - Slovian"
                                        sName3 = "SLOV"
                                End Select
                                Dim sname1 As String = Nothing
                                If ComboBox1.Text = "PL - Polish" Then sname1 = "PL"
                                Dim txt_ As String
                                Dim txt_PL As String
                                'MsgBox(lrd.GetValue(f))
                                txt_ = lrd(sname1.ToString) & "/" & lrd(sName.ToString) & "/" & lrd(sName3.ToString)  '& lrd(sName.ToString)
                                txt_PL = lrd(sname1.ToString)
                                'Dim vv = oDoc.ActiveSheet.RevisionTables.Item(1).Title
                                If oDoc.ActiveSheet.RevisionTables.Item(1).Title = "HISTORIA ZMIAN / AENDERUNGVERMERK / CHANGES" Then
                                    If txt_PL = "HISTORIA ZMIAN " Then
                                        oDoc.ActiveSheet.RevisionTables.Item(1).Title = txt_.ToUpper
                                    End If
                                End If
                                For wRev As Integer = 1 To oDoc.ActiveSheet.RevisionTables.Item(1).RevisionTableColumns.Count
                                    Dim vas = oDoc.ActiveSheet.RevisionTables.Item(1).RevisionTableColumns.Item(wRev).Title

                                    Select Case oDoc.ActiveSheet.RevisionTables.Item(1).RevisionTableColumns.Item(wRev).Title
                                        Case "OPIS/AENDERUNG/DESCRIPTION"
                                            If txt_PL = "OPIS " Then
                                                oDoc.ActiveSheet.RevisionTables.Item(1).RevisionTableColumns.Item(wRev).Title = txt_.ToUpper
                                            End If
                                        Case "DATA/DATUM/DATE"
                                            If txt_PL = "DATA " Then
                                                oDoc.ActiveSheet.RevisionTables.Item(1).RevisionTableColumns.Item(wRev).Title = txt_.ToUpper
                                            End If
                                        Case "NAZWISKO/NAME/NAME"
                                            If txt_PL = "NAZWISKO " Then
                                                oDoc.ActiveSheet.RevisionTables.Item(1).RevisionTableColumns.Item(wRev).Title = txt_.ToUpper
                                            End If
                                    End Select
                                    'oDoc.ActiveSheet.RevisionTables.Item(1).RevisionTableColumns.Item(wRev).Title = "opis"
                                    'MsgBox(oDoc.ActiveSheet.RevisionTables.Item(1).RevisionTableColumns.Item(wRev).Title)

                                Next
                                f += 0
                            End While
                            SQLcommand.Dispose()
                            w1 += 1

                        Next
                        ' Next
                        cn.Close()
                    End Using
                End If
                '---------------------------------
                'custom table
                '-----------------------------------
                If oDoc.ActiveSheet.CustomTables.Count > 0 And RadioButton1.Checked = True Then
                    Dim descriptTC As String() = Nothing
                    Dim descriptDC As String() = Nothing
                    Dim cK As String = Nothing
                    Dim ckj As String = Nothing
                    For wReC As Integer = 1 To oDoc.ActiveSheet.CustomTables.Count
                        descriptTC = oDoc.ActiveSheet.CustomTables.Item(wReC).Title.Split(New Char() {"\"c})
                        descriptDC = oDoc.ActiveSheet.CustomTables.Item(wReC).Columns(1).Title.Split(New Char() {"\"c})
                        cK = oDoc.ActiveSheet.CustomTables.Item(wReC).Title
                        ckj = oDoc.ActiveSheet.CustomTables.Item(wReC).Columns(1).Title
                    Next
                    Dim cTable(1) As String
                    cTable(0) = descriptTC(0)
                    cTable(1) = descriptDC(0)

                    Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
                        cn.Open()
                        Dim SQLcommand As New SQLite.SQLiteCommand
                        For ctabT As Integer = 0 To cTable.GetUpperBound(0)
                            'Insert Record into TranslateBase

                            SQLcommand = cn.CreateCommand

                            '  Dim dt As New Data.DataTable()
                            SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL like '" & cTable(ctabT) & "' "


                            Dim lrdT As IDataReader = SQLcommand.ExecuteReader()
                            ' Dim SQLreader As System.Data.SqlClient.SqlDataReader = SQLcommand.ExecuteReader()
                            ' Next
                            Dim fT As Integer
                            While lrdT.Read()

                                Dim sName As String
                                Dim sName3 As String
                                Select Case ComboBox2.Text.ToString
                                    Case "PL- Polish"
                                        sName = "PL"
                                    Case "DE - German"
                                        sName = "DE"
                                    Case "EN  - English"
                                        sName = "EN"
                                    Case "HU - Hungarian"
                                        sName = "HU"
                                    Case "RU - Russian"
                                        sName = "RU"
                                    Case "CZ - Czech"
                                        sName = "CZ"
                                    Case "SLOV - Slovian"
                                        sName = "SLOV"
                                End Select
                                Select Case ComboBox3.Text.ToString
                                    Case "PL- Polish"
                                        sName3 = "PL"
                                    Case "DE - German"
                                        sName3 = "DE"
                                    Case "EN  - English"
                                        sName3 = "EN"
                                    Case "HU - Hungarian"
                                        sName3 = "HU"
                                    Case "RU - Russian"
                                        sName3 = "RU"
                                    Case "CZ - Czech"
                                        sName3 = "CZ"
                                    Case "SLOV - Slovian"
                                        sName3 = "SLOV"
                                End Select
                                Dim sname1 As String
                                If ComboBox1.Text = "PL - Polish" Then sname1 = "PL"
                                Dim txt_T As String
                                'MsgBox(lrd.GetValue(f))
                                txt_T = lrdT(sname1.ToString) & "\" & Chr(32) & lrdT(sName.ToString) & "\" & lrdT(sName3.ToString) '

                                'For wReC As Integer = 1 To oDoc.ActiveSheet.CustomTables.Count
                                ' Select Case oDoc.ActiveSheet.CustomTables.Item(wReC).Title
                                '   Case "Tabela wykonañ \ Ausführung Tabele \ Version table [ szt \ Stk \ qty]"
                                'MsgBox(oDoc.ActiveSheet.CustomTables.Item(1).Title)

                                If cTable(ctabT) = "Tabela wykonañ " Then
                                    txt_T = lrdT(sname1.ToString) & "\" & Chr(32) & lrdT(sName.ToString) & "\" & Chr(32) & lrdT(sName3.ToString)
                                    oDoc.ActiveSheet.CustomTables.Item(1).WrapAutomatically = True
                                    'MsgBox(Replace(oDoc.ActiveSheet.CustomTables.Item(1).Title, "Tabela wykonañ \ Ausführung Tabele \" & Chr(10) & " Version table", txt_T))
                                    oDoc.ActiveSheet.CustomTables.Item(1).Title = Replace(oDoc.ActiveSheet.CustomTables.Item(1).Title, "Tabela wykonañ \ Ausführung Tabele \" & Chr(10) & " Version table", txt_T)
                                    'oDoc.ActiveSheet.CustomTables.Item(1).ShowTitle = True
                                    oDoc.ActiveSheet.CustomTables.Item(1).Update()
                                End If
                                If cTable(ctabT) = "Poz. " Then
                                    '   Case "Poz. \ Pos. \ Item"
                                    oDoc.ActiveSheet.CustomTables.Item(1).Columns(1).Title = txt_T
                                    'End Select
                                End If
                                ' Next

                                fT += 0
                            End While
                            SQLcommand.Dispose()
                            w1 += 1
                        Next
                        ' Next
                        cn.Close()
                    End Using
                End If
                '-------------------------------------
                'Drawing Notes
                '-----------------------------------------

                ' Rows to test a name of textboxes
                If oDoc.ActiveSheet.DrawingNotes.Count > 0 And RadioButton3.Checked = True Then
                    ' For Each dRow As DataRow In dt.Rows
                    'Insert Record into TranslateBase

                    Translation_Message.Show()
                    Translation_Message.Visible = False
                    '   Translation_Message.Visible = False
                    Dim textW As String
                    Dim remember1 As String = Nothing
                    Dim tableT() As String = Nothing
                    Dim hg As Integer = 0
                    Dim iu As Integer = 0
                    Dim su As Integer = 0
                    Dim l_k As Integer = 0
                    Translation_Message.ListBox1.Items.Clear()

                    For wDeN As Integer = 1 To oDoc.ActiveSheet.DrawingNotes.Count
                        xDN += 1
                        ProgressBar2.Value = w + wTitleblock_progress + xss + xgg + css + xcc + xbb + xrev + xDN
                        If oDoc.ActiveSheet.DrawingNotes.Count > 0 Then
                            Translation_Message.TextBox1.Clear()
                            Translation_Message.ListBox1.Items.Clear()
                            Translation_Message.ListBox2.Items.Clear()
                            Translation_Message.ListBox3.Items.Clear()
                            Translation_Message.ListBox1.Refresh()
                            Translation_Message.TextBox1.Refresh()
                            If oDoc.ActiveSheet.DrawingNotes.Item(wDeN).Text = "Wykonanie / Ausfuhrung / Version W1 - jak na rysunku / wie gezeichnet  / as shown Wykonanie / Ausfuhrung / Version W2 - lustrzane odbicie / spiegelbildlich / mirrored " Then
                                oDoc.ActiveSheet.DrawingNotes.Item(wDeN).Text = "Wykonanie W1 jak na rysunku/ Ausfuhrung W1 wie gezeichnet/Version W1 as  shown;" & Chr(32) & "Wykonanie W2 jako lustrzne odbicie / Ausfuhrung spiegelbildlich/ Version W2 is mirrored"
                            End If


                            Translation_text.TextBox3.Clear()
                            Translation_text.TextBox3.AppendText(oDoc.ActiveSheet.DrawingNotes.Item(wDeN).Text & Chr(10))

                            ' dodawanie formatki z zapisywaniem nowych danych.
                            ' Musi byc dodana druga formatka, aby dane zapisaæ do tej piwerwszej i z tej pierwszej mia³a pobraæ pêtla.

                            Translation_text.ShowDialog()

                            Translation_text.Hide()

                            For VlistBox1 As Integer = 0 To Translation_Message.ListBox2.Items.Count - 1

                                Translation_Message.ListBox3.Items.Add(Translation_Message.ListBox2.Items(VlistBox1).ToString)
                                'MsgBox("2")
                            Next '
                            textW = ""
                            If textW = "" Then
                                For VlistBox2 As Integer = 0 To Translation_Message.ListBox3.Items.Count - 1

                                    textW = textW + Translation_Message.ListBox3.Items(VlistBox2).ToString & Chr(10)
                                    'MsgBox("2")
                                Next
                                oDoc.ActiveSheet.DrawingNotes.Item(wDeN).Text = textW
                            End If


                        End If
                    Next

                    '  Translation_Message.Refresh()



                    Translation_Message.Hide()
                    ' Translation modul for a text in Multilanguage table.
                    ' Modu³ t³umaczenia tekstu dla tabelki multilanguage
                    ' For Each obox In oDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes
                    'MsgBox(obox.Text)
                End If

                '------------------------
                ' PARTLIST
                ' Partlist analyze or perse of the partlist
                ' Translate a text in a table. A short way
                '====================
                '-------------------------------------------
                If oDoc.ActiveSheet.PartsLists.Item(1).PartsListRows.Count > 0 And RadioButton2.Checked = True Then

                    Dim invPLCell1 As PartsListCell
                    Dim invPLCell2 As PartsListCell
                    Dim invPLCell3 As PartsListCell
                    Dim invPLCell4 As PartsListCell
                    Dim invPLCell5 As PartsListCell
                    Dim invPLCell6 As PartsListCell
                    Dim invPLCell7 As PartsListCell
                    Dim invPLCell8 As PartsListCell
                    Dim txt_PartLIstT As String
                    Dim txt_PartLIst As String
                    Dim txt_PartLIstD As String

                    For w3 As Integer = 1 To oDoc.ActiveSheet.PartsLists.Item(1).PartsListRows.Count
                        ' Pasek postêpu dla partlisty
                        xPartlist += 1
                        ProgressBar2.Value = w + wTitleblock_progress + xss + xgg + css + xcc + xbb + xrev + xDN + xPartlist
                        ' Definition a fields to record from partlist. Save a data to database
                        Dim invPLRow = oDoc.ActiveSheet.PartsLists.Item(1).PartsListRows.Item(w3)
                        invPLCell1 = invPLRow.Item(1)
                        invPLCell2 = invPLRow.Item(2)
                        invPLCell3 = invPLRow.Item(3)
                        invPLCell4 = invPLRow.Item(4)
                        invPLCell5 = invPLRow.Item(5)
                        invPLCell6 = invPLRow.Item(6)
                        invPLCell7 = invPLRow.Item(7)
                        invPLCell8 = invPLRow.Item(8)

                        Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
                            cn.Open()
                            Dim SQLcommand As New SQLite.SQLiteCommand

                            ' For Each dRow As DataRow In dt.Rows
                            'Insert Record into TranslateBase
                            ' modu³ do usuwania spacji przed ka¿dym tekstem
                            Dim defDescriptionPartlists As String() = invPLCell3.Value.Split(New Char() {"("c})
                            Dim defDescriptionPartlist As String() = invPLCell3.Value.Split(New Char() {"/"c})
                            Dim defDescriptionPartlist_spacia As String() = invPLCell3.Value.Split(New Char() {" "c})
                            'Translation_Message.Show()
                            'Translation_Message.RichTextBox1.Clear()
                            'Translation_Message.RichTextBox1.AppendText(invPLCell3.Value)
                            Dim tableTs As String() = Nothing

                            '   MsgBox("s")
                            Dim n_s_1 = 0
                            Dim n_s_2 = 0
                            Dim n_s_3 = 0
                            Dim n_s_4 = 0
                            Dim n_s_5 = 0
                            Dim ng As Integer
                            Dim ngh As Integer
                            Dim Iss
                            Dim ch11
                            Dim ch21
                            Dim mv1 As Integer
                            Dim mvv1 As Integer
                            Dim remem1 As String = Nothing
                            Dim s1 As Integer = 0
                            Dim hs1 As Integer = 0
                            For gNum As Integer = 0 To defDescriptionPartlist.GetUpperBound(0)
                                remem1 = Nothing
                                n_s_1 = 0
                                n_s_2 = 0
                                n_s_4 = 0
                                ngh = 0
                                If defDescriptionPartlist(gNum) <> "" Then
                                    For Iss = 1 To Len(defDescriptionPartlist(gNum))
                                        n_s_3 = Mid(defDescriptionPartlist(gNum), Iss, 1)
                                        If n_s_3 = " " Then
                                            n_s_1 += 1
                                        End If
                                        If n_s_3 <> " " Then
                                            n_s_2 += 1
                                        End If
                                        If n_s_1 > n_s_2 And n_s_2 < 2 Then
                                            n_s_4 = n_s_1 '- 1
                                            'ch1 = Mid(descriptDNFormTextSlas(gNum), I + fg, 1)
                                            ' Else
                                            'ch1 = Mid(descriptDNFormTextSlas(gNum), I, 1)
                                        End If

                                    Next
                                    For n_s_5 = 1 To Len(defDescriptionPartlist(gNum))
                                        ch11 = Mid(defDescriptionPartlist(gNum), n_s_5 + n_s_4, 1)
                                        If n_s_5 > 2 Then
                                            ch21 = Mid(defDescriptionPartlist(gNum), n_s_5 + n_s_4 - 1, 1)
                                        End If

                                        If ch11 = " " And n_s_5 > 3 And Len(defDescriptionPartlist(gNum)) >= n_s_5 Then
                                            ngh += 1
                                        End If
                                        If ch11 + ch21 = "  " Then
                                            ' pozycja z dwoma chr (32). 
                                            If mv1 < 3 Then
                                                mv1 += 1
                                                mvv1 = n_s_5 + n_s_4
                                            End If
                                            If mv1 > 0 Then
                                                ''   hs1 = 1
                                                ' remem1 = Nothing

                                            End If
                                        End If
                                        If mv1 > 0 Or ngh < 2 Then
                                            If ch11 <> " " Then
                                                ch11 = Mid(defDescriptionPartlist(gNum), n_s_5 + n_s_4, 1)
                                                If ngh >= 1 Then
                                                    ngh = 0
                                                End If
                                            Else
                                                If ngh > 1 Then
                                                    'ch11 = Mid(defDescriptionPartlist(gNum), n_s_5 + n_s_4, 1)
                                                    ch11 = Nothing
                                                End If
                                                Dim ch31 = Nothing
                                                If Len(defDescriptionPartlist(gNum)) > n_s_5 + n_s_4 Then
                                                    ch31 = Mid(defDescriptionPartlist(gNum), n_s_5 + n_s_4 + 1, 1)
                                                End If
                                                If ngh >= 1 And Len(defDescriptionPartlist(gNum)) <> n_s_5 + n_s_4 And ch31 <> " " Then
                                                    'ch11 = Mid(defDescriptionPartlist(gNum), n_s_5 + n_s_4, 1)
                                                    '   MsgBox(Len(defDescriptionPartlist(gNum)))
                                                    ch11 = " "
                                                Else
                                                    ch11 = Nothing
                                                End If
                                                If ngh > 0 And Len(defDescriptionPartlist(gNum)) = n_s_5 + n_s_4 Then
                                                    ch11 = Nothing
                                                End If
                                            End If

                                            ' remem = remem & ch1 '+ Chr(10)
                                        End If

                                        Dim zlozenie As String = Nothing
                                        If n_s_5 <= Len(defDescriptionPartlist(gNum)) Then
                                            'If iv = Len(descriptDNFormTextSlas(gNum)) Then s += 1 '
                                            ReDim Preserve tableTs(gNum + s1 + hs1)
                                            'If ch1 = descriptDNFormTextSlas(gNum).ToUpper Then
                                            Translation_Message.TextBox1.Focus()
                                            'remem1 = remem1 & ch11
                                            ' jezeli tekst bêdzie zawiera³ / wtedy mo¿e byæ w tabeli wiecej ni¿ 3 stringi
                                            ' wtedy poprzez pêtle dopisuje do stringa 3 pozosta³e informacje
                                            '---
                                            If Len(defDescriptionPartlist(gNum)) <> n_s_5 Or ch11 <> "" Then ' And ch11 <> " " Then
                                                'remem1 = remem1 & ch11
                                                'If gNum = 1 And n_s_5 = Len(defDescriptionPartlist(gNum)) - 1 Then
                                                'If ch11 <> " " Then
                                                '    remem1 = remem1 '& ch11
                                                'Else
                                                '    remem1 = remem1
                                                'End If
                                                'Else
                                                remem1 = remem1 & ch11
                                                'End If
                                                If defDescriptionPartlist.GetUpperBound(0) > 3 Then

                                                    If gNum = 2 And n_s_5 = Len(defDescriptionPartlist(gNum)) - 1 Then
                                                        For vbn As Integer = 3 To defDescriptionPartlist.GetUpperBound(0)
                                                            zlozenie = zlozenie & "/" & defDescriptionPartlist(vbn)
                                                        Next
                                                        remem1 = remem1 + Chr(32) + zlozenie
                                                    End If
                                                End If
                                            Else
                                                If ch11 = "" Then
                                                    remem1 = remem1 '& ch11
                                                End If
                                            End If
                                            '-------
                                            ' Analizowany tekst zapisuje do tabeli i generuje tekst do txt_PartLIst
                                            tableTs(gNum + s1 + hs1) = remem1
                                            'End If


                                        End If

                                    Next

                                End If
                            Next gNum
                            If remem1 = Nothing Then
                                GoTo ENDofloop8
                            End If
                            SQLcommand = cn.CreateCommand
                            ' wyszukiwanie w bazie tekstu
                            '  Dim dt As New Data.DataTable()
                            '  SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL like '" & defDescriptionPartlist(0) & Chr(32) & "' "
                            SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL like '" & tableTs(0) & Chr(32) & "' "
                            Dim lrd As IDataReader = SQLcommand.ExecuteReader()
                            ' Dim SQLreader As System.Data.SqlClient.SqlDataReader = SQLcommand.ExecuteReader()
                            ' Next
                            Dim f As Integer
                            txt_PartLIst = Nothing
                            txt_PartLIstT = Nothing
                            txt_PartLIstD = Nothing
                            While lrd.Read()

                                Dim sName As String = Nothing
                                Dim sName3 As String = Nothing
                                Select Case ComboBox2.Text.ToString
                                    Case "PL- Polish"
                                        sName = "PL"
                                    Case "DE - German"
                                        sName = "DE"
                                    Case "EN  - English"
                                        sName = "EN"
                                    Case "HU - Hungarian"
                                        sName = "HU"
                                    Case "RU - Russian"
                                        sName = "RU"
                                    Case "CZ - Czech"
                                        sName = "CZ"
                                    Case "SLOV - Slovian"
                                        sName = "SLOV"
                                End Select
                                Select Case ComboBox3.Text.ToString
                                    Case "PL- Polish"
                                        sName3 = "PL"
                                    Case "DE - German"
                                        sName3 = "DE"
                                    Case "EN  - English"
                                        sName3 = "EN"
                                    Case "HU - Hungarian"
                                        sName3 = "HU"
                                    Case "RU - Russian"
                                        sName3 = "RU"
                                    Case "CZ - Czech"
                                        sName3 = "CZ"
                                    Case "SLOV - Slovian"
                                        sName3 = "SLOV"
                                End Select
                                Dim sname1 As String
                                If ComboBox1.Text = "PL - Polish" Then sname1 = "PL"

                                ' mo¿na napisaæ program analizujacy jakie sa wpisane jezyki w tabeli. Najlepiej porównuj¹c jeden wyraz z wystêpujacymi w tabeli.
                                'We can write a programe which updated  language are recorded to table.
                                ' txt_PartLIst = tableTs(0) & Chr(32) & "/" & Chr(32) & tableTs(1) & Chr(32) & "/" & Chr(32) & tableTs(2)
                                txt_PartLIst = tableTs(0) & Chr(32) & "/" & tableTs(1) & Chr(32) & "/" & tableTs(2) & Chr(32)
                                'txt_PartLIstD = lrd("PL") & "/" & lrd(sName.ToString) & "/" & lrd(sName3.ToString)
                                txt_PartLIstD = lrd("PL") & "/" & lrd("DE") & "/" & lrd("EN")
                                txt_PartLIstT = lrd(sname1.ToString) & "/" & Chr(32) & lrd(sName.ToString) & "/" & Chr(32) & lrd(sName3.ToString) & Chr(32)
                                ' MsgBox(invPLCell3.Value)

                                '    MsgBox(invPLCell3.Value)
                                ' invPLCell3.Value = Replace(invPLCell3.Value, txt_PartLIst, txt_PartLIstT)
                                'invPLCell3.Value = Replace(invPLCell3.Value, invPLCell3.Value, txt_PartLIstT)
                                invPLCell3.Value = Replace(txt_PartLIst, txt_PartLIstD, txt_PartLIstT)

                                '    f += 0
                            End While

                            SQLcommand.Dispose()
                            cn.Close()
                        End Using
                        Dim textWPartlist As String = Nothing
                        ' otwarcie formularza, je¿eli brak t³umaczenia w bazie
EndOfLoop8:
                        If txt_PartLIst = "" Or invPLCell3.Value = txt_PartLIst Then
                            Translation_Message.Show()
                            Translation_Message.Visible = False
                            Translation_text.TextBox1.Clear()
                            Translation_text.TextBox3.Clear()
                            Translation_text.Visible = False
                            Translation_Message.TextBox1.Clear()
                            Translation_Message.ListBox2.Items.Clear()
                            Translation_Message.ListBox3.Items.Clear()
                            Translation_text.TextBox3.AppendText(invPLCell3.Value)
                            Translation_text.TextBox1.AppendText(invPLCell3.Value)
                            Translation_text.ShowDialog()
                            Translation_text.Hide()
                            ' wype³nianie tekstu jezeli brak tlumaczenia w bazie
                            For VlistBox1 As Integer = 0 To Translation_Message.ListBox2.Items.Count - 1

                                Translation_Message.ListBox3.Items.Add(Translation_Message.ListBox2.Items(VlistBox1).ToString)
                                ' MsgBox("2")
                            Next '
                            For VlistBox2 As Integer = 0 To Translation_Message.ListBox3.Items.Count - 1

                                textWPartlist = textWPartlist + Translation_Message.ListBox3.Items(VlistBox2).ToString ' & Chr(10)

                            Next
                            If textWPartlist = Nothing Then
                                invPLCell3.Value = ""
                            Else
                                invPLCell3.Value = textWPartlist
                            End If
                            ' zamkniecie formularza translatora
                            Translation_Message.Hide()
                        End If
                        '------------------------------------
                        ' Zapis listy czesci do bazy danych
                        ' save up list of the parts
                        ' --------------


                        Dim pozT As Integer = invPLCell1.Value ' pozycja
                        Dim sztT As Integer = invPLCell2.Value ' sztuk
                        Dim NazwaT As String = invPLCell3.Value ' nazwa czêœci/z³o¿enia
                        Dim MaterialT As String = invPLCell4.Value  ' material
                        Dim nr_rysT As String = invPLCell5.Value  'numer rysunku
                        Dim nr_normyT As String = invPLCell6.Value  'numer normy
                        Dim waga_JT As String = Replace(invPLCell7.Value, ",", ".")  'waga jednostkowa czêsci
                        Dim waga_ca³T As String = Replace(invPLCell8.Value, ",", ".")  'waga calkowita czêsci
                        Dim parent_drawT As String = ts ' nazwa g³owna rysunku
                        Dim parent_draw_noT As String = NR ' numer g³owny rysunku 
                        Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
                            cn.Open()
                            Dim SQLcommand As SQLite.SQLiteCommand
                            SQLcommand = cn.CreateCommand
                            SQLcommand.CommandText = "insert into PartList (POZ,SZT,PARENT_DRAW,NAZWA,MATERIAL,NR_RYS,NR_NORMY,WAGA_J,WAGA_CAL,PARENT_DRAW_NO) values(" & "'" & pozT & "'" & "," & "'" & sztT & "'" & "," & "'" & parent_drawT & "'" & "," & "'" & NazwaT & "'" & "," & "'" & MaterialT & "'" & "," & "'" & nr_rysT & "'" & "," & "'" & nr_normyT & "'" & "," & "'" & waga_JT & "'" & "," & "'" & waga_ca³T & "'" & "," & "'" & parent_draw_noT & "'" & ")"
                            SQLcommand.ExecuteNonQuery()
                            SQLcommand.Dispose()
                            'Next
                            cn.Close()
                        End Using
                    Next
                End If


                'Dim oskeychedSymbols As SketchedSymbol = oSheet.SketchedSymbols
                Dim we As Integer
                Dim wer As Integer
                '-------------------------------------------------------
                ' Translation a symbols table
                '----------------------------------------
                'Dim sketchedTable(,) As String = Nothing
                Dim arr As New ArrayList
                Dim r1 As New ArrayList
                Dim nameN As Integer = 0
                Dim textN As Integer = 0


                If oDoc.ActiveSheet.SketchedSymbols.Count > 0 And RadioButton4.Checked = True Then
                    For wer = 1 To oDoc.Sheets(1).SketchedSymbols.Count
                        '  nameN += 1
                        '  MsgBox(oDoc.Sheets(1).SketchedSymbols.Item(wer).Name)
                        For tSketch As Integer = 1 To oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Count
                            '   MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Count)
                            '   textN += 1

                            'If wer = 1 Then ReDim Preserve sketchedTable(0, tSketch - 1)
                            'If wer = 2 Then ReDim Preserve sketchedTable(1, tSketch - 1)

                            ' MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Count)

                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "Cynkowaæ" Then
                                ' sketchedTable(wer - 1, tSketch - 1) = oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name

                                '  MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                Dim ver As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                'Dim defDescriptionSketch As String() = ver.Split(New Char() {" "c, "/"c})
                                Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                    Case "Postac handlowa: Handelsform: Commercial form:"
                                        '    sketchedTable(wer - 1, tSketch - 1) = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                        arr.Add("Postac handlowa:")
                                    Case "Ocynk. / Verzinkt / Galvanized"
                                        '     sketchedTable(wer - 1, tSketch - 1) = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                        arr.Add("Ocynk.")
                                End Select

                            End If
                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "Tolerancja otworów" Then
                                'If tSketch - 1 = 0 Then
                                '    arr.Add(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name)
                                '    '      sketchedTable(wer - 1, tSketch - 1) = oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name
                                'End If
                                'MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                Dim verb As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                    Case "Tolerancja wykonania nieoznaczonych otworów"
                                        '         sketchedTable(wer - 1, tSketch - 1) = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))

                                        arr.Add(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        '12,5
                                        '"Tolerancja wykonania nieoznaczonych otworów"
                                        'oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch), "")
                                        '"Alle nicht bezeichneten Bohrungen alle umbemaßten Schweißnähte"
                                        '
                                        '"Manufacturing tolerance of unmarked holes"
                                        '
                                End Select
                            End If
                            ' "Tolerancja ogólna wykonania"
                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "Tolerancja ogólna wykonania" Then
                                ' sketchedTable(wer - 1, tSketch - 1) = oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name
                                ' MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                Dim verbs As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                    Case "Freimasztoleranzen fuer Bearbeitung"
                                        arr.Add("Stopien dokladnoœci")
                                        arr.Add("Tolerancja ogolna wykonania")
                                        'sketchedTable(wer - 1, tSketch - 1) = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                        '        sketchedTable(wer - 1, tSketch - 1) = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                        'oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 1), "t³umaczenie")
                                        '"Middle"
                                        'IT14
                                        '"Mittle"
                                        '"Stopien dokladnoœci"
                                        '"General tolerances for treatment"
                                        '"Degree of accuracy"
                                        '"Genauigkeitsgrad"
                                        '"ISO 2768"
                                        '"Tolerancja ogolna wykonania"



                                        ' oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 2), "t³umaczenie2")
                                        'oDoc.ActiveSheet.SketchedSymbols.Item(wer).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 3), "t³umaczenie4")
                                End Select
                            End If

                            ' "Tolerancja ogólna dla konstrukcji spawanych"
                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "Tolerancja ogólna dla konstrukcji spawanych" Then
                                '     MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                Dim verbs As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                    Case "Genauigkeitsgrad"
                                        arr.Add("Stopien dokladnoœci")
                                        arr.Add("Tolerancja ogolna dla konstrukcji spawanych")
                                        '"Degree of accuracy"
                                        '"B,F   "
                                        '"EN ISO 13920"
                                        '"Stopien dokladnoœci"
                                        '"Tolerancja ogolna dla konstrukcji spawanych"
                                        '"Allgemeintoleranzen  fuer Schweiszkonstruktionen"
                                        '"General tolerances  for welding construction"


                                        'oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 1), "t³umaczenie")
                                        ' oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 2), "t³umaczenie2")
                                        'oDoc.ActiveSheet.SketchedSymbols.Item(wer).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch + 3), "t³umaczenie4")
                                End Select
                            End If
                            '"Tolerancja otworów pasowanych"
                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "Tolerancja otworów pasowanych" Then
                                'MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                Dim verbse As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                    Case "Tolerancja wykonania  otworów pasowanych"
                                        arr.Add(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        '"6,3"
                                        '"Alle nicht bezeichneten Bohrungen alle umbemaßten Schweißnähte"
                                        '"Tolerance of fitting holes"

                                        'oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 1), "t³umaczenie")
                                        ' oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 2), "t³umaczenie2")
                                        ' oDoc.ActiveSheet.SketchedSymbols.Item(wer).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 3), "t³umaczenie4")
                                End Select
                            End If
                            ' "Spoiny nie oznaczone"
                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "Spoiny nie oznaczone" Then
                                '  MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                Dim verbsee As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                    Case "Unmarked seam welding"
                                        arr.Add("Spoiny nieoznaczone spawaæ")
                                        ' arr.Add(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        '"Alle umbemaßten Schweißnähte"
                                        '"Spoiny nieoznaczone spawaæ:"
                                        '"SPOINA"

                                End Select
                            End If
                            ' "Krawêdzie"
                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "Krawêdzie" Then
                                'MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                Dim vert = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                    Case "Krawedzie zalamac"
                                        arr.Add("Krawêdzie za³amaæ")
                                        ' arr.Add(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        ' value of the fill
                                        'o
                                End Select
                            End If
                            ' "Momenty Dokrêcania Œrub"
                            Dim vertr As String() = Nothing
                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "Momenty Dokrêcania Œrub" Then
                                ' MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                Dim vertrf = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                    Case "Momenty dokrêcania œrub podane w Nm dla typowych gwintów ze œrednim wspó³czynnikiem  tarcia równym 0,12 (g³adka powierzchnia smarowana lub sucha)/ Angeben in Nm fuer Regelgewinde bei einer mittleren Gleitreibungszahl  von 0,12 (gute Oberflaeche geschmiert oder trocken)/ Tightening torque data in Nm for regular type of screw thread with an average friction  factor of 0,12 (smooth surface lubricated or not)"
                                        vertr = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add(vertr(0))
                                        '"558"
                                        '392
                                        '127
                                        '23,1
                                        '9,5
                                        '5,5
                                        '194
                                        '8,8
                                        '34
                                        '14
                                        '117
                                        '68
                                        '8,1
                                        '10,9
                                        '80
                                        '46
                                        'm20
                                        'm16
                                        'm14
                                        '286
                                        '186
                                    Case "Momenty dokrêcania dla œrub ze ³bem np: DIN 912/933 itd./ Anzugsmomente fuer Schrauben mit Kopf z.B. DIN 912/933 usw./ Tightening torque for screws with head e.g. DIN 912/933 etc."
                                        'vertr = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        'arr.Add(vertr(0))
                                        arr.Add("Momenty dokrêcania dla œrub ze ³bem np: DIN 912/933 itd.")
                                        'm6
                                        'm12
                                        'm10
                                        'm8
                                        'm5
                                    Case "/Klasa œrub /Festigkeitsklasse /Strenght category"
                                        vertr = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})

                                        'arr.Add(vertr(1))
                                        arr.Add("Klasa œrub")
                                        'oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 1), "t³umaczenie")
                                        ' oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 2), "t³umaczenie2")
                                        ' oDoc.ActiveSheet.SketchedSymbols.Item(wer).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch + 3), "t³umaczenie4")
                                End Select
                            End If
                            ' "Kierunek ruchu"
                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "Kierunek ruchu" Then
                                MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                Dim vertrt As String() = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                '"Kierunek transportu/ Foerderrichtung/ Conveying direction"
                                'arr.Add(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                arr.Add(vertrt(0))
                            End If
                            '"40 HMT"
                            Dim o As String()
                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "40 HMT" Then
                                '  For tSketch9 As Integer = 1 To oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Count
                                '      MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))

                                Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))

                                    Case "Uwaga / Bemerkung / Note:"
                                        o = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add("Uwaga")
                                        ' oDoc.ActiveSheet.SketchedSymbols.Item(wer).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch), "t³umaczenie") 'opis
                                    Case "40HMT - materia³ ulepszony cieplnie 32÷40 HRC/ 40HMT - Vergüten 32÷40 HRC/ 40HMT - quenched material 32÷40 HRC"
                                        o = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add(o(0))
                                End Select
                                ' Next
                            End If


                            '"ko³o ³añcuchowe"
                            Dim verd As String() = Nothing
                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "ko³o ³añcuchowe" Then
                                'MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                ' = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                    Case "Liczba zêbów/Zähnezahl/Number of teeth"
                                        verd = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add(verd(0))
                                        'oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 1), "t³umaczenie")
                                    Case "Podzia³ka / Graduierung / Graduation"
                                        verd = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add("Podzia³ka")
                                    Case "Typ ³añcucha/Kettentyp/Chain type"
                                        verd = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add(verd(0))

                                    Case "Œrednica podzia³owa/Wirkdurchmesser/Pitch diameter"
                                        verd = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add(verd(0))
                                    Case "Dane ko³a ³añcuchowego/Kettenrad Daten/Sprocket data"
                                        verd = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add("Dane ko³a ³añcuchowego")

                                End Select
                            End If
                            ' "ko³o zêbate"\
                            Dim verds As String()
                            If oDoc.ActiveSheet.SketchedSymbols.Item(wer).Name = "ko³o zêbate" Then
                                '      MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                ' = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch))
                                    Case "Liczba zêbów/Zähnezahl/Number of teeth"
                                        verds = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add(verds(0))
                                    Case "Modu³/Modul/Module"
                                        verds = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add(verds(0))
                                    Case "Œrednica podzia³owa/Wirkdurchmesser/Pitch diameter"
                                        verds = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add(verds(0))
                                    Case "Dane ko³a zêbatego / Stirnzahnrad Daten / Sprocket data"
                                        verds = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add("Dane ko³a zêbatego")
                                    Case "Szerokoœæ zêba/Zahnbreit/Tooth width"
                                        verds = oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)).Split(New Char() {"/"c})
                                        arr.Add(verds(0))
                                End Select
                            End If
                        Next

                    Next
                End If
                Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
                    cn.Open()
                    Dim wer1 As Integer

                    Dim txt_skatchedIstDE As String
                    Dim txt_skatchedIstEN As String
                    Dim txt_skatchedIst As String
                    Dim txt_skatchedIstPL = Nothing
                    If oDoc.ActiveSheet.SketchedSymbols.Count <> 0 Then
                        For arList As Integer = 0 To arr.Count - 1
                            txt_skatchedIstPL = Nothing
                            txt_skatchedIstDE = Nothing
                            txt_skatchedIstEN = Nothing
                            txt_skatchedIst = Nothing
                            Dim SQLcommandS As New SQLite.SQLiteCommand
                            SQLcommandS = cn.CreateCommand
                            '  Dim dt As New Data.DataTable()
                            'MsgBox(arr.Item(arList))
                            '  SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL like '" & defDescriptionPartlist(0) & Chr(32) & "' "
                            SQLcommandS.CommandText = "SELECT * FROM TranslateBase where PL like '" & arr.Item(arList) & Chr(32) & "'" ' or  '" & arr.Item(arList) & " '"
                            Dim lrdS As IDataReader = SQLcommandS.ExecuteReader()
                            ' Dim SQLreader As System.Data.SqlClient.SqlDataReader = SQLcommand.ExecuteReader()
                            ' Next

                            While lrdS.Read()

                                Dim sName As String = Nothing
                                Dim sName3 As String = Nothing
                                Select Case ComboBox2.Text.ToString
                                    Case "PL- Polish"
                                        sName = "PL"
                                    Case "DE - German"
                                        sName = "DE"
                                    Case "EN  - English"
                                        sName = "EN"
                                    Case "HU - Hungarian"
                                        sName = "HU"
                                    Case "RU - Russian"
                                        sName = "RU"
                                    Case "CZ - Czech"
                                        sName = "CZ"
                                    Case "SLOV - Slovian"
                                        sName = "SLOV"
                                End Select
                                Select Case ComboBox3.Text.ToString
                                    Case "PL- Polish"
                                        sName3 = "PL"
                                    Case "DE - German"
                                        sName3 = "DE"
                                    Case "EN  - English"
                                        sName3 = "EN"
                                    Case "HU - Hungarian"
                                        sName3 = "HU"
                                    Case "RU - Russian"
                                        sName3 = "RU"
                                    Case "CZ - Czech"
                                        sName3 = "CZ"
                                    Case "SLOV - Slovian"
                                        sName3 = "SLOV"
                                End Select
                                Dim sname1 As String = Nothing
                                If ComboBox1.Text = "PL - Polish" Then sname1 = "PL"

                                ' mo¿na napisaæ program analizujacy jakie sa wpisane jezyki w tabeli. Najlepiej porównuj¹c jeden wyraz z wystêpujacymi w tabeli.
                                'We can write a programe which updated  language are recorded to table.

                                txt_skatchedIst = lrdS(sname1.ToString) & "/" & lrdS(sName.ToString) & "/" & lrdS(sName3.ToString)
                                txt_skatchedIstPL = lrdS(sname1.ToString)
                                txt_skatchedIstDE = lrdS(sName.ToString)
                                txt_skatchedIstEN = lrdS(sName3.ToString)

                                If "40HMT - materia³ ulepszony cieplnie 32÷40 HRC " = txt_skatchedIstPL Then
                                    txt_skatchedIst = lrdS(sname1.ToString) & "/" & Chr(10) & lrdS(sName.ToString) & "/" & Chr(10) & lrdS(sName3.ToString)

                                End If
                                If "Postac handlowa: " = txt_skatchedIstPL Then
                                    txt_skatchedIst = "/" & lrdS(sname1.ToString) & Chr(10) & "/" & lrdS(sName.ToString) & Chr(10) & "/" & lrdS(sName3.ToString)
                                End If
                                If "Klasa œrub " = txt_skatchedIstPL Then
                                    txt_skatchedIst = "/" & lrdS(sname1.ToString) & Chr(10) & "/" & lrdS(sName.ToString) & Chr(10) & "/" & lrdS(sName3.ToString)
                                End If


                                If "Momenty dokrêcania œrub podane w Nm dla typowych gwintów ze œrednim wspó³czynnikiem  tarcia równym 0,12 (g³adka powierzchnia smarowana lub sucha) " = txt_skatchedIstPL Then
                                    'txt_skatchedIst = "/" & lrdS(sname1.ToString) & Chr(10) & "/" & lrdS(sName.ToString) & Chr(10) & "/" & lrdS(sName3.ToString)
                                    Dim intIndexDE As Integer 'ComboBox2.Text.ToString
                                    Dim intIndexEN As Integer 'ComboBox3.Text.ToString
                                    Dim ISD As Integer
                                    Dim n_D
                                    Dim rembDE As String = Nothing
                                    Dim rembEN As String = Nothing
                                    ' Wyszukiwanie tekstu po którym bêdzie wprowadzony enter. Wartoœæ liczbowa po jakiej wystêpuje ten tekst.
                                    Select Case ComboBox2.Text.ToString
                                        Case "PL- Polish"

                                        Case "DE - German"
                                            intIndexDE = txt_skatchedIstDE.IndexOf("von", 0)
                                        Case "EN  - English"
                                            intIndexDE = txt_skatchedIstDE.IndexOf("friction", 0)
                                        Case "HU - Hungarian"
                                            intIndexDE = txt_skatchedIstDE.IndexOf("friction", 0)
                                        Case "RU - Russian"
                                            intIndexDE = txt_skatchedIstDE.IndexOf("friction", 0)
                                        Case "CZ - Czech"
                                            intIndexDE = txt_skatchedIstDE.IndexOf("friction", 0)
                                        Case "SLOV - Slovian"
                                            intIndexDE = txt_skatchedIstDE.IndexOf("friction", 0)
                                    End Select

                                    Select Case ComboBox3.Text.ToString
                                        Case "PL- Polish"
                                        Case "DE - German"
                                            intIndexEN = txt_skatchedIstEN.IndexOf("von", 0)
                                        Case "EN  - English"
                                            intIndexEN = txt_skatchedIstEN.IndexOf("friction", 0)
                                        Case "HU - Hungarian"
                                            intIndexEN = txt_skatchedIstEN.IndexOf("friction", 0)
                                        Case "RU - Russian"
                                            intIndexEN = txt_skatchedIstEN.IndexOf("friction", 0)
                                        Case "CZ - Czech"
                                            intIndexEN = txt_skatchedIstEN.IndexOf("friction", 0)
                                        Case "SLOV - Slovian"
                                            intIndexEN = txt_skatchedIstEN.IndexOf("friction", 0)
                                    End Select
                                    '   intIndexDE = txt_skatchedIstDE.IndexOf("friction", 0) 'EN ComboBox2.Text.ToString
                                    '           intIndexEN = txt_skatchedIstEN.IndexOf("von", 0) 'EN ComboBox3.Text.ToString
                                    ''
                                    ' ComboBox2.Text.ToString , wprowadzenie entera 
                                    For ISD = 1 To Len(txt_skatchedIstDE)
                                        n_D = Mid(txt_skatchedIstDE, ISD, 1)
                                        If ISD = intIndexDE Then
                                            rembDE = rembDE + Chr(10) + n_D
                                        Else
                                            rembDE = rembDE + n_D
                                        End If

                                    Next
                                    ' ComboBox3.Text.ToString , wprowadzenie entera 
                                    For ISD = 1 To Len(txt_skatchedIstEN)
                                        n_D = Mid(txt_skatchedIstEN, ISD, 1)
                                        If ISD = intIndexEN Then
                                            rembEN = rembEN + Chr(10) + n_D
                                        Else
                                            rembEN = rembEN + n_D
                                        End If

                                    Next
                                    txt_skatchedIst = "/" & "Momenty dokrêcania œrub podane w Nm dla typowych gwintów ze œrednim" & Chr(10) & " wspó³czynnikiem  tarcia równym 0,12 (g³adka powierzchnia smarowana lub sucha) " & Chr(10) & "/" & rembDE & Chr(10) & "/" & rembEN
                                End If

                                If "Momenty dokrêcania dla œrub ze ³bem np: DIN 912/933 itd. " = txt_skatchedIstPL Then
                                    txt_skatchedIst = "/" & lrdS(sname1.ToString) & Chr(10) & "/" & lrdS(sName.ToString) & Chr(10) & "/" & lrdS(sName3.ToString)
                                End If

                                If "Kierunek transportu " = txt_skatchedIstPL Then
                                    txt_skatchedIst = "/" & lrdS(sname1.ToString) & Chr(10) & "/" & lrdS(sName.ToString) & Chr(10) & "/" & lrdS(sName3.ToString)
                                End If
                            End While

                            SQLcommandS.Dispose()
                            For wer1 = 1 To oDoc.Sheets(1).SketchedSymbols.Count

                                'MsgBox(oDoc.Sheets(1).SketchedSymbols.Item(wer).Name)
                                For tSketch As Integer = 1 To oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Count
                                    'MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(1).)
                                    ' progressbar---------------
                                    If arList = 1 Then
                                        If wer1 = 1 Then xwer11 = tSketch
                                        If wer1 = 2 Then xwer21 = tSketch
                                        If wer1 = 3 Then xwer31 = tSketch
                                        If wer1 = 4 Then xwer41 = tSketch
                                        If wer1 = 5 Then xwer51 = tSketch
                                        If wer1 = 6 Then xwer61 = tSketch
                                        If wer1 = 7 Then xwer71 = tSketch
                                        If wer1 = 8 Then xwer81 = tSketch
                                        ProgressBar2.Value = w + wTitleblock_progress + xss + xgg + css + xcc + xbb + xrev + xDN + xPartlist + xwer11 + xwer21 + xwer31 + xwer41 + xwer51 + xwer61 + xwer71 + xwer81
                                    End If
                                    ' ProgressBar2.Value = w + wTitleblock_progress + xss + xgg + css + xcc + xbb + xrev + xDN + xPartlist + xwer11 + xwer21 + xwer31 + xwer41 + xwer51 + xwer61 + xwer71 + xwer81
                                    '--------------------
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "Cynkowaæ" Then
                                        ' MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        Dim ver As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                        If arr.Item(arList) = "Postac handlowa:" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Postac handlowa: Handelsform: Commercial form:" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            ' Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                            ' Case "Postac handlowa:  Handelsform: Commercial form:"
                                            ' If arr.Item(arList) = " Postac handlowa" Then
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If
                                        If arr.Item(arList) = "Ocynk." And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Ocynk. / Verzinkt / Galvanized" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            'Case "Ocynk. / Verzinkt / Galvanized"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)

                                        End If
                                    End If
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "Tolerancja otworów" Then
                                        '  MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        Dim verb As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                        If arr.Item(arList) = "Tolerancja wykonania nieoznaczonych otworów" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                                Case "Tolerancja wykonania nieoznaczonych otworów"
                                                    '12,5
                                                    '"Tolerancja wykonania nieoznaczonych otworów"
                                                    'oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch), "")
                                                    '"Alle nicht bezeichneten Bohrungen alle umbemaßten Schweißnähte"
                                                    oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch + 1), txt_skatchedIstDE)
                                                    '"Manufacturing tolerance of unmarked holes"
                                                    oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch + 2), txt_skatchedIstEN)
                                            End Select
                                        End If
                                    End If
                                    ' "Tolerancja ogólna wykonania"
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "Tolerancja ogólna wykonania" Then
                                        '  MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        Dim verbs As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))

                                        If arr.Item(arList) = "Stopien dokladnoœci" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Stopien dokladnoœci" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            ' Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))

                                            '  Case "Stopien dokladnoœci"
                                            '"Freimasztoleranzen fuer Bearbeitung"
                                            '"Middle"
                                            'IT14
                                            '"Mittle"
                                            '"Stopien dokladnoœci"
                                            '"General tolerances for treatment"

                                            '"Degree of accuracy"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch + 2), txt_skatchedIstEN)
                                            '"Genauigkeitsgrad"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch + 3), txt_skatchedIstDE)
                                            '"ISO 2768"
                                        End If
                                        If arr.Item(arList) = "Tolerancja ogolna wykonania" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Tolerancja ogolna wykonania" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then

                                            '"Tolerancja ogolna wykonania"
                                            '"Freimasztoleranzen fuer Bearbeitung"
                                            'MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch - 9)))
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch - 9), txt_skatchedIstDE)
                                            '"General tolerances for treatment"
                                            'MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch - 4)))
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch - 4), txt_skatchedIstEN)

                                            ' oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 2), "t³umaczenie2")
                                            'oDoc.ActiveSheet.SketchedSymbols.Item(wer).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 3), "t³umaczenie4")

                                        End If
                                    End If
                                    ' "Tolerancja ogólna dla konstrukcji spawanych"
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "Tolerancja ogólna dla konstrukcji spawanych" Then
                                        ' MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        Dim verbs As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))

                                        If arr.Item(arList) = "Stopien dokladnoœci" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Stopien dokladnoœci" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            ' Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                            ' "Genauigkeitsgrad"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch - 4), txt_skatchedIstDE)
                                            '"Degree of accuracy"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch - 3), txt_skatchedIstEN)
                                            '"B,F   "
                                            '"EN ISO 13920"
                                            '"Stopien dokladnoœci"
                                            '"Tolerancja ogolna dla konstrukcji spawanych"
                                        End If
                                        If arr.Item(arList) = "Tolerancja ogolna dla konstrukcji spawanych" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Tolerancja ogolna dla konstrukcji spawanych" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            '"Allgemeintoleranzen  fuer Schweiszkonstruktionen"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch + 1), txt_skatchedIstDE)
                                            '"General tolerances  for welding construction"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch + 2), txt_skatchedIstEN)
                                        End If
                                    End If
                                    '"Tolerancja otworów pasowanych"
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "Tolerancja otworów pasowanych" Then
                                        '  MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        Dim verbse As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                        If arr.Item(arList) = "Tolerancja wykonania  otworów pasowanych" And txt_skatchedIstPL = arr.Item(arList) Then
                                            Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))

                                                Case "Tolerancja wykonania  otworów pasowanych"
                                                    '"6,3"
                                                    '"Alle nicht bezeichneten Bohrungen alle umbemaßten Schweißnähte"
                                                    '"Tolerance of fitting holes"
                                                    oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch + 2), txt_skatchedIstDE)
                                                    oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch + 3), txt_skatchedIstEN)
                                            End Select
                                        End If
                                    End If
                                    ' "Spoiny nie oznaczone"
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "Spoiny nie oznaczone" Then
                                        '  MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        Dim verbsee As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                        If arr.Item(arList) = "Spoiny nieoznaczone spawaæ" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                                Case "Spoiny nieoznaczone spawaæ:"
                                                    '"Alle umbemaßten Schweißnähte"
                                                    '"Spoiny nieoznaczone spawaæ:"
                                                    '"SPOINA"
                                                    oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch - 1), txt_skatchedIstDE)
                                                    oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch - 2), txt_skatchedIstEN)
                                            End Select
                                        End If
                                    End If
                                    ' "Krawêdzie"
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "Krawêdzie" Then
                                        '  MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        Dim vert = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                        If arr.Item(arList) = "Krawêdzie za³amaæ" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            Select Case oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                                Case "Krawedzie zalamac"
                                                    'oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 1), "t³umaczenie")
                                                    oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch + 1), txt_skatchedIstDE)
                                                    oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch + 2), txt_skatchedIstEN)
                                                    ' value of the fill
                                                    'o
                                            End Select
                                        End If
                                    End If
                                    ' "Momenty Dokrêcania Œrub"
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "Momenty Dokrêcania Œrub" Then
                                        '   MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        Dim vertr = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))

                                        If arr.Item(arList) = "Momenty dokrêcania œrub podane w Nm dla typowych gwintów ze œrednim wspó³czynnikiem  tarcia równym 0,12 (g³adka powierzchnia smarowana lub sucha)" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Momenty dokrêcania œrub podane w Nm dla typowych gwintów ze œrednim wspó³czynnikiem  tarcia równym 0,12 (g³adka powierzchnia smarowana lub sucha)/ Angeben in Nm fuer Regelgewinde bei einer mittleren Gleitreibungszahl  von 0,12 (gute Oberflaeche geschmiert oder trocken)/ Tightening torque data in Nm for regular type of screw thread with an average friction  factor of 0,12 (smooth surface lubricated or not)" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                            '"558"
                                            '392
                                            '127
                                            '23,1
                                            '9,5
                                            '5,5
                                            '194
                                            '8,8
                                            '34
                                            '14
                                            '117
                                            '68
                                            '8,1
                                            '10,9
                                            '80
                                            '46
                                            'm20
                                            'm16
                                            'm14
                                            '286
                                            '186
                                        End If
                                        '"Momenty dokrêcania dla œrub ze ³bem np: DIN 912/933 itd./ Anzugsmomente fuer Schrauben mit Kopf z.B. DIN 912/933 usw./ Tightening torque for screws with head e.g. DIN 912/933 etc."
                                        If arr.Item(arList) = "Momenty dokrêcania dla œrub ze ³bem np: DIN 912/933 itd." And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Momenty dokrêcania dla œrub ze ³bem np: DIN 912/933 itd./ Anzugsmomente fuer Schrauben mit Kopf z.B. DIN 912/933 usw./ Tightening torque for screws with head e.g. DIN 912/933 etc." And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                            'm6
                                            'm12
                                            'm10
                                            'm8
                                            'm5
                                        End If
                                        '"/Klasa œrub /Festigkeitsklasse /Strenght category"
                                        If arr.Item(arList) = "Klasa œrub" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "/Klasa œrub /Festigkeitsklasse /Strenght category" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then

                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                            'oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 1), "t³umaczenie")
                                            ' oDoc.ActiveSheet.SketchedSymbols.Item(we).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(1).Definition.Sketch.TextBoxes.Item(tSketch + 2), "t³umaczenie2")
                                            'oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch + 3), "t³umaczenie4")
                                        End If
                                    End If
                                    ' "Kierunek ruchu"
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "Kierunek ruchu" Then
                                        '  Dim vertrtd = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                        If arr.Item(arList) = "Kierunek transportu" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Kierunek transportu/ Foerderrichtung/ Conveying direction" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            Dim vertrt = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))
                                            '"Kierunek transportu/ Foerderrichtung/ Conveying direction"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If
                                    End If
                                    '"40 HMT"
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "40 HMT" Then
                                        '  For tSketch9 As Integer = 1 To oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Count
                                        '  MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        Dim o = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) '.Split(New Char() {":"c, "/"c})
                                        If arr.Item(arList) = "Uwaga" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Uwaga / Bemerkung / Note:" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then

                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If
                                        If arr.Item(arList) = "40HMT - materia³ ulepszony cieplnie 32÷40 HRC" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "40HMT - materia³ ulepszony cieplnie 32÷40 HRC/ 40HMT - Vergüten 32÷40 HRC/ 40HMT - quenched material 32÷40 HRC" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then

                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)

                                            ' Next
                                        End If
                                    End If


                                    '"ko³o ³añcuchowe"
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "ko³o ³añcuchowe" Then
                                        'MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        Dim verd As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))

                                        If arr.Item(arList) = "Liczba zêbów" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Liczba zêbów/Zähnezahl/Number of teeth" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            ' "Liczba zêbów/Zähnezahl/Number of teeth"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If

                                        '"Podzia³ka / Graduierung / Graduation"
                                        If arr.Item(arList) = "Podzia³ka" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Podzia³ka / Graduierung / Graduation" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If

                                        '"Typ ³añcucha/Kettentyp/Chain type"
                                        If arr.Item(arList) = "Typ ³añcucha" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Typ ³añcucha/Kettentyp/Chain type" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If



                                        '"Œrednica podzia³owa/Wirkdurchmesser/Pitch diameter"
                                        If arr.Item(arList) = "Œrednica podzia³owa" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Œrednica podzia³owa/Wirkdurchmesser/Pitch diameter" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If
                                        ' "Dane ko³a ³añcuchowego/Kettenrad Daten/Sprocket data"
                                        If arr.Item(arList) = "Dane ko³a ³añcuchowego" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Dane ko³a ³añcuchowego/Kettenrad Daten/Sprocket data" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If


                                    End If
                                    ' "ko³o zêbate"
                                    If oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Name = "ko³o zêbate" Then
                                        '   MsgBox(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)))
                                        Dim verds As String = oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch))

                                        ' "Liczba zêbów/Zähnezahl/Number of teeth"
                                        If arr.Item(arList) = "Liczba zêbów" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Liczba zêbów/Zähnezahl/Number of teeth" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            ' "Liczba zêbów/Zähnezahl/Number of teeth"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If
                                        '"Modu³/Modul/Module"
                                        If arr.Item(arList) = "Modu³" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Modu³/Modul/Module" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            ' "Liczba zêbów/Zähnezahl/Number of teeth"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If
                                        ' "Œrednica podzia³owa/Wirkdurchmesser/Pitch diameter"
                                        If arr.Item(arList) = "Œrednica podzia³owa" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Œrednica podzia³owa/Wirkdurchmesser/Pitch diameter" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            ' "Liczba zêbów/Zähnezahl/Number of teeth"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If
                                        '"Dane ko³a zêbatego / Stirnzahnrad Daten / Sprocket data"
                                        If arr.Item(arList) = "Dane ko³a zêbatego" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Dane ko³a zêbatego / Stirnzahnrad Daten / Sprocket data" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            ' "Liczba zêbów/Zähnezahl/Number of teeth"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If
                                        '"Szerokoœæ zêba/Zahnbreit/Tooth width"
                                        If arr.Item(arList) = "Szerokoœæ zêba" And oDoc.ActiveSheet.SketchedSymbols.Item(wer1).GetResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch)) = "Szerokoœæ zêba/Zahnbreit/Tooth width" And txt_skatchedIstPL = arr.Item(arList) & Chr(32) Then
                                            ' "Liczba zêbów/Zähnezahl/Number of teeth"
                                            oDoc.ActiveSheet.SketchedSymbols.Item(wer1).SetPromptResultText(oDoc.ActiveSheet.SketchedSymbols.Item(wer1).Definition.Sketch.TextBoxes.Item(tSketch), txt_skatchedIst)
                                        End If
                                    End If
                                Next
                            Next
                        Next
                    End If
                    cn.Close()
                End Using
                ' powiêkszanie rozmaru okna i wyœwitlanie paska pastêpu. Ukrycie okna
                ' increase window size and displaying status bar
                If Me.Height = 646 Then
                    Me.Height = 539
                End If

            End If
            ' Zapisujemy do wybranego katalogu by siê dane nie pomyli³y
            If Form1.TextBox13.Text = "" Then

                MessageBox.Show("Select folder to write a files.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            'MsgBox(Form1.TextBox13.Text)

            If Form1.ListView1.Items(di).Text.Contains(".idw") = True Then
                oDoc.SaveAs(Form1.TextBox13.Text & TextBox1.Text, True)
            End If
EndOfLoop6:
            ' End If
        Next

        oInvApp.Quit()
        oDoc = Nothing
        oInvApp = Nothing
EndOfLoop7:
        'oSourceDoc = Nothing
        'oSourceTB = Nothing
        'oCustomSet = Nothing
    End Sub

    '    Public Function TitleBlockVersion(ByVal VersionNum As String) As Boolean
    '        TitleBlockVersion = False

    '        Dim oDrawDoc As DrawingDocument
    '        oDrawDoc = ThisApplication.ActiveDocument

    '        ' Create the new title block defintion.
    '        Dim oTitleBlockDef As TitleBlockDefinition
    '        On Error GoTo Errorhandler
    '        oTitleBlockDef = oDrawDoc.TitleBlockDefinitions.Item("SMALL") ' this is our standard title block


    '        Dim oSketch As DrawingSketch
    '        Call oTitleBlockDef.Edit(oSketch)

    '        Dim Counter As Integer
    '        Dim Name As String
    '        Dim VersionName As String
    '        VersionName = "TITLE BLOCK VERSION: " & VersionNum

    '        'Loop thru and find the approved by box
    '        For Counter = 1 To oSketch.TextBoxes.Count



    '            Name = oSketch.TextBoxes.Item(Counter).Text
    '            If Name = VersionName Then
    '                TitleBlockVersion = True
    '            End If

    '        Next Counter
    '        Call oTitleBlockDef.ExitEdit(True)
    '        Exit Function

    'Errorhandler:
    '        TitleBlockVersion = False
    '        Exit Function

    '    End Function


    '    Public Sub TitleBlockCopy()

    '        'Get the current path for templates
    '        Dim oApp As Inventor.Application
    '        oApp = ThisApplication
    '        'Debug.Print oApp.FileOptions.TemplatesPath


    '        Dim oCurrentDocument As DrawingDocument
    '        oCurrentDocument = ThisApplication.ActiveDocument

    '        'Check to see if the titleblock version is current
    '        Dim Current As Boolean
    '        Current = TitleBlockVersion("001")

    '        ' quit if already current
    '        If Current Then
    '            Exit Sub
    '        End If

    '        Dim TemplatePath As String
    '        TemplatePath = oApp.FileOptions.TemplatesPath & "Standard.idw"


    '        'Set a reference to the document's active title block name
    '        Dim TitleBlockName As String
    '        TitleBlockName = oCurrentDocument.ActiveSheet.TitleBlock.Name

    '        'Open the template file
    '        Dim oTemplateDocument As DrawingDocument
    '        oTemplateDocument = ThisApplication.Documents.Open(TemplatePath)

    '        'Check to see if the template has the same title block as the currentdocument
    '        Dim RefTitleBlockDef As TitleBlockDefinition
    '        Dim TitleBlockExists As Boolean
    '        TitleBlockExists = False
    '        For Each RefTitleBlockDef In oTemplateDocument.TitleBlockDefinitions
    '            If RefTitleBlockDef.Name = TitleBlockName Then
    '                TitleBlockExists = True
    '                'Debug.Print "Found Title Block"
    '            End If
    '        Next

    '        If Not TitleBlockExists Then
    '            TitleBlockName = "SMALL" ' this is our default title block
    '        End If


    '        ' Get the new source title block definition.
    '        Dim oSourceTitleBlockDef As TitleBlockDefinition
    '        oSourceTitleBlockDef = oTemplateDocument.TitleBlockDefinitions.Item(TitleBlockName)

    '        'Wipe out any references to the existing title block
    '        Dim oSheet As Sheet
    '        oCurrentDocument.Activate()
    '        For Each oSheet In oCurrentDocument.Sheets
    '            oSheet.Activate()
    '            On Error Resume Next
    '            oSheet.TitleBlock.Delete()
    '        Next

    '        'Delete the existing titleblock
    '        On Error Resume Next
    '        oCurrentDocument.TitleBlockDefinitions.Item(TitleBlockName).Delete()

    '        'Copy the Template Title Block to the current file
    '        Dim oNewTitleBlockDef As TitleBlockDefinition
    '        oNewTitleBlockDef = oSourceTitleBlockDef.CopyTo(oCurrentDocument)

    '        oTemplateDocument.Close()

    '        ' Iterate through the sheets.
    '        For Each oSheet In oCurrentDocument.Sheets
    '            oSheet.Activate()
    '            Call oSheet.AddTitleBlock(oNewTitleBlockDef)
    '        Next




    ' End Sub

    ' pasek pomocy tooltip
    Private Sub Button3_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.MouseHover
        ToolTip1.SetToolTip(Button3, "Select Single File or folder with group of the files in browser ")
    End Sub

    Private Sub Button2_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.MouseHover
        ToolTip1.SetToolTip(Button2, "Open DataBase ")
    End Sub
End Class

