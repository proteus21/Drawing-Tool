Imports System.IO
Imports System.Runtime.InteropServices
Imports iTextSharp.text
Imports System
Imports Microsoft.VisualBasic
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports iTextSharp
Imports iTextSharp.text.pdf
Imports iTextSharp.text.xml
Imports PDFTech.PDFOperation
Imports Excel = Microsoft.Office.Interop.Excel
Imports org.pdfbox.pdmodel
Imports org.pdfbox.util
Imports System.Reflection
Imports System.Collections.Generic
Imports QuickPDFAX0813
'Imports iTextSharp.text.pdf.BaseFont
Imports PDFCreator
Imports PrintPDFs.PrintPDFs.StandardAddInServer
Imports System.Environment
Imports System.Net.Dns
Imports System.Runtime.Serialization.Formatters.Binary





Public Class Form1
    Public DirectoryW As String = My.Application.Info.DirectoryPath
    Private ReadOnly LOGIN_DATA_FILE As String = Application.StartupPath + "\" + "login.dat"
    Private WithEvents PDFCreator1 As PDFCreator.clsPDFCreator
    'Dim PDFCreator1 As Object
    Dim oInvApp As Inventor.Application
    Dim ListTab(1) As String

    Private Structure SHFILEINFO
        Public hIcon As IntPtr            ' : icon
        Public iIcon As Integer           ' : icondex
        Public dwAttributes As Integer    ' : SFGAO_ flags
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> _
        Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)> _
        Public szTypeName As String
    End Structure
    'Private mRootPath As String = "c:\"

    'Property RootPath() As String
    '    Get
    '        Return mRootPath
    '    End Get
    '    Set(ByVal value As String)
    '        mRootPath = value
    '        InitializeRoot()
    '    End Set
    'End Property
    'Private Sub InitializeRoot()
    '    ' when our component is loaded, we initialize the TreeView by  adding  the root node
    '    Dim mRootNode As New TreeNode
    '    mRootNode.Text = RootPath
    '    mRootNode.Tag = RootPath
    '    mRootNode.Nodes.Add("*DUMMY*")
    '    TreeView1.Nodes.Clear()
    '    TreeView1.Nodes.Add(mRootNode)
    'End Sub
    Private Declare Auto Function SHGetFileInfo Lib "shell32.dll" _
            (ByVal pszPath As String, _
             ByVal dwFileAttributes As Integer, _
             ByRef psfi As SHFILEINFO, _
             ByVal cbFileInfo As Integer, _
             ByVal uFlags As Integer) As IntPtr

    Private Const SHGFI_ICON = &H100
    Private Const SHGFI_SMALLICON = &H1
    Private Const SHGFI_LARGEICON = &H0    ' Large icon
    Private Const MAX_PATH = 260
    Private nIndex = 0
    'Imports iTextSharp.text.pdf

    Public Function GetFormData(ByVal sourcePdf As String) As Dictionary(Of String, String)
        Dim frmData As New Dictionary(Of String, String)
        Try
            'Open the pdf using pdfreader
            Dim reader As New PdfReader(sourcePdf)
            'Get the form from the pdf
            Dim frm As AcroFields = reader.AcroFields
            'get the fields from the form
            Dim fields As System.Collections.Hashtable = frm.Fields
            'Extract the data from the fields
            Dim data As String = String.Empty
            For Each key As String In fields.Keys
                data = frm.GetField(key)
                frmData.Add(key, data)
            Next
            reader.Close()
        Catch ex As Exception
            Debug.Write(ex.Message)

        End Try
        Return frmData
    End Function


    Private Sub AddImages(ByVal strFileName As String)

        Dim shInfo As SHFILEINFO
        shInfo = New SHFILEINFO()
        shInfo.szDisplayName = New String(vbNullChar, MAX_PATH)
        shInfo.szTypeName = New String(vbNullChar, 80)
        Dim hIcon As IntPtr
        hIcon = SHGetFileInfo(strFileName, 0, shInfo, Marshal.SizeOf(shInfo), SHGFI_ICON Or SHGFI_SMALLICON)
        Dim MyIcon As Drawing.Bitmap
        MyIcon = Drawing.Icon.FromHandle(shInfo.hIcon).ToBitmap
        ImageList1.Images.Add(strFileName.ToString(), MyIcon)
        nIndex = nIndex + 1
    End Sub

    Private Sub AddAllFolders(ByVal TNode As TreeNode, ByVal FolderPath As String)
        Try
            For Each FolderNode As String In Directory.GetDirectories(FolderPath)
                Dim SubFolderNode As TreeNode = TNode.Nodes.Add(FolderNode.Substring(FolderNode.LastIndexOf("\"c) + 1))

                SubFolderNode.Tag = FolderNode

                SubFolderNode.Nodes.Add("Loading...")
                AddImages(SubFolderNode.Tag)
                SubFolderNode.ImageIndex = CInt(nIndex - 1)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


    'Private Sub treeview1_NodeMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
    '    'find out wich button was clicked
    '    If e.Button = Windows.Forms.MouseButtons.Left Then
    '        'Get the DirectoryW Info and File Info
    '        Dim MyTopDirectoryWInfo As System.IO.DirectoryWInfo
    '        Dim MyFileInfo As System.IO.FileInfo

    '        'clear the ListView
    '        ListView1.Items.Clear()
    '        'looping through all the sub directories in the selected DirectoryW and adding them to 
    '        'the listveiw
    '        For Each TopDirectoryW As String In My.Computer.FileSystem.GetDirectories(e.Node.Tag)
    '            MyTopDirectoryWInfo = My.Computer.FileSystem.GetDirectoryWInfo(TopDirectoryW)
    '            'set the "Key", "Text" and "ImageIndex" for the folder to be added
    '            'remember the "1" is the index in our imgSmall and imgLarge for "Folders"
    '            ListView1.Items.Add(MyTopDirectoryWInfo.Name, MyTopDirectoryWInfo.Name, 1)
    '            'set the item to the "Folders" group
    '            'remember the "0" is the index for the "Folders" Group
    '            ListView1.Items(MyTopDirectoryWInfo.Name).Group = ListView1.Groups(0)
    '        Next
    '        'next we will add all the files, but first get the info for the Files in 
    '        'the current DirectoryW by reading the "Tag"
    '        For Each MyFile As String In My.Computer.FileSystem.GetFiles(e.Node.Tag)
    '            MyFileInfo = My.Computer.FileSystem.GetFileInfo(MyFile)
    '            'We are not showing hidden files but you can modify this however needed
    '            If MyFileInfo.Attributes.ToString.Contains(IO.FileAttributes.Hidden.ToString) Then
    '                'hidden files
    '            Else
    '                'add the files to the listview and set the Image Index to "0" for Files
    '                ListView1.Items.Add(MyFile, MyFile, 0)
    '                'our Image Index for our "Group", "Files" is 0
    '                ListView1.Items(MyFile).Group = ListView1.Groups(1)
    '            End If
    '            'MsgBox(MyFile & "-" & MyFileInfo.Attributes.ToString)
    '        Next
    '    End If

    'End Sub

    'Private Sub Treeview1_BeforeExpand(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeExpand

    '    If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "Loading..." Then

    '        e.Node.Nodes.Clear()
    '        AddAllFolders(e.Node, CStr(e.Node.Tag))

    '    End If

    'End Sub
    Private Sub TreeView1_BeforeCollapse(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeCollapse
        ' clear the node that is being collapsed
        e.Node.Nodes.Clear()
        ' add a dummy TreeNode to the node being collapsed so it is expandable
        e.Node.Nodes.Add("*DUMMY*")
    End Sub

    Private Sub TreeView1_BeforeExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeExpand
        ' clear the expanding node so we can re-populate it, or else we end up with duplicate nodes

        e.Node.Nodes.Clear()
        ' get the DirectoryW representing this node
        Dim mNodeDirectoryW As IO.DirectoryInfo
        mNodeDirectoryW = New IO.DirectoryInfo(e.Node.Tag.ToString)
        Try
            ' add each subDirectoryW from the file system to the expanding node as a child node
            For Each mDirectoryW As IO.DirectoryInfo In mNodeDirectoryW.GetDirectories
                ' declare a child TreeNode for the next subDirectoryW
                Dim mDirectoryWNode As New TreeNode
                ' store the full path to this DirectoryW in the child TreeNode's Tag property
                mDirectoryWNode.Tag = mDirectoryW.FullName
                ' set the child TreeNodes's display text
                mDirectoryWNode.Text = mDirectoryW.Name
                ' add a dummy TreeNode to this child TreeNode to make it expandable
                mDirectoryWNode.Nodes.Add("*DUMMY*")
                ' add this child TreeNode to the expanding TreeNode
                e.Node.Nodes.Add(mDirectoryWNode)
                mDirectoryWNode.ImageKey = CacheShellIcon(mDirectoryW.FullName)
                mDirectoryWNode.SelectedImageKey = mDirectoryWNode.ImageKey
                'AddImages(mDirectoryW.FullName)

            Next
            ' add each file from the file system that is a child of the argNode that was passed in
            For Each mFile As IO.FileInfo In mNodeDirectoryW.GetFiles
                ' declare a TreeNode for this file
                Dim mFileNode As New TreeNode
                ' store the full path to this file in the file TreeNode's Tag property
                mFileNode.Tag = mFile.FullName
                ' set the file TreeNodes's display text
                mFileNode.Text = mFile.Name
                mFileNode.ImageKey = Module1.CacheShellIcon(mFile.FullName)
                mFileNode.SelectedImageKey = mFileNode.ImageKey & "-open"
                ' add this file TreeNode to the TreeNode that is being populated
                e.Node.Nodes.Add(mFileNode)
            Next
        Catch ex As Exception
            Debug.Write(ex.Message)

        End Try
    End Sub

    Private Sub TreeView1_NodeMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick
        ' only proceed if the node represents a file
        If e.Node.ImageKey = "folder" Then Exit Sub
        If e.Node.Tag = "" Then Exit Sub
        ' try to open the file
        Try
            Process.Start(e.Node.Tag)
        Catch ex As Exception
            MessageBox.Show("Error opening file: " & ex.Message)
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        'TreeView1.Sort()

        'Dim x As Integer
        ''Start looping through the Drives
        'For x = 0 To My.Computer.FileSystem.Drives.Count - 1
        '    If My.Computer.FileSystem.Drives(x).IsReady = True Then
        '        'add the Drive to the Tree Node use the Drive Name as the "Key" and "Text"
        '        Dim Tnode As TreeNode = TreeView1.Nodes.Add(My.Computer.FileSystem.Drives(x).Name, My.Computer.FileSystem.Drives(x).Name)
        '        ' 
        '        'TreeView1.Nodes.Add(My.Computer.FileSystem.Drives(x).Name, My.Computer.FileSystem.Drives(x).Name)
        '        'set the Tag Property to the Drive Name for Identification Later On
        '        TreeView1.Nodes(My.Computer.FileSystem.Drives(x).Name).Tag = My.Computer.FileSystem.Drives(x).Name
        '        'add the first level of sub directories to the TreeView

        '        For Each SubDirectoryW As String In My.Computer.FileSystem.GetDirectories(My.Computer.FileSystem.Drives(x).Name)
        '            'The Mid Function is used so our Node does not include something like
        '            '"c:\Windows" it should rather read something like "Windows".
        '            'However the Key (in our case the first part of the Add() will
        '            'have the whole path. This will be used later for Finding the 
        '            'Sub Directories)
        '            'AddAllFolders(Tnode, "c:\")
        '            TreeView1.Nodes(x).Nodes.Add(SubDirectoryW, Mid(SubDirectoryW, 4))
        '            'Here we add the Whole path to the Tag Property for Identification
        '            'later on
        '            TreeView1.Nodes(x).Nodes(SubDirectoryW).Tag = SubDirectoryW

        '        Next
        '    End If
        'Next
        'Dim Tnode1 As TreeNode = TreeView1.Nodes.Add("(Drive C:)")
        'AddAllFolders(Tnode1, "c:\")
        ListView1.View = View.Details
        '' Add a column with width 80 and left alignment
        ListView1.Columns.Add("File Name", 500, HorizontalAlignment.Left)
        ListView1.Columns.Add("File Size", 80, HorizontalAlignment.Right)
        ListView1.Columns.Add("File Type", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("Date Modified", 250, HorizontalAlignment.Left)

        'mRootNode.ImageKey = CacheShellIcon(RootPath)
        'mRootNode.SelectedImageKey = mRootNode.ImageKey
        ' when our component is loaded, we initialize the TreeView by  adding  the root node

        'Counter for our Physical Drives
        Dim x As Integer, y As Integer = 0
        'Start looping through the Drives
        For x = 0 To My.Computer.FileSystem.Drives.Count - 1
            'make sure the drive is ready
            If My.Computer.FileSystem.Drives(x).IsReady = True Then
                'add the Drive to the Tree Node use the Drive Name as the "Key" and "Text"
                TreeView1.Nodes.Add(My.Computer.FileSystem.Drives(x).Name, My.Computer.FileSystem.Drives(x).Name)
                'set the Tag Property to the Drive Name for Identification Later On
                TreeView1.Nodes(My.Computer.FileSystem.Drives(x).Name).Tag = My.Computer.FileSystem.Drives(x).Name
                'add the first level of sub directories to the TreeView
                '========
                ''AddImages(My.Computer.FileSystem.Drives(x).Name)- lub przez podanie z funkcji
                TreeView1.Nodes(My.Computer.FileSystem.Drives(x).Name).ImageKey = CacheShellIcon(My.Computer.FileSystem.Drives(x).Name)
                For Each SubDirectoryW As String In My.Computer.FileSystem.GetDirectories(My.Computer.FileSystem.Drives(x).Name)
                    Try
                        TreeView1.Nodes(y).Nodes.Add(SubDirectoryW, Mid(SubDirectoryW, 4))
                        'Here we add the Whole path to the Tag Property for Identification  
                        'later on                     
                        TreeView1.Nodes(y).Nodes(SubDirectoryW).Tag = SubDirectoryW
                        '  InitializeRoot()
                        '======== obie funkcje sa do wyswietlania ikon
                        TreeView1.Nodes(y).Nodes(SubDirectoryW).ImageKey = CacheShellIcon(SubDirectoryW)
                        ''  AddImages(SubDirectoryW)
                    Catch ex As Exception

                    End Try
                Next

                y += 1
            End If
        Next
        ' Try
        '    Dim VistaSecurity
        '    Dim key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\PDFTechnologies\PDFToText", True)
        '    If key IsNot Nothing AndAlso CInt(key.GetValue("NewVersionAvailable")) = 1 Then
        '        If Not VistaSecurity.IsVistaOrHigher() AndAlso Not VistaSecurity.IsAdmin() Then
        '            VistaSecurity.RestartElevated()
        '        End If
        '        Me.Cursor = Cursors.WaitCursor
        '        If File.Exists(New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll_") Then
        '            File.Delete(New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll")

        '            File.Move(New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll_", New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll")
        '        End If
        '        key.SetValue("NewVersionAvailable", 0)
        '        Me.Cursor = Cursors.[Default]
        '        MessageBox.Show("The updated version is successfully installed.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    End If
        'Catch ex As Exception
        '    Me.Cursor = Cursors.[Default]
        '    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        'End Try
        'Dim mRootNode As New TreeNode
        ''mRootNode.Text = RootPath
        ''mRootNode.Tag = RootPath
        ''mRootNode.Nodes.Add("*DUMMY*")
        ''TreeView1.Nodes.Add(mRootNode)
        'mRootNode.ImageKey = CacheShellIcon(RootPath)
        'mRootNode.SelectedImageKey = mRootNode.ImageKey
        ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        ListView1.Columns(0).Width = 500
        ListView1.Columns(1).Width = 80
        ListView1.Columns(2).Width = 100
        ListView1.Columns(3).Width = 250
        ' Wskazuje nazwe komputera
        ToolStripStatusLabel4.Text = Environment.MachineName
        ' Wsjkazuje ip Adress komputera lokalnego
        Dim hostInfo As System.Net.IPHostEntry = _
        System.Net.Dns.GetHostByName(System.Net.Dns.GetHostByName("LocalHost").HostName)
        Dim ipaddr As Byte() = hostInfo.AddressList(0).GetAddressBytes

        ToolStripStatusLabel6.Text = ipaddr(0) & "." & ipaddr(1) & "." & ipaddr(2) & "." & ipaddr(3)
        'MessageBox.Show("IP Address:" & ipaddr(0) & "." & ipaddr(1) & "." & _
        ' ipaddr(2) & "." & ipaddr(3) & "" & Microsoft.VisualBasic.Chr(10) & "")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Shared Function GetHFCell(ByVal header As String, ByVal footer As String, ByVal text As String) As PdfPTable
        Dim pdft As PdfPTable
        Dim hc As PdfPCell
        Dim fc As PdfPCell

        pdft = New PdfPTable(1)
        pdft.WidthPercentage = 100.0F
        pdft.DefaultCell.Border = 0

        hc = New PdfPCell(New Phrase(header))
        hc.Top = 0.0F
        hc.FixedHeight = 7.0F
        hc.HorizontalAlignment = 1
        hc.BackgroundColor = iTextSharp.text.BaseColor.BLUE
        DirectCast(hc.Phrase(0), Chunk).Font = New iTextSharp.text.Font(DirectCast(hc.Phrase(0), Chunk).Font.Family, 5.0F)

        fc = New PdfPCell(New Phrase(footer))
        hc.Top = 0.0F
        fc.FixedHeight = 7.0F
        hc.HorizontalAlignment = 1
        fc.BackgroundColor = iTextSharp.text.BaseColor.YELLOW
        DirectCast(fc.Phrase(0), Chunk).Font = New iTextSharp.text.Font(DirectCast(fc.Phrase(0), Chunk).Font.Family, 5.0F)

        pdft.AddCell(hc)
        pdft.AddCell(text)
        pdft.AddCell(fc)

        Return pdft
    End Function

    'Public Sub GeneratePDF()
    'Dim document As New Document()
    '    Try
    '        PdfWriter.GetInstance(document, New FileStream("File1.pdf", FileMode.Create))

    '        document.Open()

    'Dim table As New PdfPTable(5)
    '        table.DefaultCell.Padding = 0
    '        table.DefaultCell.BorderWidth = 2.0F
    '        For j As Integer = 1 To 5
    '            For i As Integer = 1 To 5
    ''calling GetHFCell
    '                table.AddCell(GetHFCell("header " & CInt(i + 5 * (j - 1)).ToString(), "footer " & CInt(i + 5 * (j - 1)).ToString(), "z" & j.ToString() & i.ToString()))
    '            Next
    '        Next

    '        document.Add(table)
    ''...
    '    Catch de As DocumentException
    ''...
    '    Catch ioe As IOException
    '    End Try
    '    document.Close()
    'End Sub


    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        AboutBox1.Show()
    End Sub


    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dialog1.ShowDialog()
        Dim dr As DialogResult = Dialog1.DialogResult
        If dr = Windows.Forms.DialogResult.Cancel Then
            MessageBox.Show("Open Part List form  canceled.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            GoTo LoopS
        End If

        If dr = Windows.Forms.DialogResult.OK Then
            Dim nFiles, h As Integer
            'Dim rect As util.RectangleJ
            'Dim rectA4_1, rectA4_2, rectA4_3, rectA4_4 As util.RectangleJ
            'Dim rectA3_1, rectA3_2, rectA3_3, rectA3_4, rectA3_5 As util.RectangleJ
            'Dim rectA2_1, rectA2_2, rectA2_3, rectA2_4, rectA2_5 As util.RectangleJ
            'Dim rectA1_1, rectA1_2, rectA1_3, rectA1_4, rectA1_5 As util.RectangleJ
            'Dim rectA0_1, rectA0_2, rectA0_3, rectA0_4, rectA0_5 As util.RectangleJ
            'Dim Out_Name_rys As String
            Dim myFileDialog As OpenFileDialog = New OpenFileDialog()
            Dim xlApp As Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet

            xlApp = New Excel.ApplicationClass
            xlApp.Visible = True
            If Dialog1.RadioButton1.Checked = True Then
                xlWorkBook = xlApp.Workbooks.Open(Filename:=My.Application.Info.DirectoryPath + "\" + "Zeichnungsliste.xls")
                xlWorkSheet = xlWorkBook.Worksheets("Zeichnungsliste (2)")

                'display the cells value B2
                '=============================================
                ' w pierwszej kolejności należy poslugiwać się itextem. Jeżeli w nazwie znajdzie mi wartość pustą musi przejść innym programem
                '===============================================
                Dim folder As String = Nothing
                'Dim zlicz2 As Integer = 0
                Dim b As Integer = 0
                Dim sTable As String() = Nothing  '"" ' item number 1 -numer rysunku
                Dim sTableq As String() = Nothing ' / ' item number 1 -numer rysunku
                Dim zlicz2 As Integer = 0

                For di As Integer = 0 To ListView1.Items.Count - 1
                    If ListView1.Items(di).Selected = False Then
                        ListView1.Items(di).Tag = ListView1.Items(di).Text
                        folder = ListTab(0) + "\" + CStr(ListView1.Items(di).Tag)
                        ListTab(0) = ListTab(0).ToString '+ "\" + CStr(ListView1.Items(di).Tag)
                        ' wyszukiwanie numeru rysunku w bazie / parent draw _no
                        Dim parent_draw_NO As String = Microsoft.VisualBasic.Left(CStr(ListView1.Items(di).Tag), Len(CStr(ListView1.Items(di).Tag)) - 4)
                        Dim name_file As String() = CStr(ListView1.Items(di).Tag).Split(" ")
                        Dim name_fileS As String = Microsoft.VisualBasic.Left(name_file(0), Len(name_file(0)) - 4)
                        Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
                            cn.Open()
                            Dim SQLcommand As New SQLite.SQLiteCommand
                            SQLcommand = cn.CreateCommand

                            '  Dim dt As New Data.DataTable()
                            ' SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL like '" & defDescription2(0) & "' "
                            SQLcommand.CommandText = "SELECT * FROM PartList where PARENT_DRAW_NO='" + name_file(0) + "' "
                            'SQLcommand.CommandText = "SELECT PL,DE FROM TranslateBase"
                            Dim lrd As IDataReader = SQLcommand.ExecuteReader()
                            Dim txt_1 As String
                            'MsgBox(lrd.GetValue(f))
                            DataGridView1.Rows.Clear()
                            'Tracks the current record number
                            Dim ns As Integer = 0
                            Dim i As Integer
                            While lrd.Read()

                                'MsgBox(lrd.Item(1).ToString)
                                ns = DataGridView1.Rows.Add()
                                MsgBox(lrd.Item(0).ToString & lrd.Item(1).ToString & lrd.Item(2).ToString & lrd.Item(3).ToString & lrd.Item(4).ToString)
                                DataGridView1.Rows.Item(ns).Cells(0).Value = lrd.Item(0).ToString
                                DataGridView1.Rows.Item(ns).Cells(1).Value = lrd.Item(1).ToString
                                DataGridView1.Rows.Item(ns).Cells(2).Value = lrd.Item(2).ToString
                                DataGridView1.Rows.Item(ns).Cells(3).Value = lrd.Item(3).ToString
                                DataGridView1.Rows.Item(ns).Cells(4).Value = lrd.Item(4).ToString
                                DataGridView1.Rows.Item(ns).Cells(5).Value = lrd.Item(5).ToString
                                DataGridView1.Rows.Item(ns).Cells(6).Value = lrd.Item(6).ToString
                                DataGridView1.Rows.Item(ns).Cells(7).Value = lrd.Item(7).ToString

                            End While
                            SQLcommand.Dispose()

                            cn.Close()
                        End Using

                        '   zlicz2 += 1
                        Dim n As Integer = 0
                        '  Try
                        If Not folder Is Nothing AndAlso IO.Directory.Exists(folder) Then

                            For Each file As String In IO.Directory.GetFiles(folder)
                                If Microsoft.VisualBasic.Right(file, 3) = "pdf" Then

                                    ' -------------------------------------------------------------------------------------
                                    ' Bibliteka ItexSharp
                                    ' Sprawdzanie tekstu w locie - dla standardowych plików pdf
                                    '---------------------------------------------------------------------------------------
                                    RichTextBox1.Clear()
                                    '    Dim oReader As New iTextSharp.text.pdf.PdfReader(ListView1.Items(nFiles).ImageKey)
                                    Dim oReader As New iTextSharp.text.pdf.PdfReader(file)
                                    Dim i As Integer
                                    Dim sOut = ""
                                    Dim sOut1 = ""
                                    Dim ss = ""
                                    Dim sf As String
                                    Dim strText As String = ""
                                    Try
                                        For i = 1 To oReader.NumberOfPages
                                            Dim its As New iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
                                            ' odczyt tekstu w pliku
                                            sOut1 &= iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(oReader, i, its)

                                        Next
                                    Catch es As Exception
                                        ' Let the user know what went wrong.
                                        Console.WriteLine("The file could not be read:")
                                        Console.WriteLine(es.Message)
                                    End Try
                                    ' parseUsingPDFBox(folder)

                                    ' Display the text
                                    ' Kodowanie do polskich znaków
                                    Dim helvetica As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1250, BaseFont.EMBEDDED)
                                    Dim plLarge As New iTextSharp.text.Font(helvetica, 16)
                                    ' Dekodowanie pliku do UTF8 - brak wszystkich polskich znaków
                                    sOut = (Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(sOut1))))
                                    'RichTextBox1.AppendText((Environment.NewLine + sOut.ToString()))
                                    'sf = DirectCast(Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.[Default], Encoding.UTF8, Encoding.[Default].GetBytes(sOut1))), String)
                                    ' Zapis do Stringbuildera z uwzględnieniem polskich znaków
                                    Dim sb As New StringBuilder()
                                    sb.AppendFormat(sOut, plLarge) ' do przetestowania
                                    'Dim resultsb As String = Encoding.GetEncoding("ISO-8859-2").GetString(System.Text.Encoding.GetEncoding("ISO-8859-2").GetBytes(sOut)) ' polskie znaki
                                    'MsgBox(resultsb)
                                    Dim line As String()
                                    Dim args As String
                                    line = (sOut.Split(ControlChars.Lf))

                                    'Dodanie nowych lini odczytanych po zakonczniu tekstu
                                    For Each item As String In line
                                        RichTextBox1.AppendText((item & ControlChars.Lf))
                                        '   TextBox7.Text += item & ControlChars.Lf
                                    Next
                                    If sOut = Nothing Then
                                        RichTextBox2.AppendText(("Plik jest obrazem" & "-" & folder & ControlChars.Lf))
                                        xlWorkSheet.Cells(11 + b, 6) = "Plik jest obrazem" & "-" & folder
                                    Else

                                        ''  RichTextBox1.AppendText((sOut.ToString())) - bez polskich znaków
                                        ''RichTextBox1.AppendText((Environment.NewLine + sb.ToString())) '- z polskimi znakami
                                        'RichTextBox1.AppendText((resultsb.ToString())) '- bez polskich znaków

                                        Dim xf As Integer = 0
                                        Dim x As Integer
                                        Dim siteK() As String
                                        Dim sit() As String
                                        ' Odczyt z richboxa i zapis do tablicy  której wysukuje się dane.
                                        For x = RichTextBox1.Lines.GetLowerBound(0) To RichTextBox1.Lines.GetUpperBound(0)
                                            ' MessageBox.Show(RichTextBox1.Lines(x))

                                            ReDim Preserve siteK(xf)
                                            ReDim Preserve sit(xf)
                                            siteK(xf) = RichTextBox1.Lines(x)
                                            xf = xf + 1
                                        Next
                                        ' Wyszukiwanie tekstu i zapisywanie do pliku xls
                                        'Dim result1 As New List(Of Integer)
                                        Dim result1(4) As String
                                        'Dim result1 As List(Of Integer) = New List(Of Integer)
                                        For ig As Integer = 0 To siteK.Length - 1
                                            ' określa pozycje wystąpienia znaku
                                            If siteK(ig).Contains("Zmiana:") Then result1(0) = ig ' item number 1 -numer rysunku
                                            If siteK(ig).Contains("Zatwierdził ") Then result1(1) = ig 'item number 2 - rozmiar rysunku
                                            If siteK(ig).Contains("Nr.artykułu") Then result1(2) = ig 'item number 3 - waga częsci then
                                            If siteK(ig).Contains("Ciężar") Then result1(3) = ig 'item number 4 - nazwa rysunku
                                            If siteK(ig).Contains("Format:") Then result1(4) = ig 'item number 5 - index zmian


                                            'If siteK(ig).Contains("Nr rysunku /Drawing-No. /RAJZSZÁM") Then result1.Add(ig) ' item number 1 -numer rysunku
                                            'If siteK(ig).Contains("Zatwierdził") Or siteK(ig).Contains("Zatwierdzi Bestaetigt") Or siteK(ig).Contains("ZatwierdziBestaetigt") Then result1.Add(ig) 'item number 2 - rozmiar rysunku
                                            'If siteK(ig).Contains("Approved") Then result1.Add(ig) 'item number 3 - waga częsci then
                                            'If siteK(ig).Contains("MASA/WEIGHT/HMOTNOSŤ") Then result1.Add(ig) 'item number 4 - nazwa rysunku
                                            'If siteK(ig).Contains("WERKSTOFF") Or siteK(ig).Contains("MATERIA WERKSTOFF") Then result1.Add(ig) 'item number 5 - index zmian

                                        Next
                                        Dim pos As Integer = result1(0).ToString ' item number 1 -numer rysunku
                                        Dim pos1 As Integer = result1(1).ToString 'item number 2 - rozmiar rysunku
                                        Dim pos2 As Integer = result1(2).ToString 'item number 3 - waga częsci
                                        Dim pos3 As Integer = result1(3).ToString 'item number 4 - nazwa rysunku
                                        Dim pos4 As Integer = result1(4).ToString 'item number 5 - index zmian


                                        sTable = siteK(pos + 4).Split(" ") ' item number 1 -numer rysunku
                                        sTableq = siteK(pos + 4).Split("/") ' item number 1 -numer rysunku


                                        xlWorkSheet.Cells(11 + b, 1) = b + 1 ' Numer pozycji 

                                        If Microsoft.VisualBasic.Right(ListView1.Items(di).Text, 3) = "dwg" Then
                                            xlWorkSheet.Cells(11 + b, 2) = sTable(0) ' numer rysunku dwg - number of drawing
                                        Else
                                            'xlWorkSheet.Cells(11 + b, 2) = "---------------"
                                        End If



                                        xlWorkSheet.Cells(11 + b, 3) = sTable(0) '  numer rysunku  pdf
                                        ' dodanie hiperzłącza - hiperlinks
                                        'MsgBox(".." & Microsoft.VisualBasic.Right(ListTab(0), Len(ListTab(0)) - 2) & "\" & sTable(0) & ".pdf")
                                        '=HIPERŁĄCZE(ZŁĄCZ.TEKSTY("..\3. Zeichnungen\";G68;"\";H68;".pdf");H68)

                                        'xlWorkSheet.Cells(11 + b, 3) = "=HIPERŁĄCZE(ZŁĄCZ.TEKSTY(""..\3. Zeichnungen\"";G68;""\"";H68;"".pdf"");H68)"
                                        'xlWorkSheet.Cells(11 + b, 3).Hyperlinks.Add(Anchor:=xlWorkSheet.Cells(11 + b, 3), Address:=info, SubAddress:="", TextToDisplay:=sTable(0))
                                        'xlWorkSheet.Cells(11 + b, 3).Hyperlinks.Add(Anchor:=xlWorkSheet.Cells(11 + b, 3), Address:=folder, SubAddress:="", TextToDisplay:=sTable(0))

                                        'xlWorkSheet.Cells(11 + b, 3).Value = "..\wieszak\Nowy folder\204-0720-007-00000.pdf"
                                        'xlWorkSheet.Cells(11 + b, 3).Hyperlinks.Add(xlWorkSheet.Cells(11 + b, 3), xlWorkSheet.Cells(11 + b, 3).Value)
                                        xlWorkSheet.Cells(11 + b, 3).Hyperlinks.add(Anchor:=xlWorkSheet.Cells(11 + b, 3), Address:="", SubAddress:= _
        "", TextToDisplay:="this is long text")
                                        If siteK(pos4 + 23) = "General tolerances:" Then
                                            xlWorkSheet.Cells(11 + h, 4) = ""
                                        Else
                                            If siteK(pos4 + 24) = "  " Then
                                                xlWorkSheet.Cells(11 + h, 4) = siteK(pos4 + 24)
                                            Else
                                                Dim NameReadT9 = siteK(pos3 + 7).Split("/")
                                                Dim Lsle As Integer = NameReadT9.GetUpperBound(0)
                                                If Lsle > 1 Then
                                                    xlWorkSheet.Cells(11 + h, 4) = siteK(pos4 + 22) ' index zmian
                                                Else
                                                    xlWorkSheet.Cells(11 + h, 4) = siteK(pos4 + 23)
                                                End If
                                            End If
                                        End If
                                        xlWorkSheet.Cells(11 + b, 5) = siteK(pos1 - 3) ' rozmiar rysunku- drawing size


                                        Dim NazwaReadTs, fRead As Integer
                                        Dim sread As Integer = 0
                                        Dim lNameReadTs
                                        Dim NameReadT As Integer = Len(siteK(pos3 + 8)) ' odlicza długość pierwszego wyrazu /  /
                                        Dim NameReadTs = siteK(pos3 + 8).Split("/")
                                        If NameReadTs.GetUpperBound(0) > 1 Then ' jeżeli dla pozycji p3 występuje  ciąg  ///

                                            lNameReadTs = Len(NameReadTs(2))
                                            'MsgBox(Mid(readTab(resultT.Item(0) + 1), 1, NameReadT - NazwaReadTs - 1))
                                            xlWorkSheet.Cells(11 + b, 6) = (siteK(pos3 + 8)) 'nazwa części - part name
                                        Else ' jeżeli dla pozycji p3 występuje  ciąg  /// - wtedy zmieniamy pozycję ciągu w pliku
                                            'Dim NameReadT2 As Integer = Len(siteK(posP3 + 2)) ' odlicza długość pierwszego wyrazu /  /
                                            Dim NameReadT2 = siteK(pos3 + 7).Split("/")
                                            Dim Lslesh As Integer = NameReadT2.GetUpperBound(0)
                                            If Lslesh > 0 And Lslesh <= 1 Then
                                                'Dim NameReadTs22 As Integer = Len(NameReadTs2(2))
                                                'Dim NameReadTs_ = siteK(posP3 + 2).Split(" ")
                                                'Dim NameReadTs_2 As Integer = Len(NameReadTs_(0))
                                                xlWorkSheet.Cells(11 + b, 6) = siteK(pos3 + 2) & siteK(pos3 + 3) ', 1 + NameReadTs_2, NameReadT2 - NameReadTs22 - NameReadTs_2) 'nazwa części - part name
                                                If Lslesh > 0 And Lslesh <= 2 Then
                                                    xlWorkSheet.Cells(11 + b, 6) = siteK(pos3 + 2) & siteK(pos3 + 3) & siteK(pos3 + 4)
                                                End If
                                            Else
                                                Dim NameReadT3() = siteK(pos3 + 7).Split(New Char() {"/"c})
                                                Dim Lslesh1 As Integer = NameReadT3.GetUpperBound(0)
                                                Dim rsitek1 As String = Nothing
                                                Dim rsitek2 As String = Nothing
                                                Dim dg As Integer = 0
                                                Dim dh As Integer = 0
                                                If Lslesh1 > 0 And Lslesh1 <= 2 Then
                                                    Dim NameReadT4() = siteK(pos3 + 7).Split(New Char() {" "c, "/"c})

                                                    For bd As Integer = 1 To NameReadT4.GetUpperBound(0)
                                                        If NameReadT4(bd) <> "" Then
                                                            dg += 1
                                                            If dg = 1 Then rsitek1 = NameReadT4(bd)
                                                            If dg = 2 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT4(bd)
                                                            If dg = 3 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT4(bd)
                                                            If dg > 3 Then rsitek1 = rsitek1 + Chr(32) + NameReadT4(bd)
                                                        End If
                                                        'If NameReadT3(2) = "" Then
                                                        '    dh += 1
                                                        '    Dim NameReadT5() = NameReadT3(0).Split(New Char() {" "c, "/"c})
                                                        '    For bh As Integer = 1 To NameReadT5.GetUpperBound(0)
                                                        '        If dh = 1 Then rsitek1 = NameReadT5(bd)
                                                        '        If dh = 2 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT5(bd)
                                                        '        If dh = 3 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT5(bd)
                                                        '    Next
                                                        'End If
                                                    Next
                                                    ' MsgBox(rsitek1)
                                                    xlWorkSheet.Cells(11 + b, 6) = rsitek1
                                                    'xlWorkSheet.Cells(11 + b, 6) = siteK(posP3 + 1)
                                                Else
                                                    If siteK(pos3 + 11).Contains("-") Then
                                                        xlWorkSheet.Cells(11 + b, 6) = siteK(pos3 + 8) + siteK(pos3 + 9) + Chr(32) + siteK(pos3 + 10)
                                                    Else
                                                        xlWorkSheet.Cells(11 + b, 6) = siteK(pos3 + 8) + siteK(pos3 + 9)
                                                    End If

                                                End If

                                            End If
                                        End If

                                        Dim ReadTableqW() As String = siteK(pos2 - 5).Split(" ")
                                        If siteK(pos2 - 5).Contains("Kg") Then

                                            If ReadTableqW.GetUpperBound(0) > 0 Then
                                                ReadTableqW = siteK(pos2 - 5).Split(" ")
                                            Else
                                                ReadTableqW = siteK(pos2 - 6).Split(" ")

                                            End If
                                        Else
                                            ReadTableqW = siteK(pos2 - 6).Split(" ")
                                        End If

                                        Dim rwTab() As String = Nothing
                                        Dim fReadWs As Integer
                                        Dim sreads As Integer = 0
                                        For fReadWs = ReadTableqW.GetLowerBound(0) To ReadTableqW.GetUpperBound(0)
                                            If ReadTableqW(fReadWs) <> "Kg" Then
                                                ReDim Preserve rwTab(sreads)
                                                rwTab(sreads) = ReadTableqW(fReadWs)
                                                sreads += 1
                                            End If


                                        Next
                                        Dim memorW As String = rwTab(0)

                                        xlWorkSheet.Cells(11 + b, 7) = memorW 'masa części - weight part - odczyt z pliku i tablicy readTAB

                                        Dim LSplit As String() = folder.Split("\")
                                        Dim vT As Integer = LSplit.GetUpperBound(0)
                                        Dim lPath = Len(LSplit(vT))
                                        xlWorkSheet.Cells(11 + b, 8) = Microsoft.VisualBasic.Right(folder, lPath) 'brak folderu

                                        xlWorkSheet.Cells(11 + b, 9) = sTable(0) '  numer rysunku  pdf do celów inforam
                                        xlWorkSheet.Cells(11 + b, 10) = 1 ' liczba wystąpień
                                        b += 1
                                        h += 1
                                    End If
                                End If
                                If Microsoft.VisualBasic.Right(file, 3) = "dwg" Then
                                    Dim Stab As String() = file.Split("\")
                                    Dim lStab As Integer = Stab.GetUpperBound(0)
                                    xlWorkSheet.Cells(11 + zlicz2, 2) = Stab(lStab) ' numer rysunku dwg - number of drawing
                                    'xlWorkSheet.Cells(11 + zlicz2, 2).Hyperlinks.Add(Anchor:=xlWorkSheet.Cells(11 + zlicz2, 2), Address:="", SubAddress:=folder, TextToDisplay:=CStr(ListView1.Items(di).Tag))
                                    zlicz2 += 1
                                Else
                                    'xlWorkSheet.Cells(11 + zlicz2, 2) = "---------------"
                                End If
                            Next
                            ' For zlicz1 As Integer = 0 To zlicz
                       
                        End If
                        '  End If

                        For nDGw As Integer = 0 To DataGridView1.Rows.Count - 1


                            MsgBox(DataGridView1.Rows.Item(nDGw).Cells(6).Value)

                            If Microsoft.VisualBasic.Right(ListView1.Items(di).Text, 3) = "pdf" Then ' And DataGridView1.Rows.Item(nDGw).Cells(7).Value = name_fileS Then 'And DataGridView1.Rows.Item(nDGw).Cells(11).Value = name_fileS Then

                                ' -------------------------------------------------------------------------------------
                                ' Bibliteka ItexSharp
                                ' Sprawdzanie tekstu w locie - dla standardowych plików pdf
                                '---------------------------------------------------------------------------------------
                                RichTextBox1.Clear()
                                '    Dim oReader As New iTextSharp.text.pdf.PdfReader(ListView1.Items(nFiles).ImageKey)
                                Dim oReader As New iTextSharp.text.pdf.PdfReader(folder)
                                Dim i As Integer
                                Dim sOut = ""
                                Dim sOut1 = ""
                                Dim ss = ""
                                Dim sf As String
                                Dim strText As String = ""
                                Try
                                    For i = 1 To oReader.NumberOfPages
                                        Dim its As New iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
                                        ' odczyt tekstu w pliku
                                        sOut1 &= iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(oReader, i, its)

                                    Next
                                Catch es As Exception
                                    ' Let the user know what went wrong.
                                    Console.WriteLine("The file could not be read:")
                                    Console.WriteLine(es.Message)
                                End Try
                                ' parseUsingPDFBox(folder)

                                ' Display the text
                                ' Kodowanie do polskich znaków
                                Dim helvetica As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1250, BaseFont.EMBEDDED)
                                Dim plLarge As New iTextSharp.text.Font(helvetica, 16)
                                ' Dekodowanie pliku do UTF8 - brak wszystkich polskich znaków
                                sOut = (Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(sOut1))))
                                'RichTextBox1.AppendText((Environment.NewLine + sOut.ToString()))
                                'sf = DirectCast(Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.[Default], Encoding.UTF8, Encoding.[Default].GetBytes(sOut1))), String)
                                ' Zapis do Stringbuildera z uwzględnieniem polskich znaków
                                Dim sb As New StringBuilder()
                                sb.AppendFormat(sOut, plLarge) ' do przetestowania
                                'Dim resultsb As String = Encoding.GetEncoding("ISO-8859-2").GetString(System.Text.Encoding.GetEncoding("ISO-8859-2").GetBytes(sOut)) ' polskie znaki
                                'MsgBox(resultsb)
                                Dim line As String()
                                Dim args As String
                                line = (sOut.Split(ControlChars.Lf))

                                'Dodanie nowych lini odczytanych po zakonczniu tekstu
                                For Each item As String In line
                                    RichTextBox1.AppendText((item & ControlChars.Lf))
                                    '   TextBox7.Text += item & ControlChars.Lf
                                Next
                                If sOut = Nothing Then
                                    RichTextBox2.AppendText(("Plik jest obrazem" & "-" & folder & ControlChars.Lf))
                                    xlWorkSheet.Cells(11 + b, 6) = "Plik jest obrazem" & "-" & folder
                                Else

                                    ''  RichTextBox1.AppendText((sOut.ToString())) - bez polskich znaków
                                    ''RichTextBox1.AppendText((Environment.NewLine + sb.ToString())) '- z polskimi znakami
                                    'RichTextBox1.AppendText((resultsb.ToString())) '- bez polskich znaków

                                    Dim xf As Integer = 0
                                    Dim x As Integer
                                    Dim siteK() As String
                                    Dim sit() As String
                                    ' Odczyt z richboxa i zapis do tablicy  której wysukuje się dane.
                                    For x = RichTextBox1.Lines.GetLowerBound(0) To RichTextBox1.Lines.GetUpperBound(0)
                                        ' MessageBox.Show(RichTextBox1.Lines(x))

                                        ReDim Preserve siteK(xf)
                                        ReDim Preserve sit(xf)
                                        siteK(xf) = RichTextBox1.Lines(x)
                                        xf = xf + 1
                                    Next
                                    ' Wyszukiwanie tekstu i zapisywanie do pliku xls
                                    'Dim result1 As New List(Of Integer)
                                    Dim result1(4) As String
                                    'Dim result1 As List(Of Integer) = New List(Of Integer)
                                    For ig As Integer = 0 To siteK.Length - 1
                                        ' określa pozycje wystąpienia znaku
                                        If siteK(ig).Contains("Zmiana:") Then result1(0) = ig ' item number 1 -numer rysunku
                                        If siteK(ig).Contains("Zatwierdził ") Or siteK(ig).Contains("Zatwierdził") Then result1(1) = ig 'item number 2 - rozmiar rysunku
                                        If siteK(ig).Contains("Nr.artykułu") Then result1(2) = ig 'item number 3 - waga częsci then
                                        If siteK(ig).Contains("Ciężar") Then result1(3) = ig 'item number 4 - nazwa rysunku
                                        If siteK(ig).Contains("JEDN.") Then result1(4) = ig 'item number 5 - index zmian


                                        'If siteK(ig).Contains("Nr rysunku /Drawing-No. /RAJZSZÁM") Then result1.Add(ig) ' item number 1 -numer rysunku
                                        'If siteK(ig).Contains("Zatwierdził") Or siteK(ig).Contains("Zatwierdzi Bestaetigt") Or siteK(ig).Contains("ZatwierdziBestaetigt") Then result1.Add(ig) 'item number 2 - rozmiar rysunku
                                        'If siteK(ig).Contains("Approved") Then result1.Add(ig) 'item number 3 - waga częsci then
                                        'If siteK(ig).Contains("MASA/WEIGHT/HMOTNOSŤ") Then result1.Add(ig) 'item number 4 - nazwa rysunku
                                        'If siteK(ig).Contains("WERKSTOFF") Or siteK(ig).Contains("MATERIA WERKSTOFF") Then result1.Add(ig) 'item number 5 - index zmian

                                    Next
                                    Dim pos As Integer = result1(0).ToString ' item number 1 -numer rysunku
                                    Dim pos1 As Integer = result1(1).ToString 'item number 2 - rozmiar rysunku
                                    Dim pos2 As Integer = result1(2).ToString 'item number 3 - waga częsci
                                    Dim pos3 As Integer = result1(3).ToString 'item number 4 - nazwa rysunku
                                    Dim pos4 As Integer = result1(4).ToString 'item number 5 - index zmian


                                    sTable = siteK(pos + 4).Split(" ") ' item number 1 -numer rysunku
                                    sTableq = siteK(pos + 4).Split("/") ' item number 1 -numer rysunku


                                    xlWorkSheet.Cells(11 + b, 1) = b + 1 ' Numer pozycji 


                                    Dim sTables() = siteK(pos3 + 7).Split(New Char() {" "c, "/"c})
                                    Dim dgg As Integer = 0
                                    Dim dg22 As Integer = 0
                                    Dim rsit As String
                                    For bdg As Integer = 0 To sTables.GetUpperBound(0)
                                        If sTables(bdg) <> "" Then ' And dg > 1 Then
                                            dgg += 1
                                            If dgg > 0 Then
                                                dg22 += 1
                                                If dg22 = 1 Then rsit = sTables(bdg)
                                            End If
                                        End If
                                    Next
                                    If Microsoft.VisualBasic.Right(ListView1.Items(di).Text, 3) = "dwg" Then
                                        xlWorkSheet.Cells(11 + b, 2) = rsit  ' numer rysunku dwg - number of drawing
                                    Else
                                        'xlWorkSheet.Cells(11 + b, 2) = "---------------"
                                    End If
                                    '

                                    '

                                    xlWorkSheet.Cells(11 + b, 3) = rsit  '  numer rysunku  pdf
                                    ' dodanie hiperzłącza - hiperlinks
                                    'MsgBox(".." & Microsoft.VisualBasic.Right(ListTab(0), Len(ListTab(0)) - 2) & "\" & sTable(0) & ".pdf")
                                    '=HIPERŁĄCZE(ZŁĄCZ.TEKSTY("..\3. Zeichnungen\";G68;"\";H68;".pdf");H68)

                                    'xlWorkSheet.Cells(11 + b, 3) = "=HIPERŁĄCZE(ZŁĄCZ.TEKSTY(""..\3. Zeichnungen\"";G68;""\"";H68;"".pdf"");H68)"
                                    'xlWorkSheet.Cells(11 + b, 3).Hyperlinks.Add(Anchor:=xlWorkSheet.Cells(11 + b, 3), Address:=info, SubAddress:="", TextToDisplay:=sTable(0))
                                    'xlWorkSheet.Cells(11 + b, 3).Hyperlinks.Add(Anchor:=xlWorkSheet.Cells(11 + b, 3), Address:=folder, SubAddress:="", TextToDisplay:=sTable(0))
                                    ' xlWorkSheet.Cells(11 + b, 3).Value = "..\wieszak\Nowy folder\204-0720-007-00000.pdf"
                                    ' xlWorkSheet.Cells(11 + b, 3).Hyperlinks.Add(xlWorkSheet.Cells(11 + b, 3), xlWorkSheet.Cells(11 + b, 3).Value)

                                    If siteK(pos4 - 1) = "General tolerances:" Then
                                        xlWorkSheet.Cells(11 + h, 4) = ""
                                    Else


                                        If siteK(pos4 + 24) = "  " Then
                                            xlWorkSheet.Cells(11 + h, 4) = siteK(pos4 + 39)
                                        Else
                                            Dim NameReadT9 = siteK(pos3 + 7).Split("/")
                                            Dim Lsle As Integer = NameReadT9.GetUpperBound(0)
                                            If Lsle > 1 Then
                                                xlWorkSheet.Cells(11 + h, 4) = siteK(pos4 - 1) ' index zmian
                                            Else
                                                xlWorkSheet.Cells(11 + h, 4) = siteK(pos4 - 1)
                                            End If

                                        End If
                                    End If
                                    xlWorkSheet.Cells(11 + b, 5) = siteK(pos1 - 3) ' rozmiar rysunku- drawing size


                                    Dim NazwaReadTs, fRead As Integer
                                    Dim sread As Integer = 0
                                    Dim lNameReadTs
                                    Dim NameReadT As Integer = Len(siteK(pos3 + 8)) ' odlicza długość pierwszego wyrazu /  /
                                    Dim NameReadTs = siteK(pos3 + 8).Split("/")
                                    If NameReadTs.GetUpperBound(0) > 1 Then ' jeżeli dla pozycji p3 występuje  ciąg  ///

                                        lNameReadTs = Len(NameReadTs(2))
                                        'MsgBox(Mid(readTab(resultT.Item(0) + 1), 1, NameReadT - NazwaReadTs - 1))
                                        xlWorkSheet.Cells(11 + b, 6) = siteK(pos3 + 8) + siteK(pos3 + 9) 'nazwa części - part name
                                    Else ' jeżeli dla pozycji p3 występuje  ciąg  /// - wtedy zmieniamy pozycję ciągu w pliku
                                        'Dim NameReadT2 As Integer = Len(siteK(posP3 + 2)) ' odlicza długość pierwszego wyrazu /  /
                                        Dim NameReadT2 = siteK(pos3 + 7).Split("/")
                                        Dim Lslesh As Integer = NameReadT2.GetUpperBound(0)
                                        If Lslesh > 0 And Lslesh <= 1 Then
                                            'Dim NameReadTs22 As Integer = Len(NameReadTs2(2))
                                            'Dim NameReadTs_ = siteK(posP3 + 2).Split(" ")
                                            'Dim NameReadTs_2 As Integer = Len(NameReadTs_(0))
                                            xlWorkSheet.Cells(11 + b, 6) = siteK(pos3 + 2) & siteK(pos3 + 3) ', 1 + NameReadTs_2, NameReadT2 - NameReadTs22 - NameReadTs_2) 'nazwa części - part name
                                            If Lslesh > 0 And Lslesh <= 2 Then
                                                xlWorkSheet.Cells(11 + b, 6) = siteK(pos3 + 2) & siteK(pos3 + 3) & siteK(pos3 + 4)
                                            End If
                                        Else
                                            Dim NameReadT3() = siteK(pos3 + 7).Split(New Char() {"/"c})
                                            Dim Lslesh1 As Integer = NameReadT3.GetUpperBound(0)
                                            Dim rsitek1 As String = Nothing
                                            Dim rsitek2 As String = Nothing
                                            Dim dg As Integer = 0
                                            Dim dg2 As Integer = 0
                                            Dim dg3 As Integer = 0
                                            Dim dh As Integer = 0
                                            If Lslesh1 > 0 And Lslesh1 <= 2 Then
                                                Dim NameReadT4() = siteK(pos3 + 7).Split(New Char() {"/"c})

                                                Dim NameReadT5() = NameReadT4(0).Split(New Char() {" "c})
                                                For bds As Integer = 0 To NameReadT5.GetUpperBound(0)
                                                    If NameReadT5(bds) <> "" Then
                                                        dg2 += 1
                                                        If dg2 = 2 Then rsitek1 = NameReadT5(bds)
                                                        If dg2 = 3 Then rsitek1 = rsitek1 + Chr(32) + NameReadT5(bds)
                                                        If dg2 = 4 Then rsitek1 = rsitek1 + Chr(32) + NameReadT5(bds)
                                                        If dg2 > 5 Then rsitek1 = rsitek1 + Chr(32) + NameReadT5(bds)
                                                    End If
                                                Next
                                                For bd As Integer = 0 To NameReadT4.GetUpperBound(0)
                                                    If NameReadT4(bd) <> "" Then ' And dg > 1 Then
                                                        dg += 1


                                                        If dg > 1 Then
                                                            dg3 += 1
                                                            If dg3 = 1 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT4(bd)
                                                            If dg3 = 2 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT4(bd)
                                                            If dg3 > 3 Then rsitek1 = rsitek1 + Chr(32) + NameReadT4(bd)
                                                        End If
                                                    End If

                                                Next
                                                ' MsgBox(rsitek1)
                                                xlWorkSheet.Cells(11 + b, 6) = rsitek1
                                                'xlWorkSheet.Cells(11 + b, 6) = siteK(posP3 + 1)
                                            Else
                                                If siteK(pos3 + 11).Contains("-") Then
                                                    xlWorkSheet.Cells(11 + b, 6) = siteK(pos3 + 8) + siteK(pos3 + 9) + Chr(32) + siteK(pos3 + 10)
                                                Else
                                                    xlWorkSheet.Cells(11 + b, 6) = siteK(pos3 + 8) + siteK(pos3 + 9)
                                                End If
                                                If siteK(pos3 + 7).Contains("-") Then
                                                Else
                                                    xlWorkSheet.Cells(11 + b, 6) = " "
                                                End If
                                            End If

                                        End If
                                    End If

                                    Dim ReadTableqW() As String = siteK(pos2 - 5).Split(" ")
                                    If siteK(pos2 - 5).Contains("Kg") Then

                                        If ReadTableqW.GetUpperBound(0) > 0 Then
                                            ReadTableqW = siteK(pos2 - 5).Split(" ")
                                        Else
                                            ReadTableqW = siteK(pos2 - 6).Split(" ")

                                        End If
                                    Else
                                        ReadTableqW = siteK(pos2 - 6).Split(" ")
                                    End If

                                    Dim rwTab() As String = Nothing
                                    Dim fReadWs As Integer
                                    Dim sreads As Integer = 0
                                    For fReadWs = ReadTableqW.GetLowerBound(0) To ReadTableqW.GetUpperBound(0)
                                        If ReadTableqW(fReadWs) <> "Kg" Then
                                            ReDim Preserve rwTab(sreads)
                                            rwTab(sreads) = ReadTableqW(fReadWs)
                                            sreads += 1
                                        End If


                                    Next
                                    Dim memorW As String = rwTab(0)
                                    If memorW.Contains(".") Then
                                        memorW = Microsoft.VisualBasic.Replace(memorW, ".", ",")
                                    End If
                                    xlWorkSheet.Cells(11 + b, 7) = memorW 'masa części - weight part - odczyt z pliku i tablicy readTAB
                                    xlWorkSheet.Cells(11 + b, 8) = "-----------"  'brak folderu

                                    xlWorkSheet.Cells(11 + b, 9) = rsit '  numer rysunku  pdf do celów inforam
                                    xlWorkSheet.Cells(11 + b, 10) = 1 ' liczba wystąpień
                                    b += 1
                                    h += 1
                                End If
                            End If
                        Next

                    End If

                    ' For zlicz1 As Integer = 0 To zlicz
                    If Microsoft.VisualBasic.Right(folder, 3) = "dwg" Then
                        Dim StabF As String() = folder.Split("\")
                        Dim lStabF As Integer = StabF.GetUpperBound(0)
                        xlWorkSheet.Cells(11 + zlicz2, 2) = StabF(lStabF) ' numer rysunku dwg - number of drawing

                        '  xlWorkSheet.Cells(11 + zlicz2, 2).Hyperlinks.Add(Anchor:=xlWorkSheet.Cells(11 + zlicz2, 2), Address:=folder, SubAddress:="", TextToDisplay:=CStr(ListView1.Items(di).Tag))
                        zlicz2 += 1
                    Else
                        'xlWorkSheet.Cells(11 + zlicz2, 2) = "---------------"
                    End If

                    'Next
                    'b += 1
                    'h += 1
                    ' Next



                    ' Dodawanie danych dl excela
                    xlWorkSheet.Cells(2, 3) = TextBox7.Text ' numer  kontraktu
                    xlWorkSheet.Cells(3, 6) = TextBox1.Text ' nazwa  kontraktu
                    xlWorkSheet.Cells(5, 1) = "Klient/Kunde:" & TextBox2.Text ' nazwa klienta
                    xlWorkSheet.Cells(6, 1) = "Budowa/ Baustelle:" & TextBox3.Text ' nazwa budowy
                    xlWorkSheet.Cells(7, 1) = "Wykonał/ausgeführt von:" & TextBox4.Text ' nazwa budowy
                    xlWorkSheet.Cells(8, 1) = "Data/Daten:" & TextBox5.Text ' nazwa budowy
                    xlWorkSheet.Cells.Range("C11:H300").Font.Size = 10

                    'Next
                    '    Next
                    'Catch ex As Exception
                    '  MsgBox(ex.Message)
                    ' End Try

                    'Exit For
                    ' end if

                    'End If
                    ' h += 1
                    '   Next
                    'b += 1
                    'h += 1
                    'End If
                Next

                xlWorkBook.SaveAs(TextBox6.Text & "Zeichnungsliste.xls")

                xlWorkBook.Close()
                xlApp.Quit()
                xlWorkBook = Nothing
                xlApp = Nothing
            End If

        End If

        ' Drawing list 2
        If dr = Windows.Forms.DialogResult.OK Then
            Dim nFiles, h As Integer
            'Dim rect As util.RectangleJ
            'Dim rectA4_1, rectA4_2, rectA4_3, rectA4_4 As util.RectangleJ
            'Dim rectA3_1, rectA3_2, rectA3_3, rectA3_4, rectA3_5 As util.RectangleJ
            'Dim rectA2_1, rectA2_2, rectA2_3, rectA2_4, rectA2_5 As util.RectangleJ
            'Dim rectA1_1, rectA1_2, rectA1_3, rectA1_4, rectA1_5 As util.RectangleJ
            'Dim rectA0_1, rectA0_2, rectA0_3, rectA0_4, rectA0_5 As util.RectangleJ
            'Dim Out_Name_rys As String
            Dim myFileDialog As OpenFileDialog = New OpenFileDialog()
            Dim xlApp As Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet

            xlApp = New Excel.ApplicationClass
            xlApp.Visible = True

            If Dialog1.RadioButton2.Checked = True Then
                xlWorkBook = xlApp.Workbooks.Open(Filename:=My.Application.Info.DirectoryPath + "\" + "Zeichnungsliste2.xls")
                xlWorkSheet = xlWorkBook.Worksheets("Zeichnungsliste (2)")


                'display the cells value B2
                '=============================================
                ' w pierwszej kolejności należy poslugiwać się itextem. Jeżeli w nazwie znajdzie mi wartość pustą musi przejść innym programem
                '===============================================
                Dim folder As String = Nothing
                Dim zlicz As Integer = 0
                Dim b As Integer = 0
                Dim sTable As String() = Nothing  '"" ' item number 1 -numer rysunku
                Dim sTableq As String() = Nothing ' / ' item number 1 -numer rysunku
                Dim zlicz2 As Integer = 0
                Dim zlicz3 As Integer = 0
                For di As Integer = 0 To ListView1.Items.Count - 1
                    If ListView1.Items(di).Selected = False Then
                        ListView1.Items(di).Tag = ListView1.Items(di).Text
                        folder = ListTab(0) + "\" + CStr(ListView1.Items(di).Tag)
                        ListTab(0) = ListTab(0).ToString '+ "\" + CStr(ListView1.Items(di).Tag)

                        Dim n As Integer = 0
                        '  Try
                        If Not folder Is Nothing AndAlso IO.Directory.Exists(folder) Then

                            For Each file As String In IO.Directory.GetFiles(folder)
                                If Microsoft.VisualBasic.Right(file, 3) = "pdf" Then

                                    ' -------------------------------------------------------------------------------------
                                    ' Bibliteka ItexSharp
                                    ' Sprawdzanie tekstu w locie - dla standardowych plików pdf
                                    '---------------------------------------------------------------------------------------
                                    Dim dt As New Data.DataTable()
                                    'Using cn As New SqlClient.SqlConnection(ConfigurationManager.ConnectionStrings("ConsoleApplication3.Properties.Settings.daasConnectionString").ConnectionString)
                                    Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
                                        cn.Open()
                                        Dim SQLcommand As SQLite.SQLiteCommand
                                        SQLcommand = cn.CreateCommand
                                        SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL like '" & TextBox9.Text & "'"
                                        'SQLcommand.CommandText = "SELECT * FROM TranslateBase where PL='" + defDescription(0).ToString + "' "
                                        'SQLcommand.CommandText = "SELECT PL,DE FROM TranslateBase"
                                        Dim lrd As IDataReader = SQLcommand.ExecuteReader()
                                        ' Dim SQLreader As System.Data.SqlClient.SqlDataReader = SQLcommand.ExecuteReader()
                                        DataGridView1.Rows.Clear()
                                        'Tracks the current record number
                                        '  Dim n As Integer = 0
                                        '  Dim i As Integer
                                        While lrd.Read()

                                            'MsgBox(lrd.Item(1).ToString)
                                            n = DataGridView1.Rows.Add()
                                            DataGridView1.Rows.Item(n).Cells(0).Value = lrd.Item(0).ToString
                                            DataGridView1.Rows.Item(n).Cells(1).Value = lrd.Item(1).ToString
                                            DataGridView1.Rows.Item(n).Cells(2).Value = lrd.Item(2).ToString
                                            DataGridView1.Rows.Item(n).Cells(3).Value = lrd.Item(3).ToString
                                            DataGridView1.Rows.Item(n).Cells(4).Value = lrd.Item(4).ToString
                                            DataGridView1.Rows.Item(n).Cells(5).Value = lrd.Item(5).ToString
                                            DataGridView1.Rows.Item(n).Cells(6).Value = lrd.Item(6).ToString
                                            DataGridView1.Rows.Item(n).Cells(7).Value = lrd.Item(7).ToString



                                        End While

                                        'SQLcommand.ExecuteNonQuery()
                                        SQLcommand.Dispose()
                                        'Next
                                        cn.Close()
                                    End Using


                                    RichTextBox1.Clear()
                                    '    Dim oReader As New iTextSharp.text.pdf.PdfReader(ListView1.Items(nFiles).ImageKey)
                                    Dim oReader As New iTextSharp.text.pdf.PdfReader(file)
                                    Dim i As Integer
                                    Dim sOut = ""
                                    Dim sOut1 = ""
                                    Dim ss = ""
                                    Dim sf As String
                                    Dim strText As String = ""
                                    Try
                                        For i = 1 To oReader.NumberOfPages
                                            Dim its As New iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
                                            ' odczyt tekstu w pliku
                                            sOut1 &= iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(oReader, i, its)

                                        Next
                                    Catch es As Exception
                                        ' Let the user know what went wrong.
                                        Console.WriteLine("The file could not be read:")
                                        Console.WriteLine(es.Message)
                                    End Try
                                    ' parseUsingPDFBox(folder)

                                    ' Display the text
                                    ' Kodowanie do polskich znaków
                                    Dim helvetica As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1250, BaseFont.EMBEDDED)
                                    Dim plLarge As New iTextSharp.text.Font(helvetica, 16)
                                    ' Dekodowanie pliku do UTF8 - brak wszystkich polskich znaków
                                    sOut = (Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(sOut1))))
                                    'RichTextBox1.AppendText((Environment.NewLine + sOut.ToString()))
                                    'sf = DirectCast(Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.[Default], Encoding.UTF8, Encoding.[Default].GetBytes(sOut1))), String)
                                    ' Zapis do Stringbuildera z uwzględnieniem polskich znaków
                                    Dim sb As New StringBuilder()
                                    sb.AppendFormat(sOut, plLarge) ' do przetestowania
                                    'Dim resultsb As String = Encoding.GetEncoding("ISO-8859-2").GetString(System.Text.Encoding.GetEncoding("ISO-8859-2").GetBytes(sOut)) ' polskie znaki
                                    'MsgBox(resultsb)
                                    Dim line As String()
                                    Dim args As String
                                    line = (sOut.Split(ControlChars.Lf))

                                    'Dodanie nowych lini odczytanych po zakonczniu tekstu
                                    For Each item As String In line
                                        RichTextBox1.AppendText((item & ControlChars.Lf))
                                        '   TextBox7.Text += item & ControlChars.Lf
                                    Next
                                    If sOut = Nothing Then
                                        RichTextBox2.AppendText(("Plik jest obrazem" & "-" & folder & ControlChars.Lf))
                                        xlWorkSheet.Cells(11 + b, 6) = "Plik jest obrazem" & "-" & folder
                                    Else

                                        ''  RichTextBox1.AppendText((sOut.ToString())) - bez polskich znaków
                                        ''RichTextBox1.AppendText((Environment.NewLine + sb.ToString())) '- z polskimi znakami
                                        'RichTextBox1.AppendText((resultsb.ToString())) '- bez polskich znaków

                                        Dim xf As Integer = 0
                                        Dim x As Integer
                                        Dim siteK() As String
                                        Dim sit() As String
                                        ' Odczyt z richboxa i zapis do tablicy  której wysukuje się dane.
                                        For x = RichTextBox1.Lines.GetLowerBound(0) To RichTextBox1.Lines.GetUpperBound(0)
                                            ' MessageBox.Show(RichTextBox1.Lines(x))

                                            ReDim Preserve siteK(xf)
                                            ReDim Preserve sit(xf)
                                            siteK(xf) = RichTextBox1.Lines(x)
                                            xf = xf + 1
                                        Next
                                        ' Wyszukiwanie tekstu i zapisywanie do pliku xls
                                        'Dim result1 As New List(Of Integer)
                                        Dim result1(4) As String
                                        'Dim result1 As List(Of Integer) = New List(Of Integer)
                                        For ig As Integer = 0 To siteK.Length - 1
                                            ' określa pozycje wystąpienia znaku
                                            If siteK(ig).Contains("Zmiana:") Then result1(0) = ig ' item number 1 -numer rysunku
                                            If siteK(ig).Contains("Zatwierdził ") Or siteK(ig).Contains("Zatwierdził") Then result1(1) = ig 'item number 2 - rozmiar rysunku
                                            If siteK(ig).Contains("Nr.artykułu") Then result1(2) = ig 'item number 3 - waga częsci then
                                            If siteK(ig).Contains("Ciężar") Then result1(3) = ig 'item number 4 - nazwa rysunku
                                            If siteK(ig).Contains("Format:") Then result1(4) = ig 'item number 5 - index zmian


                                            'If siteK(ig).Contains("Nr rysunku /Drawing-No. /RAJZSZÁM") Then result1.Add(ig) ' item number 1 -numer rysunku
                                            'If siteK(ig).Contains("Zatwierdził") Or siteK(ig).Contains("Zatwierdzi Bestaetigt") Or siteK(ig).Contains("ZatwierdziBestaetigt") Then result1.Add(ig) 'item number 2 - rozmiar rysunku
                                            'If siteK(ig).Contains("Approved") Then result1.Add(ig) 'item number 3 - waga częsci then
                                            'If siteK(ig).Contains("MASA/WEIGHT/HMOTNOSŤ") Then result1.Add(ig) 'item number 4 - nazwa rysunku
                                            'If siteK(ig).Contains("WERKSTOFF") Or siteK(ig).Contains("MATERIA WERKSTOFF") Then result1.Add(ig) 'item number 5 - index zmian

                                        Next
                                        Dim pos As Integer = result1(0).ToString ' item number 1 -numer rysunku
                                        Dim pos1 As Integer = result1(1).ToString 'item number 2 - rozmiar rysunku
                                        Dim pos2 As Integer = result1(2).ToString 'item number 3 - waga częsci
                                        Dim pos3 As Integer = result1(3).ToString 'item number 4 - nazwa rysunku
                                        Dim pos4 As Integer = result1(4).ToString 'item number 5 - index zmian


                                        sTable = siteK(pos + 4).Split(" ") ' item number 1 -numer rysunku
                                        sTableq = siteK(pos + 4).Split("/") ' item number 1 -numer rysunku


                                        xlWorkSheet.Cells(11 + b, 1) = b + 1 ' Numer pozycji 


                                        Dim sTables() = siteK(pos3 + 7).Split(New Char() {" "c, "/"c})
                                        Dim dggw As Integer = 0
                                        Dim dg22w As Integer = 0
                                        Dim rsit1 As String
                                        For bdg As Integer = 0 To sTables.GetUpperBound(0)
                                            If sTables(bdg) <> "" Then ' And dg > 1 Then
                                                dggw += 1
                                                If dggw > 0 Then
                                                    dg22w += 1
                                                    If dg22w = 1 Then rsit1 = sTables(bdg)
                                                End If
                                            End If
                                        Next
                                        If Microsoft.VisualBasic.Right(ListView1.Items(di).Text, 3) = "dwg" Then
                                            xlWorkSheet.Cells(11 + b, 2) = rsit1  ' numer rysunku dwg - number of drawing
                                        Else
                                            'xlWorkSheet.Cells(11 + b, 2) = "---------------"
                                        End If
                                        xlWorkSheet.Cells(11 + b, 6) = rsit1  '  numer rysunku  pdf
                                        ' dodanie hiperzłącza - hiperlinks
                                        'xlWorkSheet.Cells(11 + b, 3).Value = "..\wieszak\Nowy folder\204-0720-007-00000.pdf"
                                        'xlWorkSheet.Cells(11 + b, 3).Hyperlinks.Add(xlWorkSheet.Cells(11 + b, 3), xlWorkSheet.Cells(11 + b, 3).Value)

                                        If siteK(pos4 + 23) = "General tolerances:" Then
                                            xlWorkSheet.Cells(11 + h, 4) = ""
                                        Else
                                            If siteK(pos4 + 24) = "  " Then
                                                xlWorkSheet.Cells(11 + h, 7) = siteK(pos4 + 24)
                                            Else
                                                Dim NameReadT9 = siteK(pos3 + 7).Split("/")
                                                Dim Lsle As Integer = NameReadT9.GetUpperBound(0)
                                                If Lsle > 1 Then
                                                    xlWorkSheet.Cells(11 + h, 7) = siteK(pos4 + 22) ' index zmian
                                                Else
                                                    xlWorkSheet.Cells(11 + h, 7) = siteK(pos4 + 23)
                                                End If
                                            End If
                                        End If
                                        xlWorkSheet.Cells(11 + b, 8) = siteK(pos1 - 3) ' rozmiar rysunku- drawing size


                                        Dim NazwaReadTs, fRead As Integer
                                        Dim sread As Integer = 0
                                        Dim lNameReadTs
                                        Dim NameReadT As Integer = Len(siteK(pos3 + 8)) ' odlicza długość pierwszego wyrazu /  /
                                        Dim NameReadTs = siteK(pos3 + 8).Split("/")
                                        If NameReadTs.GetUpperBound(0) > 1 Then ' jeżeli dla pozycji p3 występuje  ciąg  ///

                                            lNameReadTs = Len(NameReadTs(2))
                                            'MsgBox(Mid(readTab(resultT.Item(0) + 1), 1, NameReadT - NazwaReadTs - 1))
                                            xlWorkSheet.Cells(11 + b, 9) = (siteK(pos3 + 8)) 'nazwa części - part name
                                        Else ' jeżeli dla pozycji p3 występuje  ciąg  /// - wtedy zmieniamy pozycję ciągu w pliku
                                            'Dim NameReadT2 As Integer = Len(siteK(posP3 + 2)) ' odlicza długość pierwszego wyrazu /  /
                                            Dim NameReadT2 = siteK(pos3 + 7).Split("/")
                                            Dim Lslesh As Integer = NameReadT2.GetUpperBound(0)
                                            If Lslesh > 0 And Lslesh <= 1 Then
                                                'Dim NameReadTs22 As Integer = Len(NameReadTs2(2))
                                                'Dim NameReadTs_ = siteK(posP3 + 2).Split(" ")
                                                'Dim NameReadTs_2 As Integer = Len(NameReadTs_(0))
                                                xlWorkSheet.Cells(11 + b, 9) = siteK(pos3 + 2) & siteK(pos3 + 3) ', 1 + NameReadTs_2, NameReadT2 - NameReadTs22 - NameReadTs_2) 'nazwa części - part name
                                                If Lslesh > 0 And Lslesh <= 2 Then
                                                    xlWorkSheet.Cells(11 + b, 9) = siteK(pos3 + 2) & siteK(pos3 + 3) & siteK(pos3 + 4)
                                                End If
                                            Else
                                                Dim NameReadT3() = siteK(pos3 + 7).Split(New Char() {"/"c})
                                                Dim Lslesh1 As Integer = NameReadT3.GetUpperBound(0)
                                                Dim rsitek1 As String = Nothing
                                                Dim rsitek2 As String = Nothing
                                                Dim dg As Integer = 0
                                                Dim dh As Integer = 0
                                                Dim bd As Integer = 1
                                                Dim df As Integer = 0
                                                If Lslesh1 > 0 And Lslesh1 <= 2 Then
                                                    Dim NameReadT4() = siteK(pos3 + 7).Split(New Char() {" "c, "/"c})
                                                    If NameReadT4(0) = "" Then
                                                        df = 1
                                                        'GoTo loopD
                                                    End If
                                                    'LoopD:
                                                    For bd = 1 + df To NameReadT4.GetUpperBound(0)

                                                        If NameReadT4(bd) <> "" Then
                                                            dg += 1
                                                            If dg = 1 Then rsitek1 = NameReadT4(bd)
                                                            If dg = 2 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT4(bd)
                                                            If dg = 3 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT4(bd)
                                                            If dg > 3 Then rsitek1 = rsitek1 + Chr(32) + NameReadT4(bd)
                                                        End If
                                                        'If NameReadT3(2) = "" Then
                                                        '    dh += 1
                                                        '    Dim NameReadT5() = NameReadT3(0).Split(New Char() {" "c, "/"c})
                                                        '    For bh As Integer = 1 To NameReadT5.GetUpperBound(0)
                                                        '        If dh = 1 Then rsitek1 = NameReadT5(bd)
                                                        '        If dh = 2 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT5(bd)
                                                        '        If dh = 3 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT5(bd)
                                                        '    Next
                                                        'End If
                                                    Next
                                                    ' MsgBox(rsitek1)
                                                    xlWorkSheet.Cells(11 + b, 9) = rsitek1
                                                    'xlWorkSheet.Cells(11 + b, 6) = siteK(posP3 + 1)
                                                Else
                                                    If siteK(pos3 + 11).Contains("-") Then
                                                        xlWorkSheet.Cells(11 + b, 9) = siteK(pos3 + 8) + siteK(pos3 + 9) + Chr(32) + siteK(pos3 + 10)
                                                    Else
                                                        xlWorkSheet.Cells(11 + b, 9) = siteK(pos3 + 8) + siteK(pos3 + 9)
                                                    End If

                                                End If

                                            End If
                                        End If

                                        Dim ReadTableqW() As String = siteK(pos2 - 5).Split(" ")
                                        If siteK(pos2 - 5).Contains("Kg") Then

                                            If ReadTableqW.GetUpperBound(0) > 0 Then
                                                ReadTableqW = siteK(pos2 - 5).Split(" ")
                                            Else
                                                ReadTableqW = siteK(pos2 - 6).Split(" ")

                                            End If
                                        Else
                                            ReadTableqW = siteK(pos2 - 6).Split(" ")
                                        End If

                                        Dim rwTab() As String = Nothing
                                        Dim fReadWs As Integer
                                        Dim sreads As Integer = 0
                                        For fReadWs = ReadTableqW.GetLowerBound(0) To ReadTableqW.GetUpperBound(0)
                                            If ReadTableqW(fReadWs) <> "Kg" Then
                                                ReDim Preserve rwTab(sreads)
                                                rwTab(sreads) = ReadTableqW(fReadWs)
                                                sreads += 1
                                            End If


                                        Next
                                        Dim memorW As String = rwTab(0)

                                        xlWorkSheet.Cells(11 + b, 10) = memorW 'masa części - weight part - odczyt z pliku i tablicy readTAB

                                        Dim LSplit As String() = folder.Split("\")
                                        Dim vT As Integer = LSplit.GetUpperBound(0)
                                        Dim lPath = Len(LSplit(vT))
                                        xlWorkSheet.Cells(11 + b, 11) = Microsoft.VisualBasic.Right(folder, lPath) 'brak folderu

                                        xlWorkSheet.Cells(11 + b, 12) = sTable(0) '  numer rysunku  pdf do celów inforam
                                        'xlWorkSheet.Cells(11 + b, 10) = 1 ' liczba wystąpień
                                        b += 1
                                        h += 1
                                    End If
                                End If
                                If Microsoft.VisualBasic.Right(file, 3) = "dwg" Then
                                    Dim StabF2 As String() = file.Split("\")
                                    Dim lStabF2 As Integer = StabF2.GetUpperBound(0)
                                    xlWorkSheet.Cells(11 + zlicz2, 5) = StabF2(lStabF2) ' numer rysunku dwg - number of drawing
                                    'xlWorkSheet.Cells(11 + zlicz2, 5).Hyperlinks.Add(Anchor:=xlWorkSheet.Cells(11 + zlicz2, 2), Address:=folder, SubAddress:="", TextToDisplay:=CStr(ListView1.Items(di).Tag))
                                  
                                    xlWorkSheet.Cells(11 + zlicz3 + zlicz2, 2).Interior.ColorIndex = 46
                                    xlWorkSheet.Cells(11 + zlicz3 + zlicz2, 3).Interior.ColorIndex = 46
                                    xlWorkSheet.Cells(11 + zlicz3 + zlicz2, 4).Interior.ColorIndex = 46
                                    zlicz2 += 1
                                Else
                                    'xlWorkSheet.Cells(11 + zlicz2, 2) = "---------------"
                                End If
                            Next
                            ' For zlicz1 As Integer = 0 To zlicz
                       
                        End If
                        '  End If




                        If Microsoft.VisualBasic.Right(folder, 3) = "pdf" Then

                            ' -------------------------------------------------------------------------------------
                            ' Bibliteka ItexSharp
                            ' Sprawdzanie tekstu w locie - dla standardowych plików pdf
                            '---------------------------------------------------------------------------------------
                            RichTextBox1.Clear()
                            '    Dim oReader As New iTextSharp.text.pdf.PdfReader(ListView1.Items(nFiles).ImageKey)
                            Dim oReader As New iTextSharp.text.pdf.PdfReader(folder)
                            Dim i As Integer
                            Dim sOut = ""
                            Dim sOut1 = ""
                            Dim ss = ""
                            Dim sf As String
                            Dim strText As String = ""
                            Try
                                For i = 1 To oReader.NumberOfPages
                                    Dim its As New iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
                                    ' odczyt tekstu w pliku
                                    sOut1 &= iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(oReader, i, its)

                                Next
                            Catch es As Exception
                                ' Let the user know what went wrong.
                                Console.WriteLine("The file could not be read:")
                                Console.WriteLine(es.Message)
                            End Try
                            ' parseUsingPDFBox(folder)

                            ' Display the text
                            ' Kodowanie do polskich znaków
                            Dim helvetica As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1250, BaseFont.EMBEDDED)
                            Dim plLarge As New iTextSharp.text.Font(helvetica, 16)
                            ' Dekodowanie pliku do UTF8 - brak wszystkich polskich znaków
                            sOut = (Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(sOut1))))
                            'RichTextBox1.AppendText((Environment.NewLine + sOut.ToString()))
                            'sf = DirectCast(Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.[Default], Encoding.UTF8, Encoding.[Default].GetBytes(sOut1))), String)
                            ' Zapis do Stringbuildera z uwzględnieniem polskich znaków
                            Dim sb As New StringBuilder()
                            sb.AppendFormat(sOut, plLarge) ' do przetestowania
                            'Dim resultsb As String = Encoding.GetEncoding("ISO-8859-2").GetString(System.Text.Encoding.GetEncoding("ISO-8859-2").GetBytes(sOut)) ' polskie znaki
                            'MsgBox(resultsb)
                            Dim line As String()
                            Dim args As String
                            line = (sOut.Split(ControlChars.Lf))

                            'Dodanie nowych lini odczytanych po zakonczniu tekstu
                            For Each item As String In line
                                RichTextBox1.AppendText((item & ControlChars.Lf))
                                '   TextBox7.Text += item & ControlChars.Lf
                            Next
                            If sOut = Nothing Then
                                RichTextBox2.AppendText(("Plik jest obrazem" & "-" & folder & ControlChars.Lf))
                                xlWorkSheet.Cells(11 + b, 6) = "Plik jest obrazem" & "-" & folder
                            Else

                                ''  RichTextBox1.AppendText((sOut.ToString())) - bez polskich znaków
                                ''RichTextBox1.AppendText((Environment.NewLine + sb.ToString())) '- z polskimi znakami
                                'RichTextBox1.AppendText((resultsb.ToString())) '- bez polskich znaków

                                Dim xf As Integer = 0
                                Dim x As Integer
                                Dim siteK() As String
                                Dim sit() As String
                                ' Odczyt z richboxa i zapis do tablicy  której wysukuje się dane.
                                For x = RichTextBox1.Lines.GetLowerBound(0) To RichTextBox1.Lines.GetUpperBound(0)
                                    ' MessageBox.Show(RichTextBox1.Lines(x))

                                    ReDim Preserve siteK(xf)
                                    ReDim Preserve sit(xf)
                                    siteK(xf) = RichTextBox1.Lines(x)
                                    xf = xf + 1
                                Next
                                ' Wyszukiwanie tekstu i zapisywanie do pliku xls
                                'Dim result1 As New List(Of Integer)
                                Dim result1(4) As String
                                'Dim result1 As List(Of Integer) = New List(Of Integer)
                                For ig As Integer = 0 To siteK.Length - 1
                                    ' określa pozycje wystąpienia znaku
                                    If siteK(ig).Contains("Zmiana:") Then result1(0) = ig ' item number 1 -numer rysunku
                                    If siteK(ig).Contains("Zatwierdził ") Or siteK(ig).Contains("Zatwierdził") Then result1(1) = ig 'item number 2 - rozmiar rysunku
                                    If siteK(ig).Contains("Nr.artykułu") Then result1(2) = ig 'item number 3 - waga częsci then
                                    If siteK(ig).Contains("Ciężar") Then result1(3) = ig 'item number 4 - nazwa rysunku
                                    If siteK(ig).Contains("Format:") Then result1(4) = ig 'item number 5 - index zmian


                                    'If siteK(ig).Contains("Nr rysunku /Drawing-No. /RAJZSZÁM") Then result1.Add(ig) ' item number 1 -numer rysunku
                                    'If siteK(ig).Contains("Zatwierdził") Or siteK(ig).Contains("Zatwierdzi Bestaetigt") Or siteK(ig).Contains("ZatwierdziBestaetigt") Then result1.Add(ig) 'item number 2 - rozmiar rysunku
                                    'If siteK(ig).Contains("Approved") Then result1.Add(ig) 'item number 3 - waga częsci then
                                    'If siteK(ig).Contains("MASA/WEIGHT/HMOTNOSŤ") Then result1.Add(ig) 'item number 4 - nazwa rysunku
                                    'If siteK(ig).Contains("WERKSTOFF") Or siteK(ig).Contains("MATERIA WERKSTOFF") Then result1.Add(ig) 'item number 5 - index zmian

                                Next
                                Dim pos As Integer = result1(0).ToString ' item number 1 -numer rysunku
                                Dim pos1 As Integer = result1(1).ToString 'item number 2 - rozmiar rysunku
                                Dim pos2 As Integer = result1(2).ToString 'item number 3 - waga częsci
                                Dim pos3 As Integer = result1(3).ToString 'item number 4 - nazwa rysunku
                                Dim pos4 As Integer = result1(4).ToString 'item number 5 - index zmian


                                sTable = siteK(pos + 4).Split(" ") ' item number 1 -numer rysunku
                                sTableq = siteK(pos + 4).Split("/") ' item number 1 -numer rysunku


                                xlWorkSheet.Cells(11 + b, 1) = b + 1 ' Numer pozycji 

                                If Microsoft.VisualBasic.Right(ListView1.Items(di).Text, 3) = "dwg" Then
                                    xlWorkSheet.Cells(11 + b, 5) = sTable(0) ' numer rysunku dwg - number of drawing
                                Else
                                    'xlWorkSheet.Cells(11 + b, 2) = "---------------"
                                End If



                                xlWorkSheet.Cells(11 + b, 6) = sTable(0) '  numer rysunku  pdf
                                ' dodanie hiperzłącza - hiperlinks
                                'MsgBox(".." & Microsoft.VisualBasic.Right(ListTab(0), Len(ListTab(0)) - 2) & "\" & sTable(0) & ".pdf")
                                '=HIPERŁĄCZE(ZŁĄCZ.TEKSTY("..\3. Zeichnungen\";G68;"\";H68;".pdf");H68)

                                'xlWorkSheet.Cells(11 + b, 3) = "=HIPERŁĄCZE(ZŁĄCZ.TEKSTY(""..\3. Zeichnungen\"";G68;""\"";H68;"".pdf"");H68)"
                                'xlWorkSheet.Cells(11 + b, 3).Hyperlinks.Add(Anchor:=xlWorkSheet.Cells(11 + b, 3), Address:=info, SubAddress:="", TextToDisplay:=sTable(0))
                                'xlWorkSheet.Cells(11 + b, 3).Hyperlinks.Add(Anchor:=xlWorkSheet.Cells(11 + b, 3), Address:=folder, SubAddress:="", TextToDisplay:=sTable(0))
                                '   xlWorkSheet.Cells(11 + b, 6).Value = "..\wieszak\Nowy folder\204-0720-007-00000.pdf"
                                '   xlWorkSheet.Cells(11 + b, 6).Hyperlinks.Add(xlWorkSheet.Cells(11 + b, 3), xlWorkSheet.Cells(11 + b, 3).Value)

                                If siteK(pos4 + 23) = "General tolerances:" Then
                                    xlWorkSheet.Cells(11 + h, 7) = ""
                                Else
                                    If siteK(pos4 + 24) = "  " Then
                                        xlWorkSheet.Cells(11 + h, 7) = siteK(pos4 + 24)
                                    Else
                                        Dim NameReadT9 = siteK(pos3 + 7).Split("/")
                                        Dim Lsle As Integer = NameReadT9.GetUpperBound(0)
                                        If Lsle > 1 Then
                                            xlWorkSheet.Cells(11 + h, 7) = siteK(pos4 + 22) ' index zmian
                                        Else
                                            xlWorkSheet.Cells(11 + h, 7) = siteK(pos4 + 23)
                                        End If
                                    End If
                                End If
                                xlWorkSheet.Cells(11 + b, 8) = siteK(pos1 - 3) ' rozmiar rysunku- drawing size


                                Dim NazwaReadTs, fRead As Integer
                                Dim sread As Integer = 0
                                Dim lNameReadTs
                                Dim NameReadT As Integer = Len(siteK(pos3 + 8)) ' odlicza długość pierwszego wyrazu /  /
                                Dim NameReadTs = siteK(pos3 + 8).Split("/")
                                If NameReadTs.GetUpperBound(0) > 1 Then ' jeżeli dla pozycji p3 występuje  ciąg  ///

                                    lNameReadTs = Len(NameReadTs(2))
                                    'MsgBox(Mid(readTab(resultT.Item(0) + 1), 1, NameReadT - NazwaReadTs - 1))
                                    xlWorkSheet.Cells(11 + b, 9) = (siteK(pos3 + 8)) 'nazwa części - part name
                                Else ' jeżeli dla pozycji p3 występuje  ciąg  /// - wtedy zmieniamy pozycję ciągu w pliku
                                    'Dim NameReadT2 As Integer = Len(siteK(posP3 + 2)) ' odlicza długość pierwszego wyrazu /  /
                                    Dim NameReadT2 = siteK(pos3 + 7).Split("/")
                                    Dim Lslesh As Integer = NameReadT2.GetUpperBound(0)
                                    If Lslesh > 0 And Lslesh <= 1 Then
                                        'Dim NameReadTs22 As Integer = Len(NameReadTs2(2))
                                        'Dim NameReadTs_ = siteK(posP3 + 2).Split(" ")
                                        'Dim NameReadTs_2 As Integer = Len(NameReadTs_(0))
                                        xlWorkSheet.Cells(11 + b, 9) = siteK(pos3 + 2) & siteK(pos3 + 3) ', 1 + NameReadTs_2, NameReadT2 - NameReadTs22 - NameReadTs_2) 'nazwa części - part name
                                        If Lslesh > 0 And Lslesh <= 2 Then
                                            xlWorkSheet.Cells(11 + b, 9) = siteK(pos3 + 2) & siteK(pos3 + 3) & siteK(pos3 + 4)
                                        End If
                                    Else
                                        Dim NameReadT3() = siteK(pos3 + 7).Split(New Char() {"/"c})
                                        Dim Lslesh1 As Integer = NameReadT3.GetUpperBound(0)
                                        Dim rsitek1 As String = Nothing
                                        Dim rsitek2 As String = Nothing
                                        Dim dg As Integer = 0
                                        Dim dh As Integer = 0
                                        If Lslesh1 > 0 And Lslesh1 <= 2 Then
                                            Dim NameReadT4() = siteK(pos3 + 7).Split(New Char() {" "c, "/"c})

                                            For bd As Integer = 1 To NameReadT4.GetUpperBound(0)
                                                If NameReadT4(bd) <> "" Then
                                                    dg += 1
                                                    If dg = 1 Then rsitek1 = NameReadT4(bd)
                                                    If dg = 2 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT4(bd)
                                                    If dg = 3 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT4(bd)
                                                    If dg > 3 Then rsitek1 = rsitek1 + Chr(32) + NameReadT4(bd)
                                                End If
                                                'If NameReadT3(2) = "" Then
                                                '    dh += 1
                                                '    Dim NameReadT5() = NameReadT3(0).Split(New Char() {" "c, "/"c})
                                                '    For bh As Integer = 1 To NameReadT5.GetUpperBound(0)
                                                '        If dh = 1 Then rsitek1 = NameReadT5(bd)
                                                '        If dh = 2 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT5(bd)
                                                '        If dh = 3 Then rsitek1 = rsitek1 + Chr(32) & "/" & Chr(32) + NameReadT5(bd)
                                                '    Next
                                                'End If
                                            Next
                                            ' MsgBox(rsitek1)
                                            xlWorkSheet.Cells(11 + b, 9) = rsitek1
                                            'xlWorkSheet.Cells(11 + b, 6) = siteK(posP3 + 1)
                                        Else
                                            If siteK(pos3 + 11).Contains("-") Then
                                                xlWorkSheet.Cells(11 + b, 9) = siteK(pos3 + 8) + siteK(pos3 + 9) + Chr(32) + siteK(pos3 + 10)
                                            Else
                                                xlWorkSheet.Cells(11 + b, 9) = siteK(pos3 + 8) + siteK(pos3 + 9)
                                            End If

                                        End If

                                    End If
                                End If

                                Dim ReadTableqW() As String = siteK(pos2 - 5).Split(" ")
                                If siteK(pos2 - 5).Contains("Kg") Then

                                    If ReadTableqW.GetUpperBound(0) > 0 Then
                                        ReadTableqW = siteK(pos2 - 5).Split(" ")
                                    Else
                                        ReadTableqW = siteK(pos2 - 6).Split(" ")

                                    End If
                                Else
                                    ReadTableqW = siteK(pos2 - 6).Split(" ")
                                End If

                                Dim rwTab() As String = Nothing
                                Dim fReadWs As Integer
                                Dim sreads As Integer = 0
                                For fReadWs = ReadTableqW.GetLowerBound(0) To ReadTableqW.GetUpperBound(0)
                                    If ReadTableqW(fReadWs) <> "Kg" Then
                                        ReDim Preserve rwTab(sreads)
                                        rwTab(sreads) = ReadTableqW(fReadWs)
                                        sreads += 1
                                    End If


                                Next
                                Dim memorW As String = rwTab(0)

                                xlWorkSheet.Cells(11 + b, 10) = memorW 'masa części - weight part - odczyt z pliku i tablicy readTAB
                                xlWorkSheet.Cells(11 + b, 11) = "-----------"  'brak folderu

                                xlWorkSheet.Cells(11 + b, 12) = sTable(0) '  numer rysunku  pdf do celów inforam
                                'xlWorkSheet.Cells(11 + b, 10) = 1 ' liczba wystąpień
                                b += 1
                                h += 1
                            End If
                        End If

                    End If

                    Dim zlicz4 As Integer
                    If Microsoft.VisualBasic.Right(folder, 3) = "dwg" Then

                        If zlicz2 > 1 Then
                            zlicz4 = 1
                        End If
                        Dim StabFF As String() = folder.Split("\")
                        Dim lStabFF As Integer = StabFF.GetUpperBound(0)
                        xlWorkSheet.Cells(11 + zlicz3, 5) = StabFF(lStabFF) ' numer rysunku dwg - number of drawing

                        'xlWorkSheet.Cells(11 + zlicz2, 5).Hyperlinks.Add(Anchor:=xlWorkSheet.Cells(11 + zlicz2, 2), Address:=folder, SubAddress:="", TextToDisplay:=CStr(ListView1.Items(di).Tag))
                        xlWorkSheet.Cells(11 + zlicz3 + zlicz2 + zlicz4, 2).Interior.ColorIndex = 46
                        xlWorkSheet.Cells(11 + zlicz3 + zlicz2 + zlicz4, 3).Interior.ColorIndex = 46
                        xlWorkSheet.Cells(11 + zlicz3 + zlicz2 + zlicz4, 4).Interior.ColorIndex = 46
                        zlicz3 += 1
                    Else

                        'xlWorkSheet.Cells(11 + zlicz2, 2) = "---------------"
                    End If
                    'Next
                    'b += 1
                    'h += 1
                    ' Next



                    ' Dodawanie danych dl excela
                    xlWorkSheet.Cells(2, 4) = TextBox7.Text ' numer  kontraktu
                    xlWorkSheet.Cells(3, 9) = TextBox1.Text ' nazwa  kontraktu
                    xlWorkSheet.Cells(5, 1) = "Klient/Kunde:" & TextBox2.Text ' nazwa klienta
                    xlWorkSheet.Cells(6, 1) = "Budowa/ Baustelle:" & TextBox3.Text ' nazwa budowy
                    xlWorkSheet.Cells(7, 1) = "Wykonał/ausgeführt von:" & TextBox4.Text ' nazwa budowy
                    xlWorkSheet.Cells(8, 1) = "Data/Daten:" & TextBox5.Text ' nazwa budowy
                    xlWorkSheet.Cells.Range("C11:H300").Font.Size = 10


                    'Next
                    '    Next
                    'Catch ex As Exception
                    '  MsgBox(ex.Message)
                    ' End Try

                    'Exit For
                    ' end if

                    'End If
                    ' h += 1
                    '   Next
                    'b += 1
                    'h += 1

                    'End If
                Next
                xlWorkBook.SaveAs(TextBox6.Text & "Zeichnungsliste2.xls")

                xlWorkBook.Close()
                xlApp.Quit()
                xlWorkBook = Nothing
                xlApp = Nothing
            End If

        End If
        'End If
LoopS:
        Exit Sub
LoopS1:
        Exit Sub
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub



    Private Sub ListView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseDoubleClick
        ' only proceed if the node represents a file
        Dim filePath As String = Nothing
        'ListTab(0) = Nothing
        For Each file As ListViewItem In ListView1.Items
            Try

                filePath = ListTab(0).ToString '& file.Text
                If file.Selected = True Then
                    Select Case MessageBox.Show("You are about to open " & filePath & ".  Are you sure?", "Open File", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                        Case DialogResult.Yes
                            Process.Start(filePath)
                            Exit Select

                        Case DialogResult.No
                            MsgBox("you decided not to open")
                            Exit Select
                    End Select
                End If
            Catch ex As Exception
                MessageBox.Show("Error opening file: " & ex.Message)
            End Try
        Next
    End Sub


    Private Sub ListView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
      


        Dim FileExtension As String
        Dim FolderExtension As String
        Dim SubItemIndex As Integer
        Dim SubItemIndexs As Integer
        Dim DateMod As String
        Dim DateMods As String
        Dim folder As String


        Dim resultv As New List(Of Integer)
        Dim folder_por As String = Nothing
        Dim folder_por2 As String = Nothing
        Dim folder_d As String = Nothing
        For i As Integer = 0 To ListView1.Items.Count - 1
            If ListView1.Items(i).Selected = True Then
                ListView1.Items(i).Tag = ListView1.Items(i).Text

                Dim indexT As String() = ListTab(0).Split("\")
                For int_v As Integer = 0 To indexT.GetUpperBound(0)
                    ' sprawdzanie czy wystepuje "."
                    If indexT(int_v).Contains(".") Then
                        ' poszukiwanie pierwszego członu z "."
                        resultv.Add(int_v)
                        Exit For
                    End If
                Next
                Dim posd As Integer
                ' Zliczanie pozycji pliku w ścieżce
                If resultv.Count > 0 Then
                    posd = resultv(0).ToString
                End If
                If posd > 0 Then
                    For pos_p As Integer = 0 To posd - 1
                        If pos_p = 0 Then
                            folder_por = indexT(pos_p)
                            folder_por2 = indexT(pos_p)
                        Else
                            folder_por = folder_por + "\" + indexT(pos_p)
                            folder_por2 = folder_por2 + "\" + indexT(pos_p)
                        End If
                    Next
                Else
                    For pod_1 As Integer = 0 To ListTab.GetUpperBound(0)
                        If pod_1 = 0 Then
                            folder_d = ListTab(pod_1)

                        Else
                            folder_d = folder_d '+ "\" + ListTab(pod_1)

                        End If

                    Next
                End If
                ' Analiza folderu i zapamiętywanie do niego scieżki 
                ' jeżeli otwierane pliki będą w tym samym folderze
                ' jeżeli nie będą pobierana jest ścieżka z listTab.
                If posd = 0 Then
                    ListTab(0) = folder_d
                Else
                    ListTab(0) = folder_por
                End If
                folder = ListTab(0) + "\" + CStr(ListView1.Items(i).Tag)
                ' zapis ścieżki do listTab
                ListTab(0) = ListTab(0).ToString + "\" + CStr(ListView1.Items(i).Tag)

                Exit For

            End If
        Next



        Dim n As Integer = 0

        If Not folder Is Nothing AndAlso IO.Directory.Exists(folder) Then
            ' MsgBox("orety")
            ListView1.Items.Clear()
            Dim subItems() As ListViewItem.ListViewSubItem
            Dim subItemsD() As ListViewItem.ListViewSubItem
            Dim item As ListViewItem = Nothing
            Dim nodeDirInfo As DirectoryInfo = New DirectoryInfo(folder)

            Try
                For Each nodeDirInfo In nodeDirInfo.GetDirectories()

                    FolderExtension = IO.Path.GetExtension(nodeDirInfo.Name)
                    DateMods = IO.Directory.GetLastWriteTime(nodeDirInfo.Name)

                    item = New ListViewItem(nodeDirInfo.Name, CacheShellIcon(nodeDirInfo.FullName))
                    'ListView1.Items.Add(nodeDirInfo.Name.Substring(nodeDirInfo.Name.LastIndexOf("\"c) + 1), mkey)
                    subItems = New ListViewItem.ListViewSubItem() _
                        {New ListViewItem.ListViewSubItem(item, ""), _
                        New ListViewItem.ListViewSubItem(item, "Folder File"), _
                        New ListViewItem.ListViewSubItem(item, _
                         IO.File.GetLastWriteTime(nodeDirInfo.FullName).ToString())}

                    item.SubItems.AddRange(subItems)
                    ListView1.Items.Add(item)
                Next


                For Each filed As String In IO.Directory.GetFiles(folder)
                    'If Not folder Is Nothing = True Then
                    ' For Each fileb In nodeDirInfo.GetFiles()
                    FileExtension = IO.Path.GetExtension(filed)
                    Dim FileSize As Double
                    FileSize = Math.Round(Module1.GetSizeKB(filed.ToString), 0)
                    DateMod = IO.File.GetLastWriteTime(filed).ToString()
                    item = New ListViewItem(filed.Substring(filed.LastIndexOf("\"c) + 1), CacheShellIcon(filed))
                    subItems = New ListViewItem.ListViewSubItem() _
                    {New ListViewItem.ListViewSubItem(item, FileSize.ToString & Chr(32) & "KB"), _
                        New ListViewItem.ListViewSubItem(item, FileExtension.ToString() & Chr(32) & "File"), _
                        New ListViewItem.ListViewSubItem(item, _
                        DateMod.ToString)}

                    'AddImages(fileb.FullName)

                    If FileExtension.ToString() = "" Then
                        item = New ListViewItem(filed.Substring(filed.LastIndexOf("\"c) + 1), 5)
                        subItemsD = New ListViewItem.ListViewSubItem() _
                        {New ListViewItem.ListViewSubItem(item, FileSize.ToString & Chr(32) & "KB"), _
                        New ListViewItem.ListViewSubItem(item, "SYS. File"), _
                        New ListViewItem.ListViewSubItem(item, _
                        DateMod.ToString)}
                        item.SubItems.AddRange(subItemsD)
                        ListView1.Items.Add(item)
                    Else
                        item.SubItems.AddRange(subItems)
                        ListView1.Items.Add(item)
                    End If
                    SubItemIndex = SubItemIndex + 1
                    n += 1
                    ' End If
                Next filed



            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        End If
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        'Counter for our Physical Drives
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim z As Integer = 0

        'Start looping through the Drives
        If My.Computer.FileSystem.Drives.Count > TreeView1.Nodes.Count Then
EndOfLoop1:
            For z = 0 To TreeView1.Nodes.Count
                Dim zd As Integer = TreeView1.Nodes.Count
                If zd = 0 Then
                    Exit For
                Else
                    TreeView1.Nodes(z).Remove()
                End If

                GoTo EndOfLoop1
            Next
            For x = 0 To My.Computer.FileSystem.Drives.Count - 1
                If My.Computer.FileSystem.Drives(x).IsReady = True Then
                    'make sure the drive is ready
                    'And My.Computer.FileSystem.Drives(x).Name <> TreeView1.Nodes.Item(x).Name Then
                    'add the Drive to the Tree Node use the Drive Name as the "Key" and "Text


                    TreeView1.Nodes.Add(My.Computer.FileSystem.Drives(x).Name, My.Computer.FileSystem.Drives(x).Name)
                    'set the Tag Property to the Drive Name for Identification Later On
                    TreeView1.Nodes(My.Computer.FileSystem.Drives(x).Name).Tag = My.Computer.FileSystem.Drives(x).Name
                    'add the first level of sub directories to the TreeView
                    '=x=======
                    ''AddImages(My.Computer.FileSystem.Drives(x).Name)- lub przez podanie z funkcji
                    TreeView1.Nodes(My.Computer.FileSystem.Drives(x).Name).ImageKey = CacheShellIcon(My.Computer.FileSystem.Drives(x).Name)
                    For Each SubDirectoryW As String In My.Computer.FileSystem.GetDirectories(My.Computer.FileSystem.Drives(x).Name)
                        Try
                            TreeView1.Nodes(y).Nodes.Add(SubDirectoryW, Mid(SubDirectoryW, 4))
                            'Here we add the Whole path to the Tag Property for Identification  
                            'later on                     
                            TreeView1.Nodes(y).Nodes(SubDirectoryW).Tag = SubDirectoryW
                            '  InitializeRoot()
                            '======== obie funkcje sa do wyswietlania ikon
                            TreeView1.Nodes(y).Nodes(SubDirectoryW).ImageKey = CacheShellIcon(SubDirectoryW)
                            ''  AddImages(SubDirectoryW)
                        Catch ex As Exception

                        End Try
                    Next

                    y += 1
                End If
            Next
        End If

        'For Each drive As String In DirectoryW.GetLogicalDrives()
        '    TreeView1.Nodes.Add(My.Computer.FileSystem.Drives(x).Name, My.Computer.FileSystem.Drives(x).Name)
        '    x += 1
        'Next drive

    End Sub


    'Private Function parseUsingPDFBox(ByVal filename As String) As String
    '    'LogFile(" Attempting to parse file: " & filename)
    '    Dim doc As PDDocument = New PDDocument()
    '    Dim stripper As PDFTextStripper = New PDFTextStripper()
    '    doc.close()
    '    doc = PDDocument.load(filename)

    '    Dim content As String = "empty"
    '    Try
    '        content = stripper.getText(doc)
    '        doc.close()
    '    Catch ex As Exception
    '        'LogFile(" Error parsing file: " & filename & vbcrlf & ex.Message)
    '    Finally
    '        doc.close()
    '    End Try
    '    'MsgBox(content)
    '    Return content

    'End Function

    'Imports System.Collections.Generic
    'Imports System.ComponentModel
    'Imports System.Data
    'Imports System.Drawing
    'Imports System.Text
    'Imports System.Windows.Forms
    'Imports System.IO
    'Imports System.Threading
    'Imports PDFTech
    'Imports System.Reflection
    'Imports System.Net
    'Imports System.Diagnostics

    'Namespace WindowsApplication1
    '        Partial Public Class Form1
    '            Inherits Form
    '            Public Sub New()
    '                InitializeComponent()
    '            End Sub


    '            Private PageCount As Integer
    '            Private Sub button14_Click(ByVal sender As Object, ByVal e As EventArgs)
    '                If openFileDialog1.ShowDialog() = DialogResult.OK Then
    '                    textBox10.Text = openFileDialog1.FileName
    '                    Dim fi As New FileInfo(textBox10.Text)
    '                    Dim DirectoryWnaem As String = fi.DirectoryWName
    '                    If DirectoryWnaem.EndsWith("\") Then
    '                        DirectoryWnaem = DirectoryWnaem.Replace("\", "")
    '                    End If
    '                    Dim textname As String = DirectoryWnaem & "\" & fi.Name.Replace(fi.Extension, "") & ".txt"
    '                    textBox9.Text = textname
    '                    saveFileDialog1.FileName = textBox9.Text


    '                    Application.DoEvents()
    '                    info = New PDFOperation(textBox10.Text, "")
    '                    PageCount = info.GetPageCount()
    '                    info.Close()
    '                    If info.[Error] <> "" Then
    '                        MessageBox.Show(info.[Error], "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
    '                        PageCount = 0
    '                        numericUpDown1.Minimum = 0
    '                        numericUpDown1.Maximum = PageCount
    '                        numericUpDown2.Minimum = 0
    '                        numericUpDown2.Maximum = PageCount
    '                        numericUpDown1.Value = 0
    '                        numericUpDown2.Value = PageCount
    '                        Return
    '                    Else
    '                        numericUpDown1.Minimum = 1
    '                        numericUpDown1.Maximum = PageCount
    '                        numericUpDown2.Minimum = 1
    '                        numericUpDown2.Maximum = PageCount
    '                        numericUpDown1.Value = 1
    '                        numericUpDown2.Value = PageCount
    '                    End If
    '                    groupBox2.Enabled = True
    '                End If
    '            End Sub

    '            Private Sub button13_Click(ByVal sender As Object, ByVal e As EventArgs)
    '                If saveFileDialog1.ShowDialog() = DialogResult.OK Then
    '                    textBox9.Text = saveFileDialog1.FileName
    '                End If
    '            End Sub

    '            Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
    '                Close()
    '            End Sub

    '            Private Sub btnGenerate_Click(ByVal sender As Object, ByVal e As EventArgs)
    '                Dim textfile As String = textBox9.Text
    '                Application.DoEvents()
    '                Dim lay As LayoutType
    '                If radioButton1.Checked Then
    '                    lay = LayoutType.Flowing
    '                ElseIf radioButton2.Checked Then
    '                    lay = LayoutType.Physical
    '                Else
    '                    lay = LayoutType.Row
    '                End If

    '                Dim typ As EncodingType
    '                If radioButton6.Checked Then
    '                    typ = EncodingType.ANSI
    '                ElseIf radioButton5.Checked Then
    '                    typ = EncodingType.Unicode
    '                Else
    '                    typ = EncodingType.UTF8
    '                End If

    '                If checkBox1.Checked Then
    '                    Dim fi As New FileInfo(textfile)
    '                    Dim DirectoryWnaem As String = fi.DirectoryWName
    '                    If DirectoryWnaem.EndsWith("\") Then
    '                        DirectoryWnaem = DirectoryWnaem.Replace("\", "")
    '                    End If
    '                    textfile = DirectoryWnaem & "\" & fi.Name.Replace(fi.Extension, "") & ".html"
    '                End If

    '                info = New PDFOperation(textBox10.Text, txtUserPass.Text)
    '                info.Progress += New ProgressEventHandler(AddressOf doc_Progress)
    '                info.ExtractText(CInt(numericUpDown1.Value), CInt(numericUpDown2.Value), lay, typ, checkBox1.Checked, checkBox3.Checked, _
    '                 checkBox4.Checked, textfile)
    '                If info.[Error] <> "" Then
    '                    Dim sr2 As New StreamWriter(textfile & ".log", False, Encoding.Unicode)
    '                    sr2.WriteLine(info.[Error])
    '                    sr2.Close()
    '                End If
    '                If checkBox2.Checked Then
    '                    System.Diagnostics.Process.Start(textfile)
    '                End If
    '                info.Close()
    '            End Sub

    '            Private progress As Progress
    '            Private info As PDFOperation
    '            Private Sub doc_Progress(ByVal sender As Object, ByVal arg As ProgressEventArgs)
    '                Select Case arg.Stage
    '                    Case ProgressStage.Starting
    '                        progress = New Progress()
    '                        progress.doc = info
    '                        progress.AllowTransparency = False
    '                        progress.Show()
    '                        progress.Refresh()
    '                        Exit Select
    '                    Case ProgressStage.Running
    '                        Application.DoEvents()
    '                        progress.progressBar1.Value = arg.PercentDone
    '                        Exit Select
    '                    Case ProgressStage.Ending
    '                        progress.Dispose()
    '                        Exit Select
    '                End Select
    '            End Sub

    '            Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
    '                Try
    '                    If Not VistaSecurity.IsVistaOrHigher() AndAlso Not VistaSecurity.IsAdmin() Then
    '                        VistaSecurity.RestartElevated()
    '                    End If
    '                    Dim info As FileVersionInfo = FileVersionInfo.GetVersionInfo(New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll")
    '                    Dim LocalCoreFileVersion As String = info.FileVersion
    '                    info = FileVersionInfo.GetVersionInfo(Me.[GetType]().Assembly.Location)
    '                    Dim LocalVersion As String = info.FileVersion

    '                    Dim asm As Assembly = Assembly.LoadFrom(New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll")
    '                    Dim LocalCoreAssemblyVersion As String = asm.GetName().Version.ToString(4)

    '                    Dim cli As New WebClient()
    '                    Dim b As Byte() = cli.DownloadData("http://www.pdf-technologies.com/downloads/products.ini")
    '                    Dim s As String = ASCIIEncoding.ASCII.GetString(b)
    '                    Dim index As Integer = s.IndexOf("PDFTechLib.dll FileVersion = ")
    '                    Dim last As Integer = s.IndexOf(vbCr & vbLf, index)
    '                    Dim CoreLibraryFileVersion As String = s.Substring(index + 29, last - index - 29)
    '                    index = s.IndexOf("PDFTechLib.dll AssemblyVersion = ")
    '                    last = s.IndexOf(vbCr & vbLf, index)
    '                    Dim CoreLibraryAssemblyVersion As String = s.Substring(index + 33, last - index - 33)
    '                    index = s.IndexOf("PDFToText.exe FileVersion = ")
    '                    last = s.IndexOf(vbCr & vbLf, index)
    '                    Dim PDFToTextVersion As String = s.Substring(index + 28, last - index - 28)

    '                    If LocalCoreAssemblyVersion <> CoreLibraryAssemblyVersion OrElse PDFToTextVersion <> LocalVersion Then
    '                        MessageBox.Show("The updated version of this software is available." & vbCr & vbLf & "Please go to http://www.pdf-technologies.com to download.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '                    ElseIf LocalCoreFileVersion <> CoreLibraryFileVersion Then
    '                        Dim ass As Assembly() = Thread.GetDomain().GetAssemblies()
    '                        Dim loaded As Boolean = False
    '                        For Each asse As Assembly In ass
    '                            If asse.FullName.IndexOf("PDFTechLib") <> -1 Then
    '                                loaded = True
    '                                Exit For
    '                            End If
    '                        Next
    '                        If DialogResult.Yes = MessageBox.Show("The updated version of this software is available. Do you wish to install?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information) Then
    '                            Application.DoEvents()
    '                            Me.Cursor = Cursors.WaitCursor
    '                            If loaded Then
    '                                cli.DownloadFile("http://www.pdf-technologies.com/downloads/PDFTechLib.dll", New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll_")
    '                            Else
    '                                cli.DownloadFile("http://www.pdf-technologies.com/downloads/PDFTechLib.dll", New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll")
    '                            End If
    '                            If loaded Then
    '                                MessageBox.Show("To install updates please restart the application", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '                                Dim key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\PDFTechnologies\PDFToText", True)
    '                                If key Is Nothing Then
    '                                    key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\PDFTechnologies\PDFToText")
    '                                End If
    '                                key.SetValue("NewVersionAvailable", 1)
    '                            End If
    '                            Me.Cursor = Cursors.[Default]
    '                            If Not loaded Then
    '                                MessageBox.Show("The updated version is successfully installed.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '                            End If
    '                        End If
    '                    Else
    '                        MessageBox.Show("Your software is up to date. No patch is required.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '                    End If
    '                    cli.Dispose()
    '                Catch ex As Exception
    '                    Me.Cursor = Cursors.[Default]
    '                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
    '                End Try
    '            End Sub

    '            Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)
    '                Try
    '                    Dim key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\PDFTechnologies\PDFToText", True)
    '                    If key IsNot Nothing AndAlso CInt(key.GetValue("NewVersionAvailable")) = 1 Then
    '                        If Not VistaSecurity.IsVistaOrHigher() AndAlso Not VistaSecurity.IsAdmin() Then
    '                            VistaSecurity.RestartElevated()
    '                        End If
    '                        Me.Cursor = Cursors.WaitCursor
    '                        If File.Exists(New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll_") Then
    '                            File.Delete(New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll")

    '                            File.Move(New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll_", New FileInfo(Me.[GetType]().Assembly.Location).DirectoryW.FullName & "\PDFTechLib.dll")
    '                        End If
    '                        key.SetValue("NewVersionAvailable", 0)
    '                        Me.Cursor = Cursors.[Default]
    '                        MessageBox.Show("The updated version is successfully installed.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '                    End If
    '                Catch ex As Exception
    '                    Me.Cursor = Cursors.[Default]
    '                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
    '                End Try
    '            End Sub

    '            Private Sub textBox10_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
    '                If textBox10.Text.EndsWith(".pdf") Then
    '                    Dim fi As New FileInfo(textBox10.Text)
    '                    Dim DirectoryWnaem As String = fi.DirectoryWName
    '                    If DirectoryWnaem.EndsWith("\") Then
    '                        DirectoryWnaem = DirectoryWnaem.Replace("\", "")
    '                    End If
    '                    Dim textname As String = DirectoryWnaem & "\" & fi.Name.Replace(fi.Extension, "") & ".txt"
    '                    textBox9.Text = textname
    '                    saveFileDialog1.FileName = textBox9.Text

    '                    Application.DoEvents()
    '                    info = New PDFOperation(textBox10.Text, "")
    '                    PageCount = info.GetPageCount()
    '                    info.Close()
    '                    If info.[Error] <> "" Then
    '                        MessageBox.Show(info.[Error], "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
    '                        PageCount = 0
    '                        numericUpDown1.Minimum = 0
    '                        numericUpDown1.Maximum = PageCount
    '                        numericUpDown2.Minimum = 0
    '                        numericUpDown2.Maximum = PageCount
    '                        numericUpDown1.Value = 0
    '                        numericUpDown2.Value = PageCount
    '                        Return
    '                    Else
    '                        numericUpDown1.Minimum = 1
    '                        numericUpDown1.Maximum = PageCount
    '                        numericUpDown2.Minimum = 1
    '                        numericUpDown2.Maximum = PageCount
    '                        numericUpDown1.Value = 1
    '                        numericUpDown2.Value = PageCount
    '                    End If
    '                    groupBox2.Enabled = True
    '                End If
    '            End Sub
    '        End Class
    '    End Namespace
    ''Private Sub RemoveUnReferencedPages(ByVal document As PdfDocument, ByVal referenceString As String)
    ''    ' this procedure removes any pages from the pdf document that do not contain
    ''    ' the reference string

    ''    Dim pageNo As Integer = -1
    ''    Dim strStreamValue As String
    ''    Dim streamValue As Byte()
    ''    Dim keepPageArray As Boolean()
    ''    keepPageArray = New Boolean(document.Pages.Count - 1) {}

    ''    ' iterate through the pages
    ''    For Each page As PdfPage In document.Pages
    ''        pageNo += 1
    ''        strStreamValue = ""

    ''        ' put the stream value for every element on the page in a string variable.
    ''        For i As Integer = 0 To page.Contents.Elements.Count - 1
    ''            Dim stream As PdfDictionary.PdfStream = page.Contents.Elements.GetDictionary(i).Stream
    ''            streamValue = stream.Value
    ''            For Each b As Byte In streamValue
    ''                strStreamValue += ChrW(b)
    ''            Next
    ''        Next
    ''        ' flag those pages that contain the reference value
    ''        keepPageArray(pageNo) = strStreamValue.Contains(referenceString)
    ''    Next

    ''    ' Now, remove the pages we identified.  We're doing this in reverse order
    ''    ' because the deletion of an earlier page moves the rest of the pages up 
    ''    ' on page.  This keeps us from deleting the wrong pages.
    ''    For i As Integer = keepPageArray.Length - 1 To -1 + 1 Step -1
    ''        If Not keepPageArray(i) Then
    ''            Dim deletePage As PdfPage = document.Pages(i)
    ''            document.Pages.Remove(deletePage)
    ''        End If
    ''    Next
    ''End Sub
    '' ''    Public Class PDFParser
    '' ''        ''' BT = Beginning of a text object operator 
    '' ''        ''' ET = End of a text object operator
    '' ''        ''' Td move to the start of next line
    '' ''        '''  5 Ts = superscript
    '' ''        ''' -5 Ts = subscript

    '' ''#Region "Fields"

    '' ''#Region "_numberOfCharsToKeep"
    '' ''        ''' <summary>
    '' ''        ''' The number of characters to keep, when extracting text.
    '' ''        ''' </summary>
    '' ''        Private Shared _numberOfCharsToKeep As Integer = 15
    '' ''#End Region

    '' ''#End Region


    '' ''#Region "ExtractTextFromPDFBytes"
    '' ''        ''' <summary>
    '' ''        ''' This method processes an uncompressed Adobe (text) object 
    '' ''        ''' and extracts text.
    '' ''        ''' </summary>
    '' ''        ''' <param name="input">uncompressed</param>
    '' ''        ''' <returns></returns>
    '' ''        Public Function ExtractTextFromPDFBytes(ByVal input As Byte()) As String
    '' ''            If input Is Nothing OrElse input.Length = 0 Then
    '' ''                Return ""
    '' ''            End If

    '' ''            Try
    '' ''                Dim resultString As String = ""

    '' ''                ' Flag showing if we are we currently inside a text object
    '' ''                Dim inTextObject As Boolean = False

    '' ''                ' Flag showing if the next character is literal 
    '' ''                ' e.g. '\\' to get a '\' character or '\(' to get '('
    '' ''                Dim nextLiteral As Boolean = False

    '' ''                ' () Bracket nesting level. Text appears inside ()
    '' ''                Dim bracketDepth As Integer = 0

    '' ''                ' Keep previous chars to get extract numbers etc.:
    '' ''                Dim previousCharacters As Char() = New Char(_numberOfCharsToKeep - 1) {}
    '' ''                For j As Integer = 0 To _numberOfCharsToKeep - 1
    '' ''                    previousCharacters(j) = " "c
    '' ''                Next


    '' ''                For i As Integer = 0 To input.Length - 1
    '' ''                    Dim c As Char = ChrW(input(i))

    '' ''                    If inTextObject Then
    '' ''                        ' Position the text
    '' ''                        If bracketDepth = 0 Then
    '' ''                            If CheckToken(New String() {"TD", "Td"}, previousCharacters) Then
    '' ''                                resultString += vbLf & vbCr
    '' ''                            Else
    '' ''                                If CheckToken(New String() {"'", "T*", """"}, previousCharacters) Then
    '' ''                                    resultString += vbLf
    '' ''                                Else
    '' ''                                    If CheckToken(New String() {"Tj"}, previousCharacters) Then
    '' ''                                        resultString += " "
    '' ''                                    End If
    '' ''                                End If
    '' ''                            End If
    '' ''                        End If

    '' ''                        ' End of a text object, also go to a new line.
    '' ''                        If bracketDepth = 0 AndAlso CheckToken(New String() {"ET"}, previousCharacters) Then

    '' ''                            inTextObject = False
    '' ''                            resultString += " "
    '' ''                        Else
    '' ''                            ' Start outputting text
    '' ''                            If (c = "("c) AndAlso (bracketDepth = 0) AndAlso (Not nextLiteral) Then
    '' ''                                bracketDepth = 1
    '' ''                            Else
    '' ''                                ' Stop outputting text
    '' ''                                If (c = ")"c) AndAlso (bracketDepth = 1) AndAlso (Not nextLiteral) Then
    '' ''                                    bracketDepth = 0
    '' ''                                Else
    '' ''                                    ' Just a normal text character:
    '' ''                                    If bracketDepth = 1 Then
    '' ''                                        ' Only print out next character no matter what. 
    '' ''                                        ' Do not interpret.
    '' ''                                        If c = "\"c AndAlso Not nextLiteral Then
    '' ''                                            nextLiteral = True
    '' ''                                        Else
    '' ''                                            If ((c >= " "c) AndAlso (c <= "~"c)) OrElse ((c >= 128) AndAlso (c < 255)) Then
    '' ''                                                resultString += c.ToString()
    '' ''                                            End If

    '' ''                                            nextLiteral = False
    '' ''                                        End If
    '' ''                                    End If
    '' ''                                End If
    '' ''                            End If
    '' ''                        End If
    '' ''                    End If

    '' ''                    ' Store the recent characters for 
    '' ''                    ' when we have to go back for a checking
    '' ''                    For j As Integer = 0 To _numberOfCharsToKeep - 2
    '' ''                        previousCharacters(j) = previousCharacters(j + 1)
    '' ''                    Next
    '' ''                    previousCharacters(_numberOfCharsToKeep - 1) = c

    '' ''                    ' Start of a text object
    '' ''                    If Not inTextObject AndAlso CheckToken(New String() {"BT"}, previousCharacters) Then
    '' ''                        inTextObject = True
    '' ''                    End If
    '' ''                Next
    '' ''                Return resultString
    '' ''            Catch
    '' ''                Return ""
    '' ''            End Try
    '' ''        End Function
    '' ''#End Region

    '' ''#Region "CheckToken"
    '' ''        ''' <summary>
    '' ''        ''' Check if a certain 2 character token just came along (e.g. BT)
    '' ''        ''' </summary>
    '' ''        ''' <param name="search">the searched token</param>
    '' ''        ''' <param name="recent">the recent character array</param>
    '' ''        ''' <returns></returns>
    '' ''        Private Function CheckToken(ByVal tokens As String(), ByVal recent As Char()) As Boolean
    '' ''            For Each token As String In tokens
    '' ''                If token.Length > 1 Then
    '' ''                    If (recent(_numberOfCharsToKeep - 3) = token(0)) AndAlso (recent(_numberOfCharsToKeep - 2) = token(1)) AndAlso ((recent(_numberOfCharsToKeep - 1) = " "c) OrElse (recent(_numberOfCharsToKeep - 1) = &HD) OrElse (recent(_numberOfCharsToKeep - 1) = &HA)) AndAlso ((recent(_numberOfCharsToKeep - 4) = " "c) OrElse (recent(_numberOfCharsToKeep - 4) = &HD) OrElse (recent(_numberOfCharsToKeep - 4) = &HA)) Then
    '' ''                        Return True
    '' ''                    End If
    '' ''                Else
    '' ''                    Return False

    '' ''                End If
    '' ''            Next
    '' ''            Return False
    '' ''        End Function
    '' ''#End Region
    '' ''    End Class
    ''Public Function ExtractText() As String
    ''    Dim outputText As [String] = ""
    ''    Try
    ''        Dim inputDocument As PdfDocument = PdfReader.Open(Me._sDirectoryW + Me._sFileName, PdfDocumentOpenMode.[ReadOnly])

    ''        For Each page As PdfPage In inputDocument.Pages
    ''            For index As Integer = 0 To page.Contents.Elements.Count - 1

    ''                Dim stream As PdfDictionary.PdfStream = page.Contents.Elements.GetDictionary(index).Stream
    ''                outputText += New PDFParser().ExtractTextFromPDFBytes(stream.Value)
    ''            Next

    ''        Next
    ''    Catch e As Exception
    ''        Dim oEx As New PDF_ParseException(Me, e)
    ''        oEx.Log()
    ''        oEx.ToPdf(Me._sDirectoryWException)
    ''    End Try
    ''    Return outputText
    ''End Function
    ''    Dim pdfTextRegexp As String = "(T[wdcm*])[\s]*(\[([^\]]*)\]|\((?<text>[^\)]*)\))[\s]*Tj"

    ''    Dim r As PdfDocument = PdfReader.Open(File)
    ''    Dim contents As PdfContents = r.Pages(0).Contents
    ''For Each o As PdfReference In contents.Elements
    ''    Dim c As PdfContent = TryCast(o.Value, PdfContent)
    ''	If c IsNot Nothing Then
    ''    Dim content As String = Encoding.[Default].GetString(c.Stream.Value)
    ''		Using sr As New StringReader(content)
    ''    Dim line As String
    ''			While (InlineAssignHelper(line, sr.ReadLine())) IsNot Nothing
    ''    Dim m As Match = Regex.Match(Line, pdfTextRegexp, RegexOptions.Compiled)
    ''				If m.Success Then
    ''					Debug.WriteLine(m.Groups("text").Value)
    ''				End If
    ''			End While
    ''		End Using
    ''	End If
    ''Next


    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            TextBox5.Text = Date.Today
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ListView1.GridLines = True
        Else
            ListView1.GridLines = False
        End If
    End Sub

    ' działający code

    'Dim nFiles, h As Integer

    'Dim myFileDialog As OpenFileDialog = New OpenFileDialog()
    'Dim xlApp As Excel.Application
    'Dim xlWorkBook As Excel.Workbook
    'Dim xlWorkSheet As Excel.Worksheet

    '    xlApp = New Excel.ApplicationClass
    '    xlApp.Visible = True
    '    xlWorkBook = xlApp.Workbooks.Open(Filename:=My.Application.Info.DirectoryWPath + "\" + "Zeichnungsliste.xls")
    '    xlWorkSheet = xlWorkBook.Worksheets("Zeichnungsliste (2)")
    ''display the cells value B2
    ''=============================================
    '' w pierwszej kolejności należy poslugiwać się itextem. Jeżeli w nazwie znajdzie mi wartość pustą musi przejść innym programem
    ''===============================================
    'Dim folder As String

    'Dim b As Integer = 0

    '    For di As Integer = 0 To ListView1.Items.Count - 1
    '        If ListView1.Items(di).Selected = False Then
    '            ListView1.Items(di).Tag = ListView1.Items(di).Text
    '            folder = ListTab(0) + "\" + CStr(ListView1.Items(di).Tag)
    '            ListTab(0) = ListTab(0).ToString '+ "\" + CStr(ListView1.Items(di).Tag)



    'Dim n As Integer = 0
    'Dim argPath As String
    'Dim mkey As String
    ''  Try
    '            If Not folder Is Nothing AndAlso IO.DirectoryW.Exists(folder) Then

    '                For Each file As String In IO.DirectoryW.GetFiles(folder)
    '                    If Microsoft.VisualBasic.Right(file, 3) = "pdf" Then
    '                        RichTextBox1.Clear()
    ''    Dim oReader As New iTextSharp.text.pdf.PdfReader(ListView1.Items(nFiles).ImageKey)
    'Dim oReader As New iTextSharp.text.pdf.PdfReader(File)
    'Dim i As Integer
    'Dim sOut = ""
    'Dim ss = ""
    '                        For i = 1 To oReader.NumberOfPages
    'Dim its As New iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy

    '                            sOut &= iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(oReader, i, its)

    '                        Next

    ''sOut = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes.(sOut)))
    '                        RichTextBox1.AppendText((Environment.NewLine + sOut.ToString()))
    'Dim xf As Integer = 0
    'Dim x As Integer
    'Dim siteK() As String

    '                        For x = RichTextBox1.Lines.GetLowerBound(0) To RichTextBox1.Lines.GetUpperBound(0)
    '' MessageBox.Show(RichTextBox1.Lines(x))

    '                            ReDim Preserve siteK(xf)
    '                            siteK(xf) = RichTextBox1.Lines(x)
    '                            xf = xf + 1
    '                        Next

    '' Wyszukiwanie tekstu i zapisywanie do pliku xls
    'Dim result1 As New List(Of Integer)
    ''Dim result1 As List(Of Integer) = New List(Of Integer)
    '                        For ig As Integer = 0 To siteK.Length - 1
    '' określa pozycje wystąpienia znaku
    '                            If siteK(ig).Contains("Nr rysunku / Zeichnungs-Nr. / Drawing-No.") Then result1.Add(ig) ' item number 1 -numer rysunku
    '                            If siteK(ig).Contains("Zatwierdził") Then result1.Add(ig) 'item number 2 - rozmiar rysunku
    '                            If siteK(ig).Contains("Approved") Then result1.Add(ig) 'item number 3 - waga częsci
    '                            If siteK(ig).Contains("MASA/MASSE/WEIGHT") Then result1.Add(ig) 'item number 4 - nazwa rysunku
    '                            If siteK(ig).Contains("WERKSTOFF") Then result1.Add(ig) 'item number 5 - index zmian
    '                        Next
    'Dim pos As Integer = result1.Item(0).ToString ' item number 1 -numer rysunku
    'Dim pos1 As Integer = result1.Item(1).ToString 'item number 2 - rozmiar rysunku
    'Dim pos2 As Integer = result1.Item(2).ToString 'item number 3 - waga częsci
    'Dim pos3 As Integer = result1.Item(3).ToString 'item number 4 - nazwa rysunku
    'Dim pos4 As Integer = result1.Item(3).ToString 'item number 5 - index zmian

    'Dim sTable As String() = Nothing  '"" ' item number 1 -numer rysunku
    'Dim sTableq As String() = Nothing ' / ' item number 1 -numer rysunku
    '                        sTable = siteK(pos + 1).Split(" ") ' item number 1 -numer rysunku
    '                        sTableq = siteK(pos + 1).Split("/") ' item number 1 -numer rysunku
    '                        xlWorkSheet.Cells(11 + b, 1) = b + 1 ' Numer pozycji 
    '                        If Microsoft.VisualBasic.Right(ListView1.Items(di).Text, 3) = "dwg" Then
    '                            xlWorkSheet.Cells(11 + b, 2) = sTable(0) ' numer rysunku dwg - number of drawing
    '                        Else
    '                            xlWorkSheet.Cells(11 + b, 2) = "---------------"
    '                        End If

    '                        xlWorkSheet.Cells(11 + b, 3) = sTable(0) '  numer rysunku  pdf


    '                        xlWorkSheet.Cells(11 + h, 4) = siteK(pos4 - 5) ' index zmian
    '                        xlWorkSheet.Cells(11 + b, 5) = siteK(pos1 - 3) ' rozmiar rysunku- drawing size

    '' Nazwa rysunku
    ' ''Dim NameLenght, ff As Integer
    ' ''Dim s As Integer = 0
    ' ''Dim NameLenghtg As Integer = Len(sTable(0)) ' odlicza długość pierwszego wyrazu /  /
    ' ''For ff = sTableq.GetLowerBound(0) To sTableq.GetUpperBound(0)
    ' ''    s += 1
    ' ''Next
    ' ''If Not (sTableq(s - 1)) Is Nothing Then
    ' ''    NameLenght = Len(sTableq(s - 1))
    ' ''Else
    ' ''    NameLenght = Len(sTableq(2))
    ' ''End If
    '' '' odlicza długość trzeciego wyrazu /  /
    ' ''Dim NameLc As Integer = Len(siteK(result.Item(0) + 1))
    ''xlWorkSheet.Cells(11+h, 6) = siteK(result.Item(0) + 1)
    ''xlWorkSheet.Cells(11 + h, 6) = Mid(siteK(result.Item(0) + 1), NameLenghtg + 1, NameLc - NameLenght - NameLenghtg - 1) 'nazwa części - part name
    'Dim NazwaReadTs, fRead As Integer
    'Dim sread As Integer = 0
    'Dim lNameReadTs
    'Dim NameReadT As Integer = Len(siteK(pos3 + 4)) ' odlicza długość pierwszego wyrazu /  /
    'Dim NameReadTs = siteK(pos3 + 4).Split("/")
    '                        If NameReadTs.GetUpperBound(0) > 1 Then ' jeżeli dla pozycji p3 występuje  ciąg  ///

    '                            lNameReadTs = Len(NameReadTs(2))
    ''MsgBox(Mid(readTab(resultT.Item(0) + 1), 1, NameReadT - NazwaReadTs - 1))
    '                            xlWorkSheet.Cells(11 + b, 6) = Mid(siteK(pos3 + 4), 1, NameReadT - lNameReadTs) 'nazwa części - part name
    ' ''Elseif   ' jeżeli dla pozycji p3 występuje  ciąg  /// - wtedy zmieniamy pozycję ciągu w pliku
    ' ''    Dim result2 As New List(Of Integer)
    ' ''    For ig2 As Integer = 0 To siteK.Length - 1
    ' ''        ' określa pozycje wystąpienia znaku
    ' ''        If siteK(ig2).Contains("Nr rysunku / Zeichnungs-Nr. / Drawing-No.") Then result2.Add(ig2) ' item number 1 -nazwa rysunku

    ' ''    Next
    ' ''    Dim posP3 As Integer = result2.Item(0).ToString ' item number 1 -numer rysunku
    ' ''    Dim NameReadT2 As Integer = Len(siteK(posP3 + 1)) ' odlicza długość pierwszego wyrazu /  /
    ' ''    Dim NameReadTs2 = siteK(posP3 + 1).Split("/")
    ' ''    Dim NameReadTs22 As Integer = Len(NameReadTs2(2))
    ' ''    Dim NameReadTs_ = siteK(posP3 + 1).Split(" ")
    ' ''    Dim NameReadTs_2 As Integer = Len(NameReadTs_(0))
    ' ''    xlWorkSheet.Cells(11 + b, 6) = Mid(siteK(posP3 + 1), 1 + NameReadTs_2, NameReadT2 - NameReadTs22 - NameReadTs_2) 'nazwa części - part name
    '                        Else
    'Dim result22 As New List(Of Integer)
    '                            For ig2 As Integer = 0 To siteK.Length - 1
    '' określa pozycje wystąpienia znaku
    '                                If siteK(ig2).Contains("Nr rysunku / Zeichnungs-Nr. / Drawing-No.") Then result22.Add(ig2) ' item number 1 -nazwa rysunku

    '                            Next
    'Dim posP31 As Integer = result22.Item(0).ToString ' item number 1 -numer rysunku
    'Dim NameReadT21 As Integer = Len(siteK(posP31 + 2)) ' odlicza długość pierwszego wyrazu /  /
    'Dim NameReadTs21 = siteK(posP31 + 2).Split("/")
    'Dim NameReadTs221 As Integer = Len(NameReadTs21(2))
    'Dim NameReadTs_1 = siteK(posP31 + 2).Split(" ")
    'Dim NameReadTs_21 As Integer = Len(NameReadTs_1(0))
    '                            xlWorkSheet.Cells(11 + b, 6) = Mid(siteK(posP31 + 2), 1 + NameReadTs_21, NameReadT21 - NameReadTs221 - NameReadTs_21) 'nazwa części - part name
    '                        End If



    'Dim ReadTableqW = siteK(pos2 + 1).Split(" ")
    'Dim rwTab() As String = Nothing
    'Dim fReadWs As Integer
    'Dim sreads As Integer = 0
    '                        For fReadWs = ReadTableqW.GetLowerBound(0) To ReadTableqW.GetUpperBound(0)
    '                            If ReadTableqW(fReadWs) <> "Kg" Then
    '                                ReDim Preserve rwTab(sreads)
    '                                rwTab(sreads) = ReadTableqW(fReadWs)
    '                                sreads += 1
    '                            End If


    '                        Next
    'Dim memorW As String = rwTab(0)
    ' ''Dim NazwaReadW3 = Len(siteK(pos2 + 1))
    ' ''Dim NazwaReadW1 = Len(rwTab(0))
    '' ''Dim NazwaReadW2 = Len(rwTab(1))
    '' ''MsgBox(Microsoft.VisualBasic.Right(siteK(pos2 - 5), NazwaReadW2))
    ' ''
    ' ''Dim fReadWs1 As Integer
    ' ''Dim sreads1 As Integer = 0
    ' ''For fReadWs1 = ReadTableqW.GetLowerBound(0) To ReadTableqW.GetUpperBound(0)
    ' ''    If Not (ReadTableqW(sreads1) Is Nothing) And ReadTableqW(sreads1).ToString = "Kg" Then
    ' ''        memorW = Mid(siteK(pos2 + 1), NazwaReadW1 + 1, NazwaReadW2 + 1)
    ' ''    Else
    ' ''        memorW = Microsoft.VisualBasic.Right(siteK(pos2 + 1), NazwaReadW2)
    ' ''    End If
    ' ''    sreads1 += 1
    ' ''Next
    '                        xlWorkSheet.Cells(11 + b, 7) = memorW 'masa części - weight part - odczyt z pliku i tablicy readTAB
    '                        xlWorkSheet.Cells(11 + b, 8) = folder 'brak folderu

    '                        xlWorkSheet.Cells(11 + b, 9) = sTable(0) '  numer rysunku  pdf do celów inforam
    '                        xlWorkSheet.Cells(11 + b, 10) = 1 ' liczba wystąpień
    '' ''Dim result As New List(Of Integer)
    '' ''For ig As Integer = 0 To siteK.Length - 1
    '' ''    ' określa pozycje wystąpienia znaku
    '' ''    If siteK(ig).Contains("Nr rysunku / Zeichnungs-Nr. / Drawing-No.") Then result.Add(ig) ' item number 1
    '' ''    If siteK(ig).Contains("Approved") Then result.Add(ig) 'item number 2
    '' ''    If siteK(ig).Contains("Zatwierdził") Then result.Add(ig) 'item number 3
    '' ''Next
    '' ''Dim pos As Integer = result.Item(0).ToString
    '' ''Dim sTable As String() = Nothing  '""
    '' ''Dim sTableq As String() = Nothing ' /
    '' ''sTable = siteK(pos + 1).Split(" ")
    '' ''sTableq = siteK(pos + 1).Split("/")

    ' '' '' ok- działa

    ' '' ''edit the cell with new value

    '' ''xlWorkSheet.Cells(11 + h, 2) = sTable(0) ' numer rysunku - number of drawing
    '' ''xlWorkSheet.Cells(11 + h, 3) = sTable(0) '  numer rysunku  pdf

    '' ''xlWorkSheet.Cells(11 + h, 5) = siteK(result.Item(1) - 3) ' rozmiar rysunku drawing size
    '' ''Dim NameLenght, ff As Integer
    '' ''Dim s As Integer = 0
    '' ''Dim NameLenghtg As Integer = Len(sTable(0)) ' odlicza długość pierwszego wyrazu /  /
    '' ''For ff = sTableq.GetLowerBound(0) To sTableq.GetUpperBound(0)
    '' ''    s += 1
    '' ''Next
    '' ''If Not (sTableq(s - 1)) Is Nothing Then
    '' ''    NameLenght = Len(sTableq(s - 1))
    '' ''Else
    '' ''    NameLenght = Len(sTableq(2))
    '' ''End If
    ' '' '' odlicza długość trzeciego wyrazu /  /
    '' ''Dim NameLc As Integer = Len(siteK(result.Item(0) + 1))
    ' '' ''xlWorkSheet.Cells(11+h, 6) = siteK(result.Item(0) + 1)
    '' '' '' xlWorkSheet.Cells(11 + h, 6) = Mid(siteK(result.Item(0) + 1), NameLenghtg + 1, NameLc - NameLenght - NameLenghtg - 1) 'nazwa części - part name
    '' ''xlWorkSheet.Cells(11 + h, 7) = siteK(result.Item(2) + 1) 'masa części - weight part
    '' ''xlWorkSheet.Cells(11 + b, 8) = folder  'brak folderu

    '                        h = h + 1
    '                        b += 1
    '                    End If
    '                Next
    ''End If

    '            ElseIf IO.File.Exists(folder) Then

    ' '' Dim b As Integer
    '' For nFiles = 0 To ListView1.Items.Count - 1
    ''   If Microsoft.VisualBasic.Right(ListView1.Items(nFiles).Text, 3) = "pdf" Then
    '                If Microsoft.VisualBasic.Right(ListView1.Items(di).Text, 3) = "pdf" Then

    '' -------------------------------------------------------------------------------------
    '' Bibliteka ItexSharp
    '' Sprawdzanie tekstu w locie - dla standardowych plików pdf
    ''---------------------------------------------------------------------------------------
    '                    RichTextBox1.Clear()
    ''    Dim oReader As New iTextSharp.text.pdf.PdfReader(ListView1.Items(nFiles).ImageKey)
    'Dim oReader As New iTextSharp.text.pdf.PdfReader(folder)
    'Dim i As Integer
    'Dim sOut = ""
    'Dim sOut1 = ""
    'Dim ss = ""
    'Dim sf As String
    'Dim strText As String = ""

    '                    For i = 1 To oReader.NumberOfPages
    'Dim its As New iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy

    '                        sOut &= iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(oReader, i, its)

    '                    Next
    '' parseUsingPDFBox(folder)

    '' Display the text
    '' Kodowanie do polskich znaków
    'Dim helvetica As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1250, BaseFont.EMBEDDED)
    'Dim plLarge As New iTextSharp.text.Font(helvetica, 16)
    '' Dekodowanie pliku do UTF8 - brak wszystkich polskich znaków
    '                    sOut = (Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(sOut))))
    ''RichTextBox1.AppendText((Environment.NewLine + sOut.ToString()))
    ''sf = DirectCast(Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.[Default], Encoding.UTF8, Encoding.[Default].GetBytes(sOut1))), String)
    '' Zapis do Stringbuildera z uwzględnieniem polskich znaków
    'Dim sb As New StringBuilder()
    '                    sb.AppendFormat(sOut, plLarge) ' do przetestowania
    ''Dim resultsb As String = Encoding.GetEncoding("ISO-8859-2").GetString(System.Text.Encoding.GetEncoding("ISO-8859-2").GetBytes(sOut)) ' polskie znaki
    ''MsgBox(resultsb)
    'Dim line As String()
    'Dim args As String
    '                    line = (sOut.Split(ControlChars.Lf))

    ''Dodanie nowych lini odczytanych po zakonczniu tekstu
    '                    For Each item As String In line
    '                        RichTextBox1.AppendText((item & ControlChars.Lf))
    ''   TextBox7.Text += item & ControlChars.Lf
    '                    Next
    '                    If sOut = Nothing Then
    '                        RichTextBox2.AppendText(("Plik jest obrazem" & folder & ControlChars.Lf))
    '                    Else

    ' ''  RichTextBox1.AppendText((sOut.ToString())) - bez polskich znaków
    ' ''RichTextBox1.AppendText((Environment.NewLine + sb.ToString())) '- z polskimi znakami
    ''RichTextBox1.AppendText((resultsb.ToString())) '- bez polskich znaków

    'Dim xf As Integer = 0
    'Dim x As Integer
    'Dim siteK() As String
    'Dim sit() As String
    '' Odczyt z richboxa i zapis do tablicy  której wysukuje się dane.
    '                        For x = RichTextBox1.Lines.GetLowerBound(0) To RichTextBox1.Lines.GetUpperBound(0)
    '' MessageBox.Show(RichTextBox1.Lines(x))

    '                            ReDim Preserve siteK(xf)
    '                            ReDim Preserve sit(xf)
    '                            siteK(xf) = RichTextBox1.Lines(x)
    '                            xf = xf + 1
    '                        Next
    '' Wyszukiwanie tekstu i zapisywanie do pliku xls
    'Dim result1 As New List(Of Integer)
    ''Dim result1 As List(Of Integer) = New List(Of Integer)
    '                        For ig As Integer = 0 To siteK.Length - 1
    '' określa pozycje wystąpienia znaku
    '                            If siteK(ig).Contains("Nr rysunku / Zeichnungs-Nr. / Drawing-No.") Then result1.Add(ig) ' item number 1 -numer rysunku
    '                            If siteK(ig).Contains("Zatwierdził") Then result1.Add(ig) 'item number 2 - rozmiar rysunku
    '                            If siteK(ig).Contains("Approved") Then result1.Add(ig) 'item number 3 - waga częsci
    '                            If siteK(ig).Contains("MASA/MASSE/WEIGHT") Then result1.Add(ig) 'item number 4 - nazwa rysunku
    '                            If siteK(ig).Contains("WERKSTOFF") Then result1.Add(ig) 'item number 5 - index zmian
    '                        Next
    'Dim pos As Integer = result1.Item(0).ToString ' item number 1 -numer rysunku
    'Dim pos1 As Integer = result1.Item(1).ToString 'item number 2 - rozmiar rysunku
    'Dim pos2 As Integer = result1.Item(2).ToString 'item number 3 - waga częsci
    'Dim pos3 As Integer = result1.Item(3).ToString 'item number 4 - nazwa rysunku
    'Dim pos4 As Integer = result1.Item(3).ToString 'item number 5 - index zmian

    'Dim sTable As String() = Nothing  '"" ' item number 1 -numer rysunku
    'Dim sTableq As String() = Nothing ' / ' item number 1 -numer rysunku
    '                        sTable = siteK(pos + 1).Split(" ") ' item number 1 -numer rysunku
    '                        sTableq = siteK(pos + 1).Split("/") ' item number 1 -numer rysunku

    '' oK działa
    '' Bibiloteka PdfTechLib
    ''--------------------------------------------------------------------------------
    ''  Dim doc As PDDocument = PDDocument.load(folder.ToString)
    ''Dim ghj As PDFTech.PDFAction = PDFTech.PDFAction
    ''If siteK(45) = Nothing Then ' sprawdzenie czy istnieje nazwa rysunku
    ''    Dim readTab As String() = Nothing
    ''    Dim operation As PDFOperation = New PDFOperation(folder, "")
    ''    Dim pageCount As Integer = operation.GetPageCount()

    ''    ' Dim info As PDFDocInfo = operation.DocumentInfo

    ''    Dim info = New PDFOperation(folder.ToString, TextBox6.Text)
    ''    'info.Progress += New ProgressEventHandler(AddressOf doc_Progress)
    ''    info.ExtractText(1, pageCount, LayoutType.Flowing, EncodingType.UTF8, False, True, _
    ''     False, "c:\f.txt")
    ''    If info.Error <> "" Then
    ''        Dim sr2 As New StreamWriter("c:\f.txt" & ".log", False, Encoding.Unicode)
    ''        sr2.WriteLine(info.Error)
    ''        sr2.Close()
    ''    End If
    ''    'If CheckBox2.Checked Then
    ''    '    System.Diagnostics.Process.Start(textfile)
    ''    'End If
    ''    info.Close()

    ''    ' open text file and read its
    ''    Try
    ''        ' Create an instance of StreamReader to read from a file.
    ''        ' The using statement also closes the StreamReader.
    ''        Using sr As New StreamReader("c:\f.txt")
    ''            Dim lines As String
    ''            Dim k As Integer = 0

    ''            ' Read and display lines from the file until the end of
    ''            ' the file is reached.
    ''            Do
    ''                ReDim Preserve readTab(k)
    ''                lines = sr.ReadLine()
    ''                If Not (lines Is Nothing) And lines <> "" And lines <> ControlChars.Lf Then
    ''                    Console.WriteLine(lines)
    ''                    readTab(k) = lines
    ''                    k += 1
    ''                End If

    ''            Loop Until lines Is Nothing
    ''        End Using
    ''    Catch es As Exception
    ''        ' Let the user know what went wrong.
    ''        Console.WriteLine("The file could not be read:")
    ''        Console.WriteLine(es.Message)
    ''    End Try



    ''    ' Dim crypto As PDFCryptoOptions = operation.ProtectionOptions
    ''    ' Dim sg As Integer = operation.GetPageCount()
    ''    'operation.ExtractText(0, pageCount, LayoutType.Flowing, EncodingType.UTF8, False, True, False, "C:\t.txt")
    ''    'operation.Close()



    ''    ' Wyszukiwanie tekstu i zapisywanie do pliku xls
    ''    Dim result As New List(Of Integer)
    ''    For ig As Integer = 0 To siteK.Length - 1
    ''        ' określa pozycje wystąpienia znaku
    ''        If siteK(ig).Contains("Nr rysunku / Zeichnungs-Nr. / Drawing-No.") Then result.Add(ig) ' item number 1
    ''        If siteK(ig).Contains("Approved") Then result.Add(ig) 'item number 2
    ''        If siteK(ig).Contains("Zatwierdził") Then result.Add(ig) 'item number 3
    ''    Next
    ''    Dim pos As Integer = result.Item(0).ToString
    ''    Dim sTable As String() = Nothing  '""
    ''    Dim sTableq As String() = Nothing ' /
    ''    sTable = siteK(pos + 1).Split(" ")
    ''    sTableq = siteK(pos + 1).Split("/")

    ''    ' Wyszukiwanie tesktu z   readTab
    ''    Dim resultT As New List(Of Integer)
    ''    For ig1 As Integer = 0 To readTab.Length - 3
    ''        ' określa pozycje wystąpienia znaku
    ''        'If readTab(ig1).Contains("Ogólne tolerancje wykonania: Allgemeintoleranzen: General tolerances:") Then resultT.Add(ig1) ' item number 1 - numer rysunku pdf
    ''        If readTab(ig1).Contains("Ogólne tolerancje wykonania: Allgemeintoleranzen: General tolerances:") Then resultT.Add(ig1) 'row item number in the table
    ''        If readTab(ig1).Contains("Konstruował Gezeichn./Design. Sprawdził Geprueft/Checked") Then resultT.Add(ig1) 'item number 3 - szuka masy całkowitej urządzenia
    ''    Next
    ''    Dim posT As Integer = resultT.Item(1).ToString
    ''    Dim posT2 As Integer = resultT.Item(0).ToString

    ''    ' Wydzielenie napisów z nazwy znajdujących sie po /
    ''    Dim ReadTableq As String() = Nothing ' /
    ''    Dim ReadTablea As String() = Nothing ' /

    ''    ReadTableq = readTab(posT - 1).Split("/")
    ''    ReadTablea = readTab(posT - 1).Split("")

    ''    ' ok- działa
    ''    ' Tabela Excel
    ''    '-----------------------------------------------------
    ''    ' Zapis danych do komórek w arkuszu kalkulacyjnym
    ''    'edit the cell with new value

    ''    xlWorkSheet.Cells(11 + b, 1) = b + 1 ' Numer pozycji 
    ''    If Microsoft.VisualBasic.Right(ListView1.Items(di).Text, 3) = "dwg" Then
    ''        xlWorkSheet.Cells(11 + b, 2) = sTable(0) ' numer rysunku dwg - number of drawing
    ''    Else
    ''        xlWorkSheet.Cells(11 + b, 2) = "---------------"
    ''    End If

    ''    xlWorkSheet.Cells(11 + b, 3) = sTable(0) '  numer rysunku  pdf
    ''    'xlWorkSheet.Cells(11 + h, 4) ' index zmian
    ''    xlWorkSheet.Cells(11 + b, 5) = siteK(result.Item(1) - 3) ' rozmiar rysunku- drawing size

    ''    ' Nazwa rysunku
    ''    ''Dim NameLenght, ff As Integer
    ''    ''Dim s As Integer = 0
    ''    ''Dim NameLenghtg As Integer = Len(sTable(0)) ' odlicza długość pierwszego wyrazu /  /
    ''    ''For ff = sTableq.GetLowerBound(0) To sTableq.GetUpperBound(0)
    ''    ''    s += 1
    ''    ''Next
    ''    ''If Not (sTableq(s - 1)) Is Nothing Then
    ''    ''    NameLenght = Len(sTableq(s - 1))
    ''    ''Else
    ''    ''    NameLenght = Len(sTableq(2))
    ''    ''End If
    ''    ' '' odlicza długość trzeciego wyrazu /  /
    ''    ''Dim NameLc As Integer = Len(siteK(result.Item(0) + 1))
    ''    'xlWorkSheet.Cells(11+h, 6) = siteK(result.Item(0) + 1)
    ''    'xlWorkSheet.Cells(11 + h, 6) = Mid(siteK(result.Item(0) + 1), NameLenghtg + 1, NameLc - NameLenght - NameLenghtg - 1) 'nazwa części - part name
    ''    Dim NazwaReadTs, fRead As Integer
    ''    Dim sread As Integer = 0
    ''    Dim NameReadT As Integer = Len(ReadTablea(0)) ' odlicza długość pierwszego wyrazu /  /
    ''    For fRead = ReadTableq.GetLowerBound(0) To ReadTableq.GetUpperBound(0)
    ''        sread += 1
    ''    Next
    ''    If Not (ReadTableq(sread - 1)) Is Nothing Then
    ''        NazwaReadTs = Len(ReadTableq(sread - 1))
    ''        'Else
    ''        '    NameReadT = Len(sTableq(2))
    ''    End If
    ''    'MsgBox(Mid(readTab(resultT.Item(0) + 1), 1, NameReadT - NazwaReadTs - 1))
    ''    xlWorkSheet.Cells(11 + b, 6) = Mid(readTab(posT - 1), 1, NameReadT - NazwaReadTs - 1) 'nazwa części - part name


    ''    Dim ReadTableqW = readTab(posT2 + 4).Split(" ")
    ''    Dim rwTab() As String = Nothing
    ''    Dim fReadWs As Integer
    ''    Dim sreads As Integer = 0
    ''    For fReadWs = ReadTableqW.GetLowerBound(0) To ReadTableqW.GetUpperBound(0)
    ''        If ReadTableqW(fReadWs) <> "Kg" Then
    ''            ReDim Preserve rwTab(sreads)
    ''            rwTab(sreads) = ReadTableqW(fReadWs)
    ''            sreads += 1
    ''        End If


    ''    Next
    ''    Dim NazwaReadW3 = Len(readTab(posT2 + 4))
    ''    Dim NazwaReadW1 = Len(rwTab(0))
    ''    Dim NazwaReadW2 = Len(rwTab(1))
    ''    'MsgBox(Microsoft.VisualBasic.Right(readTab(posT2 - 5), NazwaReadW2))
    ''    Dim memorW As String
    ''    Dim fReadWs1 As Integer
    ''    Dim sreads1 As Integer = 0
    ''    For fReadWs1 = ReadTableqW.GetLowerBound(0) To ReadTableqW.GetUpperBound(0)
    ''        If Not (ReadTableqW(sreads1) Is Nothing) And ReadTableqW(sreads1).ToString = "Kg" Then
    ''            memorW = Mid(readTab(posT2 + 4), NazwaReadW1 + 1, NazwaReadW2 + 1)
    ''        Else
    ''            memorW = Microsoft.VisualBasic.Right(readTab(posT2 + 4), NazwaReadW2)
    ''        End If
    ''        sreads1 += 1
    ''    Next
    ''    xlWorkSheet.Cells(11 + b, 7) = memorW 'masa części - weight part - odczyt z pliku i tablicy readTAB
    ''    xlWorkSheet.Cells(11 + b, 8) = "-----------"  'brak folderu

    ''    xlWorkSheet.Cells(11 + b, 9) = sTable(0) '  numer rysunku  pdf do celów inforam
    ''    xlWorkSheet.Cells(11 + b, 10) = 1 ' liczba wystąpień
    ''End If
    ' '' Wyszukiwanie tekstu i zapisywanie do pliku xls
    ''Dim result As New List(Of Integer)
    ''For ig As Integer = 0 To siteK.Length - 1
    ''    ' określa pozycje wystąpienia znaku
    ''    If siteK(ig).Contains("Nr rysunku / Zeichnungs-Nr. / Drawing-No.") Then result.Add(ig) ' item number 1
    ''    If siteK(ig).Contains("Approved") Then result.Add(ig) 'item number 2
    ''    If siteK(ig).Contains("Zatwierdził") Then result.Add(ig) 'item number 3
    ''Next
    ''Dim pos As Integer = result.Item(0).ToString
    ''Dim sTable As String() = Nothing  '""
    ''Dim sTableq As String() = Nothing ' /
    ''sTable = siteK(pos + 1).Split(" ")
    ''sTableq = siteK(pos + 1).Split("/")

    ' '' Wyszukiwanie tesktu z   readTab
    ' ''Dim resultT As New List(Of Integer)
    ' ''For ig1 As Integer = 0 To readTab.Length - 3
    ' ''    ' określa pozycje wystąpienia znaku
    ' ''    'If readTab(ig1).Contains("Ogólne tolerancje wykonania: Allgemeintoleranzen: General tolerances:") Then resultT.Add(ig1) ' item number 1 - numer rysunku pdf
    ' ''    If readTab(ig1).Contains("Ogólne tolerancje wykonania: Allgemeintoleranzen: General tolerances:") Then resultT.Add(ig1) 'row item number in the table
    ' ''    If readTab(ig1).Contains("Konstruował Gezeichn./Design. Sprawdził Geprueft/Checked") Then resultT.Add(ig1) 'item number 3 - szuka masy całkowitej urządzenia
    ' ''Next
    ' ''Dim posT As Integer = resultT.Item(1).ToString
    ' ''Dim posT2 As Integer = resultT.Item(0).ToString

    '' '' Wydzielenie napisów z nazwy znajdujących sie po /
    ' ''Dim ReadTableq As String() = Nothing ' /
    ' ''Dim ReadTablea As String() = Nothing ' /

    ' ''ReadTableq = readTab(posT - 1).Split("/")
    ' ''ReadTablea = readTab(posT - 1).Split("")

    ' '' ok- działa
    ' '' Tabela Excel
    ' ''-----------------------------------------------------
    ' '' Zapis danych do komórek w arkuszu kalkulacyjnym
    ' ''edit the cell with new value

    '                        xlWorkSheet.Cells(11 + b, 1) = b + 1 ' Numer pozycji 
    '                        If Microsoft.VisualBasic.Right(ListView1.Items(di).Text, 3) = "dwg" Then
    '                            xlWorkSheet.Cells(11 + b, 2) = sTable(0) ' numer rysunku dwg - number of drawing
    '                        Else
    '                            xlWorkSheet.Cells(11 + b, 2) = "---------------"
    '                        End If

    '                        xlWorkSheet.Cells(11 + b, 3) = sTable(0) '  numer rysunku  pdf


    '                        xlWorkSheet.Cells(11 + h, 4) = siteK(pos4 - 5) ' index zmian
    '                        xlWorkSheet.Cells(11 + b, 5) = siteK(pos1 - 3) ' rozmiar rysunku- drawing size

    '' Nazwa rysunku
    ' ''Dim NameLenght, ff As Integer
    ' ''Dim s As Integer = 0
    ' ''Dim NameLenghtg As Integer = Len(sTable(0)) ' odlicza długość pierwszego wyrazu /  /
    ' ''For ff = sTableq.GetLowerBound(0) To sTableq.GetUpperBound(0)
    ' ''    s += 1
    ' ''Next
    ' ''If Not (sTableq(s - 1)) Is Nothing Then
    ' ''    NameLenght = Len(sTableq(s - 1))
    ' ''Else
    ' ''    NameLenght = Len(sTableq(2))
    ' ''End If
    '' '' odlicza długość trzeciego wyrazu /  /
    ' ''Dim NameLc As Integer = Len(siteK(result.Item(0) + 1))
    ''xlWorkSheet.Cells(11+h, 6) = siteK(result.Item(0) + 1)
    ''xlWorkSheet.Cells(11 + h, 6) = Mid(siteK(result.Item(0) + 1), NameLenghtg + 1, NameLc - NameLenght - NameLenghtg - 1) 'nazwa części - part name
    'Dim NazwaReadTs, fRead As Integer
    'Dim sread As Integer = 0
    'Dim lNameReadTs
    'Dim NameReadT As Integer = Len(siteK(pos3 + 4)) ' odlicza długość pierwszego wyrazu /  /
    'Dim NameReadTs = siteK(pos3 + 4).Split("/")
    '                        If NameReadTs.GetUpperBound(0) > 1 Then ' jeżeli dla pozycji p3 występuje  ciąg  ///

    '                            lNameReadTs = Len(NameReadTs(2))
    ''MsgBox(Mid(readTab(resultT.Item(0) + 1), 1, NameReadT - NazwaReadTs - 1))
    '                            xlWorkSheet.Cells(11 + b, 6) = Mid(siteK(pos3 + 4), 1, NameReadT - lNameReadTs) 'nazwa części - part name
    '                        Else ' jeżeli dla pozycji p3 występuje  ciąg  /// - wtedy zmieniamy pozycję ciągu w pliku
    'Dim result2 As New List(Of Integer)
    '                            For ig2 As Integer = 0 To siteK.Length - 1
    '' określa pozycje wystąpienia znaku
    '                                If siteK(ig2).Contains("Nr rysunku / Zeichnungs-Nr. / Drawing-No.") Then result2.Add(ig2) ' item number 1 -nazwa rysunku

    '                            Next
    'Dim posP3 As Integer = result2.Item(0).ToString ' item number 1 -numer rysunku
    'Dim NameReadT2 As Integer = Len(siteK(posP3 + 2)) ' odlicza długość pierwszego wyrazu /  /
    'Dim NameReadTs2 = siteK(posP3 + 2).Split("/")
    'Dim NameReadTs22 As Integer = Len(NameReadTs2(2))
    'Dim NameReadTs_ = siteK(posP3 + 2).Split(" ")
    'Dim NameReadTs_2 As Integer = Len(NameReadTs_(0))
    '                            xlWorkSheet.Cells(11 + b, 6) = Mid(siteK(posP3 + 2), 1 + NameReadTs_2, NameReadT2 - NameReadTs22 - NameReadTs_2) 'nazwa części - part name
    '                        End If



    'Dim ReadTableqW = siteK(pos2 + 1).Split(" ")
    'Dim rwTab() As String = Nothing
    'Dim fReadWs As Integer
    'Dim sreads As Integer = 0
    '                        For fReadWs = ReadTableqW.GetLowerBound(0) To ReadTableqW.GetUpperBound(0)
    '                            If ReadTableqW(fReadWs) <> "Kg" Then
    '                                ReDim Preserve rwTab(sreads)
    '                                rwTab(sreads) = ReadTableqW(fReadWs)
    '                                sreads += 1
    '                            End If


    '                        Next
    'Dim memorW As String = rwTab(0)
    ' ''Dim NazwaReadW3 = Len(siteK(pos2 + 1))
    ' ''Dim NazwaReadW1 = Len(rwTab(0))
    '' ''Dim NazwaReadW2 = Len(rwTab(1))
    '' ''MsgBox(Microsoft.VisualBasic.Right(siteK(pos2 - 5), NazwaReadW2))
    ' ''
    ' ''Dim fReadWs1 As Integer
    ' ''Dim sreads1 As Integer = 0
    ' ''For fReadWs1 = ReadTableqW.GetLowerBound(0) To ReadTableqW.GetUpperBound(0)
    ' ''    If Not (ReadTableqW(sreads1) Is Nothing) And ReadTableqW(sreads1).ToString = "Kg" Then
    ' ''        memorW = Mid(siteK(pos2 + 1), NazwaReadW1 + 1, NazwaReadW2 + 1)
    ' ''    Else
    ' ''        memorW = Microsoft.VisualBasic.Right(siteK(pos2 + 1), NazwaReadW2)
    ' ''    End If
    ' ''    sreads1 += 1
    ' ''Next
    '                        xlWorkSheet.Cells(11 + b, 7) = memorW 'masa części - weight part - odczyt z pliku i tablicy readTAB
    '                        xlWorkSheet.Cells(11 + b, 8) = "-----------"  'brak folderu

    '                        xlWorkSheet.Cells(11 + b, 9) = sTable(0) '  numer rysunku  pdf do celów inforam
    '                        xlWorkSheet.Cells(11 + b, 10) = 1 ' liczba wystąpień
    '                    End If
    '                End If
    '            End If
    '            b += 1
    '            h += 1
    '' Next
    '        End If


    '' Dodawanie danych dl excela
    '                xlWorkSheet.Cells(2, 3) = TextBox7.Text ' numer  kontraktu
    '                xlWorkSheet.Cells(3, 6) = TextBox1.Text ' nazwa  kontraktu
    '                xlWorkSheet.Cells(5, 1) = "Klient/Kunde:" & TextBox2.Text ' nazwa klienta
    '                xlWorkSheet.Cells(6, 1) = "Budowa/ Baustelle:" & TextBox3.Text ' nazwa budowy
    '                xlWorkSheet.Cells(7, 1) = "Wykonał/ausgeführt von:" & TextBox4.Text ' nazwa budowy
    '                xlWorkSheet.Cells(8, 1) = "Data/Daten:" & TextBox5.Text ' nazwa budowy
    ''Next
    ''    Next
    ''Catch ex As Exception
    ''  MsgBox(ex.Message)
    '' End Try

    ''Exit For
    '' end if

    ''End If
    '' h += 1
    ''   Next

    ''End If
    '    Next

    '    xlWorkBook.Close()
    '    xlApp.Quit()
    '    xlWorkBook = Nothing
    '    xlApp = Nothing
    'End Sub


    '    'Test file name
    '    Dim TestFile As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Test.pdf")

    '    'Standard iTextSharp setup
    '			Using fs As New FileStream(TestFile, FileMode.Create, FileAccess.Write, FileShare.None)
    '				Using doc As New Document(PageSize.LETTER)
    '					Using w As PdfWriter = PdfWriter.GetInstance(doc, fs)
    '    'Open the document for writing
    '						doc.Open()

    '    'Will hold our current x,y coordinates;
    '    Dim curY As Single
    '    Dim curX As Single

    '    'Add a paragraph
    '						doc.Add(New Paragraph("It was the best of times"))

    '    'Get the current Y value
    '						curY = w.GetVerticalPosition(True)

    '    'The current X is just the left margin
    '						curX = doc.LeftMargin

    '    'Set a color fill
    '						w.DirectContent.SetRGBColorStroke(0, 0, 0)
    '    'Set the x,y of where to start drawing
    '						w.DirectContent.MoveTo(curX, curY)
    '    'Draw a line
    '						w.DirectContent.LineTo(doc.PageSize.Width - doc.RightMargin, curY)
    '    'Fill the line in
    '						w.DirectContent.Stroke()

    '    'Add another paragraph
    '						doc.Add(New Paragraph("It was the word of times"))

    '    'Repeat the above. curX never really changes unless you modify the document's margins
    '						curY = w.GetVerticalPosition(True)

    '						w.DirectContent.SetRGBColorStroke(0, 0, 0)
    '						w.DirectContent.MoveTo(curX, curY)
    '						w.DirectContent.LineTo(doc.PageSize.Width - doc.RightMargin, curY)
    '						w.DirectContent.Stroke()


    '    'Close the document
    '						doc.Close()
    '					End Using
    '				End Using
    '			End Using

    '			Me.Close()
    '		End Sub
    'End Class
    Private objapprenticeServerApp As New Inventor.ApprenticeServerComponent
    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        '    'setup parameters for the open dialog

        'With OpenFileDialog1
        '    .InitialDirectory = objapprenticeServerApp.FileLocations.Workspace
        '    .Title = "Select Assembly file"
        '    .DefaultExt = ".idw"
        '    .Filter = "Inventor Assembly File (*.idw) | *.idw"
        '    .ShowDialog()
        'End With

        ' Usuwanie istniejacego i nie wyłaczonego procesu drukarki
        Dim pProcess3() As Process = System.Diagnostics.Process.GetProcessesByName("PDFCreator")

        For Each p As Process In pProcess3
            p.Kill()
        Next

        ' Get the current process.
        Dim currentProcess As Process = Process.GetCurrentProcess()

        ' Get all processes running on the local computer.
        Dim localAll As Process() = Process.GetProcesses()

        'open the specified file
        Dim objapprenticeServerDocument As Inventor.ApprenticeServerDocument
        Dim Inventor As Inventor.Application
        Dim strDocName As String
        Dim folder As String
        For di As Integer = 0 To ListView1.Items.Count - 1
            If ListView1.Items(di).Selected = False Then
                If ListView1.Items(di).Text.Contains(".idw") = True Then 'Or ListView1.Items(di).Text.Contains(".dxf") = False Then
                    If ListView1.Items(di).Text.Contains(".dwf") = False Then
                        ListView1.Items(di).Tag = ListView1.Items(di).Text
                        folder = ListTab(0) + "\" + CStr(ListView1.Items(di).Tag)

                        'strDocName = objapprenticeServerApp.FileManager.GetFullDocumentName(txtFileName.Text) ', LODRep.Text)
                        objapprenticeServerDocument = objapprenticeServerApp.Open(folder)
                        'On Error GoTo Err_Control
                        Dim PDFCreator1
                        Dim killit
                        Dim numsheets As Integer
                        Dim InitPrinter

                        numsheets = 0

                        ' Set reference to active drawing
                        Dim oDrgDoc As Inventor.ApprenticeServerDocument
                        oDrgDoc = objapprenticeServerDocument

                        ' Set reference to drawing print manager
                        Dim oDrgPrintMgr As ApprenticeDrawingPrintManager
                        oDrgPrintMgr = oDrgDoc.PrintManager

                        'Read printer so it can be set back
                        InitPrinter = oDrgPrintMgr.Printer

                        If objapprenticeServerDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                            'PDFCreator1 = New PDFCreator.clsPDFCreator
                            PDFCreator1 = CreateObject("PDFCreator.clsPDFCreator")

                            With PDFCreator1
                                If .cStart("/NoProcessingAtStartup") = False Then

                                    Dim pProcess() As Process = System.Diagnostics.Process.GetProcessesByName("PDFCreator")

                                    For Each p As Process In pProcess
                                        p.Kill()
                                    Next
                                    'killit = Shell("taskkill /f/im PDFCreator.exe/t", AppWinStyle.Hide)
                                    MsgBox("There was an error starting the pdf printer, please try (click) again!")
                                    Debug.Print("Can't initialize PDFCreator.")
                                    Exit Sub
                                End If
                            End With

                            Debug.Print("PDFCreator initialized.")

                            'Set some settings and clear queue
                            With PDFCreator1
                                .cOption("UseAutosave") = 1
                                .cOption("UseAutosaveDirectory") = 1
                                .cOption("AutosaveFormat") = 0 ' 0 = PDF
                                .cClearCache()
                            End With

                            ' Set the printer to PDFCreator
                            oDrgPrintMgr.Printer = "PDFCreator"

                            Dim sht As Sheet

                            For Each sht In oDrgDoc.Sheets
                                'sht.Activate()
                                'MsgBox(sht)


                                'Set the paper size , scale and orientation
                                oDrgPrintMgr.ScaleMode = Global.Inventor.PrintScaleModeEnum.kPrintFullScale ' kPrintBestFitScale
                                ' Change the paper size to a custom size. The units are in centimeters.
                                Dim shtsize As Long
                                Dim shtorientation As Long
                                shtsize = sht.Size
                                shtorientation = sht.Orientation
                                oDrgPrintMgr.PaperSize = Global.Inventor.PaperSizeEnum.kPaperSizeCustom
                                If shtsize = 9993 Then ' A0
                                    oDrgPrintMgr.PaperHeight = 84.1
                                    oDrgPrintMgr.PaperWidth = 118.9
                                ElseIf shtsize = 9994 Then ' A1
                                    oDrgPrintMgr.PaperHeight = 59.4
                                    oDrgPrintMgr.PaperWidth = 84.1
                                ElseIf shtsize = 9995 Then ' A2
                                    oDrgPrintMgr.PaperHeight = 42
                                    oDrgPrintMgr.PaperWidth = 59.4
                                ElseIf shtsize = 9996 Then ' A3
                                    oDrgPrintMgr.PaperHeight = 29.7
                                    oDrgPrintMgr.PaperWidth = 42
                                ElseIf shtsize = 9997 Then ' A4
                                    oDrgPrintMgr.PaperHeight = 21
                                    oDrgPrintMgr.PaperWidth = 29.7
                                End If
                                oDrgPrintMgr.PrintRange = PrintRangeEnum.kPrintAllSheets
                                If shtorientation = 10242 Then 'Landscape
                                    oDrgPrintMgr.Orientation = Global.Inventor.PrintOrientationEnum.kLandscapeOrientation
                                ElseIf shtorientation = 10243 Then 'Portrait
                                    oDrgPrintMgr.Orientation = Global.Inventor.PrintOrientationEnum.kPortraitOrientation
                                End If

                                oDrgPrintMgr.AllColorsAsBlack = True
                                oDrgPrintMgr.Rotate90Degrees = True
                                With PDFCreator1
                                    .cOption("AutosaveDirectory") = TextBox10.Text '"\\Data\PMC\PDF-arkiv\Dokumentation-til-godkendelse\"
                                    .cOption("AutosaveFilename") = ListView1.Items(di).Text
                                End With

                                oDrgPrintMgr.SubmitPrint()

                                System.Threading.Thread.Sleep(1000)

                                numsheets = numsheets + 1
                            Next
                        Else
                            MsgBox("You aren't using an Inventor drawing!")
                            Exit Sub
                        End If

                        'Wait until all prints are in queue
                        Do Until PDFCreator1.cCountOfPrintjobs = numsheets
                            System.Windows.Forms.Application.DoEvents()
                            System.Threading.Thread.Sleep(1000)
                        Loop

                        'Combine sheets in queue to one pdf
                        With PDFCreator1
                            .cCombineAll()
                            .cPrinterStop = False
                        End With

                        'Wait until job done
                        Do Until PDFCreator1.cCountOfPrintjobs = 0
                            System.Windows.Forms.Application.DoEvents()
                            System.Threading.Thread.Sleep(1000)
                        Loop

                        'Set back to first sheet and set printer back
                        'oDrgDoc.Sheets(1).Activate()
                        oDrgPrintMgr.Printer = InitPrinter

                        'Clean up
                        'PDFCreator1.cClose() ''
                        'PDFCreator1.cStop = True
                        PDFCreator1.cClose()

                        Dim pProcess2() As Process = System.Diagnostics.Process.GetProcessesByName("PDFCreator")

                        For Each p As Process In pProcess2
                            p.Kill()
                        Next
                        'killit = Shell("taskkill /f/im PDFCreator.exe /t", AppWinStyle.Hide)
                        oDrgDoc = Nothing
                        oDrgPrintMgr = Nothing
                    End If
                End If
            End If
        Next

Exit_Here:
        Exit Sub
Err_Control:
        Select Case Err.Number
            'Add your Case selections here
            'Case Is = 1000
            '    'Handle error
            '    Err.Clear()
            '    Resume Exit_Here
            Case Else
                MsgBox(Err.Number & ", " & Err.Description, , "PlotPdf")
                Err.Clear()
                Resume Exit_Here
        End Select


    End Sub

    Private Sub createDWG(ByVal oApp As Object, ByVal dwgAddIn As Inventor.TranslatorAddIn, ByVal fNAME As String, ByVal sPath As String)

        Dim map As Inventor.NameValueMap
        Dim context As Inventor.TranslationContext
        Dim trans As Inventor.TransientObjects

        trans = oApp.TransientObjects
        map = trans.CreateNameValueMap
        context = trans.CreateTranslationContext
        context.Type = IOMechanismEnum.kFileBrowseIOMechanism
        Dim b As Boolean
        Dim file As Inventor.DataMedium
        file = trans.CreateDataMedium

        b = dwgAddIn.HasSaveCopyAsOptions(file, context, map)

        file.FileName = fNAME
        'specify ini file from where the setting will be pickup

        map.Value("Export_Acad_IniFile") = sPath & "dwgconfig.ini"
        dwgAddIn.SaveCopyAs(oApp.ActiveDocument, context, map, file)

    End Sub
    Function PublishToDWG(ByVal fileName As String) As Boolean

        ' Get the DWG translator Add-In.

        Dim DWGAddIn As Inventor.TranslatorAddIn
        DWGAddIn = oInvApp.ApplicationAddIns.ItemById("{C24E3AC2-122E-11D5-8E91-0010B541CD80}")

        'Set a reference to the active document (the document to be published).
        Dim oDocument As Document
        oDocument = oInvApp.ActiveEditDocument

        Dim oContext As Inventor.TranslationContext
        oContext = oInvApp.TransientObjects.CreateTranslationContext
        'oContext.Type = kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As Inventor.NameValueMap
        oOptions = oInvApp.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As Inventor.DataMedium
        oDataMedium = oInvApp.TransientObjects.CreateDataMedium

        ' Check whether the translator has 'SaveCopyAs' options
        If DWGAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then

            Dim strIniFile As String
            strIniFile = "C:\Test.ini"
            ' Create the name-value that specifies the ini file to use.
            oOptions.Value("Export_Acad_IniFile") = strIniFile
        End If

        'Set the destination file name
        oDataMedium.FileName = fileName '"c:\temp\dwgout.dwg"

        'Publish document.
        'Call DWGAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
        Call DWGAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)

    End Function

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click


        ' Create a new NameValueMap object


        Dim oInvApp As Inventor.Application
        'Launch Inventor
        oInvApp = CreateObject("Inventor.Application")
        oInvApp.Visible = True
        'Set Startup options.
        oInvApp.GeneralOptions.ShowStartupDialog = False
        Dim folder As String
        For di As Integer = 0 To ListView1.Items.Count - 1
            If ListView1.Items(di).Selected = False And ListView1.Items(di).Text.Contains(".idw") = True Then
                ListView1.Items(di).Tag = ListView1.Items(di).Text
                folder = ListTab(0) + "\" + CStr(ListView1.Items(di).Tag)


                Dim oDoc As Inventor.DrawingDocument
                Dim oDocOpenOptions As NameValueMap
                oDocOpenOptions = oInvApp.TransientObjects.CreateNameValueMap()

                ' Set the option to defer the drawing update.
                Call oDocOpenOptions.Add("DeferUpdates", True)

                ' Open the document, suppressing the warning dialog.
                oInvApp.SilentOperation = True
                'Dim oDoc As DrawingDocument

                oDoc = oInvApp.Documents.OpenWithOptions(folder, oDocOpenOptions, True)
                'oDoc.SilentOperation = False
                ' User iteration is needed in this point

                Dim addIns As ApplicationAddIns
                addIns = oDoc.Parent.ApplicationAddIns
                Dim dwgAddIn As TranslatorAddIn
                Dim i As Integer
                For i = 1 To addIns.Count
                    If addIns(i).AddInType = ApplicationAddInTypeEnum.kTranslationApplicationAddIn Then
                        If addIns(i).Description Like "*DWG*" Then
                            dwgAddIn = addIns.Item(i)
                            Exit For
                        End If
                    End If
                Next i
                'Activate AddIns
                dwgAddIn.Activate()
                Dim fNAME As String
                fNAME = oDoc.FullFileName
                fNAME = Microsoft.VisualBasic.Left(fNAME, Len(fNAME) - 3) & "dwg"
                Dim iPath As Integer
                Dim sPath As String
                iPath = InStrRev(fNAME, "\")
                sPath = Microsoft.VisualBasic.Left(fNAME, iPath)
                Call createDWG(oDoc.Parent, dwgAddIn, fNAME, sPath)

                oDoc.Close(False)
            End If
        Next
        oInvApp.Quit()
        oInvApp = Nothing

    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        ' First create a FolderBrowserDialog object
        Dim FolderBrowserDialog1 As New FolderBrowserDialog

        ' Then use the following code to create the Dialog window
        ' Change the .SelectedPath property to the default location
        With FolderBrowserDialog1
            ' Desktop is the root folder in the dialog.
            .RootFolder = Environment.SpecialFolder.Desktop
            ' Select the C:\Windows DirectoryW on entry.
            .SelectedPath = "c:\windows"
            ' Prompt the user with a custom message.
            .Description = "Select the source DirectoryW"
            If .ShowDialog = DialogResult.OK Then
                TextBox10.Text = .SelectedPath & "\"
                ' Display the selected folder if the user clicked on the OK button.
                'MessageBox.Show(.SelectedPath)
            End If
        End With
    End Sub



    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        Dictionary.Show()
    End Sub
    Private Sub TreeView1_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Dim FileExtension As String
        Dim FolderExtension As String
        Dim SubItemIndex As Integer
        Dim SubItemIndexs As Integer
        Dim DateMod As String
        Dim DateMods As String

        Dim newSelected As TreeNode = e.Node
        ListView1.Items.Clear()
        Dim nodeDirInfo As DirectoryInfo = New DirectoryInfo(newSelected.Tag)
        Dim subItems() As ListViewItem.ListViewSubItem
        Dim subItemsD() As ListViewItem.ListViewSubItem
        Dim item As ListViewItem = Nothing
        ListTab(0) = CStr(TreeView1.SelectedNode.Tag)
        Dim n As Integer = 0
        Try
            For Each nodeDirInfo In nodeDirInfo.GetDirectories()

                FolderExtension = IO.Path.GetExtension(nodeDirInfo.Name)
                DateMods = IO.Directory.GetLastWriteTime(nodeDirInfo.Name)

                item = New ListViewItem(nodeDirInfo.Name, CacheShellIcon(nodeDirInfo.FullName))
                'ListView1.Items.Add(nodeDirInfo.Name.Substring(nodeDirInfo.Name.LastIndexOf("\"c) + 1), mkey)
                subItems = New ListViewItem.ListViewSubItem() _
                    {New ListViewItem.ListViewSubItem(item, ""), _
                    New ListViewItem.ListViewSubItem(item, "Folder File"), _
                    New ListViewItem.ListViewSubItem(item, _
                     IO.File.GetLastWriteTime(nodeDirInfo.FullName).ToString())}

                item.SubItems.AddRange(subItems)
                ListView1.Items.Add(item)
            Next
            Dim folder As String = CStr(TreeView1.SelectedNode.Tag)
            Dim fileb As FileInfo
            For Each filed As String In IO.Directory.GetFiles(folder)
                ' For Each fileb In nodeDirInfo.GetFiles()
                FileExtension = IO.Path.GetExtension(filed)
                Dim FileSize As Double
                FileSize = Math.Round(Module1.GetSizeKB(filed.ToString), 0)
                DateMod = IO.File.GetLastWriteTime(filed).ToString()
                item = New ListViewItem(filed.Substring(filed.LastIndexOf("\"c) + 1), CacheShellIcon(filed))
                subItems = New ListViewItem.ListViewSubItem() _
                {New ListViewItem.ListViewSubItem(item, FileSize.ToString & Chr(32) & "KB"), _
                    New ListViewItem.ListViewSubItem(item, FileExtension.ToString() & Chr(32) & "File"), _
                    New ListViewItem.ListViewSubItem(item, _
                    DateMod.ToString)}

                'AddImages(fileb.FullName)
                If FileExtension.ToString() = "" Then
                    item = New ListViewItem(filed.Substring(filed.LastIndexOf("\"c) + 1), 5)
                    subItemsD = New ListViewItem.ListViewSubItem() _
                    {New ListViewItem.ListViewSubItem(item, FileSize.ToString & Chr(32) & "KB"), _
                    New ListViewItem.ListViewSubItem(item, "SYS. File"), _
                    New ListViewItem.ListViewSubItem(item, _
                    DateMod.ToString)}
                    item.SubItems.AddRange(subItemsD)
                    ListView1.Items.Add(item)
                Else
                    item.SubItems.AddRange(subItems)
                    ListView1.Items.Add(item)
                End If
                SubItemIndex = SubItemIndex + 1
                n += 1
            Next filed
        Catch ex As Exception
            Debug.Write(ex.Message)

        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        ' First create a FolderBrowserDialog object
        Dim FolderBrowserDialog1 As New FolderBrowserDialog

        ' Then use the following code to create the Dialog window
        ' Change the .SelectedPath property to the default location
        With FolderBrowserDialog1
            ' Desktop is the root folder in the dialog.
            .RootFolder = Environment.SpecialFolder.Desktop
            ' Select the C:\Windows DirectoryW on entry.
            .SelectedPath = "c:\windows"
            ' Prompt the user with a custom message.
            .Description = "Select the source DirectoryW"
            If .ShowDialog = DialogResult.OK Then
                TextBox6.Text = .SelectedPath & "\"
                ' Display the selected folder if the user clicked on the OK button.
                'MessageBox.Show(.SelectedPath)
            End If
        End With
    End Sub

    Private Sub ToolStripDropDownButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton1.Click
        Translation.Show()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        ' Splitcontainer is visible  or hidden
        If SplitContainer2.Panel2Collapsed = True Then
            SplitContainer2.Panel2Collapsed = False
        Else
            SplitContainer2.Panel2Collapsed = True
        End If
    End Sub


    Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox9.CheckedChanged
        'PdfManipulation2.AddWatermarkText()
    End Sub
    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged

        '    If (Not errorFound) Then
        If (Me.CheckBox3.Checked = True) Then

            Dim fs As New FileStream(LOGIN_DATA_FILE, FileMode.Create)

            Try
                ' Construct a BinaryFormatter and use it 
                ' to serialize the data to the stream.
                Dim formatter As New BinaryFormatter

                ' Construct a Version1Type object and serialize it.
                Dim P As New TestSimpleObject((TextBox7.Text), (TextBox1.Text), (TextBox2.Text), (TextBox3.Text), (TextBox4.Text), (TextBox5.Text), (TextBox11.Text), (TextBox6.Text), (TextBox10.Text), (TextBox13.Text), CheckBox2.Checked)
                'Dim Persons As New ArrayList

                formatter.Serialize(fs, P)
                MessageBox.Show("All data were recorded to memory.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch v As Runtime.Serialization.SerializationException
                Console.WriteLine("Failed to serialize. Reason: " & v.Message)
                Throw
            Finally
                fs.Close()
            End Try
        Else

            If (File.Exists(LOGIN_DATA_FILE)) Then
                ' delete saved login info if the user doesn't have the
                ' 'remember' box checked
                MessageBox.Show("All data were delated with memory.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                File.Delete(LOGIN_DATA_FILE)
            End If
            'End If
            '        

        End If
    End Sub

    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        HelpForm.Show()
    End Sub

    Private Sub CheckBox16_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox16.CheckedChanged
        If CheckBox16.Checked = True Then
            Dim dt As New Data.DataTable()
            'Using cn As New SqlClient.SqlConnection(ConfigurationManager.ConnectionStrings("ConsoleApplication3.Properties.Settings.daasConnectionString").ConnectionString)
            Using cn As New SQLite.SQLiteConnection("Data Source=" & DirectoryW & "\TranslateBase.s3db;")
                cn.Open()
                Dim SQLcommand As SQLite.SQLiteCommand
                SQLcommand = cn.CreateCommand

                ''Delete Last Created Record from TranslateBase
                SQLcommand.CommandText = "DELETE FROM PartLIst"
                SQLcommand.ExecuteNonQuery()
                SQLcommand.Dispose()
                cn.Close()
                MessageBox.Show(" Deleted all records.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        Else
        End If
    End Sub

    Private Sub CheckBox17_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox17.CheckedChanged
        Dim FolderBrowserDialog1 As New FolderBrowserDialog

        ' Then use the following code to create the Dialog window
        ' Change the .SelectedPath property to the default location
        With FolderBrowserDialog1
            ' Desktop is the root folder in the dialog.
            .RootFolder = Environment.SpecialFolder.Desktop
            ' Select the C:\Windows DirectoryW on entry.
            .SelectedPath = "c:\windows"
            ' Prompt the user with a custom message.
            .Description = "Select the source DirectoryW"
            If .ShowDialog = DialogResult.OK Then
                TextBox13.Text = .SelectedPath & "\"
                ' Display the selected folder if the user clicked on the OK button.
                'MessageBox.Show(.SelectedPath)
            End If
        End With
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        Dim FolderBrowserDialog1 As New FolderBrowserDialog

        ' Then use the following code to create the Dialog window
        ' Change the .SelectedPath property to the default location
        With FolderBrowserDialog1
            ' Desktop is the root folder in the dialog.
            .RootFolder = Environment.SpecialFolder.Desktop
            ' Select the C:\Windows DirectoryW on entry.
            .SelectedPath = "c:\windows"
            ' Prompt the user with a custom message.
            .Description = "Select the source DirectoryW"
            If .ShowDialog = DialogResult.OK Then
                TextBox11.Text = .SelectedPath & "\"
                ' Display the selected folder if the user clicked on the OK button.
                'MessageBox.Show(.SelectedPath)
            End If
        End With
    End Sub
    'Refresh tabcontrol while the splittes is moving

    Private Sub SplitContainer1_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitContainer1.SplitterMoved
        TabControl1.Refresh()
    End Sub

    Private Sub SplitContainer1_SplitterMoving(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterCancelEventArgs) Handles SplitContainer1.SplitterMoving
        ' Check to make sure the splitter won't be updated by the
    
        TabControl1.Refresh()
        'End If

    End Sub
    Private Sub SplitContainer2_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitContainer2.SplitterMoved
        TabControl1.Refresh()
    End Sub

    Private Sub SplitContainer2_SplitterMoving(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterCancelEventArgs) Handles SplitContainer2.SplitterMoving
        ' Check to make sure the splitter won't be updated by the
        ' normal move behavior also
        ' If DirectCast(sender, SplitContainer).IsSplitterFixed Then
        TabControl1.Refresh()
        'End If

    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim loginIN As TestSimpleObject '(CDbl(TextBox2.Text), CDbl(TextBox3.Text), CDbl(TextBox4.Text), CDbl(TextBox5.Text), CDbl(TextBox6.Text), CDbl(TextBox6.Text), CDbl(TextBox8.Text), CDbl(TextBox9.Text), CDbl(TextBox10.Text))
        If (File.Exists(LOGIN_DATA_FILE)) Then
            Dim fs As FileStream = New FileStream(LOGIN_DATA_FILE, FileMode.Open)
            Dim formatter As Runtime.Serialization.Formatters.Binary.BinaryFormatter = New Runtime.Serialization.Formatters.Binary.BinaryFormatter()


            Try
                ' loginIn = formatter.Deserialize(fs)
                'loginIn = CType(formatter.Deserialize(fs), TestSimpleObject)
                loginIN = DirectCast(formatter.Deserialize(fs), TestSimpleObject)
            Catch
                ' do nothing
            Finally
                fs.Close()
            End Try
            If (Not loginIN Is Nothing) Then
                Me.TextBox7.Text = loginIN.member1.ToString
                Me.TextBox1.Text = loginIN.member2.ToString
                Me.TextBox2.Text = loginIN.member3.ToString
                Me.TextBox3.Text = loginIN.member4.ToString
                Me.TextBox4.Text = loginIN.member5.ToString
                Me.TextBox5.Text = loginIN.member6.ToString
                Me.TextBox11.Text = loginIN.member7.ToString
                Me.TextBox6.Text = loginIN.member8.ToString
                Me.TextBox10.Text = loginIN.member9.ToString
                Me.TextBox13.Text = loginIN.member10.ToString
                Me.CheckBox2.Checked = loginIN.member11
            End If
        End If
    End Sub
End Class

