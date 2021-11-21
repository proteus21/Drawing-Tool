Imports iTextSharp.text.pdf
Module Module1


    Public Declare Auto Function SHGetFileInfo Lib "shell32.dll" (ByVal pszPath As String, ByVal dwFileAttributes As Integer, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Integer, ByVal uFlags As Integer) As IntPtr

    Public Const SHGFI_ICON As Integer = &H100
    Public Const SHGFI_SMALLICON As Integer = &H1
    Public Const SHGFI_LARGEICON As Integer = &H0
    Public Const SHGFI_OPENICON As Integer = &H2

    Structure SHFILEINFO
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As Integer
        <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=260)> _
        Public szDisplayName As String
        <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=80)> _
        Public szTypeName As String
    End Structure
    Function GetShellOpenIconAsImage(ByVal argPath As String) As Image
        Dim mShellFileInfo As New SHFILEINFO
        mShellFileInfo.szDisplayName = New String(Chr(0), 260)
        mShellFileInfo.szTypeName = New String(Chr(0), 80)
        SHGetFileInfo(argPath, 0, mShellFileInfo, System.Runtime.InteropServices.Marshal.SizeOf(mShellFileInfo), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_OPENICON)
        ' attempt to create a System.Drawing.Icon from the icon handle that was returned in the SHFILEINFO structure
        Dim mIcon As System.Drawing.Icon
        Dim mImage As System.Drawing.Image
        Try
            mIcon = System.Drawing.Icon.FromHandle(mShellFileInfo.hIcon)
            mImage = mIcon.ToBitmap
        Catch ex As Exception
            ' for some reason the icon could not be converted so create a blank System.Drawing.Image to return instead
            mImage = New System.Drawing.Bitmap(16, 16)
        End Try
        ' return the final System.Drawing.Image
        Return mImage
    End Function
    Function CacheShellIcon(ByVal argPath As String) As String
        Dim mKey As String = Nothing
        ' determine the icon key for the file/folder specified in argPath
        If IO.Directory.Exists(argPath) = True Then
            If argPath = IO.Directory.GetDirectoryRoot(argPath) Then
                mKey = "drive"
            Else
                mKey = "folder"
            End If
        ElseIf IO.File.Exists(argPath) = True Then
            mKey = IO.Path.GetExtension(argPath)
        End If
        ' check if an icon for this key has already been added to the collection

        If Form1.ImageList1.Images.ContainsKey(mKey) = False Then
            Form1.ImageList1.Images.Add(mKey, GetShellOpenIconAsImage(argPath))
            'Form1.ImageList1.Images.Add(mKey, GetShellOpenIconAsImage(argPath))
            If mKey = "folder" Then Form1.ImageList1.Images.Add(mKey & "-open", GetShellOpenIconAsImage(argPath))
        End If
        Return mKey
    End Function
    Public Function ParsePdfText(ByVal sourcePDF As String, _
    Optional ByVal fromPageNum As Integer = 0, _
    Optional ByVal toPageNum As Integer = 0) As String

        Dim sb As New System.Text.StringBuilder()
        Try
            Dim reader As New PdfReader(sourcePDF)
            Dim pageBytes() As Byte = Nothing
            Dim token As PRTokeniser = Nothing
            Dim tknType As Integer = -1
            Dim tknValue As String = String.Empty

            If fromPageNum = 0 Then
                fromPageNum = 1
            End If
            If toPageNum = 0 Then
                toPageNum = reader.NumberOfPages
            End If

            If fromPageNum > toPageNum Then
                Throw New ApplicationException("Parameter error: The value of fromPageNum can " & _
                "not be larger than the value of toPageNum")
            End If

            For i As Integer = fromPageNum To toPageNum Step 1
                pageBytes = reader.GetPageContent(i)
                If Not IsNothing(pageBytes) Then
                    token = New PRTokeniser(pageBytes)
                    While token.NextToken()
                        tknType = token.TokenType()
                        tknValue = token.StringValue

                        If tknType = PRTokeniser.TokType.STRING Then
                            sb.Append(token.StringValue)
                            'I need to add these additional tests to properly add whitespace to the output string
                        ElseIf tknType = 1 AndAlso tknValue = "-600" Then
                            sb.Append(" ")
                        ElseIf tknType = 10 AndAlso tknValue = "TJ" Then
                            sb.Append(" ")
                        End If
                    End While
                End If
            Next i
        Catch ex As Exception
            MessageBox.Show("Exception occured. " & ex.Message)
            Return String.Empty
        End Try
        Return sb.ToString()
    End Function
    ''' <summary>
    ''' Extract the text from pdf pages and return it as a string
    ''' </summary>
    ''' <param name="sourcePDF">Full path to the source pdf file</param>
    ''' <param name="fromPageNum">[Optional] the page number (inclusive) to start text extraction </param>
    ''' <param name="toPageNum">[Optional] the page number (inclusive) to stop text extraction</param>
    ''' <returns>A string containing the text extracted from the specified pages</returns>
    ''' <remarks>If fromPageNum is not specified, text extraction will start from page 1. If
    ''' toPageNum is not specified, text extraction will end at the last page of the source pdf file.</remarks>
    Public Function ParsePdfText1(ByVal sourcePDF As String, _
                                  Optional ByVal fromPageNum As Integer = 0, _
                                  Optional ByVal toPageNum As Integer = 0) As String

        Dim sb As New System.Text.StringBuilder()
        Try
            Dim reader As New iTextSharp.text.pdf.PdfReader(sourcePDF)
            Dim pageBytes() As Byte = Nothing
            Dim token As iTextSharp.text.pdf.PRTokeniser = Nothing
            Dim tknType As Integer = -1
            Dim tknValue As String = String.Empty

            If fromPageNum = 0 Then
                fromPageNum = 1
            End If
            If toPageNum = 0 Then
                toPageNum = reader.NumberOfPages
            End If

            If fromPageNum > toPageNum Then
                Throw New ApplicationException("Parameter error: The value of fromPageNum can " & _
                                           "not be larger than the value of toPageNum")
            End If

            For i As Integer = fromPageNum To toPageNum Step 1
                pageBytes = reader.GetPageContent(i)
                If Not IsNothing(pageBytes) Then
                    token = New iTextSharp.text.pdf.PRTokeniser(pageBytes)
                    While token.NextToken()
                        tknType = token.TokenType()
                        tknValue = token.StringValue
                        Select Case tknType
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.NUMBER      '1
                                Dim dValue As Double
                                If Double.TryParse(tknValue, dValue) Then
                                    If dValue < -8000 Then
                                        sb.Append(ControlChars.Tab)
                                    End If
                                End If
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.STRING      '2
                                sb.Append(token.StringValue)
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.NAME        '3
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.COMMENT     '4
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.START_ARRAY '5
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.END_ARRAY   '6
                                sb.Append(" ")
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.START_DIC   '7
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.END_DIC     '8
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.REF         '9
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.OTHER       '10
                                Select Case tknValue
                                    Case "TJ"
                                        sb.Append(" ")
                                    Case "ET", "TD", "Td", "Tm", "T*"
                                        sb.Append(System.Environment.NewLine)
                                End Select
                        End Select
                    End While
                End If
            Next i
            reader.Close()
        Catch ex As Exception
            MessageBox.Show("Exception occured. " & ex.Message)
            Return String.Empty
        End Try
        Return sb.ToString()
    End Function
    Public Function ParseAllPdfText(ByVal sourcePDF As String) As Dictionary(Of Integer, String)
        Dim pdfText As New Dictionary(Of Integer, String)
        Dim sb As New System.Text.StringBuilder()
        Try
            Dim reader As New iTextSharp.text.pdf.PdfReader(sourcePDF)
            Dim pageBytes() As Byte = Nothing
            Dim token As iTextSharp.text.pdf.PRTokeniser = Nothing
            Dim tknType As Integer = -1
            Dim tknValue As String = String.Empty

            For i As Integer = 1 To reader.NumberOfPages Step 1
                pageBytes = reader.GetPageContent(i)
                If Not IsNothing(pageBytes) Then
                    sb.Length = 0
                    token = New iTextSharp.text.pdf.PRTokeniser(pageBytes)
                    While token.NextToken()
                        tknType = token.TokenType()
                        tknValue = token.StringValue
                        Select Case tknType
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.NUMBER      '1
                                Dim dValue As Double
                                If Double.TryParse(tknValue, dValue) Then
                                    If dValue < -8000 Then
                                        sb.Append(ControlChars.Tab)
                                    End If
                                End If
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.STRING      '2
                                sb.Append(token.StringValue)
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.NAME        '3
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.COMMENT     '4
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.START_ARRAY '5
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.END_ARRAY   '6
                                sb.Append(" ")
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.START_DIC   '7
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.END_DIC     '8
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.REF         '9
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.OTHER       '10
                                Select Case tknValue
                                    Case "TJ"
                                        sb.Append(" ")
                                    Case "ET", "TD", "Td", "Tm", "T*"
                                        sb.Append(System.Environment.NewLine)
                                End Select
                        End Select
                    End While
                    pdfText.Add(i, sb.ToString)
                End If
            Next i
            reader.Close()
        Catch ex As Exception
            MessageBox.Show("Exception occured. " & ex.Message)
        End Try
        Return pdfText
    End Function

    '    Imports System.Collections.Generic
    'Imports System.ComponentModel
    'Imports System.Data
    'Imports System.Drawing
    'Imports System.Linq
    'Imports System.Text
    'Imports System.Windows.Forms
    'Imports iTextSharp.text.pdf
    'Imports iTextSharp.text.pdf.parser

    'Namespace iTextSharpDeneme
    '        Partial Public Class Form1
    '            Inherits Form
    '            Public Sub New()
    '                InitializeComponent()
    '            End Sub

    '            Private Sub button2_Click(ByVal sender As Object, ByVal e As EventArgs)
    '                Dim fd As FileDialog = New OpenFileDialog()
    '                'Browse butonuna tiklandiginda FileDialog penceresi acilir
    '                fd.Filter = "(*.pdf)|*.pdf"
    '                've bu alanda okunmasi istenilen pdf dosyasi secilir
    '                fd.ShowDialog()
    '                pathText.Text = fd.FileName

    '            End Sub

    '            Private Sub readBtn_Click(ByVal sender As Object, ByVal e As EventArgs)
    '                icerik.Text = ""
    '                If pathText.Equals(String.Empty) Then
    '                    'Eger dosya secilmemisse uyari ver
    '                    MessageBox.Show("Lutfen bir dosya seciniz")
    '                Else
    '                    Try
    '                        Dim reader As New PdfReader(pathText.Text)
    '                        'Bir PdfReader nesnesi yaratilir
    '                        Dim n As Integer = reader.NumberOfPages
    '                        For i As Integer = 1 To n
    '                            '1. sayfadan son sayfaya kadar
    '                            Dim ites As ITextExtractionStrategy = New iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy()
    '                            Dim s As [String] = PdfTextExtractor.GetTextFromPage(reader, n, ites)
    '                            'bir sayfa icerisindeki yazilari alir
    '                            s = DirectCast(Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.[Default], Encoding.UTF8, Encoding.[Default].GetBytes(s))), String)
    '                            'Eger bir sayfadaki yazilarin tumu isteniyorsa s ile her satir
    '                            'uzerinde ayri ayri islem yapilmak isteniyorsa line ile
    '                            'islem yapilir
    '                            Dim line As String()
    '                            line = s.Split(ControlChars.Lf)
    '                            'Satir sonuna gore ayir
    '                            For Each item As String In line
    '                                'Her bir satir icin
    '                                'Her bir satir icin islemler

    '                                'Her bir line ayri ayri alinir ve RichTextBoxa eklenir
    '                                icerik.Text += item & ControlChars.Lf
    '                            Next


    '                            reader.Close()

    '                        Next
    '                    Catch ex As Exception
    '                        MessageBox.Show("Okuma islemi sirasinda bir hata olustu. " & ex.Message)
    '                    End Try
    '                End If
    '            End Sub
    '        End Class
    '    End Namespace
    Public Function GetSizeKB(ByVal filename As String) As Double
        ' There are 1024 bytes in a kilobyte
        Return New IO.FileInfo(filename).Length / 1024
    End Function
End Module
