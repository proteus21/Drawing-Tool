Imports System
'Imports System.core
Imports System.Data
Imports System.Data.dataset
Imports System.Deployment
Imports System.Drawing
Imports System.Runtime.Serialization
'Imports System.servicemodel
Imports System.Windows.Forms
Imports System.Xml
Imports System.Xml.Linq
Imports AllFoldersandFilesTree.com.microsofttranslator.api.TranslateCompletedEventArgs
Imports AllFoldersandFilesTree.com.microsofttranslator.api.SoapService



Public Class Dictionary


    Dim ISO_language As String
    Dim ISO_language2 As String
    ' active key 7B42C67D99961F38D5EBB54955042F860DA20AD5
    Private Sub btnTranslate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTranslate.Click
        chooseLanguage()
        Dim strTranslatedText As String
        Try
            Dim client As New AllFoldersandFilesTree.com.microsofttranslator.api.SoapService
            client = New AllFoldersandFilesTree.com.microsofttranslator.api.SoapService()
            'Dim client As New TranslatorService.LanguageServiceClient
            'client = New TranslatorService.LanguageServiceClient()
            strTranslatedText = client.Translate("6CE9C85A41571C050C379F60DA173D286384E0F2", txtTraslateFrom.Text, ISO_language, ISO_language2)
            txtTranslatedText.Text = strTranslatedText
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


    Sub chooseLanguage()
        '=============================================
        'AFRIKAANS' : 'af',
        'ALBANIAN' : 'sq',
        'AMHARIC' : 'am',
        'ARABIC' : 'ar',
        'ARMENIAN' : 'hy',
        'AZERBAIJANI' : 'az',
        'BASQUE' : 'eu',
        'BELARUSIAN' : 'be',
        'BENGALI' : 'bn',
        'BIHARI' : 'bh',
        'BRETON' : 'br',
        'BULGARIAN' : 'bg',
        'BURMESE' : 'my',
        'CATALAN' : 'ca',
        'CHEROKEE' : 'chr',
        'CHINESE' : 'zh',
        'CHINESE_SIMPLIFIED' : 'zh-CN',
        'CHINESE_TRADITIONAL' : 'zh-TW',
        'CORSICAN' : 'co',
        'CROATIAN' : 'hr',
        'CZECH' : 'cs',
        'DANISH' : 'da',
        'DHIVEHI' : 'dv',
        'DUTCH': 'nl',  
        'ENGLISH' : 'en',
        'ESPERANTO' : 'eo',
        'ESTONIAN' : 'et',
        'FAROESE' : 'fo',
        'FILIPINO' : 'tl',
        'FINNISH' : 'fi',
        'FRENCH' : 'fr',
        'FRISIAN' : 'fy',
        'GALICIAN' : 'gl',
        'GEORGIAN' : 'ka',
        'GERMAN' : 'de',
        'GREEK' : 'el',
        'GUJARATI' : 'gu',
        'HAITIAN_CREOLE' : 'ht',
        'HEBREW' : 'iw',
        'HINDI' : 'hi',
        'HUNGARIAN' : 'hu',
        'ICELANDIC' : 'is',
        'INDONESIAN' : 'id',
        'INUKTITUT' : 'iu',
        'IRISH' : 'ga',
        'ITALIAN' : 'it',
        'JAPANESE' : 'ja',
        'JAVANESE' : 'jw',
        'KANNADA' : 'kn',
        'KAZAKH' : 'kk',
        'KHMER' : 'km',
        'KOREAN' : 'ko',
        'KURDISH': 'ku',
        'KYRGYZ': 'ky',
        'LAO' : 'lo',
        'LATIN' : 'la',
        'LATVIAN' : 'lv',
        'LITHUANIAN' : 'lt',
        'LUXEMBOURGISH' : 'lb',
        'MACEDONIAN' : 'mk',
        'MALAY' : 'ms',
        'MALAYALAM' : 'ml',
        'MALTESE' : 'mt',
        'MAORI' : 'mi',
        'MARATHI' : 'mr',
        'MONGOLIAN' : 'mn',
        'NEPALI' : 'ne',
        'NORWEGIAN' : 'no',
        'OCCITAN' : 'oc',
        'ORIYA' : 'or',
        'PASHTO' : 'ps',
        'PERSIAN' : 'fa',
        'POLISH' : 'pl',
        'PORTUGUESE' : 'pt',
        'PORTUGUESE_PORTUGAL' : 'pt-PT',
        'PUNJABI' : 'pa',
        'QUECHUA' : 'qu',
        'ROMANIAN' : 'ro',
        'RUSSIAN' : 'ru',
        'SANSKRIT' : 'sa',
        'SCOTS_GAELIC' : 'gd',
        'SERBIAN' : 'sr',
        'SINDHI' : 'sd',
        'SINHALESE' : 'si',
        'SLOVAK' : 'sk',
        'SLOVENIAN' : 'sl',
        'SPANISH' : 'es',
        'SUNDANESE' : 'su',
        'SWAHILI' : 'sw',
        'SWEDISH' : 'sv',
        'SYRIAC' : 'syr',
        'TAJIK' : 'tg',
        'TAMIL' : 'ta',
        'TATAR' : 'tt',
        'TELUGU' : 'te',
        'THAI' : 'th',
        'TIBETAN' : 'bo',
        'TONGA' : 'to',
        'TURKISH' : 'tr',
        'UKRAINIAN' : 'uk',
        'URDU' : 'ur',
        'UZBEK' : 'uz',
        'UIGHUR' : 'ug',
        'VIETNAMESE' : 'vi',
        'WELSH' : 'cy',
        'YIDDISH' : 'yi',
        'YORUBA' : 'yo',
        'UNKNOWN' : ''

        If ComboBox1.SelectedIndex = 0 Then
            ISO_language = "en"
        End If

        If ComboBox1.SelectedIndex = 1 Then
            ISO_language = "fr"
        End If


        If ComboBox1.SelectedIndex = 2 Then
            ISO_language = "zn"
        End If


        If ComboBox1.SelectedIndex = 3 Then
            ISO_language = "it"
        End If


        If ComboBox1.SelectedIndex = 4 Then
            ISO_language = "de"
        End If

        If ComboBox1.SelectedIndex = 5 Then
            ISO_language = "es"
        End If


        If ComboBox1.SelectedIndex = 6 Then
            ISO_language = "ru"
        End If

        If ComboBox1.SelectedIndex = 7 Then
            ISO_language = "ja"
        End If
        If ComboBox1.SelectedIndex = 8 Then
            ISO_language = "ko"
        End If

        If ComboBox1.SelectedIndex = 9 Then
            ISO_language = "ar"
        End If
        If ComboBox1.SelectedIndex = 10 Then
            ISO_language = "sk"
        End If
        If ComboBox1.SelectedIndex = 11 Then
            ISO_language = "el"
        End If

        If ComboBox1.SelectedIndex = 12 Then
            ISO_language = "hu"
        End If
        If ComboBox1.SelectedIndex = 13 Then
            ISO_language = "pl"
        End If
        If ComboBox1.SelectedIndex = 14 Then
            ISO_language = "sl"
        End If
        If ComboBox1.SelectedIndex = 15 Then
            ISO_language = "uk"
        End If


        If ComboBox2.SelectedIndex = 0 Then
            ISO_language2 = "en"
        End If

        If ComboBox2.SelectedIndex = 1 Then
            ISO_language2 = "fr"
        End If


        If ComboBox2.SelectedIndex = 2 Then
            ISO_language2 = "zn"
        End If


        If ComboBox2.SelectedIndex = 3 Then
            ISO_language2 = "it"
        End If


        If ComboBox2.SelectedIndex = 4 Then
            ISO_language2 = "de"
        End If

        If ComboBox2.SelectedIndex = 5 Then
            ISO_language2 = "es"
        End If


        If ComboBox2.SelectedIndex = 6 Then
            ISO_language2 = "ru"
        End If

        If ComboBox2.SelectedIndex = 7 Then
            ISO_language2 = "ja"
        End If
        If ComboBox2.SelectedIndex = 8 Then
            ISO_language2 = "ko"
        End If

        If ComboBox2.SelectedIndex = 9 Then
            ISO_language2 = "ar"
        End If
        If ComboBox2.SelectedIndex = 10 Then
            ISO_language2 = "sk"
        End If
        If ComboBox2.SelectedIndex = 11 Then
            ISO_language2 = "el"
        End If

        If ComboBox2.SelectedIndex = 12 Then
            ISO_language2 = "hu"
        End If
        If ComboBox2.SelectedIndex = 13 Then
            ISO_language2 = "pl"
        End If
        If ComboBox2.SelectedIndex = 14 Then
            ISO_language = "sl"
        End If
        If ComboBox2.SelectedIndex = 15 Then
            ISO_language = "uk"
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ComboBox1.Items.Add("English")
        ComboBox1.Items.Add("French")
        ComboBox1.Items.Add("Chinese")
        ComboBox1.Items.Add("Italian")
        ComboBox1.Items.Add("German")
        ComboBox1.Items.Add("Spanish")
        ComboBox1.Items.Add("Russian")
        ComboBox1.Items.Add("Japanese")
        ComboBox1.Items.Add("Korean")
        ComboBox1.Items.Add("Arabic")
        ComboBox1.Items.Add("Slovak")
        ComboBox1.Items.Add("Greek")
        ComboBox1.Items.Add("Hungarian")
        ComboBox1.Items.Add("Polish")
        ComboBox1.Items.Add("Slovenian")
        ComboBox1.Items.Add("Ukrainian")

        ComboBox2.Items.Add("English")
        ComboBox2.Items.Add("French")
        ComboBox2.Items.Add("Chinese")
        ComboBox2.Items.Add("Italian")
        ComboBox2.Items.Add("German")
        ComboBox2.Items.Add("Spanish")
        ComboBox2.Items.Add("Russian")
        ComboBox2.Items.Add("Japanese")
        ComboBox2.Items.Add("Korean")
        ComboBox2.Items.Add("Arabic")
        ComboBox2.Items.Add("Slovak")
        ComboBox2.Items.Add("Greek")
        ComboBox2.Items.Add("Hungarian")
        ComboBox2.Items.Add("Polish")
        ComboBox2.Items.Add("Slovenian")
        ComboBox2.Items.Add("Ukrainian")

        ComboBox1.SelectedIndex = 13
        ComboBox2.SelectedIndex = 0
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ComboBox1.SelectedIndex = 13
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ComboBox1.SelectedIndex = 4
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ComboBox2.SelectedIndex = 13
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        ComboBox2.SelectedIndex = 0
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ComboBox2.SelectedIndex = 4
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim remComb1 As Integer = ComboBox1.SelectedIndex
        Dim remComb2 As Integer = ComboBox2.SelectedIndex

        ComboBox2.SelectedIndex = remComb1
        ComboBox1.SelectedIndex = remComb2
    End Sub
End Class


