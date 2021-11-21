Imports System.Runtime.Serialization.Formatters.Binary
Imports System.SerializableAttribute
Module Serializable

    <Serializable()> Public Class TestSimpleObject

        Public member1 As String
        Public member2 As String
        Public member3 As String
        Public member4 As String
        Public member5 As String
        Public member6 As String
        Public member7 As String
        Public member8 As String
        Public member9 As String
        Public member10 As String
        Public member11 As Boolean


        Public Sub New(ByVal _text2 As String, ByVal _text3 As String, ByVal _text4 As String, ByVal _text5 As String, ByVal _text6 As String, ByVal _text7 As String, ByVal _text8 As String, ByVal _text9 As String, ByVal _text10 As String, ByVal _text11 As String, ByVal _text12 As Boolean)
            member1 = _text2
            member2 = _text3
            member3 = _text4
            member4 = _text5
            member5 = _text6
            member6 = _text7
            member7 = _text8
            member8 = _text9
            member9 = _text10
            member10 = _text11
            member11 = _text12

        End Sub 'New


    End Class 'TestSimpleObject

End Module
