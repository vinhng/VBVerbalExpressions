Imports System.Text.RegularExpressions

Public Class VerbalExpression
    Private _prefixes As String = ""
    Private _source As String = ""
    Private _suffixes As String = ""
    Private _modifiers As RegexOptions = RegexOptions.Multiline

    Public Sub New()

    End Sub

    Public ReadOnly Property RegexString As String
        Get
            Return _prefixes & _source & _suffixes
        End Get
    End Property

    Public ReadOnly Property RegexPattern As Regex
        Get
            Return New Regex(RegexString, _modifiers)
        End Get
    End Property

#Region "helpers"
    Public Function Sanitize(ByVal value As String) As String
        CheckInput(value)
        Return Regex.Escape(value)
    End Function

    Public Function Add(ByVal value As String, Optional ByVal sanitize As Boolean = False) As VerbalExpression
        CheckInput(value)
        If sanitize Then
            value = Me.Sanitize(value)
        End If
        _source &= value
        Return Me
    End Function

    Public Function StartOfLine(Optional ByVal enabled As Boolean = True) As VerbalExpression
        If enabled Then
            _prefixes &= "^"
        End If
        Return Me
    End Function

    Public Function EndOfLine(Optional ByVal enabled As Boolean = True) As VerbalExpression
        If enabled Then
            _suffixes &= "$"
        End If
        Return Me
    End Function

    Public Function Find(ByVal value As String) As VerbalExpression
        Return [Then](value)
    End Function

    Public Function [Then](ByVal value As String) As VerbalExpression
        CheckInput(value)
        Return Add(String.Format("({0})", value), True)
    End Function

    Public Function Maybe(ByVal value As String) As VerbalExpression
        CheckInput(value)
        Return Add(String.Format("({0})?", value), True)
    End Function

    Public Function Anything() As VerbalExpression
        Return Add("(.*)")
    End Function

    Public Function AnythingBut(ByVal value As String) As VerbalExpression
        CheckInput(value)
        Return Add(String.Format("([^{0}]*)", value), True)
    End Function

    Public Function Something() As VerbalExpression
        Return Add("(.+)")
    End Function

    Public Function SomethingBut(ByVal value As String) As VerbalExpression
        CheckInput(value)
        Return Add(String.Format("([^{0}]+)", value), True)
    End Function

    Public Function Replace(ByVal value As String) As VerbalExpression
        CheckInput(value)
        Dim temp = RegexPattern.ToString()
        If temp.Length > 0 Then
            _source = _source.Replace(temp, value)
        End If
        Return Me
    End Function

    Public Function LineBreak() As VerbalExpression
        Return Add("\n|(\r\n)")
    End Function

    Public Function Br() As VerbalExpression
        Return LineBreak()
    End Function

    Public Function Tab() As VerbalExpression
        Return Add("\t")
    End Function

    Public Function Word() As VerbalExpression
        Return Add("\w+")
    End Function

    Public Function AnyOf(ByVal value As String) As VerbalExpression
        CheckInput(value)
        Return Add("[" & value & "]", True)
    End Function

    Public Function Any(ByVal value As String) As VerbalExpression
        Return AnyOf(value)
    End Function

    Public Function Range(ByVal ParamArray params() As String) As VerbalExpression
        Dim value = ""
        For i As Integer = 0 To params.Length - 2 Step +2
            value &= String.Format("{0}-{1}", params(i), params(i + 1))
        Next
        Return Add("[" & value & "]")
    End Function

    Public Function Multiple(ByVal value As String, Optional ByVal sanitize As Boolean = True) As VerbalExpression
        CheckInput(value)
        Return Add(String.Format("({0})+", value), sanitize)
    End Function

    Sub CheckInput(ByVal value As String)
        If value Is Nothing Then Throw New Exception("value cannot be null")
    End Sub

#End Region
End Class
