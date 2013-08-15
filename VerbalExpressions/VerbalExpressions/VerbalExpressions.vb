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
        If value = vbNull Then
            Throw New Exception("value cannot be null")
        End If
        Return Regex.Escape(value)
    End Function

    Public Function Add(ByVal value As String, Optional ByVal sanitize As Boolean = False) As VerbalExpression
        If value = vbNull Then
            Throw New Exception("value cannot be null")
        End If
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
#End Region
End Class
