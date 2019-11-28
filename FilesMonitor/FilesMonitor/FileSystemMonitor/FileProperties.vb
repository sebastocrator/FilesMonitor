Imports System.IO

Public Class FileProperties
	Public Extensions As New List(Of String)
	Public NameStartsWith As New List(Of String)
	Public NameEndsWith As New List(Of String)
	Public NameContains As New List(Of String)

	Public Function HasSomeProperties(ByVal fi As FileInfo) As Boolean
		If Me.Extensions.Contains(fi.Extension) Then
			Return True
		Else
			For Each txt As String In NameStartsWith
				If fi.Name.ToLower.StartsWith(txt.ToLower) Then
					Return True
				End If
			Next
			For Each txt As String In NameEndsWith
				If fi.Name.ToLower.EndsWith(txt.ToLower) Then
					Return True
				End If
			Next
			For Each txt As String In NameContains
				If fi.Name.ToLower.Contains(txt.ToLower) Then
					Return True
				End If
			Next
		End If
		Return False
	End Function

End Class