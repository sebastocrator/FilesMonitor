Public Class SMTPMailer
	Dim SmtpClient As New Net.Mail.SmtpClient With {.EnableSsl = True}

	Public Property SMTPHost As String = String.Empty
	Public Property SMTPPort As Integer = 25
	Public Property SMTPUser As String = String.Empty
	Public Property SMTPPwd As String = String.Empty
	Public Property SMTPSsl As Boolean = False
	Public Property MsgFrom As String
	Public Property MsgTo As String

	Public Sub New()

	End Sub
	Public Function IsBruikbaar() As Boolean
		Return (SMTPHost <> String.Empty AndAlso SMTPPort > 0 AndAlso MsgFrom <> String.Empty AndAlso MsgTo <> String.Empty)
	End Function
	Public Sub Send(subject As String, body As String)
		SmtpClient.Host = SMTPHost
		SmtpClient.Port = SMTPPort
		If SMTPUser <> String.Empty Then
			SmtpClient.Credentials = New Net.NetworkCredential(SMTPUser, SMTPPwd)
		End If
		SmtpClient.EnableSsl = SMTPSsl
		SmtpClient.Send(MsgFrom, MsgTo, subject, body)
	End Sub

End Class
