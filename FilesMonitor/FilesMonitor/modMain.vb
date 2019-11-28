Public Module modMain

	Dim Mailer As New SMTPMailer
	Dim IsMailActief As Boolean = False

	Public Sub Main()
		Try

			Dim Settingsfile As String = String.Empty

			Dim root As String = String.Empty
			Dim InclusiefSubfolders As Boolean = True
			Dim IgnoreFileProperties As New FileProperties


			Dim args = My.Application.CommandLineArgs
			If args.Count = 1 Then
				If args(0).ToLower.StartsWith("/settings=") Then
					Settingsfile = args(0).Substring(Len("/settings="))
					If My.Computer.FileSystem.FileExists(Settingsfile) Then
						Dim rawsettings As String = My.Computer.FileSystem.ReadAllText(Settingsfile)
						Dim message As New System.Net.Mail.MailMessage()
						Dim settings As String() = rawsettings.Split(GetChar(vbCrLf, 1))
						For Each item As String In settings
							item = item.Replace(vbLf, "")
							Dim optie As String() = item.Split("="c)
							If optie.Length = 2 Then
								Select Case optie(0).Trim.ToLower
									Case "smtphost"
										Mailer.SMTPHost = optie(1).Trim
									Case "smtpport"
										Mailer.SMTPPort = CInt(optie(1).Trim)
									Case "smtpuser"
										Mailer.SMTPUser = optie(1).Trim
									Case "smtppwd"
										Mailer.SMTPPwd = optie(1).Trim
									Case "mailfrom"
										Mailer.MsgFrom = optie(1).Trim
									Case "mailto"
										Mailer.MsgTo = optie(1).Trim
									Case "smtpssl"
										Mailer.SMTPSsl = CBool(optie(1).Trim)
									Case "includesubfolders"
										InclusiefSubfolders = (optie(1).Trim <> "0" And Not optie(1).Trim.StartsWith("n"))
									Case "folder"
										root = optie(1).Trim
									Case "ignorefileextension"
										IgnoreFileProperties.Extensions.Add(optie(1).Trim.ToLower)
									Case "ignorefilenamestart"
										IgnoreFileProperties.NameStartsWith.Add(optie(1).Trim.ToLower)
									Case "ignorefilenameend"
										IgnoreFileProperties.NameEndsWith.Add(optie(1).Trim.ToLower)
									Case "ignorefilenamecontains"
										IgnoreFileProperties.NameContains.Add(optie(1).Trim.ToLower)
								End Select
							End If
						Next
						If Mailer.IsBruikbaar Then
							Console.WriteLine("Wijzigingen worden per mail gemeld.")
							Mailer.Send("MonitorFiles geactiveerd", "MonitorFiles is geactiveerd en wijzigingen worden per mail gemeld.")
							IsMailActief = True
						Else
							Console.WriteLine("Geen bruikbare mailinstellingen gevonden, wijzigingen worden niet per mail gemeld.")
						End If
					End If
				End If

				FileMonitor = New FolderTracker(root, InclusiefSubfolders) With {
						.IgnoreFileProperties = IgnoreFileProperties
					}
				FileMonitor.Start()
				Console.Error.WriteLine("MonitorFiles is geactiveerd. Druk op 'q' om de applicatie af te sluiten.")

				While Console.ReadKey(False).KeyChar <> "q"c
				End While

				FileMonitor.Stop()
			Else
				WriteSyntax()
			End If
		Catch ex As Exception
			Console.Error.WriteLine($"Fout: {ex.Message}")
			If ex.InnerException IsNot Nothing Then
				Console.Error.WriteLine($"{ex.InnerException.Message}")
			End If
			WriteSyntax()
		End Try

	End Sub
	Private Sub WriteSyntax()
		Console.Error.WriteLine("---------------------------------------------------------------------------------")
		Console.Error.WriteLine($"Aanroep: {My.Application.Info.AssemblyName} /Settings=<Settingsfile>")
		Console.Error.WriteLine("---------------------------------------------------------------------------------")
		Console.Error.WriteLine("Settingsfile dient de volgende waarden te bevatten:")
		Console.Error.WriteLine("")
		Console.Error.WriteLine("Folder=(de te monitoren folder)")
		Console.Error.WriteLine("IncludeSubfolders=(0 of n: geen subfolders monitoren)")
		Console.Error.WriteLine("# Vanaf hier zijn de settings optioneel:")
		Console.Error.WriteLine("SMTPHost=(SMTP host)")
		Console.Error.WriteLine("SMTPPort=(SMTP poort)")
		Console.Error.WriteLine("SMTPUser=(SMTP accountnaam)")
		Console.Error.WriteLine("SMTPPwd=(wachtwoord SMTP account)")
		Console.Error.WriteLine("MailFrom=(mailadres afzender)")
		Console.Error.WriteLine("MailTo=(mailadres bestemming)")
		Console.Error.WriteLine("# De volgende settings mogen meermaals voorkomen:")
		Console.Error.WriteLine("IgnoreFileExtension=(extensie)")
		Console.Error.WriteLine("IgnoreFileNameStart=(begin van bestandsnaam)")
		Console.Error.WriteLine("IgnoreFileNameEnd=(einde van bestandsnaam)")
		Console.Error.WriteLine("IgnoreFileNameContains=(willekeurig deel van bestandsnaam)")
		Console.Error.WriteLine("---------------------------------------------------------------------------------")
	End Sub

	Public WithEvents FileMonitor As FolderTracker

	Private Sub FileMonitor_ChangedEvent(fullPath As String) Handles FileMonitor.ChangedEvent
		Console.WriteLine($"Gewijzigd  {Now.ToString("yyyy-dd-MM HH:mm:ss")} : {fullPath}")
		If IsMailActief Then
			Try
				Mailer.Send($"Gewijzigd:  {fullPath}", $"Gewijzigd:  {fullPath}")
			Catch ex As Exception
				Console.Error.WriteLine($"Fout: {ex.Message}")
				If ex.InnerException IsNot Nothing Then
					Console.Error.WriteLine($"{ex.InnerException.Message}")
				End If
			End Try
		End If
	End Sub

	Private Sub FileMonitor_CreatedEvent(fullPath As String) Handles FileMonitor.CreatedEvent
		Console.WriteLine($"Nieuw      {Now.ToString("yyyy-dd-MM HH:mm:ss")} : {fullPath}")
		If IsMailActief Then
			Try
				Mailer.Send($"Nieuw: {fullPath}", $"Nieuw: {fullPath}")
			Catch ex As Exception
				Console.Error.WriteLine($"Fout: {ex.Message}")
				If ex.InnerException IsNot Nothing Then
					Console.Error.WriteLine($"{ex.InnerException.Message}")
				End If
			End Try
		End If
	End Sub

	Private Sub FileMonitor_DeletedEvent(fullPath As String) Handles FileMonitor.DeletedEvent
		Console.WriteLine($"Verwijderd {Now.ToString("yyyy-dd-MM HH:mm:ss")} : {fullPath}")
		If IsMailActief Then
			Try
				Mailer.Send($"Verwijderd: {fullPath}", $"Verwijderd: {fullPath}")
			Catch ex As Exception
				Console.Error.WriteLine($"Fout: {ex.Message}")
				If ex.InnerException IsNot Nothing Then
					Console.Error.WriteLine($"{ex.InnerException.Message}")
				End If
			End Try
		End If
	End Sub

	Private Sub FileMonitor_RenamedEvent(oldFullPath As String, oldName As String, fullPath As String, name As String) Handles FileMonitor.RenamedEvent
		Console.WriteLine($"Hernoemd   {Now.ToString("yyyy-dd-MM HH:mm:ss")} : {oldFullPath} naar {fullPath}")
		If IsMailActief Then
			Try
				Mailer.Send($"Hernoemd: {oldFullPath} naar {fullPath}", $"Hernoemd: {oldFullPath} naar {fullPath}")
			Catch ex As Exception
				Console.Error.WriteLine($"Fout: {ex.Message}")
				If ex.InnerException IsNot Nothing Then
					Console.Error.WriteLine($"{ex.InnerException.Message}")
				End If
			End Try
		End If
	End Sub

End Module
