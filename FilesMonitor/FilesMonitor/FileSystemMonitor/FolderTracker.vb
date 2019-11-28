Imports System.IO
Imports System.Threading

Public Class FolderTracker
	Implements IFolderTracker

	Enum FolderTrackerEventType
		ChangedEventType
		CreatedEventType
		DeletedEventType
		RenamedEventType
		MovedEventType
	End Enum

	Structure FolderTrackerKey
		Public Property OldFullPath As String
		Public Property OldName As String
		Public Property FullPath As String
		Public Property Name As String
		Public Property EventType As FolderTrackerEventType
	End Structure

	Structure FolderTrackerValue
		Public Property Timestamp As Date
		Public Property EventType As FolderTrackerEventType
	End Structure

	Public Event ChangedEvent As FileSystemEvent Implements IFolderTracker.ChangedEvent
	Public Event CreatedEvent As FileSystemEvent Implements IFolderTracker.CreatedEvent
	Public Event DeletedEvent As FileSystemEvent Implements IFolderTracker.DeletedEvent
	Public Event RenamedEvent As FileSystemRenameEvent Implements IFolderTracker.RenamedEvent


	Public Sub Start() Implements IFolderTracker.Start
		_fileSystemWatcher.EnableRaisingEvents = True
	End Sub

	Public Sub [Stop]() Implements IFolderTracker.Stop
		If _fileSystemWatcher IsNot Nothing Then
			_fileSystemWatcher.Dispose()
		End If

		_fileSystemWatcher.EnableRaisingEvents = False
	End Sub

	Public Sub New(ByVal observablePath As String, ByVal includeSubDirectories As Boolean)
		_existingPaths = Directory.EnumerateFileSystemEntries(observablePath, "*", SearchOption.AllDirectories).ToList()
		_fileSystemWatcher.Path = observablePath
		_fileSystemWatcher.IncludeSubdirectories = includeSubDirectories
		AddHandler _fileSystemWatcher.Created, New FileSystemEventHandler(AddressOf OnCreate)
		AddHandler _fileSystemWatcher.Changed, New FileSystemEventHandler(AddressOf OnChange)
		AddHandler _fileSystemWatcher.Deleted, New FileSystemEventHandler(AddressOf OnDelete)
		AddHandler _fileSystemWatcher.Renamed, New RenamedEventHandler(AddressOf OnRename)
		_fileSystemWatcher.NotifyFilter = NotifyFilters.Attributes Or NotifyFilters.CreationTime Or NotifyFilters.DirectoryName Or NotifyFilters.FileName Or NotifyFilters.LastAccess Or NotifyFilters.LastWrite Or NotifyFilters.Security Or NotifyFilters.Size
		_timer = New Timer(AddressOf OnTimeout, Nothing, Timeout.Infinite, Timeout.Infinite)
	End Sub

	Private ReadOnly _fileSystemWatcher As FileSystemWatcher = New FileSystemWatcher()
	Private ReadOnly _pendingEvents As Dictionary(Of FolderTrackerKey, FolderTrackerValue) = New Dictionary(Of FolderTrackerKey, FolderTrackerValue)()
	Private ReadOnly _timer As Timer
	Private _timerStarted As Boolean = False
	Private Shared _existingPaths As List(Of String) = New List(Of String)()


	Private Function IsFileForEvent(ByVal path As String) As Boolean
		Dim fi As FileInfo = New FileInfo(path)
		Return fi.Exists _
						AndAlso ((fi.Attributes And FileAttributes.Hidden) <> FileAttributes.Hidden) _
						AndAlso (IsFileNameForEvent(fi))
	End Function

	Public IgnoreFileProperties As New FileProperties

	Private Function IsFileNameForEvent(ByVal fi As FileInfo) As Boolean
		Return Not IgnoreFileProperties.HasSomeProperties(fi)
	End Function

	Private Function IsFileNameForEvent(ByVal path As String) As Boolean
		Return IsFileNameForEvent(New FileInfo(path))
	End Function

	Private Function FindReadyEvents(ByVal events As Dictionary(Of FolderTrackerKey, FolderTrackerValue)) As Dictionary(Of FolderTrackerKey, FolderTrackerEventType)
		Dim results As Dictionary(Of FolderTrackerKey, FolderTrackerEventType) = New Dictionary(Of FolderTrackerKey, FolderTrackerEventType)()
		Dim now As DateTime = DateTime.Now

		For Each e In events.GroupBy(Function(x) x.Key.FullPath)
			Dim hasDeletedEvent As Boolean = False
			Dim hasCreatedEvent As Boolean = False
			Dim hasChangedEvent As Boolean = False
			Dim hasRenamedEvent As Boolean = False

			For Each i In e

				Select Case i.Key.EventType
					Case FolderTrackerEventType.ChangedEventType
						hasChangedEvent = True
					Case FolderTrackerEventType.CreatedEventType
						hasCreatedEvent = True
					Case FolderTrackerEventType.DeletedEventType
						hasDeletedEvent = True
					Case FolderTrackerEventType.RenamedEventType
						hasRenamedEvent = True
				End Select
			Next

			If hasDeletedEvent AndAlso hasCreatedEvent Then
				results(e.First().Key) = FolderTrackerEventType.ChangedEventType

				SyncLock _pendingEvents

					For Each i In e
						_pendingEvents.Remove(i.Key)
					Next
				End SyncLock
			ElseIf hasCreatedEvent Then
				Dim entry = e.Where(Function(x) x.Key.EventType = FolderTrackerEventType.CreatedEventType).FirstOrDefault()
				Dim diff As Double = now.Subtract(entry.Value.Timestamp).TotalMilliseconds

				If diff >= 75 Then
					results(entry.Key) = FolderTrackerEventType.CreatedEventType
				End If
			ElseIf hasChangedEvent Then
				Dim entry = e.Where(Function(x) x.Key.EventType = FolderTrackerEventType.ChangedEventType).FirstOrDefault()
				Dim diff As Double = now.Subtract(entry.Value.Timestamp).TotalMilliseconds

				If diff >= 75 Then
					results(entry.Key) = FolderTrackerEventType.ChangedEventType
				End If
			ElseIf hasDeletedEvent Then
				Dim entry = e.Where(Function(x) x.Key.EventType = FolderTrackerEventType.DeletedEventType).FirstOrDefault()
				Dim diff As Double = now.Subtract(entry.Value.Timestamp).TotalMilliseconds

				If diff >= 75 Then
					results(entry.Key) = FolderTrackerEventType.DeletedEventType
				End If
			ElseIf hasRenamedEvent Then
				Dim entry = e.Where(Function(x) x.Key.EventType = FolderTrackerEventType.RenamedEventType).FirstOrDefault()
				Dim diff As Double = now.Subtract(entry.Value.Timestamp).TotalMilliseconds

				If diff >= 75 Then
					results(entry.Key) = FolderTrackerEventType.RenamedEventType
				End If
			End If
		Next

		Return results
	End Function



	Private Sub FireChangedEvent(ByVal key As FolderTrackerKey)
		If IsFileForEvent(key.FullPath) Then
			RaiseEvent ChangedEvent(key.FullPath)
		End If
	End Sub

	Private Sub FireDeletedEvent(ByVal key As FolderTrackerKey)
		_existingPaths.Remove(key.FullPath)
		RaiseEvent DeletedEvent(key.FullPath)
	End Sub

	Private Sub FireCreatedEvent(ByVal key As FolderTrackerKey)
		If IsFileForEvent(key.FullPath) Then
			If Not _existingPaths.Contains(key.FullPath) Then
				_existingPaths.Add(key.FullPath)
				RaiseEvent CreatedEvent(key.FullPath)
			End If
		End If
	End Sub

	Private Sub FireRenamedEvent(ByVal key As FolderTrackerKey)
		If Not _existingPaths.Contains(key.FullPath) AndAlso IsFileForEvent(key.FullPath) Then
			_existingPaths.Remove(key.OldFullPath)
			_existingPaths.Add(key.FullPath)
			If IsFileNameForEvent(key.OldFullPath) Then
				RaiseEvent RenamedEvent(key.OldFullPath, key.OldName, key.FullPath, key.Name)
			Else
				RaiseEvent CreatedEvent(key.FullPath)
			End If

		End If
	End Sub

	Private Sub OnTimeout(ByVal state As Object)
		Dim events As Dictionary(Of FolderTrackerKey, FolderTrackerEventType)

		SyncLock _pendingEvents
			events = FindReadyEvents(_pendingEvents)

			For Each e In events
				_pendingEvents.Remove(e.Key)
			Next

			If _pendingEvents.Count = 0 Then
				_timer.Change(Timeout.Infinite, Timeout.Infinite)
				_timerStarted = False
			End If
		End SyncLock

		Dim currentPaths = Directory.EnumerateFileSystemEntries(_fileSystemWatcher.Path, "*", SearchOption.AllDirectories).ToList()
		Dim newFiles = currentPaths.Except(_existingPaths).ToList()
		Dim deletedFiles = _existingPaths.Except(currentPaths).ToList()

		For Each e In events

			Select Case e.Value
				Case FolderTrackerEventType.ChangedEventType
					FireChangedEvent(e.Key)
				Case FolderTrackerEventType.CreatedEventType

					If newFiles.Contains(e.Key.FullPath) Then
						newFiles.Remove(e.Key.FullPath)
						FireCreatedEvent(e.Key)
					End If

				Case FolderTrackerEventType.DeletedEventType

					If deletedFiles.Contains(e.Key.FullPath) Then
						deletedFiles.Remove(e.Key.FullPath)
						FireDeletedEvent(e.Key)
					End If

				Case FolderTrackerEventType.RenamedEventType

					If deletedFiles.Contains(e.Key.OldFullPath) Then
						deletedFiles.Remove(e.Key.OldFullPath)
						FireRenamedEvent(e.Key)
					End If
			End Select
		Next

		_existingPaths = currentPaths
	End Sub

	Private Sub OnFolderTrackerEvent(ByVal key As FolderTrackerKey, ByVal value As FolderTrackerValue)
		SyncLock _pendingEvents
			_pendingEvents(key) = value

			If Not _timerStarted Then
				_timer.Change(100, 100)
				_timerStarted = True
			End If
		End SyncLock
	End Sub

	Private Sub OnChange(ByVal sender As Object, ByVal e As FileSystemEventArgs)
		If File.Exists(e.FullPath) OrElse Directory.Exists(e.FullPath) Then

			If Not _existingPaths.Contains(e.FullPath) Then
				Return
			End If

			Dim key As FolderTrackerKey = New FolderTrackerKey() With {
				.Name = New FileInfo(e.Name).Name,
				.FullPath = e.FullPath,
				.EventType = FolderTrackerEventType.ChangedEventType
			}
			Dim value As FolderTrackerValue = New FolderTrackerValue() With {
				.EventType = FolderTrackerEventType.ChangedEventType,
				.Timestamp = DateTime.Now
			}
			OnFolderTrackerEvent(key, value)
		End If
	End Sub

	Private Sub OnDelete(ByVal sender As Object, ByVal e As FileSystemEventArgs)
		If _existingPaths.Contains(e.FullPath) Then
			Dim key As FolderTrackerKey = New FolderTrackerKey() With {
				.Name = New FileInfo(e.Name).Name,
				.FullPath = e.FullPath,
				.EventType = FolderTrackerEventType.DeletedEventType
			}
			Dim value As FolderTrackerValue = New FolderTrackerValue() With {
				.EventType = FolderTrackerEventType.DeletedEventType,
				.Timestamp = DateTime.Now
			}
			OnFolderTrackerEvent(key, value)
		End If
	End Sub

	Private Sub OnCreate(ByVal sender As Object, ByVal e As FileSystemEventArgs)
		If File.Exists(e.FullPath) OrElse Directory.Exists(e.FullPath) Then
			Dim key As FolderTrackerKey = New FolderTrackerKey() With {
				.Name = New FileInfo(e.Name).Name,
				.FullPath = e.FullPath,
				.EventType = FolderTrackerEventType.CreatedEventType
			}
			Dim value As FolderTrackerValue = New FolderTrackerValue() With {
				.EventType = FolderTrackerEventType.CreatedEventType,
				.Timestamp = DateTime.Now
			}
			OnFolderTrackerEvent(key, value)
		End If
	End Sub

	Private Sub OnRename(ByVal sender As Object, ByVal e As RenamedEventArgs)
		Dim key As FolderTrackerKey = New FolderTrackerKey() With {
			.OldFullPath = e.OldFullPath,
			.OldName = New FileInfo(e.OldName).Name,
			.FullPath = e.FullPath,
			.Name = New FileInfo(e.Name).Name,
			.EventType = FolderTrackerEventType.RenamedEventType
		}
		Dim value As FolderTrackerValue = New FolderTrackerValue() With {
			.EventType = FolderTrackerEventType.RenamedEventType,
			.Timestamp = DateTime.Now
		}
		OnFolderTrackerEvent(key, value)
	End Sub
End Class

