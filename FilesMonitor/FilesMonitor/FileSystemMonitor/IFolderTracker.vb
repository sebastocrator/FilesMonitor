Interface IFolderTracker
	Event ChangedEvent As FileSystemEvent
	Event CreatedEvent As FileSystemEvent
	Event DeletedEvent As FileSystemEvent
	Event RenamedEvent As FileSystemRenameEvent
	Sub Start()
	Sub [Stop]()
End Interface