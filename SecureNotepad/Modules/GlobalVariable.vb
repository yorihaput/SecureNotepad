Module GlobalVariable
    Public AppName As String = "Secure Notepad"
    Public Version As String = "v.1.0"
    Public filePassword As String = ""
    Public frmPasswordCancel As Boolean = True

    <Flags()> _
    Public Enum HChangeNotifyFlags
        ' <summary>
        ' The <i>dwItem1</i> and <i>dwItem2</i> parameters are DWORD values.
        ' </summary>
        SHCNF_DWORD = &H3
        ' <summary>
        ' <i>dwItem1</i> and <i>dwItem2</i> are the addresses of ITEMIDLIST structures that
        ' represent the item(s) affected by the change.
        ' Each ITEMIDLIST must be relative to the desktop folder.
        ' </summary>
        SHCNF_IDLIST = &H0
        ' <summary>
        ' <i>dwItem1</i> and <i>dwItem2</i> are the addresses of null-terminated strings of
        ' maximum length MAX_PATH that contain the full path names
        ' of the items affected by the change.
        ' </summary>
        SHCNF_PATHA = &H1
        ' <summary>
        ' <i>dwItem1</i> and <i>dwItem2</i> are the addresses of null-terminated strings of
        ' maximum length MAX_PATH that contain the full path names
        ' of the items affected by the change.
        ' </summary>
        SHCNF_PATHW = &H5
        ' <summary>
        ' <i>dwItem1</i> and <i>dwItem2</i> are the addresses of null-terminated strings that
        ' represent the friendly names of the printer(s) affected by the change.
        ' </summary>
        SHCNF_PRINTERA = &H2
        ' <summary>
        ' <i>dwItem1</i> and <i>dwItem2</i> are the addresses of null-terminated strings that
        ' represent the friendly names of the printer(s) affected by the change.
        ' </summary>
        SHCNF_PRINTERW = &H6
        ' <summary>
        ' The function should not return until the notification
        ' has been delivered to all affected components.
        ' As this flag modifies other data-type flags it cannot by used by itself.
        ' </summary>
        SHCNF_FLUSH = &H1000
        ' <summary>
        ' The function should begin delivering notifications to all affected components
        ' but should return as soon as the notification process has begun.
        ' As this flag modifies other data-type flags it cannot by used by itself.
        ' </summary>
        SHCNF_FLUSHNOWAIT = &H2000
    End Enum
    <Flags()> _
    Public Enum HChangeNotifyEventID
        SHCNE_ALLEVENTS = &H7FFFFFFF

        ' <summary>
        ' A file type association has changed. <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/>
        ' must be specified in the <i>uFlags</i> parameter.
        ' <i>dwItem1</i> and <i>dwItem2</i> are not used and must be <see langword="null"/>.
        ' </summary>
        SHCNE_ASSOCCHANGED = &H8000000

        ' <summary>
        ' The attributes of an item or folder have changed.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the item or folder that has changed.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_ATTRIBUTES = &H800

        ' <summary>
        ' A nonfolder item has been created.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the item that was created.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_CREATE = &H2

        ' <summary>
        ' A nonfolder item has been deleted.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the item that was deleted.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_DELETE = &H4

        ' <summary>
        ' A drive has been added.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the root of the drive that was added.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_DRIVEADD = &H100

        ' <summary>
        ' A drive has been added and the Shell should create a new window for the drive.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the root of the drive that was added.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_DRIVEADDGUI = &H10000

        ' <summary>
        ' A drive has been removed. <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the root of the drive that was removed.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_DRIVEREMOVED = &H80

        ' <summary>
        ' Not currently used.
        ' </summary>
        ' SHCNE_EXTENDED_EVENT = &H4000000

        ' <summary>
        ' The amount of free space on a drive has changed.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the root of the drive on which the free space changed.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_FREESPACE = &H40000

        ' <summary>
        ' Storage media has been inserted into a drive.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the root of the drive that contains the new media.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_MEDIAINSERTED = &H20

        ' <summary>
        ' Storage media has been removed from a drive.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the root of the drive from which the media was removed.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_MEDIAREMOVED = &H40

        ' <summary>
        ' A folder has been created. <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/>
        ' or <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the folder that was created.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_MKDIR = &H8

        ' <summary>
        ' A folder on the local computer is being shared via the network.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the folder that is being shared.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_NETSHARE = &H200

        ' <summary>
        ' A folder on the local computer is no longer being shared via the network.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the folder that is no longer being shared.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_NETUNSHARE = &H400

        ' <summary>
        ' The name of a folder has changed.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the previous pointer to an item identifier list (PIDL) or name of the folder.
        ' <i>dwItem2</i> contains the new PIDL or name of the folder.
        ' </summary>
        SHCNE_RENAMEFOLDER = &H20000

        ' <summary>
        ' The name of a nonfolder item has changed.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the previous PIDL or name of the item.
        ' <i>dwItem2</i> contains the new PIDL or name of the item.
        ' </summary>
        SHCNE_RENAMEITEM = &H1

        ' <summary>
        ' A folder has been removed.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the folder that was removed.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_RMDIR = &H10

        ' <summary>
        ' The computer has disconnected from a server.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the server from which the computer was disconnected.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' </summary>
        SHCNE_SERVERDISCONNECT = &H4000

        ' <summary>
        ' The contents of an existing folder have changed
        ' but the folder still exists and has not been renamed.
        ' <see cref="HChangeNotifyFlags.SHCNF_IDLIST"/> or
        ' <see cref="HChangeNotifyFlags.SHCNF_PATH"/> must be specified in <i>uFlags</i>.
        ' <i>dwItem1</i> contains the folder that has changed.
        ' <i>dwItem2</i> is not used and should be <see langword="null"/>.
        ' If a folder has been created deleted or renamed use SHCNE_MKDIR SHCNE_RMDIR or
        ' SHCNE_RENAMEFOLDER respectively instead.
        ' </summary>
        SHCNE_UPDATEDIR = &H1000

        ' <summary>
        ' An image in the system image list has changed.
        ' <see cref="HChangeNotifyFlags.SHCNF_DWORD"/> must be specified in <i>uFlags</i>.
        ' </summary>
        SHCNE_UPDATEIMAGE = &H8000
    End Enum
End Module
