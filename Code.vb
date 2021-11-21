Module Code
    '' Kod szybki tylko je¿eli wyswielta tekst bez ikon
    '' Zmiana poszukiwania ikon i extrat
    'Private Sub treeView1_NodeMouseClick(ByVal sender As Object, _
    '    ByVal e As TreeNodeMouseClickEventArgs) _
    '        Handles TreeView1.NodeMouseClick
    '    Dim FileExtension As String
    '    Dim FolderExtension As String
    '    Dim SubItemIndex As Integer
    '    Dim SubItemIndexs As Integer
    '    Dim DateMod As String
    '    Dim DateMods As String

    '    Dim newSelected As TreeNode = e.Node
    '    listView1.Items.Clear()
    '    Dim nodeDirInfo As DirectoryWInfo = New DirectoryWInfo(newSelected.Tag)
    '    Dim subItems() As ListViewItem.ListViewSubItem
    '    Dim item As ListViewItem = Nothing
    '    'ListTab(0) = CStr(TreeView1.SelectedNode.Tag)
    '    Dim n As Integer = 0
    '    Dim argPath As String
    '    Dim mkey As String
    '    Dim dir As DirectoryWInfo
    '    For Each nodeDirInfo In nodeDirInfo.GetDirectories()

    '        FolderExtension = IO.Path.GetExtension(nodeDirInfo.Name)
    '        DateMods = IO.DirectoryW.GetLastWriteTime(nodeDirInfo.Name)

    '        item = New ListViewItem(nodeDirInfo.Name, CacheShellIcon(nodeDirInfo.FullName))
    '        'ListView1.Items.Add(nodeDirInfo.Name.Substring(nodeDirInfo.Name.LastIndexOf("\"c) + 1), mkey)
    '        subItems = New ListViewItem.ListViewSubItem() _
    '            {New ListViewItem.ListViewSubItem(item, "Folder File"), _
    '            New ListViewItem.ListViewSubItem(item, _
    '            nodeDirInfo.LastAccessTime.ToShortDateString())}

    '        item.SubItems.AddRange(subItems)
    '        ListView1.Items.Add(item)
    '    Next
    '    'Dim folder As String = CStr(TreeView1.SelectedNode.Tag)
    '    Dim fileb As FileInfo
    '    For Each fileb In nodeDirInfo.GetFiles()
    '        FileExtension = IO.Path.GetExtension(fileb.FullName)
    '        subItems = New ListViewItem.ListViewSubItem() _
    '            {New ListViewItem.ListViewSubItem(item, FileExtension.ToString() & Chr(32) & "File"), _
    '            New ListViewItem.ListViewSubItem(item, _
    '            fileb.LastAccessTime.ToShortDateString())}
    '        'AddImages(fileb.FullName)
    '        item = New ListViewItem(fileb.Name, CacheShellIcon(fileb.FullName))
    '        item.SubItems.AddRange(subItems)
    '        ListView1.Items.Add(item)
    '    Next fileb


    'End Sub
    '==========================================================

    ' kod zródlowy nie jest naszybszy przy wyciaganiu i wyswietlaniu ikon
    '=================================================================================
    'Private Sub Treeview1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
    '    Dim FileExtension As String
    '    Dim FolderExtension As String
    '    Dim SubItemIndex As Integer
    '    Dim SubItemIndexs As Integer
    '    Dim DateMod As String
    '    Dim DateMods As String

    '    ListView1.Items.Clear()

    '    'If TreeView1.SelectedNode.Nodes.Count = 1 AndAlso TreeView1.SelectedNode.Nodes(0).Text = "Loading..." Then

    '    '    TreeView1.SelectedNode.Nodes.Clear()

    '    '    AddAllFolders(TreeView1.SelectedNode, CStr(TreeView1.SelectedNode.Tag))
    '    '    'TreeView1.SelectedImageKey=
    '    'End If
    '    'Dim folder As String = CStr(e.Node.Tag)
    '    Dim folder As String = CStr(TreeView1.SelectedNode.Tag)

    '    ListTab(0) = CStr(TreeView1.SelectedNode.Tag)
    '    Dim n As Integer = 0
    '    Dim argPath As String
    '    Dim mkey As String
    '    TextBox6.Text = folder ' dodanie info o scie¿ce do folderu
    '    'If Not folder Is Nothing AndAlso IO.DirectoryW.Exists(folder) Then

    '    '        Catch ex As Exception
    '    '        MsgBox(ex.Message)
    '    '    End Try
    '    '    Next
    '    'End If
    '    'AddImages(folder)
    '    If Not folder Is Nothing AndAlso IO.DirectoryW.Exists(folder) Then
    '        ' MsgBox("orety")

    '        Try
    '            For Each FolderNode As String In DirectoryW.GetDirectories(folder)

    '                FolderExtension = IO.Path.GetExtension(FolderNode)
    '                DateMods = IO.DirectoryW.GetLastWriteTime(FolderNode)
    '                ' ImageList1.Images.Clear()
    '                '  AddImages(FolderNode)

    '                If IO.DirectoryW.Exists(FolderNode) = True Then
    '                    If argPath = IO.DirectoryW.GetDirectoryWRoot(FolderNode) Then
    '                        mkey = "drive"
    '                    Else
    '                        mkey = "folder"
    '                    End If
    '                ElseIf IO.File.Exists(FolderNode) = True Then
    '                    mkey = IO.Path.GetExtension(FolderNode)
    '                End If
    '                If mkey = "folder" Then ImageList1.Images.Add(mkey & "-open", GetShellOpenIconAsImage(FolderNode))
    '                ListView1.SmallImageList = ImageList1

    '                ListView1.Items.Add(FolderNode.Substring(FolderNode.LastIndexOf("\"c) + 1), mkey)

    '                ' Create two ImageList objects.

    '                ' Add the ListView to the control collection.
    '                ' mDirectoryWNode.ImageKey = CacheShellIcon(mDirectoryW.FullName)
    '                '  ListView1.Items(SubItemIndexs).SubItems.Add(FolderExtension.ToString() & " Folder File")
    '                ListView1.Items(SubItemIndexs).SubItems.Add("Folder File")
    '                ListView1.Items(SubItemIndexs).SubItems.Add(DateMods.ToString())
    '                SubItemIndexs += 1
    '                n = n + 1
    '            Next

    '            For Each file As String In IO.DirectoryW.GetFiles(folder)
    '                FileExtension = IO.Path.GetExtension(file)
    '                DateMod = IO.File.GetLastWriteTime(file).ToString()
    '                AddImages(file)
    '                ListView1.Items.Add(file.Substring(file.LastIndexOf("\"c) + 1), file.ToString())
    '                ListView1.Items(SubItemIndex + n).SubItems.Add(FileExtension.ToString() & " File")
    '                ListView1.Items(SubItemIndex + n).SubItems.Add(DateMod.ToString())
    '                SubItemIndex = SubItemIndex + 1

    '            Next
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try

    '    End If
    '    listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)

    'End Sub
    ' Kod szybki tylko je¿eli wyswielta tekst bez ikon

    'Private Sub treeView1_NodeMouseClick(ByVal sender As Object, _
    '    ByVal e As TreeNodeMouseClickEventArgs) _
    '        Handles TreeView1.NodeMouseClick
    '    Dim FileExtension As String
    '    Dim FolderExtension As String
    '    Dim SubItemIndex As Integer
    '    Dim SubItemIndexs As Integer
    '    Dim DateMod As String
    '    Dim DateMods As String

    '    Dim newSelected As TreeNode = e.Node
    '    listView1.Items.Clear()
    '    Dim nodeDirInfo As DirectoryWInfo = New DirectoryWInfo(newSelected.Tag)
    '    Dim subItems() As ListViewItem.ListViewSubItem
    '    Dim item As ListViewItem = Nothing
    '    'ListTab(0) = CStr(TreeView1.SelectedNode.Tag)
    '    Dim n As Integer = 0
    '    Dim argPath As String
    '    Dim mkey As String
    '    Dim dir As DirectoryWInfo
    '    For Each nodeDirInfo In nodeDirInfo.GetDirectories()

    '        FolderExtension = IO.Path.GetExtension(nodeDirInfo.Name)
    '        DateMods = IO.DirectoryW.GetLastWriteTime(nodeDirInfo.Name)
    '        ' ImageList1.Images.Clear()
    '        '  AddImages(FolderNode)

    '        'If IO.DirectoryW.Exists(nodeDirInfo.FullName) = True Then
    '        '    If argPath = IO.DirectoryW.GetDirectoryWRoot(nodeDirInfo.FullName) Then
    '        '        mkey = "drive"
    '        '    Else
    '        '        mkey = "folder"
    '        '    End If
    '        'ElseIf IO.File.Exists(nodeDirInfo.FullName) = True Then
    '        '    mkey = IO.Path.GetExtension(nodeDirInfo.FullName)
    '        'End If
    '        'If mkey = "folder" Then ImageList1.Images.Add(mkey & "-open", GetShellOpenIconAsImage(nodeDirInfo.FullName))
    '        'ListView1.SmallImageList = ImageList1
    '        item = New ListViewItem(nodeDirInfo.Name, mkey)
    '        'ListView1.Items.Add(nodeDirInfo.Name.Substring(nodeDirInfo.Name.LastIndexOf("\"c) + 1), mkey)
    '        subItems = New ListViewItem.ListViewSubItem() _
    '            {New ListViewItem.ListViewSubItem(item, "Folder File"), _
    '            New ListViewItem.ListViewSubItem(item, _
    '            nodeDirInfo.LastAccessTime.ToShortDateString())}

    '        item.SubItems.AddRange(subItems)
    '        ListView1.Items.Add(item)
    '    Next
    '    'Dim folder As String = CStr(TreeView1.SelectedNode.Tag)
    '    Dim fileb As FileInfo
    '    For Each fileb In nodeDirInfo.GetFiles()

    '        subItems = New ListViewItem.ListViewSubItem() _
    '            {New ListViewItem.ListViewSubItem(item, "File"), _
    '            New ListViewItem.ListViewSubItem(item, _
    '            fileb.LastAccessTime.ToShortDateString())}
    '        'AddImages(fileb.FullName)
    '        item = New ListViewItem(fileb.Name, fileb.ToString())
    '        item.SubItems.AddRange(subItems)
    '        ListView1.Items.Add(item)
    '    Next fileb


    'End Sub
    '=====================================
    'Private Sub Treeview1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
    '    Dim FileExtension As String
    '    Dim FolderExtension As String
    '    Dim SubItemIndex As Integer
    '    Dim SubItemIndexs As Integer
    '    Dim DateMod As String
    '    Dim DateMods As String

    '    ListView1.Items.Clear()

    '    'If TreeView1.SelectedNode.Nodes.Count = 1 AndAlso TreeView1.SelectedNode.Nodes(0).Text = "Loading..." Then

    '    '    TreeView1.SelectedNode.Nodes.Clear()

    '    '    AddAllFolders(TreeView1.SelectedNode, CStr(TreeView1.SelectedNode.Tag))
    '    '    'TreeView1.SelectedImageKey=
    '    'End If
    '    'Dim folder As String = CStr(e.Node.Tag)
    '    Dim folder As String = CStr(TreeView1.SelectedNode.Tag)

    '    ListTab(0) = CStr(TreeView1.SelectedNode.Tag)
    '    Dim n As Integer = 0
    '    Dim argPath As String
    '    Dim mkey As String
    '    TextBox6.Text = folder ' dodanie info o scie¿ce do folderu
    '    'If Not folder Is Nothing AndAlso IO.DirectoryW.Exists(folder) Then

    '    '        Catch ex As Exception
    '    '        MsgBox(ex.Message)
    '    '    End Try
    '    '    Next
    '    'End If
    '    'AddImages(folder)
    '    If Not folder Is Nothing AndAlso IO.DirectoryW.Exists(folder) Then
    '        ' MsgBox("orety")

    '        Try
    '            For Each FolderNode As String In DirectoryW.GetDirectories(folder)

    '                FolderExtension = IO.Path.GetExtension(FolderNode)
    '                DateMods = IO.DirectoryW.GetLastWriteTime(FolderNode)
    '                ' ImageList1.Images.Clear()
    '                '  AddImages(FolderNode)

    '                If IO.DirectoryW.Exists(FolderNode) = True Then
    '                    If argPath = IO.DirectoryW.GetDirectoryWRoot(FolderNode) Then
    '                        mkey = "drive"
    '                    Else
    '                        mkey = "folder"
    '                    End If
    '                ElseIf IO.File.Exists(FolderNode) = True Then
    '                    mkey = IO.Path.GetExtension(FolderNode)
    '                End If
    '                If mkey = "folder" Then ImageList1.Images.Add(mkey & "-open", GetShellOpenIconAsImage(FolderNode))
    '                ListView1.SmallImageList = ImageList1

    '                ListView1.Items.Add(FolderNode.Substring(FolderNode.LastIndexOf("\"c) + 1), mkey)

    '                ' Create two ImageList objects.

    '                ' Add the ListView to the control collection.
    '                ' mDirectoryWNode.ImageKey = CacheShellIcon(mDirectoryW.FullName)
    '                '  ListView1.Items(SubItemIndexs).SubItems.Add(FolderExtension.ToString() & " Folder File")
    '                ListView1.Items(SubItemIndexs).SubItems.Add("Folder File")
    '                ListView1.Items(SubItemIndexs).SubItems.Add(DateMods.ToString())
    '                SubItemIndexs += 1
    '                n = n + 1
    '            Next

    '            For Each file As String In IO.DirectoryW.GetFiles(folder)
    '                FileExtension = IO.Path.GetExtension(file)
    '                DateMod = IO.File.GetLastWriteTime(file).ToString()
    '                AddImages(file)
    '                ListView1.Items.Add(file.Substring(file.LastIndexOf("\"c) + 1), file.ToString())
    '                ListView1.Items(SubItemIndex + n).SubItems.Add(FileExtension.ToString() & " File")
    '                ListView1.Items(SubItemIndex + n).SubItems.Add(DateMod.ToString())
    '                SubItemIndex = SubItemIndex + 1

    '            Next
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try

    '    End If
    '    listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)

    'End Sub
    '====================
End Module
