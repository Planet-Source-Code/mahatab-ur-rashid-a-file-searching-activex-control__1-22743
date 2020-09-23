VERSION 5.00
Begin VB.UserControl SearchFile 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   660
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   450
   ScaleWidth      =   660
   ToolboxBitmap   =   "SearchFile.ctx":0000
   Begin VB.ListBox lstFoundFiles 
      Height          =   645
      Left            =   3240
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.DriveListBox drvList 
      Height          =   315
      Left            =   5160
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.DirListBox dirList 
      Height          =   1890
      Left            =   5160
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.FileListBox filList 
      Height          =   2235
      Left            =   3240
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Left            =   0
      Picture         =   "SearchFile.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "SearchFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    '==== SearchFile ActiveX Control ====
    '====    by Mahatab-Ur-Rashid    ====
    '==== rashid_mahatab@yahoo.co.uk ====

Option Explicit
Public SearchFlag As Integer   ' Used as flag for cancel and other operations.
Public SearchingInDir As String ' Currently which folder searching

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=filList,filList,-1,FileName
Public Property Get Filename() As String
Attribute Filename.VB_Description = "Returns/sets the path and filename of a selected file."
    Filename = filList.Filename
End Property

Public Property Let Filename(ByVal New_FileName As String)
    filList.Filename() = New_FileName
    PropertyChanged "FileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=filList,filList,-1,List
Public Property Get FileList(ByVal Index As Integer) As String
Attribute FileList.VB_Description = "Returns/sets the items contained in a control's list portion."
    FileList = filList.List(Index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=filList,filList,-1,ListCount
Public Property Get FileListCount() As Integer
Attribute FileListCount.VB_Description = "Returns the number of items in the list portion of a control."
    FileListCount = filList.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=filList,filList,-1,ListIndex
Public Property Get FileListIndex() As Integer
Attribute FileListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    FileListIndex = filList.ListIndex
End Property

Public Property Let FileListIndex(ByVal New_FileListIndex As Integer)
    filList.ListIndex() = New_FileListIndex
    PropertyChanged "FileListIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=filList,filList,-1,Path
Public Property Get FilePath() As String
Attribute FilePath.VB_Description = "Returns/sets the current path."
    FilePath = filList.Path
End Property

Public Property Let FilePath(ByVal New_FilePath As String)
    filList.Path() = New_FilePath
    PropertyChanged "FilePath"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=filList,filList,-1,Pattern
Public Property Get FilePattern() As String
Attribute FilePattern.VB_Description = "Returns/sets a value indicating the filenames displayed in a control at run time."
    FilePattern = filList.Pattern
End Property

Public Property Let FilePattern(ByVal New_FilePattern As String)
    filList.Pattern() = New_FilePattern
    PropertyChanged "FilePattern"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=drvList,drvList,-1,List
Public Property Get DriveList(ByVal Index As Integer) As String
Attribute DriveList.VB_Description = "Returns/sets the items contained in a control's list portion."
    DriveList = drvList.List(Index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=dirList,dirList,-1,List
Public Property Get FolderList(ByVal Index As Integer) As String
Attribute FolderList.VB_Description = "Returns/sets the items contained in a control's list portion."
    FolderList = dirList.List(Index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=dirList,dirList,-1,ListCount
Public Property Get FolderListCount() As Integer
Attribute FolderListCount.VB_Description = "Returns the number of items in the list portion of a control."
    FolderListCount = dirList.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=dirList,dirList,-1,ListIndex
Public Property Get FolderListIndex() As Integer
Attribute FolderListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    FolderListIndex = dirList.ListIndex
End Property

Public Property Let FolderListIndex(ByVal New_FolderListIndex As Integer)
    dirList.ListIndex() = New_FolderListIndex
    PropertyChanged "FolderListIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=dirList,dirList,-1,Path
Public Property Get FolderPath() As String
Attribute FolderPath.VB_Description = "Returns/sets the current path."
    FolderPath = dirList.Path
End Property

Public Property Let FolderPath(ByVal New_FolderPath As String)
    dirList.Path() = New_FolderPath
    PropertyChanged "FolderPath"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstFoundFiles,lstFoundFiles,-1,List
Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
    List = lstFoundFiles.List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    lstFoundFiles.List(Index) = New_List
    PropertyChanged "List"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstFoundFiles,lstFoundFiles,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = lstFoundFiles.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstFoundFiles,lstFoundFiles,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = lstFoundFiles.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    lstFoundFiles.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstFoundFiles,lstFoundFiles,-1,Sorted
Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = lstFoundFiles.Sorted
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstFoundFiles,lstFoundFiles,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of a control or the system Clipboard."
    lstFoundFiles.Clear
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=drvList,drvList,-1,ListCount
Public Property Get DriveListCount() As Integer
Attribute DriveListCount.VB_Description = "Returns the number of items in the list portion of a control."
    DriveListCount = drvList.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=drvList,drvList,-1,ListIndex
Public Property Get DriveListIndex() As Integer
Attribute DriveListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    DriveListIndex = drvList.ListIndex
End Property

Public Property Let DriveListIndex(ByVal New_DriveListIndex As Integer)
    drvList.ListIndex() = New_DriveListIndex
    PropertyChanged "DriveListIndex"
End Property
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer

    filList.Filename = PropBag.ReadProperty("FileName", "0")
    filList.ListIndex = PropBag.ReadProperty("FileListIndex", 0)
    filList.Path = PropBag.ReadProperty("FilePath", "0")
    filList.Pattern = PropBag.ReadProperty("FilePattern", "*.*")
    dirList.ListIndex = PropBag.ReadProperty("FolderListIndex", 0)
    dirList.Path = PropBag.ReadProperty("FolderPath", "0")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    lstFoundFiles.List(Index) = PropBag.ReadProperty("List" & Index, "0")
    lstFoundFiles.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    drvList.ListIndex = PropBag.ReadProperty("DriveListIndex", 0)
    drvList.Drive = PropBag.ReadProperty("Drive", "")
    End Sub

Private Sub UserControl_Resize()
UserControl.Height = Image1.Height
UserControl.Width = Image1.Width
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer

    Call PropBag.WriteProperty("FileName", filList.Filename, "0")
    Call PropBag.WriteProperty("FileListIndex", filList.ListIndex, 0)
    Call PropBag.WriteProperty("FilePath", filList.Path, "0")
    Call PropBag.WriteProperty("FilePattern", filList.Pattern, "*.*")
    Call PropBag.WriteProperty("FolderListIndex", dirList.ListIndex, 0)
    Call PropBag.WriteProperty("FolderPath", dirList.Path, "0")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("List" & Index, lstFoundFiles.List(Index), "0")
    Call PropBag.WriteProperty("ListIndex", lstFoundFiles.ListIndex, 0)
    Call PropBag.WriteProperty("DriveListIndex", drvList.ListIndex, 0)
    Call PropBag.WriteProperty("Drive", drvList.Drive, "")
    End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Search(Filename As String) As String
' Initialize for search, then perform recursive search.
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
  ' Check what the user did last.
    
  ' Update dirList.Path if it is different from the currently
  ' selected directory, otherwise perform the search.
    If dirList.Path <> dirList.List(dirList.ListIndex) Then
        dirList.Path = dirList.List(dirList.ListIndex)
        Exit Function         ' Exit so user can take a look before searching.
    End If

    ' Continue with the search.
    filList.Pattern = Filename
    FirstPath = dirList.Path
    DirCount = dirList.ListCount

    ' Start recursive direcory search.
    result = DirDiver(FirstPath, DirCount, "")
    filList.Path = dirList.Path
End Function


Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
'  Recursively search directories from NewPath down...
'  NewPath is searched on this recursion.
'  BackUp is origin of this recursion.
'  DirCount is number of subdirectories in this directory.
Static FirstErr As Integer

Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String, FilePath As String
Dim retval, FileCount As Integer
    SearchFlag = True
    retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    DirDiver = False            ' Set to True if there is an error.
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = dirList.ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = dirList.Path                      ' Save old path for next recursion.
        dirList.Path = NewPath
        If dirList.ListCount > 0 Then
            ' Get to the node bottom.
            dirList.Path = dirList.List(DirsToPeek - 1)
            SearchingInDir = dirList.Path
            AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Call function to enumerate files.
    If filList.ListCount Then
        If Len(dirList.Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = dirList.Path                  ' If at root level, leave as is...
        Else
            ThePath = dirList.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        For ind = 0 To filList.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = ThePath + filList.List(ind)
            lstFoundFiles.AddItem entry
            
            Next ind
    End If
    If BackUp <> "" Then        ' If there is a superior directory, move it.
        dirList.Path = BackUp
    End If
    Exit Function

DirDriverHandler:
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiver = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "Out of Memory! Abandoning search...", vbCritical
        Exit Function           ' Note that the exit procedure resets Err to 0.
    End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function ResetSearch() As Variant
lstFoundFiles.Clear
    SearchFlag = False                  ' Flag indicating search in progress.
    dirList.Path = CurDir: drvList.Drive = dirList.Path ' Reset the path.
End Function

Private Sub DirList_Change()
    ' Update the file list box to synchronize with the directory list box.
    filList.Path = dirList.Path
End Sub
'

Private Sub DrvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub

DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=drvList,drvList,-1,Drive
Public Property Get Drive() As String
Attribute Drive.VB_Description = "Returns/sets the selected drive at run time."
    Drive = drvList.Drive
End Property

Public Property Let Drive(ByVal New_Drive As String)
    drvList.Drive() = New_Drive
    PropertyChanged "Drive"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13
Public Function DriveLatterOnly(DriveLatterWithLabelName As String) As String
    DriveLatterOnly = Left(DriveLatterWithLabelName, 2) + "\"
End Function


Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "About The Author"
Attribute ShowAbout.VB_UserMemId = -552
    About.Show vbModal
    Unload About
    Set About = Nothing
End Sub
