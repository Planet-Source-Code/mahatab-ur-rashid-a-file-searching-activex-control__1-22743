VERSION 5.00
Object = "*\ASearchFile.vbp"
Begin VB.Form SearchFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Finder"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin Project1.SearchFile SearchFile1 
      Left            =   2160
      Top             =   360
      _ExtentX        =   609
      _ExtentY        =   635
      FileName        =   ""
      FileListIndex   =   -1
      FilePath        =   "D:\Program Files\Microsoft Visual Studio\VB98"
      FolderListIndex =   -1
      FolderPath      =   "D:\Program Files\Microsoft Visual Studio\VB98"
      List0           =   ""
      ListIndex       =   -1
      DriveListIndex  =   2
      Drive           =   "d: [MMSTUDIO]"
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2640
      Top             =   360
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "Form1.frx":0442
      Left            =   120
      List            =   "Form1.frx":0444
      TabIndex        =   4
      Top             =   2580
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "*.exe"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblTotalFiles 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1920
      TabIndex        =   7
      Top             =   4080
      Width           =   45
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   3120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total Number Of File(s):"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   1665
   End
   Begin VB.Label lblpath 
      Caption         =   "File List:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Search In Drive:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter File Name To Search:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1965
   End
End
Attribute VB_Name = "SearchFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CmdSearch_Click()

Dim TotalNumberOfFileFound As String
Dim i As Integer
Dim Res

If CmdSearch.Caption = "&Search" Then
    List1.Clear
    lblTotalFiles.Caption = 0
    CmdSearch.Caption = "&Stop"
    Timer1.Enabled = True
    SearchFile1.Search (txtSearch.Text)  ' Search for the file
    TotalNumberOfFileFound = SearchFile1.ListCount - 1   ' Number of file found

For i = 1 To TotalNumberOfFileFound
    List1.AddItem SearchFile1.List(i)      ' Add file to the List box
Next i

lblTotalFiles.Caption = TotalNumberOfFileFound
ReSet                                      ' Reset Search

If SearchFile1.SearchFlag = True Then      ' End of Search
    MsgBox "End of Search!", vbInformation
End If
Else
If CmdSearch.Caption = "&Stop" Then
    Res = MsgBox("Are you sure?", vbYesNo)
    If Res = vbYes Then
        ReSet
        List1.Clear
    Else
    End If
End If
End If

End Sub

Private Sub Combo1_Click()

ReSet       ' Reset Search

End Sub

Private Sub Form_Load()

Dim TotalNumberOfDrive As Integer, i As Integer

TotalNumberOfDrive = SearchFile1.DriveListCount - 1  ' Total number of Drive
For i = 0 To TotalNumberOfDrive
    Combo1.AddItem SearchFile1.DriveList(i)          ' Add Drive name to combo box
Next i

Combo1.ListIndex = 1
ReSet

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload SearchFile
Set SearchFile = Nothing
End Sub

Private Sub Timer1_Timer()

If SearchFile1.SearchFlag Then
    lblpath.Caption = SearchFile1.SearchingInDir     ' Show the folder currently searching
Else
    lblpath.Caption = "File List:"
    Timer1.Enabled = False
End If

End Sub

Public Sub ReSet()

SearchFile1.ReSetSearch   ' Reset Search process in the ActiveX control
SearchFile1.Drive = Combo1.List(Combo1.ListIndex)  ' Default drive to search
SearchFile1.FolderPath = SearchFile1.DriveLatterOnly(SearchFile1.Drive) ' Default folder to search. Here it is root folder (i.e, c:\)
' DriveLatterOnly Method cut the label of the drive and return only drive latter (i.e, C: [Master] -> C:\)
CmdSearch.Caption = "&Search"

End Sub
