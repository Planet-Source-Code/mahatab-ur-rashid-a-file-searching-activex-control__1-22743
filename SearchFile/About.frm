VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Author"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3750
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "rashid_mahatab@yahoo.co.uk"
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "by Mahatab-Ur-Rashid"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "SearchFile ActiveX Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   510
      TabIndex        =   0
      Top             =   240
      Width           =   2805
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
