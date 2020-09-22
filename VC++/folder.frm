VERSION 5.00
Begin VB.Form Folder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose a folder"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Dir:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2040
      MouseIcon       =   "folder.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4920
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3120
      MouseIcon       =   "folder.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4920
      Width           =   630
   End
End
Attribute VB_Name = "Folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
Text1.Text = Dir1.path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub

Private Sub Form_Load()
forward Me
End Sub

Private Sub Label1_Click()
If Text1.Text <> "" Or Text2.Text <> "" Then
MkDir Text1.Text & Text2.Text
MsgBox "Directory created!"
Dir1.path = Dir1.path & Text2.Text
Text1.Text = Dir1.path
Text2.Text = ""
Else
MsgBox "no directory selected!"
End If
End Sub

Private Sub Label6_Click()
Builder.d_folder.Text = Text1.Text
Builder.m_folder.Text = Text1.Text
Unload Me
End Sub
