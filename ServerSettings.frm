VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DayZ Server "
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   4320
      TabIndex        =   7
      Top             =   4560
      Width           =   5895
   End
   Begin VB.TextBox txtdzmodinfo 
      Height          =   2535
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "ServerSettings.frx":0000
      Top             =   1800
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next >"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6960
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label lblPath 
      Caption         =   "Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblDrive 
      Caption         =   "Drive:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblServerModLoc 
      Caption         =   "DayZ Mod Location - Server "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()
    ' when the drive changes, set the directory box path to the new drive
    On Error GoTo DriveError
    Dir1.Path = Drive1.Drive
    GoTo DriveEnd
DriveError:
        'Drive not Acessible
        Drive1.Drive = Dir1.Path
DriveEnd:
    On Error GoTo 0
    
End Sub

Private Sub Dir1_Change()
    ' when the directory changes, change the file box path to the directory box path
    File1.Path = Dir1.Path
End Sub

Private Sub Command1_Click()
    ' when the user clicks a filename in the file box,
    ' you can tell what file was selected by using
    ' "File1.List(File1.ListIndex)"
    Dim strFullFilename As String
    strFullFilename = Dir1.Path & "\" & File1.List(File1.ListIndex)
    Form1.txtdzmodinfo.Text = Form1.txtdzmodinfo.Text + "DayZ Mod Path: " & strFullFilename + vbCrLf
    SModPath = strFullFilename
    ' MsgBox "You selected " & strFullFilename
End Sub

Private Sub Form_Load()
Form1.txtdzmodinfo.Text = "DayZ Mod Info" + vbCrLf
Form1.File1.Visible = False

End Sub



