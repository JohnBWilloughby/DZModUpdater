VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Path History Implementation"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   8655
   End
   Begin VB.CommandButton cmdLoadPath 
      Caption         =   "LOAD  PATH  FROM  HISTORY"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   4920
      Width           =   3855
   End
   Begin VB.ComboBox cmbSavedPaths 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5640
      Width           =   8655
   End
   Begin VB.CommandButton cmdSavePath 
      Caption         =   "SAVE  PATH  TO  HISTORY"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   3855
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   5640
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   3060
      Left            =   2760
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "VBFORUMS  -  MANIPULATING   THE   PATH   HISTORY"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MANIPULATING THE PATH HISTORY - CODE BY lone_REBEL OF WWW.VBFORUMS.COM
'----------------------------------------------------------------------
'
'this coding shows how to successfully create, save and load directory and file
'paths anywhere on the hard disk using a drive listbox, a directory listbox and a files
'listbox

'to understand this code better, start with reading the code for DRIVE1 events, then
'DIR1 events and then FILE1 events. clicking on these lists' values allows the user to
'browse through the hard disk. a variable declared as fullpath (STRING) saves the path
'which is being browsed by the user.

'after understanding how the browse-path is created as a string, proceed with reading
'the coding for the command buttons which are used for saving and loading browsing
'history.
'there are two command buttons. one of them saves the current path being browsed by the
'user while the other loads the saved path. the saved paths are stored in a combo box

Option Explicit

Dim fullpath As String


Private Sub cmdLoadPath_Click()
    'these variables are used to get the relative values from the saved path
    Dim drive As String, directory As String, file As String
    'we load cmbSavedPaths.Text to a simpler variable so that its easy to use in our code
    Dim path As String
    'plc is used to get the location of "\" in the path and i is used as a loop counter
    Dim plc As Integer, i As Integer
    
    'if there is no path to load, then simply exit
    If cmbSavedPaths.ListIndex = -1 Then Exit Sub
    'if there is a path available to load, first copy it to our string variable
    path = cmbSavedPaths.Text
    'the drive information is present in the first two letters of the path
    'e.g. if path is G:\Songs\Tracks then drive would be G:
    drive = Left(path, 2)
    
    'select the saved drive letter in the drives listbox
    For i = 0 To Drive1.ListCount
        If Left(UCase(Drive1.List(i)), 2) = drive Then
            Drive1.ListIndex = i
            Exit For
        End If
    Next i
    
    'now get the directory path. if the user has saved the path with a file name at
    'the end, then the last character of the path will not be "\", otherwise (if the
    'user saved a directory path to the save list), then the last character will be "\"
    plc = InStrRev(path, "\")
    
    'if only the directory path has been saved, then there's no need to get the file
    'name from the saved value
    If plc = Len(path) Then
        fullpath = path
        Dir1.path = fullpath
        File1.path = fullpath
        
    'if the filename has also been saved (by clicking on a file before pressing the save
    'button) then we need to get the file name also from the path
    Else
        file = Mid(path, plc + 1)
        fullpath = Left(path, plc)
        Dir1.path = fullpath
        Dir1.Refresh
        File1.path = fullpath
        File1.Refresh
        
        'after getting the file name, we should select that file in the files listbox
        For i = 0 To File1.ListCount - 1
            If File1.List(i) = file Then
                File1.ListIndex = i
                Exit Sub
            End If
        Next i
    End If
End Sub


'the save button simply adds FULLPATH to cmbSavePath list
Private Sub cmdSavePath_Click()
    cmbSavedPaths.AddItem fullpath
    If cmbSavedPaths.ListIndex = -1 Then cmbSavedPaths.ListIndex = 0
End Sub


'when the directory listbox is clicked, we set the path to the path of directory which
'the user has clicked in the list
Private Sub Dir1_Click()
    fullpath = Dir1.List(Dir1.ListIndex)
    'if the backslash (\) symbol is not included in the directory name, then add it
    If Right(fullpath, 1) <> "\" Then fullpath = fullpath & "\"
    'now set this path as the directory and file path
    Dir1.path = fullpath
    File1.path = fullpath
    txtPath.Text = fullpath
End Sub


'if the user pressed ENTER on any item in the directory listbox, repeat the procedure
'defined above
Private Sub Dir1_KeyPress(KeyAscii As Integer)
    If Dir1.ListIndex = -1 Then Exit Sub
    If KeyAscii = 13 Then
        fullpath = Dir1.List(Dir1.ListIndex)
        If Right(fullpath, 1) <> "\" Then fullpath = fullpath & "\"
        Dir1.path = fullpath
        File1.path = fullpath
        txtPath.Text = fullpath
    End If
End Sub


'when the drive letter is changed, we need to reset the whole browse-path to the
'current drive letter
Private Sub Drive1_Change()
    'we get the drive letter and the ":" sign from the listbox (which also displays
    'the drive name/title)
    fullpath = UCase(Left(Drive1.drive, 2) & "\")
    txtPath.Text = fullpath
    'set this drive letter as the path for directory and file listboxes
    Dir1.path = fullpath
    File1.path = fullpath
End Sub


'when the user pressed ENTER on the drive listbox, simply repeat the change procedure
'defined above
Private Sub Drive1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fullpath = UCase(Left(Drive1.drive, 2) & "\")
        txtPath.Text = fullpath
        Dir1.path = fullpath
        File1.path = fullpath
    End If
End Sub


'when any item in the files listbox is clicked, we add the file name to the browse-path
Private Sub File1_Click()
    'this plc variable is used to remember the value of the last \ in the browse path
    Dim plc As Integer
    'if there is no selected value (i.e. the files list is empty) then just exit
    If File1.ListIndex = -1 Then Exit Sub
    'find the path of the last \ so that we get the path upto the last directory level
    plc = InStrRev(fullpath, "\")
    'now remove the previous file name (if any) from the browse path and add the new
    'file name to the browse path
    fullpath = Left(fullpath, plc) & File1.List(File1.ListIndex)
    'finally set this path and filename as the new browse path
    txtPath.Text = fullpath
End Sub


Private Sub Form_Load()
    'at the beginning of the project, set the path to C:\
    Drive1_KeyPress 13
End Sub
