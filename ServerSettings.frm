VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DayZ Server "
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   8760
      TabIndex        =   16
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10800
      TabIndex        =   15
      Top             =   7560
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   6360
      TabIndex        =   14
      Top             =   4440
      Width           =   6135
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write >>>"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   7680
      Width           =   1335
   End
   Begin VB.ComboBox cmbDriveList 
      Height          =   315
      ItemData        =   "ServerSettings.frx":0000
      Left            =   360
      List            =   "ServerSettings.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton cmdNext2 
      Caption         =   "Next >>"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox txtcModPath 
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Text            =   "Program Files (x86)\Steam\steamapps\common\DayZ\!Workshop"
      Top             =   3600
      Width           =   5775
   End
   Begin VB.TextBox txtcIPAdd 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Text            =   "Client IP Addres"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   3600
      TabIndex        =   7
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtdzmodinfo 
      Height          =   3135
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "ServerSettings.frx":0004
      Top             =   1200
      Width           =   6135
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   7680
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   360
      TabIndex        =   1
      Top             =   3600
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label lblClientIP 
      Caption         =   "Client IP Address"
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
      TabIndex        =   12
      Top             =   720
      Width           =   2535
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
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3000
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
      Top             =   1800
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
      Left            =   4440
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

Private Sub cmdNext_Click()
    ' when the user clicks a filename in the file box,
    ' you can tell what file was selected by using
    ' "File1.List(File1.ListIndex)"
    
    strFullFilename = Dir1.Path & "\" & File1.List(File1.ListIndex)
    Form1.txtdzmodinfo.Text = Form1.txtdzmodinfo.Text + "Server Mod Path: " & strFullFilename + vbCrLf
    sModPath = strFullFilename
    FindTheMods = sModPath
    Call FindTheModFolders
    ' MsgBox "You selected " & strFullFilename
    Form1.lblServerModLoc.Caption = "DayZ Mod Location - Client"
    Form1.lblClientIP.Visible = True
    Form1.Drive1.Visible = False
    Form1.Dir1.Visible = False
    Form1.txtcIPAdd.Visible = True
    Form1.txtcModPath.Visible = True
    Form1.cmbDriveList.Visible = True
    Form1.cmdNext2.Visible = True
    
    
End Sub

Private Sub cmdNext2_Click()

RemoteDrive = Replace(cmbDriveList.Text, ":\", "$")
cDriveLetter = RemoteDrive
cCompName = txtcIPAdd.Text

'Form1.txtcModPath.Text = "Program Files (x86)\Steam\steamapps\common\DayZ\!Workshop"
If PingIP(txtcIPAdd.Text) Then
    strFullFilename = "\\" + txtcIPAdd.Text + "\" + RemoteDrive + "\" + txtcModPath.Text
    FindTheMods = strFullFilename
    Call FindTheModFolders
    cCompName = txtcIPAdd.Text
    Form1.txtdzmodinfo.Text = Form1.txtdzmodinfo.Text + "Client Mod Path: " & strFullFilename + vbCrLf
    Form1.txtdzmodinfo.Text = Form1.txtdzmodinfo.Text + vbCrLf + vbCrLf + sMessage
    Form1.cmdNext2.Visible = False
    Form1.cmdWrite.Visible = True
Else
    MsgBox "Unable to Find Client Computer", vbCritical
    End
End If

End Sub

Private Sub cmdWrite_Click()

    Call WriteConfigtoFile
    Form1.cmdUpdate.Visible = True
    
  
    
End Sub

Private Sub WriteConfigtoFile()

' Write Configuration to File.


End Sub
Private Sub cmdClose_Click()
    
    Unload Form1
    Unload Form2
    End
    
End Sub

Private Sub cmdUpdate_Click()
    Unload Form1
    Load Form2
    End
    

End Sub
Private Sub Form_Load()

'ConfigFile = App.Path + "\settings.config"
sMessage = "Click on Write to Write Configuration to File" + vbCrLf + "Click on Update to open DayZModUpdater" + vbCrLf
If Not Dir$(App.Path + "\settings.config") = "" Then
    Form1.Visible = False
    Unload Form1
    Load Form2
    Form2.Visible = True
    
'    MsgBox "Settings are Set "
Else

Form1.txtdzmodinfo.Text = "DayZ Mod Info" + vbCrLf
Form1.File1.Visible = False
Form1.cmdUpdate.Visible = False

Form1.lblClientIP.Visible = False
Form1.txtcIPAdd.Visible = False
Form1.txtcModPath.Visible = False
Form1.cmdNext2.Visible = False
Form1.cmdWrite.Visible = False
Form1.cmbDriveList.Visible = False

Form1.cmbDriveList.AddItem ("C:\")
Form1.cmbDriveList.AddItem ("D:\")
Form1.cmbDriveList.AddItem ("E:\")
Form1.cmbDriveList.AddItem ("F:\")
Form1.cmbDriveList.AddItem ("G:\")
Form1.cmbDriveList.AddItem ("H:\")
Form1.cmbDriveList.AddItem ("I:\")
Form1.cmbDriveList.AddItem ("J:\")
Form1.cmbDriveList.AddItem ("K:\")
Form1.cmbDriveList.AddItem ("L:\")
Form1.cmbDriveList.AddItem ("M:\")
Form1.cmbDriveList.AddItem ("N:\")
Form1.cmbDriveList.AddItem ("O:\")
Form1.cmbDriveList.AddItem ("P:\")
Form1.cmbDriveList.AddItem ("Q:\")
Form1.cmbDriveList.AddItem ("R:\")
Form1.cmbDriveList.AddItem ("S:\")
Form1.cmbDriveList.AddItem ("T:\")
Form1.cmbDriveList.AddItem ("U:\")
Form1.cmbDriveList.AddItem ("V:\")
Form1.cmbDriveList.AddItem ("W:\")
Form1.cmbDriveList.AddItem ("X:\")
Form1.cmbDriveList.AddItem ("Y:\")
Form1.cmbDriveList.AddItem ("Z:\")
End If

End Sub

Function PingIP(IP)
    Dim objWMIService
    Dim colItems
    Dim objItem
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_PingStatus Where timeout = 1000 and Address='" & IP & "'")
    For Each objItem In colItems
        If objItem.StatusCode = 0 Then
            PingIP = True
        Else
            PingIP = False
        End If
    Next
End Function

' **************************************************************
' Testing Mapping a drive

Private Sub Command1_Click()
Dim lret
lret = MapNetworkDrive(WhattoMap, "", WheretoMap, "Sorry")
End Sub

Private Sub Command2_Click()
DisconnectNetworkDrive WheretoMap, 1, "Sorry"
End Sub

Private Sub ThisIsATest()
Command1.Caption = "Connect Drive"
Command2.Caption = "Release Drive"
Text1.Text = "\\" + cCompName + "\" + cDriveLetter
WhattoMap = Text1.Text
Text2.Text = "U:"
WheretoMap = Text2.Text
End Sub


' ********************************************************
Private Sub FindTheModFolders()
Dim sFolder As String
sFolder = Dir$(FindTheMods, vbDirectory)
Debug.Print sModPath
Debug.Print sFolder
Do While sFolder <> ""
    If InStr(sFolder, "@") <> 0 Then
        Form1.List1.AddItem (sFolder)
    End If
    sFolder = Dir$
Loop

End Sub
