VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB Ripper"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   5520
      TabIndex        =   8
      Text            =   "0"
      Top             =   540
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3000
      TabIndex        =   6
      Text            =   "0"
      Top             =   540
      Width           =   1275
   End
   Begin VB.ComboBox TrackList 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   540
      Width           =   1755
   End
   Begin VB.ComboBox DriveList 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   6735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Version info"
      Height          =   555
      Left            =   3660
      TabIndex        =   1
      Top             =   1020
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rip Track"
      Height          =   555
      Left            =   180
      TabIndex        =   0
      Top             =   1020
      Width           =   3435
   End
   Begin VB.Label Label2 
      Caption         =   "End Address"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   540
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Start Address"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   540
      Width           =   855
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00%"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   1740
      Width           =   6675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This project and modifications to AKRip was made by Arto Rusanen
' http://www.4dsoftware.8m.com


' Credits:

' AKRip was orginally made by Andy Key and you can find AKRip and its source from
' http://akrip.sourceforge.net/




Option Explicit
Dim CurDrive As Long

' This one was first test that does my lil trick on AKRip work..
Private Sub Command3_Click()
  Dim ver As DWORD
  ver = GetAKRipDllVersion()
  MsgBox ver.LOWORD & "." & ver.HIWORD
End Sub

' This one was second...
Private Sub Command1_Click()
  Call RipTrack(val(Text1.Text), val(Text2.Text), "Track " & TrackList.ListIndex + 1 & ".wav")
End Sub


Private Sub DriveList_Click()
  On Error Resume Next
  TrackList.Clear
  Call DeInitCDDrive
  If Not InitCDDrive(DriveList.ListIndex) Then Exit Sub
  
  Dim i As Long
  
  Do While MSB2LONG(DiscToc.tracks(i + 2).addr) <> 0
    TrackList.AddItem "Track " & i + 1
    i = i + 1
  Loop
  TrackList.ListIndex = 0
End Sub


' And  following came after I succesfully riped first track of my CD
Private Sub Form_Load()
  Dim DriveCount As Long
  Dim MyInfo As CDREC
  ChDrive App.Path
  ChDir App.Path

  DriveCount = GetNumAdapters + 1
  
  Dim i As Long
  For i = 1 To DriveCount
    MyInfo = GetDriveInformation(i - 1)
    DriveList.AddItem StripNulls(MyInfo.id)
  Next i
  
  DriveList.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call DeInitCDDrive
End Sub


Private Sub TrackList_Click()
  Text1.Text = MSB2LONG(DiscToc.tracks(TrackList.ListIndex + 1).addr)
  Text2.Text = MSB2LONG(DiscToc.tracks(TrackList.ListIndex + 2).addr)
End Sub
