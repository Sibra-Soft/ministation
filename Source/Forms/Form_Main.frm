VERSION 5.00
Object = "{69DBEE3D-E09E-4122-9CAA-E6734195BEEC}#1.0#0"; "d3DLine.ocx"
Object = "{0A362340-2E5E-11D3-85BF-00105AC8B715}#1.0#0"; "isDigitalLibrary.ocx"
Object = "{BF3128D8-55B8-11D4-8ED4-00E07D815373}#1.0#0"; "MBPrgBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{40F6D89D-D6BF-4EAD-B885-E1869BDF4E31}#41.0#0"; "AdioLibrary.ocx"
Begin VB.Form Form_Main 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ministation"
   ClientHeight    =   2070
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin AdioLibrary.AdioCore AdioCore 
      Left            =   120
      Top             =   1080
      _ExtentX        =   2778
      _ExtentY        =   873
      Begin AdioLibrary.AdioPlaylist AdioPlaylist 
         Left            =   1080
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioMediaPlayer AdioMediaPlayer 
         Left            =   600
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
   End
   Begin VB.PictureBox Picturebox_PanelBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   4665
      TabIndex        =   2
      Top             =   1770
      Width           =   4665
      Begin MBProgressBar.ProgressBar Progressbar_CurrentPos 
         Height          =   180
         Left            =   45
         TabIndex        =   3
         Top             =   75
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   318
         BorderStyle     =   2
         CaptionType     =   0
         Smooth          =   -1  'True
         BackColor       =   12632256
         BarStartColor   =   16711680
         BarEndColor     =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackPicture     =   "Form_Main.frx":030A
         BarPicture      =   "Form_Main.frx":0326
      End
      Begin Graphical_Line.d3DLine d3DLine3 
         Height          =   30
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   53
      End
   End
   Begin VB.PictureBox Picturebox_PanelPlaySettings 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   4695
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
      Begin Ministation.ButtonBig Button_CancelSettings 
         Height          =   390
         Left            =   3720
         TabIndex        =   20
         Top             =   1150
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         Caption         =   "Cancel"
         TextAlignment   =   0
      End
      Begin VB.OptionButton Option_Shuffle 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Shuffle (only for playlists)"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton Option_RepeatTrack 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reapeat Track"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option_NormalPlay 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Normal Play"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Value           =   -1  'True
         Width           =   1455
      End
      Begin Ministation.ButtonBig Button_SaveSettings 
         Height          =   390
         Left            =   2760
         TabIndex        =   21
         Top             =   1150
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         Caption         =   "Save"
         TextAlignment   =   0
      End
   End
   Begin MSComctlLib.ImageList Imagelist_Digits 
      Left            =   4080
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":0818
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":0CEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":11C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":169A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":1B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":2046
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":251C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":29F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":2EC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer_Main 
      Interval        =   500
      Left            =   2640
      Top             =   1320
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3360
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Button_MediaPlayer 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1620
      Picture         =   "Form_Main.frx":339E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Button_MediaPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1140
      Picture         =   "Form_Main.frx":3928
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox Picturebox_Display 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   3915
      TabIndex        =   10
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox Picturebox_KBit 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   2520
         ScaleHeight     =   150
         ScaleWidth      =   375
         TabIndex        =   15
         Top             =   400
         Width           =   375
      End
      Begin VB.PictureBox Picturebox_Khz 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   2520
         ScaleHeight     =   150
         ScaleWidth      =   375
         TabIndex        =   14
         Top             =   220
         Width           =   375
      End
      Begin isDigitalLibrary.iSevenSegmentIntegerX SegmentDisplay_TrackCount 
         Height          =   540
         Left            =   75
         TabIndex        =   11
         Top             =   135
         Width           =   615
         Value           =   0
         ShowSign        =   0   'False
         DigitCount      =   2
         LeadingStyle    =   2
         AutoSize        =   -1  'True
         DigitSpacing    =   6
         SegmentMargin   =   4
         SegmentColor    =   65280
         SegmentSeperation=   1
         SegmentSize     =   1
         ShowOffSegments =   -1  'True
         PowerOff        =   0   'False
         BackGroundColor =   0
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         SegmentOffColor =   8421504
         AutoSegmentOffColor=   -1  'True
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   41
         Object.Height          =   36
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSevenSegmentIntegerX SegmentDisplay_Minutes 
         Height          =   540
         Left            =   840
         TabIndex        =   12
         Top             =   135
         Width           =   615
         Value           =   0
         ShowSign        =   0   'False
         DigitCount      =   2
         LeadingStyle    =   2
         AutoSize        =   -1  'True
         DigitSpacing    =   6
         SegmentMargin   =   4
         SegmentColor    =   65280
         SegmentSeperation=   1
         SegmentSize     =   1
         ShowOffSegments =   -1  'True
         PowerOff        =   0   'False
         BackGroundColor =   0
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         SegmentOffColor =   8421504
         AutoSegmentOffColor=   -1  'True
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   41
         Object.Height          =   36
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSevenSegmentIntegerX SegmentDisplay_Seconds 
         Height          =   540
         Left            =   1560
         TabIndex        =   13
         Top             =   135
         Width           =   615
         Value           =   0
         ShowSign        =   0   'False
         DigitCount      =   2
         LeadingStyle    =   2
         AutoSize        =   -1  'True
         DigitSpacing    =   6
         SegmentMargin   =   4
         SegmentColor    =   65280
         SegmentSeperation=   1
         SegmentSize     =   1
         ShowOffSegments =   -1  'True
         PowerOff        =   0   'False
         BackGroundColor =   0
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         SegmentOffColor =   8421504
         AutoSegmentOffColor=   -1  'True
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   41
         Object.Height          =   36
         OPCItemCount    =   0
      End
      Begin VB.Image Image9 
         Height          =   240
         Left            =   3360
         Picture         =   "Form_Main.frx":3EB2
         Top             =   200
         Width           =   450
      End
      Begin VB.Image Image_Pause 
         Height          =   255
         Left            =   2280
         Picture         =   "Form_Main.frx":44B4
         Top             =   120
         Width           =   195
      End
      Begin VB.Image Image_Play 
         Height          =   255
         Left            =   2250
         Picture         =   "Form_Main.frx":479E
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image Image8 
         Height          =   150
         Left            =   2950
         Picture         =   "Form_Main.frx":4A88
         Top             =   420
         Width           =   495
      End
      Begin VB.Image Image7 
         Height          =   165
         Left            =   2520
         Picture         =   "Form_Main.frx":4EB2
         Top             =   10
         Width           =   390
      End
      Begin VB.Image Image6 
         Height          =   165
         Left            =   2950
         Picture         =   "Form_Main.frx":5264
         Top             =   240
         Width           =   330
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   2280
         Picture         =   "Form_Main.frx":5592
         Top             =   120
         Width           =   150
      End
      Begin VB.Image Image4 
         Height          =   150
         Left            =   1850
         Picture         =   "Form_Main.frx":57D4
         Top             =   10
         Width           =   315
      End
      Begin VB.Image Image3 
         Height          =   225
         Left            =   1440
         Picture         =   "Form_Main.frx":5A96
         Top             =   240
         Width           =   120
      End
      Begin VB.Image Image2 
         Height          =   150
         Left            =   1120
         Picture         =   "Form_Main.frx":5C40
         Top             =   15
         Width           =   300
      End
      Begin VB.Image Image1 
         Height          =   165
         Left            =   120
         Picture         =   "Form_Main.frx":5EDA
         Top             =   10
         Width           =   525
      End
   End
   Begin VB.CommandButton Button_MediaPlayer 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "Form_Main.frx":63C0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Button_MediaPlayer 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   660
      Picture         =   "Form_Main.frx":698A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Button_MediaPlayer 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2160
      Picture         =   "Form_Main.frx":6F14
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin Graphical_Line.d3DLine d3DLine2 
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   238
   End
   Begin Graphical_Line.d3DLine d3DLine1 
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   238
   End
   Begin VB.Menu MenuItem_File 
      Caption         =   "&File"
      Begin VB.Menu MenuItem_File_PlayFile 
         Caption         =   "&Play File..."
      End
      Begin VB.Menu MenuItem_File_PlayLocation 
         Caption         =   "&Play Location..."
      End
      Begin VB.Menu MenuItem_File_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuItem_File_Exit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenuItem_Options 
      Caption         =   "&Options"
      Begin VB.Menu MenuItem_Options_Play 
         Caption         =   "&Play"
      End
   End
   Begin VB.Menu MenuItem_Help 
      Caption         =   "&Help"
      Begin VB.Menu MenuItem_Help_About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SavedLocations As New Collection
Public Sub AddToPlaylist(Files As String)
On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

AdioPlaylist.AddMultipleFiles Files

'Restore default
Screen.MousePointer = vbDefault
Exit Sub

ErrorHandler:
Screen.MousePointer = vbDefault

Select Case Err.Number
    Case 0
    Case 35602
    Case Else: Debug.Print Err.Number & " - " & Err.Description
End Select
End Sub

Private Sub Button_CancelSettings_Click()
Picturebox_PanelPlaySettings.Visible = False
End Sub

Private Sub Button_MediaPlayer_Click(index As Integer)
Select Case index
    Case 0: Call AdioMediaPlayer.LoadFile(AdioPlaylist.GetTrack(PLS_PREV))
    Case 1: Call AdioMediaPlayer.StopPlay
    Case 2
        If AdioPlaylist.GetList.Count = 0 Then
            MenuItem_File_PlayFile_Click
        Else
            AudiostationMP3Player.StartPlay
        End If

    Case 3: Call AdioMediaPlayer.PausePlay
    Case 4: Call AdioMediaPlayer.LoadFile(AdioPlaylist.GetTrack(PLS_NEXT))
End Select
End Sub

Private Sub Button_SaveSettings_Click()

If Option_NormalPlay.value Then
    PlayMode = NormalPlay
ElseIf Option_RepeatTrack.value Then
    PlayMode = RepeatTrack
ElseIf Option_Shuffle.value Then
    PlayMode = Shuffle
End If

Picturebox_PanelPlaySettings.Visible = False
End Sub

Private Sub Form_Load()
Dim Locations As String

Locations = Settings.ReadSetting("Sibra-Soft", "Ministation", "SavedLocations", "")
Set SavedLocations = Extensions.StringToCollection(Locations, vbNewLine)

ChDrive App.path
ChDir App.path
End Sub

Private Sub MenuItem_File_Exit_Click()
Unload Me
End Sub

Private Sub MenuItem_File_PlayFile_Click()
Dim Files As String

On Error GoTo ErrorHandler
With CommonDialog
    .CancelError = True
    .MaxFileSize = 9999
    .DialogTitle = "Play audio file(s)"
    .Filter = "MPEG-1 Layer 3 (*.mp3)|*.mp3|Microsoft WaveForm Audio (*.wav)|*.wav"
    .flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
    .ShowOpen

    If .filename <> vbNullString Then
        Files = Extensions.CommondialogFilesToList(.filename)
        
        AudiostationMP3Player.TrackNr = 1
        
        Call AddToPlaylist(Files)
        Call AudiostationMP3Player.StartPlay
    End If
End With

ErrorHandler:
Select Case Err.Number
    Case 0
    Case Else: Debug.Print Err.Description
End Select
End Sub

Private Sub MenuItem_File_PlayLocation_Click()
Dim LocalFile As String

Form_OpenLocation.Show vbModal, Me

If Form_OpenLocation.DialogResult = vbOK Then
    If Not Extensions.CollectionContains(SavedLocations, Form_OpenLocation.Location) Then
        SavedLocations.Add Form_OpenLocation.Location
    End If
    
    LocalFile = App.path & "\temp.mp3"
    
    Call Extensions.RemoveFile(LocalFile)
    Call Extensions.DownloadFile(Form_OpenLocation.Location, LocalFile)
    
    Call BASS_StreamFree(chan)
    Call BASS_MusicFree(chan)
    
    chan = BASS_StreamCreateFile(BASSFALSE, StrPtr(LocalFile), 0, 0, BASS_STREAM_AUTOFREE)
    If chan = 0 Then chan = BASS_MusicLoad(BASSFALSE, StrPtr(LocalFile), 0, 0, BASS_STREAM_AUTOFREE, 1)
    
    Call BASS_ChannelPlay(chan, True)
    
    If BASS_IsStarted Then
        Call Settings.WriteSetting("Sibra-Soft", "Ministation", "SavedLocations", Extensions.CollectionToString(SavedLocations, vbNewLine))
        
        TrackNr = 1
        PlayState = Playing
    End If
End If
End Sub

Private Sub MenuItem_Help_About_Click()
Form_About.Show vbModal, Me
End Sub

Private Sub MenuItem_Options_Play_Click()
Option_Shuffle.Enabled = True
If AdioPlaylist.RepeatMode = PLS_SHUFFLE Then Option_Shuffle.Enabled = False

Picturebox_PanelPlaySettings.Visible = True
End Sub

