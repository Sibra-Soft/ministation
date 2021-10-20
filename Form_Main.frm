VERSION 5.00
Object = "{69DBEE3D-E09E-4122-9CAA-E6734195BEEC}#1.0#0"; "d3DLine.ocx"
Object = "{0A362340-2E5E-11D3-85BF-00105AC8B715}#1.0#0"; "isDigitalLibrary.ocx"
Object = "{BF3128D8-55B8-11D4-8ED4-00E07D815373}#1.0#0"; "MBPrgBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form_Main 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ministation"
   ClientHeight    =   2040
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   4665
   Icon            =   "Form_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
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
            Picture         =   "Form_Main.frx":0422
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":08F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":0DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":12A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":177A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":1C50
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":2126
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":25FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":2AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":2FA8
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
      Height          =   375
      Index           =   3
      Left            =   1620
      Picture         =   "Form_Main.frx":347E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Button_MediaPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   2
      Left            =   1140
      Picture         =   "Form_Main.frx":3A08
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox Picturebox_Display 
      BackColor       =   &H00000000&
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
         Picture         =   "Form_Main.frx":3F92
         Top             =   200
         Width           =   450
      End
      Begin VB.Image Image_Pause 
         Height          =   255
         Left            =   2280
         Picture         =   "Form_Main.frx":4594
         Top             =   120
         Width           =   195
      End
      Begin VB.Image Image_Play 
         Height          =   255
         Left            =   2250
         Picture         =   "Form_Main.frx":487E
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image Image8 
         Height          =   150
         Left            =   2950
         Picture         =   "Form_Main.frx":4B68
         Top             =   420
         Width           =   495
      End
      Begin VB.Image Image7 
         Height          =   165
         Left            =   2520
         Picture         =   "Form_Main.frx":4F92
         Top             =   10
         Width           =   390
      End
      Begin VB.Image Image6 
         Height          =   165
         Left            =   2950
         Picture         =   "Form_Main.frx":5344
         Top             =   240
         Width           =   330
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   2280
         Picture         =   "Form_Main.frx":5672
         Top             =   120
         Width           =   150
      End
      Begin VB.Image Image4 
         Height          =   150
         Left            =   1850
         Picture         =   "Form_Main.frx":58B4
         Top             =   10
         Width           =   315
      End
      Begin VB.Image Image3 
         Height          =   225
         Left            =   1440
         Picture         =   "Form_Main.frx":5B76
         Top             =   240
         Width           =   120
      End
      Begin VB.Image Image2 
         Height          =   150
         Left            =   1120
         Picture         =   "Form_Main.frx":5D20
         Top             =   15
         Width           =   300
      End
      Begin VB.Image Image1 
         Height          =   165
         Left            =   120
         Picture         =   "Form_Main.frx":5FBA
         Top             =   10
         Width           =   525
      End
   End
   Begin VB.CommandButton Button_MediaPlayer 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "Form_Main.frx":64A0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Button_MediaPlayer 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   1
      Left            =   660
      Picture         =   "Form_Main.frx":6A6A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Button_MediaPlayer 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   4
      Left            =   2160
      Picture         =   "Form_Main.frx":6FF4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox Picturebox_PanelBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   4665
      TabIndex        =   2
      Top             =   1740
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
         BackPicture     =   "Form_Main.frx":75BE
         BarPicture      =   "Form_Main.frx":75DA
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
Public Sub AddToPlaylist(file As String)
Dim MediaDuration As String
Dim lstItem As ListItem
Dim Files
Dim FilesToAddCount As Integer
Dim i As Integer
Dim FirstIndex As Integer
Dim AddingMoreThanOne As Boolean
Dim FileNotFoundCount As Integer

On Error GoTo ErrorHandler
Files = Split(file, vbNewLine)
FilesToAddCount = UBound(Files)

If FilesToAddCount > 0 Then: FirstIndex = 0
If FilesToAddCount > 1 Then: AddingMoreThanOne = True

'Add files to the playlist
For i = FirstIndex To FilesToAddCount
    'Get current file to process
    file = Files(i)
    
    Playlist.Add file
Next

'Restore default
Screen.MousePointer = vbDefault

Exit Sub
ErrorHandler:
Select Case Err.Number
    Case 0
    Case 35602
    Case Else: Debug.Print Err.Number & " - " & Err.Description
End Select
End Sub

Private Sub Button_MediaPlayer_Click(index As Integer)
Select Case index
    Case 0: AudiostationMP3Player.PreviousTrack
    Case 1: AudiostationMP3Player.StopPlay
    Case 2
        If Playlist.Count = 0 Then
            MenuItem_File_PlayFile_Click
        Else
            AudiostationMP3Player.StartPlay
        End If

    Case 3: AudiostationMP3Player.Pause
    Case 4: AudiostationMP3Player.NextTrack
End Select
End Sub

Private Sub Form_Load()
ChDrive App.path
ChDir App.path

' Check the correct BASS was loaded
If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
    MsgBox "An incorrect version of BASS.DLL was loaded", vbCritical
    End
End If

' Initialize BASS
If (BASS_Init(-1, 44100, 0, Me.hwnd, 0) = 0) Then
    MsgBox es & vbCrLf & vbCrLf & "error code: " & BASS_ErrorGetCode, vbExclamation, "Error"
    End
End If

Call BASS_SetConfig(BASS_CONFIG_NET_PLAYLIST, 1) ' enable playlist processing
Call BASS_SetConfig(BASS_CONFIG_NET_PREBUF, 0) ' minimize automatic pre-buffering, so we can do it (and display it) instead

Picturebox_Khz.PaintPicture Imagelist_Digits.ListImages(1).Picture, 0, 0
Picturebox_Khz.PaintPicture Imagelist_Digits.ListImages(2).Picture, 120, 0
Picturebox_Khz.PaintPicture Imagelist_Digits.ListImages(3).Picture, 240, 0

Picturebox_KBit.PaintPicture Imagelist_Digits.ListImages(1).Picture, 0, 0

AudiostationMP3Player.Init
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call BASS_ChannelFree(chan)
Call BASS_Free
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
    .Filter = ComboxboxToCommondialogFilter
    .flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
    .ShowOpen

    If .filename <> vbNullString Then
        Files = Extensions.CommondialogFilesToList(.filename)
        
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

Private Sub MenuItem_Help_About_Click()
Form_About.Show vbModal, Me
End Sub

Private Sub Timer_Main_Timer()
Dim length, Pos As Long
Dim Totaltime, Elapsedtime, Remainingtime  As Double
Dim TimeSerial As String

If PlayState = Playing Then
    length = BASS_ChannelGetLength(chan, BASS_POS_BYTE)
    Pos = BASS_ChannelGetPosition(chan, BASS_POS_BYTE)
    Totaltime = BASS_ChannelBytes2Seconds(chan, length)
    Elapsedtime = BASS_ChannelBytes2Seconds(chan, Pos)
    Remainingtime = Totaltime - Elapsedtime
    
    If Not Totaltime <= 0 Then
        Progressbar_CurrentPos.max = Totaltime
        Progressbar_CurrentPos.value = Elapsedtime
    End If
    
    TimeSerial = Extensions.SecondsToTimeSerial(Elapsedtime, SmallTimeSerial)
    
    SegmentDisplay_Minutes.value = Extensions.Explode(TimeSerial, ":", 0)
    SegmentDisplay_Seconds.value = Extensions.Explode(TimeSerial, ":", 1)
    
    Image_Pause.Visible = False
    Image_Play.Visible = True
    
ElseIf PlayState = MediaEnded Or Stopped Then

    Image_Pause.Visible = False
    Image_Play.Visible = False
    
ElseIf PlayState = Paused Then
    
    Image_Play.Visible = False
    Image_Pause.Visible = True

End If
           
SegmentDisplay_TrackCount.value = AudiostationMP3Player.TrackNr
End Sub
