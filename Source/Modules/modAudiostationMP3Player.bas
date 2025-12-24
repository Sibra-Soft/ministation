Attribute VB_Name = "AudiostationMP3Player"
Public PlayState As enumPlayStates
Public PlayMode As enumPlayMode
Public PlayStateMediaMode As enumMediaMode

Public Playlist As New Collection

Public ShowElapsedTime As Boolean
Public MediaFilename As String
Public TrackNr As Integer
Public Sub Init()
PlayState = Stopped
PlayMode = NormalPlay
AudiostationMP3Player.ShowElapsedTime = True
End Sub
Public Sub Rewind()
Dim Pos As Long

Pos = BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(chan, BASS_POS_BYTE))

Call BASS_ChannelSetPosition(chan, BASS_ChannelSeconds2Bytes(chan, Pos + 5), BASS_POS_BYTE)
End Sub
Public Sub Forward()
Dim Pos As Long

Pos = BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(chan, BASS_POS_BYTE))

Call BASS_ChannelSetPosition(chan, BASS_ChannelSeconds2Bytes(chan, Pos - 5), BASS_POS_BYTE)
End Sub
Public Sub Pause()
Call BASS_ChannelPause(chan)
PlayState = Paused
End Sub
Public Sub StartPlay()
PlayStateMediaMode = MP3MediaMode

If Playlist.Count = 0 Then: Exit Sub

If PlayState = Paused Then
    Call BASS_ChannelPlay(chan, False)
Else
    If TrackNr = 0 Then: TrackNr = 1
   
    Call BASS_StreamFree(chan)
    Call BASS_MusicFree(chan)
    
    MediaFilename = Playlist(TrackNr)

    chan = BASS_StreamCreateFile(BASSFALSE, StrPtr(MediaFilename), 0, 0, BASS_STREAM_AUTOFREE)
    If chan = 0 Then chan = BASS_MusicLoad(BASSFALSE, MediaFilename, 0, 0, BASS_STREAM_AUTOFREE, 1)

    Call BASS_ChannelPlay(chan, True)
End If

PlayState = Playing
End Sub
Public Sub StopPlay()
Call BASS_ChannelStop(chan)
PlayState = Stopped
End Sub
Public Sub NextTrack(Optional TrackNumber As Integer, Optional Force = False)
Dim MediaFilename As String

If Playlist.Count = 0 Then: Exit Sub
If TrackNr >= Playlist.Count - 1 And PlayMode = NormalPlay Then: Exit Sub

If TrackNumber > 0 Then
    'Track number is set by parameter
    AudiostationMP3Player.TrackNr = TrackNumber
Else
    Dim NextTrackNumber As Integer
    Randomize
    
    If PlayMode = RepeatTrack Then NextTrackNumber = TrackNr: GoTo DoNext
    If Force Then NextTrackNumber = AudiostationMP3Player.TrackNr + 1: GoTo DoNext
    
    NextTrackNumber = AudiostationMP3Player.TrackNr + 1
    
DoNext:
    'Auto select track number
    AudiostationMP3Player.TrackNr = NextTrackNumber
End If

AudiostationMP3Player.TrackNr = TrackNr
MediaFilename = Playlist.Item(TrackNr)

CurrentMediaFilename = MediaFilename

Call StartPlay
End Sub
Public Sub PreviousTrack()
Dim MediaFilename As String

If Playlist.Count = 0 Or TrackNr = 1 Then: Exit Sub

AudiostationMP3Player.TrackNr = TrackNr - 1

MediaFilename = Playlist.Item(TrackNr)
CurrentMediaFilename = MediaFilename

Call StartPlay
End Sub
