VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Returns - Total duration from seconds as Time format: HH:MM:SS
Public Function GetTime(ByVal seconds As Long) As String
    Dim hour As Single, min As Single, sec As Single

    hour = seconds / 60 / 60
    sec = seconds Mod 60
    min = (hour - Int(hour)) * 60

    GetTime = Format(Int(hour), "00") & ":" & Format(Int(min), "00") & ":" & Format(Int(sec), "00")
End Function

' Returns - Playing position in seconds
Public Function GetPlayingPos(ByVal handle As Long) As Single
    GetPlayingPos = BASS_ChannelBytes2Seconds(handle, BASS_ChannelGetPosition(handle, BASS_POS_BYTE))
End Function

' Returns - Total duration in seconds
Public Function GetDuration(ByVal handle As Long) As Single
    GetDuration = BASS_ChannelBytes2Seconds(handle, BASS_ChannelGetLength(handle, BASS_POS_BYTE))
End Function
' Returns - Kilo Bits Per Second
Public Function GetBitsPerSecond(ByVal handle As Long, ByVal FileLength As Long) As Long
    GetBitsPerSecond = CInt(((FileLength * 8) / GetDuration(handle)) / 1000)
End Function

' Returns - Sample Rate [Frequency]
Public Function GetFrequency(ByVal handle As Long) As Long
    Dim freq As Single
    Call BASS_ChannelGetAttribute(handle, BASS_ATTRIB_FREQ, freq)
    GetFrequency = freq
End Function

' Returns - DirectX version
Public Function GetDXver() As Byte
    Dim bi As BASS_INFO
    Call BASS_GetInfo(bi)
    GetDXver = bi.dsver
End Function
