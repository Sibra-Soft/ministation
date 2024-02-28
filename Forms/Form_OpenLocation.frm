VERSION 5.00
Begin VB.Form Form_OpenLocation 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Location"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_OpenLocation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Ministation.ButtonBig Button_CancelSettings 
      Height          =   390
      Left            =   3270
      TabIndex        =   2
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   688
      Caption         =   "Cancel"
      TextAlignment   =   0
   End
   Begin Ministation.ButtonBig Button_OK 
      Height          =   390
      Left            =   2190
      TabIndex        =   3
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   688
      Caption         =   "OK"
      TextAlignment   =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Enter a Uniform Resource Location (URL) "
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox Combox_Location 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "Combox_Location"
         Top             =   480
         Width           =   5895
      End
   End
End
Attribute VB_Name = "Form_OpenLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DialogResult As VbMsgBoxResult
Public Location As String
Private Sub Button_CancelSettings_Click()
DialogResult = vbCancel
Unload Me
End Sub

Private Sub Button_OK_Click()
Location = Combox_Location.Text

If Location = vbNullString Then
    DialogResult = vbCancel
Else
    DialogResult = vbOK
End If

Unload Me
End Sub

Private Sub Form_Load()
Dim I As Integer

Combox_Location.Text = vbNullString
Combox_Location.Clear

For I = 1 To Form_Main.SavedLocations.Count
    Combox_Location.AddItem Form_Main.SavedLocations(I)
Next
End Sub
