VERSION 5.00
Begin VB.Form frmReminder 
   Caption         =   "Calendar reminders"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReminder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtReminders 
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblReminders 
      AutoSize        =   -1  'True
      Caption         =   "&Reminders"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Calendar
'Copyright Alasdair King, 2010, http://www.alasdairking.me.uk
'Released under the GNU Public Licence, Version 3.

Option Explicit

Private voice As SpVoice

Public Sub SpeakReminder()
    Call voice.Speak(CALENDAR_REMINDERS & " " & txtReminders.text, SVSFlagsAsync)
End Sub

Private Sub cmdClose_Click()
    On Error Resume Next
    Call Unload(Me)
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Call voice.Speak("", SVSFPurgeBeforeSpeak)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set voice = New SpVoice
    Call modI18N.ApplyUILanguageToThisForm(Me)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lblReminders.Top = 90
    lblReminders.Left = 90
    txtReminders.Top = lblReminders.Top + lblReminders.Height + 90
    txtReminders.Left = 90
    txtReminders.Width = Me.ScaleWidth - 180
    txtReminders.Height = Me.ScaleHeight - cmdClose.Height - 180 - txtReminders.Top
    cmdClose.Top = Me.ScaleHeight - cmdClose.Height - 90
    cmdClose.Left = Me.ScaleWidth / 2 - cmdClose.Width / 2
End Sub

Private Sub txtReminders_GotFocus()
    On Error Resume Next
    txtReminders.SelStart = 0
    txtReminders.SelLength = Len(txtReminders.text)
End Sub
