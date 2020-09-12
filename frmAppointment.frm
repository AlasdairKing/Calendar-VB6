VERSION 5.00
Begin VB.Form frmAppointment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Appointment"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAppointment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   8880
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   8880
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtAppointment 
      Height          =   375
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   $"frmAppointment.frx":151A
      Top             =   1080
      Width           =   8655
   End
   Begin VB.ComboBox cboDate 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Text            =   "Enter the date here (press down arrow)"
      Top             =   360
      Width           =   8655
   End
   Begin VB.Label lblAppointment 
      AutoSize        =   -1  'True
      Caption         =   "&Appointment"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Date"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "frmAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Calendar
'Copyright Alasdair King, 2010, http://www.alasdairking.me.uk
'Released under the GNU Public Licence, Version 3.

Option Explicit

Private Const JUMP_FORWARDS As Integer = 1
Private Const JUMP_BACKWARDS As Integer = -1
Public gAppointmentBeingEdited As Long

Private Sub cboDate_Click()
    On Error Resume Next
    cmdOK.Enabled = (Len(txtAppointment.text) > 0) And (cboDate.ListIndex <> -1 Or (cboDate.text <> "" And cboDate.text <> DEFAULT_DATE_TEXT)) And txtAppointment.text <> DEFAULT_APPOINTMENT_TEXT
End Sub

Private Sub cboDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        
    If KeyCode = vbKeyPageDown And (Shift And vbCtrlMask) Then
        If cboDate.text = DEFAULT_DATE_TEXT Then
            cboDate.ListIndex = cboDate.ListIndex + 30
        Else
            Call JumpAMonth(JUMP_FORWARDS)
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeyPageUp And (Shift And vbCtrlMask) Then
        If cboDate.text = DEFAULT_DATE_TEXT Then
            Call Beep
        Else
            Call JumpAMonth(JUMP_BACKWARDS)
        End If
        KeyCode = 0
    End If
End Sub

Private Sub JumpAMonth(direction As Integer)
    On Error Resume Next
    Dim currentDate As Date
    Dim targetDay As Integer
    Dim targetMonth As String
    Dim targetYear As Integer
    Dim jump As Integer
    Dim d As String
    Dim i As Long
    Dim quitWhile As Boolean
    targetDay = GetDay(cboDate.text)
    targetMonth = GetMonth(cboDate.text)
    If direction = JUMP_FORWARDS Then
        If targetMonth = "Jan" Or targetMonth = "Mar" Or targetMonth = "May" Or targetMonth = "Jul" Or targetMonth = "Aug" Or targetMonth = "Oct" Or targetMonth = "Dec" Then
            jump = 31
        ElseIf targetMonth = "Apr" Or targetMonth = "Jun" Or targetMonth = "Sep" Or targetMonth = "Nov" Then
            jump = 30
        Else
            'February, special case.
            targetYear = GetYear(cboDate.text)
            If IsLeapYear(targetYear) Then
                jump = 29
            Else
                jump = 28
            End If
        End If
    ElseIf direction = JUMP_BACKWARDS Then
        If targetMonth = "Jan" Or targetMonth = "Feb" Or targetMonth = "Apr" Or targetMonth = "Jun" Or targetMonth = "Aug" Or targetMonth = "Sep" Or targetMonth = "Nov" Then
            jump = 31
        ElseIf targetMonth = "May" Or targetMonth = "Jul" Or targetMonth = "Oct" Or targetMonth = "Dec" Then
            jump = 30
        Else
            'March, special case.
            targetYear = GetYear(cboDate.text)
            If IsLeapYear(targetYear) Then
                jump = 29
            Else
                jump = 28
            End If
        End If
    End If
    If direction = JUMP_FORWARDS Then
        If cboDate.ListIndex < cboDate.ListCount - jump Then
            cboDate.ListIndex = cboDate.ListIndex + jump
        Else
            Call Beep
        End If
    ElseIf direction = JUMP_BACKWARDS Then
        If cboDate.ListIndex > jump Then
            cboDate.ListIndex = cboDate.ListIndex - jump
        Else
            Call Beep
        End If
    End If
End Sub

Private Function GetDay(dateAsString As String) As Integer
    On Error Resume Next
    GetDay = CInt(Val(Right(dateAsString, Len(dateAsString) - InStr(1, dateAsString, " "))))
End Function
    
Private Function GetMonth(dateAsString As String) As String
    On Error Resume Next
    Dim s As String
    s = Right(dateAsString, Len(dateAsString) - InStr(1, dateAsString, " "))
    s = Right(s, Len(s) - InStr(1, s, " "))
    GetMonth = Left(s, InStr(1, s, " "))
    GetMonth = Trim(GetMonth)
End Function

Private Function GetYear(dateAsString As String) As Integer
    On Error Resume Next
    Dim s As String
    s = Right(dateAsString, Len(dateAsString) - InStrRev(dateAsString, " "))
    GetYear = CInt(Val(s))
End Function

Private Function IsLeapYear(year As Integer) As Boolean
    On Error Resume Next
    If year Mod 400 = 0 Then
        IsLeapYear = True
    ElseIf year Mod 100 = 0 Then
        IsLeapYear = False
    ElseIf year Mod 4 = 0 Then
        IsLeapYear = True
    Else
        IsLeapYear = False
    End If
End Function

Private Sub cboDate_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    cmdOK.Enabled = (Len(txtAppointment.text) > 0) And (cboDate.ListIndex <> -1 Or (cboDate.text <> "" And cboDate.text <> DEFAULT_DATE_TEXT)) And txtAppointment.text <> DEFAULT_APPOINTMENT_TEXT
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    Call Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    Dim d As String
    Dim clashCount As Long
    Dim clashReport As String
    
    d = cboDate.text
    d = Right(d, Len(d) - InStr(1, d, " "))
    If gEditing Then
        Call gCalendar.UpdateAppointment(gAppointmentBeingEdited + 1, CDate(d), txtAppointment.text)
    Else
        Call gCalendar.AddAppointment(txtAppointment.text, CDate(d), clashCount, clashReport)
        If clashCount > 0 Then
            MsgBox GetText("You have") & " " & clashCount & " " & GetText("appointments on") & cboDate.text & ":" & vbNewLine & clashReport, vbExclamation
        End If
    End If
    Call Me.Hide
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Dim i As Long
    
    Call cboDate.Clear
    For i = 1 To 732
        Call cboDate.AddItem(Format(Now + i, "dddd dd mmm yyyy"))
    Next i
    If gEditing Then
        'Editing existing appointment
        txtAppointment.text = gCalendar.appointments(gAppointmentBeingEdited + 1).text
        cboDate.text = Format(gCalendar.appointments(gAppointmentBeingEdited + 1).day, "dddd dd mmm yyyy")
        Call txtAppointment.SetFocus
    Else
        'New appointment
        cboDate.text = DEFAULT_DATE_TEXT
        txtAppointment.text = DEFAULT_APPOINTMENT_TEXT
        Call cboDate.SetFocus
    End If
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call modI18N.ApplyUILanguageToThisForm(Me)
End Sub

Private Sub txtAppointment_Change()
    On Error Resume Next
    cmdOK.Enabled = (Len(txtAppointment.text) > 0) And (cboDate.ListIndex <> -1 Or (cboDate.text <> "" And cboDate.text <> DEFAULT_DATE_TEXT)) And txtAppointment.text <> DEFAULT_APPOINTMENT_TEXT
End Sub

Private Sub txtAppointment_GotFocus()
    On Error Resume Next
    txtAppointment.SelStart = 0
    txtAppointment.SelLength = Len(txtAppointment.text)
End Sub
