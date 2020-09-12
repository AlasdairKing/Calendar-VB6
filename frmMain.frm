VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Calendar"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   825
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
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlgExport 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "rtf"
      DialogTitle     =   "Export Calendar"
      FileName        =   "Calendar.rtf"
      Filter          =   "RTF Files (*.rtf)|*.rtf"
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Height          =   975
      Left            =   2760
      Picture         =   "frmMain.frx":151A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "D&elete"
      Enabled         =   0   'False
      Height          =   975
      Left            =   1440
      Picture         =   "frmMain.frx":19C9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer tmrShow 
      Interval        =   139
      Left            =   1800
      Top             =   1320
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "A&dd"
      Height          =   975
      Left            =   120
      Picture         =   "frmMain.frx":1EAD
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox lstAppointments 
      Height          =   1260
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label lblAppointments 
      AutoSize        =   -1  'True
      Caption         =   "&Appointments"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1170
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExportcalendard 
         Caption         =   "&Export calendar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileHide 
         Caption         =   "&Hide"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAppointments 
      Caption         =   "A&ppointments"
      Begin VB.Menu mnuAppointmentsAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuAppointmentsDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAppointmentsChange 
         Caption         =   "&Change"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAppointmentsFind 
         Caption         =   "&Find"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuAppointmentsFindnext 
         Caption         =   "Find &Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsRuncalendareverytimeyoulogin 
         Caption         =   "Run Calendar every time you log in"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpManual 
         Caption         =   "&Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Calendar
'Copyright Alasdair King, 2010, http://www.alasdairking.me.uk
'Released under the GNU Public Licence, Version 3.

Option Explicit

'1.0.3
'   31 March 2009 Release for beta testing.
'1.0.4
'   Fixed XP Style bug.
'1.5.0
'   Added Task Schedule control of running Calendar on logon.
'   Added buttons and images.
'   Enabled the editing of existing tasks.
'2.0.1
'   Scrapped Task Schedule for creating a shortcut because acted oddly on XP. 22 Sep 2009.
'2.1.0
'   Added a Find Appointments function so you can search your appointments.
'   Changed reminders to operate for any event in the next week, not just exactly one week away.
'   Added alert if clashing appointments are identified.
'   Fixed Appointments form not showing in task bar.
'   Fixed program ALWAYS writing/deleting Startup folder shortcut because it chokes anti-startup
'       programs.
'2.1.1
'   Fixed invalid XML characters (<, &) breaking appointments loading.
'   Fixed wrong icon on Add form.
'   Fixed I18N not being applied.
'2.2.0 16 August 2010
'   Fixed not saving details when minimized to system tray/notification area.
'   Now does US-style calendar if you have "en-us" selected as your language.
'   Added ability to export calendar to 14pt RTF file.


Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                 (ByVal lpBuffer As String, nSize As Long) As Long

Private mFindText As String

Private Sub cmdAdd_Click()
    On Error Resume Next
    Call frmAppointment.Show(vbModal, Me)
    Call Display
End Sub

Private Sub cmdChange_Click()
    On Error Resume Next
    Call lstAppointments_KeyPress(vbKeyReturn)
    Call lstAppointments.SetFocus
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    Call lstAppointments_KeyDown(vbKeyDelete, 0)
    Call lstAppointments.SetFocus
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call modI18N.ApplyUILanguageToThisForm(Me)
    Debug.Print modI18N.GetLanguage
    mnuOptionsRuncalendareverytimeyoulogin.Checked = modPath.GetSettingIni("Calendar", "Options", "RunOnLogon", "true")
    mnuAppointmentsDelete.Caption = mnuAppointmentsDelete.Caption & vbTab & GetText("Delete")
    mnuAppointmentsChange.Caption = mnuAppointmentsChange.Caption & vbTab & GetText("Return")
    Call SetSchedule
    Call Display
    lstAppointments.ListIndex = 0
End Sub

Private Sub SetSchedule(Optional forceShortcutWrite As Boolean)
    On Error Resume Next
    Dim name As String
    Dim nameLength As Long
    Dim result As Long

    name = Space$(255)
    nameLength = Len(name)
    result = GetUserName(name, nameLength)
    If result <> 0 Then
        'Success
        name = Left(name, nameLength - 1)
    Else
        Debug.Print "WARNING: Failed to get username in SetSchedule: " & Err.LastDllError
        name = ""
    End If

    If modPath.runningLocal Then
        'Nope, we're running off a USB stick. Can't schedule.
        mnuOptionsRuncalendareverytimeyoulogin.Enabled = False
    Else
        If mnuOptionsRuncalendareverytimeyoulogin.Checked Then
            'Run when user logs on.
            If Dir(modPath.GetSpecialFolderPath(modPath.CSIDL_STARTUP) & "\" & app.Title & ".lnk") = "" Or forceShortcutWrite Then
                'shortcut missing or force instruction to change.
                Call modCreateShortcut.CreateShortCut(app.Path & "\" & app.EXEName & ".exe", modPath.GetSpecialFolderPath(modPath.CSIDL_STARTUP), app.Title, "-showreminders", app.Path)
            End If
        Else
            If Dir(modPath.GetSpecialFolderPath(modPath.CSIDL_STARTUP) & "\" & app.Title & ".lnk") Then
                'Stop running when user logs on.
                Call Kill(modPath.GetSpecialFolderPath(modPath.CSIDL_STARTUP) & "\" & app.Title & ".lnk")
            End If
        End If
        
        
        
''' DEV This is the abandoned AT task scheduling code. Didn't work on XP when I tried it.
'''            'set to run when user logs on
'''            Call Shell("schtasks /delete /tn AccessibleCalendar /F")
'''            DoEvents
'''            If name = "" Then
'''                'Failed to get username
'''                mnuOptionsRuncalendareverytimeyoulogin.Enabled = False
'''            Else
'''                'Dev: Doesn't work with /U flag, have to run for everyone.
'''                'Dev: /F flag doesn't exist, I think - I'm having problems getting this working on XP.
'''                Dim s As String
'''                s = "schtasks /create /tn AccessibleCalendar /tr """ & App.Path & "\" & App.EXEName & ".exe -showreminders"" /sc ONLOGON"
'''                MsgBox s
'''                Call Shell(s) '/U " & name)
'''            End If
'''        Else
'''            'Stop running when user logs on
'''            Call Shell("schtasks /delete /tn AccessibleCalendar /F")
    End If
End Sub

Private Sub Display()
    On Error Resume Next
    Dim i As Long
    Dim a As clsAppointment
    
    i = lstAppointments.ListIndex
    Call lstAppointments.Clear
    For Each a In gCalendar.appointments
        Call lstAppointments.AddItem(a.Display)
    Next a
    If lstAppointments.ListCount = 0 Then
        Call lstAppointments.AddItem(NO_APPOINTMENTS_TEXT)
        mnuAppointmentsFind.Enabled = False
        mnuAppointmentsFindnext.Enabled = False
    Else
        mnuAppointmentsFind.Enabled = True
        mnuAppointmentsFindnext.Enabled = True
    End If
    If gCalendar.LastEdited > 0 Then
        lstAppointments.ListIndex = gCalendar.LastEdited - 1
        If Me.Visible Then Call lstAppointments.SetFocus
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        Me.lblAppointments.Top = 90
        lblAppointments.Left = 90
        lstAppointments.Top = lblAppointments.Top + lblAppointments.Height + 45
        cmdAdd.Left = 90
        cmdAdd.Top = Me.ScaleHeight - cmdAdd.Height - 90
        lstAppointments.Left = 90
        lstAppointments.Height = cmdAdd.Top - 90 - lstAppointments.Top
        lstAppointments.Width = Me.ScaleWidth - 90 - lstAppointments.Left
        cmdDelete.Top = cmdAdd.Top
        cmdChange.Top = cmdAdd.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim f As Form
    If mnuOptionsRuncalendareverytimeyoulogin.Enabled Then
        Call modPath.SaveSettingIni("Calendar", "Options", "RunOnLogon", CStr(mnuOptionsRuncalendareverytimeyoulogin.Checked))
    End If
    For Each f In Forms
        If f.name <> Me.name Then
            Call Unload(f)
        End If
    Next f
End Sub

Private Sub lstAppointments_Click()
    On Error Resume Next
    cmdDelete.Enabled = (lstAppointments.text <> NO_APPOINTMENTS_TEXT)
    cmdChange.Enabled = cmdDelete.Enabled
    mnuAppointmentsDelete.Enabled = cmdDelete.Enabled
    mnuAppointmentsChange.Enabled = cmdDelete.Enabled
End Sub

Private Sub lstAppointments_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim i As Long
    
    If KeyCode = vbKeyDelete Then
        If lstAppointments.ListIndex > -1 Then
            If lstAppointments.text = NO_APPOINTMENTS_TEXT Then
                Call Beep
            Else
                i = lstAppointments.ListIndex
                Call gCalendar.RemoveAppointment(lstAppointments.ListIndex + 1)
                Call lstAppointments.RemoveItem(lstAppointments.ListIndex)
                If lstAppointments.ListCount = 0 Then Call lstAppointments.AddItem(NO_APPOINTMENTS_TEXT)
                If i > lstAppointments.ListCount - 1 Then i = lstAppointments.ListCount - 1
                lstAppointments.ListIndex = i
            End If
        End If
    End If
End Sub

Private Sub lstAppointments_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        If lstAppointments.text = NO_APPOINTMENTS_TEXT Then
            Call Beep
        Else
            frmAppointment.gAppointmentBeingEdited = lstAppointments.ListIndex
            gEditing = True
            Call frmAppointment.Show(vbModal, Me)
            gEditing = False
            Call Display
        End If
    End If
End Sub

Private Sub mnuAppointmentsAdd_Click()
    On Error Resume Next
    Call cmdAdd_Click
End Sub

Private Sub mnuAppointmentsChange_Click()
    On Error Resume Next
    Call cmdChange_Click
End Sub

Private Sub mnuAppointmentsDelete_Click()
    On Error Resume Next
    Call cmdDelete_Click
End Sub

Private Sub mnuAppointmentsFind_Click()
    On Error Resume Next
    Dim textToFind As String
    
    textToFind = InputBox(GetText("Enter text to find in your appointments:"), GetText("Calendar"), mFindText)
    If textToFind <> "" Then
        mFindText = textToFind
        Call FindText(textToFind)
    End If
End Sub

Private Sub FindText(Optional text As String)
    On Error Resume Next
    Dim i As Long
    Dim found As Boolean
    
    If text = "" Then text = mFindText
    If text = "" Then
        Call Beep
    Else
        'Some text to find!
        For i = lstAppointments.ListIndex + 1 To lstAppointments.ListCount - 1
            If InStr(1, lstAppointments.List(i), text, vbTextCompare) > 0 Then
                'Found!
                found = True
                Exit For
            End If
        Next i
        If Not found Then
            For i = 0 To lstAppointments.ListIndex
                If InStr(1, lstAppointments.List(i), text, vbTextCompare) > 0 Then
                    'Found!
                    found = True
                    Exit For
                End If
            Next i
        End If
        If found Then
            'Found somewhere.
            lstAppointments.ListIndex = i
            Call lstAppointments.SetFocus
        Else
            Call MsgBox(GetText("Not found:") & " " & """" & text & """", vbExclamation)
        End If
    End If
End Sub

Private Sub mnuAppointmentsFindnext_Click()
    On Error Resume Next
    Call FindText
End Sub

Private Sub mnuFileExit_Click()
    On Error Resume Next
    Call Unload(Me)
End Sub

Private Sub mnuFileExportcalendard_Click()
    On Error Resume Next
    Dim s As String
    Dim fso As Scripting.FileSystemObject
    Dim ts As TextStream
    Dim doSaveWithFSO As Boolean
    
    'Saves your Calendar as an RTF or HTML or Text file.
    cdlgExport.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
    cdlgExport.InitDir = modPath.GetSpecialFolderPath(modPath.CSIDL_DESKTOP)
    cdlgExport.CancelError = True
    cdlgExport.DialogTitle = GetText("Export Calendar")
    On Error GoTo quit:
    Call cdlgExport.ShowSave
    On Error Resume Next
    If cdlgExport.filename <> "" Then
        'OK, let's make it!
        Set fso = New Scripting.FileSystemObject
        Select Case cdlgExport.FilterIndex
            Case 1 ' RTF
                frmRTF.rtbLayout.text = ""
                frmRTF.rtbLayout.Font.size = 14
                frmRTF.rtbLayout.Font.bold = True
                frmRTF.rtbLayout.text = gCalendar.CalendarAsText
                Call frmRTF.rtbLayout.SaveFile(cdlgExport.filename)
            Case 2 ' HTML
                s = gCalendar.CalendarAsHTML
                doSaveWithFSO = True
            Case 3 ' Plain Text
                s = gCalendar.CalendarAsText
                doSaveWithFSO = True
        End Select
        If doSaveWithFSO Then
            Set ts = fso.CreateTextFile(cdlgExport.filename, True, True)
            Call ts.Write(s)
            Call ts.Close
        End If
    End If
quit:
End Sub

Private Sub mnuFileHide_Click()
    On Error Resume Next
    Call Load(frmSystemTray)
    Me.Visible = False
End Sub

Private Sub mnuHelpAbout_Click()
    On Error Resume Next
    MsgBox app.Title & vbTab & app.Major & "." & app.Minor & "." & app.Revision & vbNewLine & "Package Version" & vbTab & modVersion.GetPackageVersion & vbNewLine & "Alasdair King http://www.webbie.org.uk", vbInformation
End Sub

Private Sub mnuHelpManual_Click()
    On Error Resume Next
'    MsgBox "When the program starts you have a list of appointments with nothing in it. Press the Add button (Alt and A) to add an appointment: select the date with arrow down and tab to enter a message, then return to save the appointment. This will be added to the appointment list. When the computer starts your appointment list will be checked and any appointments that day, tomorrow, or in a week will be read out and shown until the user closes the reminder window.", vbInformation
    frmHelp.Icon = Me.Icon
    Call frmHelp.Show(vbModeless, Me)
End Sub

Private Sub mnuOptionsRuncalendareverytimeyoulogin_Click()
    On Error Resume Next
    mnuOptionsRuncalendareverytimeyoulogin.Checked = Not mnuOptionsRuncalendareverytimeyoulogin.Checked
    Call SetSchedule(True)
End Sub

Private Sub tmrShow_Timer()
    On Error Resume Next
    Static busy As Boolean
    If busy Then
    Else
        busy = True
        tmrShow.Enabled = False
        Me.Visible = True
        DoEvents
        Call Me.SetFocus
        busy = False
    End If
End Sub
