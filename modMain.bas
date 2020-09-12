Attribute VB_Name = "modMain"
'Calendar
'Copyright Alasdair King, 2010, http://www.alasdairking.me.uk
'Released under the GNU Public Licence, Version 3.

Option Explicit

'Calendar
'1.0.0
'   First development version, send to John.
'1.0.1
'   Reminder now correctly shuts up on any key down.
'   Reminder only speaks when first started, not every time reminder gets focus.
'   Added up to two years in calendar.
'   Can enter own date.

Public gCalendar As clsCalendar
Public gEditing As Boolean

Public DEFAULT_DATE_TEXT As String
Public DEFAULT_APPOINTMENT_TEXT As String
Public NO_APPOINTMENTS_TEXT As String
Public CALENDAR_REMINDERS As String

Public Sub Main()
    On Error Resume Next
    Dim appointment As clsAppointment
    Dim text As String
    Dim shouldBeRunningOnLogonForThisUser As Boolean
    
    Call modPath.DetermineSettingsPath("WebbIE", "Calendar", "1")
    Set gCalendar = New clsCalendar
        
    shouldBeRunningOnLogonForThisUser = CBool(modPath.GetSettingIni("Calendar", "Options", "RunOnLogon", "false"))
    If LCase(Trim(Command$)) = "-showreminders" Then
        If shouldBeRunningOnLogonForThisUser Then
            CALENDAR_REMINDERS = GetText("Calendar Reminders for today.")
            For Each appointment In gCalendar.appointments
                If Int(appointment.day) = Int(Now) Then
                    text = text & GetText("Today:") & " " & appointment.text
                    If Right(text, 1) <> "." Then text = text & "."
                    text = text & vbNewLine
                ElseIf Int(appointment.day) = Int(Now) + 1 Then
                    text = text & GetText("Tomorrow:") & " " & appointment.text
                    If Right(text, 1) <> "." Then text = text & "."
                    text = text & vbNewLine
                ElseIf Int(appointment.day) <= Int(Now) + 7 Then
                    text = text & GetText("In the next week:") & " " & appointment.text
                    If Right(text, 1) <> "." Then text = text & "."
                    text = text & vbNewLine
                End If
            Next appointment
            If Len(text) > 0 Then
                Call Load(frmReminder)
                frmReminder.txtReminders.text = text
                frmReminder.Visible = True
                Call frmReminder.SpeakReminder
            End If
        Else
            'Some other user from the one who set this up to run on logon...
        End If
    ElseIf LCase(Trim(Command$)) = "-deleteshortcut" Then
        'Instruction from uninstall to delete the shortcut to this application.
        If Dir(modPath.GetSpecialFolderPath(modPath.CSIDL_STARTUP) & "\" & App.Title & ".lnk") <> "" Then
            Call Kill(modPath.GetSpecialFolderPath(modPath.CSIDL_STARTUP) & "\" & App.Title & ".lnk")
        End If
        'That's it, done.
    Else
        DEFAULT_APPOINTMENT_TEXT = GetText("Type your appointment here (e.g. ""Lunch with John"")")
        DEFAULT_DATE_TEXT = GetText("Enter the date here (press down arrow)")
        NO_APPOINTMENTS_TEXT = GetText("No appointments in your calendar yet!")
        Call Load(frmMain)
        frmMain.Visible = True
    End If
End Sub
