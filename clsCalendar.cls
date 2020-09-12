VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Calendar
'Copyright Alasdair King, 2010, http://www.alasdairking.me.uk
'Released under the GNU Public Licence, Version 3.

Option Explicit

Private mAppointments As Collection
Private mLastEdited As Long

Public Property Get LastEdited() As Long
    On Error Resume Next
    LastEdited = mLastEdited
End Property

Public Property Get appointments() As Collection
    On Error Resume Next
    Set appointments = mAppointments
End Property

Public Sub AddAppointment(text As String, day As Date, clashCount As Long, clashReport As String)
    On Error Resume Next
    Dim i As Long
    Dim newApp As clsAppointment
    Dim added As Boolean
    Dim clash As Boolean
    Dim addedThisDay As Collection
    
    Set newApp = New clsAppointment
    newApp.day = day
    newApp.text = text
    'Work out where to add
    i = 1
    added = False
    While Not added And i <= mAppointments.Count
        If mAppointments.Item(i).day > newApp.day Then
            Call mAppointments.Add(newApp, , i)
            added = True
            mLastEdited = i
        End If
        i = i + 1
    Wend
    If Not added Then
        Call mAppointments.Add(newApp) 'Add to end
        mLastEdited = mAppointments.Count
    End If
    'Is there a clash of days?
    Set addedThisDay = New Collection
    For i = 1 To mAppointments.Count
        If mAppointments.Item(i).day = newApp.day Then
            Call addedThisDay.Add(mAppointments.Item(i))
        End If
    Next i
    If addedThisDay.Count > 1 Then
        'Clash of events!
        clashCount = addedThisDay.Count
        For i = 1 To addedThisDay.Count
            clashReport = clashReport & addedThisDay.Item(i).text & vbNewLine
        Next i
    End If
    Call SaveState
End Sub

Private Sub Class_Initialize()
    On Error Resume Next
    Dim Calendar As DOMDocument30
    Dim appointment As IXMLDOMNode
    Dim newAppointment As clsAppointment
    Dim fso As New FileSystemObject
    
    Set mAppointments = New Collection
    Set Calendar = New DOMDocument30
    Calendar.async = False
    If fso.FileExists(modPath.settingsPath & "\calendar.xml") Then
        Call Calendar.Load(modPath.settingsPath & "\calendar.xml")
    Else
        Call Calendar.loadXML("<?xml version=""1.0"" encoding=""UTF-16""?><calendar />")
    End If
    If Calendar.parseError.errorCode <> 0 Then
        'Failed to load old calendar, whoops. Maybe problem with user putting "<" characters
        'into it. Give up and reload.
        Call Calendar.loadXML("<?xml version=""1.0"" encoding=""UTF-16""?><calendar />")
    End If
    For Each appointment In Calendar.documentElement.selectNodes("appointment")
        If Int(CDate(appointment.selectSingleNode("day").text)) < Int(Now) Then
            'appointment in the past, discard.
        Else
            Set newAppointment = New clsAppointment
            newAppointment.day = CDate(appointment.selectSingleNode("day").text)
            newAppointment.text = appointment.selectSingleNode("text").text
            Call mAppointments.Add(newAppointment)
        End If
    Next appointment
    If mAppointments.Count = 0 Then mLastEdited = -1 Else mLastEdited = 1
End Sub

Private Sub SaveState()
    On Error Resume Next
    Dim Calendar As DOMDocument30
    Dim appointment As clsAppointment
    Dim s As String
    
    Set Calendar = New DOMDocument30
    Calendar.async = False
    s = "<?xml version=""1.0"" encoding=""UTF-16""?><calendar>" & vbNewLine
    For Each appointment In mAppointments
        s = s & vbTab & appointment.xml & vbNewLine
    Next appointment
    s = s & "</calendar>"
    Call Calendar.loadXML(s)
    Call Calendar.save(modPath.settingsPath & "\calendar.xml")
End Sub

Public Sub RemoveAppointment(index As Long)
    On Error Resume Next
    If index >= 1 And index <= mAppointments.Count Then
        Call mAppointments.Remove(index)
        mLastEdited = index
        Call SaveState
    End If
End Sub

Public Sub UpdateAppointment(index As Long, Optional newDate As Date = CDate(0), Optional newText As String)
    On Error Resume Next
    Dim appo As clsAppointment
    
    If index >= 1 And index <= mAppointments.Count Then
        Set appo = mAppointments.Item(index)
        If newDate <> CDate(0) Then appo.day = newDate
        If newText <> "" Then appo.text = newText
        Call mAppointments.Remove(index)
        If index > mAppointments.Count Then
            Call mAppointments.Add(appo)
        Else
            Call mAppointments.Add(appo, , index)
        End If
        mLastEdited = index
        Call SaveState
    End If
End Sub

Public Function CalendarAsHTML() As String
    On Error Resume Next
    Dim appo As clsAppointment
    
    CalendarAsHTML = "<html><meta http-equiv=""content-type"" content=""text/html; charset=dbcs""><title>" & GetText("frmMain.Caption") & "</title>"
    CalendarAsHTML = CalendarAsHTML & "<style type=""text/css"">* {font-face:arial; } body { padding:5%; } h2 { font-size:16pt; } p { padding-left:2em; font-size:14pt; font-weight:bold; } h1 { font-size:18pt; }</style>"
    CalendarAsHTML = CalendarAsHTML & "</head><body>"
    CalendarAsHTML = CalendarAsHTML & "<h1>" & GetText("frmMain.Caption") & "</h1>"
    If mAppointments.Count = 0 Then
        CalendarAsHTML = CalendarAsHTML & "<p>" & NO_APPOINTMENTS_TEXT & "</p>"
    Else
        For Each appo In mAppointments
            CalendarAsHTML = CalendarAsHTML & "<p>" & appo.DisplayHTML & "</p>"
        Next appo
    End If
    CalendarAsHTML = CalendarAsHTML & "</html>"
End Function

Public Function CalendarAsText() As String
    On Error Resume Next
    Dim appo As clsAppointment
    
    CalendarAsText = GetText("frmMain.Caption") & vbNewLine & vbNewLine
    If mAppointments.Count = 0 Then
        CalendarAsText = CalendarAsText & NO_APPOINTMENTS_TEXT
    Else
        For Each appo In mAppointments
            CalendarAsText = CalendarAsText & appo.Display & vbNewLine
        Next appo
    End If
End Function

