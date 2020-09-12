VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Calendar
'Copyright Alasdair King, 2010, http://www.alasdairking.me.uk
'Released under the GNU Public Licence, Version 3.

Option Explicit

Private mDay As Date
Private mText As String ' The text of the appointment, in non-escaped form. That is,
'   it can have "<" in it.

Public Property Get xml() As String
    On Error Resume Next
    Dim text As String
    
    text = Replace(mText, "&", "&amp;") ' Do this one first!
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, """", "&quot;")
    xml = "<appointment><day>" & mDay & "</day><text>" & text & "</text></appointment>"
End Property

Public Property Let day(newDate As Date)
    On Error Resume Next
    mDay = newDate
End Property

Public Property Get day() As Date
    On Error Resume Next
    day = mDay
End Property

Public Property Let text(newText As String)
    On Error Resume Next
    mText = newText
End Property

Public Property Get text() As String
    On Error Resume Next
    text = mText
End Property

Public Property Get Display() As String
    On Error Resume Next
    If modI18N.GetLanguage = "en-us" Then
        Display = Format(mDay, "ddd, mmm dd, yyyy")
    Else
        Display = Format(mDay, "ddd, dd mmm yyyy")
    End If
    Display = Display & " - " & mText
End Property

Public Property Get DisplayHTML() As String
    On Error Resume Next
    DisplayHTML = "<h2>"
    If modI18N.GetLanguage = "en-us" Then
        DisplayHTML = DisplayHTML & Format(mDay, "ddd, mmm dd, yyyy")
    Else
        DisplayHTML = DisplayHTML & Format(mDay, "ddd, dd mmm yyyy")
    End If
    DisplayHTML = DisplayHTML & "</h2><p>" & mText & "</p>"
End Property
