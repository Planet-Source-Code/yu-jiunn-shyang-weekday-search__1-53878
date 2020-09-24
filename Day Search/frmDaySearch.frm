VERSION 5.00
Begin VB.Form frmWeekDaySearch 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Week Day Search"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   6255
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin VB.ComboBox cboDay 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5760
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox cboMonth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3120
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmWeekDaySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************'
'********                                      ********'
'********         Weekday Search               ********'
'********   Created by : Yu Jiunn Shyang       ********'
'********                                      ********'
'******************************************************'
' This is the source from Perpetual Calendar           '
'                                                      '
' This program is search for a certain weekday without '
' using Visual Basic 6 Built-In Function.              '
'                                                      '
' I and my friend have test this program using         '
' different date, month, and year. And we ensure 100%  '
' accurate before publish this source.                 '
'                                                      '
' This program can support year from 1001 to           '
' 999999(6 nine), don't more than 9999999(7 nine)      '
' or system will crash(for some PC).                   '
'                                                      '
' I do apologize for some badly structured codes,      '
' Kinda confusing, huh?                                '

' Anyhow, should there be further enquiries about my   '
' code, please email me at:   adrianyu_j_s@yahoo.com   '
' please do include your name.                         '

' PS:  Hehe... If you've downloaded this source from   '
'      Planet Source Code, please do vote for me,      '
'      vote would appreciate.  :-)                     '
'******************************************************'

Option Explicit
Dim result, counting, num(28), jan(28), feb(28), mar(28), apr(28), may(28), jun(28), july(28), aug(28), sept(28), oct(28), nov(28), dec(28), i, j, e As Integer
Dim found As Boolean
Dim yy1 As Long
Dim reday As String

Private Sub cboDay_KeyPress(KeyAscii As Integer)
KeyAscii = 0 'Disable keypress at Combo Box(Day)
End Sub

Private Sub cboMonth_Click()
cboDay = "" 'After Month is selected, Day must be cleared to make sure accuracy in day
cboDay.Visible = True
Call countfeb
End Sub

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
KeyAscii = 0 'Disable keypress at Combo Box(Month)
End Sub

Private Sub cmdAbout_Click()
frmAbout.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdOk_Click()
If Len(txtYear) >= 4 Then 'Check whether the Year is more than 4 digit
    If cboMonth = "" Then 'Check whether user have entered Month
        MsgBox "Please fill the Month box", vbOKOnly + vbCritical, "Month require"
    Else
        If cboDay = "" Then 'Check whether user have entered Day
            MsgBox "Please fill Day box", vbOKOnly + vbCritical, "Day require"
        Else
            Call SectionA

            For j = 1 To 28
                num(j) = i 'Save the i value in array (num(array))
                i = i + 1
            Next j
            found = False
            counting = 1
            Do Until found = True
                If txtYear.Text = num(counting) Then 'Find the array across the same array
                found = True
                Else
                    counting = counting + 1
                End If
            Loop
'********************January Formula*******************'
jan(1) = 4
jan(2) = 5
jan(3) = 6
jan(4) = 0
jan(5) = 2
jan(6) = 3
jan(7) = 4
jan(8) = 5
jan(9) = 0
jan(10) = 1
jan(11) = 2
jan(12) = 3
jan(13) = 5
jan(14) = 6
jan(15) = 0
jan(16) = 1
jan(17) = 3
jan(18) = 4
jan(19) = 5
jan(20) = 6
jan(21) = 1
jan(22) = 2
jan(23) = 3
jan(24) = 4
jan(25) = 6
jan(26) = 0
jan(27) = 1
jan(28) = 2
'*******************************************************'

'********************February Formula*******************'
feb(1) = 0
feb(2) = 1
feb(3) = 2
feb(4) = 3
feb(5) = 5
feb(6) = 6
feb(7) = 0
feb(8) = 1
feb(9) = 3
feb(10) = 4
feb(11) = 5
feb(12) = 6
feb(13) = 1
feb(14) = 2
feb(15) = 3
feb(16) = 4
feb(17) = 6
feb(18) = 0
feb(19) = 1
feb(20) = 2
feb(21) = 4
feb(22) = 5
feb(23) = 6
feb(24) = 0
feb(25) = 2
feb(26) = 3
feb(27) = 4
feb(28) = 5
'*******************************************************'

'********************March Formula**********************'
mar(1) = 0
mar(2) = 1
mar(3) = 2
mar(4) = 4
mar(5) = 5
mar(6) = 6
mar(7) = 0
mar(8) = 2
mar(9) = 3
mar(10) = 4
mar(11) = 5
mar(12) = 0
mar(13) = 1
mar(14) = 2
mar(15) = 3
mar(16) = 5
mar(17) = 6
mar(18) = 0
mar(19) = 1
mar(20) = 3
mar(21) = 4
mar(22) = 5
mar(23) = 6
mar(24) = 1
mar(25) = 2
mar(26) = 3
mar(27) = 4
mar(28) = 6
'*******************************************************'

'********************April Formula**********************'
apr(1) = 3
apr(2) = 4
apr(3) = 5
apr(4) = 0
apr(5) = 1
apr(6) = 2
apr(7) = 3
apr(8) = 5
apr(9) = 6
apr(10) = 0
apr(11) = 1
apr(12) = 3
apr(13) = 4
apr(14) = 5
apr(15) = 6
apr(16) = 1
apr(17) = 2
apr(18) = 3
apr(19) = 4
apr(20) = 6
apr(21) = 0
apr(22) = 1
apr(23) = 2
apr(24) = 4
apr(25) = 5
apr(26) = 6
apr(27) = 0
apr(28) = 2
'*******************************************************'

'********************May Formula************************'
may(1) = 5
may(2) = 6
may(3) = 0
may(4) = 2
may(5) = 3
may(6) = 4
may(7) = 5
may(8) = 0
may(9) = 1
may(10) = 2
may(11) = 3
may(12) = 5
may(13) = 6
may(14) = 0
may(15) = 1
may(16) = 3
may(17) = 4
may(18) = 5
may(19) = 6
may(20) = 1
may(21) = 2
may(22) = 3
may(23) = 4
may(24) = 6
may(25) = 0
may(26) = 1
may(27) = 2
may(28) = 4
'*******************************************************'

'********************June Formula***********************'
jun(1) = 1
jun(2) = 2
jun(3) = 3
jun(4) = 5
jun(5) = 6
jun(6) = 0
jun(7) = 1
jun(8) = 3
jun(9) = 4
jun(10) = 5
jun(11) = 6
jun(12) = 1
jun(13) = 2
jun(14) = 3
jun(15) = 4
jun(16) = 6
jun(17) = 0
jun(18) = 1
jun(19) = 2
jun(20) = 4
jun(21) = 5
jun(22) = 6
jun(23) = 0
jun(24) = 2
jun(25) = 3
jun(26) = 4
jun(27) = 5
jun(28) = 0
'*******************************************************'

'********************July Formula***********************'
july(1) = 3
july(2) = 4
july(3) = 5
july(4) = 0
july(5) = 1
july(6) = 2
july(7) = 3
july(8) = 5
july(9) = 6
july(10) = 0
july(11) = 1
july(12) = 3
july(13) = 4
july(14) = 5
july(15) = 6
july(16) = 1
july(17) = 2
july(18) = 3
july(19) = 4
july(20) = 6
july(21) = 0
july(22) = 1
july(23) = 2
july(24) = 4
july(25) = 5
july(26) = 6
july(27) = 0
july(28) = 2
'*******************************************************'

'********************August Formula*********************'
aug(1) = 6
aug(2) = 0
aug(3) = 1
aug(4) = 3
aug(5) = 4
aug(6) = 5
aug(7) = 6
aug(8) = 1
aug(9) = 2
aug(10) = 3
aug(11) = 4
aug(12) = 6
aug(13) = 0
aug(14) = 1
aug(15) = 2
aug(16) = 4
aug(17) = 5
aug(18) = 6
aug(19) = 0
aug(20) = 2
aug(21) = 3
aug(22) = 4
aug(23) = 5
aug(24) = 0
aug(25) = 1
aug(26) = 2
aug(27) = 3
aug(28) = 5
'*******************************************************'

'********************September Formula******************'
sept(1) = 2
sept(2) = 3
sept(3) = 4
sept(4) = 6
sept(5) = 0
sept(6) = 1
sept(7) = 2
sept(8) = 4
sept(9) = 5
sept(10) = 6
sept(11) = 0
sept(12) = 2
sept(13) = 3
sept(14) = 4
sept(15) = 5
sept(16) = 0
sept(17) = 1
sept(18) = 2
sept(19) = 3
sept(20) = 5
sept(21) = 6
sept(22) = 0
sept(23) = 1
sept(24) = 3
sept(25) = 4
sept(26) = 5
sept(27) = 6
sept(28) = 1
'*******************************************************'

'********************October Formula********************'
oct(1) = 4
oct(2) = 5
oct(3) = 6
oct(4) = 1
oct(5) = 2
oct(6) = 3
oct(7) = 4
oct(8) = 6
oct(9) = 0
oct(10) = 1
oct(11) = 2
oct(12) = 4
oct(13) = 5
oct(14) = 6
oct(15) = 0
oct(16) = 2
oct(17) = 3
oct(18) = 4
oct(19) = 5
oct(20) = 0
oct(21) = 1
oct(22) = 2
oct(23) = 3
oct(24) = 5
oct(25) = 6
oct(26) = 0
oct(27) = 1
oct(28) = 3
'*******************************************************'

'********************November Formula*******************'
nov(1) = 0
nov(2) = 1
nov(3) = 2
nov(4) = 4
nov(5) = 5
nov(6) = 6
nov(7) = 0
nov(8) = 2
nov(9) = 3
nov(10) = 4
nov(11) = 5
nov(12) = 0
nov(13) = 1
nov(14) = 2
nov(15) = 3
nov(16) = 5
nov(17) = 6
nov(18) = 0
nov(19) = 1
nov(20) = 3
nov(21) = 4
nov(22) = 5
nov(23) = 6
nov(24) = 1
nov(25) = 2
nov(26) = 3
nov(27) = 4
nov(28) = 6
'*******************************************************'

'********************December Formula*******************'
dec(1) = 2
dec(2) = 3
dec(3) = 4
dec(4) = 6
dec(5) = 0
dec(6) = 1
dec(7) = 2
dec(8) = 4
dec(9) = 5
dec(10) = 6
dec(11) = 0
dec(12) = 2
dec(13) = 3
dec(14) = 4
dec(15) = 5
dec(16) = 0
dec(17) = 1
dec(18) = 2
dec(19) = 3
dec(20) = 5
dec(21) = 6
dec(22) = 0
dec(23) = 1
dec(24) = 3
dec(25) = 4
dec(26) = 5
dec(27) = 6
dec(28) = 1
'*******************************************************'
        Call countday
        Label5.Visible = True
        Label4.Visible = True
        Label5.Caption = cboDay + ", " + cboMonth + ", " + txtYear + " is fell at"
        Label4.Caption = reday
        Exit Sub
        End If
    End If
        Else
        MsgBox "Please make sure you enter 4 digit year!", vbOKOnly + vbCritical, "4 digit year require"
End If
End Sub

Private Sub Form_Load()
With cboMonth
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
End With
End Sub

Sub countday()
If cboMonth = "January" Then
    result = cboDay + jan(counting)
End If
If cboMonth = "February" Then
    result = cboDay + feb(counting)
End If
If cboMonth = "March" Then
    result = cboDay + mar(counting)
End If
If cboMonth = "April" Then
    result = cboDay + apr(counting)
End If
If cboMonth = "May" Then
    result = cboDay + may(counting)
End If
If cboMonth = "June" Then
    result = cboDay + jun(counting)
End If
If cboMonth = "July" Then
    result = cboDay + july(counting)
End If
If cboMonth = "August" Then
    result = cboDay + aug(counting)
End If
If cboMonth = "September" Then
    result = cboDay + sept(counting)
End If
If cboMonth = "October" Then
    result = cboDay + oct(counting)
End If
If cboMonth = "November" Then
    result = cboDay + nov(counting)
End If
If cboMonth = "December" Then
    result = cboDay + dec(counting)
End If
Call SectionC
End Sub

Sub countfeb()
On Error Resume Next
If cboMonth = "January" Or cboMonth = "March" Or cboMonth = "May" Or cboMonth = "July" Or cboMonth = "August" Or cboMonth = "October" Or cboMonth = "December" Then
    For i = 1 To 31
        With cboDay
            .AddItem i
        End With
    Next i
ElseIf cboMonth = "April" Or cboMonth = "June" Or cboMonth = "September" Or cboMonth = "November" Then
    For i = 1 To 30
        With cboDay
            .AddItem i
        End With
    Next i
ElseIf cboMonth = "February" Then
    If txtYear Mod 4 = 0 Then 'Check whether the year is leap year
        For i = 1 To 29
            With cboDay
                .AddItem i
            End With
        Next i
    Else
        For i = 1 To 28
            With cboDay
                .AddItem i
            End With
        Next i
    End If
End If
End Sub

Private Sub txtYear_Change()
cboDay.Clear
Call countfeb
End Sub

Sub SectionA()
found = False
yy1 = 1001 'Start of the year
Do Until found = True
    If txtYear >= yy1 And txtYear <= yy1 + 27 Then 'Find the year that is between
        i = yy1
        found = True
    Else
        yy1 = yy1 + 28
    End If
Loop
End Sub

Sub SectionC()
If result = 1 Or result = 8 Or result = 15 Or result = 22 Or result = 29 Or result = 36 Then
    reday = "Sunday"
ElseIf result = 2 Or result = 9 Or result = 16 Or result = 23 Or result = 30 Or result = 37 Then
    reday = "Monday"
ElseIf result = 3 Or result = 10 Or result = 17 Or result = 24 Or result = 31 Then
    reday = "Tuesday"
ElseIf result = 4 Or result = 11 Or result = 18 Or result = 25 Or result = 32 Then
    reday = "Wednesday"
ElseIf result = 5 Or result = 12 Or result = 19 Or result = 26 Or result = 33 Then
    reday = "Thursday"
ElseIf result = 6 Or result = 13 Or result = 20 Or result = 27 Or result = 34 Then
    reday = "Friday"
ElseIf result = 7 Or result = 14 Or result = 21 Or result = 28 Or result = 35 Then
    reday = "Saturday"
End If
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then 'Only allow number in year
    Exit Sub
Else
    KeyAscii = 0 'Disable any key
End If
End Sub
