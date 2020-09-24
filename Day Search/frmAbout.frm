VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   5100
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7755
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3520.11
   ScaleMode       =   0  'User
   ScaleWidth      =   7282.346
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   5775
   End
   Begin VB.Line Line1 
      X1              =   338.059
      X2              =   1577.607
      Y1              =   496.957
      Y2              =   496.957
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Coding: Yu Jiunn Shyang && Goh Chun Hean   Designer: Yu Jiunn Shyang        Documentation: Lee Yew Siong"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Credit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub
