VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voodoo Unload"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4125
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quit"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   975
   End
   Begin VB.Timer T1 
      Left            =   360
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The easiest of all to create

Private Sub Command1_Click()

T1.Enabled = True

End Sub

Private Sub Form_Load()

'You can change the values
'directly from properties
T1.Enabled = False
T1.Interval = 5

End Sub

Private Sub T1_Timer()

'Me.Height goes - 200
Me.Height = Me.Height - 200

'Until Me.Height reaches 510
'where it unloads
If Me.Height = 510 Then Unload Me

End Sub
