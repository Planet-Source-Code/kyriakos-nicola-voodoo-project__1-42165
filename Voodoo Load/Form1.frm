VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   30
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   30
   ShowInTaskbar   =   0   'False
   Begin VB.Timer T2 
      Left            =   600
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quit"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Timer T1 
      Left            =   120
      Top             =   120
   End
   Begin VB.Image Img1 
      Height          =   4125
      Left            =   240
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Very simple examble to make special
'effects for your project

Private Sub Command1_Click()

T2.Enabled = True

End Sub

Private Sub Form_Load()

'As you can see I haven't changed the following values
'from the properties so that every one can understand
'what this project needs to work. If you want you can
'change them directly from the properties.
T1.Interval = 5
T2.Interval = 5

T2.Enabled = False

End Sub

Private Sub T1_Timer()

'The following commands are responsible for all the work.

'The first 2(two) lines make the form grow bigger,
Me.Height = Me.Height + 30
Me.Width = Me.Width + 30

'And the these 2(two) lines make the form move
Me.Left = Me.Left + 20
Me.Top = Me.Top + 20

'This command disables T1(Timer1) when Me.Height goes to 6000
If Me.Height = 6000 Then T1.Enabled = False
'Here you can put width too, it doesn't matter we only
'want a value where our T1 commands will stop being repeated

End Sub

Private Sub T2_Timer()

'Well T2 commands are the opposite of T1 commands
Me.Height = Me.Height - 30
Me.Width = Me.Width - 30

Me.Left = Me.Left - 20
Me.Top = Me.Top - 20

'Now we want the smallest value that our form can reach.
'As we all know the smallest value we can get in non
'negative numbers is 0(zero).
'Well, I found out that the smallest value we can get in
'any form in VB is 120 for Height and 510 for width.
'So, when Me.Height reaches 120 the project unloads!
If Me.Height = 120 Then Unload Me

End Sub
